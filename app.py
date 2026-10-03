import os
import re
import io
import json
import base64
import threading
import time
import html
from datetime import date

import docx
import pypdf
import streamlit as st
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from dotenv import load_dotenv
from openai import OpenAI

# PDF Generation Imports
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.fonts import addMapping

# Urdu / Arabic script support in PDF (optional)
try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    RTL_OK = True
except Exception:
    RTL_OK = False

# ============================================================
# LOAD ENVIRONMENT VARIABLES & PAGE CONFIG
# ============================================================

load_dotenv()

st.set_page_config(
    page_title="WSA Educational Test Series & Paper Builder",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# LLM PROVIDERS CONFIG
# ============================================================

PROVIDERS = {
    "Groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "env": "GROQ_API_KEY",
        "models": ["openai/gpt-oss-120b", "openai/gpt-oss-20b"],
        "vision_models": ["qwen/qwen3.8-27b"],
    },
    "Gemini": {
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "env": "GEMINI_API_KEY",
        "models": ["gemini-flash-latest", "gemini-3.1-flash-lite"],
        "vision_models": ["gemini-flash-latest", "gemini-3.1-flash-lite"],
    },
}

AUTO_MODE = "Auto (Fallback)"
AUTO_ORDER = ["Groq", "Gemini"]

MAX_REFERENCE_CHARS = 9000
LETTERS = "ABCD"
USER_MESSAGE_PREFIX = "[user] "

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")
UPLOAD_TYPES = ["pdf", "docx", "txt", "jpg", "jpeg", "png", "webp"]
UPLOAD_HINT = "PDF, DOCX, TXT, JPG, PNG"


def is_image(uploaded_file):
    return uploaded_file is not None and uploaded_file.name.lower().endswith(IMAGE_EXTS)


def prepare_image_b64(uploaded_file, max_side=1800):
    """Shrinks a photo and converts it to base64."""
    from PIL import Image, ImageOps

    img = Image.open(io.BytesIO(uploaded_file.getvalue()))
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


# ============================================================
# SECRETS + QUOTA PROTECTION
# ============================================================


def get_secret(name, default=""):
    try:
        val = st.secrets.get(name)
        if val:
            return str(val).strip()
    except Exception:
        pass
    return os.getenv(name, default).strip()


_groq_model_override = get_secret("GROQ_MODEL")
if _groq_model_override:
    PROVIDERS["Groq"]["models"].insert(0, _groq_model_override)

_groq_vision_override = get_secret("GROQ_VISION_MODEL")
if _groq_vision_override:
    PROVIDERS["Groq"]["vision_models"].insert(0, _groq_vision_override)

_gemini_override = get_secret("GEMINI_MODEL")
if _gemini_override:
    PROVIDERS["Gemini"]["models"].insert(0, _gemini_override)
    PROVIDERS["Gemini"]["vision_models"].insert(0, _gemini_override)

SESSION_COOLDOWN_SEC = 15
SESSION_MAX_CALLS = 20
try:
    DAILY_LIMIT = int(get_secret("DAILY_LIMIT") or 400)
except ValueError:
    DAILY_LIMIT = 400


@st.cache_resource
def usage_store():
    return {"day": date.today().isoformat(), "count": 0, "lock": threading.Lock()}


def allow_ai_call(units=1):
    ss = st.session_state
    now = time.time()
    wait = SESSION_COOLDOWN_SEC - (now - ss.get("_last_ai_call", 0))
    if wait > 0:
        st.warning(f"⏳ Please wait {int(wait) + 1} seconds before trying again.")
        return False
    if ss.get("_ai_calls", 0) + units > SESSION_MAX_CALLS:
        st.error("You have reached the usage limit for this session. Please try again later.")
        return False

    store = usage_store()
    with store["lock"]:
        today = date.today().isoformat()
        if store["day"] != today:
            store["day"], store["count"] = today, 0
        if store["count"] + units > DAILY_LIMIT:
            st.error("The daily free limit has been reached. Please come back tomorrow.")
            return False
        store["count"] += units

    ss["_last_ai_call"] = now
    ss["_ai_calls"] = ss.get("_ai_calls", 0) + units
    return True


def show_error(e):
    msg = str(e)
    low = msg.lower()
    print("AI error:", msg)
    if msg.startswith(USER_MESSAGE_PREFIX):
        st.error("❌ " + msg[len(USER_MESSAGE_PREFIX):])
    elif "429" in msg or "rate limit" in low or "quota" in low:
        st.error("⏳ The AI service is busy right now. Please wait a minute and try again.")
    elif "401" in msg or "api key" in low or "authentication" in low:
        st.error("🔑 The AI service is not set up correctly. Please contact the administrator.")
    else:
        st.error("❌ Something went wrong. Please try again.")
    if get_secret("DEBUG") == "1":
        with st.expander("Technical details (debug mode)"):
            st.code(msg)


# ============================================================
# CUSTOM CSS (system fonts, mobile-first, sidebar-friendly)
# ============================================================

st.markdown(
    """
    <style>
        :root {
            --navy: #1E3A8A;
            --blue: #2563EB;
            --sky: #38BDF8;
            --ink: #0F172A;
            --muted: #64748B;
            --line: #E2E8F0;
        }

        .stApp {
            background: linear-gradient(180deg, #EAF0FB 0%, #F8FAFC 260px);
            font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
        }
        .block-container { padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1180px; }

        #MainMenu, footer { visibility: hidden; }

        /* ---------- Sidebar ---------- */
        [data-testid="stSidebar"] {
            background: #F8FAFF;
            border-right: 1px solid #E2E8F0;
        }
        [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
            padding: 10px 12px;
            border-radius: 10px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.15s ease;
        }
        [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
            background: #E8EEF9;
        }

        /* ---------- Hero ---------- */
        .hero {
            background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 55%, #38BDF8 120%);
            padding: 30px 34px;
            border-radius: 18px;
            color: #fff;
            box-shadow: 0 12px 30px -10px rgba(30, 58, 138, 0.45);
            margin-bottom: 22px;
        }
        .hero-label { font-size: 0.85rem; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; opacity: 0.85; }
        .hero-title { font-size: 2.2rem; font-weight: 800; margin: 6px 0 8px 0; letter-spacing: -0.5px; line-height: 1.15; }
        .hero-subtitle { font-size: 1.02rem; opacity: 0.92; margin: 0 0 16px 0; max-width: 720px; }
        .chip {
            display: inline-block;
            background: rgba(255,255,255,0.16);
            border: 1px solid rgba(255,255,255,0.28);
            padding: 5px 12px;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 500;
            margin: 0 8px 6px 0;
        }

        /* ---------- Cards ---------- */
        div[class*="st-key-card"] {
            background: #FFFFFF;
            border: 1px solid var(--line) !important;
            border-radius: 14px;
            box-shadow: 0 2px 12px rgba(15, 23, 42, 0.04);
        }
        [data-testid="stMetric"] {
            background: #F5F8FF;
            border: 1px solid var(--line);
            padding: 12px 16px;
            border-radius: 12px;
        }
        [data-testid="stMetricValue"] { color: var(--navy); font-weight: 800; }

        /* ---------- Buttons ---------- */
        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
            color: #fff;
            font-weight: 700;
            border: none;
            border-radius: 12px;
            padding: 12px 24px;
            box-shadow: 0 6px 16px -6px rgba(37, 99, 235, 0.6);
        }
        .stButton > button[kind="primary"]:hover { color: #fff; }
        .stDownloadButton > button, .stButton > button[kind="secondary"] {
            border-radius: 12px;
            border: 1px solid #C7D5F0;
            font-weight: 600;
            color: var(--navy);
            background: #F8FAFF;
        }
        .stDownloadButton > button:hover, .stButton > button[kind="secondary"]:hover {
            border-color: var(--blue);
            color: var(--blue);
        }

        /* ---------- Inputs ---------- */
        [data-testid="stFileUploaderDropzone"] {
            border-radius: 12px;
            border: 1.5px dashed #93B4F5;
            background: #F8FAFF;
        }
        .section-title { font-size: 1.05rem; font-weight: 700; color: var(--navy); margin: 2px 0 2px 0; }
        .section-hint { font-size: 0.86rem; color: var(--muted); margin-bottom: 8px; }
        .ai-note { font-size: 0.8rem; color: var(--muted); text-align: center; margin-top: 6px; }

        /* ---------- Footer ---------- */
        .branding-card {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            color: #F8FAFC;
            padding: 26px 22px;
            border-radius: 16px;
            border-top: 4px solid #3B82F6;
            text-align: center;
            margin-top: 40px;
            box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.3);
        }
        .hadith-quote { font-size: 1.02rem; font-style: italic; font-weight: 500; color: #F1F5F9; line-height: 1.6; margin-bottom: 6px; }
        .hadith-ref { font-size: 0.82rem; color: #60A5FA; font-weight: 600; letter-spacing: 0.3px; margin-bottom: 16px; }
        .footer-divider { border: 0; border-top: 1px solid #334155; margin: 16px auto; width: 80%; }
        .branding-name { font-size: 1.1rem; font-weight: 700; color: #38BDF8; margin-bottom: 4px; }
        .branding-tag { font-size: 0.82rem; color: #94A3B8; letter-spacing: 0.6px; text-transform: uppercase; }

        /* ---------- Responsive ---------- */
        @media (max-width: 760px) {
            .block-container { padding-top: .65rem; padding-bottom: 1.2rem; }
            .hero { padding: 22px 18px; margin-bottom: 15px; border-radius: 17px; }
            .hero-label { font-size: .72rem; letter-spacing: .08em; }
            .hero-title { font-size: clamp(1.45rem, 6vw, 1.9rem); line-height: 1.16; }
            .hero-subtitle { font-size: .91rem; line-height: 1.55; }
            .hero .chip { display: none; }
            [data-testid="stVerticalBlockBorderWrapper"] { border-radius: 14px; }
            [data-testid="stMetric"] { padding: 10px 11px; }
            [data-testid="stMetricLabel"] { font-size: .76rem; }
            [data-testid="stMetricValue"] { font-size: 1.25rem; }
            .section-title { font-size: 1rem; }
            .section-hint { font-size: .82rem; line-height: 1.5; }
            .stButton > button, .stDownloadButton > button { width: 100%; min-height: 48px; font-size: .91rem; }
            .branding-card { padding: 22px 15px; margin-top: 26px; }
            .hadith-quote { font-size: .91rem; }
            .hadith-ref, .branding-tag { font-size: .73rem; line-height: 1.5; }
            [data-testid="stSidebar"] {
                width: 80vw !important;
                min-width: 80vw !important;
            }
            [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
                padding: 13px 12px;
                font-size: .95rem;
            }
        }
        @media (max-width: 420px) {
            .block-container { padding-left: 10px; padding-right: 10px; }
            .hero { padding: 19px 15px; }
            .hero-title { font-size: 1.42rem; }
            [data-testid="stMetricValue"] { font-size: 1.12rem; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PDF FONTS (LAZY LOADED)
# ============================================================


def register_pdf_fonts():
    candidates = [
        (
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/ariali.ttf",
            "C:/Windows/Fonts/arialbi.ttf",
        ),
        (
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/System/Library/Fonts/Supplemental/Arial Italic.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf",
        ),
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
        ),
    ]
    for reg, bold, ital, bi in candidates:
        if all(os.path.exists(p) for p in (reg, bold, ital, bi)):
            try:
                pdfmetrics.registerFont(TTFont("WSA", reg))
                pdfmetrics.registerFont(TTFont("WSA-B", bold))
                pdfmetrics.registerFont(TTFont("WSA-I", ital))
                pdfmetrics.registerFont(TTFont("WSA-BI", bi))
                addMapping("WSA", 0, 0, "WSA")
                addMapping("WSA", 1, 0, "WSA-B")
                addMapping("WSA", 0, 1, "WSA-I")
                addMapping("WSA", 1, 1, "WSA-BI")
                return "WSA", "WSA-B", "WSA-I"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


_PDF_FONTS = None


def get_pdf_fonts():
    global _PDF_FONTS
    if _PDF_FONTS is None:
        _PDF_FONTS = register_pdf_fonts()
    return _PDF_FONTS


ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")
EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FFFF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u200d]"
)


def is_rtl_line(line):
    letters = [c for c in line if c.isalpha()]
    if not letters:
        return False
    arabic = sum(1 for c in letters if ARABIC_RE.match(c))
    return arabic / len(letters) > 0.5


def shape_text(line):
    if RTL_OK and ARABIC_RE.search(line):
        try:
            return get_display(arabic_reshaper.reshape(line))
        except Exception:
            return line
    return line


def clean_text(text):
    text = str(text or "")
    text = text.replace("**", "").replace("`", "")
    return text.strip()


def rich(text, style):
    lines = clean_text(text).split("\n")
    body = "<br/>".join(html.escape(shape_text(ln), quote=False) for ln in lines)
    non_empty = [ln for ln in lines if ln.strip()]
    if RTL_OK and non_empty and all(is_rtl_line(ln) for ln in non_empty):
        style = ParagraphStyle(name=style.name + "_rtl", parent=style, alignment=TA_RIGHT)
    return Paragraph(body, style)


# ============================================================
# JSON SCHEMA + PARSING
# ============================================================

JSON_SCHEMA_TEXT = """{
  "mcqs":  [ {"q": "question text", "options": ["option 1", "option 2", "option 3", "option 4"], "answer": "B"} ],
  "short": [ {"q": "question text", "answer": "model answer in 2-3 lines"} ],
  "long":  [ {"q": "question text (use (a), (b), (c) parts on new lines if suitable)", "answer": "key points of model answer"} ]
}"""


def normalize_paper(obj, mcq_count, short_count, long_count):
    mcqs = []
    for item in obj.get("mcqs") or []:
        if not isinstance(item, dict):
            continue
        q = clean_text(item.get("q") or item.get("question") or "")
        opts = item.get("options") or []
        if isinstance(opts, dict):
            opts = [opts[k] for k in sorted(opts)]
        cleaned = []
        for i, o in enumerate(list(opts)[:4]):
            o = clean_text(o)
            m = re.match(r"^\(?([A-Da-d])[\)\.:]\s+", o)
            if m and m.group(1).upper() == LETTERS[i]:
                o = o[m.end():].strip()
            cleaned.append(o)
        if not q or len(cleaned) < 2:
            continue

        ans = clean_text(item.get("answer", ""))
        letter = ""
        m = re.match(r"^\(?([A-Da-d])\)?(?:[\)\.:\s-]|$)", ans)
        if m:
            letter = m.group(1).upper()
        else:
            for i, o in enumerate(cleaned):
                if ans and ans.lower() == o.lower():
                    letter = LETTERS[i]
                    break
        if letter and LETTERS.index(letter) >= len(cleaned):
            letter = ""
        mcqs.append({"q": q, "options": cleaned, "answer": letter or "?"})

    def simple_items(key):
        out = []
        for item in obj.get(key) or []:
            if isinstance(item, str):
                item = {"q": item}
            if not isinstance(item, dict):
                continue
            q = clean_text(item.get("q") or item.get("question") or "")
            a = item.get("answer", "")
            if isinstance(a, list):
                a = "\n".join(str(x) for x in a)
            if q:
                out.append({"q": q, "answer": clean_text(a)})
        return out

    return {
        "mcqs": mcqs[:mcq_count],
        "short": simple_items("short")[:short_count],
        "long": simple_items("long")[:long_count],
    }


def parse_paper_json(raw, mcq_count, short_count, long_count):
    txt = (raw or "").strip()
    txt = re.sub(r"^```(?:json)?", "", txt, flags=re.I).strip()
    txt = re.sub(r"```$", "", txt).strip()
    start, end = txt.find("{"), txt.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("JSON not found in AI response")
    obj = json.loads(txt[start : end + 1])
    return normalize_paper(obj, mcq_count, short_count, long_count)


# ============================================================
# PAPER -> TEXT / MARKDOWN
# ============================================================

INSTRUCTIONS = [
    "Attempt all questions.",
    "In MCQs, choose only one correct option.",
    "Write neatly and clearly. Cutting / overwriting will not be entertained.",
]


def section_list(data, meta):
    out = []
    letters = iter("ABC")
    if data["mcqs"]:
        n, m = len(data["mcqs"]), meta["mcq_marks"]
        out.append((next(letters), "Multiple Choice Questions", f"{n} × {m} = {n * m} Marks", "mcq", data["mcqs"]))
    if data["short"]:
        n, m = len(data["short"]), meta["short_marks"]
        out.append((next(letters), "Short Questions", f"{n} × {m} = {n * m} Marks", "short", data["short"]))
    if data["long"]:
        n, m = len(data["long"]), meta["long_marks"]
        out.append((next(letters), "Long Questions", f"{n} × {m} = {n * m} Marks", "long", data["long"]))
    return out


def compute_total_marks(data, meta):
    return (
        len(data["mcqs"]) * meta["mcq_marks"]
        + len(data["short"]) * meta["short_marks"]
        + len(data["long"]) * meta["long_marks"]
    )


def answer_key_lines(data, meta):
    lines = []
    for letter, title, _marks, kind, items in section_list(data, meta):
        lines.append(f"SECTION {letter} — {title}")
        if kind == "mcq":
            for i, it in enumerate(items, 1):
                lines.append(f"  Q{i}: {it['answer']}")
        else:
            for i, it in enumerate(items, 1):
                lines.append(f"  Q{i}: {it['answer'] or '-'}")
        lines.append("")
    return lines


def paper_to_text(data, meta, include_key=True):
    L = [
        meta["institute"].upper(),
        meta["category"],
        "=" * 64,
        f"Subject: {meta['topic']}",
        f"Time Allowed: {meta['time_min']} Minutes    Total Marks: {meta['total_marks']}",
        "Name: ______________________   Roll No: __________   Date: __________",
        "=" * 64,
        "Instructions:",
    ]
    L += [f"  {i}. {t}" for i, t in enumerate(INSTRUCTIONS, 1)]
    L.append("")

    for letter, title, marks, kind, items in section_list(data, meta):
        L.append(f"SECTION {letter} — {title}   ({marks})")
        L.append("-" * 64)
        for i, it in enumerate(items, 1):
            L.append(f"Q{i}. {it['q']}")
            if kind == "mcq":
                for j, opt in enumerate(it["options"]):
                    L.append(f"     ({LETTERS[j]}) {opt}")
            L.append("")
        L.append("")

    L.append("— End of Paper —")

    if include_key:
        L += ["", "=" * 64, "ANSWER KEY (Teacher Copy)", "=" * 64]
        L += answer_key_lines(data, meta)
    return "\n".join(L)


def paper_to_markdown(data, meta):
    def br(t):
        return t.replace("\n", "  \n")

    L = [
        f"## 🎓 {meta['institute']}",
        f"**{meta['category']}**",
        "",
        f"**Subject:** {meta['topic']}  |  **Time:** {meta['time_min']} min  |  **Total Marks:** {meta['total_marks']}",
        "",
        "**Instructions:** " + " ".join(INSTRUCTIONS),
        "",
        "---",
    ]
    for letter, title, marks, kind, items in section_list(data, meta):
        L += ["", f"### Section {letter}: {title} *({marks})*", ""]
        for i, it in enumerate(items, 1):
            L.append(f"**Q{i}.** {br(it['q'])}")
            if kind == "mcq":
                L.append("")
                for j, opt in enumerate(it["options"]):
                    L.append(f"&nbsp;&nbsp;&nbsp;&nbsp;**({LETTERS[j]})** {br(opt)}  ")
            L.append("")
    return "\n".join(L)


def answer_key_markdown(data, meta):
    L = []
    for letter, title, _marks, kind, items in section_list(data, meta):
        L.append(f"**Section {letter}: {title}**")
        if kind == "mcq":
            L.append("  ".join(f"`Q{i}: {it['answer']}`" for i, it in enumerate(items, 1)))
        else:
            for i, it in enumerate(items, 1):
                L.append(f"- **Q{i}:** {it['answer'] or '-'}".replace("\n", " "))
        L.append("")
    return "\n".join(L)


# ============================================================
# PROFESSIONAL PDF
# ============================================================

NAVY = colors.HexColor("#1E3A8A")
SLATE = colors.HexColor("#334155")
LIGHT = colors.HexColor("#E8EEF9")
GRID = colors.HexColor("#94A3B8")


def pdf_styles():
    PDF_FONT, PDF_FONT_B, PDF_FONT_I = get_pdf_fonts()
    base = getSampleStyleSheet()["Normal"]

    def mk(name, **kw):
        kw.setdefault("fontName", PDF_FONT)
        kw.setdefault("fontSize", 10.5)
        kw.setdefault("leading", 14)
        kw.setdefault("textColor", colors.HexColor("#0F172A"))
        return ParagraphStyle(name, parent=base, **kw)

    return {
        "inst": mk("inst", fontName=PDF_FONT_B, fontSize=17, leading=21, textColor=NAVY, alignment=TA_CENTER),
        "exam": mk("exam", fontName=PDF_FONT_B, fontSize=11.5, leading=15, textColor=SLATE, alignment=TA_CENTER, spaceAfter=6),
        "cell": mk("cell", fontSize=9.5, leading=12.5),
        "cellb": mk("cellb", fontName=PDF_FONT_B, fontSize=9.5, leading=12.5),
        "sec": mk("sec", fontName=PDF_FONT_B, fontSize=11, textColor=NAVY),
        "secr": mk("secr", fontName=PDF_FONT_B, fontSize=10, textColor=NAVY, alignment=TA_RIGHT),
        "qn": mk("qn", fontName=PDF_FONT_B),
        "q": mk("q"),
        "opt": mk("opt", fontSize=10.2, leading=13.5),
        "small": mk("small", fontSize=8.5, leading=11.5, textColor=SLATE),
        "center": mk("center", fontSize=9, alignment=TA_CENTER, textColor=SLATE),
        "hadith": mk("hadith", fontName=PDF_FONT_I, fontSize=8.5, leading=12, alignment=TA_CENTER, textColor=SLATE),
        "hadithref": mk("hadithref", fontSize=7.5, alignment=TA_CENTER, textColor=colors.HexColor("#2563EB")),
    }


def _page_footer(canvas, doc):
    PDF_FONT = get_pdf_fonts()[0]
    canvas.saveState()
    w, _h = A4
    canvas.setStrokeColor(GRID)
    canvas.setLineWidth(0.4)
    canvas.line(18 * mm, 14 * mm, w - 18 * mm, 14 * mm)
    canvas.setFont(PDF_FONT, 8)
    canvas.setFillColor(SLATE)
    canvas.drawString(18 * mm, 9.5 * mm, "WSA Educational Test Series")
    canvas.drawCentredString(w / 2, 9.5 * mm, f"Page {doc.page}")
    canvas.drawRightString(w - 18 * mm, 9.5 * mm, "Designed by Waheed Ali Hamouzai")
    canvas.restoreState()


def _header_flowables(meta, S, width, title_line, with_student_fields=True):
    fl = [
        rich(meta["institute"].upper(), S["inst"]),
        rich(title_line, S["exam"]),
    ]
    rows = [
        [rich("Subject / Topic:", S["cellb"]), rich(meta["topic"], S["cell"]), rich("Time Allowed:", S["cellb"]), rich(f"{meta['time_min']} Minutes", S["cell"])],
        [rich("Exam:", S["cellb"]), rich(meta["category"], S["cell"]), rich("Total Marks:", S["cellb"]), rich(str(meta["total_marks"]), S["cell"])],
    ]
    if with_student_fields:
        rows.append([rich("Student Name:", S["cellb"]), rich("", S["cell"]), rich("Roll No:", S["cellb"]), rich("", S["cell"])])
    cw = [36 * mm, width - 36 * mm - 32 * mm - 32 * mm, 32 * mm, 32 * mm]
    t = Table(rows, colWidths=cw)
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.8, NAVY),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
        ("BACKGROUND", (0, 0), (0, -1), LIGHT),
        ("BACKGROUND", (2, 0), (2, -1), LIGHT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    fl.append(t)
    fl.append(Spacer(1, 8))
    return fl


def _section_bar(letter, title, marks, S, width):
    bar = Table(
        [[rich(f"SECTION {letter}  —  {title}", S["sec"]), rich(marks, S["secr"])]],
        colWidths=[width * 0.65, width * 0.35],
    )
    bar.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("LINEBELOW", (0, 0), (-1, -1), 1.2, NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return bar


def _question_row(num_label, text, S, width):
    t = Table(
        [[rich(num_label, S["qn"]), rich(text, S["q"])]],
        colWidths=[13 * mm, width - 13 * mm],
    )
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _options_table(options, S, width):
    indent, lab = 13 * mm, 8 * mm
    use_two_cols = all(len(o) <= 38 and "\n

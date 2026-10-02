import base64
import html
import io
import json
import os
import re
import threading
import time
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

# Urdu / Arabic script support in PDF (optional - the app still works without it)
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
    initial_sidebar_state="collapsed",
)

# ============================================================
# LLM PROVIDERS CONFIG (all use OpenAI-compatible endpoints)
# ------------------------------------------------------------
# Model names and free limits change often. If a model is retired, set
# GROQ_MODEL / GROQ_VISION_MODEL / GEMINI_MODEL in the secrets (no code change).
#   Groq:        https://console.groq.com/keys
#   Gemini:      https://aistudio.google.com/apikey
# ============================================================

PROVIDERS = {
    "Groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "env": "GROQ_API_KEY",
        "models": ["openai/gpt-oss-120b", "openai/gpt-oss-20b"],
        # Used to read text from photos (Groq vision models are in preview)
        "vision_models": ["qwen/qwen3.8-27b"],
    },
    "Gemini": {
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "env": "GEMINI_API_KEY",
        # The "-latest" alias always points to Google's current Flash model,
        # so the app keeps working when an older model is retired.
        "models": ["gemini-flash-latest", "gemini-3.1-flash-lite"],
        "vision_models": ["gemini-flash-latest", "gemini-3.1-flash-lite"],
    },
}

AUTO_MODE = "Auto (Fallback)"
AUTO_ORDER = ["Groq", "Gemini"]

# Free tiers have small token limits, so reference text is capped
MAX_REFERENCE_CHARS = 9000

LETTERS = "ABCD"

# Messages starting with this prefix are safe to show to users as-is
USER_MESSAGE_PREFIX = "[user] "

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")
UPLOAD_TYPES = ["pdf", "docx", "txt", "jpg", "jpeg", "png", "webp"]
UPLOAD_HINT = "PDF, DOCX, TXT, JPG, PNG"


def is_image(uploaded_file):
    return uploaded_file is not None and uploaded_file.name.lower().endswith(IMAGE_EXTS)


def prepare_image_b64(uploaded_file, max_side=1800):
    """Shrinks a photo and converts it to base64 (keeps it within API size limits)."""
    from PIL import Image, ImageOps

    img = Image.open(io.BytesIO(uploaded_file.getvalue()))
    img = ImageOps.exif_transpose(img)  # fixes sideways / upside-down phone photos
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


# ============================================================
# SECRETS + SHARED-QUOTA PROTECTION
# ------------------------------------------------------------
# API keys are never shown to users. They live only on the server, in the
# app "secrets" or in environment variables:
#   - Streamlit Cloud: App settings -> Secrets
#   - Local computer : .streamlit/secrets.toml  or  .env file
# ============================================================


def get_secret(name, default=""):
    try:
        val = st.secrets.get(name)
        if val:
            return str(val).strip()
    except Exception:
        pass
    return os.getenv(name, default).strip()


# Optional: if a model is retired, set GROQ_MODEL in the secrets (no code change needed)
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

SESSION_COOLDOWN_SEC = 15  # minimum gap between two AI requests from one user
SESSION_MAX_CALLS = 20  # maximum AI requests in one browser session
try:
    DAILY_LIMIT = int(get_secret("DAILY_LIMIT") or 400)  # total AI requests per day (all users)
except ValueError:
    DAILY_LIMIT = 400


@st.cache_resource
def usage_store():
    """Counter shared by all users (lives as long as the server runs)."""
    return {"day": date.today().isoformat(), "count": 0, "lock": threading.Lock()}


def allow_ai_call(units=1):
    """Protects the shared free API quota. Returns True if the request may go ahead."""
    ss = st.session_state
    now = time.time()
    wait = SESSION_COOLDOWN_SEC - (now - ss.get("_last_ai_call", 0))
    if wait > 0:
        st.warning(f"⏳ Please wait {int(wait) + 1} seconds before trying again.")
        return False
    if ss.get("_ai_calls", 0) + units > SESSION_MAX_CALLS:
        st.error(
            "You have reached the usage limit for this session. Please try again later."
        )
        return False

    store = usage_store()
    with store["lock"]:
        today = date.today().isoformat()
        if store["day"] != today:
            store["day"], store["count"] = today, 0
        if store["count"] + units > DAILY_LIMIT:
            st.error(
                "The daily free limit has been reached. Please come back tomorrow."
            )
            return False
        store["count"] += units

    ss["_last_ai_call"] = now
    ss["_ai_calls"] = ss.get("_ai_calls", 0) + units
    return True


def show_error(e):
    """Shows a friendly message. Technical details only appear when DEBUG=1 is set."""
    msg = str(e)
    low = msg.lower()
    print("AI error:", msg)  # visible in the server logs only
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
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

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
            font-family: 'Inter', sans-serif;
        }
        .block-container { padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1180px; }

        /* Hide Streamlit menu, footer and the sidebar (settings live on the server) */
        #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"],
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapsedControl"] { display: none !important; }

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

        /* ---------- Tabs ---------- */
        .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
            background: #E3EAF8;
            padding: 6px;
            border-radius: 14px;
            border-bottom: none;
        }
        .stTabs [data-baseweb="tab"] {
            height: 44px;
            padding: 0 20px;
            border-radius: 10px;
            font-weight: 600;
            color: #475569;
            background: transparent;
        }
        .stTabs [aria-selected="true"] {
            background: #FFFFFF;
            color: var(--navy);
            box-shadow: 0 1px 6px rgba(15, 23, 42, 0.12);
        }
        .stTabs [data-baseweb="tab-highlight"],
        .stTabs [data-baseweb="tab-border"] { display: none; }

        /* ---------- Buttons ---------- */
        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
            color: #fff;
            font-weight: 700;
            border: none;
            border-radius: 12px;
            padding: 12px 24px;
            box-shadow: 0 6px 16px -6px rgba(37, 99, 235, 0.6);
            transition: transform 0.12s ease, box-shadow 0.12s ease;
        }
        .stButton > button[kind="primary"]:hover {
            transform: translateY(-1px);
            box-shadow: 0 10px 20px -8px rgba(37, 99, 235, 0.7);
            color: #fff;
        }
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
        .branding-name { font-size: 1.1rem; font-weight: 700; color: #38BDF8; margin-bottom: 4px; letter-spacing: 0.2px; }
        .branding-tag { font-size: 0.82rem; color: #94A3B8; letter-spacing: 0.6px; text-transform: uppercase; }

        @media (max-width: 640px) {
            .hero { padding: 22px 20px; }
            .hero-title { font-size: 1.6rem; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PDF FONT + URDU (RTL) HELPERS
# ============================================================


def register_pdf_fonts():
    """Finds a Unicode font (Windows Arial / Linux DejaVu). Falls back to Helvetica."""
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


PDF_FONT, PDF_FONT_B, PDF_FONT_I = register_pdf_fonts()

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
    """Prepares Urdu/Arabic text for the PDF (joined letters + right-to-left)."""
    if RTL_OK and ARABIC_RE.search(line):
        try:
            return get_display(arabic_reshaper.reshape(line))
        except Exception:
            return line
    return line


def clean_text(text):
    """Removes markdown characters (** and `) from AI output."""
    text = str(text or "")
    text = text.replace("**", "").replace("`", "")
    return text.strip()


def rich(text, style):
    """Safe Paragraph: HTML escape + Urdu shaping + line breaks."""
    lines = clean_text(text).split("\n")
    body = "<br/>".join(html.escape(shape_text(ln), quote=False) for ln in lines)
    non_empty = [ln for ln in lines if ln.strip()]
    if RTL_OK and non_empty and all(is_rtl_line(ln) for ln in non_empty):
        style = ParagraphStyle(
            name=style.name + "_rtl", parent=style, alignment=TA_RIGHT
        )
    return Paragraph(body, style)


# ============================================================
# PROMPT + JSON PARSING FOR STRUCTURED PAPER
# ============================================================

JSON_SCHEMA_TEXT = """{
  "mcqs":  [ {"q": "question text", "options": ["option 1", "option 2", "option 3", "option 4"], "answer": "B"} ],
  "short": [ {"q": "question text", "answer": "model answer in 2-3 lines"} ],
  "long":  [ {"q": "question text (use (a), (b), (c) parts on new lines if suitable)", "answer": "key points of model answer"} ]
}"""


def normalize_paper(obj, mcq_count, short_count, long_count):
    """Converts the AI's JSON into a clean, safe structure."""
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
            # remove the prefix only when it is the correct letter for that position
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
# PAPER -> TEXT / MARKDOWN (on-screen + txt download)
# ============================================================

INSTRUCTIONS = [
    "Attempt all questions.",
    "In MCQs, choose only one correct option.",
    "Write neatly and clearly. Cutting / overwriting will not be entertained.",
]


def section_list(data, meta):
    """Har section: (letter, title, marks_text, kind, items)."""
    out = []
    letters = iter("ABC")
    if data["mcqs"]:
        n, m = len(data["mcqs"]), meta["mcq_marks"]
        out.append(
            (
                next(letters),
                "Multiple Choice Questions",
                f"{n} × {m} = {n * m} Marks",
                "mcq",
                data["mcqs"],
            )
        )
    if data["short"]:
        n, m = len(data["short"]), meta["short_marks"]
        out.append(
            (
                next(letters),
                "Short Questions",
                f"{n} × {m} = {n * m} Marks",
                "short",
                data["short"],
            )
        )
    if data["long"]:
        n, m = len(data["long"]), meta["long_marks"]
        out.append(
            (
                next(letters),
                "Long Questions",
                f"{n} × {m} = {n * m} Marks",
                "long",
                data["long"],
            )
        )
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
    """Markdown version for the screen (each option on its own line, in order)."""

    def br(t):
        return t.replace("\n", "  \n")

    L = [
        f"## 🎓 {meta['institute']}",
        f"**{meta['category']}**",
        "",
        f"**Subject:** {meta['topic']}  |  **Time:** {meta['time_min']} min  |  "
        f"**Total Marks:** {meta['total_marks']}",
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
            L.append(
                "  ".join(f"`Q{i}: {it['answer']}`" for i, it in enumerate(items, 1))
            )
        else:
            for i, it in enumerate(items, 1):
                L.append(f"- **Q{i}:** {it['answer'] or '-'}".replace("\n", " "))
        L.append("")
    return "\n".join(L)


# ============================================================
# PROFESSIONAL PDF (question paper / answer key)
# ============================================================

NAVY = colors.HexColor("#1E3A8A")
SLATE = colors.HexColor("#334155")
LIGHT = colors.HexColor("#E8EEF9")
GRID = colors.HexColor("#94A3B8")


def pdf_styles():
    base = getSampleStyleSheet()["Normal"]

    def mk(name, **kw):
        kw.setdefault("fontName", PDF_FONT)
        kw.setdefault("fontSize", 10.5)
        kw.setdefault("leading", 14)
        kw.setdefault("textColor", colors.HexColor("#0F172A"))
        return ParagraphStyle(name, parent=base, **kw)

    return {
        "inst": mk("inst", fontName=PDF_FONT_B, fontSize=17, leading=21,
                   textColor=NAVY, alignment=TA_CENTER),
        "exam": mk("exam", fontName=PDF_FONT_B, fontSize=11.5, leading=15,
                   textColor=SLATE, alignment=TA_CENTER, spaceAfter=6),
        "cell": mk("cell", fontSize=9.5, leading=12.5),
        "cellb": mk("cellb", fontName=PDF_FONT_B, fontSize=9.5, leading=12.5),
        "sec": mk("sec", fontName=PDF_FONT_B, fontSize=11, textColor=NAVY),
        "secr": mk("secr", fontName=PDF_FONT_B, fontSize=10, textColor=NAVY,
                   alignment=TA_RIGHT),
        "qn": mk("qn", fontName=PDF_FONT_B),
        "q": mk("q"),
        "opt": mk("opt", fontSize=10.2, leading=13.5),
        "small": mk("small", fontSize=8.5, leading=11.5, textColor=SLATE),
        "center": mk("center", fontSize=9, alignment=TA_CENTER, textColor=SLATE),
        "hadith": mk("hadith", fontName=PDF_FONT_I, fontSize=8.5, leading=12,
                     alignment=TA_CENTER, textColor=SLATE),
        "hadithref": mk("hadithref", fontSize=7.5, alignment=TA_CENTER,
                        textColor=colors.HexColor("#2563EB")),
    }


def _page_footer(canvas, doc):
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
        [
            rich("Subject / Topic:", S["cellb"]),
            rich(meta["topic"], S["cell"]),
            rich("Time Allowed:", S["cellb"]),
            rich(f"{meta['time_min']} Minutes", S["cell"]),
        ],
        [
            rich("Exam:", S["cellb"]),
            rich(meta["category"], S["cell"]),
            rich("Total Marks:", S["cellb"]),
            rich(str(meta["total_marks"]), S["cell"]),
        ],
    ]
    if with_student_fields:
        rows.append(
            [
                rich("Student Name:", S["cellb"]),
                rich("", S["cell"]),
                rich("Roll No:", S["cellb"]),
                rich("", S["cell"]),
            ]
        )
    cw = [36 * mm, width - 36 * mm - 32 * mm - 32 * mm, 32 * mm, 32 * mm]
    t = Table(rows, colWidths=cw)
    t.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.8, NAVY),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                ("BACKGROUND", (0, 0), (0, -1), LIGHT),
                ("BACKGROUND", (2, 0), (2, -1), LIGHT),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    fl.append(t)
    fl.append(Spacer(1, 8))
    return fl


def _section_bar(letter, title, marks, S, width):
    bar = Table(
        [
            [
                rich(f"SECTION {letter}  —  {title}", S["sec"]),
                rich(marks, S["secr"]),
            ]
        ],
        colWidths=[width * 0.65, width * 0.35],
    )
    bar.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                ("LINEBELOW", (0, 0), (-1, -1), 1.2, NAVY),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return bar


def _question_row(num_label, text, S, width):
    t = Table(
        [[rich(num_label, S["qn"]), rich(text, S["q"])]],
        colWidths=[13 * mm, width - 13 * mm],
    )
    t.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    return t


def _options_table(options, S, width):
    indent, lab = 13 * mm, 8 * mm
    use_two_cols = all(len(o) <= 38 and "\n" not in o for o in options)
    if use_two_cols:
        text_w = (width - indent - 2 * lab) / 2
        rows = []
        for i in range(0, len(options), 2):
            row = ["", rich(f"({LETTERS[i]})", S["opt"]), rich(options[i], S["opt"])]
            if i + 1 < len(options):
                row += [
                    rich(f"({LETTERS[i + 1]})", S["opt"]),
                    rich(options[i + 1], S["opt"]),
                ]
            else:
                row += ["", ""]
            rows.append(row)
        t = Table(rows, colWidths=[indent, lab, text_w, lab, text_w])
    else:
        text_w = width - indent - lab
        rows = [
            ["", rich(f"({LETTERS[i]})", S["opt"]), rich(o, S["opt"])]
            for i, o in enumerate(options)
        ]
        t = Table(rows, colWidths=[indent, lab, text_w])
    t.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ]
        )
    )
    return t


def _brand_footer(S):
    return [
        Spacer(1, 14),
        rich("— End of Paper —", S["center"]),
        Spacer(1, 10),
        rich(
            '"Whoever travels a path in search of knowledge, Allah will make'
            ' easy for him a path to Paradise."',
            S["hadith"],
        ),
        rich(
            "— Prophet Muhammad (PBUH) | Sahih Muslim, Book 35, Hadith 6518",
            S["hadithref"],
        ),
    ]


def build_paper_pdf(data, meta, mode="paper"):
    """mode='paper' -> student question paper, mode='key' -> answer key."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=20 * mm,
        title=f"{meta['topic']} - {'Answer Key' if mode == 'key' else 'Question Paper'}",
        author="WSA Educational Test Series",
    )
    width = A4[0] - 36 * mm
    S = pdf_styles()
    story = []

    if mode == "paper":
        story += _header_flowables(meta, S, width, meta["category"])
        inst = Table(
            [
                [
                    rich(
                        "Instructions:  "
                        + "  ".join(
                            f"{i}) {t}" for i, t in enumerate(INSTRUCTIONS, 1)
                        ),
                        S["small"],
                    )
                ]
            ],
            colWidths=[width],
        )
        inst.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 0.5, GRID),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story += [inst, Spacer(1, 10)]

        for letter, title, marks, kind, items in section_list(data, meta):
            story.append(_section_bar(letter, title, marks, S, width))
            story.append(Spacer(1, 7))
            for i, it in enumerate(items, 1):
                block = [_question_row(f"Q{i}.", it["q"], S, width)]
                if kind == "mcq":
                    block.append(_options_table(it["options"], S, width))
                    block.append(Spacer(1, 7))
                else:
                    block.append(Spacer(1, 10 if kind == "short" else 16))
                story.append(KeepTogether(block))
            story.append(Spacer(1, 6))
        story += _brand_footer(S)

    else:  # answer key
        story += _header_flowables(
            meta, S, width, "ANSWER KEY (Teacher Copy)", with_student_fields=False
        )
        for letter, title, marks, kind, items in section_list(data, meta):
            story.append(_section_bar(letter, title, marks, S, width))
            story.append(Spacer(1, 7))
            if kind == "mcq":
                per_row = 5
                cells = [
                    rich(f"Q{i}:  {it['answer']}", S["cellb"])
                    for i, it in enumerate(items, 1)
                ]
                while len(cells) % per_row:
                    cells.append("")
                rows = [
                    cells[i : i + per_row] for i in range(0, len(cells), per_row)
                ]
                t = Table(rows, colWidths=[width / per_row] * per_row)
                t.setStyle(
                    TableStyle(
                        [
                            ("BOX", (0, 0), (-1, -1), 0.5, GRID),
                            ("INNERGRID", (0, 0), (-1, -1), 0.3, GRID),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )
                story += [t, Spacer(1, 10)]
            else:
                for i, it in enumerate(items, 1):
                    story.append(
                        KeepTogether(
                            [
                                _question_row(f"Q{i}.", it["q"], S, width),
                                _options_answer(it["answer"], S, width),
                                Spacer(1, 6),
                            ]
                        )
                    )
        story.append(Spacer(1, 8))

    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    buffer.seek(0)
    return buffer


def _options_answer(answer, S, width):
    t = Table(
        [["", rich("Ans: " + (answer or "-"), S["small"])]],
        colWidths=[13 * mm, width - 13 * mm],
    )
    t.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return t


# ============================================================
# WORD (.docx) EXPORT - the most reliable option for Urdu text
# ============================================================


def build_paper_docx(data, meta):
    d = docx.Document()
    for s in d.sections:
        s.left_margin = s.right_margin = Cm(2)
        s.top_margin = s.bottom_margin = Cm(1.8)
    normal = d.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    def add_text(par, text, bold=False, size=None, color=None):
        parts = clean_text(text).split("\n")
        for idx, part in enumerate(parts):
            run = par.add_run(part)
            run.bold = bold
            if size:
                run.font.size = Pt(size)
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            if idx < len(parts) - 1:
                run.add_break()

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(p, meta["institute"].upper(), bold=True, size=17, color="1E3A8A")
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(p, meta["category"], bold=True, size=12, color="334155")

    t = d.add_table(rows=3, cols=4)
    t.style = "Table Grid"
    info = [
        ("Subject / Topic:", meta["topic"], "Time Allowed:", f"{meta['time_min']} Minutes"),
        ("Exam:", meta["category"], "Total Marks:", str(meta["total_marks"])),
        ("Student Name:", "", "Roll No:", ""),
    ]
    for r, row in enumerate(info):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            cell.text = ""
            add_text(cell.paragraphs[0], val, bold=(c % 2 == 0), size=10)

    p = d.add_paragraph()
    add_text(
        p,
        "Instructions: "
        + "  ".join(f"{i}) {x}" for i, x in enumerate(INSTRUCTIONS, 1)),
        size=9,
    )

    for letter, title, marks, kind, items in section_list(data, meta):
        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        add_text(
            p, f"SECTION {letter} — {title}      ({marks})",
            bold=True, size=12, color="1E3A8A",
        )
        for i, it in enumerate(items, 1):
            p = d.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(1)
            add_text(p, f"Q{i}.  ", bold=True)
            add_text(p, it["q"])
            if kind == "mcq":
                for j, opt in enumerate(it["options"]):
                    op = d.add_paragraph()
                    op.paragraph_format.left_indent = Cm(1.2)
                    op.paragraph_format.space_after = Pt(0)
                    add_text(op, f"({LETTERS[j]})  {opt}")
            elif kind == "long":
                d.add_paragraph()

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    add_text(p, "— End of Paper —", size=9, color="64748B")
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(
        p,
        "Designed by Waheed Ali Hamouzai • WSA Educational Community",
        size=8,
        color="64748B",
    )

    buf = io.BytesIO()
    d.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# SIMPLE TEXT -> PDF (evaluation / diagnostic reports)
# ============================================================


def create_pdf_from_text(title, content):
    """Converts markdown/text content into a downloadable PDF with Hadith and author branding."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=45,
    )
    S = pdf_styles()
    body_style = ParagraphStyle(
        "BodyStyle", parent=S["q"], fontSize=10, leading=14, spaceAfter=6
    )
    head_style = ParagraphStyle(
        "HeadStyle", parent=S["q"], fontName=PDF_FONT_B, fontSize=11.5,
        leading=15, textColor=NAVY, spaceBefore=8, spaceAfter=4,
    )
    title_style = ParagraphStyle(
        "TitleStyle", parent=S["inst"], fontSize=18, spaceAfter=12
    )

    story = [
        rich("Generated via WSA Educational Test Series & Paper Builder (AI-Powered)", S["center"]),
        Spacer(1, 6),
        rich(title, title_style),
        Spacer(1, 6),
    ]

    for line in content.split("\n"):
        line = EMOJI_RE.sub("", line).rstrip()
        if not line.strip():
            story.append(Spacer(1, 5))
            continue
        stripped = line.strip()
        if stripped.startswith("#") or (
            stripped.startswith("**") and stripped.endswith("**") and len(stripped) < 90
        ):
            story.append(rich(stripped.lstrip("#").strip(), head_style))
        elif set(stripped) <= set("-=_ "):
            continue
        else:
            story.append(rich(stripped, body_style))

    story += [Spacer(1, 12)] + _brand_footer(S)[2:]
    story.append(
        rich("Designed by Waheed Ali Hamouzai • WSA Educational Community", S["center"])
    )
    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    buffer.seek(0)
    return buffer


# ============================================================
# FILE EXTRACTION FUNCTION
# ============================================================


def process_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".txt"):
        try:
            return uploaded_file.getvalue().decode("utf-8").strip()
        except Exception:
            return ""

    elif filename.endswith(".docx"):
        try:
            document = docx.Document(uploaded_file)
            full_text = [p.text for p in document.paragraphs if p.text]
            return "\n".join(full_text).strip()
        except Exception:
            return ""

    elif filename.endswith(".pdf"):
        text = ""
        try:
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
            text = text.strip()
        except Exception:
            text = ""

        return text[:MAX_REFERENCE_CHARS]

    return ""


# ============================================================
# UNIVERSAL LLM CALL (Groq / Gemini) WITH FALLBACK
# ============================================================


def call_llm(
    llm_cfg,
    prompt_text,
    system_instruction="You are an expert educational examiner.",
    temperature=0.3,
    json_mode=False,
    image_b64=None,
):
    """
    image_b64: if given, the image is read with a vision model.
    llm_cfg = {
        "keys": {"Groq": "...", "Gemini": "..."},
        "provider": "Auto (Fallback)" | "Groq" | "Gemini",
        "custom_model": "" (optional),
    }
    """
    keys = llm_cfg["keys"]
    provider_choice = llm_cfg["provider"]
    custom_model = (llm_cfg.get("custom_model") or "").strip()

    if provider_choice == AUTO_MODE:
        provider_order = [p for p in AUTO_ORDER if keys.get(p)]
    else:
        provider_order = [provider_choice] if keys.get(provider_choice) else []

    if image_b64:
        # only providers that have a vision model
        provider_order = [p for p in provider_order if PROVIDERS[p]["vision_models"]]

    if not provider_order:
        raise Exception(
            "No AI key is configured on the server"
            + (" for reading photos (a Groq or Gemini key is needed)." if image_b64 else ".")
        )

    errors = []

    for provider in provider_order:
        cfg = PROVIDERS[provider]
        client = OpenAI(api_key=keys[provider], base_url=cfg["base_url"], timeout=90)

        if image_b64:
            models_to_try = cfg["vision_models"]
        elif custom_model and provider_choice == provider:
            models_to_try = [custom_model]
        else:
            models_to_try = cfg["models"]

        for model_name in models_to_try:
            if image_b64:
                # for vision models the instruction goes inside the user message
                messages = [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": system_instruction + "\n\n" + prompt_text},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"},
                            },
                        ],
                    }
                ]
            else:
                messages = [
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": prompt_text},
                ]
            use_json = json_mode and not image_b64
            attempts = [{"response_format": {"type": "json_object"}}, {}] if use_json else [{}]
            for extra in attempts:
                try:
                    response = client.chat.completions.create(
                        model=model_name,
                        messages=messages,
                        temperature=temperature,
                        **extra,
                    )
                    text = response.choices[0].message.content or ""
                    # some reasoning models also return <think>...</think> blocks
                    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
                    if text:
                        return text
                    errors.append(f"{provider}/{model_name}: empty response")
                except Exception as e:
                    err = str(e)
                    if keys.get(provider):
                        err = err.replace(keys[provider], "***")
                    errors.append(f"{provider}/{model_name}: {err[:160]}")
                    continue

    raise Exception("All AI providers failed:\n- " + "\n- ".join(errors))


# ============================================================
# IMAGE (PHOTO / SCAN) -> TEXT
# ============================================================


def extract_text_from_image(llm_cfg, uploaded_file):
    """Reads text from a photo (paper, answer sheet, result card). Urdu + English."""
    image_b64 = prepare_image_b64(uploaded_file)
    text = call_llm(
        llm_cfg,
        prompt_text=(
            "Transcribe ALL text visible in this image exactly as written (English"
            " and Urdu, printed or handwritten). Keep question numbers, options,"
            " marks, ticks/circled answers and the original line order. Do not"
            " explain, summarise or add anything. If nothing readable is present,"
            " reply with exactly: UNREADABLE"
        ),
        system_instruction="You are a precise OCR engine for exam papers and answer sheets.",
        temperature=0.0,
        image_b64=image_b64,
    )
    if "UNREADABLE" in text.upper() and len(text) < 40:
        raise Exception(
            USER_MESSAGE_PREFIX
            + "Could not read any text from the photo. Please upload a clear,"
            " well-lit and straight photo."
        )
    return text


def read_upload(uploaded_file, llm_cfg):
    """Reads PDF/DOCX/TXT directly. Photos are read with the AI vision model."""
    if is_image(uploaded_file):
        return extract_text_from_image(llm_cfg, uploaded_file)
    return process_uploaded_file(uploaded_file)


def preview_if_image(uploaded_file):
    if is_image(uploaded_file):
        st.image(uploaded_file, width=260)


# ============================================================
# AI GENERATION & EVALUATION LOGIC
# ============================================================


def generate_test_paper(
    llm_cfg,
    topic,
    uploaded_pdf,
    test_type,
    language,
    mcq_count,
    short_count,
    long_count,
    diff_level,
    level_note="",
):
    pdf_text = read_upload(uploaded_pdf, llm_cfg)[:MAX_REFERENCE_CHARS]

    if language == "Urdu":
        language_instruction = (
            "Write EVERYTHING (questions, options, answers) strictly in URDU language."
        )
    elif language == "Bilingual (English + Urdu)":
        language_instruction = (
            "BILINGUAL: every question is written in English, then a new line"
            " (\\n), then its Urdu translation. Every option is written as"
            " 'English text / اردو ترجمہ' on ONE line. Answers may be in English."
        )
    else:
        language_instruction = "Write everything in ENGLISH language."

    reference_block = f"REFERENCE TEXT (base the questions on this):\n{pdf_text}" if pdf_text else ""

    prompt = f"""
Create a professional examination paper as a JSON object.

EXAM CATEGORY / STYLE: {test_type}
TOPIC / SUBJECT: {topic}
DIFFICULTY LEVEL: {diff_level}
LEARNER LEVEL: {level_note or "General audience"}
LANGUAGE: {language}
{language_instruction}

EXACT QUESTION COUNTS:
- mcqs: {mcq_count}
- short: {short_count}
- long: {long_count}
(If a count is 0, return an empty list for it.)

QUALITY RULES:
1. Every MCQ has EXACTLY 4 options in the "options" list, WITHOUT letter prefixes
   (write "Paris", NOT "A) Paris"). The "answer" is only the letter: "A", "B", "C" or "D".
2. Only ONE option is correct; the other three must be plausible distractors.
   Avoid overusing "All of the above" / "None of the above".
3. Spread the correct answers evenly across A, B, C and D (do not favour one letter).
4. No repeated or near-duplicate questions. Cover different sub-topics.
5. Order questions from easier to harder within each section.
6. Short questions need a 2-3 line answer; long questions are descriptive and may
   have parts (a), (b), (c) on separate lines. Give a model answer / key points for each.
7. Do not use markdown (no ** or #) inside any text.

Return ONLY a valid JSON object in exactly this structure, with no extra text:
{JSON_SCHEMA_TEXT}

{reference_block}
"""

    system = (
        "You are a senior paper-setter for educational boards and competitive"
        " testing services (NTS, PPSC, FPSC, school boards). You output ONLY"
        " valid JSON, never prose."
    )

    last_error = None
    for attempt in range(2):
        raw = call_llm(
            llm_cfg,
            prompt_text=prompt if attempt == 0 else prompt + "\nIMPORTANT: your previous reply was not valid JSON. Reply with JSON ONLY.",
            system_instruction=system,
            temperature=0.3,
            json_mode=True,
        )
        try:
            data = parse_paper_json(raw, mcq_count, short_count, long_count)
            if data["mcqs"] or data["short"] or data["long"]:
                return data
            last_error = "the AI returned no questions"
        except Exception as e:
            last_error = str(e)

    raise Exception(
        USER_MESSAGE_PREFIX
        + "The AI did not return the paper in the expected format"
        f" ({last_error}). Please click Generate again."
    )


def evaluate_student_answers(llm_cfg, paper_text, answers_text):
    prompt = f"""
Evaluate the student's answer sheet against the provided question paper and answer key.

QUESTION PAPER / ANSWER KEY:
{paper_text}

STUDENT ANSWERS:
{answers_text}

Provide a structured evaluation report:
- Total Marks & Obtained Marks
- Percentage & Grade
- Question-wise Analysis
- Strengths & Weaknesses
"""

    return call_llm(
        llm_cfg,
        prompt_text=prompt,
        system_instruction=(
            "You are an experienced examiner evaluating student answers"
            " accurately and providing constructive feedback."
        ),
        temperature=0.2,
    )


def analyze_random_test(llm_cfg, test_content, additional_context=""):
    context = (
        additional_context
        if additional_context
        else "General test analysis and performance diagnostic."
    )
    prompt = f"""
Analyze the following random test paper, solved sheet, or test result provided by the user.

USER CONTEXT / GOAL:
{context}

TEST DATA / CONTENT:
{test_content}

Provide a comprehensive Diagnostic & Improvement Report structured as follows:
1. 📌 **Executive Performance Overview**: High-level estimation of score, accuracy, or completion quality.
2. 🎯 **Key Weaknesses & Knowledge Gaps**: Identify exact topics, question types, or concepts where performance is lacking.
3. 🔍 **Priority Focus Areas**: Highlight top 3 critical subjects/topics the student MUST prioritize immediately.
4. 🚀 **Actionable Improvement Strategy**: Step-by-step study recommendations, practice methods, and revision plan.
5. 💡 **Recommended Resources & Next Steps**: Suggested topics to solve next or key formulas/concepts to memorize.
"""

    return call_llm(
        llm_cfg,
        prompt_text=prompt,
        system_instruction=(
            "You are a senior academic mentor and diagnostic expert specializing"
            " in test analysis and student performance optimization."
        ),
        temperature=0.3,
    )


def safe_download_data(builder, *args, **kwargs):
    """If building a PDF/Word file fails, show a warning instead of crashing the app."""
    try:
        return builder(*args, **kwargs)
    except Exception as e:
        st.warning("Could not create this file. Please try again.")
        print("File build error:", e)
        return None


# ============================================================
# EXAM CATEGORIES: group -> class / level -> subjects, difficulty, time, AI note
# ============================================================

EXAM_GROUPS = {
    "School (Class 1–10)": [f"Class {i}" for i in range(1, 11)],
    "College (Class 11–12)": ["Class 11 (FA / FSc / ICS)", "Class 12 (FA / FSc / ICS)"],
    "University": ["Bachelor (BS / BA / BSc)", "Master (MS / MA / MSc)"],
    "Competitive Tests": [
        "General Test (NTS / CTSP / SBK)",
        "PPSC / FPSC",
        "BPSC",
        "CSS",
        "CS / IT Screening Test",
    ],
    "Custom": ["Custom Mock Test"],
}

OTHER_SUBJECT = "Other (type below)"
DIFF_LABELS = {1: "Very Easy", 2: "Easy", 3: "Medium", 4: "Hard", 5: "Very Hard"}
LANGUAGES = ["English", "Urdu", "Bilingual (English + Urdu)"]

COMPETITIVE_DEFAULTS = {
    "General Test (NTS / CTSP / SBK)": (3, 60),
    "PPSC / FPSC": (4, 90),
    "BPSC": (4, 90),
    "CSS": (5, 120),
    "CS / IT Screening Test": (4, 60),
}


def subjects_for(group, level):
    """Common subjects for the selected class / level."""
    if group == "School (Class 1–10)":
        n = int(re.search(r"\d+", level).group())
        if n <= 2:
            base = ["English", "Urdu", "Mathematics", "General Knowledge", "Islamiyat"]
        elif n <= 5:
            base = ["English", "Urdu", "Mathematics", "General Science", "Social Studies",
                    "Islamiyat", "Computer"]
        elif n <= 8:
            base = ["English", "Urdu", "Mathematics", "General Science", "Social Studies",
                    "Islamiyat", "Computer Science", "Geography", "History"]
        else:
            base = ["English", "Urdu", "Mathematics", "Physics", "Chemistry", "Biology",
                    "Computer Science", "Pakistan Studies", "Islamiyat", "General Science"]
    elif group == "College (Class 11–12)":
        base = ["Physics", "Chemistry", "Biology", "Mathematics", "Computer Science",
                "English", "Urdu", "Pakistan Studies", "Islamiyat", "Statistics",
                "Economics", "Accounting"]
    elif group == "University":
        base = ["Computer Science", "Information Technology", "Software Engineering",
                "Mathematics", "Physics", "Chemistry", "Biology", "English",
                "Business Administration", "Economics", "Education"]
    elif group == "Competitive Tests":
        base = ["General Knowledge", "English", "Urdu", "Pakistan Studies", "Islamic Studies",
                "Current Affairs", "Everyday Science", "Quantitative / Mathematics",
                "Analytical Reasoning / IQ", "Computer Science / IT"]
    else:
        base = ["General"]
    return base + [OTHER_SUBJECT]


def level_profile(group, level):
    """Default difficulty (1-5), time (minutes) and a note for the AI, based on the class / level."""
    curriculum = "Follow the Pakistani national curriculum / textbook board syllabus."

    if group == "School (Class 1–10)":
        n = int(re.search(r"\d+", level).group())
        if n <= 2:
            diff, time_min = 1, 30
            band = "Use very simple words and very short sentences. Keep every question basic and use everyday examples."
        elif n <= 5:
            diff, time_min = (1 if n == 3 else 2), 45
            band = "Use simple words and short sentences. Questions must suit primary-level learners."
        elif n <= 8:
            diff, time_min = (2 if n < 8 else 3), 60
            band = "Use clear, simple language suitable for middle-level students."
        else:
            diff, time_min = 3, 90
            band = "Match the Secondary School Certificate (SSC / Matric) exam standard."
        note = (
            f"Class {n} students (about {n + 5}-{n + 6} years old). {band} {curriculum}"
        )
    elif group == "College (Class 11–12)":
        is_12 = level.startswith("Class 12")
        diff, time_min = (4, 120) if is_12 else (3, 90)
        note = f"{level} students. Match the Intermediate (HSSC) exam standard. {curriculum}"
    elif group == "University":
        diff, time_min = 4, 120
        note = f"University students ({level}). Use university-level terminology and conceptual depth."
    elif group == "Competitive Tests":
        diff, time_min = COMPETITIVE_DEFAULTS.get(level, (3, 60))
        note = (
            f"Candidates preparing for {level} in Pakistan. Match the style and standard of"
            " such recruitment / screening tests: factual, concept-based and application questions."
        )
    else:
        diff, time_min, note = 3, 60, ""

    return {"difficulty": diff, "time_min": time_min, "note": note}


def apply_level_defaults():
    """Sets difficulty, time and subject automatically when the class / level changes."""
    ss = st.session_state
    prof = level_profile(ss["exam_group"], ss["exam_level"])
    ss["difficulty"] = prof["difficulty"]
    ss["time_min"] = prof["time_min"]
    ss["subject"] = subjects_for(ss["exam_group"], ss["exam_level"])[0]


def on_group_change():
    st.session_state["exam_level"] = EXAM_GROUPS[st.session_state["exam_group"]][0]
    apply_level_defaults()


def on_level_change():
    apply_level_defaults()


def use_generated_paper():
    st.session_state["q_paper_text"] = st.session_state.get("generated_paper", "")


# First-run defaults
_first_group = list(EXAM_GROUPS)[0]
_defaults = {
    "exam_group": _first_group,
    "exam_level": EXAM_GROUPS[_first_group][0],
    "chapter": "",
    "subject_other": "",
    "language": "English",
    "mcq_count": 10,
    "short_count": 5,
    "long_count": 2,
    "institute": "WSA Educational Community",
    "mcq_marks": 1,
    "short_marks": 2,
    "long_marks": 5,
    "q_paper_text": "",
}
for _k, _v in _defaults.items():
    st.session_state.setdefault(_k, _v)
if "difficulty" not in st.session_state:
    apply_level_defaults()


# ============================================================
# AI KEYS (server side only) + HEADER
# ============================================================

llm_cfg = {
    "keys": {name: get_secret(cfg["env"]) for name, cfg in PROVIDERS.items()},
    "provider": AUTO_MODE,
    "custom_model": "",
}
has_any_key = any(llm_cfg["keys"].values())

NO_KEY_MSG = "⚠️ The AI service is not configured yet. Please contact the administrator."

st.markdown(
    """<div class="hero">
<div class="hero-label">🎓 WSA Educational Community</div>
<div class="hero-title">Test Series & Paper Builder</div>
<div class="hero-subtitle">Create exam papers for any class or competitive test, check answer sheets, and find weak areas in minutes.</div>
<span class="chip">English · Urdu · Bilingual</span>
<span class="chip">PDF & Word export</span>
<span class="chip">Answer key included</span>
<span class="chip">Photo upload</span>
</div>""",
    unsafe_allow_html=True,
)

if not has_any_key:
    st.error(NO_KEY_MSG)
    st.caption("Administrator: add GROQ_API_KEY (and optionally GEMINI_API_KEY) in the app secrets.")

tab1, tab2, tab3 = st.tabs(
    [
        "📝 Create Paper",
        "📊 Check Answers",
        "🎯 Test Diagnostic",
    ]
)


# ============================================================
# TAB 1 — CREATE PAPER
# ============================================================

with tab1:
    with st.container(border=True, key="card_1"):
        st.markdown('<div class="section-title">1. Class and subject</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="section-hint">Pick the class or test. Difficulty and time are set for you.</div>',
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            st.selectbox(
                "Category",
                list(EXAM_GROUPS),
                key="exam_group",
                on_change=on_group_change,
            )
        with c2:
            st.selectbox(
                "Class / Test",
                EXAM_GROUPS[st.session_state["exam_group"]],
                key="exam_level",
                on_change=on_level_change,
            )
        with c3:
            st.selectbox(
                "Subject",
                subjects_for(st.session_state["exam_group"], st.session_state["exam_level"]),
                key="subject",
            )

        if st.session_state["subject"] == OTHER_SUBJECT:
            st.text_input(
                "Subject name",
                key="subject_other",
                placeholder="e.g., Home Economics",
            )

        t1, t2, t3 = st.columns([2, 1, 1])
        with t1:
            st.text_input(
                "Chapter / Topic (optional)",
                key="chapter",
                placeholder="e.g., Chapter 3 — Force and Motion",
            )
        with t2:
            st.selectbox("Language", LANGUAGES, key="language")
        with t3:
            st.slider("Difficulty", 1, 5, key="difficulty")
        st.caption(
            f"Difficulty: **{DIFF_LABELS[st.session_state['difficulty']]}** "
            f"(set automatically for {st.session_state['exam_level']}, you can change it)."
        )

    with st.container(border=True, key="card_2"):
        st.markdown('<div class="section-title">2. Questions</div>', unsafe_allow_html=True)
        q1, q2, q3 = st.columns(3)
        with q1:
            st.number_input("Number of MCQs", 0, 50, key="mcq_count")
        with q2:
            st.number_input("Number of Short Questions", 0, 20, key="short_count")
        with q3:
            st.number_input("Number of Long Questions", 0, 10, key="long_count")

        uploaded_pdf = st.file_uploader(
            f"Reference material (optional) — {UPLOAD_HINT}",
            type=UPLOAD_TYPES,
            help="Upload a chapter, notes or a textbook page photo and the questions will be based on it.",
        )
        preview_if_image(uploaded_pdf)

        with st.expander("🏫 Paper header, marks and time (optional)"):
            h1, h2 = st.columns(2)
            with h1:
                st.text_input("Institute / Academy name", key="institute")
                st.number_input("Time allowed (minutes)", 5, 300, key="time_min", step=5)
            with h2:
                st.number_input("Marks per MCQ", 1, 10, key="mcq_marks")
                st.number_input("Marks per Short Question", 1, 20, key="short_marks")
                st.number_input("Marks per Long Question", 1, 50, key="long_marks")

    ss = st.session_state
    subject_name = (
        ss["subject_other"].strip() if ss["subject"] == OTHER_SUBJECT else ss["subject"]
    )
    topic_text = subject_name + (f" — {ss['chapter'].strip()}" if ss["chapter"].strip() else "")

    if st.button("🚀 Generate Test Paper", type="primary", use_container_width=True):
        if not has_any_key:
            st.error(NO_KEY_MSG)
        elif not subject_name:
            st.error("⚠️ Please enter the subject name.")
        elif ss["mcq_count"] + ss["short_count"] + ss["long_count"] == 0:
            st.error("⚠️ Please ask for at least one question.")
        elif not allow_ai_call(units=1 + int(is_image(uploaded_pdf))):
            pass  # allow_ai_call already showed the reason
        else:
            profile = level_profile(ss["exam_group"], ss["exam_level"])
            with st.spinner("Creating your test paper... this takes a few seconds."):
                try:
                    paper = generate_test_paper(
                        llm_cfg,
                        topic_text,
                        uploaded_pdf,
                        ss["exam_level"],
                        ss["language"],
                        ss["mcq_count"],
                        ss["short_count"],
                        ss["long_count"],
                        DIFF_LABELS[ss["difficulty"]],
                        profile["note"],
                    )
                    meta = {
                        "institute": ss["institute"].strip() or "WSA Educational Community",
                        "category": ss["exam_level"],
                        "topic": topic_text,
                        "language": ss["language"],
                        "time_min": int(ss["time_min"]),
                        "mcq_marks": int(ss["mcq_marks"]),
                        "short_marks": int(ss["short_marks"]),
                        "long_marks": int(ss["long_marks"]),
                    }
                    meta["total_marks"] = compute_total_marks(paper, meta)

                    ss["paper_data"] = paper
                    ss["paper_meta"] = meta
                    ss["generated_paper"] = paper_to_text(paper, meta, include_key=True)

                    got = (len(paper["mcqs"]), len(paper["short"]), len(paper["long"]))
                    want = (ss["mcq_count"], ss["short_count"], ss["long_count"])
                    if got != want:
                        st.warning(
                            f"The AI created {got[0]} MCQs, {got[1]} short and {got[2]} long"
                            f" questions, but you asked for {want[0]}, {want[1]} and {want[2]}."
                            " You can click Generate again."
                        )
                except Exception as e:
                    show_error(e)

    if "paper_data" in st.session_state:
        paper = st.session_state["paper_data"]
        meta = st.session_state["paper_meta"]

        st.markdown("&nbsp;", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total marks", meta["total_marks"])
        m2.metric("Time", f"{meta['time_min']} min")
        m3.metric("MCQs", len(paper["mcqs"]))
        m4.metric("Short / Long", f"{len(paper['short'])} / {len(paper['long'])}")

        with st.container(border=True, key="card_3"):
            st.markdown(paper_to_markdown(paper, meta))

        with st.expander("🔑 Answer key (teacher copy)"):
            st.markdown(answer_key_markdown(paper, meta))

        st.markdown('<div class="section-title">Download</div>', unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        d3, d4 = st.columns(2)

        paper_pdf = safe_download_data(build_paper_pdf, paper, meta, "paper")
        key_pdf = safe_download_data(build_paper_pdf, paper, meta, "key")
        word_file = safe_download_data(build_paper_docx, paper, meta)

        with d1:
            if paper_pdf:
                st.download_button(
                    "⬇️ Question Paper (PDF)",
                    data=paper_pdf,
                    file_name="WSA_Question_Paper.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        with d2:
            if key_pdf:
                st.download_button(
                    "🔑 Answer Key (PDF)",
                    data=key_pdf,
                    file_name="WSA_Answer_Key.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        with d3:
            if word_file:
                st.download_button(
                    "📝 Question Paper (Word)",
                    data=word_file,
                    file_name="WSA_Question_Paper.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
        with d4:
            st.download_button(
                "⬇️ Paper + Key (Text)",
                data=st.session_state["generated_paper"],
                file_name="WSA_Test_Paper.txt",
                mime="text/plain",
                use_container_width=True,
            )

        if meta["language"] != "English":
            st.info("Tip: For Urdu or Bilingual papers, the Word (.docx) file is the most reliable option.")
        st.markdown(
            '<div class="ai-note">AI-generated content can contain mistakes. Please review the paper before using it.</div>',
            unsafe_allow_html=True,
        )


# ============================================================
# TAB 2 — CHECK ANSWERS
# ============================================================

with tab2:
    with st.container(border=True, key="card_4"):
        st.markdown('<div class="section-title">Check student answers</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="section-hint">Add the question paper (with answer key) and the student\'s answer sheet. You can upload a file or a photo, or paste text.</div>',
            unsafe_allow_html=True,
        )

        col_paper, col_answer = st.columns(2)

        with col_paper:
            st.markdown("**1️⃣ Question paper / answer key**")
            paper_file = st.file_uploader(
                f"Upload paper ({UPLOAD_HINT})",
                type=UPLOAD_TYPES,
                key="p_up",
            )
            preview_if_image(paper_file)
            if st.session_state.get("generated_paper"):
                st.button(
                    "📄 Use the paper I just created",
                    on_click=use_generated_paper,
                    use_container_width=True,
                )
            question_paper_text = st.text_area(
                "Or paste text",
                height=200,
                key="q_paper_text",
            )

        with col_answer:
            st.markdown("**2️⃣ Student answer sheet**")
            student_file = st.file_uploader(
                f"Upload answers ({UPLOAD_HINT})",
                type=UPLOAD_TYPES,
                key="s_up",
            )
            preview_if_image(student_file)
            student_answers_text = st.text_area(
                "Or paste text", height=200, key="s_answers_text"
            )

    if st.button("📊 Check Answers", type="primary", use_container_width=True):
        if not has_any_key:
            st.error(NO_KEY_MSG)
        else:
            p_text = process_uploaded_file(paper_file)  # photos are read later with AI
            a_text = process_uploaded_file(student_file)
            paper_pasted = question_paper_text.strip()
            answers_pasted = student_answers_text.strip()
            n_images = int(is_image(paper_file)) + int(is_image(student_file))

            if not (p_text or paper_pasted or is_image(paper_file)):
                st.error("⚠️ The question paper is missing. Please upload it or paste the text.")
            elif not (a_text or answers_pasted or is_image(student_file)):
                st.error("⚠️ The student's answers are missing. Please upload them or paste the text.")
            elif not allow_ai_call(units=1 + n_images):
                pass  # allow_ai_call already showed the reason
            else:
                with st.spinner("Checking the answers... please wait."):
                    try:
                        if is_image(paper_file):
                            p_text = read_upload(paper_file, llm_cfg)
                        if is_image(student_file):
                            a_text = read_upload(student_file, llm_cfg)
                        eval_res = evaluate_student_answers(
                            llm_cfg, p_text or paper_pasted, a_text or answers_pasted
                        )
                        st.session_state["evaluation"] = eval_res
                    except Exception as e:
                        show_error(e)

    if "evaluation" in st.session_state:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        with st.container(border=True, key="card_5"):
            st.markdown('<div class="section-title">📋 Evaluation result</div>', unsafe_allow_html=True)
            st.markdown(st.session_state["evaluation"])

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                "⬇️ Download report (.txt)",
                data=st.session_state["evaluation"],
                file_name="WSA_Student_Evaluation.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with col_d2:
            eval_pdf = safe_download_data(
                create_pdf_from_text,
                "Student Evaluation Report",
                st.session_state["evaluation"],
            )
            if eval_pdf:
                st.download_button(
                    "⬇️ Download report (.pdf)",
                    data=eval_pdf,
                    file_name="WSA_Student_Evaluation.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        st.markdown(
            '<div class="ai-note">AI marking can make mistakes. Please review important results yourself.</div>',
            unsafe_allow_html=True,
        )


# ============================================================
# TAB 3 — TEST DIAGNOSTIC
# ============================================================

with tab3:
    with st.container(border=True, key="card_6"):
        st.markdown('<div class="section-title">Find strengths and weak areas</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="section-hint">Upload or paste any solved test, question paper or result card to get a clear improvement plan.</div>',
            unsafe_allow_html=True,
        )

        col_diag1, col_diag2 = st.columns(2)

        with col_diag1:
            st.markdown("**📄 Test file or raw data**")
            random_test_file = st.file_uploader(
                f"Upload test or result ({UPLOAD_HINT})",
                type=UPLOAD_TYPES,
                key="random_test_up",
            )
            preview_if_image(random_test_file)
            random_test_text = st.text_area(
                "Or paste test content / scores",
                height=200,
                key="random_test_text",
                placeholder=(
                    "e.g., Question 1: Incorrect answer selected...\nQuestion 2:"
                    " Correct...\nOr paste the raw quiz content here."
                ),
            )

        with col_diag2:
            st.markdown("**⚙️ Goal (optional)**")
            user_context = st.text_area(
                "Target exam / student goal",
                height=270,
                key="user_context",
                placeholder=(
                    "e.g., Preparing for the BPSC Computer Science test."
                    " Target score is 80%+."
                ),
            )

    if st.button("🔍 Analyze Test", type="primary", use_container_width=True):
        if not has_any_key:
            st.error(NO_KEY_MSG)
        else:
            extracted_test = process_uploaded_file(random_test_file)  # photos are read later with AI
            pasted_test = random_test_text.strip()

            if not (extracted_test or pasted_test or is_image(random_test_file)):
                st.error("⚠️ Please upload a test file (PDF, DOCX, TXT or photo) or paste the text.")
            elif not allow_ai_call(units=1 + int(is_image(random_test_file))):
                pass  # allow_ai_call already showed the reason
            else:
                with st.spinner("Analyzing the test and building your improvement plan..."):
                    try:
                        if is_image(random_test_file):
                            extracted_test = read_upload(random_test_file, llm_cfg)
                        diag_res = analyze_random_test(
                            llm_cfg, extracted_test or pasted_test, user_context
                        )
                        st.session_state["diagnostic_analysis"] = diag_res
                    except Exception as e:
                        show_error(e)

    if "diagnostic_analysis" in st.session_state:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        with st.container(border=True, key="card_7"):
            st.markdown('<div class="section-title">📈 Diagnostic report and improvement plan</div>', unsafe_allow_html=True)
            st.markdown(st.session_state["diagnostic_analysis"])

        col_rd1, col_rd2 = st.columns(2)
        with col_rd1:
            st.download_button(
                "⬇️ Download report (.txt)",
                data=st.session_state["diagnostic_analysis"],
                file_name="WSA_Test_Diagnostic.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with col_rd2:
            diag_pdf = safe_download_data(
                create_pdf_from_text,
                "Test Diagnostic & Improvement Plan",
                st.session_state["diagnostic_analysis"],
            )
            if diag_pdf:
                st.download_button(
                    "⬇️ Download report (.pdf)",
                    data=diag_pdf,
                    file_name="WSA_Test_Diagnostic.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        st.markdown(
            '<div class="ai-note">AI analysis is a guide, not a final judgement. Please use your own judgement too.</div>',
            unsafe_allow_html=True,
        )


# ============================================================
# MAIN FOOTER
# ============================================================

st.markdown(
    """<div class="branding-card">
    <div class="hadith-quote">"Whoever travels a path in search of knowledge, Allah will make easy for him a path to Paradise."</div>
    <div class="hadith-ref">— Prophet Muhammad (PBUH) | Sahih Muslim, Book 35, Hadith 6518</div>
    <hr class="footer-divider">
    <div class="branding-name">Designed by Waheed Ali Hamouzai</div>
    <div class="branding-tag">WSA Educational Community • AI Powered Learning</div>
</div>""",
    unsafe_allow_html=True,
)

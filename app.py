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

# Urdu / Arabic script support in PDF (optional - app bina iske bhi chalti hai)
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
)

# ============================================================
# LLM PROVIDERS CONFIG (all use OpenAI-compatible endpoints)
# ------------------------------------------------------------
# Model names / free limits change often. Agar koi model "not found"
# de, toh secrets mai GROQ_MODEL = "naya-model-id" likh do.
#   Groq:        https://console.groq.com/keys
#   Gemini:      https://aistudio.google.com/apikey
# ============================================================

PROVIDERS = {
    "Groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "env": "GROQ_API_KEY",
        "models": ["openai/gpt-oss-120b", "openai/gpt-oss-20b"],
        # Photo se text nikalne ke liye (Groq ke vision models preview mai hain)
        "vision_models": ["qwen/qwen3.8-27b"],
    },
    "Gemini": {
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "env": "GEMINI_API_KEY",
        # "-latest" alias hamesha Google ke current Flash model par rehta hai,
        # is liye purane model band hone par app nahi tootti.
        "models": ["gemini-flash-latest", "gemini-3.1-flash-lite"],
        "vision_models": ["gemini-flash-latest", "gemini-3.1-flash-lite"],
    },
}

AUTO_MODE = "Auto (Fallback)"
AUTO_ORDER = ["Groq", "Gemini"]

# Free tiers ki token limits chhoti hoti hain, isliye reference text cap
MAX_REFERENCE_CHARS = 9000

LETTERS = "ABCD"

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")
UPLOAD_TYPES = ["pdf", "docx", "txt", "jpg", "jpeg", "png", "webp"]
UPLOAD_HINT = "PDF, DOCX, TXT, JPG, PNG"


def is_image(uploaded_file):
    return uploaded_file is not None and uploaded_file.name.lower().endswith(IMAGE_EXTS)


def prepare_image_b64(uploaded_file, max_side=1800):
    """Photo ko chhota (compress) karke base64 mai badalta hai (API limit ke andar rehne ke liye)."""
    from PIL import Image, ImageOps

    img = Image.open(io.BytesIO(uploaded_file.getvalue()))
    img = ImageOps.exif_transpose(img)  # mobile photo ka ulta / lait hona theek karta hai
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


# ============================================================
# SECRETS + SHARED-QUOTA PROTECTION
# ------------------------------------------------------------
# API key user ko kabhi nahi dikhti: wo server ke "secrets" mai rehti hai.
#   - Streamlit Cloud: App settings -> Secrets
#   - Local computer : .streamlit/secrets.toml  ya  .env file
# ============================================================


def get_secret(name, default=""):
    try:
        val = st.secrets.get(name)
        if val:
            return str(val).strip()
    except Exception:
        pass
    return os.getenv(name, default).strip()


# Optional: model retire ho jaye toh secrets mai GROQ_MODEL likh do (code na badalna pade)
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

SESSION_COOLDOWN_SEC = 15  # ek user do requests ke beech kam az kam itna gap
SESSION_MAX_CALLS = 20  # ek browser session mai zyada se zyada AI requests
try:
    DAILY_LIMIT = int(get_secret("DAILY_LIMIT") or 400)  # sab users ki mila kar roz ki hadd
except ValueError:
    DAILY_LIMIT = 400


@st.cache_resource
def usage_store():
    """Sab users ke darmiyan shared counter (server chalta rahe tak)."""
    return {"day": date.today().isoformat(), "count": 0, "lock": threading.Lock()}


def allow_ai_call(using_own_key, units=1):
    """Shared free key ko bachane ke liye limits. Apni key wale users par koi limit nahi."""
    if using_own_key:
        return True

    ss = st.session_state
    now = time.time()
    wait = SESSION_COOLDOWN_SEC - (now - ss.get("_last_ai_call", 0))
    if wait > 0:
        st.warning(f"⏳ Thora sabr karo, {int(wait) + 1} second baad dobara try karo.")
        return False
    if ss.get("_ai_calls", 0) + units > SESSION_MAX_CALLS:
        st.error(
            "Is session ki free limit khatam ho gayi. Page refresh karo ya sidebar"
            " mai apni free Groq key dalo."
        )
        return False

    store = usage_store()
    with store["lock"]:
        today = date.today().isoformat()
        if store["day"] != today:
            store["day"], store["count"] = today, 0
        if store["count"] + units > DAILY_LIMIT:
            st.error(
                "Aaj ki shared free limit khatam ho gayi. Kal dobara aana, ya sidebar"
                " mai apni free Groq key (console.groq.com/keys) dal kar chalao."
            )
            return False
        store["count"] += units

    ss["_last_ai_call"] = now
    ss["_ai_calls"] = ss.get("_ai_calls", 0) + units
    return True


def show_error(e):
    msg = str(e)
    low = msg.lower()
    if "429" in msg or "rate limit" in low or "quota" in low:
        st.error(
            "⏳ Free AI server abhi busy hai (limit lag gayi). 1-2 minute ruk kar"
            " dobara try karo, ya sidebar mai apni free Groq key dalo."
        )
    elif "401" in msg or "api key" in low or "authentication" in low:
        st.error("🔑 API key ka masla hai. Key dobara check karo.")
    else:
        st.error("❌ Error occurred. Dobara try karo.")
    with st.expander("Technical details"):
        st.code(msg)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
        .stApp { background-color: #F8FAFC; }

        .hero-container {
            background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
            padding: 28px 32px;
            border-radius: 12px;
            color: white;
            box-shadow: 0 4px 12px rgba(30, 58, 138, 0.15);
            margin-bottom: 25px;
        }
        .hero-title { font-size: 2.3rem; font-weight: 800; margin: 0; letter-spacing: -0.5px; }
        .hero-subtitle { font-size: 1.05rem; opacity: 0.9; margin-top: 8px; margin-bottom: 0; }

        .edu-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            margin-bottom: 20px;
        }

        .branding-card {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            color: #F8FAFC;
            padding: 24px 20px;
            border-radius: 12px;
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

        .stButton > button {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
            color: white;
            font-weight: 600;
            border-radius: 8px;
            border: none;
            padding: 10px 24px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PDF FONT + URDU (RTL) HELPERS
# ============================================================


def register_pdf_fonts():
    """Unicode font dhoondta hai (Windows Arial / Linux DejaVu). Na mile toh Helvetica."""
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
    """Urdu/Arabic text ko PDF ke liye sahi shakal deta hai (joined letters + RTL)."""
    if RTL_OK and ARABIC_RE.search(line):
        try:
            return get_display(arabic_reshaper.reshape(line))
        except Exception:
            return line
    return line


def clean_text(text):
    """AI ke output se markdown ki ** aur ` hata deta hai."""
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
    """AI ke JSON ko saaf-suthri aur safe structure mai badalta hai."""
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
            # sirf tab prefix hatao jab wo us position ka sahi letter ho
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
    """Screen par dikhane ke liye (options har line par, sequence mai)."""

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
# WORD (.docx) EXPORT - Urdu ke liye sab se bharosemand option
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
    image_b64: agar diya ho toh vision model se image padhwati hai.
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
        # sirf wo providers jinke paas vision model hai
        provider_order = [p for p in provider_order if PROVIDERS[p]["vision_models"]]

    if not provider_order:
        raise Exception(
            "Koi API key nahi mili"
            + (" jo photo (vision) padh sake. Groq ya Gemini key chahiye." if image_b64 else ".")
            + " Sidebar mai apni Groq key dalo."
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
                # vision models ke liye instruction user message ke andar hi rakhte hain
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
                    # kuch reasoning models <think>...</think> bhi bhej dete hain
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

    raise Exception("Sab providers fail ho gaye:\n- " + "\n- ".join(errors))


# ============================================================
# IMAGE (PHOTO / SCAN) -> TEXT
# ============================================================


def extract_text_from_image(llm_cfg, uploaded_file):
    """Photo (paper, answer sheet, result card) se text nikalta hai. Urdu + English."""
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
            "Photo se text nahi nikal saka. Saaf, roshan aur seedhi tasveer upload karo."
        )
    return text


def read_upload(uploaded_file, llm_cfg):
    """PDF/DOCX/TXT seedha padhta hai, photo ho toh AI (vision) se text nikalta hai."""
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
            last_error = "AI ne koi question nahi diya."
        except Exception as e:
            last_error = str(e)

    raise Exception(
        f"AI ne paper sahi format mai nahi diya ({last_error}). Dobara Generate dabao"
        " ya sidebar mai koi aur provider select karo."
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
    """PDF/Word banane mai koi masla aaye toh poori app crash na ho."""
    try:
        return builder(*args, **kwargs)
    except Exception as e:
        st.warning(f"File banane mai masla aaya: {e}")
        return None


# ============================================================
# SIDEBAR & HEADER
# ============================================================

with st.sidebar:
    st.header("⚙️ Settings")

    server_keys = {name: get_secret(cfg["env"]) for name, cfg in PROVIDERS.items()}
    server_ready = any(server_keys.values())

    if server_ready:
        st.success("✅ AI tayyar hai. Koi API key dalne ki zaroorat nahi.")
    else:
        st.warning("Server par API key set nahi hai. Neeche apni Groq key dalo.")

    with st.expander("🔑 Apni Groq key use karo (optional)", expanded=not server_ready):
        own_key = st.text_input(
            "Groq API Key",
            type="password",
            help="Free key: console.groq.com/keys. Apni key par koi limit nahi lagti.",
        ).strip()

    using_own_key = bool(own_key)
    if using_own_key:
        keys = {name: "" for name in PROVIDERS}
        keys["Groq"] = own_key
    else:
        keys = server_keys

    llm_cfg = {"keys": keys, "provider": AUTO_MODE, "custom_model": ""}
    has_any_key = any(keys.values())

    st.markdown("---")
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            padding: 16px;
            border-radius: 10px;
            border-top: 3px solid #3B82F6;
            text-align: center;
            color: white;
            box-shadow: 0 4px 10px rgba(0,0,0,0.15);
        ">
            <div style="font-size: 0.88rem; font-style: italic; color: #F1F5F9; line-height: 1.4; margin-bottom: 6px;">
                "Whoever travels a path in search of knowledge, Allah will make easy for him a path to Paradise."
            </div>
            <div style="font-size: 0.75rem; color: #60A5FA; font-weight: 600; margin-bottom: 10px;">
                — Prophet Muhammad (PBUH)<br><b>Sahih Muslim, Book 35, Hadith 6518</b>
            </div>
            <hr style="border: 0; border-top: 1px solid #334155; margin: 10px 0;">
            <div style="font-size: 0.9rem; font-weight: 700; color: #38BDF8;">
                Designed by Waheed Ali Hamouzai
            </div>
            <div style="font-size: 0.7rem; color: #94A3B8; letter-spacing: 0.5px; text-transform: uppercase; margin-top: 2px;">
                WSA Educational Community
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

NO_KEY_MSG = "⚠️ AI key set nahi hai. Sidebar mai apni Groq key dalo."

st.markdown(
    """<div class="hero-container">
<div class="hero-title">🎓 WSA Educational Test Series & Paper Builder</div>
<div class="hero-subtitle">AI-powered exam paper generation, student evaluation, and performance diagnostics.</div>
</div>""",
    unsafe_allow_html=True,
)

tab1, tab2, tab3 = st.tabs(
    [
        "📝 Generate Test Paper",
        "📊 Evaluate Student Answers",
        "🎯 Quick Test Diagnostic & Focus Areas",
    ]
)


# ============================================================
# TAB 1 — GENERATE PAPER
# ============================================================

with tab1:
    st.markdown('<div class="edu-card">', unsafe_allow_html=True)
    st.subheader("Create New Test Paper")

    col1, col2 = st.columns(2)

    with col1:
        test_type = st.selectbox(
            "Exam Category",
            [
                "General Competitive Test (CTSP / SBK / NTS)",
                "BPSC",
                "School / College (9th–12th)",
                "CS / IT Screening Test",
                "Custom Mock Test",
            ],
        )
        topic = st.text_input(
            "Topic / Subject", placeholder="e.g., Computer Science / Pak Studies"
        )
        language = st.selectbox(
            "Language Mode", ["English", "Urdu", "Bilingual (English + Urdu)"]
        )
        difficulty = st.slider("Difficulty Level", 1, 5, 3)
        diff_labels = {
            1: "Very Easy",
            2: "Easy",
            3: "Medium",
            4: "Hard",
            5: "Very Hard",
        }
        diff_level = diff_labels[difficulty]

    with col2:
        uploaded_pdf = st.file_uploader(
            f"Upload Reference Document or Photo (Optional) — {UPLOAD_HINT}",
            type=UPLOAD_TYPES,
        )
        preview_if_image(uploaded_pdf)
        mcq_count = st.number_input("Number of MCQs", 0, 50, 10)
        short_count = st.number_input("Number of Short Questions", 0, 20, 5)
        long_count = st.number_input("Number of Long Questions", 0, 10, 2)

    with st.expander("🏫 Paper Header, Marks & Time (optional)"):
        h1, h2 = st.columns(2)
        with h1:
            institute = st.text_input(
                "Institute / Academy Name", value="WSA Educational Community"
            )
            time_min = st.number_input("Time Allowed (minutes)", 5, 300, 60, step=5)
        with h2:
            mcq_marks = st.number_input("Marks per MCQ", 1, 10, 1)
            short_marks = st.number_input("Marks per Short Question", 1, 20, 2)
            long_marks = st.number_input("Marks per Long Question", 1, 50, 5)

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🚀 Generate Test Paper", use_container_width=True):
        if not has_any_key:
            st.error(NO_KEY_MSG)
        elif not topic.strip():
            st.error("⚠️ Please enter a topic or subject name.")
        elif mcq_count + short_count + long_count == 0:
            st.error("⚠️ Kam az kam ek question chahiye.")
        elif not allow_ai_call(using_own_key, units=1 + int(is_image(uploaded_pdf))):
            pass  # limit message allow_ai_call ne dikha diya
        else:
            with st.spinner("Generating professional test paper via AI... Please wait."):
                try:
                    paper = generate_test_paper(
                        llm_cfg,
                        topic,
                        uploaded_pdf,
                        test_type,
                        language,
                        mcq_count,
                        short_count,
                        long_count,
                        diff_level,
                    )
                    meta = {
                        "institute": institute.strip() or "WSA Educational Community",
                        "category": test_type,
                        "topic": topic.strip(),
                        "language": language,
                        "time_min": int(time_min),
                        "mcq_marks": int(mcq_marks),
                        "short_marks": int(short_marks),
                        "long_marks": int(long_marks),
                    }
                    meta["total_marks"] = compute_total_marks(paper, meta)

                    st.session_state["paper_data"] = paper
                    st.session_state["paper_meta"] = meta
                    st.session_state["generated_paper"] = paper_to_text(
                        paper, meta, include_key=True
                    )

                    got = (len(paper["mcqs"]), len(paper["short"]), len(paper["long"]))
                    want = (mcq_count, short_count, long_count)
                    if got != want:
                        st.warning(
                            f"AI ne MCQ/Short/Long = {got} diye, aap ne {want} maange thay."
                            " Zaroorat ho toh dobara Generate karo."
                        )
                except Exception as e:
                    show_error(e)

    if "paper_data" in st.session_state:
        paper = st.session_state["paper_data"]
        meta = st.session_state["paper_meta"]

        st.divider()
        st.subheader("📄 Generated Test Paper")
        st.markdown(paper_to_markdown(paper, meta))

        with st.expander("🔑 Answer Key (Teacher Copy)"):
            st.markdown(answer_key_markdown(paper, meta))

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

        if meta["language"] != "English" and not RTL_OK:
            st.info(
                "Urdu PDF ke liye: `pip install arabic-reshaper python-bidi` chalao."
                " Word (.docx) file mai Urdu hamesha theek aata hai."
            )


# ============================================================
# TAB 2 — EVALUATE ANSWERS
# ============================================================

with tab2:
    st.subheader("📊 Evaluate Student Answers")
    default_paper = st.session_state.get("generated_paper", "")

    col_paper, col_answer = st.columns(2)

    with col_paper:
        st.markdown("### 1️⃣ Question Paper / Answer Key")
        paper_file = st.file_uploader(
            f"Upload Paper ({UPLOAD_HINT})",
            type=UPLOAD_TYPES,
            key="p_up",
        )
        preview_if_image(paper_file)
        question_paper_text = st.text_area(
            "Or Paste Text Directly",
            value=default_paper,
            height=200,
            key="q_paper_text",
        )

    with col_answer:
        st.markdown("### 2️⃣ Student Answer Sheet")
        student_file = st.file_uploader(
            f"Upload Answers ({UPLOAD_HINT})",
            type=UPLOAD_TYPES,
            key="s_up",
        )
        preview_if_image(student_file)
        student_answers_text = st.text_area(
            "Or Paste Text Directly", height=200, key="s_answers_text"
        )

    if st.button("📊 Evaluate Answers", use_container_width=True):
        if not has_any_key:
            st.error(NO_KEY_MSG)
        else:
            p_text = process_uploaded_file(paper_file)  # photo ho toh abhi khali
            a_text = process_uploaded_file(student_file)
            paper_pasted = question_paper_text.strip()
            answers_pasted = student_answers_text.strip()
            n_images = int(is_image(paper_file)) + int(is_image(student_file))

            if not (p_text or paper_pasted or is_image(paper_file)):
                st.error(
                    "⚠️ Question paper content is missing. Please upload or paste"
                    " text."
                )
            elif not (a_text or answers_pasted or is_image(student_file)):
                st.error(
                    "⚠️ Student answer content is missing. Please upload or paste"
                    " text."
                )
            elif not allow_ai_call(using_own_key, units=1 + n_images):
                pass  # limit message allow_ai_call ne dikha diya
            else:
                with st.spinner("Evaluating student answers via AI... Please wait."):
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
        st.divider()
        st.subheader("📋 Evaluation Result")
        st.markdown(st.session_state["evaluation"])

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                "⬇️ Download Text Report (.txt)",
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
                    "⬇️ Download PDF Report (.pdf)",
                    data=eval_pdf,
                    file_name="WSA_Student_Evaluation.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )


# ============================================================
# TAB 3 — RANDOM TEST DIAGNOSTIC & IMPROVEMENT
# ============================================================

with tab3:
    st.subheader("🎯 Test Performance Diagnostic & Focus Area Planner")
    st.markdown(
        "Upload or paste any random solved test, question paper, or result card to"
        " receive a detailed breakdown of strengths, weaknesses, focus areas, and"
        " improvement plans."
    )

    col_diag1, col_diag2 = st.columns(2)

    with col_diag1:
        st.markdown("### 📄 Test File / Raw Data")
        random_test_file = st.file_uploader(
            f"Upload Test or Result Document ({UPLOAD_HINT})",
            type=UPLOAD_TYPES,
            key="random_test_up",
        )
        preview_if_image(random_test_file)
        random_test_text = st.text_area(
            "Or Paste Raw Test Content / Scores Directly",
            height=200,
            key="random_test_text",
            placeholder=(
                "e.g., Question 1: Incorrect answer selected...\nQuestion 2:"
                " Right...\nOR paste raw quiz content here."
            ),
        )

    with col_diag2:
        st.markdown("### ⚙️ Context & Target Goals (Optional)")
        user_context = st.text_area(
            "Target Exam / Student Goal",
            height=270,
            key="user_context",
            placeholder=(
                "e.g., Preparing for BPSC Computer Science / SBK Screening Test."
                " Target score is 80%+."
            ),
        )

    if st.button("🔍 Analyze Test & Generate Focus Plan", use_container_width=True):
        if not has_any_key:
            st.error(NO_KEY_MSG)
        else:
            extracted_test = process_uploaded_file(random_test_file)  # photo ho toh abhi khali
            pasted_test = random_test_text.strip()

            if not (extracted_test or pasted_test or is_image(random_test_file)):
                st.error(
                    "⚠️ Please upload a test file (PDF, DOCX, TXT, photo) or paste"
                    " text to analyze."
                )
            elif not allow_ai_call(using_own_key, units=1 + int(is_image(random_test_file))):
                pass  # limit message allow_ai_call ne dikha diya
            else:
                with st.spinner(
                    "Analyzing test data and designing personalized improvement"
                    " plan..."
                ):
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
        st.divider()
        st.subheader("📈 Diagnostic Report & Improvement Plan")
        st.markdown(st.session_state["diagnostic_analysis"])

        col_rd1, col_rd2 = st.columns(2)
        with col_rd1:
            st.download_button(
                "⬇️ Download Text Diagnostic (.txt)",
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
                    "⬇️ Download PDF Diagnostic (.pdf)",
                    data=diag_pdf,
                    file_name="WSA_Test_Diagnostic.pdf",
                    mime="application/pdf",
                    use_container_width=True,
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

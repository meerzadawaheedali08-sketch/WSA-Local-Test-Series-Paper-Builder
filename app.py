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

# Urdu / Arabic script support in PDF
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
# SECRETS + QUOTA MANAGEMENT
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
# MOBILE-FIRST ENHANCED CUSTOM CSS
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
        .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1180px; }

        #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"],
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapsedControl"] { display: none !important; }

        /* Mobile Layout Optimizations */
        @media (max-width: 768px) {
            .block-container {
                padding-left: 0.75rem !important;
                padding-right: 0.75rem !important;
                padding-top: 0.5rem !important;
            }
            .hero {
                padding: 18px 16px !important;
                border-radius: 14px !important;
                margin-bottom: 12px !important;
            }
            .hero-title { font-size: 1.35rem !important; line-height: 1.2 !important; }
            .hero-subtitle { font-size: 0.88rem !important; }

            /* Smooth horizontal tab scroll on small screens */
            .stTabs [data-baseweb="tab-list"] {
                display: flex !important;
                flex-wrap: nowrap !important;
                overflow-x: auto !important;
                scroll-snap-type: x mandatory;
                -webkit-overflow-scrolling: touch;
                gap: 6px !important;
                padding: 4px !important;
            }
            .stTabs [data-baseweb="tab"] {
                flex: 0 0 auto !important;
                scroll-snap-align: start;
                font-size: 0.85rem !important;
                height: 40px !important;
                padding: 0 14px !important;
            }

            /* Prevent auto-zooming on mobile inputs */
            input, textarea, select, [data-baseweb="select"] {
                font-size: 16px !important;
            }

            /* Stack all horizontal blocks cleanly on mobile */
            [data-testid="stHorizontalBlock"] {
                flex-direction: column !important;
                gap: 0.6rem !important;
            }
            [data-testid="stHorizontalBlock"] > [data-testid="column"] {
                width: 100% !important;
                min-width: 100% !important;
            }

            /* Touch-friendly full-width buttons */
            .stButton > button, .stDownloadButton > button {
                width: 100% !important;
                min-height: 48px !important;
                font-size: 1rem !important;
            }

            /* Option styling for mobile view readability */
            .option-block {
                display: block;
                margin-bottom: 6px;
                padding: 6px 10px;
                background: #F1F5F9;
                border-radius: 8px;
            }
        }

        /* Card and Container styling */
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
        
        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
            color: #fff;
            font-weight: 700;
            border-radius: 12px;
            padding: 12px 24px;
        }
        
        .branding-card {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            color: #F8FAFC;
            padding: 24px 18px;
            border-radius: 16px;
            text-align: center;
            margin-top: 30px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PDF FONT + UTILITY FUNCTIONS
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
ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\FB50-\uFDFF\uFE70-\uFEFF]")


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
        style = ParagraphStyle(
            name=style.name + "_rtl", parent=style, alignment=TA_RIGHT
        )
    return Paragraph(body, style)


# ============================================================
# PAPER STRUCTURE & MARKDOWN CONVERSION
# ============================================================

INSTRUCTIONS = [
    "Attempt all questions.",
    "In MCQs, choose only one correct option.",
    "Write neatly and clearly.",
]


def section_list(data, meta):
    out = []
    letters = iter("ABC")
    if data.get("mcqs"):
        n, m = len(data["mcqs"]), meta["mcq_marks"]
        out.append((next(letters), "Multiple Choice Questions", f"{n} × {m} = {n * m} Marks", "mcq", data["mcqs"]))
    if data.get("short"):
        n, m = len(data["short"]), meta["short_marks"]
        out.append((next(letters), "Short Questions", f"{n} × {m} = {n * m} Marks", "short", data["short"]))
    if data.get("long"):
        n, m = len(data["long"]), meta["long_marks"]
        out.append((next(letters), "Long Questions", f"{n} × {m} = {n * m} Marks", "long", data["long"]))
    return out


def paper_to_markdown(data, meta):
    """Formats exam paper nicely for screens with mobile-optimized layout."""
    def br(t):
        return t.replace("\n", "<br/>")

    L = [
        f"## 🎓 {meta['institute']}",
        f"**{meta['category']}**",
        "",
        f"**Subject:** {meta['topic']} | **Time:** {meta['time_min']} mins | **Marks:** {meta['total_marks']}",
        "",
        "---",
    ]
    for letter, title, marks, kind, items in section_list(data, meta):
        L += ["", f"### Section {letter}: {title} *({marks})*", ""]
        for i, it in enumerate(items, 1):
            L.append(f"**Q{i}.** {br(it['q'])}")
            if kind == "mcq":
                L.append("")
                # Render vertically for easy mobile reading
                for j, opt in enumerate(it["options"]):
                    L.append(f"<div class='option-block'><b>({LETTERS[j]})</b> {br(opt)}</div>")
            L.append("")
    return "\n".join(L)

# ============================================================
# APP INTERFACE LAYOUT
# ============================================================

def main():
    st.markdown(
        """
        <div class="hero">
            <div class="hero-label">WSA Educational Portal</div>
            <div class="hero-title">Test Series & Paper Builder</div>
            <div class="hero-subtitle">Generate professional, high-quality exam papers with mobile-friendly ease.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab1, tab2 = st.tabs(["📄 Paper Builder", "ℹ️ Help & Info"])

    with tab1:
        st.subheader("Configure Exam Paper")
        
        with st.container():
            col1, col2 = st.columns(2)
            with col1:
                institute = st.text_input("Institute / School Name", "WSA Academy")
                topic = st.text_input("Subject / Topic Name", "Computer Science")
            with col2:
                category = st.selectbox("Exam Category", ["First Term Test", "Midterm Exam", "Final Board Exam", "Class Test"])
                time_min = st.number_input("Time Allowed (Minutes)", min_value=10, max_value=180, value=60)

        st.markdown("---")
        st.subheader("Questions Distribution")

        col_q1, col_q2, col_q3 = st.columns(3)
        with col_q1:
            mcq_count = st.number_input("MCQs Count", 0, 50, 5)
            mcq_marks = st.number_input("Marks per MCQ", 1, 5, 1)
        with col_q2:
            short_count = st.number_input("Short Questions Count", 0, 20, 3)
            short_marks = st.number_input("Marks per Short Qs", 1, 10, 2)
        with col_q3:
            long_count = st.number_input("Long Questions Count", 0, 10, 2)
            long_marks = st.number_input("Marks per Long Qs", 1, 20, 5)

        total_marks = (mcq_count * mcq_marks) + (short_count * short_marks) + (long_count * long_marks)
        st.info(f"📊 **Calculated Total Marks:** {total_marks}")

        if st.button("🚀 Generate Paper", type="primary"):
            meta = {
                "institute": institute,
                "topic": topic,
                "category": category,
                "time_min": time_min,
                "total_marks": total_marks,
                "mcq_marks": mcq_marks,
                "short_marks": short_marks,
                "long_marks": long_marks,
            }
            # Mock paper state for demonstration
            sample_data = {
                "mcqs": [
                    {"q": "Which component is known as the brain of the computer?", "options": ["RAM", "CPU", "Hard Disk", "Monitor"], "answer": "B"},
                    {"q": "What does RAM stand for?", "options": ["Read Access Memory", "Random Access Memory", "Rapid Action Memory", "Run Auto Memory"], "answer": "B"},
                ],
                "short": [
                    {"q": "Define System Software with two examples.", "answer": "Software that manages basic computer operations."},
                ],
                "long": [
                    {"q": "Explain the architecture of a computer system in detail.", "answer": "Detail answer covering CPU, Input/Output, and Memory units."},
                ],
            }
            st.session_state["generated_paper"] = (sample_data, meta)
            st.success("Paper generated successfully!")

        if "generated_paper" in st.session_state:
            data, meta = st.session_state["generated_paper"]
            st.markdown("---")
            st.subheader("📋 Generated Exam Paper Preview")
            st.markdown(paper_to_markdown(data, meta), unsafe_allow_html=True)

    with tab2:
        st.write("This application helps teachers generate exams easily on both smartphones and desktops.")

    # Footer Branding
    st.markdown(
        """
        <div class="branding-card">
            <div class="hadith-quote">"Seeking knowledge is an obligation upon every Muslim."</div>
            <div class="hadith-ref">Sunan Ibn Majah 224</div>
            <hr class="footer-divider" style="border: 0; border-top: 1px solid #334155; margin: 12px 0;">
            <div class="branding-name">WSA Educational Test Series</div>
            <div class="branding-tag">Designed by Waheed Ali Hamouzai</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()

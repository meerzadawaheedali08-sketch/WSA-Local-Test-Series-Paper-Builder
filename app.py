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

def is_image(uploaded_file):
    return uploaded_file is not None and uploaded_file.name.lower().endswith(IMAGE_EXTS)

def prepare_image_b64(uploaded_file, max_side=1800):
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
# SECRETS & QUOTA MANAGEMENT
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
        st.error("You have reached the usage limit for this session.")
        return False

    store = usage_store()
    with store["lock"]:
        today = date.today().isoformat()
        if store["day"] != today:
            store["day"], store["count"] = today, 0
        if store["count"] + units > DAILY_LIMIT:
            st.error("The daily free limit has been reached.")
            return False
        store["count"] += units

    ss["_last_ai_call"] = now
    ss["_ai_calls"] = ss.get("_ai_calls", 0) + units
    return True

# ============================================================
# CUSTOM DASHBOARD CSS (MATCHING IMAGE 1 DESIGN)
# ============================================================

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        :root {
            --bg-dark: #091527;
            --card-bg: #112239;
            --text-main: #ffffff;
            --text-sub: #94a3b8;
        }

        .stApp {
            background-color: var(--bg-dark);
            font-family: 'Inter', sans-serif;
            color: var(--text-main);
        }

        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1100px;
        }

        #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none !important; }

        .hero-banner { margin-bottom: 20px; }
        .hero-title { font-size: 2rem; font-weight: 800; color: #ffffff; margin-bottom: 4px; }
        .hero-sub { color: var(--text-sub); font-size: 0.95rem; margin-bottom: 12px; }
        .badge-btn {
            background: rgba(16, 185, 129, 0.2);
            color: #10b981;
            border: 1px solid #10b981;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.8rem;
            font-weight: 600;
            display: inline-block;
        }

        .section-heading { font-size: 1.3rem; font-weight: 700; margin: 24px 0 12px 0; color: #ffffff; }

        /* Dashboard Grid Cards */
        .grid-container {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 14px;
            margin-bottom: 24px;
        }
        .action-card {
            padding: 18px 20px;
            border-radius: 12px;
            color: white;
            cursor: pointer;
        }
        .card-blue { background: linear-gradient(135deg, #1d6bf3, #1557c0); }
        .card-green { background: linear-gradient(135deg, #0d8a6a, #096b52); }
        .card-purple { background: linear-gradient(135deg, #6b46c1, #53349c); }
        .card-orange { background: linear-gradient(135deg, #c05621, #9c4115); }

        .card-icon { font-size: 1.4rem; margin-bottom: 6px; }
        .card-title { font-size: 1.1rem; font-weight: 700; margin-bottom: 2px; }
        .card-desc { font-size: 0.82rem; opacity: 0.85; }

        /* UI Select Boxes & Inputs */
        div[data-baseweb="select"] > div {
            background-color: var(--card-bg) !important;
            border: 1px solid #1e3a5f !important;
            border-radius: 10px !important;
            color: white !important;
            min-height: 48px !important;
        }
        input {
            background-color: var(--card-bg) !important;
            border: 1px solid #1e3a5f !important;
            color: white !important;
            border-radius: 10px !important;
            min-height: 48px !important;
            font-size: 16px !important;
        }

        .stButton > button[kind="primary"] {
            width: 100%;
            background: linear-gradient(90deg, #1d6bf3, #0052d4);
            color: white;
            font-weight: 700;
            border: none;
            border-radius: 12px;
            height: 50px;
            font-size: 1rem;
            box-shadow: 0 4px 15px rgba(29, 107, 243, 0.4);
        }

        @media (max-width: 640px) {
            .grid-container { grid-template-columns: 1fr; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# PDF & TEXT GENERATION UTILITIES
# ============================================================

def register_pdf_fonts():
    candidates = [
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/ariali.ttf", "C:/Windows/Fonts/arialbi.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf"),
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
                for j, opt in enumerate(it["options"]):
                    L.append(f"&nbsp;&nbsp;&nbsp;&nbsp;**({LETTERS[j]})** {br(opt)}")
            L.append("")
    return "\n".join(L)

# ============================================================
# MAIN APPLICATION INTERFACE
# ============================================================

def main():
    # Top Dashboard Header
    st.markdown(
        """
        <div class="hero-banner">
            <div class="hero-title">Learn. Practice. Improve.</div>
            <div class="hero-sub">Create papers, practice MCQs and track your progress.</div>
            <span class="badge-btn">Your Learning Dashboard</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Action Cards Grid Section
    st.markdown('<div class="section-heading">What do you want to do?</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="grid-container">
            <div class="action-card card-blue">
                <div class="card-icon">📄</div>
                <div class="card-title">Create Paper</div>
                <div class="card-desc">Build a new exam</div>
            </div>
            <div class="action-card card-green">
                <div class="card-icon">📋</div>
                <div class="card-title">Test Series</div>
                <div class="card-desc">Practice by subject</div>
            </div>
            <div class="action-card card-purple">
                <div class="card-icon">📊</div>
                <div class="card-title">My Results</div>
                <div class="card-desc">Review performance</div>
            </div>
            <div class="action-card card-orange">
                <div class="card-icon">🎯</div>
                <div class="card-title">MCQ Practice</div>
                <div class="card-desc">Improve your concepts</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Form Configurator Section
    st.markdown('<div class="section-heading">Start Creating Your Paper</div>', unsafe_allow_html=True)
    st.caption("Class, subject, difficulty and question settings")

    col1, col2 = st.columns(2)
    with col1:
        institute = st.text_input("Institute / School Name", "WSA Academy")
        topic = st.text_input("Subject / Topic Name", "Computer Science")
        class_opt = st.selectbox("Select Class", ["9th Class", "10th Class", "11th Class (FSc)", "12th Class (FSc)", "BS Computer Science"], index=0)
    with col2:
        category = st.selectbox("Exam Category", ["First Term Test", "Midterm Exam", "Final Board Exam", "Class Test"])
        time_min = st.number_input("Time Allowed (Minutes)", min_value=10, max_value=180, value=60)

    st.markdown("---")
    st.subheader("Questions Settings")

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

    if st.button("Continue ➔", type="primary"):
        meta = {
            "institute": institute,
            "topic": f"{class_opt} - {topic}",
            "category": category,
            "time_min": time_min,
            "total_marks": total_marks,
            "mcq_marks": mcq_marks,
            "short_marks": short_marks,
            "long_marks": long_marks,
        }
        
        # Generated Data Structure
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
        st.success("Paper Generated Successfully!")

    if "generated_paper" in st.session_state:
        data, meta = st.session_state["generated_paper"]
        st.markdown("---")
        st.subheader("📋 Generated Exam Paper Preview")
        st.markdown(paper_to_markdown(data, meta), unsafe_allow_html=True)

if __name__ == "__main__":
    main()

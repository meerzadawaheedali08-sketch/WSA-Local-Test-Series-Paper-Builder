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

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    RTL_OK = True
except Exception:
    RTL_OK = False

load_dotenv()

st.set_page_config(
    page_title="WSA Educational Test Series & Paper Builder",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================
# MODERN DASHBOARD & MOBILE RESPONSIVE CSS
# ============================================================

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        :root {
            --bg-dark: #091527;
            --card-bg: #112239;
            --blue-card: #1d6bf3;
            --green-card: #0d8a6a;
            --purple-card: #6b46c1;
            --orange-card: #c05621;
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

        /* Header / Banner Styling */
        .hero-banner {
            margin-bottom: 20px;
        }
        .hero-title {
            font-size: 2rem;
            font-weight: 800;
            color: #ffffff;
            margin-bottom: 4px;
        }
        .hero-sub {
            color: var(--text-sub);
            font-size: 0.95rem;
            margin-bottom: 12px;
        }
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

        .section-heading {
            font-size: 1.3rem;
            font-weight: 700;
            margin: 20px 0 12px 0;
            color: #ffffff;
        }

        /* Action Cards Grid */
        .grid-container {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 14px;
            margin-bottom: 28px;
        }

        .action-card {
            padding: 18px 20px;
            border-radius: 12px;
            color: white;
            cursor: pointer;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .action-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.3);
        }

        .card-blue { background: linear-gradient(135deg, #1d6bf3, #1557c0); }
        .card-green { background: linear-gradient(135deg, #0d8a6a, #096b52); }
        .card-purple { background: linear-gradient(135deg, #6b46c1, #53349c); }
        .card-orange { background: linear-gradient(135deg, #c05621, #9c4115); }

        .card-icon { font-size: 1.4rem; margin-bottom: 8px; }
        .card-title { font-size: 1.1rem; font-weight: 700; margin-bottom: 2px; }
        .card-desc { font-size: 0.82rem; opacity: 0.85; }

        /* Form Controls Overhaul */
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

        .stButton > button {
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

        /* Mobile Adjustments */
        @media (max-width: 640px) {
            .grid-container {
                grid-template-columns: 1fr;
                gap: 10px;
            }
            .hero-title { font-size: 1.5rem; }
            .action-card { padding: 14px 16px; }
            .block-container { padding-left: 0.8rem; padding-right: 0.8rem; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# DASHBOARD INTERFACE
# ============================================================

def main():
    # Top Banner Section
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

    # Action Cards Grid
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

    # Creation Form Section
    st.markdown('<div class="section-heading">Start Creating Your Paper</div>', unsafe_allow_html=True)
    st.caption("Class, subject, difficulty and question settings")

    class_opt = st.selectbox("Select Class", ["9th Class", "10th Class", "11th Class (FSc)", "12th Class (FSc)", "BS Computer Science"], index=0)
    subject_opt = st.selectbox("Select Subject", ["Computer Science", "Physics", "Chemistry", "Mathematics", "English"], index=0)

    col1, col2 = st.columns(2)
    with col1:
        mcq_count = st.number_input("MCQs Count", min_value=1, max_value=50, value=10)
    with col2:
        short_count = st.number_input("Short Questions Count", min_value=0, max_value=20, value=5)

    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("Continue ➔"):
        st.success(f"{class_opt} - {subject_opt} ke liye paper process ho raha hai!")

if __name__ == "__main__":
    main()

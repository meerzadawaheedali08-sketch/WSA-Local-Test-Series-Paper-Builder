```python
import os
import time
import streamlit as st
import pypdf
from google import genai
from google.genai import types
from dotenv import load_dotenv


# ============================================================
# LOAD ENVIRONMENT VARIABLES
# ============================================================

load_dotenv()


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="WSA Educational Test Series & Paper Builder",
    page_icon="🎓",
    layout="wide"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
        .stApp {
            background-color: #F8FAFC;
        }

        .hero-container {
            background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
            padding: 28px 32px;
            border-radius: 12px;
            color: white;
            box-shadow: 0 4px 12px rgba(30, 58, 138, 0.15);
            margin-bottom: 25px;
        }

        .hero-title {
            font-size: 2.3rem;
            font-weight: 800;
            margin: 0;
            letter-spacing: -0.5px;
        }

        .hero-subtitle {
            font-size: 1.05rem;
            opacity: 0.9;
            margin-top: 8px;
            margin-bottom: 0;
        }

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
            padding: 16px;
            border-radius: 10px;
            border-left: 4px solid #3B82F6;
            text-align: center;
            margin-top: 25px;
        }

        .branding-name {
            font-size: 1.05rem;
            font-weight: 700;
            color: #60A5FA;
            margin-bottom: 4px;
        }

        .branding-tag {
            font-size: 0.8rem;
            color: #94A3B8;
            letter-spacing: 0.5px;
        }

        .stButton > button {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
            color: white;
            font-weight: 600;
            border-radius: 8px;
            border: none;
            padding: 10px 24px;
            box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
            transition: all 0.2s ease;
        }

        .stButton > button:hover {
            background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
            box-shadow: 0 4px 8px rgba(37, 99, 235, 0.3);
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(uploaded_file):
    """Extract text from a PDF file."""

    try:
        reader = pypdf.PdfReader(uploaded_file)

        text = ""

        for page in reader.pages:
            extracted = page.extract_text()

            if extracted:
                text += extracted + "\n"

        return text.strip()

    except Exception:
        return ""


# ============================================================
# GEMINI API CALL WITH RETRY
# ============================================================

def call_gemini_with_retry(
    client,
    contents,
    temperature=0.3,
    max_tokens=3500,
    retries=3
):
    """
    Call Gemini API with automatic retry for temporary
    rate-limit/server errors.
    """

    for attempt in range(retries):

        try:

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=temperature,
                    max_output_tokens=max_tokens,
                )
            )

            if not response.text:
                raise Exception("Gemini returned an empty response.")

            return response.text

        except Exception as e:

            error_text = str(e)

            # Retry temporary errors
            temporary_error = (
                "429" in error_text
                or "503" in error_text
                or "UNAVAILABLE" in error_text
                or "RESOURCE_EXHAUSTED" in error_text
                or "rate limit" in error_text.lower()
            )

            if temporary_error and attempt < retries - 1:

                wait_time = 2 * (attempt + 1)

                time.sleep(wait_time)

                continue

            raise e


# ============================================================
# GENERATE TEST PAPER
# ============================================================

def generate_test_paper(
    api_key,
    topic,
    uploaded_pdf,
    test_type,
    mcq_count,
    short_count,
    long_count,
    diff_level
):
    """
    Generate complete question paper and answer key.
    """

    try:

        client = genai.Client(api_key=api_key)

        prompt = f"""
You are an expert examiner for competitive testing agencies
and educational boards in Balochistan/Pakistan.

Generate a professional Question Paper and complete Answer Key.

Target Test Category/Role:
{test_type}

Topic / Subject:
{topic}

Difficulty Level:
{diff_level}

Paper Pattern:

MCQs:
{mcq_count}

Short Questions:
{short_count}

Long/Descriptive Questions:
{long_count}

IMPORTANT REQUIREMENTS:

1. MCQs must have exactly four options:
A, B, C, D.

2. Questions must match the selected topic.

3. Difficulty must match the selected level.

4. If a reference PDF is provided, use it as the primary
   reference material.

5. Do not invent information that contradict

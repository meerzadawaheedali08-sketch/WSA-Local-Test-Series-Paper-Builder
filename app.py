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
            background: linear-gradient(
                135deg,
                #1E3A8A 0%,
                #3B82F6 100%
            );
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
            background: linear-gradient(
                135deg,
                #0F172A 0%,
                #1E293B 100%
            );
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
            background: linear-gradient(
                135deg,
                #2563EB 0%,
                #1D4ED8 100%
            );
            color: white;
            font-weight: 600;
            border-radius: 8px;
            border: none;
            padding: 10px 24px;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(uploaded_file):
    """Extract text from uploaded PDF."""
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
# GEMINI API CALL WITH ROBUST FALLBACK & AUTO-DISCOVERY
# ============================================================

def call_gemini_with_retry(
    client,
    contents,
    temperature=0.3,
    max_tokens=3500,
    retries=2
):
    """
    Call Gemini API with robust model fallback using the new google-genai SDK.
    """
    # Active Gemini model identifier list for the google-genai SDK
    candidate_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-2.5-pro"
    ]

    last_error_message = ""

    for model_name in candidate_models:
        for attempt in range(retries):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        temperature=temperature,
                        max_output_tokens=max_tokens
                    )
                )

                if response and response.text:
                    return response.text

            except Exception as e:
                error_text = str(e)
                last_error_message = error_text

                # Invalid API Key / Authentication Error -> Fail fast
                if "API_KEY_INVALID" in error_text or "403" in error_text or "PermissionDenied" in error_text:
                    raise Exception(
                        "❌ Invalid API Key! Kripya Google AI Studio se sahi API Key copy karke Enter karein."
                    )

                # If 404/NOT_FOUND -> Move immediately to the next candidate model
                if "404" in error_text or "NOT_FOUND" in error_text or "not found" in error_text.lower():
                    break

                # Rate Limit / Transient Error -> Retry with exponential backoff
                temporary_error = (
                    "429" in error_text
                    or "503" in error_text
                    or "UNAVAILABLE" in error_text
                    or "RESOURCE_EXHAUSTED" in error_text
                    or "rate limit" in error_text.lower()
                )

                if temporary_error and attempt < retries - 1:
                    time.sleep(2 * (attempt + 1))
                    continue

    # Final Exception Handler with User Instructions
    raise Exception(
        f"API Connection Failed!\n\n"
        f"Kripya niche diye gaye steps check karein:\n"
        f"1. **API Key Verify Karein:** Google AI Studio (aistudio.google.com) par jakar new API key generate karein.\n"
        f"2. **Project Billing / Quota:** Check karein ki aapke Google account par Free Tier Quota limit end toh nahi ho gayi.\n"
        f"3. **Network Connection:** Apna Internet connection check karein.\n\n"
        f"Error Details: {last_error_message}"
    )


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
    Handles both normal text PDFs and scanned image PDFs.
    """
    client = genai.Client(api_key=api_key)
    contents = []

    pdf_text = ""
    if uploaded_pdf is not None:
        pdf_text = extract_pdf_text(uploaded_pdf)
        if len(pdf_text) > 20000:
            pdf_text = pdf_text[:20000]

        # Scanned PDF Fallback: Direct Part attachment for OCR
        if not pdf_text:
            uploaded_pdf.seek(0)
            bytes_data = uploaded_pdf.getvalue()
            contents.append(
                types.Part.from_bytes(
                    data=bytes_data,
                    mime_type="application/pdf"
                )
            )

    prompt = f"""
You are an expert examiner for competitive testing agencies
and educational boards in Balochistan and Pakistan.

Create a professional examination paper.

TARGET TEST CATEGORY / ROLE:
{test_type}

TOPIC / SUBJECT:
{topic}

DIFFICULTY LEVEL:
{diff_level}

QUESTION COUNTS:
MCQs: {mcq_count}
SHORT QUESTIONS: {short_count}
LONG / DESCRIPTIVE QUESTIONS: {long_count}

IMPORTANT REQUIREMENTS:
1. MCQs must have exactly four options: A, B, C and D.
2. Every question must be relevant to the selected topic.
3. Match the requested difficulty level.
4. Avoid duplicate questions.
5. MCQs must have only one clearly correct answer.
6. Short questions should require concise answers.
7. Long questions should require detailed answers.
8. Provide a complete answer key.
9. Keep the paper professional and suitable for Pakistani educational or competitive testing.
10. If reference material is provided below or as an attached PDF, use it as the primary source.
11. Clearly separate the QUESTION PAPER from the ANSWER KEY.

REQUIRED OUTPUT FORMAT:

========================================
WSA EDUCATIONAL TEST SERIES
========================================

TEST CATEGORY: {test_type}
TOPIC: {topic}
DIFFICULTY: {diff_level}

========================================
SECTION A — MCQs
========================================
1. Question
   A) Option
   B) Option
   C) Option
   D) Option

========================================
SECTION B — SHORT QUESTIONS
========================================
1. Question

========================================
SECTION C — LONG QUESTIONS
========================================
1. Question

========================================
ANSWER KEY
========================================

MCQs:
1. A

SHORT QUESTIONS:
1. Model Answer:

LONG QUESTIONS:
1. Model Answer:

REFERENCE MATERIAL FROM TEXT:
{pdf_text if pdf_text else "Attached PDF processed directly via Vision/OCR."}
"""

    contents.append(prompt)

    result = call_gemini_with_retry(
        client=client,
        contents=contents,
        temperature=0.3,
        max_tokens=6000,
        retries=2
    )

    return result


# ============================================================
# EVALUATE STUDENT ANSWERS
# ============================================================

def evaluate_student_answers(
    api_key,
    question_paper,
    student_answers
):
    """
    Evaluate student answers using Gemini.
    """
    client = genai.Client(api_key=api_key)

    prompt = f"""
You are an experienced examiner.

Evaluate the student's answers against the
provided question paper and answer key.

QUESTION PAPER:
{question_paper}

STUDENT ANSWERS:
{student_answers}

Provide the evaluation in this format:

========================================
STUDENT EVALUATION
========================================

Total Marks:
Obtained Marks:
Percentage:
Performance:

========================================
QUESTION-WISE EVALUATION
========================================

Question 1:
Correct / Incorrect / Partially Correct

Marks:
Explanation:

========================================
FINAL FEEDBACK
========================================

Strengths:
- 

Weak Areas:
- 

Suggestions:
- 
"""

    result = call_gemini_with_retry(
        client=client,
        contents=[prompt],
        temperature=0.2,
        max_tokens=5000,
        retries=2
    )

    return result


# ============================================================
# SIDEBAR
# ============================================================

    api_key = api_key_input or env_api_key

    st.divider()

    st.subheader("About")
    st.write(
        "WSA Educational Test Series & Paper Builder "
        "is an AI-powered educational tool for creating "
        "practice papers and evaluating student answers."
    )
    st.info(
        "Your Gemini API key is used only for generating "
        "and evaluating content."
    )


# ============================================================
# HERO SECTION
# ============================================================

st.markdown(
    """<div class="hero-container">
<div class="hero-title">🎓 WSA Educational Test Series & Paper Builder</div>
<div class="hero-subtitle">AI-powered exam paper generation, practice testing and answer evaluation.</div>
</div>""",
    unsafe_allow_html=True
)

with st.sidebar:
    st.header("⚙️ Settings")
    st.subheader("Gemini API")

    env_api_key = os.getenv("GEMINI_API_KEY", "")
    api_key_input = st.text_input(
        "Enter Gemini API Key",
        value=env_api_key,
        type="password",
        placeholder="AIza..."
    )



# ============================================================
# TABS
# ============================================================

tab1, tab2 = st.tabs(
    [
        "📝 Generate Test Paper",
        "📊 Evaluate Student Answers"
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
                "Custom Mock Test"
            ]
        )

        topic = st.text_input(
            "Topic / Subject",
            placeholder="Example: Computer Science, English Grammar, Pakistan Studies"
        )

        difficulty = st.slider(
            "Difficulty Level",
            min_value=1,
            max_value=5,
            value=3
        )

        difficulty_labels = {
            1: "Very Easy",
            2: "Easy",
            3: "Medium",
            4: "Hard",
            5: "Very Hard"
        }

        diff_level = difficulty_labels[difficulty]

    with col2:
        uploaded_pdf = st.file_uploader(
            "Upload Syllabus / Chapter PDF (Optional)",
            type=["pdf"]
        )

        mcq_count = st.number_input(
            "Number of MCQs",
            min_value=0,
            max_value=60,
            value=10,
            step=1
        )

        short_count = st.number_input(
            "Number of Short Questions",
            min_value=0,
            max_value=20,
            value=5,
            step=1
        )

        long_count = st.number_input(
            "Number of Long Questions",
            min_value=0,
            max_value=10,
            value=2,
            step=1
        )

    st.markdown("</div>", unsafe_allow_html=True)

    generate_button = st.button(
        "🚀 Generate Test Paper",
        use_container_width=True
    )

    if generate_button:
        if not api_key:
            st.error("⚠️ Please enter your Gemini API key in the sidebar.")
        elif not topic.strip():
            st.error("⚠️ Please enter a topic or subject.")
        elif mcq_count == 0 and short_count == 0 and long_count == 0:
            st.error("⚠️ Please select at least one question count.")
        else:
            with st.spinner("Generating your test paper..."):
                try:
                    generated_result = generate_test_paper(
                        api_key=api_key,
                        topic=topic,
                        uploaded_pdf=uploaded_pdf,
                        test_type=test_type,
                        mcq_count=mcq_count,
                        short_count=short_count,
                        long_count=long_count,
                        diff_level=diff_level
                    )

                    st.session_state["generated_paper"] = generated_result

                except Exception as e:
                    st.error("❌ Failed to Generate Test Paper")
                    st.warning(str(e))

    if "generated_paper" in st.session_state:
        st.divider()
        st.subheader("📄 Generated Test Paper")
        st.markdown(st.session_state["generated_paper"])

        st.download_button(
            label="⬇️ Download Test Paper",
            data=st.session_state["generated_paper"],
            file_name="WSA_Test_Paper.txt",
            mime="text/plain",
            use_container_width=True
        )


# ============================================================
# TAB 2 — EVALUATE ANSWERS
# ============================================================

with tab2:
    st.subheader("📊 Evaluate Student Answers")

    default_paper = st.session_state.get("generated_paper", "")

    question_paper = st.text_area(
        "Paste Question Paper / Answer Key",
        value=default_paper,
        height=300,
        placeholder="Paste the generated question paper and answer key here..."
    )

    student_answers = st.text_area(
        "Paste Student Answers",
        height=300,
        placeholder="Paste the student's answers here..."
    )

    evaluate_button = st.button(
        "📊 Evaluate Answers",
        use_container_width=True
    )

    if evaluate_button:
        if not api_key:
            st.error("⚠️ Please enter your Gemini API key in the sidebar.")
        elif not question_paper.strip():
            st.error("⚠️ Please paste the question paper.")
        elif not student_answers.strip():
            st.error("⚠️ Please paste the student's answers.")
        else:
            with st.spinner("Evaluating student answers..."):
                try:
                    evaluation = evaluate_student_answers(
                        api_key=api_key,
                        question_paper=question_paper,
                        student_answers=student_answers
                    )

                    st.session_state["evaluation"] = evaluation

                except Exception as e:
                    st.error("❌ Failed to Evaluate Answers")
                    st.warning(str(e))

    if "evaluation" in st.session_state:
        st.divider()
        st.subheader("📋 Evaluation Result")
        st.markdown(st.session_state["evaluation"])

        st.download_button(
            label="⬇️ Download Evaluation",
            data=st.session_state["evaluation"],
            file_name="WSA_Student_Evaluation.txt",
            mime="text/plain",
            use_container_width=True
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """<div class="branding-card">
<div class="branding-name">Designed by Waheed Ali Hamouzai</div>
<div class="branding-tag">WSA Educational Community • AI Powered Learning</div>
</div>""",
    unsafe_allow_html=True
)

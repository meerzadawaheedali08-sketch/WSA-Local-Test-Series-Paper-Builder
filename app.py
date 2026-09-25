import os
import time

import streamlit as st
import pypdf
import docx

from google import genai
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
# HELPER FUNCTIONS FOR FILE EXTRACTION & HANDLING
# ============================================================

def process_uploaded_file(uploaded_file):
    """
    Extracts text from PDF/DOCX/TXT files or returns raw binary data
    with MIME type for Gemini Multimodal API (Image / Scanned PDF).
    """
    if uploaded_file is None:
        return "", None

    filename = uploaded_file.name.lower()

    # 1. Plain Text File (.txt)
    if filename.endswith(".txt"):
        try:
            return uploaded_file.getvalue().decode("utf-8").strip(), None
        except Exception:
            return "", None

    # 2. Word Document (.docx)
    elif filename.endswith(".docx"):
        try:
            doc = docx.Document(uploaded_file)
            full_text = [p.text for p in doc.paragraphs if p.text]
            return "\n".join(full_text).strip(), None
        except Exception:
            return "", None

    # 3. PDF File (.pdf)
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

        # If text extracted successfully, return it
        if text:
            if len(text) > 15000:
                text = text[:15000]
            return text, None
        else:
            # Scanned PDF Fallback (Pass bytes directly for OCR)
            uploaded_file.seek(0)
            return "", {
                "mime_type": "application/pdf",
                "data": uploaded_file.getvalue()
            }

    # 4. Images (.png, .jpg, .jpeg)
    elif filename.endswith((".png", ".jpg", ".jpeg")):
        mime_type = "image/png" if filename.endswith(".png") else "image/jpeg"
        uploaded_file.seek(0)
        return "", {
            "mime_type": mime_type,
            "data": uploaded_file.getvalue()
        }

    return "", None


# ============================================================
# GEMINI API CALL WITH RELIABLE FALLBACK
# ============================================================

def call_gemini_with_retry(
    api_key,
    prompt_text,
    file_attachments=None,
    temperature=0.3
):
    """
    Call Gemini API using official SDK with multi-model fallback.
    """
    client = genai.Client(api_key=api_key)

    candidate_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash"
    ]

    last_error = ""

    for model_name in candidate_models:
        try:
            contents = []

            # Append any binary file parts (Images / Scanned PDFs)
            if file_attachments:
                for attachment in file_attachments:
                    if attachment:
                        contents.append(attachment)

            # Append Prompt Text
            contents.append(prompt_text)

            response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config={
                    "temperature": temperature
                }
            )

            if response and response.text:
                return response.text

        except Exception as e:
            last_error = str(e)
            if "404" in last_error or "NOT_FOUND" in last_error or "not found" in last_error.lower():
                continue
            time.sleep(1)

    raise Exception(
        f"Model connection failed. Google AI Studio se new API key check karein.\n\nError: {last_error}"
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
    pdf_text, pdf_part = process_uploaded_file(uploaded_pdf)
    file_attachments = [pdf_part] if pdf_part else None

    prompt = f"""
You are an expert examiner for competitive testing agencies and educational boards in Pakistan.

Create a professional examination paper.

TARGET TEST CATEGORY / ROLE: {test_type}
TOPIC / SUBJECT: {topic}
DIFFICULTY LEVEL: {diff_level}

QUESTION COUNTS:
MCQs: {mcq_count}
SHORT QUESTIONS: {short_count}
LONG / DESCRIPTIVE QUESTIONS: {long_count}

REQUIREMENTS:
1. MCQs must have four options (A, B, C, D) with 1 correct option.
2. Provide short questions and long questions as requested.
3. Provide a complete ANSWER KEY at the bottom.
4. Keep the paper clean and professional.

{f"REFERENCE TEXT:\n{pdf_text}" if pdf_text else ""}
"""

    return call_gemini_with_retry(
        api_key=api_key,
        prompt_text=prompt,
        file_attachments=file_attachments,
        temperature=0.3
    )


# ============================================================
# EVALUATE STUDENT ANSWERS
# ============================================================

def evaluate_student_answers(
    api_key,
    paper_text,
    paper_attachment,
    answers_text,
    answers_attachment
):
    file_attachments = []
    if paper_attachment:
        file_attachments.append(paper_attachment)
    if answers_attachment:
        file_attachments.append(answers_attachment)

    prompt = f"""
You are an experienced examiner and paper grader.

Evaluate the student's answer sheet against the provided question paper and answer key.

QUESTION PAPER / ANSWER KEY CONTENT:
{paper_text if paper_text else "Question Paper attached as document/image."}

STUDENT ANSWERS CONTENT:
{answers_text if answers_text else "Student answers attached as document/image."}

Provide a comprehensive, objective, and clear evaluation using this exact structure:

========================================
STUDENT EVALUATION SUMMARY
========================================
Total Marks:
Obtained Marks:
Percentage:
Grade / Performance:

========================================
QUESTION-WISE BREAKDOWN
========================================
(For each question present in the paper)
1. Question:
   Status: Correct / Partially Correct / Incorrect
   Marks Awarded:
   Feedback / Explanation:

========================================
STRENGTHS & WEAKNESSES
========================================
Strengths:
- 

Areas for Improvement:
- 

Suggestions for Improvement:
- 
"""

    return call_gemini_with_retry(
        api_key=api_key,
        prompt_text=prompt,
        file_attachments=file_attachments if file_attachments else None,
        temperature=0.2
    )


# ============================================================
# SIDEBAR
# ============================================================

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

    api_key = api_key_input or env_api_key

    st.divider()
    st.subheader("About")
    st.write("WSA Educational Test Series & Paper Builder")


# ============================================================
# HERO SECTION
# ============================================================

st.markdown(
    """<div class="hero-container">
<div class="hero-title">🎓 WSA Educational Test Series & Paper Builder</div>
<div class="hero-subtitle">AI-powered exam paper generation and student answer evaluation.</div>
</div>""",
    unsafe_allow_html=True
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
            placeholder="Example: Computer Science, English Grammar"
        )

        difficulty = st.slider("Difficulty Level", 1, 5, 3)
        diff_labels = {1: "Very Easy", 2: "Easy", 3: "Medium", 4: "Hard", 5: "Very Hard"}
        diff_level = diff_labels[difficulty]

    with col2:
        uploaded_pdf = st.file_uploader(
            "Upload Syllabus / Reference Document (PDF, DOCX, TXT)",
            type=["pdf", "docx", "txt"]
        )
        mcq_count = st.number_input("Number of MCQs", 0, 50, 10)
        short_count = st.number_input("Number of Short Questions", 0, 20, 5)
        long_count = st.number_input("Number of Long Questions", 0, 10, 2)

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🚀 Generate Test Paper", use_container_width=True):
        if not api_key:
            st.error("⚠️ Please enter your Gemini API key in the sidebar.")
        elif not topic.strip():
            st.error("⚠️ Please enter a topic.")
        else:
            with st.spinner("Generating test paper..."):
                try:
                    res = generate_test_paper(
                        api_key, topic, uploaded_pdf, test_type,
                        mcq_count, short_count, long_count, diff_level
                    )
                    st.session_state["generated_paper"] = res
                except Exception as e:
                    st.error(f"❌ Error: {e}")

    if "generated_paper" in st.session_state:
        st.divider()
        st.subheader("📄 Generated Test Paper")
        st.markdown(st.session_state["generated_paper"])
        st.download_button(
            "⬇️ Download Paper",
            st.session_state["generated_paper"],
            file_name="WSA_Test_Paper.txt",
            use_container_width=True
        )


# ============================================================
# TAB 2 — EVALUATE ANSWERS (UPDATED WITH FILE UPLOADS)
# ============================================================

with tab2:
    st.subheader("📊 Evaluate Student Answers")

    default_paper = st.session_state.get("generated_paper", "")

    col_paper, col_answer = st.columns(2)

    with col_paper:
        st.markdown("### 1️⃣ Question Paper / Answer Key")
        paper_file = st.file_uploader(
            "Upload Question Paper / Answer Key (PDF, DOCX, TXT, Image)",
            type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
            key="paper_uploader"
        )
        question_paper_text = st.text_area(
            "Or Paste Question Paper / Answer Key Text Here",
            value=default_paper,
            height=250,
            placeholder="Paste text if you don't upload a file..."
        )

    with col_answer:
        st.markdown("### 2️⃣ Student Answer Sheet")
        student_file = st.file_uploader(
            "Upload Student Answer Sheet (PDF, DOCX, TXT, Image)",
            type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
            key="student_uploader"
        )
        student_answers_text = st.text_area(
            "Or Paste Student Answers Text Here",
            height=250,
            placeholder="Paste student answers if you don't upload a file..."
        )

    if st.button("📊 Evaluate Answers", use_container_width=True):
        if not api_key:
            st.error("⚠️ Please enter your Gemini API key in the sidebar.")
        else:
            # Process Paper Input
            paper_extracted_text, paper_part = process_uploaded_file(paper_file)
            final_paper_text = paper_extracted_text or question_paper_text.strip()

            # Process Answer Input
            answer_extracted_text, answer_part = process_uploaded_file(student_file)
            final_answer_text = answer_extracted_text or student_answers_text.strip()

            # Validation
            has_paper = bool(final_paper_text or paper_part)
            has_answer = bool(final_answer_text or answer_part)

            if not has_paper:
                st.error("⚠️ Please upload or paste the Question Paper / Answer Key.")
            elif not has_answer:
                st.error("⚠️ Please upload or paste the Student's Answers.")
            else:
                with st.spinner("Evaluating student answers using AI..."):
                    try:
                        eval_res = evaluate_student_answers(
                            api_key=api_key,
                            paper_text=final_paper_text,
                            paper_attachment=paper_part,
                            answers_text=final_answer_text,
                            answers_attachment=answer_part
                        )
                        st.session_state["evaluation"] = eval_res
                    except Exception as e:
                        st.error(f"❌ Evaluation Error: {e}")

    if "evaluation" in st.session_state:
        st.divider()
        st.subheader("📋 Evaluation Result")
        st.markdown(st.session_state["evaluation"])

        st.download_button(
            label="⬇️ Download Evaluation Report",
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

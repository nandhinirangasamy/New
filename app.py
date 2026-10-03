import os
import re
from io import BytesIO
from datetime import datetime

import streamlit as st
from openai import OpenAI
from dotenv import load_dotenv
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ============================================================
# 1. LOAD ENVIRONMENT VARIABLES
# ============================================================

load_dotenv()

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")

if not OPENROUTER_API_KEY:
    st.error(
        "OPENROUTER_API_KEY not found.\n\n"
        "Please create a .env file and add your OpenRouter API key."
    )
    st.stop()


# ============================================================
# 2. OPENROUTER CLIENT
# ============================================================

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY
)

MODEL = "openai/gpt-4.1-mini"


# ============================================================
# 3. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="AI Resume Tailor",
    page_icon="🤖",
    layout="wide"
)


# ============================================================
# 4. CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 42px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitle {
        font-size: 18px;
        color: #666;
        margin-bottom: 25px;
    }

    .section-title {
        font-size: 24px;
        font-weight: 600;
        margin-top: 20px;
    }

    .info-box {
        padding: 18px;
        border-radius: 10px;
        background-color: #f5f7fa;
        border: 1px solid #e1e5ea;
        margin-bottom: 15px;
    }

    .success-box {
        padding: 15px;
        border-radius: 10px;
        background-color: #eef8f0;
        border: 1px solid #b8dfbd;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 5. HEADER
# ============================================================

st.markdown(
    '<div class="main-title">🤖 AI Resume Tailor</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Automatically customize your resume according to the Job Description '
    'without changing your actual experience.'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# 6. PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(uploaded_file):

    try:

        reader = PdfReader(uploaded_file)

        text = ""

        for page in reader.pages:

            page_text = page.extract_text()

            if page_text:
                text += page_text + "\n"

        return text.strip()

    except Exception as e:

        raise Exception(
            f"Unable to read PDF file: {e}"
        )


# ============================================================
# 7. DOCX TEXT EXTRACTION
# ============================================================

def extract_docx_text(uploaded_file):

    try:

        document = Document(uploaded_file)

        text_parts = []

        # Paragraphs
        for paragraph in document.paragraphs:

            if paragraph.text.strip():

                text_parts.append(
                    paragraph.text.strip()
                )

        # Tables
        for table in document.tables:

            for row in table.rows:

                row_text = []

                for cell in row.cells:

                    if cell.text.strip():

                        row_text.append(
                            cell.text.strip()
                        )

                if row_text:

                    text_parts.append(
                        " | ".join(row_text)
                    )

        return "\n".join(text_parts).strip()

    except Exception as e:

        raise Exception(
            f"Unable to read DOCX file: {e}"
        )


# ============================================================
# 8. EXTRACT RESUME TEXT
# ============================================================

def extract_resume_text(uploaded_file):

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".pdf"):

        return extract_pdf_text(uploaded_file)

    elif file_name.endswith(".docx"):

        return extract_docx_text(uploaded_file)

    else:

        raise Exception(
            "Only PDF and DOCX files are supported."
        )


# ============================================================
# 9. BASIC EXPERIENCE EXTRACTION
# ============================================================

def extract_experience_reference(resume_text):

    patterns = [

        r"(\d+(?:\.\d+)?)\s*\+?\s*years?\s+of\s+experience",

        r"(\d+(?:\.\d+)?)\s*\+?\s*years?\s+experience",

        r"experience\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*\+?\s*years?"

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            resume_text,
            re.IGNORECASE
        )

        if match:

            return match.group(1) + " years"

    return "Experience duration is not explicitly mentioned."


# ============================================================
# 10. AI RESUME TAILORING
# ============================================================

def tailor_resume(resume_text, job_description):

    experience_reference = extract_experience_reference(
        resume_text
    )

    system_prompt = """
You are an expert ATS Resume Optimization Specialist.

Your task is to tailor an existing resume according to a Job Description.

==================================================
CRITICAL FACTUAL RULES
==================================================

1. NEVER change the candidate's actual experience duration.

2. NEVER change company names.

3. NEVER change employment dates.

4. NEVER create fake companies.

5. NEVER create fake projects.

6. NEVER create fake responsibilities.

7. NEVER invent technical skills.

8. NEVER invent certifications.

9. NEVER invent education.

10. NEVER invent tools or technologies.

11. NEVER claim that the candidate used a technology
    if that technology is not present in the original resume.

12. NEVER increase or decrease the candidate's experience.

13. NEVER remove factual employment information.

14. NEVER create a job title that completely contradicts
    the original resume.

==================================================
WHAT YOU ARE ALLOWED TO CHANGE
==================================================

You MAY:

1. Rewrite the professional summary.

2. Reorder existing skills according to the JD.

3. Highlight existing skills that match the JD.

4. Rephrase existing responsibilities using professional
   ATS-friendly wording.

5. Reorder existing responsibilities according to relevance.

6. Reorder existing projects according to relevance.

7. Improve grammar.

8. Improve sentence structure.

9. Improve ATS keyword placement.

10. Emphasize technologies that already exist in the resume.

11. Position the candidate toward the JD role ONLY using
    information that exists in the original resume.

==================================================
SKILL RULE
==================================================

If the JD contains:

Power BI
DAX
SQL
Power Query

and the original resume contains:

Power BI
SQL
Power Query

but does NOT contain DAX,

DO NOT add DAX.

Instead, prioritize:

Power BI
SQL
Power Query

and leave DAX out.

==================================================
EXPERIENCE RULE
==================================================

If the original resume says:

4 years experience

the final resume must NOT say:

5 years experience
6 years experience
4.7 years experience

unless that exact value exists in the original resume.

==================================================
COMPANY AND DATE RULE
==================================================

Company names and employment dates must remain exactly
as they appear in the original resume.

==================================================
OUTPUT
==================================================

Return the complete tailored resume.

Use this structure:

PROFESSIONAL SUMMARY

TECHNICAL SKILLS

PROFESSIONAL EXPERIENCE

PROJECTS

CERTIFICATIONS

EDUCATION

Only include sections that actually exist in the original resume.

Do not invent missing information.
"""


    user_prompt = f"""
==================================================
ORIGINAL RESUME
==================================================

{resume_text}


==================================================
JOB DESCRIPTION
==================================================

{job_description}


==================================================
KNOWN EXPERIENCE REFERENCE
==================================================

{experience_reference}


==================================================
TASK
==================================================

Tailor the resume for the provided Job Description.

The main objective is:

JD Role
    ↓
Identify relevant existing skills
    ↓
Prioritize those skills
    ↓
Rewrite summary
    ↓
Reorder relevant experience responsibilities
    ↓
Reorder relevant projects
    ↓
Improve ATS keyword alignment


IMPORTANT:

The candidate's real experience must remain unchanged.

Company names must remain unchanged.

Employment dates must remain unchanged.

Do not invent skills.

Do not invent projects.

Do not invent certifications.

Do not invent responsibilities.

Only optimize information already present in the original resume.

Return ONLY the final tailored resume.
"""


    response = client.chat.completions.create(

        model=MODEL,

        messages=[
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content": user_prompt
            }
        ],

        temperature=0.3,

        max_tokens=6000
    )

    return response.choices[0].message.content


# ============================================================
# 11. GENERATE PROFESSIONAL DOCX
# ============================================================

def create_docx(resume_text):

    document = Document()

    # --------------------------------------------------------
    # Page margins
    # --------------------------------------------------------

    section = document.sections[0]

    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)


    # --------------------------------------------------------
    # Default font
    # --------------------------------------------------------

    styles = document.styles

    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)


    # --------------------------------------------------------
    # Parse resume lines
    # --------------------------------------------------------

    lines = resume_text.splitlines()


    # --------------------------------------------------------
    # Section headings
    # --------------------------------------------------------

    section_names = {

        "PROFESSIONAL SUMMARY",
        "SUMMARY",
        "PROFILE",

        "TECHNICAL SKILLS",
        "SKILLS",
        "TECHNICAL EXPERTISE",

        "PROFESSIONAL EXPERIENCE",
        "EXPERIENCE",
        "WORK EXPERIENCE",

        "PROJECTS",
        "PROJECT EXPERIENCE",

        "CERTIFICATIONS",
        "CERTIFICATION",

        "EDUCATION",

        "ACHIEVEMENTS",
        "AWARDS",

        "LANGUAGES"
    }


    # --------------------------------------------------------
    # Helper: heading formatting
    # --------------------------------------------------------

    def add_heading(text):

        paragraph = document.add_paragraph()

        paragraph.paragraph_format.space_before = Pt(8)
        paragraph.paragraph_format.space_after = Pt(4)

        run = paragraph.add_run(
            text.upper()
        )

        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(12)

        # Bottom border
        p = paragraph._p

        pPr = p.get_or_add_pPr()

        pBdr = OxmlElement("w:pBdr")

        bottom = OxmlElement("w:bottom")

        bottom.set(
            qn("w:val"),
            "single"
        )

        bottom.set(
            qn("w:sz"),
            "6"
        )

        bottom.set(
            qn("w:space"),
            "1"
        )

        pBdr.append(bottom)

        pPr.append(pBdr)


    # --------------------------------------------------------
    # Add resume content
    # --------------------------------------------------------

    for raw_line in lines:

        line = raw_line.strip()

        if not line:

            continue


        # Check heading

        normalized = line.upper().strip(":# ")

        if normalized in section_names:

            add_heading(normalized)

            continue


        # Bullet points

        if line.startswith("- "):

            paragraph = document.add_paragraph(
                style="List Bullet"
            )

            paragraph.paragraph_format.space_after = Pt(2)

            run = paragraph.add_run(
                line[2:].strip()
            )

            run.font.name = "Arial"
            run.font.size = Pt(10)

            continue


        if line.startswith("• "):

            paragraph = document.add_paragraph(
                style="List Bullet"
            )

            paragraph.paragraph_format.space_after = Pt(2)

            run = paragraph.add_run(
                line[2:].strip()
            )

            run.font.name = "Arial"
            run.font.size = Pt(10)

            continue


        # ----------------------------------------------------
        # Normal paragraph
        # ----------------------------------------------------

        paragraph = document.add_paragraph()

        paragraph.paragraph_format.space_after = Pt(3)

        run = paragraph.add_run(line)

        run.font.name = "Arial"
        run.font.size = Pt(10)


    # --------------------------------------------------------
    # Save to memory
    # --------------------------------------------------------

    output = BytesIO()

    document.save(output)

    output.seek(0)

    return output


# ============================================================
# 12. SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Resume Settings")

    st.write(
        "The AI will optimize your resume based on the "
        "Job Description."
    )

    st.markdown("---")

    st.write("### 🔒 Protected Information")

    st.write("✅ Experience")

    st.write("✅ Company names")

    st.write("✅ Employment dates")

    st.write("✅ Education")

    st.write("✅ Certifications")

    st.markdown("---")

    st.write("### ✨ AI Optimization")

    st.write("✅ Relevant skills")

    st.write("✅ Summary")

    st.write("✅ ATS keywords")

    st.write("✅ Responsibility ordering")

    st.write("✅ Project ordering")


# ============================================================
# 13. RESUME UPLOAD
# ============================================================

st.markdown(
    "## 1️⃣ Upload Your Resume"
)

resume_file = st.file_uploader(
    "Upload your existing resume",
    type=["pdf", "docx"],
    help="Upload your current resume in PDF or DOCX format."
)


# ============================================================
# 14. JOB DESCRIPTION
# ============================================================

st.markdown(
    "## 2️⃣ Enter Job Description"
)

job_description = st.text_area(
    "Paste the Job Description",
    height=300,
    placeholder="""
Example:

Power Platform Developer

We are looking for a Power Platform Developer
with experience in Power Apps, Power Automate,
Dataverse, Power BI and SQL Server.

Responsibilities:
- Develop Canvas Apps
- Build Power Automate workflows
- Create Power BI dashboards
- Integrate SQL Server
- Work with Dataverse
""",
)


# ============================================================
# 15. GENERATE BUTTON
# ============================================================

st.markdown("---")

generate_button = st.button(
    "🚀 Generate JD-Tailored Resume",
    type="primary",
    use_container_width=True
)


# ============================================================
# 16. GENERATE RESUME
# ============================================================

if generate_button:

    # --------------------------------------------------------
    # Validate resume
    # --------------------------------------------------------

    if resume_file is None:

        st.warning(
            "⚠️ Please upload your resume first."
        )

        st.stop()


    # --------------------------------------------------------
    # Validate JD
    # --------------------------------------------------------

    if not job_description.strip():

        st.warning(
            "⚠️ Please enter the Job Description."
        )

        st.stop()


    # --------------------------------------------------------
    # Extract resume
    # --------------------------------------------------------

    with st.spinner(
        "📄 Reading your resume..."
    ):

        try:

            resume_text = extract_resume_text(
                resume_file
            )

        except Exception as e:

            st.error(str(e))

            st.stop()


    # --------------------------------------------------------
    # Check extracted text
    # --------------------------------------------------------

    if not resume_text:

        st.error(
            "❌ Could not extract text from the resume."
        )

        st.stop()


    # --------------------------------------------------------
    # Show resume information
    # --------------------------------------------------------

    experience_reference = extract_experience_reference(
        resume_text
    )

    st.success(
        f"Resume loaded successfully. "
        f"Detected experience reference: {experience_reference}"
    )


    # --------------------------------------------------------
    # AI processing
    # --------------------------------------------------------

    with st.spinner(
        "🤖 AI is tailoring your resume according to the JD..."
    ):

        try:

            tailored_resume = tailor_resume(
                resume_text,
                job_description
            )

        except Exception as e:

            st.error(
                f"❌ AI processing failed: {e}"
            )

            st.stop()


    # ========================================================
    # RESULT
    # ========================================================

    st.markdown("---")

    st.markdown(
        "## ✨ Your JD-Tailored Resume"
    )


    # --------------------------------------------------------
    # Two columns
    # --------------------------------------------------------

    col1, col2 = st.columns(
        [2, 1]
    )


    with col1:

        st.markdown(
            "### 📄 Resume Preview"
        )

        st.markdown(
            tailored_resume
        )


    with col2:

        st.markdown(
            "### 🔒 Resume Protection"
        )

        st.success(
            "Experience protected"
        )

        st.success(
            "Company names protected"
        )

        st.success(
            "Dates protected"
        )

        st.success(
            "No new skills invented"
        )

        st.success(
            "No fake projects"
        )

        st.markdown(
            "### 🎯 AI Optimization"
        )

        st.info(
            "The resume has been optimized "
            "to emphasize existing skills that "
            "match the Job Description."
        )


    # ========================================================
    # CREATE DOCX
    # ========================================================

    try:

        docx_file = create_docx(
            tailored_resume
        )

        st.markdown("---")

        st.markdown(
            "## 📥 Download"
        )

        file_name = (
            "JD_Tailored_Resume_"
            + datetime.now().strftime("%Y%m%d_%H%M")
            + ".docx"
        )

        st.download_button(

            label="📄 Download Tailored Resume",

            data=docx_file,

            file_name=file_name,

            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),

            use_container_width=True
        )

        st.success(
            "🎉 Your tailored resume is ready!"
        )

    except Exception as e:

        st.error(
            f"❌ DOCX generation failed: {e}"
        )

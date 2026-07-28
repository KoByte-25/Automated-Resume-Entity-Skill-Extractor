import re
import io
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.title("Automated Resume Entity & Skill Extractor")

# ---------- Session states ----------
if "messages" not in st.session_state:
    st.session_state.messages = []
if "name" not in st.session_state:
    st.session_state.name = None
if "DOB" not in st.session_state:
    st.session_state.DOB = None
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "step" not in st.session_state:
    # steps: "intro", "name", "dob", "done"
    st.session_state.step = "intro"

# ---------- Intro (only first time) ----------
if not st.session_state.messages:
    st.session_state.messages.append({
        "role": "assistant",
        "content": (
            "Hello! I am Automated Resume Entity & Skill Extractor.\n\n"
            "I will ask you a series of questions to build a resume automatically."
        )
    })
    st.session_state.step = "name"  # after intro, go to name question

# ---------- Helper methods ----------
def extract_name(text: str) -> str:
    if not text:
        return ""
    text = text.strip()

    patterns = [
        r"my name is\s+(.+)",   # My name is ...
        r"i am\s+(.+)",         # I am ...
        r"i'm\s+(.+)"           # I'm ...
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            name = m.group(1).strip(". ").strip()
            return name

    return text  # fallback: full text

def extract_dob(text: str) -> str:
    if not text:
        return ""
    text = text.strip()

    patterns = [
        r".*\sis\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})$"
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            day = m.group(1).strip()
            month = m.group(2).strip()
            year = m.group(3).strip()

            match month:
                case "01": mon = "January"
                case "02": mon = "February"
                case "03": mon = "March"
                case "04": mon = "April"
                case "05": mon = "May"
                case "06": mon = "June"
                case "07": mon = "July"
                case "08": mon = "August"
                case "09": mon = "September"
                case "10": mon = "October"
                case "11": mon = "November"
                case "12": mon = "December"
                case _: mon = "Unknown Month"

            return f"{day} {mon} {year}"

    return text  # fallback


# ---------- Render existing history ----------
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# ---------- Name step ----------
if st.session_state.step == "name" and st.session_state.name is None:
    # Ask only once
    if not any("please tell me your full name" in m["content"].lower()
               for m in st.session_state.messages):
        question = "First, please tell me your full name?"
        st.session_state.messages.append({
            "role": "assistant",
            "content": question
        })
        with st.chat_message("assistant"):
            st.write(question)

    user_input = st.chat_input("Type your name here...")

    if user_input:
        # Store user message
        st.session_state.messages.append({
            "role": "user",
            "content": user_input
        })

        # Process name
        name = extract_name(user_input)
        st.session_state.name = name

        with st.chat_message("user"):
            st.write(user_input)

        # Move to DOB step
        st.session_state.step = "dob"
        st.rerun()

# ---------- DOB step ----------
elif st.session_state.step == "dob" and st.session_state.DOB is None:
    if not any("please tell me your date of birth" in m["content"].lower()
               for m in st.session_state.messages):
        dobQuestion = "Now, can you tell me your date of birth (dd/mm/yyyy)?"
        st.session_state.messages.append({
            "role": "assistant",
            "content": dobQuestion
        })
        with st.chat_message("assistant"):
            st.write(dobQuestion)

    dobInput = st.chat_input("Type your date of birth here...")

    if dobInput:
        st.session_state.messages.append({
            "role": "user",
            "content": dobInput
        })

        dob = extract_dob(dobInput)
        st.session_state.DOB = dob

        with st.chat_message("user"):
            st.write(dobInput)

        # All basic info collected
        st.session_state.step = "done"
        st.rerun()

# ---------- Generate Word file ----------
if st.session_state.step == "done" and st.session_state.docx_bytes is None:
    with st.chat_message("assistant"):
        st.write(
            "Please wait, I am generating your resume Word file using your name..."
        )

    progress_bar = st.progress(0)
    for pct in range(0, 101, 25):
        progress_bar.progress(pct)

    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titleRun = title.add_run("Curriculum Vitae")
    titleRun.bold = True
    titleRun.underline = True
    titleRun.font.size = Pt(36)
    titleRun.font.name = "Times New Roman"

    nameP = doc.add_paragraph()
    nameRun = nameP.add_run(st.session_state.name)
    nameRun.bold = True
    nameRun.font.size = Pt(24)
    nameRun.font.name = "Times New Roman"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    st.session_state.docx_bytes = buffer.read()

    with st.chat_message("assistant"):
        st.write(
            "Your test resume Word file is ready. You can download it below."
        )

# ---------- Download button ----------
if st.session_state.docx_bytes is not None:
    st.download_button(
        label="Download CV",
        data=st.session_state.docx_bytes,
        file_name="cv.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
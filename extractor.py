import re
import io
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

st.title("Automated Resume Entity & Skill Extractor")

# ---------- Session states ----------
if "messages" not in st.session_state:
    st.session_state.messages = []
if "name" not in st.session_state:
    st.session_state.name = None
if "DOB" not in st.session_state:
    st.session_state.DOB = None
if "ADDR" not in st.session_state:
    st.session_state.ADDR = None
if "PHONE" not in st.session_state:
    st.session_state.PHONE = None
if "EMAIL" not in st.session_state:
    st.session_state.EMAIL = None
if "CAREER_OBJECTIVE" not in st.session_state:
    st.session_state.CAREER_OBJECTIVE = None
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
            "မင်္ဂလာပါ၊ ကျွန်တော်ကတော့ အလိုအလျောက် CV နဲ့ ကျွမ်းကျင်မှုတွေ ထုတ်ဖော်ရေးသားပေးသွားမှာပဲဖြစ်ပါတယ်။\n\n"
            "CV ရေးသားဖို့ ဘယ်ကနေ ဘယ်လိုစရေးရမှန်း မသိတဲ့ လူတွေအတွက် အကူအညီပေးဖို့ ရည်ရွယ်ပါတယ်။ \n CV ကို တည်ဆောက်ဖို့အတွက် ကျွန်တော်က မေးခွန်းအချို့ကို မေးမြန်းသွားမှာပဲ ဖြစ်ပါတယ်။"
        )
    })
    st.session_state.step = "name"  # after intro, go to name question

# ---------- Helper methods ----------
def insert_hr(paragraph):
    p = paragraph._p  # <w:p> element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(
        pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '10')      # thickness
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)

def extract_name(text: str) -> str:
    if not text:
        return ""
    text = text.strip()

    patterns = [
        r"my name is\s+(.+)",   # My name is ...
        r"i am\s+(.+)",         # I am ...
        r"i'm\s+(.+)",           # I'm ...
        r".*ကတော့\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*က\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$"
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            name = m.group(1).strip(". ").strip()
            return name

    return text.strip(". ")  # fallback: full text

def extract_dob(text: str) -> str:
    if not text:
        return ""
    text = text.strip(". ").strip()

    patterns = [
        r".*\sis\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})$",
        r"(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})$",
        r".*ကတော့\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*က\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$"

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

def extract_addr(text: str) -> str:
    if not text:
        return ""
    text = text.strip(". ").strip()

    patterns = [
        r".*\sis\s+(.+)\.$",
        r".*ကတော့\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*က\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*ကတော့\s+(.+)\s+မှာ.*",
        r".*က\s+(.+)\s+မှာ.*"
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            addr = m.group(1).strip()

            return addr

    return text  # fallback

def extract_phone(text: str) -> str:
    if not text:
        return ""
    text = text.strip(". ").strip()

    patterns = [
        r".*\sis\s+09([0-9]{9})\.$",
        r".*(ကတော့|က)\s+09([0-9]{9})\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            phone = m.group(2).strip()

            return f"+95 9{phone}"

    return text  # fallback

def extract_email(text: str) -> str:
    if not text:
        return ""
    text = text.strip(". ").strip()

    patterns = [
        r".*\sis\s+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\.$",
        r".*(ကတော့|က)\s+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            email = m.group(2).strip()

            return email

    return text  # fallback

def process_career_objective(text: str) -> str:
    if not text:
        return ""

    original = text.strip()
    lower = original.lower()

    # 1) Detect role / identity (very simple keyword-based)
    role = None
    if "student" in lower:
        role = "student"
    if "developer" in lower or "programmer" in lower or "engineer" in lower:
        role = "software developer"
    if "data" in lower and "analyst" in lower:
        role = "data analyst"

    # Default if nothing obvious
    if role is None:
        role = "aspiring professional"

    # 2) Detect goal phrase
    goal = None
    trigger_phrases = [
        "want to", "would like to", "aim to", "aiming to", "plan to",
        "my goal is", "my objective is", "seeking", "looking for",
        "hope to", "hoping to"
    ]
    for phrase in trigger_phrases:
        idx = lower.find(phrase)
        if idx != -1:
            # Take everything after the phrase as goal clause
            start = idx + len(phrase)
            goal = original[start:].strip(" .")
            break

    # 3) Detect skills / strengths (simple keywords)
    skills_keywords = [
        "android", "flutter", "php", "mysql", "javascript",
        "python", "nlp", "machine learning", "data science",
        "team", "leadership", "communication", "problem-solving"
    ]
    skills_found = [kw for kw in skills_keywords if kw in lower]

    # Build skill/value sentence
    value_sentence = ""
    if skills_found:
        # Make them readable (capitalize first letter)
        pretty_skills = ", ".join(s.title() for s in skills_found)
        value_sentence = (
            f" I bring skills in {pretty_skills} and am eager to contribute to real-world projects."
        )

    # 4) Compose final objective (2–3 sentences)
    if goal:
        objective = (
            f"As a {role}, I am seeking opportunities to {goal}. "
            f"My aim is to grow professionally while adding value to my organization.{value_sentence}"
        )
    else:
        # No explicit goal phrase; keep it generic
        objective = (
            f"As a {role}, I am seeking opportunities to develop my skills and gain practical experience. "
            f"I am motivated to learn, grow, and contribute to my future team.{value_sentence}"
        )

    return objective

# ---------- Render existing history ----------
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# ---------- Name step ----------
if st.session_state.step == "name" and st.session_state.name is None:
    # Ask only once
    question = "ပထမဆုံး သင့်ရဲ့ နာမည်အပြည့်အစုံကို အင်္ဂလိပ်လို ရေးပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == question
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": question
        })
        with st.chat_message("assistant"):
            st.write(question)

    user_input = st.chat_input("သင့်ရဲ့ နာမည်ကို ဒီမှာရေးပေးပါ...")

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
    dobQuestion = "အခု သင့်ရဲ့ မွေးသက္ကရာဇ်ကို (dd/mm/yyyy) ပုံစံဖြင့် အင်္ဂလိပ် ဂဏန်းများ ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == dobQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": dobQuestion
        })
        with st.chat_message("assistant"):
            st.write(dobQuestion)

    dobInput = st.chat_input("သင့်ရဲ့ မွေးသက္ကရာဇ်ကို ဒီမှာ ရေးပေးပါ...")

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
        st.session_state.step = "addr"
        st.rerun()

# ---------- Addr step ----------
elif st.session_state.step == "addr" and st.session_state.ADDR is None:
    addrQuestion = "အခု သင့်ရဲ့ နေရပ်လိပ်စာကို ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == addrQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": addrQuestion
        })
        with st.chat_message("assistant"):
            st.write(addrQuestion)

    addrInput = st.chat_input("သင့်ရဲ့ နေရပ်လိပ်စာကို ဒီမှာ ရေးပေးပါ...")

    if addrInput:
        st.session_state.messages.append({
            "role": "user",
            "content": addrInput
        })

        addr = extract_addr(addrInput)
        st.session_state.ADDR = addr

        with st.chat_message("user"):
            st.write(addrInput)

        # All basic info collected
        st.session_state.step = "phone"
        st.rerun()

elif st.session_state.step == "phone" and st.session_state.PHONE is None:
    phoneQuestion = "အခု သင့်ရဲ့ ဖုန်းနံပါတ်ကို ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == phoneQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": phoneQuestion
        })
        with st.chat_message("assistant"):
            st.write(phoneQuestion)

    phoneInput = st.chat_input("သင့်ရဲ့ ဖုန်းနံပါတ်ကို ဒီမှာ ရေးပေးပါ...")

    if phoneInput:
        st.session_state.messages.append({
            "role": "user",
            "content": phoneInput
        })

        phone = extract_phone(phoneInput)
        st.session_state.PHONE = phone

        with st.chat_message("user"):
            st.write(phoneInput)

        # All basic info collected
        st.session_state.step = "email"
        st.rerun()

elif st.session_state.step == "email" and st.session_state.EMAIL is None:
    emailQuestion = "အခု သင့်ရဲ့ အီးမေးလ်လိပ်စာကို ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == emailQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": emailQuestion
        })
        with st.chat_message("assistant"):
            st.write(emailQuestion)

    emailInput = st.chat_input("သင့်ရဲ့ အီးမေးလ်လိပ်စာကို ဒီမှာ ရေးပေးပါ...")

    if emailInput:
        st.session_state.messages.append({
            "role": "user",
            "content": emailInput
        })

        email = extract_email(emailInput)
        st.session_state.EMAIL = email

        with st.chat_message("user"):
            st.write(emailInput)

        # All basic info collected
        st.session_state.step = "done"
        st.rerun()

elif st.session_state.step == "career_objective" and st.session_state.CAREER_OBJECTIVE is None:
    careerQuestion = "အခု သင့်ရဲ့ အလုပ်ရည်ရွယ်ချက် ကို ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == careerQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": careerQuestion
        })
        with st.chat_message("assistant"):
            st.write(careerQuestion)

    careerInput = st.chat_input("သင့်ရဲ့ အလုပ်ရည်ရွယ်ချက်ကို ဒီမှာ ရေးပေးပါ...")

    if careerInput:
        st.session_state.messages.append({
            "role": "user",
            "content": careerInput
        })

        career = process_career_objective(careerInput)
        st.session_state.CAREER_OBJECTIVE = career

        with st.chat_message("user"):
            st.write(careerInput)

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

    dobP = doc.add_paragraph()
    dobBullet = dobP.add_run("\u2666")
    dobBullet.bold = True
    dobBullet.font.size = Pt(12)
    dobBullet.font.name = "Segoe UI Symbol"
    dobLable = dobP.add_run("Date of Birth\t: ")
    dobLable.bold = True
    dobLable.font.size = Pt(12)
    dobLable.font.name = "Times New Roman"
    dobValue = dobP.add_run(st.session_state.DOB)    
    dobValue.font.size = Pt(12)
    dobValue.font.name = "Times New Roman"

    addrP = doc.add_paragraph()
    addrBullet = addrP.add_run("\u2666")
    addrBullet.bold = True
    addrBullet.font.size = Pt(12)
    addrBullet.font.name = "Segoe UI Symbol"
    addrLable = addrP.add_run("Address\t\t: ")
    addrLable.bold = True
    addrLable.font.size = Pt(12)
    addrLable.font.name = "Times New Roman"
    addrValue = addrP.add_run(st.session_state.ADDR)    
    addrValue.font.size = Pt(12)
    addrValue.font.name = "Times New Roman"

    phoneP = doc.add_paragraph()
    phoneBullet = phoneP.add_run("\u2666")
    phoneBullet.bold = True
    phoneBullet.font.size = Pt(12)
    phoneBullet.font.name = "Segoe UI Symbol"
    phoneLable = phoneP.add_run("Phone\t\t: ")
    phoneLable.bold = True
    phoneLable.font.size = Pt(12)
    phoneLable.font.name = "Times New Roman"
    phoneValue = phoneP.add_run(st.session_state.PHONE)
    phoneValue.font.size = Pt(12)
    phoneValue.font.name = "Times New Roman"

    emailP = doc.add_paragraph()
    emailBullet = emailP.add_run("\u2666")
    emailBullet.bold = True
    emailBullet.font.size = Pt(12)
    emailBullet.font.name = "Segoe UI Symbol"
    emailLable = emailP.add_run("Email\t\t: ")
    emailLable.bold = True
    emailLable.font.size = Pt(12)
    emailLable.font.name = "Times New Roman"
    emailValue = emailP.add_run(st.session_state.EMAIL)
    emailValue.font.size = Pt(12)
    emailValue.font.name = "Times New Roman"

    hr1P = doc.add_paragraph()
    insert_hr(hr1P)

    cbP = doc.add_paragraph()
    cbRun = cbP.add_run("Career Objective")
    cbRun.bold = True
    cbRun.underline = True
    cbRun.font.size = Pt(16)
    cbRun.font.name = "Times New Roman"

    cbValueP = doc.add_paragraph()
    cbValueRun = cbValueP.add_run(st.session_state.CAREER_OBJECTIVE)
    cbValueRun.font.size = Pt(12)
    cbValueRun.font.name = "Times New Roman"

    hr2P = doc.add_paragraph()
    insert_hr(hr2P)

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
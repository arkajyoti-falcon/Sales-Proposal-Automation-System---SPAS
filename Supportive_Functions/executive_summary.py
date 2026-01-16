import os
import re
from io import BytesIO
from datetime import date

import streamlit as st
from groq import Groq
import pdfplumber
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from dotenv import load_dotenv
load_dotenv()
# ---------------------- CONFIG ---------------------- #

PROMPT_EXEC_SUMMARY = """
You are a senior sales consultant at Falcon Autotech writing executive summaries that SELL.

## 🎯 SALES-FIRST MINDSET (CRITICAL)
- This is NOT a technical document—it's a SALES pitch in summary form
- Every bullet point should make the client think: "This solves MY problem"
- Lead with CLIENT BENEFITS, not Falcon's features
- Create urgency and excitement—make them WANT this solution

## 💡 WHY + WHAT (Always answer WHY)
- DON'T: "Loop Cross-Belt Sorter with 40 destinations"
- DO: "Loop Cross-Belt Sorter with 40 destinations—enabling you to sort 4,000+ items/hour with 99.9% accuracy"
- DON'T: "Automated induction system"
- DO: "Automated induction that reduces manual handling by 70% and prevents bottlenecks during peak hours"

## 🗣️ HUMAN STORYTELLING STYLE
- Write like a trusted advisor explaining the solution over coffee, NOT an AI generating text
- Use simple, warm English—avoid stiff corporate language
- Vary your sentences: short punchy ones for impact, longer flowing ones for explanation
- Sound confident but not arrogant; helpful but not salesy
- AVOID: "We are pleased to present..." "It is our privilege..." "The aforementioned solution..."
- INSTEAD USE: "We've designed this specifically for..." "Here's what makes this work for you..." "Your team will see immediate improvements in..."

## Falcon Autotech Context
Falcon designs, manufactures, supplies, implements, and maintains warehouse automation solutions—such as sortation systems, conveyor automation, pick/put-to-light, ASRS robotics, and dimension & weight scanning—for industries including e-commerce, fashion, FMCG, pharma, groceries, and CE-P.

Your task is to generate **unique, client-tailored Executive Summaries** based on the “Proposed System Description” section of Falcon proposals.  
The summary must always reflect Falcon’s style but **no two summaries should ever be identical**. Introduce subtle variations in wording, phrasing, and sentence structure while keeping the same professional tone.  

### Writing Rules

**Opening Section**
- Begin with Falcon Autotech’s commitment and strong interest in responding to the client’s requirement.  
- Mention Falcon’s partnership approach, customization, and proven track record.  
- Use varied sentence structures and synonyms so every generation feels different.  

**Bullet Points**
- Provide exactly **4–5 high-level system features or modules**.  
- Each bullet MUST be short, clear, and client-friendly (e.g., “Spiral Conveyors for smooth material flow”).  
- Avoid technical specifications, sub-bullets, or repeating the same idea in different words.  
- The order of bullets should vary slightly between generations.  
- Add numeric along with the components ONLY IF extensively mentioned in Proposed System Description
- Bold the main components of the system. There can be max 2-3 bold words.

**Closing Section**
- End with a **personalized closing statement**.  
- Reaffirm that the solution is tailored to meet the client’s technical and operational requirements.  
- Mention the RFP/customization and highlight benefits like efficiency, smooth material flow, and faster TAT.  
- Closing phrasing should change between runs (use variations in tone, sentence structure, and emphasis).  

### Important Constraints
- Keep the tone formal, professional, and benefit-driven.  
- Do **not** reuse exact sentences from earlier examples.  
- Ensure variability: two runs for the same input must never produce identical text.  
- Do **not** add any extra sections outside the defined structure.  

### Output Format
1. Opening paragraph (commitment + partnership).  
2. 4–6 bullet points (system modules).  
3. Closing personalized statement.  


**DO NOT ADD ANY EXTRA TEXT OR INFORMATION OR JUSTIFICATION or "Here is an Executive Summary for the proposal:" EXCEPT THE FULL PROPOSAL**

**Eaxmple 1:**

Falcon Autotech is pleased to confirm its great interest in responding to this Dnata RFQ of Conveyor Automation for 
Dubai Location. Our team has been working closely with the relevant stakeholders, with a clear commitment to 
listening, understanding your needs, and ensuring this project's success.
Following the same objective for the Dubai CBS system, we are happy to offer a compliant solution meeting all 
technical and operational requirements at competitive price, delivering key results.
Our solution is based on the following key characteristics:
    • Conveyor Automation for Handling Shipments of Boxes 
    • 3D ASRS (NEO) System for Storage of the shipments
    • Spiral Conveyors for smooth material flow
    • Outbound sorter with Swivel Wheel Divert Units for faster TAT and TPH
    • ULD Handling System (Castor Deck)
    • Oversize Cargo Pallet Conveyor Automation

This solution has been crafted especially keeping Dnata’s technical and operational requirements, as listed in the RFP 
document, making it a tailor-made solution delivering a faster TAT and an efficient material flow.


**Example 2 :**
Falcon is pleased to confirm its great interest in responding to this Mark3 International
Requirement of SWEDI Sorter. Our team has been working closely with the relevant 
stakeholders, with a clear commitment to listening and understanding your needs and 
ensuring this project's success.
Following the same objective for the Iran system, we are happy to offer a compliant solution 
meeting all technical and operational requirements, high-performance, optimized, tailormade, fast and secure planning, and a competitive price.
Our solution is based on the following key characteristics:

    • The SWEDI sorter, based on a Swivel wheel technology, offers a designed 
    throughput of 2,400 pph.
    • A Loading conveyor system with provision of manual Shipment loading.
    • 10 Output Chute & 1 Rejection/Sort Fail Chute

A tailor-made and simple layout designed explicitly for Mark3 International. The proposed 
layout is the result of the technical requirements in the RFP document and our discussions 
with the relevant stakeholders during our meeting.

**Example 3:**
Falcon is pleased to confirm its great interest in responding to this RFQ. Our team has been 
working closely with the relevant stakeholders, with a clear commitment to listening and 
understanding your needs and ensuring this project's success.
As prime contractor, Falcon ensures its full commitment to successfully completing this 
project. Following the same objective for the system, we are happy to offer a compliant 
solution meeting all technical and operational requirements, high-performance, optimized, 
tailor-made, fast and secure planning, and a competitive price. For full transparency, a 
compliancy list forms part of our offer. 
Our solution is based on the following key characteristics:

    • One Sorter, based on Falcon’s own, well-known loop cross-belt technology
    • 7 Sets of Induct with Manual Loading 
    • 1 Sets of Infeed with Volume Distribution System (VDS)
    • 40 Pcs Pallet Chute
    • 26 Pcs PTL Chutes for 312 Bags
    • 312 Pcs Put to Light with Rack for Bags
    • 3 Pcs Rejection Chute 
    • 7 Pcs of VDS Chute 
    • 7 Pcs of NC Chute

A tailor-made and simple layout, specifically designed to VINTED The proposed layout, is 
the result of the technical requirements in the RFP document and our discussions with the 
relevant stakeholders during the feedback round.
- Simple operational conditions due to one double decker horizontal cross belt sorter.
- Easy maintenance: optimized number of conveyors and concentrated inducts area.


DO NOT ADD ANY EXTRA WORD OR INFO APART FROM THE EXECUTIVE SUMMARY
"""

MODEL_NAME = "llama-3.3-70b-versatile"  # change if you prefer another Groq model

# ---------------------- HELPERS ---------------------- #


def extract_pdf_text(uploaded_file) -> str:
    """Extract plain text from an uploaded PDF using pdfplumber."""
    if uploaded_file is None:
        return ""

    text_chunks = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_chunks.append(page_text)

    full_text = "\n\n".join(text_chunks)
    # Hard truncate to keep prompt size reasonable
    if len(full_text) > 20000:
        full_text = full_text[:20000]
    return full_text


def call_groq_exec_summary(system_text: str, client_name: str, project_title: str) -> str:
    """Call Groq API to generate the Executive Summary text."""
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY environment variable is not set.")

    client = Groq(api_key=api_key)

    user_content = (
        f"Client Name: {client_name}\n"
        f"Project / System Name: {project_title}\n\n"
        f"Proposed System Description (for context):\n{system_text}\n\n"
        "Generate the Executive Summary strictly as per the instructions."
    )

    resp = client.chat.completions.create(
        model=MODEL_NAME,
        temperature=0.4,
        max_tokens=800,
        messages=[
            {"role": "system", "content": PROMPT_EXEC_SUMMARY},
            {"role": "user", "content": user_content},
        ],
    )

    return resp.choices[0].message.content.strip()


def add_markdown_paragraph(doc: Document, text: str, style: str | None = None):
    """
    Add a paragraph with simple **bold** Markdown handling.
    """
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()

    parts = re.split(r"(\*\*[^\*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = p.add_run(part[2:-2])
            run.bold = True
        else:
            p.add_run(part)
    return p


def build_exec_summary_docx(exec_text: str) -> BytesIO:
    """Build DOCX styled like your sample and append the fixed section."""
    doc = Document()

    # Set base font to Calibri 11 for Normal style
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal_style.font.size = Pt(11)

    # "Executive Summary" heading – bold + underline, left aligned
    heading_p = doc.add_paragraph()
    heading_run = heading_p.add_run("Executive Summary")
    heading_run.bold = True
    heading_run.underline = True

    doc.add_paragraph("")  # spacing

    # Parse Groq text: paragraphs and bullets
    lines = [ln.strip() for ln in exec_text.splitlines() if ln.strip()]

    for line in lines:
        stripped = line.lstrip()
        # Bullet detection (•, -, *)
        if stripped.startswith("•"):
            content = stripped.lstrip("•").strip()
            add_markdown_paragraph(doc, content, style="List Bullet")
        elif stripped.startswith("- "):
            content = stripped[2:].strip()
            add_markdown_paragraph(doc, content, style="List Bullet")
        elif stripped.startswith("* "):
            content = stripped[2:].strip()
            add_markdown_paragraph(doc, content, style="List Bullet")
        else:
            add_markdown_paragraph(doc, stripped)

    # Gap before fixed section
    doc.add_paragraph("")

    # -------- Fixed Section 1 -------- #
    p1 = doc.add_paragraph()
    r1 = p1.add_run("1. Falcon's reliable Shipment sortation systems")
    r1.bold = True

    doc.add_paragraph(
        "Falcon Autotech has over 50+ Cross belt sorter loops installed globally which are used by most "
        "innovative brands such as Flipkart, Amazon, Delhivery, Aramex, Asendia, Fastway and many more. The "
        "main and critical components of the Falcon Autotech sorter building blocks, like wheels, motors, belts, "
        "bearings, Bus Bars, Communication platforms, PLCs etc., are sourced from industry leading suppliers "
        "in the world, such as SEW, Siemens, SICK, Vahle, Forbo etc. This strategic baseline of sourcing policy "
        "allows Falcon's customers to be fully confident in the systems' robustness and reliability."
    )

    doc.add_paragraph("")  # small gap

    # -------- Fixed Section 2 -------- #
    p2 = doc.add_paragraph()
    r2 = p2.add_run("2. Commitment to quality systems")
    r2.bold = True

    doc.add_paragraph(
        "Demonstrating Falcon's clear commitment to the CLIENT satisfaction, the sortation system, parts and "
        "services will be under warranty for 24 months from installation go-live.  "
        "Our expertise in executing world-class, projects across the MENA region has been thoroughly "
        "demonstrated and validated. We offer:"
    )

    bullets_fixed = [
        "A proven history of success and broad experience in implementing and maintaining automated systems.",
        "ISO 9001 Quality Assurance certification, ensuring consistent quality across our products, "
        "processes, and outcomes.",
        "Falcon Warehouse Control System (WCS) is designed to meet the real-time operational needs of "
        "automated materials handling systems. WCS can be molded based on the client’s requirements "
        "to enhance operations.",
        "Comprehensive service and warranty programs. We deliver both operational and emergency support "
        "with unmatched responsiveness and availability throughout the region.",
        "Strong financial stability, backed by a long-standing track record.",
    ]

    for item in bullets_fixed:
        doc.add_paragraph(item, style="List Bullet")

    # Save to memory buffer
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------------------- STREAMLIT APP ---------------------- #

st.set_page_config(page_title="Executive Summary Generator", page_icon="📝")

st.title("Executive Summary – Falcon Proposal Section")

st.markdown(
    "Upload the **Proposed System Description (PDF)** and fill the details below. "
    "The app will call **Groq** to generate a human-quality Executive Summary and "
    "return a ready-to-insert **DOCX** section in Falcon style."
)

with st.form("exec_summary_form"):
    pdf_file = st.file_uploader(
        "Upload Proposed System Description (PDF)", type=["pdf"]
    )

    client_name = st.text_input("Client Name (for personalization)", "")
    project_title = st.text_input(
        "Project / System Name (e.g., 'Loop CBS 6K PPH with PTL')", ""
    )

    letter_date = st.date_input("Letter / Proposal Date", value=date.today())

    submitted = st.form_submit_button("Generate Executive Summary DOCX")

docx_buffer = None

if submitted:
    if not pdf_file:
        st.error("Please upload the Proposed System Description PDF.")
    elif not client_name or not project_title:
        st.error("Please enter both Client Name and Project / System Name.")
    else:
        with st.spinner("Extracting PDF and generating Executive Summary via Groq..."):
            try:
                pdf_text = extract_pdf_text(pdf_file)

                # Append date context into system text (light hint, not mandatory)
                context_text = (
                    f"{pdf_text}\n\n"
                    f"(Proposal date: {letter_date.strftime('%d-%m-%Y')})"
                )

                exec_summary_text = call_groq_exec_summary(
                    context_text, client_name, project_title
                )

                docx_buffer = build_exec_summary_docx(exec_summary_text)

            except Exception as e:
                st.error(f"Error while generating Executive Summary: {e}")

if docx_buffer:
    safe_client = client_name.strip() or "Client"
    st.success("Executive Summary generated successfully.")
    st.download_button(
        label="Download Executive Summary (.docx)",
        data=docx_buffer,
        file_name=f"Executive_Summary_{safe_client}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

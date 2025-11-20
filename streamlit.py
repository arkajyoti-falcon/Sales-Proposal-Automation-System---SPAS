import os
import streamlit as st
import pandas as pd
import json
import re
from io import BytesIO
from datetime import date
from dataclasses import dataclass
from typing import Dict, List
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
from groq import Groq
from dotenv import load_dotenv
import pdfplumber

load_dotenv()

# ==================== GROQ PROMPTS & CONSTANTS ====================

# Cover Letter System Prompt
COVER_LETTER_SYSTEM_PROMPT = """
You are an AI assistant working as a professional proposal writer at Falcon Autotech. You are an expert in drafting formal, client-specific techno-commercial cover letters for proposals. Your role is to generate well-structured, personalized cover letters that follow Falcon's business communication style, maintain a professional and respectful tone, and clearly demonstrate Falcon's commitment, expertise, and partnership approach to clients.

Generate a formal techno-commercial COVER LETTER for a proposal. 
The writing style MUST be indistinguishable from natural human writing. The text should read as if drafted by an experienced professional, not an AI system. Use clear, simple, and natural language with varied sentence lengths and structures. Avoid generic phrases, repetitive patterns, or mechanical tone. Ensure that the output flows smoothly, conveys intent naturally, and would not be detected as machine-generated. The content should feel thoughtful, context-aware, and aligned with how a human proposal writer or business professional would communicate.

MAX COVER LETTER WORDS : 250 WORDS OR 1500 CHARACTER (whatever is minimum)

1. Start with:
   Kind Attention –
   Mr. {{executives}}
   M/s {{client_name}}

   Offer Ref: {{offer_ref}}; Date: {{letter_date}}

   Subject – Techno-Commercial Offer for {{project_title}}  

2. If there is only one executive, address them with:
   Dear {{first_exec_name}},
   If multiple executives, skip "Dear" and go directly to the content.
   Use Mr. for male and Ms. for female executives.

3. Opening paragraph (human way):
   - Acknowledge the invitation or requirement.
   - If invitation_date exists, mention it naturally.
   - If meeting_date exists, reference recent discussions or suggestions.
   - Wording must change between runs (not fixed sentences).

4. Body (human way):
   - Highlight Falcon's analysis, solution evaluation, and technical proposal attachment.
   - Mention Falcon's proven intralogistics technologies and experience.
   - Personalize with client_name.
   - Optionally mention project planning or timeline.

5. Closing (human way):
   - Reaffirm sender's personal commitment.
   - Encourage the client to reach out for clarifications.
   - End with "Best Regards," followed by sender_name and sender_title.

Important:
- Do not exceed word/character limit.
- Keep tone formal, professional, and client-oriented.
- Do not copy exact sentences; rephrase wording across generations.
- The cover letter MUST sound human, natural and professional. It should be clear, authentic, and warm, without feeling robotic or overly formal.
- DO NOT ADD ANY EXTRA WORD OR INFO APART FROM THE COVER LETTER.
- Highlight the main system or project name in main body (not subject line) as bold style, use ** for Bold.
"""

COVER_LETTER_USER_PROMPT_TEMPLATE = """
Use the following information to generate the cover letter:

client_name: {client_name}
project_title: {project_title}
offer_ref: {offer_ref}
letter_date: {letter_date}

executives (one per line, already with Mr./Ms. prefix):
{executives_block}

invitation_date: {invitation_date}
meeting_date: {meeting_date}

sender_name: {sender_name}
sender_title: {sender_title}

Return ONLY the cover letter text, without markdown code fences or extra commentary.
"""

# Executive Summary System Prompt
EXEC_SUMMARY_SYSTEM_PROMPT = """
You are a Proposal Writing Assistant specialized in Falcon Autotech automation projects.  
Falcon Autotech designs, manufactures, supplies, implements, and maintains warehouse automation solutions—such as sortation systems, conveyor automation, pick/put-to-light, ASRS robotics, and dimension & weight scanning—for industries including e-commerce, fashion, FMCG, pharma, groceries, and CE-P.  
The writing style must be indistinguishable from natural human writing. The text should read as if drafted by an experienced professional, not an AI system. Use clear, simple, and natural language with varied sentence lengths and structures. Avoid generic phrases, repetitive patterns, or mechanical tone. Ensure that the output flows smoothly, conveys intent naturally, and would not be detected as machine-generated. The content should feel thoughtful, context-aware, and aligned with how a human proposal writer or business professional would communicate.

Your task is to generate **unique, client-tailored Executive Summaries** based on the "Proposed System Description" section of Falcon proposals.  
The summary must always reflect Falcon's style but **no two summaries should ever be identical**. Introduce subtle variations in wording, phrasing, and sentence structure while keeping the same professional tone.  

### Writing Rules

**Opening Section**
- Begin with Falcon Autotech's commitment and strong interest in responding to the client's requirement.  
- Mention Falcon's partnership approach, customization, and proven track record.  
- Use varied sentence structures and synonyms so every generation feels different.  

**Bullet Points**
- Provide exactly **4–5 high-level system features or modules**.  
- Each bullet MUST be short, clear, and client-friendly (e.g., "Spiral Conveyors for smooth material flow").  
- Avoid technical specifications, sub-bullets, or repeating the same idea in different words.  
- The order of bullets should vary slightly between generations.  
- Add numeric along with the components ONLY IF extensively mentioned in Proposed System Description
- Bold the main components of the system. There can be max 2-3 bold words.

**Closing Section**
- End with a **personalized closing statement**.  
- Reaffirm that the solution is tailored to meet the client's technical and operational requirements.  
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

DO NOT ADD ANY EXTRA TEXT OR INFORMATION OR JUSTIFICATION or "Here is an Executive Summary for the proposal:" EXCEPT THE FULL PROPOSAL
"""

# Config paths
STATIC_ABOUT_DIR = r"Static_AboutCompany"

# Handled Shipment Spectrum Templates
@dataclass
class SorterTemplate:
    key: str
    label: str
    keywords: List[str]
    config_name: str
    item_singular: str
    subheading_51: str
    spec_table: Dict[str, Dict[str, str]]

SORTER_TEMPLATES: List[SorterTemplate] = [
    SorterTemplate(
        key="linear_dual_standard",
        label="Linear / Dual-belt CBS – standard boxes",
        keywords=["linear", "6k", "5.4k", "loop cbs + linear", "totes", "boxes"],
        config_name="Linear Cross Belt Sorter (Dual-belt configuration)",
        item_singular="shipment",
        subheading_51="Shipment size loadable on the sorter",
        spec_table={
            "Max Length": {"unit": "mm", "value": "600"},
            "Max Width":  {"unit": "mm", "value": "450"},
            "Max Height": {"unit": "mm", "value": "400"},
            "Max Weight": {"unit": "Kg", "value": "20"},
            "Min length": {"unit": "mm", "value": "100"},
            "Min Width":  {"unit": "mm", "value": "100"},
            "Min Height": {"unit": "mm", "value": "3"},
            "Min Weight": {"unit": "gm", "value": "50"},
        },
    ),
    SorterTemplate(
        key="loop_standard",
        label="Loop CBS – standard shipments",
        keywords=["loop", "double deck", "48k", "loop cbs", "main sorter"],
        config_name="Loop Cross Belt Sorter technology",
        item_singular="shipment",
        subheading_51="Shipment size loadable on the sorter",
        spec_table={
            "Max Length": {"unit": "mm", "value": "400"},
            "Max Width":  {"unit": "mm", "value": "400"},
            "Max Height": {"unit": "mm", "value": "400"},
            "Max Weight": {"unit": "Kg", "value": "40"},
            "Min length": {"unit": "mm", "value": "10"},
            "Min Width":  {"unit": "mm", "value": "100"},
            "Min Height": {"unit": "mm", "value": "50"},
            "Min Weight": {"unit": "gm", "value": "100"},
        },
    ),
    SorterTemplate(
        key="heavy_parcel",
        label="Heavy-duty CBS – parcels / bags & boxes",
        keywords=["parcel", "heavy", "bags", "bag and box", "bosta", "delhivery"],
        config_name="Heavy Duty Cross Belt Sorter",
        item_singular="parcel",
        subheading_51="Parcel size loadable on the sorter",
        spec_table={
            "Max Length": {"unit": "mm", "value": "1000"},
            "Max Width":  {"unit": "mm", "value": "800"},
            "Max Height": {"unit": "mm", "value": "800"},
            "Max Weight": {"unit": "Kg", "value": "50"},
            "Min length": {"unit": "mm", "value": "40"},
            "Min Width":  {"unit": "mm", "value": "150"},
            "Min Height": {"unit": "mm", "value": "150"},
            "Min Weight": {"unit": "Kg", "value": "0.05"},
        },
    ),
]

st.set_page_config(page_title="Falcon Proposal Generator", page_icon="📄", layout="centered")

# Custom CSS for professional look
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f3864;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.3rem;
        font-weight: 600;
        color: #1f3864;
        margin-top: 2rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #1f3864;
        padding-bottom: 0.5rem;
    }
    .stTextInput > label, .stFileUploader > label, .stTextArea > label, .stDateInput > label, .stCheckbox > label, .stSelectbox > label {
        font-weight: 500;
        color: #333;
    }
    .info-box {
        background-color: #f0f2f6;
        border-left: 4px solid #1f3864;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">Falcon Proposal Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Professional Proposal Document Generation System</div>', unsafe_allow_html=True)

st.markdown("---")

# ==================== STYLING FUNCTIONS ====================

# Deep Blue-Gray color for headings (RGB: 31, 56, 100)
HEADING_COLOR = RGBColor(31, 56, 100)

def apply_heading_style(paragraph, text, level=1):
    """Apply custom heading style: Calibri Headings 14pt, Bold, Underline, Numbered, Deep Blue-Gray"""
    paragraph.text = ""
    run = paragraph.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.underline = True
    run.font.color.rgb = HEADING_COLOR
    
    # Apply paragraph formatting
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(6)
    
    return paragraph

def apply_subheading_style(paragraph, text):
    """Apply subheading style: Calibri 12pt, Bold, Deep Blue-Gray"""
    paragraph.text = ""
    run = paragraph.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(3)
    
    return paragraph

def apply_normal_style(paragraph, text=""):
    """Apply normal text style: Calibri (Body) 11pt, Black"""
    if text:
        paragraph.text = ""
        run = paragraph.add_run(text)
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)
    else:
        for run in paragraph.runs:
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
    
    return paragraph

def apply_table_style(table):
    """Apply Medium Shading 1 Accent 1 style to table"""
    table.style = 'Medium Shading 1 Accent 1'
    return table

def add_centered_image(doc, path, width_in=5.5):
    """Add a centered image if it exists"""
    if not path or not os.path.exists(path):
        return
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_picture(path, width=Inches(width_in))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

def add_numbered_heading(doc, text, level=1, counter=None):
    """Add a numbered heading with proper formatting"""
    if counter:
        full_text = f"{counter}. {text}"  # Added period after number
    else:
        full_text = text
    
    # Use built-in heading style
    p = doc.add_heading(full_text, level=1)
    
    # Apply custom formatting to the heading
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.underline = True
        run.font.color.rgb = HEADING_COLOR
    
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    
    return p

def add_numbered_subheading(doc, text, counter=None):
    """Add a numbered subheading"""
    if counter:
        full_text = f"{counter}. {text}"  # Added period after number
    else:
        full_text = text
    
    # Use built-in heading style for subheading
    p = doc.add_heading(full_text, level=2)
    
    # Apply custom formatting
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = HEADING_COLOR
    
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    
    return p

def create_header_footer(doc, client_name, project_name, falcon_logo_path, client_logo_path):
    """Create header and footer for the document"""
    
    # Access the default section
    section = doc.sections[0]
    
    # Set page margins for better layout
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    
    # ==================== HEADER ====================
    header = section.header
    header_table = header.add_table(rows=1, cols=3, width=Inches(6.5))
    header_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Left cell - Client Logo
    left_cell = header_table.rows[0].cells[0]
    left_cell.width = Inches(1.3)
    left_cell.vertical_alignment = 1  # Center vertically
    if client_logo_path and os.path.exists(client_logo_path):
        left_para = left_cell.paragraphs[0]
        left_run = left_para.add_run()
        left_run.add_picture(client_logo_path, height=Inches(0.6))  # Fixed height for uniformity
        left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Middle cell - Header Text
    middle_cell = header_table.rows[0].cells[1]
    middle_cell.width = Inches(4.0)
    middle_cell.vertical_alignment = 1  # Center vertically
    middle_para = middle_cell.paragraphs[0]
    middle_run = middle_para.add_run(f"FALCON's Proposal to {client_name} for the {project_name}")
    middle_run.font.name = 'Calibri'
    middle_run.font.size = Pt(9)
    middle_run.font.bold = False
    middle_run.font.color.rgb = HEADING_COLOR
    middle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Right cell - Falcon Logo (use fixed path)
    right_cell = header_table.rows[0].cells[2]
    right_cell.width = Inches(1.3)
    right_cell.vertical_alignment = 1  # Center vertically
    
    # Use fixed Falcon logo path
    fixed_falcon_logo = "FIXED_IMAGE\\Falcon-Autotech_Logo-removebg-preview.png"
    
    # Try fixed path first, then uploaded logo
    falcon_logo_to_use = None
    if os.path.exists(fixed_falcon_logo):
        falcon_logo_to_use = fixed_falcon_logo
    elif falcon_logo_path and os.path.exists(falcon_logo_path):
        falcon_logo_to_use = falcon_logo_path
    
    if falcon_logo_to_use:
        right_para = right_cell.paragraphs[0]
        right_run = right_para.add_run()
        right_run.add_picture(falcon_logo_to_use, height=Inches(0.6))  # Fixed height for uniformity
        right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Remove borders from header table
    for row in header_table.rows:
        for cell in row.cells:
            tc = cell._element
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'none')
                tcBorders.append(border)
            tcPr.append(tcBorders)
    
    # Add horizontal line after header
    header_line = header.add_paragraph()
    header_line_run = header_line.add_run()
    header_line.paragraph_format.space_before = Pt(3)
    
    # ==================== FOOTER ====================
    footer = section.footer
    # Add horizontal line before footer
    footer_line = footer.add_paragraph()
    footer_line_run = footer_line.add_run()
    footer_line.paragraph_format.space_after = Pt(3)

    # Footer as a single line: copyright, clickable link, page X of Y
    para = footer.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run("© FALCON AUTOTECH 2025 Confidential: Not for Distribution. ")
    run.font.name = 'Calibri (Body)'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0, 0, 0)
    add_hyperlink(para, "https://www.falconautotech.com/", "https://www.falconautotech.com/")
    run2 = para.add_run(" | Page ")
    run2.font.name = 'Calibri (Body)'
    run2.font.size = Pt(9)
    run2.font.color.rgb = RGBColor(0, 0, 0)
    # Add page number field
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run2._r.append(fldChar1)
    run2._r.append(instrText)
    run2._r.append(fldChar2)
    run3 = para.add_run(" of ")
    run3.font.name = 'Calibri (Body)'
    run3.font.size = Pt(9)
    run3.font.color.rgb = RGBColor(0, 0, 0)
    # Add total pages field
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'begin')
    instrText2 = OxmlElement('w:instrText')
    instrText2.set(qn('xml:space'), 'preserve')
    instrText2.text = 'NUMPAGES'
    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')
    run3._r.append(fldChar3)
    run3._r.append(instrText2)
    run3._r.append(fldChar4)

def add_hyperlink(paragraph, url, text):
    """Add a hyperlink to a paragraph"""
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create a new run object (a wrapper over a <w:r> element)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Add formatting for hyperlink (blue + underline)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')  # Blue color
    rPr.append(color)
    
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    
    # Set font
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Calibri (Body)')
    rPr.append(rFonts)
    
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '18')  # 9pt = 18 half-points
    rPr.append(sz)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink

# ==================== ADDITIONAL HELPER FUNCTIONS ====================

def extract_pdf_text(uploaded_file) -> str:
    """Extract plain text from an uploaded PDF using pdfplumber."""
    if uploaded_file is None:
        return ""

    text_chunks = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text_chunks.append(page.extract_text() or "")

    full_text = "\n\n".join(text_chunks)
    # Hard truncate to keep prompt size reasonable
    if len(full_text) > 20000:
        full_text = full_text[:20000]
    return full_text

def choose_sorter_template(project_name: str) -> SorterTemplate:
    """Pick the closest template based on project name keywords."""
    text = (project_name or "").lower()

    # score by number of keyword hits
    best_tpl = SORTER_TEMPLATES[0]
    best_score = -1
    for tpl in SORTER_TEMPLATES:
        score = sum(1 for kw in tpl.keywords if kw in text)
        if score > best_score:
            best_score = score
            best_tpl = tpl

    return best_tpl

def shade_cell(cell, color_hex: str = "D9D9D9"):
    """Apply gray shading to a table cell"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)

def add_markdown_line(doc: Document, line: str):
    """Add paragraph with **bold** segments."""
    p = doc.add_paragraph()
    parts = line.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        if i % 2 == 1:
            run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
    return p

def add_markdown_paragraph(doc: Document, text: str, style: str | None = None):
    """Add a paragraph with simple **bold** Markdown handling."""
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()

    parts = re.split(r"(\*\*[^\*]+\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = p.add_run(part[2:-2])
            run.bold = True
        else:
            run = p.add_run(part)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
    return p

def add_boxed_text(doc: Document, text: str, font_size: int = 16, bold: bool = True):
    """Grey shaded single-cell table with centered text (for cover letter front page)."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    shade_cell(cell, "D9D9D9")
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    return table

def add_centered_upload_image(doc: Document, uploaded_file, width_in: float = 6.0):
    """Add uploaded image centered"""
    if not uploaded_file:
        return
    img_stream = BytesIO(uploaded_file.getvalue())
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_picture(img_stream, width=Inches(width_in))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def call_groq_cover_letter(
    client_name: str,
    project_title: str,
    offer_ref: str,
    letter_date_str: str,
    executives_block: str,
    invitation_date: str,
    meeting_date: str,
    sender_name: str,
    sender_title: str,
) -> str:
    """Call Groq API to generate the cover letter text."""
    user_prompt = COVER_LETTER_USER_PROMPT_TEMPLATE.format(
        client_name=client_name,
        project_title=project_title,
        offer_ref=offer_ref,
        letter_date=letter_date_str,
        executives_block=executives_block.strip() or "Not provided",
        invitation_date=invitation_date.strip() or "Not provided",
        meeting_date=meeting_date.strip() or "Not provided",
        sender_name=sender_name,
        sender_title=sender_title,
    )

    completion = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": COVER_LETTER_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.3,
        max_tokens=800,
    )

    text = completion.choices[0].message.content.strip()
    if text.startswith("```"):
        parts = text.split("```")
        if len(parts) >= 2:
            text = parts[1]
            if text.startswith("text\n") or text.startswith("markdown\n"):
                text = "\n".join(text.split("\n")[1:])
    return text.strip()

def call_groq_exec_summary(system_text: str, client_name: str, project_title: str) -> str:
    """Call Groq API to generate the Executive Summary text."""
    user_content = (
        f"Client Name: {client_name}\n"
        f"Project / System Name: {project_title}\n\n"
        f"Proposed System Description (for context):\n{system_text}\n\n"
        "Generate the Executive Summary strictly as per the instructions."
    )

    resp = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        temperature=0.4,
        max_tokens=800,
        messages=[
            {"role": "system", "content": EXEC_SUMMARY_SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
    )

    return resp.choices[0].message.content.strip()

# ==================== COMMERCIAL/GROQ FUNCTIONS ====================

# Groq client for price sheet generation
groq_client = Groq(api_key="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")

GROQ_SYSTEM_PROMPT = """
You are a senior commercial analyst for warehouse automation projects.

You receive the contents of an Excel sheet called "Overall Costing" as raw CSV text.
This sheet may contain many detailed costing lines, intermediate totals, taxes, and notes.

Your job is to infer the HIGH-LEVEL "Price Sheet" summary used in proposals.

The high-level Price Sheet is a short table of a few summary lines (typically 3–15),
each corresponding to a major package/component of the solution with a single rolled-up price.
Do NOT list detailed items like small sub-components or line-by-line BOM;
only show the SUMMARY building blocks that a customer would see in the commercial section.

------------------------------------------------
EXAMPLES OF TARGET PRICE SHEETS (FOR REFERENCE)
------------------------------------------------

Example 1 –

Price List- Summary
S.NO   Package                    Price
1      Conveyors Package          ₹ 28,34,27,926
2      Cross Belt Sorter Package  ₹ 29,16,52,094
3      Destinations Package       ₹ 3,83,87,721
4      Services Package           ₹ 3,46,22,786
       Total                      ₹ 64,80,90,527

"Business Cooperation Agreement"
Discount for Delhivery  4.5%
Final Total             ₹ 61,89,26,453


Example 2 – 

Price Sheet
S. No   Component                               Price (USD)
1       Loop CBS + Inducts                      (included or price)
2       Infeed + Bagging Conveyors             $ 398,105
3       Output Chutes                          $ 219,944
4       Software Package & Integration         $ 26,302
5       Packaging & forwarding                 $ 3,523
6       Project Management + Supervision cost  $ 32,268
Total (USD)                                    $ 726,386

Followed by bullets:
• Inco-Terms- Ex-works, India (Greater Noida)
• Taxes: Extra as Applicable
• Price is valid for 30 days from the date of proposal.
(These bullets are not part of the high-level table, but the items+total are.)


---------------------------------------
TASK – WHAT YOU MUST RETURN
---------------------------------------

Use the raw 'Overall Costing' CSV to reconstruct ONLY the high-level summary, in the spirit of the examples above.

1) Identify the main commercial building blocks, such as:
   - Conveyors Package
   - Cross Belt Sorter Package
   - Destinations Package
   - Services Package
   - Loop CBS + Inducts
   - Infeed + Bagging Conveyors
   - Output Chutes
   - Software package / Software Packages & SCADA
   - Steelworks / Steelwork
   - PTL
   - Installation & Commissioning
   - Project Management and Engineering Charges
   - Packaging & forwarding / Packaging & Documentation
   - Freight, Warranty, Hotline, AMC packages
   or similar high-level components used to summarize the cost.

2) For each such high-level component, return:
   - s_no: integer starting from 1 in sequence
   - label: the package/component name in clean human-readable form
   - price: the final total price for that component AS A STRING,
            including currency symbol and formatting exactly as in the sheet
            (e.g. "₹ 28,34,27,926", "SAR 10,959,208", "$ 398,105", "€ 2,871,416", "Included in CBS Price").

   IMPORTANT:
   - Do NOT invent prices.
   - Use values that actually appear in the sheet.
   - If multiple detailed rows roll up into one package, use the rolled-up total that clearly corresponds to that package.
   - Prefer the same format as used for the final summary in the data, if visible.

3) If there is a grand "Total" (for the whole solution), also return:
   - total_row: { "label": "...", "price": "..." }
   For example: { "label": "Total", "price": "₹ 64,80,90,527" }.
   If no obvious total exists, set total_row to null.

4) If there are explicit discount and final total lines (like in the Delhivery example):
   - cooperation_label: e.g. "Business Cooperation Agreement"
   - discount_label: e.g. "Discount for Delhivery"
   - discount_value: e.g. "4.5%"
   - final_total_label: e.g. "Final Total"
   - final_total_value: e.g. "₹ 61,89,26,453"
   If not present, return them as null.

5) Also return:
   - currency: "INR", "SAR", "USD", "EUR", or "MIXED" if multiple currencies appear.
   - price_sheet_title: a short label like "19.1 Price List – Summary" or "20.1 Price Sheet"
                        if visible; otherwise null.

6) OUTPUT FORMAT (VERY IMPORTANT):

Return ONLY a single valid JSON object with this exact shape:

{
  "currency": "INR" | "SAR" | "USD" | "EUR" | "MIXED" | null,
  "price_sheet_title": "string or null",
  "items": [
    {
      "s_no": 1,
      "label": "Conveyors Package",
      "price": "₹ 28,34,27,926"
    },
    ...
  ],
  "total_row": {
    "label": "Total",
    "price": "₹ 64,80,90,527"
  } or null,
  "cooperation_label": "string or null",
  "discount_label": "string or null",
  "discount_value": "string or null",
  "final_total_label": "string or null",
  "final_total_value": "string or null"
}

Do NOT wrap the JSON or ```json``` in markdown.
Do NOT add explanations or commentary.
Just return the JSON object.
"""

GROQ_USER_PROMPT_TEMPLATE = """
Below is the raw CSV export of the 'Overall Costing' sheet of an internal costing file.

Use it to construct the high-level Price Sheet summary as described in the instructions.

Raw CSV:
--------------------
{sheet_csv}
--------------------
"""

def call_groq_for_price_sheet(sheet_csv: str) -> dict:
    """Call Groq API to extract price sheet from costing CSV"""
    user_prompt = GROQ_USER_PROMPT_TEMPLATE.format(sheet_csv=sheet_csv)

    completion = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": GROQ_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.0,
    )

    raw = completion.choices[0].message.content.strip()

    # Strip markdown fences if present
    if raw.startswith("```"):
        parts = raw.split("```")
        if len(parts) >= 2:
            raw = parts[1]
            raw = raw.lstrip("json").lstrip()

    # Extract JSON from first '{' to last '}'
    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end < start:
        raise ValueError(f"Groq response does not contain a JSON object:\n{raw}")

    content = raw[start:end + 1]

    try:
        data = json.loads(content)
    except json.JSONDecodeError as e:
        raise ValueError(f"Groq response was not valid JSON: {e}\nExtracted content:\n{content}")

    return data

def parse_price_string(price_str: str):
    """Extract currency prefix and numeric value from price string"""
    if not price_str:
        return None, None, 0

    m = re.search(r"[-]?\d", price_str)
    if not m:
        return price_str.strip(), None, 0

    prefix = price_str[:m.start()].strip()
    numeric_part = price_str[m.start():].strip()

    digits_only = "".join(ch for ch in numeric_part if ch.isdigit() or ch == ".")
    if digits_only == "":
        return prefix, None, 0

    decimals_count = 0
    if "." in digits_only:
        decimals_count = len(digits_only.split(".")[1])

    try:
        value = float(digits_only)
    except ValueError:
        return prefix, None, decimals_count

    return prefix, value, decimals_count

def format_indian_number(value: float, decimals: int) -> str:
    """Format number with Indian-style digit grouping"""
    if decimals > 0:
        s = f"{value:.{decimals}f}"
    else:
        s = f"{int(round(value))}"

    if "." in s:
        int_part, frac = s.split(".")
    else:
        int_part, frac = s, None

    # Indian grouping
    if len(int_part) > 3:
        last3 = int_part[-3:]
        head = int_part[:-3]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        int_formatted = ",".join(groups + [last3])
    else:
        int_formatted = int_part

    if frac and decimals > 0:
        return int_formatted + "." + frac
    else:
        return int_formatted

def apply_bca_discount_to_price_data(price_data: dict, discount_percent: float) -> str | None:
    """Apply BCA discount on total_row.price and return discounted price string"""
    total_row = price_data.get("total_row")
    if not total_row:
        return None

    price_str = total_row.get("price")
    prefix, value, decimals = parse_price_string(price_str)
    if value is None:
        return None

    discounted = value * (1 - discount_percent / 100.0)
    formatted_number = format_indian_number(discounted, decimals)
    if prefix:
        return f"{prefix} {formatted_number}"
    else:
        return formatted_number

# ==================== INPUT COLLECTION ====================

# Professional Tabs for Input Organization
tab1, tab2, tab3, tab4 = st.tabs(["Basic Information", "Cover Letter & Front Page", "Document Sections", "Commercial & Warranty"])

# TAB 1: Basic Information
with tab1:
    st.subheader("Project Details")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client Name", value="Zepto", placeholder="Enter client name")
        project_name = st.text_input("Project Name", value="Automated Sorting System", placeholder="Enter project name")
    with col2:
        client_logo = st.file_uploader("Client Logo", type=["png", "jpg", "jpeg"])

# TAB 2: Cover Letter & Front Page
with tab2:
    st.subheader("Cover Letter Information")
    col1, col2 = st.columns(2)
    with col1:
        offer_ref = st.text_input("Offer Reference", value="F24-00524", placeholder="e.g., F24-00524")
        letter_date = st.date_input("Letter Date", value=date.today())
        invitation_date_str = st.text_input("Invitation Date (optional)", value="", placeholder="If received formal invitation")
    with col2:
        meeting_date_str = st.text_input("Meeting / Workshop Date(s) (optional)", value="", placeholder="Reference to recent meetings")
        sender_name = st.text_input("Sender Name", value="Sandeep Bansal", placeholder="Person signing the letter")
        sender_title = st.text_input("Sender Title", value="Chief Business Officer", placeholder="Sender's designation")
    
    executives_text = st.text_area(
        "Executives (one per line, include Mr./Ms.)",
        value="Mr. Rahul Didwani\nMr. Vinayak Garg",
        height=80,
        placeholder="List all executives to address"
    )
    
    st.divider()
    st.subheader("Front Page Details")
    col3, col4 = st.columns(2)
    with col3:
        contact_name = st.text_input("Contact Person Name", value="Sanyog Pratap Singh")
        contact_phone = st.text_input("Contact Phone", value="+91 8750052591")
    with col4:
        contact_email = st.text_input("Contact Email", value="Sanyog.Singh@falconautotech.com")
        layout_image = st.file_uploader("Layout Image (for front page)", type=["png", "jpg", "jpeg"])

# TAB 3: Document Sections
with tab3:
    st.subheader("Select Sections to Include")
    
    # Executive Summary
    col1, col2 = st.columns([3, 1])
    with col1:
        st.write("**Executive Summary**")
        exec_summary_pdf = st.file_uploader("Upload Proposed System Description (PDF)", type=["pdf"], key="exec_summary_pdf_upload")
    with col2:
        st.write("‎")  # Spacer
        include_exec_summary = st.checkbox("Include", value=True, key="exec_summary_check")
    
    st.divider()
    
    # Company Profile
    col1, col2 = st.columns([3, 1])
    with col1:
        st.write("**Company Profile**")
        st.caption("Uses static content and images from Static_AboutCompany folder")
    with col2:
        st.write("‎")
        include_company_profile = st.checkbox("Include", value=True, key="company_profile_check")
    
    st.divider()
    
    # Reference Projects
    col1, col2 = st.columns([3, 1])
    with col1:
        st.write("**Reference Projects**")
        st.caption("Includes 5 reference projects with site pictures")
    with col2:
        st.write("‎")
        include_ref_projects = st.checkbox("Include", value=True, key="ref_projects_check")
    
    st.divider()
    
    # Handled Shipment Spectrum
    col1, col2 = st.columns([3, 1])
    with col1:
        st.write("**Handled Shipment Spectrum**")
        st.caption("Auto-selects template based on project name (Linear/Loop/Heavy-duty CBS)")
    with col2:
        st.write("‎")
        include_handled_spectrum = st.checkbox("Include", value=True, key="handled_spectrum_check")

    # Capacity Calculations Section
    col1, col2 = st.columns([3, 1])
    with col1:
        st.write("**Sorter System Capacity Calculations**")
        capacity_excel = st.file_uploader("Upload Capacity Excel (for Sorter System Capacity section)", type=["xlsx", "xls"], key="capacity_excel_upload")
    with col2:
        st.write("‎")
        include_capacity_section = st.checkbox("Include", value=True, key="capacity_section_check")

    st.divider()
    
    # Technical Sections in Grid
    st.subheader("Technical Sections")
    col1, col2 = st.columns(2)
    
    with col1:
        elec_include = st.checkbox("Electrical System", value=True, key="elec_check")
        wcs_include = st.checkbox("Falcon WCS CONTROLIT", value=True, key="wcs_check")
        scada_include = st.checkbox("Falcon Visual Inspection System (SCADA)", value=True, key="scada_check")
        safety_include = st.checkbox("Principal of Safety", value=True, key="safety_check")
        infra_include = st.checkbox("Infrastructure", value=True, key="infra_check")
        client_resp_include = st.checkbox("Client Responsibility", value=True, key="client_check")
    
    with col2:
        key_include = st.checkbox("Key Components Make", value=True, key="key_check")
        prog_include = st.checkbox("Program Organisation", value=True, key="prog_check")
        handover_include = st.checkbox("System Handover", value=True, key="handover_check")
        commercial_include = st.checkbox("Commercial", value=True, key="commercial_check")
        warranty_include = st.checkbox("Warranty Period", value=True, key="warranty_check")
        exclusion_include = st.checkbox("Exclusions", value=True, key="exclusion_check")
    
    st.divider()
    
    # Key Components Editor (if selected)
    if key_include:
        with st.expander("Edit Key Components List"):
            default_components = [
                {"Items": "Belts", "Make": "Forbo / Derco / Habasit"},
                {"Items": "Rollers", "Make": "Falcon"},
                {"Items": "Cross Belt Carriers", "Make": "Falcon"},
                {"Items": "Linear Motors (LIM / LSM / Linear Induction)", "Make": "Falcon / SEW / FWD (as applicable)"},
                {"Items": "Feed Line Motors", "Make": "Falcon"},
                {"Items": "Volume / Barcode Scanners", "Make": "SICK / Cognex / Similar"},
                {"Items": "Weighing Scales", "Make": "Bizerba / Mettler Toledo / Equivalent"},
                {"Items": "Encoders", "Make": "SICK / Falcon"},
                {"Items": "Sensors", "Make": "SICK / Leuze / P&F"},
                {"Items": "PLC", "Make": "Siemens / Omron"},
                {"Items": "Control Panels", "Make": "Rittal / BCH"},
                {"Items": "VFDs", "Make": "Siemens / Lenze / AB / Omron"},
                {"Items": "Cables", "Make": "LAPP / Equivalent"},
                {"Items": "Switch Gear", "Make": "Schneider / Equivalent"},
                {"Items": "Bearings", "Make": "NTN / SKF / Equivalent"},
                {"Items": "Power Transmission Systems", "Make": "Vahle"},
                {"Items": "HMIs", "Make": "Siemens / Omron"},
                {"Items": "MDR", "Make": "Pulse / Itoh Denki"},
                {"Items": "Data Transmission System", "Make": "Siemens"},
            ]
            
            if "key_components_df" not in st.session_state:
                st.session_state["key_components_df"] = pd.DataFrame(default_components)
            
            key_components_edited = st.data_editor(
                st.session_state["key_components_df"],
                num_rows="dynamic",
                width='stretch',
                key="key_editor"
            )
    
    # Program Organisation Gantt (if selected)
    if prog_include:
        with st.expander("Upload Gantt Chart"):
            prog_gantt = st.file_uploader("Gantt Chart (optional)", type=["png", "jpg", "jpeg"], key="prog_gantt")

# TAB 4: Commercial & Warranty
with tab4:
    st.subheader("Commercial Details")
    costing_file = st.file_uploader("Upload Costing Excel File", type=["xlsx", "xls"], help="Excel file with 'Overall Costing' sheet")
    
    if costing_file:
        st.success("Costing file uploaded successfully")
        
        apply_bca = st.checkbox("Apply Business Cooperation Agreement Discount (4.5% on Total)", value=False, key="apply_bca_discount")
        
        st.write("**Payment Terms (editable):**")
        default_payment_terms = [
            {"Payment Percentage": "20%", "Stage": "Advance along with LOI/ PO"},
            {"Payment Percentage": "20%", "Stage": "After DAP Completion"},
            {"Payment Percentage": "40%", "Stage": "Before Dispatch"},
            {"Payment Percentage": "10%", "Stage": "Against Installation"},
            {"Payment Percentage": "10%", "Stage": "Against Handover"},
        ]
        
        if "payment_terms" not in st.session_state:
            st.session_state["payment_terms"] = default_payment_terms
        
        pt_df = pd.DataFrame(st.session_state["payment_terms"])
        edited_pt_df = st.data_editor(pt_df, num_rows="dynamic", width='stretch', key="payment_terms_editor")
        st.session_state["payment_terms"] = edited_pt_df.to_dict(orient="records")
    st.divider()
    st.subheader("Warranty Configuration")
    
    warranty_type = st.selectbox("Warranty Type", ["Standard warranty", "Comprehensive warranty"], key="warranty_type")
    
    col1, col2 = st.columns(2)
    with col1:
        warranty_duration = st.text_input("Warranty Duration", value="1 year", key="warranty_duration")
    with col2:
        warranty_start = st.selectbox(
            "Warranty Start Condition",
            [
                "from the date of beneficiary use.",
                "from the date of commissioning of the system.",
                "from the date of completion of dispatch of the materials, whichever is earlier.",
                "from the date of beneficiary use, max 30 days after readiness of commissioning.",
                "from the date of official communication of material readiness at Falcon end.",
            ],
            key="warranty_start"
        )
    
    warranty_extended = st.checkbox("Include Extended Warranty Option", value=True, key="warranty_extended")
    if warranty_extended:
        warranty_extended_text = st.text_input(
            "Extended Warranty Text",
            value="Post completion of the warranty period, the customer can opt for an Extended Warranty.",
            key="warranty_extended_text"
        )
    else:
        warranty_extended_text = ""
    
    warranty_amc = st.checkbox("Include AMC / Hotline Clause", value=False, key="warranty_amc")
    if warranty_amc:
        warranty_amc_text = st.text_input(
            "AMC Text",
            value="After the warranty year, an additional 2 years are covered under AMC and 24x7 Hotline support.",
            key="warranty_amc_text"
        )
    else:
        warranty_amc_text = ""
    
    warranty_transport = st.checkbox("Include Transportation Note", value=False, key="warranty_transport")
    if warranty_transport:
        warranty_transport_text = st.text_area(
            "Transportation Note",
            value="Note: Transportation cost for sending faulty items to Falcon factory from the site will be in the customer's scope.",
            key="warranty_transport_text"
        )
    else:
        warranty_transport_text = ""
    
    st.divider()
    st.subheader("Exclusions Configuration")
    
    st.write("**Select Exclusions to Include:**")
    
    variable_exclusions = [
        "Server PC / server system.",
        "SCADA / PC for SCADA.",
        "Workstations.",
        "Cabling from server room to Falcon control panel.",
        "Mobile carts.",
        "Collection trolleys / collection trolleys below chutes.",
        "Collection bins.",
        "Pallets at chutes.",
        "Pallets / hand-held terminals for secondary sorting.",
        "Steel works.",
        "Steel works – if not specified.",
        "Mezzanine & staircase.",
        "Mezzanine & staircase not mentioned in BOM.",
        "Maintenance platform / lift required for maintenance activity.",
        "Safety fencing / safety fencing not shown in layout.",
        "HPT/BOPT/Forklift/Hydra/Scaffoldings required for installation.",
        "Stress free mats.",
        "Insulation mats.",
        "Fans at chutes & inducts.",
        "Lighting around chutes / inducts.",
        "Irregular's provision.",
        "UPS power (separate UPS supply).",
        "CE declaration of conformity.",
    ]
    
    selected_exclusions = []
    cols = st.columns(2)
    for idx, item in enumerate(variable_exclusions):
        with cols[idx % 2]:
            if st.checkbox(item, key=f"excl_{idx}"):
                selected_exclusions.append(item)

st.divider()

# ==================== DOCUMENT GENERATION FUNCTIONS ====================

def build_cover_letter_section(doc, letter_text):
    """Build cover letter (page 1) - NO HEADER for this section"""
    HEADER_PREFIXES = (
        "Kind Attention",
        "Mr.",
        "Ms.",
        "M/s",
        "Offer Ref:",
        "Subject –",
        "Subject -",
        "Date:",
        "Location –",
    )
    
    lines = [l.rstrip() for l in letter_text.splitlines() if l.strip() != ""]
    for idx, line in enumerate(lines):
        # Check if line starts with header prefixes or contains "Dear" or ends with signature (Best Regards)
        is_header = any(line.startswith(pfx) for pfx in HEADER_PREFIXES) or "Dear " in line
        is_signature = "Best Regards" in line or idx >= len(lines) - 2

        if is_header:
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True
        elif is_signature:
            # Signature and name/title at end should be bold
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True
        else:
            if "**" in line:
                add_markdown_line(doc, line)
            else:
                p = doc.add_paragraph(line)
                apply_normal_style(p)


def build_front_page_section(doc, project_title, offer_ref, contact_name, contact_email, contact_phone, layout_image):
    """Build front page (page 2) - NO HEADER for this section"""
    doc.add_page_break()

    # Top box: "Response to RFP for"
    add_boxed_text(doc, "Response to RFP for", font_size=18, bold=True)

    # Project name box
    if project_title:
        add_boxed_text(doc, project_title, font_size=16, bold=True)

    # Proposal reference box
    if offer_ref:
        add_boxed_text(doc, f"Proposal Reference: {offer_ref}", font_size=14, bold=True)

    # Some vertical spacing
    doc.add_paragraph("")

    # Layout image
    add_centered_upload_image(doc, layout_image, width_in=6.5)

    # More spacing
    doc.add_paragraph("")

    # Contact box at bottom
    contact_lines = [
        "Falcon Autotech Private Limited",
        "Plot No. 87, Sector Ecotech-1, Extention-1, Greater Noida, Uttar Pradesh 201308.",
        "",
        f"Contact – {contact_name}",
        "Assistant Manager",
        f"Mob - {contact_phone}",
        contact_email,
    ]
    contact_text = "\n".join(contact_lines)
    table = add_boxed_text(doc, contact_text, font_size=11, bold=False)
    # make contact text paragraphs centered
    cell = table.rows[0].cells[0]
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def build_glossary_section(doc):
    """Build glossary/table of contents section with automatic TOC"""
    doc.add_page_break()
    
    p = doc.add_heading("Table of Contents", level=1)
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.underline = True
        run.font.color.rgb = HEADING_COLOR
    
    # Add automatic TOC field
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar3)
    
    # Add instruction text for users
    doc.add_paragraph("")
    p = doc.add_paragraph("Note: Right-click on the table of contents and select 'Update Field' to refresh page numbers.")
    apply_normal_style(p)
    p.runs[0].italic = True


def build_executive_summary_section(doc, exec_summary_text, counter):
    """Build Executive Summary section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Executive Summary", counter=counter)
    
    lines = exec_summary_text.strip().splitlines()
    for line in lines:
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph("")
            continue
        
        # Check if it's a bullet point line
        if stripped.startswith("•") or stripped.startswith("-"):
            bullet_text = stripped.lstrip("•- ").strip()
            p = doc.add_paragraph(style="List Bullet")
            if "**" in bullet_text:
                parts = bullet_text.split("**")
                for i, part in enumerate(parts):
                    if not part:
                        continue
                    run = p.add_run(part)
                    if i % 2 == 1:
                        run.bold = True
                    run.font.name = "Calibri"
                    run.font.size = Pt(11)
                    run.italic = True
            else:
                run = p.add_run(bullet_text)
                run.font.name = "Calibri"
                run.font.size = Pt(11)
                run.italic = True
        else:
            # Regular paragraph with possible **bold**
            if "**" in stripped:
                add_markdown_line(doc, stripped)
            else:
                p = doc.add_paragraph(stripped)
                apply_normal_style(p)


def build_company_profile_section(doc, counter):
    """Build Company Profile section with static images"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Company Profile", counter=counter)

    top_text = (
        "Falcon Autotech (Falcon) is a global intralogistics automation solutions company. "
        "With over 10 years of experience, Falcon has worked with some of the most innovative "
        "brands in E-Commerce, CEP, Fashion, Food/FMCG, Auto and Pharmaceutical Industries. "
        "With our proprietary software and robust hardware integration capabilities, Falcon designs, "
        "manufactures, supplies, implements, and maintains world-class warehouse automation systems globally. "
        "Falcon's strong research and development team and the continuous focus on innovation reflect our strong "
        "solution line around Sortation, Robotics, Conveying, Vision Systems and IOT. "
        "Falcon has done over 1,800 installations across 15 countries on four continents."
    )
    p = doc.add_paragraph(top_text)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "1.png"), width_in=4)

    bottom_text = (
        "Falcon Autotech is currently among the top 15 intralogistics automation companies; "
        "our vision is to become a top 10 intralogistics automation company in our focused product lines."
    )
    p = doc.add_paragraph(bottom_text)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "2.png"), width_in=4)

    doc.add_page_break()

    # Page 2
    top_text2 = (
        "The team started out in 2004 solving special purpose automation problems for clients and later "
        "established Falcon Autotech in 2012 with a strong focus on building a standard technology stack spanning "
        "across hardware, firmware, and software to tackle larger supply chain problems around warehouse "
        "automation and material handling. "
        "Over the decade, Falcon has made rapid strides and has carved out a niche in some of the world's most "
        "cutting-edge technologies: Sortation, Robotics, Conveying, Vision Systems and IOT."
    )
    p = doc.add_paragraph(top_text2)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "3.png"), width_in=6)

    bottom_text2 = (
        "As a leading player in the intralogistics automation space, Falcon continuously strives to improve the "
        "operational efficiencies and accuracies for its clients through its domain knowledge and experience, in "
        "addition to its wide range of products and solutions. In order to live up to the high expectations set "
        "forth by our clients, the team at Falcon realizes the importance of taking up selective applications in "
        "focused industries and delivering world-class projects in return."
    )
    p = doc.add_paragraph(bottom_text2)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "4.png"), width_in=6)

    # Page 3
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "5.png"), width_in=6)

    bottom_text3 = (
        "Falcon Autotech has successfully delivered warehouse automation solutions based on smart and innovative "
        "combinations of the above product lines for effective materials handling, sortation and movement. "
        "The process is controlled in real-time by our in-house WCS applications. These solutions considerably "
        "reduce the need for manual operations, improve working conditions and ensure the highest accuracy of the "
        "entire process up to final delivery to the recipient.\n\n"
        "Over the last 10 years, Falcon has worked with some of the most innovative brands worldwide and has "
        "established long-standing partnerships. These brands are testimony to our strong focus on delivering "
        "superior customer satisfaction and offering end-to-end intralogistics solutions."
    )
    p = doc.add_paragraph(bottom_text3)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "6.png"), width_in=6)

    doc.add_page_break()

    # Page 4
    bottom_text4 = (
        "With over 1,800 installations, Falcon's systems are used all over the globe. Falcon has a highly "
        "motivated team of 600+ employees supported by over 15 global partners who help us design, manufacture, "
        "deliver and maintain automation solutions worldwide."
    )
    p = doc.add_paragraph(bottom_text4)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "7.png"), width_in=6)

    add_numbered_subheading(doc, "Customer Engagement Model", f"{counter}.1")
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "8.png"), width_in=7)

    doc.add_page_break()

    # Page 5
    add_numbered_subheading(doc, "Falcon's Experience and Achievements in Sortation Space Globally", f"{counter}.2")

    bullet_points = [
        "Ranked among Top 10 Sortation System Suppliers globally.",
        "Currently possess one of the world's largest portfolios in sortation technologies (7 in-house technologies).",
        "Total installed capacity of 10 million shipments per day worldwide.",
        "Only company to be able to offer a fully integrated AMS.",
    ]

    for point in bullet_points:
        p = doc.add_paragraph(point, style="List Bullet")
        apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "9.png"), width_in=5)
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "10.png"), width_in=5)
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "11.png"), width_in=5)
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "12.png"), width_in=5)


def build_reference_projects_section(doc, counter):
    """Build Reference Projects section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Reference Projects", counter=counter)

    intro = (
        "Falcon has a strong legacy in Warehousing Automation solutions and references-"
    )
    p = doc.add_paragraph(intro)
    apply_normal_style(p)

    bullets_intro = [
        "Expertise in Shipment Sortation, Piece Picking and Handling, Case Picking and Handling.",
        "Lifecycle services (maintenance, spares supply chain, support).",
        "Full in-house expertise (Hardware/Software).",
        "Turn-key tailored solutions.",
        "The references list presented below focuses on Sortation Solution –",
    ]
    for text in bullets_intro:
        p = doc.add_paragraph(text, style="List Bullet")
        apply_normal_style(p)

    # Project 1
    add_numbered_subheading(doc, "Project 1- (CEP Client, India)", f"{counter}.1")

    p = doc.add_paragraph(
        "The system is equipped with two fully automated and interconnected sub-systems. "
        "Sub-System 1 is designed for handling large B2B boxes and E-commerce shipment bags while "
        "Sub-System 2 is designed to handle small E-commerce packages."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec1 = [
        "48,000 PPH (Double Deck CBS – Shipment Sorter).",
        "17,000 PPH (Double Deck CBS – Bag Sorter).",
        "Building Size: 700,000 Sq. Ft.",
    ]
    for t in spec1:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm1 = [
        "2 Sets of Double Decker CBS Sorters.",
        "Mezzanine Structures.",
        "Automated Singulators.",
        "Fully Automatic Inductions.",
        "Semi-Automatic Inductions.",
        "Telescopic Belt Conveyors.",
        "PVC Belt Conveyors.",
        "Modular Belt Conveyors.",
        "Spiral Chutes with Braking Rollers.",
        "5-Sided Scanning Tunnels.",
        "High Speed Weighing Conveyors.",
        "Direct Bagging Chutes.",
        "Put to Light Chutes.",
        "Volume Distribution Systems.",
        "High Availability Server Systems.",
        "WCS.",
    ]
    for t in ktm1:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Pictures –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj1.PNG")

    doc.add_page_break()

    # Project 2
    add_numbered_subheading(doc, "Project 2- (Client – E-Commerce, India)", f"{counter}.2")

    p = doc.add_paragraph()
    run = p.add_run("Use Case – Destination sorting of packed shipments.")
    run.bold = True
    apply_normal_style(p)

    p = doc.add_paragraph(
        "In 2019, the client was looking for a potential automation partner for design and development of a "
        "new automated sortation system for B2C shipments. The system needed to provide maximum uptime with "
        "reduced dependency on skilled manpower and better space optimization. "
        "\nThe customer chose Falcon Autotech based on its unique design that addressed these pain points, "
        "its capability for seamless WMS integration, and its life cycle support services."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec2 = [
        "Throughput: 27,600 PPH.",
        "End Destinations: 410 Direct Outputs.",
        "Building Size: 200,000 Sq. Ft.",
    ]
    for t in spec2:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm2 = [
        "Bulk Infeed Conveyors.",
        "ARB based Volume Distribution System.",
        "Integrated Presort System.",
        "Irregular Ejection System.",
        "Automatic Induct Lines.",
        "Automatic Barcode Scanner with Image Capture.",
        "Automatic Weight & Volume Measurement System.",
        "Linear Cross Belt Sorter.",
        "Smart Sliding Chutes for Direct Bagging and Cage Sorting.",
        "Bag Take-out System.",
        "WCS Software System.",
    ]
    for t in ktm2:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Pictures –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj2.PNG")

    doc.add_page_break()

    # Project 3
    add_numbered_subheading(doc, "Project 3- (Client – E-Commerce, India)", f"{counter}.3")

    p = doc.add_paragraph()
    run = p.add_run("Use Case – Destination sorting of packed shipments.")
    run.bold = True
    apply_normal_style(p)

    p = doc.add_paragraph(
        "The customer chose Falcon Autotech based on its unique design, its ability to integrate seamlessly "
        "with the WMS, and its strong life cycle support services."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec3 = [
        "Throughput: 24,000 PPH.",
        "End Destinations: 40 Collection Type Chutes.",
    ]
    for t in spec3:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm3 = [
        "Bulk Infeed Conveyors.",
        "ARB based Volume Distribution System.",
        "Irregular Ejection System.",
        "Automatic Induct Lines.",
        "Automatic Barcode Scanner with Image Capture.",
        "Automatic Weight & Volume Measurement System.",
        "Linear Cross Belt Sorter.",
        "Smart Collection Type Chutes.",
        "Bag Take-out System.",
        "WCS Software System.",
    ]
    for t in ktm3:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Pictures –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj3.PNG")

    doc.add_page_break()

    # Project 4
    add_numbered_subheading(doc, "Project 4- (CEP Client, UK)", f"{counter}.4")

    p = doc.add_paragraph(
        "This solution is designed to handle a volume of 7,200 shipments per hour. "
        "The system is equipped with three infeed conveyors integrated with an automatic label applicator "
        "before shipments enter the sortation system. Shipments are sorted using Falcon's Loop Cross Belt Sorter "
        "equipped with automatic barcode scanning, dimensioning, weighing, and image capture capabilities. "
        "The sorter is installed on the mezzanine floor and sorts directly to 58 end destinations."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec4 = [
        "Throughput: 7,200 PPH.",
        "End Destinations: 58 Nos.",
    ]
    for t in spec4:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm4 = [
        "Powered Belt Conveyors.",
        "Automatic Induct Lines.",
        "Automatic Barcode Scanner with Image Capture.",
        "Automatic Weight & Volume Measurement System.",
        "Loop Cross Belt Sorter.",
        "WCS Software System.",
    ]
    for t in ktm4:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Picture –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj4.PNG")

    doc.add_page_break()

    # Project 5
    add_numbered_subheading(doc, "Project 5- (CEP Client, Sydney)", f"{counter}.5")

    p = doc.add_paragraph(
        "This solution is designed for handling a throughput of 16,000 shipments per hour with the help of "
        "Falcon's Loop Cross Belt Sorter. The system consists of two feeding zones with a total of ten feedlines. "
        "Sorter design enables van drivers to directly drop shipments at the dock doors. It has a total of 369 end "
        "destinations achieved through a combination of direct drops and PTLs. The system is integrated with "
        "five-side automatic barcode scanning, weight and volume measurement, and automatic detection of "
        "oversize and overweight shipments."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec5 = [
        "Throughput: 16,000 PPH.",
        "End Destinations: 369 Nos.",
    ]
    for t in spec5:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm5 = [
        "Powered Belt Conveyors.",
        "2 Induct Zones.",
        "5-side Automatic Barcode Scanner.",
        "Automatic Weight & Volume Measurement System.",
        "Automatic Detection of Oversize Shipments.",
        "Loop Cross Belt Sorter.",
        "WCS Software System.",
    ]
    for t in ktm5:
        p = doc.add_paragraph(t, style="List Bullet")
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Picture –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj5.PNG")


def build_handled_spectrum_section(doc, counter, project_name, client_name):
    """Build Handled Shipment Spectrum section"""
    doc.add_page_break()  # Start on new page
    tpl = choose_sorter_template(project_name)

    item_singular = tpl.item_singular.lower()
    item_cap = item_singular.capitalize()
    item_plural = item_singular + "s"
    item_plural_cap = item_plural.capitalize()

    add_numbered_heading(doc, "Handled Shipment Spectrum", counter=counter)

    intro_1 = (
        f"{client_name} operates in a business where handling a wide spectrum of {item_plural} is critical. "
        f"Falcon has carefully analyzed the provided {item_singular} spectrum and tailored the solution to your needs."
    )
    intro_2 = (
        f"Falcon proposes to use its \"{tpl.config_name}\" to deliver maximum operational benefits to {client_name}, "
        f"ensuring reliable handling of all relevant sizes and weights for your business."
    )
    p = doc.add_paragraph(intro_1)
    apply_normal_style(p)
    p = doc.add_paragraph(intro_2)
    apply_normal_style(p)

    # Subsection 1
    add_numbered_subheading(doc, tpl.subheading_51, f"{counter}.1")

    p = doc.add_paragraph(
        f"Falcon's {tpl.config_name} has a capability to handle the below mentioned "
        f"{item_plural} sizes and weight."
    )
    apply_normal_style(p)

    # Table
    table = doc.add_table(rows=1 + len(tpl.spec_table), cols=3)
    apply_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Specification"
    hdr_cells[1].text = "Unit"
    hdr_cells[2].text = "Value"
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri (Body)'
                run.font.size = Pt(11)

    for spec, data in tpl.spec_table.items():
        if str(data["value"]).strip() == "":
            continue  # Skip empty rows
        row_cells = table.add_row().cells
        row_cells[0].text = spec
        row_cells[1].text = data["unit"]
        row_cells[2].text = data["value"]
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                apply_normal_style(paragraph)

    # Subsection 2
    add_numbered_subheading(
        doc,
        f"{item_plural_cap} to be loaded on Sorter shall have the following characteristics:",
        f"{counter}.2"
    )

    bullets_52 = [
        f"Centre of Gravity of item must not move during conveyance or sorting.",
        f"Item must not have magnetic content, otherwise behavior of {item_singular} cannot be guaranteed.",
        "Liquid or fragile material, to avoid breaking, spillage or leakage, such as wine bottles, "
        "metal cans of paint are designated as non-conveyable items.",
        f"{item_plural_cap} shall be perfectly and safely packaged: protrusion or open surfaces are not allowed.",
        "Plastic ropes shall be perfectly adherent to the surface of the package.",
        "All items with the risk of being damaged during the transport on an automatic sorting system "
        "or damaging the sorting system; they must be robust enough to avoid disintegration of container "
        "material and loss of contents in the sorting process.",
        "Item packaging shall have enough grip to be handled on the belts during the acceleration and "
        "referencing phases.",
        "Items shall not have slippery surfaces and must be able to withstand acceleration of the items "
        "on the belt during the start-stop phases (accelerations up to 0.5 g shall be assured without any "
        "sliding or tumbling of the items on the belt conveyor).",
        f"The {item_plural} must have at least one flat and regular surface providing enough stability during "
        "conveyance.",
        "All shapes are permitted except spherical, cylindrical, or alike unstable items & shapes.",
        "All usual packaging materials are permitted (including paper, carton, plastics, plastic foil, rope, "
        "tape, textile, and wood).",
    ]
    
    for text in bullets_52:
        p = doc.add_paragraph(text, style="List Number 2")
        apply_normal_style(p)

    # Subsection 3
    add_numbered_subheading(doc, f"{item_plural_cap} not loadable on the sorter", f"{counter}.3")

    bullets_53 = [
        "Unstable items with a risk to roll or tumble on the sorting system, such as spherical or cylindrical items.",
        "Items that have a spherical or cylindrical shape.",
        "Items that are packed in material that can damage the conveyors or the sorter.",
        "Items that have sharp points (e.g., Nails) or sharp edges, that can damage the conveyors or the sorter.",
        f"Fragile {item_plural} with contents not sufficiently secured.",
        "Items that have been classified as dangerous are designated.",
        "Wet items are designated.",
        "Items with anti-slip treatment.",
        "Items with protruding parts.",
        "Items with sharp edges.",
        "Inadequately packed items that could be damaged during automatic transportation.",
        "Electrostatically loaded items.",
        "Loose parts on loads and load carriers, such as adhesive tape, stickers, slips of paper, straps, "
        "wrap foil etc. are designated as non-conveyable items.",
    ]
    
    for text in bullets_53:
        p = doc.add_paragraph(text, style="List Number 2")
        apply_normal_style(p)


def build_capacity_calculations_section(doc, counter, client_name, project_name, capacity_excel):
    """Add Sorter System Capacity section using capacity_calculations.py logic"""
    from capacity_calculations import build_capacity_prompt_from_excel, call_groq_for_capacity, add_capacity_section_to_doc
    if capacity_excel:
        excel_bytes = capacity_excel.read()
        prompt = build_capacity_prompt_from_excel(excel_bytes, client_name, project_name)
        cap_data = call_groq_for_capacity(prompt)
        doc.add_page_break()
        add_capacity_section_to_doc(doc, client_name, project_name, cap_data)


def build_electrical_section(doc, counter):
    """Build Electrical System section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Electrical System", counter=counter)
    
    p = doc.add_paragraph(
        "Main power supply will supply Falcon's PDP (Power Distribution Panels) electrical cabinets. "
        "PDP cabinets supply the entire system via secondary cabinets:"
    )
    apply_normal_style(p)
    
    for item in ["Main Control Cabinet", "Induct Control Panels", "Remote Cabinets for Sorter I/O", "Scanner Control cabinets"]:
        p = doc.add_paragraph(item, style="List Bullet")
        apply_normal_style(p)
    
    add_numbered_subheading(doc, "Reference Picture of Power Distribution Panel", f"{counter}.1")
    add_centered_image(doc, "FIXED_IMAGE/elec1.PNG", width_in=4)
    
    add_numbered_subheading(doc, "Main Control Panel (Reference)", f"{counter}.2")
    add_centered_image(doc, "FIXED_IMAGE/elec2.PNG", width_in=4)
    
    add_numbered_subheading(doc, "Induct Stations Control Panel (Reference)", f"{counter}.3")
    add_centered_image(doc, "FIXED_IMAGE/elec3.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("Engines")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "Three-phase alternating current motors (Induction) will be used through a frequency converter. "
        "The engines will be coupled with a converter to improve consumption and reduce the carbon footprint."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph("All motors will have appropriate IP ratings.")
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Sensors")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The sensors will be supplied, standardized by type, with connector, with a cable length suitable for "
        "easy extraction, suitably protected from possible impacts."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Control command")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The proposed solution is based on SIEMENS Programmable Logic Controller technology (PLC) platform. "
        "The entire system will be logically divided into Zones (Sorter / Feed Line / Loop), each managed by a PLC. "
        "The planned primary communication protocol is going to be ProfiNet."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Conveyor interface")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The frequency converter of each conveyor allows the acquisition of the signals of the "
        "sensors/actuators/GIOs associated with it (e.g. conveyor end detection photocells, blockage detection "
        "photocells). Each frequency converter will be connected in series by means of the ProfiNet field bus."
    )
    apply_normal_style(p)

def build_wcs_section(doc, counter, client_name):
    """Build WCS CONTROLIT section with COMPLETE content"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Falcon's WCS CONTROLIT", counter=counter)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs1.PNG", width_in=3.0)
    
    p = doc.add_paragraph(
        "Falcon WCS (Warehouse Control System) is an in-house developed IT solution by Falcon Autotech, "
        "serving as the brain behind the company's sortation solutions. It manages the real-time movement "
        "of goods and data across the system, ensuring efficient operations in high-throughput warehouses. "
        "Falcon WCS integrates seamlessly with Warehouse Management Systems (WMS), Transport Management "
        "Systems (TMS), and other external applications via APIs to enhance operational efficiency."
    )
    apply_normal_style(p)
    
    # A. System Architecture
    add_numbered_subheading(doc, "System Architecture", f"{counter}.1")
    
    p = doc.add_paragraph()
    run = p.add_run("High-Level Design (HLD) Overview")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The Falcon WCS integrates with external systems like the Warehouse Management System (WMS) and "
        "Transport Management System (TMS). Communication occurs via APIs / WSDL / MQ communication "
        "protocols, ensuring smooth data flow for order management, shipment tracking, and other critical "
        "operations."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs2.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("Key Components:")
    run.bold = True
    apply_normal_style(p)
    
    # Presentation & Session Layer
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("Presentation and Session Layer:")
    run.bold = True
    apply_normal_style(p)
    
    for item in [
        "MySQL Database: Stores operational data, shipment details, and sortation instructions.",
        "Sorter Services: Responsible for managing sorting logic and directing parcels to appropriate destinations.",
        "Dashboard: Provides a user interface for real-time monitoring of warehouse operations and performance metrics.",
        "Integration Services: Handles communication with external systems (e.g., WMS, TMS) and ensures data consistency across platforms."
    ]:
        p = doc.add_paragraph(item, style="List Bullet 2")
        apply_normal_style(p)
    
    # Application Layer
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("Application Layer:")
    run.bold = True
    apply_normal_style(p)
    
    for item in [
        "Image Services: Processes and manages images captured during the sortation process.",
        "ICR Software: Utilizes Image Character Recognition to read parcel labels and identify shipment information.",
        "PLC Software: Interfaces with Programmable Logic Controllers to manage the physical movement of parcels and control sortation equipment."
    ]:
        p = doc.add_paragraph(item, style="List Bullet 2")
        apply_normal_style(p)
    
    # Transport Layer
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("Transport Layer:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "Sorter PLCs: Receive commands from the session layer (Sorter Services) and execute sorting operations based on real-time data.",
        style="List Bullet 2"
    )
    apply_normal_style(p)
    
    # System Communication
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("System Communication:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "All layers are connected via a stacked switch, which provides internet and intranet connectivity. "
        "Communication between the sortation system and external systems for results or shipment data occurs through this switch.",
        style="List Bullet 2"
    )
    apply_normal_style(p)
    
    # B. High Availability Architecture
    add_numbered_subheading(doc, "High Availability Architecture", f"{counter}.2")
    
    p = doc.add_paragraph(
        "The Falcon WCS architecture ensures uninterrupted operations using a High Availability (HA) server setup. "
        "The system is designed to handle both planned and unplanned downtime, providing robust mechanisms for "
        "failover, replication, and data redundancy."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs3.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("Key Components and Features of the High Availability Architecture:")
    run.bold = True
    apply_normal_style(p)
    
    # Component descriptions
    components = [
        ("Stacked Switch:", [
            "Centralizes data exchange between NAS, nodes, domain controller (DC), and peripherals.",
            "Analyzes packet headers to reduce unnecessary data transmission, enhancing LAN efficiency."
        ]),
        ("Domain Controller:", [
            "Heartbeat Monitoring: Tracks the status of nodes and initiates VM failover when necessary.",
            "Image Hosting: Stores and manages images received from the ICR (Image Character Recognition)."
        ]),
        ("NAS (Network Attached Storage):", [
            "Centralized data storage providing access to connected devices and virtual machines.",
            "Redundancy: Two NAS boxes with mirrored drives ensure data protection and availability, offering a failsafe against hardware failure."
        ]),
        ("Node:", [
            "Hyper Terminals: Nodes host and manage virtual machines (VMs) to run the warehouse control systems and related applications.",
            "Clustering: Nodes are clustered using Microsoft Windows Cluster to enable failover protection, ensuring continuous operation even in case of hardware failure."
        ]),
        ("Virtual Machine & InnoDB Cluster:", [
            "Primary VM: Hosts Falcon WCS services, while a secondary backup on the node ensures failover through network load balancing (NLB).",
            "InnoDB Cluster: Ensures data replication using a Master–Slave–Slave setup for MySQL databases, maintaining consistency and availability."
        ]),
        ("NAS Cluster:", [
            "Unified File System: NAS nodes share files across the cluster, ensuring no data loss during failover or disaster recovery.",
            "Backup NAS: Provides redundancy by replicating data between two NAS boxes, further safeguarding against failures."
        ])
    ]
    
    for comp_title, comp_items in components:
        p = doc.add_paragraph()
        run = p.add_run(comp_title)
        run.bold = True
        apply_normal_style(p)
        
        for comp_item in comp_items:
            p = doc.add_paragraph(comp_item, style="List Bullet")
            apply_normal_style(p)
    
    # Disaster Handling
    p = doc.add_paragraph()
    run = p.add_run("Disaster Handling:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph("Recovery Time Objective (RTO) & Data Loss Objective (RPO):", style="List Bullet")
    apply_normal_style(p)
    
    for disaster_item in [
        "VM Cluster Failure: RTO = 1 hour; RPO = 1 hour.",
        "Node Failure: No impact with a single failure; RTO = 4 hours if both nodes fail.",
        "NAS Failure: Backup NAS available with no downtime, ensuring continued operation."
    ]:
        p = doc.add_paragraph(disaster_item, style="List Bullet 2")
        apply_normal_style(p)
    
    # C. WCS User Interface
    add_numbered_subheading(doc, "WCS User Interface", f"{counter}.3")
    
    p = doc.add_paragraph(
        "The Falcon WCS features a robust, user-friendly dashboard that provides real-time visibility "
        "into warehouse and sortation operations. The dashboard serves as the primary interface for "
        "monitoring key system metrics, tracking performance, and ensuring smooth operations."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Dashboard Overview")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The WCS dashboard offers real-time data visualization, helping warehouse operators and IT teams "
        "make data-driven decisions. Users can monitor system health, performance, and detect anomalies "
        "through an intuitive graphical interface."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Key Features of the Dashboard:")
    run.bold = True
    apply_normal_style(p)
    
    dashboard_features = [
        "System Health Monitoring: Displays metrics such as CPU utilization, memory usage, disk performance, and system load across the infrastructure.",
        "Real-Time Sortation Monitoring: Shows the real-time movement of parcels within the sortation system, including chute assignments and shipment statuses.",
        "Error Reporting: Notifies users of system errors, network disruptions, and potential failures in real time, allowing for quick resolution and minimal downtime.",
        "Performance Metrics: Provides detailed reports on sortation throughput, parcel handling times, and system efficiency to ensure that warehouse targets are met.",
        "User Role Management: The dashboard allows different levels of access based on user roles, ensuring that the right personnel can view or manage the system as needed."
    ]
    
    for feature in dashboard_features:
        p = doc.add_paragraph(feature, style="List Bullet")
        apply_normal_style(p)
    
    p = doc.add_paragraph(
        "In the context of this IT dashboard, the following user interactive screens are provided:"
    )
    apply_normal_style(p)
    
    # Dashboard screens with images
    dashboard_screens = [
        ("Dashboard (Home Screen): Provides an overview of important metrics, data visualizations, and summary information related to the IT system or processes.", "FIXED_IMAGE\\wcs4.PNG"),
        ("Live Bags: Displays real-time information and status updates regarding bags or parcels currently in transit or being processed.", "FIXED_IMAGE\\wcs5.PNG"),
        ("Bay Status: Offers insights into the status and availability of different processing bays or areas within the system.", "FIXED_IMAGE\\wcs6.PNG"),
        ("Processed Packages: Shows details and statistics related to packages or items that have been successfully processed or handled by the system.", "FIXED_IMAGE\\wcs7.PNG"),
        ("Configuration Setting: Enables users to configure and customize various settings and parameters within the IT system or dashboard.", "FIXED_IMAGE\\wcs8.PNG")
    ]
    
    for screen_desc, screen_img in dashboard_screens:
        p = doc.add_paragraph(screen_desc, style="List Bullet")
        apply_normal_style(p)
        add_centered_image(doc, screen_img)
    
    # Additional screens without images
    additional_screens = [
        "Report & Analysis: Allows users to generate and access comprehensive reports, analytics, and insights based on the data collected by the IT dashboard.",
        "Rejection Bay Mapping: Provides functionality to map and manage rejection bays or areas where packages are deemed unsuitable for processing.",
        "Alarms: Displays alerts, notifications, or alarms related to system events, errors, or anomalies that require attention or investigation.",
        "Calibration Settings: Allows users to adjust and calibrate system settings, parameters, or sensors to ensure accurate and reliable performance.",
        "Operator Management: Offers features and tools to manage and monitor the operators or personnel responsible for operating the IT system.",
        "User Management: Provides functionality to manage user accounts, permissions, roles, and access levels within the IT dashboard.",
        "User Guide: The 'User Guide' page offers comprehensive documentation and instructions on how to use the IT dashboard effectively. It serves as a reference guide for users."
    ]
    
    for screen in additional_screens:
        p = doc.add_paragraph(screen, style="List Bullet")
        apply_normal_style(p)
    
    # D. Communication Architecture
    add_numbered_subheading(doc, "Communication Architecture", f"{counter}.4")
    
    p = doc.add_paragraph(
        "Falcon WCS operates within a highly interconnected system, ensuring seamless communication between "
        "the WCS server, on-premises devices (such as sorter PLCs, PTL devices, 1D scanners, and HHT devices), "
        "and client systems. This communication architecture facilitates real-time data exchange and operational "
        "control, optimizing sortation processes and warehouse efficiency."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs9.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("On-Premises Communication")
    run.bold = True
    apply_normal_style(p)
    
    # Communication devices
    comm_devices = [
        ("Sorter PLC Devices:", [
            "Protocol: Falcon WCS communicates with sorter PLCs using either the Siemens S7 protocol or the Omron communication protocol.",
            "Functionality: The sorter PLC devices receive sortation instructions from the WCS and execute the sorting process by directing parcels to the appropriate chute based on the system's real-time data."
        ]),
        ("PTL (Pick-to-Light) Devices:", [
            "Protocol: PTL devices communicate with Falcon WCS using the TCP/IP protocol.",
            "Functionality: The system sends commands to the PTL devices for guiding manual picking operations by lighting up indicators at the appropriate bins or shelves, improving operational accuracy and speed."
        ]),
        ("1D Scanners:", [
            "Protocol: These barcode scanners also use the TCP/IP protocol to communicate with the WCS.",
            "Functionality: The scanners capture barcode data from the parcels, and this information is sent to the WCS for processing, such as determining sorting destinations."
        ]),
        ("HHT (Handheld Terminal) Devices:", [
            "Protocol: The wireless HHT devices communicate with Falcon WCS over Wi-Fi.",
            "Functionality: The HHT devices send scan input data (e.g., barcodes) to the server over Wi-Fi. The WCS processes this data and sends the required output instructions back to the HHT device and associated PTL devices. The HHT device executes these instructions, facilitating real-time decision making and execution for operators."
        ])
    ]
    
    for device_title, device_items in comm_devices:
        p = doc.add_paragraph()
        run = p.add_run(device_title)
        run.bold = True
        apply_normal_style(p)
        
        for device_item in device_items:
            p = doc.add_paragraph(device_item, style="List Bullet")
            apply_normal_style(p)
    
    # E. Client Communication
    add_numbered_subheading(doc, "Client Communication", f"{counter}.5")
    
    p = doc.add_paragraph()
    run = p.add_run("Data Transfer Methods:")
    run.bold = True
    apply_normal_style(p)
    
    transfer_methods = [
        "API: Falcon WCS can communicate processed data to client systems through API calls, allowing for seamless integration with external software.",
        "MQ (Message Queuing): Falcon WCS can also send data via message queues, ensuring reliable delivery of messages even during network downtime.",
        "WSDL/XML: For structured data exchanges, Falcon WCS supports WSDL and XML formats for client communication.",
        f"Other Protocols: Additional methods for data transfer may include customized protocols depending on {client_name}'s requirements."
    ]
    
    for method in transfer_methods:
        p = doc.add_paragraph(method, style="List Bullet")
        apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Purpose:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The data sent to the client can include sortation results, system performance reports, and operational "
        "analytics, which can be used for further processing or reporting within external systems like Warehouse "
        "Management Systems (WMS) and Transport Management Systems (TMS).",
        style="List Bullet"
    )
    apply_normal_style(p)
    
    # F. HAA Server Specifications
    add_numbered_subheading(doc, f"HAA Server Specifications (In {client_name}'s Scope)", f"{counter}.6")
    
    p = doc.add_paragraph("20-core configuration with 128 GB RAM in T440 and 64 GB RAM in T40.")
    apply_normal_style(p)
    
    # Server spec table
    table = doc.add_table(rows=1, cols=3)
    apply_table_style(table)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "SN"
    hdr_cells[1].text = "Description"
    hdr_cells[2].text = "Qty"
    
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri (Body)'
                run.font.size = Pt(11)
    
    server_specs = [
        ("1", "DELL PowerEdge T440 Server", "2"),
        ("2", "Intel Xeon Silver 4210R 2.4G, 10C/20T, 9.6GT/s, 13.75M Cache, Turbo, HT (100W) DDR4-2400", "4"),
        ("3", "32GB RDIMM, 3200MT/s, Dual Rank", "8"),
        ("4", "480GB SSD SATA Read Intensive 6Gbps 512n 2.5in Hot-plug Drive, 1 DWPD", "4"),
        ("5", "H730P RAID Controller, 2GB NV Cache, Adapter, Low Profile", "2"),
        ("6", "Broadcom 5720 Dual Port 1Gb On-Board LOM", "2"),
        ("7", "Broadcom 57416 Dual Port 10Gb Base-T, OCP NIC 3.0", "2"),
        ("8", "Power Cord, C13, 1.8M, 250V, 10A (India BIS, IS1293)", "4"),
        ("9", "Dual, Hot-Plug, Redundant Power Supply (1+1), 750W", "2")
    ]
    
    for sn, desc, qty in server_specs:
        row_cells = table.add_row().cells
        row_cells[0].text = sn
        row_cells[1].text = desc
        row_cells[2].text = qty
        
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                apply_normal_style(paragraph)

def build_scada_section(doc, counter, client_name):
    """Build SCADA section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Falcon's Visual Inspection System (SCADA)", counter=counter)
    
    p = doc.add_paragraph(
        "SCADA stands for Supervisory Control and Data Acquisition. It is a system of hardware "
        "and software components that allows for remote monitoring, control, and data acquisition "
        "of industrial processes or facilities."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        f"The Visualization system provided by FALCON (or SCADA) allows the monitoring and control of "
        f"the different systems delivered for the {client_name}. This SCADA system receives from each "
        f"monitored sub-system all information on their operating status in real time."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph("At the system monitoring level, the functions performed are:")
    apply_normal_style(p)
    
    functions = [
        "Field data acquisition.",
        "Animated visualization of equipment.",
        "Representation of the operating mode of the system (nominal, contingency, etc.).",
        "Alarm management.",
        "Alarm history management.",
        "Diagnostic help.",
        "Failure detection.",
        "Equipment control.",
        "Statistics on equipment operation.",
        "Historical Statistical Report.",
        "Recording and archiving.",
        "Safety operator interface.",
    ]
    
    for item in functions:
        p = doc.add_paragraph(item, style="List Bullet")
        apply_normal_style(p)
    
    add_numbered_subheading(doc, "FIELD DATA ACQUISITION", f"{counter}.1")
    p = doc.add_paragraph(
        "The field data acquisition function is performed by the SCADA system connected to the sorters' PLCs. "
        "The communication with the PLCs is done using equipped CPU cards that are able to manage the "
        "communication with the PLC on the Industrial Ethernet network, without overloading the server."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\main_plc.PNG")
    
    add_numbered_subheading(doc, "ANIMATED SYSTEM VISUALIZATION", f"{counter}.2")
    p = doc.add_paragraph(
        "The animated view represents the dynamic graphical user interface that allows real-time monitoring "
        "of the controlled systems and the execution of their control procedures."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\animated_sys_v1.PNG")
    add_centered_image(doc, "FIXED_IMAGE\\animated_sys_v2.PNG")
    
    add_numbered_subheading(doc, "ALARM MANAGEMENT", f"{counter}.3")
    p = doc.add_paragraph(
        "The alarm pages display a series of information to identify the nature of the alarm or event, "
        "the elements involved and the time."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\alarm.PNG")

def build_key_components_section(doc, counter, components_df):
    """Build Key Components Make section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Key Components Make", counter=counter)
    
    table = doc.add_table(rows=1, cols=2)
    apply_table_style(table)
    
    hdr = table.rows[0].cells
    hdr[0].text = "Items"
    hdr[1].text = "Make"
    
    for run in hdr[0].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
    for run in hdr[1].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
    
    for _, row in components_df.iterrows():
        r = table.add_row().cells
        r[0].text = str(row.get("Items", ""))
        r[1].text = str(row.get("Make", ""))
        
        for cell in r:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(11)

def build_safety_section(doc, counter):
    """Build Principal of Safety section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Principal of Safety", counter=counter)
    
    p = doc.add_paragraph()
    run = p.add_run("1. E-Stops")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\e-stop.png", width_in=4.5)
    
    p = doc.add_paragraph()
    run = p.add_run("            a. At Every Conveyor Module – Both sides")
    apply_normal_style(p)
    add_centered_image(doc, "FIXED_IMAGE\\at-every.png", width_in=4.5)
    
    p = doc.add_paragraph()
    run = p.add_run("             b. VDS Chutes")
    apply_normal_style(p)
    add_centered_image(doc, "FIXED_IMAGE\\emergency-vds.png", width_in=4.5)
    
    p = doc.add_paragraph()
    run = p.add_run("2. Pull Cords Switch")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph("Required for Infeed & Takeout Conveyors")
    apply_normal_style(p)
    
    # Add pull cord images
    table_pc = doc.add_table(rows=1, cols=2)
    if os.path.exists("FIXED_IMAGE\\pull-cords.PNG"):
        p = table_pc.rows[0].cells[0].add_paragraph()
        run = p.add_run()
        run.add_picture("FIXED_IMAGE\\pull-cords.PNG", width=Inches(3))
    
    if os.path.exists("FIXED_IMAGE\\puul-cords-arch.PNG"):
        p = table_pc.rows[0].cells[1].add_paragraph()
        run = p.add_run()
        run.add_picture("FIXED_IMAGE\\puul-cords-arch.PNG", width=Inches(3))
    
    p = doc.add_paragraph()
    run = p.add_run("3. Fencing")
    run.bold = True
    apply_normal_style(p)
    
    table_fence = doc.add_table(rows=1, cols=2)
    table_fence.rows[0].cells[0].text = "a. Between Inducts \nb. Between Inducts and Sorter"
    
    if os.path.exists("FIXED_IMAGE\\fencing.png"):
        p = table_fence.rows[0].cells[1].add_paragraph()
        run = p.add_run()
        run.add_picture("FIXED_IMAGE\\fencing.png", width=Inches(2.5))
    
    p = doc.add_paragraph()
    run = p.add_run("4. Leg Guards")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\leg_gurads.png", width_in=4.5)

def build_infrastructure_section(doc, counter):
    """Build Infrastructure section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Infrastructure", counter=counter)
    
    sections = [
        ("a. Fire Protection-", 
         "Falcon's scope does not cover the design or provision of fire protection infrastructure, "
         "utilities, or related services. It is expected that the customer's sprinkler contractor "
         "will design and supply the in-rack sprinkler systems, including connectors and mounting "
         "brackets. These designs should be submitted to Falcon for review during the engineering phase."),
        
        ("b. Power Supply-",
         "The Customer must provide temporary power for installation and permanent power for "
         "commissioning. Protected multi-gang power points for workstations and peripherals will be "
         "supplied by the Customer, with planning for their locations done with the operations and "
         "IT teams."),
        
        ("c. Floor Requirements-",
         "The Customer must provide flooring with appropriate loading strength and space at the site. "
         "Falcon assumes that the floor slab will not contain corrosive materials that could affect "
         "standard fixings."),
        
        ("d. Estimated Floor Load-",
         "Estimated floor loads, including distributed and point loads, will be provided during the "
         "detailed engineering phase of the project."),
        
        ("e. Staging, Laydown and Assembly Area-",
         "The Customer is required to provide sufficient space on the same floor, adjacent to the "
         "installation site, for staging, storage, and equipment assembly."),
        
        ("f. Site Access and Unloading-",
         "The Customer is required to allocate sufficient on-site space for parking and staging "
         "shipping containers to facilitate Falcon's delivery schedule."),
        
        ("g. Lighting-",
         "All lighting is excluded from Falcon's scope of supply and must be provided by the Customer "
         "or their contractor. This includes lighting for service areas, operational areas, and beneath "
         "platforms and walkways.")
    ]
    
    for title, content in sections:
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        apply_normal_style(p)
        
        p = doc.add_paragraph(content)
        apply_normal_style(p)

def build_program_org_section(doc, counter, client_name, gantt_file):
    """Build Program Organisation section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Program Organisation", counter=counter)
    
    add_numbered_subheading(doc, "Program Schedule", f"{counter}.1")
    
    if gantt_file is not None:
        p = doc.add_paragraph()
        run = p.add_run()
        gantt_file.seek(0)
        run.add_picture(gantt_file, width=Inches(6))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p = doc.add_paragraph("Attach your timeline Gantt chart here.")
        apply_normal_style(p)
    
    add_numbered_subheading(doc, "Program Management", f"{counter}.2")
    
    p = doc.add_paragraph("For this program, proposed approach covers the following aspects:")
    apply_normal_style(p)
    
    bullets = [
        "Creation and monitoring of the project plan.",
        "Weekly/Fortnightly meeting to share project status.",
        "Scheduling of the resource management.",
        "Management of risks and opportunities.",
        "Management of the requirements.",
        "Management of the list of anomalies or reservations.",
    ]
    
    for b in bullets:
        p = doc.add_paragraph(b, style="List Bullet")
        apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The Project will be closely monitored under Falcon's Governance model as structured below."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\governence.png", width_in=5.5)
    doc.add_page_break()
    add_numbered_subheading(doc, "Project Team", f"{counter}.3")
    
    p = doc.add_paragraph(
        f"Team of 3 to 4 member from Projects team will co-ordinate on regular basis with "
        f"{client_name} and internal stakeholders for smooth execution of the project."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    
    run = p.add_run(f"{client_name}'s Team")
    run.bold = True
    run.font.size = Pt(30)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    add_centered_image(doc, "FIXED_IMAGE\\team.png", width_in=5.5)

def build_client_responsibility_section(doc, counter, client_name):
    """Build Client Responsibility section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Client Responsibility", counter=counter)
    
    sections = [
        (f"{client_name} Responsibilities During the Assembly and Commissioning Phase", [
            "Provision of the site complex and office area facilities.",
            "The possibility of authorizing access to the site and the execution of the installation work up to 7 days a week and 24 hours a day if deemed necessary and if requested by FALCON.",
            "Free provision, during the installation phase, of the power supply necessary for the installation activities (estimated at 20 kW).",
            "Provision, during the commissioning phase, of the power supply necessary for the operation of the shipment sorting system free of charge at the date of FALCON need.",
            "Provision of the IT system functionality in accordance with the specification at the date of FALCON need.",
            "The customer is responsible for a safe working environment.",
            "The customer makes arrangements for the working area(s) to be protected against direct weather influences.",
            "The customer provides adequate lighting, heating, and ventilation to create a normal working environment.",
        ]),
        (f"Responsibilities of {client_name} During the Tests", [
            "Provision of the test loads and barcode labels required for the tests.",
            "Provision of personnel required for test activities (loading and unloading operations).",
            "Provision of the necessary information to sort the shipments correctly.",
            "Verify with FALCON the quality and conformity of the test loads (labels, cartons).",
        ]),
        (f"{client_name} Responsibilities During the Training", [
            "Free from their usual work, the employees participate in the training for the duration of the training.",
            "Provision of a list of participants for each available training course 3 days before the start of the course.",
            "Provision of a classroom equipped with a whiteboard, video projector, projection screen, and enough space for desks or tables and chairs for the trainer and trained staff.",
        ])
    ]
    
    for title, bullets in sections:
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        apply_normal_style(p)
        
        for b in bullets:
            p = doc.add_paragraph(b, style="List Bullet")
            apply_normal_style(p)

def build_handover_section(doc, counter):
    """Build System Handover section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "System Handover", counter=counter)
    
    p = doc.add_paragraph(
        "The system handover will follow the workflow shown below. "
        "Each stage is described in the following sections."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\handover.PNG")
    
    sections = [
        ("Installation and Commissioning",
         "Completion of all activities required to bring the system to an operational state "
         "and ready for formal testing."),
        
        ("Pre-UAT",
         "Pre-UAT consists of checks and tests performed before the formal User Acceptance Testing. "
         "It ensures the system is stable, integrated and ready for end-user validation."),
        
        ("UAT (User Acceptance Test)",
         "UAT is performed by the client's users to verify that the solution meets agreed requirements "
         "and behaves as expected under real-world operating conditions."),
        
        ("Minor & Major Faults",
         "Any issues found during testing are categorized as minor or major faults. "
         "Minor faults affect limited areas without blocking successful test completion and are added "
         "to the snag list."),
        
        ("System Snag Points",
         "After UAT, all open minor faults are tracked as system snag points. "
         "Falcon shares a snag list with the client, detailing for each issue: description, date, location, "
         "category, responsible party, target completion date and verification / sign-off."),
        
        ("System Handover Letter",
         "After successful acceptance testing or closure of all snags, Falcon issues a handover letter "
         "confirming that the system has been installed and accepted by the client.")
    ]
    
    for title, content in sections:
        p = doc.add_paragraph()
        run = p.add_run(title + "\n")
        run.bold = True
        apply_normal_style(p)
        
        p = doc.add_paragraph(content)
        apply_normal_style(p)

def build_commercial_section(doc, counter, price_data, payment_terms, apply_bca):
    """Build Commercial section with price sheet"""
    add_numbered_heading(doc, "Commercial", counter=counter)
    
    if not price_data:
        p = doc.add_paragraph("Commercial details to be added.")
        apply_normal_style(p)
        return
    
    # Price Sheet Title
    title = price_data.get("price_sheet_title") or "Price Sheet"
    add_numbered_subheading(doc, title, f"{counter}.1")
    
    items = price_data.get("items", [])
    total_row = price_data.get("total_row")
    
    # Price Table: S. No | Component | Price
    table = doc.add_table(rows=1, cols=3)
    apply_table_style(table)
    
    hdr = table.rows[0].cells
    hdr[0].text = "S. No"
    hdr[1].text = "Component"
    hdr[2].text = "Price"
    
    for run in hdr[0].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
    for run in hdr[1].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
    for run in hdr[2].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
    
    # Add price items
    for item in items:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item.get("s_no", ""))
        row_cells[1].text = str(item.get("label", ""))
        row_cells[2].text = str(item.get("price", ""))
        
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                apply_normal_style(paragraph)
    
    # Total row
    if total_row:
        row_cells = table.add_row().cells
        row_cells[0].text = ""
        row_cells[1].text = str(total_row.get("label", "Total"))
        row_cells[2].text = str(total_row.get("price", ""))
        
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                apply_normal_style(paragraph)
    
    # Optional BCA discount row
    if apply_bca and total_row:
        final_total_str = apply_bca_discount_to_price_data(price_data, 4.5)
        if final_total_str:
            row_cells = table.add_row().cells
            row_cells[0].text = ""
            row_cells[1].text = "Final Total (after 4.5% BCA Discount)"
            row_cells[2].text = final_total_str
            
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                    apply_normal_style(paragraph)
    
    # Payment Terms section
    if payment_terms:
        doc.add_paragraph("")  # spacing
        add_numbered_subheading(doc, "Payment Terms", f"{counter}.2")
        
        pt_table = doc.add_table(rows=1, cols=2)
        apply_table_style(pt_table)
        
        pt_hdr = pt_table.rows[0].cells
        pt_hdr[0].text = "Payment Percentage"
        pt_hdr[1].text = "Stage"
        
        for run in pt_hdr[0].paragraphs[0].runs:
            run.font.bold = True
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)
        for run in pt_hdr[1].paragraphs[0].runs:
            run.font.bold = True
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)
        
        for row in payment_terms:
            perc = str(row.get("Payment Percentage", "")).strip()
            stage = str(row.get("Stage", "")).strip()
            if not perc and not stage:
                continue
            r = pt_table.add_row().cells
            r[0].text = perc
            r[1].text = stage
            
            for cell in r:
                for paragraph in cell.paragraphs:
                    apply_normal_style(paragraph)

def build_warranty_section(doc, counter, warranty_type, duration, start_cond, extended_text, amc_text, transport_text):
    """Build Warranty Period section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Warranty Period", counter=counter)
    
    intro = f"Falcon's offered System comes with a {warranty_type.lower()} of {duration} (starts {start_cond})"
    if not intro.endswith("."):
        intro += "."
    
    if extended_text:
        intro += f" {extended_text}"
    if amc_text:
        intro += f" {amc_text}"
    
    p = doc.add_paragraph(intro)
    apply_normal_style(p)
    
    p = doc.add_paragraph("The warranty covers the following support:")
    apply_normal_style(p)
    
    coverage = [
        "24 X 7 Telephonic, Email and Remote Service Support when required.",
        "Regular Software updates and Bug Fixes.",
        "Supply of Mechanical and Electrical components in case of failure (excluding damages as mentioned in the Exclusion Clause).",
    ]
    
    for item in coverage:
        p = doc.add_paragraph(item, style="List Bullet")
        apply_normal_style(p)
    
    p = doc.add_paragraph("The following items are excluded from warranty:")
    apply_normal_style(p)
    
    exclusions = [
        "Normal wear and tear.",
        "Consumables.",
        "Faulty articles continued.",
        "Failure to comply with the manufacturer's recommendations.",
        "Negligence or abnormal use of equipment.",
    ]
    
    for item in exclusions:
        p = doc.add_paragraph(item, style="List Bullet")
        apply_normal_style(p)
    
    if transport_text:
        p = doc.add_paragraph(transport_text)
        apply_normal_style(p)

def build_exclusions_section(doc, counter, selected_exclusions):
    """Build Exclusions section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Exclusions", counter=counter)
    
    intro = (
        "The scope of supply includes all parts which are defined in the Supplier's quotation.\n"
        "All other parts which are not defined in the Supplier's quotation do not belong to the Supplier's "
        "scope of supply and are excluded. The following parts are also excluded:"
    )
    
    p = doc.add_paragraph(intro)
    apply_normal_style(p)
    
    fixed_exclusions = [
        "Construction Power",
        "Building infrastructure; building structure, doors, fire exits, levelling devices, "
        "building extinguisher and fire alarm system, building heating and lighting system.",
        "Electrical power supply and wiring to the main control cabinets.",
        "UPS for Controls and Drives",
        "Network cabling up to the main server rack.",
        "Intermediate wiring to parts which are to be supplied by the Purchaser/others.",
        "Emergency/Uninterruptable power supply.",
        "Fire-alarm and fire protection devices.",
        "Traffic and route markings.",
        "Laydown area / unloading and laydown area.",
        "Ram protection devices.",
        "Cat walks, bridges, maintenance aisles and platforms.",
        "All kind of network incl. Local Area Network (LAN/WLAN), exceeding the scope described in Scope of Supply.",
        "Any kind of civil work.",
        "Any adjustment of the Supplier's scope of supply to local rules and regulations.",
        "X-Ray machines.",
        "Roller cages / pallets.",
        "Simulation and 3D animation of the sorter system.",
        "Interface with other equipment not specified in this offer.",
        "Provision of facilities for the control room (furniture, air conditioning, heating, etc.).",
        "The supply and installation of fencing around the different corridors.",
        "Any item specifically indicated as not forming part of the subject matter of the Seller's supply in the offer documentation.",
    ]
    
    all_exclusions = fixed_exclusions + selected_exclusions
    
    for item in all_exclusions:
        p = doc.add_paragraph(item, style="List Bullet")
        apply_normal_style(p)

# ==================== MAIN GENERATION ====================

st.markdown("---")
st.header("Generate Complete Document")

if st.button("Generate Final DOCX Document", type="primary", width='stretch'):
    with st.spinner("Generating your complete proposal document..."):
        try:
            # Validate required inputs
            if not client_name.strip():
                st.error("Please enter a Client Name")
                st.stop()
            
            if not project_name.strip():
                st.error("Please enter a Project Name")
                st.stop()
            
            # Generate cover letter if needed
            cover_letter_text = None
            if offer_ref and sender_name:
                with st.spinner("Generating cover letter with AI..."):
                    try:
                        cover_letter_text = call_groq_cover_letter(
                            client_name=client_name,
                            project_title=project_name,
                            offer_ref=offer_ref,
                            letter_date_str=letter_date.strftime("%B %d, %Y"),
                            executives_block=executives_text,
                            invitation_date=invitation_date_str,
                            meeting_date=meeting_date_str,
                            sender_name=sender_name,
                            sender_title=sender_title
                        )
                        st.success("Cover letter generated successfully")
                    except Exception as e:
                        st.warning(f"Could not generate cover letter: {str(e)}")
                        cover_letter_text = None
            
            # Generate executive summary if needed
            exec_summary_text = None
            if include_exec_summary and exec_summary_pdf:
                with st.spinner("📝 Generating executive summary with AI..."):
                    try:
                        system_text = extract_pdf_text(exec_summary_pdf)
                        exec_summary_text = call_groq_exec_summary(system_text, client_name, project_name)
                        st.success("Executive summary generated successfully")
                    except Exception as e:
                        st.warning(f"Could not generate executive summary: {str(e)}")
                        exec_summary_text = None
            
            # Process costing file if commercial section is included
            price_data = None
            payment_terms_data = None
            bca_discount = False
            
            if commercial_include and costing_file:
                with st.spinner("📊 Processing costing file with AI..."):
                    try:
                        # Read Overall Costing sheet
                        df = pd.read_excel(costing_file, sheet_name="Overall Costing", header=None)
                        sheet_csv = df.to_csv(index=False)
                        
                        # Call Groq to extract price sheet
                        price_data = call_groq_for_price_sheet(sheet_csv)
                        payment_terms_data = st.session_state.get("payment_terms", [])
                        bca_discount = apply_bca
                        
                        st.success("Price sheet generated from costing file")
                    except Exception as e:
                        st.warning(f"Could not process costing file: {str(e)}. Commercial section will be added as placeholder.")
                        price_data = None
            
            # Save uploaded client logo temporarily
            client_logo_path = None
            if client_logo:
                client_logo_path = f"temp_client_logo.{client_logo.name.split('.')[-1]}"
                with open(client_logo_path, "wb") as f:
                    f.write(client_logo.getbuffer())
            
            doc = Document()
            
            # Set default font for the document
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Calibri (Body)'
            font.size = Pt(11)
            
            # ==================== COVER LETTER (NO HEADER) ====================
            if cover_letter_text:
                build_cover_letter_section(doc, cover_letter_text)
            
            # ==================== FRONT PAGE (NO HEADER) ====================
            if cover_letter_text:
                build_front_page_section(doc, project_name, offer_ref, contact_name, contact_email, contact_phone, layout_image)
            
            # ==================== NOW ADD HEADER/FOOTER FOR ALL REMAINING SECTIONS ====================
            create_header_footer(doc, client_name, project_name, None, client_logo_path)
            
            # ==================== GLOSSARY ====================
            build_glossary_section(doc)
            
            # Start numbering from 1
            counter = 1
            
            # ==================== BUILD SECTIONS IN ORDER ====================
            
            # 1. Executive Summary
            if include_exec_summary and exec_summary_text:
                build_executive_summary_section(doc, exec_summary_text, counter)
                counter += 1

            # 2. Company Profile
            if include_company_profile:
                build_company_profile_section(doc, counter)
                counter += 1

            # 3. Reference Projects
            if include_ref_projects:
                build_reference_projects_section(doc, counter)
                counter += 1

            # 4. Handled Shipment Spectrum
            if include_handled_spectrum:
                build_handled_spectrum_section(doc, counter, project_name, client_name)
                counter += 1

            # 4b. Capacity Calculations Section
            if include_capacity_section and capacity_excel is not None:
                build_capacity_calculations_section(doc, counter, client_name, project_name, capacity_excel)
                counter += 1

            # 5. Electrical System
            if elec_include:
                build_electrical_section(doc, counter)
                counter += 1
            
            # 6. Falcon WCS CONTROLIT
            if wcs_include:
                build_wcs_section(doc, counter, client_name)
                counter += 1
            
            # 7. Falcon Visual Inspection System (SCADA)
            if scada_include:
                build_scada_section(doc, counter, client_name)
                counter += 1
            
            # 8. Key Components Make
            if key_include:
                build_key_components_section(doc, counter, key_components_edited)
                counter += 1
            
            # 9. Principal of Safety
            if safety_include:
                build_safety_section(doc, counter)
                counter += 1
            
            # 10. Infrastructure
            if infra_include:
                build_infrastructure_section(doc, counter)
                counter += 1
            
            # 11. Program Organisation
            if prog_include:
                build_program_org_section(doc, counter, client_name, prog_gantt)
                counter += 1
            
            # 12. Client Responsibility
            if client_resp_include:
                build_client_responsibility_section(doc, counter, client_name)
                counter += 1
            
            # 13. System Handover
            if handover_include:
                build_handover_section(doc, counter)
                counter += 1
            
            # 14. Commercial
            if commercial_include:
                build_commercial_section(doc, counter, price_data, payment_terms_data, bca_discount)
                counter += 1
            
            # 15. Warranty Period
            if warranty_include:
                build_warranty_section(doc, counter, warranty_type, warranty_duration, 
                                      warranty_start, warranty_extended_text, 
                                      warranty_amc_text, warranty_transport_text)
                counter += 1
            
            # 16. Exclusions
            if exclusion_include:
                build_exclusions_section(doc, counter, selected_exclusions)
                counter += 1
            
            # Save to buffer
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            # Clean up temporary logo files
            # No need to remove falcon_logo_path, logo is fixed from backend
            if client_logo_path and os.path.exists(client_logo_path):
                os.remove(client_logo_path)
            
            st.success("Document generated successfully!")
            
            st.download_button(
                label="📥 Download Complete Proposal Document",
                data=buffer,
                file_name=f"Falcon_Proposal_{client_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width='stretch'
            )
            
        except Exception as e:
            st.error(f"Error generating document: {str(e)}")
            st.exception(e)
            # Clean up temporary logo files in case of error
            try:
                # No need to remove falcon_logo_path, logo is fixed from backend
                if 'client_logo_path' in locals() and client_logo_path and os.path.exists(client_logo_path):
                    os.remove(client_logo_path)
            except:
                pass

st.markdown("---")
st.info("💡 **Tip:** Make sure all required images are present in the FIXED_IMAGE folder before generating the document.")

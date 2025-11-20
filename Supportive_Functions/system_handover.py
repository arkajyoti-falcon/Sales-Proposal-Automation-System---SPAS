import streamlit as st
from io import BytesIO
from docx import Document
from docx.shared import Inches
from PIL import Image

# ---------- CONFIG ----------
DIAGRAM_PATH = "FIXED_IMAGE\handover.PNG"  # put your PNG here (same as the PDF diagram)

st.set_page_config(page_title="System Handover Section", page_icon="✅")

st.title("System Handover – Section Builder")

st.write(
    "This section shows the fixed System Handover workflow and description. "
    "Click the button below to download it as a formatted DOCX section."
)

# ---------- FIXED TEXT CONTENT (PARAPHRased FROM YOUR DOC) ----------

INTRO_LINE = (
    "The system handover will follow the workflow shown below. "
    "Each stage is described in the following sections."
)

INSTALLATION_TITLE = "Installation and Commissioning"
INSTALLATION_TEXT = (
    "Completion of all activities required to bring the system to an operational state "
    "and ready for formal testing."
)

PRE_UAT_TITLE = "Pre-UAT"
PRE_UAT_TEXT = (
    "Pre-UAT consists of checks and tests performed before the formal User Acceptance Testing. "
    "It ensures the system is stable, integrated and ready for end-user validation."
)
PRE_UAT_BULLETS = [
    "Physical verification of installed equipment against requirements and specifications.",
    "Integration testing of conveyors, scanners, sorter, controls and interfaces with WMS / databases.",
    "Performance checks with trial loads to confirm the system can handle expected throughput.",
    "Functional checks across key scenarios including error handling and edge conditions.",
]

UAT_TITLE = "UAT (User Acceptance Test)"
UAT_TEXT = (
    "UAT is performed by the client’s users to verify that the solution meets agreed requirements "
    "and behaves as expected under real-world operating conditions. "
    "Falcon will share a UAT plan in advance covering prerequisites, daily schedule, "
    "roles and responsibilities, test loads and expected results."
)
UAT_PARAMS_TITLE = "User Acceptance Test Parameters"
UAT_PARAMS_BULLETS = [
    "Sorter throughput tests.",
    "Sorter accuracy tests.",
    "Barcode reading / scanning tests.",
]

FAULTS_TITLE = "Minor & Major Faults"
FAULTS_TEXT = (
    "Any issues found during testing are categorized as minor or major faults. "
    "Minor faults affect limited areas without blocking successful test completion and are added "
    "to the snag list. Major faults prevent demonstration of required functionality and may require "
    "tests to be restarted after rectification."
)

SNAG_TITLE = "System Snag Points"
SNAG_TEXT = (
    "After UAT, all open minor faults are tracked as system snag points. "
    "Falcon shares a snag list with the client, detailing for each issue: description, date, location, "
    "category, responsible party, target completion date and verification / sign-off."
)
SNAG_EXTRA = (
    "Once the snag list is published, the client can begin beneficiary use of the system and "
    "is requested to sign off on start of system warranty and conditional handover."
)

SNAG_PRACTICE_TITLE = "Practice for Snag Point Sign-Off"
SNAG_PRACTICE_BULLETS = [
    "Falcon maintains and circulates the snag list throughout UAT.",
    "For each resolved issue, the client verifies closure and signs off on the snag list.",
]

HANDOVER_LETTER_TITLE = "System Handover Letter"
HANDOVER_LETTER_TEXT = (
    "After successful acceptance testing or closure of all snags, Falcon issues a handover letter "
    "confirming that the system has been installed and accepted by the client, and indicating "
    "the final payment amount and due date."
)


# ---------- DOCX BUILDER ----------

def build_system_handover_docx(diagram_path: str) -> BytesIO:
    doc = Document()

    # Heading 2: System Handover
    doc.add_heading("System Handover", level=2)

    # Intro line
    doc.add_paragraph(INTRO_LINE)

    # Workflow image
    try:
        doc.add_picture(diagram_path, width=Inches(6.0))
    except Exception:
        # If picture not found, just leave a placeholder line
        doc.add_paragraph("[System handover workflow diagram]")

    # Installation & Commissioning
    p = doc.add_paragraph()
    p.add_run(INSTALLATION_TITLE + "\n").bold = True
    doc.add_paragraph(INSTALLATION_TEXT)

    # Pre-UAT
    p = doc.add_paragraph()
    p.add_run(PRE_UAT_TITLE + "\n").bold = True
    doc.add_paragraph(PRE_UAT_TEXT)
    for item in PRE_UAT_BULLETS:
        doc.add_paragraph(item, style="List Bullet")

    # UAT
    p = doc.add_paragraph()
    p.add_run(UAT_TITLE + "\n").bold = True
    doc.add_paragraph(UAT_TEXT)

    # UAT parameters
    p = doc.add_paragraph()
    p.add_run(UAT_PARAMS_TITLE + "\n").bold = True
    for item in UAT_PARAMS_BULLETS:
        doc.add_paragraph(item, style="List Bullet")

    # Minor & Major faults
    p = doc.add_paragraph()
    p.add_run(FAULTS_TITLE + "\n").bold = True
    doc.add_paragraph(FAULTS_TEXT)

    # System Snag Points
    p = doc.add_paragraph()
    p.add_run(SNAG_TITLE + "\n").bold = True
    doc.add_paragraph(SNAG_TEXT)
    doc.add_paragraph(SNAG_EXTRA)

    # Practice for Snag Sign-off
    p = doc.add_paragraph()
    p.add_run(SNAG_PRACTICE_TITLE + "\n").bold = True
    for item in SNAG_PRACTICE_BULLETS:
        doc.add_paragraph(item, style="List Bullet")

    # System Handover Letter
    p = doc.add_paragraph()
    p.add_run(HANDOVER_LETTER_TITLE + "\n").bold = True
    doc.add_paragraph(HANDOVER_LETTER_TEXT)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------- STREAMLIT UI ----------

st.header("System Handover Section Preview")

# Show diagram if available
try:
    img = Image.open(DIAGRAM_PATH)
    st.image(img, caption="System Handover Workflow", use_container_width=True)
except Exception:
    st.warning("Workflow diagram not found. Put it as 'System_Handover.png' next to this script.")

st.markdown(f"**{INSTALLATION_TITLE}**")
st.write(INSTALLATION_TEXT)

st.markdown(f"**{PRE_UAT_TITLE}**")
st.write(PRE_UAT_TEXT)
st.markdown("- " + "\n- ".join(PRE_UAT_BULLETS))

st.markdown(f"**{UAT_TITLE}**")
st.write(UAT_TEXT)

st.markdown(f"**{UAT_PARAMS_TITLE}**")
st.markdown("- " + "\n- ".join(UAT_PARAMS_BULLETS))

st.markdown(f"**{FAULTS_TITLE}**")
st.write(FAULTS_TEXT)

st.markdown(f"**{SNAG_TITLE}**")
st.write(SNAG_TEXT)
st.write(SNAG_EXTRA)

st.markdown(f"**{SNAG_PRACTICE_TITLE}**")
st.markdown("- " + "\n- ".join(SNAG_PRACTICE_BULLETS))

st.markdown(f"**{HANDOVER_LETTER_TITLE}**")
st.write(HANDOVER_LETTER_TEXT)

st.markdown("---")
st.subheader("Download System Handover DOCX")

docx_buffer = build_system_handover_docx(DIAGRAM_PATH)
st.download_button(
    label="Generate & Download System Handover (.docx)",
    data=docx_buffer,
    file_name="System_Handover.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)

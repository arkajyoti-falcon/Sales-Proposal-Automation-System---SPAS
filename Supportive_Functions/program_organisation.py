import streamlit as st
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Path to governance model image on backend
GOVERNANCE_IMAGE_PATH = "FIXED_IMAGE\\governence.png"  # put this file next to your script
PROJECT_TEAM_IMAGE_PATH = "FIXED_IMAGE\\team.png"  # put this file next to your script
st.set_page_config(page_title="Program Organisation Builder", page_icon="📅")

st.title("Program Organisation – Section Builder")

st.write(
    "Upload an optional Gantt chart image and enter the client name. "
    "Click **Generate** to create a DOCX section for Program Organisation."
)

# --- UI Inputs ---
client_input = st.text_input("Client name (optional)", value="Zepto")
client_name = client_input.strip() or "Client"

gantt_file = st.file_uploader(
    "Upload Gantt chart image (optional)",
    type=["png", "jpg", "jpeg"],
    help="This will be placed under 'Program Schedule'. If you skip, a placeholder text will be added."
)


def build_program_org_docx(client: str, gantt_upload) -> bytes:
    doc = Document()

    # Top-level heading
    doc.add_heading("Program Organisation", level=2)

    # -------------------------
    # 1. Program Schedule
    # -------------------------
    doc.add_heading("Program Schedule", level=3)

    if gantt_upload is not None:
        # Insert uploaded Gantt chart centered
        p = doc.add_paragraph()
        run = p.add_run()
        gantt_upload.seek(0)
        run.add_picture(gantt_upload, width=Inches(6))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph("Attach your timeline Gantt chart here.")

    doc.add_paragraph("")  # spacing

    # -------------------------
    # 2. Program Management
    # -------------------------
    doc.add_heading("Program Management", level=3)

    doc.add_paragraph("For this program, proposed approach covers the following aspects:")

    bullets = [
        "Creation and monitoring of the project plan.",
        "Weekly/Fortnightly meeting to share project status.",
        "Scheduling of the resource management.",
        "Management of risks and opportunities.",
        "Management of the requirements.",
        "Management of the list of anomalies or reservations.",
    ]
    for b in bullets:
        doc.add_paragraph(b, style="List Bullet")

    doc.add_paragraph(
        "The Project will be closely monitored under Falcon’s Governance model as structured below."
    )

    # Governance image in center (or placeholder if missing)
    try:
        p = doc.add_paragraph()
        run = p.add_run()
        run.add_picture(GOVERNANCE_IMAGE_PATH, width=Inches(5.5))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        p = doc.add_paragraph("[Governance model diagram]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")  # spacing

    # -------------------------
    # 3. Project Team
    # -------------------------
    doc.add_heading("Project Team", level=3)

    doc.add_paragraph(
        f"Team of 3 to 4 member from Projects team will co-ordinate on regular basis with "
        f"{client} and internal stakeholders for smooth execution of the project."
    )

    # Client_name's Team – centered, big font
    p = doc.add_paragraph()
    run = p.add_run(f"{client}'s Team")
    run.bold = True
    run.font.size = Pt(30)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Placeholder for client team image (center)
    p2 = run.add_picture(PROJECT_TEAM_IMAGE_PATH, width=Inches(5.5))
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER


    doc.add_paragraph(
        
        f"Falcon Project Manager is overseen and supported by Executive Sponsors. The Project Manager leads a multi-disciplinary team of members from all Falcon departments necessary to deliver the project, as shown in the diagram above."
    )

    # Export as bytes
    # -------------------------
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# --- Generate + Download logic ---
if st.button("Generate Program Organisation DOCX"):
    doc_bytes = build_program_org_docx(client_name, gantt_file)
    st.session_state["program_org_doc"] = doc_bytes
    st.success("DOCX generated. You can download it below.")

if "program_org_doc" in st.session_state:
    st.download_button(
        label="Download Program Organisation (.docx)",
        data=st.session_state["program_org_doc"],
        file_name=f"Program_Organisation_{client_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


import io
import os
import json

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches

# -------------------------------------------------
# ENV / GROQ
# -------------------------------------------------
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

# -------------------------------------------------
# SAMPLE IMAGE PATHS (replace with real paths)
# -------------------------------------------------
CROSS_BELT_SORTER_IMG = r"FIXED_IMAGE\CROS_BELT_SORTER.PNG"
CBS_CARRIER_IMG       = r"FIXED_IMAGE\CBS_CAREER.PNG"
SERVO_IMG             = r"FIXED_IMAGE\SERVO_ROLLER.PNG"
CHASSIS_IMG           = r"FIXED_IMAGE\CHASIS.PNG"
WHEEL_IMG             = r"FIXED_IMAGE\CAREER_WHEEL.PNG"
POWER_IMG             = r"FIXED_IMAGE\TRANSMISSION.PNG"
LINEAR_IMG            = r"FIXED_IMAGE\LINEAR_MOTOR_DRIVE.PNG"
FRICTION_WHEEL_IMG    = r"FIXED_IMAGE\FRICTION_WHEEL_DRIVE.PNG"
RCOAX_IMG             = r"FIXED_IMAGE\DATA.PNG"
CARRIER_POSITION_IMG  = r"FIXED_IMAGE\CPS.PNG"

# If you are passing these paths from backend, you can
# overwrite the above variables before calling build_full_doc().

# -------------------------------------------------
# STREAMLIT UI CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Description of Components + Sorter Technical Spec",
    layout="centered",
)
st.title("Description of Components of Equipment")
st.caption("Fixed components description + Technical Specification of Sorter from Loop CBS Excel.")


# -------------------------------------------------
# DOCX builder helpers
# -------------------------------------------------
def style_heading(heading, level: int):
    """Apply Calibri font & size similar to your proposals."""
    if heading is None:
        return
    for r in heading.runs:
        r.font.name = "Calibri"
        if level == 1:
            r.font.size = Pt(14)
        elif level == 2:
            r.font.size = Pt(12)
        else:
            r.font.size = Pt(11)


def add_paragraph(doc: Document, text: str):
    """Add normal body paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(11)


def add_italic_paragraph(doc: Document, text: str):
    """Add italic paragraph (for captions/placeholders)."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.italic = True
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)


def add_image_if_exists(doc: Document, img_path: str | None, caption: str, width_inch: float = 3.0):
    """
    Insert an image centered; if the path doesn't exist, insert an italic placeholder line.
    """
    if img_path and os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(img_path, width=Inches(width_inch))

        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_cap = cap.add_run(caption)
        run_cap.italic = True
        run_cap.font.name = "Calibri"
        run_cap.font.size = Pt(10.5)

        doc.add_paragraph("")  # spacing
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"[Image Placeholder]")
        run.italic = True
        run.font.name = "Calibri"
        run.font.size = Pt(10.5)
        doc.add_paragraph("")


# -------------------------------------------------
# EXCEL → TEXT (ENTIRE Loop CBS SHEET)
# -------------------------------------------------
def load_loop_cbs_sheet_from_excel(xlsx_bytes: bytes) -> tuple[str | None, pd.DataFrame | None]:
    """
    Find sheet whose name is 'Loop CBS' or 'Loop CBS upper' (case/space-insensitive)
    and return (sheet_name, DataFrame). Reads the FULL sheet (no cropping).
    """
    with pd.ExcelFile(io.BytesIO(xlsx_bytes)) as xls:
        target_name = None
        for sname in xls.sheet_names:
            normalized = sname.lower().replace(" ", "")
            if normalized == "loopcbs" or normalized == "loopcbsupper":
                target_name = sname
                break

        if target_name is None:
            return None, None

        # Read entire sheet as generic table (no header), so we keep all rows/blocks.
        df = xls.parse(target_name, header=None)
        return target_name, df


def df_to_compact_text(df: pd.DataFrame, max_rows: int = 200, max_cols: int = 20) -> str:
    """
    Convert the ENTIRE sheet (up to max_rows, max_cols) into a compact text representation
    including BOTH tables in Loop CBS sheet.
    """
    if df is None:
        return ""

    df2 = df.iloc[:max_rows, :max_cols].fillna("")

    lines: list[str] = []
    for idx in range(df2.shape[0]):
        row_vals = [str(v).strip() for v in df2.iloc[idx].tolist()]
        # keep only non-empty cells in the print
        row_vals = [v for v in row_vals if v != ""]
        if row_vals:
            lines.append(" | ".join(row_vals))

    return "\n".join(lines)


# -------------------------------------------------
# GROQ – Technical Specification of Sorter
# -------------------------------------------------
def call_groq_for_sorter_spec(sheet_name: str, sheet_text: str) -> dict:
    """
    Send the FULL Loop CBS sheet content (as text) to GROQ and get back
    a JSON object with the fields:

      sorter_carrier_type      (string, default 'Loop CBS')
      sorter_speed_mps         (string or null)
      sorter_loop_length_m     (string or null)
      sorter_height_mm         (string or null)
      actuation_technology     (string, default 'Electric')
      carrier_pitch_mm         (string or null)
      number_of_carriers       (string or null)
      motor_drive_type         (string or null)
      power_consumption        (string or null)

    We strictly tell the model NOT to invent values.
    """
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY is not set in environment / .env")

    system_prompt = """
You are a technical proposal engineer reading an Excel costing/configuration sheet
for a Loop Cross Belt Sorter ("Loop CBS").

You will receive:
- The sheet name (e.g., "Loop CBS")
- The ENTIRE sheet content as text, row by row, including ALL tables.

Your task:
Extract the following fields, strictly from the sheet content:

1) sorter_carrier_type        – fixed string "Loop CBS".
2) sorter_speed_mps           – fixed string "upto 2 m/s"
3) sorter_loop_length_m       – sorter loop length in meters (e.g., "150").
4) sorter_height_mm           – sorter height in mm (e.g., "2900").
5) actuation_technology       – fixed string "Electric".
6) carrier_pitch_mm           – carrier pitch in mm (e.g., "600", "1175", "1200"). If multiple models,
                                choose the one actually selected in the configuration area.
7) number_of_carriers         – total number of carriers in the sorter, as shown in the sheet if present.
8) motor_drive_type           – description of motor/drive type (e.g., "LIM", "LSM", "LIM + LSM", etc.).
9) power_consumption          – sorter power consumption (kW or kVA etc.) taken from the sheet.

VERY IMPORTANT RULES:
- Use ONLY information present in the provided sheet text. Do NOT guess or invent values.
- If a value is not clearly present, set it to null.
- If the sheet expresses a choice ("Select Model", "Enter Loop Length", etc.), use the chosen values.
- Keep the output values SHORT: just the numeric value or the short phrase, without explanations.

Return a single JSON object with EXACTLY these keys:

{
  "sorter_carrier_type": "...",
  "sorter_speed_mps": "... or null",
  "sorter_loop_length_m": "... or null",
  "sorter_height_mm": "... or null",
  "actuation_technology": "...",
  "carrier_pitch_mm": "... or null",
  "number_of_carriers": "... or null",
  "motor_drive_type": "... or null",
  "power_consumption": "... or null"
}
DO NOT ADD ``json``` or ``` or json in response.Only return the RAW JSON object.
No comments, no trailing text, no markdown.
"""

    user_payload = {
        "sheet_name": sheet_name,
        "sheet_text": sheet_text,
    }

    payload = {
        "model": "llama-3.3-70b-versatile",
        "temperature": 0.1,
        "max_tokens": 600,
        "messages": [
            {"role": "system", "content": system_prompt.strip()},
            {"role": "user", "content": json.dumps(user_payload, indent=2)},
        ],
    }

    resp = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()
    text = data["choices"][0]["message"]["content"].strip()

    try:
        spec = json.loads(text)
        if not isinstance(spec, dict):
            raise ValueError("Expected a JSON object")
    except Exception as exc:
        # For debug in UI, show what we got
        st.error("Failed to parse GROQ JSON for sorter specification.")
        st.code(text, language="json")
        raise exc

    # Fill fixed fields if missing
    if not spec.get("sorter_carrier_type"):
        spec["sorter_carrier_type"] = "Loop CBS"
    if not spec.get("actuation_technology"):
        spec["actuation_technology"] = "Electric"

    return spec


# -------------------------------------------------
# DOCX builder: Description of Components + Sorter Spec
# -------------------------------------------------
def build_full_doc(sorter_spec: dict | None) -> bytes:
    doc = Document()

    # Main heading
    h_main = doc.add_heading("Description of Components of Equipment", level=1)
    style_heading(h_main, level=1)
    doc.add_paragraph("")

    # Elements of the sorting system
    h_elements = doc.add_heading("Elements of the sorting system", level=2)
    style_heading(h_elements, level=2)

    add_paragraph(doc, "1. Cross Belt Sorter")
    add_paragraph(doc, "2. Scanning & Sensing on Sorter")
    add_paragraph(doc, "3. Conveyor System")
    add_paragraph(doc, "4. Steel Works – Mezzanine & Staircases")
    doc.add_paragraph("")

    # Cross Belt Sorter
    h_cbs = doc.add_heading("Cross Belt Sorter", level=2)
    style_heading(h_cbs, level=2)

    add_paragraph(
        doc,
        "Cross Belt Sorter is capable of sorting extremely high volume of versatile products in a gentle manner."
    )
    add_paragraph(
        doc,
        "Falcon’s Cross belt sorter is powered by high efficiency linear motors and is based on 100% non-touch "
        "actuation technology leading to high throughput capabilities with extremely low noise levels."
    )
    add_paragraph(
        doc,
        "Falcon’s Cross belt sorter is modular in design. It can be easily extended as per future requirements."
    )

    add_image_if_exists(doc, CROSS_BELT_SORTER_IMG, caption="")

    # CBS Carrier
    h_carrier = doc.add_heading("CBS Carrier", level=2)
    style_heading(h_carrier, level=2)

    add_paragraph(
        doc,
        "Falcon Autotech’s CBS offers one of the highest belt width to carrier pitch ratios in the market today."
    )
    add_paragraph(
        doc,
        "This additional belt width makes the system capable of handling larger product sizes without "
        "compromising on throughput. It also reduces the dead area between carrier belts, significantly reducing "
        "the number of in-betweeners and non-sortable parcel recirculation."
    )
    add_image_if_exists(doc, CBS_CARRIER_IMG, caption="")
    doc.add_paragraph("")

    # Servo Roller
    h_servo = doc.add_heading("Servo Roller", level=3)
    style_heading(h_servo, level=3)

    add_paragraph(
        doc,
        "High powered DC drive servo rollers are used to actuate the carrier belts, thereby eliminating the "
        "need for complicated drive transfer mechanisms and simplifying system installation and maintenance."
    )
    add_image_if_exists(doc, SERVO_IMG, caption="")
    doc.add_paragraph("")

    # Chassis
    h_chassis = doc.add_heading("Chassis", level=3)
    style_heading(h_chassis, level=3)

    add_paragraph(
        doc,
        "Falcon Autotech’s cross belt carrier chassis is made up of lightweight aluminium, which makes it "
        "light yet sturdy. This reduced weight leads to substantial power savings over a considerable period of usage."
    )
    add_image_if_exists(doc, CHASSIS_IMG, caption="")
    doc.add_paragraph("")

    # Carrier Wheels
    h_wheels = doc.add_heading("Carrier Wheels", level=3)
    style_heading(h_wheels, level=3)

    add_paragraph(
        doc,
        "Carrier wheels are thoroughly tested and proven for a long life cycle."
    )
    add_image_if_exists(doc, WHEEL_IMG, caption="")
    doc.add_paragraph("")

    # Friction Wheel Drive
    h_fwd = doc.add_heading("Friction Wheel Drive", level=3)
    style_heading(h_fwd, level=3)

    add_paragraph(
        doc,
        "Falcon can provide the indigenously developed Friction Wheel Drive (FWD) to drive the cross belt loop. "
        "This driving mechanism operates on the principle of friction. The unit comprises two independent "
        "motor-driven wheels that spin in opposite directions, synchronised."
    )
    add_paragraph(
        doc,
        "FWDs are highly energy-efficient drives that promote sustainability compared to traditional linear "
        "induction or synchronous motor drives."
    )
    add_image_if_exists(doc, FRICTION_WHEEL_IMG, caption="")

    # Non-contact based linear motor drive
    h_linear = doc.add_heading("Non-contact based linear motor drive", level=3)
    style_heading(h_linear, level=3)

    add_paragraph(
        doc,
        "Non-contact based linear induction motors can be configured at variable speeds depending upon "
        "operational requirements, providing maximum flexibility."
    )
    add_image_if_exists(doc, LINEAR_IMG, caption="")
    add_paragraph(
        doc,
        "The customer can choose the preferred drive system based on performance and energy-efficiency requirements."
    )
    doc.add_paragraph("")

    # Power Transmission
    h_power = doc.add_heading("Power Transmission", level=3)
    style_heading(h_power, level=3)

    add_paragraph(
        doc,
        "Power transmission to carriers is provided over sliding contacts that require low maintenance "
        "and offer high levels of reliability."
    )
    add_image_if_exists(doc, POWER_IMG, caption="")
    doc.add_paragraph("")

    # Data Transmission
    h_data = doc.add_heading("Data Transmission", level=3)
    style_heading(h_data, level=3)

    add_paragraph(
        doc,
        "The R-Coax cable is used for data distribution in the sorter. This leaky wave cable runs throughout "
        "the sorter length, transmitting data continuously. An antenna mounted on the super master carriage "
        "receives the signal from this cable while on the move, wirelessly."
    )
    add_image_if_exists(doc, RCOAX_IMG, caption="")

    # Carriers positioning system
    h_pos = doc.add_heading("Carriers positioning system", level=3)
    style_heading(h_pos, level=3)

    add_paragraph(
        doc,
        "The positioning system determines the exact location of each carrier in the loop at any given point in time."
    )
    add_paragraph(
        doc,
        "A plastic tape strip of barcodes (QR codes) runs along the sorter loop, which is continuously scanned "
        "by scanners placed on a master carrier to track and control the position of every carrier."
    )
    add_image_if_exists(doc, CARRIER_POSITION_IMG, caption="")

    # -------------------------------------------------
    # Technical Specification of Sorter (from sorter_spec)
    # -------------------------------------------------
    doc.add_page_break()

    h_ts = doc.add_heading("Technical Specification of Sorter", level=2)
    style_heading(h_ts, level=2)

    if sorter_spec is None:
        add_paragraph(
            doc,
            "Technical specification of the sorter will be finalised based on project-specific configuration."
        )
    else:
        # Safely get values (fallbacks)
        def gv(key: str, default: str = "-") -> str:
            val = sorter_spec.get(key)
            if val is None:
                return default
            s = str(val).strip()
            return s if s else default

        rows = [
            ("Sorter Carrier Type", gv("sorter_carrier_type", "Loop CBS")),
            ("Sorter Speed (m/s)", gv("sorter_speed_mps")),
            ("Sorter Loop Length (m)", gv("sorter_loop_length_m")),
            ("Sorter Height (mm)", gv("sorter_height_mm")),
            ("Sorter Actuation Technology", gv("actuation_technology", "Electric")),
            ("Carrier Pitch (mm)", gv("carrier_pitch_mm")),
            ("Number of Carriers", gv("number_of_carriers")),
            ("Motor / Drive Type", gv("motor_drive_type")),
            ("Power Consumption*", gv("power_consumption")),
        ]

        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Technical Parameter"
        hdr_cells[1].text = "Value"

        for param, value in rows:
            row_cells = table.add_row().cells
            row_cells[0].text = param
            row_cells[1].text = value

        doc.add_paragraph("")

    # Finalise DOCX -> bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# -------------------------------------------------
# STREAMLIT UI – Excel upload & DOCX generation
# -------------------------------------------------
excel_file = st.file_uploader(
    "Upload Loop CBS Excel (for Technical Specification of Sorter)",
    type=["xlsx"],
)

if st.button("Generate DOCX (Components + Sorter Spec)"):
    sorter_spec = None

    if excel_file is not None:
        try:
            sheet_name, df = load_loop_cbs_sheet_from_excel(excel_file.getvalue())
            if sheet_name is None or df is None:
                st.error("No sheet named 'Loop CBS' or 'Loop CBS upper' found in the Excel file.")
            else:
                st.write(f"Using sheet: {sheet_name}")
                sheet_text = df_to_compact_text(df)
                st.text_area("Loop CBS sheet (compact text sent to GROQ):", sheet_text, height=250)

                try:
                    sorter_spec = call_groq_for_sorter_spec(sheet_name, sheet_text)
                    st.json(sorter_spec)
                except Exception as e:
                    st.error(f"GROQ call for sorter technical specification failed: {e}")
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")

    else:
        st.info("No Excel uploaded. Technical specification table will be a generic placeholder.")

    # Build DOCX (always) – sorter_spec may be None if Excel/GROQ failed
    try:
        doc_bytes = build_full_doc(sorter_spec)
        st.success("DOCX generated successfully.")

        st.download_button(
            "Download DOCX",
            data=doc_bytes,
            file_name="Description_of_Components_and_Sorter_Spec.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"Error while building DOCX: {e}")
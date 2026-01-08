# handled_spectrum_app.py
# Streamlit app to generate “Handled Shipment Spectrum” section as a DOCX
# User only provides: Project Name + Client Name

import io
from dataclasses import dataclass
from typing import Dict, List

import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT  # just to avoid NameError if you later use it


# ---------- TEMPLATES ----------

@dataclass
class SorterTemplate:
    key: str
    label: str                  # human-readable name
    keywords: List[str]         # for fuzzy match on project name
    config_name: str            # used in intro paragraph
    item_singular: str          # "shipment" / "parcel"
    subheading_51: str          # text after "5.1 "
    spec_table: Dict[str, Dict[str, str]]  # spec -> {"unit": "...", "value": "..."}


TEMPLATES: List[SorterTemplate] = [
    # Generic Linear / Dual-belt sorter (based mainly on Example 8 – Boxes)
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
    # Loop CBS – standard shipments (based mainly on Example 6 – Auto Induct)
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
    # Heavy-duty / parcel sorter (based mainly on Example 7 & 9)
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


def choose_template(project_name: str) -> SorterTemplate:
    """Pick the closest template based on project name keywords."""
    text = (project_name or "").lower()

    # score by number of keyword hits
    best_tpl = TEMPLATES[0]
    best_score = -1
    for tpl in TEMPLATES:
        score = sum(1 for kw in tpl.keywords if kw in text)
        if score > best_score:
            best_score = score
            best_tpl = tpl

    return best_tpl


# ---------- DOCX BUILDERS ----------

def add_section_heading(doc: Document, section_number: str, title: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(f"{section_number}. {title}")
    run.bold = True
    run.underline = True
    run.font.size = Pt(12)
    run.font.name = "Calibri"


def add_paragraph(doc: Document, text: str, bold: bool = False, italic: bool = False) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(11)
    run.font.name = "Calibri"


def add_numbered_list(doc: Document, lines: List[str]) -> None:
    for line in lines:
        p = doc.add_paragraph(style="List Number")
        run = p.add_run(line)
        run.font.size = Pt(11)
        run.font.name = "Calibri"


def build_handled_spectrum_section(
    project_name: str,
    client_name: str,
    section_number: str = "5",
) -> Document:
    tpl = choose_template(project_name)

    doc = Document()

    item_singular = tpl.item_singular.lower()           # shipment / parcel
    item_cap = item_singular.capitalize()               # Shipment / Parcel
    item_plural = item_singular + "s"                   # shipments / parcels
    item_plural_cap = item_plural.capitalize()

    # 5. Handled Shipment Spectrum
    add_section_heading(doc, section_number, "Handled Shipment Spectrum")

    intro_1 = (
        "As per the {0} spectrum data provided in the RFP documents, Falcon has studied and "
        "analysed the {0} spectrum in detail.".format(item_singular)
    )
    intro_2 = (
        f"Falcon proposes to use its “{tpl.config_name}” to provide the maximum benefits to "
        f"{client_name} in terms of handling various sizes and weight."
    )
    add_paragraph(doc, intro_1)
    add_paragraph(doc, intro_2)

    # 5.1 ...
    add_paragraph(doc, f"{section_number}.1 {tpl.subheading_51}", bold=True)

    add_paragraph(
        doc,
        f"Falcon’s {tpl.config_name} has a capability to handle the below mentioned "
        f"{item_plural} sizes and weight.",
    )

    # Table (Specification / Unit / Value)
    table = doc.add_table(rows=1 + len(tpl.spec_table), cols=3)
    table.style = "Light Shading Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Specification"
    hdr_cells[1].text = "Unit"
    hdr_cells[2].text = "Value"
    for cell in hdr_cells:
        for r in cell.paragraphs[0].runs:
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(11)

    for spec, data in tpl.spec_table.items():
        row_cells = table.add_row().cells
        row_cells[0].text = spec
        row_cells[1].text = data["unit"]
        row_cells[2].text = data["value"]
        for idx in range(3):
            for r in row_cells[idx].paragraphs[0].runs:
                r.font.name = "Calibri"
                r.font.size = Pt(11)

    # 5.2 characteristics
    add_paragraph(
        doc,
        f"{section_number}.2 {item_plural_cap} to be loaded on Sorter shall have the following characteristics:",
        bold=True,
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
    add_numbered_list(doc, bullets_52)

    # 5.3 not loadable
    add_paragraph(
        doc,
        f"{section_number}.3 {item_plural_cap} not loadable on the sorter",
        bold=True,
    )

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
    add_numbered_list(doc, bullets_53)

    return doc


# ---------- STREAMLIT UI ----------

st.set_page_config(page_title="Handled Shipment Spectrum – DOCX Generator", layout="centered")

st.title("Handled Shipment Spectrum – DOCX Generator")

with st.form("handled_spectrum_form"):
    client_name = st.text_input("Client Name", value="")
    project_name = st.text_input("Project / Proposal Title", value="")
    section_number = st.text_input("Section Number (e.g., 5)", value="5")

    submitted = st.form_submit_button("Generate DOCX")

if submitted:
    if not client_name or not project_name:
        st.error("Please enter both Client Name and Project / Proposal Title.")
    else:
        tpl = choose_template(project_name)
        st.info(f"Detected sorter template: **{tpl.label}**")

        doc = build_handled_spectrum_section(
            project_name=project_name,
            client_name=client_name,
            section_number=section_number.strip() or "5",
        )

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_name = f"Handled_Shipment_Spectrum_{client_name.replace(' ', '_')}.docx"

        st.download_button(
            label="Download Handled Spectrum DOCX",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

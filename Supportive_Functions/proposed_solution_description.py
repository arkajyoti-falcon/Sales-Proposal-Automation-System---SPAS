# proposed_system_description_app.py
#
# Streamlit app to generate "Proposed System Description" section:
#  - Objective (fixed text, client name injected)
#  - Summary of the System (layout PNG from DXF or uploaded PNG)
#  - Process Flow of the System (via GROQ, based on DXF components)
#  - Main Benefits of the Proposed Solution (fixed bullets)
#
# Requirements:
#   pip install streamlit python-dotenv ezdxf convertapi python-docx requests pillow
#
# Env vars in .env:
#   GROQ_API_KEY=....
#   CONVERTAPI_SECRET=....

import os
import io
import json
import re
import tempfile
from collections import Counter, defaultdict
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv
import ezdxf
import convertapi
import requests
from docx import Document
from docx.shared import Pt, Inches
from PIL import Image

# --------------------------------------------------------------------
# Init / config
# --------------------------------------------------------------------
load_dotenv()
FEWSHOT_FILE = Path(__file__).with_name("process_flow.json")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
CONVERTAPI_SECRET = os.getenv("CONVERTAPI_SECRET")

st.set_page_config(page_title="Proposed System Description Generator", layout="wide")

st.title("Proposed System Description Generator")
st.subheader(
    "Generate Objective, Layout Summary, Process Flow, and Main Benefits from DXF + PNG"
)


# --------------------------------------------------------------------
# DXF → compact JSON summary for process-flow prompt
# --------------------------------------------------------------------

UNITS = {
    0: "Unitless",
    1: "inches",
    2: "feet",
    3: "miles",
    4: "millimeters",
    5: "centimeters",
    6: "meters",
    7: "kilometers",
}


def _is_noise_block(name: str) -> bool:
    """
    Filter out anonymous / noise blocks like *U69, *D123, etc.
    """
    n = name.strip()
    if re.match(r"^\*U\d+$", n, re.IGNORECASE):
        return True
    if re.match(r"^\*D\d+$", n, re.IGNORECASE):
        return True
    if n.startswith("*"):
        return True
    return False


def _normalize_group_name(name: str) -> str:
    """
    Normalize raw block name to a group name similar to your sample JSON:
      - lowercase
      - underscores / hyphens / multiple spaces → single spaces
      - strip trailing numbers / variants
    """
    n = name.strip()
    # remove xref prefix if any
    if "|" in n:
        n = n.split("|")[-1]
    # replace separators with space
    n = re.sub(r"[_\-]+", " ", n)
    # collapse spaces
    n = re.sub(r"\s+", " ", n).strip()
    # strip trailing numbers and brackets, e.g. "Leg Guard-01" → "Leg Guard"
    n = re.sub(r"\s*\(?\d+\)?$", "", n).strip()
    return n.lower()


def extract_dxf_components_for_process_flow(dxf_path: Path) -> dict:
    """
    Minimal DXF extraction: only component names + counts, shaped like your
    example JSON (groups + raw_block_counts). Filters out *U### noise blocks.
    """
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    hdr = doc.header
    units_code = hdr.get("$INSUNITS", None)
    try:
        units_code = int(units_code) if units_code is not None else None
    except Exception:
        units_code = None

    extmin = hdr.get("$EXTMIN", None)
    extmax = hdr.get("$EXTMAX", None)

    raw_counts: Counter[str] = Counter()

    for e in msp:
        try:
            if e.dxftype() == "INSERT":
                name = e.dxf.name or ""
                if not name:
                    continue
                if _is_noise_block(name):
                    continue
                raw_counts[name] += 1
        except Exception:
            continue

    # Grouping
    group_map: dict[str, dict] = defaultdict(
        lambda: {"total_count": 0, "examples": Counter()}
    )

    for raw_name, cnt in raw_counts.items():
        gname = _normalize_group_name(raw_name)
        if not gname:
            continue
        group_map[gname]["total_count"] += cnt
        group_map[gname]["examples"][raw_name] += cnt

    groups = []
    for gname, data in group_map.items():
        ex_list = [
            {"name": n, "count": c}
            for n, c in data["examples"].most_common(5)
        ]
        groups.append(
            {
                "group": gname,
                "total_count": int(data["total_count"]),
                "tokens": [],
                "examples": ex_list,
            }
        )

    # Sort groups by total_count desc
    groups.sort(key=lambda x: -x["total_count"])

    result = {
        "file": dxf_path.name,
        "units_code": units_code,
        "units_name": UNITS.get(units_code, "unknown") if units_code is not None else None,
        "extents": {
            "min": list(extmin) if extmin is not None else None,
            "max": list(extmax) if extmax is not None else None,
        },
        "groups": groups,
        "raw_block_counts": {k: int(v) for k, v in raw_counts.items()},
    }
    return result


# --------------------------------------------------------------------
# GROQ: Process Flow generation
# --------------------------------------------------------------------


def _summarise_components_for_prompt(dxf_json: dict) -> str:
    groups = dxf_json.get("groups", [])
    if not groups:
        return "No component groups detected."

    lines = []
    for g in groups[:40]:
        name = g.get("group", "")
        total = g.get("total_count", 0)
        ex = g.get("examples", [])
        top_example = ex[0]["name"] if ex else ""
        lines.append(f"- {name} (count: {total}, example: {top_example})")
    return "\n".join(lines)

def _load_fewshot_examples(max_examples: int = 6) -> list[dict]:
    """
    Load input/output few-shot examples from fewshot_examples.json.

    Structure of file:
    [
      {"input": "<DXF JSON string>", "output": "<process flow text>"},
      ...
    ]
    """
    if not FEWSHOT_FILE.exists():
        return []

    try:
        with FEWSHOT_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return []

    if not isinstance(data, list):
        return []

    # Limit to avoid blowing context window; tune as needed
    return data[:max_examples]

def _normalise_to_numbered_steps(raw_text: str) -> str:
    """
    Take raw GROQ output and force a clean 1..N numbered list.
    """
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

    if len(lines) == 1:
        # Try to split internal numbered segments in one long paragraph
        parts = re.split(r'(?:(?<=\.)\s+)(?=\d+\.)', lines[0])
        lines = [p.strip() for p in parts if p.strip()]

    steps = []
    for ln in lines:
        m = re.match(r"^(\d+)[\.\)\-]\s*(.*)$", ln)
        content = m.group(2).strip() if m else ln
        if content:
            steps.append(content)

    # De-duplicate near-identical lines
    dedup = []
    seen = set()
    for s in steps:
        key = re.sub(r"\s+", " ", s.lower())
        if key not in seen:
            seen.add(key)
            dedup.append(s)

    # limit 5–9 steps
    if len(dedup) < 5:
        max_steps = len(dedup)
    else:
        max_steps = min(len(dedup), 9)

    out_lines = []
    for i, content in enumerate(dedup[:max_steps], start=1):
        out_lines.append(f"{i}. {content}")
    return "\n".join(out_lines)


def call_groq_for_process_flow(client_name: str, project_name: str, dxf_json: dict):
    """
    Call GROQ and return:
      clean_steps, raw_text, groq_payload

    Now uses real input/output few-shot examples loaded from fewshot_examples.json.
    """
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY is not set in environment or .env")

    # Remove heavy part to keep prompt manageable
    safe_dxf_json = {k: v for k, v in dxf_json.items() if k != "raw_block_counts"}

    comp_summary = _summarise_components_for_prompt(safe_dxf_json)

    # Load real training examples
    fewshot_examples = _load_fewshot_examples(max_examples=6)

    fewshot_blocks: list[str] = []
    for idx, ex in enumerate(fewshot_examples, start=1):
        inp = (ex.get("input") or "").strip()
        out = (ex.get("output") or "").strip()
        if not inp or not out:
            continue

        block = (
            f"Example {idx} – INPUT (DXF JSON):\n"
            f"{inp}\n\n"
            f"Example {idx} – OUTPUT (Process Flow):\n"
            f"{out}"
        )
        fewshot_blocks.append(block)

    fewshot_text = "\n\n".join(fewshot_blocks) if fewshot_blocks else "None."

    system_prompt = """
You are a senior solution engineer at Falcon Autotech.
You write the section "Process Flow of the System" for Falcon's CBS-based proposals.

You MUST obey all of these rules:

1) OUTPUT FORMAT
- Output ONLY a numbered list of steps, one step per line.
- Each step MUST start with: "<number>. <Short Title>: <description>"
  Example: "1. Infeed System: Boxes are loaded on the infeed conveyors..."
- Aim for 5 to 9 main steps in total.
- You MAY include short sub-points within a step (e.g., a., b., c.) if needed for chute types,
  but keep them compact.
- Do NOT add headings, introductions, conclusions, or separate summary paragraphs.

2) GROUNDING IN DXF AND EXAMPLES
- Treat the DXF JSON (groups + counts) as the PRIMARY source of truth for the layout.
- Every step must be grounded in one or more of the component groups present in the DXF input,
  plus generic conveyor/sorter behaviour.
- You are also given multiple REAL examples where:
    - INPUT = DXF-derived JSON of groups.
    - OUTPUT = The manually written "Process Flow of the System".
  Use these examples to learn HOW component names map to stages in the flow and
  HOW details like chute types, counts, and VDS/PTL/Bagging modules are described.

3) WHAT YOU MAY OR MAY NOT INVENT
- If the DXF group names clearly indicate a module, you MAY describe it explicitly, e.g.:
    - Names containing "telescopico" → telescopic belt conveyors.
    - Names containing "infeed" or "feedline" → infeed/feed lines.
    - Names containing "auto induct", "autoinduct" → auto-induct lines.
    - Names containing "cbs", "cross belt" → sorter loop.
    - Names containing "gravity", "mini gravity", "bagging", "bulk", "reject",
      "irregular chute", "big parcel chute" → different chute / output types.
    - Names containing "ptl", "ptl racks" → PTL secondary sort.
    - Names containing "vds" or "VDS Arm" → Volume Distribution System.
- You may use generic terms like "infeed conveyors", "cross-belt sorter loop",
  "output chutes", and "operators" to connect the stages.
- You MUST NOT introduce modules that are neither:
    - visible in the DXF group names, nor
    - reasonably implied by the training examples for similar group patterns.
- Do NOT invent extra technologies, counts, or process stages that cannot be
  inferred from either the DXF or the patterns in the examples.

4) USE OF COUNTS AND VARIANTS
- When useful, use the 'total_count' to mention approximate numbers:
  e.g., "about 58 gravity chutes", "several telescopic conveyors", "multiple PTL racks", etc.
- When the DXF clearly shows different chute categories (e.g., gravity vs mini gravity vs bulk),
  describe them as sub-points in a single sorting/output step rather than flattening them.

5) FLOW LOGIC
- Follow a physically realistic material flow, consistent with the examples:
  loading / infeed → any distribution or buffering → induct lines →
  CBS loop → chutes / bagging / PTL / pallets / recirculation, etc. as indicated by group names.
- If only chutes + operators are visible (no telescopic, no auto-induct, no VDS),
  keep the flow simple: loading → sorter → chutes → operator collection.

6) STYLE
- Match the tone and level of detail of the training outputs:
  engineering proposal language, not marketing language.
- Focus only on the physical and logical movement of shipments/bags, not on
  scalability, redundancy, maintenance, SLAs, or electrical details.
- Avoid repeating the same idea across multiple steps.
- DO NOT ADD SCINTIFIC COMPONENTS NAMES LIKE "fs002v02", "fal s005v", ETC.
"""

    user_prompt = f"""
Client: {client_name}
Project: {project_name}

You are given several real training examples where each example has:
- INPUT: DXF-derived JSON (groups, counts, etc.)
- OUTPUT: The corresponding "Process Flow of the System" text that was written manually.

=== TRAINING EXAMPLES START ===
{fewshot_text}
=== TRAINING EXAMPLES END ===

Now you are given a NEW DXF input for a different project.

NEW INPUT – DXF JSON (sanitised):
{json.dumps(safe_dxf_json, indent=2)}

Readable component summary (from the same DXF):
{comp_summary}

Task:
Using ONLY this DXF input (plus your understanding from the training examples),
write the "Process Flow of the System" as a numbered list of 5–9 main steps.

Requirements:
- Base each step on one or more of the actual component groups from this NEW DXF.
- Use literal hints in the group names (telescopic, infeed, VDS, auto induct, CBS, gravity chute,
  mini gravity chute, bagging chute, big parcel chute, irregular chute, bulk, PTL, pallet, etc.)
  to infer which modules are present.
- Include counts where they add clarity (e.g., number of chutes, induct lines, telescopic conveyors).
- The output format must follow the rules in the system prompt and be similar in flavour and depth
  to the training outputs, but tailored to this NEW layout.

Now produce ONLY the final numbered list of steps.
"""

    groq_payload = {
        "model": "llama-3.3-70b-versatile",
        "temperature": 0.2,
        "max_tokens": 900,
        "messages": [
            {"role": "system", "content": system_prompt.strip()},
            {"role": "user", "content": user_prompt.strip()},
        ],
    }
    print("GROQ Payload:", groq_payload)
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json",
    }

    resp = requests.post(url, headers=headers, json=groq_payload, timeout=90)
    resp.raise_for_status()
    data = resp.json()
    raw_text = data["choices"][0]["message"]["content"].strip()

    clean_steps = _normalise_to_numbered_steps(raw_text)
    return clean_steps, raw_text, groq_payload




# --------------------------------------------------------------------
# DXF → PNG via ConvertAPI
# --------------------------------------------------------------------


def convert_dxf_to_png(dxf_path: Path) -> Path:
    if not CONVERTAPI_SECRET:
        raise RuntimeError(
            "CONVERTAPI_SECRET is not set (needed for DXF → PNG) and no PNG was uploaded."
        )
    convertapi.api_credentials = CONVERTAPI_SECRET
    result = convertapi.convert("png", {"File": str(dxf_path)}, from_format="dxf")
    out_files = result.save_files(str(dxf_path.parent))
    for f in out_files:
        if str(f).lower().endswith(".png"):
            return Path(f)
    return Path(out_files[0])


# --------------------------------------------------------------------
# DOCX building (SDI-style headings)
# --------------------------------------------------------------------


def _add_para(doc: Document, text: str, bold: bool = False, style: str | None = None):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    if bold:
        run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(11)


def build_proposed_system_doc(
    client_name: str,
    project_name: str,
    process_flow_text: str,
    png_path: Path | None,
) -> bytes:
    """
    Build SDI-style "Proposed System Description" page:
      5.0 Proposed System Description
      5.1 Objective
      5.2 Summary of the System (PNG)
      5.3 Process Flow of the System
      5.4 Main Benefits of the Proposed Solution
    """
    doc = Document()

    # Main heading
    h = doc.add_heading("5.0 Proposed System Description", level=1)
    for r in h.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(14)

    # 5.1 Objective
    h_obj = doc.add_heading("5.1 Objective", level=2)
    for r in h_obj.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(12)

    objective_text = (
        "The purpose of this proposal is to present the design, manufacturing, "
        "installation, commissioning, testing, and acceptance testing of the Cross Belt Sorter "
        f"system for sorting shipments, as per {client_name} requirements."
    )
    _add_para(doc, objective_text)

    doc.add_paragraph("")

    # 5.2 Summary of the System (layout PNG)
    h_sum = doc.add_heading("5.2 Summary of the System", level=2)
    for r in h_sum.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(12)

    _add_para(
        doc,
        "The following layout view illustrates the overall arrangement of infeed conveyors, sorter loop, "
        "and output chutes for the proposed system.",
    )

    if png_path and png_path.exists():
        doc.add_paragraph("")
        # Insert PNG (scale to width ~6.5 inch)
        doc.add_picture(str(png_path), width=Inches(6.5))
        doc.add_paragraph("")  # spacer
    else:
        _add_para(
            doc,
            "The detailed layout is provided separately in the attached drawing.",
        )

    # 5.3 Process Flow of the System
    h_pf = doc.add_heading("5.3 Process Flow of the System", level=2)
    for r in h_pf.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(12)

    for line in process_flow_text.splitlines():
        line = line.strip()
        if not line:
            continue
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.name = "Calibri"
        run.font.size = Pt(11)

    doc.add_paragraph("")

    # 5.4 Main Benefits
    h_ben = doc.add_heading("5.4 Main Benefits of the Proposed Solution", level=2)
    for r in h_ben.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(12)

    benefits = [
        "High operational throughput.",
        "Low occupancy of floor space in the building.",
        "Narrow discharge centers for the increased number of splits in limited space.",
        (
            "FALCON’s CBS can adapt to changing business requirements by adjusting its speed "
            "to match the operational throughput requirement, thereby leading to power savings "
            "and reduced system wear & tear."
        ),
    ]
    for b in benefits:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(b)
        run.font.name = "Calibri"
        run.font.size = Pt(11)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------
# Streamlit UI – Inputs
# --------------------------------------------------------------------
with st.form("psd_form"):
    client_name = st.text_input("Client Name", placeholder="e.g. Landmark Group")
    project_name = st.text_input(
        "Project / System Name", placeholder="e.g. Loop CBS 12K Sortation System"
    )

    dxf_file = st.file_uploader("Upload DXF Layout", type=["dxf"])
    png_file = st.file_uploader(
        "Optional: Upload Layout PNG (skips DXF→PNG conversion)", type=["png", "jpg", "jpeg"]
    )

    submitted = st.form_submit_button("Generate Proposed System Description DOCX")

# --------------------------------------------------------------------
# Main processing
# --------------------------------------------------------------------
if submitted:
    if not client_name:
        st.error("Client Name is required.")
    elif not project_name:
        st.error("Project / System Name is required.")
    elif not dxf_file:
        st.error("DXF layout file is required.")
    elif not GROQ_API_KEY:
        st.error("GROQ_API_KEY is not set in environment / .env.")
    else:
        tmp_dir = Path(tempfile.mkdtemp(prefix="psd_"))
        dxf_path = tmp_dir / dxf_file.name
        dxf_path.write_bytes(dxf_file.getvalue())

        if png_file is not None:
            png_path = tmp_dir / png_file.name
            png_path.write_bytes(png_file.getvalue())
            st.info("PNG uploaded → DXF→PNG conversion will be skipped.")
        else:
            png_path = None

        # 1) DXF extraction
        st.markdown("### Step 1 – Extracting Components from DXF")
        try:
            dxf_json = extract_dxf_components_for_process_flow(dxf_path)
            st.caption("DXF component summary (input to GROQ):")
            st.json(dxf_json)
        except Exception as e:
            st.error(f"DXF parsing failed: {e}")
            st.stop()

        # 2) GROQ Process Flow
        st.markdown("### Step 2 – Calling GROQ for Process Flow")
        try:
            process_flow_steps, groq_raw, groq_payload = call_groq_for_process_flow(
                client_name=client_name,
                project_name=project_name,
                dxf_json=dxf_json,
            )

            # Debug: payload without secrets
            st.caption("GROQ payload (sanitised, without API key):")
            safe_payload = dict(groq_payload)
            st.code(json.dumps(safe_payload, indent=2), language="json")

            st.caption("RAW GROQ output:")
            st.text_area("Raw GROQ Response", groq_raw, height=200)

            st.caption("Normalised Process Flow (used in DOCX):")
            st.text_area("Process Flow Steps", process_flow_steps, height=220)

        except Exception as e:
            st.error(f"GROQ call failed: {e}")
            st.stop()

        # 3) DXF → PNG if needed
        if png_path is None:
            st.markdown("### Step 3 – Converting DXF to PNG (ConvertAPI)")
            try:
                png_path = convert_dxf_to_png(dxf_path)
                st.success(f"DXF converted to PNG: {png_path.name}")
                try:
                    img = Image.open(png_path)
                    st.image(img, caption="Layout PNG (preview)", use_column_width=True)
                except Exception:
                    st.caption("PNG generated (preview failed, but file is available).")
            except Exception as e:
                st.error(f"DXF → PNG conversion failed: {e}")
                st.stop()
        else:
            st.markdown("### Step 3 – Skipping DXF→PNG (using uploaded PNG)")

        # 4) Build DOCX
        st.markdown("### Step 4 – Building Proposed System Description DOCX")
        try:
            doc_bytes = build_proposed_system_doc(
                client_name=client_name,
                project_name=project_name,
                process_flow_text=process_flow_steps,
                png_path=png_path,
            )
            st.success("DOCX generated successfully.")

            out_name = f"Proposed_System_Description_{project_name.replace(' ', '_')}.docx"
            st.download_button(
                "Download Proposed System Description DOCX",
                data=doc_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"Error while building DOCX: {e}")

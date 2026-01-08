#!/usr/bin/env python3
# concept_description_app.py
#
# Streamlit app:
#  - Input: Process Flow (text)
#  - Call GROQ to generate Mermaid flowchart code
#  - Render Mermaid → PNG via Kroki / mermaid.ink
#  - Build DOCX with a "Concept Description" section
#  - Flowchart image is hyperlinked to draw.io (or any URL you set)
#
# Requirements:
#   pip install streamlit python-dotenv python-docx requests groq
#
# Env vars in .env:
#   GROQ_API_KEY=....

import os
import io
import base64
import json
import re
from typing import Tuple, Optional
import copy

import requests
import streamlit as st
from dotenv import load_dotenv
from groq import Groq

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ---------------------------------------------------------------------
# Init
# ---------------------------------------------------------------------
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

st.set_page_config(page_title="Concept Description Flowchart Generator", layout="wide")
st.title("Concept Description – Flowchart Generator")
st.markdown(
    """
This tool takes a **Process Flow** (text), generates a **Mermaid flowchart** using GROQ,
renders it to **PNG**, and embeds it into a **DOCX** as a clickable image (linked to draw.io).
"""
)


# ---------------------------------------------------------------------
# Helpers: Load GROQ Client
# ---------------------------------------------------------------------
def load_groq() -> Optional[Groq]:
    """Load Groq client from environment variable."""
    key = (os.getenv("GROQ_API_KEY") or "").strip()
    if not key:
        return None
    try:
        return Groq(api_key=key)
    except Exception:
        return None


# ---------------------------------------------------------------------
# Helpers: Mermaid Sanitization
# ---------------------------------------------------------------------
def sanitize_mermaid_for_render(code: str) -> str:
    """
    Minimal sanitization - only fix obvious issues.
    Don't over-process as it can break valid Mermaid syntax.
    """
    if not code:
        return code
    
    # Remove any markdown code fences
    code = code.strip()
    if code.startswith("```"):
        lines = code.split("\n")
        if lines[0].strip().startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        code = "\n".join(lines).strip()
    
    # Remove leading "mermaid" keyword if present
    if code.lower().startswith("mermaid"):
        code = code[7:].strip()
    
    # Ensure it starts with flowchart
    if not code.lower().startswith("flowchart"):
        return code
    
    # Normalize line endings
    code = code.replace("\r\n", "\n").replace("\r", "\n")
    
    return code


# ---------------------------------------------------------------------
# Helpers: Mermaid → PNG Rendering
# ---------------------------------------------------------------------
def generate_mermaid_flowchart_png(mermaid_code: str) -> Tuple[bytes, str]:
    """
    Render Mermaid diagram to PNG with improved error handling.
    Tries multiple strategies:
    1. Kroki PNG with JSON payload
    2. Kroki PNG with plain text
    3. Mermaid.ink with base64 encoding
    """
    logs = []
    
    # Clean the mermaid code
    mermaid_str = sanitize_mermaid_for_render(mermaid_code)
    logs.append(f"Cleaned Mermaid code:\n{mermaid_str}\n")
    
    # Strategy 1: Kroki PNG with proper JSON payload
    try:
        kroki_url = "https://kroki.io/mermaid/png"
        logs.append(f"[1/3] Trying Kroki PNG endpoint: {kroki_url}")
        
        # Try with JSON payload (Kroki's preferred method)
        payload = {
            "diagram_source": mermaid_str,
            "diagram_type": "mermaid",
            "output_format": "png"
        }
        
        resp = requests.post(
            kroki_url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30,
        )
        
        logs.append(f"Kroki PNG HTTP status: {resp.status_code}")
        
        if resp.ok and resp.content and len(resp.content) > 100:
            logs.append(f"✓ Kroki PNG (JSON) returned {len(resp.content)} bytes.")
            return resp.content, "\n".join(logs)
        else:
            snippet = resp.text[:500]
            logs.append(f"Kroki PNG (JSON) error: {snippet}")
    except Exception as e:
        logs.append(f"Kroki PNG (JSON) exception: {repr(e)}")
    
    # Strategy 2: Kroki PNG with plain text
    try:
        kroki_url = "https://kroki.io/mermaid/png"
        logs.append(f"[2/3] Trying Kroki PNG with plain text")
        
        resp = requests.post(
            kroki_url,
            data=mermaid_str.encode("utf-8"),
            headers={"Content-Type": "text/plain"},
            timeout=30,
        )
        
        logs.append(f"Kroki PNG (text) HTTP status: {resp.status_code}")
        
        if resp.ok and resp.content and len(resp.content) > 100:
            logs.append(f"✓ Kroki PNG (text) returned {len(resp.content)} bytes.")
            return resp.content, "\n".join(logs)
        else:
            snippet = resp.text[:500]
            logs.append(f"Kroki PNG (text) error: {snippet}")
    except Exception as e:
        logs.append(f"Kroki PNG (text) exception: {repr(e)}")
    
    # Strategy 3: Mermaid.ink with corrected encoding
    try:
        logs.append("[3/3] Trying mermaid.ink")
        
        # Properly encode for mermaid.ink
        json_payload = {
            "code": mermaid_str,
            "mermaid": {"theme": "default"}
        }
        json_str = json.dumps(json_payload)
        
        # Encode: UTF-8 -> bytes -> base64 -> URL-safe
        utf8_bytes = json_str.encode('utf-8')
        b64_bytes = base64.b64encode(utf8_bytes)
        b64_str = b64_bytes.decode('ascii')
        
        img_url = f"https://mermaid.ink/img/{b64_str}"
        logs.append(f"mermaid.ink URL (first 200 chars): {img_url[:200]}")
        
        resp = requests.get(img_url, timeout=30)
        logs.append(f"mermaid.ink HTTP status: {resp.status_code}")
        
        if resp.ok and resp.content and len(resp.content) > 100:
            logs.append(f"✓ mermaid.ink returned {len(resp.content)} bytes.")
            return resp.content, "\n".join(logs)
        else:
            snippet = resp.text[:500]
            logs.append(f"mermaid.ink error: {snippet}")
    except Exception as e:
        logs.append(f"mermaid.ink exception: {repr(e)}")
    
    # All strategies failed
    full_log = "\n".join(logs)
    raise RuntimeError(
        f"Failed to render Mermaid diagram via all strategies.\n\n"
        f"Debug log:\n{full_log}\n\n"
        f"Original Mermaid code:\n{mermaid_code}"
    )


# ---------------------------------------------------------------------
# Helpers: GROQ – process flow → Mermaid code
# ---------------------------------------------------------------------
def call_groq_for_mermaid(process_flow_text: str, groq_client: Groq) -> Tuple[str, dict]:
    """
    Call GROQ to convert a textual Process Flow into Mermaid flowchart code.
    Enhanced prompt for better Mermaid v11 compatibility with colors and vertical layout.

    Returns:
      mermaid_code (str), raw_response (dict)
    """
    system_prompt = """You are a diagram expert specializing in Mermaid v11 flowcharts.

CRITICAL MERMAID v11 RULES:
1. Start with exactly: flowchart TD (Top Down - vertical layout)
2. Use ONLY simple alphanumeric IDs: A, B, C, D1, E2 (no special chars)
3. Use square brackets for labels: A[Start Process]
4. Keep labels SHORT (2-5 words max), Keep the flowchart as much detailed as possible.
5. NO quotes, NO colons, NO semicolons in labels
6. Use --> for simple arrows
7. For decisions, use diamond syntax: D{Decision?}
8. Apply semantic colors using style/class definitions:
   - Green (#90EE90) for start/input nodes
   - Blue (#87CEEB) for processing nodes
   - Yellow (#FFE97F) for sorting/decision nodes
   - Orange (#FFB366) for collection/storage nodes
   - Red (#FFB3B3) for rejection/error nodes
   - Purple (#DDA0DD) for recirculation/loop nodes
9. Use classDef to define color styles
10. Output ONLY the mermaid code, no explanation, no backticks

COLOR CODING EXAMPLE:
flowchart TD
    classDef inputStyle fill:#90EE90,stroke:#333,stroke-width:2px
    classDef processStyle fill:#87CEEB,stroke:#333,stroke-width:2px
    classDef sortStyle fill:#FFE97F,stroke:#333,stroke-width:2px
    classDef collectStyle fill:#FFB366,stroke:#333,stroke-width:2px
    classDef rejectStyle fill:#FFB3B3,stroke:#333,stroke-width:2px
    classDef recircStyle fill:#DDA0DD,stroke:#333,stroke-width:2px
    
    A[Infeed System]:::inputStyle
    B[Processing]:::processStyle
    C{Sort Decision}:::sortStyle
    D[Collection]:::collectStyle
    E[Rejection]:::rejectStyle
    F[Recirculation]:::recircStyle
    
    A --> B --> C
    C -->|Valid| D
    C -->|Invalid| E
    E --> F --> B
"""

    user_prompt = f"""Convert this process flow to a vertical Mermaid flowchart with appropriate colors:

{process_flow_text}

Requirements:
- Use flowchart TD for TOP-DOWN vertical layout
- Simple node IDs (A, B, C, D, E, etc.)
- Short labels in brackets (2-5 words)
- Show main flow vertically downward
- Apply color coding:
  * Green for Infeed/Input systems
  * Blue for Inducts/Processing
  * Yellow for CBS/Sorting operations
  * Orange for Live Chutes/Collection
  * Red for Rejection Chute
  * Purple for Recirculation Line
- Use classDef at the top to define styles
- Apply styles using :::className syntax
- Connect nodes with simple arrows
- NO markdown backticks in output
- Keep the diagram compact to fit on one page"""

    try:
        resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt.strip()},
                {"role": "user", "content": user_prompt.strip()},
            ],
            temperature=0.0,  # More deterministic
            max_tokens=2000,
        )
        
        mermaid_raw = (resp.choices[0].message.content or "").strip()
        
        # Aggressive cleanup
        mermaid_raw = re.sub(r"^```(?:mermaid)?\s*", "", mermaid_raw, flags=re.MULTILINE)
        mermaid_raw = re.sub(r"\s*```$", "", mermaid_raw, flags=re.MULTILINE)
        mermaid_raw = mermaid_raw.strip()
        
        # Ensure starts with flowchart
        if not mermaid_raw.lower().startswith("flowchart"):
            lines = mermaid_raw.split("\n")
            for i, line in enumerate(lines):
                if line.strip().lower().startswith("flowchart"):
                    mermaid_raw = "\n".join(lines[i:])
                    break
        
        data = {
            "choices": [{"message": {"content": mermaid_raw}}],
            "model": "llama-3.3-70b-versatile"
        }
        
        return mermaid_raw, data
        
    except Exception as e:
        raise RuntimeError(f"GROQ API call failed: {str(e)}")


# ---------------------------------------------------------------------
# Helpers: DOCX building with clickable image (single page)
# ---------------------------------------------------------------------
def build_concept_description_doc(png_bytes: bytes, drawio_url: str) -> bytes:
    """
    Build a DOCX with ONLY the "Concept Description" heading and flowchart.
    No process flow text - keep it simple and single page.
    Flowchart is centered and hyperlinked to draw.io.
    """
    doc = Document()
    
    # Set narrow margins for more space
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # Main heading
    h = doc.add_heading("Concept Description", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in h.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue

    # Add small spacing
    doc.add_paragraph()

    # Insert hyperlinked flowchart (centered)
    try:
        image_stream = io.BytesIO(png_bytes)
        
        # Add picture to paragraph
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        
        # Calculate width to fit on one page (leave some margin)
        # Portrait: ~6.5 inches width is safe for letter size with margins
        inline_shape = run.add_picture(image_stream, width=Inches(6.5))
        
        # Create external hyperlink relationship
        rel_id = doc.part.relate_to(drawio_url, RT.HYPERLINK, is_external=True)
        
        # Get the drawing element (inline)
        inline = inline_shape._inline
        
        # Create a new hyperlink element
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), rel_id)
        
        # Clone the inline element to avoid parent issues
        inline_copy = copy.deepcopy(inline)
        
        # Add the cloned inline to hyperlink
        hyperlink.append(inline_copy)
        
        # Replace the original inline with the hyperlink in the run's XML
        run._r.replace(inline, hyperlink)
        
    except Exception as e:
        # Fallback: just insert the image without hyperlink
        st.warning(f"Hyperlink creation failed: {str(e)}. Image added without link.")
        try:
            image_stream = io.BytesIO(png_bytes)
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run()
            run.add_picture(image_stream, width=Inches(6.5))
        except Exception as e2:
            st.error(f"Image insertion also failed: {str(e2)}")

    # Add footer note
    doc.add_paragraph()
    footer_para = doc.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_para.add_run("Click the diagram to open in draw.io for editing")
    footer_run.font.size = Pt(9)
    footer_run.font.color.rgb = RGBColor(128, 128, 128)  # Gray
    footer_run.font.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------
with st.form("concept_form"):
    process_flow_input = st.text_area(
        "Paste the Process Flow text here",
        height=300,
        placeholder="1. Infeed System: ...\n2. Inducts: ...\n3. Loop CBS: ...\n...",
    )

    drawio_url = st.text_input(
        "Draw.io URL to open on image click",
        value="https://app.diagrams.net/",
        help="When the flowchart image in DOCX is clicked, Word will open this URL."
    )

    docx_filename = st.text_input(
        "Output DOCX file name",
        value="Concept_Description_Flowchart.docx",
    )

    submitted = st.form_submit_button("Generate Concept Description DOCX")

# ---------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------
if submitted:
    if not process_flow_input.strip():
        st.error("Please paste the Process Flow text.")
    elif not GROQ_API_KEY:
        st.error("GROQ_API_KEY is not set in environment / .env.")
    else:
        # Load GROQ client
        groq_client = load_groq()
        if not groq_client:
            st.error("Failed to initialize GROQ client.")
            st.stop()
        
        # 1) GROQ → Mermaid
        try:
            with st.spinner("Calling GROQ to generate Mermaid flowchart code..."):
                mermaid_code, groq_raw = call_groq_for_mermaid(process_flow_input, groq_client)
        except Exception as e:
            st.error(f"GROQ call failed: {e}")
            st.stop()

        st.subheader("Mermaid Flowchart Code (from GROQ)")
        st.code(mermaid_code, language="mermaid")
        
        # Add manual verification link
        st.info("💡 Tip: You can test this code at [mermaid.live](https://mermaid.live) to verify it renders correctly")

        # 2) Mermaid → PNG
        try:
            with st.spinner("Rendering Mermaid diagram to PNG..."):
                png_bytes, render_debug = generate_mermaid_flowchart_png(mermaid_code)
        except Exception as e:
            # Show full debug
            err_msg = str(e)
            st.error("❌ Flowchart render failed.")
            with st.expander("📋 Render Debug Log"):
                st.text(err_msg)
            st.warning("Try copying the Mermaid code above and testing it at mermaid.live to see if it's valid.")
            st.stop()

        # Show renderer debug log even if successful (in expander)
        st.success("✅ Flowchart rendered successfully.")
        with st.expander("📋 Render Debug Log"):
            st.text(render_debug)

        # Preview PNG inside Streamlit
        st.subheader("Flowchart Preview")
        st.image(png_bytes, caption="Generated Flowchart (Mermaid → PNG)", use_container_width=True)

        # 3) Build DOCX (simplified - no process flow text)
        try:
            with st.spinner("Building Concept Description DOCX..."):
                doc_bytes = build_concept_description_doc(
                    png_bytes=png_bytes,
                    drawio_url=drawio_url,
                )
        except Exception as e:
            st.error(f"Error while building DOCX: {e}")
            st.stop()

        st.success("✅ DOCX generated successfully - Single page with flowchart only!")

        st.download_button(
            "📥 Download Concept Description DOCX",
            data=doc_bytes,
            file_name=docx_filename or "Concept_Description_Flowchart.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # Optional: show RAW GROQ JSON in an expander for debugging
        with st.expander("🔍 Show RAW GROQ JSON (debug)"):
            st.code(json.dumps(groq_raw, indent=2), language="json")
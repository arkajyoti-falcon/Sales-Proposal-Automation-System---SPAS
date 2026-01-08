# dxf_groq_detector_app.py
"""
Basic DXF -> Full Extract (nested + layouts + layers) -> Groq detect components
Model: openai/gpt-oss-120b (hardcoded)

Run:
  pip install streamlit ezdxf requests
  (PowerShell)  $env:GROQ_API_KEY="YOUR_KEY"
  streamlit run dxf_groq_detector_app.py
"""

import os
import re
import json
import time
import tempfile
from pathlib import Path
from collections import Counter, defaultdict
from typing import Dict, Optional, Set, Tuple, Any

import requests
import ezdxf
import streamlit as st


# =========================
# CONFIG
# =========================

GROQ_MODEL = "openai/gpt-oss-120b"
GROQ_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions"

UNITS = {
    0: "Unitless", 1: "inches", 2: "feet", 3: "miles",
    4: "millimeters", 5: "centimeters", 6: "meters", 7: "kilometers"
}

# remove anonymous U-block noise: *U119 / U119
U_NOISE_RE = re.compile(r"^\*?U\d+$", re.IGNORECASE)

def is_u_noise(name: str) -> bool:
    return bool(U_NOISE_RE.match((name or "").strip()))


# =========================
# FULL DXF EXTRACTION
# =========================

def get_units(doc) -> Tuple[Optional[int], Optional[str]]:
    units_code = doc.header.get("$INSUNITS", None)
    try:
        units_code = int(units_code) if units_code is not None else None
    except Exception:
        units_code = None
    units_name = UNITS.get(units_code, "unknown") if units_code is not None else None
    return units_code, units_name


def safe_attrib_dump(insert_entity) -> Dict[str, str]:
    out = {}
    try:
        for a in getattr(insert_entity, "attribs", []):
            try:
                out[str(a.dxf.tag)] = str(a.dxf.text)
            except Exception:
                continue
    except Exception:
        pass
    return out


def scan_space(space,
               entity_counts_by_type: Counter,
               entity_counts_by_layer: Counter,
               entity_counts_by_layer_type: Dict[str, Counter],
               insert_counts: Counter,
               insert_attr_samples: Dict[str, Dict[str, Any]],
               max_attr_samples_per_block: int = 3) -> None:
    for e in space:
        try:
            t = e.dxftype()
            entity_counts_by_type[t] += 1

            layer = ""
            try:
                layer = str(e.dxf.layer) if hasattr(e, "dxf") and hasattr(e.dxf, "layer") else ""
            except Exception:
                layer = ""
            entity_counts_by_layer[layer] += 1
            entity_counts_by_layer_type[layer][t] += 1

            if t == "INSERT":
                bname = str(e.dxf.name)

                # ✅ drop U-noise everywhere
                if is_u_noise(bname):
                    continue

                insert_counts[bname] += 1

                # sample attributes (useful for station/scanner hints in some DXFs)
                if bname not in insert_attr_samples:
                    insert_attr_samples[bname] = {"samples": []}
                if len(insert_attr_samples[bname]["samples"]) < max_attr_samples_per_block:
                    attrs = safe_attrib_dump(e)
                    if attrs:
                        insert_attr_samples[bname]["samples"].append(attrs)
        except Exception:
            continue


def expand_nested_inserts(doc, top_level_inserts: Counter, max_depth: int = 25) -> Counter:
    """
    Returns a Counter of ALL nested INSERT block names (no U-noise), with multiplied counts.
    This is the key to avoiding missing components hidden inside assemblies.
    """
    all_recursive = Counter()

    def walk(block_name: str, multiplier: int, depth: int, stack: Set[str]):
        if depth > max_depth:
            return
        if block_name in stack:
            return
        stack.add(block_name)

        try:
            blk = doc.blocks.get(block_name)
        except Exception:
            stack.remove(block_name)
            return

        local = Counter()
        for e in blk:
            try:
                if e.dxftype() == "INSERT":
                    child = str(e.dxf.name)
                    if is_u_noise(child):
                        continue
                    local[child] += 1
            except Exception:
                continue

        for child, cnt in local.items():
            total = multiplier * cnt
            all_recursive[child] += total
            walk(child, total, depth + 1, stack)

        stack.remove(block_name)

    for bname, cnt in top_level_inserts.items():
        if is_u_noise(bname):
            continue
        all_recursive[bname] += cnt
        walk(bname, cnt, 0, set())

    return all_recursive


def extract_full_dxf_info(dxf_path: Path,
                          include_layouts: bool = True,
                          nested_depth: int = 25) -> dict:
    doc = ezdxf.readfile(str(dxf_path))
    units_code, units_name = get_units(doc)

    # modelspace
    msp = doc.modelspace()
    m_type = Counter()
    m_layer = Counter()
    m_layer_type = defaultdict(Counter)
    m_inserts = Counter()
    insert_attr_samples = {}

    scan_space(msp, m_type, m_layer, m_layer_type, m_inserts, insert_attr_samples)

    # layouts
    layout_summaries = {}
    layout_inserts_total = Counter()

    if include_layouts:
        for layout in doc.layouts:
            if layout.name.lower() == "model":
                continue
            ec_type = Counter()
            ec_layer = Counter()
            ec_layer_type = defaultdict(Counter)
            ec_inserts = Counter()
            _tmp_attrs = {}

            scan_space(layout, ec_type, ec_layer, ec_layer_type, ec_inserts, _tmp_attrs)

            layout_summaries[layout.name] = {
                "entity_counts_by_type": dict(ec_type),
                "entity_counts_by_layer": dict(ec_layer),
                "entity_counts_by_layer_type": {k: dict(v) for k, v in ec_layer_type.items()},
                "insert_counts": dict(ec_inserts),
            }
            layout_inserts_total.update(ec_inserts)

    top_level_inserts_all = Counter(m_inserts)
    top_level_inserts_all.update(layout_inserts_total)

    recursive_inserts = expand_nested_inserts(doc, top_level_inserts_all, max_depth=nested_depth)

    # top layers summary (helps detect geometry-only components)
    layer_stats = defaultdict(lambda: {"total_entities": 0, "by_type": Counter()})
    for layer, type_counts in m_layer_type.items():
        layer_stats[layer]["total_entities"] += sum(type_counts.values())
        layer_stats[layer]["by_type"].update(type_counts)

    top_layers = sorted(layer_stats.items(), key=lambda x: -x[1]["total_entities"])[:60]
    top_layers_summary = [
        {"layer": layer, "total_entities": int(info["total_entities"]), "by_type": dict(info["by_type"])}
        for layer, info in top_layers
    ]

    return {
        "file": dxf_path.name,
        "units_code": units_code,
        "units_name": units_name,

        "extraction_scope": {
            "scanned_modelspace": True,
            "scanned_layouts": bool(include_layouts),
            "layout_names": list(layout_summaries.keys()),
            "nested_depth": int(nested_depth),
            "u_noise_removed": True,
        },

        "top_level_inserts": {
            "modelspace": dict(m_inserts),
            "layouts_total": dict(layout_inserts_total),
            "all_top_level": dict(top_level_inserts_all),
        },

        "nested_inserts_all_recursive": dict(recursive_inserts),

        "modelspace_entities": {
            "counts_by_type": dict(m_type),
            "counts_by_layer": dict(m_layer),
            "counts_by_layer_type": {k: dict(v) for k, v in m_layer_type.items()},
            "top_layers_summary": top_layers_summary,
        },

        "layouts": layout_summaries,
        "insert_attribute_samples": insert_attr_samples,
    }


# =========================
# LLM EVIDENCE PACK
# =========================

def build_llm_evidence_pack(full: dict) -> dict:
    """
    Evidence-rich but bounded-size payload to reduce misses.
    We send:
      - all recursive inserts (possibly capped)
      - top-level inserts (possibly capped)
      - top layers summary
      - entity counts
      - attribute samples
    """
    rec = full.get("nested_inserts_all_recursive", {}) or {}
    top = full.get("top_level_inserts", {}).get("all_top_level", {}) or {}
    layers = full.get("modelspace_entities", {}).get("top_layers_summary", []) or []

    # cap to keep payload stable
    def cap_counter(d: dict, n: int) -> dict:
        items = sorted(d.items(), key=lambda x: -x[1])
        return dict(items[:n])

    rec_capped = cap_counter(rec, 3000)  # raise if you want
    top_capped = cap_counter(top, 1200)

    return {
        "file": full.get("file"),
        "units_name": full.get("units_name"),
        "extraction_scope": full.get("extraction_scope", {}),
        "top_level_inserts_top": top_capped,
        "nested_inserts_all_recursive_top": rec_capped,
        "modelspace_entities_counts_by_type": full.get("modelspace_entities", {}).get("counts_by_type", {}),
        "modelspace_top_layers_summary": layers,
        "insert_attribute_samples": full.get("insert_attribute_samples", {}),
        "notes": {
            "recursive_inserts_total_unique": len(rec),
            "recursive_inserts_sent_unique": len(rec_capped),
            "top_level_inserts_total_unique": len(top),
            "top_level_inserts_sent_unique": len(top_capped),
        }
    }


# =========================
# ENHANCED PROMPT (fixed flow + coverage rules)
# =========================

def build_groq_prompt(evidence_pack: dict) -> str:
    return f"""
You are a CAD-to-system “System Description” detector for warehouse automation drawings.

You will receive a DXF extraction JSON (evidence pack) produced by a parser.
YOU MUST NOT use any external assumptions.
You must ONLY mark something as present if you can cite evidence from the JSON fields.

Your output MUST follow this fixed flow/order (always same order):
1. Infeed System
    1.1. Telescopic Belt Conveyor (ONLY if present)
    1.2. Infeed Conveyor 
        1.2.1. Infeed Conveyor subcomponent 1 
        1.2.2. Infeed Conveyor subcomponent 2
    1.3. VDS Loop (ONLY if present)
    
2. Induction to Sorter
   2.1 Feedlines
        2.1.1. Feedline subcomponent 1
        2.1.2. Feedline subcomponent 2
3. Manual Induction Station (ONLY if present)
3. Main Loop (Linear CBS / Loop CBS)
4. Barcode Scanning System (ONLY if present)
5. Output Chutes
6. Bag Takeaway Conveyor (ONLY if present)
6. Exception Handling Area (ONLY if present)
7. Recirculation & Manual Feedline (ONLY if present)

CANONICAL SUBCOMPONENTS TO DETECT

1) Infeed Conveyor subcomponents:
- Straight Conveyor
- Inclined Conveyor
- Straight & Inclined Conveyor (use if not separable)
- Plastic Modular Conveyor
- Curve Conveyor
- Buffer Conveyor
- Angle Merge Conveyor
- Aligning Conveyor
- Infeed Conveyor (generic, use only if infeed exists but type unclear)

2) Feedlines (under 2.1 Feedlines):
- Spacing Conveyor
- Loading Conveyor / Orientation conveyor
- Buffer Conveyor
- Weighing Conveyor
- Receiving Conveyor
- Intelligent Merge Conveyor
- Rejection Chutes

3) Output Chutes subcomponents:
- Gravity Chutes
- Mini-Gravity Chutes
- Bulk Chutes
- Sliding Chutes
- Collection Chutes
- Live Chutes
- Manual Chutes
- Technical Chutes
- Direct Bagging Chutes
- Non-sort Collection Chutes
- Dispression / Dispersion Chutes
- Bagging type PTL
- Rejection chutes (also valid here if clearly chute-related)

OPTIONAL SECTIONS LOGIC
- Barcode Scanning System: present if any evidence contains scan/barcode/dws/dimension/reader OR weigh/weight/scale (weighing is treated as scanning/DWS signal).
- Exception Handling Area: present if any evidence contains reject/rejection/exception/sortfail/overweight/technical chute.
- Recirculation & Manual Feedline: present if any evidence contains recirculation/return/refeed OR manual/operator/induct station.

SPECIAL DXF-SPECIFIC MAPPINGS (MANDATORY):
1) If any block name contains "FAL_FS002V02" OR contains both "FS002" and "WITHOUT WEIGHING" (case-insensitive),
   then you MUST classify it as:
   - Feedlines → "Buffer Conveyor" (present=true)
   Evidence must cite the exact matched block name and count.

2) If any block name or layer name contains "VDS" OR contains "VDS LOOP" OR contains both "VDS" and "LOOP" (case-insensitive),
   then you MUST mark:
   - Infeed System → "Buffer Conveyor" (present=true) AND/OR set a signal in unmapped_signals if you cannot decide placement
   Also set: "VDS Loop" = present as a detected feature under unmapped_signals.blocks (why_relevant="VDS loop detected").
   Evidence must cite the exact matched text and count (block) or layer name (layer).


EVIDENCE SOURCES YOU MAY USE (ONLY these fields):
- nested_inserts_all_recursive_top (block names + counts)
- top_level_inserts_top (block names + counts)
- modelspace_top_layers_summary (layer names + entity types)
- insert_attribute_samples
- modelspace_entities_counts_by_type

MAPPING RULES (STRICT, TEXT-CLUE BASED)
You may map by literal text clues only:
- telescopic OR telescopico -> Telescopic Belt Conveyor
- "straight", "line", "inline" -> Straight Conveyor
- "incline", "inclined", "elevation", "ramp" -> Inclined Conveyor
- If infeed exists but straight vs incline not separable -> Straight & Inclined Conveyor
- "pmc" OR "modular" OR "plastic modular" -> Plastic Modular Conveyor
- "turn" OR "curve" OR "90" OR "45" OR "deg" -> Curve Conveyor (if it looks like a conveyor/turn module)
- "buffer" OR "accumulation" -> Buffer Conveyor
- "merge" + (30/45/60/90/deg/degree) -> Angle Merge Conveyor
- "align" OR "positioning" OR "singulat" -> Aligning Conveyor
- "spacing" OR "gap" -> Spacing / Loading Conveyor
- "weigh" OR "weight" OR "scale" -> Weighing Conveyor (and supports Barcode Scanning System)
- "receive" OR "receiving" -> Receiving Conveyor
- "intelligent merge" OR "merge" (if inside feedline context) -> Intelligent Merge Conveyor
- "chute" -> Output Chutes exist (then classify type by presence of words: mini/gravity/bulk/slide/live/reject/dispersion etc.)
- "disperson" OR "dispersion" -> Dispression / Dispersion Chutes
- "bag", "bagging", "ptl", "takeaway", "roller cage", "trolley" -> Bagging type PTL or Bag Takeaway Conveyor (choose best match by text)

COVERAGE RULES (to prevent missing obvious components)
- If any conveyor/infeed evidence exists (telescopic/infeed/ingresso/feedline/buffer/merge/etc.), Infeed System must be present.
  If you cannot prove a specific type, include "Infeed Conveyor" (generic) with low confidence.
- If any "feedline" / "spacing" / "weigh" / "merge" evidence exists, Feedlines must be present.
- If any "chute" evidence exists (block or layer), Output Chutes must be present.
  If chute type is generic/unclear, include "Gravity Chutes" with low confidence as a placeholder (type unknown), citing the generic chute evidence.
- If any "weigh"/"scan"/"barcode"/"dws" evidence exists, Barcode Scanning System must be present.

OUTPUT MINIMIZATION (MANDATORY)
- Output ONLY present items. Do NOT output any "Not Present" keys.
- Omit optional sections entirely if not present (sections 4, 6, 7).
- For mandatory sections (1,2,3,5), always include them in the fixed order.
- Use booleans true/false and a confidence score.

OUTPUT FORMAT (STRICT JSON ONLY, NO MARKDOWN)
Return a single JSON object with this schema:

{{
  "file": "<string>",
  "system_description_flow": [
    {{
      "section": "1. Infeed System",
      "present": true,
      "subcomponents": [
        {{
          "name": "<canonical name>",
          "confidence": 0.0-1.0,
          "evidence": [{{"source":"...", "path":"...", "matched_text":"...", "count": <number|null>}}],
          "notes": "<optional>"
        }}
      ]
    }},
    {{
      "section": "2. Induction to Sorter",
      "present": true,
      "subsections": [
        {{
          "section": "2.1 Feedlines",
          "present": true,
          "subcomponents": [ ...same structure... ]
        }}
      ]
    }},
    {{
      "section": "3. Loop CBS / Linear CBS",
      "present": true,
      "details": {{
        "type": "<Loop CBS|Linear CBS|Unknown>",
        "confidence": 0.0-1.0,
        "evidence": [ ... ]
      }}
    }},
    {{
      "section": "5. Output Chutes",
      "present": true,
      "subcomponents": [ ... ]
    }}
  ],
  "unmapped_signals": {{
    "blocks": [{{"name":"...", "count": <n>, "why_relevant":"..."}}],
    "layers": [{{"layer":"...", "total_entities": <n>, "why_relevant":"..."}}]
  }}
}}
**DO NOT ADD ``` or json or ```json``` or any other text in the response**
Now analyze this DXF evidence pack (verbatim) and output ONLY valid JSON:

{json.dumps(evidence_pack, ensure_ascii=False)}
""".strip()


# =========================
# GROQ CALL + JSON PARSE
# =========================

def parse_json_strict(text: str) -> dict:
    """
    Robustly parse JSON even if model adds stray text.
    """
    text = (text or "").strip()
    # try direct
    try:
        return json.loads(text)
    except Exception:
        pass

    # try extracting first {...} block
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        candidate = text[start:end+1]
        return json.loads(candidate)

    raise ValueError("Model response was not valid JSON.")


def groq_chat_json(api_key: str, prompt: str,
                   temperature: float = 0.0,
                   max_tokens: int = 3500,
                   retries: int = 2) -> dict:
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": "Return ONLY valid JSON. No markdown. No extra text."},
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    last_err = None
    for _ in range(retries + 1):
        try:
            r = requests.post(GROQ_ENDPOINT, headers=headers, json=payload, timeout=180)
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            return parse_json_strict(content)
        except Exception as e:
            last_err = e
            time.sleep(1.2)
    raise RuntimeError(f"Groq call failed: {last_err}")


# =========================
# BASIC UI
# =========================

def main():
    st.set_page_config(page_title="DXF → Components", layout="centered")
    uploaded = st.file_uploader("", type=["dxf"])

    if uploaded is None:
        # Keep UI minimal (no extra text)
        return

    api_key = (os.getenv("GROQ_API_KEY") or "").strip()
    if not api_key:
        st.error("Missing GROQ_API_KEY environment variable.")
        return

    # Write uploaded DXF to temp
    suffix = Path(uploaded.name).suffix or ".dxf"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tf:
        tf.write(uploaded.getvalue())
        temp_path = Path(tf.name)

    with st.spinner(""):
        full = extract_full_dxf_info(
            temp_path,
            include_layouts=True,
            nested_depth=25
        )
        evidence_pack = build_llm_evidence_pack(full)
        prompt = build_groq_prompt(evidence_pack)
        detected = groq_chat_json(api_key=api_key, prompt=prompt)

    # Show ONLY the Groq-generated detection JSON
    st.json(detected)


if __name__ == "__main__":
    main()

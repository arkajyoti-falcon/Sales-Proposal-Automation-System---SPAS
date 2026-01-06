# sys_desc_latest.py
# Streamlit app: DXF + Costing Excel -> System Description (DOCX)
# Uses GROQ (openai/gpt-oss-120b) for detection + writing + judge pass.
# Inserts ONLY Conveyor BOQ table (as marker [[CONVEYOR_BOQ_TABLE]] placed by LLM).

import os
import re
import json
import tempfile
from pathlib import Path
from collections import Counter
from typing import Dict, Any, List, Tuple, Optional

import requests
import ezdxf
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from dotenv import load_dotenv

load_dotenv()

# -----------------------------
# CONFIG
# -----------------------------
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
GROQ_BASE_URL = os.getenv("GROQ_BASE_URL", "https://api.groq.com/openai/v1/chat/completions")
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")

TEMPLATE_PATH = os.getenv("CBS_TEMPLATE_PATH", "CBS_SYSTEM_DESC.txt")

UNITS = {
    0: "Unitless", 1: "inches", 2: "feet", 3: "miles",
    4: "millimeters", 5: "centimeters", 6: "meters", 7: "kilometers"
}

# -----------------------------
# NOISE FILTERS (*U### / U### etc)
# -----------------------------
def is_noise_block(name: str) -> bool:
    if not name:
        return True
    n = name.strip()
    if re.match(r"^\*?U\d+$", n, flags=re.IGNORECASE):
        return True
    if re.match(r"^\*[UDXATE]\d+$", n, flags=re.IGNORECASE):
        return True
    if n.startswith("*") or n.startswith("~") or n.startswith("A$C"):
        return True
    return False


def safe_int(x) -> Optional[int]:
    try:
        return int(x)
    except Exception:
        return None


def load_template_text() -> str:
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return ""

# -----------------------------
# DXF FULL EXTRACTION
# -----------------------------
def _collect_from_space(space, top_inserts: Counter, layers: Counter, text_snips: Counter, entity_types: Counter):
    for e in space:
        # ...existing code...
        pass

    total = 0

    # Count from inserts (real quantities)
    for name, cnt in nested.items():
        if any(r.search(name) for r in rx):
            total += int(cnt)

    for name, cnt in top.items():
        if any(r.search(name) for r in rx):
            total += int(cnt)

    # If not found in inserts, fall back to layer/text/block_defs as presence signal (count=1)
    if total == 0:
        for lname in layers.keys():
            if any(r.search(lname) for r in rx):
                return 1
        for t in texts.keys():
            if any(r.search(t) for r in rx):
                return 1
        for bn in block_defs:
            if any(r.search(bn) for r in rx):
                return 1

    return total


def extract_angle_degrees(d: Dict[str, int]) -> List[int]:
    """
    Extract degrees from block names, e.g. "30°", "45 deg", "_90_deg", "(1000mm_90_deg)"
    Returns unique sorted degrees.
    """
    degrees = set()
    for name in d.keys():
        s = name.lower()

        # 30°, 45°, 60°
        for m in re.findall(r"(\d{1,3})\s*°", s):
            degrees.add(int(m))

        # 90 deg / 45deg / _90_deg
        for m in re.findall(r"(\d{1,3})\s*deg", s):
            degrees.add(int(m))
        for m in re.findall(r"_(\d{1,3})_deg", s):
            degrees.add(int(m))

        # "45 Deg Turn"
        for m in re.findall(r"(\d{1,3})\s*deg\s*turn", s):
            degrees.add(int(m))

    return sorted(degrees)


def compute_dxf_metrics(full: Dict[str, Any]) -> Dict[str, Any]:
    units_name = full.get("units_name") or "unknown"
    fname = (full.get("file") or "").lower()

    # CBS type from filename (fallback), LLM can refine from inserts too
    cbs_type = "Linear CBS" if "linear" in fname else "Loop CBS"

    # FEEDLINES (count)
    feedline_count = _search_counts_multi(full, [
        r"feedline", r"feed\s*line", r"fal.*feed", r"\bfs\d{3,4}\b", r"\bfs0\d+\b", r"induct",
        r"fal[_\-\s]*fs002", r"\bfs002\b", r"fs002v02", r"auto[_\-\s]*induct"
    ])

    # Special mapping: FAL_FS002V02 (Without weighing) => Buffer Conveyor
    fs002_without_weighing = _search_counts_multi(full, [
        r"fal[_\-\s]*fs002v02", r"\bfs002v02\b", r"without\s*weigh"
    ])

    # INFEED / TELESCOPIC
    infeed_telescopic = _search_counts_multi(full, [r"telescopic", r"telescopico", r"\btbc\b"])
    infeed_generic = _search_counts_multi(full, [
        r"infeed", r"ingresso", r"receiving", r"highway", r"singulat", r"pvc\s*belt"
    ])

    # ANGLES
    degrees = extract_angle_degrees(full.get("nested_inserts", {}) or {})

    # VDS LOOP (expanded patterns)
    vds_loop = _search_counts_multi(full, [
        r"\bvds\b",
        r"vds\s*loop",
        r"distribution\s*loop",
        r"distribut.*loop",
        r"static[_\-\s]*buffer",
        r"buffer.*loop", r"loop.*buffer",
        r"\bvds[_\-\s]*return\b",
        r"fal[_\-\s]*f001", r"\bf001\b",
        r"vipacsystem", r"vipac",
        r"return\s*line", r"return\s*conveyor",
        r"boom.*conv", r"\bboom\b",  # Boom conveyors are VDS components
        r"fal.*blk.*boom", r"blk.*boom",
    ])
    has_vds_loop = vds_loop > 0

    # RECIRCULATION
    recirc_count = _search_counts_multi(full, [
        r"recirculation", r"recirculate", r"refeed", r"return.*conv", r"loop.*back", r"re[-\s]*cir"
    ])
    has_recirculation = recirc_count > 0

    # CHUTES
    rejection_chute_count  = _search_counts_multi(full, [r"reject", r"rejection", r"sortfail", r"exception"])
    dispersion_chute_count = _search_counts_multi(full, [r"disperson", r"dispersion"])
    collection_chute_count = _search_counts_multi(full, [r"collection", r"friction", r"accumulation"])
    gravity_chute_count    = _search_counts_multi(full, [r"\bgravity\b"])
    mini_gravity_count     = _search_counts_multi(full, [r"mini.*gravity", r"chutes\$0\$mini\s*gravity"])
    chute_total            = _search_counts_multi(full, [r"chute", r"slide", r"sliding", r"discharge"])

    # SCANNER / DWS
    scan_count = _search_counts_multi(full, [
        r"scanner", r"\bscan\b", r"barcode", r"\bdws\b", r"dimension", r"reader"
    ])
    weighing_signal = _search_counts_multi(full, [r"weigh", r"weight", r"scale"])
    has_scanner = (scan_count > 0) or (weighing_signal > 0)

    # MANUAL INDUCT
    manual_station_count = _search_counts_multi(full, [
        r"manual", r"operator", r"workstation", r"induct.*station"
    ])
    has_manual = manual_station_count > 0

    # FEEDLINE SUBCOMPONENT SIGNALS (presence-driven) with actual counts
    loading_sig   = _search_counts_multi(full, [r"\bloading\b", r"\bspacing\b", r"\bspacer\b", r"gap\s*optimizer"])
    buffer_sig    = _search_counts_multi(full, [r"\bbuffer\b", r"accumulation"])
    receiving_sig = _search_counts_multi(full, [r"\breceiving\b", r"\binlet\b", r"\bintake\b"])
    merge_sig     = _search_counts_multi(full, [r"intelligent\s*merge", r"angle\s*merge", r"belt\s*merge", r"\bmerge\b"])
    reject_sig    = _search_counts_multi(full, [r"rejection\s*chute", r"\breject\b", r"sort\s*fail", r"exception"])
    
    # Get actual counts for induct conveyors (used for detailed induct section)
    loading_conveyor_count = _search_counts_multi(full, [
        r"loading\s*conveyor", r"orientation\s*conveyor", r"static.*loading"
    ])
    buffer_conveyor_count = _search_counts_multi(full, [
        r"buffer\s*conveyor", r"static.*buffer", r"accumulation\s*conveyor"
    ])
    intelligent_merge_count = _search_counts_multi(full, [
        r"intelligent\s*merge", r"angle\s*merge", r"static.*merge", 
        r"fal[_\-\s]*f007", r"fal[_\-\s]*f012", r"fal[_\-\s]*f002"
    ])
    weighing_conveyor_count = _search_counts_multi(full, [
        r"weighing\s*conveyor", r"static.*weighing", r"scale\s*conveyor",
        r"fal[_\-\s]*f015"
    ])
    spacing_conveyor_count = _search_counts_multi(full, [
        r"spacing\s*conveyor", r"gap.*conveyor", r"positioning\s*system",
        r"fal[_\-\s]*f013"
    ])
    receiving_conveyor_count = _search_counts_multi(full, [
        r"receiving\s*conveyor", r"static.*receiving", r"fal[_\-\s]*f003"
    ])

    # TOTAL “conveyor-like” (rough)
    conveyor_like_total = _search_counts_multi(full, [
        r"conv", r"conveyor", r"feedline", r"feed\s*line", r"buffer", r"merge", r"infeed", r"telescop", r"\btbc\b"
    ])

    return {
        "UNITS": units_name,
        "TYPE OF CBS": cbs_type,

        "FEEDLINE COUNT": feedline_count,

        "TEL. BELT CONVEYOR COUNT": infeed_telescopic,
        "INFEED CONVEYOR COUNT": infeed_generic,

        "VDS LOOP COUNT": vds_loop,
        "HAS_VDS_LOOP": has_vds_loop,

        "HAS_REIRCULATION": has_recirculation,
        "RECIRCULATION COUNT": recirc_count,

        "HAS_MANUAL_INDUCT": has_manual,
        "MANUAL STATION COUNT": manual_station_count,

        "HAS_SCANNER": has_scanner,
        "SCANNER SIGNAL COUNT": scan_count,
        "WEIGHING SIGNAL COUNT": weighing_signal,

        "COUNT OF REJECTION CHUTE": rejection_chute_count,
        "COUNT OF DISPERSION CHUTE": dispersion_chute_count,
        "COUNT OF COLLECTION CHUTE": collection_chute_count,

        "CHUTE TOTAL": chute_total,
        "GRAVITY CHUTE COUNT": gravity_chute_count,
        "MINI GRAVITY CHUTE COUNT": mini_gravity_count,

        "DEGREE OF ANGLE MERGE": ", ".join(str(x) for x in degrees) if degrees else "",

        "TOTAL CONVEYOR / SUBCOMPONENT COUNT": conveyor_like_total,

        # special mapping visibility
        "FS002 WITHOUT WEIGHING COUNT": fs002_without_weighing,

        # Induct subcomponent actual counts
        "LOADING CONVEYOR COUNT": loading_conveyor_count,
        "BUFFER CONVEYOR COUNT": buffer_conveyor_count,
        "INTELLIGENT MERGE COUNT": intelligent_merge_count,
        "WEIGHING CONVEYOR COUNT": weighing_conveyor_count,
        "SPACING CONVEYOR COUNT": spacing_conveyor_count,
        "RECEIVING CONVEYOR COUNT": receiving_conveyor_count,

        # feedline subcomponent presence flags
        "FEEDLINE_HAS_LOADING_OR_SPACING": loading_sig > 0,
        "FEEDLINE_HAS_BUFFER": (buffer_sig > 0) or (fs002_without_weighing > 0),
        "FEEDLINE_HAS_WEIGHING": weighing_signal > 0,
        "FEEDLINE_HAS_RECEIVING": receiving_sig > 0,
        "FEEDLINE_HAS_MERGE": (merge_sig > 0) or bool(degrees),
        "FEEDLINE_HAS_REJECTION": (reject_sig > 0) or (rejection_chute_count > 0),
    }


def build_detected_json(metrics: Dict[str, Any]) -> Dict[str, Any]:
    """
    Deterministic fallback detected JSON (present-only).
    Feedline module counts default to 1, but ONLY included if present in DXF signals.
    """
    det: Dict[str, Any] = {}

    # Infeed System
    infeed: Dict[str, Any] = {}
    if metrics.get("TEL. BELT CONVEYOR COUNT", 0) > 0:
        infeed["Telescopic Belt Conveyor"] = int(metrics["TEL. BELT CONVEYOR COUNT"])
    if metrics.get("INFEED CONVEYOR COUNT", 0) > 0:
        infeed["Infeed Conveyors"] = int(metrics["INFEED CONVEYOR COUNT"])
    if metrics.get("HAS_VDS_LOOP"):
        infeed["VDS Loop Conveyor"] = max(1, int(metrics.get("VDS LOOP COUNT", 1)))
    if infeed:
        det["Infeed System"] = infeed

    # Parcel Inducts / Induction to Sorter
    if metrics.get("FEEDLINE COUNT", 0) > 0:
        feedlines: Dict[str, Any] = {"Feedline Count": int(metrics["FEEDLINE COUNT"]), "Subcomponents": {}}

        # ONLY include if present
        if metrics.get("FEEDLINE_HAS_LOADING_OR_SPACING"):
            feedlines["Subcomponents"]["Loading / Spacing Conveyor"] = 1
        if metrics.get("FEEDLINE_HAS_RECEIVING"):
            feedlines["Subcomponents"]["Receiving Conveyor"] = 1
        if metrics.get("FEEDLINE_HAS_WEIGHING"):
            feedlines["Subcomponents"]["Weighing Conveyor"] = 1
        if metrics.get("FEEDLINE_HAS_BUFFER"):
            feedlines["Subcomponents"]["Buffer Conveyor"] = 1
        if metrics.get("FEEDLINE_HAS_MERGE"):
            feedlines["Subcomponents"]["Intelligent / Angle Merge Conveyor"] = 1
            if metrics.get("DEGREE OF ANGLE MERGE"):
                feedlines["Angle Merge Degrees"] = metrics["DEGREE OF ANGLE MERGE"]
        if metrics.get("FEEDLINE_HAS_REJECTION"):
            feedlines["Subcomponents"]["Rejection Chute"] = 1

        det["Parcel Inducts / Induction to Sorter"] = {"Feedlines": feedlines}

    if metrics.get("HAS_MANUAL_INDUCT"):
        det.setdefault("Parcel Inducts / Induction to Sorter", {})
        det["Parcel Inducts / Induction to Sorter"]["Manual Induct Stations"] = {
            "Manual Induct Station Count": int(metrics.get("MANUAL STATION COUNT", 1))
        }

    # CBS
    det["CBS"] = {"Type": metrics.get("TYPE OF CBS", "")}

    # Barcode scanning
    if metrics.get("HAS_SCANNER"):
        det["Barcode Scanning System"] = {"Present": True}

    # Output Chutes (text-only, no tables)
    out: Dict[str, Any] = {}
    if metrics.get("CHUTE TOTAL", 0) > 0:
        out["Total Chutes"] = int(metrics["CHUTE TOTAL"])
    if metrics.get("MINI GRAVITY CHUTE COUNT", 0) > 0:
        out["Mini-Gravity Chutes"] = int(metrics["MINI GRAVITY CHUTE COUNT"])
    if metrics.get("GRAVITY CHUTE COUNT", 0) > 0:
        out["Gravity Chutes"] = int(metrics["GRAVITY CHUTE COUNT"])
    if metrics.get("COUNT OF COLLECTION CHUTE", 0) > 0:
        out["Collection Chutes"] = int(metrics["COUNT OF COLLECTION CHUTE"])
    if metrics.get("COUNT OF DISPERSION CHUTE", 0) > 0:
        out["Dispersion Chutes"] = int(metrics["COUNT OF DISPERSION CHUTE"])
    if metrics.get("COUNT OF REJECTION CHUTE", 0) > 0:
        out["Rejection Chutes"] = int(metrics["COUNT OF REJECTION CHUTE"])
    if out:
        det["Output Chutes"] = out

    # Recirculation
    if metrics.get("HAS_REIRCULATION"):
        det["Recirculation"] = {"Recirculation Count": int(metrics.get("RECIRCULATION COUNT", 1))}

    return det

# -----------------------------
# COSTING EXCEL EXTRACTION (Loop CBS sheet + Conveyor BOQ)
# -----------------------------
def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def read_block_table(ws, header_row: int, ncols: int = None, max_rows: int = 200) -> List[List[str]]:
    """
    Read a table from Excel starting at header_row.
    If ncols is None, auto-detect by finding the last non-empty column in header row.
    """
    table: List[List[str]] = []
    hdr: List[str] = []
    
    # Auto-detect column count if not specified
    if ncols is None:
        ncols = 1
        for c in range(1, 50):  # Check up to 50 columns
            v = ws.cell(header_row, c).value
            if v is not None and str(v).strip() != "":
                ncols = c
            elif ncols > 1 and v is None:
                # Found the end (empty cell after non-empty cells)
                break
    
    for c in range(1, ncols + 1):
        v = ws.cell(header_row, c).value
        hdr.append("" if v is None else str(v).strip())
    table.append(hdr)

    for r in range(header_row + 1, header_row + 1 + max_rows):
        v0 = ws.cell(r, 1).value
        if v0 is None or str(v0).strip() == "":
            break
        row: List[str] = []
        for c in range(1, ncols + 1):
            v = ws.cell(r, c).value
            row.append("" if v is None else str(v).strip())
        table.append(row)

    return table


def extract_costing_tables(xlsx_path: Path) -> Dict[str, List[List[str]]]:
    """
    Keeps your original multi-table extraction.
    We will ONLY USE "Conveyor BOQ" in DOCX output, but we are not removing anything.
    """
    wb = load_workbook(filename=str(xlsx_path), data_only=True)
    tables: Dict[str, List[List[str]]] = {}

    if "Conveyors" in wb.sheetnames:
        ws = wb["Conveyors"]
        tables["Conveyor BOQ"] = read_block_table(ws, header_row=2, ncols=None)

        # Bagging Conveyor BOQ (filtered)
        full = tables["Conveyor BOQ"]
        hdr = full[0]
        name_idx = hdr.index("Name") if "Name" in hdr else 1
        bag_rows = [hdr]
        sno = 1
        for row in full[1:]:
            nm = (row[name_idx] or "").lower()
            if "bagging" in nm:
                row2 = row[:]
                row2[0] = str(sno)
                sno += 1
                bag_rows.append(row2)
        if len(bag_rows) > 1:
            tables["Bagging Conveyor BOQ"] = bag_rows

    # Fallback: try to find a conveyors-like sheet even if name differs
    if "Conveyor BOQ" not in tables:
        for sname in wb.sheetnames:
            if "convey" in sname.lower():
                ws = wb[sname]
                # try header row 1..10
                for hr in range(1, 11):
                    row_vals = [ws.cell(hr, c).value for c in range(1, 10)]
                    row_txt = " | ".join([str(v).strip().lower() for v in row_vals if v is not None])
                    if "name" in row_txt and ("qty" in row_txt or "quantity" in row_txt):
                        tables["Conveyor BOQ"] = read_block_table(ws, header_row=hr, ncols=7)
                        break
            if "Conveyor BOQ" in tables:
                break

    if "Inducts" in wb.sheetnames:
        tables["Feedlines BOQ"] = read_block_table(wb["Inducts"], header_row=2, ncols=7)

    if "Bag Sorter Induct" in wb.sheetnames:
        tables["Bag Induct BOQ"] = read_block_table(wb["Bag Sorter Induct"], header_row=2, ncols=7)

    if "Destinations" in wb.sheetnames:
        tables["Output Destinations"] = read_block_table(wb["Destinations"], header_row=2, ncols=6)

    if "Weighing Conveyors" in wb.sheetnames:
        tables["Weighing Conveyor BOQ"] = read_block_table(wb["Weighing Conveyors"], header_row=2, ncols=7)

    return tables


def extract_costing_values(xlsx_path: Path) -> Dict[str, str]:
    """
    Keeps your original Loop CBS extraction but makes it resilient.
    """
    wb = load_workbook(filename=str(xlsx_path), data_only=True)

    # Prefer exact "Loop CBS" sheet, else any sheet containing "loop cbs"
    ws = None
    if "Loop CBS" in wb.sheetnames:
        ws = wb["Loop CBS"]
    else:
        for s in wb.sheetnames:
            if "loop cbs" in s.lower():
                ws = wb[s]
                break
    if ws is None:
        ws = wb[wb.sheetnames[0]]

    carrier_pitch = ""
    sorter_height = ""

    for r in range(1, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and b.strip().lower() == "carrier pitch":
            v = ws.cell(r, 3).value
            if v is not None:
                carrier_pitch = str(v).strip()
                break

    for r in range(1, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and "select the sorter height" in b.strip().lower():
            v = ws.cell(r, 4).value or ws.cell(r, 3).value
            if v is not None:
                sorter_height = str(v).strip()
                break

    return {
        "COSTING_SHEET_NAME": ws.title,
        "CBS HEIGHT FROM GROUND": sorter_height,
        "PITCH LENGTH": carrier_pitch,
    }

# -----------------------------
# GROQ CALL
# -----------------------------
def groq_chat(messages: List[Dict[str, str]], temperature: float = 0.1, max_tokens: int = 4500) -> str:
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY is not set.")

    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    body = {
        "model": GROQ_MODEL,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    r = requests.post(GROQ_BASE_URL, headers=headers, data=json.dumps(body), timeout=180)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def extract_json(txt: str) -> Dict[str, Any]:
    if not txt:
        return {}
    txt = txt.strip()
    try:
        return json.loads(txt)
    except Exception:
        pass
    a = txt.find("{")
    b = txt.rfind("}")
    if a != -1 and b != -1 and b > a:
        try:
            return json.loads(txt[a:b+1])
        except Exception:
            return {}
    return {}


def normalize_detected(d: Dict[str, Any], fallback: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalize detected JSON from GROQ into canonical keys used by our prompts.
    If it doesn't match, fall back to deterministic detected JSON.
    """
    if not isinstance(d, dict) or not d:
        return fallback

    # If GROQ returned {"detected": {...}}
    if "detected" in d and isinstance(d["detected"], dict):
        d = d["detected"]

    # Minimal sanity: must have CBS type or any major component
    has_any = any(k in d for k in ["Infeed System", "CBS", "Parcel Inducts / Induction to Sorter", "Output Chutes"])
    if not has_any:
        return fallback

    return d

# -----------------------------
# PROMPTS
# -----------------------------
def prompt_detect_components(full_dxf: Dict[str, Any], metrics: Dict[str, Any]) -> List[Dict[str, str]]:
    system = r"""
ROLE
You are a DXF component detector for Cross-Belt Sorter (CBS) systems.

INPUTS YOU WILL RECEIVE
- DXF FULL JSON: insert names + nested inserts + layers + text snippets + block definitions.
- COMPUTED METRICS: precomputed counts/presence hints.

YOUR TASK
Return a single JSON object listing ONLY the components/subcomponents that are PRESENT in the system, with best-effort counts.

STRICT OUTPUT RULES
- Output MUST be a single VALID JSON object (no markdown, no commentary).
- Include ONLY PRESENT items. Do NOT output "Not present".
- Use the CANONICAL OUTPUT SCHEMA exactly as below.
- For feedline subcomponents, set count = 1 (module presence) unless DXF clearly indicates >1.
- Enforce mandatory mappings:
  1) If any block/text/layer contains "FAL_FS002V02" AND "(Without weighing)" OR "without weigh",
     then detect Feedlines -> Buffer Conveyor as present (count=1).
  2) If any block/text/layer indicates "VDS" or "distribution loop" OR matches common VDS patterns,
     then detect Infeed System -> VDS Loop Conveyor as present (count>=1).
  3) Detect CBS type:
     - If DXF indicates "Linear" -> Type="Linear CBS" else if "Loop" -> Type="Loop CBS".

CANONICAL OUTPUT SCHEMA (ONLY THESE KEYS)
{
  "Infeed System": {
    "<subcomponent name>": <count>
  },
  "Parcel Inducts / Induction to Sorter": {
    "Feedlines": {
      "Feedline Count": <int>,
      "Subcomponents": {
        "<subcomponent name>": <int>
      },
      "Angle Merge Degrees": "<string optional>"
    },
    "Manual Induct Stations": {
      "Manual Induct Station Count": <int>
    }
  },
  "CBS": { "Type": "Loop CBS or Linear CBS" },
  "Barcode Scanning System": { "Present": true },
  "Output Chutes": {
    "<chute type>": <int>
  },
  "Recirculation": { "Recirculation Count": <int> }
}

NOTES
- You may omit any top-level key if nothing is present for it.
- Prefer counts from nested inserts when available.
- If you only have presence evidence from layers/text/definitions, use count=1.

NOW RETURN THE JSON ONLY.
"""
    user = f"""DXF FULL JSON:
{json.dumps(full_dxf, ensure_ascii=False)}

COMPUTED METRICS:
{json.dumps(metrics, ensure_ascii=False)}
"""
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


# Your original (fixed-flow) prompt is kept (do not remove anything).
def prompt_generate_system_description(template_text: str,
                                       detected: Dict[str, Any],
                                       variables: Dict[str, str]) -> List[Dict[str, str]]:
    system = """
ROLE
You are a senior solution engineer writing the “System Description” section for a Cross-Belt Sorter (CBS) proposal.

INPUTS YOU WILL RECEIVE
1) TEMPLATE TEXT (canonical wording library; may contain [VAR] placeholders)
2) DETECTED JSON (the single source of truth for which components/subcomponents are PRESENT and their counts/types)
3) VARIABLES MAP (values for placeholders; some may be missing)

PRIMARY OBJECTIVE
Generate the final “System Description” using the TEMPLATE TEXT as the base wording, but include ONLY those sections and subcomponents that are PRESENT in DETECTED JSON.

NON-NEGOTIABLE HARD RULES
1) TEMPLATE ANCHORING
- Use TEMPLATE TEXT as the canonical description wording.
- Keep descriptions the SAME in meaning and style as the template.
- You may reorder sentences slightly ONLY to fit the fixed flow or to insert detected counts.
- Do NOT introduce new claims/specifications not supported by template or detected JSON.

2) PRESENCE FILTER
- Include ONLY components/subcomponents that exist in DETECTED JSON.
- Never mention missing components.
- Never write “Not present / Not available / not detected”.

3) PLACEHOLDERS
- Replace placeholders like [VAR] using VARIABLES MAP.
- If any placeholder is missing, keep it exactly as [VAR] (do not guess).
- Do not invent numbers or values.

4) IMAGE PLACEHOLDERS
- After each component/subcomponent description block, add on a new line:
  [IMAGE PLACEHOLDER: <exact component/subcomponent name>]

5) OUTPUT FORMAT
- Output plain text only. No JSON. No Markdown.

FIXED FLOW (MUST FOLLOW EXACT ORDER)
1. Infeed System
2. Induction to Sorter
   2.1 Feedlines
   2.2 Manual Induct Stations (if present)
3. Loop CBS / Linear CBS
4. Barcode Scanning System (if present)
5. Output Chutes
6. Exception Handling Area (if present)
7. Recirculation & Manual feedline (if present)
"""
    user = f"""TEMPLATE TEXT:
{template_text}

DETECTED JSON:
{json.dumps(detected, ensure_ascii=False)}

VARIABLES (already resolved):
{json.dumps(variables, ensure_ascii=False)}
"""
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


# NEW: Dynamic system description prompt (this is the one we will call)
def prompt_generate_system_description_dynamic(template_text: str,
                                               detected: Dict[str, Any],
                                               variables: Dict[str, str]) -> List[Dict[str, str]]:
    system = r"""
ROLE
You are a senior solution engineer writing the “System Description” section for a Cross-Belt Sorter (CBS) proposal.

INPUTS
1) TEMPLATE TEXT: canonical wording library (may contain [VAR] placeholders)
2) DETECTED JSON: the ONLY source of truth for what is present
3) VARIABLES MAP: values for placeholders

CORE STYLE TARGET
Match typical CBS proposal “System Description” style:
- Sections may vary and order is NOT fixed across proposals.
- The flow must remain operationally logical.
- Infeed System usually contains multiple sub-sections and ends with “Conveyor BOQ” as the LAST subsection under Infeed System.

HARD RULES
1) DYNAMIC ORDER (NOT FIXED)
- Decide the best order based on DETECTED JSON, similar to professional CBS PDFs.
- Keep a logical progression (infeed → induction/feedlines → main sorter → scanning → outputs → optional areas).

2) PRESENCE FILTER
- Include ONLY items that exist in DETECTED JSON. Never mention missing items.

3) PLACEHOLDERS
- Replace [VAR] using VARIABLES MAP when available.
- If missing, keep [VAR] unchanged. Do NOT guess.

4) MAIN SORTER HEADING
- If DETECTED JSON -> CBS -> Type is "Loop CBS": use heading exactly "Main Loop".
- If Type is "Linear CBS": use heading exactly "Main Linear CBS".

5) CONVEYOR BOQ TABLE MARKER
- Under Infeed System, include a subsection heading exactly: “Conveyor BOQ”
- It MUST be the LAST subsection inside Infeed System.
- Immediately after that heading, output a single line marker exactly:
  [[CONVEYOR_BOQ_TABLE]]
- Do NOT output any other table markers elsewhere.

6) FEEDLINE SUBCOMPONENT SUBSECTIONS
- Under Feedlines, you may add sub-subsections ONLY for names present in:
  DETECTED JSON -> "Parcel Inducts / Induction to Sorter" -> "Feedlines" -> "Subcomponents"
- If a subcomponent is not present there, do NOT create its subsection.
- If you mention a feedline subcomponent, keep its count as 1 module unless DETECTED JSON provides another count.

7) OUTPUT CHUTES
- Describe Output Chutes using text only (counts ok). No tables.

8) IMAGE PLACEHOLDERS
- After each SECTION or SUBSECTION description block, add:
  [IMAGE PLACEHOLDER: <exact heading text>]
- Use placeholders only. Do not embed images.

9) OUTPUT FORMAT
- Output plain text only. No markdown. No JSON.
- Use clean hierarchical numbering that suits the document (e.g., 1 / 1.1 / 1.1.1).
- Headings must be on their own line.

TEMPLATE USAGE
- Prefer TEMPLATE TEXT wording for each component/subcomponent.
- Keep meaning and tone the same.
- If template does not contain an exact paragraph for a detected item, write a short neutral paragraph (2–3 lines) without adding new specs.

NOW GENERATE the system description.
"""
    user = f"""TEMPLATE TEXT:
{template_text}

DETECTED JSON:
{json.dumps(detected, ensure_ascii=False)}

VARIABLES:
{json.dumps(variables, ensure_ascii=False)}
"""
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


# NEW: Judge pass to enforce structure before DOCX
def prompt_judge_fix_system_description(detected: Dict[str, Any], draft_text: str) -> List[Dict[str, str]]:
    system = r"""
ROLE
You are a strict QA judge for CBS System Description text.

INPUTS
- DETECTED JSON (truth)
- DRAFT TEXT

YOUR JOB
- If the draft violates ANY rule, rewrite it to comply.
- If it already complies, return it unchanged.

RULES TO ENFORCE (STRICT)
1) Include ONLY items present in DETECTED JSON.
2) Infeed System must contain "Conveyor BOQ" as the LAST subsection under Infeed System, and it must contain marker [[CONVEYOR_BOQ_TABLE]] on the next line.
3) Main sorter heading:
   - Loop CBS -> "Main Loop"
   - Linear CBS -> "Main Linear CBS"
4) Feedline subcomponent subsections ONLY if present in DETECTED JSON feedline subcomponents.
5) No tables anywhere except the single marker [[CONVEYOR_BOQ_TABLE]].
6) Every heading block must be followed by an [IMAGE PLACEHOLDER: ...] line.
7) Output must be plain text only.

RETURN ONLY the final corrected text (no explanations).
"""
    user = f"""DETECTED JSON:
{json.dumps(detected, ensure_ascii=False)}

DRAFT TEXT:
{draft_text}
"""
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]

# -----------------------------
# DOCX WRITER
# -----------------------------
def add_docx_table(doc: Document, table_data: List[List[str]]):
    if not table_data or len(table_data) < 2:
        return
    rows = len(table_data)
    cols = len(table_data[0])
    t = doc.add_table(rows=rows, cols=cols)
    t.style = "Table Grid"
    for r in range(rows):
        for c in range(cols):
            t.cell(r, c).text = table_data[r][c]
    doc.add_paragraph("")


# Keeps your signature (tables + detected) so you don’t break older calls.
# But behavior is enhanced:
# - Only inserts Conveyor BOQ at marker [[CONVEYOR_BOQ_TABLE]]
# - No other tables inserted
def build_docx(system_description_text: str,
               out_path: Path,
               title: str = "System Description",
               tables: Optional[Dict[str, List[List[str]]]] = None,
               detected: Optional[Dict[str, Any]] = None) -> None:

    tables = tables or {}
    detected = detected or {}

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Title
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(16)

    doc.add_paragraph("")

    conveyor_boq = tables.get("Conveyor BOQ", [])

    for line in system_description_text.splitlines():
        if line.strip() == "[[CONVEYOR_BOQ_TABLE]]":
            add_docx_table(doc, conveyor_boq)
            continue

        doc.add_paragraph(line)

    doc.save(str(out_path))

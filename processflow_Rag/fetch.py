
# dxf_ingest_pinecone_no_grpc.py
import os
import re
import tempfile
import json
import time
from pathlib import Path
from collections import Counter, defaultdict
from typing import Any, List
from datetime import datetime
import logging

import streamlit as st
from dotenv import load_dotenv

import ezdxf
from docx import Document

# Pinecone (modern)
from pinecone import Pinecone, ServerlessSpec

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("dxf-ingest-no-grpc")

# ----------------- CONFIG -----------------
PINECONE_API_KEY = os.getenv("PINECONE_API_KEY", "pcsk_8akoe_FxzXaW2zvAsEd1uiHqxiMrosvumSujgFyrWAB9vyqG87DGWpnDc6rSxaDYrkP3v")
PINECONE_INDEX_NAME = os.getenv("PINECONE_INDEX_NAME", "spas-dxf-samples")
PINECONE_CLOUD = os.getenv("PINECONE_CLOUD", "aws")
PINECONE_REGION = os.getenv("PINECONE_REGION", "us-east-1")
PINECONE_HOST = os.getenv("PINECONE_HOST","https://spas-dxf-samples-cxyv4ex.svc.aped-4627-b74a.pinecone.io")  # optional

GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

# Pinecone uses OpenAI's text-embedding-3-large via inference
EMBED_MODEL = "llama-text-embed-v2"
EMBED_DIM = 1048
EMBED_NAMESPACE = "v1-dxf"


# ---------------- DXF extraction helpers (same logic you provided) ----------------
UNITS = {
    0: "Unitless", 1: "inches", 2: "feet", 3: "miles",
    4: "millimeters", 5: "centimeters", 6: "meters", 7: "kilometers",
}

COMPONENT_PATTERNS = {
    "AUTO_INDUCT": [r"fal.*fs\d+", r"fal.*feed", r"fal.*induct", r"auto.*induct", r"feed.*line", r"feedline", r"fs\d{3}"],
    "CONVEYOR_INFEED": [r"telescopic", r"infeed.*conv", r"in.*feed", r"receiving.*conv", r"inclined.*conv", r"incline", r"elevation.*conv"],
    "VDS_BUFFER": [r"vds", r"distribution.*loop", r"buffer", r"arm.*vds"],
    "OPERATOR_STATION": [r"operator(?!.*safety)", r"manual.*station", r"induct.*station"],
    "CHUTE": [r"^chute", r"gravity.*chute", r"live.*chute", r"slide.*chute", r"sliding.*chute", r"reject.*chute", r"collection.*chute", r"mini.*chute", r"bulk.*chute", r"discharge"],
    "PTL": [r"ptl", r"put.*to.*light", r"pick.*to.*light", r"light.*rack"],
    "BAG_SYSTEM": [r"bag.*conv", r"bag.*takeaway", r"bagging"],
    "RECIRCULATION": [r"recirculation", r"recirculate", r"refeed", r"return.*conv"],
    "CBS_SORTER": [r"cbs", r"cross.*belt", r"crossbelt", r"sorter.*module", r"carrier", r"loop.*sorter"],
    "SCANNER": [r"scanner", r"scan.*tunnel", r"barcode.*read", r"dimension.*sys", r"dws", r"volume.*scan"],
}

STRUCTURAL_PATTERNS = [r"leg.*guard", r"guard(?!.*operator)", r"fenc", r"railing", r"handrail", r"safety.*(?!operator)", r"pallet(?!.*conv)",
                       r"step", r"stair", r"ladder", r"door", r"panel(?!.*control)", r"bracket", r"bolt", r"mount(?!.*scanner)"]

def _is_noise_block(name: str) -> bool:
    n = name.strip()
    if re.match(r"^\*[UDXATE]\d+$", n, re.IGNORECASE): return True
    if n.startswith("*") or n.startswith("~"): return True
    return False

def _is_structural(name: str) -> bool:
    n_lower = name.lower()
    for pattern in STRUCTURAL_PATTERNS:
        if re.search(pattern, n_lower): return True
    return False

def _categorize_component(name: str) -> str:
    n_lower = name.lower()
    for category, patterns in COMPONENT_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, n_lower):
                return category
    return "UNCATEGORIZED"

def _normalize_group_name(name: str) -> str:
    n = name.strip()
    if "|" in n: n = n.split("|")[-1]
    if not n.startswith("FAL"):
        n = re.sub(r"[_\-]+", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    if not re.search(r"V\d+$", n, re.IGNORECASE):
        n = re.sub(r"\s*\(?\d+\)?$", "", n).strip()
    return n.lower()

def _detect_cbs_type(project_name: str) -> str:
    pn_lower = project_name.lower()
    if "linear" in pn_lower or "linear cbs" in pn_lower or "linear sorter" in pn_lower: return "Linear CBS"
    if "loop cbs" in pn_lower or "loop sorter" in pn_lower: return "Loop CBS"
    return "Loop CBS"

def _analyze_chute_types(components: dict) -> dict:
    chute_analysis = {"total": 0, "by_type": defaultdict(int), "has_type_info": False}
    for comp_name, count in components.items():
        name_lower = comp_name.lower()
        if "chute" in name_lower:
            chute_analysis["total"] += count
            if "gravity" in name_lower or "collection" in name_lower:
                chute_analysis["by_type"]["gravity"] += count; chute_analysis["has_type_info"] = True
            elif "live" in name_lower or "active" in name_lower:
                chute_analysis["by_type"]["live"] += count; chute_analysis["has_type_info"] = True
            elif "slide" in name_lower or "sliding" in name_lower:
                chute_analysis["by_type"]["sliding"] += count; chute_analysis["has_type_info"] = True
            elif "reject" in name_lower or "exception" in name_lower:
                chute_analysis["by_type"]["rejection"] += count; chute_analysis["has_type_info"] = True
            elif "mini" in name_lower:
                chute_analysis["by_type"]["mini_gravity"] += count; chute_analysis["has_type_info"] = True
            elif "bulk" in name_lower:
                chute_analysis["by_type"]["bulk"] += count; chute_analysis["has_type_info"] = True
            elif "non-sort" in name_lower or "nonsort" in name_lower:
                chute_analysis["by_type"]["non_sort"] += count; chute_analysis["has_type_info"] = True
    return chute_analysis

def extract_dxf_components(dxf_path: Path, project_name: str = "") -> dict:
    """Extract and categorize components from DXF file (compact but faithful to the ingest script)."""
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    hdr = doc.header

    units_code = hdr.get("$INSUNITS", None)
    try: units_code = int(units_code) if units_code is not None else None
    except: units_code = None

    extmin = hdr.get("$EXTMIN", None)
    extmax = hdr.get("$EXTMAX", None)

    raw_counts: Counter[str] = Counter()
    for e in msp:
        try:
            if e.dxftype() == "INSERT":
                bname = e.dxf.name
                if not _is_noise_block(bname) and not _is_structural(bname):
                    raw_counts[bname] += 1
        except Exception:
            continue

    categorized: dict = defaultdict(lambda: defaultdict(lambda: {"count": 0, "examples": []}))
    for raw_name, cnt in raw_counts.items():
        category = _categorize_component(raw_name)
        gname = _normalize_group_name(raw_name)
        categorized[category][gname]["count"] += cnt
        categorized[category][gname]["examples"].append(raw_name)

    cbs_type = _detect_cbs_type(project_name if project_name else dxf_path.name)
    chute_analysis = _analyze_chute_types(raw_counts)

    has_auto_induct = len(categorized.get("AUTO_INDUCT", {})) > 0
    has_operators = len(categorized.get("OPERATOR_STATION", {})) > 0
    has_vds = len(categorized.get("VDS_BUFFER", {})) > 0

    if has_auto_induct and has_operators:
        induction_type = "MIXED (Auto + Manual)"
    elif has_auto_induct:
        induction_type = "AUTO"
    elif has_operators:
        induction_type = "MANUAL"
    else:
        induction_type = "UNKNOWN"

    category_summary = {}
    total_components = 0
    for cat, items in categorized.items():
        count = sum(item["count"] for item in items.values())
        category_summary[cat] = count
        total_components += count

    return {
        "file": dxf_path.name,
        "units_code": units_code,
        "units_name": UNITS.get(units_code, "unknown") if units_code is not None else None,
        "extents": {"min": list(extmin) if extmin else None, "max": list(extmax) if extmax else None},
        "cbs_type": cbs_type,
        "induction_type": induction_type,
        "has_vds": has_vds,
        "total_components": total_components,
        "category_summary": dict(category_summary),
        "categorized_components": {cat: {name: data["count"] for name, data in items.items()} for cat, items in categorized.items()},
        "chute_analysis": chute_analysis,
        "raw_block_counts": {k: int(v) for k, v in raw_counts.items()},
    }

def _summarise_components_for_prompt(dxf_json: dict) -> str:
    """Compact human-readable summary (same as ingest app) for embedding/prompting."""
    lines = ["=" * 70]
    lines.append("DXF COMPONENT ANALYSIS FOR CBS PROCESS FLOW GENERATION")
    lines.append("=" * 70)
    lines.append("")
    lines.append(f"FILE: {dxf_json.get('file', 'Unknown')}")
    lines.append(f"UNITS: {dxf_json.get('units_name', 'Unknown')}")
    lines.append(f"CBS TYPE: {dxf_json.get('cbs_type', 'Unknown')} (detected from filename)")
    lines.append(f"TOTAL COMPONENTS: {dxf_json.get('total_components', 0)}")
    lines.append("")
    lines.append("SYSTEM CONFIGURATION:")
    lines.append(f"  • Induction Type: {dxf_json.get('induction_type', 'Unknown')}")
    lines.append(f"  • VDS/Buffer System: {'YES' if dxf_json.get('has_vds') else 'NO'}")
    chute_analysis = dxf_json.get('chute_analysis', {})
    if chute_analysis.get('total', 0) > 0:
        lines.append(f"  • Total Chutes: {chute_analysis['total']}")
        if chute_analysis.get('has_type_info'):
            lines.append("  • Chute Types Detected: YES")
        else:
            lines.append("  • Chute Types Detected: NO (describe generically)")
    lines.append("")
    cat_summary = dxf_json.get("category_summary", {})
    if cat_summary:
        lines.append("COMPONENT CATEGORIES:")
        priority = ["AUTO_INDUCT", "OPERATOR_STATION", "VDS_BUFFER", "CONVEYOR_INFEED",
                    "CBS_SORTER", "CHUTE", "PTL", "BAG_SYSTEM", "RECIRCULATION", "SCANNER"]
        for cat in priority:
            if cat in cat_summary:
                lines.append(f"  • {cat}: {cat_summary[cat]} units")
        for cat, count in cat_summary.items():
            if cat not in priority:
                lines.append(f"  • {cat}: {count} units")
        lines.append("")
    if chute_analysis.get('by_type'):
        lines.append("CHUTE TYPE BREAKDOWN:")
        for ctype, count in sorted(chute_analysis['by_type'].items(), key=lambda x: -x[1]):
            lines.append(f"  • {ctype.replace('_', ' ').title()}: {count} chutes")
        lines.append("")
    lines.append("=" * 70)
    lines.append("PROCESS FLOW GENERATION GUIDANCE:")
    lines.append(f"  • System Type: {dxf_json.get('cbs_type', 'Unknown')}")
    lines.append(f"  • Induction: {dxf_json.get('induction_type', 'Unknown')}")
    if dxf_json.get('has_vds'):
        lines.append("  • Include VDS/Buffer section in Infeed System")
    if dxf_json.get('induction_type') == "MIXED (Auto + Manual)":
        lines.append("  • Include BOTH Auto Induct Line AND Manual Induct Station sections")
    elif dxf_json.get('induction_type') == "AUTO":
        lines.append("  • Include Auto Induct Line section only")
    elif dxf_json.get('induction_type') == "MANUAL":
        lines.append("  • Include Manual Induct Station section only")
    if chute_analysis.get('has_type_info'):
        lines.append("  • Chute types available - use detailed breakdown")
    else:
        lines.append("  • Chute types NOT available - describe generically")
    if "PTL" in cat_summary:
        lines.append("  • Include Put To Light System section")
    if "BAG_SYSTEM" in cat_summary:
        lines.append("  • Include Bag Takeaway Conveyor section")
    if "RECIRCULATION" in cat_summary:
        lines.append("  • Include Recirculation/Exception Refeeding section")
    lines.append("=" * 70)
    return "\n".join(lines)

# ---------------- Pinecone helpers ----------------
def get_pinecone_client_and_index(create_if_missing: bool = False):
    if not PINECONE_API_KEY:
        raise RuntimeError("PINECONE_API_KEY is not set.")
    pc = Pinecone(api_key=PINECONE_API_KEY)
    try:
        raw = pc.list_indexes()

        # Newer pinecone client has .names()
        if hasattr(raw, "names"):
            existing = set(raw.names())
        # Raw list of strings or objects/dicts
        elif isinstance(raw, list):
            names = []
            for item in raw:
                if isinstance(item, str):
                    names.append(item)
                elif isinstance(item, dict) and item.get("name"):
                    names.append(item.get("name"))
                elif hasattr(item, "name"):
                    names.append(getattr(item, "name"))
            existing = set(names)
        # Dict response shape (older SDKs)
        elif isinstance(raw, dict):
            existing = {i.get("name") for i in raw.get("indexes", []) if i.get("name")}
        else:
            existing = set()
    except Exception as e:
        raise RuntimeError(f"Failed to list Pinecone indexes: {e}")

    if PINECONE_INDEX_NAME not in existing:
        if not create_if_missing:
            raise RuntimeError(f"Index '{PINECONE_INDEX_NAME}' not found. Set create_if_missing=True if you want to auto-create.")
        try:
            spec = ServerlessSpec(cloud="aws", region="us-east-1")
            pc.create_index(name=PINECONE_INDEX_NAME, dimension=EMBED_DIM, metric="cosine", spec=spec)
        except Exception:
            pc.create_index(name=PINECONE_INDEX_NAME, dimension=EMBED_DIM, metric="cosine")
        time.sleep(1.2)

    index = pc.Index(PINECONE_INDEX_NAME)
    return pc, index

def _extract_vector_from_item(item: Any):
    if isinstance(item, dict):
        for k in ("values", "embedding", "vector", "embeddings"):
            if k in item and item[k] is not None:
                v = item[k]
                if hasattr(v, "tolist"):
                    try: return v.tolist()
                    except: pass
                if isinstance(v, (list, tuple)): return list(v)
                if isinstance(v, (float, int)): return [float(v)]
    for attr in ("values", "embedding", "vector", "embeddings"):
        if hasattr(item, attr):
            v = getattr(item, attr)
            if hasattr(v, "tolist"):
                try: return v.tolist()
                except: pass
            if isinstance(v, (list, tuple)): return list(v)
            if isinstance(v, (float, int)): return [float(v)]
    if isinstance(item, (list, tuple)): return list(item)
    return None

def pinecone_embed(pc: Pinecone, texts: List[str], model: str = EMBED_MODEL, input_type: str = "passage", truncate: str = "END") -> List[List[float]]:
    if not texts: return []
    inputs_payload = [{"text": t} for t in texts]
    parameters = {"input_type": input_type}
    if truncate: parameters["truncate"] = truncate
    try:
        resp = pc.inference.embed(model=model, inputs=inputs_payload, parameters=parameters)
    except Exception as e:
        # fallback to local sentence-transformers if available
        if _SENTENCE_MODEL is not None:
            arr = _SENTENCE_MODEL.encode(texts, show_progress_bar=False, convert_to_numpy=True).tolist()
            return arr
        raise RuntimeError(f"Pinecone inference.embed failed and no local fallback available: {repr(e)}")
    vectors = []
    if hasattr(resp, "data") and isinstance(resp.data, list):
        for item in resp.data:
            v = _extract_vector_from_item(item)
            if v is None:
                raise RuntimeError("Unable to extract vector from Pinecone inference response item.")
            vectors.append([float(x) for x in v])
        return vectors
    if isinstance(resp, dict) and "data" in resp and isinstance(resp["data"], list):
        for item in resp["data"]:
            v = _extract_vector_from_item(item)
            if v is None:
                raise RuntimeError("Unable to extract vector from resp['data'] item.")
            vectors.append([float(x) for x in v])
        return vectors
    if isinstance(resp, list):
        if not resp: return []
        if isinstance(resp[0], (list, tuple)):
            for it in resp: vectors.append([float(x) for x in it])
            return vectors
        else:
            return [[float(x) for x in resp]]
    # last resort
    raise RuntimeError("Unrecognized embed response shape from Pinecone inference.")

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="DXF → Top-K Pinecone Lookup", layout="centered")
st.title("Query Pinecone: Upload DXF → show top-K matching process-flows")

st.markdown(
    "Upload a DXF file. This tool extracts a DXF summary, computes an embedding, queries your Pinecone index and returns top-K matches (process_flow snippets + metadata)."
)

with st.form("query"):
    uploaded_dxf = st.file_uploader("Upload DXF (.dxf)", type=["dxf"])
    project_name = st.text_input("Project name (optional)", value="")
    k = st.number_input("Top K", min_value=1, max_value=10, value=2)
    create_index = st.checkbox("Create index if missing (admin)", value=False)
    submit = st.form_submit_button("Find top matches")

if submit:
    if not uploaded_dxf:
        st.error("Please upload a DXF file.")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        dxf_path = td / uploaded_dxf.name
        dxf_path.write_bytes(uploaded_dxf.read())

        try:
            dxf_json = extract_dxf_components(dxf_path, project_name=project_name or uploaded_dxf.name)
            dxf_summary = _summarise_components_for_prompt(dxf_json)
        except Exception as e:
            st.exception(f"DXF extraction failed: {e}")
            st.stop()

        st.subheader("DXF summary (used for retrieval)")
        st.code(dxf_summary[:4000])  # show first chunk
        st.download_button("Download full DXF summary", data=dxf_summary, file_name=f"{Path(uploaded_dxf.name).stem}_summary.txt")

        # pinecone client + index
        try:
            pc, index = get_pinecone_client_and_index(create_if_missing=create_index)
        except Exception as e:
            st.exception(f"Pinecone init failed: {e}")
            st.stop()

        st.info("Computing embedding (Pinecone inference)...")
        try:
            vectors = pinecone_embed(pc, [dxf_summary], model=EMBED_MODEL, input_type="passage", truncate="END")
            vec = vectors[0]
        except Exception as e:
            st.exception(f"Embedding failed: {e}")
            st.stop()

        st.info(f"Querying index '{PINECONE_INDEX_NAME}' (top_k={k}) ...")
        try:
            # modern query signature: vector=..., top_k=..., include_metadata=True
            resp = index.query(vector=vec, top_k=int(k), include_metadata=True, namespace=EMBED_NAMESPACE, include_values=False)
        except Exception as e:
            st.exception(f"Index query failed: {e}")
            st.stop()

        # normalize matches
        matches = []
        if hasattr(resp, "matches"):
            matches = resp.matches
        elif isinstance(resp, dict) and "matches" in resp:
            matches = resp["matches"]
        elif isinstance(resp, list):
            matches = resp
        else:
            st.warning("No matches found (unexpected response shape).")
            matches = []

        if not matches:
            st.info("No matches returned by index.")
            st.stop()

        st.success(f"Top {min(len(matches), int(k))} matches:")
        for i, m in enumerate(matches[:k], start=1):
            # extract id, score, metadata robustly
            if isinstance(m, dict):
                mid = m.get("id")
                score = m.get("score") or m.get("distance") or m.get("similarity")
                meta = m.get("metadata") or {}
            else:
                # object with attributes
                mid = getattr(m, "id", None)
                score = getattr(m, "score", None) or getattr(m, "distance", None) or getattr(m, "similarity", None)
                meta = getattr(m, "metadata", {}) or {}

            st.markdown(f"### Match #{i}: `{mid}`  — score: **{score:.4f}**")
            # metadata preview
            pf = meta.get("process_flow") or meta.get("process_flow_snippet") or meta.get("dxf_summary") or meta.get("dxf_summary_compact") or ""
            client = meta.get("client") or meta.get("project_name") or meta.get("dxf_file")
            st.write("Client / id:", client)
            if pf:
                # show a short snippet and provide download
                snippet = pf if len(pf) < 2000 else pf[:2000] + "..."
                st.text_area("Process flow snippet", value=snippet, height=200)
                st.download_button(f"Download process flow (match #{i})", data=pf, file_name=f"{mid}_process_flow.txt")
            else:
                st.write("No process_flow text found in metadata for this match.")
            # show full metadata (sanitized)
            md_preview = {k: meta[k] for k in meta.keys() if k in ("dxf_file", "cbs_type", "induction_type", "total_components", "created_at", "client", "project_name")}
            st.json(md_preview)

        # prepare combined RAG payload (concatenate top-K process flows + new dxf summary)
        rag_pieces = []
        for m in matches[:k]:
            meta = m.metadata if hasattr(m, "metadata") else (m.get("metadata") if isinstance(m, dict) else {})
            pf = meta.get("process_flow") or meta.get("process_flow_snippet") or ""
            if pf:
                rag_pieces.append(f"---MATCH---\n{pf}\n")

        rag_prompt = "\n\n".join(rag_pieces) + "\n\n---INPUT DXF SUMMARY---\n" + dxf_summary
        st.subheader("RAG payload (top matches + current DXF summary)")
        st.text_area("RAG payload (ready to send to your LLM/groq)", value=rag_prompt[:20000], height=300)
        st.download_button("Download RAG payload", data=rag_prompt, file_name=f"{Path(uploaded_dxf.name).stem}_rag_payload.txt")

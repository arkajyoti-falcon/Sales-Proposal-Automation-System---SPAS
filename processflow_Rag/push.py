
# # dxf_ingest_pinecone_no_grpc.py
# import os
# import re
# import tempfile
# import json
# import time
# from pathlib import Path
# from collections import Counter, defaultdict
# from datetime import datetime
# import logging

# import streamlit as st
# from dotenv import load_dotenv

# import ezdxf
# from docx import Document

# # Pinecone (modern)
# from pinecone import Pinecone, ServerlessSpec

# load_dotenv()
# logging.basicConfig(level=logging.INFO)
# logger = logging.getLogger("dxf-ingest-no-grpc")

# # ----------------- CONFIG -----------------
# PINECONE_API_KEY = os.getenv("PINECONE_API_KEY", "pcsk_8akoe_FxzXaW2zvAsEd1uiHqxiMrosvumSujgFyrWAB9vyqG87DGWpnDc6rSxaDYrkP3v")
# PINECONE_INDEX_NAME = os.getenv("PINECONE_INDEX_NAME", "spas-dxf-samples")
# PINECONE_CLOUD = os.getenv("PINECONE_CLOUD", "aws")
# PINECONE_REGION = os.getenv("PINECONE_REGION", "us-east-1")
# PINECONE_HOST = os.getenv("PINECONE_HOST","https://spas-dxf-samples-cxyv4ex.svc.aped-4627-b74a.pinecone.io")  # optional

# GROQ_API_KEY = os.getenv("GROQ_API_KEY")
# GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

# # Pinecone uses OpenAI's text-embedding-3-large via inference
# EMBED_MODEL = "llama-text-embed-v2"
# EMBED_DIM = 1048
# EMBED_NAMESPACE = "v1-dxf"


# # ----------------- DXF extractor helpers (condensed, from your code) -----------------
# UNITS = {0:"Unitless",1:"inches",2:"feet",3:"miles",4:"millimeters",5:"centimeters",6:"meters",7:"kilometers"}

# COMPONENT_PATTERNS = {
#     "AUTO_INDUCT":[r"fal.*fs\d+",r"fal.*feed",r"fal.*induct",r"auto.*induct",r"feed.*line",r"feedline",r"fs\d{3}"],
#     "CONVEYOR_INFEED":[r"telescopic",r"infeed.*conv",r"in.*feed",r"receiving.*conv",r"inclined.*conv",r"incline",r"elevation.*conv"],
#     "VDS_BUFFER":[r"vds",r"distribution.*loop",r"buffer",r"arm.*vds"],
#     "OPERATOR_STATION":[r"operator(?!.*safety)",r"manual.*station",r"induct.*station"],
#     "CHUTE":[r"^chute",r"gravity.*chute",r"live.*chute",r"slide.*chute",r"sliding.*chute",r"reject.*chute",r"collection.*chute",r"mini.*chute",r"bulk.*chute",r"discharge"],
#     "PTL":[r"ptl",r"put.*to.*light",r"pick.*to.*light",r"light.*rack"],
#     "BAG_SYSTEM":[r"bag.*conv",r"bag.*takeaway",r"bagging"],
#     "RECIRCULATION":[r"recirculation",r"recirculate",r"refeed",r"return.*conv"],
#     "CBS_SORTER":[r"cbs",r"cross.*belt",r"crossbelt",r"sorter.*module",r"carrier",r"loop.*sorter"],
#     "SCANNER":[r"scanner",r"scan.*tunnel",r"barcode.*read",r"dimension.*sys",r"dws",r"volume.*scan"]
# }
# STRUCTURAL_PATTERNS = [
#     r"leg.*guard", r"guard(?!.*operator)", r"fenc", r"railing", r"handrail",
#     r"safety.*(?!operator)", r"pallet(?!.*conv)", r"step", r"stair", r"ladder",
#     r"door", r"panel(?!.*control)", r"bracket", r"bolt", r"mount(?!.*scanner)"
# ]

# def _is_noise_block(name: str) -> bool:
#     """Filter out anonymous noise blocks."""
#     n = name.strip()
#     if re.match(r"^\*[UDXATE]\d+$", n, re.IGNORECASE):
#         return True
#     if n.startswith("*") or n.startswith("~"):
#         return True
#     return False

# def _is_structural(name: str) -> bool:
#     """Check if component is structural (non-flow)."""
#     n_lower = name.lower()
#     for pattern in STRUCTURAL_PATTERNS:
#         if re.search(pattern, n_lower):
#             return True
#     return False

# def _categorize_component(name: str) -> str:
#     """Categorize component by name pattern."""
#     n_lower = name.lower()
    
#     for category, patterns in COMPONENT_PATTERNS.items():
#         for pattern in patterns:
#             if re.search(pattern, n_lower):
#                 return category
    
#     return "UNCATEGORIZED"

# def _normalize_group_name(name: str) -> str:
#     """Normalize raw block name."""
#     n = name.strip()
#     if "|" in n:
#         n = n.split("|")[-1]
    
#     # Keep underscores in FAL codes
#     if not n.startswith("FAL"):
#         n = re.sub(r"[_\-]+", " ", n)
    
#     n = re.sub(r"\s+", " ", n).strip()
    
#     # Don't remove numbers from version codes
#     if not re.search(r"V\d+$", n, re.IGNORECASE):
#         n = re.sub(r"\s*\(?\d+\)?$", "", n).strip()
    
#     return n.lower()

# def _detect_cbs_type(project_name: str) -> str:
#     """Detect CBS type from project name."""
#     pn_lower = project_name.lower()
#     # Check for linear CBS indicators
#     if "linear" in pn_lower or "linear cbs" in pn_lower or "linear sorter" in pn_lower:
#         return "Linear CBS"
#     # Check for loop CBS indicators (explicitly mentioned)
#     if "loop cbs" in pn_lower or "loop sorter" in pn_lower:
#         return "Loop CBS"
#     # Default to Loop CBS if not specified
#     return "Loop CBS"

# def _analyze_chute_types(components: dict) -> dict:
#     """Analyze chute breakdown by examining component names."""
#     chute_analysis = {
#         "total": 0,
#         "by_type": defaultdict(int),
#         "has_type_info": False
#     }
    
#     for comp_name, count in components.items():
#         name_lower = comp_name.lower()
#         if "chute" in name_lower:
#             chute_analysis["total"] += count
            
#             # Try to detect type
#             if "gravity" in name_lower or "collection" in name_lower:
#                 chute_analysis["by_type"]["gravity"] += count
#                 chute_analysis["has_type_info"] = True
#             elif "live" in name_lower or "active" in name_lower:
#                 chute_analysis["by_type"]["live"] += count
#                 chute_analysis["has_type_info"] = True
#             elif "slide" in name_lower or "sliding" in name_lower:
#                 chute_analysis["by_type"]["sliding"] += count
#                 chute_analysis["has_type_info"] = True
#             elif "reject" in name_lower or "exception" in name_lower:
#                 chute_analysis["by_type"]["rejection"] += count
#                 chute_analysis["has_type_info"] = True
#             elif "mini" in name_lower:
#                 chute_analysis["by_type"]["mini_gravity"] += count
#                 chute_analysis["has_type_info"] = True
#             elif "bulk" in name_lower:
#                 chute_analysis["by_type"]["bulk"] += count
#                 chute_analysis["has_type_info"] = True
#             elif "non-sort" in name_lower or "nonsort" in name_lower:
#                 chute_analysis["by_type"]["non_sort"] += count
#                 chute_analysis["has_type_info"] = True
    
#     return chute_analysis

# def extract_dxf_components(dxf_path: Path, project_name: str = "") -> dict:
#     """Extract and categorize components from DXF file."""
#     doc = ezdxf.readfile(str(dxf_path))
#     msp = doc.modelspace()
#     hdr = doc.header
    
#     # Get units
#     units_code = hdr.get("$INSUNITS", None)
#     try:
#         units_code = int(units_code) if units_code is not None else None
#     except:
#         units_code = None
    
#     # Get extents
#     extmin = hdr.get("$EXTMIN", None)
#     extmax = hdr.get("$EXTMAX", None)
    
#     # Extract block references
#     raw_counts: Counter[str] = Counter()
#     for e in msp:
#         try:
#             if e.dxftype() == "INSERT":
#                 bname = e.dxf.name
#                 if not _is_noise_block(bname) and not _is_structural(bname):
#                     raw_counts[bname] += 1
#         except:
#             continue
    
#     # Categorize components
#     categorized: dict = defaultdict(lambda: defaultdict(lambda: {"count": 0, "examples": []}))
    
#     for raw_name, cnt in raw_counts.items():
#         category = _categorize_component(raw_name)
#         gname = _normalize_group_name(raw_name)
        
#         categorized[category][gname]["count"] += cnt
#         categorized[category][gname]["examples"].append(raw_name)
    
#     # Detect system characteristics from project name
#     cbs_type = _detect_cbs_type(project_name if project_name else dxf_path.name)
#     chute_analysis = _analyze_chute_types(raw_counts)
    
#     # Determine induction type
#     has_auto_induct = len(categorized.get("AUTO_INDUCT", {})) > 0
#     has_operators = len(categorized.get("OPERATOR_STATION", {})) > 0
#     has_vds = len(categorized.get("VDS_BUFFER", {})) > 0
    
#     if has_auto_induct and has_operators:
#         induction_type = "MIXED (Auto + Manual)"
#     elif has_auto_induct:
#         induction_type = "AUTO"
#     elif has_operators:
#         induction_type = "MANUAL"
#     else:
#         induction_type = "UNKNOWN"
    
#     # Build category summary
#     category_summary = {}
#     total_components = 0
#     for cat, items in categorized.items():
#         count = sum(item["count"] for item in items.values())
#         category_summary[cat] = count
#         total_components += count
    
#     return {
#         "file": dxf_path.name,
#         "units_code": units_code,
#         "units_name": UNITS.get(units_code, "unknown") if units_code is not None else None,
#         "extents": {
#             "min": list(extmin) if extmin else None,
#             "max": list(extmax) if extmax else None
#         },
#         "cbs_type": cbs_type,
#         "induction_type": induction_type,
#         "has_vds": has_vds,
#         "total_components": total_components,
#         "category_summary": dict(category_summary),
#         "categorized_components": {
#             cat: {name: data["count"] for name, data in items.items()}
#             for cat, items in categorized.items()
#         },
#         "chute_analysis": chute_analysis,
#         "raw_block_counts": {k: int(v) for k, v in raw_counts.items()},
#     }

# def _summarise_components_for_prompt(dxf_json: dict) -> str:
#     """Generate comprehensive summary for LLM prompt."""
#     lines = ["=" * 70]
#     lines.append("DXF COMPONENT ANALYSIS FOR CBS PROCESS FLOW GENERATION")
#     lines.append("=" * 70)
#     lines.append("")
    
#     # File and system info
#     lines.append(f"FILE: {dxf_json.get('file', 'Unknown')}")
#     lines.append(f"UNITS: {dxf_json.get('units_name', 'Unknown')}")
#     lines.append(f"CBS TYPE: {dxf_json.get('cbs_type', 'Unknown')} (detected from filename)")
#     lines.append(f"TOTAL COMPONENTS: {dxf_json.get('total_components', 0)}")
#     lines.append("")
    
#     # System configuration
#     lines.append("SYSTEM CONFIGURATION:")
#     lines.append(f"  • Induction Type: {dxf_json.get('induction_type', 'Unknown')}")
#     lines.append(f"  • VDS/Buffer System: {'YES' if dxf_json.get('has_vds') else 'NO'}")
    
#     chute_analysis = dxf_json.get('chute_analysis', {})
#     if chute_analysis.get('total', 0) > 0:
#         lines.append(f"  • Total Chutes: {chute_analysis['total']}")
#         if chute_analysis.get('has_type_info'):
#             lines.append("  • Chute Types Detected: YES")
#         else:
#             lines.append("  • Chute Types Detected: NO (describe generically)")
#     lines.append("")
    
#     # Component categories
#     cat_summary = dxf_json.get("category_summary", {})
#     if cat_summary:
#         lines.append("COMPONENT CATEGORIES:")
#         priority = ["AUTO_INDUCT", "OPERATOR_STATION", "VDS_BUFFER", "CONVEYOR_INFEED",
#                    "CBS_SORTER", "CHUTE", "PTL", "BAG_SYSTEM", "RECIRCULATION", "SCANNER"]
        
#         for cat in priority:
#             if cat in cat_summary:
#                 lines.append(f"  • {cat}: {cat_summary[cat]} units")
        
#         # Add any remaining categories
#         for cat, count in cat_summary.items():
#             if cat not in priority:
#                 lines.append(f"  • {cat}: {count} units")
#         lines.append("")
    
#     # Chute breakdown
#     if chute_analysis.get('by_type'):
#         lines.append("CHUTE TYPE BREAKDOWN:")
#         for ctype, count in sorted(chute_analysis['by_type'].items(), key=lambda x: -x[1]):
#             lines.append(f"  • {ctype.replace('_', ' ').title()}: {count} chutes")
#         lines.append("")
    
#     # Detailed inventory by category
#     lines.append("DETAILED COMPONENT INVENTORY:")
#     lines.append("-" * 70)
    
#     categorized = dxf_json.get("categorized_components", {})
#     priority = ["AUTO_INDUCT", "OPERATOR_STATION", "VDS_BUFFER", "CONVEYOR_INFEED",
#                "CBS_SORTER", "CHUTE", "PTL", "BAG_SYSTEM", "RECIRCULATION", "SCANNER"]
    
#     for cat in priority:
#         if cat in categorized and categorized[cat]:
#             lines.append("")
#             lines.append(f"[{cat}]")
#             for name, count in sorted(categorized[cat].items(), key=lambda x: -x[1]):
#                 lines.append(f"  • {name}: {count} units")
    
#     # Add uncategorized if any
#     if "UNCATEGORIZED" in categorized and categorized["UNCATEGORIZED"]:
#         lines.append("")
#         lines.append("[UNCATEGORIZED]")
#         for name, count in sorted(categorized["UNCATEGORIZED"].items(), key=lambda x: -x[1]):
#             lines.append(f"  • {name}: {count} units")
    
#     lines.append("")
#     lines.append("=" * 70)
#     lines.append("PROCESS FLOW GENERATION GUIDANCE:")
#     lines.append(f"  • System Type: {dxf_json.get('cbs_type', 'Unknown')}")
#     lines.append(f"  • Induction: {dxf_json.get('induction_type', 'Unknown')}")
    
#     if dxf_json.get('has_vds'):
#         lines.append("  • Include VDS/Buffer section in Infeed System")
    
#     if dxf_json.get('induction_type') == "MIXED (Auto + Manual)":
#         lines.append("  • Include BOTH Auto Induct Line AND Manual Induct Station sections")
#     elif dxf_json.get('induction_type') == "AUTO":
#         lines.append("  • Include Auto Induct Line section only")
#     elif dxf_json.get('induction_type') == "MANUAL":
#         lines.append("  • Include Manual Induct Station section only")
    
#     if chute_analysis.get('has_type_info'):
#         lines.append("  • Chute types available - use detailed breakdown")
#     else:
#         lines.append("  • Chute types NOT available - describe generically")
    
#     if "PTL" in cat_summary:
#         lines.append("  • Include Put To Light System section")
    
#     if "BAG_SYSTEM" in cat_summary:
#         lines.append("  • Include Bag Takeaway Conveyor section")
    
#     if "RECIRCULATION" in cat_summary:
#         lines.append("  • Include Recirculation/Exception Refeeding section")
    
#     lines.append("=" * 70)
    
#     return "\n".join(lines)

# def read_process_flow_file(path: Path) -> str:
#     if path.suffix.lower() == ".txt":
#         return path.read_text(encoding="utf-8", errors="ignore")
#     if path.suffix.lower() == ".docx":
#         doc = Document(str(path))
#         return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
#     return path.read_text(encoding="utf-8", errors="ignore")

# # ----------------- Pinecone helpers (no gRPC) -----------------
# def get_pinecone_client_and_index(create_if_missing: bool = False):
#     """
#     Returns (pc, index_handle).
#     If create_if_missing is True, attempts to create the index with ServerlessSpec fallback.
#     """
#     if not PINECONE_API_KEY:
#         raise RuntimeError("PINECONE_API_KEY is not set in environment.")

#     pc = Pinecone(api_key=PINECONE_API_KEY)

#     # check existing indexes
#     try:
#         existing = {idx["name"] for idx in pc.list_indexes()}
#     except Exception as e:
#         raise RuntimeError(f"Failed to list Pinecone indexes: {e}")

#     if PINECONE_INDEX_NAME not in existing:
#         if not create_if_missing:
#             raise RuntimeError(f"Index '{PINECONE_INDEX_NAME}' does not exist in Pinecone. Set create_if_missing=True to auto-create.")
#         # try to create using ServerlessSpec if available, else simple create
#         try:
#             spec = ServerlessSpec(cloud="aws", region="us-east-1")
#             pc.create_index(name=PINECONE_INDEX_NAME, dimension=EMBED_DIM, metric="cosine", spec=spec)
#         except Exception:
#             pc.create_index(name=PINECONE_INDEX_NAME, dimension=EMBED_DIM, metric="cosine")
#         # slight pause so index becomes available
#         time.sleep(1.5)

#     index = pc.Index(PINECONE_INDEX_NAME)
#     return pc, index

# import inspect
# import numbers
# from typing import List

# import json
# from typing import List, Any

# def _extract_vector_from_item(item: Any):
#     """Return plain python list of floats from possible response item shapes."""
#     if isinstance(item, dict):
#         for k in ("values", "embedding", "vector", "embeddings"):
#             if k in item and item[k] is not None:
#                 v = item[k]
#                 if hasattr(v, "tolist"):
#                     try:
#                         return v.tolist()
#                     except Exception:
#                         pass
#                 if isinstance(v, (list, tuple)):
#                     return list(v)
#                 if isinstance(v, (float, int)):
#                     return [float(v)]
#     # object with attributes
#     for attr in ("values", "embedding", "vector", "embeddings"):
#         if hasattr(item, attr):
#             v = getattr(item, attr)
#             if hasattr(v, "tolist"):
#                 try:
#                     return v.tolist()
#                 except Exception:
#                     pass
#             if isinstance(v, (list, tuple)):
#                 return list(v)
#             if isinstance(v, (float, int)):
#                 return [float(v)]
#     # item itself could be a list/tuple
#     if isinstance(item, (list, tuple)):
#         return list(item)
#     return None


# def pinecone_embed(pc, texts: List[str], model: str = EMBED_MODEL, input_type: str = "passage", truncate: str = "END") -> List[List[float]]:
#     """
#     Embed `texts` using Pinecone inference for models that require `parameters.input_type`.
#     - pc: Pinecone client (Pinecone(...))
#     - texts: list of strings
#     - model: model name (e.g. "llama-text-embed-v2")
#     - input_type: 'passage' or 'query' (model-specific)
#     - truncate: optional truncate policy (e.g. "END")
#     Returns: list-of-vectors (one per text).
#     """

#     if not texts:
#         return []

#     # Build inputs according to Pinecone inference / embed API
#     inputs_payload = [{"text": t} for t in texts]

#     # Default parameters for llama-text-embed-v2: requires input_type
#     parameters = {"input_type": input_type}
#     if truncate:
#         parameters["truncate"] = truncate

#     # Helpful check: fetch model details (if supported) to see default params
#     try:
#         model_info = pc.inference.get_model(model_name=model)
#         # model_info may contain guidance; we don't absolutely need it here but log
#         # (you can print/log model_info if troubleshooting)
#     except Exception:
#         model_info = None

#     # Make the embed call with canonical shape:
#     # pc.inference.embed(model=..., inputs=[{"text":...}], parameters={...})
#     try:
#         resp = pc.inference.embed(model=model, inputs=inputs_payload, parameters=parameters)
#     except Exception as e:
#         # surface clear error with context about the parameter requirement
#         raise RuntimeError(
#             "Pinecone inference.embed failed. This model requires `parameters.input_type` and `inputs` "
#             "as a list of {\"text\": ...} objects. "
#             f"Attempted model='{model}', input_type='{input_type}'.\n"
#             f"Underlying error: {repr(e)}"
#         )

#     # Normalize response => List[List[float]]
#     vectors = []

#     # resp.data list-of-items (common)
#     if hasattr(resp, "data") and isinstance(resp.data, list):
#         for item in resp.data:
#             v = _extract_vector_from_item(item)
#             if v is None:
#                 raise RuntimeError("Unable to extract vector from the inference response item.")
#             vectors.append([float(x) for x in v])
#         return vectors

#     # dict-like with 'data'
#     if isinstance(resp, dict) and "data" in resp and isinstance(resp["data"], list):
#         for item in resp["data"]:
#             v = _extract_vector_from_item(item)
#             if v is None:
#                 raise RuntimeError("Unable to extract vector from resp['data'] item.")
#             vectors.append([float(x) for x in v])
#         return vectors

#     # plain list-of-vectors
#     if isinstance(resp, list):
#         if not resp:
#             return []
#         if isinstance(resp[0], (list, tuple)):
#             for it in resp:
#                 vectors.append([float(x) for x in it])
#             return vectors
#         else:
#             # single flat vector
#             return [[float(x) for x in resp]]

#     # resp.embeddings style
#     for attr in ("embeddings", "vectors", "values", "embedding"):
#         if hasattr(resp, attr):
#             cand = getattr(resp, attr)
#             if isinstance(cand, list):
#                 for item in cand:
#                     v = _extract_vector_from_item(item)
#                     if v is None:
#                         raise RuntimeError(f"Unable to extract vector from resp.{attr} item.")
#                     vectors.append([float(x) for x in v])
#                 return vectors

#     # Unknown shape — include short repr for debugging
#     raise RuntimeError("Unrecognized embedding response shape from Pinecone inference. Response repr: " + repr(resp)[:1000])




# # ----------------- Streamlit UI -----------------
# st.set_page_config(page_title="DXF → Pinecone (no gRPC)", layout="centered")
# st.title("Ingest DXF + Process Flow → Pinecone (Pinecone-only embedding & upsert)")

# st.markdown(
#     """
# Upload a DXF and its matching Process Flow (.docx or .txt).  
# This app will:
# 1. Extract DXF components and create a short summary.
# 2. Create an embedding using Pinecone inference (`llama-text-embed-v2`).
# 3. Upsert the vector into your Pinecone index using the management client (no gRPC).
# """
# )

# with st.form("form"):
#     c1, c2 = st.columns(2)
#     with c1:
#         dxf_file = st.file_uploader("DXF (.dxf)", type=["dxf"])
#         project_name = st.text_input("Project name (optional)")
#     with c2:
#         pf_file = st.file_uploader("Process Flow (.docx/.txt)", type=["docx","txt"])
#         client_name = st.text_input("Client name (optional)")
#     create_index = st.checkbox("Create index if missing", value=False, help="Automatically create index if it does not exist (requires Pinecone account privileges).")
#     submitted = st.form_submit_button("Ingest into Pinecone")

# if submitted:
#     if not dxf_file or not pf_file:
#         st.error("Please upload both DXF and Process Flow files.")
#         st.stop()

#     with tempfile.TemporaryDirectory() as td:
#         td = Path(td)
#         dxf_path = td / dxf_file.name
#         dxf_path.write_bytes(dxf_file.read())
#         pf_path = td / pf_file.name
#         pf_path.write_bytes(pf_file.read())

#         # extract
#         try:
#             dxf_json = extract_dxf_components(dxf_path, project_name=project_name or dxf_file.name)
#             dxf_summary = _summarise_components_for_prompt(dxf_json)
#         except Exception as e:
#             st.exception(f"DXF parse failed: {e}")
#             st.stop()

#         pf_text = read_process_flow_file(pf_path)

#         # Pinecone client + index
#         try:
#             pc, index = get_pinecone_client_and_index(create_if_missing=create_index)
#         except Exception as e:
#             st.exception(f"Pinecone client/index init failed: {e}")
#             st.stop()

#         # embedding
#         st.info("Embedding DXF summary via Pinecone inference...")
#         try:
#             vectors = pinecone_embed(pc, [dxf_summary])
#             vec = vectors[0]
#         except Exception as e:
#             st.exception(f"Embedding failed: {e}")
#             st.stop()

#         # prepare metadata + id
#         safe_proj = (project_name or dxf_file.name).strip().replace(" ", "_")
#         client_tag = (client_name or "unknown_client").strip().replace(" ", "_")
#         vec_id = f"{client_tag}::{safe_proj}::{dxf_file.name}".lower()

#         metadata = {
#             "dxf_file": dxf_file.name,
#             "dxf_summary": dxf_summary,
#             "process_flow": pf_text,
#             "cbs_type": dxf_json.get("cbs_type"),
#             "induction_type": dxf_json.get("induction_type"),
#             "has_vds": dxf_json.get("has_vds"),
#             "total_components": dxf_json.get("total_components"),
#             "created_at": datetime.utcnow().isoformat() + "Z",
#         }
        
#         # Only add optional fields if they have values (Pinecone doesn't accept null)
#         if client_name:
#             metadata["client"] = client_name
#         if project_name:
#             metadata["project_name"] = project_name

#         # metadata size guard (Pinecone has limits)
#         md_size = len(json.dumps(metadata).encode("utf-8"))
#         if md_size > 32_000:  # tune based on your account limits
#             st.warning("Metadata is large ({} bytes). Consider storing full process_flow externally and keeping only a short summary in metadata.".format(md_size))
#             # If too large, store just the summary and essential data
#             metadata_compact = {
#                 "dxf_file": dxf_file.name,
#                 "dxf_summary_compact": f"FILE: {dxf_json['file']} | COMPONENTS: {dxf_json['total_components']} | CBS: {dxf_json['cbs_type']} | INDUCTION: {dxf_json['induction_type']}",
#                 "process_flow": pf_text,
#                 "cbs_type": dxf_json.get("cbs_type"),
#                 "induction_type": dxf_json.get("induction_type"),
#                 "has_vds": dxf_json.get("has_vds"),
#                 "total_components": dxf_json.get("total_components"),
#                 "created_at": datetime.utcnow().isoformat() + "Z",
#             }
#             if client_name:
#                 metadata_compact["client"] = client_name
#             if project_name:
#                 metadata_compact["project_name"] = project_name
#             metadata = metadata_compact

#         # upsert
#         try:
#             index.upsert(vectors=[{"id": vec_id, "values": list(vec), "metadata": metadata}], namespace=EMBED_NAMESPACE)
#         except Exception as e:
#             st.exception(f"Pinecone upsert failed: {e}")
#             st.stop()

#         st.success(f"Upsert successful — id={vec_id} into index `{PINECONE_INDEX_NAME}` namespace `{EMBED_NAMESPACE}`")
#         st.json({
#             "id": vec_id,
#             "namespace": EMBED_NAMESPACE,
#             "metadata_preview": {
#                 "dxf_file": metadata["dxf_file"],
#                 "cbs_type": metadata["cbs_type"],
#                 "induction_type": metadata["induction_type"],
#                 "has_vds": metadata["has_vds"],
#                 "total_components": metadata["total_components"],
#                 "created_at": metadata["created_at"]
#             }
#         })

#         st.download_button("Download DXF summary", data=dxf_summary, file_name=f"{safe_proj}_dxf_summary.txt")
#         st.download_button("Download process flow", data=pf_text, file_name=f"{safe_proj}_process_flow.txt")











"""
push.py - Ingest DXF + Process Flow to Pinecone
================================================
CRITICAL FIX: Now uses dxf_extractor.py for consistent extraction
"""

import os
import re
import tempfile
import json
import time
from pathlib import Path
from datetime import datetime
import logging

import streamlit as st
from dotenv import load_dotenv
from docx import Document
from pinecone import Pinecone, ServerlessSpec

# CRITICAL: Import unified extractor
from dxf_extractor import (
    extract_dxf_components,
    create_dxf_summary_for_embedding,  # Use SAME format as combine.py
)

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("dxf-ingest")

# ----------------- CONFIG -----------------
PINECONE_API_KEY = os.getenv("PINECONE_API_KEY", "pcsk_8akoe_FxzXaW2zvAsEd1uiHqxiMrosvumSujgFyrWAB9vyqG87DGWpnDc6rSxaDYrkP3v")
PINECONE_INDEX_NAME = os.getenv("PINECONE_INDEX_NAME", "spas-dxf-samples")
EMBED_MODEL = "llama-text-embed-v2"
EMBED_DIM = 1048
EMBED_NAMESPACE = "v1-dxf"


def read_process_flow_file(path: Path) -> str:
    """Read process flow from .txt or .docx"""
    if path.suffix.lower() == ".txt":
        return path.read_text(encoding="utf-8", errors="ignore")
    if path.suffix.lower() == ".docx":
        doc = Document(str(path))
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    return path.read_text(encoding="utf-8", errors="ignore")


def get_pinecone_client_and_index(create_if_missing: bool = False):
    """Get Pinecone client and index"""
    if not PINECONE_API_KEY:
        raise RuntimeError("PINECONE_API_KEY is not set")

    pc = Pinecone(api_key=PINECONE_API_KEY)

    # Check existing indexes
    try:
        existing = {idx["name"] for idx in pc.list_indexes()}
    except Exception as e:
        raise RuntimeError(f"Failed to list Pinecone indexes: {e}")

    if PINECONE_INDEX_NAME not in existing:
        if not create_if_missing:
            raise RuntimeError(f"Index '{PINECONE_INDEX_NAME}' does not exist")
        
        # Create index
        try:
            spec = ServerlessSpec(cloud="aws", region="us-east-1")
            pc.create_index(
                name=PINECONE_INDEX_NAME,
                dimension=EMBED_DIM,
                metric="cosine",
                spec=spec
            )
            time.sleep(2)  # Wait for index to be ready
        except Exception as e:
            raise RuntimeError(f"Failed to create index: {e}")

    index = pc.Index(PINECONE_INDEX_NAME)
    return pc, index


def pinecone_embed(pc, texts: list, model: str = EMBED_MODEL) -> list:
    """
    Embed texts using Pinecone inference.
    Returns list of vectors (list of floats).
    """
    if not texts:
        return []

    inputs_payload = [{"text": t} for t in texts]
    parameters = {"input_type": "passage", "truncate": "END"}

    try:
        resp = pc.inference.embed(
            model=model,
            inputs=inputs_payload,
            parameters=parameters
        )
    except Exception as e:
        raise RuntimeError(f"Pinecone embed failed: {e}")

    # Extract vectors
    vectors = []
    if hasattr(resp, "data") and isinstance(resp.data, list):
        for item in resp.data:
            if hasattr(item, "values"):
                vectors.append([float(x) for x in item.values])
            elif isinstance(item, dict) and "values" in item:
                vectors.append([float(x) for x in item["values"]])
            else:
                raise RuntimeError("Unexpected embedding response format")
    else:
        raise RuntimeError("Unexpected embedding response structure")

    return vectors


# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="DXF → Pinecone Ingestion", layout="centered")
st.title("🔄 Ingest DXF + Process Flow → Pinecone")

st.markdown("""
Upload a DXF and its matching Process Flow (.docx or .txt).

**This will:**
1. Extract DXF components using **unified extractor** (same as combine.py)
2. Create embedding using Pinecone inference
3. Upsert to Pinecone with metadata
""")

with st.form("form"):
    c1, c2 = st.columns(2)
    with c1:
        dxf_file = st.file_uploader("DXF (.dxf)", type=["dxf"])
        project_name = st.text_input("Project name (optional)")
    with c2:
        pf_file = st.file_uploader("Process Flow (.docx/.txt)", type=["docx", "txt"])
        client_name = st.text_input("Client name (required)", value="")
    
    create_index = st.checkbox(
        "Create index if missing",
        value=False,
        help="Auto-create index if it doesn't exist"
    )
    submitted = st.form_submit_button("🚀 Ingest into Pinecone")

if submitted:
    if not dxf_file or not pf_file:
        st.error("❌ Please upload both DXF and Process Flow files")
        st.stop()
    
    if not client_name:
        st.error("❌ Client name is required for proper organization")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        dxf_path = td / dxf_file.name
        dxf_path.write_bytes(dxf_file.read())
        pf_path = td / pf_file.name
        pf_path.write_bytes(pf_file.read())

        # CRITICAL: Use unified extractor (same as combine.py)
        st.info("📊 Extracting DXF components...")
        try:
            dxf_json = extract_dxf_components(
                dxf_path,
                project_name=project_name or dxf_file.name
            )
            # CRITICAL: Use SAME summary format as combine.py for embedding
            dxf_summary = create_dxf_summary_for_embedding(dxf_json)
        except Exception as e:
            st.exception(f"❌ DXF extraction failed: {e}")
            st.stop()

        # Read process flow
        st.info("📄 Reading process flow...")
        try:
            pf_text = read_process_flow_file(pf_path)
        except Exception as e:
            st.exception(f"❌ Failed to read process flow: {e}")
            st.stop()

        # Pinecone setup
        st.info("🔌 Connecting to Pinecone...")
        try:
            pc, index = get_pinecone_client_and_index(create_if_missing=create_index)
        except Exception as e:
            st.exception(f"❌ Pinecone connection failed: {e}")
            st.stop()

        # Embed DXF summary
        st.info("🧠 Generating embedding...")
        try:
            vectors = pinecone_embed(pc, [dxf_summary])
            vec = vectors[0]
        except Exception as e:
            st.exception(f"❌ Embedding failed: {e}")
            st.stop()

        # Prepare metadata and ID
        safe_proj = (project_name or dxf_file.name).strip().replace(" ", "_")
        client_tag = client_name.strip().replace(" ", "_")
        vec_id = f"{client_tag}::{safe_proj}::{dxf_file.name}".lower()

        # Build metadata (Pinecone only accepts flat values)
        # Flatten category/chute counts so Pinecone metadata stays scalar-only
        cats = dxf_json.get("category_summary", {}) or {}
        chute = dxf_json.get("chute_analysis", {}) or {}

        metadata = {
            "dxf_file": dxf_file.name,
            "client": client_name,
            "process_flow": pf_text,
            "cbs_type": dxf_json.get("cbs_type"),
            "induction_type": dxf_json.get("induction_type"),
            "has_vds": dxf_json.get("has_vds"),
            "has_recirculation": dxf_json.get("has_recirculation"),
            "has_scanner": dxf_json.get("has_scanner"),
            "total_components": dxf_json.get("total_components"),
            # Serialized summaries for back-compat parsing
            "category_summary_json": json.dumps(cats),
            "chute_analysis_json": json.dumps(chute),
            # Flattened component counts (numeric fields allowed by Pinecone)
            "cat_AUTO_INDUCT": cats.get("AUTO_INDUCT", 0),
            "cat_OPERATOR_STATION": cats.get("OPERATOR_STATION", 0),
            "cat_CONVEYOR_INFEED": cats.get("CONVEYOR_INFEED", 0),
            "cat_VDS_BUFFER": cats.get("VDS_BUFFER", 0),
            "cat_CHUTE": cats.get("CHUTE", 0),
            "cat_RECIRCULATION": cats.get("RECIRCULATION", 0),
            "cat_PTL": cats.get("PTL", 0),
            "cat_BAG_SYSTEM": cats.get("BAG_SYSTEM", 0),
            "cat_SCANNER": cats.get("SCANNER", 0),
            "cat_CBS_SORTER": cats.get("CBS_SORTER", 0),
            # Flattened chute breakdown
            "chute_total": chute.get("total", 0),
            "chute_live": chute.get("by_type", {}).get("live", 0),
            "chute_collection": chute.get("by_type", {}).get("collection", 0),
            "chute_rejection": chute.get("by_type", {}).get("rejection", 0),
            "chute_sliding": chute.get("by_type", {}).get("sliding", 0),
            "chute_mini_gravity": chute.get("by_type", {}).get("mini_gravity", 0),
            "chute_bulk": chute.get("by_type", {}).get("bulk", 0),
            "chute_big_parcel": chute.get("by_type", {}).get("big_parcel", 0),
            "chute_gravity": chute.get("by_type", {}).get("gravity", 0),
            "created_at": datetime.utcnow().isoformat() + "Z",
        }
        
        if project_name:
            metadata["project_name"] = project_name

        # Check metadata size
        md_size = len(json.dumps(metadata).encode("utf-8"))
        if md_size > 35_000:  # Pinecone limit is ~40KB
            st.warning(f"⚠️ Metadata is large ({md_size} bytes). Compacting...")
            # Store compact process flow
            metadata["process_flow"] = pf_text[:15000] + "\n...[truncated]"

        # Upsert to Pinecone
        st.info(f"☁️ Upserting to Pinecone (ID: {vec_id})...")
        try:
            index.upsert(
                vectors=[{
                    "id": vec_id,
                    "values": list(vec),
                    "metadata": metadata
                }],
                namespace=EMBED_NAMESPACE
            )
        except Exception as e:
            st.exception(f"❌ Pinecone upsert failed: {e}")
            st.stop()

        # Success!
        st.success(f"✅ Successfully ingested to Pinecone!")
        
        st.json({
            "id": vec_id,
            "namespace": EMBED_NAMESPACE,
            "index": PINECONE_INDEX_NAME,
            "metadata_preview": {
                "client": client_name,
                "dxf_file": dxf_file.name,
                "cbs_type": dxf_json["cbs_type"],
                "induction_type": dxf_json["induction_type"],
                "total_components": dxf_json["total_components"],
                "has_vds": dxf_json["has_vds"],
            }
        })

        # Display summary
        with st.expander("📊 DXF Summary (embedded)"):
            st.code(dxf_summary, language="text")
        
        with st.expander("📄 Process Flow"):
            st.text(pf_text[:2000] + ("..." if len(pf_text) > 2000 else ""))

        # Download buttons
        st.download_button(
            "⬇️ Download DXF Summary",
            data=dxf_summary,
            file_name=f"{safe_proj}_dxf_summary.txt"
        )
        st.download_button(
            "⬇️ Download Process Flow",
            data=pf_text,
            file_name=f"{safe_proj}_process_flow.txt"
        )
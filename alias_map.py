import json
from collections import defaultdict
from pathlib import Path
from typing import Dict, Any, Tuple

ALIAS_MAP_PATH = Path("component_alias_map.json")


def _load_alias_map() -> Dict[str, Dict[str, Any]]:
    if ALIAS_MAP_PATH.exists():
        try:
            return json.loads(ALIAS_MAP_PATH.read_text())
        except Exception:
            return {}
    return {}


def _save_alias_map(data: Dict[str, Dict[str, Any]]):
    try:
        ALIAS_MAP_PATH.write_text(json.dumps(data, indent=2))
    except Exception:
        pass


def analyze_and_update_aliases(dxf_json: Dict[str, Any]) -> Dict[str, int]:
    """
    Lightweight alias mapper:
    - Groups unknown blocks by (block_name|layer)
    - Uses nearby text evidence if present (heuristic from dxf_json)
    - Updates alias map JSON (no LLM here to keep it safe)
    - Returns alias_counts for common labels (e.g., Infeed Conveyor, Gravity Chute)

    Expected dxf_json fields (best-effort):
    - unknown_blocks: list of {name, layer, evidence: [text,..]}
    """
    alias_map = _load_alias_map()
    counts = defaultdict(int)

    for blk in dxf_json.get("unknown_blocks", []):
        name = str(blk.get("name", "")).strip()
        layer = str(blk.get("layer", "")).strip()
        evidence = " ".join(blk.get("evidence", [])).lower()
        key = f"{name}|{layer}"
        # Simple heuristics: map by evidence keywords
        if "infeed" in evidence or "in feed" in evidence:
            label = "Infeed Conveyor"
        elif "chute" in evidence or "gravity" in evidence:
            label = "Gravity Chute"
        elif "scanner" in evidence or "barcode" in evidence:
            label = "Barcode Scanner"
        else:
            label = alias_map.get(key, {}).get("label", "unknown")
        # Update alias map
        alias_map[key] = {
            "label": label,
            "confidence": 0.65 if label != "unknown" else 0.0,
            "evidence": blk.get("evidence", []),
        }
        # Aggregate counts for labels we care about
        if label == "Gravity Chute":
            counts["gravity_chutes"] += 1
        if label == "Infeed Conveyor":
            counts["infeed_conveyors_total"] += 1
        if label == "Barcode Scanner":
            counts["barcode_scanners"] += 1

    _save_alias_map(alias_map)
    return dict(counts)

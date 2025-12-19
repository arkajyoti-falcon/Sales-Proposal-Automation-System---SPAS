"""
dxf_extractor.py - Unified DXF Component Extraction with Enhanced Embedding
============================================================================
CRITICAL IMPROVEMENTS:
1. Component-first embedding format for better discrimination
2. Includes RAW block names for transparency
3. Enhanced pattern matching (fewer UNCATEGORIZED)
4. Hybrid retrieval support with component similarity scoring
"""

import re
from pathlib import Path
from collections import Counter, defaultdict
from typing import Dict, List
import ezdxf

# ==================== SHARED CONFIGURATION ====================

UNITS = {
    0: "Unitless", 1: "inches", 2: "feet", 3: "miles",
    4: "millimeters", 5: "centimeters", 6: "meters", 7: "kilometers"
}

# ENHANCED: More comprehensive patterns to reduce UNCATEGORIZED
COMPONENT_PATTERNS = {
    "AUTO_INDUCT": [
        r"fal.*fs\d+", r"fal.*feed", r"feedline", r"feed.*line",
        r"auto.*induct", r"induct.*auto", r"fs\d{3}",
        r"transfer.*plate", r"induct.*conv"
    ],
    "CONVEYOR_INFEED": [
        r"telescopic", r"infeed.*conv", r"in.*feed", r"infeed",
        r"inclined.*conv", r"incline", r"elevation.*conv",
        r"receiving.*conv", r"highway", r"singulat"
    ],
    "VDS_BUFFER": [
        r"vds", r"distribution.*loop", r"distribution", r"buffer",
        r"arm.*vds", r"fal.*s013"
    ],
    "OPERATOR_STATION": [
        r"operator(?!.*safety)", r"manual.*station", r"induct.*station",
        r"manual.*load"
    ],
    "CHUTE": [
        r"chute", r"slide", r"sliding", r"gravity",
        r"live(?!.*load)", r"reject", r"collection.*chute",
        r"sortfail", r"exception.*chute", r"discharge",
        r"output.*chute", r"parcel.*chute"
    ],
    "PTL": [
        r"ptl", r"put.*to.*light", r"pick.*to.*light",
        r"light.*rack", r"pallet.*setup"
    ],
    "BAG_SYSTEM": [
        r"bag", r"bagging", r"takeaway", r"trolley",
        r"roller.*cage"
    ],
    "RECIRCULATION": [
        r"recirculation", r"recirculate", r"refeed",
        r"return.*conv", r"return.*line", r"loop.*back"
    ],
    "CBS_SORTER": [
        r"cbs", r"cross.*belt", r"crossbelt", r"sorter",
        r"carrier", r"loop.*sorter", r"linear.*sorter"
    ],
    "SCANNER": [
        r"scanner", r"scan", r"tunnel", r"barcode",
        r"dimension", r"dws", r"volume.*scan", r"reader"
    ],
    "CONVEYOR_TAKEAWAY": [
        r"takeaway.*conv", r"discharge.*conv", r"outfeed",
        r"pvc.*belt", r"tbc", r"take.*away"
    ],
    "COLLECTION": [
        r"collection.*bin", r"pallet(?!.*setup)",
        r"friction.*roller", r"accumulation"
    ]
}

STRUCTURAL_PATTERNS = [
    r"leg.*guard", r"guard(?!.*operator)", r"fenc", r"railing",
    r"handrail", r"safety.*(?!operator)", r"pallet(?!.*conv)",
    r"step", r"stair", r"ladder", r"door", r"panel(?!.*control)",
    r"bracket", r"bolt", r"mount(?!.*scanner)",
    r"crossover.*main", r"end.*joint", r"platform"
]

# ==================== HELPER FUNCTIONS ====================

def _is_noise_block(name: str) -> bool:
    """Filter out anonymous noise blocks."""
    n = name.strip()
    if re.match(r"^\*[UDXATE]\d+$", n, re.IGNORECASE):
        return True
    if n.startswith("*") or n.startswith("~") or n.startswith("A$C"):
        return True
    return False


def _is_structural(name: str) -> bool:
    """Check if component is structural (non-flow)."""
    n_lower = name.lower()
    return any(re.search(p, n_lower) for p in STRUCTURAL_PATTERNS)


def _categorize_component(name: str) -> str:
    """Categorize component by name pattern."""
    n_lower = name.lower()
    
    for category, patterns in COMPONENT_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, n_lower):
                return category
    
    return "UNCATEGORIZED"


def _normalize_group_name(name: str) -> str:
    """Normalize raw block name with consistent rules."""
    n = name.strip()
    
    if "|" in n:
        n = n.split("|")[-1]
    
    if not n.startswith("FAL"):
        n = re.sub(r"[_\-]+", " ", n)
    
    n = re.sub(r"\s+", " ", n).strip()
    
    if not re.search(r"V\d+$", n, re.IGNORECASE):
        n = re.sub(r"\s*\(?\d+\)?$", "", n).strip()
    
    return n.lower()


def _detect_cbs_type(project_name: str) -> str:
    """Detect CBS type from project name."""
    pn_lower = project_name.lower()
    if "linear" in pn_lower:
        return "Linear CBS"
    return "Loop CBS"


def _analyze_chute_types(components: dict) -> dict:
    """Analyze chute breakdown by examining component names."""
    chute_analysis = {
        "total": 0,
        "by_type": defaultdict(int),
        "has_type_info": False
    }
    
    for comp_name, count in components.items():
        name_lower = comp_name.lower()
        if "chute" in name_lower or "slide" in name_lower or "discharge" in name_lower:
            chute_analysis["total"] += count
            
            # Detect types (order matters - most specific first)
            if "reject" in name_lower or "sortfail" in name_lower or "exception" in name_lower:
                chute_analysis["by_type"]["rejection"] += count
                chute_analysis["has_type_info"] = True
            elif "collection" in name_lower or "friction" in name_lower or "accumulation" in name_lower:
                chute_analysis["by_type"]["collection"] += count
                chute_analysis["has_type_info"] = True
            elif "live" in name_lower or "active" in name_lower:
                chute_analysis["by_type"]["live"] += count
                chute_analysis["has_type_info"] = True
            elif "sliding" in name_lower or "slide" in name_lower:
                chute_analysis["by_type"]["sliding"] += count
                chute_analysis["has_type_info"] = True
            elif "mini" in name_lower:
                chute_analysis["by_type"]["mini_gravity"] += count
                chute_analysis["has_type_info"] = True
            elif "bulk" in name_lower:
                chute_analysis["by_type"]["bulk"] += count
                chute_analysis["has_type_info"] = True
            elif "big parcel" in name_lower or "parcel" in name_lower:
                chute_analysis["by_type"]["big_parcel"] += count
                chute_analysis["has_type_info"] = True
            elif "gravity" in name_lower:
                chute_analysis["by_type"]["gravity"] += count
                chute_analysis["has_type_info"] = True
    
    return chute_analysis


# ==================== MAIN EXTRACTION FUNCTION ====================

def extract_dxf_components(dxf_path: Path, project_name: str = "") -> dict:
    """
    Extract and categorize components from DXF file.
    MUST produce identical output whether called from push.py or combine.py.
    """
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    hdr = doc.header
    
    # Get units
    units_code = hdr.get("$INSUNITS", None)
    try:
        units_code = int(units_code) if units_code is not None else None
    except:
        units_code = None
    
    # Extract block references
    raw_counts: Counter[str] = Counter()
    for e in msp:
        try:
            if e.dxftype() == "INSERT":
                bname = e.dxf.name
                if not _is_noise_block(bname) and not _is_structural(bname):
                    raw_counts[bname] += 1
        except:
            continue
    
    # Categorize components
    categorized: dict = defaultdict(lambda: defaultdict(lambda: {"count": 0, "examples": []}))
    
    for raw_name, cnt in raw_counts.items():
        category = _categorize_component(raw_name)
        gname = _normalize_group_name(raw_name)
        
        categorized[category][gname]["count"] += cnt
        categorized[category][gname]["examples"].append(raw_name)
    
    # Detect system characteristics
    cbs_type = _detect_cbs_type(project_name if project_name else dxf_path.name)
    chute_analysis = _analyze_chute_types(raw_counts)
    
    # Determine features
    has_auto_induct = len(categorized.get("AUTO_INDUCT", {})) > 0
    has_operators = len(categorized.get("OPERATOR_STATION", {})) > 0
    has_vds = len(categorized.get("VDS_BUFFER", {})) > 0
    has_recirculation = len(categorized.get("RECIRCULATION", {})) > 0
    has_scanner = len(categorized.get("SCANNER", {})) > 0
    
    if has_auto_induct and has_operators:
        induction_type = "MIXED (Auto + Manual)"
    elif has_auto_induct:
        induction_type = "AUTO"
    elif has_operators:
        induction_type = "MANUAL"
    else:
        induction_type = "UNKNOWN"
    
    # Build category summary
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
        "cbs_type": cbs_type,
        "induction_type": induction_type,
        "has_vds": has_vds,
        "has_recirculation": has_recirculation,
        "has_scanner": has_scanner,
        "total_components": total_components,
        "category_summary": dict(category_summary),
        "categorized_components": {
            cat: {name: data["count"] for name, data in items.items()}
            for cat, items in categorized.items()
        },
        "chute_analysis": chute_analysis,
        "raw_block_counts": {k: int(v) for k, v in raw_counts.items()},
        "raw_components_by_category": {
            cat: {name: data["examples"] for name, data in items.items()}
            for cat, items in categorized.items()
        }
    }


# ==================== SUMMARY GENERATION ====================

def create_dxf_summary_for_embedding(dxf_json: dict) -> str:
    """
    COMPONENT-FIRST embedding format for better discrimination.
    
    CRITICAL: This format is used by BOTH:
    - push.py (during ingestion)
    - combine.py (during querying)
    
    Must be IDENTICAL in both places for high similarity scores.
    
    KEY CHANGE: Put unique component data FIRST, metadata LAST.
    This makes the embedding focus on what matters.
    """
    lines = []
    
    # SECTION 1: Component Fingerprint (MOST IMPORTANT - Goes First!)
    lines.append("SYSTEM FINGERPRINT:")
    
    cats = dxf_json.get("category_summary", {})
    
    # Build component signature string
    signature_parts = []
    priority = ["AUTO_INDUCT", "OPERATOR_STATION", "CONVEYOR_INFEED", 
                "VDS_BUFFER", "CHUTE", "RECIRCULATION", "PTL", 
                "BAG_SYSTEM", "SCANNER", "CBS_SORTER"]
    
    for cat in priority:
        if cat in cats and cats[cat] > 0:
            signature_parts.append(f"{cat}={cats[cat]}")
    
    lines.append(" | ".join(signature_parts))
    lines.append("")
    
    # SECTION 2: Chute Breakdown (CRITICAL for differentiation)
    chute = dxf_json.get("chute_analysis", {})
    if chute.get("total", 0) > 0:
        lines.append(f"CHUTES BREAKDOWN: TOTAL={chute['total']}")
        if chute.get("by_type"):
            chute_parts = []
            for ct, cnt in sorted(chute["by_type"].items(), key=lambda x: -x[1]):
                chute_parts.append(f"{ct}={cnt}")
            lines.append(" | ".join(chute_parts))
        lines.append("")
    
    # SECTION 3: System Type (Differentiator)
    lines.append(f"CBS: {dxf_json['cbs_type']}")
    lines.append(f"INDUCTION: {dxf_json['induction_type']}")
    lines.append(f"VDS: {'YES' if dxf_json['has_vds'] else 'NO'}")
    lines.append(f"RECIRCULATION: {'YES' if dxf_json.get('has_recirculation') else 'NO'}")
    lines.append("")
    
    # SECTION 4: Raw Component Examples (For exact matching)
    lines.append("KEY COMPONENTS:")
    raw_by_cat = dxf_json.get("raw_components_by_category", {})
    
    for cat in ["AUTO_INDUCT", "CHUTE", "RECIRCULATION", "SCANNER"]:
        if cat in raw_by_cat and raw_by_cat[cat]:
            for gname, examples in list(raw_by_cat[cat].items())[:2]:
                example = examples[0] if examples else gname
                lines.append(f"{cat}: {example}")
    
    lines.append("")
    
    # SECTION 5: Metadata (LEAST IMPORTANT - Goes Last)
    lines.append(f"FILE: {dxf_json['file']}")
    lines.append(f"TOTAL_COMPONENTS: {dxf_json['total_components']}")
    
    return "\n".join(lines)


def create_dxf_summary_verbose(dxf_json: dict) -> str:
    """
    VERBOSE summary for LLM prompts (generation phase only).
    Includes ALL raw block names and detailed guidance.
    """
    lines = ["=" * 70]
    lines.append("COMPLETE DXF ANALYSIS FOR PROCESS FLOW GENERATION")
    lines.append("=" * 70)
    lines.append("")
    
    lines.append(f"FILE: {dxf_json.get('file', 'Unknown')}")
    lines.append(f"UNITS: {dxf_json.get('units_name', 'Unknown')}")
    lines.append(f"CBS TYPE: {dxf_json.get('cbs_type', 'Unknown')}")
    lines.append(f"TOTAL COMPONENTS: {dxf_json.get('total_components', 0)}")
    lines.append("")
    
    # System configuration
    lines.append("SYSTEM CONFIGURATION:")
    lines.append(f"  • Induction Type: {dxf_json.get('induction_type', 'Unknown')}")
    lines.append(f"  • VDS/Buffer: {'YES' if dxf_json.get('has_vds') else 'NO'}")
    lines.append(f"  • Recirculation: {'YES' if dxf_json.get('has_recirculation') else 'NO'}")
    lines.append(f"  • Scanner/DWS: {'YES' if dxf_json.get('has_scanner') else 'NO'}")
    
    chute = dxf_json.get('chute_analysis', {})
    if chute.get('total', 0) > 0:
        lines.append(f"  • Total Chutes: {chute['total']}")
        if chute.get('by_type'):
            for ct, cnt in chute['by_type'].items():
                lines.append(f"    - {ct.title()}: {cnt}")
    lines.append("")
    
    # Component categories with RAW names
    lines.append("CATEGORIZED COMPONENTS (with raw block names):")
    lines.append("-" * 70)
    
    categorized = dxf_json.get("categorized_components", {})
    raw_by_cat = dxf_json.get("raw_components_by_category", {})
    
    priority = ["AUTO_INDUCT", "OPERATOR_STATION", "CONVEYOR_INFEED", "VDS_BUFFER",
               "CBS_SORTER", "SCANNER", "CHUTE", "PTL", "BAG_SYSTEM", 
               "CONVEYOR_TAKEAWAY", "RECIRCULATION", "COLLECTION"]
    
    for cat in priority:
        if cat in categorized and categorized[cat]:
            lines.append("")
            lines.append(f"[{cat}] - {sum(categorized[cat].values())} total")
            
            for name, count in sorted(categorized[cat].items(), key=lambda x: -x[1]):
                lines.append(f"  • {name}: {count}")
                
                if cat in raw_by_cat and name in raw_by_cat[cat]:
                    examples = raw_by_cat[cat][name]
                    if len(examples) <= 3:
                        for ex in examples:
                            lines.append(f"      Raw: {ex}")
                    else:
                        lines.append(f"      Raw: {examples[0]}, ... ({len(examples)} variations)")
    
    # UNCATEGORIZED
    if "UNCATEGORIZED" in categorized and categorized["UNCATEGORIZED"]:
        lines.append("")
        lines.append(f"[UNCATEGORIZED] - {sum(categorized['UNCATEGORIZED'].values())} total")
        lines.append("⚠️  Review these for missing components!")
        lines.append("")
        
        for name, count in sorted(categorized["UNCATEGORIZED"].items(), key=lambda x: -x[1]):
            lines.append(f"  • {name}: {count}")
            if "UNCATEGORIZED" in raw_by_cat and name in raw_by_cat["UNCATEGORIZED"]:
                examples = raw_by_cat["UNCATEGORIZED"][name]
                for ex in examples[:2]:
                    lines.append(f"      Raw: {ex}")
    
    lines.append("")
    lines.append("=" * 70)
    lines.append("PROCESS FLOW GENERATION GUIDANCE:")
    lines.append("")
    
    cats = dxf_json.get("category_summary", {})
    
    lines.append(f"System Type: {dxf_json.get('cbs_type')}")
    lines.append(f"Induction: {dxf_json.get('induction_type')}")
    lines.append("")
    lines.append("Required Sections:")
    
    if cats.get("CONVEYOR_INFEED", 0) > 0:
        lines.append(f"  ✓ Infeed System ({cats['CONVEYOR_INFEED']} conveyors)")
    
    if dxf_json.get('has_vds'):
        lines.append(f"  ✓ VDS/Buffer ({cats.get('VDS_BUFFER', 0)} units)")
    
    if dxf_json.get('induction_type') == "MIXED (Auto + Manual)":
        lines.append(f"  ✓ Auto Induct ({cats.get('AUTO_INDUCT', 0)} lines)")
        lines.append(f"  ✓ Manual Stations ({cats.get('OPERATOR_STATION', 0)} stations)")
    elif cats.get("AUTO_INDUCT", 0) > 0:
        lines.append(f"  ✓ Auto Induct Line ({cats['AUTO_INDUCT']} lines)")
    elif cats.get("OPERATOR_STATION", 0) > 0:
        lines.append(f"  ✓ Manual Induct ({cats['OPERATOR_STATION']} stations)")
    
    if dxf_json.get('has_scanner'):
        lines.append(f"  ✓ Scanner/DWS in CBS section ({cats.get('SCANNER', 0)} units)")
    
    lines.append(f"  ✓ {dxf_json.get('cbs_type')} (sorting)")
    
    if chute.get('total', 0) > 0:
        lines.append(f"  ✓ Output Chutes ({chute['total']} total)")
        if chute.get('by_type'):
            for ct, cnt in sorted(chute['by_type'].items(), key=lambda x: -x[1]):
                lines.append(f"      - {ct.title()}: {cnt} chutes")
    
    if cats.get("PTL", 0) > 0:
        lines.append(f"  ✓ Put To Light ({cats['PTL']} locations)")
    
    if cats.get("BAG_SYSTEM", 0) > 0 or cats.get("CONVEYOR_TAKEAWAY", 0) > 0:
        total_bag = cats.get("BAG_SYSTEM", 0) + cats.get("CONVEYOR_TAKEAWAY", 0)
        lines.append(f"  ✓ Bag/Takeaway System ({total_bag} units)")
    
    if dxf_json.get('has_recirculation'):
        lines.append(f"  ✓ Recirculation Line ({cats.get('RECIRCULATION', 0)} units)")
    
    if cats.get("UNCATEGORIZED", 0) > 0:
        lines.append("")
        lines.append(f"  ⚠️  UNCATEGORIZED: {cats['UNCATEGORIZED']} components")
        lines.append("     Review raw names above - may contain:")
        lines.append("     - Additional conveyors, chutes, or systems")
    
    lines.append("")
    lines.append("=" * 70)
    
    return "\n".join(lines)


def calculate_component_similarity(query_cats: Dict, stored_cats: Dict,
                                   query_chute: Dict, stored_meta: Dict) -> float:
    """
    Calculate component-level similarity score (0-1).
    
    Higher score = more similar component counts.
    1.0 = identical, 0.0 = completely different
    
    Args:
        query_cats: Component counts from query DXF
        stored_cats: Component counts from stored DXF
        query_chute: Chute analysis from query DXF
        stored_meta: Stored metadata (may contain chute_analysis)
    
    Returns:
        float: Similarity score between 0 and 1
    """
    
    # Key categories to compare
    important_cats = [
        "AUTO_INDUCT", "OPERATOR_STATION", "CONVEYOR_INFEED",
        "VDS_BUFFER", "CHUTE", "RECIRCULATION", "PTL",
        "BAG_SYSTEM", "SCANNER"
    ]
    
    similarities = []
    
    for cat in important_cats:
        query_count = query_cats.get(cat, 0)
        stored_count = stored_cats.get(cat, 0)
        
        if query_count == 0 and stored_count == 0:
            # Both don't have it - perfect match for this category
            similarities.append(1.0)
        elif query_count == 0 or stored_count == 0:
            # One has it, other doesn't - poor match
            similarities.append(0.0)
        else:
            # Both have it - compare counts
            ratio = min(query_count, stored_count) / max(query_count, stored_count)
            similarities.append(ratio)
    
    # Average similarity across categories
    avg_similarity = sum(similarities) / len(similarities)
    
    # Bonus: Exact chute breakdown match
    query_chute_types = set(query_chute.get("by_type", {}).keys())
    
    # Try to extract from stored metadata (might not be structured)
    stored_chute_types = set()
    if "chute_analysis" in stored_meta:
        stored_chute_types = set(stored_meta["chute_analysis"].get("by_type", {}).keys())
    
    if query_chute_types and stored_chute_types:
        chute_type_overlap = len(query_chute_types & stored_chute_types) / len(query_chute_types | stored_chute_types)
        avg_similarity = (avg_similarity * 0.7) + (chute_type_overlap * 0.3)
    
    return avg_similarity
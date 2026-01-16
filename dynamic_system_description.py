"""
Dynamic System Description Generator

Generates system description by:
1. Loading component registry from Excel
2. Extracting DXF components
3. Matching DXF to Excel sheets via exact/alias/fuzzy matching
4. Rendering with deterministic templates + optional LLM polish
5. Injecting tables explicitly beneath components
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import re
from dataclasses import dataclass, asdict
import os

# For fuzzy matching
try:
    from rapidfuzz import fuzz
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    RAPIDFUZZ_AVAILABLE = False
    # Fallback: simple ratio function
    def fuzz_ratio(a: str, b: str) -> int:
        """Simple fallback fuzzy matching"""
        if a.lower() == b.lower():
            return 100
        # Count matching characters
        matches = sum(1 for x, y in zip(a.lower(), b.lower()) if x == y)
        return int((matches / max(len(a), len(b))) * 100) if max(len(a), len(b)) > 0 else 0
    
    class fuzz:
        @staticmethod
        def ratio(a: str, b: str) -> int:
            return fuzz_ratio(a, b)

logger = logging.getLogger(__name__)

# Phase order for component organization
PHASES = ["INFEED", "INDUCT", "MAIN_LOOP", "OUTPUT", "BAGGING"]

# Directory for catalogs
SCRIPT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
CATALOG_PATH = SCRIPT_DIR / "component_catalog.json"


def normalize_name(name: str) -> str:
    """Normalize component name for matching"""
    if not name:
        return ""
    return re.sub(r"[\s_\-]+", " ", str(name)).strip().lower()


@dataclass
class DXFComponent:
    """Represents a component found in DXF"""
    name: str
    stage: Optional[str] = None
    raw: Optional[str] = None  # Original DXF token/category
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class EligibleComponent:
    """Represents a component that was matched and is eligible for inclusion"""
    dxf_component: DXFComponent
    excel_sheet_name: str  # Original Excel sheet name
    normalized_sheet_name: str  # Normalized version
    table: Optional[Any] = None  # DataFrame or list of rows
    key_values: Optional[Dict[str, Any]] = None  # Extracted fields from sheet
    matching_method: str = "unknown"  # exact, alias, fuzzy
    matching_score: float = 0.0
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            "dxf_component": self.dxf_component.to_dict() if self.dxf_component else None,
            "excel_sheet_name": self.excel_sheet_name,
            "normalized_sheet_name": self.normalized_sheet_name,
            "matching_method": self.matching_method,
            "matching_score": self.matching_score,
            "key_values": self.key_values,
        }


def load_catalog() -> Dict[str, Dict[str, Any]]:
    """Load component catalog from JSON"""
    if CATALOG_PATH.exists():
        try:
            with open(CATALOG_PATH, "r") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Failed to load catalog: {e}")
    return {}


def get_component_stage(component_name: str, catalog: Dict[str, Dict[str, Any]]) -> Optional[str]:
    """Get stage for component from catalog"""
    norm_name = normalize_name(component_name)
    
    # Exact match
    if norm_name in catalog:
        return catalog[norm_name].get("stage")
    
    # Alias match
    for cat_comp, cat_data in catalog.items():
        aliases = [normalize_name(a) for a in cat_data.get("aliases", [])]
        if norm_name in aliases:
            return cat_data.get("stage")
    
    return None


def get_component_priority(component_name: str, catalog: Dict[str, Dict[str, Any]]) -> int:
    """Get priority for component from catalog (lower = first)"""
    norm_name = normalize_name(component_name)
    
    # Exact match
    if norm_name in catalog:
        return catalog[norm_name].get("priority", 999)
    
    # Alias match
    for cat_comp, cat_data in catalog.items():
        aliases = [normalize_name(a) for a in cat_data.get("aliases", [])]
        if norm_name in aliases:
            return cat_data.get("priority", 999)
    
    return 999


def match_dxf_to_excel(
    dxf_components: List[DXFComponent],
    excel_registry: Dict[str, Dict[str, Any]],
    catalog: Dict[str, Dict[str, Any]],
    fuzzy_threshold: int = 90
) -> Tuple[List[EligibleComponent], List[str]]:
    """
    Match DXF components to Excel sheets.
    
    Returns:
        (eligible_components, unmatched_dxf_names)
    """
    eligible = []
    unmatched = []
    
    for dxf_comp in dxf_components:
        norm_dxf_name = normalize_name(dxf_comp.name)
        matched = False
        
        # Strategy 1: Exact match against normalized excel sheet names
        if norm_dxf_name in excel_registry:
            excel_entry = excel_registry[norm_dxf_name]
            eligible.append(EligibleComponent(
                dxf_component=dxf_comp,
                excel_sheet_name=excel_entry["sheet_name"],
                normalized_sheet_name=norm_dxf_name,
                table=excel_entry.get("table"),
                key_values=excel_entry.get("key_values"),
                matching_method="exact",
                matching_score=100.0
            ))
            matched = True
        
        # Strategy 2: Alias match via catalog
        if not matched and catalog:
            for cat_comp, cat_data in catalog.items():
                aliases = [normalize_name(a) for a in cat_data.get("aliases", [])]
                if norm_dxf_name in aliases:
                    # Found via alias, now look for exact match in excel
                    cat_norm = normalize_name(cat_comp)
                    if cat_norm in excel_registry:
                        excel_entry = excel_registry[cat_norm]
                        eligible.append(EligibleComponent(
                            dxf_component=dxf_comp,
                            excel_sheet_name=excel_entry["sheet_name"],
                            normalized_sheet_name=cat_norm,
                            table=excel_entry.get("table"),
                            key_values=excel_entry.get("key_values"),
                            matching_method="alias",
                            matching_score=95.0
                        ))
                        matched = True
                        break
        
        # Strategy 3: Fuzzy match against normalized excel sheet names
        if not matched and RAPIDFUZZ_AVAILABLE:
            best_score = 0
            best_match = None
            
            for excel_norm, excel_entry in excel_registry.items():
                score = fuzz.ratio(norm_dxf_name, excel_norm)
                if score > best_score and score >= fuzzy_threshold:
                    best_score = score
                    best_match = excel_entry
            
            if best_match:
                eligible.append(EligibleComponent(
                    dxf_component=dxf_comp,
                    excel_sheet_name=best_match["sheet_name"],
                    normalized_sheet_name=best_match["sheet_name"].lower(),
                    table=best_match.get("table"),
                    key_values=best_match.get("key_values"),
                    matching_method="fuzzy",
                    matching_score=float(best_score)
                ))
                matched = True
        
        # If still not matched, include but mark as missing table
        if not matched:
            eligible.append(EligibleComponent(
                dxf_component=dxf_comp,
                excel_sheet_name="",
                normalized_sheet_name="",
                table=None,
                key_values=None,
                matching_method="none",
                matching_score=0.0
            ))
            unmatched.append(dxf_comp.name)
    
    return eligible, unmatched


def sort_components_by_phase_and_priority(
    eligible: List[EligibleComponent],
    catalog: Dict[str, Dict[str, Any]]
) -> List[EligibleComponent]:
    """Sort components by phase order, then by priority within phase"""
    
    def get_sort_key(comp: EligibleComponent) -> Tuple[int, int]:
        """Return (phase_index, priority)"""
        stage = comp.dxf_component.stage
        
        if stage and stage in PHASES:
            phase_idx = PHASES.index(stage)
        else:
            phase_idx = 999
        
        priority = get_component_priority(comp.dxf_component.name, catalog)
        return (phase_idx, priority)
    
    return sorted(eligible, key=get_sort_key)


def build_deterministic_section(
    component: EligibleComponent,
    catalog: Dict[str, Dict[str, Any]]
) -> str:
    """
    Build sales-focused component text using values from Excel.
    Emphasizes WHY (benefits) not just WHAT (features).
    Written in human storytelling style.
    
    Example output:
    "Weighing Conveyors: Each parcel is accurately weighed before induction,
    ensuring you have precise shipping costs and zero billing disputes. From there..."
    """
    comp_name = component.dxf_component.name
    key_vals = component.key_values or {}
    
    # Start with component name
    section = f"{comp_name.title()}: "
    
    # Add sales-focused, benefit-driven context based on stage
    stage = component.dxf_component.stage
    
    if stage == "INFEED":
        section += "Incoming shipments arrive here, where they're smoothly prepared for processing—no bottlenecks, no delays. "
        if key_vals.get("quantity"):
            section += f"With capacity for {key_vals['quantity']} units, your peak volumes are handled effortlessly. "
        if key_vals.get("speed"):
            section += f"Operating at {key_vals['speed']} m/min keeps your throughput consistently high. "
        section += "From here, parcels move seamlessly to the induction system. "
    
    elif stage == "INDUCT":
        section += "Parcels are positioned and oriented for the sorter, ensuring smooth handoff without jams. "
        if key_vals.get("quantity"):
            section += f"Processing {key_vals['quantity']} items per cycle means your operators stay productive. "
        if key_vals.get("speed"):
            section += f"At {key_vals['speed']} m/min, you maintain the pace your operation demands. "
        section += "This feeds directly into the main sorting loop for high-speed processing. "
    
    elif stage == "MAIN_LOOP":
        section += "This is where the magic happens—parcels are sorted at high speed with pinpoint accuracy. "
        if key_vals.get("speed"):
            section += f"Running at {key_vals['speed']} m/min, you get the throughput you need. "
        if key_vals.get("capacity"):
            section += f"With {key_vals['capacity']} parcels/hour capacity, even your busiest days are covered. "
        section += "Barcodes are scanned instantly, and each parcel is routed to its correct destination without manual intervention. "
    
    elif stage == "OUTPUT":
        section += "Sorted parcels slide smoothly into their designated chutes, ready for dispatch. "
        if key_vals.get("quantity"):
            section += f"With {key_vals['quantity']} output points, you have flexibility for multiple destinations. "
        section += "Each chute serves a specific route or criterion—keeping your downstream operations efficient. "
    
    elif stage == "BAGGING":
        section += "Parcels are collected and prepared for bagging, streamlining your packaging workflow. "
        if key_vals.get("quantity"):
            section += f"Handling {key_vals['quantity']} parcels/hour keeps pace with your sorting output. "
        section += "This final step ensures packages are ready for shipment with minimal handling. "
    
    return section


def render_system_description(
    eligible_components: List[EligibleComponent],
    catalog: Dict[str, Dict[str, Any]],
    use_lm_polish: bool = False,
    groq_api_key: Optional[str] = None,
) -> Tuple[str, Dict[str, Any]]:
    """
    Render full system description from eligible components.
    
    Returns:
        (final_text, diagnostics)
    """
    diagnostics = {
        "total_components": len(eligible_components),
        "components_with_tables": 0,
        "components_missing_tables": 0,
        "unmatched_components": [],
        "phase_distribution": {}
    }
    
    # Sort by phase and priority
    sorted_comps = sort_components_by_phase_and_priority(eligible_components, catalog)
    
    # Group by phase
    phases_content = {phase: [] for phase in PHASES}
    
    for comp in sorted_comps:
        stage = comp.dxf_component.stage or "MAIN_LOOP"  # Default stage
        
        if stage in phases_content:
            # Build deterministic section
            section_text = build_deterministic_section(comp, catalog)
            
            # Add table reference if available
            if comp.table is not None:
                diagnostics["components_with_tables"] += 1
                section_text += f"\n\n[[TABLE:{comp.excel_sheet_name}]]"
            else:
                diagnostics["components_missing_tables"] += 1
                if comp.matching_method == "none":
                    diagnostics["unmatched_components"].append(comp.dxf_component.name)
            
            phases_content[stage].append(section_text)
            
            # Track phase distribution
            diagnostics["phase_distribution"][stage] = diagnostics["phase_distribution"].get(stage, 0) + 1
    
    # Assemble final text with sales-oriented phase headers
    final_parts = []
    for phase in PHASES:
        if phases_content[phase]:
            # Sales-friendly headers that emphasize the journey
            phase_header = {
                "INFEED": "Where It All Begins: Infeed System",
                "INDUCT": "Precision Handling: Induction System",
                "MAIN_LOOP": "High-Speed Sorting: The Main Loop",
                "OUTPUT": "Ready for Dispatch: Output Chutes",
                "BAGGING": "Final Touch: Bagging System"
            }.get(phase, phase.replace("_", " ").title())
            
            final_parts.append(f"\n## {phase_header}\n")
            final_parts.extend(phases_content[phase])
    
    final_text = "\n\n".join(final_parts)
    
    # Optional LLM polish
    if use_lm_polish and groq_api_key:
        final_text = _polish_with_llm(final_text, groq_api_key)
    
    return final_text, diagnostics


def _polish_with_llm(text: str, api_key: str) -> str:
    """
    Polish wording with LLM without changing structure/facts.
    
    Note: Requires groq integration. For now, returns text as-is.
    """
    # TODO: Implement LLM polish call
    # For now, return as-is
    return text


def generate_dynamic_system_description(
    dxf_components: List[DXFComponent],
    excel_registry: Dict[str, Dict[str, Any]],
    catalog: Optional[Dict[str, Dict[str, Any]]] = None,
    use_lm_polish: bool = False,
    groq_api_key: Optional[str] = None,
    fuzzy_threshold: int = 90
) -> Tuple[str, Dict[str, Any]]:
    """
    Main entry point: Generate system description dynamically.
    
    Args:
        dxf_components: List of DXFComponent objects from DXF
        excel_registry: Registry from load_component_sheets()
        catalog: Component catalog (loaded from JSON if not provided)
        use_lm_polish: Whether to polish with LLM
        groq_api_key: API key for LLM polish
        fuzzy_threshold: Fuzzy match threshold (0-100)
    
    Returns:
        (system_description_text, diagnostics)
    """
    if catalog is None:
        catalog = load_catalog()
    
    # Match DXF to Excel
    eligible, unmatched = match_dxf_to_excel(
        dxf_components,
        excel_registry,
        catalog,
        fuzzy_threshold=fuzzy_threshold
    )
    
    # Render with diagnostics
    sys_desc, diagnostics = render_system_description(
        eligible,
        catalog,
        use_lm_polish=use_lm_polish,
        groq_api_key=groq_api_key
    )
    
    # Log diagnostics
    logger.info(f"System Description Generated:")
    logger.info(f"  DXF Components: {len(dxf_components)}")
    logger.info(f"  Rendered: {diagnostics['total_components']}")
    logger.info(f"  With Tables: {diagnostics['components_with_tables']}")
    logger.info(f"  Missing Tables: {diagnostics['components_missing_tables']}")
    logger.info(f"  Unmatched: {diagnostics['unmatched_components']}")
    logger.info(f"  Phase Distribution: {diagnostics['phase_distribution']}")
    
    return sys_desc, diagnostics

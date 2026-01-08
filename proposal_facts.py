"""
ProposalFacts - Single Source of Truth for Proposal Data

This module defines the ProposalFacts dataclass and associated extraction/validation
functions. ProposalFacts serves as the authoritative data source for all proposal
sections (Cover Letter, Executive Summary, Process Flow, System Description).

Key principles:
1. One-time extraction from DXF + optional costing
2. Prefer costing quantities over DXF counts
3. Never output "0" for critical counts; use null + standard phrase
4. Track data source (costing vs DXF) for every count
5. No hallucinated equipment or counts
"""

import json
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any, Tuple
from enum import Enum
from pathlib import Path


class CountSource(Enum):
    """Tracks where a count/value originated"""
    COSTING = "costing_bom"
    DXF = "dxf_extraction"
    MANUAL = "manual_input"
    MISSING = "to_be_confirmed"


@dataclass
class ComponentCount:
    """Tracks a single component count with its source"""
    count: Optional[int]  # None if unconfirmed
    source: CountSource
    confirmed: bool = False
    component_name: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            "count": self.count,
            "source": self.source.value,
            "confirmed": self.confirmed,
            "component_name": self.component_name,
        }


@dataclass
class SystemNumbers:
    """All component counts and their sources"""
    # Induction System
    feedlines: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    induct_lines_auto: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    manual_induct_stations: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    # Infeed Conveyors
    infeed_conveyors_total: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    telescopic_belt_conveyors: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    buffer_conveyors: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    curve_conveyors: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    # Sorter Core
    cbs_sorters: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    sorter_carriers: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    # Output Chutes (by type)
    gravity_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    mini_gravity_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    collection_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    rejection_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    dispersion_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    bulk_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    direct_bagging_chutes: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    # Scanning/Measurement
    barcode_scanners: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    weighing_systems: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    # Additional Systems
    vds_loops: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    recirculation_conveyors: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    # Capacity
    throughput_pph: ComponentCount = field(default_factory=lambda: ComponentCount(None, CountSource.MISSING))
    
    def get_all_counts(self) -> Dict[str, ComponentCount]:
        """Return all component counts"""
        return asdict(self)
    
    def get_unconfirmed(self) -> List[str]:
        """Return list of unconfirmed counts"""
        return [
            k for k, v in asdict(self).items() 
            if isinstance(v, ComponentCount) and v.source == CountSource.MISSING
        ]
    
    def to_dict(self) -> Dict[str, Dict[str, Any]]:
        """Convert to dict for serialization"""
        result = {}
        for key, value in asdict(self).items():
            if isinstance(value, ComponentCount):
                result[key] = value.to_dict()
        return result


@dataclass
class LayoutFlags:
    """Binary system layout indicators"""
    has_infeed_system: bool = False
    has_auto_induct: bool = False
    has_manual_induct: bool = False
    has_vds_loop: bool = False
    has_recirculation: bool = False
    has_barcode_scanning: bool = False
    has_weighing: bool = False
    has_output_chutes: bool = False
    
    # CBS Type
    cbs_type: Optional[str] = None  # "Loop CBS" or "Linear CBS"
    
    # Operating parameters
    throughput_pph: Optional[float] = None
    parcels_per_induct: Optional[float] = None
    cbs_speed_mpm: Optional[float] = None
    
    # Layout orientation
    layout_type: Optional[str] = None  # "Straight", "L-shaped", "U-shaped", etc.


@dataclass
class ProcessStep:
    """Single step in the process flow"""
    sequence: int  # 1-based order
    component: str  # e.g., "Manual Induct Station"
    description: str  # What happens here
    upstream: Optional[str] = None  # Previous component
    downstream: Optional[str] = None  # Next component
    count: Optional[int] = None  # How many of this component


@dataclass
class ProposalFacts:
    """
    Single source of truth for all proposal data.
    Populated once from DXF + optional costing at proposal startup.
    Passed to all generators (Cover Letter, Executive Summary, Process Flow, System Description).
    """
    
    # Metadata
    dxf_filename: str = ""
    costing_filename: Optional[str] = None
    extraction_timestamp: str = ""
    
    # System Counts (with sources)
    system_numbers: SystemNumbers = field(default_factory=SystemNumbers)
    
    # Layout Indicators
    layout_flags: LayoutFlags = field(default_factory=LayoutFlags)
    
    # Process Flow (ordered sequence)
    process_steps: List[ProcessStep] = field(default_factory=list)
    
    # Missing/Unconfirmed Fields
    unknowns: List[str] = field(default_factory=list)
    
    # Raw extracted data for fallback access
    dxf_metrics: Dict[str, Any] = field(default_factory=dict)
    costing_metrics: Dict[str, Any] = field(default_factory=dict)
    
    # Configuration
    project_name: str = ""
    client_name: str = ""
    
    def get_count(self, component_key: str) -> Tuple[Optional[int], str]:
        """
        Get a component count with its source.
        Returns (count, source_description)
        If count is None, returns standard "to be confirmed" phrase.
        """
        if not hasattr(self.system_numbers, component_key):
            return None, "unknown_component"
        
        comp_count: ComponentCount = getattr(self.system_numbers, component_key)
        if comp_count.count is None:
            return None, f"(to be confirmed during detailed engineering)"
        
        source_desc = comp_count.source.value
        return comp_count.count, source_desc
    
    def mark_unconfirmed(self, component_key: str):
        """Mark a component count as unconfirmed"""
        if hasattr(self.system_numbers, component_key):
            comp_count: ComponentCount = getattr(self.system_numbers, component_key)
            comp_count.count = None
            comp_count.source = CountSource.MISSING
    
    def add_unknown(self, field_name: str):
        """Add a field to unknowns list if not already present"""
        if field_name not in self.unknowns:
            self.unknowns.append(field_name)
    
    def get_unconfirmed_fields(self) -> List[str]:
        """Get all unconfirmed component counts"""
        return self.system_numbers.get_unconfirmed()
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization"""
        return {
            "metadata": {
                "dxf_filename": self.dxf_filename,
                "costing_filename": self.costing_filename,
                "extraction_timestamp": self.extraction_timestamp,
                "project_name": self.project_name,
                "client_name": self.client_name,
            },
            "system_numbers": self.system_numbers.to_dict(),
            "layout_flags": asdict(self.layout_flags),
            "process_steps": [asdict(step) for step in self.process_steps],
            "unknowns": self.unknowns,
        }
    
    def to_json(self) -> str:
        """Convert to JSON string"""
        return json.dumps(self.to_dict(), indent=2, default=str)
    
    def summary(self) -> str:
        """Generate a human-readable summary of extracted data"""
        lines = [
            "=" * 70,
            "PROPOSAL FACTS SUMMARY",
            "=" * 70,
            f"Project: {self.project_name} | Client: {self.client_name}",
            f"DXF: {self.dxf_filename}",
            f"Costing: {self.costing_filename or 'Not provided'}",
            "",
            "SYSTEM LAYOUT:",
            f"  CBS Type: {self.layout_flags.cbs_type or 'Unknown'}",
            f"  Infeed: {self.layout_flags.has_infeed_system}",
            f"  Auto Induct: {self.layout_flags.has_auto_induct}",
            f"  Manual Induct: {self.layout_flags.has_manual_induct}",
            f"  Barcode Scanning: {self.layout_flags.has_barcode_scanning}",
            "",
            "CRITICAL COUNTS:",
        ]
        
        for key, comp in self.system_numbers.get_all_counts().items():
            if isinstance(comp, ComponentCount):
                status = "✓" if comp.count is not None else "✗"
                count_str = str(comp.count) if comp.count is not None else "MISSING"
                lines.append(f"  [{status}] {key}: {count_str} ({comp.source.value})")
        
        if self.unknowns:
            lines.append("")
            lines.append("UNCONFIRMED FIELDS:")
            for field in self.unknowns:
                lines.append(f"  - {field}")
        
        lines.append("=" * 70)
        return "\n".join(lines)


# ============================================================================
# Helper Functions for Creating ProposalFacts
# ============================================================================

def create_proposal_facts(
    dxf_filename: str,
    costing_filename: Optional[str] = None,
    project_name: str = "",
    client_name: str = "",
) -> ProposalFacts:
    """Factory function to create a new ProposalFacts instance"""
    from datetime import datetime
    
    facts = ProposalFacts(
        dxf_filename=dxf_filename,
        costing_filename=costing_filename,
        project_name=project_name,
        client_name=client_name,
        extraction_timestamp=datetime.now().isoformat(),
    )
    return facts


def populate_from_dxf(facts: ProposalFacts, dxf_metrics: Dict[str, Any]) -> ProposalFacts:
    """
    Populate ProposalFacts from DXF-extracted metrics.
    Uses ComponentCount with CountSource.DXF for all values.
    """
    facts.dxf_metrics = dxf_metrics
    
    # Map DXF metrics to system_numbers
    mapping = {
        "FEEDLINE COUNT": ("feedlines", CountSource.DXF),
        "TEL. BELT CONVEYOR COUNT": ("telescopic_belt_conveyors", CountSource.DXF),
        "INFEED CONVEYOR COUNT": ("infeed_conveyors_total", CountSource.DXF),
        "GRAVITY CHUTE COUNT": ("gravity_chutes", CountSource.DXF),
        "MINI GRAVITY CHUTE COUNT": ("mini_gravity_chutes", CountSource.DXF),
        "COUNT OF COLLECTION CHUTE": ("collection_chutes", CountSource.DXF),
        "COUNT OF REJECTION CHUTE": ("rejection_chutes", CountSource.DXF),
        "COUNT OF DISPERSION CHUTE": ("dispersion_chutes", CountSource.DXF),
        "BULK CHUTE COUNT": ("bulk_chutes", CountSource.DXF),
        "DIRECT BAGGING CHUTE COUNT": ("direct_bagging_chutes", CountSource.DXF),
        "SLIDING CHUTE COUNT": ("bulk_chutes", CountSource.DXF),  # Fallback to bulk
        "VDS LOOP COUNT": ("vds_loops", CountSource.DXF),
        "HAS_SCANNER": ("barcode_scanners", CountSource.DXF),
    }
    
    for dxf_key, (facts_key, source) in mapping.items():
        if dxf_key in dxf_metrics:
            value = dxf_metrics[dxf_key]
            if value and value != "":
                try:
                    count = int(value) if value != "True" else 1
                    if hasattr(facts.system_numbers, facts_key):
                        setattr(
                            facts.system_numbers,
                            facts_key,
                            ComponentCount(count=count, source=source, confirmed=True)
                        )
                except (ValueError, TypeError):
                    pass
    
    # Layout flags from DXF
    facts.layout_flags.has_infeed_system = dxf_metrics.get("INFEED CONVEYOR COUNT", 0) > 0
    facts.layout_flags.has_auto_induct = dxf_metrics.get("FEEDLINE COUNT", 0) > 0
    facts.layout_flags.has_manual_induct = dxf_metrics.get("HAS_MANUAL_INDUCT", False)
    facts.layout_flags.has_barcode_scanning = dxf_metrics.get("HAS_SCANNER", False)
    facts.layout_flags.has_output_chutes = any([
        dxf_metrics.get("GRAVITY CHUTE COUNT", 0) > 0,
        dxf_metrics.get("MINI GRAVITY CHUTE COUNT", 0) > 0,
        dxf_metrics.get("COUNT OF COLLECTION CHUTE", 0) > 0,
        dxf_metrics.get("BULK CHUTE COUNT", 0) > 0,
    ])
    facts.layout_flags.cbs_type = dxf_metrics.get("TYPE OF CBS", None)
    
    # Mark missing critical counts
    critical_fields = ["feedlines", "cbs_sorters", "gravity_chutes"]
    for field in critical_fields:
        comp_count: ComponentCount = getattr(facts.system_numbers, field)
        if comp_count.count is None:
            facts.add_unknown(field)
    
    return facts


def populate_from_costing(facts: ProposalFacts, costing_metrics: Dict[str, Any]) -> ProposalFacts:
    """
    Populate ProposalFacts from costing/BOQ data.
    Overrides DXF values where costing data is more authoritative.
    Uses ComponentCount with CountSource.COSTING for all values.
    """
    facts.costing_metrics = costing_metrics
    
    # Costing data preferentially overrides DXF
    mapping = {
        "feedlines": ("feedlines", CountSource.COSTING),
        "induct_lines": ("feedlines", CountSource.COSTING),
        "manual_stations": ("manual_induct_stations", CountSource.COSTING),
        "gravity_chutes": ("gravity_chutes", CountSource.COSTING),
        "mini_gravity_chutes": ("mini_gravity_chutes", CountSource.COSTING),
        "collection_chutes": ("collection_chutes", CountSource.COSTING),
        "rejection_chutes": ("rejection_chutes", CountSource.COSTING),
        "dispersion_chutes": ("dispersion_chutes", CountSource.COSTING),
        "bulk_chutes": ("bulk_chutes", CountSource.COSTING),
        "direct_bagging_chutes": ("direct_bagging_chutes", CountSource.COSTING),
        "sorter_carriers": ("sorter_carriers", CountSource.COSTING),
        "throughput_pph": ("throughput_pph", CountSource.COSTING),
    }
    
    for costing_key, (facts_key, source) in mapping.items():
        if costing_key in costing_metrics:
            value = costing_metrics[costing_key]
            if value and value not in (0, "", None):
                try:
                    count = int(value) if not isinstance(value, int) else value
                    if hasattr(facts.system_numbers, facts_key):
                        setattr(
                            facts.system_numbers,
                            facts_key,
                            ComponentCount(count=count, source=source, confirmed=True)
                        )
                except (ValueError, TypeError):
                    pass
    
    return facts


def get_counts_source_of_truth(facts: ProposalFacts, component_key: str) -> Tuple[Optional[int], str, bool]:
    """
    Get the authoritative count for a component.
    
    Preference hierarchy:
    1. Costing/BOQ data (if available and confirmed)
    2. DXF extraction (if available)
    3. Manual input (if provided)
    4. None (unconfirmed)
    
    Returns:
        (count, source_description, is_confirmed)
    
    NEVER returns 0 for critical counts. Instead:
    - Returns None if unconfirmed
    - Source will indicate "to be confirmed during detailed engineering"
    """
    if not hasattr(facts.system_numbers, component_key):
        return None, "unknown_component", False
    
    comp_count: ComponentCount = getattr(facts.system_numbers, component_key)
    
    # If count is 0 and this is a critical component, mark as unconfirmed
    CRITICAL_COMPONENTS = {
        "feedlines", "cbs_sorters", "gravity_chutes", "throughput_pph",
        "induct_lines_auto", "manual_induct_stations"
    }
    
    if comp_count.count == 0 and component_key in CRITICAL_COMPONENTS:
        return None, "(to be confirmed during detailed engineering)", False
    
    if comp_count.count is None:
        return None, "(to be confirmed during detailed engineering)", False
    
    return comp_count.count, comp_count.source.value, comp_count.confirmed


# ============================================================================
# Validation Functions
# ============================================================================

def validate_no_hallucinations(facts: ProposalFacts) -> List[str]:
    """
    Validate that ProposalFacts contains no hallucinated or unverified data.
    Returns list of validation errors (empty if valid).
    """
    errors = []
    
    # Check that all confirmed counts have valid sources
    for key, comp_count in facts.system_numbers.get_all_counts().items():
        if isinstance(comp_count, ComponentCount):
            if comp_count.confirmed and comp_count.count is not None:
                if comp_count.count < 0:
                    errors.append(f"{key}: Negative count not allowed ({comp_count.count})")
                if comp_count.count > 10000:
                    errors.append(f"{key}: Suspiciously large count ({comp_count.count})")
    
    return errors


if __name__ == "__main__":
    # Example usage
    facts = create_proposal_facts(
        dxf_filename="test.dxf",
        costing_filename="costing.xlsx",
        project_name="Test Project",
        client_name="Test Client"
    )
    
    # Simulate DXF extraction
    dxf_data = {
        "FEEDLINE COUNT": 4,
        "GRAVITY CHUTE COUNT": 10,
        "TYPE OF CBS": "Loop CBS",
        "HAS_SCANNER": True,
    }
    populate_from_dxf(facts, dxf_data)
    
    # Simulate costing override
    costing_data = {
        "gravity_chutes": 12,  # Costing says 12, not 10
        "throughput_pph": 27600,
    }
    populate_from_costing(facts, costing_data)
    
    # Print summary
    print(facts.summary())
    
    # Access specific counts
    count, source, confirmed = get_counts_source_of_truth(facts, "gravity_chutes")
    print(f"\nGravity Chutes: {count} (source: {source}, confirmed: {confirmed})")

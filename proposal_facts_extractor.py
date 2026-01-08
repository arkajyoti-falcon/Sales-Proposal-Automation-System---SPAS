"""
ProposalFacts Extractor - Unified extraction from DXF + Costing

This module coordinates the extraction and population of ProposalFacts
from DXF files and optional costing Excel workbooks.

Main entry point: extract_proposal_facts(dxf_path, costing_file=None)
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import logging
import sys

from proposal_facts import (
    ProposalFacts,
    ComponentCount,
    CountSource,
    ProcessStep,
    create_proposal_facts,
    populate_from_dxf,
    populate_from_costing,
    validate_no_hallucinations,
)
from costing_sheet_mapper import CostingSheetMapper

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


class ProposalFactsExtractor:
    """Unified extractor for DXF + costing data"""
    
    def __init__(self, dxf_path: str, costing_path: Optional[str] = None,
                 project_name: str = "", client_name: str = ""):
        """
        Initialize extractor.
        
        Args:
            dxf_path: Path to DXF file
            costing_path: Path to costing Excel file (optional)
            project_name: Project name for metadata
            client_name: Client name for metadata
        """
        self.dxf_path = Path(dxf_path)
        self.costing_path = Path(costing_path) if costing_path else None
        self.project_name = project_name
        self.client_name = client_name
        
        self.facts: Optional[ProposalFacts] = None
        self.dxf_metrics: Dict[str, Any] = {}
        self.costing_metrics: Dict[str, Any] = {}
        self.costing_mapper: Optional[CostingSheetMapper] = None
    
    def extract(self) -> ProposalFacts:
        """
        Extract all data and populate ProposalFacts.
        
        Returns:
            Populated ProposalFacts object
        """
        # Create facts object
        self.facts = create_proposal_facts(
            dxf_filename=self.dxf_path.name,
            costing_filename=self.costing_path.name if self.costing_path else None,
            project_name=self.project_name,
            client_name=self.client_name,
        )
        
        # Extract from DXF
        logger.info(f"Extracting DXF data from {self.dxf_path}")
        self._extract_dxf_metrics()
        self.facts = populate_from_dxf(self.facts, self.dxf_metrics)
        
        # Extract from costing (optional)
        if self.costing_path and self.costing_path.exists():
            logger.info(f"Extracting costing data from {self.costing_path}")
            self._extract_costing_metrics()
            self.facts = populate_from_costing(self.facts, self.costing_metrics)
        
        # Build process flow (optional)
        self._build_process_flow()
        
        # Validate
        errors = validate_no_hallucinations(self.facts)
        if errors:
            logger.warning(f"Validation issues found: {errors}")
        
        # Log summary
        logger.info(self.facts.summary())
        
        return self.facts
    
    def _extract_dxf_metrics(self):
        """Extract metrics from DXF file"""
        try:
            # Try to import and use st_sys_desc module
            try:
                import st_sys_desc as sd_sys
                
                full = sd_sys.extract_dxf_full_json(self.dxf_path)
                self.dxf_metrics = sd_sys.compute_dxf_metrics(full)
                logger.info(f"Successfully extracted DXF metrics: {len(self.dxf_metrics)} fields")
                
            except ImportError:
                logger.warning("st_sys_desc module not available; using fallback")
                self.dxf_metrics = self._extract_dxf_fallback()
        
        except Exception as e:
            logger.error(f"Failed to extract DXF metrics: {e}")
            self.dxf_metrics = {}
    
    def _extract_dxf_fallback(self) -> Dict[str, Any]:
        """Fallback DXF extraction if st_sys_desc unavailable"""
        # Placeholder for fallback extraction
        return {}
    
    def _extract_costing_metrics(self):
        """Extract metrics from costing Excel file"""
        try:
            self.costing_mapper = CostingSheetMapper(str(self.costing_path))
            
            # Get all sheet mappings
            sheet_mappings = self.costing_mapper.get_all_sheet_mappings()
            logger.info(f"Found {len(sheet_mappings)} sheets in costing workbook")
            
            # Extract counts for each component type
            for component_type in [
                "Loop CBS", "Linear CBS", "Conveyors", "Destinations",
                "Steelworks", "PTL", "Induct Lines", "Control System", "Safety Equipment"
            ]:
                try:
                    counts = self.costing_mapper.extract_component_counts(component_type)
                    if counts:
                        self.costing_metrics[component_type.lower()] = counts
                        logger.info(f"Extracted {component_type}: {len(counts)} items")
                except Exception as e:
                    logger.debug(f"Could not extract {component_type}: {e}")
        
        except Exception as e:
            logger.error(f"Failed to extract costing metrics: {e}")
            self.costing_metrics = {}
    
    def _build_process_flow(self):
        """Build process flow from extracted data"""
        if not self.facts:
            return
        
        flow_steps: List[ProcessStep] = []
        sequence = 1
        previous_component = None
        
        # Build flow based on layout flags
        if self.facts.layout_flags.has_infeed_system:
            flow_steps.append(ProcessStep(
                sequence=sequence,
                component="Infeed System",
                description="Bulk infeed conveyors receive and prepare parcels",
                upstream=None,
                downstream="Induction System",
            ))
            sequence += 1
            previous_component = "Infeed System"
        
        if self.facts.layout_flags.has_manual_induct:
            flow_steps.append(ProcessStep(
                sequence=sequence,
                component="Manual Induct Stations",
                description="Operators manually place parcels onto induction conveyors",
                upstream=previous_component,
                downstream="Induction System" if self.facts.layout_flags.has_auto_induct else "CBS",
            ))
            sequence += 1
            previous_component = "Manual Induct Stations"
        
        if self.facts.layout_flags.has_auto_induct:
            count, _, _ = self._get_count("feedlines")
            desc = f"Automatic induct lines prepare and position parcels for sorter"
            if count:
                desc += f" ({count} feedlines)"
            
            flow_steps.append(ProcessStep(
                sequence=sequence,
                component="Automatic Induct Lines",
                description=desc,
                upstream=previous_component or "Infeed System",
                downstream="CBS",
                count=count,
            ))
            sequence += 1
            previous_component = "Automatic Induct Lines"
        
        if self.facts.layout_flags.cbs_type:
            flow_steps.append(ProcessStep(
                sequence=sequence,
                component=f"{self.facts.layout_flags.cbs_type} Sorter",
                description=f"Main sorting mechanism using {self.facts.layout_flags.cbs_type} technology",
                upstream=previous_component,
                downstream="Output Chutes",
            ))
            sequence += 1
            previous_component = f"{self.facts.layout_flags.cbs_type} Sorter"
        
        if self.facts.layout_flags.has_barcode_scanning:
            flow_steps.append(ProcessStep(
                sequence=sequence,
                component="Barcode Scanning System",
                description="Scanners read barcodes to determine parcel destinations",
                upstream=previous_component,
                downstream="Output Chutes",
            ))
            sequence += 1
        
        if self.facts.layout_flags.has_output_chutes:
            total_chutes = self._count_total_chutes()
            desc = "Parcels are discharged to designated output chutes"
            if total_chutes:
                desc += f" ({total_chutes} chutes total)"
            
            flow_steps.append(ProcessStep(
                sequence=sequence,
                component="Output Chutes",
                description=desc,
                upstream=previous_component,
                downstream=None,
                count=total_chutes,
            ))
        
        self.facts.process_steps = flow_steps
    
    def _get_count(self, component_key: str) -> Tuple[Optional[int], str, bool]:
        """Get a count from facts"""
        if not self.facts:
            return None, "unknown", False
        
        count, source, confirmed = self.facts.get_count(component_key)
        return count, source, confirmed
    
    def _count_total_chutes(self) -> Optional[int]:
        """Sum all chute counts"""
        if not self.facts:
            return None
        
        total = 0
        chute_keys = [
            "gravity_chutes", "mini_gravity_chutes", "collection_chutes",
            "rejection_chutes", "dispersion_chutes", "bulk_chutes",
            "direct_bagging_chutes",
        ]
        
        for key in chute_keys:
            if hasattr(self.facts.system_numbers, key):
                comp = getattr(self.facts.system_numbers, key)
                if comp.count:
                    total += comp.count
        
        return total if total > 0 else None
    
    def get_sheet_text_for_ai(self, component_type: str) -> Optional[str]:
        """
        Get full sheet text for a component type to send to AI.
        
        Args:
            component_type: e.g., "Loop CBS", "Conveyors"
        
        Returns:
            Formatted sheet text, or None if not found
        """
        if not self.costing_mapper:
            return None
        
        try:
            return self.costing_mapper.extract_sheet_as_text(component_type, max_rows=None)
        except Exception as e:
            logger.warning(f"Could not extract sheet text for {component_type}: {e}")
            return None
    
    def log_extraction_summary(self) -> str:
        """Generate a detailed extraction summary"""
        if not self.facts:
            return "No facts extracted"
        
        lines = [
            "=" * 80,
            "PROPOSAL FACTS EXTRACTION SUMMARY",
            "=" * 80,
            f"DXF File: {self.dxf_path.name}",
            f"Costing File: {self.costing_path.name if self.costing_path else 'Not provided'}",
            "",
            "EXTRACTED DATA SOURCES:",
            f"  DXF Metrics: {len(self.dxf_metrics)} fields",
            f"  Costing Metrics: {len(self.costing_metrics)} entries",
            "",
            "SYSTEM LAYOUT:",
            f"  CBS Type: {self.facts.layout_flags.cbs_type or 'Not detected'}",
            f"  Infeed System: {self.facts.layout_flags.has_infeed_system}",
            f"  Auto Induct: {self.facts.layout_flags.has_auto_induct}",
            f"  Manual Induct: {self.facts.layout_flags.has_manual_induct}",
            f"  Barcode Scanning: {self.facts.layout_flags.has_barcode_scanning}",
            "",
            "EXTRACTED COUNTS (with sources):",
        ]
        
        for key, comp_count in self.facts.system_numbers.get_all_counts().items():
            if isinstance(comp_count, ComponentCount):
                if comp_count.count is not None:
                    status = "✓" if comp_count.confirmed else "?"
                    lines.append(
                        f"  [{status}] {key}: {comp_count.count} "
                        f"({comp_count.source.value})"
                    )
        
        unconfirmed = self.facts.get_unconfirmed_fields()
        if unconfirmed:
            lines.append("")
            lines.append("UNCONFIRMED COUNTS (need manual input):")
            for field in unconfirmed:
                lines.append(f"  - {field}")
        
        lines.append("")
        lines.append(f"Total Process Steps: {len(self.facts.process_steps)}")
        
        lines.append("=" * 80)
        return "\n".join(lines)


# ============================================================================
# Top-level convenience function
# ============================================================================

def extract_proposal_facts(
    dxf_path: str,
    costing_path: Optional[str] = None,
    project_name: str = "",
    client_name: str = "",
    verbose: bool = True,
) -> ProposalFacts:
    """
    Extract and populate ProposalFacts from DXF and optional costing file.
    
    This is the main entry point for extracting all proposal data.
    
    Args:
        dxf_path: Path to DXF file
        costing_path: Path to costing Excel file (optional)
        project_name: Project name
        client_name: Client name
        verbose: Print extraction summary
    
    Returns:
        Populated ProposalFacts object
    
    Example:
        facts = extract_proposal_facts(
            "design.dxf",
            costing_path="F24-00276-BOSTA CAIRO_Loop CBS_Rev-11.xlsx",
            project_name="BOSTA Cairo",
            client_name="BOSTA Egypt"
        )
        
        # Use facts in all generators
        cover_letter = generate_cover_letter(facts)
        exec_summary = generate_exec_summary(facts)
        process_flow = generate_process_flow(facts)
        system_desc = generate_system_description(facts)
    """
    extractor = ProposalFactsExtractor(
        dxf_path=dxf_path,
        costing_path=costing_path,
        project_name=project_name,
        client_name=client_name,
    )
    
    facts = extractor.extract()
    
    if verbose:
        print(extractor.log_extraction_summary())
    
    return facts


if __name__ == "__main__":
    # Test extraction
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python proposal_facts_extractor.py <dxf_file> [costing_file]")
        sys.exit(1)
    
    dxf_file = sys.argv[1]
    costing_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    facts = extract_proposal_facts(
        dxf_file,
        costing_path=costing_file,
        project_name="Test Project",
        client_name="Test Client",
        verbose=True
    )
    
    # Save facts to JSON for inspection
    output_path = Path(dxf_file).stem + "_facts.json"
    with open(output_path, "w") as f:
        f.write(facts.to_json())
    print(f"\nFacts saved to {output_path}")

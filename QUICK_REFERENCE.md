"""
ProposalFacts Quick Reference Guide

═══════════════════════════════════════════════════════════════════════════════

WHAT IS ProposalFacts?
A single object containing ALL proposal data (counts, layout, flow, unknowns)
extracted ONCE from DXF + costing, then used by ALL generators.

KEY BENEFIT: No duplicate extraction, consistent counts, no hallucinations.

═══════════════════════════════════════════════════════════════════════════════

QUICK START (3 Steps)
═════════════════════

1) EXTRACT DATA ONCE:

   from proposal_facts_extractor import extract_proposal_facts
   
   facts = extract_proposal_facts(
       dxf_path="design.dxf",
       costing_path="costing.xlsx",
       project_name="Project X",
       client_name="Client Name",
   )

2) USE IN GENERATORS:

   # Get any count with source preference (costing > DXF)
   count, source, confirmed = facts.get_count("gravity_chutes")
   
   if count is None:
       output = f"(to be confirmed during detailed engineering)"
   else:
       output = f"{count} Nos ({source})"

3) PASS TO ALL GENERATORS:

   cover_letter = gen_cover_letter(facts, executives, offer_ref, ...)
   exec_summary = gen_exec_summary(facts)
   process_flow = gen_process_flow(facts)
   system_desc = gen_system_description(facts)

═══════════════════════════════════════════════════════════════════════════════

CORE OBJECTS (What to Use)
══════════════════════════

ProposalFacts (THE MAIN ONE - USE THIS)
├─ system_numbers: All component counts with sources
│  └─ feedlines, gravity_chutes, rejection_chutes, ...
├─ layout_flags: System configuration (yes/no & types)
│  └─ has_auto_induct, has_manual_induct, cbs_type, ...
├─ process_steps: Ordered process flow
│  └─ [Step 1, Step 2, ..., Step N]
├─ unknowns: Missing/unconfirmed fields
│  └─ ["feedlines", "throughput_pph"]

ComponentCount (TRACKS DATA SOURCE)
├─ count: The actual number (or None if missing)
├─ source: Where it came from (COSTING, DXF, MANUAL, MISSING)
├─ confirmed: Whether we're confident
└─ component_name: e.g., "Feedlines"

═══════════════════════════════════════════════════════════════════════════════

COMMON OPERATIONS
═════════════════

GET A SINGLE COUNT:
    count, source, confirmed = facts.get_count("feedlines")
    # Returns: (4, "costing_bom", True) or (None, "to_be_confirmed", False)

GET ALL COUNTS:
    for key, comp_count in facts.system_numbers.get_all_counts().items():
        if comp_count.count is not None:
            print(f"{key}: {comp_count.count} ({comp_count.source.value})")

CHECK SYSTEM LAYOUT:
    if facts.layout_flags.has_auto_induct:
        print(f"CBS Type: {facts.layout_flags.cbs_type}")

LIST UNCONFIRMED FIELDS:
    unconfirmed = facts.get_unconfirmed_fields()
    for field in unconfirmed:
        print(f"⚠️  {field}")

GET PROCESS FLOW:
    for step in facts.process_steps:
        print(f"{step.sequence}. {step.component}: {step.description}")

═══════════════════════════════════════════════════════════════════════════════

FORMATTING OUTPUT (Examples)
════════════════════════════

SAFE WAY (Never hallucinate):
    count, text, source = facts.get_count("feedlines")
    if count:
        output = f"Feedlines: {count} Nos ({source})"
    else:
        output = f"Feedlines: {text}"  # (to be confirmed...)

WHEN COUNT CAN'T BE 0:
    # WRONG:
    if facts.system_numbers.feedlines.count == 0:
        print("0 feedlines")  # ❌

    # CORRECT:
    count, text, source = facts.get_count("feedlines")
    if count is None:
        print("Feedlines: " + text)  # ✓ (to be confirmed...)

BUILDING DESCRIPTIONS:
    parts = []
    for step in facts.process_steps:
        if step.count:
            parts.append(f"{step.count} {step.component}")
        else:
            parts.append(step.component)
    
    flow = " → ".join(parts)
    # Output: "4 Feedlines → 1 Loop CBS → 50 Gravity Chutes"

═══════════════════════════════════════════════════════════════════════════════

COMPONENT COUNTS AVAILABLE
═══════════════════════════

Induction System:
  - feedlines
  - induct_lines_auto
  - manual_induct_stations

Infeed Conveyors:
  - infeed_conveyors_total
  - telescopic_belt_conveyors
  - buffer_conveyors
  - curve_conveyors

Output Chutes:
  - gravity_chutes
  - mini_gravity_chutes
  - collection_chutes
  - rejection_chutes
  - dispersion_chutes
  - bulk_chutes
  - direct_bagging_chutes

Sorter & Scanning:
  - cbs_sorters
  - sorter_carriers
  - barcode_scanners
  - weighing_systems

Additional:
  - vds_loops
  - recirculation_conveyors
  - throughput_pph

═══════════════════════════════════════════════════════════════════════════════

SOURCE PREFERENCE HIERARCHY
══════════════════════════

When getting a count, the system checks:
  1. COSTING data (most accurate) ← PREFERRED
  2. DXF extraction (secondary)
  3. MANUAL input (if provided)
  4. MISSING (unconfirmed)

Example:
  DXF says: 50 gravity chutes
  Costing says: 52 gravity chutes
  facts.get_count("gravity_chutes") → (52, "costing_bom", True) ✓

═══════════════════════════════════════════════════════════════════════════════

STREAMLIT INTEGRATION
════════════════════

from proposal_facts_streamlit_integration import (
    integrate_proposal_facts_upload,
    display_facts_summary,
    get_component_count_with_fallback,
)

# Upload and extract in one function:
facts = integrate_proposal_facts_upload(
    client_name="Client",
    project_name="Project",
    dxf_file_upload=uploaded_dxf,
    costing_file_upload=uploaded_costing,  # Optional
    verbose=True,
)

# Show summary in Streamlit:
if facts:
    display_facts_summary(facts, show_details=True)

# Get count with UI-friendly text:
count, text, source = get_component_count_with_fallback(
    facts, "feedlines", human_name="Induct Lines"
)
st.write(f"Induct Lines: {text}")

═══════════════════════════════════════════════════════════════════════════════

MIGRATION CHECKLIST
═══════════════════

[✓] Phase 1: Create ProposalFacts system
    [✓] proposal_facts.py - Core dataclasses
    [✓] costing_sheet_mapper.py - Excel detection
    [✓] proposal_facts_extractor.py - Unified extraction
    [✓] proposal_facts_streamlit_integration.py - Streamlit helpers
    [✓] PROPOSAL_FACTS_GUIDE.md - Full documentation
    [✓] IMPLEMENTATION_SUMMARY.md - Summary

[ ] Phase 2: Refactor generators
    [ ] call_groq_for_cover_letter(facts, ...)
    [ ] call_groq_exec_summary(facts, ...)
    [ ] call_groq_for_process_flow(facts, ...)
    [ ] generate_system_description(facts, ...)

[ ] Phase 3: Update streamlit.py
    [ ] Import ProposalFacts modules
    [ ] Call extract_proposal_facts() once
    [ ] Pass facts to all generators
    [ ] Remove old extraction logic

[ ] Phase 4: Validation & testing
    [ ] Test with sample DXF + costing files
    [ ] Verify counts match reference proposals
    [ ] Check for hallucinations
    [ ] Measure quality improvement

═══════════════════════════════════════════════════════════════════════════════

CRITICAL RULES (Never Violate)
══════════════════════════════

❌ DON'T:
  - Output "0" for critical counts
  - Call DXF extraction multiple times
  - Guess missing component counts
  - Mix costing and DXF data without preference

✓ DO:
  - Use facts.get_count() for all counts
  - Check if count is None before output
  - Use standard "to be confirmed" phrase
  - Track data sources
  - Flag unconfirmed fields

═══════════════════════════════════════════════════════════════════════════════

FILE LOCATIONS
══════════════

d:\Projects\1. Propsal_Automation\V9\
├─ proposal_facts.py                      [650 lines] Core dataclasses
├─ costing_sheet_mapper.py               [450 lines] Excel detection
├─ proposal_facts_extractor.py           [350 lines] Unified extraction
├─ proposal_facts_streamlit_integration.py [300 lines] Streamlit UI
├─ PROPOSAL_FACTS_GUIDE.md              [900+ lines] Full documentation
└─ IMPLEMENTATION_SUMMARY.md             [400+ lines] Summary

═══════════════════════════════════════════════════════════════════════════════

TROUBLESHOOTING
══════════════

Q: Count is always None?
A: Check facts.get_unconfirmed_fields() - field might be missing from costing/DXF

Q: Costing data not being used?
A: Verify costing file loaded - check facts.costing_metrics
   Ensure CostingSheetMapper found the right sheet

Q: Seeing 0 in output for important component?
A: get_count() should return None, not 0
   Check get_counts_source_of_truth() logic

Q: Need sheet names from costing workbook?
A: mapper = CostingSheetMapper("costing.xlsx")
   print(mapper.get_all_sheet_mappings())

Q: Want to add new component type?
A: 1. Add field to SystemNumbers dataclass
   2. Add to ComponentCount initialization
   3. Update mapping in populate_from_dxf/costing

═══════════════════════════════════════════════════════════════════════════════

NEXT STEPS FOR YOUR TEAM
═════════════════════════

1. Review PROPOSAL_FACTS_GUIDE.md
2. Test with sample DXF + costing file:
   facts = extract_proposal_facts("test.dxf", "test.xlsx")
   print(facts.summary())
3. Pick one generator to refactor first (e.g., call_groq_for_cover_letter)
4. Use examples in proposal_facts_streamlit_integration.py as template
5. Update streamlit.py to call extract_proposal_facts() once
6. Gradually migrate all generators to use ProposalFacts

═══════════════════════════════════════════════════════════════════════════════

QUESTIONS? 
See PROPOSAL_FACTS_GUIDE.md for detailed documentation and examples.
"""

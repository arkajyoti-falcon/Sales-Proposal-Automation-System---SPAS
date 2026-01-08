"""
PROPOSALFS ARCHITECTURE & INTEGRATION GUIDE

# ProposalFacts: Single Source of Truth for Proposal Generation

This document explains how to use the new ProposalFacts system to ensure
consistency, avoid hallucinated data, and improve proposal quality.

## Overview

ProposalFacts is a unified data container that:
1. **Extracts data ONCE** from DXF + optional costing files at proposal startup
2. **Tracks data sources** - every count knows if it came from costing or DXF
3. **Prevents hallucination** - missing data is marked "to be confirmed", not guessed
4. **Enables consistency** - all generators use the SAME data, no duplicates
5. **Provides fallbacks** - prefers costing data over DXF; uses standard phrases for unknowns

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  INPUTS: DXF File + Optional Costing Excel                      │
└────────────┬────────────────────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────────────────────┐
│  proposal_facts_extractor.extract_proposal_facts()              │
│  - Loads DXF using st_sys_desc.extract_dxf_full_json()        │
│  - Loads costing using CostingSheetMapper                       │
│  - Populates ProposalFacts from DXF metrics                     │
│  - OVERRIDES with costing data where available                 │
│  - Builds process_steps from layout_flags                       │
└────────────┬────────────────────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────────────────────┐
│  ProposalFacts Object (SINGLE SOURCE OF TRUTH)                 │
│                                                                  │
│  ├─ Metadata:                                                   │
│  │  ├─ dxf_filename, costing_filename                          │
│  │  ├─ project_name, client_name                               │
│  │  └─ extraction_timestamp                                    │
│  │                                                              │
│  ├─ SystemNumbers (counts with sources):                        │
│  │  ├─ feedlines: ComponentCount(count=4, source=DXF, ✓)      │
│  │  ├─ gravity_chutes: ComponentCount(count=50, source=COSTING, ✓) │
│  │  ├─ throughput_pph: ComponentCount(count=27600, source=COSTING, ✓) │
│  │  └─ ... (all components with sources tracked)              │
│  │                                                              │
│  ├─ LayoutFlags (boolean indicators):                           │
│  │  ├─ has_auto_induct: true                                   │
│  │  ├─ has_manual_induct: false                                │
│  │  ├─ cbs_type: "Loop CBS"                                    │
│  │  └─ ... (all system configuration flags)                   │
│  │                                                              │
│  ├─ ProcessSteps (ordered flow):                                │
│  │  ├─ Step 1: Manual Induct → ...                            │
│  │  ├─ Step 2: Auto Induct Lines → ...                        │
│  │  └─ Step N: Output Chutes → ...                            │
│  │                                                              │
│  └─ Unknowns (missing/unconfirmed fields):                      │
│     └─ ["feedlines", "throughput_pph"]  ← flagged as MISSING  │
│                                                                  │
└────────────┬────────────────────────────────────────────────────┘
             │
    ┌────────┴────────┬─────────────┬──────────────┬─────────────┐
    │                 │             │              │             │
    ▼                 ▼             ▼              ▼             ▼
┌────────────┐  ┌──────────────┐ ┌──────────────┐┌─────────────┐┌───────┐
│   COVER    │  │   EXECUTIVE  │ │   PROCESS    ││   SYSTEM    ││ ...   │
│   LETTER   │  │   SUMMARY    │ │     FLOW     ││DESCRIPTION ││       │
└────────────┘  └──────────────┘ └──────────────┘└─────────────┘└───────┘

ALL generators use:
  count, source, confirmed = facts.get_count("component_key")
  - Never duplicate extraction
  - Consistent count preference (costing > DXF)
  - Clear "to be confirmed" phrases
```

## Core Classes

### 1. ComponentCount (dataclass)
Tracks a single component count with its source and verification status.

```python
@dataclass
class ComponentCount:
    count: Optional[int]  # The actual count (None if unconfirmed)
    source: CountSource   # Where it came from: COSTING, DXF, MANUAL, or MISSING
    confirmed: bool       # Whether we're confident in this count
    component_name: str   # e.g., "Feedlines"

# Usage:
comp = ComponentCount(count=4, source=CountSource.COSTING, confirmed=True)
```

### 2. SystemNumbers (dataclass)
Container for all component counts with their sources.

```python
@dataclass
class SystemNumbers:
    feedlines: ComponentCount
    gravity_chutes: ComponentCount
    mini_gravity_chutes: ComponentCount
    rejection_chutes: ComponentCount
    # ... (all component types)
    
    def get_all_counts(self) -> Dict[str, ComponentCount]:
        """Return all counts"""
    
    def get_unconfirmed(self) -> List[str]:
        """Return list of unconfirmed field names"""
```

### 3. LayoutFlags (dataclass)
Boolean indicators for system configuration.

```python
@dataclass
class LayoutFlags:
    has_infeed_system: bool
    has_auto_induct: bool
    has_manual_induct: bool
    has_vds_loop: bool
    has_barcode_scanning: bool
    cbs_type: Optional[str]  # "Loop CBS" or "Linear CBS"
    throughput_pph: Optional[float]
    # ... (other layout characteristics)
```

### 4. ProcessStep (dataclass)
Single step in the process flow.

```python
@dataclass
class ProcessStep:
    sequence: int         # 1-based order
    component: str        # e.g., "Manual Induct Station"
    description: str      # What happens here
    upstream: Optional[str]   # Previous component
    downstream: Optional[str] # Next component
    count: Optional[int]      # How many of this component
```

### 5. ProposalFacts (dataclass) - THE MAIN OBJECT
Complete proposal data container - this is what all generators use.

```python
@dataclass
class ProposalFacts:
    # Metadata
    dxf_filename: str
    costing_filename: Optional[str]
    extraction_timestamp: str
    project_name: str
    client_name: str
    
    # Core data (single source of truth)
    system_numbers: SystemNumbers        # All component counts with sources
    layout_flags: LayoutFlags            # System configuration flags
    process_steps: List[ProcessStep]     # Ordered process flow
    unknowns: List[str]                  # Unconfirmed fields
    
    # Raw data for fallback access
    dxf_metrics: Dict[str, Any]
    costing_metrics: Dict[str, Any]
    
    # Key methods:
    def get_count(self, component_key: str) -> Tuple[Optional[int], str]:
        """Get count with source; returns (count, source_description)"""
    
    def get_unconfirmed_fields(self) -> List[str]:
        """Get all unconfirmed component counts"""
    
    def to_dict(self) -> Dict:
        """Convert to dictionary for serialization"""
    
    def summary(self) -> str:
        """Get human-readable summary"""
```

## Usage in Streamlit

### Step 1: Accept Files

```python
# In your streamlit.py:
dxf_file = st.file_uploader("DXF Layout", type=["dxf"])
costing_file = st.file_uploader("Costing Sheet", type=["xlsx"], required=False)

client_name = st.text_input("Client Name")
project_name = st.text_input("Project Name")
```

### Step 2: Extract ProposalFacts ONCE

```python
from proposal_facts_streamlit_integration import integrate_proposal_facts_upload

facts = integrate_proposal_facts_upload(
    client_name=client_name,
    project_name=project_name,
    dxf_file_upload=dxf_file,
    costing_file_upload=costing_file,
    verbose=True,  # Show extraction summary
)

if not facts:
    st.stop()  # Stop if extraction failed
```

### Step 3: Display Summary (Optional)

```python
from proposal_facts_streamlit_integration import display_facts_summary

with st.expander("View Extracted Data"):
    display_facts_summary(facts)
```

### Step 4: Use Facts in Generators

```python
# Instead of passing individual metrics, pass facts:

# BEFORE (OLD WAY):
cover_letter = call_groq_for_cover_letter(
    dxf_json,           # Separate DXF data
    system_metrics,     # Separate system data
    client_name,
    # ... 10+ more parameters
)

# AFTER (NEW WAY):
cover_letter = call_groq_for_cover_letter(
    facts,              # Single source of truth
    executives_list,
    offer_reference,
    letter_date,
    # ... only relevant non-facts parameters
)
```

## Generator Refactoring Template

### Before: Using Individual Parameters
```python
def call_groq_for_cover_letter(dxf_json, system_text, client_name, 
                                process_flow_summary, pph_count, ...):
    """OLD: 15+ parameters, repeated extraction"""
    
    # Extract from DXF again (!)
    feedline_count = dxf_json.get('FEEDLINE COUNT', 0)
    gravity_chutes = dxf_json.get('GRAVITY CHUTE COUNT', 0)
    
    # Build summary again (!)
    component_counts = f"Feedlines: {feedline_count}, Chutes: {gravity_chutes}"
    
    # Call AI with scattered data
    response = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": PROMPT},
            {"role": "user", "content": f"Data: {component_counts}"},
        ],
    )
    return response.choices[0].message.content
```

### After: Using ProposalFacts
```python
def call_groq_for_cover_letter(facts: ProposalFacts,
                                executives, offer_reference, letter_date):
    """NEW: Single facts parameter, no extraction duplication"""
    
    # Get counts from facts (with source preference: costing > DXF)
    feedlines, _, _ = facts.get_count("feedlines")
    gravity_chutes, _, _ = facts.get_count("gravity_chutes")
    cbs_type = facts.layout_flags.cbs_type
    
    # Build summary from facts
    component_counts = f"Feedlines: {feedlines}, CBS: {cbs_type}, Chutes: {gravity_chutes}"
    
    # Call AI with facts-based data
    response = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": PROMPT},
            {"role": "user", "content": f"Data: {component_counts}"},
        ],
    )
    return response.choices[0].message.content
```

## Getting Component Counts

### Method 1: Simple Get
```python
# Get feedline count
feedlines, source, confirmed = facts.get_count("feedlines")

# If feedlines is None, count was unconfirmed/missing
if feedlines:
    print(f"Feedlines: {feedlines} ({source})")
else:
    print(f"Feedlines: (to be confirmed during detailed engineering)")
```

### Method 2: With Fallback Helper
```python
from proposal_facts_streamlit_integration import get_component_count_with_fallback

count, display_text, source = get_component_count_with_fallback(
    facts, "gravity_chutes", human_name="Gravity Chutes"
)
# Output: count=50, display_text="50 Nos", source="costing_bom"
```

### Method 3: All Counts At Once
```python
all_counts = facts.system_numbers.get_all_counts()

for component_key, comp_count in all_counts.items():
    if comp_count.count is not None:
        print(f"{component_key}: {comp_count.count} ({comp_count.source.value})")
```

## Data Sources & Preference

The system uses this hierarchy:

```
Preference Hierarchy:
  1. Costing/BOQ data (most accurate)
     └─ Extracted from "Quote Master" sheet
  2. DXF extraction (secondary)
     └─ Component counts from block references
  3. Manual input (least preferred)
     └─ If user provides in costing
  4. Missing/Unconfirmed
     └─ Marked as None with "to be confirmed" phrase
```

### Never Outputs "0" for Critical Counts

If a critical component count is 0 or missing, the system:
1. Sets count to None
2. Marks source as MISSING
3. Uses phrase: "(to be confirmed during detailed engineering)"

**Critical components:**
- feedlines, induct_lines_auto, manual_induct_stations
- gravity_chutes, rejection_chutes, bulk_chutes
- throughput_pph, cbs_sorters

```python
# Example:
facts.system_numbers.gravity_chutes = ComponentCount(
    count=None,  # NOT 0!
    source=CountSource.MISSING,
    confirmed=False
)

# In output:
gravity_chutes_count, display = facts.get_count("gravity_chutes")
# Returns: (None, "(to be confirmed during detailed engineering)")
```

## Validation & No Hallucination

### Check for Hallucinations
```python
from proposal_facts import validate_no_hallucinations

errors = validate_no_hallucinations(facts)
if errors:
    for error in errors:
        print(f"⚠️  {error}")
        # e.g., "gravity_chutes: Negative count not allowed (-5)"
        # e.g., "throughput_pph: Suspiciously large count (999999)"
```

### Flag Unconfirmed Data
```python
unconfirmed = facts.get_unconfirmed_fields()
if unconfirmed:
    print(f"⚠️  {len(unconfirmed)} fields need confirmation:")
    for field in unconfirmed:
        print(f"   - {field}")
    
    # Don't include these in output or mark as TBD
```

## Costing Sheet Mapper

The CostingSheetMapper automatically detects and maps component sheets.

### Supported Sheet Types
```python
COMPONENT_TYPE_PATTERNS = {
    "Loop CBS": [...patterns...],
    "Linear CBS": [...patterns...],
    "Conveyors": [...patterns...],
    "Destinations": [...patterns...],
    "Steelworks": [...patterns...],
    "PTL": [...patterns...],
    "Induct Lines": [...patterns...],
    "Control System": [...patterns...],
    "Safety Equipment": [...patterns...],
}
```

### Extract Sheet Data for AI
```python
from costing_sheet_mapper import CostingSheetMapper

mapper = CostingSheetMapper("costing.xlsx")

# Get sheet name for component type
loop_cbs_sheet = mapper.get_sheet_by_component_type("Loop CBS")

# Extract full sheet as formatted text for AI
sheet_text = mapper.extract_sheet_as_text("Loop CBS")

# Send to AI for fact building
ai_response = groq_client.chat.completions.create(
    model="llama-3.3-70b-versatile",
    messages=[
        {"role": "system", "content": "Extract Loop CBS specifications..."},
        {"role": "user", "content": sheet_text},
    ],
)
```

## Testing & Debugging

### Log Extraction Summary
```python
from proposal_facts_extractor import ProposalFactsExtractor

extractor = ProposalFactsExtractor(
    dxf_path="design.dxf",
    costing_path="costing.xlsx",
    project_name="Test Project",
    client_name="Test Client"
)

facts = extractor.extract()

# Show detailed summary
print(extractor.log_extraction_summary())
```

### Export Facts to JSON
```python
import json

facts_json = facts.to_json()
with open("proposal_facts_dump.json", "w") as f:
    f.write(facts_json)

# Now inspect the JSON file to verify data extraction
```

### Check Individual Components
```python
# Check if component was extracted
if facts.layout_flags.has_auto_induct:
    count, source, confirmed = facts.get_count("feedlines")
    print(f"Feedlines: {count} ({source}, confirmed={confirmed})")
else:
    print("No auto induct system detected")
```

## Migration Path

### Phase 1: Add ProposalFacts (Current)
1. ✅ Create ProposalFacts dataclass
2. ✅ Create CostingSheetMapper
3. ✅ Create ProposalFactsExtractor
4. ✅ Create Streamlit integration module
5. Next: Update generators one by one

### Phase 2: Update Generators
1. Refactor call_groq_for_cover_letter(facts) - FIRST
2. Refactor call_groq_exec_summary(facts)
3. Refactor process flow generation
4. Refactor system description generation

### Phase 3: Validation & Testing
1. Add proposal facts validation in streamlit UI
2. Test with sample DXF + costing files
3. Verify counts against reference proposals
4. Measure quality improvement

### Phase 4: Full Rollout
1. Update main streamlit.py to use ProposalFacts
2. Remove old extraction logic from generators
3. Update documentation
4. Train team on new workflow

## File Organization

```
d:\Projects\1. Propsal_Automation\V9\
├─ proposal_facts.py                      # Core dataclasses
├─ costing_sheet_mapper.py               # Excel sheet detection
├─ proposal_facts_extractor.py           # Unified extraction
├─ proposal_facts_streamlit_integration.py  # Streamlit helpers
├─ PROPOSAL_FACTS_GUIDE.md               # This file
│
└─ streamlit.py (UPDATED)
   ├─ Import ProposalFacts modules
   ├─ Create facts object once at startup
   ├─ Pass facts to all generators
   └─ Display facts summary in UI
```

## API Reference

### extract_proposal_facts()
```python
facts = extract_proposal_facts(
    dxf_path: str,
    costing_path: Optional[str] = None,
    project_name: str = "",
    client_name: str = "",
    verbose: bool = True,
) -> ProposalFacts
```

### integrate_proposal_facts_upload() [Streamlit]
```python
facts = integrate_proposal_facts_upload(
    client_name: str,
    project_name: str,
    dxf_file_upload: object,
    costing_file_upload: Optional[object] = None,
    verbose: bool = True,
) -> Optional[ProposalFacts]
```

### get_counts_source_of_truth()
```python
count, source, confirmed = get_counts_source_of_truth(
    facts: ProposalFacts,
    component_key: str,
) -> Tuple[Optional[int], str, bool]
```

## Common Patterns

### Always Prefer Costing Data
```python
# Correct - Costing data automatically preferred
count, source, confirmed = facts.get_count("gravity_chutes")

if count:
    print(f"Gravity chutes: {count}")
    print(f"(From {source})")
else:
    print("Gravity chutes: to be confirmed during detailed engineering")
```

### Never Output 0 for Critical Counts
```python
# WRONG:
if facts.system_numbers.feedlines.count == 0:
    output = "0 feedlines"  # ❌ Never!

# CORRECT:
count, phrase, source = facts.get_count("feedlines")
if count is None:
    output = "Feedlines: " + phrase  # ✓ "to be confirmed..."
else:
    output = f"Feedlines: {count} Nos"
```

### Build Flow Descriptions
```python
parts = []
for step in facts.process_steps:
    if step.count:
        parts.append(f"{step.count} x {step.component}")
    else:
        parts.append(step.component)

flow_desc = " → ".join(parts)
# Output: "4 x Automatic Induct Lines → 1 x Loop CBS → 180 x Output Chutes"
```

---

**Last Updated:** January 8, 2026
**Version:** 1.0
**Author:** Falcon Autotech AI Systems
"""

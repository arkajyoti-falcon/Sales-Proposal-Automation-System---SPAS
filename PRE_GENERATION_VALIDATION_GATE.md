# Pre-Generation Validation Gate Implementation

## Overview
A pre-generation validation gate has been integrated into the proposal system to enforce non-contradiction and prevent hallucinated counts before document sections are generated.

## Key Components Added

### 1. Helper Functions (streamlit.py, lines 85-103)

#### `log_context_counts(ctx: ProposalContext)`
- **Purpose**: Log normalized counts and their sources for auditability and debugging
- **Logs**: Feedlines, telescopic conveyors, total chutes, and PPH with source attribution
- **Usage**: Called immediately after ProposalContext creation (line 11440)

#### `validate_context_counts(ctx: ProposalContext) -> List[str]`
- **Purpose**: Validate that required counts are present and non-zero; detect unsafe states
- **Returns**: List of gating issues (empty if all counts valid)
- **Checks**:
  - Feedlines = 0 or None → "Feedlines/Induct Lines are not specified; numbers will not be invented."
  - Total chutes = 0 → "Total chutes are not specified; using safe phrasing."
- **Usage**: Called at multiple points before generation to trigger safe phrasing injection

---

## Integration Points

### 2. ProposalContext Creation & Logging (streamlit.py, ~line 11426)
```python
# Log normalized counts and validate pre-generation
try:
    log_context_counts(context)
    issues = validate_context_counts(context)
    if issues:
        st.session_state.pre_gen_issues = issues
        logger.warning(f"Pre-generation validation issues: {issues}")
    else:
        st.session_state.pre_gen_issues = []
except Exception as e:
    logger.warning(f"Context logging/validation failed: {e}")
```
- Stores pre-generation issues in session state for later display in Audit UI

### 3. Cover Letter Generation (streamlit.py, ~line 6485)
```python
gating_issues = validate_context_counts(context) if context and ENABLE_CONTEXT_UNIFICATION else []
user_prompt = (
    counts_block + COVER_LETTER_USER_PROMPT_TEMPLATE.format(...) + (
        "\n\nIf any counts are missing, avoid inventing numbers; use safe phrasing without quantities." 
        if gating_issues else ""
    )
)
```
- Injects safe phrasing instruction into prompt if counts are missing/zero

### 4. Executive Summary Generation (streamlit.py, ~line 6544)
```python
gating_issues = validate_context_counts(context) if context and ENABLE_CONTEXT_UNIFICATION else []
user_content += (
    ...
    "Avoid numeric hallucinations; prefer qualitative descriptions where counts are unspecified."
    if gating_issues else ""
)
```
- Prevents numeric hallucinations by instructing to use qualitative descriptions

### 5. Process Flow Generation (streamlit.py, ~line 2495)
```python
if context:
    safe_dxf_json["counts_block_text"] = context.counts_block_text()
    gating_issues = validate_context_counts(context) if ENABLE_CONTEXT_UNIFICATION else []
    if gating_issues:
        safe_dxf_json["safe_phrasing"] = True
        safe_dxf_json["gating_issues"] = gating_issues
        logger.warning(f"Pre-generation validation gate (Process Flow) triggered: {gating_issues}")
```
- Marks flow as requiring safe phrasing; passes gating issues to agentY for safe generation

### 6. System Description Generation (streamlit.py, ~line 11299)
```python
if context:
    constraints = context.counts_block_text() + "\n\n"
    template_text = constraints + template_text
    gating_issues = validate_context_counts(context) if ENABLE_CONTEXT_UNIFICATION else []
    if gating_issues:
        template_text = template_text + "\n\nIf quantities are not specified, avoid inventing counts; use safe qualitative phrasing."
```
- Prepends safe phrasing guidance to template when counts are missing

---

## Audit UI Enhancement (streamlit.py, ~line 11693)

```python
audit_container = st.expander("Audit", expanded=False)
with audit_container:
    st.write("**Normalized Context Counts:**")
    if ctx:
        st.code(ctx.counts_block_text(), language="text")
    
    # Show pre-generation issues
    pre_gen_issues = st.session_state.get("pre_gen_issues", [])
    if pre_gen_issues:
        st.info("**Pre-Generation Gating Issues:**\n" + "\n".join([f"• {i}" for i in pre_gen_issues]))
    
    if violations:
        st.warning("**Consistency Violations Detected:**")
        for v in violations:
            st.warning(v)
        # Apply corrective pass...
    else:
        st.success("✓ Consistency audit passed: no violations detected")
```

**UI Components**:
1. **Normalized Context Counts**: Shows exact counts and their sources
2. **Pre-Generation Gating Issues**: Lists validation issues from context creation
3. **Consistency Violations**: Shows post-generation mismatches (if any)
4. **Corrective Pass**: Minimally aligns numbers if violations detected

---

## Feature Flags

Three feature flags control the validation gate behavior:

```python
ENABLE_CONTEXT_UNIFICATION = True       # Use unified ProposalContext
ENABLE_CONSISTENCY_AUDIT = True         # Run post-generation audit
ENABLE_ALIAS_MAP = True                 # Enrich DXF with alias hints
```

---

## Flow Diagram

```
1. Extract DXF + Costing → ProposalFacts
                              ↓
2. Build ProposalContext (unify counts)
                              ↓
3. log_context_counts() → stdout/logs
                              ↓
4. validate_context_counts() → gating_issues
                              ↓
5. Store in session_state.pre_gen_issues
                              ↓
6. Inject gating_issues into ALL generator prompts
   (Cover Letter, Exec Summary, Process Flow, System Description)
                              ↓
7. Generators use HARD CONSTRAINTS + safe phrasing if issues present
                              ↓
8. Generate sections
                              ↓
9. audit_section_consistency() → violations (if any)
                              ↓
10. correct_section_numbers() → minimal alignment (if needed)
                              ↓
11. Display Audit UI:
    - Normalized counts
    - Pre-gen issues
    - Violations
    - Pass/Fail summary
```

---

## Safety Guarantees

1. **No Invented Counts**: If feedlines/chutes are missing/zero, safe phrasing is mandated in prompts
2. **Transparent Logging**: Every count normalization and gating decision is logged
3. **Post-Generation Audit**: Sections are scanned for contradictions post-generation
4. **Minimal Corrections**: Corrective pass only adjusts numbers, never rewrite structure
5. **UI Visibility**: All counts, sources, and issues visible in Audit expander

---

## Example Scenarios

### Scenario 1: Feedlines = 0 (Unknown)
- **Gating Issue**: "Feedlines/Induct Lines are not specified; numbers will not be invented."
- **Action**: Prompt added to all generators: "avoid inventing numbers; use safe phrasing"
- **Output**: "the system includes induct lines sized to operational needs" (no count)

### Scenario 2: Total Chutes = 0 (Unknown)
- **Gating Issue**: "Total chutes are not specified; using safe phrasing."
- **Action**: Prompt prepended with safe phrasing guidance
- **Output**: "chutes are provided for each product category" (no count)

### Scenario 3: All Counts Specified
- **Gating Issues**: None
- **Action**: Prompts injected with HARD CONSTRAINTS only; no safe phrasing override
- **Output**: Exact numbers used (e.g., "3 feedlines", "5 gravity chutes")

---

## Testing

To verify the validation gate:

1. **Launch app**: `streamlit run streamlit.py`
2. **Upload DXF** with unknown components (e.g., no induct lines in blueprint)
3. **Check logs** for `Pre-generation validation gate` entries
4. **View Audit expander** to confirm:
   - Normalized counts shown
   - Pre-generation gating issues listed
   - Consistency audit passed/failed
5. **Inspect generated sections** for safe phrasing (no invented numbers)

---

## Notes

- The validation gate is **non-blocking**: It guides prompts but doesn't stop generation
- To make it **blocking** (reject unsafe proposals entirely), add `st.error()` and `st.stop()` after gating check
- Gating issues are **additive**: Multiple issues can be present and displayed together
- Corrective pass runs **only if violations detected**; otherwise original text is retained


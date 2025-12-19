"""
ITERATIVE PROCESS FLOW GENERATOR WITH EVALUATION
==================================================
Multi-step refinement with structural coherence scoring
Fixed: Progressive improvement, early stopping, gibberish detection
"""

import os
import re
import json
import tempfile
import logging
from pathlib import Path
from typing import Dict, List, Tuple
from collections import defaultdict

import streamlit as st
from dotenv import load_dotenv
import torch
from sentence_transformers import SentenceTransformer, util as st_util
from bert_score import score as bert_score

# Import from existing code
from combine_old import (
    extract_dxf_components,
    create_dxf_summary,
    get_pinecone_index,
    query_similar_flows,
    call_groq,
    get_cbs_knowledge,
)
from dxf_extractor import create_dxf_summary_for_embedding

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Load sentence transformer for coherence scoring
@st.cache_resource
def load_sentence_model():
    return SentenceTransformer('all-MiniLM-L6-v2')

sentence_model = load_sentence_model()

# ============================================================================
# CLEANING AND VALIDATION FUNCTIONS
# ============================================================================

def clean_generated_flow(flow: str) -> str:
    """Remove meta-commentary and debugging output from generated flow."""
    # Remove everything after "Note:" or "Changes Made:"
    flow = re.split(r'\n(?:Note:|Changes Made:|The revised flow|DXF Constraints:|However,|By making|To further)', flow)[0]
    
    # Remove leading/trailing whitespace
    flow = flow.strip()
    
    # Remove any "=== REVISED FLOW ===" headers
    flow = re.sub(r'=== REVISED FLOW.*?===\s*', '', flow)
    
    return flow


def validate_flow_quality(flow: str) -> Tuple[bool, str]:
    """
    Validate that the generated flow is not gibberish or corrupted.
    Returns: (is_valid, error_message)
    """
    if not flow or len(flow.strip()) < 100:
        return False, "Flow is too short or empty"
    
    # Check for excessive repetition (gibberish detection)
    words = flow.split()
    if len(words) > 20:
        # Count repeated words in sequence
        max_repeat = 1
        current_repeat = 1
        for i in range(1, len(words)):
            if words[i] == words[i-1]:
                current_repeat += 1
                max_repeat = max(max_repeat, current_repeat)
            else:
                current_repeat = 1
        
        if max_repeat > 5:  # Same word repeated more than 5 times
            return False, f"Detected gibberish: word '{words[i]}' repeated {max_repeat} times"
    
    # Check for presence of key sections
    required_sections = ["Process Flow", "Infeed", "CBS", "Output"]
    found_sections = sum(1 for section in required_sections if section.lower() in flow.lower())
    
    if found_sections < 2:
        return False, "Missing key sections in flow"
    
    # Check for excessive non-alphanumeric characters
    alpha_ratio = sum(c.isalnum() or c.isspace() for c in flow) / len(flow)
    if alpha_ratio < 0.85:
        return False, "Too many special characters (possible corruption)"
    
    return True, ""


# ============================================================================
# EVALUATION FUNCTIONS (keeping existing ones)
# ============================================================================

def split_into_sentences(text: str) -> List[str]:
    """Split text into sentences."""
    sentences = re.split(r'[.!?]+', text)
    return [s.strip() for s in sentences if s.strip()]


def compute_bert_scores(original: str, generated: str) -> Tuple[float, float, float]:
    """Compute BERT precision, recall, F1 scores."""
    if not original.strip() and not generated.strip():
        return 1.0, 1.0, 1.0
    if not original.strip() or not generated.strip():
        return 0.0, 0.0, 0.0
    
    P, R, F1 = bert_score(
        [generated],
        [original],
        lang="en",
        rescale_with_baseline=False,
    )
    return float(P[0]), float(R[0]), float(F1[0])


def compute_structural_coherence(original: str, generated: str) -> float:
    """
    Compute structural coherence score (0-100).
    
    Evaluates:
    - Sentence order similarity
    - Paragraph structure
    - Bullet point usage
    - Transition smoothness
    """
    o_sents = split_into_sentences(original)
    g_sents = split_into_sentences(generated)
    
    if not o_sents or not g_sents:
        return 0.0
    
    # Encode sentences
    o_emb = sentence_model.encode(o_sents, convert_to_tensor=True)
    g_emb = sentence_model.encode(g_sents, convert_to_tensor=True)
    
    # 1. Sentence order similarity
    min_len = min(len(o_sents), len(g_sents))
    if min_len == 0:
        return 0.0
    
    sim_pairs = st_util.cos_sim(o_emb[:min_len], g_emb[:min_len])
    order_sim = float(sim_pairs.diag().mean().item())
    
    # 2. Paragraph structure similarity
    def split_paragraphs(t):
        return [p.strip() for p in re.split(r"\n\s*\n", t) if p.strip()]
    
    o_paras = split_paragraphs(original)
    g_paras = split_paragraphs(generated)
    para_sim = 1.0 - min(1.0, abs(len(o_paras) - len(g_paras)) / max(len(o_paras), 1))
    
    # 3. Bullet point similarity
    def count_bullets(t):
        return sum(
            1 for line in t.split("\n")
            if re.match(r"^\s*[\-\*\•\da-z]+[\.\)]\s+", line.strip())
        )
    
    o_bullets = count_bullets(original)
    g_bullets = count_bullets(generated)
    if max(o_bullets, g_bullets) == 0:
        bullet_sim = 1.0
    else:
        bullet_sim = 1.0 - min(1.0, abs(o_bullets - g_bullets) / max(o_bullets, g_bullets))
    
    structure_sim = para_sim * 0.6 + bullet_sim * 0.4
    
    # 4. Transition smoothness (adjacent sentence similarity)
    if len(g_sents) > 1:
        adj_sims = []
        for i in range(len(g_sents) - 1):
            sims = float(st_util.cos_sim(g_emb[i], g_emb[i + 1]).item())
            adj_sims.append(sims)
        transition_sim = sum(adj_sims) / len(adj_sims)
    else:
        transition_sim = 1.0
    
    # Weighted combination
    coherence = 0.45 * order_sim + 0.35 * structure_sim + 0.20 * transition_sim
    return float(max(0.0, min(1.0, coherence)) * 100.0)


def analyze_style_differences(reference: str, generated: str, dxf_json: dict) -> List[str]:
    """
    Generate actionable style and tone feedback by comparing with reference.
    Focus on what matters: structure, language, flow - NOT numbers.
    """
    feedback = []
    
    ref_lower = reference.lower()
    gen_lower = generated.lower()
    
    # 1. Check section structure
    ref_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', reference, re.MULTILINE)
    gen_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', generated, re.MULTILINE)
    
    if len(ref_sections) != len(gen_sections):
        feedback.append(f"Section count mismatch: Reference has {len(ref_sections)} sections, yours has {len(gen_sections)}. Match the reference structure.")
    
    # 2. Check for bullet point style (a., b., c. vs - or •)
    ref_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', reference, re.MULTILINE))
    gen_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', generated, re.MULTILINE))
    ref_has_dashes = bool(re.search(r'^\s*[-\•]\s+', generated, re.MULTILINE))
    
    if ref_has_letters and not gen_has_letters:
        feedback.append("Use lettered sub-points (a., b., c.) for output chutes, matching the reference style")
    elif not ref_has_letters and ref_has_dashes:
        feedback.append("Remove dash bullets, use plain sub-points like the reference")
    
    # 3. Check for key phrases from reference
    key_phrases = [
        ("dumped in bulk", "Use 'dumped in bulk' for infeed arrival (from reference)"),
        ("ascend to a higher level", "Use 'ascend to a higher level' for vertical movement (from reference)"),
        ("picks and positions", "Use 'picks and positions' for operator action (from reference)"),
        ("efficiently sorts", "Use 'efficiently sorts' for CBS operation (from reference)"),
        ("discharged into", "Use 'discharged into' for chute output (from reference)"),
        ("utilizing the data", "Use 'utilizing the data provided by [Client]' (from reference)"),
    ]
    
    for phrase, suggestion in key_phrases:
        if phrase in ref_lower and phrase not in gen_lower:
            feedback.append(suggestion)
    
    # 4. Check for unwanted technical terms
    technical_terms = ["VDS_BUFFER", "AUTO_INDUCT", "OPERATOR_STATION", "BAG_SYSTEM"]
    found_terms = [term for term in technical_terms if term in generated]
    if found_terms:
        feedback.append(f"Remove technical category names: {', '.join(found_terms)}. Use natural language instead")
    
    # 5. Check tone - formal vs casual
    casual_indicators = ["the system features", "there are", "is designed to"]
    formal_indicators = ["shipments are", "parcels enter", "the operator picks"]
    
    casual_count = sum(1 for phrase in casual_indicators if phrase in gen_lower)
    formal_count = sum(1 for phrase in formal_indicators if phrase in ref_lower)
    
    if formal_count > 3 and casual_count > 2:
        feedback.append("Match reference tone: use active voice ('shipments are', 'parcels enter') instead of passive constructions")
    
    # 6. Check paragraph breaks
    ref_para_count = len(re.split(r'\n\s*\n', reference))
    gen_para_count = len(re.split(r'\n\s*\n', generated))
    
    if abs(ref_para_count - gen_para_count) > 2:
        feedback.append(f"Paragraph structure: Reference has {ref_para_count} paragraphs, yours has {gen_para_count}. Match the reference pacing")
    
    # 7. Check for client name usage
    client_name = dxf_json.get('client', 'the client')
    if client_name and client_name.lower() in ref_lower and client_name.lower() not in gen_lower:
        feedback.append(f"Include client name '{client_name}' in sorting logic description, like the reference")
    
    # 8. Check section order
    if ref_sections and gen_sections:
        # Compare first 3 sections
        for i in range(min(3, len(ref_sections), len(gen_sections))):
            if ref_sections[i].lower().strip() != gen_sections[i].lower().strip():
                feedback.append(f"Section order: Position {i+1} should be '{ref_sections[i]}' (reference has this order)")
                break
    
    return feedback

def generate_ai_feedback(reference: str, generated: str, dxf_json: dict, current_score: float, target_score: float) -> List[str]:
    """
    🔧 ULTRA-CONSERVATIVE: Generate MICRO language improvements only
    - NO structural changes
    - NO content additions/removals
    - ONLY phrasing/word choice adjustments
    """
    
    # Extract current established patterns that MUST NOT CHANGE
    cbs_type_match = re.search(r'(Loop CBS|Linear CBS)', generated)
    established_cbs_type = cbs_type_match.group(1) if cbs_type_match else "CBS"
    
    # Extract component counts that must be preserved
    component_counts = {}
    count_patterns = [
        (r'(\d+)\s+(?:manual\s+)?induct\s+station', 'induct_stations'),
        (r'(\d+)\s+output\s+chute', 'output_chutes'),
        (r'(\d+)\s+(?:live\s+)?chute', 'live_chutes'),
        (r'(\d+)\s+PTL', 'ptl_locations'),
    ]
    for pattern, key in count_patterns:
        match = re.search(pattern, generated, re.IGNORECASE)
        if match:
            component_counts[key] = match.group(1)
    
    current_client = dxf_json.get('client', 'UNKNOWN')
    dxf_cats = dxf_json.get("category_summary", {})
    dxf_components_str = ", ".join([f"{cat} ({count})" for cat, count in dxf_cats.items() if count > 0])
    
    system_prompt = f"""You are an expert at suggesting MICRO language improvements for technical documentation.

## 🎯 YOUR MISSION: Suggest 2-3 TINY Word-Level Changes

**Current Score:** {current_score:.1f}/100
**Target:** {target_score}/100
**Gap:** {target_score - current_score:.1f} points

## 🚫 ABSOLUTE RULES (VIOLATION = FAILURE)

### PROTECTED ELEMENTS (DO NOT TOUCH):
- CBS Type: **{established_cbs_type}** (NEVER change this)
- Component Counts: {json.dumps(component_counts)} (NEVER change these)
- Client Name: **{current_client}** (NEVER change this)
- Section Structure: Keep ALL sections in same order

### WHAT YOU CAN CHANGE (ONLY THESE):
1. ✅ Replace 1-2 words with synonyms from reference
2. ✅ Add transition word ("Upon", "Then", "After")
3. ✅ Adjust verb tense slightly (present → present continuous)
4. ✅ Reorder words in sentence (same meaning)

### WHAT YOU CANNOT CHANGE:
1. ❌ Section names or order
2. ❌ Numbers or counts
3. ❌ Component types or names
4. ❌ CBS type or client name
5. ❌ Any structural elements

## 📝 MICRO-CHANGE EXAMPLES

### ✅ APPROVED (These are the ONLY types of changes allowed):
- "are inducted" → "are subsequently inducted" (synonym)
- "Shipments sort efficiently" → "Upon arrival, shipments sort efficiently" (transition)
- "using data" → "utilizing data" (synonym)
- "The system directs" → "The system then directs" (transition word)

### ❌ FORBIDDEN (These will cause score drops):
- "Loop CBS" → "Linear CBS" (breaks established pattern)
- "32 chutes" → "202 chutes" (changes count)
- Remove any section (breaks structure)
- "efficiently sorts" → "system processes" (changes domain terminology)

## 🔍 VERIFICATION CHECKLIST

Before suggesting ANY change, ask:
1. ✅ Is it ONLY 1-2 words?
2. ✅ Does it preserve all counts?
3. ✅ Does it keep CBS type as {established_cbs_type}?
4. ✅ Does it keep client as {current_client}?
5. ✅ Does it match reference vocabulary?

If ANY answer is NO → DO NOT suggest that change

## 📤 OUTPUT FORMAT

Return ONLY a JSON array with 2-3 micro-suggestions:

[
  "Change 'are inducted' to 'are subsequently inducted' to match reference flow",
  "Add 'Upon arrival' before 'at the induct zone' for smoother transition"
]

**CRITICAL:**
- Each suggestion changes MAX 1-2 words
- NO structural changes
- NO content removal
- NO number changes
- Return ONLY JSON array (no markdown, no code blocks)"""

    user_prompt = f"""## CURRENT FLOW (Score: {current_score:.1f})

```
{generated[:1500]}
```

## REFERENCE FLOW (Target Style)

```
{reference[:1500]}
```

## AVAILABLE COMPONENTS (DXF - DO NOT INVENT OTHERS)
{dxf_components_str}

---

## TASK: Suggest 2-3 MICRO word-level improvements

**CONSTRAINTS:**
- CBS Type = {established_cbs_type} (MUST NOT CHANGE)
- Counts = {json.dumps(component_counts)} (MUST NOT CHANGE)
- Client = {current_client} (MUST NOT TOUCH)

**ALLOWED:**
- Replace 1-2 words with reference vocabulary
- Add transition words
- Minor verb adjustments

**FORBIDDEN:**
- Structural changes
- Section reordering
- Count changes
- Content removal

Return ONLY JSON array of 2-3 suggestions."""

    try:
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        result = call_groq(messages, temp=0.2, max_tok=800)  # Lower temp for consistency
        
        # Clean response
        result_clean = result.strip()
        result_clean = re.sub(r'^```json\s*', '', result_clean)
        result_clean = re.sub(r'^```\s*', '', result_clean)
        result_clean = re.sub(r'\s*```$', '', result_clean)
        result_clean = result_clean.strip()
        
        feedback_list = json.loads(result_clean)
        
        if isinstance(feedback_list, list) and len(feedback_list) > 0:
            # Limit to 3 items max for conservative approach
            return feedback_list[:3]
        else:
            return ["Match reference phrasing more closely"]
            
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse AI feedback: {e}")
        return ["Apply incremental language improvements"]
    except Exception as e:
        logger.error(f"Error generating AI feedback: {e}")
        return ["Focus on minor phrasing adjustments"]


def evaluate_process_flow(generated: str, reference: str, dxf_json: dict, 
                          target_score: float = 85) -> Dict:
    """
    Comprehensive evaluation of generated process flow.
    Now uses AI-generated feedback instead of static rules.
    
    Returns:
        dict with scores and feedback
    """
    # Clean generated flow first
    generated_clean = clean_generated_flow(generated)
    
    evaluation = {
        "structural_coherence": 0.0,
        "bert_precision": 0.0,
        "bert_recall": 0.0,
        "bert_f1": 0.0,
        "component_coverage": 0.0,
        "style_match": 0.0,
        "feedback": [],
    }
    
    # 1. Structural coherence
    evaluation["structural_coherence"] = compute_structural_coherence(reference, generated_clean)
    
    # 2. BERT scores
    P, R, F1 = compute_bert_scores(reference, generated_clean)
    evaluation["bert_precision"] = P * 100
    evaluation["bert_recall"] = R * 100
    evaluation["bert_f1"] = F1 * 100
    
    # 3. Component coverage check (keep this as is - it's data validation)
    dxf_cats = dxf_json.get("category_summary", {})
    generated_lower = generated_clean.lower()
    
    coverage_keywords = {
        "AUTO_INDUCT": ["feedline", "automatic", "auto induct", "automatically induct"],
        "OPERATOR_STATION": ["operator", "manual", "positions", "manually position"],
        "VDS_BUFFER": ["vds", "buffer", "distribution"],
        "CHUTE": ["chute", "output", "discharged"],
        "PTL": ["ptl", "put to light", "put-to-light"],
        "BAG_SYSTEM": ["bag", "bagging", "takeaway"],
        "RECIRCULATION": ["recirculation", "recirculate"],
        "CONVEYOR_INFEED": ["infeed", "conveyor", "loaded onto"],
    }
    
    covered = 0
    total_components = 0
    missing_components = []
    
    for cat, count in dxf_cats.items():
        if count > 0 and cat in coverage_keywords:
            total_components += 1
            keywords = coverage_keywords[cat]
            if any(kw in generated_lower for kw in keywords):
                covered += 1
            else:
                # Only flag if it's a major component
                if count > 5 or cat in ["CHUTE", "CBS_SORTER"]:
                    missing_components.append(f"{cat.replace('_', ' ').title()} ({count} units)")
    
    evaluation["component_coverage"] = (covered / total_components * 100) if total_components > 0 else 100
    
    # Add missing components to feedback if any
    if missing_components:
        evaluation["feedback"].append(f"Missing major components in description: {', '.join(missing_components)}")
    
    # 4. AI-GENERATED STYLE FEEDBACK (REPLACING STATIC RULES)
    # Only generate feedback if score is below target
    if evaluation["structural_coherence"] < target_score:
        ai_feedback = generate_ai_feedback(
            reference, 
            generated_clean, 
            dxf_json,
            evaluation["structural_coherence"],
            target_score
        )
        evaluation["feedback"].extend(ai_feedback)
    
    # 5. Style match score (based on feedback count and structural coherence)
    # If structural coherence is high and no feedback, style match should be high
    if evaluation["structural_coherence"] >= target_score and len(evaluation["feedback"]) == 0:
        evaluation["style_match"] = 100.0
    else:
        # Base style match on structural coherence and feedback count
        base_score = evaluation["structural_coherence"]
        feedback_penalty = len(evaluation["feedback"]) * 10
        evaluation["style_match"] = max(0, min(100, base_score - feedback_penalty))
    
    return evaluation

# ============================================================================
# GENERATION FUNCTIONS (keep your existing ones - no changes needed)
# ============================================================================

# [SKIP - Use your existing generate_initial_flow, generate_second_flow_with_chunks functions]

def generate_initial_flow(client_name: str, dxf_json: dict) -> str:
    """Step 1: Generate initial flow from DXF data only."""
    
    dxf_summary = create_dxf_summary(dxf_json)
    # Get CBS knowledge and actually use it
    cbs_knowledge_text = get_cbs_knowledge()
    if not cbs_knowledge_text:
        cbs_knowledge_text = """
Process Flow – Generalized Falcon CBS Ecosystem 

This description covers three major families of flows: 

Forward / primary parcel sorting – loose shipments or boxes/totes from trucks/FCs. 

Returns sorting & store-order consolidation – carton-based returns broken down into eaches + PTL. 

Bag-level outbound sorting – bags/semi-large shipments sorted again on an outbound CBS. 

Wherever there are variants, the text explicitly calls them out (“in some configurations…”, “alternatively…”). 

Shape 

1. Infeed System – Getting Shipments from Dock/Upstream to Induct 

Inbound source patterns 

Shipments typically enter the CBS ecosystem in one of these ways: 

Existing upstream conveyors: 
Boxes/totes are already on an existing conveyor network, usually lengthwise oriented. From there, they are routed onto highway lines that deliver them in a singulated manner to the CBS induct zone.  

 

Multiple dedicated infeed lines: 
Systems may have multiple infeed lines (e.g., 5 for small and big parcels). Operators load shipments onto their respective infeed belts, which then feed a Volume Distribution System (VDS) loop for balancing across inducts.  

 

Telescopic belt conveyors from dock: 
Parcels are manually loaded on telescopic belt conveyors, often 5 + 1 optional telescopics, which extend into trucks at dock level, then carry parcels up to a mezzanine level. From there they connect to infeed conveyors and aligning conveyors before auto-induct.  

 

Bulk dump from FC / marketplace / bags: 
Shipments from FCs, marketplaces, or bags that are unsealed are dumped in bulk onto infeed lines. These infeed lines may be equipped with telescopic belts and then ascend to a higher level where the flow enters a VDS loop that redistributes load across inducts. 

Balancing and pre-induct buffering (VDS) 

A Volume Distribution System (VDS) loop is often used when there are multiple inducts. 

Shipments circulate in the loop and are picked manually from VDS chutes and fed to induct conveyors. 

In some advanced layouts, Arm VDS technology is used to evenly distribute shipments across inducts with minimal manual bias.  

 

Auto infeed vs simple infeed 

In some systems, the auto infeed lines themselves are configured in a singulator manner, ensuring single-piece spacing and smooth transfer up an inclined conveyor to the induct level.  

 

Shape 

2. Inducts – Getting Shipments onto the CBS Carriers 

There are two main induct philosophies: fully auto-induct and manual-assisted induct, and many systems use both. 

Auto-induct lines 

Auto-inducts receive aligned parcels from upstream conveyors. 

Feed lines automatically place shipments onto the Cross-Belt Sorter loop with precise timing, often using dimension/weight/position information to decide carrier readiness. 

Auto-induct is typically used where high throughput and high consistency are required. 

Manual induct stations 

Operators pick shipments from VDS chutes or infeed lines, align them, and place them on induct conveyors or directly onto the CBS loop. 

Standard rule: barcode must face upwards for reliable scanning downstream. 

In some layouts, mezzanine-level manual induct stations allow operators to collect parcels and place them directly on the loop (useful for handling slower flows or exceptions alongside auto-induct).  

 

Common induct “rules” 

Shipments should be: 

Singulated (one per position) 

Aligned (no crosswise blocking) 

Barcode up (to avoid no-reads) 

If these conditions are not met, the system will experience more rejects and recirculation, which should be captured in the process flow. 

Shape 

3. On-Sorter Scanning and Data Capture 

Once shipments are on the CBS loop (loop or linear), a typical pattern is: 

Shipments pass through a barcode scanning tunnel where: 

Barcodes are read from the top side. 

Volumetric details (L×W×H) are captured for each piece. 

The control system now has: 

Unique ID (AWB/shipment ID or bag ID). 

Volume and sometimes weight. 

Current carrier ID and position. 

Sorting logic (always client- or project-specific) uses this data to determine: 

Which chute/destination the shipment should go to (zone, route, store, bagging lane, etc.). 

In some flows, volume thresholds are used to decide whether a shipment goes to a collection chute vs a direct bagging chute or “large shipment chute”.  

 

Shape 

4. Main Sorter Types – Loop CBS vs Linear CBS 

Loop CBS 

A closed loop of cross-belt carriers. 

Shipments are inducted onto the loop and routed to multiple types of chutes around the loop. 

Often used for primary parcel sorting and double-decker configurations with large numbers of chutes. 

Linear CBS 

A linear cross-belt sorter integrated with upstream and downstream conveyors. 

Typically used when: 

The building geometry is more linear. 

Sorting is done at bag or semi-large shipment level, or for a single long row of chutes (e.g., large shipment chutes, live chutes, collection chutes). 

Common sorter behavior 

The CBS control keeps track of: 

Which shipment is on which carrier. 

When the carrier will reach the target chute. 

When the carrier is over the right chute, the cross-belt runs sideways to discharge the shipment into that chute. 

Shape 

5. Output Chutes – Types and Typical Behavior 

Across projects, several chute archetypes repeat: 

Sliding / live chutes 

Used for live loading into trolleys, bags, or conveyors. 

Often sliding-type chutes integrated with PVC belt conveyors or takeaway belt conveyors (TBCs) for continuous evacuation. 

Collection chutes (roller or gravity) 

Friction roller-based collection chutes that gradually accumulate parcels until an operator clears them.  

 

Gravity chutes and mini gravity chutes (typical implementations use tens of such chutes, with capacities around 100–170 parcels each) for passive accumulation.  

 

Direct bagging chutes 

High-volume destinations use direct bagging chutes, where shipments are collected and bagged directly at the chute. 

These are often L-type chutes in double-decker loops and are treated as high-volume chutes for key lanes. 

Secondary chutes feeding PTL 

Secondary chutes are used where shipments require a second-level sort (e.g., store-wise split). 

Shipments from secondary chutes go to PTL racks, usually with a fixed mapping like “each secondary chute → set of PTL locations”, sometimes double-decker and arranged in L-shape for density. 

Non-sort / bulk / special chutes 

Non-sort chutes are used for flows that are not fully sorted at this stage and need further sortation via PTL or manual layouts (e.g., to pallets).  

 

Bulk output chutes send high-volume flows onto a takeaway conveyor (often TBC) for bulk outbound activity.  

 

Strand chutes or special-purpose chutes may be reserved for specific customer-defined flows (e.g., special products, future logic).  

 

Double-decker chutes 

In high-density systems, chutes are arranged in double-decker fashion, with transferring plates connected to the loop. Upper and lower decks support different flows (e.g., direct bagging vs secondary). 

Rejection / sortfail chutes 

Dedicated rejection or sortfail chutes collect: 

No-reads or barcode failures. 

Dimension/logic mismatches. 

Unsorted or mis-sorted shipments. 

Shape 

6. Bagging, PTL, and Post-Sort Handling 

Direct bagging at chutes 

At direct bagging chutes, operators: 

Collect shipments into bags for specific lanes/routes. 

Once a bag is full, they scan the bag barcode, often at a bagging induct line, and the bag itself becomes a unit that can be sorted again (e.g., on a linear CBS for large shipments). 

PTL-based secondary sorting 

For secondary chutes and non-sort chutes, the second-level logic is often PTL-driven: 

Shipments are moved from chute to PTL rack. 

Operator scans an item (or carton). 

The PTL light at the target bin glows. 

Operator places the item in that bin and confirms via an acknowledgment button. 

PTL racks can be: 

Double-decker, L-shaped arrays with thousands of locations (e.g., 3000 PTL locations, 30 locations per chute).  

 

Used for bagging PTL (creating route bags) or store-order PTL (consolidation by store). 

Bag takeaway conveyor 

After bagging (whether from direct bagging chutes or PTL), bags are loaded onto a bag takeaway conveyor placed below the sorter. 

This conveyor carries bags out of the sorter area to outbound docks or an outbound sorter. 

Shape 

7. Returns Sorting and Store-Order Consolidation 

Returns sorting process 

Returns arrive in cartons: 

Operator places the return carton on an infeed conveyor leading to the CBS feeding point. 

Contents are removed; empty cartons are sent to a trash takeaway conveyor for disposal. 

Returned eaches are placed manually on CBS carriers, barcode up. 

Items pass through the barcode tunnel, volumetrics captured, and the linear CBS sorts each item to its assigned chute. 

Downstream, PTL is used to: 

Scan each item. 

Light the correct bin. 

Place item and confirm via acknowledgment button. 

Rejected or failed scans are diverted to rejection/sortfail chutes for manual correction and potential refeed. 

Store-order consolidation (optional / if applicable) 

Picked cartons for store orders are placed on an idler conveyor. 

Via PTL: 

Operator scans carton barcode. 

PTL light at the correct location turns on. 

Items are placed in the highlighted location. 

After PTL processing, cartons move to the next consolidation zone for packing or dispatch. 

Shape 

8. Bag-Level Outbound Sorting on Linear CBS 

Some sites run a secondary / outbound linear CBS that sorts bags and semi-large shipments coming from the primary sorter and cross-dock: 

Outbound infeed 

Bags from the primary Shipment Sorter plus cross-dock bags and semi-large shipments are loaded onto outbound infeed conveyors. 

Semi-large shipments can be loose but are typically non-conveyable at the primary level, so they are handled here. 

Shipments ascend to a higher level and enter a dump VDS loop, which balances load across outbound inducts.  

 

Outbound inducts 

Operators pick bags from the VDS loop, place and scan them on the induct line. 

Bags/shipments are then automatically inducted onto the linear CBS.  

 

Outbound linear CBS sorting 

The outbound linear CBS sorts bags/semi-large shipments into: 

Low volume collection chutes – typically feeding manual sorting layouts downstream. 

High volume sliding chutes – directly discharge into placed trolleys or similar devices. 

Sortfail chute – for rejected or non-sorted bags/shipments.  

 

Shape 

9. Recirculation and Exception Refeeding 

Recirculation lines 

Some linear CBS systems include a recirculation line that automatically returns sort-fail parcels to the induct area. 

There is usually an integrated manual loading point on this line so that reprocessed boxes/totes from the rejection chute can be fed back into the sorter. 

Exception handling rules 

Parcels sent to rejection or sortfail chutes undergo: 

Barcode correction / relabeling. 

Data correction in WMS/WCS if needed. 

Manual refeed via infeed or recirculation. 

Shape 

10. How to Use These Variants in a Model 

When your model generates a process flow for a CBS-based site, it should pick and combine elements based on the scenario: 

Loose parcels from FC / marketplace → talk about bulk dump onto infeed lines, VDS loop, manual pick to induct, loop CBS, sliding + non-sort chutes, bagging PTL, bag takeaway conveyors.  

 

Boxes/totes on existing conveyors → mention lengthwise loading, highway lines, fully automatic induct, linear CBS, live/collection/rejection chutes, and recirculation line.  

 

Truck-based telescopic docks and mezzanine → emphasize telescopic belts, inclined conveyors, aligning conveyors, auto infeed in singulator mode, and auto-induct to loop CBS. 

High volume direct bagging site → highlight direct bagging chutes, L-type double-decker chutes, bagging + PTL, bag takeaway conveyor, and possibly a secondary outbound linear CBS for bags/semi-large shipments. 

Returns site → use the carton-based returns flow, trash takeaway for empties, eaches on sorter, PTL bins, rejection/sortfail chutes, and optional store order consolidation via PTL. 


DXF_COMPONENT_NAME_MAP = {
    # -----------------------
    # PTL SYSTEM
    # -----------------------
    "PTL_Rack": "Put-to-Light (PTL) System (PTL rack)",
    "PTL frame": "Put-to-Light (PTL) System (PTL rack frame)",
    "Ptl racks 4x3 and 4x3": "Put-to-Light (PTL) System (multi-level PTL racks)",
    "3x4 PTL Frame (FAL_P006V01) T-1199": "Put-to-Light (PTL) System (3x4 PTL frame)",
    "3x3 PTL Frame (FAL_P005V01) T-1199": "Put-to-Light (PTL) System (3x3 PTL frame)",
    "PTL_Gen2": "Put-to-Light (PTL) System (PTL display/device)",
    "PTL Chute (T-1266)": "Put-to-Light (PTL) System (PTL chute)",
    "PTL Chute (FAL_C010V01)": "Put-to-Light (PTL) System (PTL chute)",

    # -----------------------
    # SCANNER / DWS TUNNEL
    # -----------------------
    "ECDS (FAL_S012V01)": "Scanner & Dimensioning Tunnel (DWS)",
    "ECDS (FAL_S012V02)": "Scanner & Dimensioning Tunnel (DWS)",
    "ATR Vipacsystem +": "Barcode Scanner Tunnel",
    "Static_ATR Vipacsystem +_001": "Barcode Scanner Tunnel",
    "Static_Weighing Conveyor (FAL_F015V02)_01": "Weighing Conveyor (DWS)",
    "Static_Static-Weighing Conveyor (FAL_F015V01)_008": "Weighing Conveyor (DWS)",
    "WEIGHING_CONVEYOR_3000PPH": "Weighing Conveyor (DWS)",

    # -----------------------
    # AUTO INDUCT LINE (IFU / POSITION / MERGE)
    # -----------------------
    "Static_IFU Conveyor (FAL_F011V02)_01": "Auto Induct IFU Conveyor",
    "Static_XL IFU Conveyor(FAL_F029V01)_01": "Auto Induct IFU Conveyor (XL)",
    "Static_Positioning System (FAL_F013V02)_01": "Induct Positioning System",
    "Static_Positioning System (FAL_F013V01)_1": "Induct Positioning System",
    "FAL_S011V01 (Positioning System)": "Induct Positioning System",
    "Static_Intelligent Merge 30 Deg (FAL_F012V02)_01": "Intelligent Merge Conveyor (Induct to sorter)",
    "Static_Intelligent Merge 30 Deg (FAL_F007V01)_1": "Intelligent Merge Conveyor (Induct to sorter)",
    "Static_Intelligent Merge 60Deg (FAL_F002V01)_1": "Intelligent Merge Conveyor (Induct to sorter)",
    "Static_XL Position Detection System(FAL_F032V01)_01": "Position Detection / Sizing System (Induct)",
    "FAL_S010V02 (Centering System)": "Centering / Aligning Conveyor",
    "FAL_PRC8V01_aligner": "Aligning Conveyor on Infeed/Induct",
    "FAL_PC4V01 (S3_MDR Driven PVC Belt Conveyor)": "Powered Belt Conveyor (MDR driven, induct section)",

    # -----------------------
    # INFEED / RECEIVING / BUFFER
    # -----------------------
    "Static_Infeed Or Orientation Conveyor (FAL_F014V02)_01": "Infeed / Orientation Conveyor",
    "Static_Receiving Conveyor (FAL_F003V01)_1": "Receiving Conveyor (Infeed)",
    "Static_Buffer Conveyor (FAL_F001V01)_1": "Buffer Conveyor (Infeed)",
    "Static_Buffer Conveyor (FAL_F001V01)_2": "Buffer Conveyor (Infeed)",
    "Static_Buffer Conveyor (FAL_F001V01)_3": "Buffer Conveyor (Infeed)",
    "Receiving Conveyor": "Receiving Conveyor (Infeed)",
    "LOADING_CONVEYOR_2000PPH": "Manual Infeed / Loading Conveyor (Each loading)",
    "Mannual Inbound Conveyor": "Manual Inbound Conveyor",

    # -----------------------
    # TELESCOPIC INFEED
    # -----------------------
    "FAL_BLK_Boom Conveyor": "Telescopic Belt Conveyor",
    "FAl_BLK_Boom Conveyors": "Telescopic Belt Conveyor",
    "Telescopico a nastro_ingresso": "Telescopic Belt Conveyor",

    # -----------------------
    # FEEDLINES / HIGHWAYS
    # -----------------------
    "FAL_BLK_Feed lineW800 @30°": "Feedline / Highway Conveyor",
    "FAL_BLK_Feed lineW1000 @30°": "Feedline / Highway Conveyor",
    "Feedline_2.4k": "Feedline / Highway Conveyor",
    "F1_Feedline_Conveyor_Block_V1": "Feedline / Highway Conveyor",
    "3k Dap Feedline": "Feedline / Highway Conveyor",

    # -----------------------
    # CBS SORTER (LOOP / LINEAR)
    # -----------------------
    "Dual CBS Carrier (FAL_S007V01)": "Cross-Belt Sorter Carrier (CBS)",
    "Single Belt Carrier (FAL_S003V01)": "Cross-Belt Sorter Carrier (Single-belt CBS)",
    "Dual Belt CBS Straight Section (FAL_S006V02)": "CBS Straight Section",
    "Dual Belt CBS Straight Section (FAL_S006V01)": "CBS Straight Section",
    "Dual Belt CBS 45 Deg Turn (FAL_S005V02)": "CBS 45° Curve Section",
    "Dual Belt CBS 45 Deg Turn (FAL_S005V01)": "CBS 45° Curve Section",
    "SIngle Belt CBS Straight Section (FAL_S002V02)": "CBS Straight Section (Single belt)",
    "Single Belt CBS 30 Deg Turn (FAL_S001V02)": "CBS 30° Curve Section",
    "CBS_Legs_SingleDecker(FAL_S026V02)": "CBS Support Legs (Single deck)",
    "CBS_Legs_DoubleDecker(FAL_S027V01)": "CBS Support Legs (Double deck)",
    "Dual Belt CBS Leg Structure (FAL_S008V01)": "CBS Support Structure",

    # -----------------------
    # VDS LOOP
    # -----------------------
    "VDS Chute (T2242)": "VDS Collection Chute",
    "FAL_S013V01 (VDS Arm)": "VDS Distribution Arm",

    # -----------------------
    # GENERIC / IRREGULAR CHUTES
    # -----------------------
    "Chute-01": "Generic Sortation Chute",
    "Chute-001": "Generic Sortation Chute",
    "Chute-002": "Generic Sortation Chute",
    "Chute-02": "Generic Sortation Chute",
    "Chute_1": "Generic Sortation Chute",
    "Roller Cage Chute (T-1241)": "Roller Cage Collection Chute",
    "Fix Bin Chute": "Fixed Bin Chute",
    "big parcel chute 2000 mm wide": "Big-parcel Gravity Chute",
    "Chute for live parcel": "Live Parcel Chute",
    "Chute for Ob Live on TBC": "Live Dock Chute (to TBC)",
    "Chute of OB Live Parcels 01": "Live Dock Chute",
    "Chute for LR Pickup Parcels": "Parcel Pickup Chute",
    "IRChute-01": "Irregular Parcel Chute",
    "IRChute-02": "Irregular Parcel Chute",
    "Irregular Parcel Chute-01": "Irregular Parcel Chute",
    "Irregular Parcel Chute-001": "Irregular Parcel Chute",
    "Irregular Parcel Chute-0111": "Irregular Parcel Chute",
    "Irregular chute with 30 Deg angle": "Irregular Parcel Chute (angled)",
    "Irregular Chutes re Shift": "Irregular Parcel Chute Group",

    # -----------------------
    # DIRECT BAGGING CHUTES
    # -----------------------
    "Direct Bagging Chute (FAL_C006V01)": "Direct Bagging Chute",
    "Direct Bagging Churtes 1200 mm pitch": "Direct Bagging Chute (1200mm pitch)",
    "Direct Bagging Chute 1.1 m pitch": "Direct Bagging Chute (1.1m pitch)",
    "Direct Bagging Chute 900 mm pitch": "Direct Bagging Chute (900mm pitch)",
    "Db Chute Updated": "Direct Bagging Chute (double deck)",

    # -----------------------
    # COLLECTION / SECONDARY CHUTES
    # -----------------------
    "Collection Chute 125": "Collection Chute",
    "Collection Chute (T-1538)": "Collection Chute",
    "Collection Chute (FAL_C008V01)": "Collection Chute",
    "Collection Chute for Double Deck Straight": "Double-deck Collection Chute",
    "Collection Type Chute": "Collection Chute (generic)",

    # -----------------------
    # MINI GRAVITY / MINI OUTPUT
    # -----------------------
    "Chutes$0$Mini gravity 01": "Mini Gravity Chute",
    "Chutes$0$Mini Gravity Chute": "Mini Gravity Chute Group",

    # -----------------------
    # BULK OUTPUT
    # -----------------------
    "Bulk Chute": "Bulk Output Chute",

    # -----------------------
    # REJECTION / OW / EXCEPTIONS
    # -----------------------
    "Over Weight Chute and Over Dim Chute": "Overweight / Oversize Rejection Chute",
    "Overweight Chute": "Overweight Rejection Chute",
    "Overweight_03": "Overweight Rejection Chute",
    "Rejection Spiral Chute": "Rejection Spiral Chute",
    "Empty Chute": "Empty Carton / Tote Chute",
    "Empty Tote Chute": "Empty Tote Chute",

    # (Shadowfax outbound sortfail is mostly layer-based; block naming can be added here if needed.)

    # -----------------------
    # LIVE DOCK CHUTES (MONDIAL / AMAZON STYLE)
    # -----------------------
    "Ob Live Chute for Outbound Docks": "Live Dock Chute (Outbound docks)",
    "Outbound Live Chutes": "Live Dock Chute Group",
    "Live Dock Chute Assembly": "Live Dock Chute Assembly",
    "Live Chute Connected with TBC": "Live Dock Chute (connected to TBC)",

    # -----------------------
    # SPIRAL CHUTES
    # -----------------------
    "Spiral Chute Double Decker": "Spiral Chute (Double deck)",
    "Spiral Chute Type -D": "Spiral Chute",

    # -----------------------
    # TRANSFER PLATES / RECIRCULATION
    # -----------------------
    "Transfer Plate for Rec line and manual line": "Recirculation / Manual Transfer Plate",
    "FeedLineTransferPlate": "Feedline Transfer Plate",
    "FeedLineTransferPlate with 600 mm": "Feedline Transfer Plate (600mm)",
    "Lower Deck Live Dock Transfer Plate": "Live Dock Transfer Plate (lower deck)",

    # -----------------------
    # BAG TAKEAWAY / UNDER-CBS CONVEYORS
    # -----------------------
    "Filled+Empty tote Conveyor Below CBS_on Ground$0$_ACMFILLEDHALF": "Filled/Empty Tote Conveyor Below CBS",
    "FAL_FS002V02 1000mmW": "Powered Belt Conveyor (1000mm wide)",
    "FAL_FS003V01": "Powered Belt Conveyor",
    "FAL_PC1V02": "Powered Belt Conveyor",
    "FAL_PC2V02": "Powered Belt Conveyor",
    "FAL_PC5V02": "Powered Belt Conveyor",
    "FAL_PMC7V01(1000mm_60_deg)": "Powered Modular Conveyor (60°)",
    "FAL_PMC9V01(1000mm_30_deg)": "Powered Modular Conveyor (30°)",

    # -----------------------
    # OPERATORS / PANELS
    # -----------------------
    "Operator": "Operator Workstation",
    "Operators and Panel": "Operator Workstation with Control Panel",
    "Operator_Panel_with_HHS": "Operator Panel with Handheld Scanner",

    # -----------------------
    # STORAGE / LOAD UNITS
    # -----------------------
    "pallet": "Pallet",
    "FAL_BLK_DET_Pallet": "Pallet",
    "FAL_DET_BLK_Pallet": "Pallet",
    "FAL_BLK_Roller Cage 1300x600": "Roller Cage (1300x600)",
    "Roller Cage 1000 x 1200": "Roller Cage (1000x1200)",
    "Roller Cage 1000x1200": "Roller Cage (1000x1200)",
    "Trolley 1000x1000": "Parcel / Bag Trolley (1000x1000)",
    "Trolley": "Parcel / Bag Trolley",
    "Trolly": "Parcel / Bag Trolley",
}

 
"""
    system_prompt  = """## ROLE
You are a senior solution engineer writing the "Process Flow of the System" section for Cross-Belt Sorter (CBS) proposals. Your output must **exactly match** the style, structure, and content depth of professional CBS proposals while strictly adhering to provided data.

---
**ABOUT CROSS BELT SORTERS (CBS):** {cbs_knowledge_text}
---

## 🚨 CRITICAL FORMATTING RULES (NON-NEGOTIABLE)

### Rule 1: NO SECTION NUMBERING
- ❌ WRONG: `1. Infeed System:` `2. Auto Induct Line:`
- ✅ CORRECT: `Infeed System:` `Auto Induct Line:`
- **This is the #1 cause of low structural_coherence scores**

### Rule 2: NO SUB-POINT PARENT NUMBERING  
- ❌ WRONG: `5. a. Live Chutes` `6. b. Collection Chutes`
- ✅ CORRECT: `a. Live Chutes` `b. Collection Chutes`

### Rule 3: Start with "Process Flow" Header Only
- ✅ CORRECT: `Process Flow` (standalone line)
- ❌ WRONG: `Process flow of the Loop CBS System-`
- ❌ WRONG: `1. Process Flow`

### Output Structure Template:
```
Process Flow
Infeed System: - <Description>

Auto Induct Line: <Description>

Loop CBS: - <Description>

Output Chutes: - <Description>
a. <Type> - <Description>
b. <Type> - <Description>

<Conditional Section>: <Description>
```

---

## 📊 DATA PRIORITY HIERARCHY

### Priority 1: METADATA (Highest)
If metadata explicitly states something, use that **exact wording**:
- "existing conveyor" → use "existing conveyor"
- "lengthwise orientation" → use "lengthwise orientation"
- "based on dimensions and weight" → use "based on their dimensions and weight"
- Client name "Amazon" → use "using data provided by Amazon"
- "Falcon's fully automatic induct line" → use this exact phrase

### Priority 2: DXF GUIDANCE
Follow the DXF analysis guidance section provided.

### Priority 3: Generic Functional Descriptions
When data is sparse, use simple, generic descriptions based on system type.

### Priority 4: NEVER INVENT
DO NOT add details not supported by metadata or DXF guidance.

---

## 🚫 CRITICAL "DO NOT" RULES

### DO NOT Invent Scanner Details
- ❌ "Barcode scanners are positioned along the infeed line"
- ❌ "Top-side barcode scanner reads each parcel's barcode"
- ❌ "Side barcode scanners are mounted on the infeed conveyors"
- ✅ ONLY mention scanners if metadata explicitly describes them

### DO NOT Expose CAD Codes
- ❌ "Falcon FS002V02 auto-induct units"
- ❌ "fal_fs002v02 and feedline transfer plates"
- ❌ "supported by transfer plates"
- ✅ Use generic: "auto induct lines", "automatically inducted"

### DO NOT Add Unnecessary Technical Details
- ❌ "ensuring accurate identification before induction"
- ❌ "without operator intervention"
- ❌ "rapidly routes each shipment"
- ✅ Keep it simple and direct like actual examples

### DO NOT Invent Metadata Terms
- ❌ "existing conveyor" (unless metadata says this)
- ❌ "highway line" (unless metadata says this)
- ❌ "lengthwise orientation" (unless metadata says this)

### DO NOT Describe What Happens Inside Equipment
- ❌ "The parcels are lifted via an inclined conveyor"
- ✅ "The shipments ascend to a higher level via an inclined conveyor"

---

## 🎯 LANGUAGE & STYLE MATCHING

### Use These Sentence Patterns:

**Arrivals:**
- "Boxes and totes are loaded onto..."
- "Shipments are placed on..."
- "Bags containing shipments are unsealed and dumped..."
- "Shipments from [source] are dumped in bulk onto..."

**Movement:**
- "The shipments ascend to a higher level..."
- "travel from lower level to Mezzanine level"
- "From there, they are directed to..."

**Transitions:**
- "Once the [items] are [state], they [action]..."
- "Upon arrival at the induct zone..."
- "After the shipments are collected..."

**Operations:**
- "The operator picks and positions each shipment..."
- "Feedlines automatically induct the parcels..."
- "efficiently sorts the shipments into their respective output chutes"
- "by utilizing the data provided by [Client]'s sorting logic"

**Counts:**
- "there are [X] chutes present in the system"
- "A total of [X] chutes are designed to..."
- "Within the system, there are a total of [X]..."

### Tone Characteristics:
- **Direct and factual**, not flowery
- **Active voice** preferred
- **Present tense**
- **Specific over generic** when data available
- **Natural flow** with transitions

---

## ✅ QUALITY ASSURANCE CHECKLIST

### Before Generating Output, Verify:

**Structure (structural_coherence):**
- [ ] NO section numbering (`1.`, `2.`, `3.`)
- [ ] NO sub-point parent numbers (`5. a.`, `6. b.`)
- [ ] Starts with "Process Flow" header only
- [ ] Proper blank lines between sections
- [ ] Lettered sub-points only in Output Chutes section

**Content Accuracy (content_coverage & numeric_accuracy):**
- [ ] Used exact client name from metadata
- [ ] Included all sections specified in DXF guidance
- [ ] Used exact chute counts from DXF/metadata
- [ ] NO invented scanner details at infeed
- [ ] NO exposed CAD codes (FAL_FS002V02, etc.)
- [ ] Included VDS if guidance specified
- [ ] Used metadata terminology exactly where provided

**Language Quality (semantic_similarity & domain_tone):**
- [ ] Matches sentence patterns from actual examples
- [ ] Natural transitions between sections
- [ ] Professional but accessible tone
- [ ] No unnecessary technical elaboration
- [ ] Active voice predominant

---

## 🎯 FINAL REMINDER

**Your output will be evaluated on:**
- **semantic_similarity**: Match actual language patterns and phrasing
- **content_coverage**: Include all relevant details from metadata/DXF, no invented content
- **structural_coherence**: Perfect formatting (no section numbers, proper sub-points)
- **numeric_accuracy**: Exact counts from DXF/metadata
- **domain_tone**: Professional, direct style matching actual examples

**Keys to Success:**
1. **NO section numbering** - this alone will boost structural_coherence by 20+ points
2. **NO invented scanners** - boosts content_coverage and numeric_accuracy
3. **Use metadata verbatim** - boosts semantic_similarity
4. **Match sentence patterns** - boosts semantic_similarity and domain_tone
5. **Follow DXF guidance exactly** - boosts content_coverage

**Example 1**
Infeed System: Boxes and totes are loaded onto the existing conveyor in a lengthwise orientation. From there, they are directed to their assigned highway line, which transports them to the CBS induct zone in a singulated manner.
Inducts: Upon arrival at the induct zone, Falcon's fully automatic induct line accurately and smoothly inducts the parcels onto the Linear CBS, based on their dimensions and weight.
Linear CBS: Once the parcels enter the main Linear CBS, the Cross-Belt Sorter (CBS) capture the barcode details & volume data after which it efficiently sorts the boxes and totes into their designated output chutes using data provided by Amazon.
Output Chutes: The Totes/Boxes are discharged into below output chutes.
a. Live Chutes - There are 9 sliding-type live chutes within the Linear CBS system, integrated with PVC belt conveyors and TBCs for live loading.
b. Collection chute – A total of 20 friction roller-based chutes are designed to collect and gradually accumulate the parcels.
c. Rejection Chute- One friction roller-based chute handles rejected shipments.
Recirculation Line: A recirculation line is available to automatically feed sortfail parcels back into the Linear CBS. It is also integrated with a manual loading point for reprocessed boxes and totes collected from the rejection chute.

**Example 2**
Infeed System: - Shipments from FC and Market place are dumped in bulk onto the infeed
lines . The shipments ascend to a higher level and arrive at the VDS loop system. Once the
shipments are within the VDS loop, they are picked manually and are fed among all inducts.
Inducts : - The operator picks and positions each shipment on the induct line, ensuring that
the shipment is properly aligned and that its barcode is facing upwards. The feedlines then
automatically induct the shipments onto the Cross -Belt Sorter Loop.
Loop CBS: - Once the shipments have entered the main loop, the Cross -Belt Sorter (CBS)
efficiently sorts the shipments into their respective output chutes by utilizing the data
provided by Noon 's sorting logic.
Output Chutes : - The shipments are discharged into two types of chutes.
a. Sliding Chutes - Within the loop CBS system, there are a total of 50 Sliding chutes for
each zone . The Shipments collected in Roller Cage trolleys, then they are
consolidated into bags using bagging type PTL racks .
b. Non -Sort Chutes - Within the loop CBS system, there are a total of 13 Non-Sort
Chutes per zone . Shipments collected within these chutes further undergo sortation
via PTL setup into Pallets.
c. Rejection Chute s- Two Rejection Chutes per zone are present to handle rejected
Shipments .
Bag Takeaway Conveyor: - Following the Secondary Sorting process of bagging type PTL , the
shipments are placed into bags and then manually loaded onto a bag takeaway conveyor
located beneath the CBS loop. This conveyor transports the bags out of shipment sorter area
to outbound docks.


**Example 3**
Infeed System: - Boxes and totes are loaded onto the existing conveyor in a lengthwise orientation. From there, they are directed to their assigned highway line, which transports them to the CBS induct zone in a singulated manner.
Auto Induct Line: - Upon arrival at the induct zone, the fully automatic induct line smoothly inducts the parcels onto the Loop Cross Belt Sorter, based on their dimensions and weight.
Manual Induct Station: - Operators pick and position each shipment on the induct line, aligning the parcels for induction. The feedlines then automatically induct the shipments onto the Loop Cross Belt Sorter.
Loop CBS: - Once the shipments have entered the main Loop CBS, the Cross‑Belt Sorter efficiently sorts the shipments into their respective output chutes by utilizing the data provided by Bosta's sorting logic.
Output Chutes: - The sorted shipments are discharged into output chutes.
Generic Chutes – A total of 51 chutes are provided to collect the parcels after sorting.
Put To Light System: - In the system there are 10 PTL locations. Each secondary chute is linked to these PTL locations. The PTL racks are placed in an L‑Shape arrangement.
Bag Takeaway Conveyor: - Following the sorting process, the shipments are placed into bags and then manually loaded onto a bag takeaway conveyor located beneath the CBS loop. This conveyor transports the bags out of the shipment sorter area to the outbound docks.
 

---

**ONLY RETURN A CLEAN PROCESS FLOW, NO EXPLANATION OR EXTRA TEXT IS NEEDED**
**Generate output that a human expert would write, not a template filler.**"""
    
    user_prompt = f"""Generate process flow for:

CLIENT: {client_name}
CBS TYPE: {dxf_json['cbs_type']}
INDUCTION: {dxf_json['induction_type']}

COMPONENTS:
{dxf_summary}

Output ONLY the process flow text. No notes or explanations."""
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    result = call_groq(messages, temp=0.2, max_tok=2000)
    return clean_generated_flow(result)


def generate_second_flow_with_chunks(
    initial_flow: str, 
    dxf_json: dict, 
    reference_chunks: List[Dict],
    client_name: str
) -> str:
    """
    Step 2: Refine with reference chunks, add missing components.
    """
    
    dxf_cats = dxf_json.get("category_summary", {})
    
    # Build reference context
    ref_context = ""
    for i, ref in enumerate(reference_chunks[:2], 1):
        ref_context += f"\n=== REFERENCE {i}: {ref['client']} ===\n"
        ref_context += ref["process_flow"][:1200] + "...\n"
    
    system_prompt ="""You are a senior proposal engineer refining a CBS Process Flow section to match the quality and style of actual winning proposals.

# CRITICAL SUCCESS FACTORS

## 1. 🎯 LANGUAGE TRANSFORMATION (Highest Priority)

Your PRIMARY goal is to transform technical/robotic language into natural, professional proposal language.

### âŒ WRONG (Technical/Robotic):
- "The Loop CBS system combines auto and manual induction with 24 AUTO_INDUCT units"
- "VDS/Buffer area consists of 2 VDS_BUFFER units"
- "conveyor consists of 104 BAG_SYSTEM units"
- "Shipments are loaded onto the infeed conveyor"

### âœ… CORRECT (Natural/Professional):
- "Operators manually position shipments on the induct line"
- "Once within the loop, shipments are manually picked and fed into the inducts"
- "After secondary sorting and bagging, shipments are placed in bags"
- "Shipments from FC and marketplace are dumped in bulk onto the infeed lines"

### Language Transformation Rules:
1. **NEVER expose category names** (VDS_BUFFER, AUTO_INDUCT, BAG_SYSTEM, etc.)
2. **Use descriptive verbs**: "dumped in bulk", "manually picked", "ascend to", "discharged into"
3. **Focus on the shipment journey**, not equipment specifications
4. **Use natural transitions**: "Once within...", "After...", "From there..."
5. **Prefer active descriptions**: "Operators position shipments" vs "Shipments are positioned"

---

## 2. 📖 NARRATIVE FLOW (Tell the Story)

Each section should connect to create a shipment's journey through the system.

### Story Arc Template:
```
Arrival → Distribution → Preparation → Induction → Sorting → Collection → Dispatch
```

### Connection Phrases to Use:
- "Once the shipments..." / "Once within..."
- "After [action], they..."
- "From there, they..."
- "The shipments then..."
- "Following [process]..."

### Example of Good Flow:
"Shipments from FC and marketplace are dumped in bulk onto the infeed lines, where they **ascend to a higher level** and **enter the VDS loop system**. **Once within the loop**, shipments are manually picked and fed into the inducts."

---

## 3. 🔢 COMPONENT INTEGRATION RULES

### Rule A: Add Missing Sections
IF component exists in BOTH:
- DXF data (count > 0)
- Reference flows (described/mentioned)

AND component is missing from initial flow
→ ADD it using reference language

### Rule B: Component Count Usage
- **DXF counts are sacred** - use them exactly as provided
- Reference counts are for language/style ONLY
- When adding sections, extract the count from DXF, not references
- 

### Rule C: Translation of Technical Categories
When you see these in DXF, translate naturally:

| DXF Category | Natural Language Options |
|--------------|-------------------------|
| VDS_BUFFER | "VDS loop system", "distribution loop", "buffer system" |
| AUTO_INDUCT | "automatically induct", "feedlines" |
| OPERATOR_STATION | "manually position", "operators pick and place", "manual induct" |
| CHUTE (by type) | "sliding chutes", "collection chutes", "rejection chutes" |
| PTL | "PTL racks", "PTL setup", "Put-to-Light system" |
| BAG_SYSTEM | "bags", "bagging process", "roller cage trolleys" |
| COLLECTION | "trolleys", "pallets", "collection points" |

---

## 4. 🎨 STYLE MATCHING FROM REFERENCES

Extract and apply these elements from reference flows:

### A. Sentence Structures
Study how references construct sentences:
- "Shipments from [source] are dumped in bulk onto..."
- "The operator picks and positions each shipment on..."
- "Once the shipments have entered..., [equipment] efficiently sorts..."
- "Within the system, there are [X] chutes..."

### B. Technical Detail Level
- References balance detail with readability
- Count mentions: "50 chutes per zone", "13 chutes per zone"
- BUT avoid: "consisting of X units of Y type"

### C. Section Structure
From references, maintain:
- Section numbering IF references use it
- Colon after section title
- Sub-points (a., b., c.) for chute types
- Descriptive details after the count

---

## 5. 🚫 CRITICAL "DO NOT" RULES

### NEVER Do These:
1. âŒ Expose category names: "VDS_BUFFER units", "AUTO_INDUCT units"
2. âŒ Use technical equipment specs: "consists of X BAG_SYSTEM units"
3. âŒ Write disconnected sections without transitions
4. âŒ Add components that have 0 count in DXF
5. âŒ Use reference counts (only DXF counts)
6. âŒ Create duplicate sections
7. âŒ Remove sections from initial flow
8. âŒ Add debug output or category labels
9. âŒ Expose CAD codes (FS002V02, etc.)
10. âŒ Invent details not in DXF or references

### ALWAYS Do These:
1. âœ… Use natural, professional language from references
2. âœ… Create narrative flow with transitions
3. âœ… Use exact DXF counts
4. âœ… Translate category names naturally
5. âœ… Match reference tone and style
6. âœ… Add missing sections if in DXF + references
7. âœ… Keep client-specific terminology (e.g., "Noon's sorting logic")

---

## 6. 📋 SECTION ENHANCEMENT GUIDE

For each section type, here's how to refine:

### Infeed System:
- Start with shipment arrival: "dumped in bulk", "loaded onto"
- Describe movement: "ascend to", "enter the", "arrive at"
- Include distribution if VDS present: "Once within the loop"

### Inducts/Induction:
- Focus on operator action: "Operators manually position"
- Mention barcode orientation: "ensuring proper alignment and barcode visibility"
- Describe automation: "feedlines automatically induct"

### CBS (Loop or Linear):
- Keep it concise
- Emphasize sorting logic: "using [Client]'s sorting logic"
- Describe efficiency: "efficiently sorts shipments into"

### Output Chutes:
- Lead with discharge action: "Shipments are discharged into"
- Use sub-points (a., b., c.) for types
- Include count AND purpose for each type
- Natural descriptions: "50 chutes per zone collect shipments in roller cage trolleys"

### PTL System:
- Focus on function, not just count
- "PTL racks for consolidation", "further sorted via PTL setup"

### Bag Takeaway:
- Describe the post-sorting journey
- "After secondary sorting and bagging"
- Include destination: "transports them to the outbound docks"

---

## 7. ✅ QUALITY CHECKLIST

Before finalizing, verify:

**Language Quality:**
- [ ] NO category names exposed (VDS_BUFFER, AUTO_INDUCT, etc.)
- [ ] Natural verbs and descriptions used throughout
- [ ] Reads like it was written by a human expert
- [ ] Matches reference tone and phrasing

**Narrative Flow:**
- [ ] Each section connects to the next
- [ ] Tells the shipment journey from arrival to dispatch
- [ ] Transition phrases used between sections
- [ ] Logical progression maintained

**Content Accuracy:**
- [ ] All DXF counts used exactly as provided
- [ ] No reference counts used
- [ ] Missing sections added if in DXF + references
- [ ] No invented components
- [ ] Client name used correctly

**Structure:**
- [ ] Section numbering matches references (if applicable)
- [ ] Sub-points only for chute types
- [ ] Proper spacing between sections
- [ ] No duplicate sections

---

## 8. 🎯 FINAL REMINDER

Your output should read like this actual example:

**GOOD:**
"Shipments from FC and marketplace are dumped in bulk onto the infeed lines, where they ascend to a higher level and enter the VDS loop system. Once within the loop, shipments are manually picked and fed into the inducts."

**NOT like this:**
"The infeed system consists of 5 CONVEYOR_INFEED units. Shipments are loaded onto the conveyor and directed to the VDS/Buffer area, which consists of 2 VDS_BUFFER units."

Transform technical data into professional narrative.
Use DXF for accuracy, references for style.
Tell the shipment's story.
**Sample Examples (Only to understand style, do NOT copy content):** 
Process Flow

Infeed System: - Shipments from FC and Market place are dumped in bulk onto the infeed
lines . The shipments ascend to a higher level and arrive at the VDS loop system. Once the
shipments are within the VDS loop, they are picked manually and are fed among all inducts.
Inducts : - The operator picks and positions each shipment on the induct line, ensuring that
the shipment is properly aligned and that its barcode is facing upwards. The feedlines then
automatically induct the shipments onto the Cross -Belt Sorter Loop.
Loop CBS: - Once the shipments have entered the main loop, the Cross -Belt Sorter (CBS)
efficiently sorts the shipments into their respective output chutes by utilizing the data
provided by Noon 's sorting logic.
Output Chutes : - The shipments are discharged into two types of chutes.
a. Sliding Chutes - Within the loop CBS system, there are a total of 50 Sliding chutes for
each zone . The Shipments collected in Roller Cage trolleys, then they are
consolidated into bags using bagging type PTL racks .
b. Non -Sort Chutes - Within the loop CBS system, there are a total of 13 Non-Sort
Chutes per zone . Shipments collected within these chutes further undergo sortation
via PTL setup into Pallets.
c. Rejection Chute s- Two Rejection Chutes per zone are present to handle rejected
Shipments .
Bag Takeaway Conveyor: - Following the Secondary Sorting process of bagging type PTL , the
shipments are placed into bags and then manually loaded onto a bag takeaway conveyor
located beneath the CBS loop. This conveyor transports the bags out of shipment sorter area
to outbound docks.


Process Flow
Infeed System:
Bags containing shipments are unsealed and dumped in bulk onto the infeed lines of the Cross Belt Sorter equipped with Telescopic Belt Conveyor. The shipments ascend to a higher level and arrive at the VDS system. Once the shipments are within the VDS loop, they are evenly distributed among all inducts using Arm VDS technology.
Inducts:
After the shipments are collected in the VDS chute, an operator picks and positions each shipment on the induct line, ensuring that the shipment is properly aligned and that its barcode is facing upwards. The feed lines then automatically induct the shipments onto the Cross-Belt Sorter Loop.
Loop CBS:
Once the shipments have entered the main loop, the Cross-Belt Sorter (CBS) efficiently sorts the shipments into their respective output chutes by utilizing the data provided by Shadowfax's sorting logic.
Output Chutes:
The shipments are discharged into two types of chutes.
•	Direct Bagging Chutes (L-type):
Within a double decker loop CBS system, there are a total of 104 L-Type direct bagging chutes. Shipments collected within these direct bagging chutes are bagged and will be treated as high volume chutes.
•	Secondary Chutes (L-type):
Within a double decker loop CBS system, there are a total of 100 L-Type Secondary chutes. Shipments collected within these secondary chutes further undergo sortation via PTL setup placed at two levels.
•	Rejection Chute:
Four rejection chutes are present to handle rejected shipments.
Put To Light System:
In the system there are 3000 PTL locations. Each secondary chute is linked to 30 PTL locations. The PTL racks are placed in L-Shape double decker arrangement.
Bag Takeaway Conveyor:
Following the direct bagging process and secondary sorting process, the shipments are placed into bags and then manually loaded onto a bag takeaway conveyor located beneath the CBS loop. This conveyor transports the bags out of shipment sorter to outbound sorter located beneath base mezzanine in the approximate centre of the Loop CBS.


Process Flow
Infeed System: Boxes and totes are loaded onto the existing conveyor in a lengthwise orientation. From there, they are directed to their assigned highway line, which transports them to the CBS induct zone in a singulated manner.
Inducts: Upon arrival at the induct zone, Falcon's fully automatic induct line accurately and smoothly inducts the parcels onto the Linear CBS, based on their dimensions and weight.
Linear CBS: Once the parcels enter the main Linear CBS, the Cross-Belt Sorter (CBS) capture the barcode details & volume data after which it efficiently sorts the boxes and totes into their designated output chutes using data provided by Amazon.
Output Chutes: The Totes/Boxes are discharged into below output chutes.
a. Live Chutes - There are 9 sliding-type live chutes within the Linear CBS system, integrated with PVC belt conveyors and TBCs for live loading.
b. Collection chute – A total of 20 friction roller-based chutes are designed to collect and gradually accumulate the parcels.
c. Rejection Chute- One friction roller-based chute handles rejected shipments.
Recirculation Line: A recirculation line is available to automatically feed sortfail parcels back into the Linear CBS. It is also integrated with a manual loading point for reprocessed boxes and totes collected from the rejection chute.
---
JUST OUTPUT THE CLEAN PROCESS FLOW TEXT that matches the reference style."""
    
    # ✅ FIXED: Identify components that are BOTH in DXF AND references
    # This ensures we only add components that actually exist in current project
    missing = []
    for cat, count in dxf_cats.items():
        if count > 0:  # Only if DXF has it
            cat_lower = cat.lower()
            # Check if component is missing from initial flow
            if cat_lower not in initial_flow.lower():
                # Check if it exists in reference flows
                ref_mentions = any(cat_lower in ref["process_flow"].lower() 
                                 for ref in reference_chunks)
                if ref_mentions:
                    missing.append(f"{cat} ({count} units)")
    
    user_prompt = f"""
=== INITIAL FLOW ===
{initial_flow}

=== DXF COMPONENTS (CURRENT PROJECT - SOURCE OF TRUTH) ===
{create_dxf_summary(dxf_json)}

⚠️ CRITICAL RULE: ONLY add components that exist in CURRENT DXF (count > 0)
Do NOT add any components from references that don't exist in above DXF data.

=== COMPONENTS MISSING FROM INITIAL FLOW BUT PRESENT IN DXF ===
{', '.join(missing) if missing else 'None (initial flow has all DXF components)'}

=== REFERENCE STYLE (USE ONLY FOR TONE/LANGUAGE - NOT CONTENT) ===
{ref_context}

=== YOUR TASK ===

1. MAIN GOAL: Describe the CURRENT project's components (from DXF above)
2. HOW: Use the language/tone/style from references
3. DON'T: Add ANY components not in DXF (even if references have them)
4. DO: Use exact counts from DXF data
5. DO: Match reference's narrative flow and professional tone

Example of CORRECT approach:
- DXF shows: 50 chutes, 1 VDS, 5 PTL locations
- Reference shows: 100 chutes, 2 VDS, 3000 PTL locations
- YOUR OUTPUT: Should describe 50 chutes, 1 VDS, 5 PTL (use DXF counts exactly, match reference tone)

Example of WRONG approach:
- Adding "There are 3000 PTL locations" (reference's count, not in your DXF)
- Adding PTL section if your DXF has 0 PTL
- Mentioning components from reference that don't exist in current DXF

=== REFINEMENT INSTRUCTIONS ===

Step 1: Improve language in initial flow to match reference style
Step 2: Add missing sections ONLY if they're in DXF data
Step 3: Use exact DXF counts, never reference counts
Step 4: Create smooth narrative flow with transitions
Step 5: Ensure client name is correct: {client_name}

Output ONLY the refined process flow text. No notes or explanations."""
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    result = call_groq(messages, temp=0.15, max_tok=3500)
    return clean_generated_flow(result)

# ============================================================================
# IMPROVED ITERATIVE REFINEMENT
# ============================================================================


def iterative_refinement(
    current_flow: str,
    dxf_json: dict,
    reference_sample: str,
    evaluation: Dict,
    iteration: int,
    target_score: float
) -> str:
    """
    🔧 ULTRA-CONSERVATIVE: Apply micro-improvements only
    - Preserves ALL structure
    - Changes ONLY phrasing
    - Validates before returning
    """
    
    current_coherence = evaluation['structural_coherence']
    gap_to_target = target_score - current_coherence
    
    # Format feedback
    feedback_text = ""
    if evaluation.get("feedback"):
        feedback_text = "🎯 MICRO-IMPROVEMENTS TO APPLY:\n\n"
        for i, fb in enumerate(evaluation["feedback"], 1):
            feedback_text += f"{i}. {fb}\n"
        feedback_text += "\n⚠️ Apply these changes ONE AT A TIME. Do NOT combine or expand them."
    else:
        feedback_text = "✅ No issues found - flow is at target quality"
    
    system_prompt = f"""You are an expert at applying MICRO language improvements to technical documentation.

## 🎯 MISSION: ITERATION {iteration} - Apply Micro Word Changes

**Current Score:** {current_coherence:.1f}/100
**Target Score:** {target_score}/100
**Gap:** {gap_to_target:.1f} points

## ⚡ EXECUTION STRATEGY

### Step 1: Read Feedback
Each feedback item suggests changing 1-2 words. Read them carefully.

### Step 2: Apply Changes ONE AT A TIME
For each feedback item:
1. Find the EXACT phrase mentioned
2. Make ONLY the word change suggested
3. Leave everything else untouched

### Step 3: Verify No Breaking Changes
- ✅ All sections still present?
- ✅ All counts unchanged?
- ✅ CBS type unchanged?
- ✅ Structure intact?

## 🚫 CRITICAL RULES

**DO:**
- Change ONLY the specific words mentioned in feedback
- Keep all section structure exactly same
- Preserve all numbers and counts
- Maintain all section names

**DO NOT:**
- Rewrite entire sections
- Combine feedback items into big changes
- Remove any content
- Add new sections
- Change any numbers
- Alter CBS type or client name

## 📋 FEEDBACK TO APPLY

{feedback_text}

## ⚠️ WARNING

Making changes beyond what's suggested in feedback causes score DROPS.
Stay conservative. Small improvements add up.

Expected score increase: ~{min(gap_to_target, 5):.1f} points

## 📤 OUTPUT

Generate the refined flow with ONLY the micro-changes applied.

**MUST:**
- Start with "Process Flow"
- Keep ALL existing sections
- Change ONLY words mentioned in feedback
- Preserve all structure and formatting"""

    user_prompt = f"""## CURRENT FLOW (Apply micro-improvements to this)

```
{current_flow}
```

## REFERENCE FLOW (For style reference only)

```
{reference_sample[:1200]}
```

---

## 🎯 TASK: Apply the {len(evaluation.get('feedback', []))} micro-improvements

**IMPORTANT:**
- Apply ONLY the changes listed in feedback
- Do NOT make additional improvements
- Do NOT rewrite sections
- Keep structure identical

Generate refined flow now."""

    try:
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        # Lower temperature for more predictable results
        result = call_groq(messages, temp=0.05, max_tok=3500)
        cleaned = clean_generated_flow(result)
        
        # Validate output
        is_valid, error_msg = validate_flow_quality(cleaned)
        if not is_valid:
            logger.warning(f"Generated flow failed validation: {error_msg}")
            return current_flow  # Revert to current on failure
        
        return cleaned
        
    except Exception as e:
        logger.error(f"Error in iterative refinement: {e}")
        return current_flow  # Revert to current on error


# ============================================================================
# STREAMLIT UI WITH IMPROVED LOGIC
# ============================================================================

def main():
    st.set_page_config(page_title="Iterative Process Flow Generator", layout="wide")
    
    st.title("📄 Iterative Process Flow Generator")
    st.markdown("**Multi-step refinement with progressive structural improvement**")
    
    # Sidebar for inputs
    with st.sidebar:
        st.header("📋 Input")
        uploaded = st.file_uploader("Upload DXF", type=["dxf"])
        client = st.text_input("Client Name", "Zepto")
        project = st.text_input("Project Name", "")
        
        max_iterations = st.slider("Max Iterations", 1, 10, 5)
        target_score = st.slider("Target Structural Coherence", 70, 95, 85)
        min_improvement = st.slider("Min Improvement per Iteration", 0.5, 5.0, 1.0)
        
        generate_btn = st.button("🚀 Generate", type="primary", use_container_width=True)
    
    # Initialize session state
    if 'results' not in st.session_state:
        st.session_state.results = None
    
    if generate_btn and uploaded:
        # Save uploaded file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
            tmp.write(uploaded.read())
            tmp_path = Path(tmp.name)
        
        # Container for progress
        progress_container = st.container()
        
        with progress_container:
            results = {
                "iterations": [],
                "final_flow": None,
                "dxf_json": None,
                "references": None,
                "stopped_reason": None,
            }
            
            # ================================================================
            # STEP 0: DXF EXTRACTION
            # ================================================================
            with st.status("📊 Step 0: Extracting DXF Components...", expanded=True) as status:
                try:
                    dxf_json = extract_dxf_components(tmp_path, project or uploaded.name)
                    dxf_json['client'] = client
                    results["dxf_json"] = dxf_json
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Components", dxf_json['total_components'])
                    with col2:
                        st.metric("CBS Type", dxf_json['cbs_type'])
                    with col3:
                        st.metric("Induction", dxf_json['induction_type'])
                    
                    st.success("✅ DXF extracted successfully")
                    status.update(label="✅ Step 0: DXF Extracted", state="complete")
                except Exception as e:
                    st.error(f"❌ DXF extraction failed: {e}")
                    return
            
            # ================================================================
            # STEP 1: INITIAL GENERATION
            # ================================================================
            with st.status("✏️ Step 1: Generating Initial Flow...", expanded=True) as status:
                try:
                    initial_flow = generate_initial_flow(client, dxf_json)
                    
                    st.text_area("Initial Flow Output", initial_flow, height=300, key="step1_output")
                    st.info("ℹ️ Generated from DXF data only")
                    
                    results["iterations"].append({
                        "iteration": 0,
                        "stage": "Initial",
                        "flow": initial_flow,
                        "input": "DXF data only",
                    })
                    
                    status.update(label="✅ Step 1: Initial Flow Generated", state="complete")
                except Exception as e:
                    st.error(f"❌ Initial generation failed: {e}")
                    return
            
            # ================================================================
            # STEP 2: QUERY REFERENCES & REFINE
            # ================================================================
            with st.status("🔍 Step 2: Querying References & Refining...", expanded=True) as status:
                try:
                    # Query Pinecone
                    pc, index = get_pinecone_index()
                    dxf_summary = create_dxf_summary_for_embedding(dxf_json)
                    references = query_similar_flows(pc, index, dxf_summary, dxf_json, top_k=2, threshold=0.75)
                    results["references"] = references
                    
                    if references:
                        st.write(f"**Found {len(references)} similar references:**")
                        for i, ref in enumerate(references, 1):
                            score = ref.get("combined_score", 0)
                            st.write(f"{i}. {ref['client']} - Score: {score:.3f}")
                    else:
                        st.warning("⚠️ No references found above threshold")
                    
                    # Generate second flow
                    second_flow = generate_second_flow_with_chunks(
                        initial_flow, dxf_json, references, client
                    )
                    
                    st.text_area("Refined Flow Output", second_flow, height=300, key="step2_output")
                    
                    results["iterations"].append({
                        "iteration": 1,
                        "stage": "Refined with References",
                        "flow": second_flow,
                        "input": f"{len(references)} reference chunks",
                    })
                    
                    current_flow = second_flow
                    status.update(label="✅ Step 2: Flow Refined", state="complete")
                except Exception as e:
                    st.error(f"❌ Reference refinement failed: {e}")
                    current_flow = initial_flow
            
            # ================================================================
            # STEP 3 & 4: EVALUATION & ITERATIVE REFINEMENT (IMPROVED)
            # ================================================================
            st.markdown("---")
            st.subheader("🔄 Progressive Iterative Refinement")
            
            # Use best reference as target
            target_reference = references[0]["process_flow"] if references else current_flow
            
            # Initialize tracking variables
            iteration_num = 2
            best_score = 0
            best_flow = current_flow
            no_improvement_count = 0
            max_no_improvement = 2  # Stop if no improvement for 2 iterations
            
            for iter_count in range(max_iterations):
                with st.expander(f"**Iteration {iteration_num}**", expanded=(iter_count == 0)):
                    col1, col2 = st.columns([1, 1])
                    
                    # In the main() function, update the evaluation call:

                    with col1:
                        st.markdown("##### 📊 Evaluation")
                        
                        # Evaluate current flow - PASS target_score parameter
                        evaluation = evaluate_process_flow(
                            current_flow, 
                            target_reference,
                            dxf_json,
                            target_score=target_score  # Add this parameter
                        )
                        
                        current_score = evaluation['structural_coherence']
                        
                        # Display scores
                        score_col1, score_col2 = st.columns(2)
                        with score_col1:
                            delta_text = ""
                            if iter_count > 0:
                                delta = current_score - best_score
                                delta_text = f"+{delta:.1f}" if delta > 0 else f"{delta:.1f}"
                            
                            st.metric(
                                "Structural Coherence", 
                                f"{current_score:.1f}",
                                delta=delta_text if delta_text else None
                            )
                            st.metric("Style Match", f"{evaluation['style_match']:.1f}")
                        with score_col2:
                            st.metric("Component Coverage", f"{evaluation['component_coverage']:.1f}")
                            st.metric("BERT F1", f"{evaluation['bert_f1']:.1f}")
                        
                        # Display AI-generated feedback with special formatting
                        if evaluation['feedback']:
                            st.markdown("**🤖 AI-Generated Feedback:**")
                            for i, fb in enumerate(evaluation['feedback'], 1):
                                # Color code by priority (first items are higher priority)
                                if i <= 2:
                                    st.error(f"🔴 **Priority {i}:** {fb}")
                                elif i <= 4:
                                    st.warning(f"🟡 {fb}")
                                else:
                                    st.info(f"🔵 {fb}")
                        else:
                            st.success("✅ AI Analysis: No issues found - flow matches target!")
                    
                    with col2:
                        st.markdown("##### 📝 Current Flow")
                        st.text_area(
                            "Flow", 
                            current_flow, 
                            height=300, 
                            key=f"iter_{iteration_num}_flow",
                            label_visibility="collapsed"
                        )
                    
                    # Check if this is an improvement
                    improvement = current_score - best_score
                    
                    if current_score > best_score:
                        # This is better - accept it
                        best_score = current_score
                        best_flow = current_flow
                        no_improvement_count = 0
                        
                        st.success(f"✅ Improvement: +{improvement:.1f} points")
                    else:
                        # No improvement - keep best flow
                        no_improvement_count += 1
                        st.warning(f"⚠️ No improvement ({improvement:.1f}). Keeping best flow.")
                        current_flow = best_flow  # Revert to best
                    
                    # Save iteration results
                    results["iterations"].append({
                        "iteration": iteration_num,
                        "stage": "Evaluated",
                        "flow": current_flow,
                        "evaluation": evaluation,
                        "score": current_score,
                        "is_best": current_score == best_score,
                    })
                    
                    # Check stopping conditions
                    stop_reason = None
                    
                    # 1. Target reached
                    if current_score >= target_score:
                        stop_reason = f"✅ Target score reached: {current_score:.1f} >= {target_score}"
                        st.success(stop_reason)
                    
                    # 2. No feedback and high score
                    elif not evaluation['feedback'] and current_score >= 75:
                        stop_reason = f"✅ No issues found and score is strong: {current_score:.1f}/100"
                        st.success(stop_reason)
                    
                    # 3. No improvement for multiple iterations
                    elif no_improvement_count >= max_no_improvement:
                        stop_reason = f"⚠️ No improvement for {max_no_improvement} iterations. Stopping."
                        st.warning(stop_reason)
                    
                    # 4. Very high score already
                    elif current_score >= 90:
                        stop_reason = f"✅ Excellent score achieved: {current_score:.1f}/100"
                        st.success(stop_reason)
                    
                    if stop_reason:
                        results["stopped_reason"] = stop_reason
                        break
                    
                    # Generate next iteration if not last
                    if iter_count < max_iterations - 1:
                        st.markdown("##### 🔄 Generating Next Iteration...")
                        try:
                            # Use best flow as base for next iteration
                            next_flow = iterative_refinement(
                                best_flow,  # Always start from best
                                dxf_json,
                                target_reference,
                                evaluation,
                                iteration_num,
                                target_score
                            )
                            
                            # Validate the new flow
                            is_valid, error_msg = validate_flow_quality(next_flow)
                            if not is_valid:
                                st.error(f"❌ Generated flow is invalid: {error_msg}")
                                st.warning("Using previous best flow instead")
                                next_flow = best_flow
                            
                            current_flow = next_flow
                            iteration_num += 1
                            
                        except Exception as e:
                            st.error(f"❌ Refinement failed: {e}")
                            results["stopped_reason"] = f"Error: {e}"
                            break
            
            # Set final results
            if not results.get("stopped_reason"):
                results["stopped_reason"] = f"Completed all {max_iterations} iterations"
            
            results["final_flow"] = best_flow
            results["final_score"] = best_score
            st.session_state.results = results
        
        # Clean up
        tmp_path.unlink()
    
    # ================================================================
    # DISPLAY FINAL RESULTS
    # ================================================================
    if st.session_state.results:
        st.markdown("---")
        st.header("📊 Final Results")
        
        results = st.session_state.results
        
        # Display stop reason
        if results.get("stopped_reason"):
            if "✅" in results["stopped_reason"]:
                st.success(results["stopped_reason"])
            elif "⚠️" in results["stopped_reason"]:
                st.warning(results["stopped_reason"])
            else:
                st.info(results["stopped_reason"])
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Final Score", f"{results.get('final_score', 0):.1f}/100")
        with col2:
            st.metric("Iterations", len(results['iterations']))
        with col3:
            st.metric("References Used", len(results.get('references', [])))
        with col4:
            # Count improvements
            improvements = sum(1 for it in results['iterations'] if it.get('is_best', False))
            st.metric("Improvements", improvements)
        
        # Tabs for different views
        tab1, tab2, tab3, tab4 = st.tabs(["📄 Final Flow", "📈 Score Progression", "🔄 Iteration History", "📁 DXF Analysis"])
        
        with tab1:
            st.text_area("Final Process Flow", results['final_flow'], height=500)
            st.download_button(
                "💾 Download Flow",
                results['final_flow'],
                file_name=f"{client}_process_flow.txt",
                mime="text/plain"
            )
        
        with tab2:
            # Plot score progression
            iterations = [it['iteration'] for it in results['iterations'] if 'score' in it]
            scores = [it['score'] for it in results['iterations'] if 'score' in it]
            
            if iterations and scores:
                import pandas as pd
                df = pd.DataFrame({
                    'Iteration': iterations,
                    'Structural Coherence': scores
                })
                st.line_chart(df.set_index('Iteration'))
                
                # Show improvement summary
                st.markdown("### Improvement Summary")
                if len(scores) > 1:
                    total_improvement = scores[-1] - scores[0]
                    st.metric("Total Improvement", f"{total_improvement:+.1f} points")
                    st.metric("Best Score Achieved", f"{max(scores):.1f}/100")
        
        with tab3:
            for iter_data in results['iterations']:
                iter_num = iter_data['iteration']
                stage = iter_data['stage']
                is_best = iter_data.get('is_best', False)
                
                title = f"Iteration {iter_num}: {stage}"
                if is_best:
                    title += " ⭐ (Best)"
                
                with st.expander(title):
                    if 'evaluation' in iter_data:
                        eval_data = iter_data['evaluation']
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Structural", f"{eval_data['structural_coherence']:.1f}")
                        with col2:
                            st.metric("Style", f"{eval_data['style_match']:.1f}")
                        with col3:
                            st.metric("Coverage", f"{eval_data['component_coverage']:.1f}")
                        with col4:
                            st.metric("BERT F1", f"{eval_data['bert_f1']:.1f}")
                        
                        if eval_data.get('feedback'):
                            st.markdown("**Feedback:**")
                            for fb in eval_data['feedback']:
                                st.write(f"• {fb}")
                    
                    st.text_area("Flow", iter_data['flow'], height=200, key=f"history_{iter_num}")
        
        with tab4:
            dxf = results['dxf_json']
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("CBS Type", dxf['cbs_type'])
            with col2:
                st.metric("Induction", dxf['induction_type'])
            with col3:
                st.metric("Total Components", dxf['total_components'])
            
            st.markdown("**Component Categories:**")
            for cat, count in sorted(dxf['category_summary'].items(), key=lambda x: -x[1]):
                st.write(f"• {cat}: {count} units")
            
            if dxf.get('chute_analysis', {}).get('total', 0) > 0:
                st.markdown("**Chute Analysis:**")
                chute = dxf['chute_analysis']
                st.write(f"Total: {chute['total']} chutes")
                for chute_type, count in chute.get('by_type', {}).items():
                    st.write(f"  • {chute_type.replace('_', ' ').title()}: {count}")


if __name__ == "__main__":
    main()
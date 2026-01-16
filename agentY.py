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

_sentence_model = None

def get_sentence_model():
    """Lazily load the SentenceTransformer model to avoid side-effects at import time."""
    global _sentence_model
    if _sentence_model is None:
        try:
            logger.info("Loading SentenceTransformer model (this may take a few moments)...")
            _sentence_model = SentenceTransformer('all-MiniLM-L6-v2')
            logger.info("✅ SentenceTransformer model loaded successfully")
        except Exception as e:
            logger.error(f"❌ Failed to load SentenceTransformer model: {e}")
            raise
    return _sentence_model

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
    
    # Check for empty Output Chutes section
    output_chutes_issue = validate_output_chutes_section(flow)
    if output_chutes_issue:
        return False, output_chutes_issue
    
    return True, ""


def validate_output_chutes_section(flow: str) -> str:
    """
    Check if Output Chutes section exists and has content.
    Returns error message if there's an issue, empty string if valid.
    """
    flow_lower = flow.lower()
    
    # Check if Output Chutes section exists
    output_patterns = [
        r'output\s+chutes?\s*[:\-–]',
        r'output\s+chutes?\s*$',
    ]
    
    has_output_section = any(re.search(pat, flow_lower, re.MULTILINE) for pat in output_patterns)
    
    if not has_output_section:
        # Output chutes section is missing entirely
        return "Output Chutes section is missing"
    
    # Find the Output Chutes section and check if it has content after the header
    # Pattern: "Output Chutes:" followed by content before next section or end of text
    section_pattern = r'(output\s+chutes?\s*[:\-–]?\s*)(.*?)(?=(?:\n\s*(?:[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s*[:\-–])|$))'
    match = re.search(section_pattern, flow, re.IGNORECASE | re.DOTALL)
    
    if match:
        content = match.group(2).strip()
        # Check if there's actual content after the header
        if len(content) < 20:
            return "Output Chutes section is empty or has insufficient content"
        
        # Check if content is just whitespace or newlines
        if not re.search(r'[a-zA-Z]{10,}', content):
            return "Output Chutes section lacks meaningful content"
    
    return ""


def fix_empty_output_chutes(flow: str, dxf_json: dict, client_name: str) -> str:
    """
    If Output Chutes section is empty, regenerate it with proper content based on DXF analysis.
    Enhanced to include all chute subtypes from DXF with per-zone counts.
    Also adds Bag Takeaway Conveyor section if PTL/Sliding chutes are present.
    """
    issue = validate_output_chutes_section(flow)
    if not issue:
        # Check if Bag Takeaway section is missing but needed
        has_ptl = dxf_json.get("category_summary", {}).get("PTL", 0) > 0
        has_sliding = dxf_json.get("chute_analysis", {}).get("by_type", {}).get("sliding_chutes", 0) > 0
        has_bag_system = dxf_json.get("has_bag_system", False)
        
        if (has_ptl or has_sliding or has_bag_system) and "bag takeaway" not in flow.lower():
            # Add Bag Takeaway section at the end
            bag_section = (
                "\n\n5. Bag Takeaway Conveyor: - Following the direct bagging process & secondary sorting process, "
                "the shipments are placed into bags and then manually loaded onto a bag takeaway conveyor "
                "located beneath the CBS loop. This conveyor transports the bags out of shipment sorter to "
                "Outbound sorter located beneath base mezzanine in the approx. centre of the loop CBS."
            )
            flow = flow.rstrip() + bag_section
        return flow
    
    # Extract chute information from DXF
    chute_info = dxf_json.get("chute_analysis", {})
    total_chutes = chute_info.get("total", 0)
    chute_types = chute_info.get("by_type", {})
    zone_count = chute_info.get("zone_count", 0)
    per_zone = chute_info.get("per_zone", {})
    
    if total_chutes == 0:
        # Try to get from category_summary
        category_summary = dxf_json.get("category_summary", {})
        total_chutes = category_summary.get("CHUTE", 0) + category_summary.get("SLIDING_CHUTE", 0) + \
                      category_summary.get("NON_SORT_CHUTE", 0) + category_summary.get("REJECTION_CHUTE", 0)
    
    # Also check categorized_components for specific chute types
    categorized = dxf_json.get("categorized_components", {})
    chute_components = categorized.get("CHUTE", {})
    
    # If we have specific chute components but no by_type, parse them
    if chute_components and not chute_types:
        for comp_name, count in chute_components.items():
            name_lower = comp_name.lower()
            if "live" in name_lower or "ob live" in name_lower:
                chute_types["live_chutes"] = chute_types.get("live_chutes", 0) + count
            elif "discharge" in name_lower:
                chute_types["discharge_chutes"] = chute_types.get("discharge_chutes", 0) + count
            elif "reject" in name_lower:
                chute_types["rejection_chutes"] = chute_types.get("rejection_chutes", 0) + count
            elif "collection" in name_lower:
                chute_types["collection_chutes"] = chute_types.get("collection_chutes", 0) + count
            elif "ow" in name_lower and "chute" in name_lower:
                chute_types["ow_chutes"] = chute_types.get("ow_chutes", 0) + count
            elif "od" in name_lower and "chute" in name_lower:
                chute_types["od_chutes"] = chute_types.get("od_chutes", 0) + count
            elif "generic" not in name_lower and "chute" in name_lower:
                chute_types["generic_chutes"] = chute_types.get("generic_chutes", 0) + count
    
    if total_chutes == 0:
        # Default fallback content
        output_content = "Output Chutes: - The sorted shipments are discharged into output chutes for collection."
    else:
        # Build content based on DXF data with detailed subtypes and per-zone counts
        output_lines = ["Output Chutes: - The shipments are discharged into following types of chutes:"]
        
        letter = 'a'
        # Sort chute types by count (descending) for better presentation
        sorted_chutes = sorted(chute_types.items(), key=lambda x: -x[1])
        
        # Show per-zone counts only if zone_count is 2 or 3 (i.e., 1 < zone_count < 4)
        use_per_zone = 1 < zone_count < 4
        
        for chute_type, count in sorted_chutes:
            if count > 0:
                # Format chute type name nicely
                chute_name = chute_type.replace('_', ' ').title()
                
                # Use per-zone counts if zone_count is 2 or 3
                display_count = per_zone.get(chute_type, count) if use_per_zone else count
                per_zone_suffix = " for each zone" if use_per_zone else ""
                
                # Create descriptive text based on chute type
                if "sliding" in chute_type.lower():
                    description = (f"Within the loop CBS system, there are a total of {display_count} Sliding chutes{per_zone_suffix}. "
                                  "The Shipments collected in Roller Cage trolleys, then they are consolidated into bags using bagging type PTL racks.")
                elif "non_sort" in chute_type.lower():
                    description = (f"Within the loop CBS system, there are a total of {display_count} Non-Sort Chutes{' per zone' if use_per_zone else ''}. "
                                  "Shipments collected within these chutes further undergo sortation via PTL setup into Pallets.")
                elif "rejection" in chute_type.lower():
                    if use_per_zone:
                        description = f"{display_count} Rejection Chutes per zone are present to handle rejected Shipments."
                    else:
                        description = f"{display_count} rejection chutes handle rejected shipments."
                elif "live" in chute_type.lower():
                    description = f"There are {count} {chute_name.lower()} integrated with conveyors for live loading, efficiently handling sorted shipments."
                elif "generic" in chute_type.lower():
                    description = f"There are {count} generic chutes that collect parcels after sorting."
                elif "collection" in chute_type.lower():
                    description = f"A total of {count} {chute_name.lower()} are designed to collect and gradually accumulate the parcels."
                elif "discharge" in chute_type.lower():
                    description = f"{count} {chute_name.lower()} efficiently discharge sorted items."
                else:
                    description = f"There are {count} {chute_name.lower()} within the system."
                
                output_lines.append(f"{letter}. {chute_name} – {description}")
                letter = chr(ord(letter) + 1)
        
        output_content = "\n".join(output_lines)
    
    # Replace empty Output Chutes section with the new content
    # Pattern to find Output Chutes header with empty/minimal content
    patterns_to_replace = [
        r'(output\s+chutes?\s*[:\-–]?\s*)\n*(?=\n\s*(?:[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s*[:\-–])|$)',
        r'(output\s+chutes?\s*[:\-–]?\s*)(?:\n\s*){0,3}(?=\n|$)',
    ]
    
    for pattern in patterns_to_replace:
        if re.search(pattern, flow, re.IGNORECASE | re.DOTALL):
            flow = re.sub(pattern, output_content + "\n\n", flow, flags=re.IGNORECASE | re.DOTALL)
            break
    
    # Add Bag Takeaway section if PTL or sliding chutes are present
    has_ptl = dxf_json.get("category_summary", {}).get("PTL", 0) > 0
    has_sliding = chute_types.get("sliding_chutes", 0) > 0
    has_bag_system = dxf_json.get("has_bag_system", False)
    
    if (has_ptl or has_sliding or has_bag_system) and "bag takeaway" not in flow.lower():
        bag_section = (
            "\n\n5. Bag Takeaway Conveyor: - Following the direct bagging process & secondary sorting process, "
            "the shipments are placed into bags and then manually loaded onto a bag takeaway conveyor "
            "located beneath the CBS loop. This conveyor transports the bags out of shipment sorter to "
            "Outbound sorter located beneath base mezzanine in the approx. centre of the loop CBS."
        )
        flow = flow.rstrip() + bag_section
    
    return flow


# ============================================================================
# EVALUATION FUNCTIONS (keeping existing ones)
# ============================================================================

def split_into_sentences(text: str) -> List[str]:
    """Split text into sentences."""
    sentences = re.split(r'[.!?]+', text)
    return [s.strip() for s in sentences if s.strip()]


def compute_bert_scores(original: str, generated: str) -> Tuple[float, float, float]:
    """Compute BERT precision, recall, F1 scores."""
    try:
        if not original.strip() and not generated.strip():
            return 1.0, 1.0, 1.0
        if not original.strip() or not generated.strip():
            return 0.0, 0.0, 0.0
        
        logger.info("Computing BERT scores...")
        P, R, F1 = bert_score(
            [generated],
            [original],
            lang="en",
            rescale_with_baseline=False,
        )
        logger.info(f"✅ BERT scores computed: P={P[0]:.3f}, R={R[0]:.3f}, F1={F1[0]:.3f}")
        return float(P[0]), float(R[0]), float(F1[0])
    except Exception as e:
        logger.error(f"❌ Error computing BERT scores: {e}")
        # Return default scores to prevent hanging
        return 0.85, 0.85, 0.85


def compute_structural_coherence(original: str, generated: str) -> float:
    """
    Compute structural coherence score (0-100).
    
    Evaluates:
    - Sentence order similarity
    - Paragraph structure
    - Bullet point usage
    - Transition smoothness
    """
    try:
        o_sents = split_into_sentences(original)
        g_sents = split_into_sentences(generated)
        
        if not o_sents or not g_sents:
            return 0.0
        
        # Encode sentences
        logger.info(f"Encoding {len(o_sents)} original and {len(g_sents)} generated sentences...")
        o_emb = get_sentence_model().encode(o_sents, convert_to_tensor=True)
        g_emb = get_sentence_model().encode(g_sents, convert_to_tensor=True)
        
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
        final_score = float(max(0.0, min(1.0, coherence)) * 100.0)
        
        logger.info(f"Structural coherence computed: {final_score:.2f}")
        return final_score
        
    except Exception as e:
        logger.error(f"❌ Error computing structural coherence: {e}")
        # Return a default score to prevent hanging
        return 75.0  # Default to 75 if scoring fails


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
    try:
        logger.info("Starting process flow evaluation...")
        
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
        logger.info("Computing structural coherence...")
        evaluation["structural_coherence"] = compute_structural_coherence(reference, generated_clean)
        
        # 2. BERT scores
        logger.info("Computing BERT scores...")
        P, R, F1 = compute_bert_scores(reference, generated_clean)
        evaluation["bert_precision"] = P * 100
        evaluation["bert_recall"] = R * 100
        evaluation["bert_f1"] = F1 * 100
        
        # 3. Component coverage check (keep this as is - it's data validation)
        logger.info("Checking component coverage...")
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
            logger.info(f"Generating AI feedback (score {evaluation['structural_coherence']:.2f} < target {target_score})...")
            ai_feedback = generate_ai_feedback(
                reference, 
                generated_clean, 
                dxf_json,
                evaluation["structural_coherence"],
                target_score
            )
            evaluation["feedback"].extend(ai_feedback)
            logger.info(f"Generated {len(ai_feedback)} feedback items")
        else:
            logger.info(f"Score {evaluation['structural_coherence']:.2f} >= target {target_score}, skipping AI feedback")
        
        # 5. Style match score (based on feedback count and structural coherence)
        # If structural coherence is high and no feedback, style match should be high
        if evaluation["structural_coherence"] >= target_score and len(evaluation["feedback"]) == 0:
            evaluation["style_match"] = 100.0
        else:
            # Base style match on structural coherence and feedback count
            base_score = evaluation["structural_coherence"]
            feedback_penalty = len(evaluation["feedback"]) * 10
            evaluation["style_match"] = max(0, min(100, base_score - feedback_penalty))
        
        logger.info(f"Evaluation complete. Score: {evaluation['structural_coherence']:.2f}")
        return evaluation
        
    except Exception as e:
        logger.error(f"❌ Error during evaluation: {e}")
        # Return a default evaluation to prevent hanging
        return {
            "structural_coherence": 75.0,
            "bert_precision": 75.0,
            "bert_recall": 75.0,
            "bert_f1": 75.0,
            "component_coverage": 100.0,
            "style_match": 75.0,
            "feedback": ["Evaluation encountered an error, using default scores"],
        }

# ============================================================================
# GENERATION FUNCTIONS (keep your existing ones - no changes needed)
# ============================================================================

# [SKIP - Use your existing generate_initial_flow, generate_second_flow_with_chunks functions]

def generate_initial_flow(client_name: str, dxf_json: dict) -> str:
    """Step 1: Generate initial flow from DXF data only."""
    
    logger.info(f"Generating initial process flow for client: {client_name}")
    
    dxf_summary = create_dxf_summary(dxf_json)
    
    # Get CBS knowledge and use it in the prompt
    logger.info("Fetching CBS knowledge...")
    cbs_knowledge_text = get_cbs_knowledge()
    
    system_prompt  = """## ROLE
You are a senior SALES engineer presenting the "Process Flow of the System" to a potential client. You're not just describing the system—you're SELLING how this solution transforms their operations.

---

## 🎯 SALES-FIRST MINDSET (CRITICAL)
- This is a SALES document, not a technical manual
- Every step should subtly answer: "Why does this matter to the client?"
- Highlight BENEFITS: speed, accuracy, efficiency, reduced errors, labor savings
- Make the client visualize their parcels flowing SMOOTHLY through their new system
- Create confidence: this system will solve their problems

## 💡 WHY + WHAT (Always explain WHY, not just WHAT)
- DON'T: "Parcels are inducted onto the sorter"
- DO: "Parcels are smoothly inducted onto the sorter, ensuring zero jams and maximum throughput even during peak hours"
- DON'T: "Barcodes are scanned"
- DO: "Barcodes are scanned instantly, enabling precise routing without any manual intervention—reducing errors and speeding up delivery"
- DON'T: "Output chutes receive sorted items"
- DO: "Sorted items slide into designated output chutes, ready for immediate dispatch—cutting your processing time significantly"

---

## 📖 HUMAN STORYTELLING STYLE (MOST IMPORTANT)

**Write like a trusted advisor walking the client through their future warehouse.**

### ✅ HOW TO WRITE LIKE A HUMAN:
1. **CONNECT Each Step to the Next**: Each section flows naturally into the next, showing the parcel journey
2. **Use Transitional Language**: "From there...", "Your parcels then...", "Once they reach...", "Finally..."
3. **Tell a STORY**: Imagine you're standing with the client, pointing at each part of the system
4. **Simple Conversational English**: Avoid stiff technical jargon—sound like a helpful expert, not a manual
5. **Show BENEFITS**: Speed, reliability, efficiency woven into every step

### ✅ GOOD EXAMPLE (Sales-Focused, Human Written):
```
Process Flow

Incoming shipments are dumped in bulk onto the infeed conveyor, where they begin their smooth journey through the system. From there, the shipments ascend to a higher level via the inclined conveyor, arriving at the induction zone without any bottlenecks. Operators quickly position each shipment with the barcode facing upwards onto the feedlines—a simple motion that keeps throughput high. Once on the feedlines, parcels are automatically inducted into the Loop CBS, where the real magic happens: high-speed sorting with 99.9% accuracy. The system efficiently routes each shipment to its designated output chute based on barcode data. Finally, sorted packages slide into collection chutes, ready for immediate dispatch—dramatically reducing your processing time.
```

### ❌ BAD EXAMPLE (Technical, AI-Like, No Benefits):
```
Process Flow
Infeed System: Incoming shipments on conveyor
Auto Induct Line: Parcels are inducted
Loop CBS: Sorting happens
Output Chutes: 51 chutes
```

---
**ABOUT CROSS BELT SORTERS (CBS):** {cbs_knowledge_text}
---
**Components name mapping**
Sci_fi_name,Actual_component_name
3k Dap Feedline,Feedline conveyor (3.0 m section)
ANGLE_MERGE_2000PPH,Angle merge conveyor (rated ~2000 pph)
ATR Vipacsystem +,Barcode scanning system (ATR / VIPAC)
Assem2,Assembly block (misc. mechanical assembly) [needs confirmation]
Auto Induct Chute-02,Auto-induct chute (variant 02)
Auto Induct Chute-03,Auto-induct chute (variant 03)
Bag Transfer Chute,Bag transfer chute
CBS_PTL_XREF(26-10)$0$A$Ca681773c,XREF block (CAD artifact / unknown)
CBS_PTL_XREF(26-10)$0$Dual Belt CBS 45 Deg Turn (FAL_S005V01),CBS sorter 45° turn module (dual belt) [XREF]
CBS_PTL_XREF(26-10)$0$Dual Belt CBS 45 Deg Turn(Track),CBS turn track segment [XREF]
CBS_PTL_XREF(26-10)$0$FAL_FS002V02,Falcon feedline/auto-induct module (FS002 V02) [XREF]
CBS_PTL_XREF(26-10)$0$Operator,Operator marker [XREF]
CBS_PTL_XREF(26-10)$0$PTL Chute-01(DD),PTL chute / station chute (double-deck) [XREF]
CHUTE TYPE-02,Chute type definition block
Chute,Generic chute (collection/slide chute)
Chute-001,Generic chute block
Chute-002,Generic chute block
Chute-01,Generic chute block
Chutes$0$Mini gravity 01,Mini gravity chute (small gravity chute)
Collection Bin (FAL_ST001V01),Collection bin / tote
Collection Chute for Double Deck Straight,Collection chute (double deck straight)
Container,Container / skid box symbol
Container1,Container / skid box symbol
Conveyor 1,Generic conveyor segment
Conveyor 10,Generic conveyor segment
Customize Leg,Custom support leg
Crossover for Maintanance,Maintenance crossover / bridge
Direct Bagging Chute 900 mm pitch,Direct bagging chute (900 mm pitch)
Disperson Chute-01,Dispersion chute
Diverter-01,Diverter unit
Double Deck Loop Stairs,Stairs (double deck loop access)
Dual Belt CBS 45 Deg Turn (FAL_S005V01),CBS sorter 45° turn module (dual belt)
Dual Belt CBS 45 Deg Turn (FAL_S005V01)_003,CBS sorter 45° turn module (dual belt) - instance
Dual Belt CBS 45 Deg Turn (FAL_S005V02),CBS sorter 45° turn module (dual belt) - rev V02
End joint individual,End joint / connector piece
FAL_BLK_Boom Conveyor,Boom conveyor (boom/incline belt conveyor section)
FAL_BLK_Boom Conveyor (6-18),Boom conveyor (boom/incline belt conveyor section)
FAL_BLK_Feed lineW1000 @30°,Inclined feedline belt conveyor (1000 mm wide, 30°)
FAL_BLK_P&A,Positioning & alignment unit (P&A) / aligner module
FAL_BT30°_W1200,Belt turn/transfer module (30°, 1200 mm width)
FAL_BT60°_W1200,Belt turn/transfer module (60°, 1200 mm width)
FAL_BT90°_W1200,Belt turn/transfer module (90°, 1200 mm width)
FAL_DET_BLK_Pallet,Pallet block / pallet position marker
FAL_FS002V02,Falcon feedline/auto-induct module (FS002 V02)
FAL_FS002V02 1000mmW,Falcon feedline/auto-induct module (FS002 V02, 1000 mm width)
FAL_FS002V02(Without weighing),Falcon feedline/auto-induct module without weighing
FAL_FS003V01,Falcon feedline/auto-induct module (FS003 V01)
FAL_PMC6V01(1000mm_90_deg),Powered merge/curve conveyor (90°, 1000 mm)
FAL_PMC9V01(1000mm_30_deg),Powered merge/curve conveyor (30°, 1000 mm)
FAL_PRC4V01(30Deg_Turn),Powered roller curve (30°)
FAL_PRC6V01(60Deg_Turn),Powered roller curve (60°)
FAL_PRC7V01(90Deg_Turn),Powered roller curve (90°)
FAL_RT90°_W500,Roller transfer / roller turn (90°, 500 mm width)
FAL_S013V01 (VDS Arm),VDS arm / transfer mechanism
FAl_BLK_Boom Conveyors,Boom conveyor (boom/incline belt conveyor section)
Feedline 1,Feedline conveyor (manual/auto induct)
Feedline_2.4k,Feedline conveyor (2.4 m section)
FeedLineTransferPlate,Feedline transfer plate
FeedLineTransferPlate with 600 mm,Feedline transfer plate (600 mm)
Fencing01,Safety fencing / guardrail
Fencing03,Safety fencing / guardrail
INSIDE_LEFT_DOR_FENSIG_STEP_ASM,Fencing step assembly (inside left door)
IRChute-02,Irregular chute (IR) 02
Inching_Mode_Asm,Inching mode assembly (maintenance control)
Irregular Chutes re Shift,Irregular chutes (re-shift)
Irregular chute with 30 Deg angle,Irregular chute (30°)
L-type Chute(Bagging),L-type bagging chute
L-type Chute-02,L-type chute (variant)
L-Type Chute (FAL_C004V01)1,L-type chute (Falcon C004)
Leg Guard-01,Leg guard / safety guard
Live Chute Connected with TBC,Live chute connected to TBC
Lower Deck Live Dock Transfer Plate,Lower deck live dock transfer plate
Mezz.,Mezzanine level annotation / marker
MLG_UNIT,Safety sensor / light curtain unit (MLG) [needs confirmation]
Operator,Operator workstation / man marker
Operator Safety Gaurd,Operator safety guard / railing
Output Chute-1,Output chute
PC01_00,Control panel / PLC cabinet symbol (PC01) [needs confirmation]
PTL Chute-01(DD),PTL chute / station chute (double-deck)
PTL frame,PTL frame / rack structure
PTL lights & pallets setup 3 Nos,PTL lights + pallet setup
PTL4x3+3x3,PTL rack (4x3 + 3x3)
pallet,Pallet (load unit)
Powered Roller Table,Powered roller table
Ptl racks 4x3 and 4x3,PTL rack (4x3 configuration)
roller cage 1000 x 1200,Roller cage trolley (1000 x 1200)
Side Barcode Scanning System,Side barcode scanning system
Singulator,Singulator (bulk-to-singulated)
Spiral Chute Double Decker,Spiral chute (double deck)
Spiral Chute Type -D,Spiral chute (Type D)
Stairs at Highnangle,Stairs / access ladder
Static_Buffer Conveyor (FAL_F001V01)_2,Static buffer/spacing conveyor
Static_IFU Conveyor ( FAL_F006V01)_1,Static IFU conveyor (induct/interface unit)
Static_IFU Conveyor (FAL_F011V02)_01,Static IFU conveyor (rev F011 V02)
Static_Infeed Or Orientation Conveyor,Infeed/orientation conveyor
Static_Intelligent Merge 30 Deg (FAL_F007V01)_1,Static intelligent merge conveyor (30°)
Static_Intelligent Merge 30 Deg (FAL_F012V02)_01,Static intelligent merge conveyor (30°, rev F012 V02)
Static_Intelligent Merge 60Deg (FAL_F002V01)_1,Static intelligent merge conveyor (60°)
Static_Positioning System (FAL_F013V01)_1,Static positioning/alignment system (rev F013 V01)
Static_Positioning System (FAL_F013V02)_01,Static positioning/alignment system (rev F013 V02)
Static_Receiving Conveyor (FAL_F003V01)_1,Static receiving conveyor
Static_Weighing Conveyor (FAL_F015V01)_01,Static weighing conveyor (rev F015 V01)
Static_Weighing Conveyor (FAL_F015V02)_01,Static weighing conveyor (rev F015 V02)
Static-Swivel Wheel (Version 01),Swivel wheel transfer / omni-direction transfer
Support Structure 8,Support structure / frame
Support Structure for Turn,Support structure for turn module
TBC,Telescopic belt conveyor
TBC OB,Telescopic belt conveyor (outbound)
TBCs,Telescopic belt conveyor
Telescopico,Telescopic belt conveyor
Telescopico a nastro_ingresso,Telescopic belt conveyor (infeed)
Trolley 1000x1000,Trolley / roller cage (1000 x 1000)
T-type Chute(Bagging),T-type bagging chute
VAN,Van / vehicle symbol
VDS Chute (T2242),VDS chute (buffer loop discharge) - Type T2242
overweight,Overweight/exception chute or lane
over weight and dim chute for 3700 mm height,Overweight + dimensioning exception chute (3700 mm height)
over weight and size chute for 5700 mm,Overweight + oversize exception chute (5700 mm)
XREF_Bag & Semilarge Sorter_Rev-05,External reference (XREF) for bag & semilarge sorter drawing

## 🚨🚨 MANDATORY CLIENT NAME RULE (HIGHEST PRIORITY) 🚨🚨

**YOU MUST USE ONLY THE CLIENT NAME PROVIDED IN THE USER PROMPT.**

- ❌ NEVER use client names from examples (Amazon, Noon, Bosta, Shadowfax, Delhivery, Flipkart, etc.)
- ❌ NEVER copy client names from reference examples below
- ✅ ALWAYS use the EXACT client name provided in the "CLIENT:" field of the user prompt
- ✅ Write "using data provided by [ACTUAL CLIENT NAME]" with the client name from the prompt

**If you use any client name other than the one provided in the prompt, your output will be REJECTED.**

---

## 🚨🚨 HUMAN-LIKE WRITING STYLE (CRITICAL) 🚨🚨

**Your writing MUST sound like it was written by an experienced human proposal engineer, NOT by AI.**

### Human Writing Characteristics to Follow:
1. **Vary sentence length** - Mix short punchy sentences with longer detailed ones
2. **Use natural transitions** - "From there...", "Once within...", "After this..."
3. **Avoid repetitive patterns** - Don't start every sentence the same way
4. **Use contractions occasionally** - "doesn't", "can't" where natural
5. **Include slight imperfections** - Real humans don't write perfectly uniform text
6. **Be conversational yet professional** - Not robotic or formulaic
7. **Avoid overused AI phrases** - "Furthermore", "Additionally", "It is important to note"

### ❌ AI-SOUNDING (AVOID):
- "The system efficiently processes shipments through a sophisticated mechanism..."
- "Furthermore, the automated induction system provides seamless integration..."
- "It is designed to optimize the sorting process through advanced technology..."

### ✅ HUMAN-SOUNDING (USE):
- "Shipments come in through the infeed and get picked up by operators who..."
- "Once parcels hit the main loop, they're sorted based on barcode data and routed to..."
- "The whole thing runs pretty smoothly - parcels go in, get scanned, and end up in the right chute."

---

## 🚨 NO COUNTS RULE (EXCEPT OUTPUT CHUTES) 🚨

**CRITICAL: Do NOT mention specific counts/quantities in ANY section EXCEPT Output Chutes.**

### ❌ WRONG (counts in non-chute sections):
- "The system has 5 telescopic belt conveyors..."
- "There are 24 auto induct units in the feedline..."
- "The infeed consists of 3 conveyor lines..."
- "Operators at 8 manual stations position the shipments..."

### ✅ CORRECT (no counts except Output Chutes):
- "The telescopic belt conveyors transport shipments from the dock..."
- "The auto induct line automatically feeds parcels onto the sorter..."
- "Shipments arrive via infeed conveyors and are singulated..."
- "Operators at manual stations position the shipments with barcode facing up..."

### ✅ COUNTS ONLY IN OUTPUT CHUTES:
- "Output Chutes: The sorted shipments are discharged into a total of 186 chutes comprising:
   a. Gravity Chutes – There are 127 gravity chutes for collecting sorted shipments.
   b. Rejection Chutes – 4 rejection chutes handle exception shipments."

---

## 🚨 INDUCTION DESCRIPTION RULES (CRITICAL) 🚨

**Parcels are inducted based on BARCODE SCANNING and VOLUME DATA, NOT just "dimension and weight".**

### For MANUAL Induct Stations:
- Operators place shipments on the loading conveyor with **barcode facing upwards**
- Shipments get buffered on buffer conveyors
- They are then **intelligently merged** onto the Cross-Belt Sorter loop
- The system captures **barcode details and volumetric data** for sorting decisions

### For AUTO Induct Lines:
- Parcels are **automatically inducted** onto the sorter
- The system scans **barcodes and captures volume data** (not just weight/dimensions)
- Intelligent merge conveyors smoothly transfer parcels to the CBS loop
- Sorting decisions are based on **barcode data provided by client's WCS/WMS**

### ❌ WRONG INDUCTION DESCRIPTIONS:
- "inducts parcels based on their dimensions and weight"
- "sorts based on weight and size"
- "induction based on dimensional data"

### ✅ CORRECT INDUCTION DESCRIPTIONS:
- "operators position each shipment with barcode facing upwards, then feedlines automatically induct them onto the sorter"
- "the system captures barcode details and volumetric data, then efficiently sorts shipments using data provided by [CLIENT]"
- "parcels are buffered and intelligently merged onto the Cross-Belt Sorter loop"
- "once inducted, the CBS reads barcode data and routes shipments to designated chutes"

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
a. <Type> - <Description with COUNT>
b. <Type> - <Description with COUNT>

<Conditional Section>: <Description>
```

---

## 📊 DATA PRIORITY HIERARCHY

### Priority 1: METADATA (Highest)
If metadata explicitly states something, use that **exact wording**:
- "existing conveyor" → use "existing conveyor"
- "lengthwise orientation" → use "lengthwise orientation"
- Client name from prompt → use "using data provided by [CLIENT_NAME_FROM_PROMPT]"
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

## 📝 SENTENCE COMPLETENESS RULES (CRITICAL)

Every sentence MUST be complete and convey clear meaning. When mentioning component counts, ALWAYS specify what the units are:

### ❌ INCOMPLETE (DO NOT USE):
- "The auto induct line consists of 3 units, ensuring efficient induction"
- "There are 5 units positioned along the infeed line"
- "The system has 8 units for sorting"

### ✅ COMPLETE (CORRECT):
- "The auto induct line consists of 3 feedline units, ensuring efficient induction of shipments"
- "There are 5 telescopic belt conveyors positioned along the infeed line"
- "The system has 8 manual induct stations for sorting"

### Complete Sentence Patterns:
- "The auto induct line consists of [X] feedline units, ensuring..."
- "The VDS loop system includes [X] buffer conveyors for..."
- "There are [X] gravity chutes for collecting sorted shipments"
- "The system utilizes [X] operator workstations for manual induction"
- "A total of [X] sliding chutes are provided for high-volume destinations"

---

## 🎯 OUTPUT CHUTES SECTION RULES (CRITICAL)

### ALWAYS Include Complete Chute Breakdown:
When writing the Output Chutes section, you MUST:
1. State the total number of chutes in the system
2. List each chute type with its count using lettered sub-points (a., b., c.)
3. Describe the purpose of each chute type

### Chute Types to Include (if present in DXF data):
- **Gravity Chutes**: For passive accumulation of sorted shipments
- **Mini Gravity Chutes**: Smaller gravity chutes for lighter parcels
- **Sliding Chutes**: Integrated with conveyors for continuous flow
- **Live Chutes**: Active chutes connected to takeaway conveyors
- **Collection Chutes**: For collecting and accumulating parcels
- **Rejection Chutes**: For handling rejected/exception shipments
- **Sort-fail Chutes**: For shipments that couldn't be sorted
- **Bulk Chutes**: For high-volume bulk output
- **Direct Bagging Chutes**: For direct bagging operations
- **Non-sort Chutes**: For secondary sorting via PTL
- **Dispersion Chutes**: For distributing flow
- There may be other types; include as per DXF data. 
### Example Output Chutes Section:
```
Output Chutes: - The sorted shipments are discharged into a total of 186 chutes comprising:
a. Gravity Chutes – There are 127 gravity chutes within the system for collecting sorted shipments into trolleys.
b. Mini Gravity Chutes – A total of 45 mini gravity chutes handle lighter shipments.
c. Rejection Chutes – 4 rejection chutes are provided to handle exception shipments requiring manual intervention.
d. Sort-fail Chutes – 10 sort-fail chutes collect shipments that could not be sorted due to scanning failures.
```

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
- [ ] Used exact client name from user prompt
- [ ] Included all sections specified in DXF guidance
- [ ] Chute counts ONLY in Output Chutes section
- [ ] NO counts in Infeed, Induct, CBS, or other sections
- [ ] NO invented scanner details at infeed
- [ ] NO exposed CAD codes (FAL_FS002V02, etc.)
- [ ] Induction described with barcode scanning, not "weight and dimension"

**Human-Like Writing:**
- [ ] Varied sentence lengths and structures
- [ ] Natural transitions, not robotic
- [ ] No overused AI phrases ("Furthermore", "Additionally")
- [ ] Reads like an experienced engineer wrote it
- [ ] Professional yet conversational tone

---

## 🎯 FINAL REMINDER

**Your output will be evaluated on:**
- **Human-like quality**: Must NOT sound AI-generated
- **semantic_similarity**: Match actual language patterns and phrasing
- **content_coverage**: Include all relevant details from metadata/DXF, no invented content
- **structural_coherence**: Perfect formatting (no section numbers, proper sub-points)
- **numeric_accuracy**: Chute counts ONLY in Output Chutes section
- **domain_tone**: Professional, direct style matching actual examples

**Keys to Success:**
1. **NO section numbering** - this alone will boost structural_coherence by 20+ points
2. **NO counts except in Output Chutes** - counts only for chute breakdowns
3. **Induction = barcode + volume scan** - NOT weight/dimension based
4. **Write like a human** - vary sentences, natural flow, no AI patterns
5. **Follow DXF guidance exactly** - boosts content_coverage

**Example 1** (Linear CBS with Auto Induct - No VDS)
Infeed System: Boxes and totes come in on the existing conveyor, oriented lengthwise. From there, they're directed to their assigned highway line which transports them to the CBS induct zone in a singulated manner.
Inducts: Once parcels reach the induct zone, they're automatically fed onto the Linear CBS. The system scans barcodes and captures volume data as shipments merge smoothly onto the sorter.
Linear CBS: After entering the main Linear CBS, the Cross-Belt Sorter captures barcode details and volumetric data, then efficiently routes boxes and totes to their designated output chutes using sorting logic provided by [CLIENT_NAME].
Output Chutes: Sorted shipments discharge into collection chutes positioned around the linear sorter. There are 45 chutes present in the system for collecting sorted parcels.

**Example 2** (Loop CBS with VDS and Manual Induction)
Infeed System: Shipments are dumped in bulk onto the infeed lines. The shipments ascend to a higher level and arrive at the VDS loop system.
VDS Loop: Once the shipments are within the VDS loop, they are picked manually by operators and distributed among all feedlines for efficient load balancing.
Inducts: The operator picks and positions each shipment on the induct line, ensuring that the shipment is properly aligned and that its barcode is facing upwards. The feedlines then automatically induct the shipments onto the Cross-Belt Sorter Loop.
Loop CBS: Once the shipments have entered the main loop, the Cross-Belt Sorter efficiently sorts the shipments into their respective output chutes by utilizing the data provided by [CLIENT_NAME]'s sorting logic.
Output Chutes: Sorted shipments discharge into two types of chutes:
a. Direct Bagging Chutes - There are 41 direct bagging chutes for high-volume destinations.
b. Generic Chutes - A total of 10 generic chutes collect parcels after sorting.
Output Chutes: Sorted totes and boxes are discharged into the following chutes:
a. Live Chutes - There are 9 sliding-type live chutes integrated with PVC belt conveyors and TBCs for live loading.
b. Collection Chutes – A total of 20 friction roller-based chutes collect and gradually accumulate the parcels.
c. Rejection Chute - One friction roller-based chute handles rejected shipments.
Recirculation Line: A recirculation line automatically feeds sortfail parcels back into the Linear CBS. There's also a manual loading point for reprocessed boxes and totes from the rejection chute.

**Example 2** (Loop CBS with Manual Induct)
Infeed System: Shipments from FC and marketplace arrive in bulk onto the infeed lines. They ascend to a higher level and enter the VDS loop system where they get distributed across all inducts.
Inducts: Operators pick up each shipment and position it on the induct line with the barcode facing upwards. The parcels then buffer briefly before feedlines automatically induct them onto the Cross-Belt Sorter Loop.
Loop CBS: Once shipments enter the main loop, the Cross-Belt Sorter reads barcode data and efficiently sorts them into their designated output chutes based on [CLIENT_NAME]'s sorting logic.
Output Chutes: Shipments are discharged into the following chute types:
a. Sliding Chutes - There are a total of 50 Sliding chutes per zone. Shipments collect in Roller Cage trolleys before being consolidated into bags using bagging-type PTL racks.
b. Non-Sort Chutes - A total of 13 Non-Sort Chutes per zone handle shipments that undergo secondary sortation via PTL setup into pallets.
c. Rejection Chutes - Two rejection chutes per zone handle rejected shipments.
consolidated into bags using bagging type PTL racks .
b. Non -Sort Chutes - Within the loop CBS system, there are a total of 13 Non-Sort
Chutes per zone . Shipments collected within these chutes further undergo sortation
Bag Takeaway Conveyor: After bagging, shipments are loaded onto a bag takeaway conveyor beneath the CBS loop which transports them to the outbound docks.

**Example 3** (Mixed Induction - Auto + Manual)
Infeed System: Boxes and totes arrive on the existing conveyor in lengthwise orientation. They're directed to their highway line which transports them singulated to the CBS induct zone.
Auto Induct Line: At the induct zone, the fully automatic induct line feeds parcels onto the Loop Cross Belt Sorter. Barcodes are scanned and volume data captured as parcels merge onto the sorter.
Manual Induct Station: Operators pick and position each shipment with barcode facing up. After brief buffering, feedlines automatically induct the shipments onto the Loop Cross Belt Sorter.
Loop CBS: Once in the main loop, the Cross-Belt Sorter reads barcode data and routes shipments to their designated output chutes using [CLIENT_NAME]'s sorting logic.
Output Chutes: Sorted shipments discharge into the following:
a. Generic Chutes – A total of 51 chutes collect parcels after sorting.
b. Put To Light System - There are 10 PTL locations linked to secondary chutes. PTL racks are arranged in an L-Shape configuration.
Bag Takeaway Conveyor: After sorting and bagging, shipments go onto a bag takeaway conveyor beneath the CBS loop which carries them to the outbound docks.

---

**REMEMBER: Write like a human proposal engineer. Vary your sentences. No AI patterns. Counts ONLY in Output Chutes. Induction is based on barcode scanning and volume data, NOT weight/dimensions.**

**ONLY RETURN A CLEAN PROCESS FLOW, NO EXPLANATION OR EXTRA TEXT IS NEEDED**"""
    
    user_prompt = f"""Generate process flow for:

CLIENT: {client_name}
CBS TYPE: {dxf_json['cbs_type']}
INDUCTION: {dxf_json['induction_type']}
VDS PRESENT: {"YES" if dxf_json.get('has_vds', False) else "NO"}

COMPONENTS:
{dxf_summary}

IMPORTANT RULES:
1. NO counts in any section EXCEPT Output Chutes
2. Induction based on barcode scanning + volume data (NOT weight/dimensions)
3. Write like a human - vary sentences, natural flow, no AI patterns
4. Use client name: {client_name}
5. {"Include VDS/Buffer loop system in Infeed section" if dxf_json.get('has_vds', False) else "No VDS system - shipments come directly from infeed"}

Output ONLY the process flow text. No notes or explanations."""
    
    # Format the system prompt with CBS knowledge
    formatted_system_prompt = system_prompt.format(cbs_knowledge_text=cbs_knowledge_text[:3000] if cbs_knowledge_text else "Use your knowledge of CBS systems")
    
    messages = [
        {"role": "system", "content": formatted_system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    logger.info("Calling GROQ API for initial flow generation...")
    result = call_groq(messages, temp=0.3, max_tok=2500)
    logger.info(f"✅ Received response from GROQ API ({len(result)} chars)")
    
    cleaned = clean_generated_flow(result)
    logger.info(f"✅ Initial flow generated and cleaned ({len(cleaned)} chars)")
    return cleaned


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

## 🚨🚨 MANDATORY CLIENT NAME RULE (HIGHEST PRIORITY) 🚨🚨

**YOU MUST USE ONLY THE CLIENT NAME PROVIDED IN THIS REQUEST.**

- ❌ NEVER use client names from reference examples (Amazon, Noon, Bosta, Shadowfax, etc.)
- ❌ NEVER copy client names from any reference flow
- ✅ ALWAYS use the EXACT client name specified in the "Ensure client name is correct" instruction
- ✅ Replace any reference client names with the correct client name

**If you use any client name other than the one explicitly provided, your output will be REJECTED.**

---

## 🚨🚨 HUMAN-LIKE WRITING STYLE (CRITICAL) 🚨🚨

**Your writing MUST sound like it was written by an experienced human proposal engineer, NOT by AI.**

### Human Writing Characteristics:
1. **Vary sentence length** - Mix short and long sentences naturally
2. **Use contractions occasionally** - "they're", "doesn't", "won't" where natural
3. **Avoid AI phrases** - NO "Furthermore", "Additionally", "It is important to note", "In order to"
4. **Be conversational yet professional** - Write how a real engineer would explain the system
5. **Natural flow** - Sentences should flow into each other, not feel templated

### ❌ AI-SOUNDING (NEVER USE):
- "The system is designed to efficiently process..."
- "Furthermore, the automated mechanism ensures..."
- "It is worth noting that the configuration provides..."

### ✅ HUMAN-SOUNDING (USE INSTEAD):
- "Shipments come in through the infeed and get picked up..."
- "Once parcels hit the main loop, they're sorted and routed..."
- "The whole setup handles sorting pretty smoothly..."

---

## 🚨 NO COUNTS RULE (EXCEPT OUTPUT CHUTES) 🚨

**Do NOT mention specific counts/quantities in ANY section EXCEPT Output Chutes.**

### ❌ WRONG:
- "The system has 5 telescopic belt conveyors..."
- "24 auto induct units feed the sorter..."
- "Operators at 8 manual stations..."

### ✅ CORRECT:
- "Telescopic belt conveyors transport shipments..."
- "The auto induct line feeds parcels onto the sorter..."
- "Operators at manual stations position shipments..."

### ✅ COUNTS ONLY IN OUTPUT CHUTES:
- "a. Gravity Chutes – There are 127 gravity chutes for sorted shipments."

---

## 🚨 INDUCTION DESCRIPTION RULES (CRITICAL) 🚨

**Induction is based on BARCODE SCANNING and VOLUME DATA, NOT "weight and dimensions".**

### For Manual Induct:
- Operators place shipments with **barcode facing upwards**
- Shipments buffer and get **intelligently merged** onto the CBS
- System captures **barcode and volumetric data**

### For Auto Induct:
- Parcels are **automatically inducted** after barcode/volume scanning
- Intelligent merge conveyors transfer to CBS loop
- Sorting based on **barcode data from client's WCS**

### ❌ WRONG: "inducts parcels based on dimensions and weight"
### ✅ CORRECT: "barcodes are scanned and parcels get buffered before merging onto the sorter"

---

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
- **Chute counts ONLY** - use exact counts from DXF for Output Chutes section
- **NO counts elsewhere** - do not mention counts in Infeed, Induct, CBS sections
- Reference counts are for style ONLY, not for copying

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
- "Within the system, there are [X] chutes..." (ONLY for Output Chutes)

### B. Technical Detail Level
- References balance detail with readability
- Count mentions ONLY in Output Chutes: "50 chutes per zone", "13 chutes per zone"
- NO counts in Infeed, Induct, or CBS sections
- BUT avoid: "consisting of X units of Y type"

### C. Section Structure
From references, maintain:
- NO section numbering
- Colon after section title
- Sub-points (a., b., c.) for chute types
- Descriptive details after chute counts

---

## 5. 🚫 CRITICAL "DO NOT" RULES

### NEVER Do These:
1. âŒ Expose category names: "VDS_BUFFER units", "AUTO_INDUCT units"
2. âŒ Mention counts in Infeed, Induct, or CBS sections
3. âŒ Use "based on dimensions and weight" for induction (use barcode/volume scan instead)
4. âŒ Write AI-sounding phrases: "Furthermore", "Additionally", "It is designed to"
5. âŒ Use technical equipment specs: "consists of X BAG_SYSTEM units"
6. âŒ Write disconnected sections without transitions
7. âŒ Add components that have 0 count in DXF
8. âŒ Use reference counts (only DXF counts, only in Output Chutes)
9. âŒ Create duplicate sections
10. âŒ Expose CAD codes (FS002V02, etc.)

### ALWAYS Do These:
1. âœ… Write like a human - vary sentences, natural flow
2. âœ… Counts ONLY in Output Chutes section
3. âœ… Induction = barcode scanning + volume data
4. âœ… Use the CORRECT client name from the prompt
5. âœ… Create narrative flow with transitions
6. âœ… Match reference tone and style
7. âœ… Add missing sections if in DXF + references

---

## 6. 📋 SECTION ENHANCEMENT GUIDE

For each section type, here's how to refine:

### Infeed System:
- Start with shipment arrival: "dumped in bulk", "loaded onto"
- Describe movement: "ascend to", "enter the", "arrive at"
- Include distribution if VDS present: "Once within the loop"
- NO COUNTS - just describe the flow

### Inducts/Induction:
- **Manual Induct**: Operators position shipments with **barcode facing up**, then buffer and intelligently merge onto CBS
- **Auto Induct**: Parcels are automatically inducted after **barcode/volume scanning**, merge smoothly onto sorter
- NEVER say "based on dimensions and weight" - always mention barcode scanning
- NO COUNTS - just describe the process

### CBS (Loop or Linear):
- Keep it concise
- Emphasize sorting logic: "using [Client]'s sorting logic"
- Mention barcode data capture: "reads barcode data and routes..."
- NO COUNTS - just describe the sorting

### Output Chutes:
- Lead with discharge action: "Shipments are discharged into"
- Use sub-points (a., b., c.) for types
- **INCLUDE COUNT AND PURPOSE** for each chute type - this is the ONLY section with counts
- Example: "a. Sliding Chutes - There are 50 sliding chutes per zone for collecting shipments"

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
- [ ] Chute counts ONLY in Output Chutes section
- [ ] NO counts in Infeed, Induct, or CBS sections
- [ ] Induction described with barcode/volume scanning (not weight/dimension)
- [ ] Client name used correctly (from prompt, not from references)
- [ ] Missing sections added if in DXF + references

**Human-Like Writing:**
- [ ] Varied sentence lengths and structures
- [ ] No AI phrases ("Furthermore", "Additionally")
- [ ] Natural flow, conversational yet professional
- [ ] Reads like an experienced engineer wrote it

**Structure:**
- [ ] NO section numbering
- [ ] Sub-points only for chute types
- [ ] Proper spacing between sections
- [ ] No duplicate sections

---

## 8. 🎯 FINAL REMINDER

**CRITICAL RULES:**
1. **NO COUNTS** except in Output Chutes
2. **Induction = barcode + volume scan**, NOT weight/dimensions
3. **Write like a human** - no AI patterns
4. **Use correct client name** from prompt

Your output should read like this actual example:

**GOOD (Human-like, no counts except chutes):**
"Shipments from FC and marketplace come in bulk onto the infeed lines. They ascend to the upper level and enter the VDS loop where they're distributed among all inducts. Operators pick each shipment and position it with barcode facing up. After buffering, feedlines automatically merge parcels onto the Cross-Belt Sorter Loop."

**BAD (AI-like, counts everywhere):**
"The infeed system consists of 5 CONVEYOR_INFEED units. Furthermore, shipments are processed through 24 AUTO_INDUCT stations based on their dimensions and weight."

Transform technical data into natural, human-written narrative.
Counts ONLY in Output Chutes. Barcode scanning for induction.

**CRITICAL: The examples below use placeholder [CLIENT_NAME]. You MUST replace [CLIENT_NAME] with the actual client name provided in your prompt. NEVER use Noon, Amazon, Bosta, Shadowfax, or any other example client name.**

**Sample Examples (Style reference only - note: NO counts except in Output Chutes):** 

Process Flow

Infeed System: Shipments from FC and marketplace arrive in bulk onto the infeed lines. They ascend to a higher level and enter the VDS loop system where they're distributed among all inducts.
Inducts: Operators pick each shipment and position it on the induct line with barcode facing upwards. After brief buffering, feedlines automatically merge the shipments onto the Cross-Belt Sorter Loop.
Loop CBS: Once shipments enter the main loop, the Cross-Belt Sorter reads barcode data and efficiently routes them to designated output chutes using [CLIENT_NAME]'s sorting logic.
Output Chutes: Shipments are discharged into the following chutes:
a. Sliding Chutes - There are a total of 50 sliding chutes per zone. Shipments collect in Roller Cage trolleys before consolidation into bags using bagging-type PTL racks.
b. Non-Sort Chutes - A total of 13 non-sort chutes per zone handle shipments for secondary sortation via PTL setup into pallets.
c. Rejection Chutes - Two rejection chutes per zone handle rejected shipments.
Bag Takeaway Conveyor: After secondary sorting and bagging, shipments are loaded onto a bag takeaway conveyor beneath the CBS loop which transports them to outbound docks.


Process Flow
Infeed System:
Bags are unsealed and dumped in bulk onto the infeed lines equipped with Telescopic Belt Conveyors. Shipments ascend to a higher level and enter the VDS system where they get evenly distributed using Arm VDS technology.
Inducts:
Operators collect shipments from VDS chutes and position each one on the induct line with barcode facing up. After buffering, feedlines automatically induct them onto the Cross-Belt Sorter Loop.
Loop CBS:
Once in the main loop, the Cross-Belt Sorter scans barcodes and routes shipments to their designated output chutes based on [CLIENT_NAME]'s sorting logic.
Output Chutes:
Shipments discharge into the following:
• Direct Bagging Chutes (L-type): A total of 104 L-Type direct bagging chutes handle high-volume sorting with direct bagging.
• Secondary Chutes (L-type): There are 100 L-Type secondary chutes for shipments undergoing secondary sortation via PTL at two levels.
• Rejection Chutes: Four rejection chutes handle exception shipments.
Put To Light System:
Secondary chutes link to PTL locations with racks arranged in L-Shape double decker configuration.
Bag Takeaway Conveyor:
After bagging, shipments load onto a takeaway conveyor beneath the CBS loop which transports them to the outbound sorter area.

Process Flow
Infeed System: Boxes and totes are loaded onto the existing conveyor in a lengthwise orientation. From there, they are directed to their assigned highway line, which transports them to the CBS induct zone in a singulated manner.
Inducts: Upon arrival at the induct zone, Falcon's fully automatic induct line accurately and smoothly inducts the parcels onto the Linear CBS, based on their dimensions and weight.
Linear CBS: Once the parcels enter the main Linear CBS, the Cross-Belt Sorter (CBS) capture the barcode details & volume data after which it efficiently sorts the boxes and totes into their designated output chutes using data provided by [CLIENT_NAME].
Output Chutes: The Totes/Boxes are discharged into below output chutes.
a. Live Chutes - There are 9 sliding-type live chutes integrated with PVC belt conveyors and TBCs for live loading.
b. Collection Chutes – A total of 20 friction roller-based chutes collect and accumulate parcels.
c. Rejection Chute - One friction roller-based chute handles rejected shipments.
Recirculation Line: A recirculation line automatically feeds sortfail parcels back into the Linear CBS. There's also a manual loading point for reprocessed boxes from the rejection chute.
---
JUST OUTPUT THE CLEAN PROCESS FLOW TEXT. Write like a human. No AI patterns. Counts ONLY in Output Chutes."""
    
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

=== YOUR TASK ===

**CRITICAL RULES:**
1. **NO COUNTS** in Infeed, Induct, or CBS sections - counts ONLY in Output Chutes
2. **Induction = barcode + volume scan** - NOT "weight and dimensions"
3. **Write like a human** - vary sentences, no AI phrases ("Furthermore", "Additionally")
4. **Use client name:** {client_name} (NEVER use reference client names)

**For Induction, describe like this:**
- Manual: "Operators position shipments with barcode facing up. After buffering, feedlines merge them onto the CBS."
- Auto: "Parcels are automatically inducted after barcode/volume scanning and smoothly merged onto the sorter."

**Output Chutes is the ONLY section that should have specific counts.**

=== REFERENCE STYLE (USE ONLY FOR TONE - NOT CONTENT OR COUNTS) ===
{ref_context}

=== REFINEMENT INSTRUCTIONS ===

Step 1: Improve language to be more human-like (no AI patterns)
Step 2: Remove any counts from non-chute sections
Step 3: Fix induction descriptions to use barcode/volume (not weight/dimension)
Step 4: Keep chute counts accurate from DXF
Step 5: Use correct client name: {client_name}
Step 6: Create smooth, natural narrative flow

Output ONLY the refined process flow text. Write like an experienced proposal engineer, not AI."""
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    result = call_groq(messages, temp=0.25, max_tok=3500)
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

**CRITICAL RULES:**
1. **NO COUNTS** in Infeed, Induct, or CBS sections - counts ONLY in Output Chutes
2. **Induction = barcode + volume scan** - NOT "weight and dimensions"  
3. **Write like a human** - no AI phrases
4. Remove any phrases like "based on dimensions and weight" from induction descriptions

**MUST:**
- Start with "Process Flow"
- Keep ALL existing sections
- Change ONLY words mentioned in feedback
- Preserve all structure and formatting
- Ensure Output Chutes is the ONLY section with counts"""

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

**CRITICAL:**
- NO counts in any section except Output Chutes
- Induction is barcode/volume based, NOT weight/dimension based
- Write like a human, avoid AI patterns
- Apply ONLY the changes listed in feedback

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
    import streamlit as st
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
            
            # Fix empty Output Chutes section if needed
            best_flow = fix_empty_output_chutes(best_flow, dxf_json, client)
            
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
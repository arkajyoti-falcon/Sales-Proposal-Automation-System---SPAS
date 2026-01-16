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
    Use AI to generate intelligent, context-aware feedback for improvement.
    Focuses on structural coherence and style matching.
    """
    
    # Extract structural information
    ref_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', reference, re.MULTILINE)
    gen_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', generated, re.MULTILINE)
    
    ref_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', reference, re.MULTILINE))
    gen_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', generated, re.MULTILINE))
    
    ref_para_count = len(re.split(r'\n\s*\n', reference))
    gen_para_count = len(re.split(r'\n\s*\n', generated))
    
    # Get first few sentences from reference as examples
    ref_sentences = [s.strip() for s in re.split(r'[.!?]+', reference) if len(s.strip()) > 20][:8]
    
    system_prompt = f"""You are an expert technical writing evaluator specializing in Cross-Belt Sorter (CBS) process flow documentation.

## YOUR TASK

Analyze the generated flow against the reference and provide 3-5 actionable feedback items that will improve the **Structural Coherence score**.

**Current Structural Coherence:** {current_score:.1f}/100
**Target Score:** {target_score}/100
**Gap:** {target_score - current_score:.1f} points

## WHAT IS STRUCTURAL COHERENCE?

Structural Coherence (0-100) measures:
1. **Section Order Match** (45% weight) - Are sections in the same order?
2. **Formatting Match** (35% weight) - Bullets, paragraphs, numbering
3. **Flow Smoothness** (20% weight) - Transitions and sentence rhythm

## STRUCTURAL ANALYSIS

**Reference has {len(ref_sections)} sections:** {', '.join(ref_sections[:5])}
**Generated has {len(gen_sections)} sections:** {', '.join(gen_sections[:5])}

**Reference formatting:**
- Uses lettered sub-points (a., b., c.): {'YES' if ref_has_letters else 'NO'}
- Paragraph count: {ref_para_count}

**Generated formatting:**
- Uses lettered sub-points (a., b., c.): {'YES' if gen_has_letters else 'NO'}
- Paragraph count: {gen_para_count}

## FEEDBACK GUIDELINES

Generate feedback that:
1. **Prioritizes structural issues** - section order, formatting, structure (these have highest impact on score)
2. **Is specific and actionable** - "Change X to Y" not "Improve X"
3. **References the reference document** - "Reference uses..." or "Match reference by..."
4. **Focuses on score improvement** - explain HOW the change will improve coherence
5. **Limits to 3-5 items** - most critical issues only
6. **No generic advice** - every item must be specific to this comparison

## FEEDBACK PRIORITY ORDER

**Priority 1: Structural Issues** (Fix these first - highest score impact)
- Section order mismatch
- Missing or extra sections
- Section title differences

**Priority 2: Formatting Issues** (Medium score impact)
- Bullet point style (a., b., c. vs • vs -)
- Paragraph structure mismatch
- Sub-point formatting

**Priority 3: Flow Issues** (Lower score impact, but important)
- Missing transition phrases
- Sentence structure differences
- Key phrase omissions

## OUTPUT FORMAT

Return feedback as a JSON array of strings. Each item should be:
- One sentence or two max
- Specific and actionable
- Focused on structural/formatting changes

Example format:
[
  "Section order mismatch: Move 'Inducts' section before 'Auto Induct Line' to match reference structure (positions 2 and 3 are swapped)",
  "Use lettered sub-points (a., b., c.) for Output Chutes section instead of bullet points (•) to match reference formatting",
  "Add 'Bag Takeaway Conveyor' section at the end - reference has this as final section but yours is missing it"
]

**CRITICAL:** Return ONLY the JSON array, no other text."""

    user_prompt = f"""Compare these two flows and generate 3-5 actionable feedback items to improve structural coherence.

## REFERENCE FLOW (TARGET STYLE)
```
{reference[:1200]}
```

Reference example sentences (for style):
{chr(10).join(f'• {s}' for s in ref_sentences[:5])}

---

## GENERATED FLOW (TO BE IMPROVED)
```
{generated[:1200]}
```

---

## DXF CONTEXT
Client: {dxf_json.get('client', 'Unknown')}
CBS Type: {dxf_json.get('cbs_type', 'Unknown')}
Total Components: {dxf_json.get('total_components', 0)}

---

Generate 3-5 specific, actionable feedback items that will increase Structural Coherence from {current_score:.1f} to {target_score}.

Focus on:
1. Section order/structure differences
2. Formatting mismatches (bullets, paragraphs)
3. Missing/extra sections

Return ONLY a JSON array of feedback strings."""

    try:
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        result = call_groq(messages, temp=0.3, max_tok=800)
        
        # Try to parse as JSON
        # Clean up the response - remove markdown code blocks if present
        result_clean = result.strip()
        result_clean = re.sub(r'^```json\s*', '', result_clean)
        result_clean = re.sub(r'^```\s*', '', result_clean)
        result_clean = re.sub(r'\s*```$', '', result_clean)
        result_clean = result_clean.strip()
        
        feedback_list = json.loads(result_clean)
        
        if isinstance(feedback_list, list) and len(feedback_list) > 0:
            # Limit to 5 items max
            return feedback_list[:5]
        else:
            logger.warning("AI feedback did not return a valid list")
            return ["Continue refining the flow structure to match reference formatting"]
            
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse AI feedback as JSON: {e}")
        logger.error(f"Raw response: {result}")
        # Fallback: try to extract lines that look like feedback
        lines = [line.strip() for line in result.split('\n') if line.strip() and len(line.strip()) > 20]
        if lines:
            return lines[:5]
        return ["AI feedback generation failed - continue with structural improvements"]
        
    except Exception as e:
        logger.error(f"Error generating AI feedback: {e}")
        return [f"Error generating feedback: {str(e)}"]


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
    system_prompt  = """## ROLE
You are a senior SALES engineer presenting the "Process Flow of the System" to a potential client. You're not just describing—you're SELLING how this solution transforms their operations.

## SALES-FIRST MINDSET (CRITICAL)
- This is a SALES document, not a technical manual
- Every step should answer: "Why does this matter to the client?"
- Highlight BENEFITS: speed, accuracy, efficiency, reduced errors, labor savings
- Make the client visualize parcels flowing SMOOTHLY through their new system

## WHY + WHAT (Always explain WHY, not just WHAT)
- DON'T: "Parcels are inducted onto the sorter"
- DO: "Parcels are smoothly inducted, ensuring zero jams and maximum throughput"

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

### Induction + VDS Logic (MANDATORY)
- If VDS/BUFFER is YES **and** induction includes AUTO: explicitly state that operators place shipments on the loading conveyor with the barcode facing up, shipments are buffered in the VDS loop, and then intelligently merged onto the Cross Belt Sorter loop via auto induct.
- If VDS/BUFFER is NO **and** induction includes AUTO: explicitly state that shipments coming from the infeed conveyors are automatically inducted onto the Loop Cross Belt Sorter (no VDS buffer involved).
- Keep language natural (no category codes like VDS_BUFFER/AUTO_INDUCT); just describe the behavior.

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
VDS/BUFFER: {'YES' if dxf_json.get('has_vds') else 'NO'}

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
    
    # Identify missing components
    missing = []
    for cat, count in dxf_cats.items():
        if count > 0:
            cat_lower = cat.lower()
            if cat_lower not in initial_flow.lower():
                ref_mentions = any(cat_lower in ref["process_flow"].lower() 
                                 for ref in reference_chunks)
                if ref_mentions:
                    missing.append(f"{cat} ({count} units)")
    
    user_prompt = f"""
=== INITIAL FLOW ===
{initial_flow}

=== DXF COMPONENTS ===
{create_dxf_summary(dxf_json)}

=== MISSING COMPONENTS ===
{', '.join(missing) if missing else 'None'}

=== REFERENCE STYLE (MATCH THIS EXACTLY) ===
{ref_context}

TASK: Refine the initial flow to match the reference style exactly.
- Use reference language patterns and phrasing
- Add missing components if any
- Match the tone and structure of references

Output ONLY the refined process flow text. No notes or explanations."""
    
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    result = call_groq(messages, temp=0.15, max_tok=3500)
    return clean_generated_flow(result)

# ============================================================================
# BREAKING CHANGE DETECTION
# ============================================================================

def detect_breaking_changes(original: str, refined: str) -> List[str]:
    """
    Detect if refinement introduced breaking changes that must be reverted.
    Returns list of breaking changes detected (empty if safe).
    """
    breaking_issues = []
    
    # 1. Check CBS type hasn't changed
    original_cbs_match = re.search(r'(Loop CBS|Linear CBS|Straight CBS)', original, re.IGNORECASE)
    refined_cbs_match = re.search(r'(Loop CBS|Linear CBS|Straight CBS)', refined, re.IGNORECASE)
    if original_cbs_match and refined_cbs_match:
        if original_cbs_match.group(1).lower() != refined_cbs_match.group(1).lower():
            breaking_issues.append(f"CRITICAL: CBS type changed from '{original_cbs_match.group(1)}' to '{refined_cbs_match.group(1)}'")
    
    # 2. Check component counts haven't been removed/changed drastically
    count_patterns = [
        (r'(\d+)\s+(?:manual\s+)?induct\s+station', 'induct stations'),
        (r'(\d+)\s+output\s+chute', 'output chutes'),
        (r'(\d+)\s+(?:live\s+)?chute', 'live chutes'),
        (r'(\d+)\s+PTL', 'PTL locations'),
        (r'(\d+)\s+(?:direct\s+)?bagging?\s+chute', 'bagging chutes'),
    ]
    
    for pattern, label in count_patterns:
        orig_match = re.search(pattern, original, re.IGNORECASE)
        refined_match = re.search(pattern, refined, re.IGNORECASE)
        
        if orig_match and not refined_match:
            breaking_issues.append(f"CRITICAL: Component count removed - {label}: {orig_match.group(1)} disappeared")
        elif orig_match and refined_match:
            if orig_match.group(1) != refined_match.group(1):
                breaking_issues.append(f"WARNING: {label} count changed from {orig_match.group(1)} to {refined_match.group(1)}")
    
    # 3. Check sections haven't been removed
    original_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', original, re.MULTILINE)
    refined_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', refined, re.MULTILINE)
    
    for sec in original_sections:
        if not any(s.lower() == sec.lower() for s in refined_sections):
            if sec.lower() in ['infeed', 'induct', 'loop', 'output chute', 'chute']:  # Critical sections
                breaking_issues.append(f"CRITICAL: Section removed - '{sec}'")
    
    # 4. Check for unexplained new technical terms
    new_terms = {'singulator', 'telescopic', 'double decker', 'vds', 'put to light', 'ptl', 'tbcs'}
    original_lower = original.lower()
    refined_lower = refined.lower()
    
    for term in new_terms:
        if term not in original_lower and term in refined_lower:
            # Only flag if it's a major introduction (appears 2+ times)
            count = refined_lower.count(term)
            if count >= 2:
                breaking_issues.append(f"WARNING: New technical term '{term}' introduced {count} times (not in original)")
    
    return breaking_issues

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
    Step 3/4: Iteratively refine based on AI-generated evaluation feedback.
    Enhanced with:
    - Ultra-conservative approach (micro changes only)
    - Breaking change detection and prevention
    - Score regression protection
    """
    
    # Calculate gap to target
    current_coherence = evaluation['structural_coherence']
    gap_to_target = target_score - current_coherence
    
    # Extract established patterns from current flow (these are SACRED)
    cbs_type_match = re.search(r'(Loop CBS|Linear CBS|Straight CBS)', current_flow, re.IGNORECASE)
    established_cbs_type = cbs_type_match.group(1) if cbs_type_match else "CBS"
    
    component_counts = {}
    count_patterns = [
        (r'(\d+)\s+(?:manual\s+)?induct\s+station', 'induct_stations'),
        (r'(\d+)\s+output\s+chute', 'output_chutes'),
        (r'(\d+)\s+(?:live\s+)?chute', 'live_chutes'),
        (r'(\d+)\s+PTL', 'ptl_locations'),
    ]
    for pattern, key in count_patterns:
        match = re.search(pattern, current_flow, re.IGNORECASE)
        if match:
            component_counts[key] = match.group(1)
    
    # Format AI-generated feedback (REDUCED TO 2-4 items)
    feedback_text = ""
    feedback_items = evaluation.get("feedback", [])
    
    # Limit to max 4 feedback items to prevent overwhelming refinement
    feedback_items = feedback_items[:4]
    
    if feedback_items:
        feedback_text = "🎯 MICRO-IMPROVEMENT SUGGESTIONS (Language Polish Only):\n\n"
        for i, fb in enumerate(feedback_items, 1):
            feedback_text += f"{i}. {fb}\n"
        feedback_text += "\nIMPORTANT: These are MICRO suggestions only. Keep all component counts, CBS type, and section names unchanged."
    else:
        feedback_text = "✅ No improvements needed. Flow is optimal quality."
    
    # Extract structural information for context
    ref_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', reference_sample, re.MULTILINE)
    current_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', current_flow, re.MULTILINE)
    
    system_prompt = f"""You are an expert technical writer specializing in CBS process flow documentation.

## 🎯 ULTRA-CONSERVATIVE REFINEMENT - ITERATION {iteration}

**Current Score:** {current_coherence:.1f}/100 | **Target:** {target_score}/100

## 🔥 NON-NEGOTIABLE RULES

### ABSOLUTE PROTECTION (WILL REVERT IF BROKEN):
1. CBS Type: MUST STAY "{established_cbs_type}"
2. Component Counts: MUST STAY {json.dumps(component_counts)}
3. Section Names: MUST STAY {json.dumps(current_sections)}
4. Content Details: ALL MUST REMAIN - no removals

### ALLOWED CHANGES (Micro only):
- ✅ Synonym changes: "rapidly" → "efficiently"
- ✅ Phrase additions: Insert transition words
- ✅ Voice adjustments: Slight active/passive tweaks
- ✅ Phrasing: Reorder for flow (NOT meaning)

### FORBIDDEN CHANGES:
- ❌ CBS type changes
- ❌ Component count changes
- ❌ Detail removals
- ❌ New technical terms
- ❌ Structural rewrites
- ❌ Section reordering

## 🎯 MICRO-IMPROVEMENT SUGGESTIONS

{feedback_text}

## 📤 OUTPUT

Apply ONLY the micro-suggestions above.
Output ONLY the process flow text.
If ANY breaking change would occur, output the current flow unchanged."""

    user_prompt = f"""Current flow (Score {current_coherence:.1f}):
```
{current_flow}
```

Reference style:
```
{reference_sample[:1000]}
```

Micro-suggestions to apply:
{feedback_text}

CRITICAL: 
- Apply ONLY language micro-improvements
- Keep CBS type: {established_cbs_type}
- Keep component counts: {json.dumps(component_counts)}
- Keep all sections and details
- Return ONLY the flow text

If any breaking change would occur, return the current flow unchanged."""

    try:
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        result = call_groq(messages, temp=0.1, max_tok=3500)
        refined = clean_generated_flow(result)
        
        # CRITICAL: Check for breaking changes
        breaking_changes = detect_breaking_changes(current_flow, refined)
        if breaking_changes:
            logger.warning(f"⚠️ Breaking changes detected! Reverting refinement:")
            for change in breaking_changes:
                logger.warning(f"  - {change}")
            return current_flow  # Return original unchanged
        
        # Validate output
        is_valid, error_msg = validate_flow_quality(refined)
        if not is_valid:
            logger.warning(f"Generated flow failed validation: {error_msg}. Reverting.")
            return current_flow
        
        return refined
        
    except Exception as e:
        logger.error(f"Error in iterative refinement: {e}")
        return current_flow


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
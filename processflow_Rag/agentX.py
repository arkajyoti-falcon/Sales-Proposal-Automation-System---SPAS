"""
ITERATIVE PROCESS FLOW GENERATOR WITH EVALUATION
==================================================
Multi-step refinement with structural coherence scoring
Fixed: Better feedback focused on style/tone, not numbers
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
# CLEANING FUNCTIONS
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

# ============================================================================
# EVALUATION FUNCTIONS
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
            if re.match(r"^\s*[\-\*\u2022\da-z]+[\.\)]\s+", line.strip())
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
    ref_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-\u2013]?', reference, re.MULTILINE)
    gen_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-\u2013]?', generated, re.MULTILINE)
    
    if len(ref_sections) != len(gen_sections):
        feedback.append(f"Section count mismatch: Reference has {len(ref_sections)} sections, yours has {len(gen_sections)}. Match the reference structure.")
    
    # 2. Check for bullet point style (a., b., c. vs - or •)
    ref_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', reference, re.MULTILINE))
    gen_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', generated, re.MULTILINE))
    ref_has_dashes = bool(re.search(r'^\s*[-\u2022]\s+', generated, re.MULTILINE))
    
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


def evaluate_process_flow(generated: str, reference: str, dxf_json: dict) -> Dict:
    """
    Comprehensive evaluation of generated process flow.
    
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
    
    # 3. Component coverage check
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
    
    if missing_components:
        evaluation["feedback"].append(f"Missing major components: {', '.join(missing_components)}")
    
    # 4. Style matching analysis (MAIN FOCUS)
    style_feedback = analyze_style_differences(reference, generated_clean, dxf_json)
    evaluation["feedback"].extend(style_feedback)
    
    # 5. Style match score (based on feedback count)
    max_feedback_items = 8
    evaluation["style_match"] = max(0, 100 - (len(style_feedback) * (100 / max_feedback_items)))
    
    return evaluation

# ============================================================================
# GENERATION FUNCTIONS
# ============================================================================

def generate_initial_flow(client_name: str, dxf_json: dict) -> str:
    """Step 1: Generate initial flow from DXF data only."""
    
    dxf_summary = create_dxf_summary(dxf_json)
    # Get CBS knowledge and actually use it
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


def iterative_refinement(
    current_flow: str,
    dxf_json: dict,
    reference_sample: str,
    evaluation: Dict,
    iteration: int
) -> str:
    """
    Step 3/4: Iteratively refine based on evaluation feedback.
    Uses enhanced chain-of-thought prompting for systematic improvement.
    """
    
    # Format feedback with numbers
    feedback_text = "\n".join(f"{i+1}. {fb}" for i, fb in enumerate(evaluation.get("feedback", [])))
    
    # Extract structural analysis
    ref_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', reference_sample, re.MULTILINE)
    current_sections = re.findall(r'^([A-Z][A-Za-z\s]+):\s*[-–]?', current_flow, re.MULTILINE)
    
    # Build section comparison
    section_analysis = "REFERENCE SECTIONS:\n"
    for i, sec in enumerate(ref_sections[:7], 1):
        section_analysis += f"  {i}. {sec}\n"
    
    section_analysis += "\nYOUR CURRENT SECTIONS:\n"
    for i, sec in enumerate(current_sections, 1):
        section_analysis += f"  {i}. {sec}\n"
    
    # Identify section mismatches
    section_issues = ""
    for i in range(min(len(ref_sections), 5)):
        if i < len(current_sections):
            if ref_sections[i].lower().strip() != current_sections[i].lower().strip():
                section_issues += f"  ⚠️ Position {i+1}: You wrote '{current_sections[i]}' but reference uses '{ref_sections[i]}'\n"
        else:
            section_issues += f"  ⚠️ Position {i+1}: MISSING - Reference has '{ref_sections[i]}'\n"
    
    # Extract key phrases from reference (sentence by sentence)
    ref_sentences = [s.strip() for s in re.split(r'[.!?]+', reference_sample) if len(s.strip()) > 20][:10]
    key_phrase_examples = "\n".join(f"  • {sent}" for sent in ref_sentences[:5])
    
    # Analyze bullet/numbering style
    ref_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', reference_sample, re.MULTILINE))
    ref_has_bullets = bool(re.search(r'^\s*[•\-]\s+', reference_sample, re.MULTILINE))
    current_has_letters = bool(re.search(r'^\s*[a-z]\.\s+', current_flow, re.MULTILINE))
    current_has_bullets = bool(re.search(r'^\s*[•\-]\s+', current_flow, re.MULTILINE))
    
    bullet_guidance = ""
    if ref_has_letters and not current_has_letters:
        bullet_guidance = "⚠️ Reference uses lettered sub-points (a., b., c.) - YOU MUST USE THIS FORMAT"
    elif not ref_has_letters and current_has_letters:
        bullet_guidance = "⚠️ Reference does NOT use lettered sub-points - DO NOT USE (a., b., c.)"
    elif ref_has_bullets and not current_has_bullets:
        bullet_guidance = "⚠️ Reference uses bullet points (• or -) - YOU MUST USE THIS FORMAT"
    
    # Extract DXF numbers for emphasis
    dxf_cats = dxf_json.get("category_summary", {})
    chute_total = dxf_json.get("chute_analysis", {}).get("total", 0)
    
    dxf_numbers_text = ""
    if chute_total > 0:
        dxf_numbers_text += f"  • Total Chutes: {chute_total}\n"
    for cat, count in dxf_cats.items():
        if count > 5:  # Only major components
            dxf_numbers_text += f"  • {cat.replace('_', ' ').title()}: {count} units\n"
    
    system_prompt = f"""You are an expert technical writer specializing in Cross-Belt Sorter (CBS) process flow documentation.

## 🎯 YOUR MISSION (ITERATION {iteration})

Your current draft scores:
- Structural Coherence: {evaluation['structural_coherence']:.1f}/100 (TARGET: 85+)
- Style Match: {evaluation['style_match']:.1f}/100
- Component Coverage: {evaluation['component_coverage']:.1f}/100

You MUST systematically improve this draft by following a precise chain-of-thought process.

## 📋 CHAIN-OF-THOUGHT REFINEMENT PROCESS

### STEP 1: ANALYZE THE REFERENCE STRUCTURE
First, understand EXACTLY how the reference is structured:

{section_analysis}

{section_issues if section_issues else "✓ Section structure matches!"}

### STEP 2: MATCH THE FORMATTING STYLE
Check bullet point and numbering format:

{bullet_guidance if bullet_guidance else "✓ Bullet style matches!"}

### STEP 3: ADOPT THE REFERENCE LANGUAGE PATTERNS
Study how the reference writes sentences. Here are key examples:

{key_phrase_examples}

Notice:
- Sentence length and rhythm
- Use of active vs passive voice  
- Transition words ("Once...", "From there...", "Upon arrival...")
- Technical terminology level (formal but accessible)

### STEP 4: ADDRESS SPECIFIC FEEDBACK
You received {len(evaluation.get('feedback', []))} specific feedback items. Address EACH ONE:

{feedback_text if feedback_text else "No specific issues - focus on general polish"}

### STEP 5: VERIFY DXF DATA ACCURACY
These numbers are SACRED - use them EXACTLY:

{dxf_numbers_text}

## 🔧 YOUR SYSTEMATIC REFINEMENT TASK

Now, mentally go through your current draft line by line:

**FOR EACH SECTION:**
1. Does the section title match the reference? If not, change it.
2. Does the content use similar language patterns as the reference? If not, rewrite.
3. Are transitions smooth? Add "Once...", "From there...", etc.
4. Are DXF numbers present and correct? Verify.

**FOR BULLET POINTS:**
1. Does the format match reference (a., b., c. vs • vs -)?
2. Is the content structure similar (count + description)?

**FOR OVERALL FLOW:**
1. Read it aloud mentally - does it sound like the reference?
2. Is the tone professional yet accessible?
3. Are shipments/parcels tracked through their journey?

## ⚠️ CRITICAL EXECUTION RULES

1. **EXACT STRUCTURE MATCH**: Your sections must be in the SAME ORDER as reference
2. **EXACT STYLE MATCH**: Your sentences must SOUND LIKE the reference
3. **EXACT DATA MATCH**: Use DXF numbers precisely
4. **NO META-COMMENTARY**: Output ONLY the process flow text itself

DO NOT WRITE:
- "Here is the revised flow..."
- "Changes made:"
- "Note:"
- "The flow has been improved..."
- Any explanation or commentary

WRITE ONLY:
```
Process Flow

Infeed System: ...

[Next Section]: ...
```

## 🎯 QUALITY CHECKLIST (Mental Check Before Output)

Before you output, verify:
- [ ] Section order matches reference exactly
- [ ] Bullet style matches reference exactly  
- [ ] Language patterns sound like reference
- [ ] All feedback items addressed
- [ ] DXF numbers used correctly
- [ ] No meta-commentary included
- [ ] Smooth transitions between sections
- [ ] Professional yet accessible tone

## 📤 OUTPUT INSTRUCTION

Now generate the improved process flow. Output ONLY the clean process flow text starting with "Process Flow" - absolutely nothing else."""

    user_prompt = f"""## YOUR CURRENT DRAFT (TO BE IMPROVED)
```
{current_flow}
```

## TARGET REFERENCE (YOUR STYLE GUIDE)
```
{reference_sample[:1500]}
```

## DXF COMPONENT DATA (YOUR FACTS)
```
{create_dxf_summary(dxf_json)}
```

---

## 🚀 EXECUTE THE REFINEMENT NOW

Following the 5-step chain-of-thought process from the system prompt:

**STEP 1**: Verify your section structure matches the reference
**STEP 2**: Match the bullet/numbering format  
**STEP 3**: Adopt the reference language patterns
**STEP 4**: Address each feedback item systematically
**STEP 5**: Verify all DXF numbers are correct

Generate the improved process flow now. Output starts with "Process Flow" and contains ONLY the document text."""

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    result = call_groq(messages, temp=0.1, max_tok=3500)
    return clean_generated_flow(result)

# ============================================================================
# STREAMLIT UI
# ============================================================================

def main():
    st.set_page_config(page_title="Iterative Process Flow Generator", layout="wide")
    
    st.title("🔄 Iterative Process Flow Generator")
    st.markdown("**Multi-step refinement with style-focused evaluation**")
    
    # Sidebar for inputs
    with st.sidebar:
        st.header("📋 Input")
        uploaded = st.file_uploader("Upload DXF", type=["dxf"])
        client = st.text_input("Client Name", "Zepto")
        project = st.text_input("Project Name", "")
        
        max_iterations = st.slider("Max Iterations", 1, 10, 5)
        target_score = st.slider("Target Score", 50, 95, 85)
        
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
            }
            
            # ================================================================
            # STEP 0: DXF EXTRACTION
            # ================================================================
            with st.status("📊 Step 0: Extracting DXF Components...", expanded=True) as status:
                try:
                    dxf_json = extract_dxf_components(tmp_path, project or uploaded.name)
                    dxf_json['client'] = client  # Add client name for evaluation
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
            with st.status("✍️ Step 1: Generating Initial Flow...", expanded=True) as status:
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
            # STEP 3 & 4: EVALUATION & ITERATIVE REFINEMENT
            # ================================================================
            st.markdown("---")
            st.subheader("🔄 Iterative Refinement Loop")
            
            # Use best reference as target
            target_reference = references[0]["process_flow"] if references else current_flow
            
            iteration_num = 2
            best_score = 0
            best_flow = current_flow
            
            for iter_count in range(max_iterations):
                with st.expander(f"**Iteration {iteration_num}**", expanded=(iter_count == 0)):
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.markdown("##### 📊 Evaluation")
                        evaluation = evaluate_process_flow(
                            current_flow, 
                            target_reference,
                            dxf_json
                        )
                        
                        # Display scores
                        score_col1, score_col2 = st.columns(2)
                        with score_col1:
                            st.metric(
                                "Structural Coherence", 
                                f"{evaluation['structural_coherence']:.1f}",
                                delta=f"Target: {target_score}"
                            )
                            st.metric("Style Match", f"{evaluation['style_match']:.1f}")
                        with score_col2:
                            st.metric("Component Coverage", f"{evaluation['component_coverage']:.1f}")
                            st.metric("BERT F1", f"{evaluation['bert_f1']:.1f}")
                        
                        # Feedback
                        if evaluation['feedback']:
                            st.markdown("**🎯 Actionable Feedback:**")
                            for fb in evaluation['feedback']:
                                st.write(f"• {fb}")
                        else:
                            st.success("✅ No issues found - flow matches target!")
                    
                    with col2:
                        st.markdown("##### 📝 Current Flow")
                        st.text_area(
                            "Flow", 
                            current_flow, 
                            height=300, 
                            key=f"iter_{iteration_num}_flow",
                            label_visibility="collapsed"
                        )
                    
                    # Check if target reached
                    current_score = evaluation['structural_coherence']
                    if current_score > best_score:
                        best_score = current_score
                        best_flow = current_flow
                    
                    results["iterations"].append({
                        "iteration": iteration_num,
                        "stage": "Evaluated",
                        "flow": current_flow,
                        "evaluation": evaluation,
                        "score": current_score,
                    })
                    
                    if current_score >= target_score:
                        st.success(f"🎉 Target score reached: {current_score:.1f} >= {target_score}")
                        break
                    
                    # Generate next iteration
                    if iter_count < max_iterations - 1:
                        st.markdown("##### 🔄 Generating Next Iteration...")
                        try:
                            current_flow = iterative_refinement(
                                current_flow,
                                dxf_json,
                                target_reference,
                                evaluation,
                                iteration_num
                            )
                            iteration_num += 1
                        except Exception as e:
                            st.error(f"❌ Refinement failed: {e}")
                            break
            
            # Final results
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
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Final Score", f"{results.get('final_score', 0):.1f}/100")
        with col2:
            st.metric("Iterations", len(results['iterations']))
        with col3:
            st.metric("References Used", len(results.get('references', [])))
        
        # Tabs for different views
        tab1, tab2, tab3 = st.tabs(["📄 Final Flow", "📈 Iteration History", "🔍 DXF Analysis"])
        
        with tab1:
            st.text_area("Final Process Flow", results['final_flow'], height=500)
            st.download_button(
                "💾 Download Flow",
                results['final_flow'],
                file_name=f"{client}_process_flow.txt",
                mime="text/plain"
            )
        
        with tab2:
            for iter_data in results['iterations']:
                iter_num = iter_data['iteration']
                stage = iter_data['stage']
                
                with st.expander(f"Iteration {iter_num}: {stage}"):
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
        
        with tab3:
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
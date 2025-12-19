
import os
import io
import streamlit as st
import pandas as pd
import json
import re
import base64
import copy
import tempfile
import time
from io import BytesIO
from datetime import date, datetime
from typing import Optional, Tuple
from dataclasses import dataclass
from typing import Dict, List
from pathlib import Path
from collections import Counter, defaultdict
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from PIL import Image
from groq import Groq
from dotenv import load_dotenv
from docxcompose.composer import Composer
import pdfplumber
import ezdxf
import requests
import convertapi
import sys
from sentence_transformers import SentenceTransformer, util as st_util
from bert_score import score as bert_score
from Supportive_Functions.feedback_manager import (
    load_feedback_rules,
    save_feedback_rule,
    get_all_rules,
    deactivate_rule,
    update_rule,
    create_local_rule,
    get_rule_stats
)

load_dotenv()

# Import the standalone system-description module (we will reuse its deterministic
# DXF/costing extraction + GROQ prompt flow to produce the System Description)
import sys_desc_latest as sd_sys

# Ensure processflow_Rag modules are importable
sys.path.append(str(Path(__file__).parent / "processflow_Rag"))
import combine_old as combine_old_mod
from dxf_extractor import create_dxf_summary_for_embedding

# Import agentY functions for process flow generation
from agentY import (
    generate_initial_flow as agentY_generate_initial_flow,
    generate_second_flow_with_chunks,
    iterative_refinement,
    clean_generated_flow as agentY_clean_generated_flow,
    validate_flow_quality as agentY_validate_flow_quality,
    evaluate_process_flow as agentY_evaluate_process_flow
)

import torch
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ==================== SENTENCE MODEL CACHE ====================
@st.cache_resource
def load_sentence_model():
    """Load SentenceTransformer model for semantic similarity scoring."""
    return SentenceTransformer('all-MiniLM-L6-v2')

sentence_model = load_sentence_model()

# ==================== EVALUATION & VALIDATION FUNCTIONS ====================

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
    - Sentence order similarity (45% weight)
    - Paragraph structure (35% weight)
    - Bullet point usage
    - Transition smoothness (20% weight)
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
        
        result = combine_old_mod.call_groq(messages, temp=0.3, max_tok=800)
        
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
        evaluation["feedback"].append(f"Missing coverage: {', '.join(missing_components[:3])}")
    
    # 4. AI-GENERATED STYLE FEEDBACK (REPLACING STATIC RULES)
    # Only generate feedback if score is below target
    if evaluation["structural_coherence"] < target_score:
        ai_feedback = generate_ai_feedback(reference, generated_clean, dxf_json, evaluation["structural_coherence"], target_score)
        evaluation["feedback"].extend(ai_feedback)
    
    # 5. Style match score (based on feedback count and structural coherence)
    # If structural coherence is high and no feedback, style match should be high
    if evaluation["structural_coherence"] >= target_score and len(evaluation["feedback"]) == 0:
        evaluation["style_match"] = 100.0
    else:
        # Style match is correlation of structural coherence and low feedback count
        feedback_penalty = max(0, len(evaluation["feedback"]) - 1) * 5
        evaluation["style_match"] = max(0, min(100, evaluation["structural_coherence"] - feedback_penalty))
    
    return evaluation

# ==================== GROQ CLIENT SETUP ====================
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
CONVERTAPI_SECRET = os.getenv("CONVERTAPI_SECRET")

if not GROQ_API_KEY:
    st.sidebar.error("⚠️ GROQ_API_KEY not found in .env file!")
else:
    groq_client = Groq(api_key=GROQ_API_KEY)

if CONVERTAPI_SECRET:
    convertapi.api_credentials = CONVERTAPI_SECRET

# ==================== CLIENT LOGOS MAPPING ====================
CLIENT_LOGOS = {
    "Zepto": "FIXED_IMAGE/clients/zepto.png",
    "Flipkart": "FIXED_IMAGE/clients/flipkart.png",
    "Shiprocket": "FIXED_IMAGE/clients/shiprocket.png",
    "Amazon": "FIXED_IMAGE/clients/amazon.png",
    "Delhivery": "FIXED_IMAGE/clients/delhivery.png",
    "Swiggy": "FIXED_IMAGE/clients/swiggy.png",
    "Mondial": "FIXED_IMAGE/clients/mondial.jpg",
    "Zomato": "FIXED_IMAGE/clients/zomato.png",
}

# ==================== COMPONENT DESCRIPTION IMAGE PATHS ====================
COMPONENT_IMAGES = {
    "cross_belt_sorter": r"FIXED_IMAGE\CROS_BELT_SORTER.PNG",
    "cbs_carrier": r"FIXED_IMAGE\CBS_CAREER.PNG",
    "servo_roller": r"FIXED_IMAGE\SERVO_ROLLER.PNG",
    "chassis": r"FIXED_IMAGE\CHASIS.PNG",
    "wheel": r"FIXED_IMAGE\CAREER_WHEEL.PNG",
    "power": r"FIXED_IMAGE\TRANSMISSION.PNG",
    "linear": r"FIXED_IMAGE\LINEAR_MOTOR_DRIVE.PNG",
    "friction_wheel": r"FIXED_IMAGE\FRICTION_WHEEL_DRIVE.PNG",
    "rcoax": r"FIXED_IMAGE\DATA.PNG",
    "carrier_position": r"FIXED_IMAGE\CPS.PNG",
}

GLOSSARY_ENTRIES = [
    ("RFQ", "Request For Quotation"),
    ("RFP", "Request For Proposal"),
    ("PPH", "Parcels / Shipments Per Hour"),
    ("ARB", "Actuated Roller Balls"),
    ("DBO", "Damaged Barcode"),
    ("VDS", "Volume Distribution System"),
    ("ICR", "Intelligent Character Recognition"),
    ("MEZZ", "Mezzanine"),
    ("LIM", "Linear Induction Motor"),
    ("LSM", "Linear Synchronous Motor"),
    ("FWD", "Friction Wheel Drive"),
    ("ECDS", "Empty Carrier Detection System"),
    ("AC", "Alternating Current"),
    ("DC", "Direct Current"),
    ("PLC", "Programmable Logic Controller"),
    ("IT", "Information Technology"),
    ("BOQ", "Bill Of Quantity"),
    ("I/O", "Input/ Output"),
    ("PDP", "Power Distribution Panel"),
    ("PC", "Personal Computer"),
    ("UPS", "Uninterrupted Power Supply"),
    ("CBS", "Cross Belt Sorter"),
    ("MDR", "Motor Driven Roller"),
    ("IPP", "Individual Productivity Potential"),
    ("VM", "Virtual Machine"),
    ("MENA", "Middle East North Africa"),
    ("FOC", "Free of Cost"),
    ("CEP", "Courier Express Parcel"),
    ("DAP", "Design Approval Phase"),
    ("DNF", "Data Not Found"),
    ("LGM", "Logic Mismatch"),
    ("RTVC", "Real Time Video Coding"),
    ("NL Shipments", "Non-Large Shipments"),
    ("SL Shipments", "Semi-Large Shipments"),
    ("NO(S)", "Piece(s)"),
]

# ==================== RETRY WRAPPER FOR RATE LIMITS ====================
def call_groq_with_retry(api_call_func, max_retries=5, initial_delay=2):
    """Wrapper to retry GROQ API calls with exponential backoff on rate limits."""
    for attempt in range(max_retries):
        try:
            return api_call_func()
        except Exception as e:
            error_str = str(e)
            # Check if it's a rate limit error
            if "429" in error_str or "rate_limit_exceeded" in error_str.lower():
                if attempt < max_retries - 1:
                    # Extract wait time from error message if available
                    wait_match = re.search(r'try again in ([0-9.]+)s', error_str)
                    if wait_match:
                        wait_time = float(wait_match.group(1)) + 1  # Add 1 second buffer
                    else:
                        wait_time = initial_delay * (2 ** attempt)  # Exponential backoff
                    
                    st.warning(f"Rate limit hit. Waiting {wait_time:.1f}s before retry {attempt + 1}/{max_retries}...")
                    time.sleep(wait_time)
                else:
                    raise  # Re-raise on final attempt
            else:
                raise  # Re-raise non-rate-limit errors immediately
    
    raise RuntimeError(f"Failed after {max_retries} retries")

# ==================== DXF COMPONENT EXTRACTION ====================

UNITS = {
    0: "Unitless", 1: "inches", 2: "feet", 3: "miles",
    4: "millimeters", 5: "centimeters", 6: "meters", 7: "kilometers",
}

# Component categorization keywords
COMPONENT_PATTERNS = {
    "AUTO_INDUCT": [
        r"fal.*fs\d+", r"fal.*feed", r"fal.*induct", r"auto.*induct", 
        r"feed.*line", r"feedline", r"fs\d{3}"
    ],
    "CONVEYOR_INFEED": [
        r"telescopic", r"infeed.*conv", r"in.*feed", r"receiving.*conv",
        r"inclined.*conv", r"incline", r"elevation.*conv"
    ],
    "VDS_BUFFER": [
        r"vds", r"distribution.*loop", r"buffer", r"arm.*vds"
    ],
    "OPERATOR_STATION": [
        r"operator(?!.*safety)", r"manual.*station", r"induct.*station"
    ],
    "CHUTE": [
        r"^chute", r"gravity.*chute", r"live.*chute", r"slide.*chute",
        r"sliding.*chute", r"reject.*chute", r"collection.*chute",
        r"mini.*chute", r"bulk.*chute", r"discharge"
    ],
    "PTL": [
        r"ptl", r"put.*to.*light", r"pick.*to.*light", r"light.*rack"
    ],
    "BAG_SYSTEM": [
        r"bag.*conv", r"bag.*takeaway", r"bagging"
    ],
    "RECIRCULATION": [
        r"recirculation", r"recirculate", r"refeed", r"return.*conv"
    ],
    "CBS_SORTER": [
        r"cbs", r"cross.*belt", r"crossbelt", r"sorter.*module",
        r"carrier", r"loop.*sorter"
    ],
    "SCANNER": [
        r"scanner", r"scan.*tunnel", r"barcode.*read", r"dimension.*sys",
        r"dws", r"volume.*scan"
    ],
}

# Structural components to filter out
STRUCTURAL_PATTERNS = [
    r"leg.*guard", r"guard(?!.*operator)", r"fenc", r"railing",
    r"handrail", r"safety.*(?!operator)", r"pallet(?!.*conv)",
    r"step", r"stair", r"ladder", r"door", r"panel(?!.*control)",
    r"bracket", r"bolt", r"mount(?!.*scanner)"
]

def _is_noise_block(name: str) -> bool:
    """Filter out anonymous noise blocks."""
    n = name.strip()
    if re.match(r"^\*[UDXATE]\d+$", n, re.IGNORECASE):
        return True
    if n.startswith("*") or n.startswith("~"):
        return True
    return False

def _is_structural(name: str) -> bool:
    """Check if component is structural (non-flow)."""
    n_lower = name.lower()
    for pattern in STRUCTURAL_PATTERNS:
        if re.search(pattern, n_lower):
            return True
    return False

def _categorize_component(name: str) -> str:
    """Categorize component by name pattern."""
    n_lower = name.lower()
    
    for category, patterns in COMPONENT_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, n_lower):
                return category
    
    return "UNCATEGORIZED"

def _normalize_group_name(name: str) -> str:
    """Normalize raw block name."""
    n = name.strip()
    if "|" in n:
        n = n.split("|")[-1]
    
    # Keep underscores in FAL codes
    if not n.startswith("FAL"):
        n = re.sub(r"[_\-]+", " ", n)
    
    n = re.sub(r"\s+", " ", n).strip()
    
    # Don't remove numbers from version codes
    if not re.search(r"V\d+$", n, re.IGNORECASE):
        n = re.sub(r"\s*\(?\d+\)?$", "", n).strip()
    
    return n.lower()

def _detect_cbs_type(project_name: str) -> str:
    """Detect CBS type from project name."""
    pn_lower = project_name.lower()
    # Check for linear CBS indicators
    if "linear" in pn_lower or "linear cbs" in pn_lower or "linear sorter" in pn_lower:
        return "Linear CBS"
    # Check for loop CBS indicators (explicitly mentioned)
    if "loop cbs" in pn_lower or "loop sorter" in pn_lower:
        return "Loop CBS"
    # Default to Loop CBS if not specified
    return "Loop CBS"

def _analyze_chute_types(components: dict) -> dict:
    """Analyze chute breakdown by examining component names."""
    chute_analysis = {
        "total": 0,
        "by_type": defaultdict(int),
        "has_type_info": False
    }
    
    for comp_name, count in components.items():
        name_lower = comp_name.lower()
        if "chute" in name_lower:
            chute_analysis["total"] += count
            
            # Try to detect type
            if "gravity" in name_lower or "collection" in name_lower:
                chute_analysis["by_type"]["gravity"] += count
                chute_analysis["has_type_info"] = True
            elif "live" in name_lower or "active" in name_lower:
                chute_analysis["by_type"]["live"] += count
                chute_analysis["has_type_info"] = True
            elif "slide" in name_lower or "sliding" in name_lower:
                chute_analysis["by_type"]["sliding"] += count
                chute_analysis["has_type_info"] = True
            elif "reject" in name_lower or "exception" in name_lower:
                chute_analysis["by_type"]["rejection"] += count
                chute_analysis["has_type_info"] = True
            elif "mini" in name_lower:
                chute_analysis["by_type"]["mini_gravity"] += count
                chute_analysis["has_type_info"] = True
            elif "bulk" in name_lower:
                chute_analysis["by_type"]["bulk"] += count
                chute_analysis["has_type_info"] = True
            elif "non-sort" in name_lower or "nonsort" in name_lower:
                chute_analysis["by_type"]["non_sort"] += count
                chute_analysis["has_type_info"] = True
    
    return chute_analysis

def extract_dxf_components(dxf_path: Path, project_name: str = "") -> dict:
    """Extract and categorize components from DXF file."""
    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()
    hdr = doc.header
    
    # Get units
    units_code = hdr.get("$INSUNITS", None)
    try:
        units_code = int(units_code) if units_code is not None else None
    except:
        units_code = None
    
    # Get extents
    extmin = hdr.get("$EXTMIN", None)
    extmax = hdr.get("$EXTMAX", None)
    
    # Extract block references
    raw_counts: Counter[str] = Counter()
    for e in msp:
        try:
            if e.dxftype() == "INSERT":
                bname = e.dxf.name
                if not _is_noise_block(bname) and not _is_structural(bname):
                    raw_counts[bname] += 1
        except:
            continue
    
    # Categorize components
    categorized: dict = defaultdict(lambda: defaultdict(lambda: {"count": 0, "examples": []}))
    
    for raw_name, cnt in raw_counts.items():
        category = _categorize_component(raw_name)
        gname = _normalize_group_name(raw_name)
        
        categorized[category][gname]["count"] += cnt
        categorized[category][gname]["examples"].append(raw_name)
    
    # Detect system characteristics from project name
    cbs_type = _detect_cbs_type(project_name if project_name else dxf_path.name)
    chute_analysis = _analyze_chute_types(raw_counts)
    
    # Determine induction type
    has_auto_induct = len(categorized.get("AUTO_INDUCT", {})) > 0
    has_operators = len(categorized.get("OPERATOR_STATION", {})) > 0
    has_vds = len(categorized.get("VDS_BUFFER", {})) > 0
    
    if has_auto_induct and has_operators:
        induction_type = "MIXED (Auto + Manual)"
    elif has_auto_induct:
        induction_type = "AUTO"
    elif has_operators:
        induction_type = "MANUAL"
    else:
        induction_type = "UNKNOWN"
    
    # Build category summary
    category_summary = {}
    total_components = 0
    for cat, items in categorized.items():
        count = sum(item["count"] for item in items.values())
        category_summary[cat] = count
        total_components += count
    
    return {
        "file": dxf_path.name,
        "units_code": units_code,
        "units_name": UNITS.get(units_code, "unknown") if units_code is not None else None,
        "extents": {
            "min": list(extmin) if extmin else None,
            "max": list(extmax) if extmax else None
        },
        "cbs_type": cbs_type,
        "induction_type": induction_type,
        "has_vds": has_vds,
        "total_components": total_components,
        "category_summary": dict(category_summary),
        "categorized_components": {
            cat: {name: data["count"] for name, data in items.items()}
            for cat, items in categorized.items()
        },
        "chute_analysis": chute_analysis,
        "raw_block_counts": {k: int(v) for k, v in raw_counts.items()},
    }

def _summarise_components_for_prompt(dxf_json: dict) -> str:
    """Generate comprehensive summary for LLM prompt."""
    lines = ["=" * 70]
    lines.append("DXF COMPONENT ANALYSIS FOR CBS PROCESS FLOW GENERATION")
    lines.append("=" * 70)
    lines.append("")
    
    # File and system info
    lines.append(f"FILE: {dxf_json.get('file', 'Unknown')}")
    lines.append(f"UNITS: {dxf_json.get('units_name', 'Unknown')}")
    lines.append(f"CBS TYPE: {dxf_json.get('cbs_type', 'Unknown')} (detected from filename)")
    lines.append(f"TOTAL COMPONENTS: {dxf_json.get('total_components', 0)}")
    lines.append("")
    
    # System configuration
    lines.append("SYSTEM CONFIGURATION:")
    lines.append(f"  • Induction Type: {dxf_json.get('induction_type', 'Unknown')}")
    lines.append(f"  • VDS/Buffer System: {'YES' if dxf_json.get('has_vds') else 'NO'}")
    
    chute_analysis = dxf_json.get('chute_analysis', {})
    if chute_analysis.get('total', 0) > 0:
        lines.append(f"  • Total Chutes: {chute_analysis['total']}")
        if chute_analysis.get('has_type_info'):
            lines.append("  • Chute Types Detected: YES")
        else:
            lines.append("  • Chute Types Detected: NO (describe generically)")
    lines.append("")
    
    # Component categories
    cat_summary = dxf_json.get("category_summary", {})
    if cat_summary:
        lines.append("COMPONENT CATEGORIES:")
        priority = ["AUTO_INDUCT", "OPERATOR_STATION", "VDS_BUFFER", "CONVEYOR_INFEED",
                   "CBS_SORTER", "CHUTE", "PTL", "BAG_SYSTEM", "RECIRCULATION", "SCANNER"]
        
        for cat in priority:
            if cat in cat_summary:
                lines.append(f"  • {cat}: {cat_summary[cat]} units")
        
        # Add any remaining categories
        for cat, count in cat_summary.items():
            if cat not in priority:
                lines.append(f"  • {cat}: {count} units")
        lines.append("")
    
    # Chute breakdown
    if chute_analysis.get('by_type'):
        lines.append("CHUTE TYPE BREAKDOWN:")
        for ctype, count in sorted(chute_analysis['by_type'].items(), key=lambda x: -x[1]):
            lines.append(f"  • {ctype.replace('_', ' ').title()}: {count} chutes")
        lines.append("")
    
    # Detailed inventory by category
    lines.append("DETAILED COMPONENT INVENTORY:")
    lines.append("-" * 70)
    
    categorized = dxf_json.get("categorized_components", {})
    for cat in priority:
        if cat in categorized and categorized[cat]:
            lines.append("")
            lines.append(f"[{cat}]")
            for name, count in sorted(categorized[cat].items(), key=lambda x: -x[1]):
                lines.append(f"  • {name}: {count} units")
    
    # Add uncategorized if any
    if "UNCATEGORIZED" in categorized and categorized["UNCATEGORIZED"]:
        lines.append("")
        lines.append("[UNCATEGORIZED]")
        for name, count in sorted(categorized["UNCATEGORIZED"].items(), key=lambda x: -x[1]):
            lines.append(f"  • {name}: {count} units")
    
    lines.append("")
    lines.append("=" * 70)
    lines.append("PROCESS FLOW GENERATION GUIDANCE:")
    lines.append(f"  • System Type: {dxf_json.get('cbs_type', 'Unknown')}")
    lines.append(f"  • Induction: {dxf_json.get('induction_type', 'Unknown')}")
    
    if dxf_json.get('has_vds'):
        lines.append("  • Include VDS/Buffer section in Infeed System")
    
    if dxf_json.get('induction_type') == "MIXED (Auto + Manual)":
        lines.append("  • Include BOTH Auto Induct Line AND Manual Induct Station sections")
    elif dxf_json.get('induction_type') == "AUTO":
        lines.append("  • Include Auto Induct Line section only")
    elif dxf_json.get('induction_type') == "MANUAL":
        lines.append("  • Include Manual Induct Station section only")
    
    if chute_analysis.get('has_type_info'):
        lines.append("  • Chute types available - use detailed breakdown")
    else:
        lines.append("  • Chute types NOT available - describe generically")
    
    if "PTL" in cat_summary:
        lines.append("  • Include Put To Light System section")
    
    if "BAG_SYSTEM" in cat_summary:
        lines.append("  • Include Bag Takeaway Conveyor section")
    
    if "RECIRCULATION" in cat_summary:
        lines.append("  • Include Recirculation/Exception Refeeding section")
    
    lines.append("=" * 70)
    
    return "\n".join(lines)

def convert_dxf_to_png(dxf_path: Path) -> Path:
    """Convert DXF file to PNG using ConvertAPI."""
    if not CONVERTAPI_SECRET:
        raise RuntimeError(
            "CONVERTAPI_SECRET is not set in .env file. Cannot convert DXF to PNG without it."
        )
    
    try:
        # Convert DXF to PNG using ConvertAPI
        result = convertapi.convert("png", {"File": str(dxf_path)}, from_format="dxf")
        out_files = result.save_files(str(dxf_path.parent))
        
        # Find the PNG file
        for f in out_files:
            if str(f).lower().endswith(".png"):
                return Path(f)
        
        # Return first file if no .png extension found
        return Path(out_files[0]) if out_files else None
    except Exception as e:
        raise RuntimeError(f"Failed to convert DXF to PNG: {str(e)}")

def _normalise_to_numbered_steps(raw_text: str) -> str:
    """Force clean 1..N numbered list from GROQ output."""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    if len(lines) == 1:
        parts = re.split(r'(?:(?<=\.)\s+)(?=\d+\.)', lines[0])
        lines = [p.strip() for p in parts if p.strip()]
    
    steps = []
    for ln in lines:
        m = re.match(r"^(\d+)[\.\)\-]\s*(.*)$", ln)
        content = m.group(2).strip() if m else ln
        if content: steps.append(content)
    
    dedup = []
    seen = set()
    for s in steps:
        key = re.sub(r"\s+", " ", s.lower())
        if key not in seen:
            seen.add(key)
            dedup.append(s)
    
    max_steps = min(len(dedup), 9) if len(dedup) >= 5 else len(dedup)
    return "\n".join([f"{i}. {content}" for i, content in enumerate(dedup[:max_steps], start=1)])

# ==================== LOOP CBS EXCEL PARSING ====================

def load_loop_cbs_sheet_from_excel(xlsx_bytes: bytes) -> tuple[str | None, pd.DataFrame | None]:
    """
    Find sheet whose name is 'Loop CBS' or 'Loop CBS upper' (case/space-insensitive)
    and return (sheet_name, DataFrame). Reads the FULL sheet (no cropping).
    """
    with pd.ExcelFile(io.BytesIO(xlsx_bytes)) as xls:
        target_name = None
        for sname in xls.sheet_names:
            normalized = sname.lower().replace(" ", "")
            if normalized == "loopcbs" or normalized == "loopcbsupper":
                target_name = sname
                break

        if target_name is None:
            return None, None

        # Read entire sheet as generic table (no header), so we keep all rows/blocks.
        df = xls.parse(target_name, header=None)
        return target_name, df


def df_to_compact_text(df: pd.DataFrame, max_rows: int = 200, max_cols: int = 20) -> str:
    """
    Convert the ENTIRE sheet (up to max_rows, max_cols) into a compact text representation
    including BOTH tables in Loop CBS sheet.
    """
    if df is None:
        return ""

    df2 = df.iloc[:max_rows, :max_cols].fillna("")

    lines: list[str] = []
    for idx in range(df2.shape[0]):
        row_vals = [str(v).strip() for v in df2.iloc[idx].tolist()]
        # keep only non-empty cells in the print
        row_vals = [v for v in row_vals if v != ""]
        if row_vals:
            lines.append(" | ".join(row_vals))

    return "\n".join(lines)

def df_to_compact_text_quote_master(df: pd.DataFrame) -> str:
    """
    Convert Quote Master dataframe to compact text for GROQ.
    Every non-empty row becomes a pipe-separated list of "Col=Value".
    """
    lines = []
    headers = list(df.columns)

    for _, row in df.iterrows():
        if row.isna().all():
            continue

        pairs = []
        for h, v in zip(headers, row.values):
            if pd.isna(v) or h is None:
                continue
            v_str = str(v).strip()
            if not v_str:
                continue
            pairs.append(f"{h}={v_str}")

        if pairs:
            lines.append(" | ".join(pairs))

    return "\n".join(lines)

# ==================== PROCESS FLOW GENERATION ====================


def call_groq_for_process_flow(client_name: str, project_name: str, dxf_json: dict):
    """
    Generate process flow using agentY logic with PROGRESSIVE IMPROVEMENT PROTECTION.
    
    🔧 ENHANCED FEATURES:
    - Allows up to 6 iterations
    - Detects and PREVENTS score regression
    - Ensures progressive improvement each iteration
    - Tracks best score and reverts if needed
    
    Args:
        client_name: Name of the client
        project_name: Name of the project
        dxf_json: Parsed DXF component data
    
    Returns:
        tuple: (final_flow_text, iteration_details_list)
    """
    import logging
    logger = logging.getLogger(__name__)
    
    safe_dxf_json = {k: v for k, v in dxf_json.items() if k != "raw_block_counts"}
    safe_dxf_json['client'] = client_name
    iteration_details = []
    
    target_score = 88.0
    max_iterations = 6  # ✅ Up to 6 iterations
    min_improvement = 0.5  # Require at least 0.5 point improvement

    try:
        # ================================================================
        # STEP 1: INITIAL GENERATION (from DXF only)
        # ================================================================
        initial_flow = agentY_generate_initial_flow(client_name, safe_dxf_json)
        if not initial_flow or len(initial_flow.strip()) < 50:
            return "", []

        # ================================================================
        # STEP 2: QUERY REFERENCES & REFINE
        # ================================================================
        reference_flows = []
        try:
            pc, index = combine_old_mod.get_pinecone_index()
            dxf_summary = create_dxf_summary_for_embedding(safe_dxf_json)
            reference_flows = combine_old_mod.query_similar_flows(
                pc, index, dxf_summary, safe_dxf_json, top_k=2, threshold=0.70
            )
        except Exception as e:
            logger.warning(f"Failed to query Pinecone: {e}")
            reference_flows = []
        
        # Generate second flow with references
        if reference_flows:
            try:
                second_flow = generate_second_flow_with_chunks(
                    initial_flow, safe_dxf_json, reference_flows, client_name
                )
                current_flow = second_flow
            except Exception as e:
                logger.warning(f"Failed to generate second flow: {e}")
                current_flow = initial_flow
        else:
            current_flow = initial_flow
        
        # Get reference text for evaluation
        reference_text = reference_flows[0]["process_flow"] if reference_flows else current_flow
        
        # ================================================================
        # STEP 3 & 4: EVALUATION & ITERATIVE REFINEMENT (ENHANCED PROTECTION)
        # ================================================================
        best_score = 0.0
        best_flow = current_flow
        best_iteration = 0
        no_improvement_iterations = 0
        last_score = 0.0
        score_regression_count = 0
        
        for iteration_num in range(1, max_iterations + 1):
            # Evaluate current flow
            evaluation = agentY_evaluate_process_flow(
                current_flow,
                reference_text,
                safe_dxf_json,
                target_score=target_score
            )
            
            current_score = evaluation.get('structural_coherence', 0.0)
            feedback_items = evaluation.get('feedback', [])
            
            # Store iteration details with FULL FLOW TEXT
            iteration_details.append({
                "iteration": iteration_num,
                "score": round(current_score, 2),
                "target_score": target_score,
                "gap": round(target_score - current_score, 2),
                "feedback": feedback_items,
                "flow": current_flow  # ✅ FULL FLOW TEXT
            })
            
            logger.info(f"Iteration {iteration_num}: Score={current_score:.2f}, Target={target_score}, Feedback={len(feedback_items)} items")
            
            # ===================================================================
            # 🔧 SCORE REGRESSION PROTECTION
            # ===================================================================
            
            # Check for significant score drop (>3 points)
            if iteration_num > 1 and current_score < last_score - 3:
                score_regression_count += 1
                logger.warning(f"⚠️ Score regression detected: {last_score:.2f} → {current_score:.2f}")
                
                # If 2+ regressions, revert to best and stop trying to improve
                if score_regression_count >= 2:
                    logger.info(f"Too many regressions ({score_regression_count}). Reverting to best flow.")
                    current_flow = best_flow
                    no_improvement_iterations = 10  # Force stop
                    break
            else:
                score_regression_count = 0  # Reset counter on improvement/stable
            
            last_score = current_score
            
            # ===================================================================
            # 🔧 BEST SCORE TRACKING (Accept if better)
            # ===================================================================
            
            if current_score > best_score + min_improvement:
                # This is a meaningful improvement - accept it
                improvement = current_score - best_score
                best_score = current_score
                best_flow = current_flow
                best_iteration = iteration_num
                no_improvement_iterations = 0  # Reset counter
                
                logger.info(f"✅ New best score: {best_score:.2f} (+{improvement:.2f})")
                
            elif current_score >= best_score - 1.0:
                # Within 1 point of best - acceptable, update flow but not best
                logger.info(f"✓ Score stable: {current_score:.2f} (best: {best_score:.2f})")
                no_improvement_iterations += 1
                
            else:
                # Worse than best by >1 point - revert to best
                logger.warning(f"⚠️ Score dropped to {current_score:.2f}. Reverting to best: {best_score:.2f}")
                current_flow = best_flow  # Revert to best
                no_improvement_iterations += 1
            
            # ===================================================================
            # 🔧 ENHANCED STOPPING CONDITIONS
            # ===================================================================
            
            # 1. Target reached
            if current_score >= target_score:
                logger.info(f"✅ Target reached! Score {current_score:.2f} >= {target_score}")
                break
            
            # 2. Excellent score
            if current_score >= 90:
                logger.info(f"✅ Excellent score! {current_score:.2f} >= 90")
                break
            
            # 3. No improvement for 3 consecutive iterations
            if no_improvement_iterations >= 3:
                logger.info(f"⚠️ No improvement for {no_improvement_iterations} iterations. Stopping.")
                break
            
            # 4. Max iterations reached
            if iteration_num >= max_iterations:
                logger.info(f"Max iterations ({max_iterations}) reached")
                break
            
            # ===================================================================
            # 🔧 REFINEMENT STEP (if not stopping)
            # ===================================================================
            
            if iteration_num < max_iterations:
                try:
                    logger.info(f"Applying conservative refinement to improve from {current_score:.2f}...")
                    
                    # ALWAYS use BEST flow as base (not current if it regressed)
                    next_flow = iterative_refinement(
                        best_flow,  # ✅ Always refine from best
                        safe_dxf_json,
                        reference_text,
                        evaluation,
                        iteration_num,
                        target_score
                    )
                    
                    # Validate generated flow
                    is_valid, error_msg = validate_flow_quality(next_flow)
                    if is_valid:
                        current_flow = next_flow  # Move to next iteration
                        logger.info("✓ Refinement successful, proceeding to next iteration")
                    else:
                        logger.warning(f"⚠️ Generated flow failed validation: {error_msg}")
                        current_flow = best_flow  # Keep best flow
                        no_improvement_iterations += 1
                        
                except Exception as e:
                    logger.error(f"❌ Refinement error: {e}")
                    current_flow = best_flow  # Keep best flow on error
                    no_improvement_iterations += 1
        
        # ===================================================================
        # 🔧 FINAL RETURN
        # ===================================================================
        
        final_flow = best_flow.strip()
        logger.info(f"Final result: Best score {best_score:.2f} achieved in iteration {best_iteration}")
        logger.info(f"Total iterations: {len(iteration_details)}")
        
        return final_flow, iteration_details
    
    except Exception as e:
        logger.error(f"❌ Error in call_groq_for_process_flow: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return "", []


# ==================== MERMAID FLOWCHART GENERATION ====================

def sanitize_mermaid_for_render(code: str) -> str:
    """Minimal sanitization for Mermaid code."""
    if not code: return code
    code = code.strip()
    if code.startswith("```"):
        lines = code.split("\n")
        if lines[0].strip().startswith("```"): lines = lines[1:]
        if lines and lines[-1].strip() == "```": lines = lines[:-1]
        code = "\n".join(lines).strip()
    if code.lower().startswith("mermaid"): code = code[7:].strip()
    if not code.lower().startswith("flowchart"): return code
    return code.replace("\r\n", "\n").replace("\r", "\n")

def generate_mermaid_png(mermaid_code: str) -> tuple:
    """Render Mermaid diagram to PNG."""
    logs = []
    mermaid_str = sanitize_mermaid_for_render(mermaid_code)
    logs.append(f"Cleaned Mermaid code:\n{mermaid_str}\n")
    
    # Strategy 1: Kroki PNG with JSON
    try:
        payload = {"diagram_source": mermaid_str, "diagram_type": "mermaid", "output_format": "png"}
        resp = requests.post("https://kroki.io/mermaid/png", json=payload, 
                           headers={"Content-Type": "application/json"}, timeout=30)
        if resp.ok and resp.content and len(resp.content) > 100:
            logs.append(f"✓ Kroki PNG returned {len(resp.content)} bytes.")
            return resp.content, "\n".join(logs)
    except Exception as e:
        logs.append(f"Kroki PNG error: {repr(e)}")
    
    # Strategy 2: Mermaid.ink
    try:
        json_payload = {"code": mermaid_str, "mermaid": {"theme": "default"}}
        b64_str = base64.b64encode(json.dumps(json_payload).encode('utf-8')).decode('ascii')
        resp = requests.get(f"https://mermaid.ink/img/{b64_str}", timeout=30)
        if resp.ok and resp.content and len(resp.content) > 100:
            logs.append(f"✓ mermaid.ink returned {len(resp.content)} bytes.")
            return resp.content, "\n".join(logs)
    except Exception as e:
        logs.append(f"mermaid.ink error: {repr(e)}")
    
    raise RuntimeError(f"Failed to render Mermaid:\n" + "\n".join(logs))

def call_groq_for_mermaid(process_flow_text: str):
    """Generate Mermaid flowchart code from process flow."""
    system_prompt = """You are a diagram expert for Mermaid v11 flowcharts.
RULES:
- Start with: flowchart TD (top-down)
- Simple IDs: A, B, C, D (no special chars)
- Square brackets for labels: A[Start]
- Short labels (2-5 words)
- Use --> for arrows
- Apply colors using classDef and :::className
- Green (#90EE90) for input, Blue (#87CEEB) for process, Yellow (#FFE97F) for sorting, 
  Orange (#FFB366) for collection, Red (#FFB3B3) for rejection
- Output ONLY mermaid code, no backticks
**MUST FOLLOW - ONLY RAW output DO NOT ADD ```mermaid``` or ``` or any extra text in response***"""

    user_prompt = f"""Convert to vertical Mermaid flowchart with colors:
{process_flow_text}

Use flowchart TD, simple node IDs (A,B,C), short labels, apply color coding with classDef."""

    def api_call():
        return groq_client.chat.completions.create(
            model="groq/compound",
            messages=[
                {"role": "system", "content": system_prompt.strip()},
                {"role": "user", "content": user_prompt.strip()},
            ],
            temperature=0.0,
            max_tokens=2000,
        )
    
    resp = call_groq_with_retry(api_call)
    mermaid_code = resp.choices[0].message.content.strip()
    mermaid_code = re.sub(r"^```(?:mermaid)?\s*", "", mermaid_code, flags=re.MULTILINE)
    mermaid_code = re.sub(r"\s*```$", "", mermaid_code, flags=re.MULTILINE)
    return mermaid_code.strip()

def extract_full_text_from_docx(doc: Document) -> str:
    """Concatenate all paragraph and table text from the DOCX."""
    parts = []
    for p in doc.paragraphs:
        if p.text:
            parts.append(p.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text:
                    parts.append(cell.text)

    return "\n".join(parts)


def find_terms_in_text(text: str):
    """
    Return list of (term, description) that actually appear in the text.
    Matching is done as a whole word, case-sensitive, to avoid accidental
    matches of 'it' vs 'IT'.
    """
    found = []
    for term, desc in GLOSSARY_ENTRIES:
        # Build safe regex – whole word / token
        pattern = r"(?<![A-Za-z0-9])" + re.escape(term) + r"(?![A-Za-z0-9])"
        if re.search(pattern, text):
            found.append((term, desc))

    # keep original order, no duplicates
    seen = set()
    unique = []
    for t, d in found:
        if t not in seen:
            seen.add(t)
            unique.append((t, d))
    return unique

# ==================== PROPOSED SYSTEM TECHNICAL DETAILS - BOM AGGREGATION ====================

PROPOSED_SYSTEM_TECHNICAL_DETAILS_PROMPT = """
You are a proposal BOM aggregation engine for Cross Belt Sorter projects.

INPUT
- You receive flattened rows from an Excel sheet called "Quote Master".
- Each row looks like:
  "S.No=1 | Category=Conveyor | Description=Straight PVC Conveyor BW 800 | UOM=m | Re=24 | ..."
- Columns can include S.No., Category, Sub-Category, Description, UOM, Costing Type, Cost,
  Unit Price, Sales Factor Requirement, Recommended Quantity, etc.

GOAL
- Build a SHORT "Proposed System Technical Details" section, split into 3 logical sub-sections:
  1) "Mechanical equipment"
  2) "Electrical Equipment"
  3) "Control System"
- The output must be highly aggregated, similar to Falcon proposal tables, NOT one row per part.

GENERAL RULES
- Strongly prefer 3–6 rows in Mechanical, 1 row in Electricals, 1 row in Control System.
- Group fine-grained parts into functional systems.
- DO NOT list every small item separately in the final table.
- DO NOT invent components that are not present in the sheet.

MECHANICAL EQUIPMENT – TYPICAL SYSTEMS
Create a small number of "systems" such as (examples, use only what makes sense):
- "Infeed System" or "Auto Infeed System"
- "Feedlines" or "Auto Induct Feedlines"
- "Sorter – 1 Loop Cross Belt Sorter"
- "Sorter Outputs"
- "Steel Works"
- Any other clear mechanical system suggested by the data.

Mapping guidance:
- Infeed / Auto Infeed System:
  - Straight PVC conveyors, S3 conveyors, gravity rollers, curves, merges, spacing conveyors that
    clearly belong to the main inbound line.
- Feedlines / Auto Induct Feedlines:
  - Receiving, weighing, spacing, buffer conveyors, angle merges that feed the sorter.
- Sorter:
  - Loop CBS sorter, carrier pitch, sorter length/height/speed, type of drive (LIM/LSM),
    supports, fencing, empty-carrier detection, product centring, dimension scanning, etc.
- Sorter Outputs:
  - PTL chutes, rejection chutes, spurs, pop-up sorter units, bag holding assemblies, bins, etc.
- Steel Works:
  - Platforms, mezzanines, stairs, ladders, fencing, safety guards, leg guards, end joints etc.

ELECTRICAL EQUIPMENT
- Usually a single line "Electricals – Consists of".
- Group all power/controls items:
  - main power distribution panel, MCC, main control panel, feedline control panels,
    sorter drive panels, VFD panels, network switches, field cabling, earthing, hooters,
    tower lamps, pull cords, emergency stops, IO cards, surge suppressors, harmonic filters, etc.
- Quantity: typically "1 Set".
- Value: usually just "Included" (unless there is a clear different summary).

CONTROL SYSTEM
- Usually a single line "Components – Consists of".
- Group all PLC / SCADA / IT / software items:
  - PLC based control system, SCADA, industrial switches, servers/IPC, CCTV/VMS,
    OCR/vision PCs, WCS / sorter control software, PTL controllers, licences, custom IT integration, etc.
- Quantity: typically "1 Set".
- Value: use simple summary such as "1 Nos" and "As per requirement" where appropriate.

DATA TO PRODUCE FOR EACH ROW
For each final row in the tables, you must produce:
- pos: integer, starting from 1 within each section.
- qty: short human-readable quantity, e.g. "1 Set", "14 Feedlines", "1 Sorter".
- description_lines: array of short text lines for the Description column:
    * Line 1: main system name (e.g. "Infeed System" or "Sorter").
    * Following lines: bullet-style sub points starting with "• ". These are
      the key components grouped into this system.
- value_lines: array of short text lines for the Value column. Use this for:
    * Total counts or dimensions of important components.
    * Example: "Straight PVC Conveyors: ~38 m total", "Curve Conveyors: 2 Nos".

IMPORTANT: think in terms of systems, not raw rows.

---------------- EXAMPLES (VERY IMPORTANT) ----------------

Example A – Aggregating an Infeed System

INPUT snippet (conceptual):
ROW: Category=Conveyor | Description=Straight PVC Conveyor BW 800 | UOM=m | Re=24
ROW: Category=Conveyor | Description=Straight PVC Conveyor BW 1000 | UOM=m | Re=14
ROW: Category=Conveyor | Description=Curve Conveyor 30 deg | UOM=Nos | Re=2

EXPECTED mechanical item:
{
  "pos": 1,
  "qty": "1 Infeed System",
  "description_lines": [
    "Infeed System",
    "• Straight PVC Conveyors",
    "• Curve Conveyors"
  ],
  "value_lines": [
    "Straight PVC Conveyors: ~38 m total",
    "Curve Conveyors: 2 Nos"
  ]
}

Example B – Aggregating Feedlines

INPUT snippet (conceptual):
Rows describing "Receiving Conveyor", "Weighing Conveyor", "Spacing Conveyor",
"Buffer Conveyor", "Angle merge" with various quantities.

EXPECTED mechanical item:
{
  "pos": 2,
  "qty": "14 Feedlines",
  "description_lines": [
    "Feedlines – Consists of",
    "• Receiving Conveyor",
    "• Weighing Conveyor",
    "• Spacing Conveyor",
    "• Buffer Conveyor",
    "• Angle merge"
  ],
  "value_lines": [
    "Receiving Conveyor: 1 Set",
    "Weighing Conveyor: 1 Set",
    "Spacing Conveyor: 3 Set",
    "Buffer Conveyor: 3 Set",
    "Angle merge: 2 Set"
  ]
}

Example C – Aggregating the Sorter

INPUT snippet (conceptual):
Rows describing a loop CBS sorter with height, length, speed, drive type,
carrier pitch and associated mechanical options.

EXPECTED mechanical item:
{
  "pos": 3,
  "qty": "1 Sorter",
  "description_lines": [
    "Sorter",
    "1 Loop Cross Belt Sorter",
    "• Sorter height",
    "• Sorter length",
    "• Sorter speed",
    "• Sorter drive",
    "• Carrier pitch",
    "Including:",
    "• Standard sorter supports",
    "• Product centring system",
    "• Dimension / barcode scanning system",
    "• Hooters and E-stops",
    "• Fencing"
  ],
  "value_lines": [
    "Height: approx 2900 mm",
    "Loop length: approx 150 m"
  ]
}

Example D – Electricals

INPUT snippet (conceptual):
Rows for power distribution panel, main control panel, feedline control panels,
sorter drive panels, network switches, field cabling, hooters, tower lamps, etc.

EXPECTED electrical section:
{
  "title": "Electrical Equipment",
  "items": [
    {
      "pos": 1,
      "qty": "1 Set",
      "description_lines": [
        "Electricals",
        "Consists of",
        "Main power distribution panel",
        "Main control panel",
        "Feedline control panels",
        "Sorter drive panels",
        "Network switches",
        "Field cabling"
      ],
      "value_lines": [
        "Included"
      ]
    }
  ]
}

Example E – Control System

INPUT snippet (conceptual):
Rows for Siemens PLC, SCADA, industrial switches, servers, sorter control software.

EXPECTED control section:
{
  "title": "Control System",
  "items": [
    {
      "pos": 1,
      "qty": "1 Set",
      "description_lines": [
        "Components",
        "Consists of",
        "PLC based control system with SCADA",
        "Industrial switch"
      ],
      "value_lines": [
        "1 Nos",
        "As per requirement"
      ]
    }
  ]
}

---------------- OUTPUT FORMAT ----------------

Return JSON ONLY in this schema:

{
  "sections": [
    {
      "title": "Mechanical equipment",
      "items": [
        {
          "pos": 1,
          "qty": "1 Set",
          "description_lines": ["..."],
          "value_lines": ["..."]
        }
      ]
    },
    {
      "title": "Electrical Equipment",
      "items": [ ... ]
    },
    {
      "title": "Control System",
      "items": [ ... ]
    }
  ]
}
"""

def call_groq_for_bom(sheet_text: str) -> dict:
    """Call GROQ API to generate aggregated BOM structure from Quote Master sheet."""
    user_prompt = (
        "Below are flattened rows from the 'Quote Master' sheet.\n\n"
        "QUOTE_MASTER_ROWS:\n"
        + sheet_text
    )

    def api_call():
        return groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": PROPOSED_SYSTEM_TECHNICAL_DETAILS_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
        )

    try:
        resp = call_groq_with_retry(api_call)
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        st.error(f"Failed to generate Proposed System Technical Details: {e}")
        return None

# ==================== GROQ PROMPTS & CONSTANTS ====================

# Cover Letter System Prompt
COVER_LETTER_SYSTEM_PROMPT = """
You are an AI assistant working as a professional proposal writer at Falcon Autotech. You are an expert in drafting formal, client-specific techno-commercial cover letters for proposals. Your role is to generate well-structured, personalized cover letters that follow Falcon's business communication style, maintain a professional and respectful tone, and clearly demonstrate Falcon's commitment, expertise, and partnership approach to clients.

Generate a formal techno-commercial COVER LETTER for a proposal that MUST fit within a single page. 
The writing style MUST be indistinguishable from natural human writing. The text should read as if drafted by an experienced professional, not an AI system. Use clear, simple, and natural language with varied sentence lengths and structures. Avoid generic phrases, repetitive patterns, or mechanical tone. Ensure that the output flows smoothly, conveys intent naturally, and would not be detected as machine-generated. The content should feel thoughtful, context-aware, and aligned with how a human proposal writer or business professional would communicate.

STRICT LENGTH LIMIT: Maximum 300 words to ensure single-page fit. Be concise and impactful.

1. Start with:
   Kind Attention –
   Mr. {{executives}}
   M/s {{client_name}}

   Offer Ref: {{offer_ref}}; Date: {{letter_date}}

   Subject – Techno-Commercial Offer for {{project_title}}  

2. If there is only one executive, address them with:
   Dear {{first_exec_name}},
   If multiple executives, skip "Dear" and go directly to the content.
   Use Mr. for male and Ms. for female executives.

3. Opening paragraph (natural, professional style):
   - Thank the client for inviting Falcon to offer for the project.
   - If invitation_date or meeting_date exists, reference it naturally (e.g., "Over the past period we worked closely together" or "In our meeting on [date], we discussed...").
   - Mention that you are pleased to submit the Techno-Commercial Offer.
   - Wording must vary between runs (not fixed sentences).

4. Middle paragraph - System Overview & Analysis (CRITICAL):
   - State that Falcon has done an in-depth data analysis and evaluated various solution options.
   - **MANDATORY: Include the high-level process flow summary if provided**. Mention key system components naturally in a single sentence (e.g., "The proposed solution includes automatic induct conveyors, cross-belt sorter with scanner systems, and output chutes for efficient sortation").
   - Highlight any specific technical values, quantities, or capacities if mentioned (e.g., "200 destinations", "5 camera scanner systems", "2 speed settings").
   - Keep this brief but informative - demonstrate technical understanding without overwhelming detail.
   - Mention that the detailed technical proposal is laid out in various sections to provide full insight into the proposed solution.

5. Commitment paragraph:
   - Reference Falcon's intralogistics automation technologies and proven track record.
   - Highlight subsequent sections covering capabilities, experiences, and references.
   - Reinforce commitment to being a strategic partner.

6. Closing (professional, warm):
   - Add sender's personal commitment on behalf of Falcon Autotech.
   - Encourage client to reach out for clarifications or further information.
   - End with "Best Regards," followed by sender_name and sender_title.

Important:
- MUST NOT EXCEED 300 words to ensure single-page fit.
- Keep tone formal, professional, and client-oriented.
- Do not copy exact sentences; rephrase wording across generations.
- The cover letter MUST sound human, natural and professional. It should be clear, authentic, and warm, without feeling robotic or overly formal.
- DO NOT ADD ANY EXTRA WORD OR INFO APART FROM THE COVER LETTER.
- Highlight the main system or project name in main body (not subject line) as bold style, use ** for Bold.
- If process_flow_summary is provided, ALWAYS incorporate it naturally into the letter.
"""

COVER_LETTER_USER_PROMPT_TEMPLATE = """
Use the following information to generate the cover letter:

client_name: {client_name}
project_title: {project_title}
offer_ref: {offer_ref}
letter_date: {letter_date}

executives (one per line, already with Mr./Ms. prefix):
{executives_block}

invitation_date: {invitation_date}
meeting_date: {meeting_date}

process_flow_summary (very high-level system components and key quantities): {process_flow_summary}

sender_name: {sender_name}
sender_title: {sender_title}

CRITICAL REQUIREMENTS:
1. The cover letter MUST fit within a single page (maximum 300 words).
2. If process_flow_summary is provided, ALWAYS incorporate it naturally into the letter body to demonstrate technical understanding.
3. Mention any specific quantities or technical details from the summary to add credibility.
4. Keep the tone professional, warm, and client-focused like the example letter provided.

Return ONLY the cover letter text, without markdown code fences or extra commentary.
"""

# Executive Summary System Prompt
EXEC_SUMMARY_SYSTEM_PROMPT = """
You are a Proposal Writing Assistant specialized in Falcon Autotech automation projects.  
Falcon Autotech designs, manufactures, supplies, implements, and maintains warehouse automation solutions—such as sortation systems, conveyor automation, pick/put-to-light, ASRS robotics, and dimension & weight scanning—for industries including e-commerce, fashion, FMCG, pharma, groceries, and CE-P.  
The writing style must be indistinguishable from natural human writing. The text should read as if drafted by an experienced professional, not an AI system. Use clear, simple, and natural language with varied sentence lengths and structures. Avoid generic phrases, repetitive patterns, or mechanical tone. Ensure that the output flows smoothly, conveys intent naturally, and would not be detected as machine-generated. The content should feel thoughtful, context-aware, and aligned with how a human proposal writer or business professional would communicate.

Your task is to generate **unique, client-tailored Executive Summaries** based on the "Proposed System Description" section of Falcon proposals.  
The summary must always reflect Falcon's style but **no two summaries should ever be identical**. Introduce subtle variations in wording, phrasing, and sentence structure while keeping the same professional tone.  

### Writing Rules

**Opening Section**
- Begin with Falcon Autotech's commitment and strong interest in responding to the client's requirement.  
- Mention Falcon's partnership approach, customization, and proven track record.  
- Use varied sentence structures and synonyms so every generation feels different.  

**Bullet Points**
- Provide exactly **4–5 high-level system features or modules**.  
- Each bullet MUST be short, clear, and client-friendly (e.g., "Spiral Conveyors for smooth material flow").  
- Avoid technical specifications, sub-bullets, or repeating the same idea in different words.  
- The order of bullets should vary slightly between generations.  
- Add numeric along with the components ONLY IF extensively mentioned in Proposed System Description
- Bold the main components of the system. There can be max 2-3 bold words.
- **CRITICAL: Always use numeric format for quantities (e.g., 3, 9, 24, 202) instead of words (e.g., three, nine, twenty-four).**

**Closing Section**
- End with a **personalized closing statement**.  
- Reaffirm that the solution is tailored to meet the client's technical and operational requirements.  
- Mention the RFP/customization and highlight benefits like efficiency, smooth material flow, and faster TAT.  
- Closing phrasing should change between runs (use variations in tone, sentence structure, and emphasis).  

### Important Constraints
- Keep the tone formal, professional, and benefit-driven.  
- Do **not** reuse exact sentences from earlier examples.  
- Ensure variability: two runs for the same input must never produce identical text.  
- Do **not** add any extra sections outside the defined structure.  

### Output Format
1. Opening paragraph (commitment + partnership).  
2. 4–6 bullet points (system modules).  
3. Closing personalized statement.  

DO NOT ADD ANY EXTRA TEXT OR INFORMATION OR JUSTIFICATION or "Here is an Executive Summary for the proposal:" EXCEPT THE FULL PROPOSAL
"""

# System Description System Prompt
ENHANCED_SYSTEM_DESCRIPTION_PROMPT = """You are an expert Material Handling System Engineer specializing in Cross-Belt Sorter systems. Your task is to generate COMPREHENSIVE, DETAILED, and EXTENSIVE system descriptions that match the depth and technical detail of professional engineering documentation.

**CRITICAL INSTRUCTIONS:**

1. **USE ONLY PROVIDED INFORMATION:**
   - Extract ALL information from the process flow input
   - Extract ALL quantities and specifications from the DXF file information
   - DO NOT use any values from training examples
   - DO NOT assume or invent specifications

2. **DXF FILE INTEGRATION:**
   You will receive DXF file information in JSON format containing:
   - File name and units
   - Block counts for components (chutes, operators, leg guards, fencing, pallets, etc.)
   - Groups with total counts
   
   **Use this DXF data to:**
   - Extract exact quantities for chutes, operators, safety equipment
   - Include specific counts in relevant sections
   - Reference the DXF file as the source of layout information
   - Add details about protection, fencing, and infrastructure based on block counts

3. **SECTION GENERATION - BE EXTREMELY DETAILED:**

   Create sections ONLY for components mentioned in process flow or DXF data. Each section must be COMPREHENSIVE with multiple paragraphs.

   **INFEED SYSTEM** (if mentioned):
   - Write 4-6 detailed paragraphs
   - Describe the overall configuration and purpose
   - Explain each conveyor type in detail (3-4 sentences each):
     * **Straight Belt Conveyor**: Modular and robust design, used for smooth conveying of products over straight paths. MS profile is used to build conveyor frame. The conveyors are supplied with necessary supports and bolts to fix them to the supporting plane, as well as junction elements allowing easy and jam-free passage from one conveyor to another. Features include low noise, maximum uptime, minimal maintenance, high safety standards, and fastest ROI.
     * **Inclined PVC Conveyor**: Used for smooth conveying of products over inclined and declined paths. Belt conveyors feature modular design with MS profile construction. Supplied with necessary supports, bolts, and junction elements for seamless integration.
     * **Buffer Conveyor**: A buffer conveyor, also known as a buffering conveyor or accumulation conveyor, is a type of conveyor system used to temporarily store or hold items in a controlled manner. Its primary purpose is to manage the flow of items between different stages of a production or handling process when there is a mismatch in the speeds or capacities of the upstream and downstream equipment. These conveyors are required to maintain the throughput of the line.
     * **Curve Conveyor**: Robust and easily maintainable design. The uniquely designed curves and belts provide smooth environment to parcels for making turns. The metal frames of the belts are not deformable to prevent belt misalignment. The belt guide assembly includes removable parts to allow quick replacement in case of damage.
   - Mention flow path from loading to induct zone
   - Include general specifications format: Belt material (PVC), load capacity, motor type (AC Geared Motor), gear motor makes, drive makes
   - Reference total conveyor counts if available from DXF

   **INDUCTION/FEEDLINE SYSTEM** (if mentioned):
   - Write 5-8 detailed paragraphs
   - Describe overall feedline configuration
   - Detail each module type with 3-4 sentences:
     * **Loading/Receiving Conveyor**: A receiving conveyor is a type of conveyor system used to receive and release the products for induction onto CBS. It serves as the connection point at turn point of entry where products are collected and conveyed to subsequent stages of the process. The receiving conveyor accurately positions parcels for smooth transfer to the main sorter.
     * **Weighing Conveyor**: A weighing conveyor, also known as a weigh belt conveyor, is a type of conveyor system specifically designed to measure the weight of materials as they move along the conveyor belt. It combines the functions of conveying and weighing into a single integrated process. Weighing conveyors are equipped with high precision load cells to capture the weight of shipments. Makes include Bizerba, Mettler Toledo, or equivalent manufacturers.
     * **Spacing Conveyor**: A spacing conveyor, also referred to as a gapping conveyor or gap optimizer, is a type of conveyor system used to create and maintain consistent gaps or spacing between items as they move along the conveyor line. Its primary purpose is to regulate the flow and spacing of products to ensure smooth operation and efficient downstream processes. This conveyor is a variable speed special purpose module that creates space between parcels as well as regulates feeding to downstream equipment.
     * **Buffer Conveyors**: Used to temporarily store or hold items in controlled manner. Primary purpose is to manage flow of items between different stages when there is mismatch in speeds or capacities of upstream and downstream equipment. Required to maintain the throughput of line.
     * **Angle Merge Conveyor**: An angle/intelligent merge conveyor incorporates advanced automation and control technologies to intelligently merge stream of materials into a single unified flow. It optimizes the merging process by dynamically adjusting the speed and position of items to ensure a smooth and efficient merge. This is typically a 30° triangular high-speed conveyor used for inducting shipments/boxes directly onto the sorter. The belts are strip belts for smooth shipment movement.
   - Explain sensor placement and functionality
   - Describe how parcels are prepared and positioned for sorter entry
   - Include number of feedlines and capacity from process flow

   **MANUAL INDUCT STATIONS** (if mentioned):
   - Write 2-3 paragraphs
   - Describe location (ground level, mezzanine)
   - Explain operator workflow in detail
   - Mention capacity and number of stations
   - Include operator count from DXF data if available

   **CROSS-BELT SORTER (Main Sorter)**:
   - Write 4-6 detailed paragraphs
   - Describe sorter type (Linear CBS or Loop CBS)
   - Installation details: height from ground, location
   - Carrier specifications: type (single/dual belt), pitch, belt dimensions
   - For Linear: top running length, total length, number of carriers
   - For Loop: loop circumference, deck configuration
   - Operation description: How parcels pass through the sorter, barcode scanning process, chute assignment logic, carrier actuation mechanism, discharge process
   - Explain the sorting sequence step by step

   **BARCODE SCANNING & DIMENSIONING SYSTEM** (if mentioned):
   - Write 3-4 paragraphs
   - Scanner type and configuration (5-side, 6-side, top-only)
   - Technology: ICR (Image Code Reader) or other
   - Manufacturer and model information
   - Capabilities: Barcode types (1D, 2D), scanning coverage, orientation
   - Additional features: Image archiving, dimension measurement accuracy
   - Integration with WCS and sorting logic

   **OUTPUT CHUTES** - BE VERY DETAILED:
   - **Use exact quantities from DXF data**
   - Write 8-12 paragraphs total covering all chute types
   
   For each chute type present:
   
   **Collection Chutes / Manual Chutes**:
   - Extract total count from DXF data (look for "chute", "Chute" in block counts)
   - Write 3-4 paragraphs describing:
     * Type: Friction roller chute or gravity chute design
     * Purpose: A friction roller chute is a type of chute used for the smooth descent of materials or objects from an elevated position to a lower level. It utilizes its roller platform to gradually descend and collect the parcel at the end.
     * Configuration: Single deck or double deck
     * Capacity calculation with example dimensions
     * Equipment per chute: Chute full sensors (quantity and function), three-color tower lights/beacon lights (to indicate chute status), push buttons (to start/stop sorting operations)
   
   **Live Chutes / Live Dock Chutes** (if mentioned):
   - Write 2-3 paragraphs
   - Describe: A live chute refers to a combination of collection chute, PVC belt conveyor, and TBC (if applicable), where the collection chute helps bringing down the sorted parcel and releases it to running conveyor for direct loading into trucks
   - Configuration and integration with conveyors
   
   **Rejection/Technical Chutes** (if mentioned):
   - Write 2-3 paragraphs
   - Purpose: Handle rejected, oversized, overweight, no-read parcels
   - Design and operation
   - Equipment included
   
   **Direct Bagging Chutes** (if applicable):
   - Write 2-3 paragraphs
   - Purpose and operation
   - Integration with bagging system

   **RECIRCULATION & MANUAL REFEED LINE** (if mentioned):
   - Write 3-4 paragraphs
   - Recirculation line: Strategically designed at the end of the sorter system to manage parcels that encounter sorting failures. This automated line efficiently gathers and transports the sort-failed parcels, refeeding them back into the sorter system without requiring additional manual labor. The entire process is seamless, ensuring parcels are automatically re-fed into the sorting system.
   - Manual refeed line: Integration for reintroduction of rejected parcels that have been manually reprocessed. This ensures that manually handled parcels are easily fed back into the sorter, maintaining operational flow and minimizing delays.

   **BAGGING SYSTEM** (if applicable):
   - Write 3-4 paragraphs
   - Bagging conveyor configuration
   - Flow from bagging chutes to bag induct
   - Bag scanning and induction process

   **SECONDARY SORTING / PALLETIZATION** (if applicable):
   - Write 2-3 paragraphs
   - Operator workflow with hand-held terminals
   - Pallet positioning and dispatch
   - Include pallet count from DXF data if available

   **TELESCOPIC BELT CONVEYORS** (if applicable):
   - Write 2-3 paragraphs
   - Quantity and placement
   - Technical specifications: base length, extended length, belt specifications
   - Purpose and operation

   **INFRASTRUCTURE & SUPPORT SYSTEMS**:
   - Write 6-10 paragraphs covering all infrastructure elements
   
   **Mezzanine Platform** (if mentioned):
   - Total area, clear height, type
   - Number of staircases
   - Deck configuration
   
   **Safety & Protection**:
   - **Extract counts from DXF data**:
     * Leg guards count (look for "leg guard", "Leg Guard" in blocks)
     * Operator safety guards (look for "operator safety" in blocks)
     * Fencing (look for "fencing", "Fencing" in blocks)
   - Write detailed paragraphs: Leg guards are protective components designed to shield the legs from external material or component. Material for leg guards is typically MS (Mild Steel). Operator safety guards protect personnel near the system. Perimeter fencing defines the loading zone and protects personnel.
   
   **Pathways**:
   - Allocated pathways for operator and vehicle movement
   
   **System Color Coding** (if applicable):
   - RAL color codes for different system components
   
   **Electrical & Controls Infrastructure**:
   - Control panels, switch racks, socket provisions
   - Cable management systems
   - Communication protocols

   **SYSTEM TECHNICAL SUMMARY**:
   - Write 3-4 paragraphs summarizing:
     * Total conveyor system metrics
     * Feedline configuration and capacity
     * Sorter specifications
     * Total chutes by type (use DXF counts)
     * Operator positions (from DXF)
     * Infrastructure elements
     * Key equipment and technologies

4. **TABLE FORMATTING (CRITICAL):**
   - When you need to present tabular data (e.g., system components, quantities, specifications), use this JSON format:
   
   ```json
   TABLE_START
   {
     "title": "Table Title Here",
     "headers": ["Column1", "Column2", "Column3"],
     "rows": [
       ["Row1Col1", "Row1Col2", "Row1Col3"],
       ["Row2Col1", "Row2Col2", "Row2Col3"]
     ]
   }
   TABLE_END
   ```
   
   - Place this JSON block on its own lines in the output
   - Do NOT use markdown tables (| --- |), ONLY use the JSON format above
   - Use tables for: System Components, Quantities from DXF, Specifications, Equipment Lists

5. **WRITING REQUIREMENTS:**
   - Each major section: 4-8 paragraphs minimum
   - Each subsection: 2-4 paragraphs minimum
   - Each component description: 3-5 sentences minimum
   - Use technical, professional language
   - Explain functionality, purpose, and integration
   - Include design rationale where applicable
   - Maintain consistent technical depth throughout
   - Use proper material handling terminology
   - For bold text, use **text** format (it will be rendered bold without asterisks)
   - **CRITICAL: Always use numeric format for quantities (e.g., 3, 9, 24, 202) instead of words (e.g., three, nine, twenty-four).**
   - For subsection headings, use the format: ## Heading Text (this will be rendered as numbered subheading)
   - Do NOT use bullet-star combinations like •	*Heading** for subsections, ONLY use ## format

6. **QUANTITY EXTRACTION FROM DXF:**
   - Total chutes: Sum all chute-related blocks
   - Operators: Look for "operator", "Operator" in block names
   - Leg guards: Look for "leg guard", "Leg Guard"
   - Fencing: Look for "fencing", "Fencing"
   - Pallets: Look for "pallet", "Pallet"
   - Safety equipment: Look for "safety", "gaurd", "guard"
   - Use these exact numbers in relevant sections

7. **OUTPUT LENGTH TARGET:**
   - Aim for 3000-5000 words total
   - Match the depth and detail of professional engineering system descriptions
   - Every component gets thorough explanation
   - Multiple paragraphs per major section

**REMEMBER:**
- Be EXTREMELY detailed and comprehensive
- Write multiple paragraphs for each section
- Use exact quantities from DXF data
- Explain every component thoroughly
- Match the professional engineering documentation style
- Generate content that is 5-10 pages when exported to Word
- Use JSON format for ALL tables (TABLE_START...TABLE_END)"""

# Config paths
STATIC_ABOUT_DIR = r"Static_AboutCompany"

# Handled Shipment Spectrum Templates
@dataclass
class SorterTemplate:
    key: str
    label: str
    keywords: List[str]
    config_name: str
    item_singular: str
    subheading_51: str
    spec_table: Dict[str, Dict[str, str]]

SORTER_TEMPLATES: List[SorterTemplate] = [
    SorterTemplate(
        key="linear_dual_standard",
        label="Linear / Dual-belt CBS – standard boxes",
        keywords=["linear", "6k", "5.4k", "loop cbs + linear", "totes", "boxes"],
        config_name="Linear Cross Belt Sorter (Dual-belt configuration)",
        item_singular="shipment",
        subheading_51="Shipment size loadable on the sorter",
        spec_table={
            "Max Length": {"unit": "mm", "value": "600"},
            "Max Width":  {"unit": "mm", "value": "450"},
            "Max Height": {"unit": "mm", "value": "400"},
            "Max Weight": {"unit": "Kg", "value": "20"},
            "Min length": {"unit": "mm", "value": "100"},
            "Min Width":  {"unit": "mm", "value": "100"},
            "Min Height": {"unit": "mm", "value": "3"},
            "Min Weight": {"unit": "gm", "value": "50"},
        },
    ),
    SorterTemplate(
        key="loop_standard",
        label="Loop CBS – standard shipments",
        keywords=["loop", "double deck", "48k", "loop cbs", "main sorter"],
        config_name="Loop Cross Belt Sorter technology",
        item_singular="shipment",
        subheading_51="Shipment size loadable on the sorter",
        spec_table={
            "Max Length": {"unit": "mm", "value": "400"},
            "Max Width":  {"unit": "mm", "value": "400"},
            "Max Height": {"unit": "mm", "value": "400"},
            "Max Weight": {"unit": "Kg", "value": "40"},
            "Min length": {"unit": "mm", "value": "10"},
            "Min Width":  {"unit": "mm", "value": "100"},
            "Min Height": {"unit": "mm", "value": "50"},
            "Min Weight": {"unit": "gm", "value": "100"},
        },
    ),
    SorterTemplate(
        key="heavy_parcel",
        label="Heavy-duty CBS – parcels / bags & boxes",
        keywords=["parcel", "heavy", "bags", "bag and box", "bosta", "delhivery"],
        config_name="Heavy Duty Cross Belt Sorter",
        item_singular="parcel",
        subheading_51="Parcel size loadable on the sorter",
        spec_table={
            "Max Length": {"unit": "mm", "value": "1000"},
            "Max Width":  {"unit": "mm", "value": "800"},
            "Max Height": {"unit": "mm", "value": "800"},
            "Max Weight": {"unit": "Kg", "value": "50"},
            "Min length": {"unit": "mm", "value": "40"},
            "Min Width":  {"unit": "mm", "value": "150"},
            "Min Height": {"unit": "mm", "value": "150"},
            "Min Weight": {"unit": "Kg", "value": "0.05"},
        },
    ),
]

st.set_page_config(page_title="Falcon Proposal Generator", page_icon="📄", layout="centered")

# Custom CSS for professional look
st.markdown("""
<style>
    /* Global Styles */
    .main > div { 
        padding-top: 2rem; 
        padding-bottom: 2rem; 
    }
    
    /* Main Header */
    .main-header {
        background: linear-gradient(90deg, #060c71 0%, #2a3bb8 35%, #f9d20e 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(6, 12, 113, 0.3);
        color: white;
    }
    
    .main-header h1 {
        color: #fff !important;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.25);
    }
    
    .main-header .subtitle {
        color: rgba(255,255,255,0.95);
        font-size: 1.1rem;
        margin-top: 0.5rem;
        font-weight: 500;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
    }
    
    /* Section Headers */
    .section-header {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-left: 4px solid #060c71;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        margin: 2rem 0 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    .section-header h3 {
        color: #060c71;
        font-weight: 700;
        margin: 0;
        font-size: 1.3rem;
    }
    
    /* Input Fields */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stDateInput > div > div > input,
    .stSelectbox > div > div > select {
        border-radius: 8px;
        border: 2px solid #e0e0e0;
        padding: 0.75rem;
        font-size: 16px;
        transition: all 0.3s ease;
    }
    
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus,
    .stDateInput > div > div > input:focus,
    .stSelectbox > div > div > select:focus {
        border-color: #060c71;
        box-shadow: 0 0 0 3px rgba(6, 12, 113, 0.1);
    }
    
    /* Labels */
    .stTextInput > label,
    .stFileUploader > label,
    .stTextArea > label,
    .stDateInput > label,
    .stCheckbox > label,
    .stSelectbox > label {
        font-weight: 600;
        color: #2a3bb8;
        font-size: 0.95rem;
    }
    
    /* File Uploader */
    .stFileUploader > div {
        border: 2px dashed #060c71;
        border-radius: 10px;
        padding: 1.5rem;
        text-align: center;
        background: rgba(6,12,113,0.02);
        transition: all 0.3s ease;
    }
    
    .stFileUploader > div:hover {
        background: rgba(6, 12, 113, 0.05);
        border-color: #f9d20e;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #060c71 0%, #2a3bb8 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        font-size: 16px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(6,12,113,0.3);
        width: 100%;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #f9d20e 0%, #ffe34a 100%);
        color: #060c71;
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(249,210,14,0.4);
    }
    
    .stButton > button:disabled {
        background: #cccccc;
        color: #666666;
        transform: none;
        box-shadow: none;
    }
    
    /* Download Button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #28a745 0%, #34ce57 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        width: 100%;
        box-shadow: 0 4px 15px rgba(40,167,69,0.3);
    }
    
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #218838 0%, #28a745 100%);
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(40,167,69,0.4);
    }
    
    /* Expanders */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-radius: 10px;
        border: 2px solid #e0e0e0;
        font-weight: 600;
        color: #060c71;
        padding: 1rem;
    }
    
    .streamlit-expanderContent {
        border: 2px solid #e0e0e0;
        border-top: none;
        border-radius: 0 0 10px 10px;
        background: white;
        padding: 1rem;
    }
    
    /* Info Boxes */
    .info-box {
        background: linear-gradient(135deg, rgba(6,12,113,0.05) 0%, rgba(42,59,184,0.05) 100%);
        border-left: 4px solid #060c71;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 8px;
    }
    
    /* Success/Warning/Error Messages */
    .stSuccess, .stWarning, .stError, .stInfo {
        border-radius: 10px;
        padding: 1rem;
    }
    
    /* Divider */
    hr {
        margin: 2rem 0;
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent 0%, #f9d20e 50%, transparent 100%);
    }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}
</style>
""", unsafe_allow_html=True)
st.markdown('''
<div class="main-header">
    <h1>Falcon Proposal Generator</h1>
    <div class="subtitle">Professional Proposal Document Generation System</div>
</div>
''', unsafe_allow_html=True)

st.markdown("---")

# ==================== STYLING FUNCTIONS ====================

# Deep Blue-Gray color for headings (RGB: 31, 56, 100)
HEADING_COLOR = RGBColor(31, 56, 100)
def render_section_header(title):
    st.markdown(f'''
    <div class="section-header">
        <h3>{title}</h3>
    </div>
    ''', unsafe_allow_html=True)
def apply_heading_style(paragraph, text, level=1):
    """Apply custom heading style: Calibri Headings 14pt, Bold, Underline, Numbered, Deep Blue-Gray"""
    paragraph.text = ""
    run = paragraph.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.underline = True
    run.font.color.rgb = HEADING_COLOR
    
    # Apply paragraph formatting
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(6)
    
    return paragraph

def apply_subheading_style(paragraph, text):
    """Apply subheading style: Calibri 12pt, Bold, Deep Blue-Gray"""
    paragraph.text = ""
    run = paragraph.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(3)
    
    return paragraph

def apply_normal_style(paragraph, text=""):
    """Apply normal text style: Calibri (Body) 11pt, Black"""
    if text:
        paragraph.text = ""
        run = paragraph.add_run(text)
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)
    else:
        for run in paragraph.runs:
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
    
    return paragraph

def apply_table_style(table):
    """Apply Medium Shading 1 Accent 1 style to table"""
    try:
        table.style = 'Medium Shading 1 Accent 1'
    except KeyError:
        # If style doesn't exist, apply manual formatting similar to Medium Shading 1 Accent 1
        # This happens when document is created from a template without this style
        try:
            table.style = 'Table Grid'
        except KeyError:
            # If even Table Grid doesn't exist, skip styling
            pass
    return table

def add_centered_image(doc, path, width_in=5.5):
    """Add a centered image if it exists"""
    if not path or not os.path.exists(path):
        return
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_picture(path, width=Inches(width_in))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

def add_numbered_heading(doc, text, level=1, counter=None):
    """Add a numbered heading with proper formatting"""
    if counter:
        full_text = f"{counter}. {text}"  # Added period after number
    else:
        full_text = text
    
    # Use built-in heading style
    p = doc.add_heading(full_text, level=1)
    
    # Apply custom formatting to the heading
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.underline = True
        run.font.color.rgb = HEADING_COLOR
    
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    
    return p

def add_numbered_subheading(doc, text, counter=None):
    """Add a numbered subheading"""
    if counter:
        full_text = f"{counter}. {text}"  # Added period after number
    else:
        full_text = text
    
    # Use built-in heading style for subheading
    p = doc.add_heading(full_text, level=2)
    
    # Apply custom formatting
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = HEADING_COLOR
    
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    
    return p

def ensure_list_styles(doc):
    """Ensure List Bullet, List Number, and Table styles exist in the document"""
    styles = doc.styles
    
    # Check if List Bullet exists, if not create it
    try:
        styles['List Bullet']
    except KeyError:
        # Create List Bullet style
        from docx.enum.style import WD_STYLE_TYPE
        list_bullet_style = styles.add_style('List Bullet', WD_STYLE_TYPE.PARAGRAPH)
        list_bullet_style.base_style = styles['Normal']
        list_bullet_style.font.name = 'Calibri'
        list_bullet_style.font.size = Pt(11)
        # Set paragraph format for bullet
        pf = list_bullet_style.paragraph_format
        pf.left_indent = Inches(0.25)
        pf.first_line_indent = Inches(-0.25)
    
    # Check if List Number exists, if not create it
    try:
        styles['List Number']
    except KeyError:
        # Create List Number style
        from docx.enum.style import WD_STYLE_TYPE
        list_number_style = styles.add_style('List Number', WD_STYLE_TYPE.PARAGRAPH)
        list_number_style.base_style = styles['Normal']
        list_number_style.font.name = 'Calibri'
        list_number_style.font.size = Pt(11)
        # Set paragraph format for numbering
        pf = list_number_style.paragraph_format
        pf.left_indent = Inches(0.25)
        pf.first_line_indent = Inches(-0.25)
    
    # Check if List Number 2 exists, if not create it
    try:
        styles['List Number 2']
    except KeyError:
        # Create List Number 2 style (deeper indentation level)
        from docx.enum.style import WD_STYLE_TYPE
        list_number_2_style = styles.add_style('List Number 2', WD_STYLE_TYPE.PARAGRAPH)
        list_number_2_style.base_style = styles['Normal']
        list_number_2_style.font.name = 'Calibri'
        list_number_2_style.font.size = Pt(11)
        pf = list_number_2_style.paragraph_format
        pf.left_indent = Inches(0.5)
        pf.first_line_indent = Inches(-0.25)

def add_hyperlink(paragraph, url, text):
    """Add a hyperlink to a paragraph"""
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create a new run object (a wrapper over a <w:r> element)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Add formatting for hyperlink (blue + underline)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')  # Blue color
    rPr.append(color)
    
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    
    # Set font
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Calibri (Body)')
    rPr.append(rFonts)
    
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '18')  # 9pt = 18 half-points
    rPr.append(sz)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink


def _clear_header_footer_part(part):
    """Remove existing content from a header/footer part."""
    if not part:
        return
    for tbl in list(part.tables):
        tbl._element.getparent().remove(tbl._element)
    for paragraph in list(part.paragraphs):
        parent = paragraph._element.getparent()
        if parent is not None:
            parent.remove(paragraph._element)


def _get_logo_stream(logo_bytes: Optional[bytes], fallback_path: str | None) -> io.BytesIO | None:
    """Return a BytesIO for the supplied logo, loading fallback path if needed."""
    if logo_bytes:
        return io.BytesIO(logo_bytes)
    if fallback_path and os.path.exists(fallback_path):
        with open(fallback_path, "rb") as fh:
            return io.BytesIO(fh.read())
    return None


def create_header_footer(
    doc: Document,
    client_name: str,
    project_name: str,
    client_logo_bytes: Optional[bytes] = None,
    falcon_logo_bytes: Optional[bytes] = None,
) -> None:
    """Apply Falcon-branded header and footer to every section after composition."""

    falcon_fallback_path = "FIXED_IMAGE/Falcon-Autotech_Logo-removebg-preview.png"

    for index, section in enumerate(doc.sections):
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        first_section = index == 0
        section.different_first_page_header_footer = first_section

        if first_section:
            _clear_header_footer_part(getattr(section, "first_page_header", None))
            _clear_header_footer_part(getattr(section, "first_page_footer", None))

        header = section.header
        footer = section.footer
        _clear_header_footer_part(header)
        _clear_header_footer_part(footer)

        header_table = header.add_table(rows=1, cols=3, width=Inches(6.5))
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        left_cell, middle_cell, right_cell = header_table.rows[0].cells
        left_cell.width = Inches(1.3)
        middle_cell.width = Inches(4.0)
        right_cell.width = Inches(1.3)

        client_logo_stream = _get_logo_stream(client_logo_bytes, None)
        if client_logo_stream:
            left_para = left_cell.paragraphs[0]
            left_run = left_para.add_run()
            client_logo_stream.seek(0)
            left_run.add_picture(client_logo_stream, height=Inches(0.6))
            left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        middle_para = middle_cell.paragraphs[0]
        middle_run = middle_para.add_run(f"FALCON's Proposal to {client_name} for the {project_name}")
        middle_run.font.name = "Calibri"
        middle_run.font.size = Pt(9)
        middle_run.font.color.rgb = HEADING_COLOR
        middle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        falcon_logo_stream = _get_logo_stream(falcon_logo_bytes, falcon_fallback_path)
        if falcon_logo_stream:
            right_para = right_cell.paragraphs[0]
            right_run = right_para.add_run()
            falcon_logo_stream.seek(0)
            right_run.add_picture(falcon_logo_stream, height=Inches(0.6))
            right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        for row in header_table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                tcBorders = OxmlElement("w:tcBorders")
                for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
                    border = OxmlElement(f"w:{border_name}")
                    border.set(qn("w:val"), "none")
                    tcBorders.append(border)
                tcPr.append(tcBorders)

        header.add_paragraph().paragraph_format.space_before = Pt(3)

        footer.add_paragraph().paragraph_format.space_after = Pt(3)
        footer_para = footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        run = footer_para.add_run("© FALCON AUTOTECH 2025 Confidential: Not for Distribution. ")
        run.font.name = "Calibri (Body)"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 0, 0)

        add_hyperlink(footer_para, "https://www.falconautotech.com/", "https://www.falconautotech.com/")

        page_run = footer_para.add_run(" | Page ")
        page_run.font.name = "Calibri (Body)"
        page_run.font.size = Pt(9)
        page_run.font.color.rgb = RGBColor(0, 0, 0)

        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = "PAGE"
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        page_run._r.append(fld_begin)
        page_run._r.append(instr)
        page_run._r.append(fld_end)

        total_run = footer_para.add_run(" of ")
        total_run.font.name = "Calibri (Body)"
        total_run.font.size = Pt(9)
        total_run.font.color.rgb = RGBColor(0, 0, 0)

        fld_total_begin = OxmlElement("w:fldChar")
        fld_total_begin.set(qn("w:fldCharType"), "begin")
        instr_total = OxmlElement("w:instrText")
        instr_total.set(qn("xml:space"), "preserve")
        instr_total.text = "NUMPAGES"
        fld_total_end = OxmlElement("w:fldChar")
        fld_total_end.set(qn("w:fldCharType"), "end")
        total_run._r.append(fld_total_begin)
        total_run._r.append(instr_total)
        total_run._r.append(fld_total_end)


def create_cover_page(
    client_logo: Optional[bytes],
    client_name: str,
    project_title: str,
) -> io.BytesIO:
    """Create a cover page using template - exactly as in main.py"""
    template_path = "FIXED_IMAGE\\Cover_Temp.docx"
    doc = Document(template_path)

    # Remove all headers and footers from template
    for sec in doc.sections:
        for part in (
            getattr(sec, "header", None),
            getattr(sec, "footer", None),
            getattr(sec, "first_page_header", None),
            getattr(sec, "first_page_footer", None),
            getattr(sec, "even_page_header", None),
            getattr(sec, "even_page_footer", None),
        ):
            if not part:
                continue
            try:
                part.is_linked_to_previous = False
            except Exception:
                pass
            try:
                for tbl in list(part.tables):
                    tbl._element.getparent().remove(tbl._element)
                for p in list(part.paragraphs):
                    p._element.getparent().remove(p._element)
            except Exception:
                pass

    # Add client logo if provided - process with PIL to ensure proper embedding
    if client_logo:
        try:
            # Open and process image
            im = Image.open(io.BytesIO(client_logo))
            if im.mode != "RGBA":
                im = im.convert("RGBA")
            alpha = im.getchannel("A")
            bbox = alpha.getbbox()
            if bbox:
                im = im.crop(bbox)
                alpha = im.getchannel("A")
            # Create white background and paste
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=alpha)

            # Save to buffer
            buf = io.BytesIO()
            bg.save(buf, format="PNG")
            buf.seek(0)

            # Insert at beginning
            first_para = doc.paragraphs[0]
            run_logo = first_para.insert_paragraph_before().add_run()
            run_logo.add_picture(buf, width=Inches(2.0))
        except Exception:
            # Fallback: insert without processing
            first_para = doc.paragraphs[0]
            run_logo = first_para.insert_paragraph_before().add_run()
            run_logo.add_picture(io.BytesIO(client_logo), width=Inches(2.0))

    # Add spacing
    for _ in range(6):
        doc.add_paragraph("")

    # Add title
    title = f"FALCON's Proposal to {client_name} for the {project_title}"
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.font.size = Pt(24)
    run.font.bold = False
    run.font.name = "Calibri"
    run.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add date
    today_str = datetime.today().strftime("%B %d, %Y")
    p2 = doc.add_paragraph()
    run2 = p2.add_run(today_str)
    run2.font.size = Pt(14)
    run2.font.name = "Calibri"
    run2.font.color.rgb = RGBColor(255, 215, 0)
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add page break after cover page
    doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==================== ADDITIONAL HELPER FUNCTIONS ====================

def extract_pdf_text(uploaded_file) -> str:
    """Extract plain text from an uploaded PDF using pdfplumber."""
    if uploaded_file is None:
        return ""

    text_chunks = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text_chunks.append(page.extract_text() or "")

    full_text = "\n\n".join(text_chunks)
    # Hard truncate to keep prompt size reasonable
    if len(full_text) > 20000:
        full_text = full_text[:20000]
    return full_text

def choose_sorter_template(project_name: str) -> SorterTemplate:
    """Pick the closest template based on project name keywords."""
    text = (project_name or "").lower()

    # score by number of keyword hits
    best_tpl = SORTER_TEMPLATES[0]
    best_score = -1
    for tpl in SORTER_TEMPLATES:
        score = sum(1 for kw in tpl.keywords if kw in text)
        if score > best_score:
            best_score = score
            best_tpl = tpl

    return best_tpl

def shade_cell(cell, color_hex: str = "D9D9D9"):
    """Apply gray shading to a table cell"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)

def add_markdown_line(doc: Document, line: str):
    """Add paragraph with **bold** segments."""
    p = doc.add_paragraph()
    parts = line.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        if i % 2 == 1:
            run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
    return p

def add_markdown_paragraph(doc: Document, text: str, style: str | None = None):
    """Add a paragraph with simple **bold** Markdown handling."""
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()

    parts = re.split(r"(\*\*[^\*]+\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = p.add_run(part[2:-2])
            run.bold = True
        else:
            run = p.add_run(part)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
    return p

def add_boxed_text(doc: Document, text: str, font_size: int = 16, bold: bool = True):
    """Grey shaded single-cell table with centered text (for cover letter front page)."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    shade_cell(cell, "D9D9D9")
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)
    p.space_before = Pt(6)
    p.space_after = Pt(6)
    return table

def add_centered_upload_image(doc: Document, uploaded_file, width_in: float = 6.0):
    """Add uploaded image centered"""
    if not uploaded_file:
        return
    img_stream = BytesIO(uploaded_file.getvalue())
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_picture(img_stream, width=Inches(width_in))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def call_groq_cover_letter(
    client_name: str,
    project_title: str,
    offer_ref: str,
    letter_date_str: str,
    executives_block: str,
    invitation_date: str,
    meeting_date: str,
    sender_name: str,
    sender_title: str,
    process_flow_summary: str = "",
) -> str:
    """Call Groq API to generate the cover letter text."""
    user_prompt = COVER_LETTER_USER_PROMPT_TEMPLATE.format(
        client_name=client_name,
        project_title=project_title,
        offer_ref=offer_ref,
        letter_date=letter_date_str,
        executives_block=executives_block.strip() or "Not provided",
        invitation_date=invitation_date.strip() or "Not provided",
        meeting_date=meeting_date.strip() or "Not provided",
        process_flow_summary=process_flow_summary.strip() or "Not provided",
        sender_name=sender_name,
        sender_title=sender_title,
    )

    def api_call():
        return groq_client.chat.completions.create(
            model="groq/compound",
            messages=[
                {"role": "system", "content": COVER_LETTER_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.3,
            max_tokens=600,  # Increased to allow for process flow details while staying under 300 words
        )
    
    completion = call_groq_with_retry(api_call)
    text = completion.choices[0].message.content.strip()
    if text.startswith("```"):
        parts = text.split("```")
        if len(parts) >= 2:
            text = parts[1]
            if text.startswith("text\n") or text.startswith("markdown\n"):
                text = "\n".join(text.split("\n")[1:])
    return text.strip()

def call_groq_exec_summary(system_text: str, client_name: str, project_title: str) -> str:
    """Call Groq API to generate the Executive Summary text."""
    user_content = (
        f"Client Name: {client_name}\n"
        f"Project / System Name: {project_title}\n\n"
        f"Proposed System Description (for context):\n{system_text}\n\n"
        "Generate the Executive Summary strictly as per the instructions."
    )

    def api_call():
        return groq_client.chat.completions.create(
            model="groq/compound",
            temperature=0.4,
            max_tokens=800,
            messages=[
                {"role": "system", "content": EXEC_SUMMARY_SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
        )
    
    resp = call_groq_with_retry(api_call)
    return resp.choices[0].message.content.strip()

def call_groq_for_system_description(process_flow: str, dxf_json: dict, project_name: str) -> str:
    """Generate comprehensive system description using Groq API"""
    # Convert DXF JSON to string for prompt
    dxf_info = json.dumps(dxf_json, indent=2, ensure_ascii=False)
    
    user_prompt = f"""Generate a COMPREHENSIVE, DETAILED system description for:

PROJECT NAME: {project_name}

PROCESS FLOW:
{process_flow}

DXF FILE INFORMATION:
{dxf_info}

REQUIREMENTS:
1. Extract ALL quantities from the DXF data (chutes, operators, leg guards, fencing, pallets)
2. Use these exact numbers in the appropriate sections
3. Generate EXTENSIVE descriptions for each component (multiple paragraphs), Add Table if needed.
4. Only include sections for components mentioned in process flow or present in DXF data
5. Write 2000-3000 words with technical depth matching professional engineering documentation
6. Each major section should have 3-4 paragraphs with subsections having 2-4 paragraphs
7. Each component description should have 3-5 sentences explaining functionality, design, and purpose
8. Each subheading will be in bold format example: **Conveyor System**
9. Do NOT ADD ```json`` or any other code block formatting in the output

Generate the detailed system description now."""

    def api_call():
        return groq_client.chat.completions.create(
            messages=[
                {
                    "role": "system",
                    "content": ENHANCED_SYSTEM_DESCRIPTION_PROMPT
                },
                {
                    "role": "user",
                    "content": user_prompt
                }
            ],
            model="llama-3.3-70b-versatile",
            temperature=0.3,
            max_tokens=3000,
            top_p=0.9
        )
    
    resp = call_groq_with_retry(api_call)
    return resp.choices[0].message.content.strip()


def generate_system_description_from_sd_sys(dxf_path: Path, costing_file_upload, project_name: str,
                                           client_name: str = "", pph_count: str = "", ipp_rate: str = "") -> tuple[str, dict, dict, dict]:
    """
    Wrapper that uses `sys_desc_latest.py` functionality to generate a final System Description text.

    Returns a tuple: (system_description_text, detected, tables, cost_vals)
        # NOTE: System Description section replaced by sys_desc_latest integration.
        # The legacy formatting/LLM output parsing was removed so the sys_desc_latest
        # pipeline (called elsewhere) will supply `system_description_text` and
        # that text will be added into the main document using the new flow.

1) Identify the main commercial building blocks, such as:
   - Conveyors Package
   - Cross Belt Sorter Package
   - Destinations Package
   - Services Package
   - Loop CBS + Inducts
   - Infeed + Bagging Conveyors
   - Output Chutes
   - Software package / Software Packages & SCADA
   - Steelworks / Steelwork
   - PTL
   - Installation & Commissioning
   - Project Management and Engineering Charges
   - Packaging & forwarding / Packaging & Documentation
   - Freight, Warranty, Hotline, AMC packages
   or similar high-level components used to summarize the cost.

2) For each such high-level component, return:
   - s_no: integer starting from 1 in sequence
   - label: the package/component name in clean human-readable form
   - price: the final total price for that component AS A STRING,
            including currency symbol and formatting exactly as in the sheet
            (e.g. "₹ 28,34,27,926", "SAR 10,959,208", "$ 398,105", "€ 2,871,416", "Included in CBS Price").

   IMPORTANT:
   - Do NOT invent prices.
   - Use values that actually appear in the sheet.
   - If multiple detailed rows roll up into one package, use the rolled-up total that clearly corresponds to that package.
   - Prefer the same format as used for the final summary in the data, if visible.

3) If there is a grand "Total" (for the whole solution), also return:
   - total_row: { "label": "...", "price": "..." }
   For example: { "label": "Total", "price": "₹ 64,80,90,527" }.
   If no obvious total exists, set total_row to null.

4) If there are explicit discount and final total lines (like in the Delhivery example):
   - cooperation_label: e.g. "Business Cooperation Agreement"
   - discount_label: e.g. "Discount for Delhivery"
   - discount_value: e.g. "4.5%"
   - final_total_label: e.g. "Final Total"
   - final_total_value: e.g. "₹ 61,89,26,453"
   If not present, return them as null.

5) Also return:
   - currency: "INR", "SAR", "USD", "EUR", or "MIXED" if multiple currencies appear.
   - price_sheet_title: a short label like "19.1 Price List – Summary" or "20.1 Price Sheet"
                        if visible; otherwise null.

6) OUTPUT FORMAT (VERY IMPORTANT):

Return ONLY a single valid JSON object with this exact shape:

{
  "currency": "INR" | "SAR" | "USD" | "EUR" | "MIXED" | null,
  "price_sheet_title": "string or null",
  "items": [
    {
      "s_no": 1,
      "label": "Conveyors Package",
      "price": "₹ 28,34,27,926"
    },
    ...
  ],
  "total_row": {
    "label": "Total",
    "price": "₹ 64,80,90,527"
  } or null,
  "cooperation_label": "string or null",
  "discount_label": "string or null",
  "discount_value": "string or null",
  "final_total_label": "string or null",
  "final_total_value": "string or null"
}

Do NOT wrap the JSON or ```json``` in markdown.
Do NOT add explanations or commentary.
Just return the JSON object.
"""

# System prompt used for extracting price-sheet JSON from costing CSV
GROQ_SYSTEM_PROMPT = """
You are an expert data extraction assistant. You will be given a CSV-like dump of an internal costing sheet.
Your job is to extract and return a single valid JSON object exactly matching the schema requested by the user.
Return ONLY the JSON object and nothing else. Do not include markdown, commentary or any extra text.
If a field is missing, set it to null. Be strict about types.
"""

GROQ_USER_PROMPT_TEMPLATE = """
Below is the raw CSV export of the 'Overall Costing' sheet of an internal costing file.

Use it to construct the high-level Price Sheet summary as described in the instructions.

Raw CSV:
--------------------
{sheet_csv}
--------------------
"""

def call_groq_for_price_sheet(sheet_csv: str) -> dict:
    """Call Groq API to extract price sheet from costing CSV"""
    user_prompt = GROQ_USER_PROMPT_TEMPLATE.format(sheet_csv=sheet_csv)

    def api_call():
        return groq_client.chat.completions.create(
            model="groq/compound",
            messages=[
                {"role": "system", "content": GROQ_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.0,
        )
    
    completion = call_groq_with_retry(api_call)
    raw = completion.choices[0].message.content.strip()

    # Strip markdown fences if present
    if raw.startswith("```"):
        parts = raw.split("```")
        if len(parts) >= 2:
            raw = parts[1]
            raw = raw.lstrip("json").lstrip()

    # Extract JSON from first '{' to last '}'
    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end < start:
        raise ValueError(f"Groq response does not contain a JSON object:\n{raw}")

    content = raw[start:end + 1]

    try:
        data = json.loads(content)
    except json.JSONDecodeError as e:
        raise ValueError(f"Groq response was not valid JSON: {e}\nExtracted content:\n{content}")

    return data

def parse_price_string(price_str: str):
    """Extract currency prefix and numeric value from price string"""
    if not price_str:
        return None, None, 0

    m = re.search(r"[-]?\d", price_str)
    if not m:
        return price_str.strip(), None, 0

    prefix = price_str[:m.start()].strip()
    numeric_part = price_str[m.start():].strip()

    digits_only = "".join(ch for ch in numeric_part if ch.isdigit() or ch == ".")
    if digits_only == "":
        return prefix, None, 0

    decimals_count = 0
    if "." in digits_only:
        decimals_count = len(digits_only.split(".")[1])

    try:
        value = float(digits_only)
    except ValueError:
        return prefix, None, decimals_count

    return prefix, value, decimals_count

def format_indian_number(value: float, decimals: int) -> str:
    """Format number with Indian-style digit grouping"""
    if decimals > 0:
        s = f"{value:.{decimals}f}"
    else:
        s = f"{int(round(value))}"

    if "." in s:
        int_part, frac = s.split(".")
    else:
        int_part, frac = s, None

    # Indian grouping
    if len(int_part) > 3:
        last3 = int_part[-3:]
        head = int_part[:-3]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        int_formatted = ",".join(groups + [last3])
    else:
        int_formatted = int_part

    if frac and decimals > 0:
        return int_formatted + "." + frac
    else:
        return int_formatted

def apply_bca_discount_to_price_data(price_data: dict, discount_percent: float) -> str | None:
    """Apply BCA discount on total_row.price and return discounted price string"""
    total_row = price_data.get("total_row")
    if not total_row:
        return None

    price_str = total_row.get("price")
    prefix, value, decimals = parse_price_string(price_str)
    if value is None:
        return None

    discounted = value * (1 - discount_percent / 100.0)
    formatted_number = format_indian_number(discounted, decimals)
    if prefix:
        return f"{prefix} {formatted_number}"
    else:
        return formatted_number

# ==================== CAPACITY CALCULATIONS FUNCTIONS ====================

def build_capacity_prompt_from_excel(
    excel_bytes: bytes,
    client_name: str,
    project_name: str
) -> str:
    """
    Read the uploaded Excel (all sheets), dump them as CSV text,
    and build a very explicit extraction prompt for GROQ.
    We do NOT try to interpret any cell ourselves.
    """
    xls = pd.ExcelFile(BytesIO(excel_bytes))

    sheet_dumps = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        # Keep as CSV-like text to preserve structure
        csv_text = df.to_csv(index=False, header=False)
        sheet_dumps.append(f"### Sheet: {sheet}\n{csv_text}")

    workbook_text = "\n\n".join(sheet_dumps)

    # IMPORTANT: we define a strict JSON schema and explicitly
    # tell GROQ to set fields to null if they are missing.
    prompt = f"""
You are an expert in interpreting throughput and capacity calculation Excel sheets
for parcel/shipment sortation systems (Loop CBS, Linear CBS, Cross Belt Sorters, etc.).

You are given a raw text dump of the complete Excel workbook used for capacity calculations.
Using ONLY the information present in the workbook (numbers and labels), you must extract
or compute the key capacity fields and return them as a single JSON object.

Context:
- Client: {client_name}
- Project: {project_name}

The workbook text follows after this instruction. It is a concatenation of all sheets, each
in CSV-like form.

IMPORTANT RULES:

1. **Use exact numbers from the workbook wherever a field is explicitly present.**
   - If a value is written in the sheet (e.g. "Sorter Speed 2 m/s", "Carrier per hour 6128"),
     prefer the sheet value instead of recomputing it.
2. **Only compute** a value if:
   - It is clearly implied (e.g. carrier_per_hour = speed_mps * 3600 / pitch_m) AND
   - It is NOT already available as a direct cell value.
3. If a field is not given and cannot be safely derived, set it explicitly to null.

KEY FIELDS (SEMANTICS):

- sorter_type:
    A short human-readable description like "Loop CBS", "Linear CBS", "Dual Belt Loop CBS"
    or "Cross Belt Sorter". Use what best matches the workbook text.

- sorter_speed_mps:
    Sorter speed in meters per second. If sheet says "Speed 2 m/s", set 2.0.

- pitch_m:
    Carrier pitch in meters. If sheet says "Pitch 1,175 mm", then pitch_m = 1.175.

- carriers_per_hour_cph:
    "Carrier per hour" / "Carriers/Hour" / "Carriers per hour" from the sheet.
    If not present, you may compute as:
      carriers_per_hour = speed_mps * 3600 / pitch_m
    and round to nearest integer.

- belts_per_hour_bph:
    "Belts per hour" / "Belts/Hour" from the sheet.
    If not present but the sorter is clearly Dual Belt, you may compute:
      belts_per_hour = carriers_per_hour * 2
    If single belt, belts_per_hour = carriers_per_hour.

- num_feedlines:
    Number of feedlines / inducts / infeed lines (e.g. "No of Feedlines", "No of Inducts").
    If the workbook has multiple such numbers, choose the one used in the capacity section.

- num_operators:
    Number of operators used in capacity calculations, if explicitly given
    (e.g. "No of Operators", "No of operators on manual induct station").
    If not given, set null.

- capacity_per_operator_pph:
    Capacity per operator in parcels/shipments per hour, if explicitly given
    (e.g. "Capacity per operator 1000 Shipments per hour"). If not given, set null.

- sorter_designed_capacity_A_pph:
    Sorter designed capacity on the parcel spectrum. Look for labels like:
    "Sorter Designed Capacity (A)", "Effective Designed Throughput of Sorter (A)",
    "Effective Designed TPH", or similar. Use the PPH/Shipments per hour value.

- feedline_designed_capacity_B_pph:
    Total feedline/induction capacity. Look for labels like:
    "Total Feedline designed capacity (B)", "Induction Capacity (B)",
    "Total Induction Capacity Designed", etc. Use the PPH value.

- effective_capacity_min_AB_pph:
    The effective designed capacity of the system.
    If the sheet already has "System designed throughput" or "Operational capacity",
    use that value.
    If not explicitly given, compute:
       effective_capacity_min_AB_pph = min(sorter_designed_capacity_A_pph,
                                           feedline_designed_capacity_B_pph)
    (if both are known).

- single_belt_pct and dual_belt_pct:
    Percentages of shipments handled on single and dual belts, if present
    (e.g. "Single Belts Shipments 91.36%", "Dual Belt Shipments 8.64%").
    Store them as numeric percentages (e.g. 91.36, 8.64).
    If not present, set them to null.

JSON SCHEMA (MANDATORY KEYS):

You MUST return exactly one JSON object with ALL of these keys:

{{{{
  "sorter_type": "Loop CBS or Linear CBS or Cross Belt Sorter etc.",
  "sorter_speed_mps": 2.0,
  "pitch_m": 1.175,
  "carriers_per_hour_cph": 0,
  "belts_per_hour_bph": 0,
  "num_feedlines": 0,
  "num_operators": null,
  "capacity_per_operator_pph": null,
  "sorter_designed_capacity_A_pph": 0,
  "feedline_designed_capacity_B_pph": 0,
  "effective_capacity_min_AB_pph": 0,
  "single_belt_pct": null,
  "dual_belt_pct": null
}}}}

RESPONSE FORMAT REQUIREMENTS (CRITICAL):

- Output MUST be **only** a JSON object.
- Do NOT include markdown, explanations, or any text outside the JSON.
- All numeric values must be raw numbers (no units, no commas, no % signs).
- If a value is unknown or not present, set it to null (not 0).
- DO NOT add ```json``` or json in the response. ONLY return raw JSON.
**MUST : DO NOT add ```json``` or json in the response. ONLY return raw JSON.**
Below is the full workbook dump:

{workbook_text}
"""
    return prompt


def call_groq_for_capacity(prompt: str) -> dict:
    """
    Call GROQ with response_format=json_object so that we reliably get JSON.
    """
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY is not set in environment variables.")

    client = Groq(api_key=GROQ_API_KEY)

    def api_call():
        return client.chat.completions.create(
            model="groq/compound",
            messages=[
                {
                    "role": "system",
                    "content": "You are a precise JSON data extractor. Always follow the schema exactly."
                },
                {
                    "role": "user",
                    "content": prompt
                },
            ],
            response_format={"type": "json_object"},
            temperature=0.0,
        )
    
    chat_completion = call_groq_with_retry(api_call)
    raw = chat_completion.choices[0].message.content
    return json.loads(raw)


def call_groq_for_sorter_spec(sheet_name: str, sheet_text: str) -> dict:
    """
    Call GROQ API to extract sorter technical specifications from Loop CBS Excel sheet.
    Returns a JSON object with sorter specifications.
    """
    system_prompt = """You are a technical proposal engineer reading an Excel costing/configuration sheet
for a Loop Cross Belt Sorter ("Loop CBS").

You will receive:
- The sheet name (e.g., "Loop CBS")
- The ENTIRE sheet content as text, row by row, including ALL tables.

Your task:
Extract the following fields, strictly from the sheet content:

1) sorter_carrier_type        – fixed string "Loop CBS".
2) sorter_speed_mps           – fixed string "upto 2 m/s"
3) sorter_loop_length_m       – sorter loop length in meters (e.g., "150").
4) sorter_height_mm           – sorter height in mm (e.g., "2900").
5) actuation_technology       – fixed string "Electric".
6) carrier_pitch_mm           – carrier pitch in mm (e.g., "600", "1175", "1200"). If multiple models,
                                choose the one actually selected in the configuration area.
7) number_of_carriers         – total number of carriers in the sorter, as shown in the sheet if present.
8) motor_drive_type           – description of motor/drive type (e.g., "LIM", "LSM", "LIM + LSM", etc.).
9) power_consumption          – sorter power consumption (kW or kVA etc.) taken from the sheet.

VERY IMPORTANT RULES:
- Use ONLY information present in the provided sheet text. Do NOT guess or invent values.
- If a value is not clearly present, set it to null.
- If the sheet expresses a choice ("Select Model", "Enter Loop Length", etc.), use the chosen values.
- Keep the output values SHORT: just the numeric value or the short phrase, without explanations.

Return a single JSON object with EXACTLY these keys:

{
  "sorter_carrier_type": "...",
  "sorter_speed_mps": "... or null",
  "sorter_loop_length_m": "... or null",
  "sorter_height_mm": "... or null",
  "actuation_technology": "...",
  "carrier_pitch_mm": "... or null",
  "number_of_carriers": "... or null",
  "motor_drive_type": "... or null",
  "power_consumption": "... or null"
}

No comments, no trailing text, no markdown."""

    user_payload = {
        "sheet_name": sheet_name,
        "sheet_text": sheet_text,
    }

    def api_call():
        return groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt.strip()},
                {"role": "user", "content": json.dumps(user_payload, indent=2)},
            ],
            response_format={"type": "json_object"},
            temperature=0.1,
        )

    try:
        resp = call_groq_with_retry(api_call)
        text = resp.choices[0].message.content.strip()
        spec = json.loads(text)
        
        if not isinstance(spec, dict):
            raise ValueError("Expected a JSON object")
        
        # Fill fixed fields if missing
        if not spec.get("sorter_carrier_type"):
            spec["sorter_carrier_type"] = "Loop CBS"
        if not spec.get("actuation_technology"):
            spec["actuation_technology"] = "Electric"
        
        return spec
    except Exception as exc:
        st.error(f"Failed to parse GROQ response for sorter specification: {exc}")
        return None


def add_capacity_section_to_doc(
    doc: Document,
    client_name: str,
    project_name: str,
    cap: dict,
    counter: int
) -> None:
    """
    Add 'Sorter System Capacity' section to an existing Document,
    using the extracted capacity dict.
    """
    # Heading
    add_numbered_heading(doc, "Sorter System Capacity", counter=counter)

    intro_para = (
        f"The following table shows the throughput calculation for the sortation system "
        f"designed based on {client_name}'s {project_name} requirements."
    )
    p = doc.add_paragraph(intro_para)
    apply_normal_style(p)

    # Table: SPECIFICATION | VALUE
    table = doc.add_table(rows=1, cols=2)
    table.style = "Medium Shading 1 Accent 1"

    hdr = table.rows[0].cells
    hdr[0].text = "SPECIFICATION"
    hdr[1].text = "VALUE"

    def fmt(value, suffix=""):
        if value is None or value == "":
            return "N/A"
        return f"{value}{suffix}"

    def add_row(label, value):
        row = table.add_row().cells
        row[0].text = label
        row[1].text = value

    # Fill rows
    add_row("Sorter Type", cap.get("sorter_type", ""))

    # Speed / pitch
    add_row("Sorter Speed", fmt(cap.get("sorter_speed_mps"), " m/s"))
    add_row("Pitch", fmt(cap.get("pitch_m"), " m"))

    # Capacity raw
    add_row("Carrier per hour", fmt(cap.get("carriers_per_hour_cph"), " CPH"))
    add_row("Belts per hour", fmt(cap.get("belts_per_hour_bph"), " BPH"))

    # Feedlines / operators
    add_row("No. of Feedlines", fmt(cap.get("num_feedlines")))
    add_row("No. of Operators", fmt(cap.get("num_operators")))
    add_row("Capacity per Operator", fmt(cap.get("capacity_per_operator_pph"), " PPH"))

    # Sorter vs Feedline capacity
    add_row(
        "Sorter Designed Capacity (A)",
        fmt(cap.get("sorter_designed_capacity_A_pph"), " PPH"),
    )
    add_row(
        "Feedline Designed Capacity (B)",
        fmt(cap.get("feedline_designed_capacity_B_pph"), " PPH"),
    )
    add_row(
        "Effective Designed Capacity (min of A & B)",
        fmt(cap.get("effective_capacity_min_AB_pph"), " PPH"),
    )

    # Optional single / dual belt %
    if cap.get("single_belt_pct") is not None or cap.get("dual_belt_pct") is not None:
        add_row(
            "Single Belt Shipments",
            fmt(cap.get("single_belt_pct"), " %"),
        )
        add_row(
            "Dual Belt Shipments",
            fmt(cap.get("dual_belt_pct"), " %"),
        )

# ==================== INPUT COLLECTION ====================

# Professional Tabs for Input Organization
# ==================== INPUT COLLECTION ====================

# Section 1: Project & Client Information
render_section_header("Section 1: Project & Client Information")

col1, col2 = st.columns([2, 1])

with col1:
    project_name = st.text_input("Project Name *", value="Automated Sorting System", placeholder="Enter project name")
    offer_ref = st.text_input("Offer Reference No *", value="F24-00524", placeholder="e.g., F24-00524")
    
    # Client dropdown with add new option
    client_options = list(CLIENT_LOGOS.keys()) + ["+ Add New Client"]
    selected_client = st.selectbox("Client Name *", client_options, index=0)
    
    # Handle new client addition
    if selected_client == "+ Add New Client":
        client_name = st.text_input("Enter New Client Name *", placeholder="Enter client name")
        client_logo = st.file_uploader("Upload Client Logo *", type=["png", "jpg", "jpeg"], key="new_client_logo")
        client_logo_path_display = None
    else:
        client_name = selected_client
        client_logo = None
        client_logo_path_display = CLIENT_LOGOS.get(selected_client)
    
    executives_text = st.text_area(
        "Client Executives (one per line, include Mr./Ms.) *",
        value="Mr. Rahul Didwani\nMr. Vinayak Garg",
        height=80,
        placeholder="Mr. John Doe\nMs. Jane Smith"
    )
    # Operational inputs used by System Description generation
    pph_count = st.text_input("PPH COUNT", value="", placeholder="Parcels per hour (e.g. 2000)")
    ipp_rate = st.text_input("IPP RATE", value="", placeholder="IPP Rate (e.g. 1200)")
    
    col1a, col1b = st.columns(2)
    with col1a:
        invitation_date = st.date_input("Invitation Date (optional)", value=None)
    with col1b:
        meeting_date = st.date_input("Meeting/Workshop Date (optional)", value=None)
    
    st.markdown("**Contact Person Details**")
    col1c, col1d = st.columns(2)
    with col1c:
        contact_name = st.text_input("Name *", value="Sanyog Pratap Singh")
        contact_phone = st.text_input("Phone *", value="+91 8750052591")
    with col1d:
        contact_email = st.text_input("Email *", value="Sanyog.Singh@falconautotech.com")
        st.write("")  # Spacer

with col2:
    st.markdown("**Client Logo Preview**")
    if client_logo_path_display and os.path.exists(client_logo_path_display):
        st.image(client_logo_path_display, width='stretch')
    elif client_logo:
        st.image(client_logo, width='stretch')
    else:
        st.info("Logo will appear here")

# Fixed values (not shown to user)
letter_date = date.today()
sender_name = "Sandeep Bansal"
sender_title = "Chief Business Officer"
invitation_date_str = invitation_date.strftime("%B %d, %Y") if invitation_date else ""
meeting_date_str = meeting_date.strftime("%B %d, %Y") if meeting_date else ""

st.markdown("---")

# Section 2: Upload Files
render_section_header("Section 2: Upload Files")

col1, col2 = st.columns(2)

with col1:
    dxf_layout_file = st.file_uploader("2.1 DXF Layout File *", type=["dxf"], key="dxf_upload")
    costing_file = st.file_uploader("2.2 Costing Sheet *", type=["xlsx", "xls"], key="costing_upload")

with col2:
    capacity_excel = st.file_uploader("2.3 Throughput Calculation Sheet *", type=["xlsx", "xls"], key="capacity_upload")
    prog_gantt = st.file_uploader("2.4 Project Timeline Chart (optional)", type=["png", "jpg", "jpeg"], key="gantt_upload")

st.markdown("")  # Spacer
have_solution_png = st.checkbox("I already have PNG of the solution", value=False)

if have_solution_png:
    layout_full_png = st.file_uploader("Upload your solution PNG here", type=["png", "jpg", "jpeg"], key="solution_png_upload")
else:
    layout_full_png = None

st.markdown("---")

# Section 2.5: Process Flow Preview & Feedback (only if DXF is uploaded)
if dxf_layout_file:
    render_section_header("Section 2.5: Process Flow Preview & Feedback")
    
    with st.expander("🔄 Preview & Refine Process Flow", expanded=False):
        st.markdown("""
        **Quick Preview & Feedback Loop**
        
        Generate a preview of the process flow from your DXF file. If you're not satisfied with the output, 
        provide feedback and regenerate until it meets your requirements.
        """)
        
        col_preview1, col_preview2 = st.columns([3, 1])
        
        with col_preview1:
            if st.button("🚀 Generate Process Flow Preview", key="generate_preview", width='stretch'):
                with st.spinner("Generating process flow preview..."):
                    try:
                        # Create temp directory and process DXF
                        tmp_dir = Path(tempfile.mkdtemp(prefix="preview_"))
                        dxf_path = tmp_dir / dxf_layout_file.name
                        dxf_path.write_bytes(dxf_layout_file.getvalue())
                        
                        # Extract DXF components
                        dxf_json = extract_dxf_components(dxf_path, project_name)
                        
                        # Call the refactored function to get process flow and iteration details
                        final_flow, iteration_details = call_groq_for_process_flow(client_name, project_name, dxf_json)
                        
                        # Save preview in session state
                        st.session_state['preview_process_flow'] = final_flow
                        st.session_state['iteration_details'] = iteration_details
                        
                        st.success("✅ Process Flow generated successfully!")
                        
                        # Display Iteration Details Expander
                        with st.expander("📊 Iteration Details & Scores (Process Flow Evolution)", expanded=True):
                            if iteration_details:
                                st.markdown("### 🔄 Refinement Process Summary")
                                
                                # Show INITIAL flow first
                                st.markdown("#### **Initial Process Flow**")
                                if iteration_details:
                                    initial_detail = iteration_details[0]
                                    st.markdown(f"**Score:** {initial_detail['score']:.2f} / {initial_detail['target_score']}")
                                    st.markdown("**Flow:**")
                                    st.text_area("Initial Flow", value=initial_detail['flow'], height=250, 
                                                disabled=False, key="initial_flow_display")
                                
                                st.markdown("---")
                                st.markdown("### 📈 Iteration-by-Iteration Refinement")
                                
                                # Show each iteration with full flow
                                for detail in iteration_details:
                                    iteration_num = detail['iteration']
                                    score = detail['score']
                                    target = detail['target_score']
                                    gap = detail['gap']
                                    flow = detail['flow']
                                    feedback = detail['feedback']
                                    
                                    # Create expander for each iteration
                                    is_success = score >= target
                                    status_icon = "✅" if is_success else "🔄"
                                    
                                    with st.expander(f"{status_icon} Iteration {iteration_num} | Score: {score:.2f}/{target} | Gap: {gap:.2f}"):
                                        col1, col2, col3 = st.columns(3)
                                        with col1:
                                            st.metric("Coherence Score", f"{score:.2f}")
                                        with col2:
                                            st.metric("Target Score", f"{target:.1f}")
                                        with col3:
                                            st.metric("Gap", f"{gap:.2f}")
                                        
                                        # Show feedback items if any
                                        if feedback:
                                            st.markdown("**💡 Feedback Applied:**")
                                            for i, fb in enumerate(feedback, 1):
                                                st.write(f"**{i}.** {fb}")
                                        else:
                                            st.info("✨ No feedback generated (already matches target)")
                                        
                                        # Show full process flow
                                        st.markdown("**📄 Process Flow:**")
                                        st.text_area(f"Flow_{iteration_num}", value=flow, height=300, 
                                                    disabled=False, key=f"flow_iter_{iteration_num}")
                                
                                # Summary table
                                st.markdown("---")
                                st.markdown("### 📊 Summary Table")
                                summary_data = []
                                for detail in iteration_details:
                                    status = "✅ Success" if detail['score'] >= detail['target_score'] else "🔄 Refining"
                                    summary_data.append({
                                        "Iteration": detail['iteration'],
                                        "Score": f"{detail['score']:.2f}",
                                        "Target": f"{detail['target_score']:.1f}",
                                        "Status": status,
                                        "Gap": f"{detail['gap']:.2f}",
                                        "Feedback Items": len(detail['feedback'])
                                    })
                                
                                import pandas as pd
                                df_summary = pd.DataFrame(summary_data)
                                st.dataframe(df_summary, use_container_width=True, hide_index=True)
                            else:
                                st.info("No iteration details available")
                        
                        st.markdown("---")
                        st.markdown("### 🏆 Final Process Flow (Best Iteration)")
                        
                        # Find best iteration details for context
                        best_iteration_detail = None
                        best_score = 0
                        for detail in iteration_details:
                            if detail['score'] > best_score:
                                best_score = detail['score']
                                best_iteration_detail = detail
                        
                        if best_iteration_detail:
                            st.markdown(f"**Best Score Achieved:** {best_iteration_detail['score']:.2f} / {best_iteration_detail['target_score']}")
                            st.markdown(f"**Achieved in Iteration:** {best_iteration_detail['iteration']}")
                        
                        st.markdown("**Optimized Process Flow:**")
                        st.code(final_flow)
                    except Exception as e:
                        st.error(f"Failed to generate process flow preview: {e}")
                        try:
                            st.exception(e)
                        except Exception:
                            pass

key_include = True
safety_include = True
infra_include = True
prog_include = True
client_resp_include = True
handover_include = True
commercial_include = True
warranty_include = True
exclusion_include = True
include_proposed_system = True
include_concept_desc = True
include_exec_summary = True
include_company_profile = True
include_handled_spectrum = True
include_capacity_section = True
elec_include = True
wcs_include= True
scada_include= True

st.markdown("---")

# Section 3: Edit/Confirm Settings
render_section_header("Section 3: Edit/Confirm Settings")

with st.expander("💰 Commercial Settings", expanded=False):
    apply_bca = st.checkbox("Apply Business Cooperation Agreement Discount (4.5%)", value=False, key="apply_bca_discount")
    
    st.markdown("**Payment Terms**")
    default_payment_terms = [
        {"Payment Percentage": "20%", "Stage": "Advance along with LOI/ PO"},
        {"Payment Percentage": "20%", "Stage": "After DAP Completion"},
        {"Payment Percentage": "40%", "Stage": "Before Dispatch"},
        {"Payment Percentage": "10%", "Stage": "Against Installation"},
        {"Payment Percentage": "10%", "Stage": "Against Handover"},
    ]
    
    if "payment_terms" not in st.session_state:
        st.session_state["payment_terms"] = default_payment_terms
    
    pt_df = pd.DataFrame(st.session_state["payment_terms"])
    edited_pt_df = st.data_editor(pt_df, num_rows="dynamic", width='stretch', key="payment_terms_editor")
    st.session_state["payment_terms"] = edited_pt_df.to_dict(orient="records")

with st.expander("📜 Warranty Configuration", expanded=False):
    warranty_type = st.selectbox("Warranty Type", ["Standard warranty", "Comprehensive warranty"], key="warranty_type")
    
    col1, col2 = st.columns(2)
    with col1:
        warranty_duration = st.text_input("Warranty Duration", value="1 year", key="warranty_duration")
    with col2:
        warranty_start = st.selectbox(
            "Warranty Start Condition",
            [
                "from the date of beneficiary use.",
                "from the date of commissioning of the system.",
                "from the date of completion of dispatch of the materials, whichever is earlier.",
                "from the date of beneficiary use, max 30 days after readiness of commissioning.",
                "from the date of official communication of material readiness at Falcon end.",
            ],
            key="warranty_start"
        )
    
    warranty_extended = st.checkbox("Include Extended Warranty Option", value=True, key="warranty_extended")
    if warranty_extended:
        warranty_extended_text = st.text_input(
            "Extended Warranty Text",
            value="Extended warranty of 2 years available on request @ 5% of the order value.",
            key="warranty_extended_text"
        )
    else:
        warranty_extended_text = None
    
    warranty_amc = st.checkbox("Include AMC / Hotline Clause", value=False, key="warranty_amc")
    if warranty_amc:
        warranty_amc_text = st.text_input(
            "AMC / Hotline Text",
            value="AMC / Hotline services available post warranty on demand.",
            key="warranty_amc_text"
        )
    else:
        warranty_amc_text = None
    
    warranty_transport = st.checkbox("Include Transportation Note", value=False, key="warranty_transport")
    if warranty_transport:
        warranty_transport_text = st.text_input(
            "Transportation Note",
            value="Transportation of defective parts to Falcon premises will be at client's cost.",
            key="warranty_transport_text"
        )
    else:
        warranty_transport_text = None

with st.expander("🎯 Feedback Rules Management", expanded=False):
    st.markdown("""
    **Manage Process Flow Feedback Rules**
    
    **🌍 Global Rules:** Saved permanently in `feedback_rules.json`, applied to ALL projects across all sessions  
    **📍 Local Rules:** Temporary, stored in current session only, disappear when you close the app
    """)
    
    # Get global rules
    all_global_rules = get_all_rules()
    
    # Get local rules from session
    local_rules = st.session_state.get('local_feedback_rules', [])
    
    # Get statistics
    stats = get_rule_stats()
    
    # Display tabs for Global vs Local
    tab_global, tab_local, tab_stats = st.tabs(["🌍 Global Rules", "📍 Local Rules", "📊 Statistics"])
    
    with tab_global:
        if not all_global_rules:
            st.info("No global rules saved yet. Create your first permanent rule from the preview section!")
        else:
            st.markdown(f"**Total Global Rules:** {len(all_global_rules)} | **Active:** {sum(1 for r in all_global_rules if r.get('active', True))}")
            st.markdown("---")
            
            # Display global rules
            for idx, rule in enumerate(all_global_rules):
                rule_id = rule.get("id", f"rule_{idx}")
                rule_type = rule.get("type", "global")
                is_active = rule.get("active", True)
                timestamp = rule.get("timestamp", "Unknown")
                user_fb = rule.get("user_feedback", "No feedback")
                rule_text = rule.get("rule", "No rule extracted")
                projects = rule.get("applied_to_projects", [])
                
                # Create a container for each rule
                status_emoji = "✅" if is_active else "🔴"
                type_badge = "🌍" if rule_type == "global" else "📍"
                
                with st.container():
                    col_rule1, col_rule2 = st.columns([4, 1])
                    
                    with col_rule1:
                        st.markdown(f"### {status_emoji} {type_badge} Rule ID: `{rule_id}`")
                        st.markdown(f"**Created:** {timestamp[:10] if len(timestamp) > 10 else timestamp}")
                        
                        if projects:
                            st.markdown(f"**Applied to:** {', '.join(projects[:3])}{' and more...' if len(projects) > 3 else ''}")
                        
                        with st.expander("📝 View Details", expanded=False):
                            st.markdown("**Original User Feedback:**")
                            st.text_area(f"Feedback_{rule_id}", value=user_fb, height=80, disabled=True, key=f"fb_{rule_id}", label_visibility="collapsed")
                            
                            st.markdown("**Extracted Rule:**")
                            
                            # Allow editing if active
                            if is_active:
                                new_rule = st.text_area(
                                    f"Rule_{rule_id}",
                                    value=rule_text,
                                    height=100,
                                    key=f"rule_edit_{rule_id}",
                                    label_visibility="collapsed"
                                )
                                
                                if new_rule != rule_text:
                                    if st.button(f"💾 Update Rule", key=f"update_{rule_id}"):
                                        if update_rule(rule_id, new_rule):
                                            st.success(f"Rule {rule_id} updated!")
                                            st.rerun()
                                        else:
                                            st.error("Failed to update rule")
                            else:
                                st.text_area(f"Rule_{rule_id}", value=rule_text, height=100, disabled=True, key=f"rule_view_{rule_id}", label_visibility="collapsed")
                    
                    with col_rule2:
                        st.markdown("**Actions:**")
                        
                        if is_active:
                            if st.button(f"🔴 Deactivate", key=f"deactivate_{rule_id}", width='stretch'):
                                if deactivate_rule(rule_id):
                                    st.success(f"Rule {rule_id} deactivated")
                                    st.rerun()
                                else:
                                    st.error("Failed to deactivate")
                        else:
                            st.info("Inactive")
                    
                    st.markdown("---")
    
    with tab_local:
        if not local_rules:
            st.info("No local rules for this session. Create temporary rules from the preview section!")
        else:
            st.markdown(f"**Total Local Rules:** {len(local_rules)} | **Active:** {sum(1 for r in local_rules if r.get('active', True))}")
            st.warning("⚠️ **Note:** Local rules disappear when you close the app. Save as Global if you want to keep them!")
            st.markdown("---")
            
            # Display local rules
            for idx, rule in enumerate(local_rules):
                rule_id = rule.get("id", f"lrule_{idx}")
                is_active = rule.get("active", True)
                timestamp = rule.get("timestamp", "Unknown")
                user_fb = rule.get("user_feedback", "No feedback")
                rule_text = rule.get("rule", "No rule extracted")
                projects = rule.get("applied_to_projects", [])
                
                status_emoji = "✅" if is_active else "🔴"
                
                with st.container():
                    col_rule1, col_rule2 = st.columns([4, 1])
                    
                    with col_rule1:
                        st.markdown(f"### {status_emoji} 📍 Local Rule: `{rule_id}`")
                        st.markdown(f"**Created:** {timestamp[:10] if len(timestamp) > 10 else timestamp}")
                        st.markdown(f"**Project:** {projects[0] if projects else 'N/A'}")
                        
                        with st.expander("📝 View Details", expanded=False):
                            st.markdown("**Original User Feedback:**")
                            st.text_area(f"Local_Feedback_{idx}", value=user_fb, height=80, disabled=True, key=f"lfb_{idx}", label_visibility="collapsed")
                            
                            st.markdown("**Extracted Rule:**")
                            st.text_area(f"Local_Rule_{idx}", value=rule_text, height=100, disabled=True, key=f"lrule_view_{idx}", label_visibility="collapsed")
                    
                    with col_rule2:
                        st.markdown("**Actions:**")
                        
                        if st.button(f"🗑️ Remove", key=f"remove_local_{idx}", width='stretch'):
                            local_rules.pop(idx)
                            st.session_state['local_feedback_rules'] = local_rules
                            st.success("Local rule removed")
                            st.rerun()
                        
                        if st.button(f"🌍 Make Global", key=f"promote_{idx}", width='stretch'):
                            # Save as global rule
                            rule_id_new = save_feedback_rule(user_fb, rule_text, projects[0] if projects else "Promoted", rule_type="global")
                            # Remove from local
                            local_rules.pop(idx)
                            st.session_state['local_feedback_rules'] = local_rules
                            st.success(f"Promoted to global rule: {rule_id_new}")
                            st.rerun()
                    
                    st.markdown("---")
    
    with tab_stats:
        st.markdown("### 📊 Rule Statistics")
        
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        
        with col_stat1:
            st.metric("Total Global Rules", stats.get("total", 0))
            st.metric("Active Global", stats.get("active", 0))
        
        with col_stat2:
            st.metric("Inactive Global", stats.get("inactive", 0))
            st.metric("Local Rules (Session)", len(local_rules))
        
        with col_stat3:
            total_projects = len(set(p for r in all_global_rules for p in r.get("applied_to_projects", [])))
            st.metric("Projects Using Rules", total_projects)
            st.metric("Total Rules (G+L)", stats.get("total", 0) + len(local_rules))
        
        st.markdown("---")
        st.markdown("**📈 Rule Breakdown:**")
        
        if all_global_rules or local_rules:
            rule_data = {
                "Type": ["Global Active", "Global Inactive", "Local Active"],
                "Count": [
                    stats.get("active", 0),
                    stats.get("inactive", 0),
                    len([r for r in local_rules if r.get("active", True)])
                ]
            }
            st.bar_chart(pd.DataFrame(rule_data).set_index("Type"))
        else:
            st.info("No rules created yet. Start generating process flows to build your rule base!")
        
        st.markdown("---")
        st.markdown("**💡 Tips:**")
        st.markdown("""
        - **Global rules** are best for organization-wide standards
        - **Local rules** are perfect for project-specific tweaks
        - Promote local rules to global if they're universally useful
        - Deactivate rules instead of deleting to preserve history
        - Review and clean up rules monthly for best performance
        """)

with st.expander("🚫 Exclusions Configuration", expanded=False):
    st.write("**Select Exclusions to Include:**")
    
    variable_exclusions = [
        "Server PC / server system.",
        "SCADA / PC for SCADA.",
        "Workstations.",
        "Cabling from server room to Falcon control panel.",
        "Mobile carts.",
        "Collection trolleys / collection trolleys below chutes.",
        "Collection bins.",
        "Pallets at chutes.",
        "Pallets / hand-held terminals for secondary sorting.",
        "Steel works.",
        "Steel works – if not specified.",
        "Mezzanine & staircase.",
        "Mezzanine & staircase not mentioned in BOM.",
        "Maintenance platform / lift required for maintenance activity.",
        "Safety fencing / safety fencing not shown in layout.",
        "HPT/BOPT/Forklift/Hydra/Scaffoldings required for installation.",
        "Stress free mats.",
        "Insulation mats.",
        "Fans at chutes & inducts.",
        "Lighting around chutes / inducts.",
        "Irregular's provision.",
        "UPS power (separate UPS supply).",
        "CE declaration of conformity.",
    ]
    
    selected_exclusions = []
    cols = st.columns(2)
    for idx, item in enumerate(variable_exclusions):
        col = cols[idx % 2]
        if col.checkbox(item, value=False, key=f"exclusion_{idx}"):
            selected_exclusions.append(item)

with st.expander("🔧 Key Components", expanded=False):
    default_components = [
        {"Items": "Belts", "Make": "Forbo / Derco / Habasit"},
        {"Items": "Rollers", "Make": "Falcon"},
        {"Items": "Cross Belt Carriers", "Make": "Falcon"},
        {"Items": "Linear Motors (LIM / LSM / Linear Induction)", "Make": "Falcon / SEW / FWD (as applicable)"},
        {"Items": "Feed Line Motors", "Make": "Falcon"},
        {"Items": "Volume / Barcode Scanners", "Make": "SICK / Cognex / Similar"},
        {"Items": "Weighing Scales", "Make": "Bizerba / Mettler Toledo / Equivalent"},
        {"Items": "Encoders", "Make": "SICK / Falcon"},
        {"Items": "Sensors", "Make": "SICK / Leuze / P&F"},
        {"Items": "PLC", "Make": "Siemens / Omron"},
        {"Items": "Control Panels", "Make": "Rittal / BCH"},
        {"Items": "VFDs", "Make": "Siemens / Lenze / AB / Omron"},
        {"Items": "Cables", "Make": "LAPP / Equivalent"},
        {"Items": "Switch Gear", "Make": "Schneider / Equivalent"},
        {"Items": "Bearings", "Make": "NTN / SKF / Equivalent"},
        {"Items": "Power Transmission Systems", "Make": "Vahle"},
        {"Items": "HMIs", "Make": "Siemens / Omron"},
        {"Items": "MDR", "Make": "Pulse / Itoh Denki"},
        {"Items": "Data Transmission System", "Make": "Siemens"},
    ]
    
    if "key_components_df" not in st.session_state:
        st.session_state["key_components_df"] = pd.DataFrame(default_components)
    
    key_components_edited = st.data_editor(
        st.session_state["key_components_df"],
        num_rows="dynamic",
        width='stretch',
        key="key_editor"
    )

st.divider()

# ==================== DOCUMENT GENERATION FUNCTIONS ====================

def build_cover_letter_section(doc, letter_text):
    """Build cover letter (page 1) - NO HEADER for this section"""
    HEADER_PREFIXES = (
        "Kind Attention",
        "Mr.",
        "Ms.",
        "M/s",
        "Offer Ref:",
        "Subject –",
        "Subject -",
        "Date:",
        "Location –",
    )
    
    lines = [l.rstrip() for l in letter_text.splitlines() if l.strip() != ""]
    for idx, line in enumerate(lines):
        # Check if line starts with header prefixes or contains "Dear" or ends with signature (Best Regards)
        is_header = any(line.startswith(pfx) for pfx in HEADER_PREFIXES) or "Dear " in line
        is_signature = "Best Regards" in line or idx >= len(lines) - 2

        if is_header:
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True
        elif is_signature:
            # Signature and name/title at end should be bold
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.bold = True
        else:
            if "**" in line:
                add_markdown_line(doc, line)
            else:
                p = doc.add_paragraph(line)
                apply_normal_style(p)


def build_front_page_section(doc, project_title, offer_ref, contact_name, contact_email, contact_phone, layout_png_path):
    """Build front page (page 2) - NO HEADER for this section"""
    doc.add_page_break()

    # Top box: "Response to RFP for"
    add_boxed_text(doc, "Response to RFP for", font_size=18, bold=True)

    # Project name box
    if project_title:
        add_boxed_text(doc, project_title, font_size=16, bold=True)

    # Proposal reference box
    if offer_ref:
        add_boxed_text(doc, f"Proposal Reference: {offer_ref}", font_size=14, bold=True)

    # Some vertical spacing
    doc.add_paragraph("")

    # Layout image - use the same image as in Proposed System Description
    # Cropped to 5.0 inches to fit everything on single page
    if layout_png_path and os.path.exists(layout_png_path):
        # Crop image before inserting
        try:
            from PIL import Image
            img = Image.open(layout_png_path)
            
            # Crop 10% from each side to remove whitespace
            width, height = img.size
            left = width * 0.1
            top = height * 0.1
            right = width * 0.9
            bottom = height * 0.9
            
            img_cropped = img.crop((left, top, right, bottom))
            
            # Save to temporary buffer
            img_buffer = io.BytesIO()
            img_cropped.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            p = doc.add_paragraph()
            run = p.add_run()
            run.add_picture(img_buffer, width=Inches(5.0))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            # Fallback to original image if cropping fails
            p = doc.add_paragraph()
            run = p.add_run()
            run.add_picture(layout_png_path, width=Inches(5.0))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p = doc.add_paragraph()
        run = p.add_run("Layout image will be provided.")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Contact box at bottom
    contact_lines = [
        "Falcon Autotech Private Limited",
        "Plot No. 87, Sector Ecotech-1, Extention-1, Greater Noida, Uttar Pradesh 201308.",
        "",
        f"Contact – {contact_name}",
        "Assistant Manager",
        f"Mob - {contact_phone}",
        contact_email,
    ]
    contact_text = "\n".join(contact_lines)
    table = add_boxed_text(doc, contact_text, font_size=11, bold=False)
    # make contact text paragraphs centered
    cell = table.rows[0].cells[0]
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def build_glossary_section(doc):
    """Build glossary/table of contents section with automatic TOC"""
    doc.add_page_break()
    
    p = doc.add_heading("Table of Contents", level=1)
    for run in p.runs:
        run.font.name = 'Calibri'
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.underline = True
        run.font.color.rgb = HEADING_COLOR
    
    # Add automatic TOC field
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar3)
    
    # Add instruction text for users
    doc.add_paragraph("")
    p = doc.add_paragraph("Note: Right-click on the table of contents and select 'Update Field' to refresh page numbers.")
    apply_normal_style(p)
    p.runs[0].italic = True

def build_glossary(doc, detected_terms):
    """Build Glossary section with table of detected terms"""
    
    # Add heading using standard formatting
    add_numbered_heading(doc, "Glossary", level=1, counter=1)
    
    if not detected_terms:
        p = doc.add_paragraph("No glossary terms detected in this document.")
        apply_normal_style(p)
        return
    
    # Create table with 3 columns: S. No., Term, Description
    table = doc.add_table(rows=1, cols=3)
    apply_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Header row
    hdr_cells = table.rows[0].cells
    headers = ["S. No.", "Term", "Description"]
    for i, text in enumerate(headers):
        p = hdr_cells[i].paragraphs[0]
        p.text = ""
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    # Data rows
    for idx, (term, desc) in enumerate(detected_terms, start=1):
        row_cells = table.add_row().cells
        
        # S. No.
        p0 = row_cells[0].paragraphs[0]
        p0.text = ""
        r0 = p0.add_run(str(idx))
        r0.font.name = 'Calibri'
        r0.font.size = Pt(11)
        r0.font.bold = True
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # Term
        p1 = row_cells[1].paragraphs[0]
        p1.text = ""
        r1 = p1.add_run(term)
        r1.font.name = 'Calibri'
        r1.font.size = Pt(11)
        r1.font.bold = True
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # Description
        p2 = row_cells[2].paragraphs[0]
        p2.text = ""
        r2 = p2.add_run(desc)
        r2.font.name = 'Calibri'
        r2.font.size = Pt(11)
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        row_cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(0.8)
    table.columns[1].width = Inches(1.5)
    table.columns[2].width = Inches(4.2)

    # Ensure next section starts on a fresh page so Glossary remains isolated
    doc.add_page_break()
def build_executive_summary_section(doc, exec_summary_text, counter):
    """Build Executive Summary section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Executive Summary", counter=counter)
    
    lines = exec_summary_text.strip().splitlines()
    for line in lines:
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph("")
            continue
        
        # Check if it's a bullet point line
        if stripped.startswith("•") or stripped.startswith("-"):
            bullet_text = stripped.lstrip("•- ").strip()
            p = doc.add_paragraph(style='List Bullet')
            if "**" in bullet_text:
                parts = bullet_text.split("**")
                for i, part in enumerate(parts):
                    if not part:
                        continue
                    run = p.add_run(part)
                    if i % 2 == 1:
                        run.bold = True
                    run.font.name = "Calibri"
                    run.font.size = Pt(11)
                    run.italic = True
            else:
                run = p.add_run(bullet_text)
                run.font.name = "Calibri"
                run.font.size = Pt(11)
                run.italic = True
        else:
            # Regular paragraph with possible **bold**
            if "**" in stripped:
                add_markdown_line(doc, stripped)
            else:
                p = doc.add_paragraph(stripped)
                apply_normal_style(p)


def build_company_profile_section(doc, counter):
    """Build Company Profile section with static images"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Company Profile", counter=counter)

    top_text = (
        "Falcon Autotech (Falcon) is a global intralogistics automation solutions company. "
        "With over 10 years of experience, Falcon has worked with some of the most innovative "
        "brands in E-Commerce, CEP, Fashion, Food/FMCG, Auto and Pharmaceutical Industries. "
        "With our proprietary software and robust hardware integration capabilities, Falcon designs, "
        "manufactures, supplies, implements, and maintains world-class warehouse automation systems globally. "
        "Falcon's strong research and development team and the continuous focus on innovation reflect our strong "
        "solution line around Sortation, Robotics, Conveying, Vision Systems and IOT. "
        "Falcon has done over 1,800 installations across 15 countries on four continents."
    )
    p = doc.add_paragraph(top_text)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "1.png"), width_in=4)

    bottom_text = (
        "Falcon Autotech is currently among the top 15 intralogistics automation companies; "
        "our vision is to become a top 10 intralogistics automation company in our focused product lines."
    )
    p = doc.add_paragraph(bottom_text)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "2.png"), width_in=4)

    doc.add_page_break()

    # Page 2
    top_text2 = (
        "The team started out in 2004 solving special purpose automation problems for clients and later "
        "established Falcon Autotech in 2012 with a strong focus on building a standard technology stack spanning "
        "across hardware, firmware, and software to tackle larger supply chain problems around warehouse "
        "automation and material handling. "
        "Over the decade, Falcon has made rapid strides and has carved out a niche in some of the world's most "
        "cutting-edge technologies: Sortation, Robotics, Conveying, Vision Systems and IOT."
    )
    p = doc.add_paragraph(top_text2)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "3.png"), width_in=6)

    bottom_text2 = (
        "As a leading player in the intralogistics automation space, Falcon continuously strives to improve the "
        "operational efficiencies and accuracies for its clients through its domain knowledge and experience, in "
        "addition to its wide range of products and solutions. In order to live up to the high expectations set "
        "forth by our clients, the team at Falcon realizes the importance of taking up selective applications in "
        "focused industries and delivering world-class projects in return."
    )
    p = doc.add_paragraph(bottom_text2)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "4.png"), width_in=6)

    # Page 3
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "5.png"), width_in=6)

    bottom_text3 = (
        "Falcon Autotech has successfully delivered warehouse automation solutions based on smart and innovative "
        "combinations of the above product lines for effective materials handling, sortation and movement. "
        "The process is controlled in real-time by our in-house WCS applications. These solutions considerably "
        "reduce the need for manual operations, improve working conditions and ensure the highest accuracy of the "
        "entire process up to final delivery to the recipient.\n\n"
        "Over the last 10 years, Falcon has worked with some of the most innovative brands worldwide and has "
        "established long-standing partnerships. These brands are testimony to our strong focus on delivering "
        "superior customer satisfaction and offering end-to-end intralogistics solutions."
    )
    p = doc.add_paragraph(bottom_text3)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "6.png"), width_in=6)

    doc.add_page_break()

    # Page 4
    bottom_text4 = (
        "With over 1,800 installations, Falcon's systems are used all over the globe. Falcon has a highly "
        "motivated team of 600+ employees supported by over 15 global partners who help us design, manufacture, "
        "deliver and maintain automation solutions worldwide."
    )
    p = doc.add_paragraph(bottom_text4)
    apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "7.png"), width_in=6)

    add_numbered_subheading(doc, "Customer Engagement Model", f"{counter}.1")
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "8.png"), width_in=7)

    doc.add_page_break()

    # Page 5
    add_numbered_subheading(doc, "Falcon's Experience and Achievements in Sortation Space Globally", f"{counter}.2")

    bullet_points = [
        "Ranked among Top 10 Sortation System Suppliers globally.",
        "Currently possess one of the world's largest portfolios in sortation technologies (7 in-house technologies).",
        "Total installed capacity of 10 million shipments per day worldwide.",
        "Only company to be able to offer a fully integrated AMS.",
    ]

    for point in bullet_points:
        p = doc.add_paragraph(point, style='List Bullet')
        apply_normal_style(p)

    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "9.png"), width_in=5)
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "10.png"), width_in=5)
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "11.png"), width_in=5)
    add_centered_image(doc, os.path.join(STATIC_ABOUT_DIR, "12.png"), width_in=5)


def build_reference_projects_section(doc, counter):
    """Build Reference Projects section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Reference Projects", counter=counter)

    intro = (
        "Falcon has a strong legacy in Warehousing Automation solutions and references-"
    )
    p = doc.add_paragraph(intro)
    apply_normal_style(p)

    bullets_intro = [
        "Expertise in Shipment Sortation, Piece Picking and Handling, Case Picking and Handling.",
        "Lifecycle services (maintenance, spares supply chain, support).",
        "Full in-house expertise (Hardware/Software).",
        "Turn-key tailored solutions.",
        "The references list presented below focuses on Sortation Solution –",
    ]
    for text in bullets_intro:
        p = doc.add_paragraph(text, style='List Bullet')
        apply_normal_style(p)

    # Project 1
    add_numbered_subheading(doc, "Project 1- (CEP Client, India)", f"{counter}.1")

    p = doc.add_paragraph(
        "The system is equipped with two fully automated and interconnected sub-systems. "
        "Sub-System 1 is designed for handling large B2B boxes and E-commerce shipment bags while "
        "Sub-System 2 is designed to handle small E-commerce packages."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec1 = [
        "48,000 PPH (Double Deck CBS – Shipment Sorter).",
        "17,000 PPH (Double Deck CBS – Bag Sorter).",
        "Building Size: 700,000 Sq. Ft.",
    ]
    for t in spec1:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm1 = [
        "2 Sets of Double Decker CBS Sorters.",
        "Mezzanine Structures.",
        "Automated Singulators.",
        "Fully Automatic Inductions.",
        "Semi-Automatic Inductions.",
        "Telescopic Belt Conveyors.",
        "PVC Belt Conveyors.",
        "Modular Belt Conveyors.",
        "Spiral Chutes with Braking Rollers.",
        "5-Sided Scanning Tunnels.",
        "High Speed Weighing Conveyors.",
        "Direct Bagging Chutes.",
        "Put to Light Chutes.",
        "Volume Distribution Systems.",
        "High Availability Server Systems.",
        "WCS.",
    ]
    for t in ktm1:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Pictures –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj1.PNG")

    doc.add_page_break()

    # Project 2
    add_numbered_subheading(doc, "Project 2- (Client – E-Commerce, India)", f"{counter}.2")

    p = doc.add_paragraph()
    run = p.add_run("Use Case – Destination sorting of packed shipments.")
    run.bold = True
    apply_normal_style(p)

    p = doc.add_paragraph(
        "In 2019, the client was looking for a potential automation partner for design and development of a "
        "new automated sortation system for B2C shipments. The system needed to provide maximum uptime with "
        "reduced dependency on skilled manpower and better space optimization. "
        "\nThe customer chose Falcon Autotech based on its unique design that addressed these pain points, "
        "its capability for seamless WMS integration, and its life cycle support services."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec2 = [
        "Throughput: 27,600 PPH.",
        "End Destinations: 410 Direct Outputs.",
        "Building Size: 200,000 Sq. Ft.",
    ]
    for t in spec2:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm2 = [
        "Bulk Infeed Conveyors.",
        "ARB based Volume Distribution System.",
        "Integrated Presort System.",
        "Irregular Ejection System.",
        "Automatic Induct Lines.",
        "Automatic Barcode Scanner with Image Capture.",
        "Automatic Weight & Volume Measurement System.",
        "Linear Cross Belt Sorter.",
        "Smart Sliding Chutes for Direct Bagging and Cage Sorting.",
        "Bag Take-out System.",
        "WCS Software System.",
    ]
    for t in ktm2:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Pictures –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj2.PNG")

    doc.add_page_break()

    # Project 3
    add_numbered_subheading(doc, "Project 3- (Client – E-Commerce, India)", f"{counter}.3")

    p = doc.add_paragraph()
    run = p.add_run("Use Case – Destination sorting of packed shipments.")
    run.bold = True
    apply_normal_style(p)

    p = doc.add_paragraph(
        "The customer chose Falcon Autotech based on its unique design, its ability to integrate seamlessly "
        "with the WMS, and its strong life cycle support services."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec3 = [
        "Throughput: 24,000 PPH.",
        "End Destinations: 40 Collection Type Chutes.",
    ]
    for t in spec3:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm3 = [
        "Bulk Infeed Conveyors.",
        "ARB based Volume Distribution System.",
        "Irregular Ejection System.",
        "Automatic Induct Lines.",
        "Automatic Barcode Scanner with Image Capture.",
        "Automatic Weight & Volume Measurement System.",
        "Linear Cross Belt Sorter.",
        "Smart Collection Type Chutes.",
        "Bag Take-out System.",
        "WCS Software System.",
    ]
    for t in ktm3:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Pictures –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj3.PNG")

    doc.add_page_break()

    # Project 4
    add_numbered_subheading(doc, "Project 4- (CEP Client, UK)", f"{counter}.4")

    p = doc.add_paragraph(
        "This solution is designed to handle a volume of 7,200 shipments per hour. "
        "The system is equipped with three infeed conveyors integrated with an automatic label applicator "
        "before shipments enter the sortation system. Shipments are sorted using Falcon's Loop Cross Belt Sorter "
        "equipped with automatic barcode scanning, dimensioning, weighing, and image capture capabilities. "
        "The sorter is installed on the mezzanine floor and sorts directly to 58 end destinations."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec4 = [
        "Throughput: 7,200 PPH.",
        "End Destinations: 58 Nos.",
    ]
    for t in spec4:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm4 = [
        "Powered Belt Conveyors.",
        "Automatic Induct Lines.",
        "Automatic Barcode Scanner with Image Capture.",
        "Automatic Weight & Volume Measurement System.",
        "Loop Cross Belt Sorter.",
        "WCS Software System.",
    ]
    for t in ktm4:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Picture –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj4.PNG")

    doc.add_page_break()

    # Project 5
    add_numbered_subheading(doc, "Project 5- (CEP Client, Sydney)", f"{counter}.5")

    p = doc.add_paragraph(
        "This solution is designed for handling a throughput of 16,000 shipments per hour with the help of "
        "Falcon's Loop Cross Belt Sorter. The system consists of two feeding zones with a total of ten feedlines. "
        "Sorter design enables van drivers to directly drop shipments at the dock doors. It has a total of 369 end "
        "destinations achieved through a combination of direct drops and PTLs. The system is integrated with "
        "five-side automatic barcode scanning, weight and volume measurement, and automatic detection of "
        "oversize and overweight shipments."
    )
    apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Solution Specifications –")
    run.bold = True
    apply_normal_style(p)
    
    spec5 = [
        "Throughput: 16,000 PPH.",
        "End Destinations: 369 Nos.",
    ]
    for t in spec5:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Key Technology Modules –")
    run.bold = True
    apply_normal_style(p)
    
    ktm5 = [
        "Powered Belt Conveyors.",
        "2 Induct Zones.",
        "5-side Automatic Barcode Scanner.",
        "Automatic Weight & Volume Measurement System.",
        "Automatic Detection of Oversize Shipments.",
        "Loop Cross Belt Sorter.",
        "WCS Software System.",
    ]
    for t in ktm5:
        p = doc.add_paragraph(t, style='List Bullet')
        apply_normal_style(p)

    p = doc.add_paragraph()
    run = p.add_run("Site Picture –")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\proj5.PNG")


def build_handled_spectrum_section(doc, counter, project_name, client_name):
    """Build Handled Shipment Spectrum section"""
    doc.add_page_break()  # Start on new page
    tpl = choose_sorter_template(project_name)

    item_singular = tpl.item_singular.lower()
    item_cap = item_singular.capitalize()
    item_plural = item_singular + "s"
    item_plural_cap = item_plural.capitalize()

    add_numbered_heading(doc, "Handled Shipment Spectrum", counter=counter)

    intro_1 = (
        f"{client_name} operates in a business where handling a wide spectrum of {item_plural} is critical. "
        f"Falcon has carefully analyzed the provided {item_singular} spectrum and tailored the solution to your needs."
    )
    intro_2 = (
        f"Falcon proposes to use its \"{tpl.config_name}\" to deliver maximum operational benefits to {client_name}, "
        f"ensuring reliable handling of all relevant sizes and weights for your business."
    )
    p = doc.add_paragraph(intro_1)
    apply_normal_style(p)
    p = doc.add_paragraph(intro_2)
    apply_normal_style(p)

    # Subsection 1
    add_numbered_subheading(doc, tpl.subheading_51, f"{counter}.1")

    p = doc.add_paragraph(
        f"Falcon's {tpl.config_name} has a capability to handle the below mentioned "
        f"{item_plural} sizes and weight."
    )
    apply_normal_style(p)

    # Table
    table = doc.add_table(rows=1 + len(tpl.spec_table), cols=3)
    apply_table_style(table)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Specification"
    hdr_cells[1].text = "Unit"
    hdr_cells[2].text = "Value"
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri (Body)'
                run.font.size = Pt(11)

    # Add data rows only (skip header row creation, we already have it)
    row_idx = 1
    for spec, data in tpl.spec_table.items():
        if str(data["value"]).strip() == "":
            continue  # Skip empty rows
        if row_idx >= len(table.rows):
            row_cells = table.add_row().cells
        else:
            row_cells = table.rows[row_idx].cells
        row_idx += 1
        row_cells[0].text = spec
        row_cells[1].text = data["unit"]
        row_cells[2].text = data["value"]
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                apply_normal_style(paragraph)

    # Subsection 2
    add_numbered_subheading(
        doc,
        f"{item_plural_cap} to be loaded on Sorter shall have the following characteristics:",
        f"{counter}.2"
    )

    bullets_52 = [
        f"Centre of Gravity of item must not move during conveyance or sorting.",
        f"Item must not have magnetic content, otherwise behavior of {item_singular} cannot be guaranteed.",
        "Liquid or fragile material, to avoid breaking, spillage or leakage, such as wine bottles, "
        "metal cans of paint are designated as non-conveyable items.",
        f"{item_plural_cap} shall be perfectly and safely packaged: protrusion or open surfaces are not allowed.",
        "Plastic ropes shall be perfectly adherent to the surface of the package.",
        "All items with the risk of being damaged during the transport on an automatic sorting system "
        "or damaging the sorting system; they must be robust enough to avoid disintegration of container "
        "material and loss of contents in the sorting process.",
        "Item packaging shall have enough grip to be handled on the belts during the acceleration and "
        "referencing phases.",
        "Items shall not have slippery surfaces and must be able to withstand acceleration of the items "
        "on the belt during the start-stop phases (accelerations up to 0.5 g shall be assured without any "
        "sliding or tumbling of the items on the belt conveyor).",
        f"The {item_plural} must have at least one flat and regular surface providing enough stability during "
        "conveyance.",
        "All shapes are permitted except spherical, cylindrical, or alike unstable items & shapes.",
        "All usual packaging materials are permitted (including paper, carton, plastics, plastic foil, rope, "
        "tape, textile, and wood).",
    ]
    
    for text in bullets_52:
        p = doc.add_paragraph(text, style="List Number 2")
        apply_normal_style(p)

    # Subsection 3
    add_numbered_subheading(doc, f"{item_plural_cap} not loadable on the sorter", f"{counter}.3")

    bullets_53 = [
        "Unstable items with a risk to roll or tumble on the sorting system, such as spherical or cylindrical items.",
        "Items that have a spherical or cylindrical shape.",
        "Items that are packed in material that can damage the conveyors or the sorter.",
        "Items that have sharp points (e.g., Nails) or sharp edges, that can damage the conveyors or the sorter.",
        f"Fragile {item_plural} with contents not sufficiently secured.",
        "Items that have been classified as dangerous are designated.",
        "Wet items are designated.",
        "Items with anti-slip treatment.",
        "Items with protruding parts.",
        "Items with sharp edges.",
        "Inadequately packed items that could be damaged during automatic transportation.",
        "Electrostatically loaded items.",
        "Loose parts on loads and load carriers, such as adhesive tape, stickers, slips of paper, straps, "
        "wrap foil etc. are designated as non-conveyable items.",
    ]
    
    for text in bullets_53:
        p = doc.add_paragraph(text, style="List Number 2")
        apply_normal_style(p)


def build_description_of_components_section(doc, counter, sorter_spec: dict | None):
    """
    Build Description of Components of Equipment section with fixed content and 
    Technical Specification of Sorter from Loop CBS Excel.
    """
    # Main heading
    add_numbered_heading(doc, "Description of Components of Equipment", counter=counter)
    doc.add_paragraph()

    # Elements of the sorting system
    add_numbered_subheading(doc, "Elements of the sorting system", f"{counter}.1")
    
    p = doc.add_paragraph()
    p.add_run("1. Cross Belt Sorter\n").font.name = 'Calibri'
    p.add_run("2. Scanning & Sensing on Sorter\n").font.name = 'Calibri'
    p.add_run("3. Conveyor System\n").font.name = 'Calibri'
    p.add_run("4. Steel Works – Mezzanine & Staircases").font.name = 'Calibri'
    apply_normal_style(p)
    doc.add_paragraph()

    # Cross Belt Sorter
    add_numbered_subheading(doc, "Cross Belt Sorter", f"{counter}.2")
    
    add_paragraph(doc, "Cross Belt Sorter is capable of sorting extremely high volume of versatile products in a gentle manner.")
    add_paragraph(doc, "Falcon's Cross belt sorter is powered by high efficiency linear motors and is based on 100% non-touch actuation technology leading to high throughput capabilities with extremely low noise levels.")
    add_paragraph(doc, "Falcon's Cross belt sorter is modular in design. It can be easily extended as per future requirements.")
    
    add_centered_image(doc, COMPONENT_IMAGES["cross_belt_sorter"], width_in=4.0)

    # CBS Carrier
    add_numbered_subheading(doc, "CBS Carrier", f"{counter}.3")
    
    add_paragraph(doc, "Falcon Autotech's CBS offers one of the highest belt width to carrier pitch ratios in the market today.")
    add_paragraph(doc, "This additional belt width makes the system capable of handling larger product sizes without compromising on throughput. It also reduces the dead area between carrier belts, significantly reducing the number of in-betweeners and non-sortable parcel recirculation.")
    
    add_centered_image(doc, COMPONENT_IMAGES["cbs_carrier"], width_in=4.0)
    doc.add_paragraph()

    # Servo Roller (Level 3 subheading)
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.1 Servo Roller")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "High powered DC drive servo rollers are used to actuate the carrier belts, thereby eliminating the need for complicated drive transfer mechanisms and simplifying system installation and maintenance.")
    add_centered_image(doc, COMPONENT_IMAGES["servo_roller"], width_in=3.5)
    doc.add_paragraph()

    # Chassis
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.2 Chassis")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Falcon Autotech's cross belt carrier chassis is made up of lightweight aluminium, which makes it light yet sturdy. This reduced weight leads to substantial power savings over a considerable period of usage.")
    add_centered_image(doc, COMPONENT_IMAGES["chassis"], width_in=3.5)
    doc.add_paragraph()

    # Carrier Wheels
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.3 Carrier Wheels")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Carrier wheels are thoroughly tested and proven for a long life cycle.")
    add_centered_image(doc, COMPONENT_IMAGES["wheel"], width_in=3.5)
    doc.add_paragraph()

    # Friction Wheel Drive
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.4 Friction Wheel Drive")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Falcon can provide the indigenously developed Friction Wheel Drive (FWD) to drive the cross belt loop. This driving mechanism operates on the principle of friction. The unit comprises two independent motor-driven wheels that spin in opposite directions, synchronised.")
    add_paragraph(doc, "FWDs are highly energy-efficient drives that promote sustainability compared to traditional linear induction or synchronous motor drives.")
    add_centered_image(doc, COMPONENT_IMAGES["friction_wheel"], width_in=4.0)

    # Non-contact based linear motor drive
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.5 Non-contact based linear motor drive")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Non-contact based linear induction motors can be configured at variable speeds depending upon operational requirements, providing maximum flexibility.")
    add_centered_image(doc, COMPONENT_IMAGES["linear"], width_in=4.0)
    add_paragraph(doc, "The customer can choose the preferred drive system based on performance and energy-efficiency requirements.")
    doc.add_paragraph()

    # Power Transmission
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.6 Power Transmission")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Power transmission to carriers is provided over sliding contacts that require low maintenance and offer high levels of reliability.")
    add_centered_image(doc, COMPONENT_IMAGES["power"], width_in=4.0)
    doc.add_paragraph()

    # Data Transmission
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.7 Data Transmission")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "The R-Coax cable is used for data distribution in the sorter. This leaky wave cable runs throughout the sorter length, transmitting data continuously. An antenna mounted on the super master carriage receives the signal from this cable while on the move, wirelessly.")
    add_centered_image(doc, COMPONENT_IMAGES["rcoax"], width_in=4.0)

    # Carriers positioning system
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.8 Carriers positioning system")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "The positioning system determines the exact location of each carrier in the loop at any given point in time.")
    add_paragraph(doc, "A plastic tape strip of barcodes (QR codes) runs along the sorter loop, which is continuously scanned by scanners placed on a master carrier to track and control the position of every carrier.")
    add_centered_image(doc, COMPONENT_IMAGES["carrier_position"], width_in=4.0)

    # Technical Specification of Sorter
    add_numbered_subheading(doc, "Technical Specification of Sorter", f"{counter}.4")

    if sorter_spec is None:
        add_paragraph(doc, "Technical specification of the sorter will be finalised based on project-specific configuration.")
    else:
        # Safely get values (fallbacks)
        def gv(key: str, default: str = "-") -> str:
            val = sorter_spec.get(key)
            if val is None:
                return default
            s = str(val).strip()
            return s if s else default

        rows_data = [
            ("Sorter Carrier Type", gv("sorter_carrier_type", "Loop CBS")),
            ("Sorter Speed (m/s)", gv("sorter_speed_mps")),
            ("Sorter Loop Length (m)", gv("sorter_loop_length_m")),
            ("Sorter Height (mm)", gv("sorter_height_mm")),
            ("Sorter Actuation Technology", gv("actuation_technology", "Electric")),
            ("Carrier Pitch (mm)", gv("carrier_pitch_mm")),
            ("Number of Carriers", gv("number_of_carriers")),
            ("Motor / Drive Type", gv("motor_drive_type")),
            ("Power Consumption*", gv("power_consumption")),
        ]

        table = doc.add_table(rows=1, cols=2)
        apply_table_style(table)
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Technical Parameter"
        hdr_cells[1].text = "Value"

        for param, value in rows_data:
            row_cells = table.add_row().cells
            row_cells[0].text = param
            row_cells[1].text = value

        doc.add_paragraph()


def add_paragraph(doc: Document, text: str):
    """Add normal body paragraph with Calibri 11pt."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    return p


def build_description_of_components_section(doc, counter, sorter_spec: dict | None):
    """
    Build Description of Components of Equipment section with fixed content and 
    Technical Specification of Sorter from Loop CBS Excel.
    """
    doc.add_page_break()  # Start on new page
    
    # Main heading
    add_numbered_heading(doc, "Description of Components of Equipment", counter=counter)
    doc.add_paragraph()

    # Elements of the sorting system
    add_numbered_subheading(doc, "Elements of the sorting system", f"{counter}.1")
    
    p = doc.add_paragraph()
    p.add_run("1. Cross Belt Sorter\\n").font.name = 'Calibri'
    p.add_run("2. Scanning & Sensing on Sorter\\n").font.name = 'Calibri'
    p.add_run("3. Conveyor System\\n").font.name = 'Calibri'
    p.add_run("4. Steel Works – Mezzanine & Staircases").font.name = 'Calibri'
    apply_normal_style(p)
    doc.add_paragraph()

    # Cross Belt Sorter
    add_numbered_subheading(doc, "Cross Belt Sorter", f"{counter}.2")
    
    add_paragraph(doc, "Cross Belt Sorter is capable of sorting extremely high volume of versatile products in a gentle manner.")
    add_paragraph(doc, "Falcon's Cross belt sorter is powered by high efficiency linear motors and is based on 100% non-touch actuation technology leading to high throughput capabilities with extremely low noise levels.")
    add_paragraph(doc, "Falcon's Cross belt sorter is modular in design. It can be easily extended as per future requirements.")
    
    add_centered_image(doc, COMPONENT_IMAGES["cross_belt_sorter"], width_in=4.0)

    # CBS Carrier
    add_numbered_subheading(doc, "CBS Carrier", f"{counter}.3")
    
    add_paragraph(doc, "Falcon Autotech's CBS offers one of the highest belt width to carrier pitch ratios in the market today.")
    add_paragraph(doc, "This additional belt width makes the system capable of handling larger product sizes without compromising on throughput. It also reduces the dead area between carrier belts, significantly reducing the number of in-betweeners and non-sortable parcel recirculation.")
    
    add_centered_image(doc, COMPONENT_IMAGES["cbs_carrier"], width_in=4.0)
    doc.add_paragraph()

    # Servo Roller (Level 3 subheading)
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.1 Servo Roller")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "High powered DC drive servo rollers are used to actuate the carrier belts, thereby eliminating the need for complicated drive transfer mechanisms and simplifying system installation and maintenance.")
    add_centered_image(doc, COMPONENT_IMAGES["servo_roller"], width_in=3.5)
    doc.add_paragraph()

    # Chassis
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.2 Chassis")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Falcon Autotech's cross belt carrier chassis is made up of lightweight aluminium, which makes it light yet sturdy. This reduced weight leads to substantial power savings over a considerable period of usage.")
    add_centered_image(doc, COMPONENT_IMAGES["chassis"], width_in=3.5)
    doc.add_paragraph()

    # Carrier Wheels
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.3 Carrier Wheels")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Carrier wheels are thoroughly tested and proven for a long life cycle.")
    add_centered_image(doc, COMPONENT_IMAGES["wheel"], width_in=3.5)
    doc.add_paragraph()

    # Friction Wheel Drive
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.4 Friction Wheel Drive")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Falcon can provide the indigenously developed Friction Wheel Drive (FWD) to drive the cross belt loop. This driving mechanism operates on the principle of friction. The unit comprises two independent motor-driven wheels that spin in opposite directions, synchronised.")
    add_paragraph(doc, "FWDs are highly energy-efficient drives that promote sustainability compared to traditional linear induction or synchronous motor drives.")
    add_centered_image(doc, COMPONENT_IMAGES["friction_wheel"], width_in=4.0)

    # Non-contact based linear motor drive
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.5 Non-contact based linear motor drive")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Non-contact based linear induction motors can be configured at variable speeds depending upon operational requirements, providing maximum flexibility.")
    add_centered_image(doc, COMPONENT_IMAGES["linear"], width_in=4.0)
    add_paragraph(doc, "The customer can choose the preferred drive system based on performance and energy-efficiency requirements.")
    doc.add_paragraph()

    # Power Transmission
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.6 Power Transmission")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "Power transmission to carriers is provided over sliding contacts that require low maintenance and offer high levels of reliability.")
    add_centered_image(doc, COMPONENT_IMAGES["power"], width_in=4.0)
    doc.add_paragraph()

    # Data Transmission
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.7 Data Transmission")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "The R-Coax cable is used for data distribution in the sorter. This leaky wave cable runs throughout the sorter length, transmitting data continuously. An antenna mounted on the super master carriage receives the signal from this cable while on the move, wirelessly.")
    add_centered_image(doc, COMPONENT_IMAGES["rcoax"], width_in=4.0)

    # Carriers positioning system
    p = doc.add_paragraph()
    run = p.add_run(f"{counter}.3.8 Carriers positioning system")
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = HEADING_COLOR
    
    add_paragraph(doc, "The positioning system determines the exact location of each carrier in the loop at any given point in time.")
    add_paragraph(doc, "A plastic tape strip of barcodes (QR codes) runs along the sorter loop, which is continuously scanned by scanners placed on a master carrier to track and control the position of every carrier.")
    add_centered_image(doc, COMPONENT_IMAGES["carrier_position"], width_in=4.0)

    # Technical Specification of Sorter
    add_numbered_subheading(doc, "Technical Specification of Sorter", f"{counter}.4")

    if sorter_spec is None:
        add_paragraph(doc, "Technical specification of the sorter will be finalised based on project-specific configuration.")
    else:
        # Safely get values (fallbacks)
        def gv(key: str, default: str = "-") -> str:
            val = sorter_spec.get(key)
            if val is None:
                return default
            s = str(val).strip()
            return s if s else default

        rows_data = [
            ("Sorter Carrier Type", gv("sorter_carrier_type", "Loop CBS")),
            ("Sorter Speed (m/s)", gv("sorter_speed_mps")),
            ("Sorter Loop Length (m)", gv("sorter_loop_length_m")),
            ("Sorter Height (mm)", gv("sorter_height_mm")),
            ("Sorter Actuation Technology", gv("actuation_technology", "Electric")),
            ("Carrier Pitch (mm)", gv("carrier_pitch_mm")),
            ("Number of Carriers", gv("number_of_carriers")),
            ("Motor / Drive Type", gv("motor_drive_type")),
            ("Power Consumption*", gv("power_consumption")),
        ]

        table = doc.add_table(rows=1, cols=2)
        apply_table_style(table)
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Technical Parameter"
        hdr_cells[1].text = "Value"

        for param, value in rows_data:
            row_cells = table.add_row().cells
            row_cells[0].text = param
            row_cells[1].text = value

        doc.add_paragraph()


def build_capacity_calculations_section(doc, counter, client_name, project_name, capacity_excel):
    """Add Sorter System Capacity section using capacity_calculations.py logic"""
    if capacity_excel:
        excel_bytes = capacity_excel.read()
        prompt = build_capacity_prompt_from_excel(excel_bytes, client_name, project_name)
        cap_data = call_groq_for_capacity(prompt)
        doc.add_page_break()
        add_capacity_section_to_doc(doc, client_name, project_name, cap_data, counter)


def build_electrical_section(doc, counter):
    """Build Electrical System section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Electrical System", counter=counter)
    
    p = doc.add_paragraph(
        "Main power supply will supply Falcon's PDP (Power Distribution Panels) electrical cabinets. "
        "PDP cabinets supply the entire system via secondary cabinets:"
    )
    apply_normal_style(p)
    
    for item in ["Main Control Cabinet", "Induct Control Panels", "Remote Cabinets for Sorter I/O", "Scanner Control cabinets"]:
        p = doc.add_paragraph(item, style='List Bullet')
        apply_normal_style(p)
    
    add_numbered_subheading(doc, "Reference Picture of Power Distribution Panel", f"{counter}.1")
    add_centered_image(doc, "FIXED_IMAGE/elec1.PNG", width_in=4)
    
    add_numbered_subheading(doc, "Main Control Panel (Reference)", f"{counter}.2")
    add_centered_image(doc, "FIXED_IMAGE/elec2.PNG", width_in=4)
    
    add_numbered_subheading(doc, "Induct Stations Control Panel (Reference)", f"{counter}.3")
    add_centered_image(doc, "FIXED_IMAGE/elec3.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("Engines")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "Three-phase alternating current motors (Induction) will be used through a frequency converter. "
        "The engines will be coupled with a converter to improve consumption and reduce the carbon footprint."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph("All motors will have appropriate IP ratings.")
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Sensors")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The sensors will be supplied, standardized by type, with connector, with a cable length suitable for "
        "easy extraction, suitably protected from possible impacts."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Control command")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The proposed solution is based on SIEMENS Programmable Logic Controller technology (PLC) platform. "
        "The entire system will be logically divided into Zones (Sorter / Feed Line / Loop), each managed by a PLC. "
        "The planned primary communication protocol is going to be ProfiNet."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Conveyor interface")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The frequency converter of each conveyor allows the acquisition of the signals of the "
        "sensors/actuators/GIOs associated with it (e.g. conveyor end detection photocells, blockage detection "
        "photocells). Each frequency converter will be connected in series by means of the ProfiNet field bus."
    )
    apply_normal_style(p)

def build_wcs_section(doc, counter, client_name):
    """Build WCS CONTROLIT section with COMPLETE content"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Falcon's WCS CONTROLIT", counter=counter)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs1.PNG", width_in=3.0)
    
    p = doc.add_paragraph(
        "Falcon WCS (Warehouse Control System) is an in-house developed IT solution by Falcon Autotech, "
        "serving as the brain behind the company's sortation solutions. It manages the real-time movement "
        "of goods and data across the system, ensuring efficient operations in high-throughput warehouses. "
        "Falcon WCS integrates seamlessly with Warehouse Management Systems (WMS), Transport Management "
        "Systems (TMS), and other external applications via APIs to enhance operational efficiency."
    )
    apply_normal_style(p)
    
    # A. System Architecture
    add_numbered_subheading(doc, "System Architecture", f"{counter}.1")
    
    p = doc.add_paragraph()
    run = p.add_run("High-Level Design (HLD) Overview")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The Falcon WCS integrates with external systems like the Warehouse Management System (WMS) and "
        "Transport Management System (TMS). Communication occurs via APIs / WSDL / MQ communication "
        "protocols, ensuring smooth data flow for order management, shipment tracking, and other critical "
        "operations."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs2.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("Key Components:")
    run.bold = True
    apply_normal_style(p)
    
    # Presentation & Session Layer
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("Presentation and Session Layer:")
    run.bold = True
    apply_normal_style(p)
    
    for item in [
        "MySQL Database: Stores operational data, shipment details, and sortation instructions.",
        "Sorter Services: Responsible for managing sorting logic and directing parcels to appropriate destinations.",
        "Dashboard: Provides a user interface for real-time monitoring of warehouse operations and performance metrics.",
        "Integration Services: Handles communication with external systems (e.g., WMS, TMS) and ensures data consistency across platforms."
    ]:
        p = doc.add_paragraph(item, style="List Bullet 2")
        apply_normal_style(p)
    
    # Application Layer
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("Application Layer:")
    run.bold = True
    apply_normal_style(p)
    
    for item in [
        "Image Services: Processes and manages images captured during the sortation process.",
        "ICR Software: Utilizes Image Character Recognition to read parcel labels and identify shipment information.",
        "PLC Software: Interfaces with Programmable Logic Controllers to manage the physical movement of parcels and control sortation equipment."
    ]:
        p = doc.add_paragraph(item, style="List Bullet 2")
        apply_normal_style(p)
    
    # Transport Layer
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("Transport Layer:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "Sorter PLCs: Receive commands from the session layer (Sorter Services) and execute sorting operations based on real-time data.",
        style="List Bullet 2"
    )
    apply_normal_style(p)
    
    # System Communication
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run("System Communication:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "All layers are connected via a stacked switch, which provides internet and intranet connectivity. "
        "Communication between the sortation system and external systems for results or shipment data occurs through this switch.",
        style="List Bullet 2"
    )
    apply_normal_style(p)
    
    # B. High Availability Architecture
    add_numbered_subheading(doc, "High Availability Architecture", f"{counter}.2")
    
    p = doc.add_paragraph(
        "The Falcon WCS architecture ensures uninterrupted operations using a High Availability (HA) server setup. "
        "The system is designed to handle both planned and unplanned downtime, providing robust mechanisms for "
        "failover, replication, and data redundancy."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs3.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("Key Components and Features of the High Availability Architecture:")
    run.bold = True
    apply_normal_style(p)
    
    # Component descriptions
    components = [
        ("Stacked Switch:", [
            "Centralizes data exchange between NAS, nodes, domain controller (DC), and peripherals.",
            "Analyzes packet headers to reduce unnecessary data transmission, enhancing LAN efficiency."
        ]),
        ("Domain Controller:", [
            "Heartbeat Monitoring: Tracks the status of nodes and initiates VM failover when necessary.",
            "Image Hosting: Stores and manages images received from the ICR (Image Character Recognition)."
        ]),
        ("NAS (Network Attached Storage):", [
            "Centralized data storage providing access to connected devices and virtual machines.",
            "Redundancy: Two NAS boxes with mirrored drives ensure data protection and availability, offering a failsafe against hardware failure."
        ]),
        ("Node:", [
            "Hyper Terminals: Nodes host and manage virtual machines (VMs) to run the warehouse control systems and related applications.",
            "Clustering: Nodes are clustered using Microsoft Windows Cluster to enable failover protection, ensuring continuous operation even in case of hardware failure."
        ]),
        ("Virtual Machine & InnoDB Cluster:", [
            "Primary VM: Hosts Falcon WCS services, while a secondary backup on the node ensures failover through network load balancing (NLB).",
            "InnoDB Cluster: Ensures data replication using a Master–Slave–Slave setup for MySQL databases, maintaining consistency and availability."
        ]),
        ("NAS Cluster:", [
            "Unified File System: NAS nodes share files across the cluster, ensuring no data loss during failover or disaster recovery.",
            "Backup NAS: Provides redundancy by replicating data between two NAS boxes, further safeguarding against failures."
        ])
    ]
    
    for comp_title, comp_items in components:
        p = doc.add_paragraph()
        run = p.add_run(comp_title)
        run.bold = True
        apply_normal_style(p)
        
        for comp_item in comp_items:
            p = doc.add_paragraph(comp_item, style='List Bullet')
            apply_normal_style(p)
    
    # Disaster Handling
    p = doc.add_paragraph()
    run = p.add_run("Disaster Handling:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph("Recovery Time Objective (RTO) & Data Loss Objective (RPO):", style='List Bullet')
    apply_normal_style(p)
    
    for disaster_item in [
        "VM Cluster Failure: RTO = 1 hour; RPO = 1 hour.",
        "Node Failure: No impact with a single failure; RTO = 4 hours if both nodes fail.",
        "NAS Failure: Backup NAS available with no downtime, ensuring continued operation."
    ]:
        p = doc.add_paragraph(disaster_item, style="List Bullet 2")
        apply_normal_style(p)
    
    # C. WCS User Interface
    add_numbered_subheading(doc, "WCS User Interface", f"{counter}.3")
    
    p = doc.add_paragraph(
        "The Falcon WCS features a robust, user-friendly dashboard that provides real-time visibility "
        "into warehouse and sortation operations. The dashboard serves as the primary interface for "
        "monitoring key system metrics, tracking performance, and ensuring smooth operations."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Dashboard Overview")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The WCS dashboard offers real-time data visualization, helping warehouse operators and IT teams "
        "make data-driven decisions. Users can monitor system health, performance, and detect anomalies "
        "through an intuitive graphical interface."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Key Features of the Dashboard:")
    run.bold = True
    apply_normal_style(p)
    
    dashboard_features = [
        "System Health Monitoring: Displays metrics such as CPU utilization, memory usage, disk performance, and system load across the infrastructure.",
        "Real-Time Sortation Monitoring: Shows the real-time movement of parcels within the sortation system, including chute assignments and shipment statuses.",
        "Error Reporting: Notifies users of system errors, network disruptions, and potential failures in real time, allowing for quick resolution and minimal downtime.",
        "Performance Metrics: Provides detailed reports on sortation throughput, parcel handling times, and system efficiency to ensure that warehouse targets are met.",
        "User Role Management: The dashboard allows different levels of access based on user roles, ensuring that the right personnel can view or manage the system as needed."
    ]
    
    for feature in dashboard_features:
        p = doc.add_paragraph(feature, style='List Bullet')
        apply_normal_style(p)
    
    p = doc.add_paragraph(
        "In the context of this IT dashboard, the following user interactive screens are provided:"
    )
    apply_normal_style(p)
    
    # Dashboard screens with images
    dashboard_screens = [
        ("Dashboard (Home Screen): Provides an overview of important metrics, data visualizations, and summary information related to the IT system or processes.", "FIXED_IMAGE\\wcs4.PNG"),
        ("Live Bags: Displays real-time information and status updates regarding bags or parcels currently in transit or being processed.", "FIXED_IMAGE\\wcs5.PNG"),
        ("Bay Status: Offers insights into the status and availability of different processing bays or areas within the system.", "FIXED_IMAGE\\wcs6.PNG"),
        ("Processed Packages: Shows details and statistics related to packages or items that have been successfully processed or handled by the system.", "FIXED_IMAGE\\wcs7.PNG"),
        ("Configuration Setting: Enables users to configure and customize various settings and parameters within the IT system or dashboard.", "FIXED_IMAGE\\wcs8.PNG")
    ]
    
    for screen_desc, screen_img in dashboard_screens:
        p = doc.add_paragraph(screen_desc, style='List Bullet')
        apply_normal_style(p)
        add_centered_image(doc, screen_img)
    
    # Additional screens without images
    additional_screens = [
        "Report & Analysis: Allows users to generate and access comprehensive reports, analytics, and insights based on the data collected by the IT dashboard.",
        "Rejection Bay Mapping: Provides functionality to map and manage rejection bays or areas where packages are deemed unsuitable for processing.",
        "Alarms: Displays alerts, notifications, or alarms related to system events, errors, or anomalies that require attention or investigation.",
        "Calibration Settings: Allows users to adjust and calibrate system settings, parameters, or sensors to ensure accurate and reliable performance.",
        "Operator Management: Offers features and tools to manage and monitor the operators or personnel responsible for operating the IT system.",
        "User Management: Provides functionality to manage user accounts, permissions, roles, and access levels within the IT dashboard.",
        "User Guide: The 'User Guide' page offers comprehensive documentation and instructions on how to use the IT dashboard effectively. It serves as a reference guide for users."
    ]
    
    for screen in additional_screens:
        p = doc.add_paragraph(screen, style='List Bullet')
        apply_normal_style(p)
    
    # D. Communication Architecture
    add_numbered_subheading(doc, "Communication Architecture", f"{counter}.4")
    
    p = doc.add_paragraph(
        "Falcon WCS operates within a highly interconnected system, ensuring seamless communication between "
        "the WCS server, on-premises devices (such as sorter PLCs, PTL devices, 1D scanners, and HHT devices), "
        "and client systems. This communication architecture facilitates real-time data exchange and operational "
        "control, optimizing sortation processes and warehouse efficiency."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\wcs9.PNG")
    
    p = doc.add_paragraph()
    run = p.add_run("On-Premises Communication")
    run.bold = True
    apply_normal_style(p)
    
    # Communication devices
    comm_devices = [
        ("Sorter PLC Devices:", [
            "Protocol: Falcon WCS communicates with sorter PLCs using either the Siemens S7 protocol or the Omron communication protocol.",
            "Functionality: The sorter PLC devices receive sortation instructions from the WCS and execute the sorting process by directing parcels to the appropriate chute based on the system's real-time data."
        ]),
        ("PTL (Pick-to-Light) Devices:", [
            "Protocol: PTL devices communicate with Falcon WCS using the TCP/IP protocol.",
            "Functionality: The system sends commands to the PTL devices for guiding manual picking operations by lighting up indicators at the appropriate bins or shelves, improving operational accuracy and speed."
        ]),
        ("1D Scanners:", [
            "Protocol: These barcode scanners also use the TCP/IP protocol to communicate with the WCS.",
            "Functionality: The scanners capture barcode data from the parcels, and this information is sent to the WCS for processing, such as determining sorting destinations."
        ]),
        ("HHT (Handheld Terminal) Devices:", [
            "Protocol: The wireless HHT devices communicate with Falcon WCS over Wi-Fi.",
            "Functionality: The HHT devices send scan input data (e.g., barcodes) to the server over Wi-Fi. The WCS processes this data and sends the required output instructions back to the HHT device and associated PTL devices. The HHT device executes these instructions, facilitating real-time decision making and execution for operators."
        ])
    ]
    
    for device_title, device_items in comm_devices:
        p = doc.add_paragraph()
        run = p.add_run(device_title)
        run.bold = True
        apply_normal_style(p)
        
        for device_item in device_items:
            p = doc.add_paragraph(device_item, style='List Bullet')
            apply_normal_style(p)
    
    # E. Client Communication
    add_numbered_subheading(doc, "Client Communication", f"{counter}.5")
    
    p = doc.add_paragraph()
    run = p.add_run("Data Transfer Methods:")
    run.bold = True
    apply_normal_style(p)
    
    transfer_methods = [
        "API: Falcon WCS can communicate processed data to client systems through API calls, allowing for seamless integration with external software.",
        "MQ (Message Queuing): Falcon WCS can also send data via message queues, ensuring reliable delivery of messages even during network downtime.",
        "WSDL/XML: For structured data exchanges, Falcon WCS supports WSDL and XML formats for client communication.",
        f"Other Protocols: Additional methods for data transfer may include customized protocols depending on {client_name}'s requirements."
    ]
    
    for method in transfer_methods:
        p = doc.add_paragraph(method, style='List Bullet')
        apply_normal_style(p)
    
    p = doc.add_paragraph()
    run = p.add_run("Purpose:")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The data sent to the client can include sortation results, system performance reports, and operational "
        "analytics, which can be used for further processing or reporting within external systems like Warehouse "
        "Management Systems (WMS) and Transport Management Systems (TMS).",
        style='List Bullet'
    )
    apply_normal_style(p)
    
    # F. HAA Server Specifications
    add_numbered_subheading(doc, f"HAA Server Specifications (In {client_name}'s Scope)", f"{counter}.6")
    
    p = doc.add_paragraph("20-core configuration with 128 GB RAM in T440 and 64 GB RAM in T40.")
    apply_normal_style(p)
    
    # Server spec table
    table = doc.add_table(rows=1, cols=3)
    apply_table_style(table)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "SN"
    hdr_cells[1].text = "Description"
    hdr_cells[2].text = "Qty"
    
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Calibri (Body)'
                run.font.size = Pt(11)
    
    server_specs = [
        ("1", "DELL PowerEdge T440 Server", "2"),
        ("2", "Intel Xeon Silver 4210R 2.4G, 10C/20T, 9.6GT/s, 13.75M Cache, Turbo, HT (100W) DDR4-2400", "4"),
        ("3", "32GB RDIMM, 3200MT/s, Dual Rank", "8"),
        ("4", "480GB SSD SATA Read Intensive 6Gbps 512n 2.5in Hot-plug Drive, 1 DWPD", "4"),
        ("5", "H730P RAID Controller, 2GB NV Cache, Adapter, Low Profile", "2"),
        ("6", "Broadcom 5720 Dual Port 1Gb On-Board LOM", "2"),
        ("7", "Broadcom 57416 Dual Port 10Gb Base-T, OCP NIC 3.0", "2"),
        ("8", "Power Cord, C13, 1.8M, 250V, 10A (India BIS, IS1293)", "4"),
        ("9", "Dual, Hot-Plug, Redundant Power Supply (1+1), 750W", "2")
    ]
    
    for sn, desc, qty in server_specs:
        row_cells = table.add_row().cells
        row_cells[0].text = sn
        row_cells[1].text = desc
        row_cells[2].text = qty
        
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                apply_normal_style(paragraph)

def build_scada_section(doc, counter, client_name):
    """Build SCADA section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Falcon's Visual Inspection System (SCADA)", counter=counter)
    
    p = doc.add_paragraph(
        "SCADA stands for Supervisory Control and Data Acquisition. It is a system of hardware "
        "and software components that allows for remote monitoring, control, and data acquisition "
        "of industrial processes or facilities."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph(
        f"The Visualization system provided by FALCON (or SCADA) allows the monitoring and control of "
        f"the different systems delivered for the {client_name}. This SCADA system receives from each "
        f"monitored sub-system all information on their operating status in real time."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph("At the system monitoring level, the functions performed are:")
    apply_normal_style(p)
    
    functions = [
        "Field data acquisition.",
        "Animated visualization of equipment.",
        "Representation of the operating mode of the system (nominal, contingency, etc.).",
        "Alarm management.",
        "Alarm history management.",
        "Diagnostic help.",
        "Failure detection.",
        "Equipment control.",
        "Statistics on equipment operation.",
        "Historical Statistical Report.",
        "Recording and archiving.",
        "Safety operator interface.",
    ]
    
    for item in functions:
        p = doc.add_paragraph(item, style='List Bullet')
        apply_normal_style(p)
    
    add_numbered_subheading(doc, "FIELD DATA ACQUISITION", f"{counter}.1")
    p = doc.add_paragraph(
        "The field data acquisition function is performed by the SCADA system connected to the sorters' PLCs. "
        "The communication with the PLCs is done using equipped CPU cards that are able to manage the "
        "communication with the PLC on the Industrial Ethernet network, without overloading the server."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\main_plc.PNG")
    
    add_numbered_subheading(doc, "ANIMATED SYSTEM VISUALIZATION", f"{counter}.2")
    p = doc.add_paragraph(
        "The animated view represents the dynamic graphical user interface that allows real-time monitoring "
        "of the controlled systems and the execution of their control procedures."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\animated_sys_v1.PNG")
    add_centered_image(doc, "FIXED_IMAGE\\animated_sys_v2.PNG")
    
    add_numbered_subheading(doc, "ALARM MANAGEMENT", f"{counter}.3")
    p = doc.add_paragraph(
        "The alarm pages display a series of information to identify the nature of the alarm or event, "
        "the elements involved and the time."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\alarm.PNG")

def build_key_components_section(doc, counter, components_df):
    """Build Key Components Make section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Key Components Make", counter=counter)
    
    table = doc.add_table(rows=1, cols=2)
    apply_table_style(table)
    
    hdr = table.rows[0].cells
    hdr[0].text = "Items"
    hdr[1].text = "Make"
    
    for run in hdr[0].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
    for run in hdr[1].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
    
    for _, row in components_df.iterrows():
        r = table.add_row().cells
        r[0].text = str(row.get("Items", ""))
        r[1].text = str(row.get("Make", ""))
        
        for cell in r:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(11)

def build_safety_section(doc, counter):
    """Build Principal of Safety section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Principal of Safety", counter=counter)
    
    p = doc.add_paragraph()
    run = p.add_run("1. E-Stops")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\e-stop.png", width_in=4.5)
    
    p = doc.add_paragraph()
    run = p.add_run("            a. At Every Conveyor Module – Both sides")
    apply_normal_style(p)
    add_centered_image(doc, "FIXED_IMAGE\\at-every.png", width_in=4.5)
    
    p = doc.add_paragraph()
    run = p.add_run("             b. VDS Chutes")
    apply_normal_style(p)
    add_centered_image(doc, "FIXED_IMAGE\\emergency-vds.png", width_in=4.5)
    
    p = doc.add_paragraph()
    run = p.add_run("2. Pull Cords Switch")
    run.bold = True
    apply_normal_style(p)
    
    p = doc.add_paragraph("Required for Infeed & Takeout Conveyors")
    apply_normal_style(p)
    
    # Add pull cord images
    table_pc = doc.add_table(rows=1, cols=2)
    if os.path.exists("FIXED_IMAGE\\pull-cords.PNG"):
        p = table_pc.rows[0].cells[0].add_paragraph()
        run = p.add_run()
        run.add_picture("FIXED_IMAGE\\pull-cords.PNG", width=Inches(3))
    
    if os.path.exists("FIXED_IMAGE\\puul-cords-arch.PNG"):
        p = table_pc.rows[0].cells[1].add_paragraph()
        run = p.add_run()
        run.add_picture("FIXED_IMAGE\\puul-cords-arch.PNG", width=Inches(3))
    
    p = doc.add_paragraph()
    run = p.add_run("3. Fencing")
    run.bold = True
    apply_normal_style(p)
    
    table_fence = doc.add_table(rows=1, cols=2)
    table_fence.rows[0].cells[0].text = "a. Between Inducts \nb. Between Inducts and Sorter"
    
    if os.path.exists("FIXED_IMAGE\\fencing.png"):
        p = table_fence.rows[0].cells[1].add_paragraph()
        run = p.add_run()
        run.add_picture("FIXED_IMAGE\\fencing.png", width=Inches(2.5))
    
    p = doc.add_paragraph()
    run = p.add_run("4. Leg Guards")
    run.bold = True
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\leg_gurads.png", width_in=4.5)

def build_infrastructure_section(doc, counter):
    """Build Infrastructure section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Infrastructure", counter=counter)
    
    sections = [
        ("a. Fire Protection-", 
         "Falcon's scope does not cover the design or provision of fire protection infrastructure, "
         "utilities, or related services. It is expected that the customer's sprinkler contractor "
         "will design and supply the in-rack sprinkler systems, including connectors and mounting "
         "brackets. These designs should be submitted to Falcon for review during the engineering phase."),
        
        ("b. Power Supply-",
         "The Customer must provide temporary power for installation and permanent power for "
         "commissioning. Protected multi-gang power points for workstations and peripherals will be "
         "supplied by the Customer, with planning for their locations done with the operations and "
         "IT teams."),
        
        ("c. Floor Requirements-",
         "The Customer must provide flooring with appropriate loading strength and space at the site. "
         "Falcon assumes that the floor slab will not contain corrosive materials that could affect "
         "standard fixings."),
        
        ("d. Estimated Floor Load-",
         "Estimated floor loads, including distributed and point loads, will be provided during the "
         "detailed engineering phase of the project."),
        
        ("e. Staging, Laydown and Assembly Area-",
         "The Customer is required to provide sufficient space on the same floor, adjacent to the "
         "installation site, for staging, storage, and equipment assembly."),
        
        ("f. Site Access and Unloading-",
         "The Customer is required to allocate sufficient on-site space for parking and staging "
         "shipping containers to facilitate Falcon's delivery schedule."),
        
        ("g. Lighting-",
         "All lighting is excluded from Falcon's scope of supply and must be provided by the Customer "
         "or their contractor. This includes lighting for service areas, operational areas, and beneath "
         "platforms and walkways.")
    ]
    
    for title, content in sections:
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        apply_normal_style(p)
        
        p = doc.add_paragraph(content)
        apply_normal_style(p)

def build_program_org_section(doc, counter, client_name, gantt_file):
    """Build Program Organisation section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Program Organisation", counter=counter)
    
    add_numbered_subheading(doc, "Program Schedule", f"{counter}.1")
    
    if gantt_file is not None:
        p = doc.add_paragraph()
        run = p.add_run()
        gantt_file.seek(0)
        run.add_picture(gantt_file, width=Inches(6))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p = doc.add_paragraph("Attach your timeline Gantt chart here.")
        apply_normal_style(p)
    
    add_numbered_subheading(doc, "Program Management", f"{counter}.2")
    
    p = doc.add_paragraph("For this program, proposed approach covers the following aspects:")
    apply_normal_style(p)
    
    bullets = [
        "Creation and monitoring of the project plan.",
        "Weekly/Fortnightly meeting to share project status.",
        "Scheduling of the resource management.",
        "Management of risks and opportunities.",
        "Management of the requirements.",
        "Management of the list of anomalies or reservations.",
    ]
    
    for b in bullets:
        p = doc.add_paragraph(b, style='List Bullet')
        apply_normal_style(p)
    
    p = doc.add_paragraph(
        "The Project will be closely monitored under Falcon's Governance model as structured below."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\governence.png", width_in=5.5)
    doc.add_page_break()
    add_numbered_subheading(doc, "Project Team", f"{counter}.3")
    
    p = doc.add_paragraph(
        f"Team of 3 to 4 member from Projects team will co-ordinate on regular basis with "
        f"{client_name} and internal stakeholders for smooth execution of the project."
    )
    apply_normal_style(p)
    
    p = doc.add_paragraph()
    
    run = p.add_run(f"{client_name}'s Team")
    run.bold = True
    run.font.size = Pt(30)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    add_centered_image(doc, "FIXED_IMAGE\\team.png", width_in=5.5)

def build_client_responsibility_section(doc, counter, client_name):
    """Build Client Responsibility section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Client Responsibility", counter=counter)
    
    sections = [
        (f"{client_name} Responsibilities During the Assembly and Commissioning Phase", [
            "Provision of the site complex and office area facilities.",
            "The possibility of authorizing access to the site and the execution of the installation work up to 7 days a week and 24 hours a day if deemed necessary and if requested by FALCON.",
            "Free provision, during the installation phase, of the power supply necessary for the installation activities (estimated at 20 kW).",
            "Provision, during the commissioning phase, of the power supply necessary for the operation of the shipment sorting system free of charge at the date of FALCON need.",
            "Provision of the IT system functionality in accordance with the specification at the date of FALCON need.",
            "The customer is responsible for a safe working environment.",
            "The customer makes arrangements for the working area(s) to be protected against direct weather influences.",
            "The customer provides adequate lighting, heating, and ventilation to create a normal working environment.",
        ]),
        (f"Responsibilities of {client_name} During the Tests", [
            "Provision of the test loads and barcode labels required for the tests.",
            "Provision of personnel required for test activities (loading and unloading operations).",
            "Provision of the necessary information to sort the shipments correctly.",
            "Verify with FALCON the quality and conformity of the test loads (labels, cartons).",
        ]),
        (f"{client_name} Responsibilities During the Training", [
            "Free from their usual work, the employees participate in the training for the duration of the training.",
            "Provision of a list of participants for each available training course 3 days before the start of the course.",
            "Provision of a classroom equipped with a whiteboard, video projector, projection screen, and enough space for desks or tables and chairs for the trainer and trained staff.",
        ])
    ]
    
    for title, bullets in sections:
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        apply_normal_style(p)
        
        for b in bullets:
            p = doc.add_paragraph(b, style='List Bullet')
            apply_normal_style(p)

def build_handover_section(doc, counter):
    """Build System Handover section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "System Handover", counter=counter)
    
    p = doc.add_paragraph(
        "The system handover will follow the workflow shown below. "
        "Each stage is described in the following sections."
    )
    apply_normal_style(p)
    
    add_centered_image(doc, "FIXED_IMAGE\\handover.PNG")
    
    sections = [
        ("Installation and Commissioning",
         "Completion of all activities required to bring the system to an operational state "
         "and ready for formal testing."),
        
        ("Pre-UAT",
         "Pre-UAT consists of checks and tests performed before the formal User Acceptance Testing. "
         "It ensures the system is stable, integrated and ready for end-user validation."),
        
        ("UAT (User Acceptance Test)",
         "UAT is performed by the client's users to verify that the solution meets agreed requirements "
         "and behaves as expected under real-world operating conditions."),
        
        ("Minor & Major Faults",
         "Any issues found during testing are categorized as minor or major faults. "
         "Minor faults affect limited areas without blocking successful test completion and are added "
         "to the snag list."),
        
        ("System Snag Points",
         "After UAT, all open minor faults are tracked as system snag points. "
         "Falcon shares a snag list with the client, detailing for each issue: description, date, location, "
         "category, responsible party, target completion date and verification / sign-off."),
        
        ("System Handover Letter",
         "After successful acceptance testing or closure of all snags, Falcon issues a handover letter "
         "confirming that the system has been installed and accepted by the client.")
    ]
    
    for title, content in sections:
        p = doc.add_paragraph()
        run = p.add_run(title + "\n")
        run.bold = True
        apply_normal_style(p)
        
        p = doc.add_paragraph(content)
        apply_normal_style(p)

def build_commercial_section(doc, counter, price_data, payment_terms, apply_bca):
    """Build Commercial section with price sheet"""
    add_numbered_heading(doc, "Commercial", counter=counter)
    
    if not price_data:
        p = doc.add_paragraph("Commercial details to be added.")
        apply_normal_style(p)
        return
    
    # Price Sheet Title
    title = price_data.get("price_sheet_title") or "Price Sheet"
    add_numbered_subheading(doc, title, f"{counter}.1")
    
    items = price_data.get("items", [])
    total_row = price_data.get("total_row")
    
    # Price Table: S. No | Component | Price
    table = doc.add_table(rows=1, cols=3)
    apply_table_style(table)
    
    hdr = table.rows[0].cells
    hdr[0].text = "S. No"
    hdr[1].text = "Component"
    hdr[2].text = "Price"
    
    for run in hdr[0].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
    for run in hdr[1].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
    for run in hdr[2].paragraphs[0].runs:
        run.font.bold = True
        run.font.name = 'Calibri (Body)'
        run.font.size = Pt(11)
    
    # Add price items
    for item in items:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item.get("s_no", ""))
        row_cells[1].text = str(item.get("label", ""))
        row_cells[2].text = str(item.get("price", ""))
        
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                apply_normal_style(paragraph)
    
    # Total row
    if total_row:
        row_cells = table.add_row().cells
        row_cells[0].text = ""
        row_cells[1].text = str(total_row.get("label", "Total"))
        row_cells[2].text = str(total_row.get("price", ""))
        
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                apply_normal_style(paragraph)
    
    # Optional BCA discount row
    if apply_bca and total_row:
        final_total_str = apply_bca_discount_to_price_data(price_data, 4.5)
        if final_total_str:
            row_cells = table.add_row().cells
            row_cells[0].text = ""
            row_cells[1].text = "Final Total (after 4.5% BCA Discount)"
            row_cells[2].text = final_total_str
            
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                    apply_normal_style(paragraph)
    
    # Payment Terms section
    if payment_terms:
        doc.add_paragraph("")  # spacing
        add_numbered_subheading(doc, "Payment Terms", f"{counter}.2")
        
        pt_table = doc.add_table(rows=1, cols=2)
        apply_table_style(pt_table)
        
        pt_hdr = pt_table.rows[0].cells
        pt_hdr[0].text = "Payment Percentage"
        pt_hdr[1].text = "Stage"
        
        for run in pt_hdr[0].paragraphs[0].runs:
            run.font.bold = True
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)
        for run in pt_hdr[1].paragraphs[0].runs:
            run.font.bold = True
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)
        
        for row in payment_terms:
            perc = str(row.get("Payment Percentage", "")).strip()
            stage = str(row.get("Stage", "")).strip()
            if not perc and not stage:
                continue
            r = pt_table.add_row().cells
            r[0].text = perc
            r[1].text = stage
            
            for cell in r:
                for paragraph in cell.paragraphs:
                    apply_normal_style(paragraph)

def build_warranty_section(doc, counter, warranty_type, duration, start_cond, extended_text, amc_text, transport_text):
    """Build Warranty Period section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Warranty Period", counter=counter)
    
    intro = f"Falcon's offered System comes with a {warranty_type.lower()} of {duration} (starts {start_cond})"
    if not intro.endswith("."):
        intro += "."
    
    if extended_text:
        intro += f" {extended_text}"
    if amc_text:
        intro += f" {amc_text}"
    
    p = doc.add_paragraph(intro)
    apply_normal_style(p)
    
    p = doc.add_paragraph("The warranty covers the following support:")
    apply_normal_style(p)
    
    coverage = [
        "24 X 7 Telephonic, Email and Remote Service Support when required.",
        "Regular Software updates and Bug Fixes.",
        "Supply of Mechanical and Electrical components in case of failure (excluding damages as mentioned in the Exclusion Clause).",
    ]
    
    for item in coverage:
        p = doc.add_paragraph(item, style='List Bullet')
        apply_normal_style(p)
    
    p = doc.add_paragraph("The following items are excluded from warranty:")
    apply_normal_style(p)
    
    exclusions = [
        "Normal wear and tear.",
        "Consumables.",
        "Faulty articles continued.",
        "Failure to comply with the manufacturer's recommendations.",
        "Negligence or abnormal use of equipment.",
    ]
    
    for item in exclusions:
        p = doc.add_paragraph(item, style='List Bullet')
        apply_normal_style(p)
    
    if transport_text:
        p = doc.add_paragraph(transport_text)
        apply_normal_style(p)

def build_exclusions_section(doc, counter, selected_exclusions):
    """Build Exclusions section"""
    doc.add_page_break()  # Start on new page
    add_numbered_heading(doc, "Exclusions", counter=counter)
    
    intro = (
        "The scope of supply includes all parts which are defined in the Supplier's quotation.\n"
        "All other parts which are not defined in the Supplier's quotation do not belong to the Supplier's "
        "scope of supply and are excluded. The following parts are also excluded:"
    )
    
    p = doc.add_paragraph(intro)
    apply_normal_style(p)
    
    fixed_exclusions = [
        "Construction Power",
        "Building infrastructure; building structure, doors, fire exits, levelling devices, "
        "building extinguisher and fire alarm system, building heating and lighting system.",
        "Electrical power supply and wiring to the main control cabinets.",
        "UPS for Controls and Drives",
        "Network cabling up to the main server rack.",
        "Intermediate wiring to parts which are to be supplied by the Purchaser/others.",
        "Emergency/Uninterruptable power supply.",
        "Fire-alarm and fire protection devices.",
        "Traffic and route markings.",
        "Laydown area / unloading and laydown area.",
        "Ram protection devices.",
        "Cat walks, bridges, maintenance aisles and platforms.",
        "All kind of network incl. Local Area Network (LAN/WLAN), exceeding the scope described in Scope of Supply.",
        "Any kind of civil work.",
        "Any adjustment of the Supplier's scope of supply to local rules and regulations.",
        "X-Ray machines.",
        "Roller cages / pallets.",
        "Simulation and 3D animation of the sorter system.",
        "Interface with other equipment not specified in this offer.",
        "Provision of facilities for the control room (furniture, air conditioning, heating, etc.).",
        "The supply and installation of fencing around the different corridors.",
        "Any item specifically indicated as not forming part of the subject matter of the Seller's supply in the offer documentation.",
    ]
    
    all_exclusions = fixed_exclusions + selected_exclusions
    
    for item in all_exclusions:
        p = doc.add_paragraph(item, style='List Bullet')
        apply_normal_style(p)

def build_proposed_system_technical_details_section(doc, counter, bom_json):
    """Build Proposed System Technical Details section with 3 subsections"""
    doc.add_page_break()
    
    add_numbered_heading(doc, "Proposed System Technical Details", counter=counter)
    
    if not bom_json or "sections" not in bom_json:
        p = doc.add_paragraph("No technical details available.")
        apply_normal_style(p)
        return
    
    section_number = 1
    for section in bom_json.get("sections", []):
        title = section.get("title", "")
        items = section.get("items", [])
        if not items:
            continue
        
        # Add subsection heading (e.g., "14.1 Mechanical equipment")
        add_numbered_subheading(doc, title, f"{counter}.{section_number}")
        section_number += 1
        
        # Create table with 4 columns: Pos., Qty., Description, Value
        table = doc.add_table(rows=1, cols=4)
        apply_table_style(table)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header row
        hdr = table.rows[0].cells
        hdr[0].text = "Pos."
        hdr[1].text = "Qty."
        hdr[2].text = "Description"
        hdr[3].text = "Value"
        
        # Format header row
        for cell in hdr:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(11)
                    run.font.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # Add data rows
        for item in items:
            row = table.add_row().cells
            
            # Pos.
            row[0].text = str(item.get("pos", ""))
            
            # Qty.
            row[1].text = item.get("qty") or ""
            
            # Description (multiple lines)
            desc_lines = item.get("description_lines") or []
            row[2].text = "\n".join(desc_lines)
            
            # Value (multiple lines)
            value_lines = item.get("value_lines") or []
            row[3].text = "\n".join(value_lines)
            
            # Format each cell
            for cell in row:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(11)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # Set column widths
        table.columns[0].width = Inches(0.5)
        table.columns[1].width = Inches(1.2)
        table.columns[2].width = Inches(3.5)
        table.columns[3].width = Inches(1.3)
        
        # Add spacing after table
        doc.add_paragraph()

def build_proposed_system_description_section(doc, counter, client_name, project_name, 
                                              process_flow_text, layout_png_path):
    """Build Proposed System Description section (5.0)"""
    doc.add_page_break()
    
    add_numbered_heading(doc, "Proposed System Description", counter=counter)
    
    # 5.1 Objective
    add_numbered_subheading(doc, "Objective", f"{counter}.1")
    objective_text = (
        "The purpose of this proposal is to present the design, manufacturing, "
        "installation, commissioning, testing, and acceptance testing of the Cross Belt Sorter "
        f"system for sorting shipments, as per {client_name} requirements."
    )
    p = doc.add_paragraph(objective_text)
    apply_normal_style(p)
    doc.add_paragraph("")
    
    # 5.2 Summary of the System (layout PNG)
    add_numbered_subheading(doc, "Summary of the System", f"{counter}.2")
    
    if layout_png_path and os.path.exists(layout_png_path):
        p = doc.add_paragraph(
            "The following layout view illustrates the overall arrangement of infeed conveyors, sorter loop, "
            "and output chutes for the proposed system."
        )
        apply_normal_style(p)
        doc.add_paragraph("")
        p = doc.add_paragraph()
        run = p.add_run()
        run.add_picture(layout_png_path, width=Inches(6.5))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("")
    else:
        p = doc.add_paragraph("The detailed layout is provided separately in the attached drawing.")
        apply_normal_style(p)
        doc.add_paragraph("")
    
    # 5.3 Process Flow of the System
    add_numbered_subheading(doc, "Process Flow of the System", f"{counter}.3")
    for line in process_flow_text.splitlines():
        line = line.strip()
        if not line: continue
        
        # Parse and apply bold formatting for **text**
        p = doc.add_paragraph()
        parts = re.split(r'(\*\*[^\*]+\*\*)', line)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                # Bold text
                text = part[2:-2]
                run = p.add_run(text)
                run.bold = True
            else:
                # Normal text
                run = p.add_run(part)
            run.font.name = "Calibri"
            run.font.size = Pt(11)
    doc.add_paragraph("")
    
    # 5.4 Main Benefits
    add_numbered_subheading(doc, "Main Benefits of the Proposed Solution", f"{counter}.4")
    benefits = [
        "High operational throughput.",
        "Low occupancy of floor space in the building.",
        "Narrow discharge centers for the increased number of splits in limited space.",
        (
            "FALCON's CBS can adapt to changing business requirements by adjusting its speed "
            "to match the operational throughput requirement, thereby leading to power savings "
            "and reduced system wear & tear."
        ),
    ]
    for b in benefits:
        p = doc.add_paragraph(b, style='List Bullet')
        apply_normal_style(p)

def generate_system_description_from_sd_sys(dxf_path: Path, costing_file_upload, project_name: str,
                                           client_name: str = "", pph_count: str = "", ipp_rate: str = "") -> tuple[str, dict, dict, dict]:
    """Generate System Description using `sys_desc_latest` utilities.

    Returns: (system_description_text, detected_json, tables_dict, cost_vals_dict)
    """
    # Build DXF JSON using existing extractor in the main app
    full = {}
    try:
        full = extract_dxf_components(dxf_path, project_name)
    except Exception:
        # fallback: try sys_desc_latest if it offers a full extractor
        try:
            full = sd_sys.extract_full_dxf_info(dxf_path)
        except Exception:
            full = {}

    # Compute metrics (prefer sd_sys implementation when available)
    try:
        metrics = sd_sys.compute_dxf_metrics(full)
    except Exception:
        logger.warning("sd_sys.compute_dxf_metrics failed; falling back to empty metrics")
        metrics = {}

    # Extract costing tables/values if provided
    tables = {}
    cost_vals = {}
    if costing_file_upload:
        try:
            if hasattr(costing_file_upload, "getvalue"):
                data = costing_file_upload.getvalue()
            elif isinstance(costing_file_upload, (bytes, bytearray)):
                data = bytes(costing_file_upload)
            else:
                data = costing_file_upload.read()

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tx:
                tx.write(data)
                tmp_xlsx = Path(tx.name)

            try:
                tables = sd_sys.extract_costing_tables(tmp_xlsx)
            except Exception:
                tables = {}
            try:
                cost_vals = sd_sys.extract_costing_values(tmp_xlsx)
            except Exception:
                cost_vals = {}
        except Exception:
            tables = {}
            cost_vals = {}

    # Prepare variables map
    variables = {}
    variables["UNIT"] = metrics.get("UNITS", "")
    variables["Type of CBS"] = metrics.get("TYPE OF CBS", "")
    variables["COUNT"] = str(metrics.get("TOTAL CONVEYOR / SUBCOMPONENT COUNT", ""))
    variables["Feedline count"] = str(metrics.get("FEEDLINE COUNT", ""))
    variables["DEGREE OF ANGLE MERGE"] = metrics.get("DEGREE OF ANGLE MERGE", "")
    variables["COUNT OF REJECTION CHUTE"] = str(metrics.get("COUNT OF REJECTION CHUTE", ""))
    variables["COUNT OF DISPERSION CHUTE"] = str(metrics.get("COUNT OF DISPERSION CHUTE", ""))
    variables["COUNT OF COLLECTION CHUTE"] = str(metrics.get("COUNT OF COLLECTION CHUTE", ""))
    variables["CBS HEIGHT FROM GROUND"] = cost_vals.get("CBS HEIGHT FROM GROUND", "")
    variables["PITCH LENGTH"] = cost_vals.get("PITCH LENGTH", "")
    variables["CLIENT'S NAME"] = (client_name or "").strip()
    variables["CLIENT NAME"] = variables["CLIENT'S NAME"]
    variables["PPH COUNT"] = (pph_count or "").strip()
    variables["IPP Rate"] = (ipp_rate or "").strip()
    variables["IPP RATE"] = (ipp_rate or "").strip()

    # GROQ detection
    det_txt = sd_sys.groq_chat(sd_sys.prompt_detect_components(full, metrics), temperature=0.1, max_tokens=2500)
    det_raw = sd_sys.extract_json(det_txt)
    detected_fallback = sd_sys.build_detected_json(metrics)
    detected = sd_sys.normalize_detected(det_raw, detected_fallback)

    # Enforce mandatory mappings
    if metrics.get("FS002 WITHOUT WEIGHING COUNT", 0) > 0:
        detected.setdefault("Parcel Inducts / Induction to Sorter", {})
        detected["Parcel Inducts / Induction to Sorter"].setdefault("Feedlines", {"Feedline Count": int(metrics.get("FEEDLINE COUNT", 1)), "Subcomponents": {}})
        detected["Parcel Inducts / Induction to Sorter"]["Feedlines"].setdefault("Subcomponents", {})
        detected["Parcel Inducts / Induction to Sorter"]["Feedlines"]["Subcomponents"]["Buffer Conveyor"] = 1

    if metrics.get("HAS_VDS_LOOP"):
        detected.setdefault("Infeed System", {})
        if "VDS Loop Conveyor" not in detected["Infeed System"]:
            detected["Infeed System"]["VDS Loop Conveyor"] = max(1, int(metrics.get("VDS LOOP COUNT", 1)))

    # Generate draft and run judge pass
    template_text = sd_sys.load_template_text() if hasattr(sd_sys, "load_template_text") else ""
    sys_desc_draft = sd_sys.groq_chat(sd_sys.prompt_generate_system_description_dynamic(template_text, detected, variables), temperature=0.2, max_tokens=4500)
    sys_desc_final = sd_sys.groq_chat(sd_sys.prompt_judge_fix_system_description(detected, sys_desc_draft), temperature=0.0, max_tokens=4500)

    return sys_desc_final, detected, tables, cost_vals
            

# Concept Description helper: insert the generated flowchart PNG into the main DOCX
def build_concept_description_section(doc, counter, png_bytes: bytes):
    """Add a Concept Description section containing the rendered flowchart PNG.

    This function inserts a numbered heading and places the PNG centered on the page.
    """
    try:
        doc.add_page_break()
    except Exception:
        pass

    add_numbered_heading(doc, "Concept Description", counter=counter)
    doc.add_paragraph()

    # Subheading for the flowchart
    add_numbered_subheading(doc, "Flowchart", f"{counter}.1")

    if not png_bytes:
        p = doc.add_paragraph("Flowchart image not available.")
        apply_normal_style(p)
        return

    try:
        p = doc.add_paragraph()
        run = p.add_run()
        image_stream = BytesIO(png_bytes)
        run.add_picture(image_stream, width=Inches(6.5))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

        # Small caption/note
        note = doc.add_paragraph("Click the diagram to open in draw.io for editing")
        note.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_normal_style(note)
    except Exception as e:
        p = doc.add_paragraph(f"Could not insert flowchart image: {e}")
        apply_normal_style(p)

# ==================== MAIN GENERATION ====================

st.markdown("---")
st.header("Generate Complete Document")

if st.button("Generate Final DOCX Document", type="primary", width='stretch'):
    with st.spinner("Generating your complete proposal document..."):
        try:
            # Validate required inputs
            if not client_name.strip():
                st.error("Please enter a Client Name")
                st.stop()
            
            if not project_name.strip():
                st.error("Please enter a Project Name")
                st.stop()
            
            # Executive summary will be generated after process flow is created from DXF
            exec_summary_text = None
            
            # Process DXF and generate Process Flow & Mermaid Flowchart if needed
            process_flow_text = None
            system_description_text = None
            flowchart_png_bytes = None
            layout_png_path = None
            dxf_json = None
            
            if (include_proposed_system or include_concept_desc) and dxf_layout_file:
                with st.spinner("🔧 Processing DXF and generating AI content..."):
                    try:
                        # Create temp directory
                        tmp_dir = Path(tempfile.mkdtemp(prefix="proposal_"))
                        
                        # Save DXF file
                        dxf_path = tmp_dir / dxf_layout_file.name
                        dxf_path.write_bytes(dxf_layout_file.getvalue())
                        
                        # Extract DXF components
                        st.info("📐 Extracting DXF components...")
                        dxf_json = extract_dxf_components(dxf_path, project_name)
                        
                        # Print raw DXF extraction to console
                        print("\n" + "="*80)
                        print("RAW DXF EXTRACTION RESULT")
                        print("="*80)
                        print(json.dumps(dxf_json, indent=2, ensure_ascii=False))
                        print("="*80 + "\n")
                        
                        # Generate Process Flow if Proposed System is included
                        if include_proposed_system:
                            # If a preview was already generated earlier in the session, reuse it
                            if st.session_state.get('preview_process_flow'):
                                process_flow_text = st.session_state.get('preview_process_flow')
                                iteration_details = st.session_state.get('iteration_details')
                                st.info("Using previously generated Process Flow from preview")
                            else:
                                st.info("✍️ Generating Process Flow with AI...")
                                process_flow_text, iteration_details = call_groq_for_process_flow(
                                    client_name, project_name, dxf_json
                                )
                                # Cache preview in session_state for potential reuse
                                try:
                                    st.session_state['preview_process_flow'] = process_flow_text
                                    st.session_state['iteration_details'] = iteration_details
                                except Exception:
                                    pass
                                st.success("✅ Process Flow generated")
                                time.sleep(2)  # Delay to avoid rate limits
                    except Exception as e:
                        st.warning(f"Could not process DXF for initial flow: {str(e)}")
                        dxf_json = None
                        process_flow_text = None
            
            # Generate cover letter AFTER process flow is created (to include high-level summary)
            cover_letter_text = None
            if offer_ref and sender_name:
                with st.spinner("Starting Build..."):
                    try:
                        # Create enhanced high-level summary from process flow and DXF data
                        process_flow_summary = ""
                        if process_flow_text:
                            lines = process_flow_text.strip().split('\n')
                            # Extract main system components from process flow
                            summary_components = []
                            for line in lines[:5]:  # Look at first 5 lines for better coverage
                                # Extract component names (remove numbering and description after colon)
                                if ':' in line:
                                    component = line.split(':')[0].strip()
                                    # Remove numbering (1., 2., etc.)
                                    component = component.lstrip('0123456789. ')
                                    if component and len(component) > 3:  # Avoid empty or very short strings
                                        summary_components.append(component)
                            
                            # Add key quantities from DXF if available
                            quantities = []
                            if dxf_json:
                                if dxf_json.get('total_chutes', 0) > 0:
                                    quantities.append(f"{dxf_json['total_chutes']} chutes")
                                if dxf_json.get('total_operators', 0) > 0:
                                    quantities.append(f"{dxf_json['total_operators']} operator stations")
                                # Add other relevant quantities if present
                                if dxf_json.get('scanner_systems', 0) > 0:
                                    quantities.append(f"{dxf_json['scanner_systems']} scanner systems")
                            
                            # Combine components and quantities into natural summary
                            if summary_components:
                                process_flow_summary = ", ".join(summary_components[:3])  # First 3 components
                                if quantities:
                                    process_flow_summary += f" with {', '.join(quantities[:2])}"  # Add up to 2 quantities
                        
                        cover_letter_text = call_groq_cover_letter(
                            client_name=client_name,
                            project_title=project_name,
                            offer_ref=offer_ref,
                            letter_date_str=letter_date.strftime("%B %d, %Y"),
                            executives_block=executives_text,
                            invitation_date=invitation_date_str,
                            meeting_date=meeting_date_str,
                            sender_name=sender_name,
                            sender_title=sender_title,
                            process_flow_summary=process_flow_summary
                        )
                        #st.success("Cover letter generated successfully")
                    except Exception as e:
                        st.warning(f"Could not generate cover letter: {str(e)}")
                        cover_letter_text = None
            
            # Continue processing DXF if needed
            if (include_proposed_system or include_concept_desc) and dxf_layout_file and process_flow_text:
                with st.spinner("🔧 Continuing AI content generation..."):
                    try:
                        # Reuse temp directory from DXF processing
                        if 'tmp_dir' not in locals():
                            tmp_dir = Path(tempfile.mkdtemp(prefix="proposal_"))
                        if 'dxf_path' not in locals() and dxf_layout_file:
                            dxf_path = tmp_dir / dxf_layout_file.name
                            if not dxf_path.exists():
                                dxf_path.write_bytes(dxf_layout_file.getvalue())
                        
                        # Generate System Description from process flow and DXF if Proposed System is included
                        if include_proposed_system and process_flow_text and dxf_json:
                            st.info("📋 Generating comprehensive System Description with AI...")
                            try:
                                # Use sys_desc_latest pipeline to produce final system description text
                                sys_desc_final, detected, sd_tables, sd_cost_vals = generate_system_description_from_sd_sys(
                                    dxf_path=dxf_path,
                                    costing_file_upload=costing_file if 'costing_file' in locals() else None,
                                    project_name=project_name,
                                    client_name=client_name,
                                    pph_count=pph_count,
                                    ipp_rate=ipp_rate,
                                )
                                system_description_text = sys_desc_final
                                # merge any extracted tables/cost_vals for potential later use
                                try:
                                    if sd_tables:
                                        tables.update(sd_tables)
                                except Exception:
                                    pass
                                st.success("✅ System Description generated")
                            except Exception as e:
                                st.warning(f"System Description generation failed: {e}")
                            time.sleep(2)  # Delay to avoid rate limits
                        
                        # Generate Executive Summary from process flow if included
                        if include_exec_summary and process_flow_text:
                            st.info("📝 Generating Executive Summary with AI...")
                            exec_summary_text = call_groq_exec_summary(process_flow_text, client_name, project_name)
                            st.success("✅ Executive Summary generated")
                            time.sleep(2)  # Delay to avoid rate limits
                        
                        # Generate Mermaid Flowchart if Concept Description is included
                        if include_concept_desc and process_flow_text:
                            st.info("🗺️ Generating Mermaid flowchart...")
                            mermaid_code = call_groq_for_mermaid(process_flow_text)
                            flowchart_png_bytes, render_log = generate_mermaid_png(mermaid_code)
                            st.success("✅ Flowchart rendered")
                            time.sleep(2)  # Delay to avoid rate limits
                        
                        # Handle layout PNG - either uploaded or convert from DXF
                        if layout_full_png:
                            # User uploaded a PNG - use it
                            layout_png_path = tmp_dir / layout_full_png.name
                            layout_png_path.write_bytes(layout_full_png.getvalue())
                            layout_png_path = str(layout_png_path)
                            st.success("✅ Using uploaded layout PNG")
                        else:
                            # No PNG uploaded - try to convert DXF to PNG
                            if CONVERTAPI_SECRET:
                                try:
                                    st.info("🔄 Converting DXF to PNG for layout visualization...")
                                    png_path = convert_dxf_to_png(dxf_path)
                                    if png_path and png_path.exists():
                                        layout_png_path = str(png_path)
                                        st.success("✅ DXF converted to PNG successfully")
                                    else:
                                        st.warning("⚠️ DXF to PNG conversion did not produce a file")
                                        layout_png_path = None
                                except Exception as e:
                                    st.warning(f"⚠️ Could not convert DXF to PNG: {str(e)}")
                                    layout_png_path = None
                            else:
                                st.warning("⚠️ CONVERTAPI_SECRET not configured. Cannot convert DXF to PNG. Please upload a PNG manually.")
                                layout_png_path = None
                        
                    except Exception as e:
                        st.warning(f"Could not process DXF file: {str(e)}")
                        process_flow_text = None
                        flowchart_png_bytes = None
                        layout_png_path = None
            
            # Process costing file if commercial section is included
            price_data = None
            payment_terms_data = None
            bca_discount = False
            
            if commercial_include and costing_file:
                with st.spinner("📊 Processing costing file with AI..."):
                    try:
                        # Read Overall Costing sheet
                        df = pd.read_excel(costing_file, sheet_name="Overall Costing", header=None)
                        sheet_csv = df.to_csv(index=False)
                        
                        # Call Groq to extract price sheet
                        price_data = call_groq_for_price_sheet(sheet_csv)
                        payment_terms_data = st.session_state.get("payment_terms", [])
                        bca_discount = apply_bca
                        
                        st.success("Price sheet generated from costing file")
                    except Exception as e:
                        st.warning(f"Could not process costing file: {str(e)}. Commercial section will be added as placeholder.")
                        price_data = None
            
            # Get client logo path - either from dropdown selection or uploaded file
            client_logo_path = None
            if selected_client != "None" and selected_client in CLIENT_LOGOS:
                # Use logo from dropdown selection
                client_logo_path = CLIENT_LOGOS[selected_client]
            elif client_logo:
                # Use uploaded logo - save temporarily
                client_logo_path = f"temp_client_logo.{client_logo.name.split('.')[-1]}"
                with open(client_logo_path, "wb") as f:
                    f.write(client_logo.getbuffer())
            
            # ==================== START WITH FRESH DOCUMENT ====================
            # Always start with a fresh document that has all standard Word styles
            doc = Document()
            
            # Ensure required list styles exist
            ensure_list_styles(doc)
            
            # Set default font for the document
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Calibri (Body)'
            font.size = Pt(11)
            
            # ==================== COVER LETTER (with header/footer) ====================
            if cover_letter_text:
                build_cover_letter_section(doc, cover_letter_text)
            
            # ==================== FRONT PAGE (WITH HEADER) ====================
            if cover_letter_text:
                build_front_page_section(doc, project_name, offer_ref, contact_name, contact_email, contact_phone, layout_png_path)
            
            # ==================== GLOSSARY ====================
            build_glossary_section(doc)
            
            # Start numbering from 2 because Glossary occupies number 1
            counter = 2
            
            # ==================== BUILD SECTIONS IN ORDER ====================
            
            # 1. Executive Summary
            if include_exec_summary and exec_summary_text:
                build_executive_summary_section(doc, exec_summary_text, counter)
                counter += 1
            
            # 2. Company Profile
            if include_company_profile:
                build_company_profile_section(doc, counter)
                counter += 1

            # 3. Handled Shipment Spectrum
            if include_handled_spectrum:
                build_handled_spectrum_section(doc, counter, project_name, client_name)
                counter += 1

            # 4. Proposed System Description
            if include_proposed_system and process_flow_text:
                build_proposed_system_description_section(doc, counter, client_name, project_name, 
                                                         process_flow_text, layout_png_path)
                counter += 1
            
            # 4.1 System Description (Detailed)
            if include_proposed_system and system_description_text:
                doc.add_page_break()
                add_numbered_heading(doc, "System Description", counter=counter)
                # Insert the LLM-produced system description text preserving markers
                for line in (system_description_text or "").splitlines():
                    # Preserve blank lines
                    if not line.strip():
                        doc.add_paragraph("")
                        continue
                    p = doc.add_paragraph(line)
                    apply_normal_style(p)
                counter += 1
            
            # 4.2 Description of Components (from Loop CBS Excel if available)
            sorter_spec = None
            if costing_file is not None:
                try:
                    sheet_name, df = load_loop_cbs_sheet_from_excel(costing_file.getvalue())
                    if sheet_name and df is not None:
                        sheet_text = df_to_compact_text(df)
                        sorter_spec = call_groq_for_sorter_spec(sheet_name, sheet_text)
                        if sorter_spec:
                            st.success("✓ Extracted sorter specifications from Loop CBS sheet")
                except Exception as e:
                    st.warning(f"Could not extract Loop CBS data for component description: {e}")
            
            build_description_of_components_section(doc, counter, sorter_spec)
            counter += 1
            
            # 5. Proposed System Technical Details (from Quote Master Excel if available)
            bom_json = None
            if costing_file is not None:
                try:
                    # Load Quote Master sheet from costing file
                    with pd.ExcelFile(io.BytesIO(costing_file.getvalue())) as xls:
                        quote_master_sheet = None
                        for sheet in xls.sheet_names:
                            if sheet.lower().strip() == "quote master":
                                quote_master_sheet = sheet
                                break
                        
                        if quote_master_sheet:
                            df_quote = pd.read_excel(xls, sheet_name=quote_master_sheet)
                            sheet_text = df_to_compact_text_quote_master(df_quote)
                            bom_json = call_groq_for_bom(sheet_text)
                            if bom_json:
                                st.success("✓ Generated Proposed System Technical Details from Quote Master")
                        else:
                            st.info("ℹ️ Quote Master sheet not found in costing file - skipping technical details section")
                except Exception as e:
                    st.warning(f"Could not extract Quote Master data: {e}")
            
            if bom_json:
                build_proposed_system_technical_details_section(doc, counter, bom_json)
                counter += 1
            
            # 6. Concept Description
            if include_concept_desc and flowchart_png_bytes:
                build_concept_description_section(doc, counter, flowchart_png_bytes)
                counter += 1

            # 7. Capacity Calculations Section (optional)
            if include_capacity_section and capacity_excel is not None:
                build_capacity_calculations_section(doc, counter, client_name, project_name, capacity_excel)
                counter += 1

            # 8. Electrical System
            if elec_include:
                build_electrical_section(doc, counter)
                counter += 1
            
            # 9. Falcon WCS CONTROLIT
            if wcs_include:
                build_wcs_section(doc, counter, client_name)
                counter += 1
            
            # 10. Falcon Visual Inspection System (SCADA)
            if scada_include:
                build_scada_section(doc, counter, client_name)
                counter += 1
            
            # 11. Key Components Make
            if key_include:
                build_key_components_section(doc, counter, key_components_edited)
                counter += 1
            
            # 12. Principal of Safety
            if safety_include:
                build_safety_section(doc, counter)
                counter += 1
            
            # 13. Infrastructure
            if infra_include:
                build_infrastructure_section(doc, counter)
                counter += 1
            
            # 14. Program Organisation
            if prog_include:
                build_program_org_section(doc, counter, client_name, prog_gantt)
                counter += 1
            
            # 15. Client Responsibility
            if client_resp_include:
                build_client_responsibility_section(doc, counter, client_name)
                counter += 1
            
            # 16. System Handover
            if handover_include:
                build_handover_section(doc, counter)
                counter += 1
            
            # 17. Commercial
            if commercial_include:
                build_commercial_section(doc, counter, price_data, payment_terms_data, bca_discount)
                counter += 1
            
            # 18. Warranty Period
            if warranty_include:
                build_warranty_section(doc, counter, warranty_type, warranty_duration, 
                                      warranty_start, warranty_extended_text, 
                                      warranty_amc_text, warranty_transport_text)
                counter += 1
            
            # 19. Exclusions
            if exclusion_include:
                build_exclusions_section(doc, counter, selected_exclusions)
                counter += 1
            
            # ==================== DETECT AND INSERT GLOSSARY ====================
            # Scan the entire document for glossary terms
            try:
                full_text = extract_full_text_from_docx(doc)
                detected_terms = find_terms_in_text(full_text)
                if detected_terms:
                    st.info(f"📖 Detected {len(detected_terms)} glossary terms in the document")
                    # Create a temporary document with just the glossary section
                    temp_glossary_doc = Document()
                    build_glossary(temp_glossary_doc, detected_terms)
                    # Save to buffer and reload
                    temp_buffer = io.BytesIO()
                    temp_glossary_doc.save(temp_buffer)
                    temp_buffer.seek(0)
                    glossary_doc = Document(temp_buffer)
                    # Find the index of the Executive Summary heading
                    glossary_placeholder_index = None
                    word_ns_tag = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'
                    for idx, el in enumerate(doc.element.body):
                        if hasattr(el, 'iter'):
                            for t in el.iter(word_ns_tag):
                                if t.text and 'Executive Summary' in t.text:
                                    glossary_placeholder_index = idx
                                    break
                        if glossary_placeholder_index is not None:
                            break
                    # Insert glossary elements before Executive Summary or at end if not found
                    glossary_elements = []
                    for element in glossary_doc.element.body:
                        if element.tag.endswith('sectPr'):
                            continue
                        glossary_elements.append(element)

                    word_ns_uri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                    word_paragraph_tag = f'{{{word_ns_uri}}}p'
                    word_break_tag = f'{{{word_ns_uri}}}br'

                    def is_page_break(elem):
                        if elem is None or elem.tag != word_paragraph_tag:
                            return False
                        for br in elem.iter(word_break_tag):
                            if br.get(f'{{{word_ns_uri}}}type') == 'page':
                                return True
                        return False

                    def build_page_break_element():
                        para = OxmlElement('w:p')
                        run = OxmlElement('w:r')
                        br = OxmlElement('w:br')
                        br.set(qn('w:type'), 'page')
                        run.append(br)
                        para.append(run)
                        return para

                    if glossary_placeholder_index is not None:
                        prev_element = None
                        if glossary_placeholder_index > 0:
                            prev_element = doc.element.body[glossary_placeholder_index - 1]
                        if not is_page_break(prev_element):
                            glossary_elements.insert(0, build_page_break_element())
                        for i, element in enumerate(glossary_elements):
                            doc.element.body.insert(glossary_placeholder_index + i, element)
                        st.success("✓ Glossary section inserted before Executive Summary")
                    else:
                        prev_element = doc.element.body[-1] if len(doc.element.body) > 0 else None
                        if not is_page_break(prev_element):
                            glossary_elements.insert(0, build_page_break_element())
                        for element in glossary_elements:
                            doc.element.body.append(element)
                        st.info("Executive Summary heading not found, glossary inserted at end.")
                else:
                    st.info("ℹ️ No glossary terms detected in the document")
            except Exception as e:
                st.warning(f"Could not generate glossary: {e}")
            
            # ==================== INSERT COVER PAGE AT BEGINNING ====================
            
            # ==================== INSERT COVER PAGE AT BEGINNING ====================
            # Now prepend cover page at the beginning if cover letter was generated
            if cover_letter_text:
                try:
                    # Get client logo bytes for cover page
                    cover_client_logo_bytes = None
                    if client_logo_path and os.path.exists(client_logo_path):
                        with open(client_logo_path, "rb") as f:
                            cover_client_logo_bytes = f.read()
                    
                    # Create cover page using template
                    cover_page_buffer = create_cover_page(
                        client_logo=cover_client_logo_bytes,
                        client_name=client_name,
                        project_title=project_name
                    )
                    
                    # Save main document to temp buffer
                    temp_main_buffer = io.BytesIO()
                    doc.save(temp_main_buffer)
                    temp_main_buffer.seek(0)
                    
                    # Load cover page document (from template)
                    cover_doc = Document(cover_page_buffer)
                    
                    # Load main content document (all our generated content with images)
                    main_doc = Document(temp_main_buffer)
                    
                    # Try using Composer for proper merge (preserves all relationships including images)
                    try:
                        composer = Composer(cover_doc)
                        composer.append(main_doc)
                        
                        # Save composed document

                        composed_buffer = io.BytesIO()
                        composer.save(composed_buffer)
                        composed_buffer.seek(0)
                        # Load as final document
                        doc = Document(composed_buffer)

                    except (ImportError, NameError, AttributeError):
                        # Fallback: If Composer not available, use element insertion
                        cover_elements = []
                        for element in cover_doc.element.body:
                            if element.tag.endswith('sectPr'):
                                continue
                            cover_elements.append(element)
                        for i, element in enumerate(cover_elements):
                            main_doc.element.body.insert(i, element)
                        page_break_xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>'
                        page_break_element = parse_xml(page_break_xml)
                        main_doc.element.body.insert(len(cover_elements), page_break_element)
                        doc = main_doc

                except Exception as e:
                    st.warning(f"Could not insert cover page: {str(e)}. Cover page will be skipped.")

            # === Apply header/footer to the finalized document ===
            try:
                client_logo_bytes = None
                if client_logo_path and os.path.exists(client_logo_path):
                    with open(client_logo_path, "rb") as fh:
                        client_logo_bytes = fh.read()

                create_header_footer(
                    doc,
                    client_name=client_name,
                    project_name=project_name,
                    client_logo_bytes=client_logo_bytes,
                )
            except Exception as e:
                st.warning(f"Could not add header/footer: {str(e)}")

            # Save to buffer
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            # Clean up temporary logo files
            # No need to remove falcon_logo_path, logo is fixed from backend
            if client_logo_path and os.path.exists(client_logo_path):
                os.remove(client_logo_path)
            
            st.success("Document generated successfully!")
            
            st.download_button(
                label="📥 Download Complete Proposal Document",
                data=buffer,
                file_name=f"Falcon_Proposal_{client_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width='stretch'
            )
            
        except Exception as e:
            st.error(f"Error generating document: {str(e)}")
            st.exception(e)
            # Clean up temporary logo files in case of error
            try:
                # No need to remove falcon_logo_path, logo is fixed from backend
                if 'client_logo_path' in locals() and client_logo_path and os.path.exists(client_logo_path):
                    os.remove(client_logo_path)
            except:
                pass

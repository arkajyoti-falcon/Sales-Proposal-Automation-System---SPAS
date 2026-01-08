import re
from typing import Dict, List, Optional, Tuple

from proposal_context import ProposalContext

KEYWORDS = {
    "induct": ["induct", "feedline", "feed line", "induction point", "feed lines"],
    "telescopic": ["telescopic", "tbc"],
    "chutes": ["chute", "chutes", "gravity", "rejection", "collection", "dispersion", "bulk", "mini", "direct bagging", "output chute", "generic"],
    "pph": ["pph", "throughput", "parcels per hour"],
}

# Dimension units that indicate a number is NOT a count
DIMENSION_UNITS = re.compile(
    r'^\s*(mm|cm|m|m/s|ms|sec|s|seconds?|minutes?|hours?|%|kg|kn|kN|×|x|\*|cph|bph|sqm)\b',
    re.IGNORECASE
)

# Patterns that look like dimensions, not counts (e.g., "495 × 800", "1175 mm")
DIMENSION_PATTERNS = [
    re.compile(r'\d+\s*[×xX\*]\s*\d+'),  # 495 × 800
    re.compile(r'\d+\s*mm\b', re.IGNORECASE),  # 1175 mm
    re.compile(r'\d+\s*cm\b', re.IGNORECASE),  # 50 cm
    re.compile(r'\d+\s*m\b(?!/s)', re.IGNORECASE),  # 5 m (but not m/s)
    re.compile(r'\d+\.?\d*\s*m/s\b', re.IGNORECASE),  # 2.5 m/s
    re.compile(r'\d+\.?\d*\s*kn\b', re.IGNORECASE),  # force in kN
    re.compile(r'\d+\.?\d*\s*kg\b', re.IGNORECASE),  # weight
    re.compile(r'pitch[:\s]+\d+', re.IGNORECASE),  # pitch: 1175
    re.compile(r'belt\s+carrier\s+pitch[:\s]+\d+', re.IGNORECASE),  # belt carrier pitch: 1175
]


def _is_dimension_context(text: str, match_start: int, number: str) -> bool:
    """Check if a number appears in a dimension context (not a count).
    
    Only checks if THIS specific number is followed by a unit, not if other 
    numbers in the context have units.
    """
    # Check text immediately after the number for dimension units
    after_text = text[match_start + len(number):match_start + len(number) + 15]
    if DIMENSION_UNITS.match(after_text):
        return True
    
    # Check if this number is immediately preceded by dimension indicator within 10 chars
    before_start = max(0, match_start - 15)
    before_text = text[before_start:match_start].lower()
    if re.search(r'(pitch|width|height|length)\s*[:]?\s*$', before_text):
        return True
    
    # Check if this specific number is part of a dimension pattern (e.g., 495 × 800)
    # Look for "X × Y" or "X x Y" pattern where this number is X or Y
    context_start = max(0, match_start - 10)
    context_end = min(len(text), match_start + len(number) + 10)
    context = text[context_start:context_end]
    
    # Check for multiplication pattern
    if re.search(rf'\b{number}\s*[×xX\*]\s*\d+', context) or re.search(rf'\d+\s*[×xX\*]\s*{number}\b', context):
        return True
    
    return False


def _extract_feedline_counts(text: str) -> List[int]:
    if not text:
        return []
    nums = []
    pattern = r'(\d{1,5})\s*(nos\.?)?\s*(fully\s*)?(automatic\s*)?(feed|induct)[\s-]*lines?'
    for m in re.finditer(pattern, text, flags=re.IGNORECASE):
        num_str = m.group(1)
        if _is_dimension_context(text, m.start(1), num_str):
            continue
        val = int(num_str)
        if val > 10000:
            continue
        nums.append(val)
    return nums


def _extract_numbers(text: str, keywords: List[str], check_dimensions: bool = True) -> List[int]:
    """Extract numbers that are tightly coupled to provided keywords (no generic bullet fallback)."""
    if not text:
        return []
    nums: List[int] = []
    for kw in keywords:
        # keyword within ~3 words of the number (either before or after), allowing short phrases between
        pattern_kw_before = rf'\b{re.escape(kw)}\b[^\d]{{0,50}}?(\d{{1,5}})'
        pattern_kw_after = rf'(\d{{1,5}})[^\d]{{0,50}}?\b{re.escape(kw)}\b'
        for pat in [pattern_kw_before, pattern_kw_after]:
            for m in re.finditer(pat, text, flags=re.IGNORECASE):
                num_str = m.group(1)
                if check_dimensions and _is_dimension_context(text, m.start(1), num_str):
                    continue
                val = int(num_str)
                if val > 10000:
                    continue
                if val not in nums:
                    nums.append(val)
    return nums


def _extract_chute_counts(text: str) -> Tuple[int, Dict[str, int]]:
    """Extract total chutes and per-type breakdown from text.
    
    Returns (total_detected, {chute_type: count})
    """
    if not text:
        return 0, {}
    
    per_type = {}
    
    # Patterns for specific chute types
    chute_types = {
        'direct_bagging': [r'(\d+)\s*(?:nos\.?)?\s*direct[\s-]?bagging\s*chutes?'],
        'gravity': [r'(\d+)\s*(?:nos\.?)?\s*gravity\s*chutes?'],
        'rejection': [r'(\d+)\s*(?:nos\.?)?\s*rejection\s*chutes?'],
        'collection': [r'(\d+)\s*(?:nos\.?)?\s*collection\s*chutes?'],
        'dispersion': [r'(\d+)\s*(?:nos\.?)?\s*dispersion\s*chutes?'],
        'bulk': [r'(\d+)\s*(?:nos\.?)?\s*bulk\s*chutes?'],
        'mini': [r'(\d+)\s*(?:nos\.?)?\s*mini[\s-]?gravity\s*chutes?'],
        'generic': [r'(\d+)\s*(?:nos\.?)?\s*(?:generic|standard)\s*chutes?'],
        'output': [r'(\d+)\s*(?:nos\.?)?\s*output\s*chutes?'],
    }
    
    for chute_type, patterns in chute_types.items():
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                per_type[chute_type] = max(int(m) for m in matches)
    
    # Look for total chutes explicitly mentioned
    total_patterns = [
        r'total\s+(?:of\s+)?(\d+)\s*(?:nos\.?)?\s*chutes?',
        r'(\d+)\s*(?:nos\.?)?\s*total\s*chutes?',
        r'a\s+total\s+of\s+(\d+)\s*chutes?',
        r'total\s+output\s+chutes?\s*:?\s*(\d+)',
    ]
    
    explicit_total = None
    for pattern in total_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            explicit_total = int(match.group(1))
            break
    
    # Priority: explicit total > sum of per-type > max of generic numbers
    if explicit_total is not None:
        return explicit_total, per_type
    elif per_type:
        return sum(per_type.values()), per_type
    else:
        # Fallback: look for numbers near "chute" keyword (but not dimensions)
        fallback_nums = _extract_numbers(text, ["chute", "chutes", "output chute", "output chutes"], check_dimensions=True)
        if fallback_nums:
            return max(fallback_nums), {}
        return 0, {}


def audit_section_consistency(context: ProposalContext, cover: str, exec_summary: str, process_flow: str, system_desc: str) -> List[str]:
    violations: List[str] = []
    
    # Guard: if context is None, return empty violations
    if context is None:
        return violations

    def check_feedlines(texts: Dict[str, str]):
        meta = context.get("feedlines")
        if meta.value is None:
            return
        for sec_name, sec_text in texts.items():
            nums = _extract_feedline_counts(sec_text or "")
            if not nums:
                continue
            for n in nums:
                if n != meta.value:
                    violations.append(f"{sec_name} mentions Induct Lines {n} but context says {meta.value}")
                    break

    def check_pph(texts: Dict[str, str]):
        if context.pph is None:
            return
        for sec_name, sec_text in texts.items():
            nums = _extract_numbers(sec_text or "", KEYWORDS["pph"], check_dimensions=True)
            if nums and max(nums) != context.pph:
                violations.append(f"{sec_name} mentions PPH {max(nums)} but context is {context.pph}")

    texts = {
        "Cover Letter": cover,
        "Executive Summary": exec_summary,
        "Process Flow": process_flow,
        "System Description": system_desc,
    }

    check_feedlines(texts)

    # Telescopic conveyors check (retain keyword search but with filters)
    meta_tel = context.get("telescopic_belt_conveyors")
    if meta_tel.value is not None:
        for sec_name, sec_text in texts.items():
            nums = _extract_numbers(sec_text or "", KEYWORDS["telescopic"], check_dimensions=True)
            if nums and nums[0] != meta_tel.value:
                violations.append(f"{sec_name} mentions Telescopic Conveyors {nums[0]} but context says {meta_tel.value}")

    # Aggregate chutes as total if available
    chute_total_meta = context.get("total_chutes")
    chute_total = chute_total_meta.value or 0
    if chute_total:
        for sec_name, sec_text in texts.items():
            detected_total, per_type = _extract_chute_counts(sec_text or "")
            if detected_total > 0 and detected_total != chute_total:
                violations.append(f"{sec_name} mentions total chutes {detected_total} but context total is {chute_total}")

    # PPH check
    check_pph(texts)

    return violations


# ==================== SELF-TESTS ====================
def _run_self_tests():
    """Run self-tests for audit edge cases."""
    errors = []
    
    # Test 1: Belt carrier pitch should NOT produce chute count
    test_text_1 = "Belt carrier pitch: 1175 mm. The system has 51 output chutes."
    nums = _extract_numbers(test_text_1, ["chute", "chutes"], check_dimensions=True)
    if 1175 in nums:
        errors.append("FAIL: 1175 mm pitch was incorrectly extracted as chute count")
    if 51 not in nums:
        errors.append("FAIL: 51 chutes was not extracted")
    
    # Test 2: Chute summation
    test_text_2 = "There are 41 direct bagging chutes and 10 generic chutes."
    total, per_type = _extract_chute_counts(test_text_2)
    if total != 51:
        errors.append(f"FAIL: Expected total 51 chutes, got {total}")
    if per_type.get('direct_bagging') != 41:
        errors.append(f"FAIL: Expected 41 direct bagging, got {per_type.get('direct_bagging')}")
    
    # Test 3: Dimension patterns should be filtered
    test_text_3 = "Conveyor width: 495 × 800 mm"
    nums = _extract_numbers(test_text_3, ["conveyor"], check_dimensions=True)
    if 495 in nums or 800 in nums:
        errors.append("FAIL: Dimension 495 × 800 was incorrectly extracted")
    
    # Test 4: Explicit total should take priority
    test_text_4 = "The system has a total of 75 chutes including 41 direct bagging and 10 generic chutes."
    total, per_type = _extract_chute_counts(test_text_4)
    if total != 75:
        errors.append(f"FAIL: Expected explicit total 75, got {total}")
    
    if errors:
        for e in errors:
            print(e)
        return False
    else:
        print("All self-tests passed!")
        return True


if __name__ == "__main__":
    _run_self_tests()

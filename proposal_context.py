from dataclasses import dataclass, field
from typing import Optional, Dict, Tuple, Any

from proposal_facts import ProposalFacts, get_counts_source_of_truth


@dataclass
class CountMeta:
    value: Optional[int]
    source: str
    confirmed: bool


@dataclass
class ProposalContext:
    client_name: str = ""
    project_name: str = ""
    pph: Optional[int] = None
    cbs_type: str = ""

    # Flags
    has_manual_induct: bool = False

    # Counts (store meta per key)
    counts: Dict[str, CountMeta] = field(default_factory=dict)

    # Raw blocks for heuristics/reference
    raw_costing: Dict[str, Any] = field(default_factory=dict)
    raw_dxf: Dict[str, Any] = field(default_factory=dict)

    def get(self, key: str) -> CountMeta:
        return self.counts.get(key, CountMeta(None, "unknown", False))

    def counts_block_text(self) -> str:
        lines = [
            "HARD CONSTRAINTS:",
            "- Use these counts exactly if present.",
            "- If a value is Not specified, do not invent it.",
            "",
            f"Client: {self.client_name}",
            f"Project: {self.project_name}",
            f"PPH (Throughput): {self.pph if self.pph is not None else 'Not specified'}",
            f"CBS Type: {self.cbs_type or 'Not specified'}",
            f"Manual Induct: {'Present' if self.has_manual_induct else 'Not specified'}",
            "",
        ]
        # Total chutes headline
        total_meta = self.get("total_chutes")
        total_val = total_meta.value if total_meta.value is not None else "Not specified"
        lines.append(f"Total Chutes: {total_val}")

        # Breakdown shown only when values are confirmed
        for label, key in [
            ("Direct Bagging Chutes", "direct_bagging_chutes"),
            ("Generic Chutes", "generic_chutes"),
            ("Gravity Chutes", "gravity_chutes"),
            ("Mini Gravity Chutes", "mini_gravity_chutes"),
            ("Collection Chutes", "collection_chutes"),
            ("Rejection Chutes", "rejection_chutes"),
            ("Dispersion Chutes", "dispersion_chutes"),
            ("Bulk Chutes", "bulk_chutes"),
        ]:
            meta = self.get(key)
            if meta.value is not None:
                lines.append(f"{label}: {meta.value}")
        # Feedlines / telescopic after chutes for clarity
        for label, key in [
            ("Induct Lines (Automatic)", "feedlines"),
            ("Telescopic Conveyors", "telescopic_belt_conveyors"),
        ]:
            meta = self.get(key)
            val = meta.value if meta.value is not None else "Not specified"
            lines.append(f"{label}: {val}")
        return "\n".join(lines)


def _count_from_costing(costing: Dict[str, Any], key: str) -> Optional[int]:
    val = costing.get(key)
    try:
        if val not in (None, "", 0):
            return int(val)
    except Exception:
        return None
    return None


def _count_from_dxf_heuristics(dxf_json: Dict[str, Any], key: str) -> Optional[int]:
    if not dxf_json:
        return None
    # Heuristics mapping
    if key == "feedlines":
        cat = dxf_json.get("category_summary", {})
        v = cat.get("AUTO_INDUCT", 0)
        return v if v else None
    if key == "telescopic_belt_conveyors":
        tel = 0
        for name, count in dxf_json.get("categorized_components", {}).get("CONVEYOR_INFEED", {}).items():
            nm = str(name).lower()
            if "telescopic" in nm or "tbc" in nm:
                tel += int(count)
        return tel if tel else None
    # Chutes
    if key in {
        "gravity_chutes",
        "mini_gravity_chutes",
        "collection_chutes",
        "rejection_chutes",
        "dispersion_chutes",
        "bulk_chutes",
        "direct_bagging_chutes",
    }:
        br = dxf_json.get("chute_analysis", {}).get("breakdown", {})
        # Map DXF type names to keys (best-effort)
        mapping = {
            "gravity": "gravity_chutes",
            "mini": "mini_gravity_chutes",
            "collection": "collection_chutes",
            "rejection": "rejection_chutes",
            "dispersion": "dispersion_chutes",
            "bulk": "bulk_chutes",
            "direct bagging": "direct_bagging_chutes",
        }
        target_type = None
        for k, v in mapping.items():
            if v == key:
                target_type = k
                break
        if target_type is not None:
            for tname, tcount in br.items():
                if target_type in str(tname).lower():
                    return int(tcount) if int(tcount) > 0 else None
    # Throughput
    if key == "throughput_pph":
        v = dxf_json.get("pph", None)
        try:
            return int(v) if v not in (None, "", 0) else None
        except Exception:
            return None
    return None


def build_proposal_context(
    facts: Optional[ProposalFacts],
    costing_metrics: Optional[Dict[str, Any]],
    dxf_json: Optional[Dict[str, Any]],
) -> ProposalContext:
    ctx = ProposalContext()
    if facts:
        ctx.client_name = facts.client_name
        ctx.project_name = facts.project_name
        # CBS type & manual induct flag
        ctx.cbs_type = facts.layout_flags.cbs_type or ""
        ctx.has_manual_induct = bool(facts.layout_flags.has_manual_induct)
    ctx.raw_costing = (
        costing_metrics
        if costing_metrics is not None
        else (facts.costing_metrics if (facts and getattr(facts, "costing_metrics", None)) else {})
    )
    ctx.raw_dxf = dxf_json or {}

    def resolve(key: str) -> CountMeta:
        # Priority 1: ProposalFacts
        if facts:
            val, src, conf = get_counts_source_of_truth(facts, key)
            if val is not None:
                return CountMeta(val, src, conf)
        # Priority 2: costing explicit
        cval = _count_from_costing(ctx.raw_costing, key)
        if cval is not None:
            return CountMeta(cval, "costing_explicit", True)
        # Priority 3: DXF heuristics
        dval = _count_from_dxf_heuristics(ctx.raw_dxf, key)
        if dval is not None:
            return CountMeta(dval, "dxf_heuristics", False)
        return CountMeta(None, "not_specified", False)

    for k in [
        "feedlines",
        "telescopic_belt_conveyors",
        "gravity_chutes",
        "mini_gravity_chutes",
        "collection_chutes",
        "rejection_chutes",
        "dispersion_chutes",
        "bulk_chutes",
        "direct_bagging_chutes",
        "generic_chutes",
        "total_chutes",
        "throughput_pph",
    ]:
        ctx.counts[k] = resolve(k)

    # Compute total chutes explicitly as sum of confirmed chute types (including generic)
    chute_keys = [
        "gravity_chutes",
        "mini_gravity_chutes",
        "collection_chutes",
        "rejection_chutes",
        "dispersion_chutes",
        "bulk_chutes",
        "direct_bagging_chutes",
        "generic_chutes",
    ]
    total_val = sum([(ctx.get(k).value or 0) for k in chute_keys])
    ctx.counts["total_chutes"] = CountMeta(total_val if total_val else None, "computed_sum", total_val > 0)

    # PPH and CBS type from facts/context
    if facts:
        pph_meta = resolve("throughput_pph")
        ctx.pph = pph_meta.value
        if not ctx.cbs_type:
            ctx.cbs_type = facts.layout_flags.cbs_type or ""
    else:
        ctx.pph = _count_from_costing(ctx.raw_costing, "throughput_pph") or _count_from_dxf_heuristics(ctx.raw_dxf, "throughput_pph")

    return ctx

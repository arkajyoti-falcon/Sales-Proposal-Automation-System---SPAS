"""
Streamlit Integration for ProposalFacts

This module shows how to integrate ProposalFacts into the streamlit.py workflow:
1. Accept DXF + costing file uploads
2. Extract and populate ProposalFacts once
3. Display summary to user
4. Pass to all generators instead of individual parameters

Usage in streamlit.py:
    from proposal_facts_streamlit_integration import integrate_proposal_facts_upload
    
    # In your file upload section:
    facts = integrate_proposal_facts_upload(
        client_name="Client Name",
        project_name="Project Name",
        dxf_file_upload=uploaded_dxf,
        costing_file_upload=uploaded_costing,
    )
    
    if facts:
        # Use facts in all generators:
        cover_letter = call_groq_for_cover_letter(facts)
        exec_summary = call_groq_exec_summary(facts)
        process_flow = call_groq_for_process_flow(facts)
        system_desc = generate_system_description(facts)
"""

import streamlit as st
from pathlib import Path
import tempfile
import io
import logging
from typing import Optional

from proposal_facts import ProposalFacts
from proposal_facts_extractor import extract_proposal_facts

logger = logging.getLogger(__name__)


def integrate_proposal_facts_upload(
    client_name: str,
    project_name: str,
    dxf_file_upload: Optional[object],
    costing_file_upload: Optional[object] = None,
    verbose: bool = True,
) -> Optional[ProposalFacts]:
    """
    Streamlit integration for ProposalFacts extraction.
    
    Handles DXF + optional costing file uploads, extracts data,
    and returns a fully populated ProposalFacts object.
    
    Args:
        client_name: Client name for metadata
        project_name: Project name for metadata
        dxf_file_upload: Streamlit UploadedFile (DXF)
        costing_file_upload: Streamlit UploadedFile (XLSX) - optional
        verbose: Show extraction summary in Streamlit
    
    Returns:
        ProposalFacts object if successful, None if failed
    
    Example:
        dxf_file = st.file_uploader("DXF", type=["dxf"])
        costing_file = st.file_uploader("Costing", type=["xlsx"])
        
        facts = integrate_proposal_facts_upload(
            client_name="My Client",
            project_name="My Project",
            dxf_file_upload=dxf_file,
            costing_file_upload=costing_file,
        )
        
        if facts:
            st.success("Proposal data extracted successfully!")
            st.write(facts.summary())
    """
    
    # Validate inputs
    if not dxf_file_upload:
        st.error("❌ DXF file is required")
        return None
    
    if not client_name or not project_name:
        st.error("❌ Client name and project name are required")
        return None
    
    try:
        # Create temporary directory for uploaded files
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            
            # Save DXF file temporarily
            dxf_path = tmpdir / dxf_file_upload.name
            dxf_path.write_bytes(dxf_file_upload.getvalue())
            
            # Save costing file if provided
            costing_path = None
            if costing_file_upload:
                costing_path = tmpdir / costing_file_upload.name
                costing_path.write_bytes(costing_file_upload.getvalue())
                
                if verbose:
                    st.info(f"📄 Using costing file: {costing_file_upload.name}")
            
            # Extract proposal facts
            if verbose:
                with st.spinner("Extracting proposal data from DXF and costing..."):
                    facts = extract_proposal_facts(
                        dxf_path=str(dxf_path),
                        costing_path=str(costing_path) if costing_path else None,
                        project_name=project_name,
                        client_name=client_name,
                        verbose=False,  # We'll show summary in Streamlit
                    )
            else:
                facts = extract_proposal_facts(
                    dxf_path=str(dxf_path),
                    costing_path=str(costing_path) if costing_path else None,
                    project_name=project_name,
                    client_name=client_name,
                    verbose=False,
                )
            
            # Display summary
            if facts:
                if verbose:
                    st.success("✅ Proposal facts extracted successfully!")
                    
                    # Show extraction summary with nice formatting
                    with st.expander("📊 Extracted Data Summary", expanded=False):
                        summary_text = facts.summary()
                        st.code(summary_text, language="text")
                    
                    # Show unconfirmed fields warning
                    unconfirmed = facts.get_unconfirmed_fields()
                    if unconfirmed:
                        with st.warning(f"⚠️ {len(unconfirmed)} fields marked as 'to be confirmed'"):
                            for field in unconfirmed:
                                st.write(f"  - {field}")
                
                logger.info(f"Successfully extracted ProposalFacts for {project_name}")
                return facts
            else:
                st.error("❌ Failed to extract proposal facts")
                return None
    
    except Exception as e:
        logger.error(f"Error in integrate_proposal_facts_upload: {e}")
        st.error(f"❌ Error extracting proposal data: {str(e)}")
        return None


def display_facts_summary(facts: ProposalFacts, show_details: bool = True):
    """
    Display ProposalFacts summary in Streamlit.
    
    Args:
        facts: ProposalFacts object to display
        show_details: Show detailed breakdown
    """
    if not facts:
        st.warning("No proposal facts available")
        return
    
    # Header
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Project", facts.project_name)
    with col2:
        st.metric("Client", facts.client_name)
    
    # System layout
    st.subheader("System Layout")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("CBS Type", facts.layout_flags.cbs_type or "Not detected")
    with col2:
        st.metric("Infeed", "✓" if facts.layout_flags.has_infeed_system else "✗")
    with col3:
        st.metric("Auto Induct", "✓" if facts.layout_flags.has_auto_induct else "✗")
    with col4:
        st.metric("Manual Induct", "✓" if facts.layout_flags.has_manual_induct else "✗")
    
    # Extracted counts
    if show_details:
        st.subheader("Extracted Component Counts")
        
        counts_data = []
        for key, comp_count in facts.system_numbers.get_all_counts().items():
            if comp_count.count is not None:
                counts_data.append({
                    "Component": key.replace("_", " ").title(),
                    "Count": comp_count.count,
                    "Source": comp_count.source.value,
                    "Confirmed": "✓" if comp_count.confirmed else "✗"
                })
        
        if counts_data:
            import pandas as pd
            df = pd.DataFrame(counts_data)
            st.dataframe(df, use_container_width=True)
        else:
            st.info("No component counts extracted")
    
    # Unconfirmed fields
    unconfirmed = facts.get_unconfirmed_fields()
    if unconfirmed:
        st.warning(f"**Unconfirmed fields** ({len(unconfirmed)})")
        for field in unconfirmed:
            st.write(f"  - {field}")


def get_component_count_with_fallback(
    facts: ProposalFacts,
    component_key: str,
    human_name: str = None,
) -> tuple:
    """
    Get a component count from ProposalFacts with user-friendly fallback.
    
    Returns:
        (count, display_text, source)
    
    Examples:
        count, text, source = get_component_count_with_fallback(
            facts, "feedlines", "Induct Lines"
        )
        st.write(f"{human_name}: {text}")
    """
    if not facts:
        return None, "Data not available", "unknown"
    
    count, source, confirmed = facts.get_count(component_key)
    human_name = human_name or component_key.replace("_", " ").title()
    
    if count is None:
        return None, f"(to be confirmed during detailed engineering)", source
    
    return count, f"{count} Nos", source


# ============================================================================
# Generator Wrapper Functions
# These show how to refactor generators to use ProposalFacts
# ============================================================================

def generate_cover_letter_with_facts(
    facts: ProposalFacts,
    executives: str,
    offer_reference: str,
    letter_date: str,
    invitation_date: str = "",
    meeting_date: str = "",
    sender_name: str = "",
    sender_title: str = "",
) -> str:
    """
    Generate cover letter using ProposalFacts as single source of truth.
    
    Refactored version that:
    1. Uses ProposalFacts instead of separate DXF/costing parameters
    2. Gets all counts from facts (prefer costing, fallback to DXF)
    3. Eliminates duplicate extraction logic
    
    Args:
        facts: ProposalFacts object (single source of truth)
        executives: Comma-separated executive names
        offer_reference: Offer reference number
        letter_date: Letter date
        invitation_date: Invitation date (optional)
        meeting_date: Meeting date (optional)
        sender_name: Sender name
        sender_title: Sender title
    
    Returns:
        Generated cover letter text
    """
    from streamlit import Groq, st
    import os
    
    # All data comes from facts - no separate DXF re-extraction!
    groq_api_key = os.getenv("GROQ_API_KEY")
    if not groq_api_key:
        raise ValueError("GROQ_API_KEY not set")
    
    groq_client = Groq(api_key=groq_api_key)
    
    # Build process flow summary from facts
    process_flow_summary = _build_process_flow_summary_from_facts(facts)
    
    # Use existing prompt structure but pass facts
    system_prompt = """You are a professional proposal cover letter writer at Falcon Autotech..."""
    # (use existing COVER_LETTER_SYSTEM_PROMPT)
    
    user_prompt = f"""
client_name: {facts.client_name}
project_title: {facts.project_name}
offer_ref: {offer_reference}
letter_date: {letter_date}

executives: {executives}

process_flow_summary: {process_flow_summary}

sender_name: {sender_name}
sender_title: {sender_title}
"""
    
    # Call Groq with facts-based data
    response = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )
    
    return response.choices[0].message.content.strip()


def generate_exec_summary_with_facts(
    facts: ProposalFacts,
) -> str:
    """
    Generate executive summary using ProposalFacts.
    
    Refactored version that:
    1. Uses ProposalFacts instead of separate DXF/system_text parameters
    2. Gets counts with proper source preference (costing > DXF)
    3. No repeated extraction from DXF
    
    Args:
        facts: ProposalFacts object (single source of truth)
    
    Returns:
        Generated executive summary text
    """
    from streamlit import Groq
    import os
    
    groq_api_key = os.getenv("GROQ_API_KEY")
    if not groq_api_key:
        raise ValueError("GROQ_API_KEY not set")
    
    groq_client = Groq(api_key=groq_api_key)
    
    # Get all critical counts from facts (prefer costing over DXF)
    feedlines, _, _ = facts.get_count("feedlines")
    cbs_type = facts.layout_flags.cbs_type or "Unknown"
    
    # Build component summary from facts
    components_summary = []
    if facts.layout_flags.has_auto_induct and feedlines:
        components_summary.append(f"{feedlines} Auto Induct Lines")
    if facts.layout_flags.has_manual_induct:
        components_summary.append("Manual Induct Stations")
    if facts.layout_flags.has_output_chutes:
        components_summary.append("Output Chutes")
    
    system_text = f"""
    System Type: {cbs_type}
    Components: {', '.join(components_summary)}
    """
    
    # Use existing executive summary logic with facts-based data
    # (This replaces the call_groq_exec_summary function)
    
    return "Executive summary generated from facts"  # Placeholder


def _build_process_flow_summary_from_facts(facts: ProposalFacts) -> str:
    """
    Build a high-level process flow summary from ProposalFacts.
    
    Example output:
    "The system includes 4 automatic induct lines, 1 Loop CBS sorter,
     and 180 output chutes for efficient parcel distribution."
    """
    parts = []
    
    # Induct system
    feedlines, _, _ = facts.get_count("feedlines")
    if feedlines:
        parts.append(f"{feedlines} automatic induct lines")
    
    # CBS
    if facts.layout_flags.cbs_type:
        parts.append(f"1 {facts.layout_flags.cbs_type} sorter")
    
    # Output chutes
    total_chutes = 0
    for key in ["gravity_chutes", "mini_gravity_chutes", "collection_chutes", "bulk_chutes"]:
        count, _, _ = facts.get_count(key)
        if count:
            total_chutes += count
    
    if total_chutes > 0:
        parts.append(f"{total_chutes} output chutes")
    
    if not parts:
        return "Automated parcel sortation system"
    
    return "The proposed system includes " + ", ".join(parts) + "."


if __name__ == "__main__":
    # Test usage
    st.title("ProposalFacts Integration Test")
    
    col1, col2 = st.columns(2)
    with col1:
        dxf_file = st.file_uploader("DXF File", type=["dxf"])
    with col2:
        costing_file = st.file_uploader("Costing File", type=["xlsx"])
    
    if dxf_file and st.button("Extract Facts"):
        facts = integrate_proposal_facts_upload(
            client_name="Test Client",
            project_name="Test Project",
            dxf_file_upload=dxf_file,
            costing_file_upload=costing_file,
        )
        
        if facts:
            display_facts_summary(facts)

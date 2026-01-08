#!/usr/bin/env python3
"""
Test script to debug Mermaid flowchart generation and rendering.
"""

import os
import io
import zlib
import base64
import re
from typing import Tuple
from dotenv import load_dotenv
from groq import Groq
import requests

# Load environment
load_dotenv()

# Test process flow text
PROCESS_FLOW = """
1. Infeed System: Parcels arrive at the infeed conveyor
2. VDS (Volume Distribution System): Measures and sorts parcels
3. Inducts: Multiple induct stations load parcels onto the sorter
4. Loop CBS (Cross Belt Sorter): Main sorting system with carriers
5. Output Chutes: Parcels are diverted to different chutes
6. PTL (Pick to Light): Manual sorting with light indicators
7. Rejection Chute: Non-sortable items are rejected
"""

def load_groq():
    """Load Groq client from environment variable."""
    key = (os.getenv("GROQ_API_KEY") or "").strip()
    if not key:
        return None
    try:
        return Groq(api_key=key)
    except Exception:
        return None

def sanitize_mermaid_for_render(code: str) -> str:
    """Sanitize Mermaid code for rendering."""
    if not code:
        return code
    s = code.replace("\r\n", "\n").replace("\r", "\n")

    def _fix(s: str, left: str, right: str) -> str:
        pat = re.compile(re.escape(left) + r"(.*?)" + re.escape(right), re.DOTALL)
        def sub(m):
            inner = m.group(1)
            inner = inner.replace("\\n", "<br/>").replace("\n", "<br/>")
            inner = re.sub(r"\s{2,}", " ", inner).strip()
            return f"{left}{inner}{right}"
        return pat.sub(sub, s)

    s = _fix(s, "[[", "]]")
    s = _fix(s, "((", "))")
    s = _fix(s, "[/", "/]")
    s = _fix(s, "[", "]")
    s = _fix(s, "(", ")")
    s = _fix(s, "{", "}")
    return s

def call_groq_for_mermaid(process_flow_text: str) -> str:
    """Call GROQ to generate Mermaid code."""
    client = load_groq()
    if not client:
        raise RuntimeError("GROQ_API_KEY is not set")

    system_prompt = """
You are a senior solution engineer and diagram expert.

Your task:
- Convert a warehouse/process "Process Flow" description into a Mermaid flowchart.

STRICT RULES FOR MERMAID v11+ COMPATIBILITY:
- Output ONLY Mermaid code, no backticks, no explanations.
- Use 'flowchart TD' at the top (top-down layout).
- Use simple alphanumeric node IDs (A, B, C, A1, B1, etc.) - NO special characters in IDs.
- Use brackets [] for rectangular nodes, () for rounded, {} for diamond decisions.
- Keep labels SHORT and SIMPLE - avoid special characters, quotes, or complex text.
- Use --> for arrows (no text on arrows unless absolutely necessary).
- NO semicolons, NO colons in labels, NO quotes in labels.
- Capture branching where appropriate (e.g., multiple chute types).
- Keep the diagram reasonably compact (avoid exploding into too many tiny nodes).
"""

    user_prompt = f"""
SOLUTION_TEXT:
{process_flow_text}

Return only Mermaid code (start with: flowchart TD).

CRITICAL REQUIREMENTS FOR MERMAID v11:
- Use ONLY simple node IDs: A, B, C, D, E, F, G, etc. (single letters or A1, B1, etc.)
- Node labels must be simple text in brackets: A[Infeed System]
- NO special characters in node IDs (no hyphens, underscores, numbers alone)
- NO quotes, colons, or semicolons in labels
- Keep labels concise (2-4 words max per node)
- Represent main steps as nodes (Infeed System, VDS, Inducts, Loop CBS, Output Chutes, PTL, etc.)
- Show major branches as separate nodes
- Connect nodes with simple arrows: A --> B
- Do NOT wrap in ```mermaid``` - just raw code starting with "flowchart TD"

Example format:
flowchart TD
    A[Infeed System]
    B[VDS]
    C[Inducts]
    A --> B
    B --> C
"""

    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": system_prompt.strip()},
            {"role": "user", "content": user_prompt.strip()},
        ],
        temperature=0.1,
        max_tokens=1200,
    )
    
    mermaid_raw = (resp.choices[0].message.content or "").strip()
    
    # Clean using regex
    mermaid_raw = re.sub(r"^```(?:mermaid)?\s*", "", mermaid_raw)
    mermaid_raw = re.sub(r"\s*```$", "", mermaid_raw)
    
    # Ensure it starts with flowchart
    if not mermaid_raw.lower().startswith("flowchart"):
        m = re.search(r"(flowchart\s+TD[\s\S]+)$", mermaid_raw, re.IGNORECASE)
        if m:
            mermaid_raw = m.group(1).strip()
    
    return mermaid_raw

def _deflate_b64_urlsafe(data: str) -> str:
    """Compress and encode for mermaid.ink."""
    compressed = zlib.compress(data.encode("utf-8"))
    return base64.urlsafe_b64encode(compressed).decode("ascii")

def test_render(mermaid_code: str):
    """Test rendering the Mermaid code."""
    print("\n" + "="*80)
    print("TESTING MERMAID RENDERING")
    print("="*80)
    
    # Sanitize
    sanitized = sanitize_mermaid_for_render(mermaid_code)
    print(f"\n✓ Sanitized code:\n{sanitized}\n")
    
    # Try Kroki PNG
    print("\n[1/2] Testing Kroki PNG...")
    try:
        resp = requests.post(
            "https://kroki.io/mermaid/png",
            data=sanitized.encode("utf-8"),
            headers={"Content-Type": "text/plain"},
            timeout=25,
        )
        print(f"Kroki status: {resp.status_code}")
        if resp.ok and resp.content:
            print(f"✓ Kroki SUCCESS: {len(resp.content)} bytes")
            with open("test_flowchart_kroki.png", "wb") as f:
                f.write(resp.content)
            print("✓ Saved to: test_flowchart_kroki.png")
            return True
        else:
            print(f"✗ Kroki FAILED: {resp.text[:500]}")
    except Exception as e:
        print(f"✗ Kroki ERROR: {e}")
    
    # Try mermaid.ink PNG
    print("\n[2/2] Testing mermaid.ink PNG...")
    try:
        payload = _deflate_b64_urlsafe(sanitized)
        url = f"https://mermaid.ink/img/{payload}"
        print(f"URL: {url}")
        resp = requests.get(url, timeout=25, headers={"Accept": "image/png"})
        print(f"mermaid.ink status: {resp.status_code}")
        if resp.ok and resp.content:
            print(f"✓ mermaid.ink SUCCESS: {len(resp.content)} bytes")
            with open("test_flowchart_mermaid_ink.png", "wb") as f:
                f.write(resp.content)
            print("✓ Saved to: test_flowchart_mermaid_ink.png")
            return True
        else:
            print(f"✗ mermaid.ink FAILED: {resp.text[:500]}")
    except Exception as e:
        print(f"✗ mermaid.ink ERROR: {e}")
    
    return False

if __name__ == "__main__":
    print("="*80)
    print("MERMAID FLOWCHART GENERATION TEST")
    print("="*80)
    
    # Step 1: Generate Mermaid code
    print("\nStep 1: Generating Mermaid code via GROQ...")
    try:
        mermaid_code = call_groq_for_mermaid(PROCESS_FLOW)
        print(f"\n✓ Generated Mermaid code:\n{mermaid_code}\n")
    except Exception as e:
        print(f"\n✗ GROQ FAILED: {e}")
        exit(1)
    
    # Step 2: Test rendering
    success = test_render(mermaid_code)
    
    if success:
        print("\n" + "="*80)
        print("✓ TEST PASSED - Flowchart generated successfully!")
        print("="*80)
    else:
        print("\n" + "="*80)
        print("✗ TEST FAILED - Could not render flowchart")
        print("="*80)
        exit(1)

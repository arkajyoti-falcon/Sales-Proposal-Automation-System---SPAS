#!/usr/bin/env python3
"""
Print DXF Raw Information to Console/CMD
This script extracts and displays raw DXF component information
"""

import sys
import os

# Set UTF-8 encoding for terminal output
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None

from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from dxf_extractor import extract_dxf_components, create_dxf_summary_verbose
import json

def print_dxf_raw_info(dxf_path):
    """Extract and print DXF component information"""
    dxf_file = Path(dxf_path)
    
    if not dxf_file.exists():
        print("[ERROR] DXF file not found:", dxf_path)
        return
    
    print("=" * 80)
    print("DXF RAW INFORMATION EXTRACTOR")
    print("File:", dxf_file.name)
    print("Path:", dxf_file.absolute())
    print("=" * 80)
    print()
    
    try:
        # Extract DXF components
        dxf_json = extract_dxf_components(dxf_file)
        
        print("[SUCCESS] DXF components extracted successfully")
        print()
        
        # Print basic statistics
        print("BASIC STATISTICS:")
        print("-" * 80)
        print("  Total Components:", dxf_json.get('total_components', 0))
        print("  Has Scanner:", dxf_json.get('has_scanner', False))
        print("  CBS Type:", dxf_json.get('cbs_type', 'Unknown'))
        print()
        
        # Print category summary
        print("CATEGORY SUMMARY:")
        print("-" * 80)
        cat_summary = dxf_json.get('category_summary', {})
        for cat, count in sorted(cat_summary.items(), key=lambda x: -x[1]):
            print("  {}: {}".format(cat, count))
        print()
        
        # Print chute analysis
        print("CHUTE ANALYSIS:")
        print("-" * 80)
        chute_analysis = dxf_json.get('chute_analysis', {})
        print("  Total Chutes:", chute_analysis.get('total', 0))
        print("  By Type:")
        for chute_type, count in sorted(chute_analysis.get('by_type', {}).items(), key=lambda x: -x[1]):
            print("    - {}: {}".format(chute_type, count))
        print()
        
        # Print raw block counts
        print("RAW BLOCK COUNTS (Top 30):")
        print("-" * 80)
        raw_counts = dxf_json.get('raw_block_counts', {})
        for i, (block_name, count) in enumerate(sorted(raw_counts.items(), key=lambda x: -x[1])[:30]):
            print("  {}: {}".format(block_name, count))
        print()
        
        # Print categorized components
        print("CATEGORIZED COMPONENTS:")
        print("-" * 80)
        categorized = dxf_json.get('categorized_components', {})
        for category, components in sorted(categorized.items()):
            print()
            print("  {}:".format(category))
            for comp_name, comp_count in sorted(components.items(), key=lambda x: -x[1])[:10]:
                print("    - {}: {}".format(comp_name, comp_count))
        print()
        
        # Print full verbose summary
        print("FULL VERBOSE SUMMARY:")
        print("-" * 80)
        verbose_summary = create_dxf_summary_verbose(dxf_json)
        print(verbose_summary)
        print()
        
        # Save to JSON file for reference
        output_json = dxf_file.parent / "{}_extracted.json".format(dxf_file.stem)
        with open(output_json, 'w') as f:
            json.dump(dxf_json, f, indent=2)
        print("[SUCCESS] Extracted JSON saved to:", output_json)
        print()
        
    except Exception as e:
        print("[ERROR] Error processing DXF file:", e)
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python print_dxf_info.py <path_to_dxf_file>")
        print()
        print("Example:")
        print('  python print_dxf_info.py "OUTPUT/DXF Upload/2111161009_0_BOSTA CAIRO_Loop CBS_12K(Rev11)_20260108_185927.dxf"')
        sys.exit(1)
    
    dxf_path = sys.argv[1]
    print_dxf_raw_info(dxf_path)

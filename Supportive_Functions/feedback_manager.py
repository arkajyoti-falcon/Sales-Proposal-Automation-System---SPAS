"""
Feedback Management Module for Process Flow Generation

This module provides utilities for:
1. Loading and saving feedback rules from JSON (GLOBAL - persistent across sessions)
2. Managing local/session rules (LOCAL - project-specific, temporary)
3. Applying feedback rules to process flow generation
4. Managing feedback lifecycle (add, edit, deactivate)

RULE TYPES:
- GLOBAL: Saved in feedback_rules.json, persistent, applied to all projects
- LOCAL: Stored in session state, temporary, project-specific only
- Future: Will support SQL database for enterprise deployment
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional


# Global rules file - persistent across sessions
FEEDBACK_FILE = Path(__file__).parent.parent / "feedback_rules.json"


def load_feedback_rules(rule_type: str = "global") -> List[Dict]:
    """
    Load active feedback rules.
    
    Args:
        rule_type: "global" (from JSON file) or "local" (from session, handled in streamlit)
    
    Returns:
        List of active rules
    """
    if rule_type != "global":
        # Local rules are managed in Streamlit session state
        return []
    
    if not FEEDBACK_FILE.exists():
        return []
    
    try:
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Return only active global rules
        rules = data.get("process_flow_feedback_rules", [])
        return [r for r in rules if r.get("active", True) and r.get("type", "global") == "global"]
    except Exception as e:
        print(f"Error loading feedback rules: {e}")
        return []


def save_feedback_rule(user_feedback: str, extracted_rule: str, project_name: str, rule_type: str = "global") -> str:
    """
    Save a new feedback rule.
    
    Args:
        user_feedback: Original feedback text from user
        extracted_rule: AI-extracted generalized rule
        project_name: Name of the project this rule came from
        rule_type: "global" (persistent, saved to JSON) or "local" (session only)
    
    Returns:
        New rule ID
    """
    if not FEEDBACK_FILE.exists():
        # Initialize with empty structure
        data = {
            "process_flow_feedback_rules": [],
            "metadata": {
                "last_updated": datetime.now().isoformat(),
                "total_rules": 0,
                "total_global_rules": 0,
                "version": "2.0"
            }
        }
    else:
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    
    # Generate new rule ID
    existing_rules = data.get("process_flow_feedback_rules", [])
    rule_id = f"{rule_type[0]}rule_{len(existing_rules) + 1:03d}"  # grule_001 or lrule_001
    
    # Create new rule
    new_rule = {
        "id": rule_id,
        "type": rule_type,  # "global" or "local"
        "timestamp": datetime.now().isoformat(),
        "user_feedback": user_feedback,
        "rule": extracted_rule,
        "applied_to_projects": [project_name],
        "active": True
    }
    
    # Only save global rules to JSON (local rules stay in session state)
    if rule_type == "global":
        # Add to rules list
        data["process_flow_feedback_rules"].append(new_rule)
        
        # Update metadata
        data["metadata"]["last_updated"] = datetime.now().isoformat()
        data["metadata"]["total_rules"] = len(data["process_flow_feedback_rules"])
        data["metadata"]["total_global_rules"] = len([r for r in data["process_flow_feedback_rules"] if r.get("type", "global") == "global"])
        
        # Save back to file
        with open(FEEDBACK_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    return rule_id


def get_feedback_rules_as_text(local_rules: List[Dict] = None) -> str:
    """
    Get all active feedback rules formatted as text for prompt injection.
    Combines GLOBAL rules (from JSON) and LOCAL rules (from session).
    
    Args:
        local_rules: Optional list of local/session rules to include
    
    Returns:
        Formatted text of all rules for AI prompt
    """
    # Load global rules
    global_rules = load_feedback_rules(rule_type="global")
    
    # Combine with local rules if provided
    all_rules = []
    
    if global_rules:
        all_rules.extend(global_rules)
    
    if local_rules:
        all_rules.extend([r for r in local_rules if r.get("active", True)])
    
    if not all_rules:
        return ""
    
    rules_text = "IMPORTANT: Apply these learned feedback rules:\n"
    rules_text += "GLOBAL RULES (organization-wide standards):\n" if global_rules else ""
    
    global_count = 0
    for idx, rule in enumerate(all_rules, 1):
        rule_type = rule.get("type", "global")
        if rule_type == "global" and global_rules:
            global_count += 1
            rules_text += f"  {global_count}. {rule['rule']}\n"
    
    if local_rules and any(r.get("type") == "local" for r in all_rules):
        rules_text += "\nLOCAL RULES (project-specific for this session):\n"
        local_count = 0
        for rule in all_rules:
            if rule.get("type") == "local":
                local_count += 1
                rules_text += f"  {local_count}. {rule['rule']}\n"
    
    return rules_text


def deactivate_rule(rule_id: str) -> bool:
    """Deactivate a specific rule by ID."""
    if not FEEDBACK_FILE.exists():
        return False
    
    try:
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        rules = data.get("process_flow_feedback_rules", [])
        for rule in rules:
            if rule.get("id") == rule_id:
                rule["active"] = False
                
                # Update metadata
                data["metadata"]["last_updated"] = datetime.now().isoformat()
                
                # Save back
                with open(FEEDBACK_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                
                return True
        
        return False
    except Exception as e:
        print(f"Error deactivating rule: {e}")
        return False


def get_all_rules() -> List[Dict]:
    """Get all rules (including inactive ones) for management UI."""
    if not FEEDBACK_FILE.exists():
        return []
    
    try:
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("process_flow_feedback_rules", [])
    except Exception as e:
        print(f"Error loading all rules: {e}")
        return []


def update_rule(rule_id: str, new_rule_text: str) -> bool:
    """Update the rule text for an existing rule (global rules only)."""
    if not FEEDBACK_FILE.exists():
        return False
    
    try:
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        rules = data.get("process_flow_feedback_rules", [])
        for rule in rules:
            if rule.get("id") == rule_id:
                rule["rule"] = new_rule_text
                rule["timestamp"] = datetime.now().isoformat()
                
                # Update metadata
                data["metadata"]["last_updated"] = datetime.now().isoformat()
                
                # Save back
                with open(FEEDBACK_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                
                return True
        
        return False
    except Exception as e:
        print(f"Error updating rule: {e}")
        return False


def create_local_rule(user_feedback: str, extracted_rule: str, project_name: str) -> Dict:
    """
    Create a local (session-only) rule without saving to JSON.
    Returns the rule dict to be stored in session state.
    
    Args:
        user_feedback: Original feedback text
        extracted_rule: AI-extracted rule
        project_name: Project name
    
    Returns:
        Rule dictionary for session state
    """
    import random
    rule_id = f"lrule_{random.randint(1000, 9999)}"
    
    return {
        "id": rule_id,
        "type": "local",
        "timestamp": datetime.now().isoformat(),
        "user_feedback": user_feedback,
        "rule": extracted_rule,
        "applied_to_projects": [project_name],
        "active": True
    }


def get_rule_stats() -> Dict:
    """
    Get statistics about rules.
    
    Returns:
        Dictionary with stats: total, active, inactive, global, local
    """
    try:
        if not FEEDBACK_FILE.exists():
            return {
                "total": 0,
                "active": 0,
                "inactive": 0,
                "global": 0,
                "local": 0
            }
        
        with open(FEEDBACK_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        rules = data.get("process_flow_feedback_rules", [])
        
        active_count = sum(1 for r in rules if r.get("active", True))
        global_count = sum(1 for r in rules if r.get("type", "global") == "global")
        
        return {
            "total": len(rules),
            "active": active_count,
            "inactive": len(rules) - active_count,
            "global": global_count,
            "local": len(rules) - global_count
        }
    except Exception as e:
        print(f"Error getting stats: {e}")
        return {"total": 0, "active": 0, "inactive": 0, "global": 0, "local": 0}

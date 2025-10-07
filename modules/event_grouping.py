"""
Event Grouping Module

Groups visits into event categories and creates visit windows with offset calculations
from protocol JSON files.
"""

import json
import re
import logging
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple


def load_json(path: str) -> Dict[str, Any]:
    """Load JSON file."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def find_all_soa_tables(node: Dict[str, Any], soa_tables: List[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Recursively find SOA tables in the JSON structure."""
    if soa_tables is None:
        soa_tables = []
    
    if isinstance(node, dict):
        if node.get("name", "").startswith("Table") and "children" in node:
            for row in node["children"]:
                if "children" not in row or not row["children"]:
                    continue
                first_cell = row["children"][0]
                if "children" in first_cell and first_cell["children"]:
                    for p in first_cell["children"]:
                        if p.get("text") and "Procedure" in p["text"]:
                            soa_tables.append(node)
                            break
        for child in node.get("children", []):
            find_all_soa_tables(child, soa_tables)
    elif isinstance(node, list):
        for item in node:
            find_all_soa_tables(item, soa_tables)
    
    return soa_tables


def normalize_visit_name(v: str, config: Dict[str, Any]) -> Optional[str]:
    """Normalize visit names according to configuration."""
    pattern = config.get('visit_normalization', {}).get('pattern', r'^([VP]\d+)(?:\s([a-zA-Z]+))?$')
    keep_suffix_length = config.get('visit_normalization', {}).get('keep_suffix_length', 1)
    special_cases = config.get('visit_normalization', {}).get('special_cases', [])
    
    m = re.match(pattern, v.strip())
    if not m:
        # Keep specific names if they're in special cases
        if v.strip().upper() in special_cases:
            return v.strip().upper()
        return None
    
    base, suffix = m.groups()
    
    # Handle special cases
    if base.upper() in special_cases:
        return base.upper()
    
    if suffix:
        if len(suffix) == keep_suffix_length:  # single-letter suffix → normalize
            return base
        else:  # multi-letter suffix → remove completely
            return None
    
    return base


def extract_visits_and_weeks(tables: List[Dict[str, Any]], config: Dict[str, Any]) -> pd.DataFrame:
    """Extract visits and study weeks from SOA tables."""
    visit_names = []
    study_weeks = []
    
    table_detection = config.get('table_detection', {})
    soa_keywords = table_detection.get('soa_keywords', ['Procedure'])
    visit_short_name_keywords = table_detection.get('visit_short_name_keywords', ['visit short name'])
    study_week_keywords = table_detection.get('study_week_keywords', ['study week'])
    
    for table in tables:
        rows = table.get("children", [])
        for row in rows:
            if "children" not in row:
                continue
            first_cell = row["children"][0]
            if not first_cell or "children" not in first_cell:
                continue
            first_text = first_cell["children"][0].get("text", "").strip() if first_cell["children"] else ""
            
            if any(keyword in first_text.lower() for keyword in visit_short_name_keywords):
                for cell in row["children"][1:]:
                    if "children" in cell:
                        for p in cell["children"]:
                            txt = p.get("text", "").strip()
                            norm = normalize_visit_name(txt, config)
                            if norm:  # only keep valid normalized names
                                visit_names.append(norm)
            elif any(keyword in first_text.lower() for keyword in study_week_keywords):
                for cell in row["children"][1:]:
                    if "children" in cell:
                        for p in cell["children"]:
                            txt = p.get("text", "").strip()
                            try:
                                study_weeks.append(int(''.join(filter(lambda x: x in '-0123456789', txt))))
                            except:
                                study_weeks.append(None)
    
    # Align lengths after filtering
    min_len = min(len(visit_names), len(study_weeks))
    visit_names = visit_names[:min_len]
    study_weeks = study_weeks[:min_len]
    
    df = pd.DataFrame({'Visit Name': visit_names, 'Study Week': study_weeks})
    return df


def find_element_by_text(data: Dict[str, Any], text_to_find: str) -> Optional[Dict[str, Any]]:
    """Recursively search the JSON for an element containing specific text."""
    if isinstance(data, dict):
        if text_to_find.lower() in data.get('text', '').lower():
            return data
        for child in data.get('children', []):
            found = find_element_by_text(child, text_to_find)
            if found:
                return found
    elif isinstance(data, list):
        for item in data:
            found = find_element_by_text(item, text_to_find)
            if found:
                return found
    return None


def extract_extension_week(doc: Dict[str, Any], config: Dict[str, Any]) -> int:
    """Find the extension start week from the protocol document."""
    extension_config = config.get('extension_detection', {})
    search_section = extension_config.get('search_section', 'Study rationale')
    pattern = extension_config.get('pattern', r'(\d+)\s*weeks on treatment')
    case_insensitive = extension_config.get('case_insensitive', True)
    
    rationale_section = find_element_by_text(doc, search_section)
    if rationale_section:
        full_text = json.dumps(rationale_section)
        flags = re.IGNORECASE if case_insensitive else 0
        match = re.search(pattern, full_text, flags)
        if match:
            week = int(match.group(1))
            logging.info(f"Found extension start at {week} weeks.")
            return week
    
    logging.warning("Could not determine extension start week from JSON. Check 'Study rationale' section.")
    return float('inf')  # Return a very large number if not found


def get_event_group(row: pd.Series, extension_start_week: int, config: Dict[str, Any]) -> str:
    """Classify a visit into an Event Group based on configuration rules."""
    visit_name = row['Visit Name']
    study_week = row['Study Week']
    
    event_groups = config.get('event_groups', {})
    
    # Check hardcoded rules first
    for group_name, group_config in event_groups.items():
        visit_names = group_config.get('visit_names', [])
        if visit_name in visit_names:
            return group_config.get('group_name', group_name.title())
    
    # Apply logic-based rules for remaining visits
    if study_week >= extension_start_week:
        return 'Extension'
    else:
        return 'Main Study'


def generate_visits_with_groups(input_protocol_json: str, output_xlsx: str, 
                               config: Dict[str, Any] = None) -> pd.DataFrame:
    """Generate visits with event groups, offsets and windows and save to Excel."""
    if config is None:
        config = {}
    
    logging.info(f"Generating visits with groups from {input_protocol_json}")
    
    doc = load_json(input_protocol_json)
    
    soa_tables = find_all_soa_tables(doc)
    soa_df = extract_visits_and_weeks(soa_tables, config)
    
    # Keep first occurrence only (no duplicates)
    soa_df = soa_df.drop_duplicates(subset=['Visit Name']).reset_index(drop=True)
    
    # Add Event Group Column
    extension_start_week = extract_extension_week(doc, config)
    soa_df['Event Group'] = soa_df.apply(
        lambda row: get_event_group(row, extension_start_week, config), axis=1
    )
    
    # Calculate offset days and visit windows
    visit_windows = config.get('visit_windows', {})
    offset_calculation = visit_windows.get('offset_calculation', 'study_week * 7')
    early_window = visit_windows.get('early_window', -3)
    late_window = visit_windows.get('late_window', 3)
    offset_types = visit_windows.get('offset_types', {})
    
    soa_df['Offset Days'] = soa_df['Study Week'] * 7
    soa_df['Visit Window Start'] = soa_df['Offset Days'] + early_window
    soa_df['Visit Window End'] = soa_df['Offset Days'] + late_window
    
    # Add Offset Type
    offset_types_list = []
    for i, row in soa_df.iterrows():
        if i == 0:
            offset_types_list.append(offset_types.get('first_visit', 'Specific: V1 a'))
        else:
            offset_types_list.append(offset_types.get('other_visits', 'Previous'))
    soa_df['Offset Type'] = offset_types_list
    
    # Rename columns first
    column_mapping = {
        'Visit Window Start': 'Day Range - Early',
        'Visit Window End': 'Day Range - Late'
    }
    soa_df.rename(columns=column_mapping, inplace=True)
    
    # Reorder and select columns for final output
    output_columns = config.get('output_columns', [
        'Event Group', 'Visit Name', 'Study Week', 'Offset Days', 
        'Offset Type', 'Day Range - Early', 'Day Range - Late'
    ])
    
    # Only select columns that exist
    available_columns = [col for col in output_columns if col in soa_df.columns]
    final_df = soa_df[available_columns].copy()
    
    # Save to Excel
    final_df.to_excel(output_xlsx, index=False)
    logging.info(f"Visits with groups saved to {output_xlsx}")
    
    return final_df


def group_events(protocol_json: str, output_xlsx: str, config: Dict[str, Any] = None) -> str:
    """
    Group events and create visit windows from protocol JSON.
    
    Args:
        protocol_json: Path to protocol JSON file
        output_xlsx: Path to output Excel file
        config: Configuration dictionary
        
    Returns:
        Path to output Excel file
    """
    if config is None:
        config = {}
    
    logging.info(f"Grouping events from {protocol_json}")
    
    try:
        final_df = generate_visits_with_groups(protocol_json, output_xlsx, config)
        return output_xlsx
    except Exception as e:
        logging.error(f"Error grouping events: {e}")
        raise

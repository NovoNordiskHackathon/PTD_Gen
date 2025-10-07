"""
Form Extractor Module

Extracts form information from eCRF JSON files, including form labels, names,
sources, visits, dynamic triggers, and required status.
"""

import json
import csv
import re
import logging
from typing import Dict, List, Any, Optional, Set, Tuple


def get_text(node: Dict[str, Any]) -> str:
    """Extract text from a node safely."""
    if not isinstance(node, dict):
        return ""
    return (node.get("text") or "").strip()


def extract_visit_strings(text: str, patterns: List[str]) -> Set[str]:
    """Extract visit patterns from text using configured patterns."""
    visits = set()
    for pattern in patterns:
        visit_pattern = re.compile(pattern, re.IGNORECASE)
        matches = visit_pattern.findall(text)
        visits.update(matches)
    return visits


def determine_form_source(form_name: str, form_text: str = "", context_text: str = "", 
                         document_context: str = "", config: Dict[str, Any] = None) -> str:
    """
    Determine the source of a form based on naming patterns, context, and document analysis.
    Returns: "Library", "New", or "Ref. Study"
    """
    if config is None:
        config = {}
    
    # Clean the form name
    clean_name = re.sub(r'[\[\]()]', '', form_name).strip()
    clean_name = re.sub(r'\s*–.*', '', clean_name)
    clean_name = re.sub(r'\s*-\s*(Non-)?[Rr]epeating.*', '', clean_name)
    
    # Extract base form name
    base_form_match = re.match(r'([A-Z][A-Z_]*?)(?:_\d+|_[A-Z]+|\d+)?(?:\s|$)', clean_name.upper())
    base_form_name = base_form_match.group(1) if base_form_match else clean_name.upper()
    
    # Combine all text for analysis
    all_text = f"{form_name} {form_text} {context_text} {document_context}".lower()
    
    # Check reference study indicators
    ref_patterns = config.get('source_classification', {}).get('reference_study_indicators', [])
    for pattern in ref_patterns:
        if re.search(pattern, all_text):
            return "Ref. Study"
    
    # Check new form indicators
    new_patterns = config.get('source_classification', {}).get('new_indicators', [])
    for pattern in new_patterns:
        if re.search(pattern, all_text):
            return "New"
    
    # Check library indicators
    library_patterns = config.get('source_classification', {}).get('library_indicators', [])
    for pattern in library_patterns:
        if re.search(pattern, all_text):
            return "Library"
    
    # Standard form database (simplified version)
    standard_domains = {
        'DEMOGRAPHY', 'DEMOGRAPHIC', 'DEMOGRAPHICS', 'DEMO',
        'INCLUSION', 'EXCLUSION', 'INCLUSIONEXCLUSION', 'ELIGIBILITY',
        'INFORMED_CONSENT', 'CONSENT', 'ICF',
        'MEDICAL_HIST', 'MEDICAL_HISTORY', 'MEDHIST',
        'PHYSICAL_EXAM', 'PHYSICALEXAM', 'PE', 'PHYSEXAM',
        'VITAL_SIGNS', 'VITALSIGNS', 'VITALS', 'VS',
        'LAB', 'LABORATORY', 'LABS', 'LABVALUE', 'LABRESULT',
        'ECG', 'ELECTROCARDIOGRAM', 'EKG',
        'AE', 'ADVERSE_EVENT', 'ADVERSEEVENT', 'SAE', 'SERIOUS_AE',
        'CONMED', 'CONCOMITANT_MEDICATION', 'CONCOMITANTMEDICATION',
        'RANDOMIZATION', 'RANDOMISATION', 'RTSM', 'IVRS', 'IWRS'
    }
    
    if base_form_name in standard_domains:
        return "Library"
    
    # Pattern-based classification
    if re.match(r'^[A-Z][A-Z0-9_]*$', clean_name.upper()):
        if len(base_form_name) <= 15:
            return "Library"
        else:
            return "New"
    
    return "Library"  # Default fallback


def extract_trigger_info(text: str, patterns: List[str]) -> Optional[str]:
    """Extract and clean trigger information using configured patterns."""
    if not text or len(text.split()) < 4:
        return None
    
    # Check for match in patterns
    for pattern in patterns:
        if re.search(pattern, text, re.IGNORECASE):
            trigger_text = re.sub(r'\s+', ' ', text.strip())
            if len(trigger_text) > 300:
                trigger_text = trigger_text[:297] + "..."
            return trigger_text
    
    return None


def is_valid_form_name(text: str, config: Dict[str, Any]) -> bool:
    """Check if text is a valid form name using configured patterns."""
    if not text:
        return False
    
    form_patterns = config.get('form_name_patterns', {})
    valid_brackets = form_patterns.get('valid_brackets', r'\[([A-Z0-9_\-]{3,})\]')
    valid_repeating = form_patterns.get('valid_repeating', r'.*\b(Non-)?[Rr]epeating\b.*')
    invalid_patterns = form_patterns.get('invalid_patterns', [])
    
    form_name_pattern = re.compile(f'(?:{valid_brackets}|{valid_repeating})', re.IGNORECASE)
    match = form_name_pattern.search(text)
    
    if not match:
        return False
    
    if match.group(1):  # Bracketed match
        bracketed_content = match.group(1)
        if not bracketed_content.isupper():
            return False
        return True
    
    # Repeating match
    if len(text) < 10 or len(text) > 80:
        return False
    
    for pattern in invalid_patterns:
        if re.match(pattern, text, re.IGNORECASE):
            return False
    
    return True


def is_valid_form_label(text: str) -> bool:
    """Check if text is a valid form label."""
    if not text or len(text) < 3 or len(text) > 100:
        return False
    
    invalid_patterns = [
        r'^\s*V\d+[A-Z]*\s*$',
        r'Design\s*Notes?\s*:?$',
        r'Oracle\s*item\s*design\s*notes?\s*:?$',
        r'General\s*item\s*design\s*notes?\s*:?$',
        r'^\s*Non-Visit\s*Related\s*$',
        r'^Data from.*',
        r'^Hidden item.*',
        r'^\d+\s+',
        r'^\s*(Non-)?[Rr]epeating(\s+form)?\s*$',
    ]
    
    for pattern in invalid_patterns:
        if re.match(pattern, text, re.IGNORECASE):
            return False
    
    return True


def deep_search_visits(node: Dict[str, Any], patterns: List[str]) -> Set[str]:
    """Recursively search all descendants for visit strings."""
    visits = set()
    if not isinstance(node, dict):
        return visits
    
    text = get_text(node)
    visits.update(extract_visit_strings(text, patterns))
    
    for child in node.get("children", []):
        visits.update(deep_search_visits(child, patterns))
    
    return visits


def deep_search_triggers(node: Dict[str, Any], patterns: List[str], max_depth: int = 5, current_depth: int = 0) -> List[Dict[str, Any]]:
    """Recursively search for triggers with increased depth for better coverage."""
    triggers = []
    if not isinstance(node, dict) or current_depth > max_depth:
        return triggers
    
    text = get_text(node)
    trigger_info = extract_trigger_info(text, patterns)
    if trigger_info:
        triggers.append({
            'text': trigger_info,
            'depth': current_depth
        })
    
    for child in node.get("children", []):
        triggers.extend(deep_search_triggers(child, patterns, max_depth, current_depth + 1))
    
    return triggers


def find_all_required_patterns_globally(data: Dict[str, Any]) -> Dict[str, List[str]]:
    """Find all required patterns and map them to forms."""
    required_mappings = {}
    all_form_nodes = []
    all_required_nodes = []
    
    def collect_nodes(node: Dict[str, Any], path_ancestry: List = [], depth: int = 0):
        if not isinstance(node, dict):
            return
        
        current_path = path_ancestry + [node]
        text = get_text(node)
        node_path = node.get('path', '')
        node_name = node.get('name', '')
        
        # Collect form nodes (simplified)
        if '[' in text and ']' in text and len(text) > 5:
            all_form_nodes.append({
                'node': node,
                'text': text,
                'path': node_path,
                'name': node_name,
                'ancestry': current_path,
                'depth': depth
            })
        
        # Collect required pattern nodes
        required_pattern = re.compile(r'.*Key\s*:\s*\[\*\]\s*=\s*Item\s+is\s+required\.?\s*.*', re.IGNORECASE)
        if re.search(required_pattern, text):
            all_required_nodes.append({
                'node': node,
                'text': text,
                'path': node_path,
                'name': node_name,
                'ancestry': current_path,
                'depth': depth
            })
        
        # Recurse through children
        for child in node.get("children", []):
            collect_nodes(child, current_path, depth + 1)
    
    collect_nodes(data)
    
    # Simple mapping: map each required pattern to the closest form
    for req_info in all_required_nodes:
        req_path = req_info['path']
        req_section_num = 0
        
        # Extract section number from path
        matches = re.findall(r'\[(\d+)\]', req_path)
        if matches:
            req_section_num = int(matches[0])
        
        closest_form = None
        min_distance = float('inf')
        
        for form_info in all_form_nodes:
            form_path = form_info['path']
            form_section_num = 0
            
            matches = re.findall(r'\[(\d+)\]', form_path)
            if matches:
                form_section_num = int(matches[0])
            
            distance = abs(req_section_num - form_section_num)
            if distance < min_distance:
                min_distance = distance
                closest_form = form_info
        
        if closest_form and min_distance <= 5:
            form_text = closest_form['text']
            if form_text not in required_mappings:
                required_mappings[form_text] = []
            required_mappings[form_text].append(req_info['text'])
    
    return required_mappings


def extract_forms_with_corrections(data: Dict[str, Any], config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Extract forms with all corrections and source detection."""
    results = []
    seen_forms = set()
    all_triggers = []
    h1_sections = []
    
    # Get configuration
    visit_patterns = config.get('visit_patterns', [r'\bV\d+[A-Z]*(?:-\d+)?\b'])
    trigger_patterns = config.get('trigger_patterns', [])
    ignore_patterns = config.get('ignore_patterns', [])
    
    # Find required patterns
    required_mappings = find_all_required_patterns_globally(data)
    
    def gather_h1_sections(node: Dict[str, Any]):
        if not isinstance(node, dict):
            return
        if node.get('name', '').startswith('H1'):
            h1_sections.append(node)
        for child in node.get('children', []):
            gather_h1_sections(child)
    
    gather_h1_sections(data)
    
    # Extract document context
    def extract_document_context(node: Dict[str, Any], context_parts: List[str] = None) -> str:
        if context_parts is None:
            context_parts = []
        if not isinstance(node, dict):
            return " ".join(context_parts)
        text = get_text(node)
        if text and len(text) > 10:
            context_parts.append(text[:200])
        for child in node.get("children", []):
            extract_document_context(child, context_parts)
            if len(context_parts) > 20:
                break
        return " ".join(context_parts)
    
    document_context = extract_document_context(data)
    
    for idx, h1_node in enumerate(h1_sections):
        h1_text = get_text(h1_node)
        if not is_valid_form_label(h1_text):
            h1_text = "Unknown Section"
        
        section_visits = deep_search_visits(h1_node, visit_patterns)
        section_triggers = deep_search_triggers(h1_node, trigger_patterns, max_depth=6)
        
        # Extract section context
        section_context = extract_document_context(h1_node)
        
        def find_forms_in_node(node: Dict[str, Any], current_label: str = None, parent_siblings: List = None, ancestors: List = None):
            if ancestors is None:
                ancestors = []
            if not isinstance(node, dict):
                return
            
            node_name = node.get("name", "")
            node_text = get_text(node)
            children = node.get("children", [])
            
            # Update label logic for H2 sections
            if node_name.startswith("H2") and is_valid_form_label(node_text) and not is_valid_form_name(node_text, config):
                current_label = node_text
            
            if is_valid_form_name(node_text, config):
                form_name = node_text
                form_label = current_label if current_label else h1_text
                
                # Skip if matches ignore patterns
                skip_form = False
                for pattern in ignore_patterns:
                    if re.search(pattern, form_name, re.IGNORECASE):
                        skip_form = True
                        break
                if skip_form:
                    return
                
                form_visits = deep_search_visits(node, visit_patterns)
                if not form_visits:
                    form_visits = section_visits
                
                visits_str = ", ".join(sorted(form_visits, key=lambda x: (
                    int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 9999,
                    x
                )))
                
                form_key = (form_label, form_name, visits_str)
                
                if form_key not in seen_forms:
                    # Enhanced trigger search
                    form_triggers = deep_search_triggers(node, trigger_patterns, max_depth=7)
                    if not form_triggers:
                        form_triggers = section_triggers
                    
                    # Special handling for ENR form
                    if '[ENR]' in form_name:
                        form_triggers = []
                    
                    # Deduplicate triggers
                    unique_triggers = list(set(t['text'] for t in form_triggers if extract_trigger_info(t['text'], trigger_patterns)))
                    all_triggers.extend(unique_triggers)
                    
                    has_trigger = len(unique_triggers) > 0
                    trigger_details = unique_triggers[0] if has_trigger else ""
                    
                    # Determine source
                    node_context = extract_document_context(node)
                    source = determine_form_source(
                        form_name=form_name,
                        form_text=node_text,
                        context_text=f"{section_context} {node_context}",
                        document_context=document_context,
                        config=config
                    )
                    
                    # Check if required
                    required_flag = "Yes" if form_name in required_mappings else "No"
                    
                    results.append({
                        "Form Label": form_label,
                        "Form Name": form_name,
                        "Source": source,
                        "Visits": visits_str,
                        "Dynamic Trigger": "Yes" if has_trigger else "No",
                        "Trigger Details": trigger_details,
                        "Required": required_flag
                    })
                    seen_forms.add(form_key)
            
            for child in children:
                find_forms_in_node(child, current_label, children, ancestors + [node])
        
        find_forms_in_node(h1_node, None, None, [])
    
    return results


def extract_forms(ecrf_json: str, output_csv: str, config: Dict[str, Any] = None) -> str:
    """
    Extract forms from eCRF JSON and save to CSV.
    
    Args:
        ecrf_json: Path to eCRF JSON file
        output_csv: Path to output CSV file
        config: Configuration dictionary
        
    Returns:
        Path to output CSV file
    """
    if config is None:
        config = {}
    
    logging.info(f"Extracting forms from {ecrf_json}")
    
    try:
        with open(ecrf_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        extracted_forms = extract_forms_with_corrections(data, config)
        
        # Write to CSV
        with open(output_csv, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = config.get('required_keys', [
                "Form Label", "Form Name", "Source", "Visits", 
                "Dynamic Trigger", "Trigger Details", "Required"
            ])
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in extracted_forms:
                writer.writerow(row)
        
        logging.info(f"Extracted {len(extracted_forms)} forms to {output_csv}")
        return output_csv
        
    except Exception as e:
        logging.error(f"Error extracting forms: {e}")
        raise


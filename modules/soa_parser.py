"""
SoA Parser Module

Parses schedule of activities from protocol JSON files, extracting visit patterns,
procedures, and creating a structured schedule CSV.
"""

import json
import re
import logging
import pandas as pd
from typing import Dict, List, Any, Optional, Set, Tuple


def load_json(file_path: str) -> Dict[str, Any]:
    """Load JSON file."""
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_node_text(node: Dict[str, Any]) -> str:
    """Extract text from a node and its children."""
    if not node:
        return ""
    text = node.get("text", "") or ""
    for child in node.get("children", []):
        text += " " + get_node_text(child)
    return text.replace('\n', ' ').replace('\r', ' ').strip()


def find_nodes_by_name(root: Dict[str, Any], name_prefix: str) -> List[Dict[str, Any]]:
    """Find all nodes with names starting with the given prefix."""
    found = []
    
    def walk(node: Any):
        if isinstance(node, dict):
            if node.get("name", "").startswith(name_prefix):
                found.append(node)
            for child in node.get("children", []):
                walk(child)
        elif isinstance(node, list):
            for item in node:
                walk(item)
    
    walk(root)
    return found


def flatten_row(row: Dict[str, Any]) -> List[str]:
    """Flatten a table row into a list of text values."""
    texts = []
    for cell in row.get("children", []):
        texts.append(get_node_text(cell))
    return texts


def cell_has_marker(text: str, markers: List[str]) -> bool:
    """Check if cell text contains any of the configured markers."""
    if not isinstance(text, str):
        return False
    for pattern in markers:
        if re.search(pattern, text, flags=re.IGNORECASE):
            return True
    return False


def extract_complete_visit_identifier(text: str, patterns: List[str]) -> Optional[str]:
    """Extract visit identifier with stricter matching to avoid false positives."""
    if not isinstance(text, str):
        return None
    
    text = text.strip()
    
    # Try patterns with word boundaries
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            longest_match = max(matches, key=len)
            
            # Check if the match is a significant portion of the text
            text_no_spaces = text.replace(' ', '').replace('(', '').replace(')', '').replace('-', '')
            match_proportion = len(longest_match) / len(text_no_spaces)
            
            if match_proportion > 0.3:  # Match must be >30% of the text
                return longest_match
    
    return None


def detect_visit_header_row(all_rows: List[List[str]], config: Dict[str, Any]) -> Optional[List[str]]:
    """Detect the row containing visit headers."""
    best_row = None
    best_score = 0
    
    visit_patterns = config.get('visit_patterns', [])
    header_keywords = config.get('header_keywords', [])
    min_visit_count = config.get('min_visit_count', 3)
    
    for row_idx, row in enumerate(all_rows):
        if not row:
            continue
        
        visit_count = 0
        unique_visits = set()
        
        for cell in row:
            visit_id = extract_complete_visit_identifier(str(cell), visit_patterns)
            if visit_id:
                visit_count += 1
                unique_visits.add(visit_id.upper())
        
        score = len(unique_visits)
        row_text = ' '.join(str(cell).lower() for cell in row)
        
        for keyword in header_keywords:
            if re.search(keyword, row_text):
                score += 2
        
        if score > best_score and score >= min_visit_count:
            best_score = score
            best_row = row
    
    return best_row


def find_schedule_end(all_rows: List[List[str]], column_to_visit: Dict[int, str], 
                     start_from: int = 0, config: Dict[str, Any] = None) -> int:
    """Find where schedule procedures end."""
    if config is None:
        config = {}
    
    cell_markers = config.get('cell_markers', [r'\b(?:X|YES|Y)\b'])
    section_breaks = config.get('section_breaks', [])
    min_procedures = config.get('min_procedures', 25)
    consecutive_threshold = config.get('consecutive_non_procedures_threshold', 25)
    
    procedure_count = 0
    consecutive_non_procedures = 0
    
    total_procedures = 0
    for i, row in enumerate(all_rows[start_from:], start_from):
        if not row:
            continue
        has_markers = any(cell_has_marker(str(row[col]), cell_markers) if col < len(row) else False
                          for col in column_to_visit.keys())
        if has_markers:
            total_procedures += 1
    
    logging.info(f"Found {total_procedures} total rows with visit markers")
    
    for i, row in enumerate(all_rows[start_from:], start_from):
        if not row:
            continue
        
        first_cell = str(row[0]).strip()
        has_markers = any(cell_has_marker(str(row[col]), cell_markers) if col < len(row) else False
                          for col in column_to_visit.keys())
        
        if has_markers:
            procedure_count += 1
            consecutive_non_procedures = 0
        else:
            consecutive_non_procedures += 1
            
            if procedure_count >= min_procedures and consecutive_non_procedures > consecutive_threshold:
                logging.info(f"Found schedule end at row {i} ({procedure_count} procedures found)")
                return i
            
            if procedure_count >= min_procedures:
                for pattern in section_breaks:
                    if re.match(pattern, first_cell, re.IGNORECASE):
                        logging.info(f"Found section break at row {i}: '{first_cell}' ({procedure_count} procedures)")
                        return i
    
    logging.info(f"No clear end found, processing all {len(all_rows)} rows")
    return len(all_rows)


def merge_broken_tables(tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Merge tables that may have been split during parsing."""
    if not tables:
        return []
    
    merged = []
    buffer = None
    
    for table in tables:
        rows = find_nodes_by_name(table, "TR")
        table_content = [flatten_row(row) for row in rows]
        
        has_visits = False
        for row in table_content:
            visit_count = sum(1 for cell in row if extract_complete_visit_identifier(str(cell), [r'\b(?:V|P)\d+[A-Za-z]*\b']))
            if visit_count >= 2:
                has_visits = True
                break
        
        if buffer is None:
            buffer = table
            buffer_has_visits = has_visits
            continue
        
        if not has_visits:
            buffer["children"].extend(rows)
        else:
            if buffer_has_visits:
                merged.append(buffer)
                buffer = table
                buffer_has_visits = True
            else:
                buf_rows = find_nodes_by_name(buffer, "TR")
                table["children"] = buf_rows + table.get("children", [])
                buffer = table
                buffer_has_visits = True
    
    if buffer is not None:
        merged.append(buffer)
    
    return merged


def find_all_schedule_tables(root: Dict[str, Any], config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Find all tables that contain schedule information."""
    tables = find_nodes_by_name(root, "Table")
    merged_tables = merge_broken_tables(tables)
    
    visit_patterns = config.get('visit_patterns', [])
    min_visit_count = config.get('min_visit_count', 3)
    
    schedule_tables = []
    for table in merged_tables:
        rows = find_nodes_by_name(table, "TR")
        table_content = [flatten_row(row) for row in rows]
        
        has_visit_patterns = False
        for row in table_content:
            visit_count = sum(1 for cell in row if extract_complete_visit_identifier(str(cell), visit_patterns))
            if visit_count >= min_visit_count:
                has_visit_patterns = True
                break
        
        if has_visit_patterns:
            schedule_tables.append(table)
    
    return schedule_tables


def parse_protocol_schedule(protocol_data: Dict[str, Any], config: Dict[str, Any]) -> Tuple[Optional[Dict[str, List[str]]], Optional[List[str]], Optional[List[str]]]:
    """Parse the protocol schedule and extract visit-procedure mappings."""
    schedule = {}
    tables = find_all_schedule_tables(protocol_data, config)
    
    if not tables:
        logging.error("No schedule tables found")
        return None, None, None
    
    all_rows = []
    for table in tables:
        rows = find_nodes_by_name(table, "TR")
        all_rows.extend([flatten_row(row) for row in rows])
    
    visit_row = detect_visit_header_row(all_rows, config)
    
    if not visit_row:
        logging.error("Could not find visit header row")
        return None, None, None
    
    logging.info("Found visit header row")
    
    visit_patterns = config.get('visit_patterns', [])
    cell_markers = config.get('cell_markers', [])
    procedure_filters = config.get('procedure_filters', [])
    
    column_to_visit = {}
    visit_order = []
    seen_visits = set()
    
    for i, cell in enumerate(visit_row):
        visit_id = extract_complete_visit_identifier(str(cell), visit_patterns)
        if visit_id:
            original_visit = visit_id
            counter = 1
            while visit_id in seen_visits:
                visit_id = f"{original_visit}_{counter}"
                counter += 1
            
            column_to_visit[i] = visit_id
            visit_order.append(visit_id)
            seen_visits.add(visit_id)
    
    logging.info(f"Detected {len(visit_order)} unique visits: {visit_order}")
    
    if len(visit_order) == 0:
        logging.error("No visit columns detected")
        return None, None, None
    
    header_row_index = -1
    for i, row in enumerate(all_rows):
        if row == visit_row:
            header_row_index = i
            break
    
    end_index = find_schedule_end(all_rows, column_to_visit, header_row_index + 1, config)
    logging.info(f"Processing rows {header_row_index + 1} to {end_index}")
    
    procedure_order = []
    
    for i, row in enumerate(all_rows[header_row_index + 1:end_index], header_row_index + 1):
        if not row:
            continue
        
        first_cell = str(row[0]).strip() if len(row) > 0 else ""
        
        # Skip if first cell matches procedure filters
        skip_row = False
        for filter_term in procedure_filters:
            if first_cell.lower() == filter_term.lower() or extract_complete_visit_identifier(first_cell, visit_patterns):
                skip_row = True
                break
        
        if skip_row:
            continue
        
        procedure = first_cell
        
        has_markers = any(cell_has_marker(str(row[col]), cell_markers) if col < len(row) else False
                          for col in column_to_visit.keys())
        
        if has_markers:
            if procedure not in procedure_order:
                procedure_order.append(procedure)
            
            for col, visit_name in column_to_visit.items():
                cell_text = row[col] if col < len(row) else ""
                if cell_has_marker(cell_text, cell_markers):
                    schedule.setdefault(visit_name, []).append(procedure)
    
    return schedule, visit_order, procedure_order


def save_schedule_to_csv(schedule: Dict[str, List[str]], visit_order: List[str], 
                        procedure_order: List[str], output_path: str) -> None:
    """Save the schedule to CSV format."""
    if not schedule:
        logging.error("Schedule is empty, not saving CSV.")
        return
    
    df = pd.DataFrame(index=procedure_order, columns=visit_order)
    df = df.fillna('')
    
    for visit, procedures in schedule.items():
        for proc in procedures:
            if proc in df.index:
                df.loc[proc, visit] = 'X'
    
    df.index.name = "Procedure"
    df.to_csv(output_path)
    logging.info(f"Schedule saved to '{output_path}'")
    logging.info(f"Total procedures: {len(procedure_order)}")
    logging.info(f"Total visits: {len(visit_order)}")


def parse_soa(protocol_json: str, output_csv: str, config: Dict[str, Any] = None) -> str:
    """
    Parse schedule of activities from protocol JSON and save to CSV.
    
    Args:
        protocol_json: Path to protocol JSON file
        output_csv: Path to output CSV file
        config: Configuration dictionary
        
    Returns:
        Path to output CSV file
    """
    if config is None:
        config = {}
    
    logging.info(f"Parsing SoA from {protocol_json}")
    
    try:
        protocol_data = load_json(protocol_json)
        schedule, visit_order, procedure_order = parse_protocol_schedule(protocol_data, config)
        
        if schedule:
            save_schedule_to_csv(schedule, visit_order, procedure_order, output_csv)
            return output_csv
        else:
            raise ValueError("Failed to parse schedule from protocol JSON")
            
    except Exception as e:
        logging.error(f"Error parsing SoA: {e}")
        raise


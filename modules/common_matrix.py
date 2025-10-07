"""
Common Matrix Module

Creates an ordered SoA matrix by merging eCRF forms with protocol procedures
using fuzzy matching to map forms to their appropriate positions.
"""

import pandas as pd
import logging
from difflib import SequenceMatcher
from typing import Dict, Any, List, Optional


def fuzzy_match(a: str, b: str, case_insensitive: bool = True) -> float:
    """Calculate similarity ratio between two strings."""
    if case_insensitive:
        a, b = a.lower(), b.lower()
    return SequenceMatcher(None, a, b).ratio()


def generate_ordered_soa_matrix(ecrf_file: str, schedule_file: str, output_file: str, 
                               config: Dict[str, Any] = None) -> pd.DataFrame:
    """
    Generate SoA matrix with per-visit ordering using fuzzy matching.
    
    Args:
        ecrf_file: Path to extracted forms CSV
        schedule_file: Path to schedule CSV
        output_file: Path to output CSV
        config: Configuration dictionary
        
    Returns:
        DataFrame containing the ordered SoA matrix
    """
    if config is None:
        config = {}
    
    logging.info(f"Generating ordered SoA matrix from {ecrf_file} and {schedule_file}")
    
    # Load configuration
    threshold = config.get('fuzzy_threshold', 0.5)
    include_unmapped = config.get('include_unmapped', False)
    case_insensitive = config.get('fuzzy_matching', {}).get('case_insensitive', True)
    
    # Column mappings
    visit_mapping = config.get('visit_column_mapping', {})
    output_columns = config.get('output_columns', {})
    
    form_label_col = visit_mapping.get('form_label_column', 'Form Label')
    form_name_col = visit_mapping.get('form_name_column', 'Form Name')
    source_col = visit_mapping.get('source_column', 'Source')
    visits_col = visit_mapping.get('visits_column', 'Visits')
    dynamic_trigger_col = visit_mapping.get('dynamic_trigger_column', 'Dynamic Trigger')
    trigger_details_col = visit_mapping.get('trigger_details_column', 'Trigger Details')
    required_col = visit_mapping.get('required_column', 'Required')
    
    # Load data
    try:
        extracted = pd.read_csv(ecrf_file)
        schedule = pd.read_csv(schedule_file)
    except Exception as e:
        logging.error(f"Error loading input files: {e}")
        raise
    
    # Procedure order from schedule
    proc_order = list(schedule['Procedure'])
    
    # Fuzzy mapping: Best match per form
    form_order_map = {}
    unmapped_forms = []
    
    for form_label in extracted[form_label_col].unique():
        best_score = 0
        best_idx = 9999  # High index for unmapped
        best_proc = None
        
        for idx, proc in enumerate(proc_order):
            score = fuzzy_match(proc, form_label, case_insensitive)
            if score >= threshold and score > best_score:
                best_score = score
                best_idx = idx
                best_proc = proc
        
        if best_proc:
            form_order_map[form_label] = {'index': best_idx, 'procedure': best_proc}
        else:
            unmapped_forms.append(form_label)
            if include_unmapped:
                form_order_map[form_label] = {'index': 9999, 'procedure': 'Unmapped'}
    
    logging.info(f"Unmapped forms: {unmapped_forms}")
    
    # Sort extracted forms based on mapping
    extracted['SortIndex'] = extracted[form_label_col].map(
        lambda x: form_order_map.get(x, {'index': 9999})['index']
    )
    ex_sorted = extracted.sort_values('SortIndex').reset_index(drop=True)
    
    # Visit order from schedule
    visits = [col for col in schedule.columns if col != 'Procedure']
    
    # Initialize matrix
    data_rows = []
    visit_parsing = config.get('visit_parsing', {})
    separator = visit_parsing.get('separator', ',')
    strip_whitespace = visit_parsing.get('strip_whitespace', True)
    
    for _, row in ex_sorted.iterrows():
        visits_list = []
        if pd.notna(row[visits_col]):
            # Robust parsing of visits
            visits_str = str(row[visits_col])
            if strip_whitespace:
                visits_list = [v.strip() for v in visits_str.split(separator) if v.strip()]
            else:
                visits_list = [v for v in visits_str.split(separator) if v]
        
        row_dict = {
            output_columns.get('form_label', 'Form Label'): row[form_label_col],
            output_columns.get('form_name', 'Form Name'): row[form_name_col],
            output_columns.get('source', 'Source'): row.get(source_col, ''),
            output_columns.get('is_dynamic', 'Is Form Dynamic?'): row.get(dynamic_trigger_col, 'No'),
            output_columns.get('dynamic_criteria', 'Form Dynamic Criteria'): row.get(trigger_details_col, '')
        }
        
        # Initialize visit columns
        for visit in visits:
            row_dict[visit] = ''
        
        data_rows.append(row_dict)
    
    matrix_df = pd.DataFrame(data_rows)
    matrix_df = matrix_df.astype({v: object for v in visits})
    
    # Assign sequential numbers per visit
    for visit in visits:
        counter = 1
        for idx, row in matrix_df.iterrows():
            visits_list = []
            if pd.notna(ex_sorted.loc[idx, visits_col]):
                visits_str = str(ex_sorted.loc[idx, visits_col])
                if strip_whitespace:
                    visits_list = [v.strip() for v in visits_str.split(separator) if v.strip()]
                else:
                    visits_list = [v for v in visits_str.split(separator) if v]
            
            if visit in visits_list:
                matrix_df.at[idx, visit] = counter
                counter += 1
            else:
                matrix_df.at[idx, visit] = ''
    
    # Save to CSV
    matrix_df.to_csv(output_file, index=False)
    logging.info(f"SoA matrix saved to {output_file}")
    
    return matrix_df


def merge_common_matrix(ecrf_csv: str, schedule_csv: str, output_csv: str, 
                       config: Dict[str, Any] = None) -> str:
    """
    Merge eCRF forms with schedule to create ordered SoA matrix.
    
    Args:
        ecrf_csv: Path to eCRF forms CSV
        schedule_csv: Path to schedule CSV
        output_csv: Path to output CSV
        config: Configuration dictionary
        
    Returns:
        Path to output CSV file
    """
    if config is None:
        config = {}
    
    logging.info(f"Merging common matrix from {ecrf_csv} and {schedule_csv}")
    
    try:
        matrix_df = generate_ordered_soa_matrix(ecrf_csv, schedule_csv, output_csv, config)
        return output_csv
    except Exception as e:
        logging.error(f"Error merging common matrix: {e}")
        raise


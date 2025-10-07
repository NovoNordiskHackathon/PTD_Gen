# Clinical Trial Data Processing Pipeline

A modular pipeline for processing clinical trial protocol and eCRF JSON files to generate comprehensive schedule grids and study-specific forms for clinical trial planning.

## Available Pipelines

### 1. PTD Schedule Grid Generator (`generate_ptd.py`)
Generates comprehensive schedule grids for clinical trial planning.

### 2. Study Specific Forms Generator (`generate_study_specific_forms.py`)
Generates study-specific forms from eCRF JSON files with detailed item analysis.

## Overview

This tool refactors and merges multiple existing scripts into a single, clean, modular pipeline that produces a `schedule_grid.csv` file. The pipeline is designed to be configurable and reusable by others.

## Features

- **Modular Design**: Each processing stage is a separate, configurable module
- **Configuration-Driven**: JSON configuration files for each module allow customization without code changes
- **Comprehensive Logging**: Detailed logging with configurable levels
- **Intermediate File Management**: Option to keep or clean up intermediate files
- **Error Handling**: Robust error handling with informative messages

## PTD Schedule Grid Pipeline Stages

1. **Form Extraction** (`extract_forms`): Extract form information from eCRF JSON
2. **SoA Parsing** (`parse_soa`): Parse schedule of activities from protocol JSON
3. **Common Matrix** (`merge_common_matrix`): Create ordered SoA matrix with fuzzy matching
4. **Event Grouping** (`group_events`): Generate visit groups with event windows
5. **Schedule Layout** (`generate_schedule_grid`): Create final schedule grid layout

## Study Specific Forms Pipeline

The Study Specific Forms Generator processes eCRF JSON files to extract detailed form information and generate comprehensive Excel reports with:

- **Form Analysis**: Extracts form labels, names, and hierarchical structure
- **Item Extraction**: Identifies and processes form items with proper validation
- **Data Type Detection**: Automatically determines data types (Text, Codelist, Date/Time, etc.)
- **Item Grouping**: Analyzes repeating and non-repeating item groups
- **Validation Rules**: Applies business rules for required fields, codelists, and data validation
- **Excel Output**: Generates formatted Excel files with proper structure and styling

## Installation

No additional dependencies beyond the existing project requirements. The pipeline uses the same libraries as the original scripts.

## Usage

### Basic Usage

```bash
python generate_ptd.py --protocol hierarchical_output_final_protocol.json --ecrf hierarchical_output_final_ecrf.json
```

### Advanced Usage (Recommended)

```bash
python generate_ptd.py \
  --protocol hierarchical_output_final_protocol.json \
  --ecrf hierarchical_output_final_ecrf.json \
  --out ./output/my_schedule_grid.xlsx \
  --keep-intermediates \
  --log-level DEBUG
```

## Study Specific Forms Generator Usage

### Basic Usage

```bash
python generate_study_specific_forms.py \
  --ecrf hierarchical_output_final_ecrf.json \
  --out study_specific_forms.xlsx
```

### Advanced Usage

```bash
python generate_study_specific_forms.py \
  --ecrf hierarchical_output_final_ecrf.json \
  --out ./output/forms.xlsx \
  --log-level DEBUG
```

### Legacy Usage

```bash
python generate_ptd.py \
  --protocol hierarchical_output_final_protocol.json \
  --ecrf hierarchical_output_final_ecrf.json \
  --output-dir ./output \
  --output-file my_schedule_grid.xlsx \
  --keep-intermediates \
  --log-level DEBUG
```

### Command Line Options

- `--protocol`: Path to protocol JSON file (required)
- `--ecrf`: Path to eCRF JSON file (required)
- `--out`: Final output file path (e.g., output_folder/schedule_grid.xlsx). If not provided, uses --output-dir/schedule_grid.xlsx
- `--output-dir`: Output directory for generated files (default: ./output)
- `--output-file`: Final output filename (default: schedule_grid.xlsx). Ignored if --out is provided.
- `--keep-intermediates`: Keep intermediate files for debugging
- `--log-level`: Logging level (DEBUG, INFO, WARNING, ERROR) (default: INFO)
- `--config-dir`: Directory containing configuration files (default: ./config)

## Configuration

Each module has its own JSON configuration file in the `config/` directory:

### config_form_extractor.json
Configures form extraction from eCRF JSON:
- Input/output paths
- Visit patterns and trigger patterns
- Source classification rules
- Form name validation patterns

### config_soa_parser.json
Configures schedule of activities parsing:
- Visit patterns and cell markers
- Header keywords and section breaks
- Procedure filtering rules
- Table detection parameters

### config_common_matrix.json
Configures the SoA matrix generation:
- Fuzzy matching threshold
- Column mappings
- Visit parsing options
- Output column configuration

### config_event_grouping.json
Configures event grouping and visit windows:
- Visit normalization patterns
- Event group definitions
- Extension detection rules
- Visit window calculations

### config_schedule_layout.json
Configures the final schedule grid layout:
- Column mappings
- Event name patterns
- Triggering rules
- Styling options

## Output Files

### Final Output
- `schedule_grid.xlsx`: The main output file containing the complete schedule grid with proper Excel formatting, visit windows, and dynamic properties

### Intermediate Files (when --keep-intermediates is used)
- `extracted_forms.csv`: Forms extracted from eCRF JSON
- `schedule.csv`: Schedule of activities parsed from protocol JSON
- `soa_matrix.csv`: Ordered SoA matrix with fuzzy matching
- `visits_with_groups.xlsx`: Visit groups with event windows

## Module Structure

```
modules/
├── __init__.py
├── form_extractor.py      # Extract forms from eCRF JSON
├── soa_parser.py          # Parse schedule of activities
├── common_matrix.py       # Create ordered SoA matrix
├── event_grouping.py      # Group events and create visit windows
└── schedule_layout.py     # Generate final schedule grid
```

## Configuration Examples

### Example: Custom Visit Patterns

```json
{
  "visit_patterns": [
    "\\bV\\d+[A-Za-z]*\\b",
    "\\bP\\d+[A-Za-z]*\\b",
    "\\bS\\d+D[\\s-]?\\d+[A-Za-z]*\\b"
  ]
}
```

### Example: Custom Fuzzy Matching Threshold

```json
{
  "fuzzy_threshold": 0.7,
  "include_unmapped": true
}
```

### Example: Custom Event Groups

```json
{
  "event_groups": {
    "screening": {
      "visit_names": ["V1"],
      "group_name": "Screening"
    },
    "randomisation": {
      "visit_names": ["V2"],
      "group_name": "Randomisation"
    }
  }
}
```

## Error Handling

The pipeline includes comprehensive error handling:
- File not found errors
- JSON parsing errors
- Configuration validation
- Data processing errors
- Graceful cleanup on failure

## Logging

The pipeline provides detailed logging:
- Progress through each stage
- Configuration loading
- File processing statistics
- Error messages and stack traces
- Performance metrics

Logs are written to both console and `ptd_generation.log` file.

## Migration from Original Scripts

This pipeline replaces the following original scripts:
- `form_label_form_name_extractor.py` → `modules/form_extractor.py`
- `soa_works_for_all.py` → `modules/soa_parser.py`
- `extracting_commonform_visits.py` → `modules/common_matrix.py`
- `event_grouping_and_event_window_configuration.py` → `modules/event_grouping.py`
- `schedule_grid_final_layout.py` → `modules/schedule_layout.py`
- `main_integration.py` → `generate_ptd.py`

## Troubleshooting

### Common Issues

1. **Configuration file not found**: Ensure config files are in the `config/` directory
2. **JSON parsing errors**: Verify input JSON files are valid
3. **Missing columns**: Check that input files have expected column names
4. **Permission errors**: Ensure write permissions for output directory

### Debug Mode

Use `--log-level DEBUG` to get detailed information about the processing steps.

### Keeping Intermediate Files

Use `--keep-intermediates` to examine intermediate files for debugging.

## Contributing

When modifying the pipeline:
1. Update the relevant module in `modules/`
2. Update the corresponding configuration file in `config/`
3. Test with both existing and new data
4. Update documentation as needed

## License

Same as the original project.

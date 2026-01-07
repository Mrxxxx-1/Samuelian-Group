# Samuelian-Group
A interview project for Samuelian Group

## Financing Costs Parser

This tool extracts and validates financing costs from CTCAC application Excel files for benchmarking purposes.

### Features

- **Automatic Download**: Scrapes and downloads Excel files from the CTCAC applications page
- **Data Extraction**: Extracts financing costs from two key sections:
  - Construction Interest & Fees
  - Permanent Financing
- **Validation**: Verifies that section totals match the sum of line items
- **Derived Metrics**: Calculates:
  - Combined financing costs
  - Financing costs per unit
  - Financing costs per square foot (uses "Total square footage of all project structures (excluding commercial/retail)" from Application tab, column D)
  - Financing costs as percentage of hard costs
- **Summary Report**: Generates an Excel report with all extracted data and metrics

### Installation

1. Install required Python packages:
```bash
pip install -r requirements.txt
```

### Project Structure

The project is modularized into separate files:
- `main.py` - Main entry point
- `data_models.py` - Data structures and models
- `download.py` - File downloading functionality
- `parser.py` - Excel parsing logic
- `report_generator.py` - Report generation

### Usage

Run the parser:
```bash
python main.py
```

The script will:
1. Download the first 10 Excel files from the CTCAC applications page (for testing)
2. Parse each application to extract financing costs
3. Validate the extracted data
4. Generate a summary report (`financing_costs_summary.xlsx`)

To skip downloading and use existing files:
```bash
python main.py --skip-download
```

### Output

The parser generates:
- **financing_costs_summary.xlsx**: Excel file with all extracted data, including:
  - Individual line items from both sections
  - Section totals
  - Combined financing costs
  - Project metrics (units, square footage, new construction total)
  - Derived metrics (per unit, per SF, % of hard costs)
  - Validation errors and warnings

### Data Extracted

#### Construction Interest & Fees
- Construction Loan Interest
- Origination Fee
- Credit Enhancement/Application Fee
- Bond Premium
- Cost of Issuance
- Title & Recording
- Taxes
- Insurance
- Other (with description)
- **Total Construction Interest & Fees**

#### Permanent Financing
- Loan Origination Fee
- Credit Enhancement/Application Fee
- Title & Recording
- Taxes
- Insurance
- Other (with description)
- **Total Permanent Financing Costs**

### Validation

The parser validates:
- Total Construction Interest & Fees equals sum of line items
- Total Permanent Financing Costs equals sum of line items
- Flags missing data (units, square footage, new construction total)

Applications with validation errors are clearly flagged in the output.

### Notes

- Downloaded Excel files are saved in the `applications/` directory
- The parser uses fixed positions for data extraction:
  - **Total number of units**: Extracted from the "Total number of units" row in column D (Application tab)
  - **Total square footage**: Extracted from the "Total square footage of all project structures (excluding commercial/retail)" row in column D (Application tab)
- Missing data is flagged but doesn't prevent processing
- If "New Construction" total is not found, it's flagged as a warning

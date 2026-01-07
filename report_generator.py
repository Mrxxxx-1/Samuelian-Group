"""
Generate summary reports from parsed application data
"""

import pandas as pd
from openpyxl.utils import get_column_letter
from typing import List

from data_models import ApplicationData


def generate_summary_report(applications: List[ApplicationData], output_file: str = "financing_costs_summary.xlsx"):
    """Generate Excel summary report with all extracted data"""
    # Prepare data for DataFrame
    rows = []
    for app in applications:
        row = {
            'Application Name': app.application_name,
            # Construction Interest & Fees
            'Construction Loan Interest': app.construction_interest_fees.construction_loan_interest,
            'Origination Fee': app.construction_interest_fees.origination_fee,
            'Credit Enhancement/Application Fee (Const)': app.construction_interest_fees.credit_enhancement_fee,
            'Bond Premium': app.construction_interest_fees.bond_premium,
            'Cost of Issuance': app.construction_interest_fees.cost_of_issuance,
            'Title & Recording (Const)': app.construction_interest_fees.title_recording,
            'Taxes (Const)': app.construction_interest_fees.taxes,
            'Insurance (Const)': app.construction_interest_fees.insurance,
            'Other (Const)': app.construction_interest_fees.other_amount,
            'Other Description (Const)': app.construction_interest_fees.other_description,
            'Total Construction Interest & Fees': app.construction_interest_fees.total,
            # Permanent Financing
            'Loan Origination Fee': app.permanent_financing.loan_origination_fee,
            'Credit Enhancement/Application Fee (Perm)': app.permanent_financing.credit_enhancement_fee,
            'Title & Recording (Perm)': app.permanent_financing.title_recording,
            'Taxes (Perm)': app.permanent_financing.taxes,
            'Insurance (Perm)': app.permanent_financing.insurance,
            'Other (Perm)': app.permanent_financing.other_amount,
            'Other Description (Perm)': app.permanent_financing.other_description,
            'Total Permanent Financing Costs': app.permanent_financing.total,
            # Combined totals
            'Combined Financing Costs': app.get_combined_financing_costs(),
            # Project metrics
            'Total Units': app.total_units,
            'Total Square Feet': app.total_square_feet,
            'New Construction Total': app.new_construction_total,
            # Derived metrics
            'Financing Costs per Unit': app.get_financing_costs_per_unit(),
            'Financing Costs per SF': app.get_financing_costs_per_sf(),
            'Financing Costs % of Hard Costs': app.get_financing_costs_pct_hard_costs(),
            # Validation
            'Validation Errors': '; '.join(app.validation_errors) if app.validation_errors else '',
            'Validation Warnings': '; '.join(app.validation_warnings) if app.validation_warnings else '',
        }
        rows.append(row)
    
    df = pd.DataFrame(rows)
    
    # Write to Excel with formatting
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Summary']
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = min(max_length + 2, 50)
    
    print(f"\nSummary report saved to: {output_file}")


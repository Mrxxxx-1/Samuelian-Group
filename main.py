"""
Main entry point for financing costs parser
"""

import os
import sys

from download import download_excel_files, find_excel_files_in_directory
from parser import ExcelParser
from report_generator import generate_summary_report


def main():
    """Main execution function"""
    url = "https://www.treasurer.ca.gov/ctcac/2025/thirdround/4percent/application/index.asp"

    # Check if user wants to skip download
    skip_download = "--skip-download" in sys.argv or "--no-download" in sys.argv

    excel_files = []

    if not skip_download:
        print("Step 1: Downloading all Excel files...")
        excel_files = download_excel_files(url)  # Download all available files

    # If no files downloaded, check for existing files in directory
    if not excel_files:
        print("Checking for existing Excel files in 'applications' directory...")
        excel_files = find_excel_files_in_directory()

    if not excel_files:
        print("No Excel files found.")
        print("Options:")
        print("  1. Run without --skip-download to download from URL")
        print("  2. Manually place Excel files in the 'applications/' directory")
        return

    print(f"\nStep 2: Parsing {len(excel_files)} applications...")
    applications = []

    for file_path in excel_files:
        try:
            print(f"Parsing: {os.path.basename(file_path)}")
            parser = ExcelParser(file_path)
            app_data = parser.parse()
            applications.append(app_data)

            # Print validation status
            if app_data.validation_errors:
                print(f"  ⚠️  ERRORS: {', '.join(app_data.validation_errors)}")
            if app_data.validation_warnings:
                print(f"  ⚠️  WARNINGS: {', '.join(app_data.validation_warnings)}")
        except Exception as e:
            print(f"  ❌ Error parsing {file_path}: {str(e)}")

    print("\nStep 3: Generating summary report...")
    generate_summary_report(applications)

    # Print summary statistics
    print("\n" + "=" * 80)
    print("SUMMARY STATISTICS")
    print("=" * 80)

    valid_apps = [app for app in applications if not app.validation_errors]
    print(f"Total applications parsed: {len(applications)}")
    print(f"Valid applications (no errors): {len(valid_apps)}")
    print(f"Applications with errors: {len(applications) - len(valid_apps)}")

    if valid_apps:
        combined_costs = [app.get_combined_financing_costs() for app in valid_apps]
        per_unit = [
            app.get_financing_costs_per_unit()
            for app in valid_apps
            if app.get_financing_costs_per_unit()
        ]
        per_sf = [
            app.get_financing_costs_per_sf()
            for app in valid_apps
            if app.get_financing_costs_per_sf()
        ]

        print("\nCombined Financing Costs:")
        print(f"  Average: ${sum(combined_costs) / len(combined_costs):,.2f}")
        print(f"  Median: ${sorted(combined_costs)[len(combined_costs) // 2]:,.2f}")
        print(f"  Min: ${min(combined_costs):,.2f}")
        print(f"  Max: ${max(combined_costs):,.2f}")

        if per_unit:
            print("\nFinancing Costs per Unit:")
            print(f"  Average: ${sum(per_unit) / len(per_unit):,.2f}")
            print(f"  Median: ${sorted(per_unit)[len(per_unit) // 2]:,.2f}")

        if per_sf:
            print("\nFinancing Costs per SF:")
            print(f"  Average: ${sum(per_sf) / len(per_sf):,.2f}")
            print(f"  Median: ${sorted(per_sf)[len(per_sf) // 2]:,.2f}")


if __name__ == "__main__":
    main()

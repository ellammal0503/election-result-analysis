import pandas as pd
import os

# Input Excel file
excel_file = "palayamkottai_results.xlsx"

# Output directory
output_dir = "csv_output"
os.makedirs(output_dir, exist_ok=True)

# Load Excel file
xls = pd.ExcelFile(excel_file)

print(f"Found sheets: {xls.sheet_names}")

# Iterate through each sheet
for sheet_name in xls.sheet_names:
    try:
        # Read sheet
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Clean sheet name (remove special chars for filename safety)
        safe_sheet_name = "".join(c if c.isalnum() else "_" for c in sheet_name)

        # Output file path
        output_file = os.path.join(output_dir, f"{safe_sheet_name}.csv")

        # Save as CSV
        df.to_csv(output_file, index=False)

        print(f"✅ Saved: {output_file}")

    except Exception as e:
        print(f"❌ Error processing sheet '{sheet_name}': {e}")

print("🎯 All sheets converted successfully!")
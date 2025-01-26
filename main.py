import pandas as pd
import os

# Load the uploaded Excel file
file_path = "C:/Users/sunny/Downloads/vendor_mis_project/input_files/payment_data_12_to_18_jan.xlsx"
excel_data = pd.ExcelFile(file_path)

# Create an output directory
output_dir = "C:/Users/sunny/Downloads/vendor_mis_project/vendor_output_directory"
os.makedirs(output_dir, exist_ok=True)

# Load all sheet data into a dictionary
sheets_data = {sheet_name: excel_data.parse(sheet_name) for sheet_name in excel_data.sheet_names}

# Collect all unique OUTSOURCINGNAMEs across sheets
outsourcing_names = set()
for sheet_data in sheets_data.values():
    if 'OUTSOURCINGNAME' in sheet_data.columns:
        outsourcing_names.update(sheet_data['OUTSOURCINGNAME'].dropna().unique())

# Process each OUTSOURCINGNAME and create workbooks based on the condition
for outsourcing_name in outsourcing_names:
    vendor_data = {}
    valid_for_workbook = False  # Flag to check if workbook should be created

    for sheet_name, sheet_data in sheets_data.items():
        if 'OUTSOURCINGNAME' in sheet_data.columns and 'PANNO' in sheet_data.columns and 'OLD_USER_ID' in sheet_data.columns:
            # Filter rows for the current OUTSOURCINGNAME
            filtered_data = sheet_data[sheet_data['OUTSOURCINGNAME'] == outsourcing_name]

            if not filtered_data.empty:
                # Group by PANNO and check unique OLD_USER_ID counts
                unique_user_counts = filtered_data.groupby('PANNO')['OLD_USER_ID'].nunique()
                if (unique_user_counts > 1).any():
                    valid_for_workbook = True
                    vendor_data[sheet_name] = filtered_data  # Store data for this sheet

    # Write workbook only if valid_for_workbook is True
    if valid_for_workbook:
        output_file = os.path.join(output_dir, f"{outsourcing_name}.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, filtered_data in vendor_data.items():
                filtered_data.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Workbooks for vendors with the condition met have been created in: {output_dir}")

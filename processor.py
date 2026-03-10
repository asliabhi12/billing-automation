import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP
import json

with open("customer_map.json", encoding="utf-8") as f:
    CUSTOMER_MAP = json.load(f)


def process_files(files):

    output = BytesIO()

    # Function to process the "Summary" sheet and extract data
    def extract_summary_data(file):
        file.seek(0)
        wb = load_workbook(file, data_only=True)
        ws = wb['Summary']
        
        def format_date(date_value):
            if isinstance(date_value, datetime):
                return date_value.strftime('%B %d, %Y')
            elif isinstance(date_value, str):
                try:
                    # Try parsing the string date if it's in an unexpected format
                    parsed_date = datetime.strptime(date_value, '%a, %b %d, %Y')
                    return parsed_date.strftime('%B %d, %Y')
                except ValueError:
                    return ''  # Return empty if parsing fails
            return ''
        
        start_date = format_date(ws['C8'].value)
        end_date = format_date(ws['C9'].value)
        
        subscription_id = ws['C5'].value.split('/')[-1]

        customer_name = CUSTOMER_MAP.get(subscription_id, "")

        data = {
            "Customer Name": customer_name,
            "Start Date": start_date,
            "End Date": end_date,
            "Subscription ID": subscription_id
        }
        return data

    # Function to process each file
    def process_file(file_path, summary_data):
        file_path.seek(0)

        df = pd.read_excel(file_path, sheet_name='Data', engine="openpyxl")
        
        # Rename the columns to match the desired format
        df.rename(columns={
            'Meter': 'Meter Name',
            'ServiceName': 'Service Type',
            'ResourceLocation': 'Region',
            'ResourceType': 'Resource Name',
            'Cost': 'Total Cost'
        }, inplace=True)

        # Select the desired columns
        df_transformed = df[['Meter Name', 'Service Type', 'Resource Name', 'Region', 'Total Cost']].copy()

        df_transformed["Total Cost"] = pd.to_numeric(
            df_transformed["Total Cost"], errors="coerce"
        ).fillna(0)
        
        # Add summary data to the DataFrame
        for col, val in summary_data.items():
            df_transformed[col] = val

        # Reorder columns to match the desired output
        df_transformed = df_transformed[[
            "Customer Name", "Start Date", "End Date", "Subscription ID",
            "Meter Name", "Service Type", "Resource Name", "Region", "Total Cost"
        ]]

        total_cost_row = pd.DataFrame({
        "Customer Name": [''],
        "Start Date": [''],
        "End Date": [''],
        "Subscription ID": [''],
        "Meter Name": [''],
        "Service Type": [''],
        "Resource Name": [''],
        "Region": ['Total Cost'],
        "Total Cost": ['FORMULA_PLACEHOLDER']
    })
        df_transformed = pd.concat([df_transformed, total_cost_row], ignore_index=True)
        
        return df_transformed


    with pd.ExcelWriter(output, engine="openpyxl") as writer:

            from openpyxl.utils import get_column_letter

            for file in files:

                summary_data = extract_summary_data(file)
                df_transformed = process_file(file, summary_data)

                sheet_name = file.original_name.replace('.xlsx', '')
                sheet_name = sheet_name[:31].replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')

                df_transformed.to_excel(writer, sheet_name=sheet_name, index=False)

                worksheet = writer.sheets[sheet_name]

                total_col = df_transformed.columns.get_loc("Total Cost") + 1
                col_letter = get_column_letter(total_col)

                last_row = len(df_transformed) + 1  # +1 because Excel rows start at 1 and include header

                worksheet[f"{col_letter}{last_row}"] = f"=SUM({col_letter}2:{col_letter}{last_row-1})"
                col_idx = df_transformed.columns.get_loc("Total Cost") + 1
                col_letter = get_column_letter(col_idx)

                for cell in worksheet[col_letter]:
                    cell.number_format = "0.00"

                # Auto-fit column widths
                for i, col in enumerate(df_transformed.columns):

                    max_length = len(str(col))

                    for value in df_transformed[col]:
                        length = len(str(value))
                        if length > max_length:
                            max_length = length

                    worksheet.column_dimensions[get_column_letter(i + 1)].width = min(max_length + 2, 50)
                    worksheet.freeze_panes = "A2"
                    worksheet.auto_filter.ref = worksheet.dimensions
 


    output.seek(0)
    return output
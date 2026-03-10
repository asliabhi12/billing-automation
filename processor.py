import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO


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
        
        data = {
            "Customer Name": "",
            "Start Date": start_date,
            "End Date": end_date,
            "Subscription ID": ws['C5'].value.split('/')[-1]
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

        # Ensure Total Cost is numeric
        df_transformed["Total Cost"] = (
        pd.to_numeric(df_transformed["Total Cost"], errors="coerce")
        .fillna(0)
        .astype("float32")
         )
        
        # Add summary data to the DataFrame
        for col, val in summary_data.items():
            df_transformed[col] = val

        # Reorder columns to match the desired output
        df_transformed = df_transformed[[
            "Customer Name", "Start Date", "End Date", "Subscription ID",
            "Meter Name", "Service Type", "Resource Name", "Region", "Total Cost"
        ]]

        # Calculate the total cost and append it as the last row
        total_cost_sum = df_transformed['Total Cost'].sum()
        total_cost_row = pd.DataFrame({
            "Customer Name": [''],
            "Start Date": [''],
            "End Date": [''],
            "Subscription ID": [''],
            "Meter Name": [''],
            "Service Type": [''],
            "Resource Name": [''],
            "Region": ['Total Cost'],
            "Total Cost": [total_cost_sum]
        })
        df_transformed = pd.concat([df_transformed, total_cost_row], ignore_index=True)
        
        return df_transformed

    # Create a Pandas Excel writer object
    # output_file_path = 'output-file/transformed_data-2.xlsx'  # Replace with your desired output file path
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for file in files:
            # Extract summary data
            summary_data = extract_summary_data(file)
            
            # Process each file
            df_transformed = process_file(file, summary_data)
            
            # Get the sheet name from the input file name
            sheet_name = file.filename.replace('.xlsx', '')
            
            # Ensure the sheet name is valid (sheet names must be <= 31 characters and cannot contain certain characters)
            sheet_name = sheet_name[:31].replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')
            
            # Write the transformed data to a new sheet
            df_transformed.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Autofit column widths
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for file in files:
            summary_data = extract_summary_data(file)
            df_transformed = process_file(file, summary_data)

            sheet_name = file.filename.replace('.xlsx', '')
            sheet_name = sheet_name[:31].replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')

            df_transformed.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output
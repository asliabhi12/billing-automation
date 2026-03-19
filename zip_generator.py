import zipfile
import os

def create_zip(excel_path, pdf_folder, zip_path):

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:

        # Add Excel only if exists
        if excel_path:
            zipf.write(excel_path, "Azure_Billing_Report.xlsx")

        # Add PDFs
        for file in os.listdir(pdf_folder):
            file_path = os.path.join(pdf_folder, file)
            zipf.write(file_path, f"pdfs/{file}")

    return zip_path
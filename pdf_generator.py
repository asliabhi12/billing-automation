import pandas as pd
import os

from reportlab.lib.pagesizes import A3, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from decimal import Decimal, ROUND_HALF_UP

def generate_pdfs(excel_path, output_folder="pdfs"):

    os.makedirs(output_folder, exist_ok=True)

    styles = getSampleStyleSheet()
    sheets = pd.read_excel(excel_path, sheet_name=None)

    pdf_paths = []

    for sheet_name, df in sheets.items():

        # ===== Extract metadata =====
        customer_name = df["Customer Name"].iloc[0] # [cite: 2]
        start_date = df["Start Date"].iloc[0] # [cite: 3]
        end_date = df["End Date"].iloc[0] # [cite: 3]
        subscription_id = df["Subscription ID"].iloc[0] # [cite: 4]

        # ===== Clean table =====
        df_table = df[
            ["Meter Name", "Service Type", "Resource Name", "Region", "Total Cost"]
        ].copy() # 

        df_table = df_table[df_table["Total Cost"].notna()]

        def to_decimal(x):
            try:
                return Decimal(str(x))
            except:
                return Decimal("0.00")

        df_table["Total Cost"] = df_table["Total Cost"].apply(to_decimal)
        df_table["Total Cost"] = df_table["Total Cost"].apply(
            lambda x: x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        )

        # ===== Calculate total BEFORE formatting =====
        total_value = sum(df_table["Total Cost"]).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )

        # ===== Format for display =====
        display_df = df_table.copy()
        display_df["Total Cost"] = display_df["Total Cost"].apply(lambda x: f"{x:.2f}")

        # ===== Build table data =====
        table_data = [display_df.columns.tolist()] + display_df.values.tolist()

        # ✅ Total Cost Row (Second last row) 
        # Matches reference: Empty, Empty, Empty, "Total Cost", Value
        table_data.append(["", "", "", "Total Cost", f"{total_value:.2f}"])

        # ✅ Subscription ID Row (Last row) [cite: 4]
        table_data.append([f"Subscription ID : {subscription_id}"] + [""] * 4)

        # ===== Row indexes =====
        total_row_index = len(table_data) - 2
        subscription_row_index = len(table_data) - 1

        # ===== Column widths =====
        col_widths = [230, 150, 260, 120, 120]

        pdf_path = os.path.join(output_folder, f"{sheet_name}.pdf")

        title = Paragraph(
            f"""
            <para align=center>
            <font size=14><b>Microsoft Azure Utilization Report for {customer_name}</b></font><br/>
            <font size=11>{start_date} to {end_date}</font>
            </para>
            """,
            styles["Normal"],
        )

        table = Table(table_data, colWidths=col_widths)

        table.setStyle(TableStyle([
    # ── Header ──────────────────────────────────────────────────────────────
    ("BACKGROUND", (0, 0),  (-1, 0),                  colors.HexColor("#1f4e79")),
    ("TEXTCOLOR",  (0, 0),  (-1, 0),                  colors.white),
    ("FONTNAME",   (0, 0),  (-1, 0),                  "Helvetica-Bold"),

    # ── Data rows (explicit white, overrides ReportLab default) ─────────────
    ("BACKGROUND", (0, 1),  (-1, total_row_index - 1), colors.white),

    # ── Grid & default alignment ─────────────────────────────────────────────
    ("GRID",  (0, 0), (-1, -1), 0.3, colors.black),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),               # global default

    # ── Data rows: right-align Region + Total Cost columns ───────────────────
    ("ALIGN", (3, 1), (4, total_row_index - 1), "RIGHT"),

    # ── Total Cost row ───────────────────────────────────────────────────────
    ("BACKGROUND", (0, total_row_index), (-1, total_row_index), colors.white),

    # Override global CENTER → right-align the label so it sits flush
    # against the value column (professional invoice look)
    ("ALIGN",    (3, total_row_index), (3, total_row_index), "RIGHT"),
    ("ALIGN",    (4, total_row_index), (4, total_row_index), "RIGHT"),

    # Bold ONLY the label cell and the value cell — nothing else in the row
    ("FONTNAME", (3, total_row_index), (3, total_row_index), "Helvetica-Bold"),
    ("FONTNAME", (4, total_row_index), (4, total_row_index), "Helvetica-Bold"),

    # Keep the empty cells (cols 0-2) in plain Helvetica so they don't
    # accidentally inherit bold from any surrounding rule
    ("FONTNAME", (0, total_row_index), (2, total_row_index), "Helvetica"),

    # Thin top border to visually separate the summary from data rows
    ("LINEABOVE", (0, total_row_index), (-1, total_row_index), 0.8, colors.black),

    # ── Subscription ID row (last row) ───────────────────────────────────────
    ("SPAN",       (0, subscription_row_index), (-1, subscription_row_index)),
    ("ALIGN",      (0, subscription_row_index), (-1, subscription_row_index), "CENTER"),
    ("FONTNAME",   (0, subscription_row_index), (-1, subscription_row_index), "Helvetica-Bold"),
    ("BACKGROUND", (0, subscription_row_index), (-1, subscription_row_index), colors.white),
]))
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=landscape(A3),
            leftMargin=20, rightMargin=20, topMargin=30, bottomMargin=20
        )

        doc.build([title, Spacer(1, 15), table])
        pdf_paths.append(pdf_path)

    return pdf_paths
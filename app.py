from flask import Flask, request, render_template, send_from_directory
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from pathlib import Path
from datetime import datetime
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("excel_file")
        username = request.form.get("username")

        if not file or not username:
            return "❌ Please upload a file and enter username.", 400
            return redirect(url_for("index"))

        try:
            xl = pd.ExcelFile(file)
            df_list = xl.parse("Client_List")
            df_data = xl.parse("Data")
        except Exception as e:
            return f"❌ Error reading Excel file: {e}", 500
            return redirect(url_for("index"))

        base_path = Path(
            fr"C:/Users/{username}/OneDrive - Shree Plan Your Journey Pvt Ltd/Test Excel/Test"
        )

        count = 0
        for _, row in df_list.iterrows():
            try:
                main_folder = str(row.iloc[0]).strip()
                sub_folder = str(row.iloc[1]).strip()
                date_value = row.iloc[2]
            except:
                continue

            if not main_folder or not sub_folder:
                continue

            try:
                month_name = pd.to_datetime(date_value).strftime('%B')
            except:
                month_name = "Unknown"

            folder_path = base_path / main_folder / sub_folder / month_name
            folder_path.mkdir(parents=True, exist_ok=True)

            client_name = sub_folder
            clean_name = ''.join(c for c in client_name if c not in r'[]:*?/\\')[:31] or "Sheet"
            df_client = df_data[df_data.iloc[:, 0].astype(str).str.strip() == client_name]

            if df_client.empty:
                continue

            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"

            max_col = df_data.shape[1]

            # Styles
            bold_font = Font(bold=True)
            center_align = Alignment(horizontal="center", vertical="center")
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))

            # Header Row 1: Client Name
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
            ws.cell(row=1, column=1, value=f"{client_name.upper()}").font = Font(size=14, bold=True)
            ws.cell(row=1, column=1).alignment = center_align

            # Header Row 2: Dynamic Date Range from Booking Date
            booking_col = None
            for col in df_client.columns:
                if "booking" in col.lower() and "date" in col.lower():
                    booking_col = col
                    break

            if booking_col:
                try:
                    df_client[booking_col] = pd.to_datetime(df_client[booking_col], errors='coerce')
                    min_dt = df_client[booking_col].min()
                    max_dt = df_client[booking_col].max()

                    if pd.notnull(min_dt) and pd.notnull(max_dt):
                        first_day = min_dt.replace(day=1)
                        last_day = max_dt.replace(day=1) + pd.offsets.MonthEnd(0)
                        date_range = f"STATEMENT ON {first_day.strftime('%d/%m/%Y')} TO {last_day.strftime('%d/%m/%Y')}"
                    else:
                        date_range = "STATEMENT DATE UNKNOWN"
                except:
                    date_range = "STATEMENT DATE UNKNOWN"
            else:
                date_range = "STATEMENT DATE NOT FOUND"

            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
            ws.cell(row=2, column=1, value=date_range).font = Font(size=12, bold=True)
            ws.cell(row=2, column=1).alignment = center_align

            # Column Headers (Row 3)
            ws.append(df_data.columns.tolist())
            for cell in ws[3]:
                cell.font = bold_font
                cell.alignment = center_align
                cell.border = thin_border
                cell.fill = PatternFill("solid", fgColor="FFFFCC")

            # Data Rows starting from row 4
            for idx, record in enumerate(df_client.itertuples(index=False), start=4):
                row_data = list(record)
                ws.append(row_data)

                for col_index, cell in enumerate(ws[idx], start=1):
                    cell.border = thin_border
                    cell.alignment = center_align

                    # Format Booking Date column to show only the date
                    if booking_col and df_client.columns[col_index - 1] == booking_col:
                        cell.number_format = 'DD/MM/YYYY'

            # Total Row
            total_row_idx = ws.max_row + 1
            ws.cell(row=total_row_idx, column=max_col - 1, value="TOTAL")
            ws.cell(row=total_row_idx, column=max_col - 1).fill = yellow_fill
            ws.cell(row=total_row_idx, column=max_col - 1).font = bold_font
            ws.cell(row=total_row_idx, column=max_col - 1).alignment = center_align

            # Auto Sum last column (assumes last column has numeric values to total)
            g_total_col_letter = chr(64 + max_col)
            ws.cell(row=total_row_idx, column=max_col).value = f"=SUM({g_total_col_letter}4:{g_total_col_letter}{total_row_idx - 1})"
            ws.cell(row=total_row_idx, column=max_col).fill = yellow_fill
            ws.cell(row=total_row_idx, column=max_col).font = bold_font
            ws.cell(row=total_row_idx, column=max_col).alignment = center_align

            # Save Excel
            wb.save(folder_path / f"{clean_name}.xlsx")
            wb.close()
            count += 1

        return f"✅ {count} file(s) created successfully in '{base_path}'"

    return render_template("upload.html")

@app.route("/output/<path:filename>")
def download_file(filename):
    return send_from_directory(base_path, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

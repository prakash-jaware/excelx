from flask import Flask, request, render_template, send_from_directory
import pandas as pd
from openpyxl import Workbook
from pathlib import Path
from datetime import datetime
import os
import tempfile

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()
OUTPUT_FOLDER = Path.home()
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("excel_file")
        if not file:
            return "No file uploaded", 400

        try:
            xl = pd.ExcelFile(file)
            df_list = xl.parse("Client_List")
            df_data = xl.parse("Data")
        except Exception as e:
            return f"Error reading Excel file: {e}", 500

        count = 0
        for _, row in df_list.iterrows():
            try:
                main_folder = str(row.iloc[0]).strip()
                sub_folder = str(row.iloc[1]).strip()
                date_value = row.iloc[2]
            except:
                continue

            if not main_folder:
                continue

            try:
                month_name = pd.to_datetime(date_value).strftime('%B')
            except:
                month_name = "Unknown"

            folder_path = OUTPUT_FOLDER / main_folder / month_name / sub_folder
            folder_path.mkdir(parents=True, exist_ok=True)

            client_name = sub_folder
            clean_name = ''.join(c for c in client_name if c not in r'[]:*?/\\')[:31] or "Sheet"
            df_client = df_data[df_data.iloc[:, 0].astype(str).str.strip() == client_name]

            if not df_client.empty:
                wb = Workbook()
                ws = wb.active
                ws.title = clean_name
                ws.append(df_data.columns.tolist())
                for record in df_client.itertuples(index=False):
                    ws.append(list(record))
                wb.save(folder_path / f"{clean_name}.xlsx")
                wb.close()
                count += 1

        return f"✅ {count} file(s) saved in '{OUTPUT_FOLDER}'"

    return render_template("upload.html")


@app.route("/output/<path:filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

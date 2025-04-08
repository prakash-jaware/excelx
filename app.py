from flask import Flask, request, render_template, send_file
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
import zipfile
import tempfile
import os

app = Flask(__name__)

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

        # Use user-specific temp directory
        user_output_folder = Path(tempfile.gettempdir()) / "excel_outputs"
        output_folder = user_output_folder / "output"
        output_folder.mkdir(parents=True, exist_ok=True)

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

            folder_path = output_folder / main_folder / month_name
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

        # Create ZIP
        zip_path = user_output_folder / "processed_output.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, _, files in os.walk(output_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_folder)
                    zipf.write(file_path, arcname)

        return render_template("upload.html", message=f"✅ {count} file(s) saved!", download_link="/download")

    return render_template("upload.html")

@app.route("/download")
def download():
    zip_path = Path(tempfile.gettempdir()) / "excel_outputs" / "processed_output.zip"
    return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

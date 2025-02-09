from flask import Flask, request, render_template, send_file, send_from_directory
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

app = Flask(__name__, static_folder="static", template_folder="templates")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def highlight_matches(file1_path, file2_path, col1_name, col2_name):
    try:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        if col1_name not in df1.columns or col2_name not in df2.columns:
            return None, f"Errore: una delle colonne '{col1_name}' o '{col2_name}' non esiste nei file."

        cod_set = set(df1[col1_name].astype(str).str.strip())
        wb = load_workbook(file2_path)
        ws = wb.active
        col_idx = df2.columns.get_loc(col2_name) + 1
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for row in range(2, ws.max_row + 1):
            cell_value = str(ws.cell(row=row, column=col_idx).value).strip()
            if cell_value in cod_set:
                ws.cell(row=row, column=col_idx).fill = yellow_fill

        output_path = os.path.join(UPLOAD_FOLDER, "File_2_Highlighted.xlsx")
        wb.save(output_path)
        wb.close()

        return output_path, None
    except Exception as e:
        return None, f"Errore nel confronto dei file: {str(e)}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        file1 = request.files["file1"]
        file2 = request.files["file2"]
        col1_name = request.form["col1_name"].strip()
        col2_name = request.form["col2_name"].strip()

        if not file1 or not file2 or not col1_name or not col2_name:
            return "Errore: Carica entrambi i file e specifica le colonne", 400

        file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
        file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)

        file1.save(file1_path)
        file2.save(file2_path)

        output_path, error = highlight_matches(file1_path, file2_path, col1_name, col2_name)

        if error:
            return error, 400

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return f"Errore durante l'upload dei file: {str(e)}", 500

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory(app.static_folder, filename)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
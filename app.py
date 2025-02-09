from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

app = Flask(__name__)

# Cartella per il caricamento dei file
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def home():
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

        # Percorsi dei file caricati
        file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
        file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)

        file1.save(file1_path)
        file2.save(file2_path)

        # Evidenziare le corrispondenze
        output_path, error = highlight_matches(file1_path, file2_path, col1_name, col2_name)

        if error:
            return error, 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Errore nel confronto: {str(e)}", 500

def highlight_matches(file1_path, file2_path, col1_name, col2_name):
    try:
        # Carica i file Excel
        df1 = pd.read_excel(file1_path, dtype=str)
        df2 = pd.read_excel(file2_path, dtype=str)

        # Verifica che le colonne esistano
        if col1_name not in df1.columns or col2_name not in df2.columns:
            return None, f"Errore: Colonne non trovate nei file"

        # Set con i valori della prima colonna
        cod_set = set(df1[col1_name].dropna().str.strip())

        # Carica il file Excel per evidenziare le corrispondenze
        wb = load_workbook(file2_path)
        ws = wb.active

        # Colore per evidenziare (giallo)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Trova l'indice della colonna corretta
        col_idx = df2.columns.get_loc(col2_name) + 1

        # Applica la colorazione
        for row in range(2, ws.max_row + 1):  # Salta l'intestazione
            cell_value = str(ws.cell(row=row, column=col_idx).value).strip()
            if cell_value in cod_set:
                ws.cell(row=row, column=col_idx).fill = yellow_fill

        # Salva il nuovo file con le corrispondenze evidenziate
        output_path = os.path.join(UPLOAD_FOLDER, "File_2_Highlighted.xlsx")
        wb.save(output_path)
        wb.close()

        return output_path, None

    except Exception as e:
        return None, f"Errore nel confronto dei file: {str(e)}"

if __name__ == "__main__":
    app.run(debug=True)
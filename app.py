from flask import Flask, request, render_template, send_file, jsonify
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        # Recupera i file e i nomi delle colonne
        file1 = request.files["file1"]
        file2 = request.files["file2"]
        col1_name = request.form["col1_name"].strip()
        col2_name = request.form["col2_name"].strip()

        # Controlla che i file e i nomi delle colonne siano stati forniti
        if not file1 or not file2 or not col1_name or not col2_name:
            return jsonify({"error": "Errore: Carica entrambi i file e specifica le colonne"}), 400

        # Salva i file temporaneamente
        file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
        file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)

        file1.save(file1_path)
        file2.save(file2_path)

        # Esegui il confronto e genera il file evidenziato
        output_path, error_message = highlight_matches(file1_path, file2_path, col1_name, col2_name)

        if error_message:
            return jsonify({"error": error_message}), 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": f"Errore nel confronto dei file: {str(e)}"}), 500

def highlight_matches(file1_path, file2_path, col1_name, col2_name):
    try:
        df1 = pd.read_excel(file1_path, dtype=str)
        df2 = pd.read_excel(file2_path, dtype=str)

        if col1_name not in df1.columns or col2_name not in df2.columns:
            return None, f"Errore: le colonne '{col1_name}' o '{col2_name}' non esistono nei file caricati."

        # Ottieni i valori unici della colonna nel primo file
        cod_set = set(df1[col1_name].dropna().str.strip())

        # Apri il secondo file con OpenPyXL per evidenziare le corrispondenze
        wb = load_workbook(file2_path)
        ws = wb.active

        # Trova l'indice della colonna cercata
        col_idx = None
        for col in range(1, ws.max_column + 1):
            if ws.cell(row=1, column=col).value == col2_name:
                col_idx = col
                break

        if col_idx is None:
            return None, f"Errore: la colonna '{col2_name}' non è stata trovata nel file 2."

        # Definisci il colore giallo per evidenziare
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Scansiona la colonna e applica il colore alle corrispondenze
        for row in range(2, ws.max_row + 1):
            cell_value = str(ws.cell(row=row, column=col_idx).value).strip()
            if cell_value in cod_set:
                ws.cell(row=row, column=col_idx).fill = yellow_fill

        # Salva il file modificato
        output_path = os.path.join(UPLOAD_FOLDER, "File_2_Highlighted.xlsx")
        wb.save(output_path)
        wb.close()

        return output_path, None

    except Exception as e:
        return None, f"Errore durante l'elaborazione: {str(e)}"

if __name__ == "__main__":
    app.run(debug=True)

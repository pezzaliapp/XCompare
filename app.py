from flask import Flask, render_template, request, send_file
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def highlight_matches(file1_path, file2_path, col1_name, col2_name):
    # Carica i file Excel
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
    
    # Normalizza i nomi delle colonne
    df1.columns = df1.columns.str.lower().str.strip()
    df2.columns = df2.columns.str.lower().str.strip()
    col1_name = col1_name.lower().strip()
    col2_name = col2_name.lower().strip()
    
    # Debug: Mostra le colonne disponibili
    print("Colonne disponibili nel File 1:", df1.columns.tolist())
    print("Colonne disponibili nel File 2:", df2.columns.tolist())
    
    # Verifica che le colonne esistano
    if col1_name not in df1.columns or col2_name not in df2.columns:
        return None, f"Colonna '{col1_name}' o '{col2_name}' non trovata nei file."
    
    # Set dei valori nella colonna del primo file
    cod_set = set(df1[col1_name].astype(str).str.strip())
    
    # Apri il file 2 con openpyxl per evidenziare
    wb = load_workbook(file2_path)
    ws = wb.active
    col_idx = list(df2.columns).index(col2_name) + 1  # Indice della colonna in Excel (1-based)
    
    # Definiamo il colore giallo per evidenziare
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    # Scansioniamo la colonna e applichiamo il colore alle corrispondenze
    for row in range(2, ws.max_row + 1):  # Partiamo dalla seconda riga per ignorare l'intestazione
        cell_value = str(ws.cell(row=row, column=col_idx).value).strip()
        if cell_value in cod_set:
            ws.cell(row=row, column=col_idx).fill = yellow_fill
    
    output_path = os.path.join(UPLOAD_FOLDER, "File_2_Highlighted.xlsx")
    wb.save(output_path)
    wb.close()  # Chiude il file per garantire che non ci siano collegamenti attivi
    
    return output_path, None

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_files():
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

if __name__ == "__main__":
    app.run(debug=True)

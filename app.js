document.getElementById("compareBtn").addEventListener("click", async () => {
    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];
    
    const colName1 = document.getElementById("columnName1").value.trim();
    const colName2 = document.getElementById("columnName2").value.trim();

    if (!file1 || !file2) {
        alert("Carica entrambi i file Excel prima di procedere.");
        return;
    }

    if (!colName1 || !colName2) {
        alert("Inserisci il nome delle colonne da confrontare.");
        return;
    }
    
    try {
        const [data1, data2] = await Promise.all([readExcel(file1), readExcel(file2)]);
        compareData(data1, data2, colName1, colName2);
    } catch (err) {
        console.error("Errore nella lettura dei file:", err);
        alert("C’è stato un errore nella lettura dei file Excel.");
    }
});

async function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            // Legge il workbook
            const workbook = XLSX.read(data, { type: "array" });
            // Prende il PRIMO foglio (se ce ne sono più di uno, adattare)
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            // Converte in array di array (header = 1 significa riga 0 come array "puro")
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            resolve(jsonData);
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

function compareData(data1, data2, colName1, colName2) {
    // data1[0] e data2[0] contengono gli header (nomi colonne)
    const header1 = data1[0] || [];
    const header2 = data2[0] || [];

    // Troviamo gli indici delle colonne, cercando i nomi che l’utente ha digitato
    const idx1 = header1.indexOf(colName1);
    const idx2 = header2.indexOf(colName2);

    if (idx1 === -1) {
        alert(`Colonna "${colName1}" non trovata nel primo file.`);
        return;
    }
    if (idx2 === -1) {
        alert(`Colonna "${colName2}" non trovata nel secondo file.`);
        return;
    }

    // Creiamo un set dei valori della colonna scelta del primo file
    // Partiamo dalla riga 1 (saltiamo l’header) fino alla fine
    const setFile1 = new Set();
    for (let i = 1; i < data1.length; i++) {
        const row = data1[i];
        if (row && row[idx1] !== undefined) {
            setFile1.add(row[idx1]);
        }
    }

    // Ora scorriamo il secondo file (dalla riga 1 in poi) e, se troviamo corrispondenza, evidenziamo
    for (let j = 1; j < data2.length; j++) {
        const row = data2[j];
        if (!row) continue;

        const value2 = row[idx2];
        // Se corrisponde, evidenziamo
        if (setFile1.has(value2)) {
            row[idx2] = `**${value2}**`;
        }
    }

    // Generiamo un nuovo Excel dal file 2 con colonna evidenziata
    generateExcel(data2);
}

function generateExcel(data) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Risultato");
    // Scarichiamo il file con la colonna 2 evidenziata
    XLSX.writeFile(wb, "File_2_Highlighted.xlsx");
}

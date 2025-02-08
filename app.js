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
    alert("Inserisci il nome delle colonne da confrontare in entrambi i campi.");
    return;
  }

  try {
    const [data1, data2] = await Promise.all([readExcel(file1), readExcel(file2)]);
    addMatchColumn(data1, data2, colName1, colName2);
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
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]]; // Assicurati di leggere il foglio giusto
      
      const jsonData = XLSX.utils.sheet_to_json(sheet, {
        header: 1,   
        blankrows: true,  // Mantiene righe vuote
        defval: ""  // Mantiene le celle vuote invece di eliminarle
      });

      console.log("Numero di righe lette:", jsonData.length); // DEBUG
      resolve(jsonData);
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

function addMatchColumn(data1, data2, colName1, colName2) {
  const header1 = data1[0] || [];
  const header2 = data2[0] || [];

  const idx1 = findColumnIndex(header1, colName1);
  const idx2 = findColumnIndex(header2, colName2);

  if (idx1 === -1) {
    alert(`Colonna "${colName1}" non trovata nel primo file.`);
    return;
  }
  if (idx2 === -1) {
    alert(`Colonna "${colName2}" non trovata nel secondo file.`);
    return;
  }

  const setFile1 = new Set();
  for (let i = 1; i < data1.length; i++) {
    const row = data1[i];
    if (row && row[idx1] !== undefined) {
      setFile1.add(row[idx1].toString().trim().toLowerCase());
    }
  }

  // **Aggiungere una nuova colonna accanto alla colonna di confronto**
  header2.push("MATCH"); // Aggiunge l'intestazione per la nuova colonna
  for (let j = 1; j < data2.length; j++) {
    const row = data2[j];
    if (!row) {
      row = new Array(header2.length).fill(""); // Mantieni la riga vuota
    }

    const val = row[idx2] ? row[idx2].toString().trim().toLowerCase() : "";
    row.push(setFile1.has(val) ? "MATCH" : "");  // Nuova colonna con "MATCH" o vuoto
  }

  console.log("Numero di righe finali:", data2.length); // DEBUG

  generateExcel(data2);
}

function findColumnIndex(headerRow, userColName) {
  const normalizedHeaderRow = headerRow.map(x => (x || "").toLowerCase().trim());
  const normalizedColName = userColName.toLowerCase().trim();
  return normalizedHeaderRow.indexOf(normalizedColName);
}

function generateExcel(data) {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Risultato");

  XLSX.writeFile(wb, "File_2_Highlighted.xlsx");
}

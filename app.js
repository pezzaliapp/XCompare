document.getElementById("compareBtn").addEventListener("click", async () => {
    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];
    
    if (!file1 || !file2) {
        alert("Carica entrambi i file Excel prima di procedere.");
        return;
    }
    
    const [data1, data2] = await Promise.all([readExcel(file1), readExcel(file2)]);
    compareData(data1, data2);
});

async function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            resolve(jsonData);
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

function compareData(data1, data2) {
    const column1 = data1.map(row => row[0]);
    const column2 = data2.map(row => row[10]); // Colonna 11 nel file 2 (indice 10)
    
    const highlightedData = data2.map(row => {
        if (column1.includes(row[10])) {
            row[10] = `**${row[10]}**`; // Simulazione evidenziazione (da gestire in Excel)
        }
        return row;
    });
    
    generateExcel(highlightedData);
}

function generateExcel(data) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Risultato");
    XLSX.writeFile(wb, "File_2_Highlighted.xlsx");
}

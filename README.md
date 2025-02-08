# XCompare - PWA per il Confronto di File Excel

## Descrizione
**XCompare** è un'applicazione web progressiva (PWA) progettata per confrontare due colonne di dati provenienti da file Excel diversi. L'app evidenzia in giallo nel secondo file tutte le corrispondenze presenti nel primo file e genera un file Excel aggiornato.

## Funzionalità
- **Caricamento di due file Excel (.xlsx)**
- **Confronto automatico tra le colonne selezionate**
- **Evidenziazione delle corrispondenze nel file 2**
- **Download del file Excel aggiornato**
- **Compatibilità mobile e desktop**
- **Funzionamento offline grazie al service worker**

## Struttura del Progetto
- `index.html` - Interfaccia principale dell'app
- `styles.css` - Stile grafico della PWA
- `app.js` - Logica di confronto e manipolazione dei file Excel
- `manifest.json` - Configurazione della PWA per l'installazione
- `service-worker.js` - Gestione della cache e funzionamento offline
- `XTM-192.png`, `XTM-512.png` - Icone della PWA

## Installazione
1. Clona o scarica il repository
2. Apri `index.html` in un browser supportato
3. Per attivare il service worker, assicurati di ospitare l'app su un server (es. con `Live Server` in VS Code)

## Tecnologie Utilizzate
- HTML, CSS, JavaScript (Vanilla JS)
- SheetJS (XLSX) per la gestione dei file Excel
- OpenPyXL per la generazione del file aggiornato (opzionale per backend Python)

## Utilizzo
1. **Carica i due file Excel** utilizzando i pulsanti di upload
2. **Clicca sul pulsante "Confronta"** per elaborare i dati
3. **Scarica il file aggiornato** con le corrispondenze evidenziate

## Requisiti
- Browser moderno con supporto a JavaScript
- Connessione a internet per il primo utilizzo (poi può funzionare offline)

## Autore
Sviluppato da Alessandro Pezzali per semplificare il confronto di dati in Excel tramite una PWA intuitiva e veloce.

## Licenza
Distribuito sotto licenza MIT.


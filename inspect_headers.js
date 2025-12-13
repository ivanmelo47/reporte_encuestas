
const XLSX = require("xlsx");
const path = require("path");
const fs = require('fs');

const FILES_TO_INSPECT = [
    'estadisticas_encuesta_3_Princess.xlsx',
    'estadisticas_encuesta_2_Palacio.xlsx'
];

// Helper to get formatted value
function getVal(row, index) {
    return row && row[index] !== undefined ? row[index] : '[Empty]';
}

async function processFile(filePath) {
    console.log(`Reading file: ${filePath}`);
    
    if (!fs.existsSync(filePath)) {
        console.log("File not found");
        return;
    }

    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = 'Worksheet'; // Assuming data is here
        if (!workbook.Sheets[sheetName]) {
            console.log(`Sheet '${sheetName}' not found`);
            return;
        }

        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (json.length > 12) {
            const headerRow = json[10]; // Row 11 (Index 10)
            const dataRow = json[11];   // Row 12 (Index 11) - Data Start
            
            console.log("Header Mapping (First 12 columns):");
            // Inspect first 12 columns
            for (let i = 0; i < 12; i++) {
                const hVal = getVal(headerRow, i);
                const dVal = getVal(dataRow, i);
                console.log(`  Index ${i}: Header='${hVal}' | Data='${dVal}'`);
            }
        } else {
            console.log("Not enough rows to inspect.");
        }

    } catch (e) {
        console.error(`Error processing ${filePath}: ${e.message}`);
    }
}

async function inspectAll() {
    for (const file of FILES_TO_INSPECT) {
        const filePath = path.join(__dirname, file);
        console.log(`\n\n========== INSPECTING: ${file} ==========`);
        await processFile(filePath);
    }
}

inspectAll();

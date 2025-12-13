
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const INPUT_FILE = path.join(__dirname, 'estadisticas_encuesta_3_Princess.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'analisis_encuesta_princess.xlsx');
const SHEET_NAME = 'Worksheet';
const HEADER_ROW_INDEX = 8; // Row 9 is index 8 (0-based)

function analyzeSurvey() {
    console.log(`Starting analysis...`);
    console.log(`Reading input file: ${INPUT_FILE}`);

    if (!fs.existsSync(INPUT_FILE)) {
        console.error(`Error: File not found: ${INPUT_FILE}`);
        return;
    }

    const workbook = XLSX.readFile(INPUT_FILE);
    if (!workbook.Sheets[SHEET_NAME]) {
        console.error(`Error: Sheet "${SHEET_NAME}" not found. Available sheets: ${workbook.SheetNames.join(', ')}`);
        return;
    }

    const worksheet = workbook.Sheets[SHEET_NAME];
    
    // Read all data as JSON (array of arrays to handle headers manually)
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (data.length <= HEADER_ROW_INDEX) {
        console.error("Error: Not enough rows for header.");
        return;
    }

    const headers = data[HEADER_ROW_INDEX];
    console.log(`Found ${headers.length} columns in header row.`);

    // Extract data rows (rows after header)
    const rawDataRows = data.slice(HEADER_ROW_INDEX + 1);
    console.log(`Processing ${rawDataRows.length} data rows.`);

    const statistics = [];
    const detailedDistributions = [];

    // Iterate over each column identified in the header
    headers.forEach((header, colIndex) => {
        if (!header) return; // Skip empty headers

        const columnValues = [];
        rawDataRows.forEach(row => {
            // Row might be shorter than header length, check index
            if (row[colIndex] !== undefined && row[colIndex] !== null && row[colIndex] !== '') {
                columnValues.push(row[colIndex]);
            }
        });

        const totalResponses = columnValues.length;
        
        // Count frequencies
        const frequencyMap = {};
        columnValues.forEach(val => {
            const key = String(val).trim();
            frequencyMap[key] = (frequencyMap[key] || 0) + 1;
        });

        // Find top answer
        let topAnswer = 'N/A';
        let topCount = 0;
        Object.entries(frequencyMap).forEach(([answer, count]) => {
            if (count > topCount) {
                topCount = count;
                topAnswer = answer;
            }
        });

        const uniqueValuesCount = Object.keys(frequencyMap).length;

        // Add to summary stats
        statistics.push({
            'Pregunta/Código': header,
            'Total Respuestas': totalResponses,
            'Valores Únicos': uniqueValuesCount,
            'Respuesta Más Común': topAnswer,
            'Frecuencia Top': topCount
        });

        // Prepare detailed distribution data
        // We'll format this later for a separate sheet or structure
        detailedDistributions.push({
            header: header,
            distribution: frequencyMap
        });
    });

    console.log(`Analyzed ${statistics.length} columns.`);

    // --- Generate Output Excel ---
    const newWorkbook = XLSX.utils.book_new();

    // 1. Summary Sheet
    const summarySheet = XLSX.utils.json_to_sheet(statistics);
    XLSX.utils.book_append_sheet(newWorkbook, summarySheet, "Resumen General");

    // 2. Details Sheet
    // Transform detailed distributions into a readable format for Excel
    // Format: Question | Answer | Count
    const detailsRows = [];
    detailedDistributions.forEach(item => {
        Object.entries(item.distribution).forEach(([answer, count]) => {
            detailsRows.push({
                'Pregunta': item.header,
                'Respuesta': answer,
                'Cantidad': count,
                'Porcentaje': (count / statistics.find(s => s['Pregunta/Código'] === item.header)['Total Respuestas'] * 100).toFixed(2) + '%'
            });
        });
    });

    const detailsSheet = XLSX.utils.json_to_sheet(detailsRows);
    XLSX.utils.book_append_sheet(newWorkbook, detailsSheet, "Detalle Frecuencias");

    XLSX.writeFile(newWorkbook, OUTPUT_FILE);
    console.log(`Analysis complete. Output saved to: ${OUTPUT_FILE}`);
}

analyzeSurvey();

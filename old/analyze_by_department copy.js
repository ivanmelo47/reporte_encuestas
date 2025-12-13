
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const INPUT_FILE = path.join(__dirname, 'estadisticas_encuesta_3_Princess.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'analisis_departamentos.xlsx');
const SHEET_NAME = 'Worksheet';

const HEADER_ROW_INDEX = 10; // Row 11 is index 10
const DATA_START_INDEX = 11; // Data starts at Row 12 (Index 11)
const COL_DEPT_INDEX = 9;    // Column J is index 9 (0-based: A=0... J=9)
const COL_QUESTIONS_START = 15; // Column P is index 15

const SCORE_MAP = {
    'SIEMPRE': 4,
    'CASI SIEMPRE': 3,
    'ALGUNAS VECES': 2,
    'CASI NUNCA': 1,
    'NUNCA': 0,
    'NUCA': 0 // Handling typo
};

function getScore(value) {
    if (typeof value !== 'string') return null;
    const cleanVal = value.trim().toUpperCase();
    if (SCORE_MAP.hasOwnProperty(cleanVal)) {
        return SCORE_MAP[cleanVal];
    }
    return null; // Not a valid scored answer (e.g. empty or other text)
}

function analyzeByDepartment() {
    console.log(`Starting Department Analysis...`);
    console.log(`Reading input file: ${INPUT_FILE}`);

    if (!fs.existsSync(INPUT_FILE)) {
        console.error(`Error: File not found: ${INPUT_FILE}`);
        return;
    }

    const workbook = XLSX.readFile(INPUT_FILE);
    if (!workbook.Sheets[SHEET_NAME]) {
        console.error(`Error: Sheet "${SHEET_NAME}" not found.`);
        return;
    }

    const worksheet = workbook.Sheets[SHEET_NAME];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (data.length <= HEADER_ROW_INDEX) {
        console.error("Error: Not enough rows for header.");
        return;
    }

    const headers = data[HEADER_ROW_INDEX];
    console.log(`Headers found at row ${HEADER_ROW_INDEX + 1}. Total columns: ${headers.length}`);

    // Group rows by Department
    const departments = {};
    let processedRows = 0;

    for (let i = DATA_START_INDEX; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const deptName = row[COL_DEPT_INDEX] ? String(row[COL_DEPT_INDEX]).trim() : 'Sin Departamento';
        
        if (!departments[deptName]) {
            departments[deptName] = [];
        }
        departments[deptName].push(row);
        processedRows++;
    }

    console.log(`Processed ${processedRows} rows. Found ${Object.keys(departments).length} departments.`);

    // Analyze each department
    const newWorkbook = XLSX.utils.book_new();

    Object.keys(departments).forEach(dept => {
        const rows = departments[dept];
        const deptStats = [];

        // Iterate through questions columns
        for (let col = COL_QUESTIONS_START; col < headers.length; col++) {
            const questionText = headers[col];
            if (!questionText) continue; // Skip empty headers

            let totalScore = 0;
            let validResponses = 0;
            const distribution = {
                'Siempre': 0,
                'Casi siempre': 0,
                'Algunas veces': 0,
                'Casi nunca': 0,
                'Nunca': 0
            };

            rows.forEach(row => {
                const val = row[col];
                const score = getScore(val);
                
                if (score !== null) {
                    totalScore += score;
                    validResponses++;
                    
                    // Update distribution count
                    // We match back to Title Case keys for the report
                    const upperVal = String(val).trim().toUpperCase().replace('NUCA', 'NUNCA');
                    const displayKey = upperVal.charAt(0) + upperVal.slice(1).toLowerCase(); 
                    // Manual mapping to ensure cleaner keys if needed, but Title Case usually works
                    // Let's rely on exact map keys for clean output
                    if (upperVal === 'SIEMPRE') distribution['Siempre']++;
                    else if (upperVal === 'CASI SIEMPRE') distribution['Casi siempre']++;
                    else if (upperVal === 'ALGUNAS VECES') distribution['Algunas veces']++;
                    else if (upperVal === 'CASI NUNCA') distribution['Casi nunca']++;
                    else if (upperVal === 'NUNCA') distribution['Nunca']++;
                }
            });

            if (validResponses > 0) {
                const avgScore = totalScore / validResponses;
                const score100 = (avgScore / 4) * 100;

                deptStats.push({
                    'Pregunta': questionText,
                    'Nivel Promedio': avgScore.toFixed(2),
                    'Calificación (0-100)': score100.toFixed(2),
                    'Total Respuestas': validResponses,
                    'Siempre': distribution['Siempre'],
                    'Casi siempre': distribution['Casi siempre'],
                    'Algunas veces': distribution['Algunas veces'],
                    'Casi nunca': distribution['Casi nunca'],
                    'Nunca': distribution['Nunca']
                });
            }
        }

        if (deptStats.length > 0) {
            const ws = XLSX.utils.json_to_sheet(deptStats);
            // Excel sheet names max 31 chars
            const safeSheetName = dept.replace(/[\*\?\[\]\:\/\\]/g, '').substring(0, 31) || "Dept";
            XLSX.utils.book_append_sheet(newWorkbook, ws, safeSheetName);
        }
    });

    XLSX.writeFile(newWorkbook, OUTPUT_FILE);
    console.log(`Analysis complete. Saved to: ${OUTPUT_FILE}`);
}

analyzeByDepartment();

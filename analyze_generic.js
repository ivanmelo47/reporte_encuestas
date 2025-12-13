
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILES_CONFIG = [
    { input: 'estadisticas_encuesta_2_Palacio.xlsx', sheetName: 'Palacio' },
    { input: 'Estadisticas_encuesta_1_Pierre.xlsx', sheetName: 'Pierre' }
];

const SHEET_DATA_NAME = 'Worksheet'; 
const HEADER_ROW_INDEX = 10; 
const DATA_START_INDEX = 11; 

// GENERIC INDICES (Based on debug_log.txt for Palacio/Pierre)
// No mysterious shift, just explicit alignment:
const COL_GENDER = 4;
const COL_AGE = 5;
const COL_CIVIL = 6;
const COL_SCHOOL = 7;
// Propiedad does not exist
const COL_DEPT = 8;      
const COL_TYPE = 9;      
const COL_TENURE = 12;   // 12. Tiempo en el puesto actual

const COL_QUESTIONS_START = 14; // 14. El espacio donde trabajo...

const DEMO_COLS = {
    'Género': COL_GENDER,
    'Edad': COL_AGE,
    'Estado Civil': COL_CIVIL,
    'Nivel de Estudios': COL_SCHOOL,
    'Tipo de Puesto': COL_TYPE,  
    'Tiempo en Puesto': COL_TENURE 
};

const SCORE_MAP = {
    'SIEMPRE': 4,
    'CASI SIEMPRE': 3,
    'ALGUNAS VECES': 2,
    'CASI NUNCA': 1,
    'NUNCA': 0,
    'NUCA': 0
};

function getScore(value) {
    if (typeof value !== 'string') return null;
    const cleanVal = value.trim().toUpperCase();
    if (SCORE_MAP.hasOwnProperty(cleanVal)) {
        return SCORE_MAP[cleanVal];
    }
    return null;
}

function analyzeDemographics(rows) {
    const stats = {};
    Object.keys(DEMO_COLS).forEach(key => {
        stats[key] = { total: 0, distinct: {} };
    });

    rows.forEach(row => {
        Object.keys(DEMO_COLS).forEach(key => {
            const colIndex = DEMO_COLS[key];
            const val = row[colIndex];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
                const cleanVal = String(val).trim();
                stats[key].total++;
                stats[key].distinct[cleanVal] = (stats[key].distinct[cleanVal] || 0) + 1;
            }
        });
    });

    const resultRows = [];
    Object.keys(stats).forEach(key => {
        resultRows.push([key.toUpperCase(), '', '']);
        resultRows.push(['Opción', 'Cantidad', 'Porcentaje']);
        const total = stats[key].total;
        Object.entries(stats[key].distinct).forEach(([option, count]) => {
            const pct = total > 0 ? (count / total * 100).toFixed(2) + '%' : '0%';
            resultRows.push([option, count, pct]);
        });
        resultRows.push(['', '', '']);
    });
    return resultRows;
}


function analyzeQuestions(rows, headers) {
    const stats = [];
    
    for (let col = COL_QUESTIONS_START; col < headers.length; col++) {
        const questionText = headers[col];
        if (!questionText) continue;

        let totalScore = 0;
        let validResponses = 0;
        const distribution = { 'Siempre': 0, 'Casi siempre': 0, 'Algunas veces': 0, 'Casi nunca': 0, 'Nunca': 0 };

        rows.forEach(row => {
            const val = row[col];
            const score = getScore(val);
            if (score !== null) {
                totalScore += score;
                validResponses++;
                const upperVal = String(val).trim().toUpperCase().replace('NUCA', 'NUNCA');
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
            stats.push({
                q: questionText,
                avg: avgScore.toFixed(2),
                score100: score100.toFixed(2),
                total: validResponses,
                d1: distribution['Siempre'], d2: distribution['Casi siempre'], d3: distribution['Algunas veces'], d4: distribution['Casi nunca'], d5: distribution['Nunca']
            });
        }
    }
    return stats;
}

function processFile(filePath, config) {
    console.log(`Processing: ${config.sheetName}`);
    if (!fs.existsSync(filePath)) {
        console.error("File not found");
        return null;
    }
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[SHEET_DATA_NAME];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headers = data[HEADER_ROW_INDEX];

    // Group by Department
    const departments = {};
    const allRows = [];
    for (let i = DATA_START_INDEX; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        allRows.push(row);
        const deptName = row[COL_DEPT] ? String(row[COL_DEPT]).trim() : 'Sin Departamento';
        if (!departments[deptName]) departments[deptName] = [];
        departments[deptName].push(row);
    }

    const tableHeader = ['Pregunta', 'Nivel Promedio', 'Calificación (0-100)', 'Total Respuestas', 'Siempre', 'Casi siempre', 'Algunas veces', 'Casi nunca', 'Nunca'];

    // 1. Analisis General (All Rows)
    const generalStats = analyzeQuestions(allRows, headers);
    const generalDataRows = [];
    if (generalStats.length > 0) {
        generalDataRows.push(['ANALISIS GENERAL (TODOS LOS DEPARTAMENTOS)', '', '', '', '', '', '', '', '']);
        generalDataRows.push(tableHeader);
        generalStats.forEach(stat => {
            generalDataRows.push([stat.q, stat.avg, stat.score100, stat.total, stat.d1, stat.d2, stat.d3, stat.d4, stat.d5]);
        });
    }
    const wsAnalisisGeneral = XLSX.utils.aoa_to_sheet(generalDataRows);
    wsAnalisisGeneral['!cols'] = [{ wch: 50 }, { wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];


    // 2. Demographics General
    const demoGeneralRows = analyzeDemographics(allRows);
    const wsDemoGeneral = XLSX.utils.aoa_to_sheet(demoGeneralRows);
    wsDemoGeneral['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 15 }];

    // 3. By Dept Demographics
    const demoDeptRows = [];
    Object.keys(departments).forEach(dept => {
        demoDeptRows.push([`DEPARTAMENTO: ${dept.toUpperCase()}`, '', '']);
        const deptDemoStats = analyzeDemographics(departments[dept]);
        deptDemoStats.forEach(r => demoDeptRows.push(r));
        demoDeptRows.push(['', '', '']);
        demoDeptRows.push(['-------------------------', '', '']); 
        demoDeptRows.push(['', '', '']); 
    });
    const wsDemoDept = XLSX.utils.aoa_to_sheet(demoDeptRows);
    wsDemoDept['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 15 }];

    // 4. Scoring (Dept)
    const allDataRows = [];
    Object.keys(departments).forEach(dept => {
        const dRows = departments[dept];
        const deptStats = analyzeQuestions(dRows, headers);

        if (deptStats.length > 0) {
            allDataRows.push(['', '', '', '', '', '', '', '', '']); 
            allDataRows.push([`DEPARTAMENTO: ${dept}`.toUpperCase()]); 
            allDataRows.push(tableHeader);
            deptStats.forEach(stat => {
                allDataRows.push([stat.q, stat.avg, stat.score100, stat.total, stat.d1, stat.d2, stat.d3, stat.d4, stat.d5]);
            });
        }
    });

    let wsMain = null;
    if (allDataRows.length > 0) {
        wsMain = XLSX.utils.aoa_to_sheet(allDataRows);
        wsMain['!cols'] = [{ wch: 50 }, { wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
    }

    return { 
        analisisGeneral: wsAnalisisGeneral,
        main: wsMain, 
        demoGeneral: wsDemoGeneral, 
        demoDept: wsDemoDept 
    };
}

function analyzeGeneric() {
    console.log(`Starting Generic Analysis (Palacio/Pierre)...`);
    const outputDir = path.join(__dirname, 'analisis');
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    FILES_CONFIG.forEach(config => {
        const fullPath = path.join(__dirname, config.input);
        const sheets = processFile(fullPath, config);
        
        if (sheets && sheets.main) {
            const newWorkbook = XLSX.utils.book_new();
            
            XLSX.utils.book_append_sheet(newWorkbook, sheets.analisisGeneral, "Analisis General");
            XLSX.utils.book_append_sheet(newWorkbook, sheets.main, config.sheetName);
            XLSX.utils.book_append_sheet(newWorkbook, sheets.demoGeneral, "Demografía General");
            XLSX.utils.book_append_sheet(newWorkbook, sheets.demoDept, "Demografía Dept");
            
            const outputFileName = `analisis_${config.sheetName.toLowerCase()}.xlsx`;
            XLSX.writeFile(newWorkbook, path.join(outputDir, outputFileName));
            console.log(`Saved: ${outputFileName}`);
        }
    });

    console.log("Generic analysis complete.");
}

analyzeGeneric();

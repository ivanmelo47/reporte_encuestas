
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILE_CONFIG = {
    input: 'estadisticas_encuesta_3_Princess.xlsx',
    sheetName: 'Princess' // Base name
};

const SHEET_DATA_NAME = 'Worksheet'; 

const HEADER_ROW_INDEX = 10; 
const DATA_START_INDEX = 11; 

// PRINCESS INDICES (No Shift relative to header, but includes Propiedad)
const COL_GENDER = 4;
const COL_AGE = 5;
const COL_CIVIL = 6;
const COL_SCHOOL = 7;
const COL_PROPIEDAD = 8; // Row I
const COL_DEPT = 9;      // Row J
const COL_TYPE = 10;     // Job Type
const COL_TENURE = 13;   // Based on assumption, verify? Let's use 13 (usually 11 in generic, shifted? No, if generic is 11... in Princess it's probably 12 or 13.
                        // In Generic (Standard): Tenure is 11.
                        // In Princess: If Propiedad inserted at 8... 0-7 same. 8 is New. 9 is old 8. 10 is old 9. 11 is old 10. 12 is old 11.
                        // So Tenure should be 12?
                        // Let's assume 12 for now.

const COL_QUESTIONS_START = 15; // Column P

const DEMO_COLS = {
    'Género': 4,
    'Edad': 5,
    'Estado Civil': 6,
    'Nivel de Estudios': 7,
    'Tipo de Puesto': 10,  
    'Tiempo en Puesto': 13 // Fixed: Index 13 is Tenure, 12 is Rotation
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

function processGroup(rows, groupName, headers) {
    const tableHeader = ['Pregunta', 'Nivel Promedio', 'Calificación (0-100)', 'Total Respuestas', 'Siempre', 'Casi siempre', 'Algunas veces', 'Casi nunca', 'Nunca'];

    // 1. Analisis General (All Rows)
    const generalStats = analyzeQuestions(rows, headers);
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
    const demoGeneralRows = analyzeDemographics(rows);
    const wsDemoGeneral = XLSX.utils.aoa_to_sheet(demoGeneralRows);
    wsDemoGeneral['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 15 }];

    // 3. By Department (Demographics)
    const departments = {};
    rows.forEach(row => {
        const deptName = row[COL_DEPT] ? String(row[COL_DEPT]).trim() : 'Sin Departamento';
        if (!departments[deptName]) departments[deptName] = [];
        departments[deptName].push(row);
    });

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

    // 4. By Department (Scoring)
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

function analyzePrincess() {
    console.log(`Starting Princess Analysis...`);
    const outputDir = path.join(__dirname, 'analisis');
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

    const fullPath = path.join(__dirname, FILE_CONFIG.input);
    if (!fs.existsSync(fullPath)) {
        console.error("File not found");
        return;
    }

    const workbook = XLSX.readFile(fullPath);
    const sheet = workbook.Sheets[SHEET_DATA_NAME];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headers = data[HEADER_ROW_INDEX];

    // Group by PROPIEDAD
    const properties = {};
    for (let i = DATA_START_INDEX; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        const propName = row[COL_PROPIEDAD] ? String(row[COL_PROPIEDAD]).trim() : 'Desconocida';
        if (!properties[propName]) properties[propName] = [];
        properties[propName].push(row);
    }

    Object.keys(properties).forEach(propName => {
        console.log(`Processing Property group: ${propName}`);
        const rows = properties[propName];
        const sheets = processGroup(rows, propName, headers);

        if (sheets.main) {
            const newWorkbook = XLSX.utils.book_new();
            // Sanitize sheet name
            const safeSheetName = propName.substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
            
            XLSX.utils.book_append_sheet(newWorkbook, sheets.analisisGeneral, "Analisis General");
            XLSX.utils.book_append_sheet(newWorkbook, sheets.main, safeSheetName);
            XLSX.utils.book_append_sheet(newWorkbook, sheets.demoGeneral, "Demografía General");
            XLSX.utils.book_append_sheet(newWorkbook, sheets.demoDept, "Demografía Dept");
            
            const outputFileName = `analisis_princess_${safeSheetName.replace(/\s+/g,'_').toLowerCase()}.xlsx`;
            XLSX.writeFile(newWorkbook, path.join(outputDir, outputFileName));
            console.log(`Saved: ${outputFileName}`);
        }
    });

    console.log("Princess analysis complete.");
}

analyzePrincess();

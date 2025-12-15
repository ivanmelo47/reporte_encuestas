const { INDICES_GENERIC, DEMO_MAP_KEYS, HEADER_ROW_INDEX, DATA_START_INDEX, QUESTION_TYPE_ROW_INDEX } = require('../config/constants');
const Statistics = require('../core/Statistics');
const Demographics = require('../core/Demographics');
const ExcelReader = require('../services/ExcelReader');

class GenericProcessor {
    constructor() {
        // Build demo map
        this.demoMap = {};
        Object.keys(DEMO_MAP_KEYS).forEach(key => {
            const indexKey = DEMO_MAP_KEYS[key];
            this.demoMap[key] = INDICES_GENERIC[indexKey];
        });
        
        this.colWidthsMain = [{ wch: 50 }, { wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
        this.colWidthsDemo = [{ wch: 30 }, { wch: 10 }, { wch: 15 }];
        this.headerMain = ['Pregunta', 'Nivel Promedio', 'Calificación (0-100)', 'Total Respuestas', 'Siempre', 'Casi siempre', 'Algunas veces', 'Casi nunca', 'Nunca'];
    }

    process(filePath, sheetName) {
        // 1. Read
        const data = ExcelReader.read(filePath, 'Worksheet'); // Assuming sheet name is always 'Worksheet' for data
        const headers = data[HEADER_ROW_INDEX];
        const allRows = [];
        const departments = {};

        // 2. Group
        for (let i = DATA_START_INDEX; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            
            // Fix: Exclude Summary/Footer rows by checking for essential metadata (e.g. Gender or Dept)
            // If header row or summary row is read as data, it might lack these columns.
            // INDICES_GENERIC.GENDER = 4
            // Check if Gender column has value. If not, skip.
            if (!row[INDICES_GENERIC.GENDER] && !row[INDICES_GENERIC.DEPT]) continue;

            allRows.push(row);
            
            const deptCol = INDICES_GENERIC.DEPT;
            const deptName = row[deptCol] ? String(row[deptCol]).trim() : 'Sin Departamento';
            if (!departments[deptName]) departments[deptName] = [];
            departments[deptName].push(row);
        }

        // 3. Stats Generation
        
        const questionTypes = data[QUESTION_TYPE_ROW_INDEX];

        // A. Analisis General
        const generalStats = Statistics.analyzeQuestions(allRows, headers, INDICES_GENERIC.QUESTIONS_START, questionTypes);
        const generalRows = this._formatStatsOutput('ANALISIS GENERAL (TODOS LOS DEPARTAMENTOS)', generalStats);
        
        // B. Demografia General
        const demoGeneralRows = Demographics.analyzeDemographics(allRows, this.demoMap);

        // C. Demografia Dept
        const demoDeptRows = [];
        Object.keys(departments).forEach(dept => {
            demoDeptRows.push([`DEPARTAMENTO: ${dept.toUpperCase()}`, '', '']);
            const results = Demographics.analyzeDemographics(departments[dept], this.demoMap);
            results.forEach(r => demoDeptRows.push(r));
            demoDeptRows.push(['', '', '']);
            demoDeptRows.push(['-------------------------', '', '']);
            demoDeptRows.push(['', '', '']);
        });

        // D. Scoring by Dept (Main)
        const mainRows = [];
        Object.keys(departments).forEach(dept => {
            const stats = Statistics.analyzeQuestions(departments[dept], headers, INDICES_GENERIC.QUESTIONS_START, questionTypes);
            if (stats.length > 0) {
                mainRows.push(['', '', '', '', '', '', '', '', '']);
                const formatted = this._formatStatsOutput(`DEPARTAMENTO: ${dept.toUpperCase()}`, stats);
                // remove table header from subsequent if desired, but original keeps it.
                // original adds header for each block? Yes.
                // My helper _formatStatsOutput includes header.
                formatted.forEach(r => mainRows.push(r));
            }
        });

        return {
            outputName: `analisis_${sheetName.toLowerCase()}.xlsx`,
            sheets: {
                'Analisis General': { data: generalRows, cols: this.colWidthsMain },
                [sheetName]: { data: mainRows, cols: this.colWidthsMain },
                'Demografía General': { data: demoGeneralRows, cols: this.colWidthsDemo },
                'Demografía Dept': { data: demoDeptRows, cols: this.colWidthsDemo }
            }
        };
    }

    _formatStatsOutput(title, stats) {
        const rows = [];
        rows.push([title, '', '', '', '', '', '', '', '']);
        rows.push(this.headerMain);
        
        let totalSum = 0;
        let count = 0;

        stats.forEach(stat => {
            rows.push([stat.q, stat.avg, stat.score100, stat.total, stat.d1, stat.d2, stat.d3, stat.d4, stat.d5]);
            
            const val = parseFloat(stat.score100);
            if (!isNaN(val)) {
                totalSum += val;
                count++;
            }
        });

        const globalAvg = count > 0 ? (totalSum / count).toFixed(2) : '0.00';
        rows.push(['', '', '', '', '', '', '', '', '']);
        // Column 2 is Score
        rows.push(['PROMEDIO GENERAL', '', `${globalAvg}%`, '', '', '', '', '', '']);

        return rows;
    }
}

module.exports = GenericProcessor;

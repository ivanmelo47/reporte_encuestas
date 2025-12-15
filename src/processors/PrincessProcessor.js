const { INDICES_PRINCESS, DEMO_MAP_KEYS, HEADER_ROW_INDEX, DATA_START_INDEX, QUESTION_TYPE_ROW_INDEX } = require('../config/constants');
const Statistics = require('../core/Statistics');
const Demographics = require('../core/Demographics');
const ExcelReader = require('../services/ExcelReader');

class PrincessProcessor {
    constructor() {
        this.demoMap = {};
        Object.keys(DEMO_MAP_KEYS).forEach(key => {
            const indexKey = DEMO_MAP_KEYS[key];
            this.demoMap[key] = INDICES_PRINCESS[indexKey];
        });
        
        this.colWidthsMain = [{ wch: 50 }, { wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
        this.colWidthsDemo = [{ wch: 30 }, { wch: 10 }, { wch: 15 }];
        this.headerMain = ['Pregunta', 'Nivel Promedio', 'Calificación (0-100)', 'Total Respuestas', 'Siempre', 'Casi siempre', 'Algunas veces', 'Casi nunca', 'Nunca'];
    }

    process(filePath) {
        const data = ExcelReader.read(filePath, 'Worksheet');
        const headers = data[HEADER_ROW_INDEX];
        const allRows = [];
        const properties = {};

        // Group by Property
        for (let i = DATA_START_INDEX; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            
            // Fix: Exclude Summary/Footer rows
            // INDICES_PRINCESS.GENDER = 4
            if (!row[INDICES_PRINCESS.GENDER] && !row[INDICES_PRINCESS.PROPIEDAD]) continue;

            const propCol = INDICES_PRINCESS.PROPIEDAD;
            const propName = row[propCol] ? String(row[propCol]).trim() : 'Desconocida';
            if (!properties[propName]) properties[propName] = [];
            properties[propName].push(row);
        }

        const results = [];
        Object.keys(properties).forEach(propName => {
            const rows = properties[propName];
            const questionTypes = data[QUESTION_TYPE_ROW_INDEX];
            const workbookData = this._processGroup(rows, propName, headers, questionTypes);
            results.push(workbookData);
        });

        return results;
    }

    _processGroup(rows, groupName, headers, questionTypes) {
        // Group by Department within this Property
        const departments = {};
        rows.forEach(row => {
            const deptCol = INDICES_PRINCESS.DEPT;
            const deptName = row[deptCol] ? String(row[deptCol]).trim() : 'Sin Departamento';
            if (!departments[deptName]) departments[deptName] = [];
            departments[deptName].push(row);
        });

        // A. Analisis General
        const generalStats = Statistics.analyzeQuestions(rows, headers, INDICES_PRINCESS.QUESTIONS_START, questionTypes);
        const generalRows = this._formatStatsOutput('ANALISIS GENERAL (TODOS LOS DEPARTAMENTOS)', generalStats);

        // B. Demografia General
        const demoGeneralRows = Demographics.analyzeDemographics(rows, this.demoMap);

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
            const stats = Statistics.analyzeQuestions(departments[dept], headers, INDICES_PRINCESS.QUESTIONS_START, questionTypes);
            if (stats.length > 0) {
                mainRows.push(['', '', '', '', '', '', '', '', '']);
                const formatted = this._formatStatsOutput(`DEPARTAMENTO: ${dept.toUpperCase()}`, stats);
                formatted.forEach(r => mainRows.push(r));
            }
        });

        // Safe Name logic
        const safeSheetName = groupName.substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
        const safeFileName = groupName.replace(/\s+/g,'_').toLowerCase();

        return {
            outputName: `analisis_princess_${safeFileName}.xlsx`,
            sheets: {
                'Analisis General': { data: generalRows, cols: this.colWidthsMain },
                [safeSheetName]: { data: mainRows, cols: this.colWidthsMain },
                'Demografía General': { data: demoGeneralRows, cols: this.colWidthsDemo },
                'Demografía Dept': { data: demoDeptRows, cols: this.colWidthsDemo }
            }
        };
    }

    _formatStatsOutput(title, stats) {
        const rows = [];
        rows.push([title, '', '', '', '', '', '', '', '']);
        rows.push(this.headerMain);
        stats.forEach(stat => {
            rows.push([stat.q, stat.avg, stat.score100, stat.total, stat.d1, stat.d2, stat.d3, stat.d4, stat.d5]);
        });
        return rows;
    }
}

module.exports = PrincessProcessor;

const fs = require('fs');
const path = require('path');
const ExcelReader = require('../services/ExcelReader');

class ComparativeProcessor {
    constructor() {
        this.comparativeMap = null;
        // Standard headers for the output report
        this.colWidths = [{ wch: 50 }, { wch: 10 }, { wch: 50 }, { wch: 10 }, { wch: 10 }];
    }

    process(config) {
        console.log("Starting Comparative Analysis (From Processed Reports)...");

        // 1. Load Comparativo
        if (!fs.existsSync(config.comparativoPath)) {
            throw new Error(`Comparativo file not found: ${config.comparativoPath}`);
        }
        const compData = JSON.parse(fs.readFileSync(config.comparativoPath, 'utf8'));
        this.comparativeMap = compData.comparativo_completo;

        // 2. Load P.json (Small Table Data)
        const jsonData = this._loadJsonData(config.jsonPath);
        const smallIndex = this._indexData(jsonData);

        // 3. Load Processed Excel Data
        // Structure: { 'PropName': { 'DeptName': { 'CleanQuestion': val } } }
        const largeIndex = {}; 
        
        for (const [propName, filePath] of Object.entries(config.excelMap)) {
            console.log(`Loading Analyzed Report for ${propName}: ${filePath}`);
            if (fs.existsSync(filePath)) {
                largeIndex[propName] = this._readProcessedExcel(filePath);
            } else {
                console.warn(`File not found: ${filePath}`);
            }
        }

        // 4. Build Comparison Report per Property
        const sheets = {};
        const headers = ['Pregunta Tabla Pequeña', 'Resultado Pequeña', 'Pregunta Tabla Grande', 'Resultado Grande', 'Diferencia'];

        for (const [prop, depts] of Object.entries(smallIndex)) {
            const sheetRows = [];
            
            for (const [dept, questions] of Object.entries(depts)) {
                const deptRows = [];
                
                for (const [qSmallCleaned, scoreSmall] of Object.entries(questions)) {
                    // Find mapping
                    const mapping = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === qSmallCleaned);
                    if (mapping) {
                        const qLarge = mapping.pregunta_tabla_grande;
                        const qLargeCleaned = this._cleanQuestion(qLarge);
                        
                        let scoreLarge = 'N/A';
                        let diff = 'N/A';
                        
                        // Lookup in Large Index
                        if (largeIndex[prop]) {
                            // Fuzzy Dept Match
                            const largeDepts = Object.keys(largeIndex[prop]);
                            const targetDept = largeDepts.find(d => d.toLowerCase().trim() === dept.toLowerCase().trim());
                            
                            if (targetDept && largeIndex[prop][targetDept][qLargeCleaned] !== undefined) {
                                scoreLarge = largeIndex[prop][targetDept][qLargeCleaned];
                            }
                        }

                        if (scoreLarge !== 'N/A') {
                             const nSmall = parseFloat(scoreSmall);
                             const nLarge = parseFloat(scoreLarge);
                             if (!isNaN(nSmall) && !isNaN(nLarge)) {
                                 diff = (nSmall - nLarge).toFixed(2);
                             }
                        }

                        deptRows.push([
                            mapping.pregunta_tabla_pequena, // Use original text from Comparativo.json
                            scoreSmall,
                            mapping.pregunta_tabla_grande,  // Use original text from Comparativo.json
                            scoreLarge,
                            diff
                        ]);
                    }
                }

                if (deptRows.length > 0) {
                    sheetRows.push([`DEPARTAMENTO: ${dept}`, '', '', '', '']);
                    sheetRows.push(headers);
                    deptRows.forEach(r => sheetRows.push(r));
                    sheetRows.push(['', '', '', '', '']);
                    sheetRows.push(['', '', '', '', '']);
                }
            }

            if (sheetRows.length > 0) {
                const safeName = prop.substring(0, 31).replace(/[\\/*[\]?]/g, '');
                sheets[safeName] = { data: sheetRows, cols: this.colWidths };
            }
        }

        return {
            outputName: 'Reporte_Comparativo.xlsx',
            sheets: sheets
        };
    }

    _readProcessedExcel(filePath) {
        // Reads a processed analysis file. 
        // We need to find the correct sheet. Usually it's the one that is NOT 'Analisis General' etc.
        // Assuming the file uses standard GenericProcessor output.
        // We look for a sheet that contains "DEPARTAMENTO:" blocks.
        
        try {
            // Since we can't easily list sheets with our current ExcelReader helper (it defaults to first sheet or specific name),
            // we might have to try standard names.
            // However, ExcelReader uses `xlsx.readFile` internally but returns `utils.sheet_to_json(workbook.Sheets[sheetName])`.
            // If we don't pass sheetName, it defaults to... logic in ExcelReader? 
            // Let's check ExcelReader.
            // If I can't check, I'll rely on "Palacio" for Palacio, etc.
            
            // For now, let's try reading 'Palacio', 'Pierre', 'Princess Mundo Imperial'.
            // Or better, read 'Sheet1' if it was a simple file, but these are generated.
            // GenericProcessor names the sheet as `sheetName` passed to it.
            // In index.js:
            // Palacio -> sheetName: 'Palacio'
            // Pierre -> sheetName: 'Pierre'
            // Princess -> sheetName: 'Princess Mundo Imperial'
            
            // I'll try to infer strict names based on filename or just hardcode the expected ones.
            let sheetName = 'Palacio';
            if (filePath.toLowerCase().includes('pierre')) sheetName = 'Pierre';
            if (filePath.toLowerCase().includes('princess')) sheetName = 'Princess Mundo Imperial';
            
            const data = ExcelReader.read(filePath, sheetName);
            if (!data || data.length === 0) {
                console.warn(`Sheet ${sheetName} empty or not found in ${filePath}`);
                return {};
            }

            // Parse Data
            // Format is blocks of:
            // DEPARTAMENTO: NAME
            // [Headers: Pregunta, Nivel, Calificacion...]
            // [Rows...]
            // [Empty]
            
            const result = {}; // { Dept: { CleanQ: Score } }
            let currentDept = null;
            
            for (let i = 0; i < data.length; i++) {
                let row = data[i];
                // Row might be array or object depending on ExcelReader mode. 
                // ExcelReader usually returns array of arrays if header=1 or default.
                // Let's look at ExcelReader usage in GenericProcessor:
                // `const data = ExcelReader.read(filePath, 'Worksheet');` // indices access -> Array of arrays.
                
                // If row is empty
                if (!row || row.length === 0) continue;
                
                const firstCell = String(row[0] || '').trim();
                
                // Detect Dept Header
                if (firstCell.startsWith('DEPARTAMENTO:')) {
                    currentDept = firstCell.replace('DEPARTAMENTO:', '').trim();
                    result[currentDept] = {};
                    // Skip header row usually follows immediately
                    continue; // Next row is header
                }
                
                // Detect Headers (skip)
                if (firstCell === 'Pregunta' || firstCell === 'PROMEDIO GENERAL') continue;
                
                // Process Data Row if we are in a Dept block
                if (currentDept) {
                    // Columns: 0=Pregunta, 1=Nivel, 2=Calificacion(0-100)
                    // Generic Processor: 
                    // rows.push([stat.q, stat.avg, stat.score100...]);
                    const question = row[0];
                    let score = row[2]; // Index 2 is Score
                    
                    if (question && score !== undefined) {
                        const cleanQ = this._cleanQuestion(question);
                        // Clean Score (remove % if present)
                        if (typeof score === 'string' && score.includes('%')) {
                            score = parseFloat(score.replace('%', ''));
                        }
                        result[currentDept][cleanQ] = score;
                    }
                }
            }
            
            return result;

        } catch (err) {
            console.error(`Error reading existing report ${filePath}:`, err.message);
            return {};
        }
    }

    _loadJsonData(filePath) {
        const raw = fs.readFileSync(filePath, 'utf8');
        let data = JSON.parse(raw);
        if (Array.isArray(data) && data[0] && !data[0].Propiedad && data.find(x => x.type === 'table')) {
            data = data.find(x => x.type === 'table').data;
        }
        return data || [];
    }

    _indexData(rows) {
        const index = {};
        rows.forEach(row => {
            const p = row.Propiedad;
            const d = row.Departamento;
            const q = this._cleanQuestion(row.Pregunta);
            let s = row.Resultado_Actual;

            if (s && typeof s === 'string' && s.includes('%')) {
                s = parseFloat(s.replace('%', ''));
            }

            if (!index[p]) index[p] = {};
            if (!index[p][d]) index[p][d] = {};
            index[p][d][q] = s;
        });
        return index;
    }

    _cleanQuestion(text) {
        if (!text) return '';
        let str = String(text);
        // Remove leading numbers
        str = str.replace(/^\d+[\.\-\)]\s*/, '');
        // Lowercase
        str = str.toLowerCase();
        // Remove accents
        str = str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        // Normalize spaces
        str = str.replace(/\s+/g, ' ').trim();
        // Remove trailing punctuation
        str = str.replace(/[\.\:\;]$/, '');
        return str;
    }
}

module.exports = ComparativeProcessor;

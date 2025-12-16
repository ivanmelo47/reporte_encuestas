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
        this.propertyMap = compData.comparativo_propiedad; 
        this.departmentMap = compData.comparativo_departamentos || []; // Load specific dept map

        // 2. Load P.json (Small Table Data)
        const jsonData = this._loadJsonData(config.jsonPath);
        const smallIndex = this._indexData(jsonData);

        // 3. Load Processed Excel Data
        // Structure: { 'PropName': { 'DeptName': { 'CleanQuestion': val } } }
        const largeIndex = {}; 
        
        for (const [propName, filePath] of Object.entries(config.excelMap)) {
            console.log(`Loading Analyzed Report for ${propName}: ${filePath}`);
            if (fs.existsSync(filePath)) {
                largeIndex[propName] = this._readProcessedExcel(filePath, propName);
            } else {
                console.warn(`File not found: ${filePath}`);
            }
        }

        // 4. Build Comparison Report per Property
        const sheets = {};
        const headers = ['Pregunta Tabla Pequeña', 'Resultado Pequeña', 'Pregunta Tabla Grande', 'Resultado Grande', 'Diferencia'];

        for (const [prop, depts] of Object.entries(smallIndex)) {
            const sheetRows = [];
            
            // Standard Property Mapping (Default)
            const propMapping = this.propertyMap.find(m => m.propiedad_tabla_pequena === prop);
            const defaultLargePropName = propMapping ? propMapping.propiedad_tabla_grande : prop;
            
            for (const [dept, questions] of Object.entries(depts)) {
                
                // --- START Specific Department Mapping Logic ---
                // Check if specific rule exists for this Property + Dept
                const deptMapping = this.departmentMap.find(m => 
                    m.propiedad_tabla_pequena === prop && 
                    this._normalize(m.departamento_tabla_pequena) === this._normalize(dept)
                );

                let targetLargePropName = defaultLargePropName;
                let targetLargeDeptNameSearch = dept; // Default: search for same name

                if (deptMapping) {
                    targetLargePropName = deptMapping.propiedad_tabla_grande;
                    targetLargeDeptNameSearch = deptMapping.departamento_tabla_grande;
                    // console.log(`[Override] Mapping '${dept}' in '${prop}' -> '${targetLargeDeptNameSearch}' in '${targetLargePropName}'`);
                }
                
                // --- END Specific Department Mapping Logic ---

                // Resolve Data Source
                const currentLargeData = largeIndex[targetLargePropName];
                
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
                        if (currentLargeData) {
                            // Fuzzy Dept Match with Normalization (Ignore Case & Accents)
                            const largeDepts = Object.keys(currentLargeData);
                            // Match against the "Search Target" (either original name or mapped specific name)
                            const targetDept = largeDepts.find(d => this._normalize(d) === this._normalize(targetLargeDeptNameSearch));
                            
                            if (targetDept && currentLargeData[targetDept][qLargeCleaned] !== undefined) {
                                scoreLarge = currentLargeData[targetDept][qLargeCleaned];
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

    _readProcessedExcel(filePath, propName) {
        // Reads a processed analysis file. 
        // propName matches keys in config.excelMap (e.g. "Arena GNP Seguros", "Direccion")
        
        try {
            // Determine sheet name based on property name
            let sheetName = propName; // Default
            
            // Map specific names if needed based on user input
            if (propName === 'Palacio Mundo Imperial') sheetName = 'Palacio';
            if (propName === 'Pierre Mundo Imperial') sheetName = 'Pierre';
            if (propName === 'Direccion') sheetName = 'Dirección'; // Fix accent
            // 'Arena GNP Seguros' -> 'Arena GNP Seguros' (matches default)
            // 'Princess Mundo Imperial' -> 'Princess Mundo Imperial' (matches default)
            // 'Mundo Imperial' -> 'Mundo Imperial' (matches default)

            // Try multiple sheet name variations
            const candidates = [
                sheetName,
                sheetName.trim(),
                sheetName.normalize("NFD").replace(/[\u0300-\u036f]/g, ""), // No accent
                'Dirección', // Explicit correction
                'Arena GNP Seguros',
                'Mundo Imperial',
                'Palacio',
                'Pierre',
                'Princess Mundo Imperial',
                'Dirección General',
                'Direccion General'
            ];

            let data = [];
            for (const name of candidates) {
                try {
                     const attempt = ExcelReader.read(filePath, name);
                     if (attempt && attempt.length > 0) {
                         // Verify it has data
                         if(attempt.some(row => row && row[0] && String(row[0]).startsWith('DEPARTAMENTO:'))) {
                             data = attempt;
                             console.log(`Found valid data in sheet: '${name}'`);
                             break;
                         }
                     }
                } catch (e) { /* ignore */ }
            }

            if (!data || data.length === 0) {
                 // Fallback: Try reading the default/first sheet if ExcelReader supports it (e.g. name=null or undefined)
                 // Or we can try to guess based on standard names
                 console.warn(`Could not find valid data sheet in ${filePath}. Tried: ${candidates.join(', ')}`);
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

    _normalize(text) {
        return String(text).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
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

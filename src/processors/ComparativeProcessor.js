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
        const headers = ['Pregunta Tabla Grande', 'Pregunta', 'Resultado Actual', 'Resultado Anterior', 'Diferencia'];

        for (const [prop, depts] of Object.entries(smallIndex)) {
            const sheetRows = [];
            const propStats = { questions: {}, depts: [] };
            const processedLargeDepts = new Set();
            
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
                const usedLargeKeys = new Set(); 
                
                let targetDept = null;
                
                // Pre-calculate Target Dept Key if available
                if (currentLargeData) {
                    const largeDepts = Object.keys(currentLargeData);
                    targetDept = largeDepts.find(d => this._normalize(d) === this._normalize(targetLargeDeptNameSearch));
                    if (targetDept) processedLargeDepts.add(targetDept);
                }
                
                const deptStats = { sSum: 0, sCount: 0, lSum: 0, lCount: 0 };

                // 1. Process Small Table Order (Primary) - SORTED BY 'ORDEN'
                const smallQuestionEntries = Object.entries(questions);
                
                smallQuestionEntries.sort((a, b) => {
                    const qA = a[0]; // Cleaned Question A
                    const qB = b[0]; // Cleaned Question B
                    
                    const mapA = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === qA);
                    const mapB = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === qB);
                    
                    const orderA = (mapA && mapA.orden !== undefined) ? mapA.orden : 9999;
                    const orderB = (mapB && mapB.orden !== undefined) ? mapB.orden : 9999;
                    
                    return orderA - orderB;
                });

                for (const [qSmallCleaned, scoreSmall] of smallQuestionEntries) {
                    // Find mapping
                    const mapping = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === qSmallCleaned);
                    
                    let scoreLarge = 'N/A';
                    let textLarge = '---';
                    let diff = 'N/A';
                    let matchFound = false;

                    if (mapping) {
                        const qLarge = mapping.pregunta_tabla_grande;
                        // Use mapped large name as default text (even if not found in data)
                        textLarge = qLarge; 
                        
                        const qLargeCleaned = this._cleanQuestion(qLarge);
                        
                        // Lookup in Large Index
                        if (targetDept && currentLargeData[targetDept][qLargeCleaned]) {
                            const lgObj = currentLargeData[targetDept][qLargeCleaned];
                            scoreLarge = lgObj.score;
                            // Use actual text from file if available, otherwise mapped name
                            textLarge = lgObj.text; 
                            usedLargeKeys.add(qLargeCleaned); // Mark as used
                            matchFound = true;
                        }

                        if (scoreLarge !== 'N/A') {
                             const nSmall = parseFloat(scoreSmall);
                             const nLarge = parseFloat(scoreLarge);
                             if (!isNaN(nSmall) && !isNaN(nLarge)) {
                                 // Difference: Large - Small
                                 diff = (nLarge - nSmall).toFixed(2) + '%';
                             }
                        }
                    }

                    // Format scores as percentages if they are numbers
                    const fmtScoreLarge = (scoreLarge !== 'N/A' && !String(scoreLarge).includes('%')) ? scoreLarge + '%' : scoreLarge;
                    const fmtScoreSmall = (scoreSmall !== 'N/A' && !String(scoreSmall).includes('%')) ? scoreSmall + '%' : scoreSmall;

                    // Only show if mapping exists (User Request)
                    if (mapping) {
                        deptRows.push([
                            textLarge, // Large Question (Col 1)
                            mapping.pregunta_tabla_pequena, // Small Question (Col 2)
                            fmtScoreLarge, // Res Actual (Col 3)
                            fmtScoreSmall, // Res Anterior (Col 4)
                            diff
                        ]);

                        // --- AGGREGATION LOGIC ---
                        if (scoreSmall !== 'N/A' && !isNaN(parseFloat(scoreSmall))) {
                            const val = parseFloat(scoreSmall);
                            // Question Stats
                            if (!propStats.questions[qSmallCleaned]) propStats.questions[qSmallCleaned] = { sSum: 0, sCount: 0, lSum: 0, lCount: 0 };
                            propStats.questions[qSmallCleaned].sSum += val;
                            propStats.questions[qSmallCleaned].sCount++;
                            // Dept Stats
                            deptStats.sSum += val;
                            deptStats.sCount++;
                        }
                        if (scoreLarge !== 'N/A' && !isNaN(parseFloat(scoreLarge))) {
                            const val = parseFloat(scoreLarge);
                            // Question Stats
                            if (!propStats.questions[qSmallCleaned]) propStats.questions[qSmallCleaned] = { sSum: 0, sCount: 0, lSum: 0, lCount: 0 };
                            propStats.questions[qSmallCleaned].lSum += val;
                            propStats.questions[qSmallCleaned].lCount++;
                            // Dept Stats
                            deptStats.lSum += val;
                            deptStats.lCount++;
                        }
                    }
                }

                // Store Dept Averages
                const avgSmall = deptStats.sCount > 0 ? (deptStats.sSum / deptStats.sCount) : 'N/A';
                const avgLarge = deptStats.lCount > 0 ? (deptStats.lSum / deptStats.lCount) : 'N/A';
                propStats.depts.push({ name: dept, avgSmall, avgLarge });


                if (deptRows.length > 0) {
                    sheetRows.push([`DEPARTAMENTO: ${dept}`, '', '', '', '']);
                    sheetRows.push(headers);
                    deptRows.forEach(r => sheetRows.push(r));

                    // Add Average Row
                    const fmtAvgSmall = avgSmall !== 'N/A' ? avgSmall.toFixed(2) + '%' : 'N/A';
                    const fmtAvgLarge = avgLarge !== 'N/A' ? avgLarge.toFixed(2) + '%' : 'N/A';
                    
                    let fmtAvgDiff = 'N/A';
                    if (avgSmall !== 'N/A' && avgLarge !== 'N/A') {
                        const diffVal = avgLarge - avgSmall;
                        fmtAvgDiff = diffVal.toFixed(2) + '%';
                    }

                    sheetRows.push(['', 'Promedio', fmtAvgLarge, fmtAvgSmall, fmtAvgDiff]);

                    sheetRows.push(['', '', '', '', '']);
                    sheetRows.push(['', '', '', '', '']);
                }
            }



            // --- 2. Process Unmatched Large Departments ---
            // FORCE CHECK for Unmatched
            const unmatchedSourceData = largeIndex[defaultLargePropName];
            
            if (unmatchedSourceData) {
                const allLargeDepts = Object.keys(unmatchedSourceData);

                for (const lgDept of allLargeDepts) {
                    if (!processedLargeDepts.has(lgDept)) {
                        
                        const unmatchedRows = [];
                        const deptStats = { sSum: 0, sCount: 0, lSum: 0, lCount: 0 }; 
                        const seenQuestions = new Set(); 

                        // Use mappings to define rows
                        const sortedMappings = [...this.comparativeMap];
                        sortedMappings.sort((a, b) => {
                             const orderA = (a.orden !== undefined) ? a.orden : 9999;
                             const orderB = (b.orden !== undefined) ? b.orden : 9999;
                             return orderA - orderB;
                        });

                        for (const mapping of sortedMappings) {
                             const qLargeTextMap = mapping.pregunta_tabla_grande;
                             const qSmallTextMap = mapping.pregunta_tabla_pequena;
                             const qLargeCleaned = this._cleanQuestion(qLargeTextMap);
                             const qSmallCleaned = this._cleanQuestion(qSmallTextMap);
                             
                             let scoreLarge = 'N/A';
                             let textLarge = qLargeTextMap;
                             
                             if (unmatchedSourceData[lgDept] && unmatchedSourceData[lgDept][qLargeCleaned]) {
                                 const lgObj = unmatchedSourceData[lgDept][qLargeCleaned];
                                 scoreLarge = lgObj.score;
                                 textLarge = lgObj.text;
                             }

                             // FILTER 1: Only show if we actually found a result in the Large Table (Per user request)
                             if (scoreLarge === 'N/A') {
                                 continue;
                             }
                             
                             // FILTER 2: Deduplicate based on the Large Question Text
                             // Since we are iterating mappings, duplicates in mapping cause duplicates in output.
                             // For Unmatched, we only care about unique Large Questions found.
                             if (seenQuestions.has(textLarge)) {
                                 continue;
                             }
                             seenQuestions.add(textLarge);

                             const fmtScoreLarge = (scoreLarge !== 'N/A' && !String(scoreLarge).includes('%')) ? scoreLarge + '%' : scoreLarge;
                             const fmtScoreSmall = 'N/A';
                             const diff = 'N/A';

                             unmatchedRows.push([
                                 textLarge, // Large Question (Col 1)
                                 qSmallTextMap, // Small Question (Col 2)
                                 fmtScoreLarge, // Result Actual (Col 3)
                                 fmtScoreSmall, // Result Ante (Col 4)
                                 diff
                             ]);

                             if (scoreLarge !== 'N/A' && !isNaN(parseFloat(scoreLarge))) {
                                const val = parseFloat(scoreLarge);
                                if (!propStats.questions[qSmallCleaned]) propStats.questions[qSmallCleaned] = { sSum: 0, sCount: 0, lSum: 0, lCount: 0 };
                                propStats.questions[qSmallCleaned].lSum += val;
                                propStats.questions[qSmallCleaned].lCount++;
                                deptStats.lSum += val;
                                deptStats.lCount++;
                            }
                        }


                        const avgSmall = 'N/A'; 
                        const avgLarge = deptStats.lCount > 0 ? (deptStats.lSum / deptStats.lCount) : 'N/A';
                        propStats.depts.push({ name: lgDept, avgSmall, avgLarge }); 

                        if (unmatchedRows.length > 0) {
                              sheetRows.push([`DEPARTAMENTO: ${lgDept}`, '', '', '', '']);
                              sheetRows.push(headers); 
                              unmatchedRows.forEach(r => sheetRows.push(r));

                              // Add Average Row
                              const fmtAvgSmall = 'N/A';
                              const fmtAvgLarge = avgLarge !== 'N/A' ? avgLarge.toFixed(2) + '%' : 'N/A';
                              
                              let fmtAvgDiff = 'N/A';
                              // Since avgSmall is always N/A for unmatched, diff is N/A. 
                              // Use N/A explicitly or keep logic generic if avgSmall ever changes.
                              // For now, keeping N/A as per current logic.
                              
                              sheetRows.push(['', 'Promedio', fmtAvgLarge, fmtAvgSmall, fmtAvgDiff]);

                              sheetRows.push(['', '', '', '', '']);
                              sheetRows.push(['', '', '', '', '']);
                        }
                    } else {
                         // console.log(`[DEBUG_EXTENDED] Skipped Matched Dept: '${lgDept}'`);
                    }
                }
            }

            // --- GENERATE SUMMARY TABLES ---

            // 1. General Questions Table
            if (Object.keys(propStats.questions).length > 0) {
                sheetRows.push(['RESUMEN GENERAL POR PREGUNTA', '', '', '', '']);
                sheetRows.push(headers);

                // Sort using same logic as Detail
                 const summaryQEntries = Object.entries(propStats.questions);
                 summaryQEntries.sort((a, b) => {
                    const qA = a[0]; 
                    const qB = b[0]; 
                    const mapA = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === qA);
                    const mapB = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === qB);
                    const orderA = (mapA && mapA.orden !== undefined) ? mapA.orden : 9999;
                    const orderB = (mapB && mapB.orden !== undefined) ? mapB.orden : 9999;
                    return orderA - orderB;
                });

                for (const [key, stats] of summaryQEntries) {
                    const mapping = this.comparativeMap.find(m => this._cleanQuestion(m.pregunta_tabla_pequena) === key);
                    const qSmallText = mapping ? mapping.pregunta_tabla_pequena : key;
                    const qLargeText = mapping ? mapping.pregunta_tabla_grande : '---';

                    const finalAvgSmall = stats.sCount > 0 ? (stats.sSum / stats.sCount).toFixed(2) : 'N/A';
                    const finalAvgLarge = stats.lCount > 0 ? (stats.lSum / stats.lCount).toFixed(2) : 'N/A';
                    let genDiff = 'N/A';
                    
                    if (finalAvgSmall !== 'N/A' && finalAvgLarge !== 'N/A') {
                        genDiff = (parseFloat(finalAvgLarge) - parseFloat(finalAvgSmall)).toFixed(2) + '%';
                    }
                    
                    const fmtS = finalAvgSmall !== 'N/A' ? finalAvgSmall + '%' : 'N/A';
                    const fmtL = finalAvgLarge !== 'N/A' ? finalAvgLarge + '%' : 'N/A';

                    sheetRows.push([qLargeText, qSmallText, fmtL, fmtS, genDiff]);
                }
                sheetRows.push(['', '', '', '', '']);
                sheetRows.push(['', '', '', '', '']);
            }

            // 2. General Departments Table
            if (propStats.depts.length > 0) {
                 sheetRows.push(['RESUMEN GENERAL POR DEPARTAMENTO', '', '', '', '']);
                 sheetRows.push(headers); // Reuse headers but Q columns will hold Dept Name

                 propStats.depts.forEach(d => {
                    let diff = 'N/A';
                    const s = d.avgSmall;
                    const l = d.avgLarge;
                    
                    if (s !== 'N/A' && l !== 'N/A') {
                        diff = (l - s).toFixed(2) + '%';
                    }

                    const fmtS = s !== 'N/A' ? s.toFixed(2) + '%' : 'N/A';
                    const fmtL = l !== 'N/A' ? l.toFixed(2) + '%' : 'N/A';

                    // Using 'headers' columns:
                    // Col 1 (Q Large) -> Dept Name
                    // Col 2 (Q Small) -> Dept Name (Repeated or empty?) - User said "same columns". 
                    // Let's put Dept Name in both or just First. I'll put in First and "---" in second.
                    
                    sheetRows.push([
                        d.name,
                        d.name, // Repeating name for clarity if user strictly wants 2 cols. 
                        fmtL,
                        fmtS,
                        diff
                    ]);
                 });
                 sheetRows.push(['', '', '', '', '']);
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
                if (row[0] && String(row[0]).startsWith('DEPARTAMENTO:')) {
                    currentDept = String(row[0]).replace('DEPARTAMENTO:', '').trim();
                    console.log(`[DEBUG] Found Dept in Excel (${filePath}): '${currentDept}'`); // DEBUG
                    result[currentDept] = {};
                    // Skip header row usually follows immediately
                    continue; // Next row is header
                } else if (currentDept && row[0]) {
                    // This 'else if' block is intended to skip the header row that immediately follows a department header.
                    // The 'continue' from the 'if' block already handles skipping the *current* row (the department header itself).
                    // The original logic was to 'continue' after finding the department, effectively skipping the *next* row (the actual header).
                    // The instruction's placement here would cause all data rows to be skipped if currentDept is set.
                    // Reverting to original logic for skipping the header row that follows the department.
                    // The original 'continue' after setting currentDept already skips the *next* row (the header).
                    // So, no additional 'else if' is needed here for skipping the header.
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
                        // Store both score and original text
                        result[currentDept][cleanQ] = { score, text: question };
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

const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

class ExcelWriter {
    /**
     * Writes a workbook to a file using standard XLSX.
     * @param {Object} sheets - Object { 'SheetName': { data: [], cols: [] } }
     * @param {String} outputPath 
     */
    static write(sheets, outputPath) {
        const workbook = XLSX.utils.book_new();

        Object.keys(sheets).forEach(name => {
            const sheetContent = sheets[name];
            const data = sheetContent.data || [];
            const cols = sheetContent.cols || [];

            const ws = XLSX.utils.aoa_to_sheet(data);
            if (cols.length > 0) {
                ws['!cols'] = cols;
            }

            const safeName = name.substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
            XLSX.utils.book_append_sheet(workbook, ws, safeName);
        });

        XLSX.writeFile(workbook, outputPath);
    }

    /**
     * Writes a structured report with specific styling.
     * @param {Object} reportData - Structure: { outputName: string, properties: [ { name: string, departments: [ { name: string, stats: [] } ] } ] }
     */
    static async writeStyledReport(reportData) {
        const workbook = new ExcelJS.Workbook();
        
        reportData.properties.forEach(prop => {
            const safeName = prop.name.substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
            const worksheet = workbook.addWorksheet(safeName);
            let currentRow = 1;

            const deptScores = []; // For Property Summary

            prop.departments.forEach(dept => {
                // 1. Department Title
                const titleRow = worksheet.getRow(currentRow);
                titleRow.getCell(1).value = dept.name;
                titleRow.font = { size: 14, bold: true };
                currentRow++;

                // 2. Table Header
                const headerRow = worksheet.getRow(currentRow);
                headerRow.values = ['Pregunta', 'Resultado Anterior'];
                headerRow.font = { bold: true };
                headerRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
                headerRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
                
                worksheet.getColumn(1).width = 80;
                worksheet.getColumn(2).width = 20;
                currentRow++;

                // 3. Data Rows
                // stats is array of { q: string, score100: string/number ... }
                let sumScores = 0;
                let countScores = 0;

                dept.stats.forEach(item => {
                    const row = worksheet.getRow(currentRow);
                    row.getCell(1).value = item.q;
                    
                    let val = 0;
                    if (item.score100 !== null && item.score100 !== undefined) {
                        const numeric = parseFloat(item.score100);
                        if (!isNaN(numeric)) {
                            val = numeric;
                            sumScores += numeric;
                            countScores++;
                        }
                    }
                    // Output as number, fixed to 2 decimals if needed, but keeping as number type for Excel if possible
                    // val is already a number. toFixed returns string.
                    // If user wants numeric format, best is to set value as number.
                    row.getCell(2).value = parseFloat(val.toFixed(2));
                    currentRow++;
                });

                // 4. Department Average
                let deptAvg = countScores > 0 ? (sumScores / countScores) : 0;
                // Add to summary
                deptScores.push({ nombre: dept.name, promedio: deptAvg });

                const avgRow = worksheet.getRow(currentRow);
                avgRow.getCell(1).value = "Calificación Promedio";
                avgRow.getCell(2).value = deptAvg.toFixed(2);
                avgRow.font = { bold: true };
                avgRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
                avgRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

                currentRow += 2; // Spacing
            });

            // 5. Property Summary Table
            if (deptScores.length > 0) {
                currentRow++;
                const sumTitle = worksheet.getRow(currentRow);
                sumTitle.getCell(1).value = "Resumen General Propiedad";
                sumTitle.font = { size: 16, bold: true, color: { argb: 'FF000000' } };
                currentRow++;

                const sumHeader = worksheet.getRow(currentRow);
                sumHeader.values = ['Departamento', 'Calificación'];
                sumHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                sumHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
                sumHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
                currentRow++;

                let totalPropSum = 0;
                deptScores.forEach(ds => {
                    const r = worksheet.getRow(currentRow);
                    r.getCell(1).value = ds.nombre;
                    r.getCell(2).value = parseFloat(ds.promedio.toFixed(2));
                    totalPropSum += ds.promedio;
                    currentRow++;
                });

                const finalAvg = totalPropSum / deptScores.length;
                const finalRow = worksheet.getRow(currentRow);
                finalRow.getCell(1).value = "Calificación Final Propiedad";
                finalRow.getCell(2).value = parseFloat(finalAvg.toFixed(2));
                finalRow.font = { bold: true, size: 12 };
                finalRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
                finalRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
            }
            
            // Handle Extra Sheets (e2e functionality for completeness, though JsonReportProcessor doesn't use it yet)
            if (prop.extraSheets) {
                prop.extraSheets.forEach(extra => {
                    const ws = workbook.addWorksheet(extra.name);
                    if (extra.data) {
                        extra.data.forEach(r => ws.addRow(r));
                    }
                });
            }

        });

        await workbook.xlsx.writeFile(reportData.outputName);
    }

    /**
     * Writes the Comparative Report with specific styling matching user request.
     * @param {Object} sheets - { SheetName: { data: [[...]] } }
     * @param {String} outputPath 
     */
    static async writeComparativeReport(sheets, outputPath) {
        const workbook = new ExcelJS.Workbook();

        for (const [sheetName, sheetData] of Object.entries(sheets)) {
            const cleanName = sheetName.substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
            const ws = workbook.addWorksheet(cleanName);
            const data = sheetData.data || [];

            // Add all rows first
            data.forEach(row => ws.addRow(row));

            // Set Column Widths
            ws.getColumn(1).width = 80; // Pregunta
            ws.getColumn(2).width = 40; // Pregunta Pequeña (Hide? Or kept for ref? User logic kept it)
            ws.getColumn(3).width = 15; // Actual
            ws.getColumn(4).width = 15; // Anterior
            ws.getColumn(5).width = 15; // Diferencia

            // Apply Styles
            ws.eachRow((row, rowNumber) => {
                const firstVal = row.getCell(1).value ? String(row.getCell(1).value) : '';
                
                // 1. DEPARTAMENTO Header
                if (firstVal.startsWith('DEPARTAMENTO:')) {
                    row.font = { bold: true, size: 12 };
                    return;
                }

                // 2. Table Header
                if (firstVal === 'Pregunta Tabla Grande' || firstVal === 'Pregunta' || (row.getCell(3).value === 'Resultado Actual')) {
                    row.eachCell((cell) => {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF092034' } }; // Navy Blue
                        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // White
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                     // First col alignment left
                     row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
                     return;
                }

                // 3. PROMEDIO Row
                const secondVal = row.getCell(2).value ? String(row.getCell(2).value) : '';
                if (secondVal === 'Promedio') {
                    row.font = { bold: true };
                    row.eachCell((cell) => {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; // Light Grey
                        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                    
                    // Scores (Cols 3 and 4)
                    [3, 4].forEach(colIdx => {
                        const cell = row.getCell(colIdx);
                        const valStr = String(cell.value);
                         if (valStr.includes('%')) {
                            const num = parseFloat(valStr.replace('%', ''));
                            if (!isNaN(num)) {
                               if (num <= 75) cell.font = { color: { argb: 'FFFF0000' }, bold: true }; // Red (<= 75)
                               else if (num >= 85) cell.font = { color: { argb: 'FF00B050' }, bold: true }; // Green (>= 85)
                               else cell.font = { color: { argb: 'FFED7D31' }, bold: true }; // Orange (> 75 and < 85)
                            }
                        } else if (valStr === 'N/A') {
                            cell.font = { color: { argb: 'FFFF0000' }, bold: true };
                        }
                    });


                    // Difference (Col 5)
                    const diffCell = row.getCell(5);
                    const diffValStr = String(diffCell.value);
                    if (diffValStr.includes('%')) {
                         const num = parseFloat(diffValStr.replace('%', ''));
                         if (!isNaN(num)) {
                             // User Request: > 0 Green, <= 0 Red
                             if (num > 0) diffCell.font = { color: { argb: 'FF00B050' }, bold: true }; 
                             else diffCell.font = { color: { argb: 'FFFF0000' }, bold: true }; 
                         }
                    } else if (diffValStr === 'N/A') {
                        diffCell.font = { color: { argb: 'FFFF0000' }, bold: true }; // Red for N/A
                    }
                    return;
                }

                // 4. Data Rows (Values)
                // Check if it's a data row (has score in Col 3 or 4)
                const c3 = row.getCell(3).value;
                if (c3 !== null && c3 !== undefined && c3 !== '') {
                    // Zebra Striping 
                    if (rowNumber % 2 === 0) {
                         row.eachCell({ includeEmpty: true }, (cell) => {
                            if (!cell.fill || !cell.fill.type) { 
                                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }; 
                            }
                         });
                    }

                    // Score Coloring (Cols 3 and 4)
                    [3, 4].forEach(colIdx => {
                        const cell = row.getCell(colIdx);
                        const valStr = String(cell.value);

                        cell.alignment = { horizontal: 'center' };

                        if (valStr === 'N/A') {
                            cell.font = { color: { argb: 'FFFF0000' }, bold: true }; // Red
                        } else if (valStr.includes('%')) {
                            const num = parseFloat(valStr.replace('%', ''));
                            if (!isNaN(num)) {
                               if (num <= 75) cell.font = { color: { argb: 'FFFF0000' }, bold: true }; // Red (<= 75)
                               else if (num >= 85) cell.font = { color: { argb: 'FF00B050' }, bold: true }; // Green (>= 85)
                               else cell.font = { color: { argb: 'FFED7D31' }, bold: true }; // Orange (> 75 and < 85)
                            }
                        }
                    });

                    // Difference Coloring (Col 5)
                    const diffColCell = row.getCell(5);
                    const diffDataStr = String(diffColCell.value);
                    diffColCell.alignment = { horizontal: 'center' };
                    
                    if (diffDataStr.includes('%')) {
                        const num = parseFloat(diffDataStr.replace('%', ''));
                         if (!isNaN(num)) {
                             // User Request: > 0 Green, <= 0 Red
                             if (num > 0) diffColCell.font = { color: { argb: 'FF00B050' }, bold: true }; 
                             else diffColCell.font = { color: { argb: 'FFFF0000' }, bold: true }; 
                         }
                    } else if (diffDataStr === 'N/A') {
                        diffColCell.font = { color: { argb: 'FFFF0000' }, bold: true }; // Red for N/A
                    }

                    // Wrap Text for Question
                    row.getCell(1).alignment = { wrapText: true, vertical: 'middle' };
                }
            });
        }

        await workbook.xlsx.writeFile(outputPath);
    }
}

module.exports = ExcelWriter;

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
}

module.exports = ExcelWriter;

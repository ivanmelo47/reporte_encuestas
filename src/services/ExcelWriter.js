const XLSX = require('xlsx');

class ExcelWriter {
    /**
     * Writes a workbook to a file.
     * @param {Object} sheets - Object where keys are sheet names and values are objects { data: [], cols: [] }
     * @param {String} outputPath - Full output path
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

            // Sanitize name for XLSX limits (31 chars) and invalid chars
            const safeName = name.substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
            XLSX.utils.book_append_sheet(workbook, ws, safeName);
        });

        XLSX.writeFile(workbook, outputPath);
    }
}

module.exports = ExcelWriter;

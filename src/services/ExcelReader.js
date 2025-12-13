const XLSX = require('xlsx');
const fs = require('fs');

class ExcelReader {
    static read(filePath, sheetName) {
        if (!fs.existsSync(filePath)) {
            throw new Error(`File not found: ${filePath}`);
        }
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            throw new Error(`Sheet '${sheetName}' not found in ${filePath}`);
        }
        return XLSX.utils.sheet_to_json(sheet, { header: 1 });
    }
}

module.exports = ExcelReader;


const XLSX = require('xlsx');
const path = require('path');

const FILE = 'estadisticas_encuesta_2_Palacio.xlsx';
const filePath = path.join(__dirname, FILE);

const wb = XLSX.readFile(filePath);
const sheet = wb.Sheets['Worksheet'];
const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const header = json[10];
const data = json[11]; // First data row

console.log('--- Column Debug (Indices 8-15) ---');
for (let i = 8; i <= 15; i++) {
    const h = header[i] ? String(header[i]).trim() : 'undefined';
    const d = data[i] ? String(data[i]).trim() : 'undefined';
    console.log(`Index ${i}:`);
    console.log(`  Header: [${h}]`);
    console.log(`  Data  : [${d}]`);
}

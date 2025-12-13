
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILE = 'estadisticas_encuesta_3_Princess.xlsx';

async function debugPrincess() {
    const filePath = path.join(__dirname, FILE);
    console.log(`\n\nDEBUGGING: ${FILE}`);
    
    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets['Worksheet'];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    const header = json[10];
    const data = json[11];
    
    let output = 'IDX | HEADER                           | DATA sample\n';
    output += '----|----------------------------------|-----------------------------------';
    
    // Inspect 10 to 18 to encompass surrounding columns
    for (let i = 10; i < 18; i++) {
        let h = header && header[i] ? String(header[i]).substring(0, 30) : 'UNDEFINED';
        let d = data && data[i] ? String(data[i]).substring(0, 30) : 'UNDEFINED';
        
        h = h.padEnd(32, ' ');
        d = d.padEnd(35, ' ');
        
        output += `\n${String(i).padEnd(3)} | ${h} | ${d}`;
    }
    
    fs.writeFileSync('debug_princess.txt', output);
    console.log('Debug info written to debug_princess.txt');
}

debugPrincess();

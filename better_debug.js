
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

async function debugFile(filename) {
    const filePath = path.join(__dirname, filename);
    console.log(`\n\nDEBUGGING: ${filename}`);
    
    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets['Worksheet'];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    const header = json[10];
    const data = json[11];
    
    let output = 'IDX | HEADER                           | DATA sample\n';
    output += '----|----------------------------------|-----------------------------------\n';
    for (let i = 0; i < 20; i++) {
        let h = header && header[i] ? String(header[i]).substring(0, 30) : 'UNDEFINED';
        let d = data && data[i] ? String(data[i]).substring(0, 30) : 'UNDEFINED';
        
        h = h.padEnd(32, ' ');
        d = d.padEnd(35, ' ');
        
        output += `${String(i).padEnd(3)} | ${h} | ${d}\n`;
    }
    
    fs.writeFileSync('debug_log.txt', output);
    console.log('Debug log written to debug_log.txt');
}

debugFile('estadisticas_encuesta_2_Palacio.xlsx');

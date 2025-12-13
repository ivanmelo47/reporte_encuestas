
const XLSX = require("xlsx");
const path = require("path");
const fs = require('fs');

const FILES = [
    'estadisticas_encuesta_3_Princess.xlsx',
    'estadisticas_encuesta_2_Palacio.xlsx'
];

async function inspect(file) {
    const filePath = path.join(__dirname, file);
    if (!fs.existsSync(filePath)) return;
    
    console.log(`\n\n=== FILE: ${file} ===`);
    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets['Worksheet'];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    const header = json[10];
    const data = json[11];

    // Find indices
    const deptIdx = header.findIndex(h => h && h.includes('Departamento'));
    const propIdx = header.findIndex(h => h && h.includes('Propiedad'));
    const genderIdx = header.findIndex(h => h && h.includes('Género'));

    console.log(`Header Indices:`);
    console.log(`  Departamento: ${deptIdx} (${header[deptIdx]})`);
    console.log(`  Propiedad: ${propIdx} (${propIdx !== -1 ? header[propIdx] : 'NOT FOUND'})`);
    console.log(`  Género: ${genderIdx} (${header[genderIdx]})`);

    console.log(`Data Check:`);
    console.log(`  Data at Dept Index (${deptIdx}): '${data[deptIdx]}'`);
    if (propIdx !== -1) console.log(`  Data at Prop Index (${propIdx}): '${data[propIdx]}'`);
    console.log(`  Data at Gender Index (${genderIdx}): '${data[genderIdx]}'`);
    
    // Check shift
    console.log(`  Data at Index 2: '${data[2]}'`);
    console.log(`  Data at Index 4: '${data[4]}'`);
    console.log(`  Data at Index 6: '${data[6]}'`);
}

FILES.forEach(f => inspect(f));

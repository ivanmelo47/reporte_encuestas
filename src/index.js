const path = require('path');
const fs = require('fs');
const GenericProcessor = require('./processors/GenericProcessor');
const PrincessProcessor = require('./processors/PrincessProcessor');
const ExcelWriter = require('./services/ExcelWriter');

// Configuration of jobs
const JOBS = [
    {
        type: 'GENERIC',
        input: 'estadisticas_encuesta_2_Palacio.xlsx',
        sheetName: 'Palacio'
    },
    {
        type: 'GENERIC',
        input: 'Estadisticas_encuesta_1_Pierre.xlsx',
        sheetName: 'Pierre'
    },
    {
        type: 'PRINCESS',
        input: 'estadisticas_encuesta_3_Princess.xlsx',
        sheetName: 'Princess' // Base name, though processor handles splitting
    },
    {
        type: 'JSON_REPORT',
        input: 'P.json',
        sheetName: 'Reporte_Numerico'
    }
];

const OUTPUT_DIR = path.join(__dirname, '../analisis');

async function main() {
    console.log("Starting Analysis...");
    
    // Ensure output dir exists
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR);
    }

    for (const job of JOBS) {
        const fullInputPath = path.join(__dirname, '../', job.input);
        console.log(`Processing Job: ${job.input} (${job.type})`);

        try {
            if (job.type === 'GENERIC') {
                const processor = new GenericProcessor();
                const result = processor.process(fullInputPath, job.sheetName);
                
                // Result is { outputName, sheets } for Generic
                const savePath = path.join(OUTPUT_DIR, result.outputName);
                ExcelWriter.write(result.sheets, savePath);
                console.log(`Saved: ${result.outputName}`);

            } else if (job.type === 'PRINCESS') {
                const processor = new PrincessProcessor();
                const results = processor.process(fullInputPath);
                
                // Result is ARRAY of { outputName, sheets }
                for (const res of results) {
                    const savePath = path.join(OUTPUT_DIR, res.outputName);
                    ExcelWriter.write(res.sheets, savePath);
                    console.log(`Saved: ${res.outputName}`);
                }
            } else if (job.type === 'JSON_REPORT') {
                const JsonReportProcessor = require('./processors/JsonReportProcessor');
                const processor = new JsonReportProcessor();
                const result = processor.process(fullInputPath);
                
                // Result is { outputName, properties: [...] } for Styled Writer
                // We typically save to output dir, but `result.outputName` might just be filename.
                // modify outputName to be in OUTPUT_DIR? 
                // writeStyledReport takes { outputName } and writes to it.
                // Let's prepend output dir.
                result.outputName = path.join(OUTPUT_DIR, result.outputName);
                
                await ExcelWriter.writeStyledReport(result);
                console.log(`Saved: ${result.outputName}`);
            }
        } catch (error) {
            console.error(`Error processing ${job.input}:`, error.message);
            console.error(error.stack);
        }
    }

    console.log("Analysis Complete.");
}

main();

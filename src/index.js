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
    }
];

const OUTPUT_DIR = path.join(__dirname, '../analisis');

function main() {
    console.log("Starting Analysis...");
    
    // Ensure output dir exists
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR);
    }

    JOBS.forEach(job => {
        const fullInputPath = path.join(__dirname, '../', job.input);
        console.log(`Processing Job: ${job.input} (${job.type})`);

        try {
            if (job.type === 'GENERIC') {
                const processor = new GenericProcessor();
                const result = processor.process(fullInputPath, job.sheetName);
                
                // Result is a single object { outputName, sheets }
                const savePath = path.join(OUTPUT_DIR, result.outputName);
                ExcelWriter.write(result.sheets, savePath);
                console.log(`Saved: ${result.outputName}`);

            } else if (job.type === 'PRINCESS') {
                const processor = new PrincessProcessor();
                const results = processor.process(fullInputPath);
                
                // Result is an ARRAY of objects
                results.forEach(res => {
                    const savePath = path.join(OUTPUT_DIR, res.outputName);
                    ExcelWriter.write(res.sheets, savePath);
                    console.log(`Saved: ${res.outputName}`);
                });
            }
        } catch (error) {
            console.error(`Error processing ${job.input}:`, error.message);
        }
    });

    console.log("Analysis Complete.");
}

main();

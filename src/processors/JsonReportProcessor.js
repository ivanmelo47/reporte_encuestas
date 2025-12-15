const fs = require('fs');

class JsonReportProcessor {
    process(filePath) {
        console.log(`Reading JSON from ${filePath}...`);
        const rawData = fs.readFileSync(filePath, 'utf8');
        let jsonExport;
        try {
            jsonExport = JSON.parse(rawData);
        } catch (e) {
            throw new Error(`Invalid JSON in ${filePath}`);
        }

        // 2. Extraer los datos reales
        let filas = [];
        
        // Si el archivo es directamente el array de filas
        if (Array.isArray(jsonExport)) {
             // Check if first item looks like a row
             if (jsonExport.length > 0 && jsonExport[0].Propiedad) {
                 filas = jsonExport;
             } else {
                 // Try looking for data prop
                 const tablaData = jsonExport.find(item => item.type === 'table' && item.data);
                 if (tablaData) {
                     filas = tablaData.data;
                 }
             }
        } else if (jsonExport.data && Array.isArray(jsonExport.data)) {
            // Some other format?
            filas = jsonExport.data;
        }

        if (filas.length === 0) {
            console.warn("No rows found in JSON, checking if it is empty...");
             // return default or throw?
             // throw new Error("No data rows found.");
        }

        // 3. Transform to WriteStyledReport Structure
        // Structure: properties: [ { name: propName, departments: [ { name: deptName, stats: [ { q: '', score100: '' } ] } ] } ]
        
        const propertiesMap = {};

        filas.forEach(fila => {
            const nombrePropiedad = fila.Propiedad || 'Sin Propiedad';
            const nombreDepto = fila.Departamento || 'Sin Departamento';
            const textoPregunta = fila.Pregunta;
            let valorResultado = fila.Resultado_Actual;

            // Formatear valor
            // ExcelWriter expects score100 to be a number or string that can be parsed
            if (valorResultado !== null && valorResultado !== undefined) {
               // If it's a string like "90%" leave it, writer handles it? 
               // WriteStyledReport: parseFloat(item.score100)
               // So if it comes as 90 (number) or "90", it works.
               // reportes.js stripped %. 
               if (typeof valorResultado === 'string') {
                   valorResultado = valorResultado.replace('%', '');
               }
            } else {
                valorResultado = 0;
            }

            if (!propertiesMap[nombrePropiedad]) {
                propertiesMap[nombrePropiedad] = {};
            }

            if (!propertiesMap[nombrePropiedad][nombreDepto]) {
                propertiesMap[nombrePropiedad][nombreDepto] = [];
            }

            propertiesMap[nombrePropiedad][nombreDepto].push({
                q: textoPregunta,
                score100: valorResultado
            });
        });

        // Convert Map to Array structure
        const propertiesArray = Object.keys(propertiesMap).map(propName => {
            const deptsMap = propertiesMap[propName];
            const departments = Object.keys(deptsMap).map(deptName => {
                return {
                    name: deptName,
                    stats: deptsMap[deptName]
                };
            });

            return {
                name: propName,
                departments: departments
            };
        });

        return {
            outputName: 'Reporte_Encuestas_Numerico.xlsx',
            properties: propertiesArray
        };
    }
}

module.exports = JsonReportProcessor;

const fs = require('fs');

try {
    // 1. Leer el archivo original
    console.log("Leyendo P.json...");
    const rawData = fs.readFileSync('P.json', 'utf8');
    const jsonExport = JSON.parse(rawData);

    // 2. Extraer los datos reales
    // Las exportaciones de PHPMyAdmin suelen venir en un array.
    // Buscamos el objeto que sea de tipo "table" y tenga "data".
    let filas = [];
    
    // Si el archivo es directamente el array de filas
    if (Array.isArray(jsonExport) && jsonExport[0].Propiedad) {
        filas = jsonExport;
    } 
    // Si es formato exportación (header, database, table...)
    else if (Array.isArray(jsonExport)) {
        const tablaData = jsonExport.find(item => item.type === 'table' && item.data);
        if (tablaData) {
            filas = tablaData.data;
        }
    }

    if (filas.length === 0) {
        throw new Error("No se encontraron datos de filas en P.json");
    }

    console.log(`Procesando ${filas.length} registros...`);

    // 3. Estructura de transformación
    const resultadoFinal = {};

    filas.forEach(fila => {
        const nombrePropiedad = fila.Propiedad;
        const nombreDepto = fila.Departamento;
        const textoPregunta = fila.Pregunta;
        let valorResultado = fila.Resultado_Actual;

        // Formatear el porcentaje si viene como número simple
        if (valorResultado !== null && valorResultado !== undefined) {
            if (!String(valorResultado).includes('%')) {
                valorResultado = `${valorResultado}%`;
            }
        } else {
            valorResultado = "N/A";
        }

        // A. Inicializar la Propiedad si no existe
        // Nota: Usamos el nombre de la propiedad como clave principal (ej: "Palacio Mundo Imperial")
        if (!resultadoFinal[nombrePropiedad]) {
            resultadoFinal[nombrePropiedad] = {
                departamentos: []
            };
        }

        // B. Buscar si el departamento ya existe en esa propiedad
        let departamentoObj = resultadoFinal[nombrePropiedad].departamentos.find(d => d.nombre === nombreDepto);

        // Si no existe, lo creamos y agregamos
        if (!departamentoObj) {
            departamentoObj = {
                nombre: nombreDepto,
                resultados: []
            };
            resultadoFinal[nombrePropiedad].departamentos.push(departamentoObj);
        }

        // C. Agregar la pregunta y respuesta al departamento
        departamentoObj.resultados.push({
            pregunta: textoPregunta,
            resultado_actual: valorResultado
        });
    });

    // 4. Guardar el archivo JSON final (opcional, para referencia)
    const outputFilename = 'reporte_final.json';
    fs.writeFileSync(outputFilename, JSON.stringify(resultadoFinal, null, 4), 'utf8');
    console.log(`Archivo JSON generado: ${outputFilename}`);

    // 5. Generar Excel
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();

    // Recorrer cada Propiedad (cada clave del objeto resultadoFinal)
    for (const [nombrePropiedad, datosPropiedad] of Object.entries(resultadoFinal)) {
        // Crear una hoja por Propiedad (limpiar nombre si es necesario, Excel max 31 chars)
        // Cortamos a 31 chars por si acaso, aunque nombres largos pueden ser problema
        const safeName = nombrePropiedad.substring(0, 31).replace(/[\\/*[\]?]/g, ''); 
        const worksheet = workbook.addWorksheet(safeName);

        let currentRow = 1;

        // Array para guardar promedios de departamentos de esta propiedad
        const deptScores = [];

        // Recorrer Departamentos
        datosPropiedad.departamentos.forEach(depto => {
            // Título del Departamento
            const titleRow = worksheet.getRow(currentRow);
            titleRow.getCell(1).value = depto.nombre;
            titleRow.font = { size: 14, bold: true };
            currentRow++;

            // Encabezados de la Tabla
            const headerRow = worksheet.getRow(currentRow);
            headerRow.values = ['Pregunta', 'Resultado Actual'];
            headerRow.font = { bold: true };
            headerRow.getCell(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' }
            };
            headerRow.getCell(2).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' }
            };
            
            // Ajustar anchos de columna
            worksheet.getColumn(1).width = 80;
            worksheet.getColumn(2).width = 20;
            
            currentRow++;

            // Variables para calcular promedio del departamento
            let sumScores = 0;
            let countScores = 0;

            // Filas de datos
            depto.resultados.forEach(item => {
                const row = worksheet.getRow(currentRow);
                row.getCell(1).value = item.pregunta;

                // Parsear score para usar numero y para promedio
                let finalValue = item.resultado_actual; // Default string (e.g. "N/A")

                if (item.resultado_actual && typeof item.resultado_actual === 'string' && item.resultado_actual.includes('%')) {
                    const numericValue = parseFloat(item.resultado_actual.replace('%', ''));
                    if (!isNaN(numericValue)) {
                        finalValue = numericValue;
                        sumScores += numericValue;
                        countScores++;
                    }
                }
                
                row.getCell(2).value = finalValue;
                currentRow++;
            });

            // Agregar fila de Promedio del Departamento
            let deptAvg = 0;
            if (countScores > 0) {
                deptAvg = sumScores / countScores;
            }
            
            // Guardar para resumen general
            deptScores.push({ nombre: depto.nombre, promedio: deptAvg });

            const avgRow = worksheet.getRow(currentRow);
            avgRow.getCell(1).value = "Calificación Promedio";
            avgRow.getCell(2).value = parseFloat(deptAvg.toFixed(2)); // Guardar como número
            avgRow.font = { bold: true };
            avgRow.getCell(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFF2CC' } 
            };
            avgRow.getCell(2).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFF2CC' }
            };

            // Espacio entre tablas de departamento
            currentRow += 2;
        });

        // --- TABLA DE RESUMEN DE LA PROPIEDAD ---
        if (deptScores.length > 0) {
            currentRow += 1; 

            // Título Resumen
            const summaryTitleRow = worksheet.getRow(currentRow);
            summaryTitleRow.getCell(1).value = "Resumen General Propiedad";
            summaryTitleRow.font = { size: 16, bold: true, color: { argb: 'FF000000' } };
            currentRow++;

            // Header Resumen
            const summaryHeader = worksheet.getRow(currentRow);
            summaryHeader.values = ['Departamento', 'Calificación'];
            summaryHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            summaryHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; 
            summaryHeader.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
            currentRow++;

            let totalPropiedadSum = 0;

            deptScores.forEach(ds => {
                const r = worksheet.getRow(currentRow);
                r.getCell(1).value = ds.nombre;
                r.getCell(2).value = parseFloat(ds.promedio.toFixed(2)); // Numero
                totalPropiedadSum += ds.promedio;
                currentRow++;
            });

            // Calificación Final Propiedad
            const finalPropAvg = totalPropiedadSum / deptScores.length;
            
            const finalRow = worksheet.getRow(currentRow);
            finalRow.getCell(1).value = "Calificación Final Propiedad";
            finalRow.getCell(2).value = parseFloat(finalPropAvg.toFixed(2)); // Numero
            finalRow.font = { bold: true, size: 12 };
            finalRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }; 
            finalRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
        }
    }

    const excelFilename = 'Reporte_Encuestas_Numerico.xlsx';
    workbook.xlsx.writeFile(excelFilename).then(() => {
        console.log(`¡Éxito! Archivo Excel generado: ${excelFilename}`);
    });

} catch (error) {
    console.error("Ocurrió un error:", error.message);
    console.error(error.stack);
}
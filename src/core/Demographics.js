/**
 * Analyzes demographics based on provided rows and column mapping.
 * @param {Array} rows - Data rows
 * @param {Object} columnMap - Object mapping Display Name -> Column Index (e.g. {'Género': 4})
 * @returns {Array} Rows ready for XLSX sheet
 */
function analyzeDemographics(rows, columnMap) {
    const stats = {};
    Object.keys(columnMap).forEach(key => {
        stats[key] = { total: 0, distinct: {} };
    });

    rows.forEach(row => {
        Object.keys(columnMap).forEach(key => {
            const colIndex = columnMap[key];
            const val = row[colIndex];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
                const cleanVal = String(val).trim();
                stats[key].total++;
                stats[key].distinct[cleanVal] = (stats[key].distinct[cleanVal] || 0) + 1;
            }
        });
    });

    const resultRows = [];
    Object.keys(stats).forEach(key => {
        resultRows.push([key.toUpperCase(), '', '']);
        resultRows.push(['Opción', 'Cantidad', 'Porcentaje']);
        const total = stats[key].total;
        Object.entries(stats[key].distinct).forEach(([option, count]) => {
            const pct = total > 0 ? (count / total * 100).toFixed(2) + '%' : '0%';
            resultRows.push([option, count, pct]);
        });
        resultRows.push(['', '', '']);
    });
    return resultRows;
}

module.exports = {
    analyzeDemographics
};

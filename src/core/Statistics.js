const { SCORE_MAP, DISTRIBUTION_KEYS } = require('../config/constants');

function getScore(value) {
    if (typeof value !== 'string') return null;
    const cleanVal = value.trim().toUpperCase();
    if (Object.prototype.hasOwnProperty.call(SCORE_MAP, cleanVal)) {
        return SCORE_MAP[cleanVal];
    }
    return null;
}

/**
 * Calculates statistics for questions starting from a specific column index.
 * @param {Array} rows - Data rows
 * @param {Array} headers - Header row
 * @param {Number} startIndex - Column index where questions begin
 * @returns {Array} Array of stats objects
 */
function analyzeQuestions(rows, headers, startIndex) {
    const stats = [];
    
    for (let col = startIndex; col < headers.length; col++) {
        const questionText = headers[col];
        if (!questionText) continue;

        let totalScore = 0;
        let validResponses = 0;
        const distribution = { 'Siempre': 0, 'Casi siempre': 0, 'Algunas veces': 0, 'Casi nunca': 0, 'Nunca': 0 };

        rows.forEach(row => {
            const val = row[col];
            const score = getScore(val);
            if (score !== null) {
                totalScore += score;
                validResponses++;
                // Handle "NUCA" typo and normalization
                const upperVal = String(val).trim().toUpperCase().replace('NUCA', 'NUNCA');
                
                // Map to distribution keys
                if (upperVal === 'SIEMPRE') distribution['Siempre']++;
                else if (upperVal === 'CASI SIEMPRE') distribution['Casi siempre']++;
                else if (upperVal === 'ALGUNAS VECES') distribution['Algunas veces']++;
                else if (upperVal === 'CASI NUNCA') distribution['Casi nunca']++;
                else if (upperVal === 'NUNCA') distribution['Nunca']++;
            }
        });

        if (validResponses > 0) {
            const avgScore = totalScore / validResponses;
            const score100 = (avgScore / 4) * 100;
            stats.push({
                q: questionText,
                avg: avgScore.toFixed(2),
                score100: score100.toFixed(2),
                total: validResponses,
                d1: distribution['Siempre'], 
                d2: distribution['Casi siempre'], 
                d3: distribution['Algunas veces'], 
                d4: distribution['Casi nunca'], 
                d5: distribution['Nunca']
            });
        }
    }
    return stats;
}

module.exports = {
    getScore,
    analyzeQuestions
};

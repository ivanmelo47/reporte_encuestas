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
/**
 * Calculates statistics for questions starting from a specific column index.
 * @param {Array} rows - Data rows
 * @param {Array} headers - Header row
 * @param {Number} startIndex - Column index where questions begin
 * @param {Array} questionTypes - Array of question types (+/-) from row 9
 * @returns {Array} Array of stats objects
 */
function analyzeQuestions(rows, headers, startIndex, questionTypes = []) {
    const stats = [];
    
    for (let col = startIndex; col < headers.length; col++) {
        const questionText = headers[col];
        if (!questionText) continue;

        const qType = questionTypes[col] ? String(questionTypes[col]).trim() : null;

        let totalScore = 0;
        let validResponses = 0; // Total forms that have a response (even if "Siempre" is excluded in some legacy sense, validResponses usually implies denominator)
        
        // For new logic: denominator is total valid rows?
        // User said: "dividir ese conteo por pregunta entre el numero total de cuestionarios que se respondieron"
        // This implies validResponses is simply the count of rows that have ANY value? Or just the total count of rows passed?
        // Usually, we filter out empty rows. validResponses tracks rows with non-empty relevant data.
        
        let countPositive = 0; // For + (Siempre, Casi siempre)
        let countNegative = 0; // For - (Nunca, Casi nunca)
        
        // Legacy accumulation
        let legacyTotalScore = 0;

        const distribution = { 'Siempre': 0, 'Casi siempre': 0, 'Algunas veces': 0, 'Casi nunca': 0, 'Nunca': 0 };

        rows.forEach(row => {
            const val = row[col];
            // Skip empty/null responses?
            if (val === undefined || val === null || String(val).trim() === '') return;

            const score = getScore(val);            
            if (score !== null) {
                legacyTotalScore += score;
                validResponses++;
                
                // Handle "NUCA" typo and normalization
                const upperVal = String(val).trim().toUpperCase().replace(/NUCA/g, 'NUNCA'); // safer replace
                
                // Map to distribution keys
                if (upperVal === 'SIEMPRE') distribution['Siempre']++;
                else if (upperVal === 'CASI SIEMPRE' || upperVal === 'CASI SIEMPE') distribution['Casi siempre']++; // Added simple typo catch if needed, strictly keeping original keys
                else if (upperVal === 'ALGUNAS VECES') distribution['Algunas veces']++;
                else if (upperVal === 'CASI NUNCA') distribution['Casi nunca']++;
                else if (upperVal === 'NUNCA') distribution['Nunca']++;
            }
        });

        // If using +/- logic, we generally want to divide by total population (rows.length)
        // rather than just valid responses, per user request ("dividir entre el numero total de cuestionarios").
        // However, for Legacy logic, we usually divide by validResponses to get a valid average.
        
        // Fix: Filter effectively empty rows to avoid off-by-one errors (e.g. ghost rows)
        // Determine totalSurveys by counting rows that have at least one non-empty cell in the relevant data range?
        // Or just rely on a simpler check: row must have some content.
        // Since we iterate column by column, we can't easily check "whole row" here without re-iterating.
        // But rows is passed in.
        const totalSurveys = rows.filter(r => r && r.length > 0 && r.some(c => c !== null && c !== undefined && String(c).trim() !== '')).length;

        if (totalSurveys > 0) {
            let finalScore100 = 0;

            if (qType === '+') {
                // + espera Siempre o Casi Siempre
                // Denominator: Total Surveys (including N/A)
                const numerator = distribution['Siempre'] + distribution['Casi siempre'];
                finalScore100 = (numerator / totalSurveys) * 100;
            } else if (qType === '-') {
                // - espera Nunca o Casi nunca
                // Denominator: Total Surveys (including N/A)
                const numerator = distribution['Nunca'] + distribution['Casi nunca'];
                finalScore100 = (numerator / totalSurveys) * 100;
            } else {
                // Fallback to old logic (Average)
                // Denominator: Valid Responses (excluding N/A) so average isn't diluted by non-answers
                if (validResponses > 0) {
                    const avgScore = legacyTotalScore / validResponses;
                    finalScore100 = (avgScore / 4) * 100;
                } else {
                    finalScore100 = 0;
                }
            }

            // Legacy Avg Calculation (Informational)
            const avgScore = validResponses > 0 ? (legacyTotalScore / validResponses) : 0;

            stats.push({
                q: questionText,
                avg: avgScore.toFixed(2),
                score100: finalScore100.toFixed(2),
                total: validResponses, // Keep showing Count of Answers
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

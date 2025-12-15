
// Sheet config
const SHEET_DATA_NAME = 'Worksheet';
const QUESTION_TYPE_ROW_INDEX = 8; // Row 9 in Excel (0-based)
const HEADER_ROW_INDEX = 10;
const DATA_START_INDEX = 11;

// Score mapping
const SCORE_MAP = {
    'SIEMPRE': 4,
    'CASI SIEMPRE': 3,
    'ALGUNAS VECES': 2,
    'CASI NUNCA': 1,
    'NUNCA': 0,
    'NUCA': 0 // Typo handling
};

const DISTRIBUTION_KEYS = ['Siempre', 'Casi siempre', 'Algunas veces', 'Casi nunca', 'Nunca'];

// Indices for Generic Analysis (Palacio/Pierre)
const INDICES_GENERIC = {
    GENDER: 4,
    AGE: 5,
    CIVIL: 6,
    SCHOOL: 7,
    DEPT: 8,
    TYPE: 9,
    TENURE: 12, // 12. Tiempo en Puesto
    EXPERIENCE: 13, // 13. Tiempo de Experiencia
    QUESTIONS_START: 14
};

// Indices for Princess Analysis
const INDICES_PRINCESS = {
    GENDER: 4,
    AGE: 5,
    CIVIL: 6,
    SCHOOL: 7,
    PROPIEDAD: 8, // Exclusive to Princess
    DEPT: 9,
    TYPE: 10,
    TENURE: 13, // 12. Tiempo en Puesto
    EXPERIENCE: 14, // 13. Tiempo de Experiencia
    QUESTIONS_START: 15
};

const DEMO_MAP_KEYS = {
    '3. Género': 'GENDER',
    '4. Edad': 'AGE',
    '5. Estado Civil': 'CIVIL',
    '6. Nivel de Estudios': 'SCHOOL',
    '8. Departamento': 'DEPT', // Shared number
    '9. Tipo de Puesto': 'TYPE',
    '12. Tiempo en Puesto': 'TENURE',
    '13. Tiempo de Experiencia Laboral': 'EXPERIENCE'
};

module.exports = {
    SHEET_DATA_NAME,
    QUESTION_TYPE_ROW_INDEX,
    HEADER_ROW_INDEX,
    DATA_START_INDEX,
    SCORE_MAP,
    DISTRIBUTION_KEYS,
    INDICES_GENERIC,
    INDICES_PRINCESS,
    DEMO_MAP_KEYS
};

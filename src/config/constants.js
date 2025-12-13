
// Sheet config
const SHEET_DATA_NAME = 'Worksheet';
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
    TENURE: 12,
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
    TENURE: 13,
    QUESTIONS_START: 15
};

// Demographic Column Mappings (Key -> Index Key in INDICES object)
const DEMO_MAP_KEYS = {
    'Género': 'GENDER',
    'Edad': 'AGE',
    'Estado Civil': 'CIVIL',
    'Nivel de Estudios': 'SCHOOL',
    'Tipo de Puesto': 'TYPE',
    'Tiempo en Puesto': 'TENURE'
};

module.exports = {
    SHEET_DATA_NAME,
    HEADER_ROW_INDEX,
    DATA_START_INDEX,
    SCORE_MAP,
    DISTRIBUTION_KEYS,
    INDICES_GENERIC,
    INDICES_PRINCESS,
    DEMO_MAP_KEYS
};

PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS paciente (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    servicio TEXT,
    dni_hc TEXT,
    cuenta TEXT,
    sexo TEXT,
    cama TEXT,
    fecha_nacimiento TEXT,
    edad_anios INTEGER,
    edad_meses INTEGER,
    fecha_ingreso TEXT,
    fecha_alta TEXT,
    dx_medico TEXT,
    grado_instruccion_padre TEXT,
    grado_instruccion_madre TEXT,
    fecha_registro TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS antecedentes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    paciente_id INTEGER NOT NULL REFERENCES paciente(id) ON DELETE CASCADE,
    edad_gestacional TEXT,
    parto TEXT,
    peso_nacer TEXT,
    perimetro_cefalico_nacer TEXT,
    longitud_nacer TEXT,
    antecedentes_familiares TEXT
);

CREATE TABLE IF NOT EXISTS signos_clinicos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    paciente_id INTEGER NOT NULL REFERENCES paciente(id) ON DELETE CASCADE,
    signo TEXT NOT NULL,
    signo_clinico TEXT,
    probable_alteracion TEXT
);

CREATE TABLE IF NOT EXISTS interaccion_farmaco (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    paciente_id INTEGER NOT NULL REFERENCES paciente(id) ON DELETE CASCADE,
    farmaco TEXT,
    via TEXT,
    interaccion TEXT,
    recomendacion TEXT
);

CREATE TABLE IF NOT EXISTS evaluacion_bioquimica (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    paciente_id INTEGER NOT NULL REFERENCES paciente(id) ON DELETE CASCADE,
    prueba TEXT NOT NULL,
    valor_normal TEXT,
    resultado TEXT
);

CREATE INDEX IF NOT EXISTS idx_signos_paciente ON signos_clinicos(paciente_id);
CREATE INDEX IF NOT EXISTS idx_farmaco_paciente ON interaccion_farmaco(paciente_id);
CREATE INDEX IF NOT EXISTS idx_bioquimica_paciente ON evaluacion_bioquimica(paciente_id);
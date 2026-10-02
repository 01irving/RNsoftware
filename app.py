import csv
import io
import json
import math
import re
import sqlite3
import tkinter as tk
from datetime import date, datetime, timedelta
from decimal import Decimal
from tkinter import ttk, messagebox, simpledialog, filedialog
from pathlib import Path

try:
    import lactmed as _LACTMED
    LACTMED_OK = True
except ImportError:
    _LACTMED = None
    LACTMED_OK = False

try:
    from pygrowup import Observation as _PGObservation, exceptions as _PGEx
    PIGROWUP_OK = True
except ImportError:
    _PGObservation = None
    _PGEx = None
    PIGROWUP_OK = False

try:
    import anthro as _ANTHRO
    from anthro import anthro as _ANTHRO_MOD
    ANTHRO_OK = True
except ImportError:
    _ANTHRO = None
    _ANTHRO_MOD = None
    ANTHRO_OK = False

try:
    from openpyxl import Workbook as _XlsxWorkbook
    from openpyxl.styles import Font as _XlsxFont, Alignment as _XlsxAlignment, PatternFill as _XlsxFill
    from openpyxl.utils import get_column_letter as _xlsx_col
    OPENPYXL_OK = True
except ImportError:
    _XlsxWorkbook = None
    _XlsxFont = _XlsxAlignment = _XlsxFill = None
    _xlsx_col = None
    OPENPYXL_OK = False

DB = Path(__file__).resolve().parent / "hc_nutricional.db"

# Clinical Assessment of Nutritional Status score (Metcoff 1994). Cada uno de los
# 9 signos se puntúa de 1 a 4; total 9-36. Malnutrición fetal si el total < 25.
# Fuente: Martínez-Nadal S et al. An Pediatr (Barc). 2016;84(4):218-223, tabla 1.
CANS_PUNTAJE_MALNUTRICION = 25
CANS_SIGNOS = [
    (
        "Cabello",
        "Calidad y docilidad del pelo",
        {
            4: "Abundante, cubre todo el cuero cabelludo. Se peina con facilidad",
            3: "Moderada cantidad. Algunos lisos, se peina con facilidad",
            2: "Escasa cantidad. Cabello liso, se peina con dificultad",
            1: "Escasa cantidad, áreas sin pelo. Cabello liso, no se puede peinar",
        },
    ),
    (
        "Mejillas",
        "Forma de la cara y adiposidad en los carrillos",
        {
            4: "Redonda. Abundante adiposidad",
            3: "Cuadrada. Moderada adiposidad",
            2: "Ovalada. Escasa adiposidad",
            1: "Triangular. Sin adiposidad",
        },
    ),
    (
        "Barbilla y cuello",
        "Perfil de la barbilla y el cuello",
        {
            4: "Pliegues adiposos doble o triple, sin cuello",
            3: "Un solo pliegue. Se insinúa cuello sin arrugas",
            2: "Sin pliegues. Cuello bien definido",
            1: "Sin pliegues. Cuello con piel laxa y arrugas",
        },
    ),
    (
        "Brazos",
        "Coger con ambas manos el brazo y el codo, mirando la zona del tríceps, "
        "comprimir hacia el centro y observar arrugas",
        {
            4: "Sin arrugas",
            3: "Escasas arrugas superficiales",
            2: "3 a 5 arrugas gruesas",
            1: "Arrugas en acordeón",
        },
    ),
    (
        "Tórax",
        "Observar prominencias del tórax y espacio intercostal",
        {
            4: "Tórax lleno, no se aprecian las costillas",
            3: "Se insinúan algunas costillas y leves espacios intercostales debajo de las mamilas",
            2: "Se aprecian costillas y espacios intercostales debajo de las mamilas",
            1: "Costillas prominentes con pérdida del tejido intercostal",
        },
    ),
    (
        "Pliegues de la pared abdominal",
        "Observar adiposidad y consistencia de la piel",
        {
            4: "Abdomen lleno, redondo sin piel laxa",
            3: "Abdomen plano sin piel laxa con uno o 2 pliegues en la región supraumbilical",
            2: "Abdomen delgado. Pliegues en todo el abdomen",
            1: "Abdomen distendido o excavado con piel laxa, fácil de levantar, pliegues en acordeón",
        },
    ),
    (
        "Espalda",
        "Pinzar suavemente con el pulgar e índice la zona interescapular o subescapular "
        "intentando elevar la piel y el tejido subcutáneo",
        {
            4: "Difícil de sujetar y elevar",
            3: "Elevación de 5-10 mm. Pliegue grueso",
            2: "Elevación de 10-20 mm. Pliegue delgado",
            1: "Elevación > 20 mm. Pliegue delgado y laxo",
        },
    ),
    (
        "Glúteos",
        "Observar glúteos y zona posterosuperior del muslo",
        {
            4: "Cojinetes adiposos redondos y llenos",
            3: "Cojinetes aplanados, sin arrugas en glúteos ni muslos",
            2: "Tejido subcutáneo delgado. Arrugas no profundas en glúteos y muslos",
            1: "Tejido subcutáneo escaso, con piel laxa y arrugas profundas",
        },
    ),
    (
        "Piernas",
        "Coger con ambas manos, mirando la región anterior de la pierna. Fijar el pie y "
        "comprimir desde la rodilla con la finalidad de formar arrugas",
        {
            4: "Sin arrugas",
            3: "Escasas arrugas y no profundas",
            2: "3 a 5 arrugas gruesas",
            1: "Múltiples arrugas en acordeón",
        },
    ),
]
CANS_PUNTOS = (4, 3, 2, 1)
CANS_MAXIMO = len(CANS_SIGNOS) * max(CANS_PUNTOS)
CANS_MINIMO = len(CANS_SIGNOS) * min(CANS_PUNTOS)
# Indice ponderal de Rohrer: IP = peso (g) x 100 / longitud^3 (cm). El articulo
# considera malnutricion un IP < 2,2 g/cm3 (IP 2,3 = p10 e IP 2,2 = p3 de peso).
IP_CORTE_MALNUTRICION = 2.2
IP_P10 = 2.3
IP_P3 = 2.2
# Curva de indice ponderal segun edad gestacional (ambos sexos, semanas 33-42).
# Caiza ME et al. An Pediatr (Barc) 2003;59(1):48-53, tabla 3 (p10/p50/p90).
IP_REFERENCIA_EG = {
    33: (2.29, 2.64, 3.05), 34: (2.34, 2.65, 3.03), 35: (2.34, 2.66, 3.06),
    36: (2.39, 2.71, 3.05), 37: (2.40, 2.74, 3.10), 38: (2.45, 2.78, 3.14),
    39: (2.49, 2.78, 3.16), 40: (2.50, 2.81, 3.16), 41: (2.50, 2.81, 3.17),
    42: (2.52, 2.83, 3.17),
}
# Tabla 5 de Caiza et al. (2003): combinacion de IP (bajo/normal/elevado) y
# P/EG (PEG/AEG/GEG) con su letra y definicion de patron de crecimiento.
TABLA5_PATRONES = {
    ("Bajo", "PEG"): ("A", "Retraso de crecimiento intrauterino asimétrico"),
    ("Bajo", "AEG"): ("C", "Retraso de crecimiento intrauterino subclínico"),
    ("Bajo", "GEG"): ("X", "Patrón extremadamente infrecuente"),
    ("Normal", "PEG"): ("B", "Retraso de crecimiento intrauterino simétrico"),
    ("Normal", "AEG"): ("N", "Normal"),
    ("Normal", "GEG"): ("E", "Grande constitucional"),
    ("Elevado", "PEG"): ("X", "Patrón extremadamente infrecuente"),
    ("Elevado", "AEG"): ("D", "Baja talla genética (?)"),
    ("Elevado", "GEG"): ("F", "Obesidad neonatal o hijo de madre con diabetes gestacional (?)"),
}

# Desarrollo motor grueso OMS-MGRS: seis hitos evaluados desde los 4 meses
# hasta caminar solo. Fuente: WHO Multicentre Growth Reference Study Group.
# WHO Motor Development Study. Acta Paediatr Suppl. 2006;450:86-95.
DESARROLLO_MOTOR_HITOS = [
    "Sentarse sin apoyo",
    "Pararse con apoyo",
    "Gatear en manos y rodillas",
    "Caminar con apoyo",
    "Pararse solo",
    "Caminar solo",
]
DESARROLLO_MOTOR_PCT_REFERENCIA = (3, 50, 97)
DESARROLLO_MOTOR_DEFINICIONES = {
    "Sentarse sin apoyo": "Sentarse erguido con la cabeza derecha sin apoyarse en brazos, "
                          "manos, tronco ni otro soporte, durante al menos 10 segundos.",
    "Pararse con apoyo": "Pararse erguido sosteniéndose de un objeto fijo o de una persona, "
                         "durante al menos 10 segundos.",
    "Gatear en manos y rodillas": "Desplazarse en cuadrupedia (manos y rodillas) "
                                  "al menos 2 veces seguidas con el tronco elevado.",
    "Caminar con apoyo": "Caminar erguido sosteniéndose de un objeto fijo o de una persona, "
                         "dando al menos 5 pasos con apoyo.",
    "Pararse solo": "Pararse erguido sin apoyo alguno durante al menos 10 segundos.",
    "Caminar solo": "Caminar sin apoyo durante al menos 5 pasos, con los brazos libres.",
}

PRUEBAS = [
    ("Albúmina", "3.5-5.5 g/dl"),
    ("Pre-albúmina", ">18 mg/dl"),
    ("Hemoglobina", "según edad"),
    ("Leucocitos", "5-15 x10^3/uL"),
    ("% Linfocitos", "20-40 %"),
    ("R.T. Linfocitos", ">2000 cel/mm3"),
    ("Plaquetas", "150-350 x10^3/uL"),
    ("Neutrófilos", "1.8-8.2 x10^3/uL"),
    ("Proteínas Totales", "6-8 g/dl"),
    ("Glucosa", "60-100 mg/dl"),
    ("Urea", "10-50 mg/dl"),
    ("Creatinina", "0.3-1.0 mg/dl"),
    ("Sodio", "132-145 mEq/L"),
    ("Potasio", "3.1-5.1 mEq/L"),
    ("Cloro", "96-111 mEq/L"),
    ("Calcio Sérico", "9.2-11 mg/dl"),
    ("T.G.O.", "13-60 U/L"),
    ("T.G.P.", "28-100 U/L"),
    ("BUN", "5-18 mg/dl"),
    ("PCR", "<5 mg/dl"),
    ("Bilirrubina Directa", "0-0.3 mg/dl"),
    ("Fosfatasa Alcalina", "40-350 U/L"),
    ("Globulinas", "2.3-3.5 g/dl"),
    ("GGT", "10-50 U/L"),
    ("DHL", "100-300 U/L"),
    ("Fósforo", "0.5-2.2 mmol/L"),
    ("Amilasa", "25-125 U/L"),
    ("Lipasa", "<200 U/L"),
    ("Triglicéridos", "<150 mg/dl"),
    ("Colesterol", "<200 mg/dl"),
    ("PCT", "<0.5 ng/ml"),
    ("Lactato", "0.5-2.2 mmol/L"),
]

BIOQ_COLUMNAS_EDAD = ["Nacimiento", "1 semana", "1 mes", "1 1/2 mes"]

# Ecuaciones de estimación de necesidades energéticas (transcritas de las
# tablas de las pp. 54-55): FAO/WHO/UNU 2001 para lactantes e Institute of
# Medicine 2005 para recién nacidos.
ENERGIA_GRUPO_LACTANTE = "Lactante (FAO/WHO/UNU, 2001)"
ENERGIA_GRUPO_RN = "Recién nacido (Institute of Medicine, 2005)"
ENERGIA_GRUPOS = [ENERGIA_GRUPO_LACTANTE, ENERGIA_GRUPO_RN]
ENERGIA_SEXOS = ["Niño", "Niña"]
ENERGIA_ALIMENTACIONES = ["Al seno materno", "Sucedáneo de leche humana"]
ENERGIA_ALIMENTACION_TODAS = "Todos"
ENERGIA_ALIMENTACIONES_OPCIONES = [ENERGIA_ALIMENTACION_TODAS] + ENERGIA_ALIMENTACIONES
ENERGIA_ALIMENTACION_NO_APLICA = "No aplica (ecuación para recién nacido)"
# (a, b, c): energía = a * peso_kg + b + c. Ecuaciones en kilocalorías, tal
# como figuran en la tabla original (p. 54).
ENERGIA_LACTANTE = {
    "Al seno materno": (92.8, -152.0, 196),
    "Sucedáneo de leche humana": (82.6, -29.0, 196),
    ENERGIA_ALIMENTACION_TODAS: (88.6, -99.4, 196),
}
ENERGIA_RN = {"Niño": (89.0, -100.0, 180), "Niña": (89.0, -100.0, 164)}
ENERGIA_NOTA = (
    "Nota: las ecuaciones están en kilocalorías y se transcribieron de las imágenes "
    "aportadas (pp. 54-55). En el lactante la ecuación depende del tipo de alimentación "
    "y no del sexo; la tabla etiqueta «Seno materno» como niños y niñas, «Sucedáneo» como "
    "niña y una tercera fila como «Todos». Verifique la tabla original antes de usarlas "
    "en decisiones clínicas. Esta herramienta estima requerimientos energéticos y no "
    "sustituye la valoración individual."
)

# Encuesta «Taking a Feeding History»: preguntas y verificaciones cruzadas
# para una historia de alimentación precisa. Las secciones se muestran según
# el grupo etario y el tipo de alimentación declarados.
RECUENTO_GRUPO_LACTANTE = "Lactante"
RECUENTO_GRUPO_MAYOR = "Niño mayor"
RECUENTO_GRUPOS = [RECUENTO_GRUPO_LACTANTE, RECUENTO_GRUPO_MAYOR]
RECUENTO_TIPO_MATERNA = "Lactancia materna"
RECUENTO_TIPO_SUCEDANEO = "Sucedáneo de leche humana"
RECUENTO_TIPO_MIXTA = "Lactancia materna y succeedáneo"
RECUENTO_TIPO_OTRA = "Otra leche o dieta"
RECUENTO_TIPO_NO_APLICA = "No aplica (niño mayor)"
RECUENTO_TIPOS_LACTANTE = [
    RECUENTO_TIPO_MATERNA, RECUENTO_TIPO_SUCEDANEO, RECUENTO_TIPO_MIXTA,
    RECUENTO_TIPO_OTRA,
]
RECUENTO_TIPOS_TODOS = RECUENTO_TIPOS_LACTANTE + [RECUENTO_TIPO_NO_APLICA]
RECUENTO_SI_NO = ["Sí", "No"]
RECUENTO_TIPO_TEXTO = "texto"
RECUENTO_TIPO_SI_NO = "si_no"
RECUENTO_TIPO_NUMERO = "numero"
RECUENTO_TIPO_LARGO = "largo"
RECUENTO_INTERVALOS = ["Cada 2 h", "Cada 3 h", "Cada 4 h", "Irregular"]
# (clave, etiqueta, tipo de campo). Se guardan solo las contestadas.
RECUENTO_SECCIONES = [
    ("A. Lactante alimentado con leche materna", "materna", [
        ("lm_frecuencia", "¿Con qué frecuencia se amamanta y cuánto tiempo en cada pecho?",
         RECUENTO_TIPO_TEXTO),
        ("lm_posicion", "¿Se verificaron posición y técnica de la lactancia?", RECUENTO_TIPO_SI_NO),
        ("lm_suplementos", "¿Se ofrecen mamaderas suplementarias u otros alimentos?",
         RECUENTO_TIPO_SI_NO),
        ("lm_suplementos_detalle", "¿Qué mamaderas o alimentos se ofrecen además?",
         RECUENTO_TIPO_TEXTO),
    ]),
    ("B. Lactante alimentado con succeedáneo", "formula", [
        ("f_tipo_formula", "¿Qué tipo de fórmula se utiliza?", RECUENTO_TIPO_TEXTO),
        ("f_preparacion", "¿Cómo se prepara la toma y a qué concentración? (concentrado o "
                         "diluida,Verifyor el contenido energético final)", RECUENTO_TIPO_TEXTO),
        ("f_energia_100ml", "Contenido energético final de la mezcla (kcal/100 ml)",
         RECUENTO_TIPO_NUMERO),
        ("f_fresca", "¿Se prepara cada toma en el momento?", RECUENTO_TIPO_SI_NO),
        ("f_tomas_24h", "¿Cuántas tomas se ofrecen en 24 h?", RECUENTO_TIPO_NUMERO),
        ("f_intervalo", "¿Con qué frecuencia se ofrecen las tomas?", RECUENTO_TIPO_TEXTO),
        ("f_volumen_ofrecido", "¿Qué volumen se ofrece en cada toma (ml)?", RECUENTO_TIPO_TEXTO),
        ("f_volumen_tomado", "¿Cuánto se toma realmente de lo ofrecido?", RECUENTO_TIPO_TEXTO),
        ("f_duracion", "¿Cuánto tarda la toma?", RECUENTO_TIPO_TEXTO),
        ("f_agregados", "¿Se agrega algo más al biberón?", RECUENTO_TIPO_SI_NO),
        ("f_agregados_detalle", "¿Qué se agrega al biberón?", RECUENTO_TIPO_TEXTO),
    ]),
    ("C. Niño mayor", "mayor", [
        ("n_comidas_meriendas", "¿Cuántas comidas y meriendas come al día?", RECUENTO_TIPO_NUMERO),
        ("n_patron", "¿Qué come en cada comida y merienda? (muestra de 1 o 2 días)",
         RECUENTO_TIPO_LARGO),
        ("n_apetito", "¿Cómo describen los padres el apetito del niño?", RECUENTO_TIPO_TEXTO),
        ("n_lugar", "¿Dónde ingiere las comidas?", RECUENTO_TIPO_TEXTO),
        ("n_comidas_familiares", "¿Hay comidas familiares?", RECUENTO_TIPO_SI_NO),
        ("n_ambiente", "¿Son situaciones agradables y placeleras?", RECUENTO_TIPO_SI_NO),
        ("n_leche", "¿Cuánta leche ingiere?", RECUENTO_TIPO_TEXTO),
        ("n_jugo", "¿Cuánto jugo ingiere?", RECUENTO_TIPO_TEXTO),
        ("n_meriendas", "¿Con qué frecuencia come meriendas o alimentos de merienda?",
         RECUENTO_TIPO_TEXTO),
    ]),
]
RECUENTO_NOTA = (
    "Nota: una historia de alimentación cuidadosa es parte esencial de la "
    "valoración nutricional. Las preguntas y verificaciones cruzadas son las de "
    "«Taking a Feeding History»; obtenga una muestra de 1 o 2 días del patrón "
    "alimentario cuando el niño sea mayor. Este registro documenta la entrevista y "
    "no sustituye la valoración nutricional individual."
)


CSV_DEFAULT = {
    "ninos.csv": """Edad_gestacional_semanas,Percentil_10_peso_g,Percentil_50_peso_g,Percentil_90_peso_g
28,815,1147,1470
29,881,1317,1660
30,1065,1511,1800
31,1230,1615,1994
32,1364,1768,2228
33,1553,1986,2498
34,1784,2246,2692
35,1908,2442,2987
36,2168,2777,3300
37,2450,2957,3514
38,2641,3135,3690
39,2747,3254,3800
40,2825,3332,3900
41,2875,3402,3950
42,2890,3484,4100
""",
    "ninas.csv": """Edad_gestacional_semanas,Percentil_10_peso_g,Percentil_50_peso_g,Percentil_90_peso_g
28,846,1037,1352
29,854,1165,1576
30,1030,1348,1740
31,1200,1512,1910
32,1390,1730,2120
33,1588,1958,2406
34,1786,2143,2694
35,1879,2343,2862
36,2120,2635,3174
37,2379,2857,3386
38,2580,3040,3588
39,2700,3153,3682
40,2760,3247,3800
41,2788,3267,3825
42,2800,3277,3978
""",
    "clasificaciones.csv": """Percentil,Interpretación
< 10,Pequeño para la edad de gestación
10 a 90,Apropiado para la edad de gestación
> 90,Grande para la edad de gestación
""",
    "ECRN_ninos.csv": """intervalo_dias,percentil,peso_nacer_2000_2500_g,peso_nacer_2500_3000_g,peso_nacer_3000_3500_g,peso_nacer_3500_4000_g,peso_nacer_4000_mas_g,todos_g
0-7,50,150,150,150,150,50,150
0-7,25,*,0,0,0,-50,0
0-7,10,*,-150,-150,-250,-250,-150
0-7,5,*,-200,-250,-300,-250,-250
0-7,n,7,88,142,100,46,383
7-14,50,275,250,250,250,275,250
7-14,25,*,150,150,100,150,150
7-14,10,*,0,50,0,50,0
7-14,5,*,-100,-50,-50,-100,-50
7-14,n,6,88,141,100,46,381
14-28,50,600,700,650,700,725,650
14-28,25,*,550,550,500,550,550
14-28,10,*,450,450,400,400,450
14-28,5,*,450,350,350,400,350
14-28,n,7,95,154,113,48,417
28-42,50,600,550,550,550,548,550
28-42,25,*,500,450,450,450,450
28-42,10,*,350,350,350,300,350
28-42,5,*,300,300,300,300,300
28-42,n,7,95,156,113,46,417
42-60,50,450,650,650,650,611,650
42-60,25,*,550,500,500,400,500
42-60,10,*,450,400,400,300,400
42-60,5,*,450,350,350,217,350
42-60,n,7,96,153,113,47,416
""",
    "ECRN_velocidad_ninos.csv": """intervalo_dias,percentil,peso_nacer_2000_2500_g,peso_nacer_2500_3000_g,peso_nacer_3000_3500_g,peso_nacer_3500_4000_g,peso_nacer_4000_mas_g,todos_g
0-7,Mediana,21,21,21,21,7,21
0-7,P25,*,0,0,0,-7,0
0-7,P10,*,-21,-21,-36,-36,-21
0-7,P5,*,-29,-36,-43,-36,-36
0-7,n,7,88,142,100,46,383
7-14,Mediana,40,36,33,31,36,36
7-14,P25,*,21,19,14,25,19
7-14,P10,*,0,6,0,6,0
7-14,P5,*,-14,-7,-7,-14,-7
7-14,n,6,88,141,100,46,381
14-28,Mediana,43,50,46,50,50,47
14-28,P25,*,39,39,36,37,38
14-28,P10,*,34,30,29,33,32
14-28,P5,*,32,25,23,29,25
14-28,n,7,95,154,113,48,417
28-42,Mediana,40,42,40,41,40,40
28-42,P25,*,36,31,33,31,32
28-42,P10,*,27,25,24,21,25
28-42,P5,*,21,21,21,21,21
28-42,n,7,95,156,113,46,417
42-60,Mediana,24,35,34,34,34,34
42-60,P25,*,29,28,26,23,28
42-60,P10,*,25,21,22,15,22
42-60,P5,*,24,17,19,14,18
42-60,n,7,96,153,113,47,416
""",
    "ECRN_velocidad_ninas.csv": """intervalo_dias,percentil,peso_nacer_2000_2500_g,peso_nacer_2500_3000_g,peso_nacer_3000_3500_g,peso_nacer_3500_4000_g,peso_nacer_4000_mas_g,todos_g
0-7,Mediana,0,21,14,14,21,14
0-7,P25,*,0,0,0,0,0
0-7,P10,*,-14,-14,-21,-14,-14
0-7,P5,*,-21,-29,-36,-29,-29
0-7,n,18,109,147,85,25,384
7-14,Mediana,29,29,29,29,29,29
7-14,P25,*,14,14,14,14,14
7-14,P10,*,0,0,0,7,0
7-14,P5,*,-12,-7,-14,0,-7
7-14,n,18,108,147,84,25,382
14-28,Mediana,36,43,39,42,44,39
14-28,P25,*,33,32,32,31,32
14-28,P10,*,29,25,23,22,25
14-28,P5,*,21,21,18,17,21
14-28,n,20,124,176,93,28,441
28-42,Mediana,36,36,35,32,38,35
28-42,P25,*,27,28,25,26,27
28-42,P10,*,23,21,18,21,21
28-42,P5,*,21,18,15,21,18
28-42,n,20,127,174,92,28,441
42-60,Mediana,29,31,27,32,29,29
42-60,P25,*,23,21,23,20,22
42-60,P10,*,19,18,19,9,18
42-60,P5,*,17,15,13,9,15
42-60,n,18,127,175,92,28,440
""",
    "ECRN_ninas.csv": """intervalo_dias,percentil,peso_nacer_2000_2500_g,peso_nacer_2500_3000_g,peso_nacer_3000_3500_g,peso_nacer_3500_4000_g,peso_nacer_4000_mas_g,todos_g
0-7,50,0,150,100,100,150,100
0-7,25,*,0,0,0,0,0
0-7,10,*,-100,-100,-150,-100,-100
0-7,5,*,-150,-200,-250,-200,-200
0-7,n,18,109,147,85,25,384
7-14,50,200,200,200,200,200,200
7-14,25,*,100,100,100,100,100
7-14,10,*,0,0,0,50,0
7-14,5,*,-100,-50,-100,0,-50
7-14,n,18,108,147,84,25,382
14-28,50,500,600,550,550,600,550
14-28,25,*,450,436,450,450,450
14-28,10,*,400,350,300,300,350
14-28,5,*,300,300,250,200,300
14-28,n,20,124,176,93,28,441
28-42,50,500,500,465,457,525,500
28-42,25,*,382,400,325,375,382
28-42,10,*,300,300,295,300,300
28-42,5,*,300,250,200,300,250
28-42,n,20,127,174,92,28,441
42-60,50,550,550,500,585,550,550
42-60,25,*,400,400,408,334,400
42-60,10,*,300,300,350,155,300
42-60,5,*,300,289,250,150,288
42-60,n,18,127,175,92,28,440
""",
    "tabla_peso_para_la_edad.csv": """sexo,edad_dias,edad_etiqueta,P3,P15,P50,P85,P97
M,0,nacimiento,2.5,2.9,3.3,3.7,4.4
M,30,1 mes,3.4,3.9,4.5,5.0,5.8
M,61,2 meses,4.3,4.9,5.6,6.3,7.1
M,91,3 meses,5.0,5.7,6.4,7.2,8.0
M,122,4 meses,5.6,6.3,7.0,7.8,8.7
M,152,5 meses,6.0,6.7,7.5,8.3,9.2
M,183,6 meses,6.4,7.1,7.9,8.8,9.8
M,213,7 meses,6.7,7.4,8.3,9.2,10.3
M,244,8 meses,6.9,7.7,8.6,9.6,10.7
M,274,9 meses,7.1,7.9,8.9,9.9,11.0
M,305,10 meses,7.4,8.2,9.2,10.2,11.4
M,335,11 meses,7.6,8.4,9.4,10.5,11.7
M,365,12 meses,7.7,8.6,9.6,10.7,12.0
F,0,nacimiento,2.44,2.78,3.23,3.73,4.17
F,30,1 mes,3.20,3.61,4.17,4.80,5.37
F,61,2 meses,4.01,4.48,5.13,5.87,6.53
F,91,3 meses,4.60,5.12,5.84,6.65,7.40
F,122,4 meses,5.09,5.65,6.43,7.31,8.12
F,152,5 meses,5.48,6.07,6.90,7.83,8.70
F,183,6 meses,5.82,6.44,7.30,8.29,9.20
F,213,7 meses,6.09,6.74,7.64,8.68,9.64
F,244,8 meses,6.35,7.02,7.95,9.03,10.04
F,274,9 meses,6.56,7.26,8.23,9.35,10.39
F,305,10 meses,6.77,7.49,8.48,9.64,10.73
F,335,11 meses,6.96,7.69,8.72,9.92,11.04
F,365,12 meses,7.14,7.89,8.95,10.18,11.33
""",
    "tabla_velocidad_crecimiento.csv": """sexo,ventana,edad_ini_dias,edad_fin_dias,peso_nacer_min_kg,peso_nacer_max_kg,P3,P15,P50,P85,P97,nota
M,0-7 días,0,7,0,2.5,,,-0.5,2.5,7.0,* casos insuficientes para estimar percentiles inferiores
M,0-7 días,0,7,2.5,3.0,,,-0.3,2.8,7.0,*
M,0-7 días,0,7,3.0,3.5,,,-1.0,2.2,6.5,*
M,0-7 días,0,7,3.5,4.0,,,-2.0,1.5,6.0,*
M,0-7 días,0,7,4.0,4.5,,,-3.0,0.8,5.5,*
M,0-7 días,0,7,4.5,99.9,,,-4.0,0.0,5.0,*
M,7-14 días,7,14,0,2.5,0,8,17,26,34,
M,7-14 días,7,14,2.5,3.0,1,9,18,27,35,
M,7-14 días,7,14,3.0,3.5,0,8,18,28,37,
M,7-14 días,7,14,3.5,4.0,-1,8,18,29,39,
M,7-14 días,7,14,4.0,4.5,-2,7,18,30,41,
M,7-14 días,7,14,4.5,99.9,-3,6,17,30,42,
M,14-28 días,14,28,0,2.5,8,17,27,37,47,
M,14-28 días,14,28,2.5,3.0,9,18,28,38,48,
M,14-28 días,14,28,3.0,3.5,10,19,30,40,50,
M,14-28 días,14,28,3.5,4.0,11,20,30,41,52,
M,14-28 días,14,28,4.0,4.5,12,20,31,42,53,
M,14-28 días,14,28,4.5,99.9,13,21,31,43,54,
M,1-2 meses,28,61,0,99.9,15,23,31,39,47,
M,2-3 meses,61,91,0,99.9,12,19,26,33,40,
M,3-4 meses,91,122,0,99.9,10,17,23,30,36,
M,4-5 meses,122,152,0,99.9,8,15,20,26,32,
M,5-6 meses,152,183,0,99.9,7,12,18,23,29,
M,6-9 meses,183,274,0,99.9,5,9,13,17,21,
M,9-12 meses,274,365,0,99.9,3,6,9,12,16,
F,0-7 días,0,7,0,2.5,,,-1.0,2.0,6.5,*
F,0-7 días,0,7,2.5,3.0,,,-0.8,2.3,6.5,*
F,0-7 días,0,7,3.0,3.5,,,-1.5,1.7,6.0,*
F,0-7 días,0,7,3.5,4.0,,,-2.5,1.0,5.5,*
F,0-7 días,0,7,4.0,4.5,,,-3.5,0.3,5.0,*
F,0-7 días,0,7,4.5,99.9,,,-4.5,-0.5,4.5,*
F,7-14 días,7,14,0,2.5,-1,7,16,25,33,
F,7-14 días,7,14,2.5,3.0,0,8,17,26,34,
F,7-14 días,7,14,3.0,3.5,-1,7,17,27,36,
F,7-14 días,7,14,3.5,4.0,-2,7,17,28,38,
F,7-14 días,7,14,4.0,4.5,-3,6,17,29,40,
F,7-14 días,7,14,4.5,99.9,-4,5,16,29,41,
F,14-28 días,14,28,0,2.5,7,16,26,36,46,
F,14-28 días,14,28,2.5,3.0,8,17,27,37,47,
F,14-28 días,14,28,3.0,3.5,9,18,29,39,49,
F,14-28 días,14,28,3.5,4.0,10,19,29,40,51,
F,14-28 días,14,28,4.0,4.5,11,19,30,41,52,
F,14-28 días,14,28,4.5,99.9,12,20,30,42,53,
F,1-2 meses,28,61,0,99.9,14,22,29,37,44,
F,2-3 meses,61,91,0,99.9,11,18,25,31,38,
F,3-4 meses,91,122,0,99.9,9,15,21,27,33,
F,4-5 meses,122,152,0,99.9,7,13,19,24,30,
F,5-6 meses,152,183,0,99.9,6,11,16,21,26,
F,6-9 meses,183,274,0,99.9,4,8,12,15,19,
F,9-12 meses,274,365,0,99.9,3,5,8,10,14,
""",
    "tabla_velocidad_oms.csv": """indicador,sexo,intervalo_meses,mes_inicio,mes_fin,L,M,S,delta
pc,F,2,0,2,0.8807,4.3539,0.15953,0.0
pc,F,2,1,3,0.8807,3.1035,0.16706,0.0
pc,F,2,2,4,0.8807,2.3473,0.18107,0.0
pc,F,2,3,5,0.8807,1.9599,0.20017,0.0
pc,F,2,4,6,0.8807,1.6524,0.22476,0.0
pc,F,2,5,7,0.8807,1.3981,0.25281,0.0
pc,F,2,6,8,0.8807,1.1762,0.28607,0.0
pc,F,2,7,9,0.8807,0.9921,0.32161,0.0
pc,F,2,8,10,0.8807,0.8471,0.35933,0.0
pc,F,2,9,11,0.8807,0.7384,0.40034,0.0
pc,F,2,10,12,0.8807,0.6552,0.44193,0.0
pc,F,3,0,3,0.4252,5.5822,0.14155,0.0
pc,F,3,1,4,0.5408,4.1837,0.14611,0.0
pc,F,3,2,5,0.6316,3.2594,0.15573,0.0
pc,F,3,3,6,0.7066,2.7161,0.16952,0.0
pc,F,3,4,7,0.7704,2.2991,0.18602,0.0
pc,F,3,5,8,0.8262,1.9451,0.20505,0.0
pc,F,3,6,9,0.8756,1.6363,0.22644,0.0
pc,F,3,7,10,0.9201,1.3903,0.2492,0.0
pc,F,3,8,11,0.9605,1.2025,0.27282,0.0
pc,F,3,9,12,0.9976,1.069,0.29743,0.0
pc,F,4,0,4,0.5215,6.6605,0.12888,0.0
pc,F,4,1,5,0.5902,5.075,0.13444,0.0
pc,F,4,2,6,0.6469,4.0154,0.14318,0.0
pc,F,4,3,7,0.6952,3.3542,0.15397,0.0
pc,F,4,4,8,0.7374,2.8419,0.16676,0.0
pc,F,4,5,9,0.7748,2.3945,0.18192,0.0
pc,F,4,6,10,0.8085,2.0271,0.19915,0.0
pc,F,4,7,11,0.8391,1.7353,0.21812,0.0
pc,F,4,8,12,0.8672,1.5073,0.23888,0.0
pc,F,4,9,13,0.8931,1.3296,0.26147,0.0
pc,F,4,10,14,0.9172,1.1864,0.28575,0.0
pc,F,4,11,15,0.9397,1.0642,0.31149,0.0
pc,F,4,12,16,0.9608,0.9573,0.33813,0.0
pc,F,4,13,17,0.9808,0.8664,0.36514,0.0
pc,F,4,14,18,0.9996,0.794,0.39238,0.0
pc,F,4,15,19,1.0175,0.7411,0.41976,0.0
pc,F,4,16,20,1.0344,0.6953,0.44691,0.0
pc,F,4,17,21,1.0506,0.6457,0.4735,0.0
pc,F,4,18,22,1.0661,0.5953,0.49958,0.0
pc,F,4,19,23,1.081,0.5473,0.52538,0.0
pc,F,4,20,24,1.0952,0.5013,0.55112,0.0
pc,F,6,0,6,0.2608,8.2683,0.11607,0.0
pc,F,6,1,7,0.3426,6.5037,0.11997,0.0
pc,F,6,2,8,0.414,5.2206,0.1259,0.0
pc,F,6,3,9,0.4774,4.341,0.13381,0.0
pc,F,6,4,10,0.5344,3.6748,0.14324,0.0
pc,F,6,5,11,0.5863,3.1336,0.15447,0.0
pc,F,6,6,12,0.6338,2.689,0.16777,0.0
pc,F,6,7,13,0.6777,2.3226,0.18258,0.0
pc,F,6,8,14,0.7186,2.0288,0.19851,0.0
pc,F,6,9,15,0.7567,1.7978,0.21524,0.0
pc,F,6,10,16,0.7925,1.6098,0.23248,0.0
pc,F,6,11,17,0.8262,1.45,0.24989,0.0
pc,F,6,12,18,0.8581,1.3166,0.26722,0.0
pc,F,6,13,19,0.8884,1.2111,0.28419,0.0
pc,F,6,14,20,0.9172,1.1249,0.30032,0.0
pc,F,6,15,21,0.9446,1.0506,0.31516,0.0
pc,F,6,16,22,0.9708,0.9819,0.32846,0.0
pc,F,6,17,23,0.9959,0.9149,0.34024,0.0
pc,F,6,18,24,1.02,0.8494,0.35116,0.0
pc,M,2,0,2,0.9267,4.6878,0.16093,0.0
pc,M,2,1,3,0.621,3.3714,0.15634,0.0
pc,M,2,2,4,0.5607,2.517,0.16382,0.0
pc,M,2,3,5,0.6219,2.0747,0.18097,0.0
pc,M,2,4,6,0.7141,1.7184,0.207,0.0
pc,M,2,5,7,0.7879,1.4381,0.23798,0.0
pc,M,2,6,8,0.8482,1.2009,0.27392,0.0
pc,M,2,7,9,0.8985,1.0106,0.31173,0.0
pc,M,2,8,10,0.9379,0.8731,0.35212,0.0
pc,M,2,9,11,0.9577,0.7615,0.39591,0.0
pc,M,2,10,12,0.9598,0.6659,0.44007,0.0
pc,M,3,0,3,0.7558,6.0419,0.13725,0.0
pc,M,3,1,4,0.4737,4.5132,0.13624,0.0
pc,M,3,2,5,0.4137,3.4944,0.14382,0.0
pc,M,3,3,6,0.4858,2.8568,0.15846,0.0
pc,M,3,4,7,0.5992,2.384,0.17799,0.0
pc,M,3,5,8,0.7074,1.9903,0.20021,0.0
pc,M,3,6,9,0.8123,1.6704,0.22358,0.0
pc,M,3,7,10,0.9028,1.4214,0.24758,0.0
pc,M,3,8,11,0.9602,1.2321,0.27193,0.0
pc,M,3,9,12,0.9852,1.0768,0.29648,0.0
pc,M,4,0,4,0.3279,7.1133,0.1244,0.0
pc,M,4,1,5,0.4212,5.4806,0.12713,0.0
pc,M,4,2,6,0.5076,4.2702,0.13454,0.0
pc,M,4,3,7,0.584,3.5185,0.14592,0.0
pc,M,4,4,8,0.6557,2.9255,0.16106,0.0
pc,M,4,5,9,0.7179,2.4506,0.17841,0.0
pc,M,4,6,10,0.7719,2.078,0.19784,0.0
pc,M,4,7,11,0.818,1.7696,0.21918,0.0
pc,M,4,8,12,0.8532,1.5301,0.24082,0.0
pc,M,4,9,13,0.8808,1.3368,0.26394,0.0
pc,M,4,10,14,0.9009,1.1844,0.28656,0.0
pc,M,4,11,15,0.9166,1.0457,0.3097,0.0
pc,M,4,12,16,0.9278,0.9301,0.33162,0.0
pc,M,4,13,17,0.9357,0.8359,0.35292,0.0
pc,M,4,14,18,0.9414,0.7605,0.37421,0.0
pc,M,4,15,19,0.9452,0.707,0.39409,0.0
pc,M,4,16,20,0.948,0.6637,0.41387,0.0
pc,M,4,17,21,0.95,0.6277,0.43237,0.0
pc,M,4,18,22,0.9518,0.5923,0.45102,0.0
pc,M,4,19,23,0.9535,0.5575,0.46942,0.0
pc,M,4,20,24,0.955,0.5248,0.48711,0.0
pc,M,6,0,6,0.4441,8.864,0.1101,0.0
pc,M,6,1,7,0.4988,6.9439,0.11688,0.0
pc,M,6,2,8,0.5465,5.5066,0.12498,0.0
pc,M,6,3,9,0.5888,4.526,0.13512,0.0
pc,M,6,4,10,0.6269,3.7983,0.14708,0.0
pc,M,6,5,11,0.6616,3.2122,0.16026,0.0
pc,M,6,6,12,0.6933,2.7354,0.17427,0.0
pc,M,6,7,13,0.7226,2.3511,0.18884,0.0
pc,M,6,8,14,0.7499,2.0461,0.20363,0.0
pc,M,6,9,15,0.7754,1.7985,0.21837,0.0
pc,M,6,10,16,0.7993,1.5932,0.2328,0.0
pc,M,6,11,17,0.8218,1.4211,0.24672,0.0
pc,M,6,12,18,0.8431,1.2795,0.26021,0.0
pc,M,6,13,19,0.8633,1.1676,0.27339,0.0
pc,M,6,14,20,0.8825,1.0785,0.28632,0.0
pc,M,6,15,21,0.9009,1.0078,0.29903,0.0
pc,M,6,16,22,0.9184,0.9483,0.31159,0.0
pc,M,6,17,23,0.9351,0.8976,0.32406,0.0
pc,M,6,18,24,0.9512,0.8511,0.33648,0.0
peso,F,1,0,1,0.7781,1279.4834,0.21479,400
peso,F,1,1,2,0.7781,1411.1075,0.19384,400
peso,F,1,2,3,0.7781,1118.0098,0.19766,400
peso,F,1,3,4,0.7781,984.8825,0.20995,400
peso,F,1,4,5,0.7781,888.9803,0.22671,400
peso,F,1,5,6,0.7781,801.391,0.24596,400
peso,F,1,6,7,0.7781,744.3023,0.26515,400
peso,F,1,7,8,0.7781,710.6923,0.28409,400
peso,F,1,8,9,0.7781,672.6072,0.30106,400
peso,F,1,9,10,0.7781,644.6032,0.31676,400
peso,F,1,10,11,0.7781,633.2166,0.33208,400
peso,F,1,11,12,0.7781,631.7383,0.34627,400
peso,F,2,0,2,0.4599,2497.0406,0.18,600
peso,F,2,1,3,0.3294,2314.2285,0.17612,600
peso,F,2,2,4,0.3128,1907.0116,0.17761,600
peso,F,2,3,5,0.356,1673.5778,0.18421,600
peso,F,2,4,6,0.4264,1482.7466,0.19524,600
peso,F,2,5,7,0.5002,1342.3734,0.20864,600
peso,F,2,6,8,0.5699,1251.4869,0.22315,600
peso,F,2,7,9,0.6268,1181.4135,0.23586,600
peso,F,2,8,10,0.673,1116.8192,0.2468,600
peso,F,2,9,11,0.7102,1078.3961,0.25656,600
peso,F,2,10,12,0.7382,1058.4112,0.26494,600
peso,F,2,11,13,0.7605,1040.8737,0.27292,600
peso,F,2,12,14,0.7762,1027.9459,0.28011,600
peso,F,2,13,15,0.7864,1019.687,0.28705,600
peso,F,2,14,16,0.7913,1016.4898,0.29343,600
peso,F,2,15,17,0.7922,1017.5335,0.29961,600
peso,F,2,16,18,0.7902,1017.2241,0.30592,600
peso,F,2,17,19,0.7866,1012.8511,0.31201,600
peso,F,2,18,20,0.7827,1007.2711,0.31824,600
peso,F,2,19,21,0.7795,1001.8324,0.32415,600
peso,F,2,20,22,0.7771,993.3265,0.33014,600
peso,F,2,21,23,0.7755,980.7096,0.33605,600
peso,F,2,22,24,0.7743,967.2057,0.34166,600
peso,F,3,0,3,0.2298,3403.924,0.16227,800
peso,F,3,1,4,0.0924,3054.3512,0.15958,800
peso,F,3,2,5,0.0599,2618.644,0.16338,800
peso,F,3,3,6,0.13,2277.5681,0.1699,800
peso,F,3,4,7,0.2404,2030.2917,0.1796,800
peso,F,3,5,8,0.358,1855.0162,0.19157,800
peso,F,3,6,9,0.4576,1724.5802,0.20334,800
peso,F,3,7,10,0.5317,1624.4588,0.2135,800
peso,F,3,8,11,0.5891,1552.7117,0.22168,800
peso,F,3,9,12,0.6373,1506.412,0.22796,800
peso,F,3,10,13,0.6806,1476.5227,0.23285,800
peso,F,3,11,14,0.7211,1455.9527,0.23682,800
peso,F,3,12,15,0.7527,1442.0871,0.2404,800
peso,F,3,13,16,0.7679,1434.2381,0.24403,800
peso,F,3,14,17,0.7642,1431.1099,0.24794,800
peso,F,3,15,18,0.7482,1429.1551,0.25198,800
peso,F,3,16,19,0.7267,1425.3256,0.25598,800
peso,F,3,17,20,0.7032,1418.4764,0.25989,800
peso,F,3,18,21,0.6782,1409.2288,0.26384,800
peso,F,3,19,22,0.6522,1398.1693,0.26792,800
peso,F,3,20,23,0.6262,1385.3711,0.27191,800
peso,F,3,21,24,0.6013,1370.5464,0.27539,800
peso,F,4,0,4,0.0891,4009.5248,0.15636,800
peso,F,4,1,5,-0.0491,3541.1498,0.16054,800
peso,F,4,2,6,-0.0362,3002.2985,0.16706,800
peso,F,4,3,7,0.0557,2616.8194,0.17682,800
peso,F,4,4,8,0.1899,2331.5089,0.18968,800
peso,F,4,5,9,0.3154,2118.7222,0.20292,800
peso,F,4,6,10,0.4069,1963.5334,0.21481,800
peso,F,4,7,11,0.4873,1855.0205,0.22441,800
peso,F,4,8,12,0.5659,1780.1251,0.2312,800
peso,F,4,9,13,0.6483,1726.3407,0.23622,800
peso,F,4,10,14,0.7201,1691.6549,0.2398,800
peso,F,4,11,15,0.7759,1668.6291,0.24275,800
peso,F,4,12,16,0.8064,1655.384,0.24522,800
peso,F,4,13,17,0.8111,1648.5922,0.24757,800
peso,F,4,14,18,0.7924,1643.852,0.25015,800
peso,F,4,15,19,0.7577,1638.7127,0.25299,800
peso,F,4,16,20,0.7139,1631.2189,0.25623,800
peso,F,4,17,21,0.6699,1621.0212,0.25952,800
peso,F,4,18,22,0.6227,1607.4454,0.2629,800
peso,F,4,19,23,0.5732,1591.7399,0.26613,800
peso,F,4,20,24,0.5249,1576.2729,0.26914,800
peso,F,6,0,6,-0.1223,4528.9831,0.15945,450
peso,F,6,1,7,-0.028,3911.9319,0.17265,450
peso,F,6,2,8,0.0799,3327.7315,0.18755,450
peso,F,6,3,9,0.1942,2853.38,0.20514,450
peso,F,6,4,10,0.3097,2501.5063,0.22466,450
peso,F,6,5,11,0.4246,2248.588,0.24383,450
peso,F,6,6,12,0.525,2068.2742,0.25997,450
peso,F,6,7,13,0.6042,1939.2944,0.27156,450
peso,F,6,8,14,0.6644,1850.4715,0.27943,450
peso,F,6,9,15,0.7065,1793.3361,0.28481,450
peso,F,6,10,16,0.7288,1758.5512,0.2887,450
peso,F,6,11,17,0.7317,1738.3567,0.29175,450
peso,F,6,12,18,0.7206,1725.0429,0.29439,450
peso,F,6,13,19,0.7016,1713.8691,0.29696,450
peso,F,6,14,20,0.6812,1703.1167,0.29971,450
peso,F,6,15,21,0.6643,1691.6943,0.30278,450
peso,F,6,16,22,0.6534,1677.6772,0.30619,450
peso,F,6,17,23,0.6489,1659.966,0.30991,450
peso,F,6,18,24,0.6476,1640.7438,0.31376,450
peso,M,1,0,1,1.3828,1423.0783,0.22048,400
peso,M,1,1,2,0.7241,1596.347,0.19296,400
peso,M,1,2,3,0.659,1215.3989,0.19591,400
peso,M,1,3,4,0.7003,1017.0488,0.20965,400
peso,M,1,4,5,0.7419,921.6249,0.2279,400
peso,M,1,5,6,0.7668,822.1842,0.24854,400
peso,M,1,6,7,0.7688,756.5306,0.26783,400
peso,M,1,7,8,0.7624,715.6257,0.28677,400
peso,M,1,8,9,0.762,684.7459,0.30439,400
peso,M,1,9,10,0.7659,658.5809,0.32154,400
peso,M,1,10,11,0.7713,643.4374,0.33882,400
peso,M,1,11,12,0.7761,639.4743,0.35502,400
peso,M,2,0,2,0.7188,2815.612,0.17422,600
peso,M,2,1,3,0.6464,2592.0761,0.17025,600
peso,M,2,2,4,0.6071,2038.1036,0.17559,600
peso,M,2,3,5,0.5915,1744.8197,0.18708,600
peso,M,2,4,6,0.5891,1541.367,0.2013,600
peso,M,2,5,7,0.5954,1377.6979,0.21318,600
peso,M,2,6,8,0.6088,1272.5277,0.22426,600
peso,M,2,7,9,0.627,1201.4599,0.23472,600
peso,M,2,8,10,0.6486,1143.8903,0.24611,600
peso,M,2,9,11,0.6725,1101.6312,0.25918,600
peso,M,2,10,12,0.6959,1077.9049,0.27217,600
peso,M,2,11,13,0.7191,1057.9071,0.28462,600
peso,M,2,12,14,0.7399,1037.0541,0.29479,600
peso,M,2,13,15,0.7597,1014.185,0.30285,600
peso,M,2,14,16,0.7771,1000.5821,0.30864,600
peso,M,2,15,17,0.7929,999.4661,0.3129,600
peso,M,2,16,18,0.8078,1000.968,0.31615,600
peso,M,2,17,19,0.821,998.4215,0.31858,600
peso,M,2,18,20,0.8335,992.804,0.32058,600
peso,M,2,19,21,0.8447,986.9799,0.32222,600
peso,M,2,20,22,0.8554,981.7965,0.32377,600
peso,M,2,21,23,0.8655,978.4016,0.32529,600
peso,M,2,22,24,0.8748,976.3696,0.32673,600
peso,M,3,0,3,0.6854,3638.873,0.15801,650
peso,M,3,1,4,0.6503,3215.101,0.16539,650
peso,M,3,2,5,0.5884,2661.5629,0.17708,650
peso,M,3,3,6,0.5368,2231.9042,0.1885,650
peso,M,3,4,7,0.4999,1939.0717,0.19877,650
peso,M,3,5,8,0.4819,1745.5952,0.20848,650
peso,M,3,6,9,0.4866,1611.6464,0.21853,650
peso,M,3,7,10,0.5135,1514.8958,0.2294,650
peso,M,3,8,11,0.5582,1442.6013,0.24108,650
peso,M,3,9,12,0.6092,1387.884,0.25261,650
peso,M,3,10,13,0.658,1346.3553,0.26315,650
peso,M,3,11,14,0.7,1314.9304,0.27214,650
peso,M,3,12,15,0.7323,1291.3726,0.27922,650
peso,M,3,13,16,0.755,1273.886,0.28446,650
peso,M,3,14,17,0.7695,1261.0053,0.28821,650
peso,M,3,15,18,0.7769,1251.6296,0.29074,650
peso,M,3,16,19,0.7781,1244.9248,0.29231,650
peso,M,3,17,20,0.774,1240.2027,0.29311,650
peso,M,3,18,21,0.7663,1235.8993,0.2935,650
peso,M,3,19,22,0.7569,1229.8975,0.29388,650
peso,M,3,20,23,0.7475,1220.6029,0.2946,650
peso,M,3,21,24,0.7393,1206.8517,0.29591,650
peso,M,4,0,4,0.7672,4136.2992,0.15684,500
peso,M,4,1,5,0.6482,3623.4564,0.17439,500
peso,M,4,2,6,0.5632,2900.447,0.19057,500
peso,M,4,3,7,0.4863,2424.1094,0.2039,500
peso,M,4,4,8,0.4302,2106.5547,0.21598,500
peso,M,4,5,9,0.4321,1871.4914,0.22766,500
peso,M,4,6,10,0.4881,1711.6071,0.24076,500
peso,M,4,7,11,0.5825,1598.0178,0.25575,500
peso,M,4,8,12,0.6678,1526.3463,0.27024,500
peso,M,4,9,13,0.7242,1473.6287,0.2835,500
peso,M,4,10,14,0.7587,1423.7181,0.29393,500
peso,M,4,11,15,0.7822,1370.5468,0.30204,500
peso,M,4,12,16,0.7966,1334.9524,0.30757,500
peso,M,4,13,17,0.8054,1321.4376,0.31114,500
peso,M,4,14,18,0.8084,1315.2621,0.31327,500
peso,M,4,15,19,0.8041,1308.5472,0.31429,500
peso,M,4,16,20,0.7912,1297.8646,0.31482,500
peso,M,4,17,21,0.7712,1284.7539,0.3152,500
peso,M,4,18,22,0.7469,1271.459,0.31566,500
peso,M,4,19,23,0.7222,1262.1643,0.31628,500
peso,M,4,20,24,0.6991,1257.3339,0.31696,500
peso,M,6,0,6,0.5209,4929.7718,0.15679,350
peso,M,6,1,7,0.4856,4243.2925,0.17552,350
peso,M,6,2,8,0.4609,3442.915,0.19228,350
peso,M,6,3,9,0.449,2879.5905,0.20802,350
peso,M,6,4,10,0.4511,2501.8054,0.22426,350
peso,M,6,5,11,0.466,2220.6833,0.24197,350
peso,M,6,6,12,0.4895,2037.9406,0.26076,350
peso,M,6,7,13,0.5168,1903.183,0.27848,350
peso,M,6,8,14,0.5442,1794.7774,0.29319,350
peso,M,6,9,15,0.5697,1709.1588,0.30394,350
peso,M,6,10,16,0.5943,1651.415,0.31109,350
peso,M,6,11,17,0.619,1616.6162,0.31517,350
peso,M,6,12,18,0.6428,1590.6081,0.31683,350
peso,M,6,13,19,0.6649,1571.5549,0.31675,350
peso,M,6,14,20,0.6849,1557.0267,0.31549,350
peso,M,6,15,21,0.7027,1545.9058,0.31347,350
peso,M,6,16,22,0.7187,1533.6871,0.31113,350
peso,M,6,17,23,0.7336,1520.616,0.30878,350
peso,M,6,18,24,0.7478,1508.4744,0.30647,350
talla,F,2,0,2,0.9918,7.9023,0.14123,0.0
talla,F,2,1,3,0.9918,6.3775,0.15004,0.0
talla,F,2,2,4,0.9918,5.1574,0.17732,0.0
talla,F,2,3,5,0.9918,4.2877,0.21092,0.0
talla,F,2,4,6,0.9918,3.5965,0.23941,0.0
talla,F,2,5,7,0.9918,3.1827,0.25995,0.0
talla,F,2,6,8,0.9918,3.0,0.27597,0.0
talla,F,2,7,9,0.9918,2.8764,0.28638,0.0
talla,F,2,8,10,0.9918,2.7444,0.29192,0.0
talla,F,2,9,11,0.9918,2.6284,0.29751,0.0
talla,F,2,10,12,0.9918,2.5303,0.30553,0.0
talla,F,2,11,13,0.9918,2.4425,0.31612,0.0
talla,F,2,12,14,0.9918,2.3621,0.32828,0.0
talla,F,2,13,15,0.9918,2.2879,0.34112,0.0
talla,F,2,14,16,0.9918,2.2236,0.35425,0.0
talla,F,2,15,17,0.9918,2.1684,0.36737,0.0
talla,F,2,16,18,0.9918,2.1113,0.38003,0.0
talla,F,2,17,19,0.9918,2.047,0.39199,0.0
talla,F,2,18,20,0.9918,1.9822,0.40358,0.0
talla,F,2,19,21,0.9918,1.9225,0.41519,0.0
talla,F,2,20,22,0.9918,1.8682,0.42686,0.0
talla,F,2,21,23,0.9918,1.8192,0.43859,0.0
talla,F,2,22,24,0.9918,1.775,0.45033,0.0
talla,F,3,0,3,0.8538,10.5967,0.11683,0.0
talla,F,3,1,4,0.8538,8.7743,0.13505,0.0
talla,F,3,2,5,0.8538,7.1455,0.15574,0.0
talla,F,3,3,6,0.8538,5.9428,0.17798,0.0
talla,F,3,4,7,0.8538,5.1554,0.19661,0.0
talla,F,3,5,8,0.8538,4.6834,0.20988,0.0
talla,F,3,6,9,0.8538,4.3922,0.21849,0.0
talla,F,3,7,10,0.8538,4.1971,0.22383,0.0
talla,F,3,8,11,0.8538,4.0329,0.22876,0.0
talla,F,3,9,12,0.8538,3.8692,0.23503,0.0
talla,F,3,10,13,0.8538,3.7174,0.24257,0.0
talla,F,3,11,14,0.8538,3.5892,0.25105,0.0
talla,F,3,12,15,0.8538,3.4811,0.25988,0.0
talla,F,3,13,16,0.8538,3.3844,0.26843,0.0
talla,F,3,14,17,0.8538,3.2934,0.27635,0.0
talla,F,3,15,18,0.8538,3.2051,0.28388,0.0
talla,F,3,16,19,0.8538,3.1173,0.2913,0.0
talla,F,3,17,20,0.8538,3.0295,0.29869,0.0
talla,F,3,18,21,0.8538,2.9427,0.30582,0.0
talla,F,3,19,22,0.8538,2.8576,0.31251,0.0
talla,F,3,20,23,0.8538,2.7779,0.31896,0.0
talla,F,3,21,24,0.8538,2.7091,0.32567,0.0
talla,F,4,0,4,0.8123,13.0081,0.10744,0.0
talla,F,4,1,5,0.8123,10.6621,0.12126,0.0
talla,F,4,2,6,0.8123,8.7302,0.13625,0.0
talla,F,4,3,7,0.8123,7.4606,0.15069,0.0
talla,F,4,4,8,0.8123,6.5992,0.16368,0.0
talla,F,4,5,9,0.8123,6.0664,0.1724,0.0
talla,F,4,6,10,0.8123,5.7273,0.17782,0.0
talla,F,4,7,11,0.8123,5.4731,0.18189,0.0
talla,F,4,8,12,0.8123,5.2575,0.18562,0.0
talla,F,4,9,13,0.8123,5.055,0.18974,0.0
talla,F,4,10,14,0.8123,4.8763,0.19407,0.0
talla,F,4,11,15,0.8123,4.7084,0.19893,0.0
talla,F,4,12,16,0.8123,4.5658,0.20385,0.0
talla,F,4,13,17,0.8123,4.4427,0.2088,0.0
talla,F,4,14,18,0.8123,4.3256,0.21372,0.0
talla,F,4,15,19,0.8123,4.2141,0.21816,0.0
talla,F,4,16,20,0.8123,4.0974,0.2224,0.0
talla,F,4,17,21,0.8123,3.9825,0.22623,0.0
talla,F,4,18,22,0.8123,3.866,0.22998,0.0
talla,F,4,19,23,0.8123,3.7559,0.23357,0.0
talla,F,4,20,24,0.8123,3.6558,0.23694,0.0
talla,F,6,0,6,0.7138,16.4915,0.09904,0.0
talla,F,6,1,7,0.7138,13.8733,0.10884,0.0
talla,F,6,2,8,0.7138,11.8137,0.11821,0.0
talla,F,6,3,9,0.7138,10.3499,0.12639,0.0
talla,F,6,4,10,0.7138,9.3426,0.1329,0.0
talla,F,6,5,11,0.7138,8.677,0.13782,0.0
talla,F,6,6,12,0.7138,8.2244,0.14171,0.0
talla,F,6,7,13,0.7138,7.8787,0.14512,0.0
talla,F,6,8,14,0.7138,7.5879,0.14836,0.0
talla,F,6,9,15,0.7138,7.3259,0.15166,0.0
talla,F,6,10,16,0.7138,7.0897,0.15514,0.0
talla,F,6,11,17,0.7138,6.8778,0.1588,0.0
talla,F,6,12,18,0.7138,6.6823,0.16252,0.0
talla,F,6,13,19,0.7138,6.4984,0.16617,0.0
talla,F,6,14,20,0.7138,6.3217,0.16964,0.0
talla,F,6,15,21,0.7138,6.1484,0.17287,0.0
talla,F,6,16,22,0.7138,5.977,0.17591,0.0
talla,F,6,17,23,0.7138,5.8083,0.17884,0.0
talla,F,6,18,24,0.7138,5.6454,0.18169,0.0
talla,M,2,0,2,0.9497,8.482,0.134,0.0
talla,M,2,1,3,0.9497,6.9984,0.14062,0.0
talla,M,2,2,4,0.9497,5.5716,0.17179,0.0
talla,M,2,3,5,0.9497,4.4941,0.20929,0.0
talla,M,2,4,6,0.9497,3.7228,0.24323,0.0
talla,M,2,5,7,0.9497,3.2403,0.26837,0.0
talla,M,2,6,8,0.9497,2.9661,0.28481,0.0
talla,M,2,7,9,0.9497,2.8089,0.29636,0.0
talla,M,2,8,10,0.9497,2.6901,0.30505,0.0
talla,M,2,9,11,0.9497,2.5785,0.31391,0.0
talla,M,2,10,12,0.9497,2.4724,0.324,0.0
talla,M,2,11,13,0.9497,2.3818,0.33613,0.0
talla,M,2,12,14,0.9497,2.2978,0.34908,0.0
talla,M,2,13,15,0.9497,2.2138,0.36174,0.0
talla,M,2,14,16,0.9497,2.1357,0.3741,0.0
talla,M,2,15,17,0.9497,2.0675,0.38645,0.0
talla,M,2,16,18,0.9497,2.0061,0.39924,0.0
talla,M,2,17,19,0.9497,1.9495,0.41274,0.0
talla,M,2,18,20,0.9497,1.8972,0.42656,0.0
talla,M,2,19,21,0.9497,1.849,0.44029,0.0
talla,M,2,20,22,0.9497,1.803,0.45398,0.0
talla,M,2,21,23,0.9497,1.7575,0.46768,0.0
talla,M,2,22,24,0.9497,1.7133,0.48129,0.0
talla,M,3,0,3,0.8792,11.4458,0.11285,0.0
talla,M,3,1,4,0.8792,9.495,0.127,0.0
talla,M,3,2,5,0.8792,7.6058,0.15474,0.0
talla,M,3,3,6,0.8792,6.2317,0.18096,0.0
talla,M,3,4,7,0.8792,5.3243,0.20103,0.0
talla,M,3,5,8,0.8792,4.7433,0.21513,0.0
talla,M,3,6,9,0.8792,4.3594,0.22535,0.0
talla,M,3,7,10,0.8792,4.1002,0.23308,0.0
talla,M,3,8,11,0.8792,3.92,0.23935,0.0
talla,M,3,9,12,0.8792,3.7818,0.24526,0.0
talla,M,3,10,13,0.8792,3.6611,0.25157,0.0
talla,M,3,11,14,0.8792,3.543,0.25876,0.0
talla,M,3,12,15,0.8792,3.4189,0.26713,0.0
talla,M,3,13,16,0.8792,3.292,0.27641,0.0
talla,M,3,14,17,0.8792,3.1717,0.2859,0.0
talla,M,3,15,18,0.8792,3.0649,0.29508,0.0
talla,M,3,16,19,0.8792,2.9758,0.30351,0.0
talla,M,3,17,20,0.8792,2.9068,0.31089,0.0
talla,M,3,18,21,0.8792,2.8507,0.31767,0.0
talla,M,3,19,22,0.8792,2.794,0.32487,0.0
talla,M,3,20,23,0.8792,2.7265,0.33332,0.0
talla,M,3,21,24,0.8792,2.6405,0.34377,0.0
talla,M,4,0,4,1.0138,13.977,0.10113,0.0
talla,M,4,1,5,1.0138,11.4886,0.12006,0.0
talla,M,4,2,6,1.0138,9.3048,0.13954,0.0
talla,M,4,3,7,1.0138,7.7601,0.15624,0.0
talla,M,4,4,8,1.0138,6.7018,0.16955,0.0
talla,M,4,5,9,1.0138,6.0704,0.1778,0.0
talla,M,4,6,10,1.0138,5.6756,0.18311,0.0
talla,M,4,7,11,1.0138,5.3939,0.18754,0.0
talla,M,4,8,12,1.0138,5.1699,0.19138,0.0
talla,M,4,9,13,1.0138,4.9623,0.19512,0.0
talla,M,4,10,14,1.0138,4.7773,0.1988,0.0
talla,M,4,11,15,1.0138,4.6014,0.20286,0.0
talla,M,4,12,16,1.0138,4.4487,0.20707,0.0
talla,M,4,13,17,1.0138,4.3123,0.21158,0.0
talla,M,4,14,18,1.0138,4.1833,0.21644,0.0
talla,M,4,15,19,1.0138,4.068,0.22124,0.0
talla,M,4,16,20,1.0138,3.9584,0.22619,0.0
talla,M,4,17,21,1.0138,3.86,0.23092,0.0
talla,M,4,18,22,1.0138,3.7663,0.23577,0.0
talla,M,4,19,23,1.0138,3.6815,0.24063,0.0
talla,M,4,20,24,1.0138,3.6058,0.2453,0.0
talla,M,6,0,6,0.9027,17.6547,0.09452,0.0
talla,M,6,1,7,0.9027,14.711,0.10935,0.0
talla,M,6,2,8,0.9027,12.3097,0.12383,0.0
talla,M,6,3,9,0.9027,10.5768,0.1357,0.0
talla,M,6,4,10,0.9027,9.4,0.14407,0.0
talla,M,6,5,11,0.9027,8.6282,0.14919,0.0
talla,M,6,6,12,0.9027,8.1114,0.15162,0.0
talla,M,6,7,13,0.9027,7.7366,0.15255,0.0
talla,M,6,8,14,0.9027,7.4335,0.15299,0.0
talla,M,6,9,15,0.9027,7.1621,0.15364,0.0
talla,M,6,10,16,0.9027,6.9165,0.15479,0.0
talla,M,6,11,17,0.9027,6.6927,0.15649,0.0
talla,M,6,12,18,0.9027,6.483,0.15863,0.0
talla,M,6,13,19,0.9027,6.2862,0.16108,0.0
talla,M,6,14,20,0.9027,6.1061,0.16362,0.0
talla,M,6,15,21,0.9027,5.9431,0.1661,0.0
talla,M,6,16,22,0.9027,5.7899,0.16861,0.0
talla,M,6,17,23,0.9027,5.6425,0.17124,0.0
talla,M,6,18,24,0.9027,5.5018,0.17392,0.0""",
    "tabla_desarrollo_motor_oms.csv": """hito,percentil,edad_dias
Sentarse sin apoyo,1,115
Sentarse sin apoyo,3,125
Sentarse sin apoyo,5,131
Sentarse sin apoyo,10,140
Sentarse sin apoyo,25,158
Sentarse sin apoyo,50,179
Sentarse sin apoyo,75,204
Sentarse sin apoyo,90,229
Sentarse sin apoyo,95,245
Sentarse sin apoyo,97,256
Sentarse sin apoyo,99,279
Pararse con apoyo,1,147
Pararse con apoyo,3,160
Pararse con apoyo,5,167
Pararse con apoyo,10,178
Pararse con apoyo,25,200
Pararse con apoyo,50,226
Pararse con apoyo,75,256
Pararse con apoyo,90,287
Pararse con apoyo,95,307
Pararse con apoyo,97,320
Pararse con apoyo,99,348
Gatear en manos y rodillas,1,157
Gatear en manos y rodillas,3,177
Gatear en manos y rodillas,5,187
Gatear en manos y rodillas,10,202
Gatear en manos y rodillas,25,226
Gatear en manos y rodillas,50,254
Gatear en manos y rodillas,75,284
Gatear en manos y rodillas,90,319
Gatear en manos y rodillas,95,345
Gatear en manos y rodillas,97,364
Gatear en manos y rodillas,99,409
Caminar con apoyo,1,181
Caminar con apoyo,3,200
Caminar con apoyo,5,210
Caminar con apoyo,10,225
Caminar con apoyo,25,249
Caminar con apoyo,50,275
Caminar con apoyo,75,304
Caminar con apoyo,90,336
Caminar con apoyo,95,360
Caminar con apoyo,97,378
Caminar con apoyo,99,418
Pararse solo,1,211
Pararse solo,3,235
Pararse solo,5,248
Pararse solo,10,266
Pararse solo,25,296
Pararse solo,50,330
Pararse solo,75,367
Pararse solo,90,408
Pararse solo,95,438
Pararse solo,97,461
Pararse solo,99,514
Caminar solo,1,250
Caminar solo,3,274
Caminar solo,5,286
Caminar solo,10,304
Caminar solo,25,333
Caminar solo,50,365
Caminar solo,75,400
Caminar solo,90,438
Caminar solo,95,466
Caminar solo,97,487
Caminar solo,99,534
""",
    "QS.csv": """Tabla,Prueba,Unidad,Nacimiento,1 semana,1 mes,1 1/2 mes,Notas
Química sanguínea,Proteína total en suero,g/dL,4.4-7.6,,,,
Química sanguínea,Albúmina en suero,g/dL,2.9-5.5,,,,
Química sanguínea,Prealbúmina en suero,mg/dL,,4.0-22.0,9.0-27.0,,
Química sanguínea,Creatinina en suero,mg/dL,0.2-1.2,,,,
Química sanguínea,Urea,mmol/L,1.3-5.1,,,,
Química sanguínea,Colesterol total en suero,mg/dL,50-120,,,,
Química sanguínea,Triglicéridos en suero,mg/dL,20-150,,,,
Química sanguínea,Folato en suero,ng/mL,2.0-15.0,,,,
Química sanguínea,Hierro total en suero,µg/dL,55-150,,,,
Química sanguínea,Magnesio,mmol/L,0.71-0.96,,,,
Química sanguínea,Calcio en sangre ionizada,mmol/L,0.90-1.45,,,,
Química sanguínea,Calcio en suero,mg/dL,Prematuro: 6.0-10.0; a término: 7.0-12.0,,,,
Química sanguínea,Fósforo en suero,mg/dL,Prematuro: 5.6-8.0; a término: 5.0-7.8; también se alcanza a leer 4.8-8.1,,,,
Química sanguínea,Sodio en suero,mmol/L,Prematuro: 132-140; a término: 133-142,,,,
Química sanguínea,Potasio en suero,mmol/L,4.5-7.0,,,,
Biometría hemática,Hemoglobina,g/dL,14.0-22.5,13.5-20.5,11.0-13.0,,
Biometría hemática,Hematocrito,%,47-62,42-62,30-48,,
Biometría hemática,Volumen corpuscular medio,fL,100-135,100-120,84-105,,
Biometría hemática,Hemoglobina corpuscular media,pg,31-37,28-40,24-36,,
Biometría hemática,Concentración de hemoglobina corpuscular media,%,32-36,32-36,32-36,,
Biometría hemática,Leucocitos,×10⁹/L,9-30,5-21,5-19,,
Biometría hemática,Neutrófilos,×10⁹/L,15-25.0,1.5-10.0,1.0-8.0,,
Biometría hemática,Linfocitos,×10⁹/L,2-11,2-17,2-13,,
Biometría hemática,Monocitos,×10⁹/L,0.1-1.7,0.1-1.7,0.1-1.1,,"El valor de 1 mes se lee 0.1-1.1 en la imagen; verificar contra el original."
Biometría hemática,Eosinófilos,×10⁹/L,0.1-1.1,0.1-1.1,0.1-1.1,,
Biometría hemática,Plaquetas,×10⁹/L,150-600,150-600,150-600,,
Examen general de orina,Creatinina en orina,g/24 h,,,,0.8-2.8,La imagen muestra el intervalo 0.8 a 2.8 g/24 h.""",
}


def conectar():
    conn = sqlite3.connect(DB)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _texto_peso_edad(sexo, dias, peso_g):
    """Texto de la evaluación peso-para-la-edad (OMS 2006). None si no se puede calcular."""
    if not sexo or dias is None or dias < 0 or dias > 365 or peso_g is None:
        return None
    try:
        conn = conectar()
        filas = conn.execute(
            "SELECT edad_dias, p3, p15, p50, p85, p97 FROM referencia_peso_edad WHERE sexo = ? ORDER BY edad_dias",
            (sexo,),
        ).fetchall()
        conn.close()
    except Exception:
        return None
    if not filas:
        return None
    fila_ini = None
    fila_fin = None
    for fila in filas:
        if fila[0] <= dias:
            fila_ini = fila
        else:
            fila_fin = fila
            break
    if fila_ini is None:
        return None

    def _interp(iv, fv):
        if fila_fin is None or fv is None or fila_fin[0] == fila_ini[0]:
            return iv
        return iv + (fv - iv) * (dias - fila_ini[0]) / (fila_fin[0] - fila_ini[0])

    p3 = _interp(fila_ini[1], fila_fin[1] if fila_fin else None)
    p15 = _interp(fila_ini[2], fila_fin[2] if fila_fin else None)
    p50 = _interp(fila_ini[3], fila_fin[3] if fila_fin else None)
    p85 = _interp(fila_ini[4], fila_fin[4] if fila_fin else None)
    p97 = _interp(fila_ini[5], fila_fin[5] if fila_fin else None)
    if p3 is None or p15 is None or p85 is None or p97 is None:
        return None
    kg = peso_g / 1000.0
    if kg < p3:
        clas = "PESO BAJO (< P3) - Riesgo de desnutrición"
    elif kg < p15:
        clas = "Peso bajo (P3 a P15) - Riesgo"
    elif kg <= p85:
        clas = "Peso adecuado (P15 a P85)"
    elif kg <= p97:
        clas = "Peso elevado (P85 a P97) - Riesgo de sobrepeso"
    else:
        clas = "PESO ELEVADO (> P97) - Sobrepeso"
    semanas, resto = divmod(dias, 7)
    return (
        f"Edad: {semanas} semanas y {resto} días | Peso: {kg:.2f} kg\n"
        f"Percentil estimado: {clas}\n"
        f"Referencias (OMS): P3 {p3:.2f} | P15 {p15:.2f} | P50 {p50:.2f} | "
        f"P85 {p85:.2f} | P97 {p97:.2f} kg"
    )


def _intervalo_neonatal(dias):
    """Devuelve el intervalo neonatal según la referencia del nacimiento.
    0-7 incluye el día 7; 7-14 incluye el día 14; 14-28 incluye el día 28."""
    if dias is None:
        return None
    if 0 <= dias <= 7:
        return "0-7"
    if 7 < dias <= 14:
        return "7-14"
    if 14 < dias <= 28:
        return "14-28"
    return None


def _intervalo_ganancia(dias):
    """Devuelve el intervalo de ganancia ECRN hasta los 60 días."""
    if dias is None or dias < 0:
        return None
    if dias <= 7:
        return "0-7"
    if dias <= 14:
        return "7-14"
    if dias <= 28:
        return "14-28"
    if dias <= 42:
        return "28-42"
    if dias <= 60:
        return "42-60"
    return None


def _n_muestra_ganancia(sexo, intervalo, peso_nacer_g):
    columnas = {
        "peso_nacer_2000_2500_g": "n_peso_nacer_2000_2500",
        "peso_nacer_2500_3000_g": "n_peso_nacer_2500_3000",
        "peso_nacer_3000_3500_g": "n_peso_nacer_3000_3500",
        "peso_nacer_3500_4000_g": "n_peso_nacer_3500_4000",
        "peso_nacer_4000_mas_g": "n_peso_nacer_4000_mas",
        "todos_g": "n_todos",
    }
    columna = columnas[App._columna_peso_nacer(peso_nacer_g)]
    try:
        conn = conectar()
        fila = conn.execute(
            f"SELECT {columna} FROM referencia_ganancia_peso_muestra "
            "WHERE sexo = ? AND intervalo_dias = ?",
            (sexo, intervalo),
        ).fetchone()
        conn.close()
        return fila[0] if fila else None
    except Exception:
        return None


def _texto_velocidad_ecrn(sexo, dias, peso_nacer_g, peso_actual_g, incluir_detalle=False):
    intervalo = _intervalo_ganancia(dias)
    columna = App._columna_peso_nacer(peso_nacer_g)
    try:
        conn = conectar()
        filas = conn.execute(
            """SELECT percentil, peso_nacer_2000_2500_g, peso_nacer_2500_3000_g,
                      peso_nacer_3000_3500_g, peso_nacer_3500_4000_g,
                      peso_nacer_4000_mas_g, todos_g
               FROM referencia_velocidad_peso_ecrn
               WHERE sexo = ? AND intervalo_dias = ?""",
            (sexo, intervalo),
        ).fetchall()
        conn.close()
    except Exception:
        return None
    if not filas:
        return None
    columna_indice = {
        "peso_nacer_2000_2500_g": 1, "peso_nacer_2500_3000_g": 2,
        "peso_nacer_3000_3500_g": 3, "peso_nacer_3500_4000_g": 4,
        "peso_nacer_4000_mas_g": 5, "todos_g": 6,
    }[columna]
    referencia = {}
    for fila in filas:
        valor = fila[columna_indice]
        if valor is None:
            valor = fila[6]
        referencia[fila[0]] = valor
    p5, p10, p25, mediana = (
        referencia.get("P5"), referencia.get("P10"),
        referencia.get("P25"), referencia.get("Mediana"),
    )
    if None in (p5, p10, p25, mediana):
        return None
    velocidad = (peso_actual_g - peso_nacer_g) / dias
    if velocidad < p5:
        clas = "< P5"
    elif velocidad < p10:
        clas = "P5 - P10"
    elif velocidad < p25:
        clas = "P10 - P25"
    elif velocidad <= mediana:
        clas = "P25 - Mediana"
    else:
        clas = "> Mediana"
    n_muestra = _n_muestra_ganancia(sexo, intervalo, peso_nacer_g)
    muestra_txt = f" | n={n_muestra}" if n_muestra is not None else ""
    texto = (
        f"Intervalo: {intervalo} días | Peso nacer: {peso_nacer_g / 1000:.2f} kg | "
        f"Rango: {columna}{muestra_txt}\n"
        f"Velocidad promedio desde nacimiento: {velocidad:.1f} g/día --> {clas}\n"
        f"Referencias (g/día): P5 {p5:g} | P10 {p10:g} | P25 {p25:g} | Mediana {mediana:g}"
    )
    if incluir_detalle:
        return {"texto": texto, "velocidad_g_dia": velocidad, "percentil": clas}
    return texto


def _texto_velocidad(sexo, dias, peso_nacer_g, peso_actual_g):
    """Evalúa la velocidad con referencias ECRN específicas por sexo."""
    if not sexo or dias is None or dias <= 0 or dias > 60:
        return None
    if peso_nacer_g is None or peso_actual_g is None:
        return None
    if sexo in ("Masculino", "Femenino"):
        return _texto_velocidad_ecrn(sexo, dias, peso_nacer_g, peso_actual_g)
    kg_nacer = peso_nacer_g / 1000.0
    try:
        conn = conectar()
        filas = conn.execute(
            """SELECT ventana, edad_ini_dias, edad_fin_dias, peso_nacer_min_kg,
               peso_nacer_max_kg, p3, p15, p50, p85, p97
               FROM referencia_velocidad_crecimiento
               WHERE sexo = ?
               ORDER BY edad_ini_dias, peso_nacer_min_kg""",
            (sexo,),
        ).fetchall()
        conn.close()
    except Exception:
        return None
    fila = None
    for f in filas:
        if not f[3] <= kg_nacer < f[4]:
            continue
        if dias <= 28:
            ventana_esperada = f"{_intervalo_neonatal(dias)} días"
            if f[0] == ventana_esperada:
                fila = f
                break
        elif f[1] <= dias < f[2]:
            fila = f
            break
    if fila is None:
        return None
    ventana, p3, p15, p50, p85, p97 = fila[0], fila[5], fila[6], fila[7], fila[8], fila[9]
    vel = (peso_actual_g - peso_nacer_g) / dias
    if p3 is None or p15 is None:
        if vel < p50:
            clas = "Inferior a la mediana (< P50)"
        elif vel <= p85:
            clas = "Adecuada (P50 a P85)"
        elif vel <= p97:
            clas = "Alta (P85 a P97)"
        else:
            clas = "Muy alta (> P97)"
        ref = f"P50 {p50:.1f} | P85 {p85:.1f} | P97 {p97:.1f}"
    else:
        if vel < p3:
            clas = "INSUFICIENTE (< P3)"
        elif vel < p15:
            clas = "Insuficiente (P3 a P15)"
        elif vel <= p85:
            clas = "Adecuada (P15 a P85)"
        elif vel <= p97:
            clas = "Alta (P85 a P97)"
        else:
            clas = "MUY ALTA (> P97)"
        ref = f"P3 {p3:.1f} | P15 {p15:.1f} | P50 {p50:.1f} | P85 {p85:.1f} | P97 {p97:.1f}"
    return (
            f"Referencia: {ventana} | Peso nacer: {kg_nacer:.2f} kg | Días: {dias}\n"
            f"Velocidad promedio desde nacimiento: {vel:.1f} g/día --> {clas}\n"
        f"Referencias (g/día): {ref}"
    )


def _percentil_desde_z(z):
    """Convierte un z-score a percentil mediante la distribución normal estándar."""
    try:
        return 100.0 * (1 + math.erf(z / math.sqrt(2))) / 2
    except Exception:
        return None


def _texto_zscore_oms(sexo, dias, peso_g, postura="Acostado", edema=False):
    """Texto del z-score peso-para-la-edad (WHO Anthro, OMS 2006). None si no se puede calcular."""
    if not sexo or dias is None or dias < 0 or dias > 1826 or peso_g is None:
        return None
    peso_kg = peso_g / 1000.0
    z = None
    try:
        if ANTHRO_OK:
            r = _ANTHRO.compute(dict(
                sex="M" if sexo == "Masculino" else "F",
                age_days=dias,
                weight_g=peso_g,
                measure="L" if postura == "Acostado" else "H",
                oedema=edema,
                mode="day",
            ))
            z = r.get("z_wfa")
        elif PIGROWUP_OK:
            base = date(2000, 1, 1)
            obs = _PGObservation(
                sex="male" if sexo == "Masculino" else "female",
                dob=base,
                date_of_observation=base + timedelta(days=dias),
            )
            z = float(obs.wfa(Decimal(str(peso_kg))))
    except Exception:
        z = None
    if z is None or isinstance(z, str):
        return None
    if z < -3:
        clas = "PESO MUY BAJO (< -3 DE)"
    elif z < -2:
        clas = "Bajo peso (< -2 DE)"
    elif z < -1:
        clas = "Riesgo de bajo peso (-2 a -1 DE)"
    elif z <= 1:
        clas = "Normal (entre -1 y +1 DE)"
    elif z <= 2:
        clas = "Riesgo alto (+1 a +2 DE)"
    else:
        clas = "Peso elevado (> +2 DE)"
    pct = _percentil_desde_z(z)
    pct_txt = f" | Percentil: P{pct:.1f}" if pct is not None else ""
    semanas, resto = divmod(dias, 7)
    return (
        f"Edad: {semanas} semanas y {resto} días | Peso: {peso_kg:.2f} kg\n"
        f"Z-score peso/edad (WHO Anthro): {z:.2f}{pct_txt}\n"
        f"{clas}\n"
        f"(OMS 2006 - WHO Anthro, peso para la edad - tabla LMS por día)"
    )


def _clas_anthro(indicador, z):
    """Clasificación WHO (OMS 2006) por indicador según el z-score."""
    if z is None:
        return ""
    if indicador == "wfa":
        if z < -3:
            return "Peso muy bajo (< -3 DE)"
        if z < -2:
            return "Bajo peso (< -2 DE)"
        if z <= 2:
            return "Peso adecuado"
        return "Peso elevado (> +2 DE)"
    if indicador in ("lhfa",):
        if z < -3:
            return "Talla muy baja / retraso severo (< -3 DE)"
        if z < -2:
            return "Talla baja / retraso del crecimiento (< -2 DE)"
        if z <= 2:
            return "Talla adecuada"
        return "Talla alta (> +2 DE)"
    if indicador in ("wfl", "wfh"):
        if z < -3:
            return "Emaciación severa (< -3 DE)"
        if z < -2:
            return "Emaciación / desnutrición aguda (< -2 DE)"
        if z <= 1:
            return "Peso adecuado"
        if z <= 2:
            return "Riesgo de sobrepeso"
        if z <= 3:
            return "Sobrepeso"
        return "Obesidad (> +3 DE)"
    if indicador == "bmifa":
        if z < -3:
            return "Emaciación severa (< -3 DE)"
        if z < -2:
            return "Emaciación / desnutrición aguda (< -2 DE)"
        if z <= 1:
            return "IMC adecuado"
        if z <= 2:
            return "Riesgo de sobrepeso"
        if z <= 3:
            return "Sobrepeso"
        return "Obesidad (> +3 DE)"
    if indicador == "hcfa":
        if z < -2:
            return "Microcefalia (< -2 DE)"
        if z <= 2:
            return "Perímetro cefálico adecuado"
        return "Macrocefalia (> +2 DE)"
    if indicador == "acfa":
        if z < -3:
            return "Delgadez severa (< -3 DE)"
        if z < -2:
            return "Delgadez / desnutrición aguda (< -2 DE)"
        if z <= 2:
            return "MUAC adecuado"
        return "MUAC alto (> +2 DE)"
    if z < -2:
        return "Valor bajo (< -2 DE)"
    if z > 2:
        return "Valor alto (> +2 DE)"
    return "Adecuado"


def _texto_anthro_completo(sexo, dias, peso_g=None, talla_cm=None,
                           pc_cm=None, muac_cm=None, triceps_mm=None,
                           subescapular_mm=None, postura=None, edema=False):
    """Evaluación completa WHO Anthro (OMS 2006): peso/edad, talla/edad,
    peso/talla o peso/altura, PC/edad, MUAC/edad, IMC/edad,
    tríceps/edad y subescapular/edad."""

    def _z_py(metodo, *args, **kwargs):
        try:
            r = metodo(*args, **kwargs)
            return float(r) if isinstance(r, Decimal) else float(r)
        except Exception:
            return None

    if not (ANTHRO_OK or PIGROWUP_OK):
        return None
    if not sexo or dias is None or dias < 0 or dias > 1826:
        return None
    lineas = []
    recostada = postura == "Acostado" if postura else dias < 731
    lineas.append(
        f"Medición: {'Acostado' if recostada else 'Parado'} | Edema: {'Sí' if edema else 'No'}")
    if ANTHRO_OK:
        resultado = _ANTHRO.compute({
            "sex": "M" if sexo == "Masculino" else "F",
            "age_days": int(dias),
            "weight_g": peso_g,
            "height_cm": talla_cm,
            "measure": "L" if recostada else "H",
            "oedema": bool(edema),
            "muac_mm": muac_cm * 10 if muac_cm is not None else None,
            "mode": "day",
        })
        indicadores = (
            ("z_wfa", "Peso/edad", "wfa"),
            ("z_lhfa", "Longitud/edad" if recostada else "Talla/edad", "lhfa"),
            ("z_wflh", "Peso/longitud" if recostada else "Peso/talla", "wfl"),
            ("z_bmi", "IMC/edad", "bmifa"),
            ("z_acfa", "MUAC/edad", "acfa"),
        )
        for clave, etiqueta, clasificador in indicadores:
            z = resultado.get(clave)
            if z is not None:
                pct = _percentil_desde_z(z)
                lineas.append(
                    f"{etiqueta}: z={z:+.2f} | P{pct:.1f} | "
                    f"{_clas_anthro(clasificador, z)}")
        if edema and any("Oedema:" in aviso for aviso in resultado.get("warnings", [])):
            lineas.append(
                "Edema presente: interpretar peso/edad y peso/longitud-talla con cautela (OMS).")
    elif PIGROWUP_OK:
        try:
            obs = _PGObservation(
                sex="male" if sexo == "Masculino" else "female",
                age_in_days=int(dias),
            )
        except Exception:
            base = date(2000, 1, 1)
            obs = _PGObservation(
                sex="male" if sexo == "Masculino" else "female",
                dob=base,
                date_of_observation=base + timedelta(days=int(dias)),
            )
        peso_kg = peso_g / 1000.0 if peso_g is not None else None
        if peso_kg is not None:
            z = _z_py(obs.wfa, Decimal(str(peso_kg)))
            if z is not None:
                pct = _percentil_desde_z(z)
                lineas.append(
                    f"Peso/edad: z={z:+.2f} | P{pct:.1f} | {_clas_anthro('wfa', z)}")
        if talla_cm is not None:
            z = _z_py(obs.lhfa, Decimal(str(talla_cm)), recumbent=recostada)
            if z is not None:
                pct = _percentil_desde_z(z)
                etiqueta = "Longitud/edad" if recostada else "Talla/edad"
                lineas.append(
                    f"{etiqueta}: z={z:+.2f} | P{pct:.1f} | {_clas_anthro('lhfa', z)}")
        if peso_kg is not None and talla_cm is not None:
            metodo_w = obs.wfl if recostada else obs.wfh
            z = _z_py(metodo_w, Decimal(str(peso_kg)), Decimal(str(talla_cm)))
            if z is not None:
                pct = _percentil_desde_z(z)
                etiqueta = "Peso/longitud" if recostada else "Peso/talla"
                lineas.append(
                    f"{etiqueta}: z={z:+.2f} | P{pct:.1f} | {_clas_anthro('wfl', z)}")
            bmi = peso_kg / (talla_cm / 100.0) ** 2
            z = _z_py(obs.bmifa, Decimal(str(round(bmi, 2))))
            if z is not None:
                pct = _percentil_desde_z(z)
                lineas.append(
                    f"IMC/edad: z={z:+.2f} | P{pct:.1f} | {_clas_anthro('bmifa', z)}")
        if muac_cm is not None:
            z = _z_py(obs.acfa, Decimal(str(muac_cm)))
            if z is not None:
                pct = _percentil_desde_z(z)
                lineas.append(
                    f"MUAC/edad: z={z:+.2f} | P{pct:.1f} | {_clas_anthro('acfa', z)}")
    if pc_cm is not None and PIGROWUP_OK:
        try:
            obs_pc = _PGObservation(
                sex="male" if sexo == "Masculino" else "female",
                age_in_days=int(dias),
            )
            z = _z_py(obs_pc.hcfa, Decimal(str(pc_cm)))
        except Exception:
            z = None
        if z is not None:
            pct = _percentil_desde_z(z)
            lineas.append(
                f"Perímetro cefálico/edad: z={z:+.2f} | P{pct:.1f} | "
                f"{_clas_anthro('hcfa', z)}")
    if edema and not ANTHRO_OK:
        lineas.append("Edema presente: pygrowup no permite ajustar los cálculos por edema.")
    if triceps_mm is not None or subescapular_mm is not None:
        obs_pliegues = None
        if PIGROWUP_OK:
            try:
                obs_pliegues = _PGObservation(
                    sex="male" if sexo == "Masculino" else "female",
                    age_in_days=int(dias),
                )
            except Exception:
                obs_pliegues = None
        z_pliegues = {}
        for etiqueta, metodo, valor in (
            ("Tríceps/edad", "tsfa", triceps_mm),
            ("Subescapular/edad", "ssfa", subescapular_mm),
        ):
            if valor is None:
                continue
            try:
                valor_num = float(valor)
            except (TypeError, ValueError):
                valor_num = None
            z = None
            if obs_pliegues is not None and valor_num is not None:
                try:
                    z = _z_py(getattr(obs_pliegues, metodo), Decimal(str(valor_num)))
                except Exception:
                    z = None
            valor_txt = f"{valor_num:g}" if valor_num is not None else str(valor)
            if z is None:
                lineas.append(
                    f"{etiqueta}: {valor_txt} mm | z no disponible "
                    "(edad: 3 meses a 5 años; valor: 1-30 mm)")
            else:
                pct = _percentil_desde_z(z)
                pct_txt = f" | P{pct:.1f}" if pct is not None else ""
                lineas.append(
                    f"{etiqueta}: {valor_txt} mm | z={z:+.2f}{pct_txt} | "
                    f"{_clas_anthro('skinfold', z)}")
                z_pliegues[metodo] = valor_num
        if "tsfa" in z_pliegues and "ssfa" in z_pliegues:
            lineas.append(
                f"Sumatoria tríceps + subescapular: "
                f"{z_pliegues['tsfa'] + z_pliegues['ssfa']:.1f} mm")
    if not lineas:
        return None
    return "\n".join(lineas)


def _texto_ganancia(sexo, dias, peso_nacer_g, peso_actual_g):
    """Texto de la evaluación de ganancia de peso (ECRN). None si no se puede calcular."""
    if not sexo or dias is None or dias < 0 or dias > 60:
        return None
    if peso_nacer_g is None or peso_actual_g is None:
        return None
    intervalo = _intervalo_ganancia(dias)
    if intervalo is None:
        return None
    try:
        conn = conectar()
        filas = conn.execute(
            """SELECT percentil, peso_nacer_2000_2500_g, peso_nacer_2500_3000_g,
               peso_nacer_3000_3500_g, peso_nacer_3500_4000_g, peso_nacer_4000_mas_g, todos_g
               FROM referencia_ganancia_peso
               WHERE sexo = ? AND intervalo_dias = ?
               ORDER BY percentil ASC""",
            (sexo, intervalo),
        ).fetchall()
        conn.close()
    except Exception:
        return None
    if not filas:
        return None
    columna = App._columna_peso_nacer(peso_nacer_g)
    mapa = {
        "peso_nacer_2000_2500_g": 1, "peso_nacer_2500_3000_g": 2,
        "peso_nacer_3000_3500_g": 3, "peso_nacer_3500_4000_g": 4,
        "peso_nacer_4000_mas_g": 5, "todos_g": 6,
    }
    referencia = {}
    for fila in filas:
        valor = fila[mapa[columna]]
        if valor is None:
            valor = fila[6]
        referencia[fila[0]] = valor
    ganancia = peso_actual_g - peso_nacer_g
    p5, p10, p25, p50 = referencia[5], referencia[10], referencia[25], referencia[50]
    if ganancia <= p5:
        percentil = "< P5"
    elif ganancia <= p10:
        percentil = "P5 - P10"
    elif ganancia <= p25:
        percentil = "P10 - P25"
    elif ganancia <= p50:
        percentil = "P25 - P50"
    else:
        percentil = "> P50"
    n_muestra = _n_muestra_ganancia(sexo, intervalo, peso_nacer_g)
    muestra_txt = f" | n={n_muestra}" if n_muestra is not None else ""
    return (
        f"Sexo: {sexo} | Intervalo: {intervalo} días | Rango Peso Nacer: {columna}{muestra_txt}\n"
        f"Ganancia real: {int(ganancia)} g | Percentil: {percentil}\n"
        f"Referencia P5: {int(p5)} g | P10: {int(p10)} g | P25: {int(p25)} g | P50: {int(p50)} g"
    )


NOMBRES_IND_VEL_OMS = {"peso": "Peso (g)", "talla": "Talla/longitud (cm)", "pc": "Perímetro cefálico (cm)"}
NOMBRES_IND_VEL_REV = {v: k for k, v in NOMBRES_IND_VEL_OMS.items()}
INTERVALOS_VEL_OMS = {"peso": [1, 2, 3, 4, 6], "talla": [2, 3, 4, 6], "pc": [2, 3, 4, 6]}
DIAS_POR_MES = 30.4375
INTERVALOS_VEL_DIAS = {
    k: [int(round(m * DIAS_POR_MES)) for m in v]
    for k, v in INTERVALOS_VEL_OMS.items()
}


def _interpolar(x, xs, ys):
    """Interpolación lineal de x en los pares (xs, ys)."""
    if x <= xs[0]:
        return ys[0]
    if x >= xs[-1]:
        return ys[-1]
    for i in range(len(xs) - 1):
        if xs[i] <= x <= xs[i + 1]:
            f = (x - xs[i]) / (xs[i + 1] - xs[i])
            return ys[i] + f * (ys[i + 1] - ys[i])
    return ys[-1]


def _cutoff_lms(M, L, S, z):
    """Corte en la escala LMS (WHO 2009, cap. 6): M·[1+L·S·z]^(1/L)."""
    if abs(L) < 1e-9:
        return M * math.exp(S * z)
    return M * (1.0 + L * S * z) ** (1.0 / L)


def _zscore_lms_vel_oms(inc, L, M, S, delta):
    """z-score OMS de incremento con delta (peso) y corrección más allá de ±3 DE."""
    y = inc + delta
    sd3p, sd2p = _cutoff_lms(M, L, S, 3), _cutoff_lms(M, L, S, 2)
    sd2n, sd3n = _cutoff_lms(M, L, S, -2), _cutoff_lms(M, L, S, -3)
    if y <= 0:
        return -3.0 + (y - sd3n) / (sd2n - sd3n)
    if abs(L) < 1e-9:
        z = math.log(y / M) / S
    else:
        z = ((y / M) ** L - 1.0) / (S * L)
    if z > 3:
        return 3.0 + (y - sd3p) / (sd3p - sd2p)
    if z < -3:
        return -3.0 + (y - sd3n) / (sd2n - sd3n)
    return z


def _parametros_vel_oms(indicador, sexo, intervalo, t1, t2):
    """Devuelve (L, M, S, delta, nota) para velocidad OMS 0-24 m (coincidencia
    exacta o interpolación lineal en el punto medio del intervalo observado)."""
    sexo_db = "Masculino" if sexo in ("M", "Masculino") else "Femenino"
    try:
        conn = conectar()
        filas = conn.execute(
            """SELECT mes_inicio, mes_fin, l, m, s, delta
               FROM referencia_velocidad_oms
               WHERE indicador = ? AND sexo = ? AND intervalo_meses = ?
               ORDER BY mes_inicio""",
            (indicador, sexo_db, int(intervalo)),
        ).fetchall()
        conn.close()
    except Exception:
        return None
    if not filas:
        return None
    for a, b, l, m, s, delta in filas:
        if abs(t1 - a) <= 1e-6 and abs(t2 - b) <= 1e-6:
            return l, m, s, delta, ""
    if (indicador == "peso" and int(intervalo) == 1
            and abs(t1) <= 0.02
            and abs(t2 - 28.0 / DIAS_POR_MES) <= 0.02):
        for a, b, l, m, s, delta in filas:
            if abs(a) <= 1e-6 and abs(b - 1.0) <= 1e-6:
                return l, m, s, delta, ""
    if (indicador == "peso" and int(intervalo) == 1
            and abs(t1 - 28.0 / DIAS_POR_MES) <= 0.02
            and abs(t2 - 2.0) <= 0.02):
        for a, b, l, m, s, delta in filas:
            if abs(a - 1.0) <= 1e-6 and abs(b - 2.0) <= 1e-6:
                return l, m, s, delta, ""
    mp = (t1 + t2) / 2.0
    mids = [(a + b) / 2.0 for a, b, *_ in filas]
    if mp < mids[0] - 0.5 or mp > mids[-1] + 0.5:
        return None
    l = _interpolar(mp, mids, [f[2] for f in filas])
    m = _interpolar(mp, mids, [f[3] for f in filas])
    s = _interpolar(mp, mids, [f[4] for f in filas])
    delta = _interpolar(mp, mids, [f[5] for f in filas])
    return l, m, s, delta, "L, M, S interpolados (intervalo observado no tabulado)."


def _texto_velocidad_oms(indicador, sexo, meses1, meses2, v1, v2,
                         intervalo_referencia=None):
    """Texto de la evaluación de velocidad de crecimiento OMS (WHO 2009),
    0-24 meses. Los tiempos se ingresan en meses y se convierten a días solo
    para presentar la duración observada. None si no se puede calcular."""
    if indicador not in NOMBRES_IND_VEL_OMS:
        return None
    if sexo not in ("M", "F", "Masculino", "Femenino"):
        return None
    try:
        meses1 = float(meses1)
        meses2 = float(meses2)
        v1 = float(v1)
        v2 = float(v2)
    except (TypeError, ValueError):
        return None
    if meses2 <= meses1:
        return None
    t1 = meses1
    t2 = meses2
    dur_m = t2 - t1
    dur_dias = dur_m * DIAS_POR_MES
    if dur_m > 24 or t2 > 24:
        return None
    if intervalo_referencia is None:
        intervalo = min(INTERVALOS_VEL_OMS[indicador], key=lambda it: abs(dur_m - it))
    else:
        try:
            intervalo = int(intervalo_referencia)
        except (TypeError, ValueError):
            return None
        if intervalo not in INTERVALOS_VEL_OMS[indicador]:
            return None
    inc = v2 - v1
    notas = []
    if indicador != "peso" and inc <= 0:
        notas.append("Incremento <= 0 en talla/PC: la OMS lo recodifica como "
                     "'sin crecimiento' (+0.01 cm).")
        inc = 0.01
    par = _parametros_vel_oms(indicador, sexo, intervalo, t1, t2)
    if par is None:
        return None
    L, M, S, delta, nota = par
    if nota:
        notas.append(nota)
    if abs(dur_m - intervalo) > 0.5:
        if intervalo_referencia is None:
            notas.append(
                f"La duración observada {dur_dias:.0f} días ({dur_m:.2f} m) difiere "
                f"del intervalo {intervalo} m: se usó el intervalo más próximo."
            )
        else:
            notas.append(
                f"Se usó el intervalo seleccionado de {intervalo} meses; "
                f"la duración observada fue {dur_m:.2f} meses."
            )
    z = _zscore_lms_vel_oms(inc, L, M, S, delta)
    pct = _percentil_desde_z(z)
    pct_txt = f" | Percentil: P{pct:.1f}" if pct is not None else ""
    if z < -3:
        clas = "MUY BAJA (< -3 DE)"
    elif z < -2:
        clas = "Baja (-3 a -2 DE)"
    elif z <= 2:
        clas = "Normal (-2 a +2 DE)"
    elif z <= 3:
        clas = "Alta (+2 a +3 DE)"
    else:
        clas = "MUY ALTA (> +3 DE)"
    nom = NOMBRES_IND_VEL_OMS[indicador]
    sexo_txt = "Masculino" if sexo in ("M", "Masculino") else "Femenino"
    unidad = "g" if indicador == "peso" else "cm"
    d_txt = f" | δ={delta:g} g" if indicador == "peso" else ""
    cortes_txt = ""
    if indicador == "peso":
        cortes = [
            f"{z:+d} DE: {_cutoff_lms(M, L, S, z) - delta:.0f} g"
            for z in (-3, -2, -1, 0, 1, 2, 3)
        ]
        cortes_txt = "\nCortes de incremento OMS: " + " | ".join(cortes)
    texto = (
        f"Indicador: {nom} | Sexo: {sexo_txt} | Referencia OMS: {intervalo} meses\n"
        f"Duración observada: {dur_dias:.0f} días\n"
        f"Incremento: {inc:.2f} {unidad} (de {v1:g} a {v2:g} en {dur_dias:.0f} días)\n"
        f"Z-score velocidad/edad (OMS 2009): {z:.2f}{pct_txt}\n"
        f"{clas}\n"
        f"L={L:.4f} | M={M:.4f} | S={S:.4f}{d_txt}\n"
        f"(WHO Child Growth Standards: Growth velocity based on weight, length "
        f"and head circumference, 2009)"
    ) + cortes_txt
    if notas:
        texto += "\nNota: " + " ".join(notas)
    return texto


def _texto_velocidad_oms_28_dias(sexo, dias_observados, peso_nacer_g, peso_actual_g):
    """Evalúa la velocidad de peso OMS como incremento equivalente en 28 días."""
    if dias_observados <= 0 or dias_observados > 28:
        return None
    incremento = peso_actual_g - peso_nacer_g
    peso_equivalente = peso_nacer_g + incremento * 28.0 / dias_observados
    return _texto_velocidad_oms(
        "peso", sexo, 0.0, 28.0 / DIAS_POR_MES,
        peso_nacer_g, peso_equivalente,
    )


def _formula_energia(a, b, c):
    """Escribe el lado derecho de la ecuación, p. ej. ((92.8 × peso) − 152) + 196."""
    return f"(({_num(a)} × peso) {_signo(b)} {_num(abs(b))}) + {_num(c)}"


def estimar_energia(grupo, sexo, alimentacion, peso_kg):
    """Devuelve (kcal_dia, detalle) con la energía estimada, sólo en kcal/día.

    Recién nacido (IOM 2005) y lactante (FAO/WHO/UNU 2001): la tabla de la
    fuente está en kilocalorías. En el lactante la ecuación depende del tipo
    de alimentación, no del sexo.
    """
    if peso_kg <= 0:
        raise ValueError("El peso debe ser mayor que cero")
    if sexo not in ENERGIA_SEXOS:
        raise ValueError("Seleccione el sexo")
    kcal, lineas = _detalle_energia(grupo, sexo, alimentacion, peso_kg)
    detalle = "\n".join([f"Grupo: {grupo}"] + lineas)
    return kcal, detalle


def _detalle_energia(grupo, sexo, alimentacion, peso_kg):
    """Calcula una estimación y devuelve (kcal_día, líneas del detalle)."""
    if grupo == ENERGIA_GRUPO_RN:
        a, b, c = ENERGIA_RN[sexo]
        kcal = (a * peso_kg) + b + c
        return kcal, [
            f"Sexo: {sexo}",
            f"Peso: {peso_kg:.3f} kg",
            "",
            f"Fórmula aplicada: {_formula_energia(a, b, c)}",
            f"Estimación: {kcal:.2f} kcal/día",
        ]

    if alimentacion not in ENERGIA_ALIMENTACIONES_OPCIONES:
        raise ValueError("Seleccione el tipo de alimentación")
    a, b, c = ENERGIA_LACTANTE[alimentacion]
    kcal = (a * peso_kg) + b + c
    return kcal, [
        f"Sexo: {sexo}",
        f"Tipo de alimentación: {alimentacion}",
        f"Peso: {peso_kg:.3f} kg",
        "",
        f"Fórmula aplicada: {_formula_energia(a, b, c)}",
        f"Estimación: {kcal:.2f} kcal/día",
    ]


def _num(valor):
    """Formatea un coeficiente de las ecuaciones sin decimales sobrantes."""
    texto = f"{valor:g}"
    return texto


def _signo(valor):
    return "+" if valor >= 0 else "−"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Historia Clínica Nutricional Pediátrica (1-6)")
        self.geometry("900x640")
        self._paciente_id = None
        self._edit_paciente_id = None
        self._backfill_avisos = []
        self._backfill_huerfanos = 0
        self._inicializar_tablas()
        self._crear_notebook()

    def _inicializar_tablas(self):
        sql = (Path(__file__).resolve().parent / "esquema.sql").read_text(encoding="utf-8")
        conn = conectar()
        conn.executescript(sql)
        self._migrar(conn)
        self._cargar_referencias(conn)
        self._rellenar_evaluaciones_pacientes(conn)
        conn.commit()
        conn.close()

    def _rellenar_evaluaciones_pacientes(self, conn):
        """Aplica (backfill) las evaluaciones OMS nuevas a los antecedentes de
        pacientes ya guardados si están vacías, usando sus datos almacenados:
        fecha de nacimiento, sexo, peso al nacer y último peso registrado."""
        filas = conn.execute(
            """SELECT a.id, a.peso_nacer, p.sexo, p.fecha_nacimiento,
                      p.nombre, p.dni_hc,
                      COALESCE(a.valoracion_ganancia_peso,''),
                      COALESCE(a.valoracion_peso_edad,''), COALESCE(a.valoracion_velocidad_crecimiento,''),
                      COALESCE(a.valoracion_zscore_oms,''),
                      (SELECT v.peso_actual_g FROM visita_ganancia_peso v
                        WHERE v.paciente_id = a.paciente_id
                        ORDER BY v.fecha_registro DESC, v.id DESC LIMIT 1),
                      (SELECT v.fecha_visita FROM visita_ganancia_peso v
                        WHERE v.paciente_id = a.paciente_id
                        ORDER BY v.fecha_registro DESC, v.id DESC LIMIT 1)
           FROM antecedentes a JOIN paciente p ON p.id = a.paciente_id
           WHERE a.valoracion_peso_edad IS NOT NULL
              OR a.valoracion_ganancia_peso IS NULL OR a.valoracion_ganancia_peso = ''
              OR a.valoracion_velocidad_crecimiento IS NULL OR a.valoracion_velocidad_crecimiento = ''
              OR a.valoracion_zscore_oms IS NULL OR a.valoracion_zscore_oms = ''""",
        ).fetchall()
        cambios = 0
        self._backfill_avisos = []
        for ant_id, peso_nacer, sexo, fecha_nac, nombre, dni, vgan, vpe, vve, vzs, peso_actual_g, fecha_visita in filas:
            c, motivos = self._aplicar_backfill_fila(
                conn, ant_id, peso_nacer, sexo, fecha_nac,
                vgan, vpe, vve, vzs, peso_actual_g, fecha_visita,
            )
            cambios += c
            if motivos:
                etiqueta = (nombre or "Paciente sin nombre") + (f" (HC {dni})" if dni else "")
                self._backfill_avisos.append((ant_id, etiqueta, motivos))
        self._backfill_huerfanos = conn.execute(
            """SELECT COUNT(*) FROM antecedentes a
               WHERE NOT EXISTS (SELECT 1 FROM paciente p WHERE p.id = a.paciente_id)"""
        ).fetchone()[0]
        return cambios

    def _aplicar_backfill_fila(self, conn, ant_id, peso_nacer, sexo, fecha_nac,
                               vgan, vpe, vve, vzs, peso_actual_g, fecha_visita):
        """Evalúa (backfill) un antecedente individual. Devuelve (cambios, motivos)."""
        motivos = []
        bloqueado = False
        if not sexo:
            motivos.append("el paciente no tiene sexo definido")
            bloqueado = True
        nac = None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                nac = datetime.strptime(fecha_nac, fmt).date()
                break
            except (TypeError, ValueError):
                continue
        if nac is None:
            motivos.append("falta la fecha de nacimiento")
            bloqueado = True
        peso_nacer_g = None
        try:
            pn = float(str(peso_nacer).replace(",", "."))
            peso_nacer_g = int(pn * 1000) if pn < 10 else int(pn)
        except (TypeError, ValueError):
            pass
        if peso_nacer_g is None:
            motivos.append("falta el peso al nacer (la velocidad no es evaluable)")
        al_nacer = False
        if peso_actual_g is None:
            if peso_nacer_g is not None:
                peso_act_g = peso_nacer_g
                al_nacer = True
            else:
                motivos.append("no hay visitas registradas ni peso al nacer (falta el peso actual)")
                bloqueado = True
        else:
            peso_act_g = int(peso_actual_g)
        if bloqueado:
            return 0, motivos
        if al_nacer:
            referencia = nac
        else:
            referencia = None
            if fecha_visita:
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
                    try:
                        referencia = datetime.strptime(fecha_visita, fmt).date()
                        break
                    except (TypeError, ValueError):
                        continue
            if referencia is None:
                referencia = date.today()
        dias = (referencia - nac).days
        nuevo_vgan = vgan
        nuevo_vpe = vpe
        nuevo_vve = vve
        nuevo_vzs = vzs
        if not vgan and peso_nacer_g is not None:
            t = _texto_ganancia(sexo, dias, peso_nacer_g, peso_act_g)
            if t:
                nuevo_vgan = t
            elif not al_nacer:
                motivos.append("ganancia de peso sin calcular (revise edad ≤ 28 días, sexo o referencia)")
        # valoracion_peso_edad se recalcula siempre (no solo si falta): depende de
        # la tabla de referencia peso/edad, que puede corregirse entre versiones.
        t = _texto_peso_edad(sexo, dias, peso_act_g)
        if t:
            nuevo_vpe = t
        elif not vpe:
            motivos.append("peso-para-la-edad sin calcular (revise edad o referencia)")
        if not vzs:
            t = _texto_zscore_oms(sexo, dias, peso_act_g)
            if t:
                nuevo_vzs = t
            elif not PIGROWUP_OK:
                motivos.append("z-score OMS sin calcular (falta instalar pygrowup2)")
            else:
                motivos.append("z-score OMS sin calcular (revise edad o instalación de pygrowup2)")
        if (nuevo_vgan, nuevo_vpe, nuevo_vve, nuevo_vzs) != (vgan, vpe, vve, vzs):
            conn.execute(
                "UPDATE antecedentes SET valoracion_ganancia_peso=?, valoracion_peso_edad=?, valoracion_velocidad_crecimiento=?, valoracion_zscore_oms=? WHERE id=?",
                (nuevo_vgan, nuevo_vpe, nuevo_vve, nuevo_vzs, ant_id),
            )
            return 1, motivos
        return 0, motivos

    def _evaluar_paciente(self, avisar=True):
        """Evalúa los antecedentes del paciente cargado en el momento y, si algo
        no puede calcularse, avisa con los motivos."""
        if not getattr(self, "_paciente_id", None):
            return 0
        conn = conectar()
        filas = conn.execute(
            """SELECT a.id, a.peso_nacer, p.sexo, p.fecha_nacimiento,
                      COALESCE(a.valoracion_ganancia_peso,''),
                      COALESCE(a.valoracion_peso_edad,''), COALESCE(a.valoracion_velocidad_crecimiento,''),
                      COALESCE(a.valoracion_zscore_oms,''),
                      (SELECT v.peso_actual_g FROM visita_ganancia_peso v
                        WHERE v.paciente_id = a.paciente_id
                        ORDER BY v.fecha_registro DESC, v.id DESC LIMIT 1),
                      (SELECT v.fecha_visita FROM visita_ganancia_peso v
                        WHERE v.paciente_id = a.paciente_id
                        ORDER BY v.fecha_registro DESC, v.id DESC LIMIT 1)
               FROM antecedentes a JOIN paciente p ON p.id = a.paciente_id
               WHERE a.paciente_id = ?""",
            (self._paciente_id,),
        ).fetchall()
        cambios = 0
        motivos = []
        for ant_id, peso_nacer, sexo, fecha_nac, vgan, vpe, vve, vzs, peso_actual_g, fecha_visita in filas:
            c, m = self._aplicar_backfill_fila(
                conn, ant_id, peso_nacer, sexo, fecha_nac,
                vgan, vpe, vve, vzs, peso_actual_g, fecha_visita,
            )
            cambios += c
            motivos.extend(m)
        conn.commit()
        conn.close()
        if avisar and motivos:
            nombre = self.campos["nombre"].get().strip()
            mensaje = (f"{nombre}: la evaluación automática quedó incompleta:\n\n"
                       + "\n".join(f"• {m}" for m in motivos))
            messagebox.showinfo("Evaluación del paciente", mensaje)
        return cambios

    def _migrar(self, conn):
        columnas = [fila[1] for fila in conn.execute("PRAGMA table_info(antecedentes)")]
        if "clasificacion_peso_nacer" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN clasificacion_peso_nacer TEXT")
        if "valoracion_ganancia_peso" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN valoracion_ganancia_peso TEXT")
        if "valoracion_peso_edad" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN valoracion_peso_edad TEXT")
        if "valoracion_velocidad_crecimiento" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN valoracion_velocidad_crecimiento TEXT")
        if "valoracion_zscore_oms" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN valoracion_zscore_oms TEXT")
        if "valoracion_velocidad_oms" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN valoracion_velocidad_oms TEXT")
        columnas_visita = [
            fila[1] for fila in conn.execute("PRAGMA table_info(visita_ganancia_peso)")
        ]
        if "velocidad_g_dia" not in columnas_visita:
            conn.execute("ALTER TABLE visita_ganancia_peso ADD COLUMN velocidad_g_dia REAL")
        if "velocidad_percentil" not in columnas_visita:
            conn.execute("ALTER TABLE visita_ganancia_peso ADD COLUMN velocidad_percentil TEXT")
        columnas_cans = [
            fila[1] for fila in conn.execute("PRAGMA table_info(evaluacion_cans)")
        ]
        for nombre, tipo in (
            ("indice_ponderal", "REAL"),
            ("peso_nacer_g", "REAL"),
            ("longitud_nacer_cm", "REAL"),
        ):
            if nombre not in columnas_cans:
                conn.execute(f"ALTER TABLE evaluacion_cans ADD COLUMN {nombre} {tipo}")
        columnas_bio = [
            fila[1] for fila in conn.execute("PRAGMA table_info(evaluacion_bioquimica)")
        ]
        if "comparacion" not in columnas_bio:
            conn.execute("ALTER TABLE evaluacion_bioquimica ADD COLUMN comparacion TEXT")
        if "unidad" not in columnas_bio:
            conn.execute("ALTER TABLE evaluacion_bioquimica ADD COLUMN unidad TEXT")
        if "referencia_edad" not in columnas_bio:
            conn.execute("ALTER TABLE evaluacion_bioquimica ADD COLUMN referencia_edad TEXT")
        columnas_lact = [fila[1] for fila in conn.execute("PRAGMA table_info(lactmed)")]
        if "aliases" not in columnas_lact:
            conn.execute("ALTER TABLE lactmed ADD COLUMN aliases TEXT")
        col_pac = [fila[1] for fila in conn.execute("PRAGMA table_info(paciente)")]
        if "nombre" not in col_pac:
            conn.execute("ALTER TABLE paciente ADD COLUMN nombre TEXT")
        if "direccion" not in col_pac:
            conn.execute("ALTER TABLE paciente ADD COLUMN direccion TEXT")
        if "telefono" not in col_pac:
            conn.execute("ALTER TABLE paciente ADD COLUMN telefono TEXT")
        if "edad" not in col_pac:
            conn.execute("ALTER TABLE paciente ADD COLUMN edad TEXT")
        if "fecha_actual" not in col_pac:
            conn.execute("ALTER TABLE paciente ADD COLUMN fecha_actual TEXT")
        for col_vieja in ("fecha_ingreso", "fecha_alta", "edad_anios", "edad_meses"):
            if col_vieja in col_pac:
                try:
                    conn.execute(f"ALTER TABLE paciente DROP COLUMN {col_vieja}")
                except sqlite3.OperationalError:
                    pass
        col_energia = [
            fila[1] for fila in conn.execute("PRAGMA table_info(estimacion_energetica)")
        ]
        if "mj_dia" in col_energia:
            try:
                conn.execute("ALTER TABLE estimacion_energetica DROP COLUMN mj_dia")
            except sqlite3.OperationalError:
                pass
        conn.commit()

    def _cargar_referencias(self, conn):
        data_dir = Path(__file__).resolve().parent / "data"
        nombres = list(dict.fromkeys([
            "ninos.csv", "ninas.csv", "clasificaciones.csv", "ECRN_ninos.csv", "ECRN_ninas.csv",
            "ECRN_velocidad_ninos.csv", "ECRN_velocidad_ninas.csv",
            "tabla_peso_para_la_edad.csv", "tabla_velocidad_crecimiento.csv",
            "tabla_velocidad_oms.csv", "tabla_desarrollo_motor_oms.csv",
            "QS.csv",
        ]))
        for nombre in nombres:
            if conn.execute("SELECT 1 FROM archivo_csv WHERE nombre = ?", (nombre,)).fetchone():
                continue
            origen = data_dir / nombre
            if origen.exists():
                contenido = origen.read_text(encoding="utf-8-sig")
            else:
                contenido = CSV_DEFAULT[nombre]
            conn.execute(
                "INSERT INTO archivo_csv (nombre, contenido) VALUES (?,?)",
                (nombre, contenido),
            )
        conn.commit()
        if conn.execute("SELECT COUNT(*) FROM referencia_peso_nacer").fetchone()[0] == 0:
            for nombre, sexo in (("ninos.csv", "Masculino"), ("ninas.csv", "Femenino")):
                contenido = conn.execute(
                    "SELECT contenido FROM archivo_csv WHERE nombre = ?", (nombre,)
                ).fetchone()[0]
                for fila in csv.DictReader(io.StringIO(contenido)):
                    conn.execute(
                        "INSERT INTO referencia_peso_nacer (sexo, edad_gestacional_semanas, percentil_10_g, percentil_50_g, percentil_90_g) VALUES (?,?,?,?,?)",
                        (
                            sexo,
                            int(fila["Edad_gestacional_semanas"]),
                            float(fila["Percentil_10_peso_g"]),
                            float(fila["Percentil_50_peso_g"]),
                            float(fila["Percentil_90_peso_g"]),
                        ),
                    )
        if conn.execute("SELECT COUNT(*) FROM clasificacion_eg").fetchone()[0] == 0:
            contenido = conn.execute(
                "SELECT contenido FROM archivo_csv WHERE nombre = 'clasificaciones.csv'"
            ).fetchone()[0]
            for fila in csv.DictReader(io.StringIO(contenido)):
                conn.execute(
                    "INSERT INTO clasificacion_eg (percentil, interpretacion) VALUES (?,?)",
                    (fila["Percentil"], fila["Interpretación"]),
                )
        columnas_ecrn = [
            "peso_nacer_2000_2500_g", "peso_nacer_2500_3000_g",
            "peso_nacer_3000_3500_g", "peso_nacer_3500_4000_g",
            "peso_nacer_4000_mas_g", "todos_g",
        ]
        def _cargar_ecrn(contenido, sexo):
            for fila in csv.DictReader(io.StringIO(contenido)):
                intervalo = fila["intervalo_dias"].strip()
                percentil = fila["percentil"].strip().lower()
                if percentil == "n":
                    valores_n = [int(fila[col]) for col in columnas_ecrn]
                    conn.execute(
                        """INSERT OR REPLACE INTO referencia_ganancia_peso_muestra
                           (sexo, intervalo_dias, n_peso_nacer_2000_2500,
                            n_peso_nacer_2500_3000, n_peso_nacer_3000_3500,
                            n_peso_nacer_3500_4000, n_peso_nacer_4000_mas, n_todos)
                           VALUES (?,?,?,?,?,?,?,?)""",
                        (sexo, intervalo, *valores_n),
                    )
                    continue
                valores = []
                for col in columnas_ecrn:
                    valor = fila[col].strip()
                    valores.append(None if valor in ("", "*") else float(valor))
                conn.execute(
                    """INSERT OR REPLACE INTO referencia_ganancia_peso
                       (sexo, intervalo_dias, percentil,
                        peso_nacer_2000_2500_g, peso_nacer_2500_3000_g,
                        peso_nacer_3000_3500_g, peso_nacer_3500_4000_g,
                        peso_nacer_4000_mas_g, todos_g)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (sexo, intervalo, int(percentil), *valores),
                )

        if conn.execute("SELECT COUNT(*) FROM referencia_ganancia_peso").fetchone()[0] == 0:
            for nombre, sexo in (("ECRN_ninos.csv", "Masculino"), ("ECRN_ninas.csv", "Femenino")):
                contenido = conn.execute(
                    "SELECT contenido FROM archivo_csv WHERE nombre = ?", (nombre,)
                ).fetchone()[0]
                _cargar_ecrn(contenido, sexo)

        archivo_ninos = data_dir / "ECRN_ninos.csv"
        contenido_ninos = (
            archivo_ninos.read_text(encoding="utf-8-sig")
            if archivo_ninos.exists() else CSV_DEFAULT["ECRN_ninos.csv"]
        )
        conn.execute(
            "UPDATE archivo_csv SET contenido = ? WHERE nombre = 'ECRN_ninos.csv'",
            (contenido_ninos,),
        )
        conn.execute("DELETE FROM referencia_ganancia_peso WHERE sexo = 'Masculino'")
        conn.execute("DELETE FROM referencia_ganancia_peso_muestra WHERE sexo = 'Masculino'")
        _cargar_ecrn(contenido_ninos, "Masculino")
        archivo_ninas = data_dir / "ECRN_ninas.csv"
        contenido_ninas = (
            archivo_ninas.read_text(encoding="utf-8-sig")
            if archivo_ninas.exists() else CSV_DEFAULT["ECRN_ninas.csv"]
        )
        conn.execute(
            "UPDATE archivo_csv SET contenido = ? WHERE nombre = 'ECRN_ninas.csv'",
            (contenido_ninas,),
        )
        conn.execute("DELETE FROM referencia_ganancia_peso WHERE sexo = 'Femenino'")
        conn.execute("DELETE FROM referencia_ganancia_peso_muestra WHERE sexo = 'Femenino'")
        _cargar_ecrn(contenido_ninas, "Femenino")
        def _cargar_velocidad_ecrn(nombre, sexo):
            archivo = data_dir / nombre
            contenido = (
                archivo.read_text(encoding="utf-8-sig")
                if archivo.exists() else CSV_DEFAULT[nombre]
            )
            conn.execute(
                "UPDATE archivo_csv SET contenido = ? WHERE nombre = ?",
                (contenido, nombre),
            )
            conn.execute(
                "DELETE FROM referencia_velocidad_peso_ecrn WHERE sexo = ?", (sexo,)
            )
            for fila in csv.DictReader(io.StringIO(contenido)):
                intervalo = fila["intervalo_dias"].strip()
                percentil = fila["percentil"].strip()
                if percentil.lower() == "n":
                    valores_n = [int(fila[col]) for col in columnas_ecrn]
                    conn.execute(
                        """INSERT OR REPLACE INTO referencia_ganancia_peso_muestra
                           (sexo, intervalo_dias, n_peso_nacer_2000_2500,
                            n_peso_nacer_2500_3000, n_peso_nacer_3000_3500,
                            n_peso_nacer_3500_4000, n_peso_nacer_4000_mas, n_todos)
                           VALUES (?,?,?,?,?,?,?,?)""",
                        (sexo, intervalo, *valores_n),
                    )
                    continue
                valores = []
                for col in columnas_ecrn:
                    valor = fila[col].strip()
                    valores.append(None if valor in ("", "*") else float(valor))
                conn.execute(
                    """INSERT OR REPLACE INTO referencia_velocidad_peso_ecrn
                       (sexo, intervalo_dias, percentil,
                        peso_nacer_2000_2500_g, peso_nacer_2500_3000_g,
                        peso_nacer_3000_3500_g, peso_nacer_3500_4000_g,
                        peso_nacer_4000_mas_g, todos_g)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (sexo, intervalo, percentil, *valores),
                )

        for nombre, sexo in (
            ("ECRN_velocidad_ninos.csv", "Masculino"),
            ("ECRN_velocidad_ninas.csv", "Femenino"),
        ):
            _cargar_velocidad_ecrn(nombre, sexo)
        contenido_peso_edad = (
            (data_dir / "tabla_peso_para_la_edad.csv").read_text(encoding="utf-8-sig")
            if (data_dir / "tabla_peso_para_la_edad.csv").exists()
            else CSV_DEFAULT["tabla_peso_para_la_edad.csv"]
        )
        conn.execute(
            "UPDATE archivo_csv SET contenido = ? WHERE nombre = 'tabla_peso_para_la_edad.csv'",
            (contenido_peso_edad,),
        )
        conn.execute("DELETE FROM referencia_peso_edad")
        for fila in csv.DictReader(io.StringIO(contenido_peso_edad)):
            conn.execute(
                """INSERT OR REPLACE INTO referencia_peso_edad
                   (sexo, edad_dias, edad_etiqueta, p3, p15, p50, p85, p97)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    "Masculino" if fila["sexo"].strip() == "M" else "Femenino",
                    int(fila["edad_dias"]),
                    fila["edad_etiqueta"].strip(),
                    float(fila["P3"]), float(fila["P15"]), float(fila["P50"]),
                    float(fila["P85"]), float(fila["P97"]),
                ),
            )
        if conn.execute("SELECT COUNT(*) FROM referencia_velocidad_crecimiento").fetchone()[0] == 0:

            def _nulo(v):
                v = v.strip()
                return None if v in ("", "*") else float(v)

            contenido = conn.execute(
                "SELECT contenido FROM archivo_csv WHERE nombre = 'tabla_velocidad_crecimiento.csv'"
            ).fetchone()[0]
            for fila in csv.DictReader(io.StringIO(contenido)):
                conn.execute(
                    """INSERT INTO referencia_velocidad_crecimiento
                       (sexo, ventana, edad_ini_dias, edad_fin_dias,
                        peso_nacer_min_kg, peso_nacer_max_kg, p3, p15, p50, p85, p97, nota)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        "Masculino" if fila["sexo"].strip() == "M" else "Femenino",
                        fila["ventana"].strip(),
                        int(fila["edad_ini_dias"]), int(fila["edad_fin_dias"]),
                        float(fila["peso_nacer_min_kg"]), float(fila["peso_nacer_max_kg"]),
                        _nulo(fila["P3"]), _nulo(fila["P15"]), _nulo(fila["P50"]),
                        _nulo(fila["P85"]), _nulo(fila["P97"]),
                        fila.get("nota", "").strip(),
                    ),
                )
        if conn.execute("SELECT COUNT(*) FROM referencia_velocidad_oms").fetchone()[0] == 0:
            contenido = conn.execute(
                "SELECT contenido FROM archivo_csv WHERE nombre = 'tabla_velocidad_oms.csv'"
            ).fetchone()[0]
            for fila in csv.DictReader(io.StringIO(contenido)):
                conn.execute(
                    """INSERT INTO referencia_velocidad_oms
                       (indicador, sexo, intervalo_meses, mes_inicio, mes_fin, l, m, s, delta)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (
                        fila["indicador"].strip(),
                        "Masculino" if fila["sexo"].strip() == "M" else "Femenino",
                        int(fila["intervalo_meses"]),
                        float(fila["mes_inicio"]), float(fila["mes_fin"]),
                        float(fila["L"]), float(fila["M"]), float(fila["S"]),
                        float(fila["delta"]),
                    ),
                )
        if conn.execute("SELECT COUNT(*) FROM referencia_desarrollo_motor").fetchone()[0] == 0:
            contenido = conn.execute(
                "SELECT contenido FROM archivo_csv WHERE nombre = 'tabla_desarrollo_motor_oms.csv'"
            ).fetchone()[0]
            for fila in csv.DictReader(io.StringIO(contenido)):
                conn.execute(
                    "INSERT OR REPLACE INTO referencia_desarrollo_motor (hito, percentil, edad_dias) VALUES (?,?,?)",
                    (fila["hito"].strip(), int(fila["percentil"]), int(fila["edad_dias"])),
                )
        contenido_qs = (
            (data_dir / "QS.csv").read_text(encoding="utf-8-sig")
            if (data_dir / "QS.csv").exists()
            else CSV_DEFAULT["QS.csv"]
        )
        conn.execute(
            "UPDATE archivo_csv SET contenido = ? WHERE nombre = 'QS.csv'",
            (contenido_qs,),
        )
        conn.execute("DELETE FROM referencia_bioquimica")
        for fila in csv.DictReader(io.StringIO(contenido_qs)):
            conn.execute(
                """INSERT OR REPLACE INTO referencia_bioquimica
                   (grupo, prueba, unidad, ref_nacimiento, ref_1_semana, ref_1_mes, ref_1_5_mes, notas)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    fila["Tabla"].strip(), fila["Prueba"].strip(), fila["Unidad"].strip(),
                    fila.get("Nacimiento", "").strip(), fila.get("1 semana", "").strip(),
                    fila.get("1 mes", "").strip(), fila.get("1 1/2 mes", "").strip(),
                    fila.get("Notas", "").strip(),
                ),
            )
        conn.commit()

    def _crear_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self._tab_paciente()
        self._tab_antecedentes()
        self._tab_cans()
        self._tab_farmaco()
        self._tab_bioquimica()
        self._tab_ecuaciones()
        self._tab_recuento()

        barra = ttk.Frame(self)
        barra.pack(fill="x", padx=8, pady=8)
        ttk.Button(barra, text="Cerrar", command=self.destroy).pack(side="right")
        self.status = ttk.Label(self, text="", foreground="green")
        self.status.pack(fill="x", padx=8, pady=(0, 6))

    def _crear_grid(self, parent, columnas, alturas=None):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, padx=6, pady=6)
        tree = ttk.Treeview(frame, columns=columnas, show="headings", height=alturas or 10)
        for col in columnas:
            tree.heading(col, text=col)
            tree.column(col, width=160)
        tsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tsb.set)
        tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        return tree

    # ---------- Pestaña Paciente ----------
    def _tab_paciente(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Paciente")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Paciente:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_paciente).pack(side="left", padx=3)
        ttk.Button(barra, text="Nuevo Paciente", command=self._nueva_consulta_paciente).pack(side="left", padx=3)
        self.pac_status = ttk.Label(barra, text="", foreground="blue")
        self.pac_status.pack(side="left", padx=10)

        marco = ttk.LabelFrame(tab, text="Datos Generales")
        marco.pack(fill="x", padx=8, pady=8)
        campos = [
            ("Nombre", "nombre"), ("Servicio", "servicio"), ("DNI/HC", "dni"),
            ("Cuenta", "cuenta"), ("Cama", "cama"),
        ]
        self.campos = {}
        self.entradas = {}
        for i, (etiqueta, clave) in enumerate(campos):
            fila, col = divmod(i, 3)
            ttk.Label(marco, text=etiqueta).grid(row=fila, column=col * 2, sticky="w", padx=5, pady=4)
            var = tk.StringVar()
            entry = ttk.Entry(marco, textvariable=var, width=16)
            entry.grid(row=fila, column=col * 2 + 1, padx=5, pady=4)
            self.campos[clave] = var
            self.entradas[clave] = entry

        ttk.Label(marco, text="Sexo").grid(row=1, column=4, sticky="w", padx=5)
        self.sexo = tk.StringVar(value="Femenino")
        ttk.Combobox(marco, textvariable=self.sexo, values=["Femenino", "Masculino"], state="readonly", width=14).grid(
            row=1, column=5, padx=5, pady=4)

        fechas = [
            ("F. Nacimiento", "fecha_nacimiento"), ("Fecha actual", "fecha_actual"),
        ]
        for i, (etiqueta, clave) in enumerate(fechas):
            col = i * 2
            ttk.Label(marco, text=etiqueta).grid(row=2, column=col, sticky="w", padx=5)
            var = tk.StringVar()
            ttk.Entry(marco, textvariable=var, width=14).grid(row=2, column=col + 1, padx=5)
            self.campos[clave] = var

        ttk.Label(marco, text="Edad").grid(row=3, column=0, sticky="w", padx=5, pady=4)
        self.campos["edad"] = tk.StringVar()
        edad_entry = ttk.Entry(marco, textvariable=self.campos["edad"], width=22, state="readonly")
        edad_entry.grid(row=3, column=1, columnspan=2, padx=5, sticky="w")
        self.campos["edad_entry"] = edad_entry

        self.campos["fecha_nacimiento"].trace_add("write", lambda *_: self._calcular_edad())
        self.campos["fecha_actual"].trace_add("write", lambda *_: self._calcular_edad())
        self.campos["fecha_actual"].set(date.today().isoformat())

        for campo in ("nombre", "dni"):
            self.entradas[campo].bind("<KeyRelease>", lambda e, c=campo: self._mostrar_sugerencias(c))
            self.entradas[campo].bind("<FocusOut>", lambda e: self.after(150, self._ocultar_sugerencias))
        self.entradas["nombre"].bind("<KeyRelease>", lambda e: self._programar_dni_nuevo(), add="+")
        self.entradas["nombre"].bind("<FocusOut>", lambda e: self._asignar_dni_si_nuevo(), add="+")

        ttk.Label(marco, text="Grado de Instrucción Padre/Cuidador").grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        self.campos["grado_padre"] = tk.StringVar()
        ttk.Entry(marco, textvariable=self.campos["grado_padre"], width=25).grid(row=4, column=2, columnspan=2, padx=5)

        ttk.Label(marco, text="Grado de Instrucción Madre/Cuidador").grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        self.campos["grado_madre"] = tk.StringVar()
        ttk.Entry(marco, textvariable=self.campos["grado_madre"], width=25).grid(row=5, column=2, columnspan=2, padx=5)

        ttk.Label(marco, text="Dirección").grid(row=6, column=0, sticky="w", padx=5, pady=4)
        self.campos["direccion"] = tk.StringVar()
        ttk.Entry(marco, textvariable=self.campos["direccion"], width=25).grid(row=6, column=1, columnspan=2, padx=5, sticky="w")
        ttk.Label(marco, text="Teléfono").grid(row=6, column=3, sticky="w", padx=5)
        self.campos["telefono"] = tk.StringVar()
        ttk.Entry(marco, textvariable=self.campos["telefono"], width=14).grid(row=6, column=4, padx=5, sticky="w")

        marco_dx = ttk.LabelFrame(tab, text="B) Dx. Médico")
        marco_dx.pack(fill="both", expand=True, padx=8, pady=8)
        self.campos["dx_medico"] = tk.StringVar()
        ttk.Entry(marco_dx, textvariable=self.campos["dx_medico"]).pack(fill="x", padx=8, pady=8)

    # ---------- Pestaña 1 Antropometría ----------
    def _tab_antecedentes(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="1. Antropometría")
        self._tab_ant = tab

        contenedor = ttk.Frame(tab)
        contenedor.pack(fill="both", expand=True)
        canvas = tk.Canvas(contenedor, highlightthickness=0)
        scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        interior = ttk.Frame(canvas)
        ventana = canvas.create_window((0, 0), window=interior, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._canvas_ant = canvas

        def _actualizar_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _ancho_scroll(_event=None):
            canvas.itemconfigure(ventana, width=_event.width)

        def _rueda(_event):
            canvas.yview_scroll(int(-_event.delta / 120), "units")

        def _cambio_tab(_event=None):
            if self.notebook.select() == tab:
                canvas.bind_all("<MouseWheel>", _rueda)
            else:
                canvas.unbind_all("<MouseWheel>")

        interior.bind("<Configure>", _actualizar_scroll)
        canvas.bind("<Configure>", _ancho_scroll)
        self.notebook.bind("<<NotebookTabChanged>>", _cambio_tab)
        tab.bind("<Map>", _cambio_tab)
        tab.bind("<Unmap>", _cambio_tab)
        tab.bind("<Destroy>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        barra = ttk.Frame(interior)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Antecedentes:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_antecedentes).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_antecedentes).pack(side="left", padx=3)

        marco = ttk.LabelFrame(interior, text="A) Antecedentes Prenatales")
        marco.pack(fill="x", padx=8, pady=8)
        campos = [
            ("Edad Gestacional", "edad_gestacional"), ("Parto", "parto"),
            ("Peso al Nacer (kg)", "peso_nacer"), ("Perímetro Cefálico al Nacer (cm)", "perimetro_cefalico_nacer"),
            ("Longitud al Nacer (cm)", "longitud_nacer"),
        ]
        self.ant = {}
        for i, (etiqueta, clave) in enumerate(campos):
            fila, col = divmod(i, 2)
            ttk.Label(marco, text=etiqueta).grid(row=fila, column=col * 2, sticky="w", padx=5, pady=4)
            var = tk.StringVar()
            ttk.Entry(marco, textvariable=var, width=18).grid(row=fila, column=col * 2 + 1, padx=5, pady=4)
            self.ant[clave] = var

        marco_familia = ttk.LabelFrame(interior, text="Antecedentes Familiares")
        marco_familia.pack(fill="x", padx=8, pady=8)
        self.ant["familiares"] = tk.Text(marco_familia, height=6)
        self.ant["familiares"].pack(fill="both", expand=True, padx=6, pady=6)

        marco_clas = ttk.LabelFrame(interior, text="Evaluación neonatal: edad gestacional, ganancia y velocidad")
        marco_clas.pack(fill="x", padx=8, pady=8)
        ttk.Label(marco_clas, text="Sexo al nacer:").grid(row=0, column=0, sticky="w", padx=5, pady=4)
        self.sexo_nacer = tk.StringVar(value=self.sexo.get())
        ttk.Combobox(
            marco_clas, textvariable=self.sexo_nacer,
            values=["Femenino", "Masculino"], state="readonly", width=14,
        ).grid(row=0, column=1, sticky="w", padx=5, pady=4)
        ttk.Label(marco_clas, text="Peso al nacer (g):").grid(
            row=1, column=0, sticky="w", padx=5, pady=4)
        self.clas_peso_g = tk.StringVar()
        ttk.Entry(marco_clas, textvariable=self.clas_peso_g, width=10, state="readonly").grid(
            row=1, column=1, sticky="w", padx=5, pady=4)
        self.ant["peso_nacer"].trace_add("write", self._actualizar_clas_peso_g)
        ttk.Button(marco_clas, text="Clasificar", command=self._clasificar).grid(row=2, column=0, sticky="w", padx=5, pady=4)
        self.clas_resultado = tk.Label(marco_clas, text="", justify="left", foreground="blue", wraplength=700)
        self.clas_resultado.grid(row=3, column=0, columnspan=4, sticky="w", padx=5, pady=4)
        self.ant["edad_gestacional"].trace_add("write", self._clasificar_peso_nacer)
        self.sexo_nacer.trace_add("write", self._clasificar_peso_nacer)

        ttk.Separator(marco_clas, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", padx=5, pady=6)
        self.peso_actual_label = ttk.Label(marco_clas, text="Peso actual (g o kg):")
        self.peso_actual_label.grid(row=5, column=0, sticky="w", padx=5, pady=4)
        self.clas_peso_actual = tk.StringVar()
        self.peso_actual_entry = ttk.Entry(marco_clas, textvariable=self.clas_peso_actual, width=14)
        self.peso_actual_entry.grid(row=5, column=1, sticky="w", padx=5, pady=4)
        self.valorar_btn = ttk.Button(marco_clas, text="Evaluar neonato", command=self._valorar_ganancia)
        self.valorar_btn.grid(row=5, column=2, sticky="w", padx=5, pady=4)
        self.ganancia_resultado = tk.Label(marco_clas, text="", justify="left", foreground="green", wraplength=700)
        self.ganancia_resultado.grid(row=6, column=0, columnspan=4, sticky="w", padx=5, pady=4)
        self.velocidad_neonatal_resultado = tk.Label(
            marco_clas, text="", justify="left", foreground="#0B5C7A", wraplength=700)
        self.velocidad_neonatal_resultado.grid(row=7, column=0, columnspan=4, sticky="w", padx=5, pady=4)

        self.guardar_visita_btn = ttk.Button(marco_clas, text="Guardar visita", command=self._guardar_visita_ganancia)
        self.guardar_visita_btn.grid(row=8, column=0, sticky="w", padx=5, pady=4)
        self.refrescar_historial_btn = ttk.Button(marco_clas, text="Refrescar historial", command=self._refrescar_historial)
        self.refrescar_historial_btn.grid(row=8, column=1, sticky="w", padx=5, pady=4)
        self.cargar_visita_btn = ttk.Button(marco_clas, text="Cargar visita", command=self._cargar_visita)
        self.cargar_visita_btn.grid(row=8, column=2, sticky="w", padx=5, pady=4)
        self.borrar_visita_btn = ttk.Button(marco_clas, text="Borrar visita", command=self._borrar_visita)
        self.borrar_visita_btn.grid(row=8, column=3, sticky="w", padx=5, pady=4)

        self.marco_vis = ttk.LabelFrame(marco_clas, text="Historial de ganancia y velocidad de crecimiento")
        self.marco_vis.grid(row=9, column=0, columnspan=4, sticky="ew", padx=5, pady=4)
        self.visita_tree = ttk.Treeview(
            self.marco_vis,
            columns=["Fecha", "Días", "Peso actual (g)", "Ganancia (g)",
                     "Percentil ganancia", "Velocidad (g/día)", "Percentil velocidad"],
            show="headings", height=5,
        )
        anchuras = {
            "Fecha": 100, "Días": 55, "Peso actual (g)": 100, "Ganancia (g)": 95,
            "Percentil ganancia": 115, "Velocidad (g/día)": 115,
            "Percentil velocidad": 120,
        }
        for col, ancho in anchuras.items():
            self.visita_tree.heading(col, text=col)
            self.visita_tree.column(col, width=ancho, anchor="center")
        self.visita_tree.pack(side="left", fill="x", expand=True)
        tsb = ttk.Scrollbar(self.marco_vis, orient="vertical", command=self.visita_tree.yview)
        tsb.pack(side="right", fill="y")
        self.visita_tree.configure(yscrollcommand=tsb.set)
        self.clas_peso_ok = False

        self._actualizar_visibilidad_ganancia()

        marco_eval = ttk.LabelFrame(interior, text="Evaluación del crecimiento (OMS 2006 - WHO Anthro)")
        marco_eval.pack(fill="x", padx=8, pady=8)
        ttk.Label(marco_eval, text="Peso actual (g o kg):").grid(
            row=0, column=0, sticky="e", padx=5, pady=4)
        self.eval_peso_actual = tk.StringVar()
        ttk.Entry(marco_eval, textvariable=self.eval_peso_actual, width=14).grid(
            row=0, column=1, sticky="w", padx=5, pady=4)
        self.eval_peso_actual.trace_add("write", self._auto_evaluar_z)
        ttk.Label(marco_eval, text="Talla/longitud (cm):").grid(
            row=1, column=0, sticky="e", padx=5, pady=4)
        self.eval_talla = tk.StringVar()
        ttk.Entry(marco_eval, textvariable=self.eval_talla, width=10).grid(
            row=1, column=1, sticky="w", padx=5, pady=4)
        self.eval_talla.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval, text="Perím. cefálico (cm):").grid(
            row=0, column=2, sticky="e", padx=5, pady=4)
        self.eval_pc = tk.StringVar()
        ttk.Entry(marco_eval, textvariable=self.eval_pc, width=10).grid(
            row=0, column=3, sticky="w", padx=5, pady=4)
        self.eval_pc.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval, text="MUAC (cm):").grid(
            row=1, column=2, sticky="e", padx=5, pady=4)
        self.eval_muac = tk.StringVar()
        ttk.Entry(marco_eval, textvariable=self.eval_muac, width=10).grid(
            row=1, column=3, sticky="w", padx=5, pady=4)
        self.eval_muac.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval, text="Tríceps (mm):").grid(
            row=2, column=0, sticky="e", padx=5, pady=4)
        self.eval_triceps = tk.StringVar()
        ttk.Entry(marco_eval, textvariable=self.eval_triceps, width=10).grid(
            row=2, column=1, sticky="w", padx=5, pady=4)
        self.eval_triceps.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval, text="Subescapular (mm):").grid(
            row=2, column=2, sticky="e", padx=5, pady=4)
        self.eval_subescapular = tk.StringVar()
        ttk.Entry(marco_eval, textvariable=self.eval_subescapular, width=10).grid(
            row=2, column=3, sticky="w", padx=5, pady=4)
        self.eval_subescapular.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval, text="Medición:").grid(
            row=3, column=0, sticky="e", padx=5, pady=4)
        self.eval_postura = tk.StringVar(value="Acostado")
        ttk.Combobox(
            marco_eval, textvariable=self.eval_postura,
            values=["Acostado", "Parado"], state="readonly", width=12,
        ).grid(row=3, column=1, sticky="w", padx=5, pady=4)
        self.eval_postura.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval, text="Edema:").grid(
            row=3, column=2, sticky="e", padx=5, pady=4)
        self.eval_edema = tk.StringVar(value="No")
        ttk.Combobox(
            marco_eval, textvariable=self.eval_edema,
            values=["No", "Sí"], state="readonly", width=8,
        ).grid(row=3, column=3, sticky="w", padx=5, pady=4)
        self.eval_edema.trace_add("write", self._auto_evaluar_anthro)
        ttk.Label(marco_eval,
                  text="1. WHO Anthro: peso/edad, talla/edad, peso/talla, IMC/edad, PC/edad, MUAC/edad, "
                       "tríceps/edad y subescapular/edad [automático al llenar los campos; pliegues en mm, "
                       "de 3 meses a 5 años]").grid(
            row=4, column=0, columnspan=4, sticky="w", padx=5, pady=2)
        self.anthro_resultado = tk.Label(
            marco_eval, text="", justify="left", foreground="#1E6E5C", wraplength=700)
        self.anthro_resultado.grid(row=5, column=0, columnspan=4, sticky="w", padx=5, pady=2)
        self._valoracion_zscore = ""

        marco_velocidad = ttk.LabelFrame(interior, text="Velocidad de crecimiento del lactante (OMS 2009)")
        marco_velocidad.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(
            marco_velocidad,
            text="Lactante: velocidad OMS 2009 (0-24 meses)",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=5, pady=4)

        ttk.Label(marco_velocidad, text="Indicador:").grid(row=1, column=0, sticky="e", padx=5, pady=4)
        self.vel_indicador = tk.StringVar(value=NOMBRES_IND_VEL_OMS["peso"])
        ttk.Combobox(
            marco_velocidad,
            textvariable=self.vel_indicador,
            values=[NOMBRES_IND_VEL_OMS["peso"], NOMBRES_IND_VEL_OMS["talla"], NOMBRES_IND_VEL_OMS["pc"]],
            state="readonly",
            width=22,
        ).grid(row=1, column=1, sticky="w", padx=5, pady=4)

        ttk.Label(marco_velocidad, text="Sexo:").grid(row=1, column=2, sticky="e", padx=5, pady=4)
        self.vel_sexo = tk.StringVar(value="Masculino")
        ttk.Combobox(
            marco_velocidad,
            textvariable=self.vel_sexo,
            values=["Masculino", "Femenino"],
            state="readonly",
            width=14,
        ).grid(row=1, column=3, sticky="w", padx=5, pady=4)

        ttk.Label(marco_velocidad, text="Intervalo OMS (meses):").grid(
            row=2, column=0, sticky="e", padx=5, pady=4)
        self.vel_intervalo = tk.StringVar(value="1")
        self.vel_intervalo_box = ttk.Combobox(
            marco_velocidad, textvariable=self.vel_intervalo,
            values=[str(m) for m in INTERVALOS_VEL_OMS["peso"]],
            state="readonly", width=8,
        )
        self.vel_intervalo_box.grid(row=2, column=1, sticky="w", padx=5, pady=4)
        self.vel_unidades = ttk.Label(marco_velocidad, text="Unidades: gramos (g)")
        self.vel_unidades.grid(row=2, column=2, columnspan=2, sticky="w", padx=5, pady=4)

        self.vel_edad1_label = ttk.Label(marco_velocidad, text="Edad 1 (meses):")
        self.vel_edad1_label.grid(row=3, column=0, sticky="e", padx=5, pady=4)
        self.vel_t1 = tk.StringVar(value="0")
        ttk.Entry(marco_velocidad, textvariable=self.vel_t1, width=10).grid(row=3, column=1, sticky="w", padx=5, pady=4)

        ttk.Label(marco_velocidad, text="Valor 1:").grid(row=3, column=2, sticky="e", padx=5, pady=4)
        self.vel_v1 = tk.StringVar(value="")
        ttk.Entry(marco_velocidad, textvariable=self.vel_v1, width=12).grid(row=3, column=3, sticky="w", padx=5, pady=4)

        self.vel_edad2_label = ttk.Label(marco_velocidad, text="Edad 2 (meses):")
        self.vel_edad2_label.grid(row=4, column=0, sticky="e", padx=5, pady=4)
        self.vel_t2 = tk.StringVar(value="")
        ttk.Entry(marco_velocidad, textvariable=self.vel_t2, width=10).grid(row=4, column=1, sticky="w", padx=5, pady=4)

        ttk.Label(marco_velocidad, text="Valor 2:").grid(row=4, column=2, sticky="e", padx=5, pady=4)
        self.vel_v2 = tk.StringVar(value="")
        ttk.Entry(marco_velocidad, textvariable=self.vel_v2, width=12).grid(row=4, column=3, sticky="w", padx=5, pady=4)

        self.vel_indicador.trace_add("write", self._cambiar_vel_indicador)
        self._cambiar_vel_indicador()
        ttk.Button(marco_velocidad, text="Calcular velocidad", command=self._calcular_velocidad_oms).grid(
            row=5, column=0, sticky="w", padx=5, pady=6,
        )
        ttk.Button(marco_velocidad, text="Usar peso actual", command=self._usar_peso_auto).grid(
            row=5, column=1, sticky="w", padx=5, pady=6,
        )
        self.vel_resultado = tk.Label(
            marco_velocidad, text="", justify="left", foreground="#0B5C7A", wraplength=760)
        self.vel_resultado.grid(row=6, column=0, columnspan=4, sticky="w", padx=5, pady=3)
        self._unidad_edad_velocidad = "meses"
        self.campos["fecha_nacimiento"].trace_add("write", self._actualizar_unidad_edad_velocidad)
        self.campos["fecha_actual"].trace_add("write", self._actualizar_unidad_edad_velocidad)
        self._actualizar_unidad_edad_velocidad()

    # ---------- Pestaña 3 Interacción Fármaco-Nutriente ----------
    def _tab_farmaco(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="3. Interacción Fármaco-Nutriente")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Fármaco-Nutriente:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_farmaco).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_farmaco).pack(side="left", padx=3)

        marco_busca = ttk.LabelFrame(tab, text="Buscador de fármacos (base LactMed de la NIH)")
        marco_busca.pack(fill="x", padx=8, pady=4)
        barra_busca = ttk.Frame(marco_busca)
        barra_busca.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Label(barra_busca, text="Fármaco:").pack(side="left", padx=5)
        self.lact_busca = tk.StringVar()
        entry_busca = ttk.Entry(barra_busca, textvariable=self.lact_busca, width=32)
        entry_busca.pack(side="left", padx=5)
        entry_busca.bind("<KeyRelease>", self._buscar_lactmed)
        entry_busca.bind("<Return>", lambda e: self._aplicar_lactmed())
        ttk.Button(barra_busca, text="Buscar", command=self._buscar_lactmed).pack(side="left", padx=3)
        ttk.Button(barra_busca, text="Aplicar selección", command=self._aplicar_lactmed).pack(side="left", padx=3)
        ttk.Button(barra_busca, text="Actualizar base LactMed", command=self._actualizar_lactmed).pack(side="left", padx=3)
        self.lact_info = ttk.Label(marco_busca, text="", foreground="#555555")
        self.lact_info.pack(anchor="w", padx=6, pady=(0, 4))

        contenedor_lista = ttk.Frame(marco_busca)
        contenedor_lista.pack(fill="x", padx=6, pady=(0, 6))
        self.lact_lista = tk.Listbox(contenedor_lista, height=6, exportselection=False)
        tsb = ttk.Scrollbar(contenedor_lista, orient="vertical", command=self.lact_lista.yview)
        self.lact_lista.configure(yscrollcommand=tsb.set)
        self.lact_lista.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        self.lact_lista.bind("<Double-Button-1>", lambda e: self._aplicar_lactmed())
        self._refrescar_lact_info()

        self.farmaco_cards = []
        self.farmaco_card_sel = None
        contenedor_f = ttk.Frame(tab)
        contenedor_f.pack(fill="both", expand=True, padx=8, pady=4)
        self.farmaco_canvas = tk.Canvas(contenedor_f, highlightthickness=0)
        tsb_f = ttk.Scrollbar(contenedor_f, orient="vertical", command=self.farmaco_canvas.yview)
        self.farmaco_interior = ttk.Frame(self.farmaco_canvas)
        ventana_f = self.farmaco_canvas.create_window(
            (0, 0), window=self.farmaco_interior, anchor="nw")
        self.farmaco_canvas.configure(yscrollcommand=tsb_f.set)
        self.farmaco_canvas.pack(side="left", fill="both", expand=True)
        tsb_f.pack(side="right", fill="y")

        def _actualizar_scroll_f(_event=None):
            self.farmaco_canvas.configure(scrollregion=self.farmaco_canvas.bbox("all"))

        def _ancho_scroll_f(_event=None):
            self.farmaco_canvas.itemconfigure(ventana_f, width=_event.width)

        self.farmaco_interior.bind("<Configure>", _actualizar_scroll_f)
        self.farmaco_canvas.bind("<Configure>", _ancho_scroll_f)

        botonera_f = ttk.Frame(tab)
        botonera_f.pack(fill="x", padx=8, pady=4)
        ttk.Button(botonera_f, text="+ Agregar fármaco", command=self._agregar_fila).pack(side="left", padx=3)
        ttk.Button(botonera_f, text="- Quitar seleccionado", command=self._quitar_fila).pack(side="left", padx=3)

    def _refrescar_lact_info(self):
        if not hasattr(self, "lact_info"):
            return
        try:
            conn = conectar()
            fila = conn.execute(
                "SELECT (SELECT COUNT(*) FROM lactmed), "
                "(SELECT fecha FROM lactmed_meta WHERE id = 1)"
            ).fetchone()
            conn.close()
        except sqlite3.Error:
            fila = None
        if fila and fila[0]:
            self.lact_info.config(
                text=f"Base LactMed {fila[1] or ''} · {fila[0]} fármacos indexados")
        else:
            self.lact_info.config(
                text="La base LactMed aún no está descargada. Use 'Actualizar base LactMed' "
                     "(primera descarga ~200 MB; requiere internet).")

    def _buscar_lactmed(self, *_args):
        self.lact_lista.delete(0, "end")
        palabra = self.lact_busca.get().strip()
        if not palabra:
            return
        conn = conectar()
        try:
            filas = conn.execute(
                "SELECT drug_name FROM lactmed WHERE drug_name LIKE ? OR aliases LIKE ? "
                "ORDER BY drug_name LIMIT 100",
                ("%" + palabra + "%", "%" + palabra + "%"),
            ).fetchall()
            total = conn.execute("SELECT COUNT(*) FROM lactmed").fetchone()[0]
        finally:
            conn.close()
        if total == 0:
            self.lact_lista.insert("end", "Descargue primero la base LactMed")
            return
        for (nombre,) in filas:
            self.lact_lista.insert("end", nombre)
        if not filas:
            self.lact_lista.insert("end", "Sin resultados")

    def _aplicar_lactmed(self):
        sel = self.lact_lista.curselection()
        if not sel:
            return
        nombre = self.lact_lista.get(sel[0])
        if nombre in ("Sin resultados", "Descargue primero la base LactMed"):
            return
        conn = conectar()
        fila = conn.execute(
            "SELECT summary, consideration, alternatives FROM lactmed WHERE drug_name = ?",
            (nombre,),
        ).fetchone()
        conn.close()
        if fila is None:
            return
        summary, consideration, alternativas = (f or "" for f in fila)
        partes = []
        for tx in (summary, consideration):
            tx = tx.strip()
            if tx and tx not in partes:
                partes.append(tx)
        if alternativas:
            partes.append("Alternativas: " + alternativas.strip())
        self._agregar_farmaco_card(farmaco=nombre, interaccion="\n\n".join(partes))
        self.status.config(text=f"Fármaco '{nombre}' aplicado desde LactMed")
        self._refrescar_lact_info()

    def _actualizar_lactmed(self):
        if not LACTMED_OK:
            messagebox.showerror("LactMed", "El módulo lactmed.py no está disponible")
            return
        if not (_LACTMED and hasattr(_LACTMED, "descargar_e_importar")):
            messagebox.showerror("LactMed", "Versión de lactmed.py incompatible")
            return
        if not messagebox.askyesno(
            "LactMed",
            "Descargar e importar la base LactMed completa desde NCBI?\n"
            "La primera descarga son ~200 MB y puede tardar varios minutos; "
            "luego queda cacheada localmente.",
        ):
            return
        self.status.config(text="Descargando e importando base LactMed (no cierre la app)...")
        import threading

        def trabajo():
            try:
                local = Path(__file__).resolve().parent / "data" / "lactmed_raw"
                conn = conectar()
                n = _LACTMED.descargar_e_importar(conn, local)
                conn.close()
                self.after(0, lambda: self._lact_ok(n))
            except Exception as e:
                self.after(0, lambda e=e: self._lact_error(str(e)))

        threading.Thread(target=trabajo, daemon=True).start()

    def _lact_ok(self, n):
        self._refrescar_lact_info()
        self.status.config(text=f"Base LactMed actualizada: {n} fármacos ({_LACTMED.TARBALL_NAME})")
        messagebox.showinfo("LactMed", f"Base actualizada con {n} fármacos.\n{_LACTMED.TARBALL_NAME}")

    def _lact_error(self, error):
        self.status.config(text="Error al actualizar la base LactMed")
        self._refrescar_lact_info()
        messagebox.showerror("LactMed", error)

    def _agregar_farmaco_card(self, farmaco="", interaccion=""):
        cont = tk.Frame(
            self.farmaco_interior, highlightbackground="#c9c9c9", highlightthickness=1)
        cont.pack(fill="x", padx=6, pady=3)

        def _seleccionar(_e=None):
            if self.farmaco_card_sel is cont:
                return
            if self.farmaco_card_sel is not None:
                self.farmaco_card_sel.configure(highlightbackground="#c9c9c9")
            self.farmaco_card_sel = cont
            cont.configure(highlightbackground="#0078D7")

        cont.bind("<Button-1>", _seleccionar)
        encab = ttk.Frame(cont)
        encab.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Label(encab, text="Fármaco:", font=("TkDefaultFont", 10, "bold")).pack(side="left")
        var_f = tk.StringVar(value=farmaco)
        e_f = ttk.Entry(encab, textvariable=var_f, width=40)
        e_f.pack(side="left", padx=4)
        ttk.Label(cont, text="Interacción:").pack(anchor="w", padx=6)
        tx_i = tk.Text(cont, height=8, wrap="word", width=92, font=("TkDefaultFont", 10),
                       highlightthickness=1, relief="solid")
        tx_i.pack(fill="x", padx=6, pady=(0, 5))
        tx_i.insert("1.0", interaccion)
        for wgt in (e_f, tx_i):
            wgt.bind("<Button-1>", _seleccionar, add="+")
        self.farmaco_cards.append({
            "frame": cont,
            "farmaco": var_f,
            "interaccion": tx_i,
        })
        self.farmaco_interior.update_idletasks()
        self.farmaco_canvas.configure(scrollregion=self.farmaco_canvas.bbox("all"))
        return cont

    def _limpiar_farmaco_cards(self):
        for c in list(self.farmaco_cards):
            c["frame"].destroy()
        self.farmaco_cards.clear()
        self.farmaco_card_sel = None
        if hasattr(self, "farmaco_interior"):
            self.farmaco_interior.update_idletasks()
            self.farmaco_canvas.configure(scrollregion=self.farmaco_canvas.bbox("all"))

    def _agregar_fila(self):
        self._agregar_farmaco_card()

    def _quitar_fila(self):
        if self.farmaco_card_sel is None:
            return
        for i, c in enumerate(self.farmaco_cards):
            if c["frame"] is self.farmaco_card_sel:
                c["frame"].destroy()
                del self.farmaco_cards[i]
                break
        self.farmaco_card_sel = None
        self.farmaco_interior.update_idletasks()
        self.farmaco_canvas.configure(scrollregion=self.farmaco_canvas.bbox("all"))

    # ---------- Pestaña 5 Ecuaciones Predictivas ----------
    def _tab_ecuaciones(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="5. Ecuaciones Predictivas")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Necesidades energéticas:").pack(side="left", padx=5)
        ttk.Button(barra, text="Calcular estimación", command=self._calcular_energia).pack(side="left", padx=3)
        ttk.Button(barra, text="Usar datos del paciente", command=self._usar_datos_paciente_energia).pack(
            side="left", padx=3)
        ttk.Button(barra, text="Guardar", command=self._guardar_energia).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_energia).pack(side="left", padx=3)
        ttk.Button(barra, text="Limpiar", command=self._limpiar_energia).pack(side="left", padx=3)

        form = ttk.LabelFrame(tab, text="Datos de entrada")
        form.pack(fill="x", padx=8, pady=4)
        form.columnconfigure(1, weight=1)

        self.ener_grupo = tk.StringVar(value=ENERGIA_GRUPO_LACTANTE)
        self.ener_sexo = tk.StringVar(value="Niño")
        self.ener_alimentacion = tk.StringVar(value=ENERGIA_ALIMENTACION_TODAS)
        self.ener_peso = tk.StringVar()
        self.ener_resultado = {"calculado": False, "kcal": None, "detalle": ""}

        ttk.Label(form, text="Grupo / fuente:").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        combo_grupo = ttk.Combobox(
            form, textvariable=self.ener_grupo, state="readonly", values=ENERGIA_GRUPOS)
        combo_grupo.grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        ttk.Label(form, text="Sexo:").grid(row=1, column=0, sticky="w", padx=6, pady=5)
        self.ener_combo_sexo = ttk.Combobox(
            form, textvariable=self.ener_sexo, state="readonly", width=34, values=ENERGIA_SEXOS)
        self.ener_combo_sexo.grid(row=1, column=1, sticky="w", padx=6, pady=5)
        ttk.Label(form, text="Tipo de alimentación:").grid(row=2, column=0, sticky="w", padx=6, pady=5)
        self.ener_combo_alim = ttk.Combobox(
            form, textvariable=self.ener_alimentacion, state="readonly", width=34,
            values=ENERGIA_ALIMENTACIONES_OPCIONES)
        self.ener_combo_alim.grid(row=2, column=1, sticky="w", padx=6, pady=5)
        ttk.Label(form, text="Peso actual (kg):").grid(row=3, column=0, sticky="w", padx=6, pady=5)
        self.ener_entry_peso = ttk.Entry(form, textvariable=self.ener_peso, width=18)
        self.ener_entry_peso.grid(row=3, column=1, sticky="w", padx=6, pady=5)
        self.ener_entry_peso.bind("<Return>", lambda _e: self._calcular_energia())

        salida = ttk.LabelFrame(tab, text="Resultado")
        salida.pack(fill="both", expand=True, padx=8, pady=4)
        self.ener_texto = tk.Text(
            salida, height=14, wrap="word", font=("Consolas", 10),
            highlightthickness=1, relief="solid")
        self.ener_texto.pack(fill="both", expand=True, padx=6, pady=6)
        self.ener_texto.insert(
            "1.0",
            "Recién nacidos y lactantes · resultados en kcal/día\n\n"
            "Complete los datos y pulse «Calcular estimación».",
        )
        self.ener_texto.configure(state="disabled")

        ttk.Label(tab, text=ENERGIA_NOTA, wraplength=900, foreground="#555555").pack(
            anchor="w", padx=10, pady=(0, 8))

        self.ener_grupo.trace_add("write", self._actualizar_energia_alimentacion)
        self._actualizar_energia_alimentacion()

    def _actualizar_energia_alimentacion(self, *_args):
        """En recién nacido la alimentación no aplica y el combo queda bloqueado."""
        if self.ener_grupo.get() == ENERGIA_GRUPO_RN:
            self.ener_combo_alim.configure(state="disabled")
            self.ener_alimentacion.set(ENERGIA_ALIMENTACION_NO_APLICA)
            return
        if self.ener_alimentacion.get() not in ENERGIA_ALIMENTACIONES_OPCIONES:
            self.ener_alimentacion.set(ENERGIA_ALIMENTACION_TODAS)
        self.ener_combo_alim.configure(state="readonly")

    def _peso_energia_kg(self):
        """Peso en kg desde el campo, aceptando gramos; None si no es válido."""
        texto = self.ener_peso.get().strip().replace(",", ".")
        if not texto:
            return None
        try:
            peso = float(texto)
        except ValueError:
            return None
        if peso <= 0:
            return None
        return peso / 1000.0 if peso > 10 else peso

    def _escribir_energia(self, texto):
        self.ener_texto.configure(state="normal")
        self.ener_texto.delete("1.0", "end")
        self.ener_texto.insert("1.0", texto)
        self.ener_texto.configure(state="disabled")

    def _calcular_energia(self):
        peso = self._peso_energia_kg()
        if peso is None:
            messagebox.showerror(
                "Dato inválido", "Introduzca un peso válido en kg, mayor que cero.")
            return
        grupo = self.ener_grupo.get()
        sexo = self.ener_sexo.get()
        alimentacion = self.ener_alimentacion.get()
        try:
            kcal, detalle = estimar_energia(grupo, sexo, alimentacion, peso)
        except ValueError as e:
            messagebox.showerror("Dato inválido", str(e))
            return
        self.ener_resultado = {
            "calculado": True, "kcal": round(kcal, 2), "detalle": detalle,
        }
        self._escribir_energia(detalle)
        self.status.config(text=f"Energía estimada: {kcal:.2f} kcal/día")

    def _usar_datos_paciente_energia(self):
        self.ener_sexo.set("Niño" if self.sexo.get() == "Masculino" else "Niña")
        gramos = self._peso_actual_g_guardado()
        if not gramos:
            messagebox.showwarning(
                "Peso",
                "No hay un peso actual registrado.\nCargue el peso en la pestaña de "
                "antropometría o ingréselo manualmente aquí.",
            )
            return
        self.ener_peso.set(f"{gramos / 1000.0:.3f}".rstrip("0").rstrip("."))
        self._calcular_energia()

    def _limpiar_energia(self):
        self.ener_grupo.set(ENERGIA_GRUPO_LACTANTE)
        self.ener_sexo.set("Niño")
        self.ener_peso.set("")
        self.ener_resultado = {"calculado": False, "kcal": None, "detalle": ""}
        self.ener_alimentacion.set(ENERGIA_ALIMENTACION_TODAS)
        self._actualizar_energia_alimentacion()
        self._escribir_energia(
            "Complete los datos y pulse «Calcular estimación».")

    def _guardar_energia(self):
        if not self._exigir_paciente():
            return
        if not self.ener_resultado.get("calculado"):
            self._calcular_energia()
            if not self.ener_resultado.get("calculado"):
                return
        try:
            conn = conectar()
            conn.execute(
                """INSERT INTO estimacion_energetica
                   (paciente_id, grupo, sexo, alimentacion, peso_kg, kcal_dia, detalle,
                    fecha)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    self._paciente_id, self.ener_grupo.get(), self.ener_sexo.get(),
                    self.ener_alimentacion.get(), self._peso_energia_kg(),
                    self.ener_resultado["kcal"],
                    self.ener_resultado["detalle"], date.today().isoformat(),
                ),
            )
            conn.commit()
            conn.close()
            self.status.config(
                text=f"Estimación energética guardada | Paciente N° {self._paciente_id}")
            messagebox.showinfo("Guardado", "Estimación de necesidades energéticas guardada")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _editar_energia(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        conn = conectar()
        fila = conn.execute(
            """SELECT grupo, sexo, alimentacion, peso_kg, kcal_dia, detalle
               FROM estimacion_energetica WHERE paciente_id = ? ORDER BY id DESC LIMIT 1""",
            (self._paciente_id,),
        ).fetchone()
        conn.close()
        if fila is None:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay estimaciones guardadas para este paciente")
            return
        self.ener_grupo.set(fila[0] or ENERGIA_GRUPO_LACTANTE)
        self.ener_sexo.set(fila[1] if fila[1] in ENERGIA_SEXOS else "Niño")
        if fila[2] in ENERGIA_ALIMENTACIONES_OPCIONES:
            self.ener_alimentacion.set(fila[2])
        if fila[3] is not None:
            self.ener_peso.set(f"{float(fila[3]):.3f}".rstrip("0").rstrip("."))
        self.ener_resultado = {"calculado": True, "kcal": fila[4],
                               "detalle": fila[5] or ""}
        self._escribir_energia(fila[5] or "Sin detalle guardado")
        self.status.config(
            text=f"Editando estimación energética | Paciente N° {self._paciente_id}")

    # ---------- Pestaña 6 Recuento Alimentario ----------
    def _tab_recuento(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="6. Recuento Alimentario")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Historia de alimentación:").pack(side="left", padx=5)
        ttk.Button(barra, text="Ver resumen", command=self._resumen_recuento).pack(side="left", padx=3)
        ttk.Button(barra, text="Guardar", command=self._guardar_recuento).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_recuento).pack(side="left", padx=3)
        ttk.Button(barra, text="Limpiar", command=self._limpiar_recuento).pack(side="left", padx=3)

        contenedor = ttk.Frame(tab)
        contenedor.pack(fill="both", expand=True)
        canvas = tk.Canvas(contenedor, highlightthickness=0)
        scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        interior = ttk.Frame(canvas)
        ventana = canvas.create_window((0, 0), window=interior, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._canvas_recuento = canvas

        def _actualizar_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _ancho_scroll(_event=None):
            canvas.itemconfigure(ventana, width=_event.width)

        def _rueda(_event):
            canvas.yview_scroll(int(-_event.delta / 120), "units")

        def _cambio_tab(_event=None):
            if self.notebook.select() == tab:
                canvas.bind_all("<MouseWheel>", _rueda)
            else:
                canvas.unbind_all("<MouseWheel>")

        interior.bind("<Configure>", _actualizar_scroll)
        canvas.bind("<Configure>", _ancho_scroll)
        self.notebook.bind("<<NotebookTabChanged>>", _cambio_tab)
        tab.bind("<Map>", _cambio_tab)
        tab.bind("<Unmap>", _cambio_tab)
        tab.bind("<Destroy>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        marco = ttk.LabelFrame(interior, text="Alimentación del niño")
        marco.pack(fill="x", padx=8, pady=4)
        marco.columnconfigure(1, weight=1)
        ttk.Label(marco, text="Grupo:").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        self.rec_grupo = tk.StringVar(value=RECUENTO_GRUPO_LACTANTE)
        ttk.Combobox(
            marco, textvariable=self.rec_grupo, state="readonly", width=30,
            values=RECUENTO_GRUPOS).grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        ttk.Label(marco, text="Alimentación:").grid(row=1, column=0, sticky="w", padx=6, pady=5)
        self.rec_tipo = tk.StringVar(value=RECUENTO_TIPO_MATERNA)
        self.rec_combo_tipo = ttk.Combobox(
            marco, textvariable=self.rec_tipo, state="readonly", width=30,
            values=RECUENTO_TIPOS_LACTANTE)
        self.rec_combo_tipo.grid(row=1, column=1, sticky="ew", padx=6, pady=5)
        ttk.Label(marco, text="Fecha de entrevista:").grid(row=2, column=0, sticky="w", padx=6, pady=5)
        self.rec_fecha = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(marco, textvariable=self.rec_fecha, width=18).grid(
            row=2, column=1, sticky="w", padx=6, pady=5)
        self.rec_grupo.trace_add("write", self._actualizar_recuento_secciones)
        self.rec_tipo.trace_add("write", self._actualizar_recuento_secciones)

        self.rec_vars = {}
        self.rec_textos = {}
        for titulo, seccion, preguntas in RECUENTO_SECCIONES:
            marco_s = ttk.LabelFrame(interior, text=titulo)
            marco_s.pack(fill="x", padx=8, pady=4)
            marco_s.columnconfigure(1, weight=1)
            for i, (clave, etiqueta, tipo) in enumerate(preguntas):
                ttk.Label(marco_s, text=etiqueta).grid(
                    row=i, column=0, sticky="w", padx=6, pady=4)
                if tipo == RECUENTO_TIPO_SI_NO:
                    var = tk.StringVar()
                    ttk.Combobox(
                        marco_s, textvariable=var, state="readonly", width=10,
                        values=RECUENTO_SI_NO).grid(row=i, column=1, sticky="w", padx=6, pady=4)
                elif tipo == RECUENTO_TIPO_NUMERO:
                    var = tk.StringVar()
                    ttk.Entry(marco_s, textvariable=var, width=10).grid(
                        row=i, column=1, sticky="w", padx=6, pady=4)
                elif tipo == RECUENTO_TIPO_LARGO:
                    var = tk.StringVar()
                    box = tk.Text(marco_s, height=5, width=40, wrap="word")
                    box.grid(row=i, column=1, sticky="ew", padx=6, pady=4)
                    self.rec_textos[clave] = box
                else:
                    var = tk.StringVar()
                    ttk.Entry(marco_s, textvariable=var, width=40).grid(
                        row=i, column=1, sticky="ew", padx=6, pady=4)
                self.rec_vars[clave] = var
                if clave == "f_intervalo":
                    var.trace_add("write", self._completar_intervalo_recuento)

        marco_obs = ttk.LabelFrame(interior, text="Observaciones")
        marco_obs.pack(fill="x", padx=8, pady=4)
        self.rec_observaciones = tk.Text(marco_obs, height=4, wrap="word")
        self.rec_observaciones.pack(fill="x", padx=6, pady=6)

        marco_salida = ttk.LabelFrame(interior, text="Resumen de la entrevista")
        marco_salida.pack(fill="both", expand=True, padx=8, pady=4)
        self.rec_texto = tk.Text(marco_salida, height=12, wrap="word")
        self.rec_texto.pack(fill="both", expand=True, padx=6, pady=6)
        self.rec_texto.configure(state="disabled")

        ttk.Label(interior, text=RECUENTO_NOTA, wraplength=880, foreground="#555555").pack(
            anchor="w", padx=10, pady=(0, 8))

        self._rec_grupos_visibles = None
        self._actualizar_recuento_secciones()

    def _completar_intervalo_recuento(self, *_args):
        """Sugiere el intervalo («cada 2 h»…) a partir del número de tomas."""
        texto = self.rec_vars["f_tomas_24h"].get().strip()
        if self.rec_vars["f_intervalo"].get().strip() or not texto:
            return
        try:
            tomas = float(texto.replace(",", "."))
        except ValueError:
            return
        if tomas > 0:
            self.rec_vars["f_intervalo"].set(f"Cada {24.0 / tomas:.1f} h")

    def _secciones_recuento_visibles(self):
        """Secciones que aplican según grupo y tipo de alimentación."""
        if self.rec_grupo.get() == RECUENTO_GRUPO_MAYOR:
            return {"mayor"}
        tipo = self.rec_tipo.get()
        visibles = {"mayor"}
        if tipo == RECUENTO_TIPO_MATERNA:
            visibles.add("materna")
        elif tipo == RECUENTO_TIPO_SUCEDANEO:
            visibles.add("formula")
        else:
            visibles.update({"materna", "formula"})
        return visibles

    def _actualizar_recuento_secciones(self, *_args):
        """Muestra u oculta las secciones según el grupo y la alimentación."""
        lactante = self.rec_grupo.get() == RECUENTO_GRUPO_LACTANTE
        if lactante and self.rec_tipo.get() == RECUENTO_TIPO_NO_APLICA:
            self.rec_tipo.set(RECUENTO_TIPO_MATERNA)
        if not lactante:
            self.rec_tipo.set(RECUENTO_TIPO_NO_APLICA)
            self.rec_combo_tipo.configure(values=[RECUENTO_TIPO_NO_APLICA], state="disabled")
        else:
            self.rec_combo_tipo.configure(values=RECUENTO_TIPOS_LACTANTE, state="readonly")
        visibles = self._secciones_recuento_visibles()
        self._actualizar_filas_recuento(visibles)

    def _actualizar_filas_recuento(self, visibles):
        """Activa o desactiva las filas de preguntas no aplicables."""
        if visibles == self._rec_grupos_visibles:
            return
        self._rec_grupos_visibles = visibles
        for _titulo, seccion, preguntas in RECUENTO_SECCIONES:
            activa = seccion in visibles
            for clave, _etiqueta, _tipo in preguntas:
                if clave in self.rec_textos:
                    estado = "normal" if activa else "disabled"
                    self.rec_textos[clave].configure(state=estado)
                else:
                    var = self.rec_vars[clave]
                    if activa:
                        if not var.get():
                            var.set("")
                    else:
                        var.set("")

    def _respuestas_recuento(self):
        """Respuestas contestadas de las secciones visibles, por clave."""
        self.rec_texto.configure(state="normal")
        datos = {}
        for _titulo, seccion, preguntas in RECUENTO_SECCIONES:
            if seccion not in self._rec_grupos_visibles:
                continue
            for clave, _etiqueta, _tipo in preguntas:
                if clave in self.rec_textos:
                    valor = self.rec_textos[clave].get("1.0", "end").strip()
                else:
                    valor = self.rec_vars[clave].get().strip()
                if valor:
                    datos[clave] = valor
        return datos

    def _escribir_recuento(self, texto):
        self.rec_texto.configure(state="normal")
        self.rec_texto.delete("1.0", "end")
        self.rec_texto.insert("1.0", texto)
        self.rec_texto.configure(state="disabled")

    def _resumen_recuento(self):
        self._respuestas_recuento()
        visibles = self._rec_grupos_visibles
        lineas = [
            f"Fecha de entrevista: {self.rec_fecha.get().strip()}",
            f"Grupo: {self.rec_grupo.get()}",
            f"Alimentación: {self.rec_tipo.get()}",
        ]
        for titulo, seccion, preguntas in RECUENTO_SECCIONES:
            if seccion not in visibles:
                continue
            contestadas = [
                p for p in preguntas
                if (self.rec_textos[p[0]].get("1.0", "end").strip()
                    if p[0] in self.rec_textos else self.rec_vars[p[0]].get().strip())
            ]
            if not contestadas:
                continue
            lineas.append("")
            lineas.append(titulo)
            for clave, etiqueta, _tipo in contestadas:
                valor = (self.rec_textos[clave].get("1.0", "end").strip()
                         if clave in self.rec_textos else self.rec_vars[clave].get().strip())
                lineas.append(f"  · {etiqueta} {valor}")
        total = sum(
            1 for _t, s, qs in RECUENTO_SECCIONES if s in visibles for _c, _e, _ti in qs
            if (self.rec_textos[_c].get("1.0", "end").strip()
                if _c in self.rec_textos else self.rec_vars[_c].get().strip())
        )
        lineas.append("")
        lineas.append(f"Preguntas contestadas: {total}")
        lineas.append("")
        lineas.append(RECUENTO_NOTA)
        self._escribir_recuento("\n".join(lineas))
        self.status.config(text=f"Resumen del recuento alimentario | {total} respuestas")

    def _limpiar_recuento(self):
        self.rec_grupo.set(RECUENTO_GRUPO_LACTANTE)
        self.rec_tipo.set(RECUENTO_TIPO_MATERNA)
        self.rec_fecha.set(date.today().isoformat())
        for clave, var in self.rec_vars.items():
            var.set("")
        for box in self.rec_textos.values():
            box.configure(state="normal")
            box.delete("1.0", "end")
        self.rec_observaciones.delete("1.0", "end")
        self._rec_grupos_visibles = None
        self._actualizar_recuento_secciones()
        self._escribir_recuento(
            "Complete las preguntas que apliquen y pulse «Ver resumen».\n")

    def _guardar_recuento(self):
        if not self._exigir_paciente():
            return
        datos = self._respuestas_recuento()
        if not datos:
            messagebox.showwarning(
                "Datos", "Complete al menos una pregunta del recuento alimentario.")
            return
        try:
            conn = conectar()
            anterior = conn.execute(
                "SELECT id FROM recuento_alimentario WHERE paciente_id = ? "
                "ORDER BY id DESC LIMIT 1",
                (self._paciente_id,),
            ).fetchone()
            valores = (
                self.rec_fecha.get().strip(), self.rec_grupo.get(), self.rec_tipo.get(),
                json.dumps(datos, ensure_ascii=False),
                self.rec_observaciones.get("1.0", "end").strip(),
            )
            if anterior:
                conn.execute(
                    """UPDATE recuento_alimentario
                       SET fecha_entrevista=?, grupo=?, tipo_alimentacion=?,
                           respuestas_json=?, observaciones=?
                       WHERE id=?""",
                    (*valores, anterior[0]),
                )
            else:
                conn.execute(
                    """INSERT INTO recuento_alimentario
                       (paciente_id, fecha_entrevista, grupo, tipo_alimentacion,
                        respuestas_json, observaciones)
                       VALUES (?,?,?,?,?,?)""",
                    (self._paciente_id, *valores),
                )
            conn.commit()
            conn.close()
            self.status.config(
                text=f"Recuento alimentario guardado | Paciente N° {self._paciente_id}")
            messagebox.showinfo("Guardado", "Historia de alimentación guardada")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _editar_recuento(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        conn = conectar()
        fila = conn.execute(
            """SELECT fecha_entrevista, grupo, tipo_alimentacion, respuestas_json,
                      observaciones
               FROM recuento_alimentario WHERE paciente_id = ? ORDER BY id DESC LIMIT 1""",
            (self._paciente_id,),
        ).fetchone()
        conn.close()
        if fila is None:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay recuentos guardados para este paciente")
            return
        self.rec_fecha.set(fila[0] or date.today().isoformat())
        self.rec_grupo.set(fila[1] if fila[1] in RECUENTO_GRUPOS else RECUENTO_GRUPO_LACTANTE)
        self.rec_tipo.set(fila[2] if fila[2] in RECUENTO_TIPOS_LACTANTE else RECUENTO_TIPO_MATERNA)
        self._rec_grupos_visibles = None
        self._actualizar_recuento_secciones()
        try:
            datos = json.loads(fila[3] or "{}")
        except ValueError:
            datos = {}
        for clave, valor in datos.items():
            if clave in self.rec_textos:
                box = self.rec_textos[clave]
                box.configure(state="normal")
                box.delete("1.0", "end")
                box.insert("1.0", valor)
            elif clave in self.rec_vars:
                self.rec_vars[clave].set(valor)
        self.rec_observaciones.delete("1.0", "end")
        self.rec_observaciones.insert("1.0", fila[4] or "")
        self.status.config(
            text=f"Editando recuento alimentario | Paciente N° {self._paciente_id}")

    # ---------- Pestaña 4 Evaluación Bioquímica ----------
    def _tab_bioquimica(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="4. Evaluación Bioquímica")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Evaluación Bioquímica:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_bioquimica).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_bioquimica).pack(side="left", padx=3)
        ttk.Button(barra, text="Limpiar", command=self._limpiar_bioquimica).pack(side="left", padx=3)
        ttk.Label(barra, text="  Referencia para edad:").pack(side="left", padx=(10, 0))
        self.bio_edad_sel = tk.StringVar(value="Automática según edad")
        combo_edad = ttk.Combobox(
            barra, textvariable=self.bio_edad_sel, state="readonly", width=20,
            values=["Automática según edad"] + BIOQ_COLUMNAS_EDAD,
        )
        combo_edad.pack(side="left", padx=5)
        combo_edad.bind("<<ComboboxSelected>>", self._actualizar_edad_bioquimica)

        contenedor = ttk.Frame(tab)
        contenedor.pack(fill="both", expand=True)
        canvas = tk.Canvas(contenedor, highlightthickness=0)
        scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        interior = ttk.Frame(canvas)
        ventana = canvas.create_window((0, 0), window=interior, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _actualizar_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _ancho_scroll(_event=None):
            canvas.itemconfigure(ventana, width=_event.width)

        interior.bind("<Configure>", _actualizar_scroll)
        canvas.bind("<Configure>", _ancho_scroll)

        conn = conectar()
        filas = conn.execute(
            "SELECT grupo, prueba, unidad, ref_nacimiento, ref_1_semana, ref_1_mes, "
            "ref_1_5_mes, notas FROM referencia_bioquimica ORDER BY id"
        ).fetchall()
        conn.close()
        self.bioq_pruebas = []
        self.bioq_grupos = []
        self.bioq_ref = {}
        grupos = {}
        for grupo, prueba, unidad, r0, r1, r2, r3, notas in filas:
            grupos.setdefault(grupo, []).append(prueba)
            self.bioq_ref[prueba] = {
                "grupo": grupo,
                "unidad": (unidad or "").strip(),
                "refs": [(r0 or "").strip(), (r1 or "").strip(), (r2 or "").strip(), (r3 or "").strip()],
                "notas": (notas or "").strip(),
            }
            self.bioq_pruebas.append(prueba)
        self.bioq_grupos = [(g, grupos[g]) for g in grupos]

        ttk.Label(
            interior,
            text="Ingrese el resultado de cada prueba: se compara automáticamente con el valor "
                 "de referencia QS según la edad del paciente (Bajo / Normal / Alto) y se guarda "
                 "con la evaluación.",
            foreground="#555555", wraplength=850,
        ).pack(anchor="w", padx=6, pady=2)
        self.bioq_edad_info = ttk.Label(interior, text="", foreground="#555555")
        self.bioq_edad_info.pack(anchor="w", padx=6, pady=1)
        cabecera = ttk.Frame(interior)
        cabecera.pack(fill="x", padx=6, pady=2)
        for col, texto in enumerate((
            "Pruebas Bioquímicas", "Unidad", "Valor normal (según edad)",
            "Resultado (input)", "Evaluación",
        )):
            ttk.Label(cabecera, text=texto, font=("TkDefaultFont", 9, "bold")).grid(
                row=0, column=col, sticky="w", padx=4)

        self.bio_resultado = {}
        self.bio_estado = {}
        self.bio_estado_label = {}
        self.bio_ref_label = {}
        for grupo, pruebas in self.bioq_grupos:
            ttk.Label(interior, text=grupo, font=("TkDefaultFont", 9, "bold"),
                      foreground="#0B5C7A").pack(anchor="w", padx=6, pady=(6, 1))
            for prueba in pruebas:
                info = self.bioq_ref[prueba]
                fila = ttk.Frame(interior)
                fila.pack(fill="x", padx=6, pady=1)
                ttk.Label(fila, text=prueba, width=40, anchor="w").grid(
                    row=0, column=0, sticky="w", padx=4)
                ttk.Label(fila, text=info["unidad"], width=10, anchor="w",
                          foreground="#555555").grid(row=0, column=1, sticky="w", padx=4)
                lbl_ref = ttk.Label(fila, text="", width=26, anchor="w", foreground="#0B5C7A")
                lbl_ref.grid(row=0, column=2, sticky="w", padx=4)
                self.bio_ref_label[prueba] = lbl_ref
                var = tk.StringVar()
                self.bio_resultado[prueba] = var
                var.trace_add("write", lambda *_, p=prueba: self._comparar_bioquimica(p))
                ttk.Entry(fila, textvariable=var, width=16).grid(
                    row=0, column=3, sticky="w", padx=4)
                estado_var = tk.StringVar(value="—")
                self.bio_estado[prueba] = estado_var
                lbl = tk.Label(fila, textvariable=estado_var, width=22, anchor="w", foreground="#555555")
                lbl.grid(row=0, column=4, sticky="w", padx=4)
                self.bio_estado_label[prueba] = lbl

        for clave in ("fecha_nacimiento", "fecha_actual"):
            if hasattr(self.campos[clave], "trace_add"):
                self.campos[clave].trace_add("write", self._actualizar_edad_bioquimica)
        self._actualizar_edad_bioquimica()

    def _dias_vida(self):
        return self._parsedias_desde_nacimiento()

    def _columna_bioq_auto(self, dias):
        if dias is None:
            return 0
        if dias <= 7:
            return 0
        if dias <= 27:
            return 1
        if dias <= 40:
            return 2
        return 3

    def _columna_bioq(self):
        sel = self.bio_edad_sel.get()
        if sel in BIOQ_COLUMNAS_EDAD:
            return BIOQ_COLUMNAS_EDAD.index(sel)
        return self._columna_bioq_auto(self._dias_vida())

    def _referencia_bioq(self, prueba):
        """Devuelve (texto, edad_usada) de la referencia QS para la prueba."""
        refs = self.bioq_ref[prueba]["refs"]
        col = self._columna_bioq()
        texto = refs[col]
        if not texto:
            for c in range(4):
                if refs[c]:
                    texto, col = refs[c], c
                    break
        return texto, BIOQ_COLUMNAS_EDAD[col]

    def _actualizar_ref_label(self, prueba):
        texto, edad_uso = self._referencia_bioq(prueba)
        if texto and edad_uso != BIOQ_COLUMNAS_EDAD[self._columna_bioq()]:
            vista = f"{texto} [{edad_uso}]"
        else:
            vista = texto
        self.bio_ref_label[prueba].config(text=vista)

    def _actualizar_edad_bioquimica(self, *_args):
        if not hasattr(self, "bio_edad_sel"):
            return
        auto = self.bio_edad_sel.get() not in BIOQ_COLUMNAS_EDAD
        if auto:
            dias = self._dias_vida()
            if dias is None:
                cola = "Complete la fecha de nacimiento para calcular la edad."
            else:
                col = self._columna_bioq_auto(dias)
                cola = f"Edad calculada: {dias} días → referencia de '{BIOQ_COLUMNAS_EDAD[col]}'"
        else:
            cola = f"Referencia fija: '{self.bio_edad_sel.get()}'"
        self.bioq_edad_info.config(text=cola)
        for p in self.bioq_pruebas:
            self._actualizar_ref_label(p)
            self._comparar_bioquimica(p)

    # ---------- Pestaña 2 Signos Clínicos (CANS score) ----------
    def _tab_cans(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="2. Signos Clínicos (CANS)")
        self._tab_signos = tab

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="CANS Score:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_cans).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_cans).pack(side="left", padx=3)
        ttk.Button(barra, text="Limpiar", command=self._limpiar_cans).pack(side="left", padx=3)
        ttk.Button(barra, text="Guardar en Excel", command=self._exportar_cans_xls).pack(side="left", padx=3)

        marco_fecha = ttk.LabelFrame(tab, text="Datos de la valoración")
        marco_fecha.pack(fill="x", padx=8, pady=4)
        ttk.Label(marco_fecha, text="Fecha de la valoración:").grid(
            row=0, column=0, sticky="e", padx=5, pady=4)
        self.cans_fecha = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(marco_fecha, textvariable=self.cans_fecha, width=14).grid(
            row=0, column=1, sticky="w", padx=5, pady=4)
        ttk.Label(marco_fecha, text="Observaciones:").grid(
            row=0, column=2, sticky="e", padx=5, pady=4)
        self.cans_observaciones = tk.Text(marco_fecha, width=52, height=3)
        self.cans_observaciones.grid(row=0, column=3, sticky="w", padx=5, pady=4)

        contenedor = ttk.Frame(tab)
        contenedor.pack(fill="both", expand=True)
        canvas = tk.Canvas(contenedor, highlightthickness=0)
        scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        interior = ttk.Frame(canvas)
        ventana = canvas.create_window((0, 0), window=interior, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._canvas_cans = canvas

        def _actualizar_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _ancho_scroll(_event=None):
            canvas.itemconfigure(ventana, width=_event.width)

        def _rueda(_event):
            canvas.yview_scroll(int(-_event.delta / 120), "units")

        def _cambio_tab(_event=None):
            if self.notebook.select() == tab:
                canvas.bind_all("<MouseWheel>", _rueda)
            else:
                canvas.unbind_all("<MouseWheel>")

        interior.bind("<Configure>", _actualizar_scroll)
        canvas.bind("<Configure>", _ancho_scroll)
        self.notebook.bind("<<NotebookTabChanged>>", _cambio_tab)
        tab.bind("<Map>", _cambio_tab)
        tab.bind("<Unmap>", _cambio_tab)
        tab.bind("<Destroy>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        self.cans_vars = {}
        for nombre, tecnica, opciones in CANS_SIGNOS:
            marco = ttk.LabelFrame(interior, text=f"{nombre} — {tecnica}")
            marco.pack(fill="x", padx=8, pady=4)
            fila = ttk.Frame(marco)
            fila.pack(fill="x", padx=5, pady=3)
            variable = tk.IntVar(value=max(CANS_PUNTOS))
            self.cans_vars[nombre] = variable
            for puntos in CANS_PUNTOS:
                ttk.Radiobutton(
                    fila,
                    text=f"{puntos}  ·  {opciones[puntos]}",
                    value=puntos,
                    variable=variable,
                    command=self._recalcular_cans,
                ).pack(anchor="w", pady=1)

        marco_total = ttk.LabelFrame(interior, text="Resultado CANS score")
        marco_total.pack(fill="x", padx=8, pady=6)
        self.cans_total_label = tk.Label(
            marco_total, text="", justify="left", font=("TkDefaultFont", 11, "bold"),
            foreground="#1E6E5C", wraplength=860)
        self.cans_total_label.pack(anchor="w", padx=8, pady=6)
        self.cans_detalle_label = tk.Label(
            marco_total, text="", justify="left", foreground="#555555", wraplength=860)
        self.cans_detalle_label.pack(anchor="w", padx=8, pady=(0, 6))

        marco_ip = ttk.LabelFrame(interior, text="Índice ponderal de Rohrer")
        self.ip_label = tk.Label(
            marco_ip, text="", justify="left", font=("TkDefaultFont", 11, "bold"),
            foreground="#555555", wraplength=860)
        self.ip_label.pack(anchor="w", padx=8, pady=6)
        self.ip_nota_label = tk.Label(
            marco_ip, text="", justify="left", foreground="#555555", wraplength=860)
        self.ip_nota_label.pack(anchor="w", padx=8, pady=(0, 6))
        ttk.Label(
            marco_ip,
            text="IP = peso al nacer (g) × 100 / longitud al nacer (cm)³  ·  "
                 "toma los datos de '1. Antropometría' → Antecedentes Prenatales",
            foreground="#555555",
        ).pack(anchor="w", padx=8, pady=(0, 6))
        self.ip_patron_label = tk.Label(
            marco_ip, text="", justify="left", font=("TkDefaultFont", 10, "bold"),
            foreground="#1E6E5C", wraplength=860)
        self.ip_patron_label.pack(anchor="w", padx=8, pady=(0, 8))

        marco_motor = ttk.LabelFrame(interior, text="Desarrollo Motor — OMS-MGRS (2006)")
        marco_motor.pack(fill="x", padx=8, pady=6)
        barra_motor = ttk.Frame(marco_motor)
        barra_motor.pack(fill="x", padx=5, pady=3)
        ttk.Label(barra_motor, text="Fecha de la valoración:").pack(side="left", padx=3)
        self.motor_fecha = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(barra_motor, textvariable=self.motor_fecha, width=14).pack(side="left", padx=3)
        ttk.Button(barra_motor, text="Guardar", command=self._guardar_desarrollo_motor).pack(side="left", padx=3)
        ttk.Button(barra_motor, text="Editar", command=self._editar_desarrollo_motor).pack(side="left", padx=3)
        ttk.Button(barra_motor, text="Limpiar", command=self._limpiar_desarrollo_motor).pack(side="left", padx=3)
        ttk.Button(barra_motor, text="Guardar en Excel", command=self._exportar_motor_xls).pack(side="left", padx=3)
        ttk.Label(
            marco_motor,
            text="Registre la fecha en que el niño logró por primera vez cada hito o escriba "
                 "directamente la edad (días). Los días se calculan desde la fecha de nacimiento y se "
                 "ubican en los percentiles P1-P99 (ventana esperada P3-P97).",
            foreground="#555555", wraplength=850,
        ).pack(anchor="w", padx=6, pady=2)
        cabecera = ttk.Frame(marco_motor)
        cabecera.pack(fill="x", padx=6, pady=2)
        for col, texto in enumerate((
            "Hito", "Fecha de logro", "Edad (días)", "Percentil", "Clasificación",
            "Definición del hito",
        )):
            ttk.Label(cabecera, text=texto, font=("TkDefaultFont", 9, "bold")).grid(
                row=0, column=col, sticky="w", padx=4)
        self.motor_logro = {}
        self.motor_edad = {}
        self.motor_pct = {}
        self.motor_clas = {}
        self.motor_clas_label = {}
        for hito in DESARROLLO_MOTOR_HITOS:
            fila = ttk.Frame(marco_motor)
            fila.pack(fill="x", padx=6, pady=1)
            ttk.Label(fila, text=hito, width=26, anchor="w").grid(row=0, column=0, sticky="w", padx=4)
            var = tk.StringVar()
            self.motor_logro[hito] = var
            var.trace_add("write", lambda *_, h=hito: self._calcular_desarrollo_motor(h))
            ttk.Entry(fila, textvariable=var, width=12).grid(row=0, column=1, sticky="w", padx=4)
            edad_var = tk.StringVar()
            self.motor_edad[hito] = edad_var
            edad_var.trace_add("write", lambda *_, h=hito: self._calcular_desarrollo_motor(h))
            ttk.Entry(fila, textvariable=edad_var, width=8).grid(
                row=0, column=2, sticky="w", padx=4)
            pct_var = tk.StringVar()
            self.motor_pct[hito] = pct_var
            ttk.Label(fila, textvariable=pct_var, width=8, anchor="w").grid(
                row=0, column=3, sticky="w", padx=4)
            clas_var = tk.StringVar()
            self.motor_clas[hito] = clas_var
            lbl_clas = tk.Label(fila, textvariable=clas_var, width=30, anchor="w", foreground="#555555")
            lbl_clas.grid(row=0, column=4, sticky="w", padx=4)
            self.motor_clas_label[hito] = lbl_clas
            tk.Label(
                fila, text=DESARROLLO_MOTOR_DEFINICIONES.get(hito, ""),
                justify="left", anchor="w", wraplength=330, foreground="#555555",
                font=("TkDefaultFont", 9),
            ).grid(row=0, column=5, sticky="nw", padx=4)
        marco_pct = ttk.LabelFrame(marco_motor, text="Percentiles de referencia por hito (OMS-MGRS, en días)")
        marco_pct = ttk.LabelFrame(marco_motor, text="Percentiles de referencia por hito (OMS-MGRS, en días)")
        marco_pct.pack(fill="x", padx=6, pady=4)
        columnas_pct = ["Hito", "P1", "P3", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "P97", "P99"]
        self.motor_tree = ttk.Treeview(marco_pct, columns=columnas_pct, show="headings", height=7)
        for col in columnas_pct:
            self.motor_tree.heading(col, text=col)
            self.motor_tree.column(col, width=(212 if col == "Hito" else 52), anchor="center")
        self.motor_tree.column("Hito", anchor="w")
        self.motor_tree.pack(fill="both", expand=True, padx=4, pady=4)
        conn = conectar()
        for hito in DESARROLLO_MOTOR_HITOS:
            filas = conn.execute(
                "SELECT percentil, edad_dias FROM referencia_desarrollo_motor WHERE hito = ? ORDER BY percentil",
                (hito,),
            ).fetchall()
            self.motor_tree.insert("", "end", values=[hito] + [str(d) for _, d in filas])
        conn.close()

        self._recalcular_cans()
        self._recalcular_ip()
        for clave in ("peso_nacer", "longitud_nacer"):
            variable = self.ant.get(clave)
            if variable is not None:
                variable.trace_add("write", lambda *_: self._recalcular_ip())

    # ---------- Acciones ----------
    def _parsedias_desde_nacimiento(self):
        texto = self.campos["fecha_nacimiento"].get().strip()
        if not texto:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                fecha = datetime.strptime(texto, fmt).date()
                break
            except ValueError:
                continue
        else:
            return None
        hasta = self.campos["fecha_actual"].get().strip()
        if hasta:
            try:
                ref = datetime.strptime(hasta, "%Y-%m-%d").date()
            except ValueError:
                return None
        else:
            ref = date.today()
        return (ref - fecha).days

    def _actualizar_clas_peso_g(self, *args):
        texto = self.ant["peso_nacer"].get().replace(",", ".").strip()
        try:
            valor = float(texto)
            gramos = int(valor * 1000) if valor < 10 else int(valor)
            self.clas_peso_g.set(str(gramos))
        except ValueError:
            self.clas_peso_g.set("")
        self._clasificar_peso_nacer()

    def _calcular_edad(self):
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            self.campos["edad"].set("")
            return
        semanas, resto = divmod(dias, 7)
        self.campos["edad"].set(f"{semanas} semanas y {resto} días")

    def _clasificar(self):
        self._clasificar_peso_nacer(silencioso=False)

    def _clasificar_peso_nacer(self, *args, silencioso=True):
        """Clasifica el peso al nacer contra la referencia de la semana de
        gestación del paciente. Silencioso por defecto: se invoca automáticamente
        al cambiar la EG, el sexo o el peso."""
        def _aviso(titulo, mensaje):
            self._clasificacion = ""
            self._codigo_peso_eg = ""
            self.clas_peso_ok = False
            if not silencioso:
                messagebox.showwarning(titulo, mensaje)
            elif hasattr(self, "clas_resultado"):
                self.clas_resultado.config(text=mensaje, foreground="#8B4513")
            if hasattr(self, "ip_patron_label"):
                self._clasificar_patron_crecimiento()

        try:
            semanas = int(re.match(r"\d+", self.ant["edad_gestacional"].get().strip()).group())
        except (AttributeError, ValueError):
            _aviso("Clasificación", "Ingrese la edad gestacional en semanas (ej.: 39)")
            return
        try:
            peso = float(self.ant["peso_nacer"].get().replace(",", ".").strip())
            if peso < 10:
                peso *= 1000
        except ValueError:
            _aviso("Clasificación", "Ingrese el peso al nacer (en kg en Antecedentes Prenatales)")
            return
        sexo = self.sexo_nacer.get()
        conn = conectar()
        fila = conn.execute(
            "SELECT percentil_10_g, percentil_50_g, percentil_90_g FROM referencia_peso_nacer WHERE sexo = ? AND edad_gestacional_semanas = ?",
            (sexo, semanas),
        ).fetchone()
        conn.close()
        if fila is None:
            _aviso("Clasificación", f"No hay referencia para la semana {semanas} ({sexo})")
            return
        p10, p50, p90 = fila
        if peso < p10:
            codigo = "PEG"
            texto = "Pequeño para la Edad de Gestación"
            ubicacion = f"por debajo del P10 ({int(p10)} g)"
        elif peso <= p90:
            codigo = "AEG"
            texto = "Apropiado para la Edad de Gestación"
            ubicacion = f"entre P10 ({int(p10)} g) y P90 ({int(p90)} g)"
        else:
            codigo = "GEG"
            texto = "Grande para la Edad de Gestación"
            ubicacion = f"por encima del P90 ({int(p90)} g)"
        self._clasificacion = f"{codigo} - {texto}"
        self._codigo_peso_eg = codigo
        self.clas_peso_ok = True
        diferencia = peso - p50
        porcentaje = diferencia / p50 * 100.0 if p50 else 0.0
        self.clas_resultado.config(
            text=(
                f"Peso: {int(peso)} g  |  EG: {semanas} semanas  |  Sexo: {sexo}\n"
                f"Referencia {semanas} sem: P10 {int(p10)} g | P50 {int(p50)} g | P90 {int(p90)} g\n"
                f"{codigo}: {texto} ({ubicacion})\n"
                f"Respecto al P50: {diferencia:+.0f} g ({porcentaje:+.1f} %)"
            ),
            foreground="blue" if codigo == "AEG" else "#B00020",
        )
        if hasattr(self, "ip_patron_label"):
            self._clasificar_patron_crecimiento()

    def _valorar_ganancia(self):
        self.velocidad_neonatal_resultado.config(text="")
        self._valoracion_velocidad = ""
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            messagebox.showwarning("Ganancia de peso", "Ingrese la fecha de nacimiento y la fecha actual válidas")
            return
        if dias < 0:
            messagebox.showwarning("Ganancia de peso", "La fecha actual no puede ser anterior a la de nacimiento")
            return
        if dias > 60:
            messagebox.showwarning("Ganancia de peso", "No hay referencia para más de 60 días de vida")
            return
        intervalo = _intervalo_ganancia(dias)
        if intervalo is None:
            messagebox.showwarning("Ganancia de peso", "No hay referencia para el intervalo neonatal solicitado")
            return
        try:
            peso_nacer = float(self.ant["peso_nacer"].get().replace(",", ".").strip())
            if peso_nacer < 10:
                peso_nacer *= 1000
            peso_actual = float(self.clas_peso_actual.get().replace(",", ".").strip())
            if peso_actual < 10:
                peso_actual *= 1000
        except ValueError:
            messagebox.showwarning("Ganancia de peso", "Ingrese el peso al nacer y el peso actual")
            return
        sexo = self.sexo_nacer.get()
        conn = conectar()
        filas = conn.execute(
            """SELECT percentil, peso_nacer_2000_2500_g, peso_nacer_2500_3000_g,
               peso_nacer_3000_3500_g, peso_nacer_3500_4000_g, peso_nacer_4000_mas_g, todos_g
               FROM referencia_ganancia_peso
               WHERE sexo = ? AND intervalo_dias = ?
               ORDER BY percentil ASC""",
            (sexo, intervalo),
        ).fetchall()
        conn.close()
        if not filas:
            messagebox.showwarning("Ganancia de peso", f"No hay referencia para {sexo} / {intervalo} días")
            return
        columna = self._columna_peso_nacer(peso_nacer)
        referencia = {}
        for fila in filas:
            per = fila[0]
            valor = fila[1 + {"peso_nacer_2000_2500_g": 0, "peso_nacer_2500_3000_g": 1, "peso_nacer_3000_3500_g": 2,
                              "peso_nacer_3500_4000_g": 3, "peso_nacer_4000_mas_g": 4, "todos_g": 5}[columna]]
            if valor is None:
                valor = fila[6]
            referencia[per] = valor
        ganancia = peso_actual - peso_nacer
        p5, p10, p25, p50 = referencia[5], referencia[10], referencia[25], referencia[50]
        if ganancia <= p5:
            percentil = "< P5"
        elif ganancia <= p10:
            percentil = "P5 - P10"
        elif ganancia <= p25:
            percentil = "P10 - P25"
        elif ganancia <= p50:
            percentil = "P25 - P50"
        else:
            percentil = "> P50"
        velocidad = (
            _texto_velocidad_ecrn(
                sexo, dias, int(peso_nacer), int(peso_actual), incluir_detalle=True)
            if dias > 0 else None
        )
        n_muestra = _n_muestra_ganancia(sexo, intervalo, peso_nacer)
        muestra_txt = f" | n={n_muestra}" if n_muestra is not None else ""
        self._valoracion_ganancia = (
            f"Sexo: {sexo} | Intervalo: {intervalo} días | Rango Peso Nacer: {columna}{muestra_txt}\n"
            f"Ganancia real: {int(ganancia)} g | Percentil: {percentil}\n"
            f"Referencia P5: {int(p5)} g | P10: {int(p10)} g | P25: {int(p25)} g | P50: {int(p50)} g"
        )
        self._visita = {
            "sexo": sexo,
            "peso_nacer_g": int(peso_nacer),
            "peso_actual_g": int(peso_actual),
            "ganancia_g": int(ganancia),
            "percentil": percentil,
            "intervalo": intervalo,
            "dias": dias,
            "velocidad_g_dia": velocidad["velocidad_g_dia"] if velocidad else None,
            "velocidad_percentil": velocidad["percentil"] if velocidad else "",
        }
        self.ganancia_resultado.config(
            text=f"Días de vida: {dias} | Ganancia: {int(ganancia)} g\n"
                 f"Percentil estimado: {percentil}\n"
                 f"Referencias ({columna}): P5 {int(p5)} | P10 {int(p10)} | P25 {int(p25)} | P50 {int(p50)} g"
        )
        if velocidad:
            self._valoracion_velocidad = velocidad["texto"]
            self.velocidad_neonatal_resultado.config(text=velocidad["texto"])
        elif dias == 0:
            self.velocidad_neonatal_resultado.config(
                text="Velocidad de crecimiento: se requiere al menos 1 día de seguimiento.")
        else:
            self.velocidad_neonatal_resultado.config(
                text=f"Velocidad de crecimiento: sin referencia CSV para {sexo}, {dias} días y ese peso al nacer.")

    def _peso_eval_actual(self):
        texto = self.eval_peso_actual.get().strip()
        if not texto:
            texto = self.clas_peso_actual.get().strip()
        if not texto:
            return None
        peso = float(texto.replace(",", "."))
        return int(peso * 1000) if peso < 10 else int(peso)

    def _pedir_peso(self, titulo, mensaje):
        """Pide un peso válido (> 0) por diálogo. Devuelve el valor en kg o None si se cancela."""
        while True:
            valor = simpledialog.askstring(titulo, mensaje, parent=self)
            if valor is None:
                return None
            valor = valor.strip().replace(",", ".")
            if not valor:
                return None
            try:
                numero = float(valor)
            except ValueError:
                messagebox.showwarning(titulo, "Ingrese un número válido (ej.: 3.2 o 3200)")
                continue
            if numero <= 0:
                messagebox.showwarning(titulo, "El peso debe ser mayor que 0")
                continue
            return numero

    def _pedir_peso_nacer(self):
        """Pide el peso al nacer y lo guarda en el campo correspondiente."""
        numero = self._pedir_peso(
            "Velocidad de crecimiento",
            "Falta el peso al nacer.\nIngréselo (kg, ej.: 3.2):",
        )
        if numero is None:
            return None
        texto = f"{numero:.3f}".rstrip("0").rstrip(".")
        if numero < 10:
            gramos = int(numero * 1000)
        else:
            gramos = int(numero)
        self.ant["peso_nacer"].set(texto if numero < 10 else str(gramos))
        self._actualizar_clas_peso_g()
        return gramos

    def _pedir_peso_actual(self):
        """Pide el peso actual (g o kg) y lo guarda en el campo de evaluación."""
        numero = self._pedir_peso(
            "Velocidad de crecimiento",
            "Falta el peso actual del paciente.\nIngréselo (kg o g, ej.: 4.2 o 4200):",
        )
        if numero is None:
            return None
        gramos = int(numero * 1000) if numero < 10 else int(numero)
        self.eval_peso_actual.set(str(gramos))
        return gramos

    def _auto_evaluar_z(self, *args):
        """Calcula y muestra el z-score peso/edad (WHO Anthro) automáticamente
        cuando los datos están completos. Silencioso: no muestra avisos."""
        try:
            if not (ANTHRO_OK or PIGROWUP_OK):
                return
            sexo = self.sexo_nacer.get()
            dias = self._parsedias_desde_nacimiento()
            peso = self._peso_eval_actual()
            if not sexo or dias is None or peso is None:
                return
            texto = _texto_zscore_oms(
                sexo, dias, peso,
                postura=self.eval_postura.get(),
                edema=self.eval_edema.get() == "Sí",
            )
            if texto:
                self._valoracion_zscore = texto
            self._auto_evaluar_anthro()
        except Exception:
            pass

    def _auto_evaluar_anthro(self, *args):
        """Calcula y muestra la evaluación completa WHO Anthro automáticamente
        cuando los datos están completos. Silencioso."""
        try:
            if not (ANTHRO_OK or PIGROWUP_OK):
                return
            sexo = self.sexo_nacer.get()
            dias = self._parsedias_desde_nacimiento()
            peso = self._peso_eval_actual()
            talla = self._talla_eval_actual()
            pc = self._pc_eval_actual()
            muac = self._muac_eval_actual()
            triceps = self._pliegue_eval_actual("eval_triceps")
            subescapular = self._pliegue_eval_actual("eval_subescapular")
            if not sexo or dias is None:
                return
            texto = _texto_anthro_completo(
                sexo, dias, peso, talla, pc, muac,
                triceps_mm=triceps, subescapular_mm=subescapular,
                postura=self.eval_postura.get(),
                edema=self.eval_edema.get() == "Sí",
            )
            if texto:
                self.anthro_resultado.config(text=texto)
        except Exception:
            pass

    def _evaluar_anthro_manual(self):
        """Calcula la evaluación completa WHO Anthro mostrando los avisos."""
        try:
            if not (ANTHRO_OK or PIGROWUP_OK):
                messagebox.showwarning(
                    "WHO Anthro", "No está instalado pygrowup2. Ejecute: pip install pygrowup2")
                return
            sexo = self.sexo_nacer.get()
            dias = self._parsedias_desde_nacimiento()
            if not sexo:
                messagebox.showwarning("WHO Anthro", "Seleccione el sexo al nacer")
                return
            if dias is None:
                messagebox.showwarning(
                    "WHO Anthro", "Ingrese la fecha de nacimiento y la fecha actual")
                return
            if dias < 0 or dias > 1826:
                messagebox.showwarning(
                    "WHO Anthro", "La edad debe estar entre 0 y 5 años (1826 días)")
                return
            triceps = self._pliegue_eval_actual("eval_triceps")
            subescapular = self._pliegue_eval_actual("eval_subescapular")
            if (triceps is not None or subescapular is not None) and not PIGROWUP_OK:
                messagebox.showwarning(
                    "WHO Anthro", "Para evaluar tríceps y subescapular debe estar instalado pygrowup2")
                return
            texto = _texto_anthro_completo(
                sexo, dias,
                self._peso_eval_actual(),
                self._talla_eval_actual(),
                self._pc_eval_actual(),
                self._muac_eval_actual(),
                triceps_mm=triceps,
                subescapular_mm=subescapular,
            )
            if texto is None:
                messagebox.showwarning(
                    "WHO Anthro",
                    "No se pudo calcular ningún indicador. Complete peso, talla, perímetro cefálico, MUAC o los pliegues.",
                )
                return
            self.anthro_resultado.config(text=texto)
        except Exception:
            pass

    def _talla_eval_actual(self):
        texto = self.eval_talla.get().strip().replace(",", ".")
        if not texto:
            return None
        try:
            return float(texto)
        except ValueError:
            return None

    def _pc_eval_actual(self):
        texto = self.eval_pc.get().strip().replace(",", ".")
        if not texto:
            return None
        try:
            return float(texto)
        except ValueError:
            return None

    def _muac_eval_actual(self):
        texto = self.eval_muac.get().strip().replace(",", ".")
        if not texto:
            return None
        try:
            return float(texto)
        except ValueError:
            return None

    def _pliegue_eval_actual(self, nombre_variable):
        variable = getattr(self, nombre_variable)
        texto = variable.get().strip().replace(",", ".")
        if not texto:
            return None
        try:
            valor = float(texto)
        except ValueError:
            return None
        return valor if math.isfinite(valor) and valor > 0 else None

    def _ind_vel(self):
        return NOMBRES_IND_VEL_REV.get(self.vel_indicador.get(), "peso")

    def _actualizar_unidad_edad_velocidad(self, *_):
        dias = self._parsedias_desde_nacimiento()
        unidad = "semanas" if dias is not None and 0 <= dias <= 60 else "meses"
        self._unidad_edad_velocidad = unidad
        self.vel_edad1_label.config(text=f"Edad 1 ({unidad}):")
        self.vel_edad2_label.config(text=f"Edad 2 ({unidad}):")

    def _cambiar_vel_indicador(self, *_):
        ind = self._ind_vel()
        valores = [str(i) for i in INTERVALOS_VEL_OMS[ind]]
        self.vel_intervalo_box["values"] = valores
        if self.vel_intervalo.get() not in valores:
            self.vel_intervalo.set(valores[0])
        self.vel_unidades.config(
            text="Unidades: " + ("gramos (g)" if ind == "peso" else "centímetros (cm)"))

    def _calcular_velocidad_oms(self):
        ind = self._ind_vel()
        sexo = "M" if self.vel_sexo.get() == "Masculino" else "F"
        try:
            t1 = float(self.vel_t1.get().replace(",", "."))
            t2 = float(self.vel_t2.get().replace(",", "."))
            v1 = float(self.vel_v1.get().replace(",", "."))
            v2 = float(self.vel_v2.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning(
                "Velocidad OMS",
                f"Ingrese edades ({self._unidad_edad_velocidad}) y valores numéricos válidos",
            )
            return
        if self._unidad_edad_velocidad == "semanas":
            t1 *= 7.0 / DIAS_POR_MES
            t2 *= 7.0 / DIAS_POR_MES
        if t2 <= t1:
            messagebox.showwarning("Velocidad OMS", "La 2.ª edad debe ser mayor que la 1.ª")
            return
        try:
            intervalo_referencia = int(self.vel_intervalo.get())
        except ValueError:
            messagebox.showwarning("Velocidad OMS", "Seleccione un intervalo OMS válido")
            return
        texto = _texto_velocidad_oms(
            ind, sexo, t1, t2, v1, v2,
            intervalo_referencia=intervalo_referencia,
        )
        if texto is None:
            messagebox.showwarning(
                "Velocidad OMS",
                "No se pudo evaluar (revise el rango 0-24 meses, el sexo o los datos)",
            )
            return
        self._valoracion_velocidad_oms = texto
        self.vel_resultado.config(text=texto)
        self.status.config(text="Velocidad de crecimiento OMS calculada | Paciente N° %s"
                                % (self._paciente_id or "-"))

    def _usar_peso_auto(self):
        peso_nacer = self._peso_nacer_g_actual()
        peso_actual = self._peso_eval_actual()
        dias = self._parsedias_desde_nacimiento()
        if peso_nacer is None or peso_actual is None or dias is None:
            messagebox.showwarning(
                "Velocidad OMS",
                "Hacen falta el peso al nacer y el peso actual (o la edad del paciente)",
            )
            return
        self.vel_indicador.set(NOMBRES_IND_VEL_OMS["peso"])
        self.vel_t1.set("0")
        self.vel_v1.set(str(peso_nacer))
        self._actualizar_unidad_edad_velocidad()
        t2 = dias / 7.0 if self._unidad_edad_velocidad == "semanas" else dias / DIAS_POR_MES
        self.vel_t2.set(f"{t2:.4f}".rstrip("0").rstrip("."))
        self.vel_v2.set(str(peso_actual))
        self._calcular_velocidad_oms()

    def _auto_evaluar_vel_oms(self, *_):
        """Calcula automáticamente la velocidad de peso OMS desde el nacer hasta
        el peso actual (solo 0-24 meses). Silencioso."""
        try:
            if self.vel_indicador.get() != NOMBRES_IND_VEL_OMS["peso"]:
                return
            sexo = self.sexo_nacer.get()
            dias = self._parsedias_desde_nacimiento()
            peso_nacer = self._peso_nacer_g_actual()
            peso_actual = self._peso_eval_actual()
            if not sexo or not peso_nacer or not peso_actual:
                return
            if dias is None or dias <= 0:
                return
            t2 = dias / 30.4375
            if t2 > 24:
                return
            texto = _texto_velocidad_oms(
                "peso", sexo, 0.0, t2, peso_nacer, peso_actual,
                intervalo_referencia=int(self.vel_intervalo.get()),
            )
            if texto:
                self._valoracion_velocidad_oms = texto
                self.vel_resultado.config(text=texto)
        except Exception:
            pass

    def _peso_nacer_g_actual(self):
        """Peso al nacer en gramos desde el campo de Antecedentes Prenatales."""
        texto = self.ant["peso_nacer"].get().replace(",", ".").strip()
        if not texto:
            return None
        try:
            v = float(texto)
            return int(v * 1000) if v < 10 else int(v)
        except ValueError:
            return None

    def _peso_actual_g_guardado(self):
        """Peso actual en gramos: campo de evaluación, último registro o peso al nacer."""
        peso = self._peso_eval_actual()
        if peso is not None:
            return peso
        if not getattr(self, "_paciente_id", None):
            return None
        try:
            conn = conectar()
            fila = conn.execute(
                """SELECT peso_actual_g FROM visita_ganancia_peso
                   WHERE paciente_id = ? ORDER BY fecha_registro DESC, id DESC LIMIT 1""",
                (self._paciente_id,),
            ).fetchone()
            conn.close()
        except Exception:
            fila = None
        if fila and fila[0]:
            return int(fila[0])
        return self._peso_nacer_g_actual()

    def _guardar_visita_ganancia(self):
        if not self._exigir_paciente():
            return
        self._valorar_ganancia()
        if not getattr(self, "_visita", None):
            return
        fecha = self.campos["fecha_actual"].get().strip() or date.today().isoformat()
        v = self._visita
        try:
            conn = conectar()
            consulta_id = getattr(self, "_consulta_id", None)
            if consulta_id:
                conn.execute(
                    """UPDATE visita_ganancia_peso
                       SET fecha_visita=?, sexo=?, peso_nacer_g=?, peso_actual_g=?,
                           ganancia_g=?, percentil=?, intervalo_dias=?, dias_vida=?,
                           velocidad_g_dia=?, velocidad_percentil=?
                       WHERE id=? AND paciente_id=?""",
                    (
                        fecha, v["sexo"], v["peso_nacer_g"], v["peso_actual_g"],
                        v["ganancia_g"], v["percentil"], v["intervalo"], v["dias"],
                        v["velocidad_g_dia"], v["velocidad_percentil"],
                        consulta_id, self._paciente_id,
                    ),
                )
                mensaje = f"Consulta del {fecha} actualizada"
            else:
                conn.execute(
                    """INSERT INTO visita_ganancia_peso
                       (paciente_id, fecha_visita, sexo, peso_nacer_g, peso_actual_g,
                        ganancia_g, percentil, intervalo_dias, dias_vida,
                        velocidad_g_dia, velocidad_percentil)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        self._paciente_id, fecha, v["sexo"], v["peso_nacer_g"],
                        v["peso_actual_g"], v["ganancia_g"], v["percentil"],
                        v["intervalo"], v["dias"], v["velocidad_g_dia"],
                        v["velocidad_percentil"],
                    ),
                )
                mensaje = f"Consulta del {fecha} guardada"
            conn.commit()
            conn.close()
            self.status.config(text=f"Visita del {fecha} guardada | Paciente N° {self._paciente_id}")
            self._refrescar_historial()
            self._consulta_id = None
            self._evaluar_paciente(avisar=False)
            self._cargar_ultimos_datos()
            self._actualizar_visibilidad_ganancia()
            messagebox.showinfo("Visita", mensaje)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _refrescar_historial(self):
        for item in self.visita_tree.get_children():
            self.visita_tree.delete(item)
        self._visitas_ids = {}
        if not getattr(self, "_paciente_id", None):
            return
        try:
            conn = conectar()
            filas = conn.execute(
                """SELECT id, fecha_visita, dias_vida, peso_actual_g, ganancia_g,
                          percentil, velocidad_g_dia, velocidad_percentil
                   FROM visita_ganancia_peso WHERE paciente_id = ? ORDER BY fecha_visita, id""",
                (self._paciente_id,),
            ).fetchall()
            conn.close()
        except Exception:
            return
        for fila in filas:
            velocidad = f"{fila[6]:.1f}" if fila[6] is not None else ""
            item = self.visita_tree.insert(
                "", "end",
                values=(fila[1], fila[2], int(fila[3]), int(fila[4]), fila[5],
                        velocidad, fila[7] or ""),
            )
            self._visitas_ids[item] = fila[0]

    def _cargar_visita(self):
        sel = self.visita_tree.selection()
        if not sel:
            messagebox.showinfo("Visita", "Seleccione una visita del historial")
            return
        valores = self.visita_tree.item(sel[0], "values")
        consulta_id = getattr(self, "_visitas_ids", {}).get(sel[0])
        self._aplicar_visita_valores((consulta_id, *valores))

    def _borrar_visita(self):
        sel = self.visita_tree.selection()
        if not sel:
            messagebox.showinfo("Visita", "Seleccione una visita del historial para borrarla")
            return
        consulta_id = getattr(self, "_visitas_ids", {}).get(sel[0])
        paciente_id = getattr(self, "_paciente_id", None)
        if not consulta_id or not paciente_id:
            messagebox.showerror("Visita", "No se pudo identificar la visita seleccionada")
            return
        valores = self.visita_tree.item(sel[0], "values")
        fecha = valores[0] if valores else ""
        if not messagebox.askyesno(
            "Borrar visita", f"¿Desea borrar la visita del {fecha}? Esta acción no se puede deshacer."
        ):
            return
        try:
            conn = conectar()
            cursor = conn.execute(
                "DELETE FROM visita_ganancia_peso WHERE id = ? AND paciente_id = ?",
                (consulta_id, paciente_id),
            )
            conn.commit()
            conn.close()
            if cursor.rowcount != 1:
                messagebox.showinfo("Visita", "La visita ya no existe en el historial")
                self._refrescar_historial()
                return
            if getattr(self, "_consulta_id", None) == consulta_id:
                self._consulta_id = None
                self._visita = None
            self._refrescar_historial()
            self._actualizar_visibilidad_ganancia()
            self.status.config(text=f"Visita del {fecha} borrada | Paciente N° {paciente_id}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    @staticmethod
    def _columna_peso_nacer(peso_nacer):
        if peso_nacer < 2000:
            return "todos_g"
        if peso_nacer < 2500:
            return "peso_nacer_2000_2500_g"
        if peso_nacer < 3000:
            return "peso_nacer_2500_3000_g"
        if peso_nacer < 3500:
            return "peso_nacer_3000_3500_g"
        if peso_nacer < 4000:
            return "peso_nacer_3500_4000_g"
        return "peso_nacer_4000_mas_g"

    def _es_consulta_sucesiva(self):
        if not getattr(self, "_paciente_id", None):
            return False
        try:
            conn = conectar()
            n_visitas = conn.execute(
                "SELECT COUNT(*) FROM visita_ganancia_peso WHERE paciente_id = ?", (self._paciente_id,)
            ).fetchone()[0]
            n_ant = conn.execute(
                "SELECT COUNT(*) FROM antecedentes WHERE paciente_id = ?", (self._paciente_id,)
            ).fetchone()[0]
            conn.close()
        except Exception:
            return False
        return n_visitas > 0 or n_ant > 0

    def _actualizar_visibilidad_ganancia(self):
        visible = self._es_consulta_sucesiva()
        for w in (
            self.peso_actual_label, self.peso_actual_entry, self.valorar_btn,
            self.ganancia_resultado, self.velocidad_neonatal_resultado,
        ):
            w.grid()
        for w in (
            self.guardar_visita_btn, self.refrescar_historial_btn,
            self.cargar_visita_btn, self.borrar_visita_btn, self.marco_vis,
        ):
            if visible:
                w.grid()
            else:
                w.grid_remove()

    def _exigir_paciente(self):
        if not getattr(self, "_paciente_id", None):
            messagebox.showwarning("Datos", "Primero guarde un paciente en la pestaña 'Paciente'")
            return False
        if not self.campos["nombre"].get().strip():
            messagebox.showwarning("Datos", "Ingrese el nombre del paciente antes de guardar")
            return False
        return True

    # ---------- CANS Score ----------
    def _cans_datos(self):
        """Puntuación elegida por signo, en el orden de CANS_SIGNOS."""
        return [(nombre, int(self.cans_vars[nombre].get())) for nombre, _, _ in CANS_SIGNOS]

    def _cans_total(self):
        return sum(puntos for _, puntos in self._cans_datos())

    def _cans_interpretacion(self, total):
        if total < CANS_PUNTAJE_MALNUTRICION:
            return (
                True,
                f"MALNUTRICIÓN FETAL: CANS score {total} (< {CANS_PUNTAJE_MALNUTRICION})",
            )
        return False, f"Nutrición adecuada: CANS score {total} (>= {CANS_PUNTAJE_MALNUTRICION})"

    def _recalcular_cans(self):
        total = self._cans_total()
        malnutricion, interpretacion = self._cans_interpretacion(total)
        color = "#B00020" if malnutricion else "#1E6E5C"
        self.cans_total_label.config(
            text=f"CANS score: {total} / {CANS_MAXIMO}   |   {interpretacion}",
            foreground=color,
        )
        marcados = [f"{n}: {p}" for n, p in self._cans_datos() if p < max(CANS_PUNTOS)]
        self.cans_detalle_label.config(
            text="Signos con puntuación < 4: " + (", ".join(marcados) if marcados else "ninguno"),
            foreground=color if marcados else "#555555",
        )

    def _cans_observaciones(self):
        return self.cans_observaciones.get("1.0", "end").strip()

    def _indice_ponderal(self):
        """Índice ponderal de Rohrer: IP = peso (g) × 100 / longitud (cm)³.
        Devuelve (ip, peso_g, longitud_cm, error)."""
        def _numero(clave, maximo):
            texto = self.ant.get(clave, tk.StringVar()).get().strip().replace(",", ".")
            if not texto:
                return None
            try:
                valor = float(texto)
            except ValueError:
                return None
            if valor <= 0 or valor > maximo:
                return None
            return valor

        peso = _numero("peso_nacer", 10000)
        if peso is None:
            return None, None, None, "Peso al nacer no válido o vacío (máx. 10000)."
        if peso < 10:
            peso *= 1000
        longitud = _numero("longitud_nacer", 100)
        if longitud is None:
            return None, peso, None, "Longitud al nacer no válida o vacía (máx. 100 cm)."
        return peso * 100.0 / (longitud ** 3), peso, longitud, ""

    def _recalcular_ip(self):
        ip, peso, longitud, error = self._indice_ponderal()
        if error:
            self.ip_label.config(text=f"No se pudo calcular el índice ponderal: {error}", foreground="#8B4513")
            self.ip_nota_label.config(text="")
            self._clasificar_patron_crecimiento()
            return None
        malnutricion = ip < IP_CORTE_MALNUTRICION
        self.ip_label.config(
            text=f"Índice ponderal: {ip:.2f} g/cm³   |   "
                 + ("MALNUTRICIÓN: IP < 2,2" if malnutricion else "Sin malnutrición por IP"),
            foreground="#B00020" if malnutricion else "#1E6E5C",
        )
        self.ip_nota_label.config(
            text=f"Peso {peso:.0f} g, longitud {longitud:.1f} cm  ·  "
                 f"IP 2,3 = p10 e IP 2,2 = p3 de peso; < {IP_CORTE_MALNUTRICION} se considera "
                 f"malnutrición (PEG tipo II, asimétrico)."
        )
        self._clasificar_patron_crecimiento()
        return ip

    def _clasificar_patron_crecimiento(self, *args):
        """Tabla 5 de Caiza et al. (2003): combina el IP (bajo/normal/elevado
        segun la curva por EG) con la clasificacion P/EG (PEG/AEG/GEG) para
        definir el patron de crecimiento intrauterino (A-F/N/X)."""
        if not hasattr(self, "ip_patron_label"):
            return
        try:
            semanas = int(re.match(r"\d+", self.ant["edad_gestacional"].get().strip()).group())
        except (AttributeError, ValueError):
            semanas = None
        codigo_eg = getattr(self, "_codigo_peso_eg", "")
        ip, _peso, _longitud, _error = self._indice_ponderal()
        if semanas is None or not codigo_eg or ip is None:
            self.ip_patron_label.config(
                text="Patrón de crecimiento (Tabla 5 Caiza 2003): requiere la clasificación "
                     "del peso por EG (pestaña '1. Antropometría') y el índice ponderal.",
                foreground="#8B4513")
            return
        referencia = IP_REFERENCIA_EG.get(semanas)
        if referencia is None:
            self.ip_patron_label.config(
                text=f"Patrón de crecimiento (Tabla 5): la curva de IP solo cubre "
                     f"33-42 semanas (recibida: {semanas}).",
                foreground="#8B4513")
            return
        p10, _p50, p90 = referencia
        if ip < p10:
            cat_ip = "Bajo"
        elif ip <= p90:
            cat_ip = "Normal"
        else:
            cat_ip = "Elevado"
        letra, descripcion = TABLA5_PATRONES[(cat_ip, codigo_eg)]
        self.ip_patron_label.config(
            text=f"Patrón de crecimiento (Tabla 5 Caiza 2003): {letra} — {descripcion}\n"
                 f"IP {ip:.2f} = {cat_ip} (ref. {semanas} sem: P10 {p10:.2f} · P90 {p90:.2f})  ·  "
                 f"P/EG: {codigo_eg}",
            foreground="#1E6E5C" if letra == "N" else "#B00020",
        )

    def _limpiar_cans(self):
        for variable in self.cans_vars.values():
            variable.set(max(CANS_PUNTOS))
        self.cans_observaciones.delete("1.0", "end")
        self.cans_fecha.set(date.today().isoformat())
        self._recalcular_cans()
        self.status.config(text="CANS score reiniciado")

    def _guardar_cans(self):
        if not self._exigir_paciente():
            return
        total = self._cans_total()
        malnutricion, interpretacion = self._cans_interpretacion(total)
        datos = self._cans_datos()
        conn = conectar()
        try:
            anterior = conn.execute(
                "SELECT id FROM evaluacion_cans WHERE paciente_id = ? ORDER BY id DESC LIMIT 1",
                (self._paciente_id,),
            ).fetchone()
            scores_json = json.dumps(datos, ensure_ascii=False)
            eg = self.ant.get("edad_gestacional", tk.StringVar()).get().strip()
            ip, peso_nacer, longitud_nacer, _error_ip = self._indice_ponderal()
            valores = (
                self.cans_fecha.get().strip(), eg, total, 1 if malnutricion else 0,
                interpretacion, scores_json, self._cans_observaciones(),
                ip, peso_nacer, longitud_nacer,
            )
            if anterior:
                conn.execute(
                    """UPDATE evaluacion_cans
                       SET fecha_evaluacion=?, edad_gestacional=?, total=?,
                           malnutricion_fetal=?, interpretacion=?, scores_json=?,
                           observaciones=?, indice_ponderal=?, peso_nacer_g=?,
                           longitud_nacer_cm=?
                       WHERE id=?""",
                    (*valores, anterior[0]),
                )
            else:
                conn.execute(
                    """INSERT INTO evaluacion_cans
                       (paciente_id, fecha_evaluacion, edad_gestacional, total,
                        malnutricion_fetal, interpretacion, scores_json, observaciones,
                        indice_ponderal, peso_nacer_g, longitud_nacer_cm)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                    (self._paciente_id, *valores),
                )
            conn.commit()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        finally:
            conn.close()
        self.status.config(
            text=f"CANS score guardado ({total}/{CANS_MAXIMO}) | Paciente N° {self._paciente_id}")

    def _editar_cans(self, silencioso=False):
        if not self._exigir_paciente():
            return
        conn = conectar()
        fila = conn.execute(
            """SELECT fecha_evaluacion, edad_gestacional, total, malnutricion_fetal,
                      interpretacion, scores_json, observaciones
               FROM evaluacion_cans WHERE paciente_id = ? ORDER BY id DESC LIMIT 1""",
            (self._paciente_id,),
        ).fetchone()
        conn.close()
        if fila is None:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay un CANS score guardado para este paciente")
            return
        fecha, eg, total, _malnutricion, interpretacion, scores_json, observaciones = fila
        self.cans_fecha.set(fecha or date.today().isoformat())
        self.cans_observaciones.delete("1.0", "end")
        if observaciones:
            self.cans_observaciones.insert("1.0", observaciones)
        try:
            guardado = {nombre: puntos for nombre, puntos in json.loads(scores_json or "[]")}
        except (TypeError, ValueError):
            guardado = {}
        for nombre, variable in self.cans_vars.items():
            if nombre in guardado:
                variable.set(int(guardado[nombre]))
        self._recalcular_cans()
        self.status.config(
            text=f"Editando CANS score ({interpretacion}) | Paciente N° {self._paciente_id}")

    def _exportar_cans_xls(self):
        if not OPENPYXL_OK:
            messagebox.showwarning(
                "Excel", "El paquete 'openpyxl' no está instalado.\nInstálelo con:  pip install openpyxl")
            return
        total = self._cans_total()
        malnutricion, interpretacion = self._cans_interpretacion(total)
        sugerencia = f"CANS_{date.today().isoformat()}.xlsx"
        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar CANS score en Excel",
            initialfile=sugerencia,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")],
        )
        if not ruta:
            return
        libro = _XlsxWorkbook()
        hoja = libro.active
        hoja.title = "CANS score"
        titulo = _XlsxFont(bold=True, size=14)
        encabezado = _XlsxFont(bold=True, color="FFFFFF")
        relleno = _XlsxFill("solid", fgColor="1E6E5C")
        centrado = _XlsxAlignment(horizontal="center", vertical="center", wrap_text=True)
        hoja["A1"] = "Clinical Assessment of Nutritional Status (CANS) score"
        hoja["A1"].font = titulo
        hoja["A2"] = "Paciente"
        hoja["B2"] = self.campos["nombre"].get().strip() or "-"
        hoja["A3"] = "Fecha"
        hoja["B3"] = self.cans_fecha.get().strip()
        hoja["A4"] = "Edad gestacional"
        hoja["B4"] = self.ant.get("edad_gestacional", tk.StringVar()).get().strip()
        fila_tabla = 6
        encabezados = ["N°", "Signo", "Técnica", "Puntos", "Descripción del signo seleccionado"]
        for columna, texto in enumerate(encabezados, start=1):
            celda = hoja.cell(row=fila_tabla, column=columna, value=texto)
            celda.font = encabezado
            celda.fill = relleno
            celda.alignment = centrado
        fila = fila_tabla
        for numero, (nombre, tecnica, opciones) in enumerate(CANS_SIGNOS, start=1):
            puntos = int(self.cans_vars[nombre].get())
            fila += 1
            hoja.cell(row=fila, column=1, value=numero)
            hoja.cell(row=fila, column=2, value=nombre)
            hoja.cell(row=fila, column=3, value=tecnica)
            hoja.cell(row=fila, column=4, value=puntos)
            hoja.cell(row=fila, column=5, value=opciones[puntos])
        fila += 1
        hoja.cell(row=fila, column=3, value="TOTAL").font = _XlsxFont(bold=True)
        celda_total = hoja.cell(row=fila, column=4, value=total)
        celda_total.font = _XlsxFont(bold=True)
        hoja.cell(row=fila, column=5, value=interpretacion)
        fila += 2
        hoja.cell(row=fila, column=1, value="Rango posible").font = _XlsxFont(bold=True)
        hoja.cell(row=fila, column=2, value=f"{CANS_MINIMO}-{CANS_MAXIMO} puntos")
        fila += 1
        hoja.cell(row=fila, column=1, value="Punto de corte").font = _XlsxFont(bold=True)
        hoja.cell(
            row=fila, column=2,
            value=f"Malnutrición fetal si CANS score < {CANS_PUNTAJE_MALNUTRICION}",
        )
        fila += 1
        hoja.cell(row=fila, column=1, value="Clasificación").font = _XlsxFont(bold=True)
        hoja.cell(row=fila, column=2, value="MALNUTRICIÓN FETAL" if malnutricion else "NUTRICIÓN ADECUADA")
        fila += 1
        hoja.cell(row=fila, column=1, value="Observaciones").font = _XlsxFont(bold=True)
        hoja.cell(row=fila, column=2, value=self._cans_observaciones())
        fila += 2
        hoja.cell(row=fila, column=1, value="ÍNDICE PONDERAL DE ROHRER").font = titulo
        fila += 1
        ip, peso_nacer, longitud_nacer, error_ip = self._indice_ponderal()
        if error_ip:
            hoja.cell(row=fila, column=1, value=f"No se pudo calcular: {error_ip}")
        else:
            hoja.cell(row=fila, column=1, value="Fórmula").font = _XlsxFont(bold=True)
            hoja.cell(row=fila, column=2, value="IP = peso (g) × 100 / longitud (cm)³")
            fila += 1
            hoja.cell(row=fila, column=1, value="Peso al nacer").font = _XlsxFont(bold=True)
            hoja.cell(row=fila, column=2, value=f"{peso_nacer:.0f} g")
            fila += 1
            hoja.cell(row=fila, column=1, value="Longitud al nacer").font = _XlsxFont(bold=True)
            hoja.cell(row=fila, column=2, value=f"{longitud_nacer:.1f} cm")
            fila += 1
            celda_ip = hoja.cell(row=fila, column=1, value="Índice ponderal")
            celda_ip.font = _XlsxFont(bold=True)
            hoja.cell(row=fila, column=2, value=f"{ip:.2f} g/cm³")
            fila += 1
            hoja.cell(row=fila, column=1, value="Punto de corte").font = _XlsxFont(bold=True)
            hoja.cell(
                row=fila, column=2,
                value=f"Malnutrición si IP < {IP_CORTE_MALNUTRICION} g/cm³ "
                      f"(IP {IP_P10} = p10, IP {IP_P3} = p3)",
            )
            fila += 1
            hoja.cell(row=fila, column=1, value="Clasificación por IP").font = _XlsxFont(bold=True)
            hoja.cell(
                row=fila, column=2,
                value=("MALNUTRICIÓN (PEG tipo II, asimétrico)"
                       if ip < IP_CORTE_MALNUTRICION else "Sin malnutrición por IP"),
            )
        fila += 2
        hoja.cell(
            row=fila, column=1,
            value="Fuente: Metcoff 1994. Martínez-Nadal S et al. "
                  "An Pediatr (Barc). 2016;84(4):218-223, tabla 1.",
        )
        for columna, ancho in enumerate((6, 32, 58, 9, 72), start=1):
            hoja.column_dimensions[_xlsx_col(columna)].width = ancho
        hoja.freeze_panes = hoja.cell(row=fila_tabla + 1, column=1)
        try:
            libro.save(ruta)
        except PermissionError:
            messagebox.showerror(
                "Excel", f"No se pudo escribir en:\n{ruta}\n\nCierre el archivo si está abierto en Excel.")
            return
        except Exception as e:
            messagebox.showerror("Excel", str(e))
            return
        self.status.config(text=f"CANS score exportado a {ruta}")

    def _parse_fecha(self, texto):
        texto = texto.strip()
        if not texto:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                return datetime.strptime(texto, fmt).date()
            except ValueError:
                continue
        return None

    def _calcular_desarrollo_motor(self, hito, *args):
        if not hasattr(self, "motor_edad") or hito not in self.motor_edad:
            return
        texto_fecha = self.motor_logro[hito].get().strip()
        texto_dias = self.motor_edad[hito].get().strip()
        self.motor_clas_label[hito].config(foreground="#555555")
        dias = None
        if texto_dias:
            try:
                if texto_dias.lstrip("0123456789") != "":
                    raise ValueError
                dias = int(texto_dias)
            except ValueError:
                self.motor_pct[hito].set("")
                self.motor_clas[hito].set("Edad (días) no válida")
                return
        if texto_fecha:
            nacimiento = self._parse_fecha(self.campos["fecha_nacimiento"].get())
            logro = self._parse_fecha(texto_fecha)
            if nacimiento is None or logro is None:
                mensaje = "Complete la fecha de nacimiento" if nacimiento is None else "Fecha no válida"
                self.motor_pct[hito].set("")
                self.motor_clas[hito].set(mensaje)
                return
            dias_calc = (logro - nacimiento).days
            if dias_calc < 0:
                self.motor_edad[hito].set("")
                self.motor_pct[hito].set("")
                self.motor_clas[hito].set("Fecha anterior al nacimiento")
                return
            dias = dias_calc
            if self.motor_edad[hito].get() != str(dias):
                self.motor_edad[hito].set(str(dias))
        if dias is None:
            self.motor_edad[hito].set("")
            self.motor_pct[hito].set("")
            self.motor_clas[hito].set("")
            return
        pct = self._percentil_desarrollo_motor(hito, dias)
        if pct is None:
            self.motor_pct[hito].set("")
            self.motor_clas[hito].set("Sin referencia")
            return
        self.motor_pct[hito].set(f"P{pct}" if pct < 100 else ">P99")
        clasificacion = self._clasificar_percentil_motor(pct)
        self.motor_clas[hito].set(clasificacion)
        self.motor_clas_label[hito].config(
            foreground=("#B00020" if pct < 3 or pct > 97 else "#1E6E5C"))

    def _percentil_desarrollo_motor(self, hito, dias):
        conn = conectar()
        puntos = conn.execute(
            "SELECT percentil, edad_dias FROM referencia_desarrollo_motor "
            "WHERE hito = ? ORDER BY percentil",
            (hito,),
        ).fetchall()
        conn.close()
        if not puntos:
            return None
        if dias <= puntos[0][1]:
            return 1
        if dias >= puntos[-1][1]:
            return 99
        for (p1, d1), (p2, d2) in zip(puntos, puntos[1:]):
            if d1 <= dias <= d2:
                if d2 == d1:
                    return p2
                return int(round(p1 + (dias - d1) * (p2 - p1) / (d2 - d1)))
        return None

    def _clasificar_percentil_motor(self, pct):
        if pct < 3:
            return "Adquisición precoz (< P3)"
        if pct <= 97:
            return "Dentro de la ventana P3-P97"
        return "Adquisición tardía (> P97)"

    def _guardar_desarrollo_motor(self):
        if not self._exigir_paciente():
            return
        conn = conectar()
        try:
            conn.execute(
                "DELETE FROM evaluacion_desarrollo_motor WHERE paciente_id = ?",
                (self._paciente_id,),
            )
            guardados = 0
            for hito in DESARROLLO_MOTOR_HITOS:
                texto = self.motor_logro[hito].get().strip()
                dias = self.motor_edad[hito].get().strip()
                if not texto and not dias:
                    continue
                pct = self.motor_pct[hito].get().strip()
                clas = self.motor_clas[hito].get().strip()
                invalida = any(m in clas for m in (
                    "Complete", "no válida", "anterior", "Sin referencia"))
                conn.execute(
                    """INSERT INTO evaluacion_desarrollo_motor
                       (paciente_id, fecha_evaluacion, hito, fecha_logro,
                        edad_dias, percentil, clasificacion)
                       VALUES (?,?,?,?,?,?,?)""",
                    (
                        self._paciente_id, self.motor_fecha.get().strip(), hito, texto,
                        int(dias) if dias.isdigit() else None,
                        int(re.sub(r"\D", "", pct)) if pct else None,
                        "" if invalida else clas,
                    ),
                )
                guardados += 1
            conn.commit()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        finally:
            conn.close()
        self.status.config(
            text=f"Desarrollo motor guardado ({guardados}/{len(DESARROLLO_MOTOR_HITOS)} hitos) "
                 f"| Paciente N° {self._paciente_id}")

    def _editar_desarrollo_motor(self, silencioso=False):
        if not self._exigir_paciente():
            return
        conn = conectar()
        fila = conn.execute(
            "SELECT fecha_evaluacion FROM evaluacion_desarrollo_motor "
            "WHERE paciente_id = ? ORDER BY id DESC LIMIT 1",
            (self._paciente_id,),
        ).fetchone()
        filas = conn.execute(
            "SELECT hito, fecha_logro FROM evaluacion_desarrollo_motor WHERE paciente_id = ?",
            (self._paciente_id,),
        ).fetchall()
        conn.close()
        if not filas:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay valoración de desarrollo motor guardada")
            return
        self.motor_fecha.set(fila[0] or date.today().isoformat())
        for hito, logro in filas:
            if hito in self.motor_logro:
                self.motor_logro[hito].set(logro or "")
        self.status.config(text=f"Editando desarrollo motor | Paciente N° {self._paciente_id}")

    def _limpiar_desarrollo_motor(self):
        for var in self.motor_logro.values():
            var.set("")
        self.motor_fecha.set(date.today().isoformat())
        self.status.config(text="Desarrollo motor reiniciado")

    def _exportar_motor_xls(self):
        if not OPENPYXL_OK:
            messagebox.showwarning(
                "Excel", "El paquete 'openpyxl' no está instalado.\nInstálelo con:  pip install openpyxl")
            return
        sugerencia = f"DesarrolloMotor_{date.today().isoformat()}.xlsx"
        ruta = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar desarrollo motor en Excel",
            initialfile=sugerencia,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")],
        )
        if not ruta:
            return
        titulo = _XlsxFont(bold=True, size=14)
        encabezado = _XlsxFont(bold=True, color="FFFFFF")
        relleno = _XlsxFill("solid", fgColor="0B5C7A")
        centrado = _XlsxAlignment(horizontal="center", vertical="center", wrap_text=True)
        conn = conectar()
        libro = _XlsxWorkbook()
        hoja = libro.active
        hoja.title = "Desarrollo motor"
        hoja["A1"] = "Desarrollo Motor — OMS-MGRS (WHO Motor Development Study, 2006)"
        hoja["A1"].font = titulo
        hoja["A2"] = "Paciente"
        hoja["B2"] = self.campos["nombre"].get().strip() or "-"
        hoja["A3"] = "Fecha"
        hoja["B3"] = self.motor_fecha.get().strip()
        fila_tabla = 5
        encabezados = ["Hito", "Fecha de logro", "Edad (días)", "Percentil", "Clasificación"]
        for columna, texto in enumerate(encabezados, start=1):
            celda = hoja.cell(row=fila_tabla, column=columna, value=texto)
            celda.font = encabezado
            celda.fill = relleno
            celda.alignment = centrado
        for i, hito in enumerate(DESARROLLO_MOTOR_HITOS, start=1):
            fila = fila_tabla + i
            hoja.cell(row=fila, column=1, value=hito)
            hoja.cell(row=fila, column=2, value=self.motor_logro[hito].get().strip())
            hoja.cell(row=fila, column=3, value=self.motor_edad[hito].get().strip())
            hoja.cell(row=fila, column=4, value=self.motor_pct[hito].get().strip())
            hoja.cell(row=fila, column=5, value=self.motor_clas[hito].get().strip())
        fila = fila_tabla + len(DESARROLLO_MOTOR_HITOS) + 1
        hoja.cell(row=fila, column=1,
                  value="Ventana de adquisición: P3-P97. Fuente: Acta Paediatr Suppl. 2006;450:86-95.")
        for columna, ancho in enumerate((34, 16, 12, 10, 36), start=1):
            hoja.column_dimensions[_xlsx_col(columna)].width = ancho
        hoja2 = libro.create_sheet("Percentiles (referencia)")
        hoja2["A1"] = "Percentiles en días, seis hitos motores gruesos (OMS-MGRS 2006)"
        hoja2["A1"].font = titulo
        columnas_pct = ["Hito", "P1", "P3", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "P97", "P99"]
        for columna, texto in enumerate(columnas_pct, start=1):
            celda = hoja2.cell(row=2, column=columna, value=texto)
            celda.font = encabezado
            celda.fill = relleno
            celda.alignment = centrado
        for i, hito in enumerate(DESARROLLO_MOTOR_HITOS, start=1):
            filas = conn.execute(
                "SELECT percentil, edad_dias FROM referencia_desarrollo_motor "
                "WHERE hito = ? ORDER BY percentil",
                (hito,),
            ).fetchall()
            hoja2.cell(row=2 + i, column=1, value=hito)
            for columna, (_p, dias) in enumerate(filas, start=2):
                hoja2.cell(row=2 + i, column=columna, value=dias)
        for columna, ancho in enumerate((34, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8), start=1):
            hoja2.column_dimensions[_xlsx_col(columna)].width = ancho
        conn.close()
        try:
            libro.save(ruta)
        except PermissionError:
            messagebox.showerror(
                "Excel", f"No se pudo escribir en:\n{ruta}\n\nCierre el archivo si está abierto en Excel.")
            return
        except Exception as e:
            messagebox.showerror("Excel", str(e))
            return
        self.status.config(text=f"Desarrollo motor exportado a {ruta}")

    def _datos_paciente(self):
        return (
            self.campos["nombre"].get(), self.campos["servicio"].get(),
            self.campos["direccion"].get(), self.campos["telefono"].get(),
            self.campos["dni"].get(), self.campos["cuenta"].get(),
            self.sexo.get(), self.campos["cama"].get(),
            self.campos["fecha_nacimiento"].get(), self.campos["edad"].get(),
            self.campos["fecha_actual"].get(), self.campos["dx_medico"].get(),
            self.campos["grado_padre"].get(), self.campos["grado_madre"].get(),
        )

    def _proximo_dni_hc(self):
        conn = conectar()
        filas = conn.execute("SELECT dni_hc FROM paciente").fetchall()
        conn.close()
        numeros = []
        for (dni,) in filas:
            if dni:
                digitos = re.sub(r"\D", "", dni)
                if digitos:
                    numeros.append(int(digitos))
        return max(numeros) + 1 if numeros else 1

    def _asignar_dni_si_nuevo(self, event=None):
        nombre = self.campos["nombre"].get().strip()
        if not nombre or self.campos["dni"].get().strip():
            return
        if getattr(self, "_edit_paciente_id", None):
            return
        try:
            conn = conectar()
            existe = conn.execute(
                "SELECT id FROM paciente WHERE nombre = ? COLLATE NOCASE", (nombre,)
            ).fetchone()
            conn.close()
        except Exception:
            return
        if existe:
            return
        self.campos["dni"].set(str(self._proximo_dni_hc()))

    def _programar_dni_nuevo(self):
        if getattr(self, "_dnijob", None):
            self.after_cancel(self._dnijob)
        self._dnijob = self.after(600, self._asignar_dni_si_nuevo)

    def _guardar_paciente(self):
        if not self.campos["nombre"].get().strip():
            messagebox.showwarning("Datos", "Ingrese el nombre del paciente antes de guardar")
            return
        try:
            self._asignar_dni_si_nuevo()
            conn = conectar()
            cur = conn.cursor()
            datos = self._datos_paciente()
            if getattr(self, "_edit_paciente_id", None):
                cur.execute(
                    """UPDATE paciente SET nombre=?, servicio=?, direccion=?, telefono=?,
                       dni_hc=?, cuenta=?, sexo=?, cama=?, fecha_nacimiento=?, edad=?,
                       fecha_actual=?, dx_medico=?, grado_instruccion_padre=?,
                       grado_instruccion_madre=? WHERE id=?""",
                    (*datos, self._edit_paciente_id),
                )
                pid = self._edit_paciente_id
            else:
                cur.execute(
                    """INSERT INTO paciente
                       (nombre, servicio, direccion, telefono, dni_hc, cuenta, sexo, cama, fecha_nacimiento,
                        edad, fecha_actual, dx_medico, grado_instruccion_padre, grado_instruccion_madre)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    datos,
                )
                pid = cur.lastrowid
            conn.commit()
            conn.close()
            self._paciente_id = pid
            self._edit_paciente_id = pid
            self.pac_status.config(text=f"Paciente N° {pid}")
            self.status.config(text=f"Paciente N° {pid} guardado")
            messagebox.showinfo("Paciente", f"Paciente N° {pid} guardado")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _cargar_paciente_num(self, num, cargar_consulta=True):
        conn = conectar()
        fila = conn.execute("SELECT * FROM paciente WHERE id = ?", (num,)).fetchone()
        conn.close()
        if fila is None:
            return False
        self.campos["nombre"].set(fila[1] or "")
        self.campos["servicio"].set(fila[2] or "")
        self.campos["direccion"].set(fila[3] or "")
        self.campos["telefono"].set(fila[4] or "")
        self.campos["dni"].set(fila[5] or "")
        self.campos["cuenta"].set(fila[6] or "")
        self.sexo.set(fila[7] or "Femenino")
        self.sexo_nacer.set(fila[7] or "Femenino")
        self.campos["cama"].set(fila[8] or "")
        self.campos["fecha_nacimiento"].set(fila[9] or "")
        self.campos["edad"].set(fila[10] or "")
        self.campos["fecha_actual"].set(fila[11] or date.today().isoformat())
        self.campos["dx_medico"].set(fila[12] or "")
        self.campos["grado_padre"].set(fila[13] or "")
        self.campos["grado_madre"].set(fila[14] or "")
        self._paciente_id = num
        self._edit_paciente_id = num
        self.pac_status.config(text=f"Editando paciente N° {num}")
        self.status.config(text=f"Editando paciente N° {num}")
        self._refrescar_historial()
        self._evaluar_paciente()
        self._cargar_ultimos_datos()
        if cargar_consulta:
            self._preguntar_consulta()
        else:
            self._cargar_ultima_consulta()
        self._actualizar_visibilidad_ganancia()
        return True

    def _nueva_consulta(self):
        self.campos["fecha_actual"].set(date.today().isoformat())
        self.clas_peso_actual.set("")
        self.eval_peso_actual.set("")
        self.eval_triceps.set("")
        self.eval_subescapular.set("")
        self.ganancia_resultado.config(text="")
        self.velocidad_neonatal_resultado.config(text="")
        self._valoracion_velocidad = ""
        self.clas_resultado.config(text="")
        self._visita = None
        self._consulta_id = None
        self._actualizar_visibilidad_ganancia()
        self.status.config(text=f"Nueva consulta | Paciente N° {self._paciente_id}")

    def _nueva_consulta_seguimiento(self):
        if not getattr(self, "_paciente_id", None):
            messagebox.showwarning("Datos", "Primero registre o cargue un paciente para la consulta de seguimiento")
            return
        self._nueva_consulta()
        self._evaluar_paciente(avisar=False)
        self._refrescar_evaluaciones_display()
        self._ver_evaluaciones()
        self.status.config(
            text=f"Consulta de seguimiento | Datos del paciente N° {self._paciente_id} conservados"
        )

    def _nueva_consulta_paciente(self):
        self._paciente_id = None
        self._edit_paciente_id = None
        self.pac_status.config(text="")
        for var in self.campos.values():
            if isinstance(var, tk.StringVar):
                var.set("")
        self.campos["fecha_actual"].set(date.today().isoformat())
        for var in self.ant.values():
            if isinstance(var, tk.StringVar):
                var.set("")
            elif isinstance(var, tk.Text):
                var.delete("1.0", "end")
        self.sexo_nacer.set(self.sexo.get())
        self.clas_peso_g.set("")
        self.clas_resultado.config(text="")
        self.clas_peso_actual.set("")
        self.eval_peso_actual.set("")
        self.eval_triceps.set("")
        self.eval_subescapular.set("")
        self.ganancia_resultado.config(text="")
        self.velocidad_neonatal_resultado.config(text="")

        self.clas_peso_ok = False
        self._clasificacion = ""
        self._valoracion_ganancia = ""
        self._valoracion_peso_edad = ""
        self._valoracion_zscore = ""
        self._valoracion_velocidad = ""
        self._visita = None
        self._consulta_id = None
        for item in self.visita_tree.get_children():
            self.visita_tree.delete(item)
        self._limpiar_farmaco_cards()
        for var in self.bio_resultado.values():
            var.set("")
        if hasattr(self, "bio_edad_sel"):
            self.bio_edad_sel.set("Automática según edad")
            self._actualizar_edad_bioquimica()
        self._limpiar_energia()
        self._limpiar_recuento()
        self._actualizar_visibilidad_ganancia()
        self.status.config(text="Nueva consulta | Paciente nuevo")

    def _cargar_ultima_consulta(self):
        hoy = date.today().isoformat()
        try:
            conn = conectar()
            fila = conn.execute(
                """SELECT id, fecha_visita, dias_vida, peso_actual_g, ganancia_g, percentil
                   FROM visita_ganancia_peso WHERE paciente_id = ?
                   ORDER BY fecha_registro DESC, id DESC LIMIT 1""",
                (self._paciente_id,),
            ).fetchone()
            conn.close()
        except Exception:
            fila = None
        if fila and fila[1] == hoy:
            self._aplicar_visita_valores(fila)
        else:
            self._nueva_consulta()

    def _preguntar_consulta(self):
        dlg = tk.Toplevel(self)
        dlg.title("Paciente cargado")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        nombre = self.campos["nombre"].get().strip()
        tk.Label(dlg, text=f"Paciente: {nombre}", font=("Arial", 11, "bold")).pack(padx=20, pady=(15, 5))
        tk.Label(dlg, text="¿Qué desea hacer?").pack(padx=20, pady=(0, 10))

        def elegir(accion):
            dlg.destroy()
            if accion == "edicion":
                self._cargar_ultima_consulta()
            else:
                self._nueva_consulta()

        ttk.Button(dlg, text="Edición", width=26, command=lambda: elegir("edicion")).pack(padx=20, pady=3)
        ttk.Button(dlg, text="Consulta de seguimiento", width=26, command=lambda: elegir("seguimiento")).pack(padx=20, pady=(3, 15))
        dlg.update_idletasks()
        dlg.geometry(f"+{self.winfo_rootx() + 60}+{self.winfo_rooty() + 80}")

    def _aplicar_visita_valores(self, valores):
        if len(valores) == 8:
            (consulta_id, fecha, dias, peso_actual, ganancia, percentil,
             velocidad, percentil_velocidad) = valores
        elif len(valores) == 6:
            consulta_id, fecha, dias, peso_actual, ganancia, percentil = valores
            velocidad, percentil_velocidad = None, ""
        else:
            consulta_id = None
            fecha, dias, peso_actual, ganancia, percentil = valores
            velocidad, percentil_velocidad = None, ""
        if isinstance(peso_actual, float) and peso_actual.is_integer():
            peso_actual = int(peso_actual)
        if isinstance(ganancia, float) and ganancia.is_integer():
            ganancia = int(ganancia)
        self._consulta_id = consulta_id
        self.clas_peso_actual.set(str(peso_actual))
        self.eval_peso_actual.set(str(peso_actual))
        self.campos["fecha_actual"].set(str(fecha))
        self.ganancia_resultado.config(
            text=f"Consulta {fecha}: {dias} días, peso {peso_actual} g, "
                 f"ganancia {ganancia} g, percentil {percentil}"
        )
        if velocidad not in (None, ""):
            self.velocidad_neonatal_resultado.config(
                text=f"Velocidad promedio desde nacimiento: {float(velocidad):.1f} g/día"
                     f" --> {percentil_velocidad or 'sin clasificación guardada'}")
        else:
            self.velocidad_neonatal_resultado.config(text="")
        self._visita = None
        self.status.config(text=f"Editando consulta del {fecha} | Paciente N° {self._paciente_id}")

    def _mostrar_sugerencias(self, campo):
        texto = self.campos[campo].get().strip()
        if not texto:
            self._ocultar_sugerencias()
            return
        col = "nombre" if campo == "nombre" else "dni_hc"
        try:
            conn = conectar()
            filas = conn.execute(
                f"SELECT id, nombre, dni_hc FROM paciente WHERE {col} LIKE ? ORDER BY id DESC LIMIT 8",
                (f"%{texto}%",),
            ).fetchall()
            conn.close()
        except Exception:
            return
        vistos = set()
        filas_unicas = []
        for fila in filas:
            clave = fila[1].strip().lower() if fila[1] else ""
            if clave and clave not in vistos:
                vistos.add(clave)
                filas_unicas.append(fila)
        filas = filas_unicas
        if not filas:
            self._ocultar_sugerencias()
            return
        if not getattr(self, "_ac_dlg", None) or not self._ac_dlg.winfo_exists():
            self._ac_dlg = tk.Toplevel(self)
            self._ac_dlg.overrideredirect(True)
            self._ac_dlg.attributes("-topmost", True)
            self._ac_lista = tk.Listbox(self._ac_dlg, height=6, width=40, font=("Arial", 10))
            self._ac_lista.pack(fill="both", expand=True)
            self._ac_lista.bind("<Button-1>", self._elegir_sugerencia)
            self._ac_lista.bind("<Return>", self._elegir_sugerencia)
            self._ac_dlg.bind("<Escape>", lambda e: self._ocultar_sugerencias())
        self._ac_lista.delete(0, "end")
        self._ac_datos = filas
        for fila in filas:
            self._ac_lista.insert("end", f"{fila[1]}  [{fila[2] or 'sin DNI'}]")
        self._ac_lista.selection_clear(0, "end")
        entry = self.entradas[campo]
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height() + 2
        self._ac_dlg.geometry(f"+{int(x)}+{int(y)}")
        self._ac_dlg.deiconify()
        self._ac_dlg.lift()

    def _ocultar_sugerencias(self):
        if getattr(self, "_ac_dlg", None) and self._ac_dlg.winfo_exists():
            self._ac_dlg.withdraw()

    def _elegir_sugerencia(self, event=None):
        index = -1
        if event is not None and event.widget is self._ac_lista:
            tipo = str(getattr(event, "type", ""))
            if tipo in ("4", "ButtonPress", "5", "ButtonRelease"):
                try:
                    index = self._ac_lista.nearest(int(event.y))
                except (TypeError, ValueError):
                    index = -1
        if index < 0:
            sel = self._ac_lista.curselection()
            index = sel[0] if sel else -1
        if not (0 <= index < len(self._ac_datos)):
            return "break"
        fila = self._ac_datos[index]
        self._cargar_paciente_num(fila[0])
        self._ocultar_sugerencias()
        return "break"

    def _buscar_paciente(self, campo):
        if campo == "nombre":
            texto = self.campos["nombre"].get().strip()
        else:
            texto = self.campos["dni"].get().strip()
        if not texto:
            messagebox.showinfo("Buscar", "Escriba en el campo Nombre o DNI lo que desea buscar")
            return
        col = "nombre" if campo == "nombre" else "dni_hc"
        conn = conectar()
        filas = conn.execute(
            f"SELECT id, nombre, dni_hc, fecha_nacimiento FROM paciente WHERE {col} LIKE ? ORDER BY id DESC LIMIT 50",
            (f"%{texto}%",),
        ).fetchall()
        conn.close()
        if not filas:
            messagebox.showinfo("Buscar", f"No se encontró ningún paciente con '{texto}'")
            return
        dlg = tk.Toplevel(self)
        dlg.title("Resultados de búsqueda")
        dlg.geometry("620x320")
        dlg.transient(self)
        dlg.grab_set()
        tree = ttk.Treeview(dlg, columns=["id", "Nombre", "DNI/HC", "F. Nacimiento"], show="headings")
        for columna, ancho in (("id", 60), ("Nombre", 240), ("DNI/HC", 120), ("F. Nacimiento", 120)):
            tree.heading(columna, text=columna)
            tree.column(columna, width=ancho)
        for fila in filas:
            tree.insert("", "end", values=(fila[0], fila[1], fila[2], fila[3]))
        tree.pack(fill="both", expand=True, padx=6, pady=6)

        def seleccionar():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Buscar", "Seleccione un paciente de la lista")
                return
            pid = int(tree.item(sel[0], "values")[0])
            self._cargar_paciente_num(pid)
            dlg.destroy()

        def al_doble(event):
            seleccionar()

        tree.bind("<Double-1>", al_doble)
        ttk.Button(dlg, text="Seleccionar", command=seleccionar).pack(pady=5)

    def _guardar_antecedentes(self):
        if not self._exigir_paciente():
            return
        try:
            conn = conectar()
            cur = conn.cursor()
            existente = cur.execute(
                """SELECT id, clasificacion_peso_nacer, valoracion_ganancia_peso,
                          valoracion_peso_edad, valoracion_velocidad_crecimiento,
                          valoracion_zscore_oms, valoracion_velocidad_oms
                   FROM antecedentes WHERE paciente_id = ?
                   ORDER BY id DESC LIMIT 1""",
                (self._paciente_id,),
            ).fetchone()
            clas = getattr(self, "_clasificacion", "")
            if not clas and existente:
                clas = existente[1] or ""
            vgan = getattr(self, "_valoracion_ganancia", "")
            if not vgan and existente:
                vgan = existente[2] or ""
            vpe = getattr(self, "_valoracion_peso_edad", "")
            if not vpe and existente:
                vpe = existente[3] or ""
            vve = getattr(self, "_valoracion_velocidad", "")
            if not vve and existente:
                vve = existente[4] or ""
            vzs = getattr(self, "_valoracion_zscore", "")
            if not vzs and existente:
                vzs = existente[5] or ""
            vvo = getattr(self, "_valoracion_velocidad_oms", "")
            if not vvo and existente:
                vvo = existente[6] or ""
            campos = (
                self.ant["edad_gestacional"].get(), self.ant["parto"].get(),
                self.ant["peso_nacer"].get(), self.ant["perimetro_cefalico_nacer"].get(),
                self.ant["longitud_nacer"].get(),
                self.ant["familiares"].get("1.0", "end").strip(),
                clas, vgan, vpe, vve, vzs, vvo,
            )
            if existente:
                cur.execute(
                    """UPDATE antecedentes
                       SET edad_gestacional=?, parto=?, peso_nacer=?,
                           perimetro_cefalico_nacer=?, longitud_nacer=?, antecedentes_familiares=?,
                           clasificacion_peso_nacer=?, valoracion_ganancia_peso=?,
                           valoracion_peso_edad=?, valoracion_velocidad_crecimiento=?,
                           valoracion_zscore_oms=?, valoracion_velocidad_oms=?
                       WHERE id=?""",
                    (*campos, existente[0]),
                )
            else:
                cur.execute(
                    """INSERT INTO antecedentes
                       (paciente_id, edad_gestacional, parto, peso_nacer,
                        perimetro_cefalico_nacer, longitud_nacer, antecedentes_familiares,
                        clasificacion_peso_nacer, valoracion_ganancia_peso,
                        valoracion_peso_edad, valoracion_velocidad_crecimiento,
                        valoracion_zscore_oms, valoracion_velocidad_oms)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (self._paciente_id, *campos),
                )
            conn.commit()
            conn.close()
            self.status.config(text=f"Antecedentes guardados | Paciente N° {self._paciente_id}")
            self._evaluar_paciente(avisar=False)
            self._cargar_ultimos_datos()
            self._actualizar_visibilidad_ganancia()
            messagebox.showinfo("Guardado", "Antecedentes guardados")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _editar_antecedentes(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        self._evaluar_paciente(avisar=False)
        conn = conectar()
        fila = conn.execute(
            "SELECT * FROM antecedentes WHERE paciente_id = ? ORDER BY id DESC LIMIT 1",
            (self._paciente_id,),
        ).fetchone()
        conn.close()
        if fila is None:
            if silencioso:
                for clave in ("edad_gestacional", "parto", "peso_nacer",
                              "perimetro_cefalico_nacer", "longitud_nacer"):
                    self.ant[clave].set("")
                self.ant["familiares"].delete("1.0", "end")
                self.clas_peso_g.set("")
                self.clas_resultado.config(text="")
                self.ganancia_resultado.config(text="")
                self.velocidad_neonatal_resultado.config(text="")

                self._clasificacion = ""
                self._valoracion_ganancia = ""
                self._valoracion_peso_edad = ""
                self._valoracion_velocidad = ""
                self._valoracion_zscore = ""
                return
            messagebox.showinfo("Editar", "No hay antecedentes guardados para este paciente")
            return
        self.ant["edad_gestacional"].set(fila[2] or "")
        self.ant["parto"].set(fila[3] or "")
        self.ant["peso_nacer"].set(fila[4] or "")
        self.ant["perimetro_cefalico_nacer"].set(fila[5] or "")
        self.ant["longitud_nacer"].set(fila[6] or "")
        self.ant["familiares"].delete("1.0", "end")
        self.ant["familiares"].insert("1.0", fila[7] or "")
        self._clasificacion = fila[8] or ""
        self._valoracion_ganancia = fila[9] or ""
        self._valoracion_peso_edad = fila[10] or ""
        self._valoracion_velocidad = fila[11] or ""
        self._valoracion_zscore = fila[12] or ""
        self.clas_resultado.config(text=self._clasificacion)
        self.ganancia_resultado.config(text=self._valoracion_ganancia)
        self.velocidad_neonatal_resultado.config(text=self._valoracion_velocidad)

        self.status.config(text=f"Editando antecedentes | Paciente N° {self._paciente_id}")
        self._refrescar_evaluaciones_display()
        self._refrescar_historial()
        if not silencioso:
            self._ver_evaluaciones()

    def _ver_evaluaciones(self):
        """Selecciona la pestaña de Antecedentes y desplaza a las evaluaciones OMS."""
        self.notebook.select(self._tab_ant)
        if getattr(self, "_canvas_ant", None):
            self._canvas_ant.yview_moveto(1.0)

    def _refrescar_evaluaciones_display(self):
        """Vuelca las valoraciones OMS guardadas del paciente actual a la interfaz."""
        if not getattr(self, "_paciente_id", None):
            return
        conn = conectar()
        fila = conn.execute(
            """SELECT clasificacion_peso_nacer, valoracion_ganancia_peso,
                      valoracion_peso_edad, valoracion_velocidad_crecimiento,
                      valoracion_zscore_oms, valoracion_velocidad_oms
               FROM antecedentes WHERE paciente_id = ?
               ORDER BY id DESC LIMIT 1""",
            (self._paciente_id,),
        ).fetchone()
        conn.close()
        if fila is None:
            return
        self._clasificacion = fila[0] or ""
        self._valoracion_ganancia = fila[1] or ""
        self._valoracion_peso_edad = fila[2] or ""
        self._valoracion_velocidad = fila[3] or ""
        self._valoracion_zscore = fila[4] or ""
        self.clas_resultado.config(text=self._clasificacion)
        self.ganancia_resultado.config(text=self._valoracion_ganancia)
        self.velocidad_neonatal_resultado.config(text=self._valoracion_velocidad)

    def _guardar_farmaco(self):
        if not self._exigir_paciente():
            return
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("DELETE FROM interaccion_farmaco WHERE paciente_id = ?", (self._paciente_id,))
            for c in self.farmaco_cards:
                farmaco = c["farmaco"].get().strip()
                interaccion = c["interaccion"].get("1.0", "end").strip()
                if farmaco or interaccion:
                    cur.execute(
                        "INSERT INTO interaccion_farmaco (paciente_id, farmaco, interaccion) VALUES (?,?,?)",
                        (self._paciente_id, farmaco, interaccion),
                    )
            conn.commit()
            conn.close()
            self.status.config(text=f"Fármaco-nutriente guardado | Paciente N° {self._paciente_id}")
            messagebox.showinfo("Guardado", "Interacción fármaco-nutriente guardada")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _editar_farmaco(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        conn = conectar()
        filas = conn.execute(
            "SELECT farmaco, interaccion, recomendacion FROM interaccion_farmaco WHERE paciente_id = ? ORDER BY id",
            (self._paciente_id,),
        ).fetchall()
        conn.close()
        self._limpiar_farmaco_cards()
        for fila in filas:
            interaccion = (fila[1] or "").strip()
            recomendacion = (fila[2] or "").strip()
            if interaccion and recomendacion:
                interaccion = interaccion + "\n\n" + recomendacion
            elif recomendacion:
                interaccion = recomendacion
            self._agregar_farmaco_card(
                farmaco=fila[0] or "",
                interaccion=interaccion,
            )
        if not filas:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay fármacos guardados para este paciente")
        else:
            self.status.config(text=f"Editando fármaco-nutriente | Paciente N° {self._paciente_id}")

    def _parse_rango_referencia(self, texto):
        """Devuelve (bajo, alto) a partir del texto del rango normal; None si no es numérico."""
        texto = (texto or "").strip()
        if not texto:
            return None
        if "<" in texto:
            m = re.search(r"([\d]+(?:[.,]\d+)?)", texto)
            if not m:
                return None
            return (None, float(m.group(1).replace(",", ".")))
        if ">" in texto:
            m = re.search(r"([\d]+(?:[.,]\d+)?)", texto)
            if not m:
                return None
            return (float(m.group(1).replace(",", ".")), None)
        m = re.search(r"([\d]+(?:[.,]\d+)?)\s*[-–]\s*([\d]+(?:[.,]\d+)?)", texto)
        if not m:
            return None
        return (
            float(m.group(1).replace(",", ".")),
            float(m.group(2).replace(",", ".")),
        )

    def _comparar_bioquimica(self, prueba, *args):
        if not hasattr(self, "bio_estado") or prueba not in self.bio_estado:
            return
        self._actualizar_ref_label(prueba)
        texto_ref, _edad_uso = self._referencia_bioq(prueba)
        rango = self._parse_rango_referencia(texto_ref)
        texto = self.bio_resultado[prueba].get().strip()
        lbl = self.bio_estado_label[prueba]
        if not texto:
            self.bio_estado[prueba].set("—")
            lbl.config(foreground="#555555")
            return
        try:
            valor = float(texto.replace(",", "."))
        except ValueError:
            self.bio_estado[prueba].set("Valor no numérico")
            lbl.config(foreground="#B00020")
            return
        if rango is None:
            self.bio_estado[prueba].set("—")
            lbl.config(foreground="#555555")
            return
        bajo, alto = rango
        if bajo is not None and valor < bajo:
            estado = "Bajo"
        elif alto is not None and valor > alto:
            estado = "Alto"
        else:
            estado = "Normal"
        self.bio_estado[prueba].set(estado)
        lbl.config(foreground=("#1E6E5C" if estado == "Normal" else "#B00020"))

    def _guardar_bioquimica(self):
        if not self._exigir_paciente():
            return
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("DELETE FROM evaluacion_bioquimica WHERE paciente_id = ?", (self._paciente_id,))
            for prueba in self.bioq_pruebas:
                resultado = self.bio_resultado[prueba].get().strip()
                texto_ref, edad_uso = self._referencia_bioq(prueba)
                cur.execute(
                    "INSERT INTO evaluacion_bioquimica "
                    "(paciente_id, prueba, valor_normal, resultado, comparacion, unidad, referencia_edad) "
                    "VALUES (?,?,?,?,?,?,?)",
                    (self._paciente_id, prueba, texto_ref, resultado,
                     self.bio_estado[prueba].get().strip() if resultado else "",
                     self.bioq_ref[prueba]["unidad"], edad_uso),
                )
            conn.commit()
            conn.close()
            self.status.config(text=f"Evaluación bioquímica guardada | Paciente N° {self._paciente_id}")
            messagebox.showinfo("Guardado", "Evaluación bioquímica guardada")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _limpiar_bioquimica(self):
        for var in self.bio_resultado.values():
            var.set("")
        self.bio_edad_sel.set("Automática según edad")
        self._actualizar_edad_bioquimica()
        self.status.config(text="Evaluación bioquímica reiniciada")

    def _editar_bioquimica(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        conn = conectar()
        filas = {}
        for fila in conn.execute(
            "SELECT prueba, valor_normal, resultado, comparacion, unidad, referencia_edad "
            "FROM evaluacion_bioquimica WHERE paciente_id = ? ORDER BY id",
            (self._paciente_id,),
        ).fetchall():
            filas[fila[0]] = (fila[1] or "", fila[2] or "", fila[3] or "", fila[4] or "", fila[5] or "")
        conn.close()
        cargadas = False
        for prueba in self.bioq_pruebas:
            if prueba in filas and filas[prueba][1]:
                self.bio_resultado[prueba].set(filas[prueba][1])
                cargadas = True
            else:
                self.bio_resultado[prueba].set("")
        if not cargadas:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay evaluaciones bioquímicas guardadas para este paciente")
        else:
            self.status.config(text=f"Editando evaluación bioquímica | Paciente N° {self._paciente_id}")

    def _cargar_ultimos_datos(self):
        self._editar_antecedentes(silencioso=True)
        self._editar_cans(silencioso=True)
        self._editar_farmaco(silencioso=True)
        self._editar_bioquimica(silencioso=True)
        self._editar_recuento(silencioso=True)
        self.status.config(
            text=f"Datos de antecedentes, signos clínicos, fármacos y bioquímica cargados | Paciente N° {self._paciente_id}"
        )

if __name__ == "__main__":
    App().mainloop()
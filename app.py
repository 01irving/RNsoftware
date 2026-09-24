import csv
import io
import re
import sqlite3
import tkinter as tk
from datetime import date, datetime, timedelta
from decimal import Decimal
from tkinter import ttk, messagebox
from pathlib import Path

try:
    from pygrowup import Observation as _PGObservation, exceptions as _PGEx
    PIGROWUP_OK = True
except ImportError:
    _PGObservation = None
    _PGEx = None
    PIGROWUP_OK = False

DB = Path(__file__).resolve().parent / "hc_nutricional.db"

ORGANOS = [
    "Cabello", "Piel", "Cara", "Uñas", "Ojos", "Labios",
    "Dientes y Encías", "Lengua", "Esqueleto", "Tejido Subcutáneo",
]

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
    ("Magnesio", "1.7-2.55 mg/dl"),
    ("Amilasa", "25-125 U/L"),
    ("Lipasa", "<200 U/L"),
    ("Triglicéridos", "<150 mg/dl"),
    ("Colesterol", "<200 mg/dl"),
    ("PCT", "<0.5 ng/ml"),
    ("Lactato", "0.5-2.2 mmol/L"),
]


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
0-7,50,150,150,150,150,150,150
0-7,25,*,0,0,0,50,0
0-7,10,*,-150,-150,-250,-250,-150
0-7,5,*,-200,-250,-300,-250,-250
7-14,50,275,250,250,250,275,250
7-14,25,*,150,100,100,100,100
7-14,10,*,0,0,0,50,0
7-14,5,*,-100,-50,-100,-100,-50
14-28,50,600,700,650,700,725,650
14-28,25,*,550,550,500,550,550
14-28,10,*,450,450,400,400,450
14-28,5,*,450,350,350,400,350
""",
    "ECRN_ninas.csv": """intervalo_dias,percentil,peso_nacer_2000_2500_g,peso_nacer_2500_3000_g,peso_nacer_3000_3500_g,peso_nacer_3500_4000_g,peso_nacer_4000_mas_g,todos_g
0-7,50,0,150,100,100,150,100
0-7,25,*,0,0,0,0,0
0-7,10,*,-100,-100,-150,-100,-100
0-7,5,*,-150,-200,-250,-200,-200
7-14,50,200,200,200,200,200,200
7-14,25,*,100,100,100,100,100
7-14,10,*,0,0,0,50,0
7-14,5,*,-100,-50,-100,0,-50
14-28,50,500,600,550,550,600,550
14-28,25,*,450,436,450,450,450
14-28,10,*,400,350,300,300,350
14-28,5,*,300,300,250,200,300
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
F,0,nacimiento,2.4,2.8,3.2,3.5,4.2
F,30,1 mes,3.2,3.7,4.2,4.8,5.5
F,61,2 meses,3.9,4.5,5.1,5.7,6.6
F,91,3 meses,4.5,5.1,5.8,6.5,7.4
F,122,4 meses,5.0,5.6,6.4,7.1,8.0
F,152,5 meses,5.4,6.0,6.9,7.6,8.6
F,183,6 meses,5.7,6.3,7.2,7.9,9.0
F,213,7 meses,5.9,6.6,7.5,8.3,9.4
F,244,8 meses,6.1,6.8,7.7,8.6,9.7
F,274,9 meses,6.3,7.0,8.0,8.9,10.0
F,305,10 meses,6.5,7.2,8.2,9.1,10.3
F,335,11 meses,6.7,7.4,8.4,9.3,10.6
F,365,12 meses,6.9,7.7,8.7,9.6,10.9
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


def _texto_velocidad(sexo, dias, peso_nacer_g, peso_actual_g):
    """Texto de la evaluación de velocidad de crecimiento (g/día, OMS 2006). None si no se puede calcular."""
    if not sexo or dias is None or dias <= 0 or dias >= 365:
        return None
    if peso_nacer_g is None or peso_actual_g is None:
        return None
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
        if f[1] <= dias < f[2] and f[3] <= kg_nacer < f[4]:
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
        f"Ventana: {ventana} | Peso nacer: {kg_nacer:.2f} kg | Días: {dias}\n"
        f"Velocidad: {vel:.1f} g/día --> {clas}\n"
        f"Referencias (g/día): {ref}"
    )


def _texto_zscore_oms(sexo, dias, peso_g):
    """Texto del z-score peso-para-la-edad (OMS, pygrowup2). None si no se puede calcular."""
    if not PIGROWUP_OK or not sexo or dias is None or dias < 0 or dias > 1856 or peso_g is None:
        return None
    try:
        peso_kg = peso_g / 1000.0
        base = date(2000, 1, 1)
        obs = _PGObservation(
            sex="male" if sexo == "Masculino" else "female",
            dob=base,
            date_of_observation=base + timedelta(days=dias),
        )
        z = float(obs.wfa(Decimal(str(peso_kg))))
    except Exception:
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
    semanas, resto = divmod(dias, 7)
    return (
        f"Edad: {semanas} semanas y {resto} días | Peso: {peso_kg:.2f} kg\n"
        f"Z-score peso/edad (OMS): {z:.2f} --> {clas}\n"
        f"(Estándar OMS 2006, peso para la edad - tabla LMS por día)"
    )


def _texto_ganancia(sexo, dias, peso_nacer_g, peso_actual_g):
    """Texto de la evaluación de ganancia de peso (ECRN). None si no se puede calcular."""
    if not sexo or dias is None or dias < 0 or dias > 28:
        return None
    if peso_nacer_g is None or peso_actual_g is None:
        return None
    if dias <= 6:
        intervalo = "0-7"
    elif dias <= 13:
        intervalo = "7-14"
    else:
        intervalo = "14-28"
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
    return (
        f"Sexo: {sexo} | Intervalo: {intervalo} días | Rango Peso Nacer: {columna}\n"
        f"Ganancia real: {int(ganancia)} g | Percentil: {percentil}\n"
        f"Referencia P5: {int(p5)} g | P10: {int(p10)} g | P25: {int(p25)} g | P50: {int(p50)} g"
    )


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Historia Clínica Nutricional Pediátrica (1-4)")
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
               WHERE a.valoracion_ganancia_peso IS NULL OR a.valoracion_ganancia_peso = ''
                  OR a.valoracion_peso_edad IS NULL OR a.valoracion_peso_edad = ''
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
        if peso_actual_g is None:
            motivos.append("no hay visitas registradas (falta el peso actual)")
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
        if bloqueado:
            return 0, motivos
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
        peso_act_g = int(peso_actual_g)
        nuevo_vgan = vgan
        nuevo_vpe = vpe
        nuevo_vve = vve
        nuevo_vzs = vzs
        if not vgan and peso_nacer_g is not None:
            t = _texto_ganancia(sexo, dias, peso_nacer_g, peso_act_g)
            if t:
                nuevo_vgan = t
            else:
                motivos.append("ganancia de peso sin calcular (revise edad ≤ 28 días, sexo o referencia)")
        if not vpe:
            t = _texto_peso_edad(sexo, dias, peso_act_g)
            if t:
                nuevo_vpe = t
            else:
                motivos.append("peso-para-la-edad sin calcular (revise edad o referencia)")
        if not vve and peso_nacer_g is not None:
            t = _texto_velocidad(sexo, dias, peso_nacer_g, peso_act_g)
            if t:
                nuevo_vve = t
            else:
                motivos.append("velocidad de crecimiento sin calcular (revise edad o referencia)")
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
        conn.commit()

    def _cargar_referencias(self, conn):
        data_dir = Path(__file__).resolve().parent / "data"
        nombres = ("ninos.csv", "ninas.csv", "clasificaciones.csv", "ECRN_ninos.csv", "ECRN_ninas.csv",
                   "tabla_peso_para_la_edad.csv", "tabla_velocidad_crecimiento.csv")
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
        if conn.execute("SELECT COUNT(*) FROM referencia_ganancia_peso").fetchone()[0] == 0:
            for nombre, sexo in (("ECRN_ninos.csv", "Masculino"), ("ECRN_ninas.csv", "Femenino")):
                contenido = conn.execute(
                    "SELECT contenido FROM archivo_csv WHERE nombre = ?", (nombre,)
                ).fetchone()[0]
                columnas = [
                    "peso_nacer_2000_2500_g",
                    "peso_nacer_2500_3000_g",
                    "peso_nacer_3000_3500_g",
                    "peso_nacer_3500_4000_g",
                    "peso_nacer_4000_mas_g",
                    "todos_g",
                ]
                for fila in csv.DictReader(io.StringIO(contenido)):
                    valores = []
                    for col in columnas:
                        v = fila[col].strip()
                        valores.append(None if v == "*" else float(v))
                    conn.execute(
                        f"""INSERT INTO referencia_ganancia_peso
                            (sexo, intervalo_dias, percentil,
                             peso_nacer_2000_2500_g, peso_nacer_2500_3000_g,
                             peso_nacer_3000_3500_g, peso_nacer_3500_4000_g,
                             peso_nacer_4000_mas_g, todos_g)
                            VALUES (?,?,?,?,?,?,?,?,?)""",
                        (sexo, fila["intervalo_dias"].strip(), int(fila["percentil"]), *valores),
                    )
        if conn.execute("SELECT COUNT(*) FROM referencia_peso_edad").fetchone()[0] == 0:
            contenido = conn.execute(
                "SELECT contenido FROM archivo_csv WHERE nombre = 'tabla_peso_para_la_edad.csv'"
            ).fetchone()[0]
            for fila in csv.DictReader(io.StringIO(contenido)):
                conn.execute(
                    """INSERT INTO referencia_peso_edad
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
        conn.commit()

    def _crear_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        self._tab_paciente()
        self._tab_antecedentes()
        self._tab_signos()
        self._tab_farmaco()
        self._tab_bioquimica()

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

    # ---------- Pestaña 1 Antecedentes ----------
    def _tab_antecedentes(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="1. Antecedentes")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Antecedentes:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_antecedentes).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_antecedentes).pack(side="left", padx=3)

        marco = ttk.LabelFrame(tab, text="A) Antecedentes Prenatales")
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

        marco_familia = ttk.LabelFrame(tab, text="Antecedentes Familiares")
        marco_familia.pack(fill="both", expand=True, padx=8, pady=8)
        self.ant["familiares"] = tk.Text(marco_familia, height=6)
        self.ant["familiares"].pack(fill="both", expand=True, padx=6, pady=6)

        marco_clas = ttk.LabelFrame(tab, text="Clasificación según Edad Gestacional, Peso al Nacer y Sexo")
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

        ttk.Separator(marco_clas, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", padx=5, pady=6)
        self.peso_actual_label = ttk.Label(marco_clas, text="Peso actual (g o kg):")
        self.peso_actual_label.grid(row=5, column=0, sticky="w", padx=5, pady=4)
        self.clas_peso_actual = tk.StringVar()
        self.peso_actual_entry = ttk.Entry(marco_clas, textvariable=self.clas_peso_actual, width=14)
        self.peso_actual_entry.grid(row=5, column=1, sticky="w", padx=5, pady=4)
        self.valorar_btn = ttk.Button(marco_clas, text="Valorar ganancia de peso", command=self._valorar_ganancia)
        self.valorar_btn.grid(row=5, column=2, sticky="w", padx=5, pady=4)
        self.ganancia_resultado = tk.Label(marco_clas, text="", justify="left", foreground="green", wraplength=700)
        self.ganancia_resultado.grid(row=6, column=0, columnspan=4, sticky="w", padx=5, pady=4)

        self.guardar_visita_btn = ttk.Button(marco_clas, text="Guardar visita", command=self._guardar_visita_ganancia)
        self.guardar_visita_btn.grid(row=7, column=0, sticky="w", padx=5, pady=4)
        self.refrescar_historial_btn = ttk.Button(marco_clas, text="Refrescar historial", command=self._refrescar_historial)
        self.refrescar_historial_btn.grid(row=7, column=1, sticky="w", padx=5, pady=4)
        self.cargar_visita_btn = ttk.Button(marco_clas, text="Cargar visita", command=self._cargar_visita)
        self.cargar_visita_btn.grid(row=7, column=2, sticky="w", padx=5, pady=4)

        self.marco_vis = ttk.LabelFrame(marco_clas, text="Historial de visitas / ganancia de peso")
        self.marco_vis.grid(row=8, column=0, columnspan=4, sticky="ew", padx=5, pady=4)
        self.visita_tree = ttk.Treeview(
            self.marco_vis, columns=["Fecha", "Días", "Peso actual (g)", "Ganancia (g)", "Percentil"],
            show="headings", height=5,
        )
        for col in ["Fecha", "Días", "Peso actual (g)", "Ganancia (g)", "Percentil"]:
            self.visita_tree.heading(col, text=col)
            self.visita_tree.column(col, width=120, anchor="center")
        self.visita_tree.pack(side="left", fill="x", expand=True)
        tsb = ttk.Scrollbar(self.marco_vis, orient="vertical", command=self.visita_tree.yview)
        tsb.pack(side="right", fill="y")
        self.visita_tree.configure(yscrollcommand=tsb.set)
        self.clas_peso_ok = False

        ttk.Separator(marco_clas, orient="horizontal").grid(row=9, column=0, columnspan=4, sticky="ew", padx=5, pady=6)
        ttk.Label(marco_clas, text="Evaluación del crecimiento (OMS 2006):").grid(
            row=10, column=0, sticky="w", padx=5, pady=4)
        self.eval_peso_edad_btn = ttk.Button(
            marco_clas, text="1. Peso para la edad",
            command=self._valorar_peso_edad,
        )
        self.eval_peso_edad_btn.grid(row=10, column=1, sticky="w", padx=5, pady=4)
        self.eval_velocidad_btn = ttk.Button(
            marco_clas, text="2. Velocidad (g/día)",
            command=self._valorar_velocidad,
        )
        self.eval_velocidad_btn.grid(row=10, column=2, sticky="w", padx=5, pady=4)
        self.eval_zscore_btn = ttk.Button(
            marco_clas, text="3. Z-score peso/edad (OMS)",
            command=self._valorar_zscore_oms,
        )
        self.eval_zscore_btn.grid(row=10, column=3, sticky="w", padx=5, pady=4)
        self.peso_edad_resultado = tk.Label(
            marco_clas, text="", justify="left", foreground="#8B4513", wraplength=700)
        self.peso_edad_resultado.grid(row=11, column=0, columnspan=4, sticky="w", padx=5, pady=4)
        self.velocidad_resultado = tk.Label(
            marco_clas, text="", justify="left", foreground="#8B4513", wraplength=700)
        self.velocidad_resultado.grid(row=12, column=0, columnspan=4, sticky="w", padx=5, pady=4)
        self.zscore_resultado = tk.Label(
            marco_clas, text="", justify="left", foreground="#8B4513", wraplength=700)
        self.zscore_resultado.grid(row=13, column=0, columnspan=4, sticky="w", padx=5, pady=4)
        self._valoracion_peso_edad = ""
        self._valoracion_velocidad = ""
        self._valoracion_zscore = ""
        self._actualizar_visibilidad_ganancia()

    # ---------- Pestaña 2 Signos Clínicos ----------
    def _tab_signos(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="2. Signos Clínicos")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Signos Clínicos:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_signos).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_signos).pack(side="left", padx=3)

        self.signos_tree = self._crear_grid(
            tab,
            ["Órgano / Signo", "Signo Clínico", "Probable Alteración Nutricional"],
            alturas=12,
        )
        for organo in ORGANOS:
            self.signos_tree.insert("", "end", values=(organo, "", ""))

    # ---------- Pestaña 3 Interacción Fármaco-Nutriente ----------
    def _tab_farmaco(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="3. Interacción Fármaco-Nutriente")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Fármaco-Nutriente:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_farmaco).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_farmaco).pack(side="left", padx=3)

        self.farmaco_tree = self._crear_grid(
            tab, ["Fármaco", "Vía", "Interacción", "Recomendación"], alturas=8
        )
        ttk.Button(tab, text="+ Agregar fila", command=self._agregar_fila).pack(pady=4)
        ttk.Button(tab, text="- Quitar fila seleccionada", command=self._quitar_fila).pack(pady=2)

    def _agregar_fila(self):
        self.farmaco_tree.insert("", "end", values=("", "", "", ""))

    def _quitar_fila(self):
        sel = self.farmaco_tree.selection()
        if sel:
            self.farmaco_tree.delete(sel[0])

    # ---------- Pestaña 4 Evaluación Bioquímica ----------
    def _tab_bioquimica(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="4. Evaluación Bioquímica")

        barra = ttk.Frame(tab)
        barra.pack(fill="x", padx=8, pady=4)
        ttk.Label(barra, text="Evaluación Bioquímica:").pack(side="left", padx=5)
        ttk.Button(barra, text="Guardar", command=self._guardar_bioquimica).pack(side="left", padx=3)
        ttk.Button(barra, text="Editar", command=self._editar_bioquimica).pack(side="left", padx=3)

        self.bio_tree = self._crear_grid(
            tab, ["Pruebas Bioquímicas", "Valores Normales", "Resultados"], alturas=20
        )
        for prueba, valor in PRUEBAS:
            self.bio_tree.insert("", "end", values=(prueba, valor, ""))

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
        try:
            ref = datetime.strptime(hasta, "%Y-%m-%d").date()
        except ValueError:
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

    def _calcular_edad(self):
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            self.campos["edad"].set("")
            return
        semanas, resto = divmod(dias, 7)
        self.campos["edad"].set(f"{semanas} semanas y {resto} días")

    def _clasificar(self):
        try:
            semanas = int(re.match(r"\d+", self.ant["edad_gestacional"].get().strip()).group())
        except (AttributeError, ValueError):
            messagebox.showwarning("Clasificación", "Ingrese la edad gestacional en semanas (ej.: 39)")
            return
        try:
            peso = float(self.ant["peso_nacer"].get().replace(",", ".").strip())
            if peso < 10:
                peso *= 1000
        except ValueError:
            messagebox.showwarning("Clasificación", "Ingrese el peso al nacer (en kg en Antecedentes Prenatales)")
            return
        sexo = self.sexo_nacer.get()
        conn = conectar()
        fila = conn.execute(
            "SELECT percentil_10_g, percentil_50_g, percentil_90_g FROM referencia_peso_nacer WHERE sexo = ? AND edad_gestacional_semanas = ?",
            (sexo, semanas),
        ).fetchone()
        conn.close()
        if fila is None:
            messagebox.showwarning("Clasificación", f"No hay referencia para la semana {semanas} ({sexo})")
            return
        p10, p50, p90 = fila
        if peso < p10:
            codigo = "PEG"
            texto = "Pequeño para la Edad de Gestación"
        elif peso <= p90:
            codigo = "AEG"
            texto = "Apropiado para la Edad de Gestación"
        else:
            codigo = "GEG"
            texto = "Grande para la Edad de Gestación"
        self._clasificacion = f"{codigo} - {texto}"
        self.clas_peso_ok = True
        self.clas_resultado.config(
            text=f"P10: {int(p10)} g | P50: {int(p50)} g | P90: {int(p90)} g\n"
                 f"Peso: {int(peso)} g --> {codigo}: {texto}"
        )

    def _valorar_ganancia(self):
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            messagebox.showwarning("Ganancia de peso", "Ingrese la fecha de nacimiento y la fecha actual válidas")
            return
        if dias < 0:
            messagebox.showwarning("Ganancia de peso", "La fecha actual no puede ser anterior a la de nacimiento")
            return
        if dias > 28:
            messagebox.showwarning("Ganancia de peso", "No hay referencia para más de 28 días de vida")
            return
        if dias <= 6:
            intervalo = "0-7"
        elif dias <= 13:
            intervalo = "7-14"
        else:
            intervalo = "14-28"
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
        self._valoracion_ganancia = (
            f"Sexo: {sexo} | Intervalo: {intervalo} días | Rango Peso Nacer: {columna}\n"
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
        }
        self.ganancia_resultado.config(
            text=f"Días de vida: {dias} | Ganancia: {int(ganancia)} g\n"
                 f"Percentil estimado: {percentil}\n"
                 f"Referencias ({columna}): P5 {int(p5)} | P10 {int(p10)} | P25 {int(p25)} | P50 {int(p50)} g"
        )

    def _valorar_peso_edad(self):
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            messagebox.showwarning("Peso para la edad", "Ingrese la fecha de nacimiento y la fecha actual válidas")
            return
        if dias < 0:
            messagebox.showwarning("Peso para la edad", "La fecha actual no puede ser anterior a la de nacimiento")
            return
        if dias > 365:
            messagebox.showwarning("Peso para la edad", "No hay referencia para más de 12 meses de edad")
            return
        try:
            peso = float(self.clas_peso_actual.get().replace(",", ".").strip())
            if peso < 10:
                peso *= 1000
        except ValueError:
            messagebox.showwarning("Peso para la edad", "Ingrese el peso actual (g o kg)")
            return
        sexo = self.sexo_nacer.get()
        texto = _texto_peso_edad(sexo, dias, peso)
        if texto is None:
            messagebox.showwarning("Peso para la edad", f"No se pudo evaluar para {sexo} a {dias} días (revise datos o referencia)")
            return
        self._valoracion_peso_edad = texto
        self.peso_edad_resultado.config(text=self._valoracion_peso_edad)

    def _valorar_velocidad(self):
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            messagebox.showwarning("Velocidad de crecimiento", "Ingrese la fecha de nacimiento y la fecha actual válidas")
            return
        if dias < 0:
            messagebox.showwarning("Velocidad de crecimiento", "La fecha actual no puede ser anterior a la de nacimiento")
            return
        if dias <= 0:
            messagebox.showwarning("Velocidad de crecimiento", "Se necesitan al menos 1 día de vida para calcular g/día")
            return
        if dias >= 365:
            messagebox.showwarning("Velocidad de crecimiento", "No hay referencia para 12 meses o más")
            return
        try:
            peso_nacer = float(self.ant["peso_nacer"].get().replace(",", ".").strip())
            if peso_nacer < 10:
                peso_nacer *= 1000
            peso_actual = float(self.clas_peso_actual.get().replace(",", ".").strip())
            if peso_actual < 10:
                peso_actual *= 1000
        except ValueError:
            messagebox.showwarning("Velocidad de crecimiento", "Ingrese el peso al nacer y el peso actual")
            return
        sexo = self.sexo_nacer.get()
        texto = _texto_velocidad(sexo, dias, peso_nacer, peso_actual)
        if texto is None:
            messagebox.showwarning("Velocidad de crecimiento", f"No se pudo evaluar para {sexo} a {dias} días de vida (revise peso al nacer, peso actual o referencia)")
            return
        self._valoracion_velocidad = texto
        self.velocidad_resultado.config(text=self._valoracion_velocidad)

    def _valorar_zscore_oms(self):
        if not PIGROWUP_OK:
            messagebox.showwarning(
                "Z-score OMS",
                "El paquete 'pygrowup2' no está instalado.\n"
                "Instálelo con:  pip install pygrowup2",
            )
            return
        texto_nac = self.campos["fecha_nacimiento"].get().strip()
        if not texto_nac:
            messagebox.showwarning("Z-score OMS", "Ingrese la fecha de nacimiento y la fecha actual válidas")
            return
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                nacimiento = datetime.strptime(texto_nac, fmt).date()
                break
            except ValueError:
                continue
        else:
            messagebox.showwarning("Z-score OMS", "Fecha de nacimiento no válida")
            return
        try:
            ref = datetime.strptime(self.campos["fecha_actual"].get().strip(), "%Y-%m-%d").date()
        except ValueError:
            ref = date.today()
        dias = (ref - nacimiento).days
        if dias < 0:
            messagebox.showwarning("Z-score OMS", "La fecha actual no puede ser anterior a la de nacimiento")
            return
        if dias > 1856:
            messagebox.showwarning("Z-score OMS", "No hay z-score OMS para más de 5 años (61 meses)")
            return
        try:
            peso = float(self.clas_peso_actual.get().replace(",", ".").strip())
            if peso < 10:
                peso *= 1000
        except ValueError:
            messagebox.showwarning("Z-score OMS", "Ingrese el peso actual (g o kg)")
            return
        texto = _texto_zscore_oms(self.sexo_nacer.get(), dias, peso)
        if texto is None:
            messagebox.showwarning("Z-score OMS", "No se pudo calcular el z-score (revise edad dentro de 0-61 meses o la instalación de pygrowup2)")
            return
        self._valoracion_zscore = texto
        self.zscore_resultado.config(text=self._valoracion_zscore)

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
                           ganancia_g=?, percentil=?, intervalo_dias=?, dias_vida=?
                       WHERE id=? AND paciente_id=?""",
                    (
                        fecha, v["sexo"], v["peso_nacer_g"], v["peso_actual_g"],
                        v["ganancia_g"], v["percentil"], v["intervalo"], v["dias"],
                        consulta_id, self._paciente_id,
                    ),
                )
                mensaje = f"Consulta del {fecha} actualizada"
            else:
                conn.execute(
                    """INSERT INTO visita_ganancia_peso
                       (paciente_id, fecha_visita, sexo, peso_nacer_g, peso_actual_g,
                        ganancia_g, percentil, intervalo_dias, dias_vida)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (
                        self._paciente_id, fecha, v["sexo"], v["peso_nacer_g"],
                        v["peso_actual_g"], v["ganancia_g"], v["percentil"],
                        v["intervalo"], v["dias"],
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
                """SELECT id, fecha_visita, dias_vida, peso_actual_g, ganancia_g, percentil
                   FROM visita_ganancia_peso WHERE paciente_id = ? ORDER BY fecha_visita, id""",
                (self._paciente_id,),
            ).fetchall()
            conn.close()
        except Exception:
            return
        for fila in filas:
            item = self.visita_tree.insert(
                "", "end",
                values=(fila[1], fila[2], int(fila[3]), int(fila[4]), fila[5]),
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
            self.ganancia_resultado, self.guardar_visita_btn,
            self.refrescar_historial_btn, self.cargar_visita_btn, self.marco_vis,
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
        self.ganancia_resultado.config(text="")
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
        self.ganancia_resultado.config(text="")
        self.peso_edad_resultado.config(text="")
        self.velocidad_resultado.config(text="")
        self.zscore_resultado.config(text="")
        self.clas_peso_ok = False
        self._clasificacion = ""
        self._valoracion_ganancia = ""
        self._valoracion_peso_edad = ""
        self._valoracion_velocidad = ""
        self._valoracion_zscore = ""
        self._visita = None
        self._consulta_id = None
        for item in self.visita_tree.get_children():
            self.visita_tree.delete(item)
        for item in self.signos_tree.get_children():
            self.signos_tree.delete(item)
        for organo in ORGANOS:
            self.signos_tree.insert("", "end", values=(organo, "", ""))
        for item in self.farmaco_tree.get_children():
            self.farmaco_tree.delete(item)
        for item in self.bio_tree.get_children():
            self.bio_tree.delete(item)
        for prueba, valor in PRUEBAS:
            self.bio_tree.insert("", "end", values=(prueba, valor, ""))
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
        if len(valores) == 6:
            consulta_id, fecha, dias, peso_actual, ganancia, percentil = valores
        else:
            consulta_id = None
            fecha, dias, peso_actual, ganancia, percentil = valores
        if isinstance(peso_actual, float) and peso_actual.is_integer():
            peso_actual = int(peso_actual)
        if isinstance(ganancia, float) and ganancia.is_integer():
            ganancia = int(ganancia)
        self._consulta_id = consulta_id
        self.clas_peso_actual.set(str(peso_actual))
        self.campos["fecha_actual"].set(str(fecha))
        self.ganancia_resultado.config(
            text=f"Consulta {fecha}: {dias} días, peso {peso_actual} g, "
                 f"ganancia {ganancia} g, percentil {percentil}"
        )
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
            cur.execute("DELETE FROM antecedentes WHERE paciente_id = ?", (self._paciente_id,))
            cur.execute(
                """INSERT INTO antecedentes
                   (paciente_id, edad_gestacional, parto, peso_nacer,
                    perimetro_cefalico_nacer, longitud_nacer, antecedentes_familiares,
                    clasificacion_peso_nacer, valoracion_ganancia_peso,
                    valoracion_peso_edad, valoracion_velocidad_crecimiento,
                    valoracion_zscore_oms)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    self._paciente_id, self.ant["edad_gestacional"].get(), self.ant["parto"].get(),
                    self.ant["peso_nacer"].get(), self.ant["perimetro_cefalico_nacer"].get(),
                    self.ant["longitud_nacer"].get(),
                    self.ant["familiares"].get("1.0", "end").strip(),
                    getattr(self, "_clasificacion", ""),
                    getattr(self, "_valoracion_ganancia", ""),
                    getattr(self, "_valoracion_peso_edad", ""),
                    getattr(self, "_valoracion_velocidad", ""),
                    getattr(self, "_valoracion_zscore", ""),
                ),
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
                self.peso_edad_resultado.config(text="")
                self.velocidad_resultado.config(text="")
                self.zscore_resultado.config(text="")
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
        self.peso_edad_resultado.config(text=self._valoracion_peso_edad)
        self.velocidad_resultado.config(text=self._valoracion_velocidad)
        self.zscore_resultado.config(text=self._valoracion_zscore)
        self.status.config(text=f"Editando antecedentes | Paciente N° {self._paciente_id}")
        self._refrescar_historial()

    def _guardar_signos(self):
        if not self._exigir_paciente():
            return
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("DELETE FROM signos_clinicos WHERE paciente_id = ?", (self._paciente_id,))
            for item in self.signos_tree.get_children():
                signo, hallazgo, alteracion = self.signos_tree.item(item, "values")
                if hallazgo or alteracion:
                    cur.execute(
                        "INSERT INTO signos_clinicos (paciente_id, signo, signo_clinico, probable_alteracion) VALUES (?,?,?,?)",
                        (self._paciente_id, signo, hallazgo, alteracion),
                    )
            conn.commit()
            conn.close()
            self.status.config(text=f"Signos clínicos guardados | Paciente N° {self._paciente_id}")
            messagebox.showinfo("Guardado", "Signos clínicos guardados")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _editar_signos(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        conn = conectar()
        filas = conn.execute(
            "SELECT signo, signo_clinico, probable_alteracion FROM signos_clinicos WHERE paciente_id = ? ORDER BY id",
            (self._paciente_id,),
        ).fetchall()
        conn.close()
        for item in self.signos_tree.get_children():
            self.signos_tree.delete(item)
        for organo in ORGANOS:
            self.signos_tree.insert("", "end", values=(organo, "", ""))
        for i, item in enumerate(self.signos_tree.get_children()):
            if i < len(filas):
                self.signos_tree.item(item, values=(filas[i][0], filas[i][1] or "", filas[i][2] or ""))
        if not filas:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay signos clínicos guardados para este paciente")
        else:
            self.status.config(text=f"Editando signos clínicos | Paciente N° {self._paciente_id}")

    def _guardar_farmaco(self):
        if not self._exigir_paciente():
            return
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("DELETE FROM interaccion_farmaco WHERE paciente_id = ?", (self._paciente_id,))
            for item in self.farmaco_tree.get_children():
                farmaco, via, interaccion, recomendacion = self.farmaco_tree.item(item, "values")
                if any([farmaco, via, interaccion, recomendacion]):
                    cur.execute(
                        "INSERT INTO interaccion_farmaco (paciente_id, farmaco, via, interaccion, recomendacion) VALUES (?,?,?,?,?)",
                        (self._paciente_id, farmaco, via, interaccion, recomendacion),
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
            "SELECT farmaco, via, interaccion, recomendacion FROM interaccion_farmaco WHERE paciente_id = ? ORDER BY id",
            (self._paciente_id,),
        ).fetchall()
        conn.close()
        for item in self.farmaco_tree.get_children():
            self.farmaco_tree.delete(item)
        for fila in filas:
            self.farmaco_tree.insert("", "end", values=(fila[0] or "", fila[1] or "", fila[2] or "", fila[3] or ""))
        if not filas:
            if not silencioso:
                messagebox.showinfo("Editar", "No hay fármacos guardados para este paciente")
        else:
            self.status.config(text=f"Editando fármaco-nutriente | Paciente N° {self._paciente_id}")

    def _guardar_bioquimica(self):
        if not self._exigir_paciente():
            return
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("DELETE FROM evaluacion_bioquimica WHERE paciente_id = ?", (self._paciente_id,))
            for item in self.bio_tree.get_children():
                prueba, valor_normal, resultado = self.bio_tree.item(item, "values")
                cur.execute(
                    "INSERT INTO evaluacion_bioquimica (paciente_id, prueba, valor_normal, resultado) VALUES (?,?,?,?)",
                    (self._paciente_id, prueba, valor_normal, resultado),
                )
            conn.commit()
            conn.close()
            self.status.config(text=f"Evaluación bioquímica guardada | Paciente N° {self._paciente_id}")
            messagebox.showinfo("Guardado", "Evaluación bioquímica guardada")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _editar_bioquimica(self, silencioso=False):
        if silencioso and not getattr(self, "_paciente_id", None):
            return
        if not self._exigir_paciente():
            return
        conn = conectar()
        filas = conn.execute(
            "SELECT prueba, valor_normal, resultado FROM evaluacion_bioquimica WHERE paciente_id = ? ORDER BY id",
            (self._paciente_id,),
        ).fetchall()
        conn.close()
        for item in self.bio_tree.get_children():
            self.bio_tree.delete(item)
        for fila in filas:
            self.bio_tree.insert("", "end", values=(fila[0] or "", fila[1] or "", fila[2] or ""))
        if not filas:
            for prueba, valor in PRUEBAS:
                self.bio_tree.insert("", "end", values=(prueba, valor, ""))
            if not silencioso:
                messagebox.showinfo("Editar", "No hay evaluaciones bioquímicas guardadas para este paciente")
        else:
            self.status.config(text=f"Editando evaluación bioquímica | Paciente N° {self._paciente_id}")

    def _cargar_ultimos_datos(self):
        self._editar_antecedentes(silencioso=True)
        self._editar_signos(silencioso=True)
        self._editar_farmaco(silencioso=True)
        self._editar_bioquimica(silencioso=True)
        self.status.config(
            text=f"Datos de antecedentes, signos, fármacos y bioquímica cargados | Paciente N° {self._paciente_id}"
        )

if __name__ == "__main__":
    App().mainloop()
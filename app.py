import csv
import re
import sqlite3
import tkinter as tk
from datetime import date, datetime
from tkinter import ttk, messagebox
from pathlib import Path

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


def conectar():
    conn = sqlite3.connect(DB)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Historia Clínica Nutricional Pediátrica (1-4)")
        self.geometry("900x640")
        self._inicializar_tablas()
        self._crear_notebook()

    def _inicializar_tablas(self):
        sql = (Path(__file__).resolve().parent / "esquema.sql").read_text(encoding="utf-8")
        conn = conectar()
        conn.executescript(sql)
        self._migrar(conn)
        self._cargar_referencias(conn)
        conn.commit()
        conn.close()

    def _migrar(self, conn):
        columnas = [fila[1] for fila in conn.execute("PRAGMA table_info(antecedentes)")]
        if "clasificacion_peso_nacer" not in columnas:
            conn.execute("ALTER TABLE antecedentes ADD COLUMN clasificacion_peso_nacer TEXT")
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
        if conn.execute("SELECT COUNT(*) FROM referencia_peso_nacer").fetchone()[0]:
            return
        data_dir = Path(__file__).resolve().parent / "data"
        for archivo, sexo in (("ninos.csv", "Masculino"), ("ninas.csv", "Femenino")):
            with open(data_dir / archivo, encoding="utf-8-sig") as f:
                for fila in csv.DictReader(f):
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
        with open(data_dir / "clasificaciones.csv", encoding="utf-8-sig") as f:
            for fila in csv.DictReader(f):
                conn.execute(
                    "INSERT INTO clasificacion_eg (percentil, interpretacion) VALUES (?,?)",
                    (fila["Percentil"], fila["Interpretación"]),
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
        ttk.Button(barra, text="Guardar Historia Clínica", command=self.guardar).pack(side="left")
        ttk.Button(barra, text="Nuevo Paciente", command=self.limpiar).pack(side="left", padx=8)
        ttk.Button(barra, text="Salir", command=self.destroy).pack(side="right")
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

        marco = ttk.LabelFrame(tab, text="Datos Generales")
        marco.pack(fill="x", padx=8, pady=8)
        campos = [
            ("Nombre", "nombre"), ("Servicio", "servicio"), ("DNI/HC", "dni"),
            ("Cuenta", "cuenta"), ("Cama", "cama"),
        ]
        self.campos = {}
        for i, (etiqueta, clave) in enumerate(campos):
            fila, col = divmod(i, 3)
            ttk.Label(marco, text=etiqueta).grid(row=fila, column=col * 2, sticky="w", padx=5, pady=4)
            var = tk.StringVar()
            ttk.Entry(marco, textvariable=var, width=16).grid(row=fila, column=col * 2 + 1, padx=5, pady=4)
            self.campos[clave] = var

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
        ttk.Label(marco_clas, text="Peso al nacer en gramos (si está en kg, se convierte a g):").grid(
            row=1, column=0, sticky="w", padx=5, pady=4)
        self.clas_peso = tk.StringVar()
        ttk.Entry(marco_clas, textvariable=self.clas_peso, width=14).grid(row=1, column=1, sticky="w", padx=5, pady=4)
        ttk.Button(marco_clas, text="Clasificar", command=self._clasificar).grid(row=2, column=0, sticky="w", padx=5, pady=4)
        ttk.Button(marco_clas, text="Calcular EG desde F. de Nacimiento", command=self._calc_eg_desde_nacimiento).grid(
            row=2, column=1, sticky="w", padx=5, pady=4)
        self.clas_resultado = tk.Label(marco_clas, text="", justify="left", foreground="blue", wraplength=700)
        self.clas_resultado.grid(row=3, column=0, columnspan=4, sticky="w", padx=5, pady=4)
        self.clas_peso_ok = False

    # ---------- Pestaña 2 Signos Clínicos ----------
    def _tab_signos(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="2. Signos Clínicos")
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

    def _dias_desde_nacimiento(self):
        dias = self._parsedias_desde_nacimiento()
        if dias is None:
            messagebox.showwarning("Edad", "Fecha de nacimiento no válida (use AAAA-MM-DD)")
        return dias

    def _calc_eg_desde_nacimiento(self):
        dias = self._dias_desde_nacimiento()
        if dias is None:
            return
        semanas, resto = divmod(dias, 7)
        self.ant["edad_gestacional"].set(f"{semanas} semanas y {resto} días")
        self.clas_resultado.config(text=f"Desde el nacimiento: {dias} días = {semanas} semanas y {resto} días")
        self._clasificacion = ""

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
            peso = float(self.clas_peso.get().replace(",", ".").strip())
            if peso < 10:
                peso *= 1000
        except ValueError:
            messagebox.showwarning("Clasificación", "Ingrese el peso al nacer (en gramos o kg)")
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

    def guardar(self):
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute(
                """INSERT INTO paciente
                   (nombre, servicio, direccion, telefono, dni_hc, cuenta, sexo, cama, fecha_nacimiento,
                    edad, fecha_actual,
                    dx_medico, grado_instruccion_padre, grado_instruccion_madre)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    self.campos["nombre"].get(), self.campos["servicio"].get(),
                    self.campos["direccion"].get(), self.campos["telefono"].get(),
                    self.campos["dni"].get(),
                    self.campos["cuenta"].get(), self.sexo.get(), self.campos["cama"].get(),
                    self.campos["fecha_nacimiento"].get(), self.campos["edad"].get(),
                    self.campos["fecha_actual"].get(), self.campos["dx_medico"].get(),
                    self.campos["grado_padre"].get(), self.campos["grado_madre"].get(),
                ),
            )
            paciente_id = cur.lastrowid

            cur.execute(
                """INSERT INTO antecedentes
                   (paciente_id, edad_gestacional, parto, peso_nacer,
                    perimetro_cefalico_nacer, longitud_nacer, antecedentes_familiares,
                    clasificacion_peso_nacer)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    paciente_id, self.ant["edad_gestacional"].get(), self.ant["parto"].get(),
                    self.ant["peso_nacer"].get(), self.ant["perimetro_cefalico_nacer"].get(),
                    self.ant["longitud_nacer"].get(),
                    self.ant["familiares"].get("1.0", "end").strip(),
                    getattr(self, "_clasificacion", ""),
                ),
            )

            for item in self.signos_tree.get_children():
                signo, hallazgo, alteracion = self.signos_tree.item(item, "values")
                if hallazgo or alteracion:
                    cur.execute(
                        "INSERT INTO signos_clinicos (paciente_id, signo, signo_clinico, probable_alteracion) VALUES (?,?,?,?)",
                        (paciente_id, signo, hallazgo, alteracion),
                    )

            for item in self.farmaco_tree.get_children():
                farmaco, via, interaccion, recomendacion = self.farmaco_tree.item(item, "values")
                if any([farmaco, via, interaccion, recomendacion]):
                    cur.execute(
                        "INSERT INTO interaccion_farmaco (paciente_id, farmaco, via, interaccion, recomendacion) VALUES (?,?,?,?,?)",
                        (paciente_id, farmaco, via, interaccion, recomendacion),
                    )

            for item in self.bio_tree.get_children():
                prueba, valor_normal, resultado = self.bio_tree.item(item, "values")
                cur.execute(
                    "INSERT INTO evaluacion_bioquimica (paciente_id, prueba, valor_normal, resultado) VALUES (?,?,?,?)",
                    (paciente_id, prueba, valor_normal, resultado),
                )

            conn.commit()
            conn.close()
            self.status.config(text=f"Guardado correctamente. Paciente N° {paciente_id}")
            messagebox.showinfo("Guardado", f"Historia clínica guardada. Paciente N° {paciente_id}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def limpiar(self):
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
        self.clas_peso.set("")
        self.clas_resultado.config(text="")
        self.clas_peso_ok = False
        self._clasificacion = ""
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
        self.status.config(text="")


if __name__ == "__main__":
    App().mainloop()
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class RadioRiskApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Evaluación de Riesgo en Radioterapia")

        # Tamaño fijo para mantener relación de aspecto
        self.ancho_fijo = 480
        self.alto_fijo = 600
        self.root.geometry(f"{self.ancho_fijo}x{self.alto_fijo}")
        self.root.minsize(500, 650)

        self.datos_paciente = {}
        self.entries = {}  # Aquí guardaremos las variables (StringVar, BooleanVar, etc.)
        self.tecnicas_fallidas = []
        self.create_main_menu()

    def create_main_menu(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        menu_container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo)
        menu_container.place(relx=0.5, rely=0.5, anchor="center")
        menu_container.pack_propagate(False)

        tk.Label(menu_container, text="Menú Principal", font=("Arial", 18, "bold")).pack(pady=50)
        btn_cargar = tk.Button(menu_container, text="Cargar Paciente", width=25, height=2,
                               command=self.cargar_archivo, font=("Arial", 10))
        btn_cargar.pack(pady=10)

    def cargar_archivo(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar reporte",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath:
            try:
                self.tecnicas_fallidas = []
                self.extraer_datos(filepath)
                self.mostrar_detalles_paciente()
            except Exception as e:
                messagebox.showerror("Error", f"Error: {e}")

    def extraer_datos(self, path):
        df = pd.read_excel(path, header=None)

        def buscar_valor(etiqueta):
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    if str(cell).strip() == etiqueta:
                        return str(df.iloc[i, j + 1]).strip()
            return "-"

        # --- Extracción de métricas de haces ---
        mcs_values, sas_values = [], []
        en_beam_metrics = False
        for i, row in df.iterrows():
            if str(row[0]).strip() == "BEAM METRICS":
                en_beam_metrics = True
                continue
            if en_beam_metrics:
                metrica = str(row[2]).strip()
                valor_str = str(row[3]).replace(',', '.')
                try:
                    valor_num = float(valor_str)
                    if metrica == "MCS":
                        mcs_values.append(valor_num)
                    elif metrica == "SAS":
                        sas_values.append(valor_num)
                except:
                    continue

        mcs_min = min(mcs_values) if mcs_values else "-"
        sas_max = max(sas_values) if sas_values else "-"

        # Se guarda en el dicionario
        self.datos_paciente = {
            "Plan": buscar_valor("PLAN NAME"),
            "Nombre": buscar_valor("PATIENT NAME"),
            "ID": buscar_valor("PATIENT ID"),
            "Sexo": buscar_valor("PATIENT SEX"),
            "Fractions": buscar_valor("FRACTIONS"),
            "MCS": buscar_valor("MCS"),
            "SAS": buscar_valor("SAS"),
            "PMU": buscar_valor("PMU"),
            "MCSmin": str(mcs_min),
            "SASmax": str(sas_max)
        }

    def actualizar_checkbox_ca(self, *args):
        """Lógica para marcar el checkbox según la región anatómica seleccionada"""
        region = self.entries["Region"].get()
        regiones_con_ca = ["COLON/RECTO", "PULMON", "CERVIX/UTERO", "CYC"]

        if region in regiones_con_ca:
            self.entries["CA"].set(True)
        else:
            self.entries["CA"].set(False)

    def mostrar_detalles_paciente(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        self.tecnicas_fallidas = [] #reinicio la lista QA

        # Contenedor para mantener relación de aspecto (Centrado)
        self.main_container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo, bd=1, relief="flat")
        self.main_container.place(relx=0.5, rely=0.5, anchor="center")
        self.main_container.pack_propagate(False)

        tk.Label(self.main_container, text="Información del Paciente y Plan", font=("Arial", 14, "bold")).pack(pady=10)

        frame_info = tk.LabelFrame(self.main_container, text=" Datos Extraídos ", padx=15, pady=10)
        frame_info.pack(padx=10, fill="both", expand=True)

        # Listas de opciones
        op_sexo = ["M", "F", "-"]
        op_anatomica = ["MAMA", "COLON/RECTO", "PULMON", "PROSTATA", "CERVIX/UTERO", "ESOFAGO", "CYC", "PANCREAS",
                        "VEJIGA", "ENCEFALO/SNC", "MIEMBROS", "OTROS"]
        op_tecnica = ["3D", "IMRT", "VMAT", "SRS", "SBRT", "FIF"]

        # --- LÓGICA DE DETECCIÓN AUTOMÁTICA ---
        plan_name = self.datos_paciente.get("Plan", "").upper()
        palabras_plan = plan_name.split()  # Divide el nombre por espacios

        # 1. Técnica (Busca coincidencia en cualquier parte del nombre)
        tecnica_defecto = "3D"
        for t in op_tecnica:
            if t in plan_name:
                tecnica_defecto = t
                break

        # 2. Región (Busca la TERCERA palabra del nombre del plan)
        region_defecto = "OTROS"
        if len(palabras_plan) >= 3:
            tercera_palabra = palabras_plan[2]  # El índice 2 es la tercera palabra
            if tercera_palabra in op_anatomica:
                region_defecto = tercera_palabra
        # ---------------------------------------

        # Variables de control
        self.entries["Sexo"] = tk.StringVar(value=self.datos_paciente.get("Sexo", "-"))
        self.entries["Region"] = tk.StringVar(value=region_defecto)
        self.entries["Tecnica"] = tk.StringVar(value=tecnica_defecto)
        self.entries["CA"] = tk.BooleanVar()
        self.entries["PPed"] = tk.BooleanVar()

        # Rastrear cambios en la región para el Checkbox automático
        self.entries["Region"].trace_add("write", self.actualizar_checkbox_ca)

        # Ejecutar la validación inicial del Checkbox (según la región detectada)
        self.actualizar_checkbox_ca()

        # Campos de solo lectura (Superiores)
        campos_fijos = [
            ("Plan Name", "Plan"), ("Patient Name", "Nombre"), ("Patient ID", "ID"),
            ("MCS Promedio", "MCS"), ("SAS Promedio", "SAS"), ("PMU Promedio", "PMU"),
            ("MCS Mínimo", "MCSmin"), ("SAS Máximo", "SASmax")
        ]

        for label_text, key in campos_fijos:
            row = tk.Frame(frame_info)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=f"{label_text}:", width=18, anchor="w", font=("Arial", 8)).pack(side="left")
            ent = tk.Entry(row, font=("Arial", 8))
            ent.insert(0, self.datos_paciente.get(key, "-"))
            ent.config(state="readonly")
            ent.pack(side="right", expand=True, fill="x")

        #tk.Label(frame_info, text="Configuración Clínica", font=("Arial", 8, "italic"), fg="blue").pack(pady=5)

        # Desplegables
        for lab, key, ops in [("Sexo", "Sexo", op_sexo), ("Región Ant.", "Region", op_anatomica),
                              ("Técnica", "Tecnica", op_tecnica)]:
            row = tk.Frame(frame_info)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=lab + ":", width=18, anchor="w", font=("Arial", 8, "bold")).pack(side="left")
            tk.OptionMenu(row, self.entries[key], *ops).pack(side="right", expand=True, fill="x")

        # Checkbox Cambios Anatómicos
        row_ca = tk.Frame(frame_info)
        row_ca.pack(fill="x", pady=5)
        tk.Checkbutton(row_ca, text="Presencia de cambios anatómicos", variable=self.entries["CA"],
                       font=("Arial", 8)).pack(side="left")

        # Checkbox Paciente Pediátrico
        row_ca = tk.Frame(frame_info)
        row_ca.pack(fill="x", pady=5)
        tk.Checkbutton(row_ca, text="Paciente Pediátrico", variable=self.entries["PPed"],
                       font=("Arial", 8)).pack(side="left")
        # Botones finales
        btn_calcular = tk.Button(self.main_container, text="Calcular Método QA", width=25, height=2,
                                 bg="#0078D7", fg="white", font=("Arial", 10, "bold"),
                                 command=self.ejecutar_arbol_decision)
        btn_calcular.pack(pady=10)
        tk.Button(self.main_container, text="Volver", command=self.create_main_menu).pack()

    def procesar_arbol_casos(self, datos):
        """Implementación de los 3 Casos y sus escenarios específicos"""
        tecnica = datos["Tecnica"]
        lista_base = []

        # --- CASO 1: 3D / FIF ---
        if tecnica in ["3D", "FIF"]:
            lista_base = ["Plancheck", "Calculo independiente", "LogFile"]

            if datos["CambiosAnatomicos"]:
                if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")
            if datos["PacientePediatrico"]:
                if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")

        # --- CASO 2: SRS / SBRT ---
        elif tecnica in ["SRS", "SBRT"]:
            lista_base = ["Plancheck", "Portal Dosimetry"]

            if datos["CambiosAnatomicos"]:
                if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")
            if datos["PacientePediatrico"]:
                if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")

        # --- CASO 3: IMRT / VMAT ---
        elif tecnica in ["IMRT", "VMAT"]:
            lista_base = ["Plancheck", "Calculo independiente", "LogFile"]

            # Variables numéricas para validación
            try:
                mcs = float(datos["MCSmin"])
                sas = float(datos["SASmax"])
                pmu = float(datos["PMU"])
                frac = int(datos["Fractions"])
            except:
                mcs, sas, pmu, frac = 0.5, 0.5, 1000, 3

            # Escenario 3.2: Índices Altos (Complejidad)
            if mcs < 0.5 and sas > 0.5 and pmu > 1000:
                if "Portal Dosimetry" not in lista_base: lista_base.append("Portal Dosimetry")

            # Escenario 3.4: Hipofraccionado (> 3 fracciones)
            if frac > 3:
                if "Portal Dosimetry" not in lista_base: lista_base.append("Portal Dosimetry")

            # Escenario 3.5: Cambios Anatómicos
            if datos["CambiosAnatomicos"]:
                if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")

            if datos["PacientePediatrico"]:
                if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")

            # Escenario 3.3: Monofraccionado (Fracciones = 1) -> No agrega nada (se queda base)
            # Escenario 3.1: Índices Bajos -> No agrega nada (se queda base)

        # 1. Eliminar técnicas que ya fallaron
        lista_final = [t for t in lista_base if t not in self.tecnicas_fallidas]

        # 2. Agregar refuerzos si hubo fallos previos (según reglas de "si no pasa")
        for fallida in self.tecnicas_fallidas:
            # Caso 2: Refuerzo específico
            if tecnica in ["SRS", "SBRT"] and fallida in lista_base:
                if "Stereophan + Gafchromic/CI" not in lista_final:
                    lista_final.append("Stereophan + Gafchromic/CI")

            # Caso 3: Refuerzo para escenarios 3.2 y 3.4
            if tecnica in ["IMRT", "VMAT"] and fallida in lista_base:
                # Solo si era complejo o hipofraccionado
                if (mcs < 0.5 and sas > 0.5 and pmu > 1000) or (frac > 3):
                    if "ArcCheck" not in lista_final: lista_final.append("ArcCheck")
                    if "3DVH" not in lista_final: lista_final.append("3DVH")

        return lista_final

    def ejecutar_arbol_decision(self):
        # Limpiar pantalla
        for widget in self.root.winfo_children():
            widget.destroy()

        res_container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo)
        res_container.place(relx=0.5, rely=0.5, anchor="center")
        res_container.pack_propagate(False)

        # Recolección de datos
        datos_actuales = {
            "Tecnica": self.entries["Tecnica"].get(),
            "Region": self.entries["Region"].get(),
            "CambiosAnatomicos": self.entries["CA"].get(),
            "PacientePediatrico": self.entries["PPed"].get(),
            "MCSmin": self.datos_paciente.get("MCSmin", "0.5"),
            "SASmax": self.datos_paciente.get("SASmax", "0.5"),
            "PMU": self.datos_paciente.get("PMU", "1000"),
            "Fractions": self.datos_paciente.get("Fractions", "1")
        }

        # Obtener lista de méritos procesada
        recomendaciones = self.procesar_arbol_casos(datos_actuales)

        tk.Label(res_container, text="EVALUACIÓN DE RIESGO QADS", font=("Arial", 14, "bold")).pack(pady=10)

        # Cuadro informativo de contexto
        info_frame = tk.Frame(res_container, bg="#f9f9f9", padx=10, pady=5)
        info_frame.pack(fill="x", padx=20)
        txt_contexto = f"Caso detectado: {datos_actuales['Tecnica']}\n" \
                       f"Métricas: MCS {datos_actuales['MCSmin']} | SAS {datos_actuales['SASmax']} | PMU {datos_actuales['PMU']}\n" \
                       f"Fracciones: {datos_actuales['Fractions']} | Cambios Anat: {'SÍ' if datos_actuales['CambiosAnatomicos'] else 'NO'}\n" \
                       f"Pac. Pediátrico: {'SÍ' if datos_actuales['PacientePediatrico'] else 'NO'}"

        tk.Label(info_frame, text=txt_contexto, font=("Arial", 8), bg="#f9f9f9", justify="left").pack()

        # Selección de técnica
        tk.Label(res_container, text="Seleccione Técnica Realizada:", font=("Arial", 10, "bold")).pack(pady=10)

        self.qa_seleccionada = tk.StringVar()
        if recomendaciones:
            self.qa_seleccionada.set(recomendaciones[0])
            for metodo in recomendaciones:
                tk.Radiobutton(res_container, text=metodo, variable=self.qa_seleccionada,
                               value=metodo, font=("Arial", 9)).pack(anchor="w", padx=50)
        else:
            tk.Label(res_container, text="No hay más técnicas disponibles.", fg="red").pack()

        # Feedback de éxito
        tk.Label(res_container, text="¿Resultado del Control?", font=("Arial", 10, "bold")).pack(pady=10)
        self.resultado_qa = tk.StringVar(value="Exitoso")
        op_res = tk.OptionMenu(res_container, self.resultado_qa, "Exitoso", "No Exitoso")
        op_res.config(width=15)
        op_res.pack()

        # Botones
        btn_frame = tk.Frame(res_container)
        btn_frame.pack(side="bottom", pady=20)

        tk.Button(btn_frame, text="Registrar", bg="#4CAF50", fg="white", width=12,
                  command=self.validar_resultado_qa).pack(side="left", padx=5)

        self.btn_excel = tk.Button(btn_frame, text="Informe Excel", state="disabled", width=12)
        self.btn_excel.pack(side="left", padx=5)

        tk.Button(btn_frame, text="Volver", command=self.mostrar_detalles_paciente, width=12).pack(side="left", padx=5)

    def validar_resultado_qa(self):
        if self.resultado_qa.get() == "No Exitoso":
            fallida = self.qa_seleccionada.get()
            if fallida not in self.tecnicas_fallidas:
                self.tecnicas_fallidas.append(fallida)

            messagebox.showwarning("Actualización de Árbol", f"QA {fallida} fallido. Se actualizan recomendaciones.")
            self.ejecutar_arbol_decision()
        else:
            messagebox.showinfo("Éxito", "Control validado correctamente.")
            self.btn_excel.config(state="normal", bg="#0078D7", fg="white")


if __name__ == "__main__":
    root = tk.Tk()
    app = RadioRiskApp(root)
    root.mainloop()
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

        # Guardamos en el diccionario (simplificado según tu petición)
        self.datos_paciente = {
            "Plan": buscar_valor("PLAN NAME"),
            "Nombre": buscar_valor("PATIENT NAME"),
            "ID": buscar_valor("PATIENT ID"),
            "Sexo": buscar_valor("PATIENT SEX"),
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

        # Botones finales
        btn_calcular = tk.Button(self.main_container, text="Calcular Método QA", width=25, height=2,
                                 bg="#0078D7", fg="white", font=("Arial", 10, "bold"),
                                 command=self.ejecutar_arbol_decision)
        btn_calcular.pack(pady=10)
        tk.Button(self.main_container, text="Volver", command=self.create_main_menu).pack()

    def ejecutar_arbol_decision(self):
        # Aquí recolectamos TODAS las variables para el árbol
        datos_finales = {
            "MCSmin": self.datos_paciente.get("MCSmin"),
            "SASmax": self.datos_paciente.get("SASmax"),
            "Tecnica": self.entries["Tecnica"].get(),
            "Region": self.entries["Region"].get(),
            "CambiosAnatomicos": self.entries["CA"].get(),
            "Sexo": self.entries["Sexo"].get()
        }

        for widget in self.root.winfo_children():
            widget.destroy()

        result_container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo)
        result_container.place(relx=0.5, rely=0.5, anchor="center")
        result_container.pack_propagate(False)

        tk.Label(result_container, text="Evaluación de Riesgo", font=("Arial", 14, "bold")).pack(pady=20)

        # Ejemplo de cómo usar los datos en la lógica:
        info_text = f"Técnica: {datos_finales['Tecnica']}\nRegión: {datos_finales['Region']}\n"
        info_text += f"¿Cambios Anatómicos?: {'SÍ' if datos_finales['CambiosAnatomicos'] else 'NO'}"

        tk.Label(result_container, text=info_text, justify="left", font=("Arial", 10)).pack(pady=20)

        # [AQUÍ IRÁ TU ÁRBOL DE DECISIÓN]

        tk.Button(result_container, text="Volver", command=self.mostrar_detalles_paciente).pack(pady=20)


if __name__ == "__main__":
    root = tk.Tk()
    app = RadioRiskApp(root)
    root.mainloop()
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class RadioRiskApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Evaluación de Riesgo en Radioterapia")
        self.root.geometry("500x600")

        # Variables para almacenar datos
        self.datos_paciente = {}

        self.create_main_menu()

    def create_main_menu(self):
        # Limpiar ventana
        for widget in self.root.winfo_children():
            widget.destroy()

        tk.Label(self.root, text="Menú Principal", font=("Arial", 16, "bold"), pady=20).pack()

        btn_cargar = tk.Button(self.root, text="Cargar Paciente", width=25, height=2,
                               command=self.cargar_archivo)
        btn_cargar.pack(pady=10)


    def cargar_archivo(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar reporte de plan",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if filepath:
            try:
                self.extraer_datos(filepath)
                self.mostrar_detalles_paciente()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el archivo: {e}")

    def extraer_datos(self, path):
        # Leemos el excel sin cabecera para buscar por coordenadas de texto
        df = pd.read_excel(path, header=None)

        # Función auxiliar para buscar valores por etiquetas
        def buscar_valor(etiqueta):
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    if str(cell).strip() == etiqueta:
                        return str(df.iloc[i, j + 1]).strip()
            return "-"

        # --- Lógica para buscar MCS min y SAS max en BEAM METRICS ---
        mcs_values = []
        sas_values = []

        en_beam_metrics = False
        for i, row in df.iterrows():
            celda_principal = str(row[0]).strip()

            # Detectamos cuando empieza la sección de Beams
            if celda_principal == "BEAM METRICS":
                en_beam_metrics = True
                continue

            if en_beam_metrics:
                # El Excel tiene el nombre de la métrica en una columna y el valor en otra
                # Según tu archivo, la métrica está en la columna index 2 y el valor en la 3
                metrica = str(row[2]).strip()
                valor_str = str(row[3]).replace(',', '.')  # Cambiar coma decimal por punto

                try:
                    valor_num = float(valor_str)
                    if metrica == "MCS":
                        mcs_values.append(valor_num)
                    elif metrica == "SAS":
                        sas_values.append(valor_num)
                except ValueError:
                    # Ignoramos filas que no tengan números (como encabezados internos)
                    continue

        # Calculamos extremos
        mcs_min = min(mcs_values) if mcs_values else "-"
        sas_max = max(sas_values) if sas_values else "-"

        # Se guarda en el diccionario
        self.datos_paciente = {
            "Plan": buscar_valor("PLAN NAME"),
            "Nombre": buscar_valor("PATIENT NAME"),
            "ID": buscar_valor("PATIENT ID"),
            "Sexo": buscar_valor("PATIENT SEX"),
            "Patologia": buscar_valor("PATHOLOGY"),
            "Localizacion": buscar_valor("LOCATION"),
            "MCS": buscar_valor("MCS"),
            "SAS": buscar_valor("SAS"),
            "PMU": buscar_valor("PMU"),
            "MCSmin": str(mcs_min),
            "SASmax": str(sas_max)
        }

    def mostrar_detalles_paciente(self):
        # Limpiar pantalla
        for widget in self.root.winfo_children():
            widget.destroy()

        tk.Label(self.root, text="Información del Paciente y Plan", font=("Arial", 12, "bold")).pack(pady=10)

        frame_info = tk.LabelFrame(self.root, text="Datos Extraídos", padx=20, pady=20)
        frame_info.pack(padx=10, fill="both")

        # ---------------------------------------------------------
        # SECCIÓN PARA EDITAR OPCIONES DESPLEGABLES
        # ---------------------------------------------------------
        opciones_sexo = ["M", "F", "-"]

        opciones_patologia = [
            "Próstata",
            "Mama",
            "Cabeza y Cuello",
            "Pulmón",
            "SNC",
            "Ginecológico",
            "-"  # Opción por defecto si no hay dato
        ]

        opciones_localizacion = [
            "Pelvis",
            "Tórax",
            "Abdomen",
            "Cráneo",
            "Columna",
            "-"
        ]
        # ---------------------------------------------------------

        self.entries = {}

        # Definimos los campos. El tercer valor indica el tipo:
        # 'text' (solo lectura), 'entry' (editable), 'menu' (desplegable)
        campos = [
            ("Plan Name", "Plan", "text", None),
            ("Patient Name", "Nombre", "text", None),
            ("Patient ID", "ID", "text", None),
            ("Sexo", "Sexo", "menu", opciones_sexo),
            ("Patología", "Patologia", "menu", opciones_patologia),
            ("Localización", "Localizacion", "menu", opciones_localizacion),
            ("MCS Promedio", "MCS", "text", None),
            ("SAS Promedio", "SAS", "text", None),
            ("PMU Promedio", "PMU", "text", None),
            ("MCS Mínimo", "MCSmin", "text", None),
            ("SAS Máximo", "SASmax", "text", None)
        ]

        for label_text, key, tipo, opciones in campos:
            row = tk.Frame(frame_info)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"{label_text}:", width=15, anchor="w").pack(side="left")

            valor_inicial = self.datos_paciente.get(key, "-")

            if tipo == "menu":
                # Variable de control para el OptionMenu
                var_menu = tk.StringVar(self.root)
                # Si el valor del Excel está en las opciones, lo selecciona, si no usa "-"
                if valor_inicial in opciones:
                    var_menu.set(valor_inicial)
                else:
                    var_menu.set("-")

                menu = tk.OptionMenu(row, var_menu, *opciones)
                menu.pack(side="right", expand=True, fill="x")
                self.entries[key] = var_menu  # Guardamos la variable, no el widget

            else:
                ent = tk.Entry(row)
                ent.insert(0, valor_inicial)
                ent.config(state="readonly")
                ent.pack(side="right", expand=True, fill="x")
                self.entries[key] = ent

        # Botón Calcular (ya visible pero deshabilitado o habilitado según desees)
        btn_calcular = tk.Button(self.root, text="Calcular Método QA", width=25, height=2,
                                 state="normal", command=self.ejecutar_arbol_decision)
        btn_calcular.pack(pady=10)

        tk.Button(self.root, text="Volver", command=self.create_main_menu).pack(pady=10)

    def ejecutar_arbol_decision(self):
        # Aquí recuperamos los datos finales (incluyendo los elegidos en los menús)
        sexo_elegido = self.entries["Sexo"].get()
        pat_elegida = self.entries["Patologia"].get()
        loc_elegida = self.entries["Localizacion"].get()

        messagebox.showinfo("Procesando", f"Evaluando riesgo para {pat_elegida} en {loc_elegida}...")
        # Aquí irá tu lógica de árbol de decisión...


if __name__ == "__main__":
    root = tk.Tk()
    app = RadioRiskApp(root)
    root.mainloop()
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class RadioRiskApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Evaluación de Riesgo en Radioterapia")

        # Configuración de ventana
        self.ancho_fijo = 480
        self.alto_fijo = 600
        self.root.geometry(f"{self.ancho_fijo}x{self.alto_fijo}")
        self.root.minsize(500, 680)

        self.datos_paciente = {}
        self.entries = {}
        self.tecnicas_fallidas = []
        self.create_main_menu()

    def create_main_menu(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        menu_container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo)
        menu_container.place(relx=0.5, rely=0.5, anchor="center")
        menu_container.pack_propagate(False)

        tk.Label(menu_container, text="Menú Principal", font=("Arial", 18, "bold")).pack(pady=50)
        tk.Button(menu_container, text="Cargar Paciente", width=25, height=2,
                  command=self.cargar_archivo, font=("Arial", 10)).pack(pady=10)

    def cargar_archivo(self):
        filepath = filedialog.askopenfilename(title="Seleccionar reporte", filetypes=[("Excel files", "*.xlsx *.xls")])
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

        mcs_values, sas_values = [], []
        en_beam_metrics = False
        for i, row in df.iterrows():
            if str(row[0]).strip() == "BEAM METRICS":
                en_beam_metrics = True
                continue
            if en_beam_metrics:
                try:
                    metrica, valor_str = str(row[2]).strip(), str(row[3]).replace(',', '.')
                    if metrica == "MCS":
                        mcs_values.append(float(valor_str))
                    elif metrica == "SAS":
                        sas_values.append(float(valor_str))
                except:
                    continue

        self.datos_paciente = {
            "Plan": buscar_valor("PLAN NAME"),
            "Nombre": buscar_valor("PATIENT NAME"),
            "ID": buscar_valor("PATIENT ID"),
            "Sexo": buscar_valor("PATIENT SEX"),
            "Fractions": buscar_valor("FRACTIONS"),
            "MCS": buscar_valor("MCS"),
            "SAS": buscar_valor("SAS"),
            "PMU": buscar_valor("PMU"),
            "MCSmin": str(min(mcs_values)) if mcs_values else "-",
            "SASmax": str(max(sas_values)) if sas_values else "-"
        }

    def actualizar_checkbox_ca(self, *args):
        region = self.entries["Region"].get()
        regiones_con_ca = ["COLON/RECTO", "PULMON", "CERVIX/UTERO", "CYC"]
        self.entries["CA"].set(region in regiones_con_ca)

    def mostrar_detalles_paciente(self):
        for widget in self.root.winfo_children(): widget.destroy()

        container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo)
        container.place(relx=0.5, rely=0.5, anchor="center")
        container.pack_propagate(False)

        tk.Label(container, text="Información del Paciente", font=("Arial", 14, "bold")).pack(pady=10)
        frame_info = tk.LabelFrame(container, text=" Datos Extraídos ", padx=15, pady=10)
        frame_info.pack(padx=10, fill="both", expand=True)

        op_sexo, op_tecnica = ["M", "F", "-"], ["3D", "IMRT", "VMAT", "SRS", "SBRT", "FIF"]
        op_anatomica = ["MAMA", "COLON/RECTO", "PULMON", "PROSTATA", "CERVIX/UTERO", "ESOFAGO", "CYC", "PANCREAS",
                        "VEJIGA", "ENCEFALO/SNC", "MIEMBROS", "OTROS"]

        plan_name = self.datos_paciente.get("Plan", "").upper()
        palabras = plan_name.split()
        tecnica_def = next((t for t in op_tecnica if t in plan_name), "3D")
        region_def = palabras[2] if len(palabras) >= 3 and palabras[2] in op_anatomica else "OTROS"

        # Variables de control
        self.entries["Sexo"] = tk.StringVar(value=self.datos_paciente.get("Sexo", "-"))
        self.entries["Region"] = tk.StringVar(value=region_def)
        self.entries["Tecnica"] = tk.StringVar(value=tecnica_def)
        self.entries["CA"] = tk.BooleanVar()
        self.entries["PPed"] = tk.BooleanVar()

        self.entries["Region"].trace_add("write", self.actualizar_checkbox_ca)
        self.actualizar_checkbox_ca()

        # --- RESTAURACIÓN DE TODOS LOS CAMPOS ---
        campos = [
            ("ID Paciente", "ID"),
            ("Nombre", "Nombre"),
            ("Plan", "Plan"),
            ("MCS Prom.", "MCS"),
            ("SAS Prom.", "SAS"),
            ("PMU", "PMU"),
            ("MCS Mínimo", "MCSmin"),
            ("SAS Máximo", "SASmax")
        ]
        for label, key in campos:
            row = tk.Frame(frame_info)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=label, width=15, anchor="w", font=("Arial", 8)).pack(side="left")
            e = tk.Entry(row, font=("Arial", 8))
            e.insert(0, self.datos_paciente.get(key, "-"))
            e.config(state="readonly")
            e.pack(side="right", fill="x", expand=True)

        # Desplegables actualizados
        for lab, key, ops in [("Sexo", "Sexo", op_sexo), ("Región", "Region", op_anatomica),
                              ("Técnica", "Tecnica", op_tecnica)]:
            row = tk.Frame(frame_info);
            row.pack(fill="x", pady=2)
            tk.Label(row, text=lab, width=15, anchor="w", font=("Arial", 8, "bold")).pack(side="left")
            tk.OptionMenu(row, self.entries[key], *ops).pack(side="right", fill="x", expand=True)

        tk.Checkbutton(frame_info, text="Cambios Anatómicos", variable=self.entries["CA"], font=("Arial", 8)).pack(
            anchor="w")
        tk.Checkbutton(frame_info, text="Paciente Pediátrico", variable=self.entries["PPed"], font=("Arial", 8)).pack(
            anchor="w")

        tk.Button(container, text="Calcular Método QA", bg="#0078D7", fg="white", font=("Arial", 10, "bold"),
                  command=self.ejecutar_arbol_decision).pack(pady=10)
        tk.Button(container, text="Volver", command=self.create_main_menu).pack()

    def procesar_arbol_casos(self, datos):
        tecnica, lista_base = datos["Tecnica"], []
        # Casos base
        if tecnica in ["3D", "FIF"]:
            lista_base = ["Plancheck", "Calculo independiente", "LogFile"]
        elif tecnica in ["SRS", "SBRT"]:
            lista_base = ["Plancheck", "Portal Dosimetry"]
        elif tecnica in ["IMRT", "VMAT"]:
            lista_base = ["Plancheck", "Calculo independiente", "LogFile"]
            try:
                if float(datos["MCSmin"]) < 0.5 and float(datos["SASmax"]) > 0.5: lista_base.append("Portal Dosimetry")
                if int(datos["Fractions"]) > 3: lista_base.append("Portal Dosimetry")
            except:
                pass

        if datos["CambiosAnatomicos"] or datos["PacientePediatrico"]:
            if "Transit-EPID" not in lista_base: lista_base.append("Transit-EPID")

        # Filtrar fallidas y añadir refuerzos
        final = [t for t in lista_base if t not in self.tecnicas_fallidas]
        for f in self.tecnicas_fallidas:
            if tecnica in ["SRS", "SBRT"] and f in ["Plancheck", "Portal Dosimetry"]:
                if "Stereophan + Gafchromic/CI" not in final: final.append("Stereophan + Gafchromic/CI")
            if tecnica in ["IMRT", "VMAT"] and (f == "Portal Dosimetry" or f == "Transit-EPID"):
                for r in ["ArcCheck", "3DVH"]:
                    if r not in final: final.append(r)
        return final

    def ejecutar_arbol_decision(self):
        for widget in self.root.winfo_children(): widget.destroy()
        res_container = tk.Frame(self.root, width=self.ancho_fijo, height=self.alto_fijo);
        res_container.place(relx=0.5, rely=0.5, anchor="center");
        res_container.pack_propagate(False)

        datos = {
            "Tecnica": self.entries["Tecnica"].get(), "Region": self.entries["Region"].get(),
            "CambiosAnatomicos": self.entries["CA"].get(), "PacientePediatrico": self.entries["PPed"].get(),
            "MCSmin": self.datos_paciente.get("MCSmin", "0.5"), "SASmax": self.datos_paciente.get("SASmax", "0.5"),
            "PMU": self.datos_paciente.get("PMU", "1000"), "Fractions": self.datos_paciente.get("Fractions", "1")
        }

        recomendaciones = self.procesar_arbol_casos(datos)
        tk.Label(res_container, text="EVALUACIÓN DE RIESGO QADS", font=("Arial", 14, "bold")).pack(pady=10)

        self.qa_seleccionada = tk.StringVar()
        if recomendaciones:
            self.qa_seleccionada.set(recomendaciones[0])
            for m in recomendaciones: tk.Radiobutton(res_container, text=m, variable=self.qa_seleccionada,
                                                     value=m).pack(anchor="w", padx=80)

        tk.Label(res_container, text="¿Resultado del Control?", font=("Arial", 10, "bold")).pack(pady=10)
        self.resultado_qa = tk.StringVar(value="Exitoso")
        tk.OptionMenu(res_container, self.resultado_qa, "Exitoso", "No Exitoso").pack()

        btn_frame = tk.Frame(res_container);
        btn_frame.pack(side="bottom", pady=20)
        tk.Button(btn_frame, text="Registrar", bg="#4CAF50", fg="white", width=12,
                  command=self.validar_resultado_qa).pack(side="left", padx=5)
        self.btn_excel = tk.Button(btn_frame, text="Informe Excel", state="disabled", width=12,
                                   command=self.exportar_informe)
        self.btn_excel.pack(side="left", padx=5)
        tk.Button(btn_frame, text="Volver", command=self.mostrar_detalles_paciente, width=12).pack(side="left", padx=5)

    def validar_resultado_qa(self):
        if self.resultado_qa.get() == "No Exitoso":
            f = self.qa_seleccionada.get()
            if f not in self.tecnicas_fallidas: self.tecnicas_fallidas.append(f)
            self.ejecutar_arbol_decision()
        else:
            messagebox.showinfo("Validado", "Control exitoso. Ya puede exportar el informe.")
            self.btn_excel.config(state="normal", bg="#0078D7", fg="white")

    def exportar_informe(self):
        nombre_archivo = "Registro_Historico_QA.xlsx"
        nueva_fila = {
            "ID Paciente": self.datos_paciente.get("ID", "-"),
            "Nombre Paciente": self.datos_paciente.get("Nombre", "-"),
            "Sexo": self.entries["Sexo"].get(),
            "Plan": self.datos_paciente.get("Plan", "-"),
            "Región Anatómica": self.entries["Region"].get(),
            "Técnica RT": self.entries["Tecnica"].get(),
            "MCS Prom.": self.datos_paciente.get("MCS", "-"),
            "SAS Prom.": self.datos_paciente.get("SAS", "-"),
            "PMU": self.datos_paciente.get("PMU", "-"),
            "Fracciones": self.datos_paciente.get("Fractions", "-")
        }

        # Historial de intentos
        for i, tecnica in enumerate(self.tecnicas_fallidas):
            nueva_fila[f"Técnica QA {i + 1}"] = tecnica
            nueva_fila[f"Resultado {i + 1}"] = "Fallido"

        idx_final = len(self.tecnicas_fallidas) + 1
        nueva_fila[f"Técnica QA {idx_final}"] = self.qa_seleccionada.get()
        nueva_fila[f"Resultado {idx_final}"] = "Exitoso"

        try:
            if os.path.exists(nombre_archivo):
                df_existente = pd.read_excel(nombre_archivo)
                df_nuevo = pd.DataFrame([nueva_fila])
                df_final = pd.concat([df_existente, df_nuevo], ignore_index=True, sort=False)
            else:
                df_final = pd.DataFrame([nueva_fila])

            df_final.to_excel(nombre_archivo, index=False)
            messagebox.showinfo("Éxito", f"Informe guardado en {nombre_archivo}")
            self.btn_excel.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")


if __name__ == "__main__":
    root = tk.Tk();
    app = RadioRiskApp(root);
    root.mainloop()
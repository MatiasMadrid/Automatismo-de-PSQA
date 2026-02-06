# 📊 Evaluación de Riesgo QADS - Radioterapia

Este software es una herramienta de gestión de calidad diseñada para evaluar la complejidad de los planes de tratamiento en Radioterapia y determinar el método de **QA (Quality Assurance)** más adecuado según protocolos de seguridad institucional.

---

## 📂 Archivos mínimos para el funcionamiento
Para que el programa funcione correctamente (versión portable `.exe`), los siguientes archivos deben residir en la **misma carpeta**:

* **`Evaluacion_QADS.exe`**: El ejecutable del programa.
* **`costos.xlsx`**: Archivo maestro con los precios de cada técnica de QA.
* **`Registro_Historico_2026.xlsx`**: Base de datos donde se guardan los resultados de los pacientes.

> [!IMPORTANTE]
> El programa generará automáticamente archivos de configuración (`config_ruta.txt` y `umbrales.txt`) tras el primer uso. **No los elimine**, ya que contienen sus preferencias personalizadas.

---

## 🚀 Guía rápida de uso

### 1. Carga de datos
1. Inicie el programa y haga clic en **"Cargar Paciente"**.
2. Seleccione el reporte de métricas en formato Excel (programa de Laura).
3. El software extraerá automáticamente valores críticos: **MCS, SAS, PMU y Dosis por Fracción**.

### 2. Evaluación de Complejidad
* Verifique los datos extraídos en pantalla.
* El sistema activará automáticamente indicadores de cambios anatómicos si corresponde.
* Haga clic en **"Calcular Método QA"**.

### 3. Registro y Resultados
* El sistema indicará el paquete de QA recomendado (ej: *Plancheck + LogFile + Portal Dosimetry*).
* Tras realizar el control físico, registre si el resultado fue **"Exitoso"** o **"No Exitoso"**.
* **Alerta Crítica:** En caso de fallo persistente o en planes de baja complejidad, el sistema emitirá una alerta de **REHACER PLAN**.
* Use el botón **"Informe Excel"** para volcar los datos y costos al registro histórico.

---

## ⚙️ Configuración (Uso Técnico)
Desde el panel de **Configuración** se puede:
* **Umbrales:** Ajustar los valores de corte para MCS, SAS (promedios y extremos) y PMU según el criterio clínico del servicio.
* **Costos:** Abrir el archivo Excel para actualizar los valores monetarios de los procedimientos.
* **Registros:** Vincular el programa a un archivo existente o crear uno nuevo para un período diferente.

---

## ⚠️ Requisitos para Windows 7 Ultimate
Para asegurar la compatibilidad en terminales con Windows 7, verifique:
* **Service Pack 1 (SP1)** instalado.
* **Universal C Runtime (KB2999226)** instalado.
* Arquitectura compatible (ejecutar versión de 32 bits si el sistema es x86).

---

### Estructura de archivos esperada:
```text
/Carpeta_del_Programa
├── Evaluacion_QADS.exe
├── costos.xlsx
├── Registro_Historico_2026.xlsx
├── umbrales.txt (generado)
└── config_ruta.txt (generado)

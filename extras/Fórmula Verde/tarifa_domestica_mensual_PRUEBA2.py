# --- Archivo: tarifa_domestica_mensual.py ---
# (Versión Final: Impacto Ambiental 25 años con NUEVOS EMOJIS en G21:G23)

import re
import pdfplumber
import openpyxl
from openpyxl.styles import Alignment, Font
from tkinter import Tk, filedialog, messagebox
import os
import subprocess
import platform
import win32com.client
import math
import datetime

# --- Esta es la función principal que importará tu app_maestro ---
def procesar_tarifa_domestica_mensual():

    # --- FUNCIÓN INTERNA: CÁLCULO AMBIENTAL (25 AÑOS) ---
    def calcular_impacto_ambiental(energia_anual_kwh):
        """
        Calcula toneladas de CO2 y árboles equivalentes proyectados a 25 años.
        """
        FACTOR_EMISION = 0.423
        ABSORCION_ARBOL = 20.0

        co2_evitado_kg = energia_anual_kwh * FACTOR_EMISION
        co2_evitado_ton = co2_evitado_kg / 1000
        arboles = co2_evitado_kg / ABSORCION_ARBOL

        return {
            "co2_25anos": round(co2_evitado_ton * 25, 1),
            "arboles_25anos": int(arboles * 25)
        }

    # --- Helpers existentes ---
    def detecting_estado(direccion, estado_a_numero):
        abreviaturas = list(estado_a_numero.keys())
        abreviaturas_esc = [re.escape(abrev.capitalize()) for abrev in abreviaturas]
        patron = r"\b(" + "|".join(abreviaturas_esc) + r")(\.|$|\s)"
        matches = re.finditer(patron, direccion)
        for match in matches:
            abrev = match.group(1)
            for key in estado_a_numero.keys():
                if key.capitalize() == abrev:
                    return key
        return None

    estado_a_numero = {
        "AGS": 1, "BC": 2, "BCS": 2, "CAMP": 3, "CDMX": 8, "CHIS": 4, "CHIH": 5,
        "COAH": 6, "COL": 7, "DGO": 9, "GTO": 10, "GRO": 11, "HGO": 12, "JAL": 13,
        "MEX": 14, "MICH": 15, "NAY": 16, "NL": 17, "OAX": 18, "PUE": 19, "QRO": 20,
        "QROO": 21, "SLP": 22, "SIN": 23, "SON": 24, "TAMPS": 25, "TLAX": 26, "VER": 27,
        "YUC": 28, "ZAC": 29
    }

    mes_a_num = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
                 "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}

    def abrir_archivo(ruta):
        sistema = platform.system()
        try:
            if sistema == "Windows":
                os.startfile(ruta)
            elif sistema == "Darwin":
                subprocess.call(["open", ruta])
            else:
                subprocess.call(["xdg-open", ruta])
        except Exception as e:
            messagebox.showwarning("Aviso", f"Se generó el archivo, pero no se pudo abrir automáticamente.\n\nError: {e}")

    # --- Lógica principal ---
    pdf_path = filedialog.askopenfilename(title="Selecciona el recibo CFE (Doméstica Mensual)", filetypes=[("Archivos PDF", "*.pdf")])
    if not pdf_path:
        messagebox.showerror("Cancelado", "No se seleccionó ningún archivo PDF.")
        return

    excel_path = r"D:/SECOM/Cotizaciones José/COTIZACION SISTEMA FOTOVOLTAICO MENSUAL.xlsm"

    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
    except Exception as e:
        messagebox.showerror("Error al leer PDF", f"No se pudo procesar el archivo PDF.\n\nError: {e}")
        return

    # --- Verificaciones ---
    tarifa_match = re.search(r"TARIFA: *([1DAC]{1}[A-F]?)", texto)
    tarifa_encontrada = tarifa_match.group(1).strip() if tarifa_match else ""
    tarifas_validas = ["1", "1A", "1B", "1C", "1D", "1E", "1F", "DAC"]
    if tarifa_encontrada not in tarifas_validas:
        messagebox.showerror("Tarifa Incorrecta", f"Recibo es tarifa '{tarifa_encontrada}'. Selecciona el botón correcto.")
        return

    periodo_match = re.search(r"PERIODO FACTURADO:\s*(\d{2}) ([A-Z]{3}) (\d{2})\s*-\s*(\d{2}) ([A-Z]{3}) (\d{2})", texto)
    if not periodo_match:
        messagebox.showerror("Error de Periodo", "No se pudo leer el periodo del PDF.")
        return

    try:
        s_day, s_mon_str, s_yr, e_day, e_mon_str, e_yr = periodo_match.groups()
        start_date = datetime.date(int(s_yr) + 2000, mes_a_num[s_mon_str], int(s_day))
        end_date = datetime.date(int(e_yr) + 2000, mes_a_num[e_mon_str], int(e_day))
        duracion_dias = (end_date - start_date).days
    except Exception as e:
        messagebox.showerror("Error de Fecha", f"Error calculando fechas: {e}")
        return

    if duracion_dias >= 45:
        messagebox.showerror("Periodo Incorrecto", f"Recibo BIMESTRAL ({duracion_dias} días) en script MENSUAL.")
        return

    # --- Extracción de Datos ---
    nombre_match = re.search(r"\b([A-ZÁÉÍÓÚÑ]{3,}(?: [A-ZÁÉÍÓÚÑ]{2,}){1,})\b", texto)
    nombre = nombre_match.group(1) if nombre_match else ""
    nombre = re.split(r"\sTOTAL|\sRFC|\sCUENTA", nombre)[0].strip()
    if not nombre: nombre = "CLIENTE_DESCONOCIDO"

    rpu = re.search(r"NO\.? DE SERVICIO: *(\d+)", texto)
    tarifa = tarifa_encontrada
    hilos = re.search(r"NO HILOS: *(\d+)", texto)

    lineas = texto.splitlines()
    bloque_direccion = []
    encontrado_nombre = False
    excluir = ["TOTAL A PAGAR", "FOTOVOLTAICO", "$", "DESCARGA", "APP", "PESOS", "M.N."]
    for linea in lineas:
        if not encontrado_nombre:
            if nombre and nombre != "CLIENTE_DESCONOCIDO" and nombre in linea:
                encontrado_nombre = True
            continue
        if "NO. DE SERVICIO" in linea: break
        if encontrado_nombre:
            if any(palabra in linea.upper() for palabra in excluir): continue
            bloque_direccion.append(linea.strip())
    direccion = " ".join(bloque_direccion)
    direccion = re.sub(r"\$\s?[\d,]+", "", direccion)
    direccion = re.sub(r"\([^)]+\)", "", direccion).strip()
    estado_detectado = detecting_estado(direccion, estado_a_numero)
    numero_estado = estado_a_numero.get(estado_detectado, "")

    consumo_actual_match = re.search(r"(\d{1,3}[,\d]*)\s+(\d{1,3}[,\d]*)\s+(\d{1,3}[,\d]*)", texto)
    consumo_actual = int(consumo_actual_match.group(3).replace(",", "")) if consumo_actual_match else 0
    pago_actual_match = re.search(r"TOTAL A PAGAR:\s*\$?([\d,]+)", texto)
    pago_actual = float(pago_actual_match.group(1).replace(",", "")) if pago_actual_match else 0.0

    patron_hist = re.compile(r"del \d{2} [A-Z]{3} \d{2} al \d{2} [A-Z]{3} \d{2} (\d+) \$([\d,]+\.\d{2})")
    resultados = patron_hist.findall(texto)
    consumos = [consumo_actual]
    pagos = [pago_actual]
    for kwh, pago in resultados[:11]:
        consumos.insert(0, int(kwh))
        pagos.insert(0, float(pago.replace(",", "")))

    # --- Cálculo Ahorro ---
    suministro_match = re.search(r"Suministro\s+([\d,.]+)", texto)
    suministro_costo = float(suministro_match.group(1).replace(",", "")) if suministro_match else 0.0
    iva_match = re.search(r"IVA\s+(\d+)%", texto)
    iva_porcentaje = int(iva_match.group(1)) if iva_match else 0
    dap_match = re.search(r"(DAP|Derecho de Alumbrado Publico)\S*\s+([\d,.]+)", texto)
    dap_costo = float(dap_match.group(2).replace(",", "")) if dap_match else 0.0
    costo_base = (suministro_costo * (1 + iva_porcentaje / 100)) + dap_costo
    valor_C17 = 0.0
    pagos_reales = [p for p in pagos if p > 0]
    if pagos_reales: valor_C17 = sum(pagos_reales) / len(pagos_reales)
    ahorro_final = valor_C17 - costo_base

    # --- Excel ---
    try:
        wb = openpyxl.load_workbook(excel_path, keep_vba=True)
    except Exception as e:
        messagebox.showerror("Error Crítico", f"No se pudo abrir plantilla Excel.\n{e}")
        return

    h1 = wb["PROMEDIO DE CONSUMO"]
    for i in range(12):
        h1[f"K{4 + i}"] = consumos[i] if i < len(consumos) else ""
        h1[f"K{20 + i}"] = pagos[i] if i < len(pagos) else ""
    h1["C20"] = ahorro_final

    h2 = wb["FORMATO DE COTIZACION"]
    h2["E4"] = nombre
    h2["E5"] = direccion
    h2["G6"] = rpu.group(1) if rpu else ""
    h2["E7"] = tarifa

    # --- INTEGRACIÓN CLIMATE TECH (25 AÑOS) ---
    energia_anual_total = sum(consumos)
    datos_eco = calcular_impacto_ambiental(energia_anual_total)

    # Texto actualizado con NUEVOS EMOJIS
    texto_impacto = (
        f"🌍 IMPACTO AMBIENTAL A 25 años:\n"
        f"🏭 * {datos_eco['co2_25anos']} Toneladas de CO₂ menos\n"
        f"🌳 * Equivale a plantar {datos_eco['arboles_25anos']} árboles"
    )

    # Combinar celdas G21-G23
    h2.merge_cells('G21:G23')
    celda_combinada = h2["G21"]
    celda_combinada.value = texto_impacto

    # Formato visual
    celda_combinada.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    celda_combinada.font = Font(name='Calibri', bold=True, size=19, color="228B22")
    # ------------------------------------------

    h3 = wb["COTIZACIÓN"]
    h3["A11"] = int(hilos.group(1)) if hilos else ""

    h4 = wb["CALCULO DE ENERGIA"]
    h4["C5"] = numero_estado

    nombre_archivo = re.sub(r'[\\/*?:"<>|]', "", nombre)
    nuevo_path = os.path.join(r"D:/SECOM/Cotizaciones José", f"COTIZACION SISTEMA FOTOVOLTAICO DOMESTICA MENSUAL {nombre_archivo}.xlsm")

    try:
        wb.save(nuevo_path)
    except PermissionError:
         base, ext = os.path.splitext(nuevo_path)
         nuevo_path = f"{base}_NUEVO{ext}"
         messagebox.showwarning("Archivo Abierto", f"Archivo original abierto. Guardando como: {nuevo_path}")
         try: wb.save(nuevo_path)
         except Exception as e:
             messagebox.showerror("Error", f"No se pudo guardar (intento 2).\n{e}")
             return
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar Excel.\n{e}")
        return

    # Macro / Gráfico
    try:
        if not os.path.exists(nuevo_path): raise FileNotFoundError(f"No existe: {nuevo_path}")
        excel = win32com.client.Dispatch("Excel.Application")
        wb_com = excel.Workbooks.Open(nuevo_path)
        hoja_rec = wb_com.Sheets("RECUPERACION")
        valor_h34 = hoja_rec.Range("H34").Value or 0
        max_y = int(math.ceil(valor_h34 / 100.0)) * 100
        hoja_graf = wb_com.Sheets("FORMATO DE COTIZACION")
        grafico = hoja_graf.ChartObjects("Chart 1").Chart
        grafico.Axes(2).MaximumScale = max_y
        valor_i34 = hoja_rec.Range("I34").Value
        if valor_i34 and valor_i34 > 0:
            unidad = int(math.ceil(valor_i34 / 100.0)) * 100
            grafico.Axes(2).MajorUnit = unidad
        wb_com.Save()
        wb_com.Close()
        excel.Quit()
    except Exception as e:
        messagebox.showerror("Error Macro", f"Excel guardado, pero falló gráfico.\n{e}")

    abrir_archivo(nuevo_path)

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    procesar_tarifa_domestica_mensual()
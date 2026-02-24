# --- Archivo: tarifa_gdmth.py ---

import re
import pdfplumber
import openpyxl
from tkinter import Tk, filedialog, messagebox
import os
import math
import platform
import subprocess
import win32com.client
from collections import defaultdict

# --- Esta es la función principal que importará tu app_maestro ---
def procesar_tarifa_gdmth():

    # --- Moviendo helpers y constantes DENTRO de la función ---
    estado_a_numero = {
        "AGS": 1, "BC": 2, "BCS": 2, "CAMP": 3, "CDMX": 8, "CHIS": 4, "CHIH": 5,
        "COAH": 6, "COL": 7, "DGO": 9, "GTO": 10, "GRO": 11, "HGO": 12, "JAL": 13,
        "MEX": 14, "MICH": 15, "NAY": 16, "NL": 17, "OAX": 18, "PUE": 19, "QRO": 20,
        "QROO": 21, "SLP": 22, "SIN": 23, "SON": 24, "TAMPS": 25, "TLAX": 26, "VER": 27,
        "YUC": 28, "ZAC": 29
    }

    def detectar_estado(direccion, estado_a_numero):
        patron = r"\b(" + "|".join(re.escape(k) for k in estado_a_numero.keys()) + r")\.?\b"
        match = re.search(patron, direccion.upper())
        if match:
            return match.group(1)
        return None

    orden_meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                   "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
    
    def orden_clave(m):
        partes = m.split()
        if partes[0] not in orden_meses:
            return 0 
        return int(partes[1]) * 100 + orden_meses.index(partes[0])

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

    # === Seleccionar PDF ===
    pdf_path = filedialog.askopenfilename(title="Selecciona el recibo GDMTH en PDF", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        messagebox.showerror("Error", "No se seleccionó ningún archivo.")
        return 

    # --- CAMBIO 1: Ruta de la plantilla base ---
    excel_path = r"D:/SECOM/Cotizaciones José/COTIZACION SISTEMA FOTOVOLTAICO MENSUAL2.xlsm"

    # === Leer PDF ===
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto = "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())
    except Exception as e:
        messagebox.showerror("Error al leer PDF", f"No se pudo procesar el archivo PDF.\n\nError: {e}")
        return 

    # --- Verificación de Tarifa ---
    tarifa_match = re.search(r"TARIFA:\s*([A-Z0-9]+)", texto)
    tarifa_encontrada = tarifa_match.group(1).strip() if tarifa_match else ""
    tarifas_validas = ["GDMTH"]
    if tarifa_encontrada not in tarifas_validas:
        messagebox.showerror(
            "Tarifa Incorrecta",
            f"¡Error! Seleccionaste 'GDMTH', pero este recibo es de tarifa '{tarifa_encontrada}'.\n\n"
            f"Por favor, selecciona el botón correcto para esta tarifa."
        )
        return
    # --- Fin Verificación ---

    # === Extraer nombre del cliente ===
    nombre_match = re.search(r"^(.*?TOTAL A PAGAR)", texto, re.MULTILINE)
    nombre = nombre_match.group(1).replace("TOTAL A PAGAR", "").strip() if nombre_match else "CLIENTE_DESCONOCIDO"

    # === Extraer domicilio ===
    lineas = texto.splitlines()
    direccion = ""
    linea_extra = ""
    for i, linea in enumerate(lineas):
        if nombre and nombre != "CLIENTE_DESCONOCIDO" and nombre in linea:
            direccion = " ".join(lineas[i+1:i+6])
        if "NO. DE SERVICIO" in linea and i >= 1:
            linea_extra = lineas[i-1]
            break
    direccion = f"{direccion} {linea_extra}".strip()
    direccion = re.sub(r"\$\d[\d,.]*", "", direccion)
    direccion = re.sub(r"\([^)]*\)", "", direccion).strip()

    # === Detectar estado y número asociado ===
    estado_abrev = detectar_estado(direccion, estado_a_numero)
    estado_num = estado_a_numero.get(estado_abrev, "")

    # === Extraer número de servicio, tarifa y hilos ===
    rpu = re.search(r"NO\.? DE SERVICIO: *(\d+)", texto)
    tarifa = tarifa_encontrada
    hilos = re.search(r"NO HILOS:\s*(\d+)", texto)

    # === Extraer consumos históricos y precios medios ===
    tabla_hist = re.findall(r"([A-Z]{3} \d{2})\s+[\d,]+\s+([\d,]+)\s+[\d,.]+\s+[\d,.]+\s+([\d.]+)", texto)
    hist_data = defaultdict(lambda: {"kwh": 0, "precio": 0})
    for mes, kwh_str, precio_str in tabla_hist:
        kwh = int(kwh_str.replace(",", ""))
        precio = float(precio_str)
        if mes in hist_data:
            hist_data[mes]["kwh"] += kwh
            hist_data[mes]["precio"] = max(hist_data[mes]["precio"], precio)
        else:
            hist_data[mes]["kwh"] = kwh
            hist_data[mes]["precio"] = precio
    hist_ordenado = sorted(hist_data.items(), key=lambda x: orden_clave(x[0]))[-12:]
    consumos = [v["kwh"] for _, v in hist_ordenado]
    precios = [v["precio"] for _, v in hist_ordenado]
    num_faltantes = 12 - len(consumos)
    if num_faltantes > 0:
        consumos.extend([0] * num_faltantes)
        precios.extend([0] * num_faltantes) 

    # --- INICIO DE CÁLCULO DE AHORRO (LÓGICA v2) ---
    
    # 1. Extraer Suministro (Cargo Fijo)
    suministro_match = re.search(r"(Cargo Fijo|Suministro)\S*\s+([\d,.]+)", texto)
    suministro_costo = float(suministro_match.group(2).replace(",", "")) if suministro_match else 0.0

    # 2. Extraer IVA (ej. 16)
    iva_match = re.search(r"IVA\s+(\d+)%", texto)
    iva_porcentaje = int(iva_match.group(1)) if iva_match else 0

    # 3. Extraer DAP (Regex robusta)
    dap_match = re.search(r"(DAP|Derecho de Alumbrado Publico)\S*\s+([\d,.]+)", texto)
    dap_costo = float(dap_match.group(2).replace(",", "")) if dap_match else 0.0

    # 4. Calcular Costo Base Total (Suministro + IVA + DAP)
    costo_base_total = (suministro_costo * (1 + iva_porcentaje / 100)) + dap_costo

    # --- FIN DE CÁLCULO DE AHORRO ---

    # === Abrir Excel base ===
    try:
        # --- CAMBIO 2: Añadido keep_vba=True ---
        wb = openpyxl.load_workbook(excel_path, keep_vba=True)
    except FileNotFoundError:
        messagebox.showerror("Error Crítico", f"No se encontró la plantilla de Excel base en la ruta:\n{excel_path}\n\nEl programa no puede continuar.")
        return 
    except Exception as e:
        messagebox.showerror("Error Crítico", f"No se pudo abrir la plantilla de Excel base.\n\nError: {e}")
        return 

    # === Hoja PROMEDIO DE CONSUMO ===
    h1 = wb["PROMEDIO DE CONSUMO"]
    for i in range(12):
        h1[f"K{4+i}"] = consumos[i]
        h1[f"L{4+i}"] = precios[i]
    
    # Rellenar columna P usando el dap_costo que encontramos
    for fila in range(4, 16):
        h1[f"P{fila}"] = f"=O{fila}+{dap_costo}"

    # --- NUEVO: Escribir resultados del cálculo de ahorro ---
    h1["C21"] = costo_base_total # Celda de ayuda con el costo fijo
    h1["C20"] = "=C17-C21"       # Fórmula final de Ahorro

    # === Hoja FORMATO DE COTIZACION ===
    h2 = wb["FORMATO DE COTIZACION"]
    h2["E4"] = nombre
    h2["E5"] = direccion
    h2["G6"] = rpu.group(1) if rpu else ""
    h2["E7"] = tarifa 

    # === Hoja COTIZACIÓN ===
    h3 = wb["COTIZACIÓN"]
    h3["A11"] = int(hilos.group(1)) if hilos else ""

    # === Hoja CALCULO DE ENERGIA: insertar número del estado en C5
    h4 = wb["CALCULO DE ENERGIA"]
    h4["C5"] = estado_num

    # === Guardar archivo con nombre personalizado ===
    nombre_archivo = re.sub(r'[\\/*?:"<>|]', "", nombre)
    # --- CAMBIO 3: Ruta del archivo de salida ---
    nuevo_path = os.path.join(r"D:/SECOM/Cotizaciones José", f"COTIZACION SISTEMA FOTOVOLTAICO GDMTH {nombre_archivo}.xlsm")
    
    if os.path.exists(nuevo_path):
        try:
            os.remove(nuevo_path)
        except PermissionError:
            base, ext = os.path.splitext(nuevo_path)
            nuevo_path = f"{base}_NUEVO{ext}"
            messagebox.showwarning("Archivo Abierto", f"El archivo original estaba abierto.\nSe guardará como: {nuevo_path}")

    try:
        wb.save(nuevo_path)
    except Exception as e:
        messagebox.showerror("Error al Guardar", f"No se pudo guardar el archivo Excel.\n\nError: {e}")
        return 

    # === Ajustar gráfico dinámicamente ===
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb_com = excel.Workbooks.Open(nuevo_path)

        # ... (código del gráfico) ...
        hoja_rec = wb_com.Sheets("RECUPERACION")
        valor_h34 = hoja_rec.Range("H34").Value or 0
        max_y = int(math.ceil(valor_h34 / 100.0)) * 100
        hoja_graf = wb_com.Sheets("FORMATO DE COTIZACION")
        grafico = hoja_graf.ChartObjects("Chart 1").Chart
        grafico.Axes(2).MinimumScaleIsAuto = True
        grafico.Axes(2).MaximumScale = max_y
        valor_i34 = hoja_rec.Range("I34").Value
        if valor_i34 and valor_i34 > 0:
            unidad = int(math.ceil(valor_i34 / 100.0)) * 100
            grafico.Axes(2).MajorUnit = unidad

        grafico.Refresh()

        wb_com.Save()
        wb_com.Close()
        excel.Quit()
    except Exception as e:
        messagebox.showerror("Error de Gráfico", f"Se guardó el Excel, pero no se pudo ajustar el gráfico.\n\nError: {e}")

    # === Abrir archivo Excel generado ===
    abrir_archivo(nuevo_path)


# --- Bloque para probar este script de forma individual ---
if __name__ == "__main__":
    root_test = Tk()
    root_test.withdraw()
    procesar_tarifa_gdmth()
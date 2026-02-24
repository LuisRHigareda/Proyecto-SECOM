# --- Archivo: tarifa_pdbt_bimestral.py ---
# (Versión con DOBLE verificación: Tarifa + Periodo)

import re
import pdfplumber
import openpyxl
from tkinter import Tk, filedialog, messagebox
import os
import math
import platform
import subprocess
import win32com.client
import datetime # <--- ¡NUEVA IMPORTACIÓN!

# --- Esta es la función principal que importará tu app_maestro ---
def procesar_tarifa_pdbt_bimestral():

    # --- Moviendo helpers y constantes DENTRO de la función ---
    estado_a_numero = {
        "AGS": 1, "BC": 2, "BCS": 2, "CAMP": 3, "CDMX": 8, "CHIS": 4, "CHIH": 5,
        "COAH": 6, "COL": 7, "DGO": 9, "GTO": 10, "GRO": 11, "HGO": 12, "JAL": 13,
        "MEX": 14, "MICH": 15, "NAY": 16, "NL": 17, "OAX": 18, "PUE": 19, "QRO": 20,
        "QROO": 21, "SLP": 22, "SIN": 23, "SON": 24, "TAMPS": 25, "TLAX": 26, "VER": 27,
        "YUC": 28, "ZAC": 29
    }

    def detectar_estado(direccion):
        for abrev in estado_a_numero:
            patron = rf"\b{abrev}\b|{abrev}\."
            if re.search(patron, direccion.upper()):
                return estado_a_numero[abrev]
        return ""

    # Diccionario para convertir meses de texto a número
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

    # === Seleccionar PDF ===
    pdf_path = filedialog.askopenfilename(title="Selecciona el recibo PDBT BIMESTRAL en PDF", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        messagebox.showerror("Cancelado", "No se seleccionó archivo.")
        return 

    excel_path = r"D:/SECOM/Cotizaciones José/COTIZACION SISTEMA FOTOVOLTAICO BIMESTRAL.xlsm"

    # === Leer PDF (con manejo de errores) ===
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto = "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())
    except Exception as e:
        messagebox.showerror("Error al leer PDF", f"No se pudo procesar el archivo PDF.\n\nError: {e}")
        return

    # --- INICIO DE DOBLE VERIFICACIÓN ---
    
    # 1. Verificar Tarifa
    tarifa_match = re.search(r"TARIFA:\s*([A-Z0-9]+)(?=\s*NO)", texto)
    tarifa_encontrada = tarifa_match.group(1).strip() if tarifa_match else ""
    tarifas_validas = ["PDBT"]
    if tarifa_encontrada not in tarifas_validas:
        messagebox.showerror(
            "Tarifa Incorrecta",
            f"¡Error! Seleccionaste 'PDBT Bimestral', pero este recibo es de tarifa '{tarifa_encontrada}'.\n\n"
            f"Por favor, selecciona el botón correcto para esta tarifa."
        )
        return
        
    # 2. Verificar Periodo
    periodo_match = re.search(r"PERIODO FACTURADO:\s*(\d{2}) ([A-Z]{3}) (\d{2})\s*-\s*(\d{2}) ([A-Z]{3}) (\d{2})", texto)
    
    if not periodo_match:
        # Intentar formato PDBT/DAC (ej. 04 SEP 25-06 OCT 25)
        periodo_match = re.search(r"PERIODO FACTURADO:(\d{2} [A-Z]{3} \d{2})-(\d{2} [A-Z]{3} \d{2})", texto)
        if not periodo_match:
             messagebox.showerror("Error de Periodo", "No se pudo leer el 'PERIODO FACTURADO' del PDF.")
             return
        else:
             # Formato PDBT/DAC (captura diferente)
             s_full, e_full = periodo_match.groups()
             s_day, s_mon_str, s_yr = s_full.split()
             e_day, e_mon_str, e_yr = e_full.split()
    else:
        # Formato estándar (GDMTH/GDMTO)
        s_day, s_mon_str, s_yr, e_day, e_mon_str, e_yr = periodo_match.groups()

    try:
        start_date = datetime.date(int(s_yr) + 2000, mes_a_num[s_mon_str], int(s_day))
        end_date = datetime.date(int(e_yr) + 2000, mes_a_num[e_mon_str], int(e_day))
        duracion_dias = (end_date - start_date).days
    except Exception as e:
        messagebox.showerror("Error de Fecha", f"No se pudo calcular la duración del periodo.\nTexto encontrado: {periodo_match.groups()}\nError: {e}")
        return

    # 3. Aplicar Lógica (Este es el script BIMESTRAL)
    if duracion_dias < 45: # 45 es un umbral seguro para "mensual"
        messagebox.showerror(
            "Periodo Incorrecto",
            f"¡Error! Este recibo es MENSUAL ({duracion_dias} días).\n\n"
            f"Estás usando el script para recibos BIMESTRALES."
        )
        return
    # --- FIN DE DOBLE VERIFICACIÓN ---


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
    direccion = re.sub(r"(CD OBREGON,Son\.)\s+\1", r"\1", direccion) 

    # === Extraer número de servicio, tarifa y hilos ===
    rpu = re.search(r"NO\.? DE SERVICIO: *(\d+)", texto)
    tarifa = tarifa_encontrada
    hilos = re.search(r"NO HILOS:\s*(\d+)", texto)

    # === Extraer consumo actual ===
    consumo_actual = re.search(r"Energía \(kWh\).*?(\d{1,3}(?:,\d{3})*?)\s*$", texto, re.MULTILINE)
    kwh_total = int(consumo_actual.group(1).replace(",", "")) if consumo_actual else 0

    # === Extraer costo total actual ===
    costo_actual = 0.0
    costo_match = re.search(r"TOTAL A PAGAR[:\s]*\$\s*([\d,]+(?:\.\d{2})?)", texto)
    if not costo_match:
        barra_match = re.search(r"\n\$([\d,]+\.\d{2})", texto)
        if barra_match:
            costo_actual = float(barra_match.group(1).replace(",", ""))
    else:
        costo_actual = float(costo_match.group(1).replace(",", ""))

    # === Extraer PRIMEROS 5 consumos históricos ===
    tabla_hist = re.findall(r"\bdel\b.*?(\d{1,4})\s+\$(\d[\d,]*\.\d{2})", texto)
    consumos = [kwh_total]
    precios = [costo_actual] # 'precios' y 'pagos' son lo mismo en este script
    for entrada in tabla_hist[:5]:
        consumos.append(int(entrada[0].replace(",", "")))
        precios.append(float(entrada[1].replace(",", "")))
    while len(consumos) < 6:
        consumos.append(0)
        precios.append(0.0)

    # --- INICIO DE CÁLCULO DE AHORRO ---
    suministro_match = re.search(r"(Suministro|Cargo Fijo)\S*\s+([\d,.]+)", texto)
    suministro_costo = float(suministro_match.group(2).replace(",", "")) if suministro_match else 0.0
    iva_match = re.search(r"IVA\s+(\d+)%", texto)
    iva_porcentaje = int(iva_match.group(1)) if iva_match else 0
    dap_match = re.search(r"(DAP|Derecho de Alumbrado Publico)\S*\s+([\d,.]+)", texto)
    dap_costo = float(dap_match.group(2).replace(",", "")) if dap_match else 0.0
    costo_base_bimestral = (suministro_costo * (1 + iva_porcentaje / 100)) + dap_costo
    valor_promedio_bimestral = 0.0
    pagos_reales = [p for p in precios if p > 0] 
    if pagos_reales: 
        valor_promedio_bimestral = sum(pagos_reales) / len(pagos_reales)
    ahorro_final_bimestral = valor_promedio_bimestral - costo_base_bimestral
    ahorro_final_mensual = ahorro_final_bimestral / 2
    # --- FIN DE CÁLCULO DE AHORRO ---


    # === Abrir Excel (con manejo de errores) ===
    try:
        wb = openpyxl.load_workbook(excel_path, keep_vba=True)
    except FileNotFoundError:
        messagebox.showerror("Error Crítico", f"No se encontró la plantilla de Excel base en la ruta:\n{excel_path}\n\nEl programa no puede continuar.")
        return
    except Exception as e:
        messagebox.showerror("Error Crítico", f"No se pudo abrir la plantilla de Excel base.\n\nError: {e}")
        return

    # === PROMEDIO DE CONSUMO ===
    h1 = wb["PROMEDIO DE CONSUMO"]
    for i in range(6):
        h1[f"I{3+i}"] = consumos[i]
        h1[f"I{19+i}"] = precios[i]
        h1[f"N{3+i}"] = f"=M{3+i}+{dap_costo}"
    h1["C20"] = ahorro_final_mensual

    # === FORMATO DE COTIZACION ===
    h2 = wb["FORMATO DE COTIZACION"]
    h2["E5"] = nombre
    h2["E6"] = direccion
    h2["G7"] = rpu.group(1) if rpu else ""
    h2["E8"] = tarifa

    # === COTIZACIÓN ===
    h3 = wb["COTIZACIÓN"]
    h3["A9"] = int(hilos.group(1)) if hilos else ""

    # === CALCULO DE ENERGIA: estado ===
    estado_num = detectar_estado(direccion)
    wb["CALCULO DE ENERGIA"]["C5"] = estado_num

    # === Guardar archivo Excel final ===
    nombre_archivo = re.sub(r'[\\/*?:"<>|]', "", nombre)
    nuevo_path = os.path.join(r"D:/SECOM/Cotizaciones José", f"COTIZACION SISTEMA FOTOVOLTAICO PDBT BIMESTRAL {nombre_archivo}.xlsm")
    
    try:
        wb.save(nuevo_path)
    except PermissionError:
         base, ext = os.path.splitext(nuevo_path)
         nuevo_path = f"{base}_NUEVO{ext}"
         messagebox.showwarning("Archivo Abierto", f"El archivo original estaba abierto.\nSe guardará como: {nuevo_path}")
         try:
             wb.save(nuevo_path)
         except Exception as e:
             messagebox.showerror("Error al Guardar", f"No se pudo guardar el archivo Excel (intento 2).\n\nError: {e}")
             return
    except Exception as e:
        messagebox.showerror("Error al Guardar", f"No se pudo guardar el archivo Excel.\n\nError: {e}")
        return
        
    # === Ajustar gráfico ===
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb_com = excel.Workbooks.Open(nuevo_path)

        hoja_rec = wb_com.Sheets("RECUPERACION")
        valor_h29 = hoja_rec.Range("H29").Value or 0
        max_y = int(math.ceil(valor_h29 / 100.0)) * 100

        hoja_graf = wb_com.Sheets("FORMATO DE COTIZACION")
        grafico = hoja_graf.ChartObjects("Chart 1").Chart
        grafico.Axes(2).MaximumScale = max_y

        valor_i29 = hoja_rec.Range("I29").Value
        if valor_i29 and valor_i29 > 0:
            unidad = int(math.ceil(valor_i29 / 100.0)) * 100
            grafico.Axes(2).MajorUnit = unidad

        wb_com.Save()
        wb_com.Close()
        excel.Quit()
    except Exception as e:
        messagebox.showerror("Error de Gráfico/Macro", f"Se guardó el Excel, pero no se pudo ajustar el gráfico.\n\nError: {e}")

    # === Abrir archivo generado ===
    abrir_archivo(nuevo_path)


# --- Bloque para probar este script de forma individual ---
if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    procesar_tarifa_pdbt_bimestral()
# --- Archivo: subir_datos_cashvolt.py ---
# (Versión con redondeo de Ahorro y Costo)

import openpyxl
from tkinter import Tk, filedialog, messagebox, simpledialog
import time
import sys
import os

# --- Importaciones de Selenium ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains

# --- Función para encontrar el logo ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- MAPA DE CELDAS ---
MAPA_MENSUAL = {
    "cliente": ("FORMATO DE COTIZACION", "E4"),
    "ahorro": ("PROMEDIO DE CONSUMO", "C20"),
    "paneles_qty": ("CALCULO DE ENERGIA", "C33"),
    "paneles_cap": ("CALCULO DE ENERGIA", "C9"),
    "costo_proy": ("FORMATO DE COTIZACION", "K23")
}
MAPA_BIMESTRAL = {
    "cliente": ("FORMATO DE COTIZACION", "E5"),
    "ahorro": ("PROMEDIO DE CONSUMO", "C20"),
    "paneles_qty": ("CALCULO DE ENERGIA", "C29"),
    "paneles_cap": ("CALCULO DE ENERGIA", "C9"),
    "costo_proy": ("FORMATO DE COTIZACION", "K24")
}

# --- Función helper para encontrar y llenar campos ---
def find_and_fill(driver, label_text, value_to_send):
    try:
        xpath = f"//*[contains(text(), '{label_text}')]/following::input[1]"
        element = driver.find_element(By.XPATH, xpath)
        
        ActionChains(driver).scroll_to_element(element).perform()
        time.sleep(0.3) 
        
        element.send_keys(value_to_send)
        time.sleep(0.3) 
    except Exception as e:
        try:
            xpath_label = f"//label[contains(text(), '{label_text}')]/following::input[1]"
            element = driver.find_element(By.XPATH, xpath_label)
            ActionChains(driver).scroll_to_element(element).perform()
            time.sleep(0.3)
            element.send_keys(value_to_send)
            time.sleep(0.3)
        except Exception as inner_e:
            print(f"No se pudo encontrar el campo para: {label_text}")
            raise inner_e 


# --- ESTA ES LA FUNCIÓN PRINCIPAL QUE IMPORTARÁ TU APP_MAESTRO ---
def subir_datos_cashvolt():

    # --- 1. PEDIR ARCHIVO EXCEL ---
    excel_path = filedialog.askopenfilename(
        title="Selecciona el archivo .xlsm final",
        filetypes=[("Archivos Excel", "*.xlsm")]
    )
    if not excel_path:
        return

    # --- 2. CREDENCIALES (¡DATOS FIJOS!) ---
    user_email = "1571SECOMENERGIAAdmin"
    user_pass = "CV2083"

    # --- 3. LEER DATOS DEL EXCEL (CON LÓGICA DE MAPEO) ---
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "BIMESTRAL" in excel_path.upper():
            mapa_celdas = MAPA_BIMESTRAL
        else:
            mapa_celdas = MAPA_MENSUAL
        
        def get_cell_value(sheet_name, cell_ref):
            if sheet_name not in wb.sheetnames:
                raise KeyError(f"La hoja '{sheet_name}' no existe en el archivo Excel.")
            
            val = wb[sheet_name][cell_ref].value
            if val is None:
                return "0" 
            return str(val)

        asesor_nombre = "Jorge Alejandro Díaz Gaxiola"
        cliente_nombre = get_cell_value(mapa_celdas["cliente"][0], mapa_celdas["cliente"][1])
        
        # --- ¡CAMBIOS AQUÍ! ---
        # 1. Redondea el ahorro al entero más cercano
        ahorro_monto = round(float(get_cell_value(mapa_celdas["ahorro"][0], mapa_celdas["ahorro"][1])))
        
        paneles_qty = get_cell_value(mapa_celdas["paneles_qty"][0], mapa_celdas["paneles_qty"][1])
        paneles_cap = get_cell_value(mapa_celdas["paneles_cap"][0], mapa_celdas["paneles_cap"][1])
        
        # 2. Redondea el costo al entero más cercano (corrige el error de 109200.0000...1)
        costo_proy = round(float(get_cell_value(mapa_celdas["costo_proy"][0], mapa_celdas["costo_proy"][1])))


        datos = {
            "Asesor de ventas:": str(asesor_nombre),
            "Nombre del cliente:": str(cliente_nombre),
            "Ahorro mensual del proyecto:": str(ahorro_monto), # Ahora es un entero
            "Cantidad de Paneles:": str(paneles_qty),
            "Capacidad del Panel:": str(paneles_cap),
            "Costo del proyecto:": str(costo_proy)   # Ahora es un entero
        }

    except KeyError as e:
        messagebox.showerror("Error de Hoja/Celda", f"{e}\n\n"
                             f"Por favor, revisa que los mapas de celdas (MAPA_MENSUAL, MAPA_BIMESTRAL) "
                             f"en el script 'subir_datos_cashvolt.py' sean correctos.")
        return
    except Exception as e:
        messagebox.showerror("Error al leer Excel", f"No se pudo leer el archivo .xlsm.\n"
                             f"Asegúrate de que las celdas son correctas en los mapas.\n\nError: {e}")
        return

    # --- 4. INICIAR EL ROBOT (SELENIUM) ---
    # (Popup eliminado)
    
    try:
        s = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s)
        wait = WebDriverWait(driver, 15)

        # --- 5. PASO DE LOGIN (OPTIMIZADO) ---
        driver.get("https://cashvolt.mx/public/login")
        
        wait.until(EC.visibility_of_element_located((By.ID, "username")))
        driver.find_element(By.ID, "username").send_keys(user_email)
        driver.find_element(By.ID, "password").send_keys(user_pass)
        driver.find_element(By.XPATH, "//button[contains(., 'Ingresar')]").click()
        wait.until(EC.url_contains("/admin/menu"))

        # --- 6. PASO DE LLENAR DATOS (CON TEXTO EXACTO) ---
        wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Asesor de ventas:')]")))
        
        find_and_fill(driver, "Asesor de ventas:", datos["Asesor de ventas:"])
        find_and_fill(driver, "Nombre del cliente:", datos["Nombre del cliente:"])
        find_and_fill(driver, "Ahorro mensual del proyecto:", datos["Ahorro mensual del proyecto:"])
        find_and_fill(driver, "Cantidad de Paneles:", datos["Cantidad de Paneles:"])
        find_and_fill(driver, "Capacidad del Panel:", datos["Capacidad del Panel:"])
        find_and_fill(driver, "Costo del proyecto:", datos["Costo del proyecto:"])

        # --- "¡Éxito!" ELIMINADO y REEMPLAZADO por PAUSA ---
        while True:
            time.sleep(5)
            
    except Exception as e:
        print(f"El robot terminó o encontró un error: {e}")
    
# --- Bloque para probar este script de forma individual ---
if __name__ == "__main__":
    subir_datos_cashvolt()
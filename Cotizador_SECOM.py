# --- Archivo: Cotizador_SECOM.py ---

# (Versión con cierre de menú corregido)



import tkinter as tk

from tkinter import messagebox, font

import sys

import os



# --- FUNCIÓN ESPECIAL PARA ENCONTRAR ARCHIVOS EN EL .EXE ---

def resource_path(relative_path):

    try:

        base_path = sys._MEIPASS

    except Exception:

        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# -----------------------------------------------------------





# --- 1. IMPORTA LAS 7 FUNCIONES ---

try:

    from tarifa_domestica_mensual import procesar_tarifa_domestica_mensual

    from tarifa_domestica_bimestral import procesar_tarifa_domestica_bimestral

    from tarifa_pdbt import procesar_tarifa_pdbt

    from tarifa_pdbt_bimestral import procesar_tarifa_pdbt_bimestral

    from tarifa_gdmth import procesar_tarifa_gdmth

    from tarifa_gdmto import procesar_tarifa_gdmto 

    from subir_datos_cashvolt import subir_datos_cashvolt



except ImportError as e:

    messagebox.showerror(

        "Error Crítico de Módulo",

        f"No se pudo encontrar un script de tarifa o el script de CashVolt.\n\n"

        f"Asegúrate de que los 7 archivos .py estén en la misma carpeta.\n\n"

        f"Error: {e}"

    )

    sys.exit()





# --- 2. FUNCIÓN DE AYUDA PARA EJECUTAR UN SCRIPT (¡CORREGIDA!) ---

def ejecutar_procesador(script_function_to_run):

    

    # --- ¡CAMBIO DE LÓGICA! ---

    # Si es el script de la web (CashVolt)...

    if script_function_to_run == subir_datos_cashvolt:

        root.destroy() # 1. Cierra el menú INMEDIATAMENTE.

        try:

            script_function_to_run() # 2. Lanza el robot (que se quedará corriendo).

        except Exception as e:

            messagebox.showerror("Error Inesperado", f"Ocurrió un error inesperado durante la ejecución:\n\n{e}")

        # El programa principal (menú) ya murió, el robot sigue vivo.

    

    # Si es un script de Excel (lógica anterior)...

    else:

        try:

            root.withdraw() # 1. Oculta el menú.

            script_function_to_run() # 2. Lanza el script de Excel.

            

        except Exception as e:

            messagebox.showerror("Error Inesperado", f"Ocurrió un error inesperado durante la ejecución:\n\n{e}")

        finally:

            # 3. Cuando el script de Excel termina, cierra el menú (que estaba oculto).

            root.destroy()





# --- 3. CREAR LA VENTANA DEL MENÚ ---

root = tk.Tk()

root.title("Selector de Cotizaciones CFE")



try:

    logo_path = resource_path("logo.ico")

    root.iconbitmap(logo_path)

except tk.TclError:

    print("Advertencia: No se pudo encontrar 'logo.ico'. Se usará el ícono por defecto.")



root.geometry("400x520") 

root.resizable(False, False) 



font_titulo = font.Font(family="Arial", size=14, weight="bold")

font_boton = font.Font(family="Arial", size=12)



# --- 4. TÍTULO Y BOTONES ---

label = tk.Label(root, text="Selecciona el tipo de tarifa:", font=font_titulo)

label.pack(pady=20, padx=20) 



frame_botones = tk.Frame(root)

frame_botones.pack(pady=10, padx=30, fill="x") 



btn_1 = tk.Button(frame_botones, text="Doméstica Mensual", font=font_boton, height=2, 

                  command=lambda: ejecutar_procesador(procesar_tarifa_domestica_mensual))

btn_1.pack(fill="x", pady=5) 

btn_2 = tk.Button(frame_botones, text="Doméstica Bimestral", font=font_boton, height=2, 

                  command=lambda: ejecutar_procesador(procesar_tarifa_domestica_bimestral))

btn_2.pack(fill="x", pady=5)

btn_3 = tk.Button(frame_botones, text="PDBT Mensual", font=font_boton, height=2, 

                  command=lambda: ejecutar_procesador(procesar_tarifa_pdbt))

btn_3.pack(fill="x", pady=5)

btn_4 = tk.Button(frame_botones, text="PDBT Bimestral", font=font_boton, height=2, 

                  command=lambda: ejecutar_procesador(procesar_tarifa_pdbt_bimestral))

btn_4.pack(fill="x", pady=5)

btn_5 = tk.Button(frame_botones, text="GDMTH", font=font_boton, height=2, 

                  command=lambda: ejecutar_procesador(procesar_tarifa_gdmth))

btn_5.pack(fill="x", pady=5)

btn_6 = tk.Button(frame_botones, text="GDMTO", font=font_boton, height=2, 

                  command=lambda: ejecutar_procesador(procesar_tarifa_gdmto))

btn_6.pack(fill="x", pady=5)



btn_7 = tk.Button(frame_botones, text="Subir Datos a CashVolt", font=font_boton, height=2, 

                  bg="#15803d", fg="white",

                  command=lambda: ejecutar_procesador(subir_datos_cashvolt))

btn_7.pack(fill="x", pady=10, ipady=4) 





# --- 5. INICIAR LA APLICACIÓN ---

root.mainloop()
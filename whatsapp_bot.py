import time
import pandas as pd
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyperclip
import unicodedata

# Diccionario de meses para conversión
MESES = {
    "01": "ENERO", "02": "FEBRERO", "03": "MARZO", "04": "ABRIL",
    "05": "MAYO", "06": "JUNIO", "07": "JULIO", "08": "AGOSTO",
    "09": "SEPTIEMBRE", "10": "OCTUBRE", "11": "NOVIEMBRE", "12": "DICIEMBRE"
}

# Función para formatear el concepto
def formatear_concepto(concepto, año):
    if isinstance(concepto, str) and len(concepto) > 3:
        base = concepto[:-2]  # Extraer la parte textual (ej: 'LUZ', 'ALQ')
        mes_num = concepto[-2:]  # Extraer los últimos dos dígitos (ej: '01', '02')
        
        if mes_num in MESES:
            if base.upper() == "LUZ":
                return f"LUZ {MESES[mes_num].lower()}{año}"
            elif base.upper() == "ALQ":
                return f"ALQ {MESES[mes_num]}{año}"
    return concepto  # Si no coincide, devolver sin cambios

# Cargar datos desde el archivo de Excel
def cargar_datos_excel(ruta_archivo, nombre_hoja):
    df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, engine="openpyxl")
    df.columns = df.columns.str.strip().str.upper()
    return df

# Detectar nuevas entradas en la tabla PAGOS_DE_ALQ desde hoy
def detectar_nuevas_entradas(df_pagos, historial_path):
    hoy = datetime.today().date()
    if "FECHA" not in df_pagos.columns:
        print("Error: La columna 'FECHA' no existe en el archivo. Columnas encontradas:", df_pagos.columns)
        return pd.DataFrame()
    df_pagos["FECHA"] = pd.to_datetime(df_pagos["FECHA"], errors='coerce').dt.date
    df_pagos = df_pagos.dropna(subset=["FECHA"])
    df_pagos_hoy = df_pagos[df_pagos["FECHA"] >= hoy]
    
    if os.path.exists(historial_path):
        df_historial = pd.read_excel(historial_path, engine="openpyxl")
        df_historial.columns = df_historial.columns.str.strip().str.upper()
        df_nuevas = df_pagos_hoy[~df_pagos_hoy.apply(tuple, axis=1).isin(df_historial.apply(tuple, axis=1))]
    else:
        df_nuevas = df_pagos_hoy
    
    return df_nuevas

# Obtener grupo desde la hoja LISTAINQ
def obtener_grupo(df_listainq, dpto_comerc):
    if "DPTO" not in df_listainq.columns:
        print("Error: No se encontró la columna 'DPTO' en LISTAINQ. Columnas disponibles:", df_listainq.columns)
        return None
    
    grupo = df_listainq.loc[df_listainq["DPTO"] == dpto_comerc, 'WA']
    return grupo.iloc[0] if not grupo.empty else None

# Normalizar texto para evitar caracteres incompatibles
def normalizar_texto(texto):
    if isinstance(texto, str):
        return unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    return texto

# Generar mensaje automático
def generar_mensaje(row):
    concepto_formateado = formatear_concepto(row['CONCEPTO'], str(int(row['AÑO'])))
    monto_formateado = f"Gs. {int(row['MONTO']):,}.-".replace(",", ".")
    
    return (f"📢 *Pago registrado* 📢\n\n"
            f"📅 *Fecha:* {row['FECHA']}\n"
            f"🧾 *Comprobante:* {row['COMPROBANTE']}\n"
            f"📆 *Concepto:* {concepto_formateado}\n"
            f"💰 *Monto:* {monto_formateado}\n\n"
            "A facturar")

# Configurar el navegador Selenium
def iniciar_whatsapp():
    opciones = webdriver.ChromeOptions()
    opciones.add_argument("--user-data-dir=C:/Users/tu_usuario/AppData/Local/Google/Chrome/User Data")
    opciones.add_argument("--profile-directory=Default")
    opciones.add_argument("--disable-dev-shm-usage")
    opciones.add_argument("--disable-gpu")
    opciones.add_argument("--no-sandbox")
    opciones.add_argument("--remote-debugging-port=9222")
    opciones.add_argument("--start-maximized")
    
    driver = webdriver.Chrome(options=opciones)
    driver.get("https://web.whatsapp.com")
    print("Esperando a que WhatsApp Web cargue completamente...")
    WebDriverWait(driver, 90).until(EC.presence_of_element_located((By.XPATH, "//div[@id='side']")))
    print("WhatsApp Web cargado correctamente.")
    return driver

# Cargar datos y ejecutar envíos
def main():
    ruta_excel = r"C:\Users\seiko\Mi unidad\ALQUILERES\ALQ PAGOS.xlsm"
    hoja_pagos = "ALQ PAGOS"
    hoja_listainq = "LISTAINQ"
    historial_path = "HISTORIAL_ENVIOS.xlsx"
    
    df_pagos = cargar_datos_excel(ruta_excel, hoja_pagos)
    df_listainq = cargar_datos_excel(ruta_excel, hoja_listainq)
    df_nuevas = detectar_nuevas_entradas(df_pagos, historial_path)
    
    if df_nuevas.empty:
        print("No hay nuevos pagos para enviar mensajes.")
        return
    
    driver = iniciar_whatsapp()
    
    for _, row in df_nuevas.iterrows():
        nombre_grupo = obtener_grupo(df_listainq, row['DPTO/COMERC'])
        nombre_grupo = normalizar_texto(nombre_grupo)
        mensaje = generar_mensaje(row)
        pyperclip.copy(mensaje)
        time.sleep(2)
    
    df_nuevas.to_excel(historial_path, index=False, engine="openpyxl")
    driver.quit()
    print("Mensajes enviados correctamente y historial actualizado.")

if __name__ == "__main__":
    main()

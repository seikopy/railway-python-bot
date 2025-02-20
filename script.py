from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import os
import pandas as pd
from datetime import datetime

# Verificar si win32com está disponible
try:
    import win32com.client
except ImportError:
    print("win32com no está disponible en este sistema operativo.")
    win32com = None

# Carpeta en Outlook donde se almacenan los correos de Banco Atlas
FOLDER_PATH = "BANCOS - BANCARD OTROS\\ATLAS"

# Grupo de WhatsApp al que se enviarán los mensajes
WHATSAPP_GROUP = "VENUS TRANSFERENCIAS"

# Número de cuenta de la peluquería
CUENTA_PELUQUERIA = "1272612"

# Archivo de logs
LOG_FILE = "logs.txt"

def escribir_log(mensaje):
    """Guarda un mensaje en el archivo de logs con la fecha y hora."""
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(f"{datetime.now()} - {mensaje}\n")

def leer_correos():
    """Lee los correos en la carpeta 'ATLAS' y procesa las transferencias"""
    if not win32com:
        escribir_log("Error: No se puede acceder a Outlook en este sistema.")
        return None

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Acceder a la carpeta específica
    try:
        folder_bancos = outlook.Folders.Item("Bandeja de entrada").Folders.Item("BANCOS - BANCARD OTROS")
        folder_atlas = folder_bancos.Folders.Item("ATLAS")
        mensajes = folder_atlas.Items
    except Exception as e:
        escribir_log(f"Error al acceder a la carpeta: {e}")
        return None

    for mensaje in mensajes:
        if any(keyword in mensaje.Subject for keyword in ["AVISO DE TRANSFERENCIAS INTERBANCARIAS", "NOTIFICACION DE TRANSFERENCIAS INTERBANCARIAS"]):
            cuerpo = mensaje.Body
            # **Verificamos que el correo contenga la cuenta de la peluquería**
            if f"Cuenta Corriente: {CUENTA_PELUQUERIA}" in cuerpo:
                escribir_log("✅ Transferencia detectada. Procesando datos...")
                return extraer_datos(cuerpo)

    escribir_log("⚠️ No se encontró ninguna transferencia nueva.")
    return None

def extraer_datos(texto):
    """Extrae los datos clave del correo"""
    lineas = texto.split("\n")
    datos = {}

    for linea in lineas:
        if "Enviado por:" in linea:
            datos["enviado_por"] = linea.split(":")[1].strip()
        elif "Monto Crédito:" in linea:
            datos["monto"] = linea.split(":")[1].strip()
        elif "Banco Origen:" in linea:
            datos["banco_origen"] = linea.split(":")[1].strip()
        elif "Fecha:" in linea:
            datos["fecha"] = linea.split(":")[1].strip()
        elif "Hora:" in linea:
            datos["hora"] = linea.split(":")[1].strip()
        elif "Nro. Operación SIPAP:" in linea:
            datos["comprobante"] = linea.split(":")[1].strip()

    if len(datos) == 6:
        escribir_log(f"📄 Datos extraídos correctamente: {datos}")
        return datos
    else:
        escribir_log("⚠️ Error: Datos incompletos en la transferencia.")
        return None

def enviar_whatsapp(datos):
    """Envía el mensaje de transferencia al grupo de WhatsApp"""
    if not datos:
        escribir_log("⚠️ No hay datos de transferencia para enviar.")
        return

    mensaje = f"""✅ *RECIBIDO TRANSFERENCIA* ✅

👤 *Enviado por:* {datos["enviado_por"]}
💰 *Monto:* {datos["monto"]}
🏦 *Banco Origen:* {datos["banco_origen"]}
📅 *Fecha:* {datos["fecha"]}
⏰ *Hora:* {datos["hora"]}
📌 *Comprobante:* {datos["comprobante"]}

Reaccionar con 👍 este mensaje la sucursal que corresponde esta transferencia."""

    # Abre WhatsApp Web y envía el mensaje al grupo
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com/")
    input("Escanea el código QR y presiona Enter...")

    try:
        search_box = driver.find_element(By.XPATH, "//div[@title='Buscar o empezar un chat']")
        search_box.send_keys(WHATSAPP_GROUP)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        message_box = driver.find_element(By.XPATH, "//div[@title='Escribe un mensaje aquí']")
        message_box.send_keys(mensaje)
        message_box.send_keys(Keys.ENTER)
        time.sleep(2)

        escribir_log("✅ Mensaje enviado correctamente a WhatsApp.")

    except Exception as e:
        escribir_log(f"⚠️ Error al enviar el mensaje a WhatsApp: {e}")

    finally:
        driver.quit()

if __name__ == "__main__":
    escribir_log("🔄 Iniciando proceso de detección de transferencias...")
    datos_transferencia = leer_correos()
    if datos_transferencia:
        enviar_whatsapp(datos_transferencia)
    else:
        escribir_log("🚫 No se enviaron mensajes a WhatsApp.")

try:
    import win32com.client
except ImportError:
    print("win32com no está disponible en este sistema operativo.")
import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Archivo Excel para registrar pagos de inquilinos
EXCEL_FILE = "pagos_inquilinos.xlsx"

def leer_correos():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        bandeja_entrada = outlook.GetDefaultFolder(6)
        mensajes = bandeja_entrada.Items

        for mensaje in mensajes:
            if mensaje.SenderEmailAddress == "notificaciones@bancoatlas.com.py" and "transferencia" in mensaje.Subject.lower():
                return extraer_datos(mensaje.Body)
    except NameError:
        print("No se puede acceder a Outlook en este sistema.")
        return None


def extraer_datos(texto):
    lineas = texto.split("\n")
    datos = {}
    for linea in lineas:
        if "Monto" in linea:
            datos["Monto"] = linea.split(":")[1].strip()
        elif "Fecha" in linea:
            datos["Fecha"] = linea.split(":")[1].strip()
        elif "Remitente" in linea:
            datos["Remitente"] = linea.split(":")[1].strip()
        elif "Cuenta destino" in linea:
            datos["Cuenta"] = linea.split(":")[1].strip()
    return datos

def guardar_en_excel(datos):
    df = pd.DataFrame([datos])
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df_existente, df], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

def enviar_whatsapp(datos):
    mensaje = f"💰 Transferencia recibida 💰\n\nMonto: {datos['Monto']}\nFecha: {datos['Fecha']}\nRemitente: {datos['Remitente']}"
    
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com")
    input("Escanea el código QR y presiona Enter...")

    time.sleep(5)
    grupo = "Grupo del Edificio"
    
    search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
    search_box.send_keys(grupo)
    search_box.send_keys(Keys.ENTER)
    
    time.sleep(2)
    msg_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
    msg_box.send_keys(mensaje)
    msg_box.send_keys(Keys.ENTER)
    
    time.sleep(2)
    driver.quit()

def ejecutar():
    datos = leer_correos()
    if datos:
        if datos["Cuenta"] == "Cuenta Banco 1 (Peluquería)":
            enviar_whatsapp(datos)
        elif datos["Cuenta"] == "Cuenta Banco 2 (Inquilinos)":
            guardar_en_excel(datos)
    else:
        print("No hay transferencias nuevas.")

if __name__ == "__main__":
    ejecutar()

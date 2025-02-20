import os
import imaplib
import email
from email.header import decode_header
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Cargar variables de entorno
load_dotenv()

# Configuración de credenciales
EMAIL_USER = os.getenv("OUTLOOK_USER")
EMAIL_PASS = os.getenv("OUTLOOK_PASSWORD")
IMAP_SERVER = "outlook.office365.com"
FOLDER_PATH = "BANCOS - BANCARD OTROS/ATLAS"
WHATSAPP_GROUP = "VENUS TRANSFERENCIAS"

# Conectar a Outlook vía IMAP
def conectar_outlook():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select(FOLDER_PATH)
        return mail
    except Exception as e:
        print("Error al conectar con Outlook:", str(e))
        return None

# Buscar correos relevantes
def buscar_correos():
    mail = conectar_outlook()
    if not mail:
        print("No se pudo conectar a Outlook.")
        return []
    
    try:
        status, messages = mail.search(None, 'ALL')
        correos = messages[0].split()
        resultados = []
        
        for num in correos[-10:]:  # Revisar los últimos 10 correos
            status, data = mail.fetch(num, '(RFC822)')
            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes) and encoding:
                        subject = subject.decode(encoding)
                    
                    if "TRANSFERENCIAS" in subject.upper() or "BANCO ATLAS - AVISO DE TRANSFERENCIAS INTERBANCARIAS" in subject.upper():
                        cuerpo = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    cuerpo = part.get_payload(decode=True).decode()
                                    break
                        else:
                            cuerpo = msg.get_payload(decode=True).decode()
                        
                        if "Cuenta Corriente: 1272612" in cuerpo:
                            resultados.append((subject, cuerpo))
        return resultados
    except Exception as e:
        print("Error al buscar correos:", str(e))
        return []

# Extraer datos relevantes
def extraer_datos(texto):
    lineas = texto.split("\n")
    datos = {}
    
    for linea in lineas:
        if "Enviado por:" in linea:
            datos["Enviado por"] = linea.split(":")[-1].strip()
        elif "Monto Crédito:" in linea:
            datos["Monto"] = linea.split(":")[-1].strip()
        elif "Banco Origen:" in linea:
            datos["Banco Origen"] = linea.split(":")[-1].strip()
        elif "Nro. Operación SIPAP:" in linea:
            datos["Comprobante"] = linea.split(":")[-1].strip()
    return datos

# Enviar mensaje a WhatsApp
def enviar_whatsapp(datos):
    mensaje = (f"*RECIBIDO TRANSFERENCIA*\n"
               f"Enviado por: {datos.get('Enviado por', 'N/A')}\n"
               f"Monto: {datos.get('Monto', 'N/A')}\n"
               f"Banco Origen: {datos.get('Banco Origen', 'N/A')}\n"
               f"Comprobante: {datos.get('Comprobante', 'N/A')}\n\n"
               "Reaccionar con 👍 este mensaje, la sucursal que corresponde esta transferencia.")
    
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com")
    input("Escanea el código QR y presiona Enter para continuar...")
    
    time.sleep(10)
    
    search_box = driver.find_element(By.XPATH, "//div[contains(@class,'copyable-text selectable-text')]")
    search_box.send_keys(WHATSAPP_GROUP + Keys.ENTER)
    time.sleep(5)
    
    message_box = driver.find_elements(By.XPATH, "//div[contains(@class,'copyable-text selectable-text')]")[-1]
    message_box.send_keys(mensaje + Keys.ENTER)
    
    time.sleep(5)
    driver.quit()

# Ejecutar script
def main():
    correos = buscar_correos()
    for _, cuerpo in correos:
        datos = extraer_datos(cuerpo)
        enviar_whatsapp(datos)
    print("Proceso completado.")

if __name__ == "__main__":
    main()

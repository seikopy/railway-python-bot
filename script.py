import imaplib
import email
from email.header import decode_header
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Configuración de credenciales
EMAIL_USER = "seikoeliz@hotmail.com"
EMAIL_PASS = "oqdzecebrhatugsk"
IMAP_SERVER = "outlook.office365.com"
IMAP_FOLDER = "BANCOS - BANCARD OTROS/ATLAS"  # Ruta completa de la carpeta

# Grupo de WhatsApp
WHATSAPP_GROUP = "VENUS TRANSFERENCIAS"

# Función para conectar a Outlook

def conectar_outlook():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select(IMAP_FOLDER)
        return mail
    except Exception as e:
        print(f"Error al conectar con Outlook: {e}")
        return None

# Función para buscar correos de transferencias

def obtener_transferencias(mail):
    try:
        status, messages = mail.search(None, 'UNSEEN')
        email_ids = messages[0].split()
        transferencias = []
        
        for e_id in email_ids:
            status, data = mail.fetch(e_id, "(RFC822)")
            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")
                    
                    if "AVISO DE TRANSFERENCIAS INTERBANCARIAS" in subject or "NOTIFICACION DE TRANSFERENCIAS INTERBANCARIAS" in subject:
                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    body = part.get_payload(decode=True).decode("utf-8")
                                    break
                        else:
                            body = msg.get_payload(decode=True).decode("utf-8")
                        
                        if "Cuenta Corriente: 1272612" in body:  # Verifica la cuenta de la peluquería
                            transferencias.append(parsear_mensaje(body))
        return transferencias
    except Exception as e:
        print(f"Error al obtener transferencias: {e}")
        return []

# Función para extraer datos del mensaje

def parsear_mensaje(body):
    try:
        lineas = body.split("\n")
        datos = {}
        for linea in lineas:
            if "Enviado por:" in linea:
                datos["Enviado por"] = linea.split(":")[-1].strip()
            elif "Monto Crédito:" in linea:
                datos["Monto Crédito"] = linea.split(":")[-1].strip()
            elif "Banco Origen:" in linea:
                datos["Banco Origen"] = linea.split(":")[-1].strip()
            elif "Concepto:" in linea:
                datos["Comprobante"] = linea.split(":")[-1].strip()
        return datos
    except Exception as e:
        print(f"Error al parsear mensaje: {e}")
        return {}

# Función para enviar mensaje por WhatsApp

def enviar_whatsapp(mensaje):
    try:
        driver = webdriver.Chrome()
        driver.get("https://web.whatsapp.com/")
        input("Escanea el código QR y presiona Enter para continuar...")
        time.sleep(10)
        
        chat = driver.find_element(By.XPATH, f"//span[@title='{WHATSAPP_GROUP}']")
        chat.click()
        time.sleep(3)
        
        input_box = driver.find_element(By.XPATH, "//div[@title='Escribe un mensaje aquí']")
        input_box.send_keys(mensaje)
        input_box.send_keys(Keys.ENTER)
        time.sleep(2)
        driver.quit()
    except Exception as e:
        print(f"Error al enviar mensaje de WhatsApp: {e}")

# Ejecutar el proceso
if __name__ == "__main__":
    mail = conectar_outlook()
    if mail:
        transferencias = obtener_transferencias(mail)
        if transferencias:
            for trans in transferencias:
                mensaje = (f"RECIBIDO TRANSFERENCIA\n"
                           f"Enviado por: {trans.get('Enviado por', 'Desconocido')}\n"
                           f"Monto Crédito: {trans.get('Monto Crédito', 'No especificado')}\n"
                           f"Banco Origen: {trans.get('Banco Origen', 'No especificado')}\n"
                           f"Comprobante: {trans.get('Comprobante', 'No especificado')}\n\n"
                           "Reaccionar con 👍 este mensaje, la sucursal que corresponde esta transferencia.")
                enviar_whatsapp(mensaje)
        else:
            print("No se encontraron transferencias para procesar.")
        mail.logout()

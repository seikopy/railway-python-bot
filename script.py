import os
import imaplib
import email
from email.header import decode_header
import time
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993

def extraer_mensaje(mensaje):
    """ Extrae los datos clave de un correo de transferencia """
    sujeto = mensaje["Subject"]
    de = mensaje["From"]
    cuerpo = ""
    
    if mensaje.is_multipart():
        for parte in mensaje.walk():
            tipo = parte.get_content_type()
            if "text/plain" in tipo:
                cuerpo = parte.get_payload(decode=True).decode()
                break
    else:
        cuerpo = mensaje.get_payload(decode=True).decode()
    
    return sujeto, de, cuerpo

def buscar_transferencias():
    """ Conecta con Outlook mediante IMAP y busca correos de transferencias """
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
        mail.select("INBOX/BANCOS - BANCARD OTROS/ATLAS")  # Ruta de la carpeta en Outlook
        
        # Buscar correos con palabras clave
        status, mensajes = mail.search(None, 'SUBJECT "TRANSFERENCIAS INTERBANCARIAS"')
        correos = mensajes[0].split()
        
        if not correos:
            print("No se encontraron transferencias para procesar.")
            return None
        
        for num in correos[-1:]:  # Solo procesar el más reciente
            status, datos = mail.fetch(num, "(RFC822)")
            raw_email = datos[0][1]
            mensaje = email.message_from_bytes(raw_email)
            
            # Extraer información
            sujeto, remitente, cuerpo = extraer_mensaje(mensaje)
            print("Correo encontrado:", sujeto, remitente)
            return cuerpo
        
    except Exception as e:
        print("Error al conectar con Outlook:", str(e))
        return None
    finally:
        mail.logout()

if __name__ == "__main__":
    while True:
        print("Buscando transferencias...")
        transferencia = buscar_transferencias()
        if transferencia:
            print("Transferencia detectada:", transferencia[:500])  # Mostrar solo parte del cuerpo
        time.sleep(60)  # Esperar 1 minuto antes de buscar nuevamente

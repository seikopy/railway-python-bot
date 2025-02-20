import os
import imaplib
import email
from email.header import decode_header
from dotenv import load_dotenv
import time

load_dotenv()

# Credenciales de Outlook
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")
IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993

# Configuración del grupo de WhatsApp
WHATSAPP_GROUP = "VENUS TRANSFERENCIAS"

# Conectar a Outlook
try:
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
    mail.select("INBOX/BANCOS - BANCARD OTROS/ATLAS")
    print("Conectado a Outlook con éxito.")
except Exception as e:
    print(f"Error al conectar con Outlook: {e}")
    mail = None

# Buscar correos de transferencias
if mail:
    try:
        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        print(f"Correos encontrados: {len(email_ids)}")

        for email_id in email_ids[-5:]:  # Procesar los últimos 5 correos
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8")
                    print(f"Asunto del correo: {subject}")

                    # Filtrar solo transferencias
                    if "TRANSFERENCIAS INTERBANCARIAS" in subject.upper():
                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                content_type = part.get_content_type()
                                if "plain" in content_type:
                                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                    break
                        else:
                            body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

                        # Extraer información relevante
                        cuenta_corriente = "1272612"
                        if cuenta_corriente in body:
                            print("Correo corresponde a la cuenta de la peluquería.")
                            datos = {
                                "Enviado por": "",
                                "Monto Crédito": "",
                                "Banco Origen": "",
                                "Comprobante": ""
                            }
                            for line in body.split("\n"):
                                if "Enviado por:" in line:
                                    datos["Enviado por"] = line.split(":", 1)[1].strip()
                                elif "Monto Crédito:" in line:
                                    datos["Monto Crédito"] = line.split(":", 1)[1].strip()
                                elif "Banco Origen:" in line:
                                    datos["Banco Origen"] = line.split(":", 1)[1].strip()
                                elif "Concepto:" in line:
                                    datos["Comprobante"] = line.split("Ref:")[1].strip()
                            
                            mensaje_whatsapp = (f"*RECIBIDO TRANSFERENCIA*\n"
                                                f"Enviado por: {datos['Enviado por']}\n"
                                                f"Monto Crédito: {datos['Monto Crédito']}\n"
                                                f"Banco Origen: {datos['Banco Origen']}\n"
                                                f"Comprobante: {datos['Comprobante']}\n\n"
                                                f"Reaccionar con 👍 este mensaje, la sucursal que corresponde esta transferencia.")
                            
                            print("Mensaje a enviar:", mensaje_whatsapp)
                            # Aquí se enviaría el mensaje a WhatsApp
                        else:
                            print("Correo no corresponde a la cuenta de la peluquería. Ignorando.")
    except Exception as e:
        print(f"Error al procesar correos: {e}")
    finally:
        mail.logout()
else:
    print("No se pudo conectar a Outlook.")

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def enviar_correo_con_pdf(destinatario, asunto, cuerpo_texto, archivo_pdf, remitente=None, password=None):
    """
    Envía un correo con un archivo PDF adjunto.
    """
    if remitente is None:
        remitente = "TUMail@gmail.com"
    if password is None:
        password = "TU_CONTRASEÑA_DE_APLICACION"

    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto
    mensaje.attach(MIMEText(cuerpo_texto, "plain", "utf-8"))

    # Adjuntar PDF
    if os.path.exists(archivo_pdf):
        with open(archivo_pdf, "rb") as adjunto:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(adjunto.read())
            encoders.encode_base64(parte)
            parte.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(archivo_pdf)}"
            )
            mensaje.attach(parte)
    else:
        print(f"⚠️ Archivo PDF no encontrado: {archivo_pdf}")
        return False

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(remitente, password)
            servidor.send_message(mensaje)
        print("✅ Correo con PDF enviado correctamente")
        return True
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")
        raise e


def enviar_correo(destinatario, asunto, cuerpo, remitente=None, password=None):
    """
    Envía un correo simple (sin adjunto) - compatible con versión anterior.
    """
    if remitente is None:
        remitente = "TUMail@gmail.com"
    if password is None:
        password = "TU_CONTRASEÑA_DE_APLICACION"

    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto
    mensaje.attach(MIMEText(cuerpo, "plain", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(remitente, password)
            servidor.send_message(mensaje)
        print("✅ Correo enviado correctamente")
        return True
    except Exception as e:
        print(f"❌ Error al enviar correo: {e}")
        raise e
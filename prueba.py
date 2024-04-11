import pandas as pd
import smtplib
from email.message import EmailMessage

remitente = "sistema@panpacksa.com.ar"
contrasena = "Sistucu_2015"
servidor_smtp = "smtp.gmail.com"
puerto = 587  # 587 o 465

notificaciones = []


def send_mail(destinatario: str, cabecera: str, contenido: str):
    mensaje = EmailMessage()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = cabecera
    mensaje.set_content(contenido)
    with smtplib.SMTP(servidor_smtp, puerto) as smtp:
        smtp.starttls()
        smtp.login(remitente, contrasena)
        smtp.send_message(mensaje)


"""
Alerta de diferencias pesos
"""

ruta_archivo = "D:\Privada\Master Cuadro Telas.xls"
df = pd.read_excel(ruta_archivo, sheet_name="Articulo",
                   header=0)
print(df.head())
valor = df.iloc[0, 5]
print(f"** {valor} **")

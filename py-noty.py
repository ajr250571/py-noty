import pandas as pd
import smtplib
from email.message import EmailMessage

remitente = "sistema@panpacksa.com.ar"
contrasena = "Sistucu_2015"
servidor_smtp = "smtp.gmail.com"
puerto = 587  # 587 o 465

notificaciones = []


def send_mail(destinatario: str, cabecera: str, contenido: str):
    contenido += chr(13)+chr(13)
    contenido += "Sistema de Notificaciones py-Notify"+chr(13)
    contenido += "Autor: Armando Rodriguez"+chr(13)

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
Envia alerta para actualizar ProgIni
"""
try:
    ruta_archivo = "//cargas/Produccion del dia/Planilla Ingreso Produccion a Deposito.xls"
    df = pd.read_excel(ruta_archivo, sheet_name="Info Ger")
    """ print(df.head()) """
    valor_planilla = df.iloc[0, 46]
    valor_progini = df.iloc[0, 47]
    if (valor_planilla != valor_progini):
        """  Enviar Mail """
        destinatario = "arodriguez@panpacksa.com.ar"
        cabecera = "CUIDADO: Falta actualizar ProgIni en Planilla Ingreso Produccion Deposito."
        contenido = cabecera
        send_mail(destinatario, cabecera, contenido)
        """"""
        notificaciones.append(f"ProgIni -> {cabecera}")
    else:
        notificaciones.append("ProgIni -> OK")
except Exception as e:
    notificaciones.append("ProgIni -> ERROR except")


"""
Alerta de diferencias pesos Arpack vs Master Cuadro de Tela
"""
try:
    ruta_archivo = "D:\Privada\Master Cuadro Telas.xls"
    df = pd.read_excel(ruta_archivo, sheet_name="Articulo")
    """ print(df.head()) """
    valor = df.iloc[0, 5]

    if (valor != 0):
        """  Enviar Mail """
        destinatario = "arodriguez@panpacksa.com.ar"
        cabecera = f"CUIDADO: Diferencias de peso en telas en {valor} articulos. Master cuadro de Tela."
        contenido = cabecera
        send_mail(destinatario, cabecera, contenido)
        """"""
        notificaciones.append(f"Peso Tela -> {cabecera}")
    else:
        notificaciones.append("Peso Tela -> OK")
except Exception as e:
    notificaciones.append("ProgIni -> ERROR except")


"""
Alerta de EE TanPi
"""
try:
    ruta_archivo = "//extrusor/Hilatura/EE.xlsm"
    df = pd.read_excel(ruta_archivo, sheet_name="Info")
    valor = round(df.iloc[0, 4], 2)
    if (valor > 0.32):
        """  Enviar Mail """
        destinatario = "arodriguez@panpacksa.com.ar;santacruz@panpacksa.com.ar"
        cabecera = f"CUIDADO: EE TanPi = {valor}  // Menor a 0.32 -> bonificacion.  // Mayor a 0.42 -> Recargo."
        contenido = cabecera
        send_mail(destinatario, cabecera, contenido)
        """"""
        notificaciones.append(f"TanPi -> {cabecera}")
    else:
        notificaciones.append("TanPi -> OK")
except Exception as e:
    notificaciones.append("ProgIni -> ERROR except")

"""
Alerta Articulos fuera de programa
"""
try:
    ruta_archivo = "//cargas/Produccion del dia/Planilla Ingreso Produccion a Deposito.xls"
    df = pd.read_excel(ruta_archivo, sheet_name="SinProg")
    valor_tela = df.iloc[0, 0]
    df = pd.read_excel(ruta_archivo, sheet_name="Cordel_FP")
    valor_cordel = df.iloc[0, 5]

    if (valor_tela + valor_cordel > 0):
        """  Enviar Mail """
        destinatario = "arodriguez@panpacksa.com.ar"
        cabecera = f"CUIDADO: Articulos fuera de programa. (tela: {valor_tela}, cordel: {valor_cordel})"
        contenido = cabecera
        send_mail(destinatario, cabecera, contenido)
        """"""
        notificaciones.append(f"FueraProg -> {cabecera}")
    else:
        notificaciones.append("FueraProg -> OK")
except Exception as e:
    notificaciones.append("FueraProg -> ERROR except")

"""
Resumen por Mail
"""
try:
    destinatario = "arodriguez@panpacksa.com.ar"
    cabecera = "Resumen Diario py-Noty"
    contenido = "Notificaciones Procesadas:"+chr(13)+chr(13)
    for notificacion in notificaciones:
        contenido += notificacion + chr(13)

    send_mail(destinatario, cabecera, contenido)
except Exception as e:
    raise e

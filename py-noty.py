import datetime
import pandas as pd
import smtplib
from email.message import EmailMessage

remitente = "sistema@panpacksa.com.ar"
contrasena = "Sistucu_2015"
servidor_smtp = "smtp.gmail.com"
puerto = 587  # 587 o 465

notificaciones = []


def send_mail(destinatario: str, cabecera: str, contenido: str):
    """
    send_mail
    Args:
        destinatario (str): email destinatarios separado con ;
        cabecera (str): texto de cabecera del mail
        contenido (str): texto del contenido del mail, chr(13) salto de linea
    """
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
        notificaciones.append(f"PesoTela -> {cabecera}")
    else:
        notificaciones.append("PesoTela -> OK")
except Exception as e:
    notificaciones.append("PesoTela -> ERROR except")


"""
Alerta de EE TanPi
"""
try:
    ruta_archivo = "//extrusor/Hilatura/EE.xlsm"
    df = pd.read_excel(ruta_archivo, sheet_name="Info")
    valor = round(df.iloc[0, 4], 2)
    if (valor > 0.32):
        """  Enviar Mail """
        destinatario = "santacruz@panpacksa.com.ar,arodriguez@panpacksa.com.ar"
        cabecera = f"CUIDADO: EE TanPi = {valor}  // Menor a 0.32 -> bonificacion.  // Mayor a 0.42 -> Recargo."
        contenido = cabecera
        send_mail(destinatario, cabecera, contenido)
        """"""
        notificaciones.append(f"TanPi -> {cabecera}")
    else:
        notificaciones.append("TanPi -> OK")
except Exception as e:
    notificaciones.append("TanPi -> ERROR except")

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
ruta: \\cintas\Stockint
Archivo: Master Titulos.xlsm
Hoja: Control
Verifica minimo 3 titulaciones dia por extr 8.
"""
hoy = datetime.date.today()
ayer = hoy - datetime.timedelta(days=1)
dia_ayer = ayer.weekday()

""" Si el dia no es sabado o domingo """
if dia_ayer != 5 and dia_ayer != 6:
    destinatario = "suphila@panpacksa.com.ar,calidad@panpacksa.com.ar,supteje@panpacksa.com.ar,rlascano@panpacksa.com.ar,arodriguez@panpacksa.com.ar"
    ruta_archivo = "//cintas/Stockint/Master Titulos.xlsm"
    tit_min = 3
    df = pd.read_excel(ruta_archivo, sheet_name="Control")
    try:
        extr2 = round(df.iloc[5, 2], 0)
        if (extr2 > 0 and extr2 < tit_min):
            cabecera = f"CUIDADO: Extrusora 2 // Minimo: {tit_min} tit/dia // Ayer: {extr2} tit/dia)"
            contenido = cabecera
            send_mail(destinatario, cabecera, contenido)
            """"""
            notificaciones.append(f"Titulos Extr2 -> {cabecera}")
        else:
            notificaciones.append("Titulos Extr2 -> OK")

    except Exception as e:
        notificaciones.append("Titulos Extr2 -> ERROR except")

    try:
        extr6 = round(df.iloc[5, 6], 0)
        if (extr6 > 0 and extr6 < tit_min):
            cabecera = f"CUIDADO: Extrusora 6 // Minimo: {tit_min} tit/dia // Ayer: {extr6} tit/dia)"
            contenido = cabecera
            send_mail(destinatario, cabecera, contenido)
            """"""
            notificaciones.append(f"Titulos Extr6 -> {cabecera}")
        else:
            notificaciones.append("Titulos Extr6 -> OK")

    except Exception as e:
        notificaciones.append("Titulos Extr6 -> ERROR except")

    try:
        extr8 = round(df.iloc[5, 7], 0)
        if (extr8 > 0 and extr8 < tit_min):
            cabecera = f"CUIDADO: Extrusora 8 // Minimo: {tit_min} tit/dia // Ayer: {extr8} tit/dia)"
            contenido = cabecera
            send_mail(destinatario, cabecera, contenido)
            """"""
            notificaciones.append(f"Titulos Extr8 -> {cabecera}")
        else:
            notificaciones.append("Titulos Extr8 -> OK")

    except Exception as e:
        notificaciones.append("Titulos Extr8 -> ERROR except")


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

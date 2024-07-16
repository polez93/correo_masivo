import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Leer el archivo Excel sin encabezados
df = pd.read_excel('prueba.xlsx', header=None)

# Configuración del servidor de correo
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'clquinteto@gmail.com'
smtp_password = 'izrq yzmd qaij stzw'

# Crear la conexión al servidor
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(smtp_user, smtp_password)

# Mensaje base
mensaje_base = """
Hola {nombre},

Este es un mensaje personalizado para ti. Gracias por ser nuestro cliente.

Saludos,
Tu Empresa
"""

# Iterar sobre cada fila en el DataFrame y enviar un correo personalizado
for index, row in df.iterrows():
    nombre = row[0]  # Asumiendo que los nombres están en la primera columna (A)
    correo = row[22]  # Asumiendo que los correos están en la segunda columna (B)
    mensaje_personalizado = mensaje_base.format(nombre=nombre)
    
    # Crear el mensaje
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = correo
    msg['Subject'] = 'Mensaje Personalizado'
    
    msg.attach(MIMEText(mensaje_personalizado, 'plain'))
    
    # Enviar el mensaje
    server.send_message(msg)

# Cerrar la conexión al servidor
server.quit()
!pip install gspread oauth2client pandas

from google.colab import drive
drive.mount('/content/drive')

import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from IPython.display import display

# Configuración de email
EMAIL_ACCOUNT = 'carlitosaburgosj1609s@hotmail.com'
EMAIL_PASSWORD = 'cmwukclgvihgqwlk'
IMAP_SERVER = 'imap-mail.outlook.com'
SMTP_SERVER = 'smtp-mail.outlook.com'

# Configuración de Google Sheets
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1JGjg2NjMu-BkTWbIJIs1ZjiNmu_V-EDvEKVwauiDjyc/edit?gid=455436416#gid=455436416'
CREDENTIALS_PATH = '/content/drive/MyDrive/Juzgado gsheets/credentials.json'

# Diccionario para nombres de meses en español (mayúsculas)
MONTHS_ES = {
    1: 'ENERO', 2:

 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO',
    7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
}

# Inicializar cliente de Google Sheets
def get_gsheets_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, scope)
    client = gspread.authorize(creds)
    return client

def get_open_court():
    today = datetime.now()
    month_num = today.month
    month_es = MONTHS_ES.get(month_num, None)
    if not month_es:
        raise ValueError("Mes no válido")

    today_str = today.strftime('%d/%m/%Y')
    print(f"Buscando juzgado abierto para fecha: {today_str}")

    # Leer la primera hoja de Google Sheets
    client = get_gsheets_client()
    spreadsheet = client.open_by_url(SHEET_URL)
    worksheet = spreadsheet.get_worksheet(0)  # Usar la primera hoja
    df = pd.DataFrame(worksheet.get_all_records())

    # Mostrar la tabla como en Colab
    display(df)

    # Convertir fechas a DD/MM/YYYY para comparación, con manejo de errores
    def parse_date(date_str):
        try:
            if not date_str or isinstance(date_str, float) or date_str.strip() == '':
                return None
            return pd.to_datetime(date_str, dayfirst=True).strftime('%d/%m/%Y')
        except Exception as e:
            print(f"Error al parsear fecha '{date_str}': {e}")
            return None

    df['FECHA'] = df['FECHA'].apply(parse_date)
    df = df[df['FECHA'].notnull()]

    # Filtrar por fecha actual y estado "Abierto"
    row = df[(df['FECHA'] == today_str) & (df['ESTADO'].str.lower() == 'abierto')]

    if not row.empty:
        court_name = row['JUZGADO'].iloc[0]
        court_email = row['CORREO'].iloc[0]
        municipio = row['Municipio'].iloc[0]
        print(f"Juzgado abierto encontrado: {court_name} ({court_email}, {municipio})")
        return court_name, court_email, municipio
    print("No se encontró juzgado abierto para hoy.")
    return None, None, None

def read_habeas_corpus_emails():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        print("Conectado al servidor IMAP... Intentando login.")
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        print("Login exitoso en IMAP.")
        mail.select('inbox')

        # Buscar correos no leídos con "habeas corpus" en el asunto
        _, message_numbers = mail.search(None, '(UNSEEN SUBJECT "habeas corpus")')
        messages = []

        for num in message_numbers[0].split():
            _, msg_data = mail.fetch(num, '(RFC822)')
            email_body = msg_data[0][1]
            msg = email.message_from_bytes(email_body)
            subject = msg['subject']
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        break
            else:
                content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
            messages.append((num, subject, content))
            print(f"Correo encontrado: Asunto='{subject}', Contenido inicial={content[:50]}...")

        return mail, messages
    except Exception as e:
        print(f"Error en conexión IMAP: {e}")
        raise

def send_email_to_court(to_email, subject, content):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ACCOUNT
        msg['To'] = to_email
        msg['Subject'] = f'Reenviado: {subject}'
        msg.attach(MIMEText(content, 'plain'))

        server = smtplib.SMTP(SMTP_SERVER, 587)
        server.starttls()
        print("Conectado al servidor SMTP... Intentando login.")
        server.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        print("Login exitoso en SMTP.")
        server.sendmail(EMAIL_ACCOUNT, to_email, msg.as_string())
        server.quit()
        print(f"Correo enviado a: {to_email}")
    except Exception as e:
        print(f"Error en envío SMTP: {e}")
        raise

def mark_email_processed(mail, msg_id):
    mail.store(msg_id, '+FLAGS', r'\Seen')
    print(f"Correo con ID {msg_id.decode('utf-8')} marcado como leído.")

def log_to_gsheets(subject, court_email, municipio):
    client = get_gsheets_client()
    spreadsheet = client.open_by_url(SHEET_URL)
    ws = spreadsheet.worksheet('correos')

    # Obtener datos existentes para determinar la última fila e ID
    data = ws.get_all_records()
    last_row = len(data) + 1  # +1 por la fila de encabezados
    new_row = last_row + 1

    # ID: Incrementar del anterior
    if last_row > 1:
        new_id = int(data[-1].get('id', 0)) + 1
    else:
        new_id = 1

    # Fecha y hora actual
    fecha = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    usuario = "Sistema Automatico"
    correo_usuario = EMAIL_ACCOUNT
    archivo = ""

    # Agregar nueva fila (columnas A a H)
    ws.append_row([
        new_id,  # id
        fecha,   # fecha y hora
        f'=TEXTO(B{new_row},"DDDD")',  # Fórmula para día (en español para Sheets)
        usuario,  # usuario
        correo_usuario,  # correo usuario
        court_email,  # correo juzgado
        archivo,  # archivo
        municipio  # municipio
    ])

    print(f"Registro agregado en fila {new_row}: ID={new_id}, Fecha={fecha}, Juzgado={court_email}")

def main():
    while True:
        try:
            # Obtener juzgado abierto
            court_name, court_email, municipio = get_open_court()
            if not court_email:
                print("No hay juzgados abiertos hoy.")
                time.sleep(300)
                continue

            # Leer correos
            mail, messages = read_habeas_corpus_emails()
            for msg_id, subject, content in messages:
                if "habeas corpus" in subject.lower():
                    print(f"Procesando correo: {subject}")
                    send_email_to_court(court_email, subject, content)
                    mark_email_processed(mail, msg_id)
                    log_to_gsheets(subject, court_email, municipio)
                    print(f"Correo enviado a {court_name} ({court_email}) y registrado en Google Sheets.")

            mail.logout()
            print("Ciclo completado. Esperando 5 minutos...")
            time.sleep(300)

        except Exception as error:
            print(f"Error: {error}")
            time.sleep(60)  # Reintentar en 1 minuto

if __name__ == '__main__':
    main()
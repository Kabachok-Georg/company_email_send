import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dotenv import load_dotenv
import os


# Загрузить переменные окружения из файла .env
load_dotenv()


# Извлечь данные из переменных окружения
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT'))
SMTP_USERNAME = os.getenv('SMTP_USERNAME')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')


# Проверить на наличие обязательных переменных
if not all([SMTP_SERVER, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD]):
    raise ValueError("Missing required SMTP environment variables.")


def send_email(to_email, subject, html_message, from_email):
    msg = MIMEMultipart('related')
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Создать HTML-часть сообщения
    html_part = MIMEText(html_message, 'html')
    msg.attach(html_part)

    # Прикрепить изображение
    try:
        with open('img/logo.jpg', 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-ID', '<logo>')
            img.add_header('Content-Disposition', 'inline', filename='logo.jpg')
            msg.attach(img)
    except FileNotFoundError:
        print("Изображение 'logo.jpg' не найдено!")

    try:
        # Подключение к серверу SMTP
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            # Для порта 587 используем starttls()
            server.starttls()  # Защищенное соединение
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
            print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


# Загрузить данные из JSON-файла
try:
    with open('data/emails.json', 'r', encoding='utf-8') as file:
        email_data = json.load(file)
except FileNotFoundError:
    print("JSON файл с данными не найден!")
    exit()


# Загрузить HTML-шаблон из файла
try:
    with open('templates/email_template.html', 'r', encoding='utf-8') as file:
        html_template = file.read()
except FileNotFoundError:
    print("HTML шаблон не найден!")
    exit()


# Отправить письма
for entry in email_data:
    html_message = html_template.format(
        name=entry['name'],
        date=entry['date'],
        time=entry['time'],
        event_link=entry['event_link']
    )
    send_email(
        to_email=entry['email'],
        subject=entry['subject'],
        html_message=html_message,
        from_email=SMTP_USERNAME
    )

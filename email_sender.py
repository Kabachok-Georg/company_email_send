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

    # Прикрепить изображение новогодней открытки
    try:
        with open(r'C:\Users\User\Documents\GitHub\company_email_send\img\New_Year_2025.jpg', 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-ID', '<header_image>')  # Используется в HTML-шаблоне как cid:header_image
            img.add_header('Content-Disposition', 'inline', filename='New_Year_2025.jpg')
            msg.attach(img)
    except FileNotFoundError:
        print("Изображение 'New_Year_2025.jpg' не найдено!")

    # Подключение к серверу SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Защищенное соединение
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
            print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


# Загрузить данные из JSON-файла
try:
    with open('data/emails_2025.json', 'r', encoding='utf-8') as file:
        email_data = json.load(file)
except FileNotFoundError:
    print("JSON файл с данными не найден!")
    exit()


# Отправить письма
for entry in email_data:
    # Пропускаем записи с пустыми полями email или name
    if not entry['email'] or not entry['name'] or not entry['subject']:
        print(f"Пропущена запись с недостающими данными: {entry}")
        continue

    # Формируем HTML-сообщение из данных JSON
    html_message = f"""
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Поздравление с Новым 2025 годом</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background-color: #f7f7f7;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                background-color: #ffffff;
                border-radius: 10px;
                padding: 20px;
                max-width: 600px;
                margin: auto;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            }}
            h1 {{
                color: #FF6347;
                text-align: center;
                font-size: 30px;
                margin-bottom: 20px;
            }}
            h2 {{
                color: #333333;
                font-size: 22px;
                margin-bottom: 20px;
            }}
            p {{
                color: #555555;
                font-size: 16px;
                line-height: 1.6;
            }}
            .highlight {{
                color: #FF4500;
                font-weight: bold;
            }}
            .footer {{
                font-size: 14px;
                color: #888888;
                text-align: center;
                margin-top: 40px;
                border-top: 1px solid #dddddd;
                padding-top: 10px;
            }}
            .snowflakes {{
                font-size: 2em;
                color: #FF6347;
                text-align: center;
                margin-top: 20px;
            }}
            .signature {{
                font-weight: bold;
                margin-top: 30px;
                text-align: center;
            }}
            .header-image {{
                display: block;
                margin: 0 auto;
                max-width: 100%;
                border-radius: 8px;
            }}
            .button {{
                display: inline-block;
                background-color: #FF6347;
                color: white;
                padding: 12px 24px;
                font-size: 18px;
                border-radius: 5px;
                text-decoration: none;
                margin-top: 30px;
                text-align: center;
            }}
            .button:hover {{
                background-color: #FF4500;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>С Наступающим Новым 2025 Годом {entry['name']}</h1>
            <img src="cid:header_image" alt="Новогодняя открытка" class="header-image">
            <p>Дорогой {entry['name']}.</p>
            <p>Рад поздравить тебя с наступающим <span class="highlight">Новым 2025 годом</span>. Этот момент — время для новых начинаний.</p>
            <h2>Пожелания на 2025 год:</h2>
            <ul>
                <li>Новых горизонтов для личного и профессионального роста.</li>
                <li>Множества счастливых и ярких моментов с близкими.</li>
                <li>Заботы о себе и своем здоровье.</li>
                <li>И конечно, успехов во всех начинаниях.</li>
            </ul>
            <p class="signature">С наилучшими пожеланиями.<br></p>
            <a href="mailto:georgdan28@mail.ru" class="button">Отправить обратную открытку</a>
            <div class="footer">
                <p>© 2025 Георгий Данилов. Все права защищены.</p>
            </div>
            <div class="snowflakes">
                &#10052;&#10052;&#10052;
            </div>
        </div>
    </body>
    </html>
    """

    send_email(
        to_email=entry['email'],
        subject=entry['subject'],
        html_message=html_message,
        from_email=SMTP_USERNAME
    )

import imaplib
import email
from email.header import decode_header
from datetime import datetime, timezone
import pandas as pd
import re

# Функция для очистки текста от байтовых строк
def clean_text(text):
    if isinstance(text, bytes):
        text = text.decode('utf-8', errors='ignore')  # Игнорируем ошибки при декодировании
    # Удаляем байтовые строки
    return re.sub(r"b'(.*?)'", '', text)
# Настройки подключения
username = "shechenkoea@afanasy.ru"
password = "AyZfmCD3NCczmMLq0sxw"  # Используйте пароль приложения

# Подключение к почтовому серверу
mail = imaplib.IMAP4_SSL("imap.mail.ru")
mail.login(username, password)
mail.select("inbox")

# Дата начала фильтрации
start_date_str = "01-Jan-2024"  # Формат: DD-Mon-YYYY

status, messages = mail.search(None, f'SINCE "{start_date_str}"')
mail_ids = messages[0].split()

# Список для хранения данных
data = []

# Обработка каждого письма
for mail_id in mail_ids:
    # Получение письма по ID
    status, msg_data = mail.fetch(mail_id, "(RFC822)")
    msg = email.message_from_bytes(msg_data[0][1])

    # Декодирование заголовков
    date = msg["Date"]
    from_ = clean_text(decode_header(msg["From"])[0][0])
    to = clean_text(decode_header(msg["To"])[0][0])

    # Проверка наличия темы
    subject = msg["Subject"]
    if subject is not None:
        subject = clean_text(decode_header(subject)[0][0])
    else:
        subject = "Без темы"

    # Преобразование даты
    if isinstance(date, str):
        date = email.utils.parsedate_to_datetime(date)
        if date.tzinfo is None:  # Если date offset-naive, добавляем timezone
            date = date.replace(tzinfo=timezone.utc)
    else:
        date = datetime.now(timezone.utc)  # Если дата не была получена

    # Получение тела письма
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = clean_text(part.get_payload(decode=True))
                break
    else:
        body = clean_text(msg.get_payload(decode=True))

    # Добавление данных в список
    data.append({
        "Дата/время": date.strftime('%Y-%m-%d %H:%M:%S'),
        "Кому": to,
        "От кого": from_,
        "Тема письма": subject,
        "Тело письма": body
    })
    break
# Создание DataFrame и сохранение в Excel
df = pd.DataFrame(data)
df.to_excel("emails.xlsx", index=False)

# Закрытие соединения
mail.logout()

import configparser

from docx import Document
import locale
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

try:
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
except locale.Error:
    print("Локаль 'ru_RU.UTF-8' не поддерживается. Используйте другую локаль.")
    input()

date = datetime.now().strftime('«%d» %B %Yг')
filename = f"Отчет_{datetime.now().strftime('%d %B %Yг')}.docx".replace(" ", "_")

config = configparser.ConfigParser()
try:
    config.read('config.ini')
    server = config.get('SMTP', 'server')
    port = config.get('SMTP', 'port')
    email_from = config.get('SMTP', 'email_from')
    email_to = config.get('SMTP', 'email_to')
    password = config.get('SMTP', 'password')
except configparser.NoSectionError as e:
    print(f"Ошибка: {e}. Убедитесь, что секция 'SMTP' присутствует в файле config.ini.")
    input()
except configparser.NoOptionError as e:
    print(f"Ошибка: {e}. Убедитесь, что все необходимые параметры присутствуют в секции 'SMTP' в файле config.ini.")
    input()

try:
    doc = Document('Отчет.docx')
    paragraphs = doc.paragraphs
    table = doc.tables[0]
    for i in range(len(table.columns)):
        table.cell(1, i).text = table.cell(2, i).text

    table.cell(2, 0).text = datetime.now().strftime('%d %B')
    table.cell(2, 1).text = input('"Холодная вода куб. м.": ') + " куб. м."
    table.cell(2, 2).text = input('"Горячая вода куб. м.": ') + " куб. м."

    cold_water_1 = float(table.cell(1, 1).text.split()[0])
    hot_water_1 = float(table.cell(1, 2).text.split()[0])
    cold_water_2 = float(table.cell(2, 1).text.split()[0])
    hot_water_2 = float(table.cell(2, 2).text.split()[0])

    table.cell(3, 1).text = f'{round(cold_water_2 - cold_water_1, 2)} куб. м.'
    table.cell(3, 2).text = f'{round(hot_water_2 - hot_water_1, 2)} куб. м.'

    paragraphs[-2].text = f'Подпись квартиросъемщика_____________________________________Дата {date}'
    paragraphs[-1].text = f'Сведения принял______________________________________________Дата {date}'
    doc.save(filename)
except Exception as e:
    print("Ошибка. Поместите файл с названием \"Отчет.docx\" рядом с исполняемым файлом.")
    input()

send_mail = input("Отчет сформирован, отправить на почту? Y / N: \t")
while send_mail not in "YyДдNnНн":
    print("Некорректный ответ. Пожалуйста, введите Y или N.")
    send_mail = input("Отчет сформирован, отправить на почту? Y / N: \t")

if send_mail in "YyДд":
    print("Формирование письма...")
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = 'Передача данных'

    # Добавление файла в качестве вложения
    try:
        print("Прикрепление файла...")
        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        # Исправлен заголовок и добавлено кодирование имени файла
        part.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', filename))
        msg.attach(part)
        attachment.close()
    except Exception as e:
        print("Ошибка при прикреплении файла:", e)
        input()
    except FileNotFoundError:
        print("Ошибка. Файл не найден.")
        input()

    # Отправка сообщения
    try:
        print("Отправка письма...")
        server = smtplib.SMTP(server, port)
        server.starttls()
        server.login(email_from, password)
        text = msg.as_string()
        server.sendmail(email_from, email_to, text)
        server.quit()
        print("Письмо отправлено.")
        input()
    except Exception as e:
        print("Ошибка при отправке письма. Проверьте корректность заполнения config.ini.", e)
        input()
elif send_mail in "NnНн":
    print("Завершение программы.")
    input()
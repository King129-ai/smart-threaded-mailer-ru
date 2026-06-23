import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import openpyxl
import threading
import re
import os
from datetime import datetime
from dotenv import load_dotenv

# Загружаем переменные из файла .env
load_dotenv()

# --- ВАЛИДАЦИЯ EMAIL ---
def is_valid_email(email):
    """Проверяет, соответствует ли строка формату email-адреса"""
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, str(email)) is not None


# --- ЧТЕНИЕ EXCEL С ПРОВЕРКОЙ ---
def load_contacts_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    contacts = {}
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = row[0]
        email = str(row[1]).strip() if row[1] else None
        
        # Защита от кривых данных: проверяем, что имя и email есть, и email корректный
        if name and email and is_valid_email(email):
            contacts[email] = name
        else:
            if email or name:
                print(f"⚠️ Строка пропущена (неверный формат или пустые поля): {name} -> {email}")
                
    return contacts


# --- КЛАСС РАССЫЛЬЩИКА ---
class SafeThreadedMailer:
    def __init__(self, sender_email, app_password):
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 465
        self.sender_email = sender_email
        self.app_password = app_password
        self.log_lock = threading.Lock()
        # Ограничиваем количество ОДНОВРЕМЕННЫХ потоков, чтобы не получить бан от Google
        self.thread_semaphore = threading.Semaphore(3) # Не более 3 потоков одновременно

    def log_result(self, email, name, status, error_msg=""):
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.log_lock:
            with open("report.txt", "a", encoding="utf-8") as f:
                if status == "SUCCESS":
                    f.write(f"[{current_time}] УСПЕШНО: Письмо отправлено {name} ({email})\n")
                else:
                    f.write(f"[{current_time}] ОШИБКА: Не удалось отправить {name} ({email}). Причина: {error_msg}\n")

    def _send_single_email(self, email, name, subject, message_template, thread_id):
        # Занимаем «слот» в семафоре. Если 3 слота заняты, остальные потоки ждут здесь
        with self.thread_semaphore:
            personalized_body = message_template.replace("{name}", name)

            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = email
            msg['Subject'] = subject
            msg.attach(MIMEText(personalized_body, 'plain'))

            try:
                # Добавляем timeout=10 секунд, чтобы поток не завис навсегда при сбое сети
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, timeout=10)
                server.login(self.sender_email, self.app_password)
                
                server.sendmail(self.sender_email, email, msg.as_string())
                print(f"✅ [Поток-{thread_id}] Отправлено для: {name}")
                self.log_result(email, name, "SUCCESS")
                
                server.quit()
            except Exception as e:
                print(f"❌ [Поток-{thread_id}] Ошибка для {name}: {e}")
                self.log_result(email, name, "FAILED", str(e))
            
            # Пауза внутри потока, чтобы запросы к серверу шли плавно
            time.sleep(2)

    def send_bulk_newsletter_parallel(self, recipients_dict, subject, message_template):
        print(f"🚀 Запуск безопасной рассылки (база: {len(recipients_dict)} адресатов)...")
        threads = []

        for index, (email, name) in enumerate(recipients_dict.items()):
            t = threading.Thread(
                target=self._send_single_email, 
                args=(email, name, subject, message_template, index + 1)
            )
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        print("\n🎉 Рассылка полностью завершена! Логи в report.txt")


# --- ЗАПУСК ---
# Подтягиваем данные из переменных окружения
SENDER = os.getenv("EMAIL_USER")
PASSWORD = os.getenv("EMAIL_PASSWORD")

if not SENDER or not PASSWORD:
    print("❌ Ошибка: Проверьте, что файл .env создан и заполнен!")
else:
    subscribers = load_contacts_from_excel("contacts.xlsx")
    
    theme = "Коммерческое предложение"
    template = "Здравствуйте, {name}!\n\nЭто финальная версия безопасного рассыльщика."

    if subscribers:
        mailer = SafeThreadedMailer(SENDER, PASSWORD)
        mailer.send_bulk_newsletter_parallel(subscribers, theme, template)
    else:
        print("Нет корректных контактов для отправки.")
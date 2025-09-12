import win32com.client
import pythoncom
import re
import pandas as pd
from openpyxl import load_workbook
import os
import time

# === Настройки ===
OUTLOOK_FOLDER = "Inbox"  # подпапка или Inbox
SEARCH_SUBJECT = "Возврат поддонов из сетей"
EXCEL_FILE = r"C:\Users\legki\Desktop\test.xlsx"
SHEET_NAME = "Данные"
CHECK_INTERVAL = 10  # проверка новых писем каждые 10 секунд

# === Функция парсинга письма ===
def parse_email_body(body):
    data = {}
    patterns = {
        "Дата": r"(\d{2}\.\d{2}\.\d{4})\s+возврат",
        "Сеть": r"\|\s*(.*?)\s*\|\s*РЦ",
        "РЦ": r"РЦ\s+([А-Яа-яA-Za-z0-9\s]+)",
        "Тягач": r"Тягач.*?:\s*(.+)",
        "Прицеп": r"Прицеп.*?:\s*(.+)",
        "Водитель": r"Ф\.И\.О\. водителя:\s*(.+)",
        "Паспорт": r"Паспорт:\s*(.+)",
        "ВУ": r"Номер ВУ:\s*(.+)",
        "Телефон": r"Телефон:\s*(.+)",
        "ИНН": r"ИНН\s*(\d+)"
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, body)
        if match:
            data[key] = match.group(1).strip()
    return data

# === Подключение к открытому Outlook ===
try:
    pythoncom.CoInitialize()
    outlook = win32com.client.GetActiveObject("Outlook.Application")
except Exception:
    print("Не удалось подключиться к открытому Outlook. Запустите Outlook и войдите в свой аккаунт.")
    exit(1)

namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

# Если нужна подпапка
try:
    folder = inbox.Folders[OUTLOOK_FOLDER] if OUTLOOK_FOLDER.lower() != "inbox" else inbox
except Exception:
    print(f"Подпапка '{OUTLOOK_FOLDER}' не найдена. Используем основной Inbox.")
    folder = inbox

# === Словарь для отслеживания уже обработанных писем ===
processed_ids = set()

print("Начинаем мониторинг папки Outlook...")

while True:
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # сортируем по дате (новые сверху)

    for msg in messages:
        if msg.EntryID in processed_ids:
            continue  # уже обработано

        if SEARCH_SUBJECT in str(msg.Subject):
            body = msg.Body
            data = parse_email_body(body)
            if data:
                print("Извлечено:", data)

                # === Запись в Excel ===
                try:
                    if not os.path.exists(EXCEL_FILE):
                        pd.DataFrame([data]).to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
                    else:
                        book = load_workbook(EXCEL_FILE)
                        if SHEET_NAME not in book.sheetnames:
                            df_new = pd.DataFrame([data])
                            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a") as writer:
                                df_new.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                        else:
                            startrow = book[SHEET_NAME].max_row
                            df_new = pd.DataFrame([data])
                            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a") as writer:
                                writer.book = book
                                writer.sheets = {ws.title: ws for ws in book.worksheets}
                                df_new.to_excel(writer, sheet_name=SHEET_NAME, index=False, header=False, startrow=startrow)
                except Exception as e:
                    print("Ошибка записи в Excel:", e)

        # добавляем в обработанные
        processed_ids.add(msg.EntryID)

    time.sleep(CHECK_INTERVAL)

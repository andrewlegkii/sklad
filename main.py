import win32com.client
import pythoncom
import re
import pandas as pd
from openpyxl import load_workbook
import os
import time

# === Настройки ===
OUTLOOK_FOLDER = "Inbox"
SEARCH_SUBJECT = "Возврат поддонов из сетей"
EXCEL_FILE = r"C:\Users\legki\Desktop\test.xlsx"
SHEET_NAME = "Данные"
CHECK_INTERVAL = 5  # проверка каждые 5 секунд

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
    print("Не удалось подключиться к открытому Outlook. Запустите Outlook и войдите в аккаунт.")
    exit(1)

namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # Inbox

try:
    folder = inbox.Folders[OUTLOOK_FOLDER] if OUTLOOK_FOLDER.lower() != "inbox" else inbox
except Exception:
    folder = inbox

processed_ids = set()
print("Начинаем мониторинг папки Outlook...")

while True:
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    for msg in messages:
        if msg.EntryID in processed_ids:
            continue
        if SEARCH_SUBJECT in str(msg.Subject):
            body = msg.Body
            data = parse_email_body(body)
            if data:
                print("Извлечено:", data)

                # === Запись в Excel по порядку ===
                df_new = pd.DataFrame([data])

                if not os.path.exists(EXCEL_FILE):
                    # создаём новый файл
                    df_new.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
                else:
                    # открываем существующий файл
                    book = load_workbook(EXCEL_FILE)
                    if SHEET_NAME in book.sheetnames:
                        startrow = book[SHEET_NAME].max_row
                    else:
                        startrow = 0
                    # записываем в конец
                    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                        df_new.to_excel(writer, sheet_name=SHEET_NAME, index=False, header=startrow==0, startrow=startrow)

        processed_ids.add(msg.EntryID)

    time.sleep(CHECK_INTERVAL)

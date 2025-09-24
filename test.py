import win32com.client
import os
from datetime import datetime, timedelta
import sys
import logging

# === 🧪 НАСТРОЙКИ ТЕСТА ===
TEST_CURRENT_TIME = "15:00"  # Меняй на "12:00", "14:00" и т.д.

# Замени на реальные email для теста!
TEST_RECIPIENTS = {
    "x5": ["skoppss@yandex.ru"],
    "тандер": ["koppss@yandex.ru"],
    "дистры": ["koppss@yandex.ru"]
}

# Автоматически используем сегодняшнюю дату
TODAY_STR = datetime.today().strftime("%Y-%m-%d")
TOMORROW_STR = (datetime.today() + timedelta(days=1)).strftime("%Y-%m-%d")

TEST_RECORDS = [
    # Тест для X5 — напоминание в 12:00 сегодня
    {
        "entry_id": "x5-test",
        "Сеть": "X5",
        "РЦ": "Тюмень",
        "Дата возврата": TODAY_STR
    },
    # Тест для Тандер — возврат завтра → напоминание сегодня в 14:00
    {
        "entry_id": "tander-test",
        "Сеть": "Тандер",
        "РЦ": "Екатеринбург",
        "Дата возврата": TOMORROW_STR
    }
]

# === Логирование ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("test_script.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

sent_reminders = set()


def send_email(subject, body, to):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = to[0] if isinstance(to, list) else to
        mail.Body = body
        mail.Send()
        logging.info(f"✅ Отправлено: {subject} → {to}")
        del mail, outlook
    except Exception as e:
        logging.error(f"❌ Ошибка отправки: {e}")


def simulate_reminders():
    global sent_reminders

    today = datetime.today().date()
    current_time = TEST_CURRENT_TIME

    for record in TEST_RECORDS:
        entry_id = record["entry_id"]
        network = record["Сеть"].lower()
        rc = record["РЦ"]

        if "лента" in network:
            continue

        try:
            return_date = datetime.strptime(record["Дата возврата"], "%Y-%m-%d").date()
        except Exception as e:
            logging.error(f"Неверная дата: {e}")
            continue

        # X5 / Дистры
        if "x5" in network or "дистр" in network:
            if today == return_date and current_time == "12:00":
                key = (entry_id, "1200")
                if key not in sent_reminders:
                    send_email(
                        f"[ТЕСТ] X5: данные на РЦ {rc}",
                        f"Дата: {return_date.strftime('%d.%m.%Y')}\nРЦ: {rc}",
                        TEST_RECIPIENTS["x5"]
                    )
                    sent_reminders.add(key)

        # Тандер
        elif "тандер" in network:
            if return_date.weekday() in (5, 6, 0):
                days_back = (return_date.weekday() - 4) % 7
                if days_back == 0:
                    days_back = 7
                reminder_date = return_date - timedelta(days=days_back)
            else:
                reminder_date = return_date - timedelta(days=1)

            if today == reminder_date and current_time == "14:00":
                key = (entry_id, "1400_tander")
                if key not in sent_reminders:
                    send_email(
                        f"[ТЕСТ] ТАНДЕР: данные на РЦ {rc}",
                        f"Дата возврата: {return_date.strftime('%d.%m.%Y')}\nРЦ: {rc}",
                        TEST_RECIPIENTS["тандер"]
                    )
                    sent_reminders.add(key)


def main():
    logging.info("🚀 Запуск теста")
    logging.info(f"🕒 Текущее тестовое время: {TEST_CURRENT_TIME}")
    logging.info(f"📅 Сегодня: {TODAY_STR}")

    simulate_reminders()

    logging.info("✅ Тест завершён")
    input("\nНажмите Enter для выхода...")


if __name__ == "__main__":
    main()

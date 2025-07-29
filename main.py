from telethon import TelegramClient, events
import os
from openpyxl import Workbook, load_workbook

# === Налаштування ===
api_id = 24063327
api_hash = "4b78510032991cc80972c8e91235df51"
chat_username = "https://t.me/hahabitchezparty"
keywords = ["пизду"]

file_name = "messages.xlsx"

# === Ініціалізація клієнта ===
client = TelegramClient("userbot", api_id, api_hash)


def init_excel():
    """Створення Excel-файлу з заголовками, якщо він не існує"""
    if not os.path.exists(file_name):
        wb = Workbook()
        ws = wb.active
        ws.title = "Messages"
        ws.append(["ID", "Відправник", "Текст"])
        wb.save(file_name)


def message_exists(msg_id):
    """Перевірка, чи є повідомлення у файлі"""
    wb = load_workbook(file_name)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == msg_id:
            return True
    return False


def save_message_to_excel(msg_id, sender_id, text):
    """Збереження повідомлення в Excel"""
    if not message_exists(msg_id):
        wb = load_workbook(file_name)
        ws = wb.active
        ws.append([msg_id, sender_id, text])
        wb.save(file_name)


async def export_old_messages(entity):
    """Експорт усіх старих повідомлень"""
    print("Пошук старих повідомлень...")
    count = 0
    async for message in client.iter_messages(entity, limit=None):
        if message.message and any(k.lower() in message.message.lower() for k in keywords):
            save_message_to_excel(message.id, message.sender_id, message.message)
            count += 1
    print(f"✅ Знайдено {count} старих повідомлень.")


@client.on(events.NewMessage)
async def new_message_handler(event):
    """Обробка нових повідомлень"""
    if event.chat and event.chat.username == chat_username:
        text = event.message.message
        if text and any(k.lower() in text.lower() for k in keywords):
            save_message_to_excel(event.message.id, event.sender_id, text)
            print(f"[НОВЕ] {event.sender_id}: {text}")


async def main():
    init_excel()
    await client.start()
    print("Юзербот запущений...")

    # Отримуємо сутність групи
    entity = await client.get_entity(chat_username)
    print(f"✅ Підключено до чату: {entity.title} (ID: {entity.id})")

    await export_old_messages(entity)
    print("Очікування нових повідомлень...")
    await client.run_until_disconnected()


client.loop.run_until_complete(main())


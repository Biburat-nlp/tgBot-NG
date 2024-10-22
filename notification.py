import os
from telegram import Bot

TELEGRAM_API_TOKEN = 'token'
USER_OIV_MAP = {
    '1008683615': {
        'oiv_ids': [12, 13, 14, 15, 10, 11, 15, 16, 17, 18],
        'category_ids': [101, 102, 103, 104]
    },
    '5082179002': {'oiv_ids':[6, 7], 'category_ids': [101, 102, 103, 104]}, #Кто-то Нижегородский
}


MESSAGE_FILE_PATH = 'message.txt'


def get_general_message(file_path):
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read().strip()
    return None


async def send_general_message(bot, message):
    for user_id in USER_OIV_MAP.keys():
        try:
            await bot.send_message(chat_id=user_id, text=message)
        except Exception as e:
            print(f"Failed to send message to {user_id}: {e}")


async def notify_all_users():
    bot = Bot(token=TELEGRAM_API_TOKEN)
    message = get_general_message(MESSAGE_FILE_PATH)
    if message:
        await send_general_message(bot, message)
    else:
        print("No general message found or the file is empty.")

if __name__ == "__main__":
    import asyncio
    asyncio.run(notify_all_users())
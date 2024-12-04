import os
import pandas as pd
from telethon import TelegramClient
from datetime import datetime, timedelta
import asyncio

api_id = 'id'
api_hash = 'hash'
phone_number = 'number'
client = TelegramClient('session_name', api_id, api_hash)

OIV_IDS = {
    'ГБУ Жилищник района Марьино города Москвы': 1,
    'Управа Марьино': 2,
    'ГБУ «Автомобильные дороги ЮВАО»': 3,
    'ГБУ Жилищник Выхино района Выхино-Жулебино города Москвы': 4,
    'Управа Выхино-Жулебино': 5,
    'ГБУ Жилищник Нижегородского района города Москвы': 6,
    'Управа Нижегородский': 7,
    'ГБУ Жилищник района Капотня города Москвы': 8,
    'Управа Капотня': 9,
    'ГБУ Жилищник района Кузьминки города Москвы': 10,
    'Управа Кузьминки': 11,
    'ГБУ Жилищник района Лефортово города Москвы': 12,
    'Управа Лефортово': 13,
    'ГБУ Жилищник района Люблино города Москвы': 14,
    'Управа Люблино': 15,
    'ГБУ Жилищник района Некрасовка города Москвы': 16,
    'Управа Некрасовка': 17,
    'ГБУ Жилищник района Печатники города Москвы': 18,
    'Управа Печатники': 19,
    'ГБУ Жилищник района Текстильщики города Москвы': 20,
    'Управа Текстильщики': 21,
    'ГБУ Жилищник Рязанского района города Москвы': 22,
    'Управа Рязанский': 23,
    'ГБУ Жилищник Южнопортового района города Москвы': 24,
    'Управа Южнопортовый': 25,
    'Жилищная инспекция по ЮВАО': 26,
    'Префектура Юго-Восточного округа': 27
}

USERS_BY_OIV_ID = {
    12: [{'username': 'number/id', 'name': 'Иван Сидорович'}, {'username': 'number/id', 'name': 'Иван Сидорович'}],
    13: [{'username': 'number/id', 'name': 'Иван Сидорович'}, {'username': 'number/id', 'name': 'Иван Сидорович'}],
}

def get_latest_downloaded_file(download_path):
    files = os.listdir(download_path)
    files = [os.path.join(download_path, f) for f in files if os.path.isfile(os.path.join(download_path, f))]
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return files[0] if files else None


async def send_daily_report():
    download_path = 'C:\\Users\\gamag\\Downloads'
    latest_file = get_latest_downloaded_file(download_path)

    if not latest_file:
        print("Файл не найден.")
        return

    df = pd.read_excel(latest_file)
    report_time = datetime.now().strftime("%d.%m.%Y %H:%M")
    messages_by_oiv_id = {}

    for index, row in df.iterrows():
        if (row.get("Просрок (Монитор)") == 'Да' and
            row.get("Статус подготовки ответа на сообщение") == "Нет ответа"):

            responsible = row.get("Ответственный за подготовку ответа")
            oiv_id = OIV_IDS.get(responsible)

            if oiv_id:
                message_number = row.get("Номер сообщения")
                monitor_date = pd.to_datetime(row.get("Дата отображения (Монитор)"), errors='coerce')
                monitor_time = monitor_date.strftime("%d.%m.%Y %H:%M") if pd.notna(monitor_date) else "неизвестно"
                if oiv_id not in messages_by_oiv_id:
                    messages_by_oiv_id[oiv_id] = []
                messages_by_oiv_id[oiv_id].append(f"{message_number} (время просрока: {monitor_time})")

    for oiv_id, message_numbers in messages_by_oiv_id.items():
        if oiv_id in USERS_BY_OIV_ID:
            users = USERS_BY_OIV_ID[oiv_id]
            message_numbers_text = "\n".join(message_numbers)
            for user_info in users:
                username = user_info['username']
                name = user_info['name']
                message = (
                    f"Здравствуйте, {name}!\n"
                    f"На основании пункта 1.1 распоряжения префектуры Юго-Восточного административного округа города Москвы "
                    f"от 06.05.2024 № Р-132/24 об исполнительной дисциплине при отработке сообщений граждан, поступающих на "
                    f"централизованный портал 'Наш Город',\n"
                    f"по состоянию на {report_time} не было предоставлено ответов на следующие сообщения:\n"
                    f"{message_numbers_text}.\n"
                    f"В срок до 23:00 сегодняшнего дня представьте на моё имя за своей подписью объяснительные в аппарат префекта, касательно каждого "
                    f"из вышеуказанных сообщений, с приложением приказов о наказании должностных лиц, допустивших нарушения.\n"
                    f"Обеспечьте персональную ответственность и постоянный контроль."
                )
                await client.send_message(username, message)
                print(f"Сообщение отправлено {username}!")


async def schedule_daily_report():
    while True:
        now = datetime.now()
        if now.hour == 13 and now.minute == 0:
            await send_daily_report()
            await asyncio.sleep(60)
        await asyncio.sleep(10)

with client:
    client.loop.run_until_complete(schedule_daily_report())
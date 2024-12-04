import asyncio
import os
import random
from datetime import datetime, timedelta
import openpyxl
from telethon import TelegramClient
from telethon.errors.rpcerrorlist import PhoneNumberInvalidError

api_id = 'id'
api_hash = 'hash'
phone_number = 'number'

client = TelegramClient('session_name2', api_id, api_hash)

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

CATEGORY_IDS = {
    'Парки/скверы': 101,
    'ДТ': 102,
    'МКД': 103,
    'ОДХ': 104
}

USERS_BY_OIV_ID = {
    4: [{'username': 'phone/id', 'name': 'Сергей Владимирович'}],
}

DOWNLOAD_PATH = 'C:\\Users\\gamag\\Downloads'
LOG_FILE = "notifications_log.txt"
REPORT_FILE = "notifications_report.docx"

MAX_RUNS_PER_DAY = 3
last_runs = []

def log_notification(user_name, user_id, message):
    formatted_time = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as log_file:
        log_file.write(f"{formatted_time} - {user_name} ({user_id}): {message}\n")


def get_latest_downloaded_file(download_path):
    files = [os.path.join(download_path, f) for f in os.listdir(download_path) if
             os.path.isfile(os.path.join(download_path, f))]
    return max(files, key=os.path.getctime) if files else None

def format_date(date):
    return date.strftime('%H:%M:%S %d.%m.%Y')


def escape_markdown_v2(text):
    return text.replace('_', '\\_').replace('*', '\\*').replace('[', '\\[').replace(']', '\\]')


def process_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    notifications = {}
    today = datetime.now()

    valid_statuses = ["Нет ответа", "Готовится ответ", "На доработке"]
    found_user_ids_and_names = set()

    for row in sheet.iter_rows(min_row=2):
        status = row[20].value
        if status not in valid_statuses:
            continue

        oiv_name = row[18].value
        oiv_id = OIV_IDS.get(oiv_name)

        if oiv_id is None:
            continue

        address = row[6].value
        message_id = row[1].value
        problem_topic = row[12].value
        publication_date = row[3].value
        monitor_flag = row[34].value
        deadline1 = row[19].value
        deadline2 = row[48].value
        object_id = row[9].value
        category_name = row[37].value
        category_id = CATEGORY_IDS.get(category_name)

        deadline = deadline1 if monitor_flag == 'Да' and status == "Готовится ответ" else deadline2 if monitor_flag == 'Да' else deadline1

        if isinstance(deadline, datetime):
            time_left = deadline - today
            if timedelta(hours=0) < time_left < timedelta(hours=2):
                for user in USERS_BY_OIV_ID.get(oiv_id, []):
                    user_name = user.get('name', "Уважаемый пользователь")
                    user_id = user.get('username')
                    if (user_id, user_name) in found_user_ids_and_names:
                        continue

                    message = (
                        f"{user_name}, у тебя в течении 2ух часов будет просрочка в системе мониторинга. Прошу обратить внимание и устранить нарушение, согласно поручению Мэра."
                    )
                    escaped_message = escape_markdown_v2(message)

                    if user_id not in notifications:
                        notifications[user_id] = []
                    notifications[user_id].append(escaped_message)
                    found_user_ids_and_names.add((user_id, user_name))
    return notifications


async def notify_users(notifications):
    for user_id, messages in notifications.items():
        for message in messages:
            user_name = next((user['name'] for user in USERS_BY_OIV_ID.get(user_id, []) if user['username'] == user_id), None)
            try:
                await client.send_message(user_id, message)
                print(f"Сообщение отправлено {user_name} ({user_id}): {message}")
                log_notification(user_name, user_id, message)
            except Exception as e:
                print(f"Ошибка при отправке сообщения пользователю {user_id}: {e}")

def get_random_time():
    random_hour = random.randint(8, 22)
    random_minute = random.randint(0, 59)
    return datetime.now().replace(hour=random_hour, minute=random_minute, second=0, microsecond=0)

def generate_scheduled_times():
    user_schedules = {}
    for oiv_id, users in USERS_BY_OIV_ID.items():
        for user in users:
            times = []
            num_times = random.randint(2, 3)
            for _ in range(num_times):
                while True:
                    time = get_random_time()
                    if all(abs((time - t).total_seconds()) >= 2 * 3600 for t in times):
                        times.append(time)
                        break
            user_schedules[user['username']] = sorted(times)
    return user_schedules


async def schedule_notifications(user_schedules):
    print("Начало планирования уведомлений.")
    while True:
        now = datetime.now()
        for user_id, schedule in list(user_schedules.items()):
            while schedule and now >= schedule[0]:
                send_time = schedule.pop(0)
                print(f"Отправляем сообщение пользователю {user_id} в {send_time}")
                try:
                    user_name = next((user['name'] for oiv_users in USERS_BY_OIV_ID.values() for user in oiv_users if
                                      user['username'] == user_id), "Уважаемый пользователь")
                    message = f"{user_name}, у тебя в течении 2ух часов будет просрочка в системе мониторинга. Прошу обратить внимание и устранить нарушение, согласно поручению Мэра."

                    notifications = {user_id: [message]}
                    await notify_users(notifications)
                except Exception as e:
                    print(f"Ошибка при отправке сообщения пользователю {user_id}: {e}")
            if not schedule:
                del user_schedules[user_id]
        if not user_schedules:
            print("Все запланированные уведомления отправлены.")
            break
        await asyncio.sleep(30)

async def main():
    print("Генерация расписания отправки.")
    user_schedules = generate_scheduled_times()
    for user, times in user_schedules.items():
        print(f"Расписание для {user}: {[t.strftime('%H:%M:%S') for t in times]}")

    await schedule_notifications(user_schedules)


if __name__ == "__main__":
    try:
        with client:
            print("Telegram клиент подключен.")
            client.loop.run_until_complete(main())
    except PhoneNumberInvalidError:
        print("Неверный номер телефона.")
    except Exception as e:
        print(f"Произошла ошибка: {e}")
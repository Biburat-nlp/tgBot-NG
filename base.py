import os
import time
from datetime import datetime, timedelta, time as dt_time
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import openpyxl
from telegram import Bot
import asyncio

TELEGRAM_API_TOKEN = 'token'

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

USER_OIV_MAP = {
    '1008683615': {
        'oiv_ids': [12, 13],
        'category_ids': [101, 102, 103, 104]
    },
    #'309204640': {'oiv_ids':[14, 15, 24, 25], 'category_ids': [101,102,103,104]},  # Александр Геннадьевич
    #'248001485': {'oiv_ids':[14, 15, 24, 25], 'category_ids': [101,102,103,104]},  # Максим Боженко Куратор (Люблино, Лефортово, Южнопортовый)
    '342617808': {'oiv_ids':[3], 'category_ids': [101,102,103,104]}, # Юрий АВД
    '5771868721': {'oiv_ids':[12, 13], 'category_ids': [101,102,103,104]}, # Владислав Лефортово
    #'444818192': {'oiv_ids':[16, 17], 'category_ids': [101,102,103,104]}, # Руслан Мегамозг
    #'773634578': {'oiv_ids':[16, 17], 'category_ids': [101,102,103,104]}, #Ксения Витальевна
    '1859497322': {'oiv_ids':[16, 17], 'category_ids': [101,102,103,104]}, #Виктория Мочалова (Некрасовка)
    '2124882080': {'oiv_ids':[8, 9], 'category_ids': [101,102,103,104]}, #Вороненко Константин (Капонтня)
    '1104241617': {'oiv_ids':[6, 7], 'category_ids': [101,102,103,104]}, #Ромашкина Мария (Нижегородский)
    '748848833': {'oiv_ids':[18], 'category_ids': [103]}, #Ярослав Путенцов(Печатники)
    '1968513890': {'oiv_ids':[16, 17], 'category_ids': [101,102,103,104]}, #Алиева Татьяна(Некрасовка)
    '682768205': {'oiv_ids':[22, 23], 'category_ids': [101,102,103,104]}, #Бочкова Алёна Владиславовна (Рязанский Управа)
    '1159531079': {'oiv_ids':[3], 'category_ids': [101,102,103,104]}, #Екатерина Лазарева (АВД)
    '637773050': {'oiv_ids':[16, 17], 'category_ids': [101, 102, 103, 104]}, #Кальдинова Наталья Игоревна (Некрасовка)
    '1118256309': {'oiv_ids':[6, 7], 'category_ids': [101,102,103,104]}, #Будаева Анастасия Николаевна (Нижегородский Управа)
    '1747380648': {'oiv_ids':[1, 2], 'category_ids': [101,102,103,104]}, #Деминова Людмила Ивановна (Марьино)
    '879277421': {'oiv_ids':[21], 'category_ids': [101,102,103,104]}, #Симоненкова Елена Сергеевна (Управа Текстильщики)
    '498139669': {'oiv_ids':[4, 5], 'category_ids': [101,102,103,104]}, #Косов Олег (Выхино Жулебино Управа)
    '1779374656': {'oiv_ids':[22], 'category_ids': [101,102,103,104]}, #Растанова Анастасия Александровна (Рязанский гбу)
    '214038510': {'oiv_ids':[20, 21], 'category_ids': [101,102,103,104]}, #Проценко Леонид Дмитриевич (Текстильщики управа)
    '259685801': {'oiv_ids':[21], 'category_ids': [101,102,103,104]}, #Павлова Анастасия Евгеньевна (Текстильщики управа)
    '1304016323': {'oiv_ids':[4, 5], 'category_ids': [101,102,103,104]}, #Комова Екатерина Германовна(Выхино-Жулебино)
    '2111490716': {'oiv_ids':[12, 13], 'category_ids': [101,102,103,104]}, #Соломатина Елена (Лефортово ГБУ)
    '957128044': {'oiv_ids':[24, 25], 'category_ids': [101,102,103,104]}, #Малородова Евгения (Южнопортовый)
    '5681222962': {'oiv_ids':[10], 'category_ids': [103]}, #Сержантова Анна (Кузьминки ГБУ)
    '199143077': {'oiv_ids':[3], 'category_ids': [101,102,103,104]},  #Титов Сергей (АВД)
    '184804530': {'oiv_ids':[20], 'category_ids': [101,102,103,104]},  #Мацеевский Максим (Текстильщики ГБУ)
    '284643670': {'oiv_ids':[12, 13], 'category_ids': [101,102,103,104]},  #Екатерина Карпушина (Лефортово Зам главы Управы)
    '891811153': {'oiv_ids':[12], 'category_ids': [101,102,103,104]},  #Сергеева Екатерина (Лефортово глава ГБУ)
    '1787572712': {'oiv_ids':[20], 'category_ids': [101,102,103,104]},  #Ворожейкина Светлана (Текстильщики ГБУ)
    '5807151468': {'oiv_ids':[18, 19], 'category_ids': [102]},  #Амет Анна (Печатники дворы)
    '1726689160': {'oiv_ids':[14, 15], 'category_ids': [101,102,103,104]}, #Кто-то из Люблино
    '1461794477': {'oiv_ids':[3], 'category_ids': [101,102,103,104]}, #Землянский Максим Начальник СМЦ ГБУ
    '5227822667': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то из Кузьминок
    '777757256': {'oiv_ids':[18, 19], 'category_ids': [101,102,103,104]}, #Елизавета Печатники
    '6322616283': {'oiv_ids':[14, 15], 'category_ids': [101,102,103,104]}, #Пруцкова Александра Ильинична Люблино
    '737424515': {'oiv_ids':[14, 15], 'category_ids': [103]}, #Калушка Ольга Владимировна Люблино
    '916001760': {'oiv_ids':[14, 15], 'category_ids': [102,103]}, #Зайцева Полина Сергеевна Люблино
    '1104172214': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то из Кузьминок
    '990914503': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то из Кузьминок
    '5213710149': {'oiv_ids':[18, 19], 'category_ids': [101,102,103,104]}, #Кто-то Печатники
    '1208819533': {'oiv_ids':[3], 'category_ids': [101,102,103,104]}, #Кто-то из АВД
    '1065318763': {'oiv_ids':[10, 11], 'category_ids': [101,102,104]}, #Абрамов Никита Сергеевич Кузьминки
    '1479535368': {'oiv_ids':[6, 7], 'category_ids': [101,102,103,104]}, #Кто-то (Нижегородский)
    '1810562708': {'oiv_ids':[6, 7], 'category_ids': [101,102,103,104]}, #Кто-то (Нижегородский)
    '433713437': {'oiv_ids':[6, 7], 'category_ids': [101,102,103,104]}, #Киряков Александр Владиславович (Нижегородский)
    '399125424': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то из Кузьминок
    '1333864717': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то из Кузьминок
    '534595947': {'oiv_ids':[10, 11], 'category_ids': [103]}, #Кто-то из Кузьминок
    '5691778916': {'oiv_ids':[16, 17], 'category_ids': [101,102,103,104]}, #Зотов Сергей (Некрасовка. Зам.Главы Управы)
    '1958311512': {'oiv_ids':[14, 15], 'category_ids': [101,102,103,104]},  # Дурова Элина (Люблино Зам. Главы Управы)
    '1201993583': {'oiv_ids':[8, 9], 'category_ids': [101,102,103,104]}, #Шитиков Михаил (Капотня Зам. Главы Управы)
    '738501775': {'oiv_ids':[16, 17], 'category_ids': [101,102,103,104]}, #Хромова Елена (Некрасовка Глава Управы)
    '139355601': {'oiv_ids':[20, 21], 'category_ids': [101,102,103,104]}, #Кто-то Текстильщики
    '405402236': {'oiv_ids':[6, 7], 'category_ids': [101,102,103,104]}, #Хозяенок Игорь (Нижегородский Зам. Главы Управы)
    '266830829': {'oiv_ids':[24, 25], 'category_ids': [101,102,103,104]}, #Квачахия Рональд (Южнопортовый Зам. Главы Управы)
    '7383347962': {'oiv_ids':[4, 5], 'category_ids': [101,102,103,104]}, #Дивин Александр (Выхино-Жулебино Руководитель ГБУ)
    '5220557356': {'oiv_ids':[18, 19], 'category_ids': [101,102,103,104]}, #Кузьмичёв Алексей (Печатники Зам. Главы Управы)
    '2140153164': {'oiv_ids':[18, 19], 'category_ids': [101,102,103,104]}, #Кто-то (Печатники)
    '541281446': {'oiv_ids':[4, 5], 'category_ids': [101,102,103,104]}, #Кто-то (Выхино)
    '1159278532': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '458394303': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '1442199120': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '1615407360': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '1926075202': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '6778633208': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '5650866862': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]},  # Кто-то (Кузьминки)
    '420412441': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '221474889': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '5293442369': {'oiv_ids':[12, 13], 'category_ids': [101,102,103,104]}, #Кто-то (Лефортово)
    '7428321359': {'oiv_ids':[12, 13], 'category_ids': [101,102,103,104]}, #Начальник 8ого участка (Лефортово)
    '1677640950': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '755792631': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '6034511872': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '651322495': {'oiv_ids':[11], 'category_ids': [101,102,103,104]}, #Кто-то (управа Кузьминки)
    '6468976698': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '1530446114': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '5288775403': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '7529825867': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '5246328211': {'oiv_ids':[12, 13], 'category_ids': [101]}, #Почивалова Елена Викторовна (Лефортово)
    '1045446604': {'oiv_ids':[4, 5], 'category_ids': [101,102,103,104]}, #Кто-то (Выхино-Жулебино)
    '5108733398': {'oiv_ids':[10, 11], 'category_ids': [101,102,103,104]}, #Кто-то (Кузьминки)
    '690416303': {'oiv_ids':[12, 13], 'category_ids': [101,102,104]}, #Кто-то(Лефортово)
    '1348896184': {'oiv_ids':[4, 5], 'category_ids': [101,102,103,104]}, #Кто-то (Выхино-Жулебино)
    '6698335587': {'oiv_ids':[12, 13], 'category_ids': [102,103]}, #начальник участка
    '6993367364': {'oiv_ids':[12, 13], 'category_ids': [102,103]}, #начальник участка
    '5226409299': {'oiv_ids':[12, 13], 'category_ids': [101, 102, 103, 104]}, #КТо то Лефортово
    '5082179002': {'oiv_ids':[6, 7], 'category_ids': [101, 102, 103, 104]}, #Кто-то Нижегородский
}

# Параметры для входа на сайт
URL = "https://gorod.mos.ru/api/service/auth/auth"
USERNAME = "login"
PASSWORD = "password"
DOWNLOAD_PATH = 'C:\\Users\\Portal\\Downloads'
CHROMEDRIVER_PATH = "C:/Users/Portal/Downloads/chromedriver-win64/chromedriver.exe"

# Настройка Selenium
options = webdriver.ChromeOptions()
options.add_argument('--disable-gpu')
options.add_argument("--window-size=1920,1080")
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_PATH,
    "download.prompt_for_download": False,
})

service = Service(executable_path=CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

logged_in = False


def login(driver):
    global logged_in
    if not logged_in:
        driver.get(URL)
        driver.find_element(By.ID, 'input-8').send_keys(USERNAME)
        driver.find_element(By.ID, 'input-9').send_keys(PASSWORD)
        driver.find_element(By.XPATH, '//*[@id="app"]/div/main/div/div/div/form[1]/button').send_keys(Keys.RETURN)
        time.sleep(5)
        logged_in = True


def download_excel(driver):
    driver.find_element(By.XPATH, '//*[@id="app"]/div/main/div/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div/div/div').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[1]/aside/div/div[2]/div[6]/a/div[2]/div').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div/div/div/form/footer/button[3]').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/button[2]/span[2]').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="q-app"]/div/header/div[1]/div[1]/div/span').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="app"]/div/main/div/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div/div/div').click()
    time.sleep(90)
    driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/main/div/div[1]/div[2]/div[1]/table/tbody/tr[1]/td[5]').click()
    time.sleep(5)


def get_latest_downloaded_file(download_path):
    files = os.listdir(download_path)
    files = [os.path.join(download_path, f) for f in files if os.path.isfile(os.path.join(download_path, f))]
    files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return files[0] if files else None


def escape_markdown_v2(text):
    escape_chars = ['_', '*', '[', ']', '(', ')', '~', '`', '>', '#', '+', '-', '=', '|', '{', '}', '.', '!']
    for char in escape_chars:
        text = text.replace(char, '\\' + char)
    return text


def format_date(date):
    return date.strftime('%H:%M:%S %d.%m.%Y')


def process_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    notifications = {}
    today = datetime.now()

    valid_statuses = ["Нет ответа", "Готовится ответ", "На доработке"]

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
        monitor_flag = row[33].value
        deadline1 = row[19].value
        deadline2 = row[47].value
        object_id = row[9].value
        category_name = row[36].value
        category_id = CATEGORY_IDS.get(category_name)

        if monitor_flag == 'Да' and status == "Готовится ответ":
            deadline = deadline1
        else:
            deadline = deadline2 if monitor_flag == 'Да' else deadline1

        if isinstance(deadline, datetime):
            time_left = deadline - today
            if timedelta(hours=0) < time_left < timedelta(hours=2):
                if monitor_flag == 'Да':
                    message = (f"*Монитор* - До просрока менее двух часов! {format_date(deadline)}\n"
                               f"Адрес: {address}\n"
                               f"Тема: {problem_topic}\n"
                               f"Номер сообщения: {message_id}\n"
                               f"Категория: {category_name}\n"
                               f"Ответственный ОИВ: {oiv_name}\n"
                               f"https://gorod.mos.ru/objects/{object_id}/messages#{message_id}")
                else:
                    message = (f"До просрока менее двух часов! {format_date(deadline)}\n"
                               f"Адрес: {address}\n"
                               f"Тема: {problem_topic}\n"
                               f"Номер сообщения: {message_id}\n"
                               f"Категория: {category_name}\n"
                               f"Ответственный ОИВ: {oiv_name}\n"
                               f"https://gorod.mos.ru/objects/{object_id}/messages#{message_id}")

                escaped_message = escape_markdown_v2(message)

                matched_users = []
                for user_id, id_map in USER_OIV_MAP.items():
                    oiv_list = id_map['oiv_ids']
                    if monitor_flag == 'Нет':
                        if oiv_id in oiv_list:
                            matched_users.append(user_id)
                    else:
                        if oiv_id in oiv_list:
                            matched_users.append(user_id)
                if monitor_flag == 'Да':
                    for user_id in matched_users:
                        category_list = USER_OIV_MAP[user_id]['category_ids']
                        if category_id in category_list:
                            if user_id not in notifications:
                                notifications[user_id] = []
                            notifications[user_id].append(escaped_message)
                else:
                    for user_id in matched_users:
                        if user_id not in notifications:
                            notifications[user_id] = []
                        notifications[user_id].append(escaped_message)
    return notifications


async def notify_users(notifications):
    bot = Bot(token=TELEGRAM_API_TOKEN)
    for user_id, messages in notifications.items():
        for message in messages:
            try:
                await bot.send_message(chat_id=user_id, text=message, parse_mode='MarkdownV2')
            except Exception as e:
                print(f"Failed to send message to {user_id}: {e}")


def return_to_main_page(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="q-app"]/div/header/div[1]/div[1]/div/span').click()
    except Exception as e:
        print(f"Ошибка при возврате на главную страницу: {e}")


def main_loop():
    global logged_in
    login(driver)
    try:
        current_time = datetime.now().time()
        start_time = dt_time(8, 0)
        end_time = dt_time(23, 0)

        # В рабочие часы: скачиваем файл и отправляем уведомления
        if start_time <= current_time <= end_time:
            print("Скачиваем файл...")
            download_excel(driver)
            time.sleep(5)

            latest_file = get_latest_downloaded_file(DOWNLOAD_PATH)
            if latest_file:
                print("Файл найден. Обрабатываем уведомления...")
                notifications = process_excel(latest_file)
                if notifications:
                    print("Отправляем уведомления пользователям...")
                    asyncio.run(notify_users(notifications))
                    print("Уведомления отправлены!")
                else:
                    print("Нет уведомлений для отправки.")
            else:
                print("Файл не найден!")

            return_to_main_page(driver)

        # В нерабочие часы: только скачиваем файл без уведомлений
        else:
            print("Тихий режим: только скачивание файла.")
            print("Скачиваем файл в тихом режиме...")
            download_excel(driver)
            time.sleep(5)

            latest_file = get_latest_downloaded_file(DOWNLOAD_PATH)
            if latest_file:
                print("Файл успешно скачан.")
            else:
                print("Файл не найден!")

            return_to_main_page(driver)

    except Exception as e:
        print(f"Произошла ошибка: {e}")


def wait_until_next_hour():
    now = datetime.now()
    next_hour = (now + timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
    time_to_wait = (next_hour - now).total_seconds()
    time.sleep(time_to_wait)


while True:
    main_loop()
    wait_until_next_hour()
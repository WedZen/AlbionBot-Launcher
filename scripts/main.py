import datetime
import random
import time
import cv2
import numpy as np
import pyautogui
import pygetwindow as gw
from PIL import ImageGrab
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pyperclip
import easyocr
import os
import sys
import codecs

sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
sys.stderr = codecs.getwriter("utf-8")(sys.stderr.detach())

sheet_name = sys.argv[1]
sheet_name_data = sheet_name[:-1]


def handle_quality_and_levels(item_name, quality, levels, batch_main, batch_data):
    sell_order, buy_order = None, None

    if quality != "None":
        click_button("quality")
        click_button(quality)
    if levels != ["None"]:
        for level in levels:
            click_button("@")
            click_button(level)
            time.sleep(random.uniform(0.2, 0.2))

            if pyautogui.pixelMatchesColor(1033, 327, (193, 161, 130)):
                sell_order = read_price_sell_order(1004, 311, 1072, 331)
                buy_order = read_price_buy_order(1299, 307, 1366, 331)

            if pyautogui.pixelMatchesColor(1014, 327, (189, 158, 127)) and pyautogui.pixelMatchesColor(1047, 327,
                                                                                                       (187, 156, 125)):
                sell_order = read_price_sell_order(1004, 312, 1084, 331)
                buy_order = read_price_buy_order(1298, 309, 1380, 331)
            else:

                sell_order = read_price_sell_order(1004, 312, 1084, 331)
                buy_order = read_price_buy_order(1298, 309, 1380, 331)

            if not sell_order or not buy_order:
                continue

            if float(sell_order) == 0:
                print(f"Пропускаем расчет для уровня {level} (sell_order равно нулю)")
                continue

            margin = calculate_margin(sell_order, buy_order)
            ratioM_B = float(margin) / float(buy_order)

            # if ratioM_B >= min_ratios and int(buy_order) >= min_buy_order or ratioM_B >= min_ratios and margin > 100000:
            time.sleep(random.uniform(0.1, 0.1))

            if margin >= min_margin and ratioM_B >= min_ratios:

                if check_pixels(pixels_to_check):

                    if margin >= 1000000000:
                        continue

                    click_button("volume_1")
                    time.sleep(random.uniform(0.5, 0.5))
                    volume_1 = read_item_volume()
                    click_button("volume_2")
                    time.sleep(random.uniform(0.5, 0.5))
                    volume_2 = read_item_volume()
                    click_button("volume_3")
                    time.sleep(random.uniform(0.5, 0.5))
                    volume_3 = read_item_volume()

                    volume_average = average_value([volume_1, volume_2, volume_3])

                    if volume_average >= min_volume:

                        # Отримуємо поточний час
                        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        # Додаємо до основного масиву
                        batch_main.append([
                            item_name, float(level[-1]), quality, float(volume_average)
                        ])

                        # Додаємо до розширеного масиву
                        batch_data.append([
                            item_name, float(level[-1]), quality, float(sell_order), float(buy_order), float(margin),
                            float(volume_average),
                            current_time
                        ])

                        print(
                            f"Предмет:{item_name} {float(level[-1])} {quality}, Продажа: {sell_order}, Покупка: {buy_order}, "
                            f"Объем: {volume_average}, Профит: {margin}, Время: {current_time}")

                    else:
                        continue
            else:
                continue


button_coordinate = {

    "clear": ((624, 177), (657, 214)),
    "search": ((492, 190), (612, 204)),
    "order": ((1245, 371), (1371, 412)),

    "quality": ((686, 332), (846, 363)),
    "common_quality": ((690, 368), (847, 395)),
    "good_quality": ((688, 402), (844, 425)),
    "outstanding_quality": ((689, 438), (845, 460)),
    "excellent_quality": ((690, 470), (846, 492)),
    "masterpiece_quality": ((688, 501), (845, 523)),

    "@": ((502, 333), (656, 362)),
    "@0": ((503, 370), (657, 396)),
    "@1": ((506, 403), (656, 424)),
    "@2": ((505, 436), (659, 457)),
    "@3": ((502, 468), (662, 492)),

    "volume_1": ((1035, 686), (1188, 692)),
    "volume_2": ((1198, 689), (1318, 691)),
    "volume_3": ((1327, 687), (1462, 691)),

    "close": ((893, 231), (915, 251)),
}

pixels_to_check = [
    (1037, 675, (102, 74, 58)),
    (1053, 675, (102, 74, 58)),
    (1070, 675, (102, 74, 57)),
    (1087, 675, (102, 74, 57)),
    (1104, 675, (102, 74, 57)),
    (1120, 675, (102, 74, 58)),
    (1136, 675, (102, 74, 57)),
    (1152, 675, (102, 74, 58)),
    (1167, 675, (102, 74, 58)),
    (1184, 675, (102, 74, 57)),
    (1200, 675, (102, 75, 58)),
    (1216, 675, (102, 75, 58)),
    (1232, 675, (102, 74, 58)),
    (1248, 675, (102, 74, 58)),
    (1264, 675, (102, 74, 58)),
    (1280, 675, (102, 74, 58)),
    (1296, 675, (102, 74, 58)),
    (1312, 675, (102, 74, 58)),
    (1328, 675, (102, 74, 58)),
    (1344, 675, (102, 74, 58)),
    (1360, 675, (102, 74, 58)),
    (1376, 675, (102, 74, 58)),
    (1392, 675, (102, 74, 58)),
    (1408, 675, (102, 74, 58)),
    (1424, 675, (102, 74, 58)),
    (1440, 675, (102, 74, 58)),
    (1456, 675, (102, 74, 58))
]


def activate_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        if window.isMinimized:
            print(f"Окно {window_title} свернуто")
            window.restore()
        window.activate()
        return window
    except IndexError:
        print(f"Окно {window_title} не найдено")
        return None
    except Exception as e:
        print(f"Ошибка при активации окна: {window_title}")
        return None


def enlarge_image(image_np, scale=2.0):
    height, width = image_np.shape[:2]
    new_height, new_width = int(height * scale), int(width * scale)
    enlarged_image = cv2.resize(image_np, (new_width, new_height), interpolation=cv2.INTER_CUBIC)
    return enlarged_image


def capture_window(window_title):
    window = activate_window(window_title)
    if window:
        bbox = (window.left, window.top, window.right, window.bottom)
        screenshot = ImageGrab.grab(bbox)
        screenshot.save(f"{window_title.lower().replace(' ', '_')}_screenshot.png")
        return screenshot
    return None


def click_button(name, button='left', duration=0.1):
    (x1, y1), (x2, y2) = button_coordinate[name]
    final_x = random.uniform(x1, x2)
    final_y = random.uniform(y1, y2)
    final_duration = random.uniform(0.1, 0.1)

    pyautogui.moveTo(final_x, final_y, final_duration)
    pyautogui.click(button=button)


def suppress_gpu_warning():
    sys.stderr = open(os.devnull, 'w')  # Перенаправляем предупреждения в "пустоту"
    reader = easyocr.Reader(['en'])  # Инициализируем EasyOCR
    sys.stderr = sys.__stderr__  # Возвращаем стандартный вывод ошибок
    return reader


reader = suppress_gpu_warning()


def read_price_sell_order(x1, y1, x2, y2):
    bbox = (x1, y1, x2, y2)
    screenshot = ImageGrab.grab(bbox)
    image_np = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    enlarged_image = enlarge_image(image_np)

    # EasyOCR для извлечения текста
    result = reader.readtext(enlarged_image, detail=0, allowlist='0123456789,.')

    sell_order_result = ''.join(filter(str.isdigit, ''.join(result)))
    return sell_order_result


def read_price_buy_order(x1, y1, x2, y2):
    bbox = (x1, y1, x2, y2)
    screenshot = ImageGrab.grab(bbox)
    image_np = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    enlarged_image = enlarge_image(image_np)

    # EasyOCR для извлечения текста
    result = reader.readtext(enlarged_image, detail=0, allowlist='0123456789,.')

    buy_order_result = ''.join(filter(str.isdigit, ''.join(result)))
    return buy_order_result


def calculate_margin(sell_order, buy_order):
    sell_order = int(sell_order)
    buy_order = int(buy_order)
    return sell_order - buy_order


def check_pixels(pixels):
    for x, y, color in pixels:
        if not pyautogui.pixelMatchesColor(x, y, color, tolerance=10):  # Можно указать tolerance, если есть отклонения
            return False
    return True


offset_x = -250
offset_y = -30
width = 250
height = 25


def read_item_volume():
    attempts = 2
    for attempt in range(attempts):
        # Получение позиции курсора и расчет области захвата
        cursor_x, cursor_y = pyautogui.position()
        bbox = (cursor_x + offset_x, cursor_y + offset_y, cursor_x + offset_x + width,
                cursor_y + offset_y + height)
        # Захват изображения
        screenshot = ImageGrab.grab(bbox)
        image_np = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        # EasyOCR для извлечения текста
        result = reader.readtext(image_np, detail=0)

        # Фильтрация и соединение результата
        volume_result = ''.join(filter(str.isdigit, ''.join(result)))

        if volume_result:
            formatted_result = volume_result.replace('.', ',')
            return formatted_result

    return str(random.randint(1, 4))


def average_value(list_of_values):
    numeric_values = [int(value.replace(',', '')) for value in list_of_values if value.replace(',', '').isdigit()]
    return sum(numeric_values) / len(numeric_values) if numeric_values else 0


def get_item_name_from_sheets():
    credentials_path = 'C:/Users/Urusalim/Documents/Google/marginalbion-16265092a9d2.json'
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(credentials)

    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/11lMYkckati6qLuyWDjIvNzt3QTdOatAe6Gb72ujE_6Q')
    worksheet = sheet.worksheet("NameItem")

    return worksheet.col_values(1)


activate_window('Albion Online Client')
capture_window('Albion Online Client')

min_volume = 5
min_margin = 30000
min_ratios = 0.3

# Словник відповідності назв міст і таблиць Google Sheets
city_to_sheet_mapping = {
    "Lumhurst": ("ResultLHw", "ResultLH"),
    "Bridgewatch": ("ResultBWw", "ResultBW"),
    "Martlock": ("ResultMLw", "ResultML"),
    "Thetford": ("ResultTFw", "ResultTF"),
    "FortSterling": ("ResultFSw", "ResultFS"),
    "Briceleon": ("ResultBSw", "ResultBS")
}


def process_item(sheet_name, sheet_name_data):
    try:
        item_names = get_item_name_from_sheets()

        credentials_path = 'C:/Users/Urusalim/Documents/Google/marginalbion-16265092a9d2.json'
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        client = gspread.authorize(credentials)

        sheet = client.open_by_url(
            'https://docs.google.com/spreadsheets/d/11lMYkckati6qLuyWDjIvNzt3QTdOatAe6Gb72ujE_6Q')
        worksheet = sheet.worksheet(sheet_name)

        sheet_data = client.open_by_url(
            'https://docs.google.com/spreadsheets/d/1JGtr5opjzX7B1VKGXgfuubgy51RjGNo3cYX86i0T1w8').worksheet(
            sheet_name_data)

        # Очищення старих даних в обох листах
        print("Очистка старих даних з таблиць...")
        worksheet.clear()
        sheet_data.clear()
        # основні данні локально
        batch_main = []
        # Зберігаємо ордери локально
        batch_data = []

        for item_name in item_names:
            try:
                while True:

                    pyperclip.copy(item_name)

                    click_button("clear")
                    click_button("search")

                    if pyautogui.pixelMatchesColor(481, 200, (230, 222, 212), tolerance=10):

                        pyautogui.hotkey('ctrl', 'v')
                        click_button("order")

                        time.sleep(random.uniform(0.2, 0.2))

                        if pyautogui.pixelMatchesColor(620, 351, (194, 165, 132),
                                                       tolerance=10) and pyautogui.pixelMatchesColor(816, 350,
                                                                                                     (202, 171, 137),
                                                                                                     tolerance=10):
                            handle_quality_and_levels(item_name, "common_quality", ["@0", "@1", "@2", "@3"], batch_main,
                                                      batch_data)
                            handle_quality_and_levels(item_name, "good_quality", ["@0", "@1", "@2", "@3"], batch_main,
                                                      batch_data)
                            handle_quality_and_levels(item_name, "outstanding_quality", ["@0", "@1", "@2", "@3"],
                                                      batch_main, batch_data)
                            handle_quality_and_levels(item_name, "excellent_quality", ["@0", "@1", "@2", "@3"],
                                                      batch_main, batch_data)

                            click_button("close")
                            break

                        if pyautogui.pixelMatchesColor(623, 353, (123, 123, 123),
                                                       tolerance=10) and pyautogui.pixelMatchesColor(818, 351,
                                                                                                     (194, 165, 132),
                                                                                                     tolerance=10):
                            handle_quality_and_levels(item_name, "common_quality", ["None"], batch_main, batch_data)
                            handle_quality_and_levels(item_name, "good_quality", ["None"], batch_main, batch_data)
                            handle_quality_and_levels(item_name, "outstanding_quality", ["None"], batch_main,
                                                      batch_data)
                            handle_quality_and_levels(item_name, "excellent_quality", ["None"], batch_main, batch_data)

                            click_button("close")
                            break

                        if pyautogui.pixelMatchesColor(626, 351, (194, 165, 132),
                                                       tolerance=10) and pyautogui.pixelMatchesColor(818, 350,
                                                                                                     (136, 136, 136),
                                                                                                     tolerance=10):
                            handle_quality_and_levels(item_name, "None", ["@0", "@1", "@2", "@3"], batch_main,
                                                      batch_data)

                            click_button("close")
                            break

                        if pyautogui.pixelMatchesColor(634, 350, (136, 136, 136),
                                                       tolerance=10) and pyautogui.pixelMatchesColor(817, 351,
                                                                                                     (130, 130, 130),
                                                                                                     tolerance=10):

                            handle_quality_and_levels(item_name, "None", ["None"], batch_main, batch_data)

                            click_button("close")
                            break

                        else:
                            print("Пиксель не найден, повторяем попытку для текущего предмета...")
                            click_button("close")
                            time.sleep(0.1)

            except Exception as item_error:

                print(f"Ошибка при обработке предмета [{item_name}]: {item_error}", flush=True)

        # Пакетне оновлення даних у таблицях (без заголовків)
        if batch_main:
            print("Оновлення основних даних у таблиці...")
            worksheet.update("A1", batch_main)  # Дані оновлюються без заголовків

        if batch_data:
            print("Оновлення розширених даних у таблиці...")
            sheet_data.update("A1", batch_data)  # Дані оновлюються без заголовків


    except Exception as process_error:
        print(f"Критическая ошибка в process_item: {process_error}", flush=True)


def main():
    """
    Головний вхідний пункт скрипта.
    """
    try:
        # Лог аргументів
        print("Отримані аргументи:", sys.argv)

        # Перевірка обов'язкового аргументу міста
        if len(sys.argv) < 2:
            raise ValueError("Не передано ім'я міста як параметр.")

        # Обробка міста
        city = sys.argv[1]
        print(f"Скрипт запущено для міста: {city}")

        # Перевірка, чи місто є в словнику
        if city not in city_to_sheet_mapping:
            raise ValueError(f"Місто '{city}' не знайдено у списку дозволених.")

        # Встановлення назв таблиць
        sheet_name, sheet_name_data = city_to_sheet_mapping[city]

        # Виклик основної логіки
        process_item(sheet_name, sheet_name_data)

    except ValueError as val_err:
        print(f"Помилка: {val_err}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Невідома помилка: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

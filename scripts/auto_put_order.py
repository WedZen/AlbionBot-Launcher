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


def activate_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)[0]
        if window.isMinimized:
            print(f"Окно {window_title} свернуто.")
            window.restore()
        window.activate()
        return window
    except IndexError:
        print(f"Окно {window_title} не найдено.")
        return None
    except Exception as e:
        print(f"Ошибка при активации окна {window_title}: {e}")
        return None


def capture_window(window_title):
    window = activate_window(window_title)
    if window:
        bbox = (window.left, window.top, window.right, window.bottom)
        screenshot = ImageGrab.grab(bbox)
        screenshot.save(f"{window_title.lower().replace(' ', '_')}_screenshot.png")
        return screenshot
    return None


# Координаты кнопок
button_coordinate = {
    "clear": ((625, 178), (659, 213)),
    "search": ((490, 189), (607, 205)),
    "order": ((1245, 373), (1370, 413)),

    "quality": ((692, 339), (817, 359)),
    "common_quality": ((691, 372), (845, 392)),
    "good_quality": ((690, 403), (840, 425)),
    "outstanding_quality": ((691, 436), (842, 461)),
    "excellent_quality": ((693, 470), (845, 493)),
    "masterpiece_quality": ((691, 502), (846, 525)),

    "close": ((889, 231), (919, 258)),

    "plus": ((795, 618), (826, 649)),
    "count": ((453, 554), (476, 571)),
    "buy_orders": ((778, 732), (897, 773)),

    "buy_order_confirm_T": ((762, 540), (860, 567)),
    "buy_order_confirm_T_L": ((761, 562), (859, 586)),

    "buy_order_confirm_F": ((757, 583), (859, 611)),
    "buy_order_confirm_F_L": ((756, 602), (861, 635)),

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


def check_pixels(pixels):
    for x, y, color in pixels:
        if not pyautogui.pixelMatchesColor(x, y, color, tolerance=10):  # Можно указать tolerance, если есть отклонения
            return False
    return True


def click_button(name, button='left'):
    (x1, y1), (x2, y2) = button_coordinate[name]
    final_x = random.uniform(x1, x2)
    final_y = random.uniform(y1, y2)
    final_duration = random.uniform(0.1, 0.1)

    pyautogui.moveTo(final_x, final_y, final_duration)
    pyautogui.click(button=button)


def get_item_enchant_quality_from_sheets():
    credentials_path = 'C:/Users/Urusalim/Documents/Google/marginalbion-16265092a9d2.json'
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(credentials)

    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/11lMYkckati6qLuyWDjIvNzt3QTdOatAe6Gb72ujE_6Q')
    worksheet = sheet.worksheet("Filter")

    item_names = worksheet.col_values(2)[1:]
    enchants = worksheet.col_values(3)[1:]
    qualities = worksheet.col_values(4)[1:]
    volumes_raw = worksheet.col_values(5)[1:]

    # Очистка данных buy_order и volume
    cleaned_volume = []

    for raw_value in volumes_raw:
        try:
            vol_value = int(raw_value)
            if vol_value <= 7:
                cleaned_volume.append(1)
            elif vol_value <= 12:
                cleaned_volume.append(2)
            elif vol_value <= 17:
                cleaned_volume.append(3)
            else:
                vol_10_percent = max(1, int(vol_value * 0.1))  # Рассчитываем 10% от объема
                cleaned_volume.append(vol_10_percent)
        except ValueError:  # Обработать значение, которое не удалось преобразовать в int
            cleaned_volume.append(0)
        # Возвращаем собранные данные
    return item_names, enchants, qualities, volumes_raw, cleaned_volume


def combine_item_and_enchant(item_names, enchants):
    return [f"{name} {enchant}".strip() for name, enchant in zip(item_names, enchants)]


def handle_quality(quality):
    quality_buttons = {
        "common_quality": "common_quality",
        "good_quality": "good_quality",
        "outstanding_quality": "outstanding_quality",
        "excellent_quality": "excellent_quality",
        "masterpiece_quality": "masterpiece_quality",
    }

    if quality in quality_buttons:
        click_button("quality")
        click_button(quality_buttons[quality])


def search_and_order_item(item_name):
    try:

        if pyautogui.pixelMatchesColor(905, 240, (233, 185, 0), tolerance=10):
            click_button("close")  # Закрытие окна заказа

        pyperclip.copy(item_name)
        click_button("clear")
        click_button("search")

        if pyautogui.pixelMatchesColor(481, 196, (239, 229, 218), tolerance=10):

            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
        else:
            print("Ошибка: строка поиска не активна.")
            return

        if pyautogui.pixelMatchesColor(481, 196, (239, 229, 218), tolerance=10):
            print("Ошибка: название предмета не вставлено.")
            return

        click_button("order")

    except Exception as e:
        print(f"Ошибка в функции search_and_order_item: {e}")


def suppress_gpu_warning():
    sys.stderr = open(os.devnull, 'w')
    reader = easyocr.Reader(['en'])
    sys.stderr = sys.__stderr__
    return reader


reader = suppress_gpu_warning()


def enlarge_image(image_np, scale=2.0):
    height, width = image_np.shape[:2]
    new_height, new_width = int(height * scale), int(width * scale)
    return cv2.resize(image_np, (new_width, new_height), interpolation=cv2.INTER_CUBIC)


def read_price_with_fallback(x1, y1, x2, y2, fallback_bbox=None):
    main_bbox = (x1, y1, x2, y2)
    main_image = enlarge_image(cv2.cvtColor(np.array(ImageGrab.grab(main_bbox)), cv2.COLOR_RGB2BGR))
    main_result = ''.join(filter(str.isdigit, ''.join(reader.readtext(main_image, detail=0))))

    if main_result:
        return int(main_result)

    if fallback_bbox:
        fallback_image = enlarge_image(cv2.cvtColor(np.array(ImageGrab.grab(fallback_bbox)), cv2.COLOR_RGB2BGR))
        fallback_result = ''.join(filter(str.isdigit, ''.join(reader.readtext(fallback_image, detail=0))))
        if fallback_result:
            return int(fallback_result)

    return 0


def process_items(city, min_profit_percentage, min_margin, limit_buying, current_balance):
    """
    Обробляє предмети з параметрами.
    :param city: Місто, яке обрано в GUI.
    :param min_profit_percentage: Мінімальний відсоток прибутку.
    :param min_margin: Мінімальний margin.
    :param limit_buying: Ліміт на кількість покупок.
    :param current_balance: Поточний баланс користувача.
    """
    try:
        total_sum = 0
        print(f"Обробка предметів для міста: {city}")
        print(f"Мінімальний прибуток: {min_profit_percentage}%")
        print(f"Мінімальний margin: {min_margin}")
        print(f"Ліміт покупок: {limit_buying}")
        print(f"Баланс: {current_balance}")

        # Логіка залишиться такою ж, але всі параметри вже отримані з GUI
        # Далі йде існуюча логіка обробки ордерів...

    except Exception as e:
        print(f"Помилка в process_items: {e}")

        current_balance = 450000000
        limit_buying = 0.2
        limit_buying = 0.2 * 100
        total_sum = 0
        min_profit_percentage = 30
        min_margin = 30000

        # Получение данных из Google Sheets
        item_names, enchants, qualities, raw_volumes, adjusted_volumes = get_item_enchant_quality_from_sheets()
        combined_items = combine_item_and_enchant(item_names, enchants)

        # Подключение к Google Sheets
        credentials_path = 'C:/Users/Urusalim/Documents/Google/marginalbion-16265092a9d2.json'
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        client = gspread.authorize(credentials)

        error_sheet = client.open_by_url(
            'https://docs.google.com/spreadsheets/d/11lMYkckati6qLuyWDjIvNzt3QTdOatAe6Gb72ujE_6Q'
        ).worksheet("Error")

        sheet_data = client.open_by_url(
            'https://docs.google.com/spreadsheets/d/1JGtr5opjzX7B1VKGXgfuubgy51RjGNo3cYX86i0T1w8').worksheet(
            "ResultTF")

        # Очищення старих даних
        print("Очистка старих даних з аркуша...")
        sheet_data.clear()

        # Зберігаємо ордери локально
        batch_data = []

        # Обработка каждого элемента
        for item_name, quality, raw_volume, adjusted_volume in zip(combined_items, qualities, raw_volumes,
                                                                   adjusted_volumes):
            print(
                f"Обрабатываю: {item_name}, Качество: {quality}, Объем: {raw_volume}")

            if not quality:
                print(f"Пропуск {item_name}: качество не указано.")
                continue

            # Поиск и выбор предмета
            search_and_order_item(item_name)
            handle_quality(quality)

            time.sleep(0.1)
            if pyautogui.pixelMatchesColor(1395, 320, (183, 154, 127)):
                print(f"Ордер: {item_name}, не перебит")
                continue

            time.sleep(random.uniform(0.1, 0.4))

            if pyautogui.pixelMatchesColor(1033, 327, (193, 161, 130)):
                current_sell_order = read_price_with_fallback(1004, 311, 1072, 331)
                current_buy_order = read_price_with_fallback(1298, 308, 1357, 334)

            if pyautogui.pixelMatchesColor(1014, 327, (189, 158, 127)) and pyautogui.pixelMatchesColor(1047, 327,
                                                                                                       (187, 156, 125)):

                current_sell_order = read_price_with_fallback(1004, 312, 1084, 331)
                current_buy_order = read_price_with_fallback(1298, 309, 1380, 331)
            else:

                current_sell_order = read_price_with_fallback(1004, 312, 1084, 331)
                current_buy_order = read_price_with_fallback(1298, 309, 1380, 331)

            if not current_sell_order or not current_buy_order:
                continue

            current_margin = current_sell_order - current_buy_order

            base_item = item_name.rsplit(" ", 1)[0]
            enchant = item_name.rsplit(" ", 1)[1] if " " in item_name else ""

            # Отримуємо поточний час
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            batch_data.append([
                base_item,
                enchant,
                quality,
                float(raw_volume),
                float(current_sell_order),
                float(current_buy_order),
                float(current_margin),
                current_time
            ])

            print(
                f"Поточні значення: {item_name}, Якість: {quality}, Продаж: {current_sell_order}, Покупка: {current_buy_order}, "
                f"Обсяг: {raw_volume}, Профіт: {current_margin}, Час: {current_time}")

            if not current_buy_order:
                print(f"Цена для {item_name} ({quality}) не распознана. Записываю в ошибочные элементы.")
                error_sheet.append_row([item_name, quality, "PRICE_NOT_FOUND"])
                click_button("close")
                continue

            profit_percent = (current_margin / current_buy_order) * 100

            # Расчёт профита в процентах
            if current_buy_order > 0:
                profit_percentage = (current_margin / current_buy_order) * 100
            else:
                profit_percentage = 0  # Если цена некорректна, профит равен нулю

            profit_percent = int(profit_percent)

            # Проверка на минимальный профит
            if profit_percentage < min_profit_percentage or current_margin < min_margin:
                print(f"Профит {profit_percentage:.2f}% ниже порога {min_profit_percentage}%. Пропуск...")
                click_button("close")
                continue
            print(f"Ордер актуален: {profit_percent} > {min_profit_percentage}")

            if check_pixels(pixels_to_check):

                click_button("plus")
                # Установка объема заказа
                max_attempts = 3
                success = False

                for attempt in range(max_attempts):
                    print(f"Попытка установки объема: {attempt + 1}/{max_attempts}")

                    set_volume(adjusted_volume)

                    # Проверка стоимости заказа
                    calculated_order_value = read_price_with_fallback(455, 745, 557, 776,
                                                                      fallback_bbox=(455, 745, 515, 771))
                    if not calculated_order_value:
                        print("Ошибка определения стоимости заказа. Повторяю установку объема.")
                        continue

                    if calculated_order_value > (limit_buying * current_balance):
                        print(
                            f"Стоимость {calculated_order_value} превышает 20% баланса ({0.2 * current_balance}). "
                            f"Уменьшаю объем..."
                        )
                        volume = max(1, volume // 2)
                        continue

                    # Обновление суммы и проверка лимита
                    total_sum += calculated_order_value
                    print(f"Заказ подтвержден ({calculated_order_value}). Общая сумма: {total_sum}")

                    click_button("buy_orders")
                    time.sleep(0.2)

                    if pyautogui.pixelMatchesColor(805, 549, (242, 184, 59), tolerance=10):
                        click_button("buy_order_confirm_T")

                    if pyautogui.pixelMatchesColor(804, 571, (170, 134, 56), tolerance=10):
                        click_button("buy_order_confirm_T_L")

                    if pyautogui.pixelMatchesColor(804, 590, (230, 176, 58), tolerance=10):
                        click_button("buy_order_confirm_F")

                    if pyautogui.pixelMatchesColor(805, 616, (242, 184, 59), tolerance=10):
                        click_button("buy_order_confirm_F_L")

                    if total_sum >= current_balance:
                        print("Лимит покупок достигнут. Завершаю выполнение.")
                        return

                    if pyautogui.pixelMatchesColor(905, 240, (233, 185, 0), tolerance=10):
                        click_button("close")  # Закрытие окна заказа

                    break
            else:
                print(f'{item_name}: нестабильная ликвидность пропуск...')
                continue

        # Відправлення даних у Google Sheets після обробки всіх предметів
        if batch_data:
            sheet_data.append_rows(batch_data)
            print(f"Дані про {len(batch_data)} елементів успішно додано до Google Sheets.")


    except Exception as e:
        if pyautogui.pixelMatchesColor(905, 240, (233, 185, 0), tolerance=10):
            click_button("close")  # Закрытие окна заказа
        print(f"Общая ошибка в процессе: {e}")


def set_volume(volume):
    """Устанавливает объем заказа."""
    click_button("count")
    pyautogui.keyDown('backspace')
    pyautogui.keyUp('backspace')
    time.sleep(0.1)

    pyperclip.copy(str(volume))

    pyautogui.hotkey('ctrl', 'v')

    pyautogui.keyDown('enter')
    pyautogui.keyUp('enter')
    print(f"Объем установлен: {volume}")


if __name__ == "__main__":
    activate_window("Albion Online Client")
    capture_window("Albion Online Client")
    process_items()

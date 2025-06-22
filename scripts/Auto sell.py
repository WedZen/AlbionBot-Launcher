import tkinter as tk
from tkinter import messagebox
import random
import pyautogui
import time


# Функция активации окна
def activate_window(window_title):
    try:
        import pygetwindow as gw
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
        print(f"Ошибка при активации окна: {str(e)}")
        return None


# Координаты кнопок на экране
button_coordinate = {
    "cancel_1": ((1337, 351), (1368, 379)),
    "cancel_2": ((1338, 412), (1369, 446)),
    "cancel_3": ((1338, 479), (1365, 511)),

    "cancel_sell1": ((1335, 713), (1370, 747)),
    "cancel_sell2": ((1338, 780), (1369, 810)),
    "cancel_sell3": ((1338, 852), (1370, 870)),

    "sell": ((1253, 382), (1362, 406)),
    "minus": ((434, 626), (468, 656)),
    "plus"
    "sell_order": ((778, 737), (901, 777)),

    "messange_1": ((764, 737), (1104, 782)),
    "messange_2": ((768, 347), (1096, 391)),
    "delete": ((874, 881), (985, 911))
}


def click_button(name, button='left', duration=1):
    (x1, y1), (x2, y2) = button_coordinate[name]
    final_x = random.uniform(x1, x2)
    final_y = random.uniform(y1, y2)
    final_duration = random.uniform(0.1, 0.1)

    pyautogui.moveTo(final_x, final_y, final_duration)
    pyautogui.click(button=button)


# Функция - отмена ордеров
def cancel_buy_orders():
    window_title = "Albion Online Client"
    window = activate_window(window_title)
    if window:
        while True:
            click_button("cancel_1")
            click_button("cancel_2")
            click_button("cancel_3")

            pyautogui.scroll(100)

            if pyautogui.pixelMatchesColor(823, 444, (92, 65, 45), tolerance=10):
                break
            continue


def cancel_sell_orders():
    window_title = "Albion Online Client"
    window = activate_window(window_title)
    if window:
        while True:
            click_button("cancel_sell1")
            click_button("cancel_sell2")
            click_button("cancel_sell3")

            pyautogui.scroll(100)

            if pyautogui.pixelMatchesColor(877, 818, (78, 53, 39), tolerance=10):
                break
            continue


# Функция - Автопродажа предметов
def auto_sell():
    window_title = "Albion Online Client"
    window = activate_window(window_title)
    if window:
        while True:

            click_button("sell")
            click_button("minus")
            click_button("sell_order")

            if pyautogui.pixelMatchesColor(910, 391, (139, 106, 81), tolerance=10):
                break
            continue


# Функция - Автоудаление сообщений
def auto_close_messages():
    window_title = "Albion Online Client"
    window = activate_window(window_title)
    if window:
        while True:
            click_button("messange_2")

            click_button("delete")

            if pyautogui.pixelMatchesColor(903, 363, (141, 108, 83), tolerance=10):
                break
            continue


# Создание GUI
def create_interface():
    root = tk.Tk()
    root.title("Интерфейс управления")

    # Кнопка для отмены ордеров
    cancel_buy_orders_button = tk.Button(
        root, text="Отмена ордеров на покупку", command=cancel_buy_orders, width=25, height=2, bg="red", fg="white"
    )
    cancel_buy_orders_button.pack(pady=10)

    # Кнопка для автопродажи
    auto_sell_button = tk.Button(
        root, text="Автопродажа предметов", command=auto_sell, width=25, height=2, bg="blue", fg="white"
    )
    auto_sell_button.pack(pady=10)

    # Кнопка для автоудаления сообщений
    auto_close_messages_button = tk.Button(
        root, text="Автоудаление сообщений", command=auto_close_messages, width=25, height=2, bg="green", fg="white"
    )
    auto_close_messages_button.pack(pady=10)

    auto_cancel_sell_messages_button = tk.Button(
        root, text="Отмена ордеров на продажу", command=cancel_sell_orders, width=25, height=2, bg="pink", fg="white"
    )
    auto_cancel_sell_messages_button.pack(pady=10)

    root.mainloop()


# Запуск интерфейса
if __name__ == "__main__":
    create_interface()

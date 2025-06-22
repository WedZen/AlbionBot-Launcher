import sys

print("PYTHONPATH:", sys.path)

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIntValidator, QDoubleValidator
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout,
    QPushButton, QLabel, QComboBox, QTextEdit, QMenuBar, QAction, QLineEdit
)
from scripts import auto_put_order
from scripts.run_main_script import MainScriptRunner  # Ваш імпорт run_main_script для потоку
import os
import sys

# Додаємо кореневу папку до PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.append(os.path.abspath(os.path.dirname(__file__)))  # Якщо запуск здійснюється із scripts

# Спрощений словник
city_to_sheet_mapping = {
    "Lumhurst": "ResultLH",
    "Bridgewatch": "ResultBW",
    "Martlock": "ResultML",
    "Thetford": "ResultTF",
    "FortSterling": "ResultFS",
    "Briceleon": "ResultBS"
}


class AutoPutOrderRunner(QThread):
    log_signal = pyqtSignal(str)

    def __init__(self, city, min_profit_percentage, min_margin, limit_buying, current_balance):
        super().__init__()
        self.city = city
        self.min_profit_percentage = min_profit_percentage
        self.min_margin = min_margin
        self.limit_buying = limit_buying
        self.current_balance = current_balance

    def run(self):
        try:
            self.log_signal.emit(f"Запуск auto_put_order для міста {self.city}...")
            auto_put_order.process_items(
                city=self.city,
                min_profit_percentage=self.min_profit_percentage,
                min_margin=self.min_margin,
                limit_buying=self.limit_buying,
                current_balance=self.current_balance,
            )
            self.log_signal.emit("Процес завершено успішно.")
        except Exception as e:
            self.log_signal.emit(f"Помилка: {e}")


import json
import os


class AlbionBotGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Albion Bot Manager")
        self.setGeometry(200, 100, 800, 600)

        self.init_menu()
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Додаємо обидві вкладки
        self.tabs.addTab(self.create_item_check_tab(), "Перевірка предметів")
        self.tabs.addTab(self.create_auto_put_order_tab(), "Загрузка ордерів")

        # Завантаження налаштувань
        self.load_settings()

    def closeEvent(self, event):
        """Зберігає налаштування перед закриттям програми."""
        self.save_settings()
        event.accept()

    def save_settings(self):

        """Зберігає значення полів у файл (тільки якщо всі дані валідні)."""
        if not self.city_selector.currentText():
            print("Помилка: Місто для вкладки 'Перевірка предметів' не вибране!")
            return

        if not self.city_selector_put_order.currentText():
            print("Помилка: Місто для вкладки 'Загрузка ордерів' не вибране!")
            return

        """Зберігає значення полів у файл."""
        settings = {
            "selected_city_main": self.city_selector.currentText(),
            "selected_city": self.city_selector_put_order.currentText(),
            "min_profit_input": self.min_profit_input.text(),
            "min_margin_input": self.min_margin_input.text(),
            "limit_buying_input": self.limit_buying_input.text(),
            "current_balance_input": self.current_balance_input.text(),
        }
        with open("settings.json", "w") as file:
            json.dump(settings, file, indent=4)

    def load_settings(self):
        if os.path.exists("settings.json"):
            try:
                with open("settings.json", "r") as file:
                    settings = json.load(file)
                    # Встановлюємо значення
                    self.city_selector.setCurrentText(settings.get("selected_city_main", ""))
                    self.city_selector_put_order.setCurrentText(settings.get("selected_city", ""))
                    self.min_profit_input.setText(settings.get("min_profit_input", ""))
                    self.min_margin_input.setText(settings.get("min_margin_input", ""))
                    self.limit_buying_input.setText(settings.get("limit_buying_input", ""))
                    self.current_balance_input.setText(settings.get("current_balance_input", ""))
            except (json.JSONDecodeError, KeyError):
                # Виведення помилки прямо до логів
                self.item_check_logs.append("Помилка завантаження налаштувань!")
                self.auto_put_order_logs.append("Помилка у файлі налаштувань! Використовується за замовчуванням.")
                self.save_settings()  # Зберегти новий файл після помилки

    def init_menu(self):
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)

        settings_menu = menu_bar.addMenu("Налаштування")
        full_setup_action = QAction("Повна настройка", self)
        full_setup_action.triggered.connect(self.full_setup)
        settings_menu.addAction(full_setup_action)

    def full_setup(self):
        print("Функція 'Повна настройка' ще в процесі створення.")

    def create_int_validator(self, min_value, max_value):
        """Створюємо валідатор для цілих чисел."""
        return QIntValidator(min_value, max_value, self)

    def create_double_validator(self, min_value, max_value):
        """Створюємо валідатор для дробових чисел."""
        validator = QDoubleValidator(min_value, max_value, 2, self)
        validator.setNotation(QDoubleValidator.StandardNotation)
        return validator

    def create_item_check_tab(self):
        """Створюємо вкладку для перевірки предметів."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Вибір міста
        layout.addWidget(QLabel("Оберіть місто:"))
        self.city_selector = QComboBox()
        self.city_selector.addItems(city_to_sheet_mapping.keys())
        layout.addWidget(self.city_selector)

        # Кнопка запуску
        self.start_check_button = QPushButton("Запустити перевірку")
        self.start_check_button.clicked.connect(self.start_item_check)
        layout.addWidget(self.start_check_button)

        # Поле для відображення логів
        layout.addWidget(QLabel("Логи:"))
        self.item_check_logs = QTextEdit()
        self.item_check_logs.setReadOnly(True)
        font = QFont("Arial", 10)
        self.item_check_logs.setFont(font)
        layout.addWidget(self.item_check_logs)

        tab.setLayout(layout)
        return tab

    def start_item_check(self):
        """Запуск перевірки предметів."""
        selected_city = self.city_selector.currentText()
        self.item_check_logs.append(f"Запуск перевірки для міста: {selected_city}...")

        # Створюємо і запускаємо потік для main.py
        self.main_script_thread = MainScriptRunner(selected_city)
        self.main_script_thread.log_signal.connect(self.update_item_check_logs)
        self.main_script_thread.start()

    def update_item_check_logs(self, log):
        """Оновлення поля логів із потоку main.py"""
        self.item_check_logs.append(log)

    def create_auto_put_order_tab(self):
        """Створюємо вкладку для автоматичного завантаження ордерів."""
        tab = QWidget()
        layout = QVBoxLayout()

        # Вибір міста
        layout.addWidget(QLabel("Оберіть місто:"))
        self.city_selector_put_order = QComboBox()
        self.city_selector_put_order.addItems(city_to_sheet_mapping.keys())
        layout.addWidget(self.city_selector_put_order)

        # Поле для вводу мінімального прибутку
        layout.addWidget(QLabel("Мінімальний відсоток прибутку (%):"))
        self.min_profit_input = QLineEdit(self)
        self.min_profit_input.setPlaceholderText("Наприклад: 20")
        self.min_profit_input.setValidator(self.create_double_validator(0, 100))  # Тільки дробові значення
        layout.addWidget(self.min_profit_input)

        # Поле для вводу мінімальної різниці
        layout.addWidget(QLabel("Мінімальна різниця (Margin):"))
        self.min_margin_input = QLineEdit(self)
        self.min_margin_input.setPlaceholderText("Наприклад: 50000")
        self.min_margin_input.setValidator(self.create_int_validator(0, 1_000_000))  # Тільки цілі значення
        layout.addWidget(self.min_margin_input)

        # Поле для вводу ліміту покупки
        layout.addWidget(QLabel("Максимальный % на один предмет:"))
        self.limit_buying_input = QLineEdit(self)
        self.limit_buying_input.setPlaceholderText("Наприклад: 20")
        self.limit_buying_input.setValidator(self.create_int_validator(1, 100_000))  # Тільки цілі значення
        layout.addWidget(self.limit_buying_input)

        # Поле для вводу поточного балансу
        layout.addWidget(QLabel("Ваш поточний баланс:"))
        self.current_balance_input = QLineEdit(self)
        self.current_balance_input.setPlaceholderText("Наприклад: 500000")
        # Використовуємо QDoubleValidator, оскільки QIntValidator має обмеження
        self.current_balance_input.setValidator(self.create_double_validator(1, 10_000_000_000))
        layout.addWidget(self.current_balance_input)

        # Кнопка запуску
        self.start_put_order_button = QPushButton("Запустити завантаження ордерів")
        self.start_put_order_button.clicked.connect(self.start_put_order)
        layout.addWidget(self.start_put_order_button)

        # Поле для відображення логів
        layout.addWidget(QLabel("Логи:"))
        self.auto_put_order_logs = QTextEdit()
        self.auto_put_order_logs.setReadOnly(True)
        font = QFont("Arial", 10)
        self.auto_put_order_logs.setFont(font)
        layout.addWidget(self.auto_put_order_logs)

        tab.setLayout(layout)
        return tab

    def start_put_order(self):
        """Обробник виклику auto_put_order із параметрами."""
        try:
            selected_city = self.city_selector_put_order.currentText()

            # Зчитування значень
            min_profit_percentage = float(self.min_profit_input.text())
            min_margin = int(self.min_margin_input.text())
            limit_buying = int(self.limit_buying_input.text())
            current_balance = int(self.current_balance_input.text())

            self.auto_put_order_logs.append(f"Місто: {selected_city}")
            self.auto_put_order_logs.append(f"Мінімальний прибуток: {min_profit_percentage}%")
            self.auto_put_order_logs.append(f"Мінімальна різниця: {min_margin}")
            self.auto_put_order_logs.append(f"Ліміт покупки: {limit_buying}")
            self.auto_put_order_logs.append(f"Баланс: {current_balance:,}")

            # Запуск потоку
            self.auto_put_order_thread = AutoPutOrderRunner(
                city=selected_city,
                min_profit_percentage=min_profit_percentage,
                min_margin=min_margin,
                limit_buying=limit_buying,
                current_balance=current_balance,
            )
            self.auto_put_order_thread.log_signal.connect(self.update_put_order_logs)
            self.auto_put_order_thread.start()

        except Exception as e:
            self.auto_put_order_logs.append(f"Виникла помилка: {e}")

    def update_put_order_logs(self, log):
        """Оновлення логів для auto_put_order."""
        self.auto_put_order_logs.append(log)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AlbionBotGUI()
    window.show()
    sys.exit(app.exec_())

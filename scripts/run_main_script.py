import os
import subprocess

from PyQt5.QtCore import QThread, pyqtSignal


class MainScriptRunner(QThread):
    log_signal = pyqtSignal(str)  # Сигнал для передачі логів у GUI

    def __init__(self, city):
        super().__init__()
        self.city = city

    def run(self):
        """Запуск скрипта main.py з передачею параметра міста."""
        try:
            script_path = os.path.join(
                "C:\\Users\\Urusalim\\PycharmProjects\\AlbionBotMargin\\scripts",
                "main.py"
            )

            process = subprocess.Popen(
                ["python", script_path, self.city],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8"  # Застосування UTF-8
            )

            for line in process.stdout:
                self.log_signal.emit(line.strip())  # Передаємо в логи GUI

            # Обробка помилки, якщо є
            process.wait()
            if process.returncode != 0:
                error_message = process.stderr.read()
                self.log_signal.emit(f"Помилка: {error_message}")

        except Exception as e:
            self.log_signal.emit(f"Сталась помилка при виконанні скрипта: {e}")

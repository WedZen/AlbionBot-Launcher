import sys
import os

# Перенаправлення помилок у файл
log_file = os.path.join(os.path.dirname(__file__), "error_log.txt")
sys.stdout = open(log_file, "w")
sys.stderr = open(log_file, "w")

# Ваш основний код далі
from scripts.Interface import AlbionBotGUI
from PyQt5.QtWidgets import QApplication

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AlbionBotGUI()
    window.show()
    sys.exit(app.exec_())

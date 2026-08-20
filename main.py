"""
Точка входа в приложение (PyQt6 версия).
"""

import sys
from PyQt6.QtWidgets import QApplication
from app import App


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Учебная нагрузка преподавателя")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("ОИВТ - филиал СГУВТ")
    
    window = App()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
import sys

from PyQt6.QtWidgets import QApplication

from companion_gui import TruancyWindow

import logging

logging.basicConfig(level=logging.INFO)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = TruancyWindow()
    window.show()

    sys.exit(app.exec())


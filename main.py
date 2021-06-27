import sys
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from app import MyApp

# Excel Automation
if __name__ == "__main__":
    # Our Application
    app = QApplication(sys.argv)
    # Main window of our Application
    window = MyApp("Excel Automatação", 1024, 768, "Excel-automation-icon.ico")
    # Show our window
    window.show()
    sys.exit(app.exec())

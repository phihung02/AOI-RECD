import sys
from PyQt6.QtWidgets import QApplication
from src.main_window import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = MainWindow() 
    viewer.show()
    sys.exit(app.exec())
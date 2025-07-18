from gui.main_window import MainWindow
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("app.ico"))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 
from Classes.MainWindows import MainWindow
from PySide6.QtWidgets import QApplication
import sys

if __name__ == '__main__':

    import sys

    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        print('running in a PyInstaller bundle')
        print(sys._MEIPASS)
    else:
        print('running in a normal Python process')

    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
from PySide6.QtWidgets import QMainWindow, QPushButton, QMessageBox, QGridLayout, QVBoxLayout, QWidget
from PySide6.QtGui import QColor, QPalette


class Color(QWidget):
    def __init__(self, color):
        super().__init__()
        self.setAutoFillBackground(True)

        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(color))
        self.setPalette(palette)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("My App")
        self.left = 100
        self.top = 100
        self.width = 320
        self.height = 240
        self.setGeometry(self.left, self.top, self.width, self.height)
        mainlayout = QVBoxLayout()

        btn = QPushButton("red")
        btn.pressed.connect(self.button_clicked)
        mainlayout.addWidget(btn)

        mainlayout.addWidget(Color("red"))
        mainlayout.addWidget(Color("green"))
        mainlayout.addWidget(Color("blue"))
        mainlayout.addWidget(Color("purple"))

        widget = QWidget()
        widget.setLayout(mainlayout)
        self.setCentralWidget(widget)

    # tag::button_clicked[]
    def button_clicked(self, s):

        button = QMessageBox.critical(
            self,
            "Oh dear!",
            "Something went very wrong.",
            buttons=QMessageBox.Discard | QMessageBox.NoToAll | QMessageBox.Ignore,
            defaultButton=QMessageBox.Discard,
        )

        if button == QMessageBox.Discard:
            print("Discard!")
        elif button == QMessageBox.NoToAll:
            print("No to all!")
        else:
            print("Ignore!")

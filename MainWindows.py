from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QFrame, QSplitter
from PySide6.QtWidgets import QApplication
import sys
import json

from Classes.HIL_Function_Widget import HIL_Function_Widget
from Classes.TC_Highlight_Widget import TC_Highlight_Widget
from Classes.TC_Substitution_Widget import TC_Substitution_Widget


class Separation_Line(QFrame):
    def __init__(self):
        super().__init__()
        self.setFrameShape(QFrame.HLine)
        self.setMinimumHeight(3)
        self.setFrameShadow(QFrame.Sunken)
        self.setStyleSheet('background-color: rgb(255,0,0)')


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ferrari TC tools")
        self.left = 100
        self.top = 100
        self.width = 320
        self.height = 240
        self.setGeometry(self.left, self.top, self.width, self.height)

        try:
            with open('pathFile.json', 'r') as json_file:
                json_file_no_comment = ''.join(line for line in json_file if not line.startswith('#'))
                json_data = json.loads(json_file_no_comment)

        except FileNotFoundError:
            print('File pathFile.json does not exist')
            sys.exit()

        print(json_data)

        # definisci il widget principale
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        widget_hil_function = HIL_Function_Widget(json_data)
        layout_tc_highlight = TC_Highlight_Widget(json_data)
        widget_tc_substitution = TC_Substitution_Widget(json_data)

        main_layout.addWidget(widget_hil_function)
        main_layout.addWidget(Separation_Line())#QSplitter(Qt.Horizontal))
        main_layout.addWidget(layout_tc_highlight)
        main_layout.addWidget(Separation_Line())
        main_layout.addWidget(widget_tc_substitution)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()

    main_window.show()
    sys.exit(app.exec())

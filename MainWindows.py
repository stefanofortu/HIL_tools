from PySide6.QtWidgets import QMainWindow, QPushButton, QMessageBox, QGridLayout, QVBoxLayout, QWidget, QFileDialog, \
    QHBoxLayout
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication
import sys
import json

from Classes.HIL_Function_Handler import HIL_Functions_Handler
from Classes.TC_HighLight_Handler import TC_HighLight_Handler


class TC_Highlight_Widget(QWidget):
    def __init__(self, json_data):
        super().__init__()
        widget_main_layout = QHBoxLayout()

        btn_exec_tc_highlight = QPushButton("exec TC Highlight")
        btn_exec_tc_highlight.pressed.connect(self.tc_highlight_exec_conversion)
        widget_main_layout.addWidget(btn_exec_tc_highlight)

        btn_save_tc_highlight = QPushButton("no sense")
        btn_save_tc_highlight.pressed.connect(self.button_clicked)
        widget_main_layout.addWidget(btn_save_tc_highlight)

        self.setLayout(widget_main_layout)

        input_TC_file_path, input_TC_file_sheet_name, output_TC_file_path = TC_HighLight_Handler.parse_json_path_file(
            json_data)
        print(input_TC_file_path)
        self.tc_highlight_handler = TC_HighLight_Handler(tc_input_file_name=input_TC_file_path,
                                                         tc_sheet_name=input_TC_file_sheet_name,
                                                         tc_output_file_name=output_TC_file_path)
        self.tc_highlight_handler.create_output_file()

    def tc_highlight_exec_conversion(self):
        self.tc_highlight_handler.convert()

    def button_clicked(self):

        button = QMessageBox.critical(
            self,
            "Oh dear!",
            "Told you not to press this.",
            buttons=QMessageBox.Discard | QMessageBox.NoToAll | QMessageBox.Ignore,
            defaultButton=QMessageBox.Discard,
        )

        if button == QMessageBox.Discard:
            print("Discard!")
        elif button == QMessageBox.NoToAll:
            print("No to all!")
        else:
            print("Ignore!")

class HIL_Function_Widget(QWidget):
    def __init__(self, json_data):
        super().__init__()
        widget_main_layout = QHBoxLayout()

        btn_exec_tc_highlight = QPushButton("Exec substitution")
        btn_exec_tc_highlight.pressed.connect(self.hil_substitution_exec_conversion)
        widget_main_layout.addWidget(btn_exec_tc_highlight)

        btn_save_tc_highlight = QPushButton("no sense")
        btn_save_tc_highlight.pressed.connect(self.button_clicked)
        widget_main_layout.addWidget(btn_save_tc_highlight)

        self.setLayout(widget_main_layout)

        function_verify_filename, function_verify_sheetName, build_filename, source_sheet, run_file_path = \
            HIL_Functions_Handler.parse_json_path_file(json_data)

        self.hil_functions_handler = HIL_Functions_Handler(function_verify_filename=function_verify_filename,
                                                           function_verify_sheetName=function_verify_sheetName,
                                                           build_filename=build_filename,
                                                           source_sheet=source_sheet,
                                                           run_filename=run_file_path)
        # hil_functions_handler.run()

    def hil_substitution_exec_conversion(self):
        self.hil_functions_handler.run()
    # tag::button_clicked[]
    def button_clicked(self):

        button = QMessageBox.critical(
            self,
            "Oh dear!",
            "Told you not to press this.",
            buttons=QMessageBox.Discard | QMessageBox.NoToAll | QMessageBox.Ignore,
            defaultButton=QMessageBox.Discard,
        )

        if button == QMessageBox.Discard:
            print("Discard!")
        elif button == QMessageBox.NoToAll:
            print("No to all!")
        else:
            print("Ignore!")

    def run(self):
        pass

    # opening SINGLE FILES
    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            print(fileName)

    # opening MULTIPLES FILES
    def openFileNamesDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self, "QFileDialog.getOpenFileNames()", "",
                                                "All Files (*);;Python Files (*.py)", options=options)
        if files:
            print(files)

    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "QFileDialog.getSaveFileName()", "",
                                                  "All Files (*);;Text Files (*.txt)", options=options)
        if fileName:
            print(fileName)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("My App")
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

        main_layout.addWidget(widget_hil_function)
        main_layout.addWidget(layout_tc_highlight)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    # main_window = TC_Highlight_Widget()
    # main_window = HIL_Function_Widget()

    main_window.show()
    sys.exit(app.exec())

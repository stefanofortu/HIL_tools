from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QWidget, QHBoxLayout, QPushButton, QMessageBox, QVBoxLayout, QLabel, QFrame, QGridLayout, \
    QLineEdit, QFileDialog

from Classes.HIL_Function_Handler import HIL_Functions_Handler


class HIL_Function_Widget(QWidget):
    def __init__(self, json_data):
        super().__init__()

        widget_main_layout = QVBoxLayout()
        ############### TITLE LABEL ###############
        title_label = QLabel(self)
        title_label.setFrameStyle(QFrame.Sunken)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setText("HIL Tool Functions ")
        title_label.setStyleSheet('background-color: rgb(228, 0, 43)')
        widget_main_layout.addWidget(title_label)
        ############### INPUT FILE  ###############
        build_file_layout = QGridLayout()
        #
        build_file_description_label = QLabel(self)
        build_file_description_label.setText("Build file :")
        build_file_description_label.setAlignment(Qt.AlignLeft)
        build_file_layout.addWidget(build_file_description_label, 0, 0)
        #
        self.build_file_path_label = QLabel(self)
        self.build_file_path_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.build_file_path_label.setText("-")
        self.build_file_path_label.setAlignment(Qt.AlignLeft)
        build_file_layout.addWidget(self.build_file_path_label, 1, 0, 1, 8)
        #
        btn_build_file_selector = QPushButton("Add ")
        btn_build_file_selector.setIcon(QIcon('icons/folder-icon.jpg'))

        btn_build_file_selector.pressed.connect(self.openBuildFileDialog)
        build_file_layout.addWidget(btn_build_file_selector, 1, 8, 1, 1)
        #
        widget_main_layout.addLayout(build_file_layout)
        ############### INPUT SHEET  ###############
        build_file_sheet_layout = QGridLayout()
        #
        build_file_sheet_label = QLabel(self)
        build_file_sheet_label.setText("Build file sheet :")
        build_file_sheet_label.setAlignment(Qt.AlignLeft)
        build_file_sheet_layout.addWidget(build_file_sheet_label, 0, 0)
        self.build_sheet_line_edit = QLineEdit()
        build_file_sheet_layout.addWidget(self.build_sheet_line_edit, 1, 0, 1, 8)
        #
        widget_main_layout.addLayout(build_file_sheet_layout)
        ############### FUNCTIONS FILE PATH  ###############
        functions_file_layout = QGridLayout()
        #
        functions_file_label = QLabel(self)
        functions_file_label.setText("Function file :")
        functions_file_label.setAlignment(Qt.AlignLeft)
        functions_file_layout.addWidget(functions_file_label, 0, 0)
        #
        self.functions_file_path_label = QLabel(self)
        self.functions_file_path_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.functions_file_path_label.setText("-")
        self.functions_file_path_label.setAlignment(Qt.AlignLeft)
        functions_file_layout.addWidget(self.functions_file_path_label, 1, 0, 1, 8)
        #
        btn_functions_file_selector = QPushButton("Add ")
        btn_functions_file_selector.setIcon(QIcon('icons/folder-icon.jpg'))
        btn_functions_file_selector.pressed.connect(self.openFunctionsFileDialog)
        functions_file_layout.addWidget(btn_functions_file_selector, 1, 8, 1, 1)
        ##
        widget_main_layout.addLayout(functions_file_layout)
        ############### FIND & REPLACE FILE SHEET  ###############
        functions_file_sheet_layout = QGridLayout()
        #
        functions_file_sheet_label = QLabel(self)
        functions_file_sheet_label.setText("Function sheet :")
        functions_file_sheet_label.setAlignment(Qt.AlignLeft)
        functions_file_sheet_layout.addWidget(functions_file_sheet_label, 0, 0)
        self.functions_sheet_line_edit = QLineEdit()
        functions_file_sheet_layout.addWidget(self.functions_sheet_line_edit, 1, 0, 1, 8)
        #
        widget_main_layout.addLayout(functions_file_sheet_layout)
        ############### OUTPUT FILE  ###############
        run_file_layout = QGridLayout()
        #
        run_file_description_label = QLabel(self)
        run_file_description_label.setText("Output file :")

        run_file_description_label.setAlignment(Qt.AlignLeft)
        run_file_layout.addWidget(run_file_description_label, 0, 0)
        #
        self.run_file_path_label = QLabel(self)
        self.run_file_path_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.run_file_path_label.setText("-")
        self.run_file_path_label.setAlignment(Qt.AlignLeft)
        run_file_layout.addWidget(self.run_file_path_label, 1, 0, 1, 8)
        #
        btn_run_file_selector = QPushButton("Add ")
        btn_run_file_selector.setIcon(QIcon('icons/folder-icon.jpg'))
        btn_run_file_selector.pressed.connect(self.saveFileDialog)
        run_file_layout.addWidget(btn_run_file_selector, 1, 8, 1, 1)
        ##
        widget_main_layout.addLayout(run_file_layout)
        ############### START CONVERSION ###############
        exec_row_layout = QHBoxLayout()
        exec_row_layout.addStretch()
        btn_exec_hil_function = QPushButton("Execute HIL Functions substitution")
        btn_exec_hil_function.setIcon(QIcon('icons/execute-icon.jpg'))
        btn_exec_hil_function.pressed.connect(self.hil_substitution_exec_conversion)
        exec_row_layout.addWidget(btn_exec_hil_function)
        exec_row_layout.addStretch()
        widget_main_layout.addLayout(exec_row_layout)
        ###############
        self.setLayout(widget_main_layout)
        ############### GUI END

        function_verify_filename, function_verify_sheetName, build_filename, source_sheet, run_file_path = \
            HIL_Functions_Handler.parse_json_path_file(json_data)

        self.hil_functions_handler = HIL_Functions_Handler()

        self.hil_functions_handler.function_verify_filename = function_verify_filename
        self.functions_file_path_label.setText(function_verify_filename)

        self.hil_functions_handler.function_verify_sheetName = function_verify_sheetName
        self.functions_sheet_line_edit.setText(str(function_verify_sheetName))

        self.hil_functions_handler.build_filename = build_filename
        self.build_file_path_label.setText(build_filename)

        self.hil_functions_handler.source_sheet = source_sheet
        self.build_sheet_line_edit.setText(source_sheet)

        self.hil_functions_handler.run_filename = run_file_path
        self.run_file_path_label.setText(run_file_path)

    def hil_substitution_exec_conversion(self):
        self.hil_functions_handler.source_sheet = self.build_sheet_line_edit.text()
        self.hil_functions_handler.run()

    def openBuildFileDialog(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Select input xlsx file", "",
                                                  "Excel Files (*.xlsx)")  # , options=options)
        if fileName:
            # print(fileName)
            self.hil_functions_handler.build_filename = fileName
            self.build_file_path_label.setText(fileName)

    def openFunctionsFileDialog(self):
        # options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "Select Functions file", "",
                                                  "Excel Files (*.xlsx)")  # , options=options)
        if fileName:
            # print(fileName)
            self.hil_functions_handler.function_verify_filename = fileName
            self.functions_file_path_label.setText(fileName)

    def saveFileDialog(self):
        fileName, _ = QFileDialog.getSaveFileName(self, "Save output file", "",
                                                  "Excel Files (*.xlsx)")  # , options=options)
        if fileName:
            # print(fileName)
            self.hil_functions_handler.run_filename = fileName
            self.run_file_path_label.setText(fileName)

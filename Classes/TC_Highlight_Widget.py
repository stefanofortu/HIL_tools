from PySide6.QtGui import Qt
from PySide6.QtWidgets import QWidget, QHBoxLayout, QPushButton, QMessageBox, QVBoxLayout, QLabel, QFrame, QGridLayout, \
    QLineEdit, QFileDialog

from Classes.TC_HighLight_Handler import TC_HighLight_Handler


class TC_Highlight_Widget(QWidget):
    def __init__(self, json_data):
        super().__init__()
        widget_main_layout = QVBoxLayout()
        #################################
        title_label = QLabel(self)
        title_label.setFrameStyle(QFrame.Sunken)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setText("TC Highlighting")
        title_label.setStyleSheet('background-color: rgb(255,140,0)')

        widget_main_layout.addWidget(title_label)
        #################################
        input_file_layout = QGridLayout()
        #
        input_file_description_label = QLabel(self)
        input_file_description_label.setText("Input file :")
        input_file_description_label.setAlignment(Qt.AlignLeft)
        input_file_layout.addWidget(input_file_description_label, 0, 0)
        #
        self.input_file_path_label = QLabel(self)
        self.input_file_path_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.input_file_path_label.setText("-")
        self.input_file_path_label.setAlignment(Qt.AlignLeft)
        input_file_layout.addWidget(self.input_file_path_label, 1, 0, 1, 8)
        #
        btn_input_file_selector = QPushButton("-|-")
        btn_input_file_selector.pressed.connect(self.openExcelFileDialog)
        input_file_layout.addWidget(btn_input_file_selector, 1, 8, 1, 1)
        ##
        widget_main_layout.addLayout(input_file_layout)
        #################################
        input_file_sheet_label = QGridLayout()
        #
        input_file_sheet_description_label = QLabel(self)
        input_file_sheet_description_label.setText("Input sheet :")
        input_file_sheet_description_label.setAlignment(Qt.AlignLeft)
        input_file_sheet_label.addWidget(input_file_sheet_description_label, 0, 0)
        self.input_sheet_line_edit = QLineEdit()
        input_file_sheet_label.addWidget(self.input_sheet_line_edit, 1, 0, 1, 8)
        #
        widget_main_layout.addLayout(input_file_sheet_label)
        #################################
        output_file_layout = QGridLayout()
        #
        output_file_description_label = QLabel(self)
        output_file_description_label.setText("Output file :")

        output_file_description_label.setAlignment(Qt.AlignLeft)
        output_file_layout.addWidget(output_file_description_label, 0, 0)
        #
        self.output_file_path_label = QLabel(self)
        self.output_file_path_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.output_file_path_label.setText("-")
        self.output_file_path_label.setAlignment(Qt.AlignLeft)
        output_file_layout.addWidget(self.output_file_path_label, 1, 0, 1, 8)
        #
        btn_input_file_selector = QPushButton("-|-")
        btn_input_file_selector.pressed.connect(self.saveFileDialog)
        output_file_layout.addWidget(btn_input_file_selector, 1, 8, 1, 1)
        ##
        widget_main_layout.addLayout(output_file_layout)
        #################################
        exec_row_layout = QHBoxLayout()
        exec_row_layout.addStretch()
        btn_exec_tc_highlight = QPushButton("exec TC Highlight")
        btn_exec_tc_highlight.pressed.connect(self.tc_highlight_exec_conversion)
        exec_row_layout.addWidget(btn_exec_tc_highlight)
        exec_row_layout.addStretch()
        widget_main_layout.addLayout(exec_row_layout)

        self.setLayout(widget_main_layout)

        self.tc_highlight_handler = TC_HighLight_Handler()

        input_TC_file_path, input_TC_file_sheet_name, output_TC_file_path = \
            TC_HighLight_Handler.parse_json_path_file(json_data)

        self.input_file_path_label.setText(input_TC_file_path)
        self.tc_highlight_handler.tc_input_file_name = input_TC_file_path
        self.tc_highlight_handler.tc_sheet_name = input_TC_file_sheet_name

        self.output_file_path_label.setText(output_TC_file_path)
        self.tc_highlight_handler.tc_output_file_name = output_TC_file_path

    def tc_highlight_exec_conversion(self):
        self.tc_highlight_handler.tc_sheet_name = self.input_sheet_line_edit.text()
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

    # opening SINGLE FILES

    def openExcelFileDialog(self):
        # options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "Select xls file", "",
                                                  "Excel Files (*.xls)")  # , options=options)
        if fileName:
            # print(fileName)
            self.tc_highlight_handler.tc_input_file_name = fileName
            self.input_file_path_label.setText(fileName)

    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Save output file", "",
                                                  "Excel Files (*.xls)")  # , options=options)
        if fileName:
            # print(fileName)
            self.tc_highlight_handler.tc_output_file_name = fileName
            self.output_file_path_label.setText(fileName)

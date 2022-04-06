from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox
from Classes.TC_Substitution_Handler import TC_Substitution_Handler


class TC_Substitution_Widget(QWidget):
    def __init__(self, json_data):
        super().__init__()
        widget_main_layout = QVBoxLayout()

        btn_exec_tc_substitution = QPushButton("Exec TC substitution")
        btn_exec_tc_substitution.pressed.connect(self.tc_substitution_exec_conversion)
        widget_main_layout.addWidget(btn_exec_tc_substitution)

        #btn_save_tc_substitution = QPushButton("no sense")
        #btn_save_tc_substitution.pressed.connect(self.button_clicked)
        #widget_main_layout.addWidget(btn_save_tc_substitution)
        widget_main_layout.insertStretch(0)

        input_file_layout = QHBoxLayout()
        btn_exec_tc_substitution = QPushButton("Exec TC substitution")
        btn_exec_tc_substitution.pressed.connect(self.tc_substitution_exec_conversion)
        input_file_layout.addWidget(btn_exec_tc_substitution)

        input_file_layout = QHBoxLayout()

        input_sheet_layout = QHBoxLayout()
        output_file_layout = QHBoxLayout()
        input_file_layout = QHBoxLayout()

        run_layout = QHBoxLayout()
        run_layout.addStretch()
        btn_exec_tc_substitution = QPushButton("Exec TC substitution")
        btn_exec_tc_substitution.pressed.connect(self.tc_substitution_exec_conversion)
        run_layout.addWidget(btn_exec_tc_substitution)

        widget_main_layout.addLayout(input_file_layout)
        widget_main_layout.addLayout(input_sheet_layout)
        widget_main_layout.addLayout(output_file_layout)
        widget_main_layout.addLayout(run_layout)

        self.setLayout(widget_main_layout)

        in_file_path, in_file_sheet, fr_file_path, fr_file_sheet, out_file_path_v = \
            TC_Substitution_Handler.parse_json_path_file(json_data)
        self.tc_substitution_handler = TC_Substitution_Handler(input_file_path=in_file_path,
                                                               input_file_sheet=in_file_sheet,
                                                               find_replace_file_path=fr_file_path,
                                                               find_replace_file_sheet=fr_file_sheet,
                                                               output_file_path=out_file_path_v)

    def tc_substitution_exec_conversion(self):
        self.tc_substitution_handler.exec_CAN_insertion()

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

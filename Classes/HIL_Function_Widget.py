from PySide6.QtWidgets import QWidget, QHBoxLayout, QPushButton, QMessageBox

from Classes.HIL_Function_Handler import HIL_Functions_Handler


class HIL_Function_Widget(QWidget):
    def __init__(self, json_data):
        super().__init__()
        widget_main_layout = QHBoxLayout()

        btn_exec_tc_highlight = QPushButton("Exec function")
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




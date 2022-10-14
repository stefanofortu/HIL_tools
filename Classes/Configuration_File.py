import json
import sys

from PySide6.QtWidgets import QFileDialog, QWidget

from Classes.Configuration_Data import TC_Substitution_Configuration_Data, TC_Highlight_Configuration_Data, \
    HIL_Function_Configuration_Data


class Configuration_File(QWidget):
    def __init__(self):
        super().__init__()
        self.configuration_file_name = None
        self.configuration_file_data = None
        # self.hil_function_file_data = HIL_Function_Configuration_Data()
        # self.tc_highlight_data = TC_Highlight_Configuration_Data()
        # self.tc_substitution_data = TC_Substitution_Configuration_Data()
        self.new()

    def new(self):
        self.configuration_file_data = {
            "root": {
                "HIL_substitution": {},
                "CAN_highlighting": {},
                "find_replace_multiple_row": {}
            }
        }

    def open(self):
        # options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, 'Select configuration file', "", 'JSON file (*.json)')
        if fileName:
            print(fileName)

            try:
                with open(fileName, 'r') as json_file:
                    json_file_no_comment = ''.join(line for line in json_file if not line.startswith('#'))
                    file_dict = json.loads(json_file_no_comment)

            except FileNotFoundError:
                print('File %s does not exist', fileName)
                sys.exit()

            self.configuration_file_name = fileName

            try:
                HIL_substitution_dict = file_dict['root']['HIL_substitution']
            except KeyError:
                print('Field HIL_substitution in file %s does not exist', fileName)
                sys.exit()
            try:
                CAN_highlighting_dict = file_dict['root']['CAN_highlighting']
            except KeyError:
                print('Field CAN_highlighting in file %s does not exist', fileName)
                sys.exit()
            try:
                find_replace_multiple_row_dict = file_dict['root']['find_replace_multiple_row']
            except KeyError:
                print('Field find_replace_multiple_row in file %s does not exist', fileName)
                sys.exit()

            hil_function_file_data = None  # HIL_Function_Configuration_Data(HIL_substitution_dict)
            tc_highlight_data = None  # TC_Highlight_Configuration_Data(CAN_highlighting_dict)
            tc_substitution_data = TC_Substitution_Configuration_Data(find_replace_multiple_row_dict)

            self.configuration_file_data = file_dict

            return hil_function_file_data, tc_highlight_data, tc_substitution_data
        else:
            raise ValueError

    def save(self, hil_function_file_data=None, tc_highlight_data=None, tc_substitution_data=None,
             select_new_file=True):

        self.configuration_file_data['root']['HIL_substitution'] = hil_function_file_data
        self.configuration_file_data['root']['CAN_highlighting'] = tc_highlight_data
        self.configuration_file_data['root']['find_replace_multiple_row'] = tc_substitution_data

        print(self.configuration_file_data)

        # Serializing json
        json_object = json.dumps(self.configuration_file_data, indent=4)

        if not select_new_file:
            fileName = self.configuration_file_name
        else:
            fileName, _ = QFileDialog.getSaveFileName(self, "Save configuration as ", "",
                                                      'JSON file (*.json)')  # , options=options)
        if fileName:
            # Writing to sample.json
            try:
                with open(fileName, "w") as outfile:
                    outfile.write(json_object)
            except IOError:
                print("Error in writing files")

            self.configuration_file_name = fileName

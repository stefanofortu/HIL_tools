import json
import sys


class Configuration_Data:
    def return_json_dict(self):
        dict = {}
        return dict


class HIL_Function_Configuration_Data(Configuration_Data):
    def __init__(self):
        self.function_verify_filename = ""
        self.function_verify_sheetName = ""
        self.build_filename = ""
        self.source_sheet = []
        self.run_file_path = ""


class TC_Highlight_Configuration_Data(Configuration_Data):
    def __init__(self):
        self.input_TC_file_path = ""
        self.input_TC_file_sheet_name = ""
        self.output_TC_file_path = ""


class TC_Substitution_Configuration_Data(Configuration_Data):
    def __init__(self, find_replace_dictionary=None):
        if find_replace_dictionary is None:
            self.input_file_path = ""
            self.input_file_sheet = ""
            self.find_replace_file_path = ""
            self.find_replace_file_sheet = ""
            self.output_file_path = ""
            print("verified that config data is created")
        else:
            try:
                input_file = find_replace_dictionary['input_file']
                self.input_file_path = input_file['path']
                # print('filePath for input file :', input_file_path)
                self.input_file_sheet = input_file['sheet_name']
                # print('sheets in input file :', input_file_sheet)

                find_replace_file = find_replace_dictionary['find_replace_file']
                self.find_replace_file_path = find_replace_file['path']
                # print('filePath for find&Replace file :', find_replace_file_path)
                self.find_replace_file_sheet = find_replace_file['sheet_name']
                # print('sheets in find&Replace file :', find_replace_file_sheet)

                output_file = find_replace_dictionary['output_file']
                self.output_file_path = output_file['path']
                # print('filePath for output file :', output_file_path)
                print("verified that config data is ok")
            except KeyError:
                print('Field in TC_Substitution_Configuration_Data wrongly formatted')
                sys.exit()

    def return_json_dict(self):
        find_replace_dict = {
            "input_file": {
                "path": self.input_file_path,
                "sheet_name": self.input_file_sheet
            },
            "find_replace_file": {
                "path": self.find_replace_file_path,
                "sheet_name": self.find_replace_file_sheet
            },
            "output_file": {
                "path": self.output_file_path
            }
        }

        return find_replace_dict

    def set_in_file_path(self, data):
        print("checks missing")
        self.input_file_path = data

    def set_in_file_sheet(self, data):
        print("checks missing")
        self.input_file_sheet = data

    def set_fr_file_path(self, data):
        print("checks missing")
        self.find_replace_file_path = data

    def set_fr_file_sheet(self, data):
        print("checks missing")
        self.find_replace_file_sheet = data

    def set_out_file_path(self, data):
        print("checks missing")
        self.output_file_path = data

    # def parse_json_dict(self, find_replace_dictionary):

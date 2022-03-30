from Classes.TC_HighLight_Handler import TC_HighLight_Handler
from Classes.HIL_Function_Handler import HIL_Functions_Handler
import sys
import json


def main():
    try:
        with open('pathFile.json', 'r') as json_file:
            json_file_no_comment = ''.join(line for line in json_file if not line.startswith('#'))
            json_data = json.loads(json_file_no_comment)

    except FileNotFoundError:
        print('File pathFile.json does not exist')
        sys.exit()

    print(json_data)

    if json_data['root']['programSelected'] == "HIL_substitution":
        hil_substitution_json = json_data['root']['HIL_substitution']

        filesSubstitution = hil_substitution_json['filesSubstitution']
        function_verify_filename = filesSubstitution['file_path']
        print('filePath for substitution file :', function_verify_filename)
        function_verify_sheetName = filesSubstitution['sheets_name']
        for sheet_name in function_verify_sheetName:
            print('sheets in substitution file :', sheet_name)

        file_TC_Build = hil_substitution_json['file_TC_Build']
        build_filename = file_TC_Build['file_path']
        print('filePath for build file :', build_filename)
        source_sheet = file_TC_Build['sheet_name']
        print('sheets in build file :', source_sheet)

        file_TC_Run = hil_substitution_json['file_TC_Run']
        run_filename = file_TC_Run['file_path']

        hil_functions_handler = HIL_Functions_Handler(function_verify_filename=function_verify_filename,
                                                      function_verify_sheetName=function_verify_sheetName,
                                                      build_filename=build_filename,
                                                      source_sheet=source_sheet,
                                                      run_filename=run_filename)
        hil_functions_handler.run()


    elif json_data['root']['programSelected'] == "CAN_highlighting":
        CAN_highlighting_json = json_data['root']['CAN_highlighting']
        input_TC_file = CAN_highlighting_json['input_file']
        input_TC_file_path = input_TC_file['file_path']
        print('filePath for substitution file :', input_TC_file_path)
        input_TC_file_sheet_name = input_TC_file['sheet_name']

        output_TC_file = CAN_highlighting_json['output_file']
        output_TC_file_name = output_TC_file['file_path']
        print('filePath for build file :', output_TC_file_name)

        # exit()

        tc_highlight_handler = TC_HighLight_Handler(tc_input_file_name=input_TC_file_path,
                                                    tc_sheet_name=input_TC_file_sheet_name,
                                                    tc_output_file_name=output_TC_file_name)
        tc_highlight_handler.create_output_file()
        tc_highlight_handler.convert()

    else:
        print("wrong program selected")
        exit()


if __name__ == "__main__":
    main()

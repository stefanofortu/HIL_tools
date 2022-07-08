from fillers import fillTestNColumn, fillEnableColumn, fillStepIDCounter
from importDictionary import importDictionary
from utils.expressionSubtitution import substituteFunctions, removeTestTypeColumn, findExpressions
from utils.fileImporter import importFunctionFiles, importBuildFile, generateRunFileFromBuildFile, TC_Build_File, \
    Functions_File, TC_Run_File


class HIL_Functions_Handler:
    def __init__(self, function_verify_filename="", function_verify_sheetName="",
                 build_filename="", source_sheet="", run_filename=""):
        self.function_verify_filename = function_verify_filename
        self.function_verify_sheetName = function_verify_sheetName
        self.build_filename = build_filename
        self.source_sheet = source_sheet
        self.run_filename = run_filename
        self.build_file = None
        self.run_file = None

    def import_build_file(self):
        self.build_file = TC_Build_File(self.build_filename, self.source_sheet)

    @staticmethod
    def parse_json_path_file(json_data):
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
        run_file_path = file_TC_Run['file_path']
        return function_verify_filename, function_verify_sheetName, build_filename, source_sheet, run_file_path

    def run(self):

        self.import_build_file()

        print("import build file")

        functionDictionary = {}

        for verifySheet in self.function_verify_sheetName:
            wsVerify = Functions_File(self.function_verify_filename, tc_sheet_name=verifySheet)

            res = importDictionary(function_file=wsVerify, functionDictionary=functionDictionary,
                                   dictionaryType="verify")
            if res == 1:
                print(
                    "MINOR: Error of importDictionary found in " + self.function_verify_filename + ",sheet : " + verifySheet)
                exit()

        for k in functionDictionary:
            print(k)

        print("dizionario acquisito")

        self.build_file.save_copy(self.run_filename)

        self.run_file = TC_Run_File(self.run_filename, self.source_sheet)

        # exit()
        # self.build_file.generateRunFileFromBuildFile(run_filename=self.run_filename)

        # exit()

        # wbRun, wsRun = generateRunFileFromBuildFile(workbookBuild=wbBuild,
        #                                            sheetNameBuild=self.source_sheet,
        #                                            run_filename=self.run_filename)
        print("saved RUN file")

        findExpressions(wsStart=self.build_file.worksheet, substitutionDictionary=functionDictionary)
        print("find expressions done")

        substituteFunctions(wsStart=self.build_file.worksheet, wsEnd=self.run_file.worksheet,
                            substitutionDictionary=functionDictionary, copyStyle=True)
        print("substitution done")


        # disableSequences(worksheet=wsRun)
        # removeTestTypeColumn(worksheet=self.run_file.worksheet)
        fillTestNColumn(worksheet=self.run_file.worksheet)
        fillEnableColumn(worksheet=self.run_file.worksheet)
        fillStepIDCounter(worksheet=self.run_file.worksheet)

        print("other operations")

        # " Salva"
        self.run_file.save()
        #wbRun.save(filename=self.run_filename)

        print("file saved")

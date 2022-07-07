from fillers import fillTestNColumn, fillEnableColumn, fillStepIDCounter
from importDictionary import importDictionary
from utils.expressionSubtitution import substituteFunctions, removeTestTypeColumn, findExpressions
from utils.fileImporter import importFunctionFiles, importBuildFile, generateRunFileFromBuildFile


class HIL_Functions_Handler:
    def __init__(self, function_verify_filename="", function_verify_sheetName="",
                 build_filename="", source_sheet="", run_filename=""):
        self.function_verify_filename = function_verify_filename
        self.function_verify_sheetName = function_verify_sheetName
        self.build_filename = build_filename
        self.source_sheet = source_sheet
        self.run_filename = run_filename

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
        # function_verify_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175_AF_substitutions_2021_11_23.xlsx"
        # function_verify_sheetName = ["actions", "doors", "networks"]

        # function_actions_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175_AF_substitutions_2021_11_23.xlsx"
        # function_actions_sheetName = "actions"

        functionDictionary = {}

        for verifySheet in self.function_verify_sheetName:
            wsVerify = importFunctionFiles(fileName=self.function_verify_filename,
                                           sheetName=verifySheet)

            res = importDictionary(worksheetAction=wsVerify, functionDictionary=functionDictionary,
                                   dictionaryType="verify")
            if res == 1:
                print("MINOR: Error of importDictionary found in " + self.function_verify_filename + ",sheet : " + verifySheet)
                exit()

        for k in functionDictionary:
            print(k)

        print("dizionario acquisito")

        wbBuild, wsBuild = importBuildFile(fileName=self.build_filename,
                                           sheetName=self.source_sheet)
        print("import build file")

        wbRun, wsRun = generateRunFileFromBuildFile(workbookBuild=wbBuild,
                                                    sheetNameBuild=self.source_sheet,
                                                    run_filename=self.run_filename)
        print("saved RUN file")

        findExpressions(wsStart=wsBuild, substitutionDictionary=functionDictionary)
        print("find expressions done")

        substituteFunctions(swStart=wsBuild, wsEnd=wsRun, substitutionDictionary=functionDictionary, copyStyle=True)
        print("substitution done")

        # disableSequences(worksheet=wsRun)
        removeTestTypeColumn(worksheet=wsRun)
        fillTestNColumn(worksheet=wsRun)
        fillEnableColumn(worksheet=wsRun)
        fillStepIDCounter(worksheet=wsRun)

        print("other operations")

        # " Salva"
        wbRun.save(filename=self.run_filename)

        print("file saved")

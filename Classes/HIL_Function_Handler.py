from fillers import fillTestNColumn, fillEnableColumn, fillStepIDCounter
from importDictionary import importDictionaryV2
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

            res = importDictionaryV2(worksheetAction=wsVerify, functionDictionary=functionDictionary,
                                     dictionaryType="verify")
            if res == 1:
                print("MINOR: Error of importDictionary found in " + self.function_verify_filename + ",sheet : " + verifySheet)
                exit()

            # wsVerify2 = importFunctionFiles(fileName=function_verify_filename, sheetName=function_verify_sheetName[1])
            #
            # res = importDictionaryV2(worksheetAction=wsVerify2, functionDictionary=functionDictionary, dictionaryType="verify")
            # if res == 1:
            #     print("MINOR: Error found in " + function_verify_filename + ",sheet : " + function_verify_sheetName[1])
            #     exit()
            #
            # wsVerify3 = importFunctionFiles(fileName=function_verify_filename, sheetName=function_verify_sheetName[2])
            #
            # res = importDictionaryV2(worksheetAction=wsVerify3, functionDictionary=functionDictionary, dictionaryType="verify")
            # if res == 1:
            #     print("MINOR: Error found in " + function_verify_filename + ",sheet : " + function_verify_sheetName[2])
            #     exit()

        for k in functionDictionary:
            print(k)

        # wsAction = importFunctionFiles(fileName=function_actions_filename, sheetName=function_actions_sheetName)

        # res = importDictionaryV2(worksheetAction=wsAction, functionDictionary=functionDictionary, dictionaryType="verify")
        # if res == 1:
        #    print("MINOR: Error found in " + function_actions_filename + ",sheet : " + function_actions_sheetName)
        #    exit()

        print("dizionario acquisito")

        # build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175\\F175_TC_HIL_Integration_Build.xlsx"
        # build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\2-TC_Build\\F175\\TC_Integration\\F175_TC_HIL_Integration_Build.xlsx"
        # source_sheet = "Sleep"
        # source_sheet = "Test case"
        # run_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\3-TC_Run\\F175_TC_HIL_Integration_Run_20211126.xlsx"

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

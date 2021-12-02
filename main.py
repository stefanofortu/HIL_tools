from importDictionary import importDictionaryV2
from utils.fileImporter import importFunctionFiles
import sys
import json

try:
    # Opening JSON file
    f = open('pathFile.json')
except FileNotFoundError:
    print('File pathFile.json does not exist')
    sys.exit()

rootPathFile = json.load(f)
pathFile = rootPathFile['root']
for sheet in pathFile['filesSubstitution']:
    print(sheet['sheetName'])
for sheet in pathFile['file_TC_Build']:
    print(sheet['sheetName'])

# Closing file
f.close()

function_verify_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175_AF_substitutions_2021_11_23.xlsx"
function_verify_sheetName = ["actions", "doors", "networks"]

# function_actions_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175_AF_substitutions_2021_11_23.xlsx"
# function_actions_sheetName = "actions"


functionDictionary = {}

# app = QApplication(sys.argv)
# window = MainWindow()
# window.show()
# app.exec()
# exit()

for verifySheet in function_verify_sheetName:
    wsVerify = importFunctionFiles(fileName=function_verify_filename, sheetName=verifySheet)

    res = importDictionaryV2(worksheetAction=wsVerify, functionDictionary=functionDictionary, dictionaryType="verify")
    if res == 1:
        print("MINOR: Error found in " + function_verify_filename + ",sheet : " + verifySheet)
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
sys.exit(0)

build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175\\F175_TC_HIL_Integration_Build.xlsx"
# build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\2-TC_Build\\F175\\TC_Integration\\F175_TC_HIL_Integration_Build.xlsx"
source_sheet = "Sleep"
# source_sheet = "Test case"
run_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\3-TC_Run\\F175_TC_HIL_Integration_Run_20211126.xlsx"


wbBuild, wsBuild = importBuildFile(fileName=build_filename,
                                   sheetName=source_sheet)

wbRun, wsRun = generateRunFileFromBuildFile(workbookBuild=wbBuild,
                                            sheetNameBuild=source_sheet,
                                            run_filename=run_filename)

substituteFunctions(swStart=wsBuild, wsEnd=wsRun, substitutionDictionary=functionDictionary)

# disableSequences(worksheet=wsRun)
removeTestTypeColumn(worksheet=wsRun)
fillTestNColumn(worksheet=wsRun)
fillEnableColumn(worksheet=wsRun)
fillStepIDCounter(worksheet=wsRun)

# " Salva"
wbRun.save(filename=run_filename)

from fillers import fillTestNColumn, fillEnableColumn, fillStepIDCounter
from importDictionary import importDictionaryV2
from utils.expressionSubtitution import substituteFunctions, removeTestTypeColumn, findExpressions
from utils.fileImporter import importFunctionFiles, importBuildFile, generateRunFileFromBuildFile
import sys
import json

try:
    # Opening JSON file
    f = open('pathFile.json')
except FileNotFoundError:
    print('File pathFile.json does not exist')
    sys.exit()

rootPathFile = json.load(f)
filesSubstitution = rootPathFile['root']['filesSubstitution']
function_verify_filename = filesSubstitution['filepath']
print('filePath for substitution file :', function_verify_filename)
function_verify_sheetName = filesSubstitution['sheetNames']
for sheetNames in function_verify_sheetName:
    print('sheets in substitution file :' , sheetNames)

file_TC_Build = rootPathFile['root']['file_TC_Build']
build_filename = file_TC_Build['filepath']
print('filePath for build file :', build_filename)
source_sheet = file_TC_Build['sheetName']
print('sheets in build file :', source_sheet)

file_TC_Run = rootPathFile['root']['file_TC_Run']
run_filename = file_TC_Run['filepath']
# Closing file
f.close()
#sys.exit()

#function_verify_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175_AF_substitutions_2021_11_23.xlsx"
#function_verify_sheetName = ["actions", "doors", "networks"]

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
        sys.exit()


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

#build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\2-TC_Build\\F175\\F175_TC_HIL_Integration_Build.xlsx"
# build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\2-TC_Build\\F175\\TC_Integration\\F175_TC_HIL_Integration_Build.xlsx"
#source_sheet = "Sleep"
# source_sheet = "Test case"
#run_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\3-TC_Run\\F175_TC_HIL_Integration_Run_20211126.xlsx"


wbBuild, wsBuild = importBuildFile(fileName=build_filename,
                                   sheetName=source_sheet)

print("import build file")

wbRun, wsRun = generateRunFileFromBuildFile(workbookBuild=wbBuild,
                                            sheetNameBuild=source_sheet,
                                            run_filename=run_filename)
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
wbRun.save(filename=run_filename)

print("file saved")

sys.exit(0)
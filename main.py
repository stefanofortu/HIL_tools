from importDictionary import importDictionaryV2
from expressionSubtitution import substituteFunctions
from utils.fileImporter import importFunctionFiles, importBuildFile, generateRunFileFromBuildFile
from stepIDFiller import fillStepIDCounter

function_verify_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\Documentazione\\3.Tools\\F175_AF_verify_2021_06_26.xlsx"
function_verify_sheetName = "verify"

function_actions_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\Documentazione\\3.Tools\\F175_AF_actions_2021_06_26.xlsx"
function_actions_sheetName = "actions"

build_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\2-TC_Build\\F173_TC_HIL_FirstTest_Build.xlsx"
source_sheet = "Test case"
run_filename = "C:\\Users\\Stefano\\Desktop\\WIP\\HIL\\TC_AF\\3-TC_Run\\F173_TC_HIL_FirstTest_Run.xlsx"

functionDictionary = {}

wsVerify = importFunctionFiles(fileName=function_verify_filename, sheetName=function_verify_sheetName)

res = importDictionaryV2(worksheetAction=wsVerify, functionDictionary=functionDictionary, dictionaryType="verify")
if res == 1:
    print("MINOR: Error found in " + function_verify_filename + ",sheet : " + function_verify_sheetName)

wsAction = importFunctionFiles(fileName=function_actions_filename, sheetName=function_actions_sheetName)

res = importDictionaryV2(worksheetAction=wsAction, functionDictionary=functionDictionary, dictionaryType="verify")
if res == 1:
    print("MINOR: Error found in " + function_actions_filename + ",sheet : " + function_actions_sheetName)

print("dizionario acquisito")


wbBuild, wsBuild = importBuildFile(fileName=build_filename,
                          sheetName=source_sheet)

wbRun, wsRun = generateRunFileFromBuildFile(workbookBuild=wbBuild,
                                            sheetNameBuild=source_sheet,
                                            run_filename=run_filename)

substituteFunctions(worksheetStart=wsBuild, worksheetEnd=wsRun, substitutionDictionary=functionDictionary)


fillStepIDCounter(worksheet=wsRun)


# " Salva"
wbRun.save(filename=run_filename)
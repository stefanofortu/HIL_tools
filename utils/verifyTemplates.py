from utils.fileTemplateConfiguration import file_TC_BUILD_Column, file_TC_RUN_Column, file_Tools_Verify_Column
from utils.excelUtils import getColumnLetterFromString, getColumnIndexFromString
from utils.DebugPrint import debugPrint


def checkBuildFile(worksheetBuild, worksheetBuildFileName):
    try:
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['enable_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['testN_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['testID_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['testType_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['stepID_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['stepDescr_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['precondition_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['action_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['expected_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['timeStep_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['sampleTime_header'])
        getColumnLetterFromString(worksheetBuild, file_TC_BUILD_Column['tolerance_header'])

    except ValueError:
        debugPrint("In file ", worksheetBuildFileName)
        exit()


def checkRunFile(worksheetRun, worksheetRunFileName):
    try:
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['enable_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['testN_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['testID_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['stepID_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['stepDescr_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['precondition_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['action_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['expected_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['timeStep_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['sampleTime_header'])
        getColumnLetterFromString(worksheetRun, file_TC_RUN_Column['tolerance_header'])
    except ValueError:
        debugPrint("In file ", worksheetRunFileName)
        exit()


def checkActionFile(worksheetAction, worksheetActionFileName):
    try:
        getColumnIndexFromString(worksheetAction, file_Tools_Verify_Column['fieldType_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['fieldType_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['value_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['stepDescr_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['precondition_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['expected_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['timeStep_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['sampleTime_header'])
        getColumnLetterFromString(worksheetAction, file_Tools_Verify_Column['tolerance_header'])
    except ValueError:
        debugPrint("In file ", worksheetActionFileName)
        exit()

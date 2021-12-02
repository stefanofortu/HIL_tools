from utils.verifyTemplates import checkActionFile, checkBuildFile
from openpyxl import load_workbook


def importFunctionFiles(fileName, sheetName):
    wbFunctions = load_workbook(fileName)
    wsFunctions = wbFunctions[sheetName]
    checkActionFile(worksheetAction=wsFunctions, worksheetActionFileName=fileName)
    ## migliorare la gestione degli errori con python - try and catch

    print(wbFunctions.sheetnames)
    return wsFunctions


def importBuildFile(fileName, sheetName):
    wbBuild = load_workbook(fileName)
    wsBuild = wbBuild[sheetName]
    checkBuildFile(worksheetBuild=wsBuild, worksheetBuildFileName=fileName)
    return wbBuild, wsBuild


def generateRunFileFromBuildFile(workbookBuild, sheetNameBuild, run_filename):
    workbookBuild.save(filename=run_filename)
    wbRun = load_workbook(run_filename)
    wsRun = wbRun[sheetNameBuild]
    return wbRun, wsRun

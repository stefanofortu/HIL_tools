::E' possibile inserire un parametro per il build della versione di interesse
:: %0 is the program name as it was called.
:: %1 is the first command line parameter ... and so on till %9
:: Example : installer.bat 101
@echo off

IF "%1"=="" (
    ECHO WARNING: Specify the version of the executable as first parameter
) ELSE (
	ECHO Building executable file:
	pyinstaller --distpath installer --onefile --clean --workpath installer/build  --name TC_tool_v%1 --specpath installer/build main.py	
)
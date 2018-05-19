REM ***This batch file will automatically run SCANPST on every PST file in the 
REM ***directory specified by PST_FILE_MASK. 

set SCANPST_PATH="C:\Program Files (x86)\Microsoft Office\Office12\SCANPST.EXE"
set PST_FILE_MASK="D:\Users\josh\Documents\My Mail\*.pst"

REM *** CD into in the directory that contains the launched bacth file...
cd %~dp0

del cscanpst.log

for %%i in (%PST_FILE_MASK%) do (

	REM Add an N to the end of the following line of you don't want backup files

	cscanpst.exe %SCANPST_PATH% "%%i"
	if errorlevel 3 goto done
)
:done

@echo Log:
@type cscanpst.log

@pause


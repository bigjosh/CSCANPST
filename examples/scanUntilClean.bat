REM ***This batch file will automatically run SCANPST on specified PST file in the 
REM ***directory specified by PST_FILE_MASK. 
REM ***It will keep running until the file is clean or there is an error.


set SCANPST_PATH="C:\Program Files (x86)\Microsoft Office\Office14\SCANPST.EXE"
set PST_FILE="D:\Users\josh\Documents\My Mail\Mail2010.pst"


REM *** CD into in the directory that contains the launched batch file...

echo CDing into %~dp0

%~d0

cd %~p0

del cscanpst.log

:loop

cscanpst.exe %SCANPST_PATH% %PST_FILE%

if errorlevel 3 goto DONE

if errorlevel 1 goto loop

ECHO *****File scanned clean!

:done

@echo Log:
@type cscanpst.log

@pause


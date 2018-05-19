REM ***This batch file will automatically run SCANPST on every PST file in the 
REM ***directory specified by PST_FILE_MASK. 
REM ***and will keep on scanning each file until it comes up with no errors

set SCANPST_PATH="C:\Program Files (x86)\Microsoft Office\root\Office16\SCANPST.EXE"
set PST_FILE_MASK="D:\Users\josh\Documents\My Mail\*.pst"

REM *** CD into in the directory that contains the launched batch file...
echo CDing into "%~dp0"
pushd "%~dp0"

for %%i in (%PST_FILE_MASK%) do (

                
                :loop
                                echo checking files %SCANPST_PATH% "%%i" 
                                REM Add an N to the end of the following line of you don't want backup files                                
                                cscanpst.exe %SCANPST_PATH% "%%i"
                                echo Finished checking file %SCANPST_PATH% "%%i"
                                if errorlevel 3 goto exit
                                if errorlevel 1 goto loop
                :exit
                                ECHO %SCANPST_PATH% "%%i" File scanned clean!
)
:done

@echo Log:
@type cscanpst.log

popd

@pause


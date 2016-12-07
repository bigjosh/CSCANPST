;
; AutoHotkey Version: 1.x
; Language:       English / working with scanpst.exe portuguese Brazil pt-BR 
; Platform:       Win9x/NT
; Author:         josh levine (http://josh.com/cmdscanpst)
;
; Script Function:
;	Allows you to run the Outlook Repair Tool (SCANPST) from a commmand line batch file
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


FileAppend, Launched on %2%... , scanpstone.log


if 0 < 2
{
	FileAppend, ERROR: Not enough parameters specified , scanpstone.log
	MsgBox CSCANPST requires at least 2 incoming parameters but it only received %0%.
	ExitApp 11

}


;Silly temp variables param1 and param2 are needed becuase ahk can't handle a command line param in an expression!

param1 = %1%


if !FileExist( param1 )
{
	FileAppend, ERROR: SCANPST.EXE not found , scanpstone.log
	MsgBox SCANPST.EXE not found at [%1%].
	ExitApp 12
}


param2 = %2%


if !FileExist( param2 )
{
	FileAppend, ERROR: PST file [%2%] not found , scanpstone.log
	MsgBox No PST file found at [%2%].
	ExitApp 13
}

param3 = %3%

	
if param3=N
{
	FileAppend, (INFO: No backup PST file will be created)...  , scanpstone.log
}



SetTitleMatchMode 2


IfWinExist ahk_exe scanpst.exe
{
	FileAppend, ERROR: Repair Tool is already running , scanpstone.log
	ExitApp 3

}


;Run the SCANPST exe file...
Run %1%

;WinWaitActive Inbox Repair Tool
WinWaitActive ahk_exe scanpst.exe

;Enter the filename to be scanned
Send %2%

Send !I


Loop 
{

	;ifWinExist, Inbox Repair Tool, been canceled
	ifWinExist, ahk_exe scanpst.exe, cancelada
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !F
			
		FileAppend, ERROR: User canceled`n , scanpstone.log
	
		exitapp 4

	}
	


	ifWinExist, ahk_exe scanpst.exe, error prevented access
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !F
			
		FileAppend, ERROR: Could not open file`n , scanpstone.log
		exitapp 5

	}

	; in use by another
	ifWinExist, ahk_exe scanpst.exe, por outro aplicativo
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !F
			
		FileAppend, ERROR: File already in use`n , scanpstone.log
		exitapp 6

	}

	;does not exist
	ifWinExist, ahk_exe scanpst.exe, o existe
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !F
			
		FileAppend, ERROR: File not found`n , scanpstone.log
		exitapp 7

	}

	;does not recognize the file
	ifWinExist, ahk_exe scanpst.exe, o reconhece o arquivo
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !F
			
		FileAppend, ERROR: File type not recognized`n , scanpstone.log
		exitapp 8

	}


	ifWinExist, ahk_exe scanpst.exe, error has occurred
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !F
			
		FileAppend, ERROR: An error has occurred`n , scanpstone.log
		exitapp 9

	}
	
	;is read-only
	ifWinExist, ahk_exe scanpst.exe, somente leitura 
	{

		WinActivate
		WinWaitActive
		Send {ESC}
		Send !F
			
		FileAppend, ERROR: PST file was read only`n , scanpstone.log
		exitapp 10

	}

	;No errors were found
	ifWinExist, ahk_exe scanpst.exe, enhum erro

	{

	
			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose

		FileAppend, No errors found`n , scanpstone.log
	
		exitapp 0
	}
	
	;To repair these errors
	ifWinExist, ahk_exe scanpst.exe, clique em Reparar
	{

		WinActivate
		WinWaitActive

		if param3=N
		   Send !F


		Send !R
		Loop 

		{

			;The backup file 
			ifWinExist, ahk_exe scanpst.exe, O arquivo de backup
			{

				WinActivate
				WinWaitActive


				
				Send !S


				WinWaitClose

			}
		
		
			ifWinExist, ahk_exe scanpst.exe, Reparo concl

			{
				WinActivate
				WinWaitActive
				Send {ENTER}
				WinWaitClose
				FileAppend, File repaired`n , scanpstone.log
				exitapp 2

			}
		



		}			



		

	}	


	ifWinExist, ahk_exe scanpst.exe, incons;Only minor inconsistencies were found

	{

			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose

		FileAppend, Minor inconsistencies not repaired`n , scanpstone.log
	 	exitapp 1
		

	}	


}


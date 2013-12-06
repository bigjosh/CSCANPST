;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         josh levine (http://josh.com/cmdscanpst)
;
; Script Function:
;	Allows you to run the Outlook Repair Tool (SCANPST) from a commmand line batch file
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


FileAppend, Launched on %2%... , cscanpst.log


if 0 < 2
{
	FileAppend, ERROR: Not enough parameters specified , cscanpst.log
	MsgBox CSCANPST requires at least 2 incoming parameters but it only received %0%.
	ExitApp 11

}


;Silly temp variables param1 and param2 are needed becuase ahk can't handle a command line param in an expression!

param1 = %1%


if !FileExist( param1 )
{
	FileAppend, ERROR: SCANPST.EXE not found , cscanpst.log
	MsgBox SCANPST.EXE not found at [%1%].
	ExitApp 12
}


param2 = %2%


if !FileExist( param2 )
{
	FileAppend, ERROR: PST file [%2%] not found , cscanpst.log
	MsgBox No PST file found at [%2%].
	ExitApp 13
}

param3 = %3%

	
if param3=N
{
	FileAppend, (INFO: No backup PST file will be created)...  , cscanpst.log
}



SetTitleMatchMode 2


IfWinExist Microsoft Inbox Repair Tool
{
	FileAppend, ERROR: Repair Tool is already running , cscanpst.log
	ExitApp 3

}


;Run the SCANPST exe file...
Run %1%

WinWaitActive Inbox Repair Tool


;Enter the filename to be scanned
Send %2%

Send !S


Loop 
{

	ifWinExist, Inbox Repair Tool, been canceled
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: User canceled`n , cscanpst.log
	
		exitapp 4

	}
	


	ifWinExist, Inbox Repair Tool, error prevented access
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Could not open file`n , cscanpst.log
		exitapp 5

	}


	ifWinExist, Inbox Repair Tool, in use by another
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: File already in use`n , cscanpst.log
		exitapp 6

	}


	ifWinExist, Inbox Repair Tool, does not exist
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: File not found`n , cscanpst.log
		exitapp 7

	}


	ifWinExist, Inbox Repair Tool, does not recognize the file
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: File type not recognized`n , cscanpst.log
		exitapp 8

	}


	ifWinExist, Inbox Repair Tool, error has occurred
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: An error has occurred`n , cscanpst.log
		exitapp 9

	}

	ifWinExist, Inbox Repair Tool, is read-only
	{

		WinActivate
		WinWaitActive
		Send {ESC}
		Send !C
			
		FileAppend, ERROR: PST file was read only`n , cscanpst.log
		exitapp 10

	}


	ifWinExist, Inbox Repair Tool, No errors were found

	{

	
			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose

		FileAppend, No errors found`n , cscanpst.log
	
		exitapp 0
	}
	

	ifWinExist, Inbox Repair Tool, To repair these errors
	{

		WinActivate
		WinWaitActive

		if param3=N
		   Send !M


		Send !R
		Loop 

		{


			ifWinExist, Inbox Repair Tool, The backup file 
			{

				WinActivate
				WinWaitActive


				
				Send !Y


				WinWaitClose

			}
		
		
			ifWinExist, Inbox Repair Tool, Repair complete 

			{
				WinActivate
				WinWaitActive
				Send {ENTER}
				WinWaitClose
				FileAppend, File repaired`n , cscanpst.log
				exitapp 2

			}
		



		}			



		

	}	


	ifWinExist, Inbox Repair Tool, Only minor inconsistencies were found

	{

			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose

		FileAppend, Minor inconsistencies not repaired`n , cscanpst.log
	 	exitapp 1
		

	}	


}



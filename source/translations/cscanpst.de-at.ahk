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
	FileAppend, ERROR: Nicht genug Parameter spezifziert , cscanpst.log
	MsgBox CSCANPST benötigt mindestens 2 parameter, hat aber nur %0% erhalten.
	ExitApp 11

}


;Silly temp variables param1 and param2 are needed becuase ahk can't handle a command line param in an expression!

param1 = %1%


if !FileExist( param1 )
{
	FileAppend, ERROR: SCANPST.EXE nicht gefunden , cscanpst.log
	MsgBox SCANPST.EXE nicht gefunden in [%1%].
	ExitApp 12
}


param2 = %2%


if !FileExist( param2 )
{
	FileAppend, ERROR: PST datei [%2%] nicht gefunden , cscanpst.log
	MsgBox Keine PST Datei gefunden in [%2%].
	ExitApp 13
}

param3 = %3%

	
if param3=N
{
	FileAppend, (INFO: Es wird keine backup PST datei erstellt)...  , cscanpst.log
}



SetTitleMatchMode 2


IfWinExist Tool zum Reparieren des Posteingangs
{
	FileAppend, ERROR: Reparier Tool läuft bereits , cscanpst.log
	ExitApp 3

}


;Run the SCANPST exe file...
Run %1%

WinWaitActive Tool zum Reparieren des Posteingangs


;Enter the filename to be scanned
Send %2%

Send !S


Loop 
{

	ifWinExist, Tool zum Reparieren des Posteingangs, Die Prüfung wurde abgebrochen.
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Abbruch durch Benutzer`n , cscanpst.log
	
		exitapp 4

	}
	


	ifWinExist, Tool zum Reparieren des Posteingangs, error prevented access
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Konnte Datei nicht öffen`n , cscanpst.log
		exitapp 5

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, wird von einer anderen Anwendung verwendet
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Datei bereits in benutzung`n , cscanpst.log
		exitapp 6

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, ist nicht vorhanden
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Datei nicht gefunden`n , cscanpst.log
		exitapp 7

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, -Datei nicht.
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Dateityp nicht erkannt`n , cscanpst.log
		exitapp 8

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, error has occurred
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !C
			
		FileAppend, ERROR: Ein Fehler ist aufgetreten`n , cscanpst.log
		exitapp 9

	}

	ifWinExist, Tool zum Reparieren des Posteingangs, ist schreibgeschützt
	{

		WinActivate
		WinWaitActive
		Send {ESC}
		Send !C
			
		FileAppend, ERROR: PST Datei ist Schreibgeschützt`n , cscanpst.log
		exitapp 10

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, In dieser Datei wurde kein Fehler gefunden

	{

	
			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose

		FileAppend, Kein Fehler gefunden`n , cscanpst.log
	
		exitapp 0
	}
	

	ifWinExist, Tool zum Reparieren des Posteingangs, Um diese zu beheben
	{

		WinActivate
		WinWaitActive

		if param3=N
		   Send !M


		Send !R
		Loop 

		{


			ifWinExist, Tool zum Reparieren des Posteingangs, Die Sicherungskopie
			{

				WinActivate
				WinWaitActive


				
				Send !Y


				WinWaitClose

			}
		
		
			ifWinExist, Tool zum Reparieren des Posteingangs, Die Reparatur ist abgeschlossen 

			{
				WinActivate
				WinWaitActive
				Send {ENTER}
				WinWaitClose
				FileAppend, Datei Repariert`n , cscanpst.log
				exitapp 2

			}
		



		}			



		

	}	


	ifWinExist, Tool zum Reparieren des Posteingangs, In dieser Datei wuden nur unbedeutende Inkonsistenzen gefunden

	{

			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose

		FileAppend, Minor inconsistencies not repaired`n , cscanpst.log
	 	exitapp 1
		

	}	


}



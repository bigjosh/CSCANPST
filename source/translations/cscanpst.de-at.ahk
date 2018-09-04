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


FileAppend, Gestartet mit %2%...`n , cscanpst.log


if 0 < 2
{
	FileAppend, ERROR: Nicht genug Parameter spezifziert , cscanpst.log
	MsgBox CSCANPST Mindestens 2 Parameter notwendig, hat aber nur %0% erhalten.
	ExitApp 11

}


;Silly temp variables param1 and param2 are needed becuase ahk can't handle a command line param in an expression!

param1 = %1%


if !FileExist( param1 )
{
	FileAppend, ERROR: SCANPST.EXE nicht gefunden`n , cscanpst.log
	MsgBox SCANPST.EXE nicht gefunden in [%1%].
	ExitApp 12
}


param2 = %2%


if !FileExist( param2 )
{
	FileAppend, ERROR: PST datei [%2%] nicht gefunden`n , cscanpst.log
	MsgBox Keine PST Datei gefunden in [%2%].
	ExitApp 13
}

param3 = %3%

	
if param3=N
{
	FileAppend, (INFO: Es wird keine backup PST datei erstellt)...`n  , cscanpst.log
}



SetTitleMatchMode 2


IfWinExist Tool zum Reparieren des Posteingangs
{
	FileAppend, ERROR: Reparier Tool bereits gestartet`n , cscanpst.log
	ExitApp 3

}


;Run the SCANPST exe file...
Run %1%

WinWaitActive Tool zum Reparieren des Posteingangs


;Enter the filename to be scanned
Send %2%

Send !t


Loop 
{

	ifWinExist, Tool zum Reparieren des Posteingangs, wurde abgebrochen. An der
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !N
			
		FileAppend, ERROR: Abbruch durch Benutzer`n , cscanpst.log
	
		exitapp 4

	}
	


	ifWinExist, Tool zum Reparieren des Posteingangs, Zugriff verweigert
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !N
			
		FileAppend, ERROR: Konnte Datei nicht öffen`n , cscanpst.log
		exitapp 5

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, wird von einer anderen Anwendung verwendet
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !N
			
		FileAppend, ERROR: Datei bereits in benutzung`n , cscanpst.log
		exitapp 6

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, ist nicht vorhanden
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !N
			
		FileAppend, ERROR: Datei nicht gefunden`n , cscanpst.log
		exitapp 7

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, -Datei nicht.
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !N
			
		FileAppend, ERROR: Dateityp nicht erkannt`n , cscanpst.log
		exitapp 8

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, Ein Fehler ist aufgetreten
	{

		WinActivate
		WinWaitActive
		Send {ENTER}
		Send !N
			
		FileAppend, ERROR: Ein Fehler ist aufgetreten`n , cscanpst.log
		exitapp 9

	}

	ifWinExist, Tool zum Reparieren des Posteingangs, ist schreibgesch
	{

		WinActivate
		WinWaitActive
		Send {ESC}
		Send !N
			
		FileAppend, ERROR: PST Datei ist Schreibgeschützt`n , cscanpst.log
		exitapp 10

	}


	ifWinExist, Tool zum Reparieren des Posteingangs, In dieser Datei wurden keine Fehler gefunden

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
		   Send {SPACE}


		Send !R
		Loop 

		{


			ifWinExist, Tool zum Reparieren des Posteingangs, Die Sicherungskopie
			{

				WinActivate
				WinWaitActive


				
				Send !J


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


	ifWinExist, Tool zum Reparieren des Posteingangs, In dieser Datei wurden nur unbedeutende Inkonsistenzen gefunden

	{

			
		WinActivate
		WinWaitActive
		Send {ENTER}
		WinWaitClose
	

		FileAppend, Unbedeutende Inkonsistenzen nicht Repariert`n , cscanpst.log
		exitapp 1
	 	
		

	}	


}



;-----------------------------------------------------------------------------------------------------------------------------------
;-------------------------------------------------------- Restart-Skript -----------------------------------------------------------
;-------------------------------------------- Addendum für AlbisOnWindows --------------------------------------------------
;
;                                                                               	Beschreibung
;
;                          	Neuladen eines Skriptes unter Umgehen des AHK internen 'Reload' Befehl.
;       	Notwendig geworden da dieser Befehl ein Skript relativ häufig nicht neu lädt. Durch ein Beenden
;       	des Skriptes, werden gleichzeitig auch unnötig belegte Speicherresourcen wieder freigegeben.
;           	- 	mit dem 5. Parameter (ist optional) wird das im Parameter 1 übergebene Skript durch eine
;               	AutohotkeyH_U64.exe (multithread unicode 64bit Autohotkey-Version) gestartet
;
;
;
;------------------------------------- written by Ixiko -this version is from 26.08.2020 -----------------------------------------
;--------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ---------------------------------
;------------------------ use subject: "Addendum" so that you don't end up in the spam folder -----------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ---------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------

#NoENV
#NoTrayIcon

FullAppPath       	:= A_Args.1
delay                	:= A_Args.2
scriptHwnd        	:= A_Args.3
scriptPID           	:= A_Args.4
runAutohotkeyH	:= A_Args.5

SplitPath, FullAppPath, AppName
Sleep, % (delay * 1000)

AHKH_FilePath := A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"

If (runAutohotkeyH = 1)
	Run, %AHKH_FilePath% /f "%FullAppPath%"
else
	Run, Autohotkey.exe /f "%FullAppPath%"


ExitApp


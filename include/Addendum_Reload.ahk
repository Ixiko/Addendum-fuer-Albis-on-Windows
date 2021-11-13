; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                            	Addendum_Reload
;                            ~~~~~~~~~~~~
;
;     Beschreibung:     	Neuladen eines Skriptes unter Umgehung des 'Reload' Befehls.
;       					    	Reload lädt Skripte relativ häufig nicht wirklich neu. Es verbleiben Daten im RAM,
;                                  	welche vom neugeladenen Skript verwendet werden. Diese Daten (zumeist Bildhandles)
;                                 	lassen sich nicht auffrischen, so daß Änderungen nicht sichtbar werden.
;
;                                 	Das hier per commandline übergebene Skript wird als neuer, frischer Prozeß gestartet.
;                                 	Unnötig belegter RAM wieder freigegeben.
;
;		Hinweis:	    		Parameter 5 (optional) ist true:  Ausführung mit AutohotkeyH [multithread unicode 64bit Version]
;
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;									begin: 02.04.2021,	last change 04.09.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


#NoENV
#NoTrayIcon

q := Chr(0x22)

FullAppPath       	:= A_Args.1
delay                	:= A_Args.2
scriptHwnd        	:= A_Args.3
scriptPID           	:= A_Args.4
runAutohotkeyH	:= A_Args.5

SplitPath, FullAppPath, AppName
Sleep, % (delay * 1000)

AHKH_FilePath := A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"

If (runAutohotkeyH = 1)
	Run, % AHKH_FilePath " /f " q . FullAppPath . q " " q "reload" q
else
	Run, % "Autohotkey.exe /f " q . FullAppPath . q " " q "reload" q


ExitApp


; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . .
; . . . . .                                   	ADDENDUM HAUPTSKRIPT
global                           AddendumVersion:= "2.6" , DatumVom:= "14.08.2023"
; . . . . .
; . . . . .           ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . .           BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . .                       RUNS ONLY WITH AUTOHOTKEY_H 64 BIT UNICODE VERSION
; . . . . .                       THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . .
; . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . .          !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !!
; . . . . .                      THIS SCRIPT ONLY WORKS WITH AUTOHOTKEY_H V1.1.32.00+
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

/*               	A DIARY OF CHANGES  [⏵]
⓪ ① ② ③ ④ ⑤ ⑥ ⑦ ⑧ ⑨ ⑩ ⑪ ⑫ ⑬ ⑭ ⑮ ⑯ ⑰ ⑱ ⑲ ⑳
⬝ ▪ ▫ ◽ ⬞
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

FINAL VERSION - KEINE WEITERENTWICKLUNG MEHR


- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


 ▛▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▜
 ▌                 DIENST                    	▌                           	SSID                                     	 ▌
 ▌   -––––––––––––––––––––––––––––––––––––    ▌    –––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––– 	 ▌
 ▌   	CGM_ALBIS (5432) postgre sql          	▌    	S-1-5-80-845206254-3503829181-3941749774-3351807599-4094003504 	   ▌
 ▌   	CGM_ALBIS_BACKGROUND_SERVICE_(6002)   	▌   	S-1-5-80-4257249827-193045864-994999254-1414716813-2431842843	     ▌
 ▌   	CGM_ALBIS_SMARTUPDATE_SERVICE		    		▌   	S-1-5-80-4281495583-391623409-1399029959-4115513306-232410700   	 ▌
 ▙▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▟


	Ifap
	  Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{3BA8167C-E0D3-4020-AFAF-BD0AF2BCB984}\1.1\0\win32
	  Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{3BA8167C-E0D3-4020-AFAF-BD0AF2BCB984}\1.1\HELPDIR

	Updateseite anzeigen:
	  mit OnUpdate.exe aus dem albiswin\ Stamm-Verzeichnis
	  Shell: [Albis-Pfad]\OnUpdate.exe   (z.B. X:\albiswin\OnUpdate.exe)
	  AHK: Run, % q Addendum.AlbisPath "\OnUpdate.exe" q


** Ideen
	OrderEntry Labor 	- automatische Eintragung der richtigen Ausnahmekennziffer
	Addendum_Gui 	- Patientenindividuelle Errinnerungen

*/

;{01. Skriptablaufeinstellungen / Vorbereitungen

		#Persistent
		#InstallKeybdHook
		#SingleInstance              	, Force
		#MaxThreads                  	, 50
		#MaxThreadsBuffer           	, On
		#MaxThreadsPerHotkey        	, 18
		#MaxHotkeysPerInterval      	, 99000000
		#HotkeyInterval              	, 99000000
		#MaxMem                      	, 256

		;#WarnContinuableException, On
		;~ #Warn , All

		#KeyHistory                  	, Off                    	; eingeschaltet lassen für A_Priorkey (Addendum_Gui)
		ListLines                    	, On
		Process                      	, Priority,, High  	; auf hohe Priorität setzen aufgrund der Hook-Prozesse
		SetBatchLines                	, -1

		DetectHiddenWindows         	, On
		SetTitleMatchMode           	, 2
		SetTitleMatchMode           	, Fast

		CoordMode, Mouse            	, Screen
		CoordMode, Pixel            	, Screen
		CoordMode, ToolTip           	, Screen
		CoordMode, Caret            	, Screen
		CoordMode, Menu             	, Screen
;
		SendMode                     	, Input
		SetKeyDelay                  	, 25		, 15

		SetWinDelay                  	, -1
		SetControlDelay              	, -1
		AutoTrim                     	, On
		FileEncoding                 	, UTF-8

	; effektiver Schutz gegen Doppelstart , das Skript wird erst nach einer zufälligen Verzögerung weiter ausgeführt
		Random, slTime, 1, 5
		Sleep % slTime*400
		If WinExist("Addendum Message Gui ahk_class AutohotkeyGUI")
			ExitApp
		DetectHiddenWindows, Off

	; Scriptende und -fehlerbehandlungsfunktionen festlegen
		OnExit("DasEnde")
		OnError("FehlerProtokoll")

	; Client Namen feststellen
		global compname := StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft

	; startet die Windows Gdip Funktion
		global pToken
		If !(pToken := Gdip_Startup()) {
			MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
			ExitApp
		}

;}

;{02. Variablendefinitonen , individuelles Client-Traymenu wird hier erstellt
	;----------------------------------------------------------------------------------------------------------------------------------------------
	; globale Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global AddendumDir                                                 	; Skriptverzeichnis
		global TBot                                                        	; benötigt für die Kommunikation mit Telegram-Messenger-Dienst
		global q                     	:= Chr(0x22)                         	; ist dieses Zeichen -> "
		global ac32770              	:= " ahk_class #32770"              	; spart ein paar Zeichen im Code
		global DatumZuvor
		global admServer                                                  	; für Interskriptkommuniktion im LAN

	 ; Hook's
		global EHProc		    	      	:= ""                                	; der Name des Prozesses der gehookt wurde
		global EHStack              	:= Array()
		global SHookHwnd	          	:= 0                                 	; hook handle des ShellHookProc
		global ifaptimer             	:= 0                                 	; ifap produziert Kaskaden an Fehlermeldungen, flag wird bei erster Fehlermeldung gesetzt
		global EHWHState             	:= false                            	; flag falls gerade noch der Hookhandler läuft
		global ClickReady            	:= 1                               		; für den GVU Formularausfüller

	  ; Daten sammelnde Objekte
		global ScanPool             	:= Array()                           	; enthält die Dateinnamen aller im BefundOrdner vorhandenen Dateien
		global GVU                   	:= Object()                         	; enthält Daten für die Vorsorgeautomatisierung
		global PatDB                                                       	; die Patientennamendatenbank (oPat)
		global cPat                                                        	; alle Patienten aus der Patient.dbf
		global sonderzeichen

		ImpfJahr := A_MM >= 6 ? A_YYYY : A_YYYY-1

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; ** EINSTELLUNGEN AUS DER ADDENDUM.INI EINLESEN / ADDENDUM OBJEKT MIT DATEN BEFÜLLEN **
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	/*                    	Beschreibung aller enthaltenen key : value - Paare

	             ~~~~~~~ Struktur des Addendum-Objektes ~~~~~~~
		Addendum.BefundOrdner  	= Scan-Ordner für neue Befundzugänge
		Addendum.iWin                	= Infofenster im Albis Patientenfenster
																					.Y,.W,.H     	- der Einblendungsbereich
																					.LVScanPool   - Objekt enthält die größe der ScanPool Listview (W,H)
		Addendum.Telegram          	= enumeriertes Objekt (Bsp: Addendum.Telegram[1].TBotID)
																					.BotName      - Name des Bot's
																					.Token        - Telegram Bot ID
																					.ChatID     	- Telegram Chat ID
		Addendum.GVUListe          	= Anzahl der vorhandenen Untersuchungen in der GVU-Liste

	 */

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		SplitPath, AddendumDir,,,,, AddendumDrive

		global Addendum	    	:= Object()
		Addendum.Dir         	:= AddendumDir
		Addendum.Drive       	:= AddendumDrive
		Addendum.Debug       	:= false                     	; gibt Daten in meine Standardausgabe (Scite) aus wenn = true
		Addendum.ExecStatus 	:= A_Args.1                  	; Art des Skriptstarts feststellen (Erstausfühung oder Reload)
		Addendum.AlbisWinID  	:= AlbisWinID()             	; handle des ersten gefundenen Albisfensters
		Addendum.AHK_HPath  	:= A_AppData "\AutoHotkeyH\AutoHotkeyH_U64.exe"

	;	- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	;	ein Aufruf nur mit dem Verzeichnisnamen initialisiert den Ini-Pfad der Funktion
		workini := IniReadExt(AddendumDir "\Addendum.ini")
	;	- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	; Einstellungen - Funktionen aus \include\Addendum_Ini.ahk  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		AddendumProperties()
		Addendum.Flag.AutoDocCalled	:= 0                           	; # flag für AutoDoc Funktion welche nicht mehrfach gestartet werden kann

		family := { "(G)roß"   	: ["(m)utter", "(v)ater", "(t)ante", "(o)nkel", "(ni)chte", "(ne)ffe"]
					  	, "(Ur)groß"	: ["(m)utter", "(v)ater", "(t)ante", "(o)nkel", "(ni)chte", "(ne)ffe"]
					  	, "(H)alb"   	: ["(b)ruder", "(s)chwester"]
					  	, "(Sc)hw"		: ["wager" 	 , "wägerin", "ester", "ieger"]}


	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tray Menu erstellen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		func_NoNo    		:= Func("NoNo")
		hIconAddendum 	:= Create_Addendum_ico()
		AddendumIcon  	:= hIconAddendum ? "hIcon: " hIconAddendum : Addendum.Dir "\assets\ModulIcons\Addendum.ico"

		Menu, Tray, NoStandard
		Menu, Tray, Tip			, % StrReplace(A_ScriptName, ".ahk") " V." AddendumVersion " vom " DatumVom
												. 	"`nClient: " compName SubStr("                   ", 1, Floor((20 - StrLen(compName))/2))
												. 	"`nPID: " DllCall("GetCurrentProcessId")
												. 	"`nAutohotkey.exe: " A_AhkVersion
												.		"`nStartstatus: " (Addendum.ExecStatus ? "Reload" : "1. Start")

		Menu, Tray, Color    	, % "c" Addendum.Default.BGColor3
		Menu, Tray, Add				, Addendum, % func_NoNo
		Menu, Tray, Icon     	, % AddendumIcon
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Module laden
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; authorisierte Module ins Traymenu laden
		Addendum.Errors := ""
		For each, admModul in Addendum.Module {
			subname := admModul.name, subpath := admModul.command, subicopath := admModul.ico
			If (subExist := FileExist(subpath)) && IsObject(fn_tmp := Func("ModulStarter").Bind(subname, subpath)) {
				Menu, SubMenu1, Add, % subname, % fn_tmp
				If FileExist(subicopath)
					Menu, SubMenu1, Icon, % subname, % RegExReplace(subicopath, "M.ico", ".ico")
			}
			else {
				Addendum.Errors .= !Addendum.Errors ? "[Ladefehler Addendum Module/Tools]`n" : ""
				Addendum.Errors .= "M.-Nr. " each ", Name: '" subname "' " (subExist ? "BoundFounc Objektfehler" : "ist nicht im angegebenen Pfad vorhanden") "`n"
			}
		}

	; authorisierte Tools ins Traymenu laden
		For each, admTool in Addendum.Tools {
			subname := admTool.name, subpath := admTool.command, subicopath := admTool.ico
			If (subExist := FileExist(subpath)) && IsObject(fn_tmp := Func("ToolStarter").Bind(subname, subpath)) {
				Menu, SubMenu2, Add, % subname, % fn_tmp
				If FileExist(subicopath)
					Menu, SubMenu2, Icon, % subname, % RegExReplace(subicopath, "M.ico", ".ico")
			}
			else
				Addendum.Errors .= "M.-Nr. " each ", Name: '" subname "' " (subExist ? "BoundFounc Objektfehler" : "ist nicht im angegebenen Pfad vorhanden") "`n"
		}

		 If IsObject(fn_tmp := Func("ToolStarter").Bind("Windows komplett neu starten", "--")) {
			Menu, SubMenu2, Add	, % "Windows komplett neu starten", % fn_tmp
			Menu, SubMenu2, Icon	, % "Windows komplett neu starten", % Addendum.Dir "\assets\ico\Windows.ico"
		}

	; Fehlerausgabe
		If Addendum.ErrorProtocol && Addendum.Errors
			FileAppend, % Addendum.Errors, % A_Temp "\AddendumErrors.txt"

		subname := subpath := subicopath := subExist := fn_tmp :=admModul := admTool := Addendum.Errors := ""

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; per Checkbox zu- oder abschaltbare Funktionen (SubMenu3)
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; Albis                                                                        	;{
	  ; Autologin
		AAL := Addendum.Autologin = " (Logindaten fehlen)" ? Addendum.Autologin : ""
		Menu, SubMenu36, Add, % "Albis Autologin" . AAL                                  	, Menu_Autologin
		Menu, SubMenu36, % (AAL || !Addendum.Autologin ? "UnCheck":"Check")             	, % "Albis Autologin" . AAL
	  ; Automatisierung von GVU Formulare
		Menu, SubMenu36, Add, % "GVU automatisieren"                                    	, Menu_GVUAutomation
		Menu, SubMenu36, % (Addendum.GVUAutomation ? "Check":"UnCheck")                 	, % "GVU automatisieren"
	  ; PopUpMenu
		Menu, SubMenu36, Add, % "PopUpMenu"                                              	, Menu_PopUpMenu
		Menu, SubMenu36, % (Addendum.PopUpMenu ? "Check":"UnCheck")                     	, % "PopUpMenu"

	  ; schnelles Löschen im Wartezimmer
		Menu, SubMenu366, Add, % "schnelles Löschen im Wartezimmer"                      	, Menu_Wartezimmer_AutoDelete
		Menu, SubMenu366, % (Addendum.AutoDelete.WZ ? "Check":"UnCheck")                	, % "schnelles Löschen im Wartezimmer"
	  ; schnelles Löschen von Dauermedikamenten
		Menu, SubMenu366, Add, % "schnelles Löschen von Dauermedikamenten"              	, Menu_Dauermed_AutoDelete
		Menu, SubMenu366, % (Addendum.AutoDelete.Dauermed ? "Check":"UnCheck")          	, % "schnelles Löschen von Dauermedikamenten"
	  ; schnelles Löschen von Dauerdiagnosen
		Menu, SubMenu366, Add, % "schnelles Löschen von Dauerdiagnosen"                 	, Menu_Dauerdia_AutoDelete
		Menu, SubMenu366, % (Addendum.AutoDelete.Dauerdia ? "Check":"UnCheck")           	, % "schnelles Löschen von Dauerdiagnosen"

		Menu, SubMenu36, Add, % "Einträge ohne Nachfrage entfernen", :SubMenu366

	  ; Schnellrezept
		Menu, SubMenu36, Add, % "Schnellrezept"                                          	, Menu_Schnellrezept
		Menu, SubMenu36, % (Addendum.Schnellrezept ? "Check":"UnCheck")                  	, % "Schnellrezept"
	  ; Addendum Toolbar
		Menu, SubMenu36, Add, % "Addendum Toolbar"                                       	, Menu_AddendumToolbar
		Menu, SubMenu36, % (Addendum.ToolbarThread ? "Check":"UnCheck")                 	, % "Addendum Toolbar"

		Menu, SubMenu3, Add, % "Albis", :SubMenu36

	; Prüfung EBM/KRW
		Menu, SubMenu39, Add, % "EBM/KRW Hinweise"                                       	, Menu_EBMKRW
		Menu, SubMenu39, % (Addendum.EBMKRW ? "Check":"UnCheck")                         	, % "EBM/KRW Hinweise"
	; Prüfung EBM/KRW TimeOut
		Menu, SubMenu39, Add, % "Timeout EBM/KRW Hinweise (" Addendum.EBMKRWTimeout "s)"	, Menu_EBMKRWTimeout

		Menu, SubMenu3, Add, % "AutoClose", :SubMenu39
	;}

	; AutoSizer für Albis, Ifap, MS Word, FoxitReader, Sumatra PDF                	;{
		Menu, SubMenu31, Add, % "Albis AutoSize"                                         	, Menu_AlbisAutoPosition
		Menu, SubMenu31, % (Addendum.AlbisLocationChange ? "Check":"UnCheck"    )	      	, % "Albis AutoSize"
		For classnn, appname in Addendum.Windows.Proc {
			Menu, SubMenu31, Add, % appname, Menu_AutoSize
			Menu, SubMenu31, Icon, % appname, % A_ScriptDir "\res\" appname ".ico"
			Menu, SubMenu31, % (Addendum.Windows[appname].AutoPos & Addendum.MonSize ? "Check":"UnCheck"), % appname
		}
		Menu, SubMenu3, Add, % "AutoSizer", :SubMenu31
	;}

	; FoxitReader-Automatisierung                                                 	;{
		Menu, SubMenu37, Add, % "Signaturhilfe"                                          	, Menu_PDFSignatureAutomation
		Menu, SubMenu37, Add, % "signiertes Dokument schliessen"                         	, Menu_PDFSignatureAutoTabClose
		Menu, SubMenu37, % (Addendum.AutoCloseFoxitTab ? "Check":"UnCheck")             	, % "signiertes Dokument schliessen"
		Menu, SubMenu37, % (Addendum.PDFSignieren ? "Check":"UnCheck")                  	, % "Signaturhilfe"
		Menu, SubMenu37, % (Addendum.PDFSignieren ? "Enable":"Disable")                  	, % "signiertes Dokument schliessen"

		Menu, SubMenu3, Add, % "FoxitReader", :SubMenu37
	;}

	; Infofenster (AddendumGui)                                                    	;{
		Menu, SubMenu32, Add, % "Infofenster"                                            	, Menu_AddendumGui
		Menu, SubMenu32, % (Addendum.AddendumGui         	? "Check":"UnCheck")          	, % "Infofenster"
		If Addendum.AddendumGui {
			Menu, SubMenu32, Add, % "Einzelbestätigung für Importvorgänge"		          		, Menu_Einzelbestaetigung
			Menu, SubMenu32, % (Addendum.iWin.ConfirmImport	? "Check":"UnCheck")           	, % "Einzelbestätigung für Importvorgänge"
			Menu, SubMenu32, Add, % "Autozuweisung Wartezimmer"                            	, Menu_AutoWZ
			Menu, SubMenu32, % (Addendum.iWin.AutoWZ       	? "Check":"UnCheck")          	, % "Autozuweisung Wartezimmer"
			Menu, SubMenu32, Add, % "Abrechnungshelfer"		                     	           	, Menu_Abrechnungshelfer
			Menu, SubMenu32, % (Addendum.iWin.AbrHelfer    	? "Check":"UnCheck")          	, % "Abrechnungshelfer"
		}
		Menu, SubMenu3, Add, % "Infofenster", :SubMenu32
    ;}

	; Mail Management                                                             	;{
		If Addendum.Mail.Outlook {
		Menu, SubMenu38, Add, % "Outlook FaxMailManager"                                 	, Menu_FaxMailManager
		Menu, SubMenu38, % (Addendum.Mail.FaxMail ? "Check":"UnCheck")                  	, % "Outlook FaxMailManager"
		Menu, SubMenu3, Add, % "Mail Management", :SubMenu38
		}
	;}

	; Laborabruf/import automatisieren                                            	;{
		Menu, SubMenu34, Add, % "Laborabruf", Menu_LaborAbrufManuell
		Menu, SubMenu34, % (Addendum.Labor.AutoAbruf ? "Check":"UnCheck")               	, % "Laborabruf"

		; zeitgesteuerter Abruf
		;  dieses Menu ist nur auf den Clients sichtbar, deren Name in der Addendum.ini
		;  unter Labor/RunOnlyClients hinterlegt sind (Komma getrennte Liste)
		LabCallClients := StrReplace(Addendum.Labor.ExecuteOn, " ")
		If compname in %LabCallClients%
		{
			Menu, SubMenu34, Add, % "zeitgesteuerter Abruf"                               	, Menu_LaborAbrufTimer
			Menu, SubMenu34, % (Addendum.Labor.AbrufTimer 	? "Check":"UnCheck")          	, % "zeitgesteuerter Abruf"
			If Addendum.Labor.AbrufTimer && !Addendum.Labor.AutoAbruf {
				Addendum.Labor.AutoAbruf	:= true            ; Timer ist an, dann auch der AutoAbruf
				IniWrite, % "ja", % Addendum.Ini, % compname                                	, % "Laborabruf_Timer"
			}
		}
		else If Addendum.Labor.AbrufTimer {
				Addendum.Labor.AbrufTimer	:= false
				IniWrite, % "nein"	, % Addendum.Ini, % compname	                          	, % "Laborabruf_Timer"
			}

		Menu, SubMenu34, Add, % "Laborimport", Menu_LaborImportManuell
		Menu, SubMenu34, % (Addendum.Labor.AutoImport   	? "Check":"UnCheck")           	, % "Laborimport"

		Menu, SubMenu3	, Add, % "Laborabruf automatisieren", :SubMenu34
	;}

	; Laborjournal                                                                 	;{
		Menu, SubMenu35, Add, % "Täglich anzeigen"                                      	, Menu_Laborjournalanzeige
		Menu, SubMenu35, Add, % "Startzeit: " Addendum.Laborjournal.StartTime            	, Menu_LaborjournalStartzeit
		Menu, SubMenu35, % (Addendum.Laborjournal.AutoView ? "Check":"UnCheck")         	, % "Täglich anzeigen"
		Menu, SubMenu3	, Add, % "Laborjournal", :SubMenu35
	;}

	; Texterkennung                                                                	;{
		Menu, SubMenu33, Add, % "Befundordner überwachen"                               	, Menu_WatchFolder
		Menu, SubMenu33, % (Addendum.OCR.WatchFolder     	? "Check":"UnCheck")          	, % "Befundordner überwachen"
		If (Addendum.OCR.Client = compname) {
			Menu, SubMenu33, Add, % "AutoOCR"                                              	, Menu_AutoOCR
			Menu, SubMenu33, % (Addendum.OCR.AutoOCR      	? "Check":"UnCheck")          	, % "AutoOCR"
		}
		Menu, SubMenu3, Add, % "Texterkennung", :SubMenu33
	;}

	; Kopien von DICOM-CDs anfertigen                                             	;{
		If DriveTypeExist("CDROM") {                                         	; Menu wird nur Clients mit angeschlossenem CD-Laufwerk angezeigt
			Menu, SubMenu40, Add, % "Kopie anlegen automatisieren"                        	, Menu_DICOM_AutoCopy
			Menu, SubMenu40, % (Addendum.Dicom.AutoCopy ? "Check":"UnCheck")              	, % "Kopie anlegen automatisieren"
			Menu, SubMenu40, Add, % "Stammpfad für Kopien festlegen"                      	, Menu_DICOM_Dir
			Menu, SubMenu40, Add, % "Muster für Unterordner angeben"                      	, Menu_DICOM_DirPattern
			Menu, SubMenu3, Add, % "DICOM-CD", :SubMenu40
		}
	;}

	; MicroDicom Dateiexport automatisieren                                        	;{
		If GetAppImagePath("mDicom.exe") {                           		; wird nur angezeigt wenn MicroDicom installiert ist
			Menu, SubMenu3, Add, % "MicroDicom Export automatisieren"	                     	, Menu_MDicom
			Menu, SubMenu3, % (Addendum.mDicom ? "Check":"UnCheck")	                       	, % "MicroDicom Export automatisieren"
		} else {
			Addendum.mDicom := false
		}
	;}

	; Clipboardüberwachung                                                         	;{
		Menu, SubMenu45, Add, % "Patientendaten erkennen"                                	, Menu_ClipWatchPatients
		Menu, SubMenu45, % (Addendum.Clip.WatchPatients   	? "Check":"UnCheck")        	, % "Patientendaten erkennen"
		Menu, SubMenu3, Add, % "Clipboardüberwachung", :SubMenu45
	;}

	; TrayTips anzeigen                                                            	;{
		Menu, SubMenu3, Add, % "TrayTips zeigen"                                        	, Menu_TrayTips
		Menu, SubMenu3, % (Addendum.ShowTrayTips ? "Check":"UnCheck")                   	, % "TrayTips zeigen"
	;}

	; Faxbox                                                                       	;{
		fn_Faxbox := Func("Faxbox").Bind(false)
		If (Addendum.PDF.Faxbox.State = "An" && Addendum.PDF.Faxbox.Client = compname) {
			Menu, SubMenu3, Add, % "Faxabruf"                                             	, Menu_Faxbox
			Menu, SubMenu3, % (Addendum.PDF.Faxbox.State = "An" ? "Check":"UnCheck")      	, % "Faxabruf"
		}
	;}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; weitere Menupunkte
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; DATEN / PROTOKOLLE

		func_call := Func("ZeigePatDB")
		Menu, SubMenu4, Add, % "Patienten"                         	, % func_call
		Menu, SubMenu4, Add

		func_call := Func("ZeigeFehlerProtokoll")
		Menu, SubMenu4, Add, % "aktuelles Fehlerprotokoll"        	, % func_call

		func_call := Func("AddendumObjektSpeichern")
		Menu, SubMenu4, Add, % "Addendum Objekt"                   	, % func_call

		func_call := Func("ShowPDFObjects")
		Menu, SubMenu4, Add, % "PDFPool und PatDocs Objekt"       	, % func_call



	; Protokolle/Datenbanken anzeigen
		If FileExist(protokoll := Addendum.LogPath           	"\LaborabrufLog.txt")                        	{
			func_call 	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Laborabruf Protokoll"           	, % func_call
		}
		If FileExist(protokoll := Addendum.LogPath          	"\LaborimportLog.txt")                       	{
			func_call  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Laborimport Protokoll"          	, % func_call
		}
		If FileExist(protokoll := Addendum.LogPath         	"\OnExit-" A_YYYY "-Protokoll.txt")           	{
			func_call := Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "OnExit Protokoll"                	, % func_call
		}
		If FileExist(protokoll := Addendum.LogPath           	"\PDfImportLog.txt")                         	{
			func_call  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "PDF Importprotokoll"            	, % func_call
		}
		If FileExist(protokoll := Addendum.LogPath          	"\OCRTime_Log.txt")                          	{
			func_call  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "PDF OCR Protokoll"              	, % func_call
		}
		If FileExist(protokoll := Addendum.LogPath          	"\WatchFolder-Log.txt")                      	{
			func_call   	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "OCR WatchFolder"                	, % func_call
		}
		If FileExist(protokoll := Addendum.LogPath          	"\AddendumMonitorLog.txt")                  	{
			func_call   	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Addendum Monitor"               	, % func_call
		}
		If FileExist(protokoll := IniReadExt("ScanPool", "FritzFaxbox_Telefonbuch"))                      	{
			func_call   	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Telefonbuch"                    	, % func_call
		}
		If FileExist(protokoll := IniReadExt("ScanPool", "FritzFaxbox_Filterliste"))                      	{
			func_call   	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Fritz!FaxBox Filter"             	, % func_call
			func_call   	:= ""
		}
		If FileExist(protokoll := Addendum.DBPath "\Dictionary\Befundbezeichner.json")                     	{
			func_call   	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Befundbezeichner"               	, % func_call
			func_call   	:= ""
		}


	; FARBEN festlegen
		Menu, SubMenu1, Color                            	, % "c" Addendum.Default.BGColor3
		Menu, SubMenu2, Color                            	, % "c" Addendum.Default.BGColor3
		Menu, SubMenu3, Color                            	, % "c" Addendum.Default.BGColor3
		Menu, SubMenu4, Color                            	, % "c" Addendum.Default.BGColor3

	; MENU erstellen
		Menu, Tray, Add, Module starten/beenden         	, :SubMenu1
		Menu, Tray, Add, Daten / Protokolle             	, :SubMenu4
		Menu, Tray, Add, andere Tools                   	, :SubMenu2
		Menu, Tray, Add, Einstellungen                   	, :SubMenu3

		Menu, Tray, Icon, Module starten/beenden 	      	, % Addendum.Dir "\assets\ModulIcons\Tools2.ico"
		Menu, Tray, Icon, Daten / Protokolle             	, % Addendum.Dir "\assets\ModulIcons\Daten.ico"
		Menu, Tray, Icon, andere Tools                   	, % Addendum.Dir "\assets\ModulIcons\Tools1.ico"
		Menu, Tray, Icon, Einstellungen                  	, % Addendum.Dir "\assets\ModulIcons\Einstellungen.ico"

		Menu, Tray, Add,
		Menu, Tray, Add, Zeige Skript Variablen          	, scriptVars
		Menu, Tray, Add, Skript Neu Starten             	, SkriptReload
		Menu, Tray, Add, Beenden                         	, DasEnde

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Einlesen der Patientendaten aus der Patienten.txt Datei im Addendum Datenbankordner
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	  ; es sind drei Patienten Datenobjekte im Moment!
	  ; oPat enthält nur Name, Vorname, Geburtsdatum, Geschlecht und Kasse der Patienten
	  ; PatDB enthält die für den Betrieb von Addendum notwendigen Daten aller Patienten aus der PATIENT.dbf (Felder siehe >> infilter + EMailadressen)
	  ; cPat ist die PatDB als Funktionsobjekt. Per Methodenaufruf kann man im Objekt nach Daten suchen. Sogar eine unscharfe Suche ist möglich.
		PatDB 	:= ReadPatientDBF(Addendum.AlbisDBPath)     ; wird nicht mehr benötigt?
		cPat		:= new PatDBF(""	, [	"NR", "VERSART", "PRIVAT", "NAME", "VORNAME", "GEBURT", "NAMENORM", "TELEFON"
									      			, 	"TELEFON2", "FAX", "HAUSARZT", "ARBEIT", "DEL_DATE", "RKENNZ", "STRASSE", "PLZ", "ORT"
										      		, 	"HAUSNUMMER", "GESCHL", "SEIT", "LAST_BEH", "MORTAL"]
									      			, 	"moredata=true")

	;}
;}

;{03. Hotkey command's

HotkeyLabel:
	;=---------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Hotkey's die überall funktionieren
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------ ;{

	#UseHook, On
	Hotkey, ^!n                    	, SkriptReload              	, I1      	;= Überall: Addendum.ahk wird neu gestartet
	Hotkey, LControl & n     	, SkriptReload              	, I1      	;= Überall: Addendum.ahk wird neu gestartet

	fnCall := Func("SkriptReload")
	Hotkey, ^+n                 	, % fnCall                                 	;= Überall: Addendum.ahk wird neu gestartet
	Hotkey, ^!ä                  	, BereiteDichVor                       	;= Überall: beendet das Addendumskript
	Hotkey, CapsLock          	, DoNothing                             	;= Überall: Capslock kann nicht mehr versehentlich gedrückt werden

	fnCall := Func("ListviewClipboard")
	Hotkey, ^+c		    		, % fnCall                                 	;= Überall: Inhalt aus Standard-Windows-Listview Steuerelementen kopieren
	#UseHook, Off

	Hotkey, !m                    	, MenuSearch                          	;= Überall: Menu Suche - funktioniert in allen Standart Windows Programmen
	Hotkey, +!e                  	, EnableAllControls                   	;= Überall: inaktivierte Bedienfelder in externen Fenstern aktivieren
	Hotkey, Pause               	, AddendumPausieren             	;= Überall: Addendum legt eine Pause ein
	Hotkey, ^!F10               	, SendClipBoardText                	;= Überall: sendet den Inhalt des Clipboards als simulierte Tasteneingabe

	fnCall := Func("SendAlternately").Bind("„|”")
	Hotkey, ^+2            		, % fnCall                                 	;= Überall: Sendet Anführungszeichen unten und dann oben
	;}

	;=---------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Scite4AutoHotkey
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------ ;{
	fn_Send := {"^NumPad1"            	: Func("SendRawFast").Bind(";{ ")             	;= SciTe: ;{
					,	"^NumPad2"            	: Func("SendRawFast").Bind(";}")              	;= SciTe: ;}
					,	"^´"                          	: Func("SendRawFast").Bind("``r"	, "")     	;= SciTe: `r
					,	"!´"                              	: Func("SendRawFast").Bind("``n"	, "")  	;= SciTe: `n
					,	"^!´"                       		: Func("SendRawFast").Bind("`t`t"	, "")  	;= SciTe: `t
					,	"^5"                         	: Func("SendRawFast").Bind("%%", "L")     	;= SciTe:
					,	"^,"                          	: Func("SendAlternately").Bind("/* | */")  	;= SciTe: /
					,	"^-"                          	: Func("SendAlternately").Bind("{ | }")       	;= SciTe:
					,  "NumPad4"                	: "SendStrgLeft"                                       	;= SciTE:  Strg+Pfeil links
					,  "NumPad6"           	     	: "SendStrgRight"                                  	;= SciTE:  Strg+Pfeil rechts
					,  "LAlt & v"       	             	: "SciteDescriptionBlock"                      	;= SciTe: Abschnittbezeichner für Addendumskripte
					,  "LControl & Numpad4"  	: "SciteSmartCaretL"                            	;= SciTe: for language depend moving of caret
					,  "LControl & Numpad6"  	: "SciteSmartCaretR"                            	;= SciTe: for language depend moving of caret
					,  "LAlt & Left"                 	: "SendFourBackspace"                         	;= SciTE: sends 4 times backspace
					,  "LAlt & Right"               	: "SendFourSpace"}                              	;= SciTE: sends 4 times space

	Hotkey, IfWinActive, ahk_class SciTEWindow
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	For hkey, funcCall in fn_Send
		Hotkey, % hkey, % funcCall
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	fn_Send := funcCall := hkey := ""

	;}

	;=---------------------------------------------------------------------------------------------------------------------------------------------------------------
	; andere Fenster
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------ ;{
	Hotkey, IfWinActive, ScanSnap Manager - Bild erfassen und Datei speichern ahk_class #32770
	Hotkey, Enter              	, ScanBeenden

	func_GetPatNameFromExplorer := Func("GetPatNameFromExplorer")
	Hotkey, IfWinActive, % "ahk_class CabinetWClass"
	Hotkey, ~F6                	, % func_GetPatNameFromExplorer
	Hotkey, ~^RButton         	, % func_GetPatNameFromExplorer

	fn_FoxitFontSize   	:= Func("HoveredControl").Bind("classFoxitReader", "Edit\d+")
	func_DeInkrementer	:= Func("DeInkrementer")
	Hotkey, If, % fn_FoxitFontSize
	Hotkey, WheelUp 	         	, % func_DeInkrementer
	Hotkey, WheelDown         	, % func_DeInkrementer

	;}

	;=---------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Albis
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------ ;{
	fn_VhK         	:= Func("AlbisAusfuellhilfe").Bind("Muster12", "Hotkey")
	fn_Undo     	:= Func("UnRedo").Bind("Undo")
	fn_Redo     	:= Func("UnRedo").Bind("Redo")
	;~ fn_cLP       	:= Func("isActiveFocus").Bind("contraction=(lp|bg)")
	fn_cLP       	:= Func("HotS_Privatabrechnung_IF")
	fn_LPTexte		:= Func("Privatabrechnung_Begruendungen")
	fn_AIndis   	:= Func("AusIndisKorrigieren")
	fn_Abr       	:= Func("AlbisAbrechnungAktiv")
	fn_SchnellZ 	:= Func("Schnellziffer")
	fn_COVIDA    	:= Func("ScheinCOVIDPrivat")
	fn_UpDown   	:= Func("IFActive_UpDown")
	fn_ItemUD   	:= Func("ItemUpDown")
	fn_PAusweis 	:= Func("PatAusweisDruck")
	fn_ATD      	:= Func("AlbisDatum").Bind("Heute")
	fn_AZD      	:= Func("AlbisDatum").Bind("Zeilendatum")
	fn_copy      	:= Func("AlbisCopier").Bind("copy")
	fn_cut        	:= Func("AlbisCopier").Bind("cut")
	fn_paste    	:= Func("AlbisEinfuegen")
	If Addendum.CImpf.CNRRun {                                                      	; Funktion bei Hotstrings/EBM
		fn_IFLko  	:= Func("IFActive_Lko")
		fn_CNR    	:= Func("EBM_COVIDImpfung")
	}

	Hotkey, IfWinActive, ahk_class OptoAppClass			                   		;~~ ALBIS Hotkey - bei aktiviertem Fenster
	Hotkey, $^c			    		, % fn_copy                                    	;= Albis: selektierten Text ins Clipboard kopieren
	Hotkey, $^x			    		, % fn_cut               			       	   	;= Albis: selektierten Text ausschneiden
	Hotkey, $^v			    		, % fn_paste                                   	;= Albis: Clipboard Inhalt (nur Text) in ein editierbares Albisfeld einfügen
	Hotkey, $^a		        		, AlbisMarkieren                            	;= Albis: gesamten Textinhalt markieren
	Hotkey, $^7	       	        	, % fn_PAusweis                               	;= Albis: Patientenausweis sofort drucken
	Hotkey, $^z	                 	, % fn_Undo                                  	;= Albis: Eingaben rückgängig machen (Edit, RichEdit)
	Hotkey, $^d	                	, % fn_Redo                                   	;= Albis: Eingaben wiederherstellen
	Hotkey, $^!-                	, % fn_AIndis                                  	;= Albis: Ausnahmeindikationen werden korrigiert
	Hotkey, ^!i                   	, % fn_COVIDA                                	;= Albis: Abrechnungsschein für COVID-19 Impfung Privatpatienten
	Hotkey, !F5		        		, % fn_ATD             			         		;= Albis: Einstellen des Programmdatum auf den aktuellen Tag
	Hotkey, ^!F5	        		, % fn_AZD            			         		;= Albis: Einstellen des Programmdatum auf das Zeilendatum
	Hotkey, !F8		    			, DruckeLabor                                	;= Albis: Laborblatt als pdf oder auf Papier drucken
	Hotkey, ^F11		    		, Schulzettel	         			   	    	;= Albis: Schulzettel.ahk starten
	Hotkey, ^F12		    		, Hausbesuche         			     	    	;= Albis: Hausbesuche.ahk starten
	Hotkey, !Down       			, AlbisKarteiKarteSchliessen     	        	;= Albis: per Kombination die aktuelle Karteikarte schliessen
	Hotkey, !Up               		, AlbisNaechsteKarteiKarte	     	        	;= Albis: per Kombination eine geöffnete Kartekarte vorwärts
	Hotkey, !Right    				, Laborblatt    			            		;= Albis: Laborblatt des Patienten anzeigen
	Hotkey, !Left  					, Karteikarte      			            		;= Albis: Karteikarte des Patienten anzeigen
	Hotkey, ^!+l                	, Scheinloeschen                             	;= Albis: Menu Patient/Schein/Löschen aufrufen
	Hotkey, ^!+ä                	, Scheinaendern                             	;= Albis: Menu Patient/Schein/Ändern aufrufen
	Hotkey, ^!j                    	, PrivatAU                                     	;= Albis: Menu Formular/Privat-AU
	;~ Hotkey, ^F7               	, GVUListe_ProgDatum                        	;= Albis: akt. Patienten in die GVU-Liste eintragen
	;~ Hotkey, !F7                 	, GVUListe_ZeilenDatum                      	;= Albis: akt. Patienten in die GVU-Liste eintragen (Zeilendatum)

	Hotkey, IfWinExist, ahk_class OptoAppClass                          		;~~ mehrmonatiges Kalender-Addon
	Hotkey, $#H                  	, HotkeyViewer                                 	;= Albis: zeigt die Hotkeys an
	Hotkey, $F9			    		, MinusEineWoche                            	;= Albis: addiert eine Woche einem Datum hinzu, z.B. im AU Formular
	Hotkey, $F10		        	, PlusEineWoche                            	;= Albis: subtrahiert eine Woche von einem Datum, z.B. im AU Formular
	Hotkey, $F11		    		, MinusEinMonat                            	;= Albis: addiert 4x7 Tage (4 Wochen) auf ein Datum
	Hotkey, $F12		        	, PlusEinMonat                                	;= Albis: subtrahiert 4x7 Tage (4 Wochen) von einem Datum
	Hotkey, $+F3					, Kalender	                                    	;= Albis: Shift+F3 Kalender Erweiterung

	Hotkey, If, % fn_UpDown                                                         	;~~ Einträge sortieren in versch. Fenstern (Dauer..., Familie, Cave)
	Hotkey, ^Up                    	, % fn_ItemUD                                  	;= Albis: schiebt die ausgewählte Zeile um eine Zeile nach oben
	Hotkey, ^Down                  	, % fn_ItemUD                           		;= Albis: schiebt die ausgewählte Zeile um eine Zeile nach unten

	Hotkey, IfWinActive, Muster 12a ahk_class #32770                   	;~~ Ausfüllhilfe für das Formular Verordnung häusliche Krankenpflege
	Hotkey, Up                    	, % fn_Vhk                                     	;= Albis: löscht ein Datumsfeld
    Hotkey, Down                   	, % fn_Vhk                                     	;= Albis: übernimmt vorhandene Datumseinträge

	Hotkey, If, % fn_cLP
	Hotkey, $^F6                  	, % fn_LPTexte                                	;= Begründungstexte GOÄ-Ziffern anzeigen, bearbeiten, verwenden

	If IsFunc(fn_CNR) {                                                                	;~~ Chargennummern COVID-19 Impfungen aus Liste übernehmen
		Hotkey, IF	, % fn_IFLko
		Hotkey, $NumpadAdd      	, % fn_CNR
		Hotkey, !a                	, COVIDImpfA
		Hotkey, !b                	, COVIDImpfB
		Hotkey, !r                	, COVIDImpfR
	}

	;}

	;=---------------------------------------------------------------------------------------------------------------------------------------------------------------
	; flexible Hotkey's und Fensterkombinationen, je nach Einstellung in der Addendum.ini
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	If Addendum.PDFSignieren 	{ ; Signatur setzen im Foxit Reader nach Hotkey

		func_FRSignature := func("FoxitReaderSignieren")
		Hotkey, IfWinActive, % "ahk_class " Addendum.PDF.ReaderWinClass
		Hotkey, % Addendum.PDFSignieren_Kuerzel, % func_FRSignature     	;= FoxitReader: Signatur setzen

	}

	;}
;}

;{04. Hotstringabfragen erstellen  (Hotkey, If ....)

		global hsUser

		HotS_View_EBM()
		HotS_View_Privatabrechnung()
		HotS_View_Diagnosis()
		HotS_Scan_Hotstrings()

;}

;{05. Timer, Threads, Initiliasierung WinEventHooks, OnMessage, TCP-Kommunikation, eigene Dialoge anzeigen

	;-----------------------------------------------------------------------------------------------------------------------------------
	; Clipboard Change
		If IsFunc("ClipChanged")
			OnClipboardChange("ClipChanged")
		else
			SciTEOutput("ClipChanged ist keine Funktion")

	;-----------------------------------------------------------------------------------------------------------------------------------
	; Timerfunktionen
	;---------------------------------------------------------------------------------------------------------------------------------;{
	; Skript neu starten um 0 Uhr
		func_call := Func("SkriptReload").Bind("AutoRestart")
		SetTimer, % func_call, % -1*TimerTime("00:00 Uhr")

	; Albis einmal täglich neu starten (abschaltbar)
	; läuft Albis auf einem Client welcher TI Geräte nutzt, lassen sich keine Versichertenkarten mehr einlesen
		if Addendum.Albis.AutoRestart
			AlbisDailyRestart()

	; Laborabruf Timer
		LaborAbrufTimer()

	; Laborjournal Timer
		;~ If Addendum.Laborjournal.AutoView
			;~ LaborjournalTimer(Addendum.Laborjournal.StartTime)
	;}

	;-----------------------------------------------------------------------------------------------------------------------------------
	; SetWinEventHooks /Shellhook initialisieren
		InitializeWinEventHooks()

	;-----------------------------------------------------------------------------------------------------------------------------------
	; Skriptkommunikation / Anpassen von Fenstern bei Änderung der Bildschirmaufläsung
		OnMessage(0x4A    	, "Receive_WM_COPYDATA")                	; Interskriptkommunikation
		OnMessage(0x7E    	, "WM_DISPLAYCHANGE")                    	; Änderung der Bildschirmauflösung
		If (DriveTypeExist("CDROM") && Addendum.Dicom.AutoCopy)     	; Funktion wird nur eingeschaltet wenn ein CDROM Laufwerk angeschlossen ist
			OnMessage(0x219  	, "WM_DEVICECHANGE")                     	; neue CD wurde eingelegt
		;~ OnMessage(0x200 	, "ScreenEdgeDetection")                	; läßt den Mauszeiger nicht am Monitorrand kleben wenn zwei Monitore unterschiedliche
																							                		; Auflösungen haben

	;-----------------------------------------------------------------------------------------------------------------------------------
	; TCP Server starten - für LAN Kommunikation
		If Addendum.LAN.IamServer {
			global admServer
			admServer := new SocketTCP()
			admServer.bind("addr_any", 12345) ; Addendum.LAN.admServer
			admServer.listen()
			admServer.onAccept := Func("admGui_OnAccept")
			admServer.onDisconnect("admGui_OnDisconnect")
			;admStartServer()
		}

	;-----------------------------------------------------------------------------------------------------------------------------------
	; Hotfolder - Überwachung des Befundordners
		If Addendum.OCR.WatchFolder {
			WatchFolder(Addendum.BefundOrdner, "admGui_FolderWatch", False, (1+2+8+64))
			Addendum.OCR.WatchFolderStatus := "running"
		}

	;-----------------------------------------------------------------------------------------------------------------------------------
	; Fritzbox Faxabruf
		If (compname=Addendum.PDF.Faxbox.Client && Addendum.PDF.Faxbox.State="An") {
			SetTimer, % fn_Faxbox, % Addendum.PDF.Faxbox.Interval * 60 * 1000 ; Angabe in Minuten
			Faxbox(false)     ;  erster Abruf sofort
		}

	;-----------------------------------------------------------------------------------------------------------------------------------
	; Infofenster gleich anzeigen wenn eine Karteikarteikarte geöffnet ist
		If RegExMatch(Addendum.AktiveAnzeige := AlbisGetActiveWindowType(), "i)(Karteikarte|Laborblatt|Biometriedaten|Karteifenster)")	{
			If Addendum.AddendumGui
				AddendumGui()
			PatDb(AlbisTitle2Data(), "exist")
		}


;}

return

;{06. Hotstrings

;:*:#at::@
^!l::ListLines

;{ 	6.1 -- ALBIS
; --- EBM	- Leistungskomplexe                                                         	;{
#If HotS_EBM_IF()
HotS_EBM_IF()                                                                                  	{              	;-- für kontextsensitive Auslösung

	If WinActive("ahk_class OptoAppClass") {
		afocus := AlbisGetFocus()
		If (RegExMatch(afocus.childs.Kuerzel.text, "i)^lk") || InStr(AlbisGetActiveWindowType(), "i)|Abrechnung"))
			return true
	}
return false
}
HotS_View_EBM()                                                                          	{

  ; schreibe * um alle Hotstrings im Kontext anzuzeigen oder
  ; schreibe Suchmuster z.B. Zuschlag* um in den Hotstrings zu suchen
  ; die Hotstring-Suche wird erst nach der Eingabe von Space/Tab/Enter ausgeführt

  ; IF Bedingung - START
    fn_IFPAbr	:= Func("HotS_EBM_IF")
	Hotkey, If, % fn_IFPAbr
	Hotstring(":?*:*", Func("HotS_View").Bind(A_LineNumber+3, "EBM"))

return
}
; ----- ---- Hotstrings ---- ----- ;{
; Gespräch
:R*:gs::03230(x:2)                                                                        	; Gesprächziffer
:R*:ps1::35110(x:2)                                                                      	; Psychosomatikziffer 1
:R*:psy1::35110(x:2)                                                                   	; Psychosomatikziffer 1
:R*:ps2::35110(x:2)                                                                      	; Psychosomatikziffer 2
:R*:psy2::35110(x:2)                                                                  	; Psychosomatikziffer 2
:R*:neuropä::04430                                                                   	; Neuropädiatrisches Gespräch …

; Kommunikation
:R*:Term::03008-03010-                                                             	; Zuschläge Terminvermittlung (Terminvermittlung Facharzt, TSS-Terminvermittlung)
:R*:Vermit::03008-03010-                                                            	; Zuschläge Terminvermittlung (Terminvermittlung Facharzt, TSS-Terminvermittlung)
:R*:VeArzt::86900-01660-                                                            	; eArztbrief-Versandpauschale
:R*:VeBrief::86900-01660-                                                           	; eArztbrief-Versandpauschale
:R*:VeMail::86900-01660                                                            	; eArztbrief-Versandpauschale
:R*:EeArzt::86901-                                                                      	; eArztbrief-Empfangspauschale
:R*:EeBrief::86901-                                                                      	; eArztbrief-Empfangspauschale
:R*:EeMail::86901-                                                                      	; eArztbrief-Empfangspauschale
:R*:ZeArzt::01660-                                                                       	; Zuschlag zur eArztbrief-Versandpauschale (Strukturförderpauschale)
:R*:ZeBrief::01660-                                                                      	; Zuschlag zur eArztbrief-Versandpauschale (Strukturförderpauschale)
:R*:ZeMail::01660-                                                                      	; Zuschlag zur eArztbrief-Versandpauschale (Strukturförderpauschale)
:R*:ZTelek::01670-                                                                      	; Einholung eines Telekonsiliums - Zuschlag für Vers.-,Grund- oder Konsiliarp
:R*:Btelek::01671-                                                                      	; Telekonsiliarische Beurteilung einer medizinischen Fragestellung
:R*:FTelek::01672-                                                                      	; Zuschlag zur 01671 für die Fortsetzung der telekonsiliarischen Beurteilung
:R*:porto::40110-                                                                       	; Porto-Kostenpauschale
:R*:fax::40111-                                                                           	; Fax-Kostenpauschale
;:R*:fax::40120(x:2)-                                                                    	; Faxgebühr (*Gebühr wurde gestrichen)
;:R*:tel::80230(x:2)-                                                                    	; Telefongebühr mit dem Krankenhaus (*Gebühr wurde gestrichen)

; Vorsorgen
:*X:gvu::AlbisGVU()                                                                   	; GVU/HKS inkl. eHKS-Formular
:R*:aneu::01747-                                                                        	; Aufklärungsgespräch Ultraschall-Screening Bauchaortenaneurysmen
:R*:aorta::01747-																    		; nur Männer ab 65 Jahren einmalig berechnungsfähig!
:R*:colo::01740-                                                                          	; Coloskopievorsorgeempfehlung Männer/Frauen ab 45 Jahren
:R*:hks::01745-01746-                                                               	; Hautkrebsscreening - Männer/Frauen ab 35 Jahren alle 3 Jahre berechnungsfähig
:R*:kvu::01731-                                                                           	; Krebsvorsorge - nur Männer ab 55 Jahren alle 3 Jahre berechnungsfähig
:R*:U1::01711                                                                             	; U1
:R*:U2::01712                                                                             	; U2
:R*:U3::01713                                                                             	; U3
:R*:U4::01714                                                                             	; U4
:R*:U5::01715                                                                             	; U5
:R*:U6::01716                                                                             	; U6
:R*:U7 ::01717                                                                            	; U7
:R*:U7a::01723                                                                           	; U7a
:R*:U8::01718                                                                             	; U8
:R*:U9::01719                                                                             	; U9
:R*:J1::01720                                                                             	; J1
:*:Kind::                                                                                      	; Vorsorgeuntersuchungen Kinder inkl. J1
:X*:UUnt::                                                                                   	;{ Vorsorgeuntersuchungen Kinder inkl. J1
SendRaw, % "{RAW}01711-01712-01713-01714-01715-01716-01717-01718-01719-01723-01720"
return ;}
:R*:entw::03350-03351                                                               	; Orientierende entwicklungsneurologische Untersuchung eines Neugeborenen,
																									; Säuglings, Kleinkindes oder Kindes
; Formulare
:R*:Belast::01610                                                                      	; Bescheinigung zur Belastungsgrenze
:R*:Reha::01611                                                                        	; Verordnung von medizinischer Rehabilitation
:R*:Verordnung Reha::01611                                                       	; Verordnung von medizinischer Rehabilitation
:R*:besch::01620                                                                      	; Muster 50 - Bescheinigung oder Zeugnis
:R*:Muster52::01621                                                                 	; Muster 52 -
:R*:kurp::01622                                                                         	; Muster 20 - Kurplan, Gutachten, Stellungnahme
:R*:gut::01622                                                                          	; Muster 20 - Kurplan, Gutachten, Stellungnahme
:R*:mdk::01622                                                                         	; Muster 20 - Kurplan, Gutachten, Stellungnahme
:R*:kurv::01623                                                                         	; Muster 25 - Kurvorschlag
:R*:Mutter::01624                                                                     	; Muster 54 - Verordnung medizinischer Vorsorge für Mütter oder Väter
:R*:Vater::01624                                                                       	; Muster 54 - Verordnung medizinischer Vorsorge für Mütter oder Väter
:R*:Erstverordnung pall::01425-                                                  	; Erstverordnung der spezialisierten ambulanten Palliativversorgung
:R*:Vpall1::01425-                                                                    	; Erstverordnung der spezialisierten ambulanten Palliativversorgung
:R*:Vpall2::01426-                                                                    	; Folgeverordnung der spezialisierten ambulanten Palliativversorgung
:R*:Folgeverordnung pall::01426-                                               	; Folgeverordnung der spezialisierten ambulanten Palliativversorgung
:R*:Zweitverordnung pall::01426-                                               	; Folgeverordnung der spezialisierten ambulanten Palliativversorgung

; Gerätemedizin
:R*:lzrr::03324                                                                            	; Langzeitblutdruckmessung
:R*:spiro::03330                                                                            	; Spirographische Untersuchung
:R*:lufu::03330                                                                            	; Lungenfunktionstest mit dem Spirometer
:R*:farb::50100                                                                            	; Prüfung des Farbsinn

; Labor und Labor-Sonderziffern
:R*:sars::32006-88240-03230                                                     	; COVID19 - Laborbudget und Sonderziffer und Gespräch
:R*:covid19::32006-88240-03230                                               	; COVID19 - Laborbudget und Sonderziffer
:R*:app::02402-02403                                                                	; bei Corona-App Warnung abzurechnen
:R*:Muster10C::32779-32811-32816                                          	; Abrechnungsspezialitäten Corona bei Verwendung Muster 10C
:R*:bsg::32042                                                                            	; BSG - Blutsenkungsgeschwindigkeit

; Stix
:R*:sang::01737-                                                                         	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:R*:Stuhl::01737-                                                                         	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:R*:iFop::01737-                                                                          	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:R*:haemof::01737-                                                                    	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:R*:bz::32025                                                                             	; Blutzuckermessung
:R*:glucose::32025                                                                      	; Blutzuckermessung
:R*:zucker::32025                                                                       	; Blutzuckermessung
:R*:stix::32030                                                                             	; Urinstix
:R*:urin::32030                                                                           	; Urinstix
:R*:ustix::32030                                                                           	; Urinstix
:R*:urinstix::32030                                                                       	; Urinstix
:R*:pipi::32030                                                                          	; Urinstix
:R*:inr::32026-                                                                           	; INR Messstreifen
:R*:gerinnung::32026-                                                                	; INR Messstreifen

; Wunden/Verbände
:R*:Kompr::02313-                                                                      	; Kompressionsverband
:R*:chronische::02313-                                                               	; Behandlungskomplex chronische Wunde
:R*:chirurgie1::02300-                                                                 	; Kleinchirurgischer Eingriff I und/oder
:R*:Wunde1::02300-                                                                 	; Kleinchirurgischer Eingriff I und/oder
:R*:eingriff1::02300-																		; primäre Wundversorgung und/oder Epilation
:R*:chirurgie2::02301-                                                                 	; Kleinchirurgischer Eingriff II und/oder
:R*:Wunde2::02301-                                                                 	; Kleinchirurgischer Eingriff II und/oder
:R*:eingriff2::02301-																		; primäre Wundversorgung mittels Naht
:R*:chirurgie3::02302-                                                                 	; Kleinchirurgischer Eingriff III und/oder primäre Wundversorgung
:R*:Wunde3::02302-                                                                 	; Kleinchirurgischer Eingriff III und/oder primäre Wundversorgung
:R*:eingriff3::02302-    																	; bei Säuglingen, Kleinkindern und Kindern

; Impfungen
:*:FSME::89102ABR{LShift Down}{Left 3}{LShift Up}
:*:ig1::89111-
:*:gsi1::89111-
:*:ig2::89112-
:*:gsi2::89112-
:*:Grippe::89111-89112

; Hausbesuche
:R*:Verwe::01440(x:2)                                                                 	; Verweilen außerhalb der Praxis
:R*:hb0::01410-97234-03230(x:3)                                               	; Hausbesuch (bestellt) {Left 11}
:R*:hb1::01411(um:19:00)-97234-03230(x:3)               		 		; Hausbesuch nach Feierabend
:R*:hb2::01412(um:19:00)-97234-03230(x:3)                       	   	; Hausbesuch dringend, sofort, Feiertage
:R*:hbWE1::01410-97234-03230(x:3)                                          	; Hausbesuch Wochenende - geplant
:R*:hbWE2::01412(um:19:00)-97234-03230(x:3)                         	; Hausbesuch Wochenende - ungeplant
:R*:hbFei::01412(um::00)-97234-03230(x:3)                              	; Hausbesuch Feiertag
:R*:hbAbe::01411(um::00)-97237-03230(x:3)                            	; Hausbesuch dringend ausserhalb Sprechzeit {Left 21}
:R*:hbSpre::01412(um::00)-97234-03230(x:3)                          	; Hausbesuch dringend, aus der Sprechstunde heraus {Left 21}

; OP-Vorbereitung/Nachbehandlung
;~ :R*:opvor::31010-31011-31012-31013                                       	; OP-Vorbereitung
:X*:opvor::  ;{
 tosend := AgeDependFee("OP-Vorbereitung")                          	; OP-Vorbereitung
 thiskeydelay := A_KeyDelay
 ;~ SciTEOutput("this keydelay: " thiskeydelay)
 SetKeyDelay , 50
 Send, % "{RAW}" tosend
 Send, % "{RAW}-"
 SetKeyDelay, % thiskeydelay
return ;}
:R*:op vor::31010-31011-31012-31013                                     	; OP-Vorbereitung
:R*:op-vor::31010-31011-31012-31013                                     	; OP-Vorbereitung
:R*:kata::31012-31013-88115(OPS:5-144.5a:RLB)                       	; Vorbereitung Katarakt-OP auf Überweisung
:R*:Op60::31013-88115(OPS:5­144.5a:RLB)(text:CataractOP)		;  Vorbereitung Katarakt-OP auf Überweisung > 60 Jahre
:R*:OP40::31012-88115(OPS:5­144.5a:RLB)(text:CataractOP)			;  Vorbereitung Katarakt-OP auf Überweisung > 60 Jahre
:R*:postop::31600                                                                     	; postoperative Nachbehandlung (nur bei Überweisung vom Facharzt)

; Sonderziffern
:R*:pall::03370-03371-03372-03373-01425-01426                    	; Palliativziffern
:R*:1.Hi::90057-                                                                          	; Versicherter wurde am gleichen oder Vortag
:R*:Rettu::90057-																			; in der Rettungsstelle behandelt und abgewiesen
:R*:Angehör::90230                                                                     	; Gespräch mit Angehörigen nach 03230/04230 von Patienten
																									; mit Pflegegrad zur Koordination der Versorgung
																									; nach Krankenhausaufenthalt
; Corona
;:*:prior::88320                                                                          	; beendet? Ausstellung Zeugnis im Kontext der CoronaImpfV (§9)
:*:berat::88322                                                                          	; Ausschließliche Impfberatung (§ 9 Abs. 2)
:*:COVID::88331-88332-88333-88334-88323-88324             	; Corona Impfung im Hausbesuch
:*:coronaHB::88323-88324                                                       	; Corona Impfung im Hausbesuch
:*:cIm::88323-88324                                                                 	; Corona Impfung im Hausbesuch
:*:covidIm::88323-88324                                                           	; Corona Impfung im Hausbesuch
:*:covidHB::88323-88324                                                           	; Corona Impfung im Hausbesuch
:*:coronaImpfHB::88323-88324                                                	; Corona Impfung im Hausbesuch
:X*:bt1::SendGB("88331A(charge:xxxxxx)")                                     	; Comirnaty 1.Impfung
:X*:bta::SendGB("88331A(charge:xxxxxx)")                                     	; Comirnaty 1.Impfung
:X*:bt2::SendGB("88331B(charge:xxxxxx)")                                  	; Comirnaty 2.Impfung
:X*:btb::SendGB("88331B(charge:xxxxxx)")                                     	; Comirnaty 1.Impfung
:X*:bt3::SendGB("88331R(charge:xxxxxx)")                                  	; Comirnaty 3.Impfung - Auffrischung/Booster
:X*:bt4::SendGB("88331R(charge:xxxxxx)")                                  	; Comirnaty 4.Impfung - Auffrischung/Booster
:X*:btr::SendGB("88331R(charge:xxxxxx)")                                     	; Comirnaty 3.Impfung - Auffrischung/Booster
:X*:btP1::SendGB("88331G(charge:xxxxxx)")                                  	; Comirnaty 1.Impfung - Pflegeheim
:X*:btP2::SendGB("88331H(charge:xxxxxx)")                                   	; Comirnaty 2.Impfung - Pflegeheim
:X*:mo1::SendGB("88332A(charge:xxxxxx)")                                 	; Moderna 1.Impfung - normal
:X*:moA::SendGB("88332A(charge:xxxxxx)")                                 	; Moderna 1.Impfung - normal
:X*:mo2::SendGB("88332B(charge:xxxxxx)")                                  	; Moderna 2.Impfung - normal
:X*:moB::SendGB("88332B(charge:xxxxxx)")                                 	; Moderna 1.Impfung - normal
:X*:mo3::SendGB("88332R(charge:xxxxxx)")                                  	; Moderna 2.Impfung - Auffrischung/Booster
:X*:mo4::SendGB("88332R(charge:xxxxxx)")                                  	; Moderna 2.Impfung - Auffrischung/Booster
:X*:moR::SendGB("88332R(charge:xxxxxx)")                                 	; Moderna 2.Impfung - Auffrischung/Booster
 ; COVID-19-Impfzertifikat
:*:Cert1::88350                                                                        	; geimpft in eigener Praxis (Web)
:*:CertP::88351                                                                        	; geimpft in eigener Praxis (PVS)
:*:CertE1::88352                                                                        	; nicht selbst geimpft (1.Impfung)
:*:CertE2::88353                                                                        	; nicht selbst geimpft (2.Impfung)
:*:CertG::88371                                                                       	; Ausstellung Genesenenzertifikats W‒ automatisiert mithilfe des PVS
:*:genesen1::88370                                                                  	; Ausstellung eines COVID-19-Genesenenzertifikats (Web)
:*:genesenW::88370                                                                 	; Ausstellung eines COVID-19-Genesenenzertifikats (Web)
:*:genesen2::88371                                                                  	; Ausstellung Genesenenzertifikats W‒ automatisiert mithilfe des PVS
:*:genesenP::88371                                                                  	; Ausstellung Genesenenzertifikats W‒ automatisiert mithilfe des PVS
:*:igx::
	thiskeydelay := A_KeyDelay
	;~ SciTEOutput("this keydelay: " thiskeydelay)
	SetKeyDelay , 50
	Send, % "{Raw}89111-"
	Sleep 300
	SendInput, {Tab}
	Sleep 500
	SendInput, {Escape}
	Sleep 300
	SendInput, {Escape}
	Sleep 300
	Send, {Input}
	Sleep 600
	If !ControlGetText("", "ahk_id " GetFocusedControl()) {
		Send, % "{Raw}dia"
		SendInput, {Tab}
		Sleep 500
		If !ControlGetText("", "ahk_id " GetFocusedControl()) {
			Send, % "{Raw}E22, G. {Z25.1G};"
			Sleep 300
			SendInput, {Tab}
			Sleep 500
			If !ControlGetText("", "ahk_id " GetFocusedControl()) {
				SendInput, {Escape}
				Sleep 100
				SendInput, {Escape}
				Sleep 200
			}
		}
	}
	;~ If ControlGetText("", "ahk_id " GetFocusedControl()) {
		;~ SendInput, {Tab}
		;~ Sleep 500
	;~ } else {

	;~ }
	;~ }
	SetKeyDelay, % thiskeydelay
return
:*:igy::89112

; Hausarztprogramme
:*:DAK::93550-93555-93560-93565-93570                             	; HAP DAK
:*:AOK::95051-                                                                          	; AOK

; Quartalsziffern (nur einmal pro Quartal)
:*:1-::03000-03040-                                                                   	; 03000-03040
:*:2-::03220                                                                                	; 03220 - chronisch krank Kennzeichnung
:*:c1::                                                                                     		;{03220 - chronisch krank Kennzeichnung + automatisierte Überprüfung
	Send, % "{RAW}03220-"
	InChronicList(AlbisAktuellePatID())
return ;}
:*:3-::03221                                                                                	; 03221 - chronisch krank Kennzeichnung
:*:c2::                                                                                          	;{03221 - chronisch krank Kennzeichnung + automatisierte Überprüfung
	Send, % "{RAW}03221-"
	InChronicList(AlbisAktuellePatID())
return ;}
:*:4-::03360-                                                                             	; geriartrischer Basiskomplex 03360
:*:5-::03362-                                                                             	; geriartrischer Basiskomplex 03362
:*:gb::03360-                                                                           	; geriartrischer Basiskomplex 03360
:*:gb2::03362-                                                                           	; geriartrischer Basiskomplex 03362

; Gebühren für Verwaltung von Patientendaten
:*:ePADo::01431                                                                       	; Daten in der ePA erfassen, verarbeiten und/oder speichern,
:*:ePA Daten ohne::01431                                                          	;{ ohne persönlichen Arzt-Patienten-Kontakt und keine Videosprechstunde,
                                                                                                 	; (4x/Arztfall) (3 Punkte - *Orientierungspunktwert 2021: 110,1244 Cent) ;}
:*:ePADm::01647                                                                       	; Zuschlag zu den Vers.-, Grund- oder Konsiliarp. (1x/Behandlungsfall)
																									; (Zusatzpauschale ePA-Ünterstützungsleistung)
:*:VOApp::01470                                                                       	;{ Erstverordnung einer DiGA aus dem DiGA-Verzeichnis, auch im Rahmen
																									; einer Videosprechstunde, mehrfach im BHF abrechenbar, wenn mehrere
																									; DiGA verordnet wurden mit Benennung der  DiGA ;}
; sonstiges
:*:tens::30712-                                                                          	; Anleitung des Patienten zur Selbstanwendung der transkutanen elektrischen Nervenstimulation




; ### ----- ----- ----- ----- ;}
;}
; --- GOÄ	- Privatabrechnung                                                          	;{
#If HotS_Privatabrechnung_IF()
HotS_Privatabrechnung_IF()                                                            	{ ;-- für kontextsensitive Auslösung
	If WinActive("ahk_class OptoAppClass") {
		afocus := AlbisGetFocus()
		If (RegExMatch(afocus.childs.Kuerzel.text, "i)\b(lp|lbg)\b") || InStr(AlbisGetActiveWindowType(), "Privatabrechnung"))
			return true
	}
return false
}
HotS_View_Privatabrechnung()                                                           	{

  ; schreibe *Suchmuster z.B. *Zuschlag um in den Hotstrings zu suchen

  ; IF Bedingung - START

    fn_IFPAbr		:= Func("HotS_Privatabrechnung_IF")
	Hotkey, If, % fn_IFPAbr
	Hotstring(":?*:*"	, Func("HotS_View").Bind(A_LineNumber+4, "GOÄ"))    ;
	Hotstring(":?*:#"	, Func("Privatabrechnung_Begruendungen"))    ;
	HotS_Privatabrechung()

return
}
; ----- ---- Hotstrings ---- ----- ;{
; Anamnese
:*:Fremd::4-                                                                             	; Fremdanamnese/Unterweisung
:*:Unterweisung::4-                                                                      	; Fremdanamnese/Unterweisung
:*:Anamnese Kin::807-                                                                   	; biographische psychiatrische Anamnese Kind/Jugendlicher
:*:Anamnese erw::809-                                                                    	; biographische psychiatrische Anamnese Erwachsener
:*:Erheb::835-                                                                           	; Fremdanamneseerhebung

; Beratung
:R*:Bera1::1(fak:3,5:Beratung < 10 Minuten)-                                            	; Beratung < 10 Minuten
:R*:Bera2::3(fakt:2,3:Beratung mind. 10 Minuten)-                                       	; Beratung mind. 10 Minuten
:R*:Bera3::3(fakt:3,5:Beratung < 20 Minuten)-                                           	; Beratung < 20 Minuten
:R*:Bera4::34(fakt:2,3:Beratung, eingehend mindestens 20 Minuten)-                      	; Beratung eingehend mindestens 20 Minuten
:*:Lebens::34-                                                                           	; Beratung, Erörtern einer Lebensveränderung
:*:hb::50(dkm:4)-                                                                        	; Hausbesuch
:*:verweil::56-                                                                          	; Verweilen außerhalb der Praxis
:*:zuschl::A-B-C-D-K1                                                                   	; Zuschläge
:*:Außer::A-                                                                            	; Leistung außerhalb der Sprechstunde
:*:Sofort::E-                                                                            	; Sofortzuschlag
:*:Nacht::F-                                                                             	; Nachtzuschlag
:*:Tief::G-                                                                              	; Tiefenachtzuschlag
:*:WE::H-                                                                                	; Wochenend-/Feiertagzuschlag
:*:Wochen::H-                                                                           	; Wochenend-/Feiertagzuschlag
:*:Vis::45-                                                                              	; Visite
:*:Zweitv::46-                                                                           	; Zweitvisite

; kleine Chirurgie
:*:Wunde klein E::2000                                                                  	; kleine Wunde Erstversorgung
:*:Wunde klein N::2001                                                                  	; kleine Wunde Naht
:*:Wunde klein U::2002                                                                  	; kleine Wunde Umschneidung und Naht
:*:Wunde groß E::2003                                                                    	; große Wunde Erstversorgung
:*:Wunde groß N::2004                                                                    	; große Wunde Naht
:*:Wunde groß U::2005                                                                   	; große Wunde Umschneidung und Naht
:*:Nekrosena::2006                                                                      	; Nekrosenabtragung
:*:Fäden::2007                                                                          	; Fäden, Klammern, Entfernung
:*:Fistelspaltung::2008                                                                  	; Wund-/Fistelspaltung
:*:Fistels::2009                                                                         	; Hautoberfläche/Schleimhaut, Fremdkörper

; Diagnostik
:*:ekg::650
:*:lzRR::654                                                                             	; Langzeitblutdruckmessung
:*:lufu::605                                                                             	; Lungenfunktion
:*:Lungenf::605                                                                          	; Lungenfunktion
:R*:oxy::602(text:SpO2)                                                                 	; Pulsoxymetrie
:R*:Sät::602(text:SpO2)                                                                  	; Pulsoxymetrie
:R*:SaO2::602(text:SpO2)                                                                	; Pulsoxymetrie
:R*:Pulsoxy::602(text:SpO2)                                                             	; Pulsoxymetrie
:R*:sono1::410-420                                                                       	; Ultraschall von bis zu drei weiteren Organen im
																							                                      			; Anschluss an Nummern 410 bis 418, je Organ. Organe sind anzugeben.
:R*:sono2::410(organ:Leber)-420(organ:Gb)-420(organ:Niere bds.)                      	  	; Ultraschall mehrere Untersuchungen
:R*:sonoknie::410(organ:Kniegelenk )                                                    	; Ultraschall Knie
:R*:sono knie::410(organ:Kniegelenk )                                                   	; Ultraschall Knie

; Labor - Blutabname / Urin- oder Probenuntersuchung
:R*:BA::250-                                                                            	; Blutabnahme
:R*:BSG::3501-                                                                           	; BSG
:R*:BZ::                                                                                 	;{ Blutzucker
:R*:Blutzuc::                                                                            	; Blutzucker
:R*:glucose::                                                                            	; Blutzucker
:R*:glukose::3560-                                                                      	; Blutzucker
;}
:R*:vGTT::                                                                               	;{ venöser Glukosetoleranztest
:R*:venöser Gluk::250(fakt:3,5:4 weitere Blutentnahme im Abstand von 30 min)-3612-
;}
:R*:oGTT::                                                                               	;{ oraler Glukosetoleranztest
:R*:oraler gluk::250(fakt:3,5:4 weitere Blutentnahme im Abstand von 30 min)-3613-
;}
:R*:inr::3530-                                                                         		; INR mit Gerät
:*:sang::3500-                                                                           	;{ Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:Stuhl::3500-                                                                         	;  Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:iFop::3500-                                                                           	;  Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:haemof::3500-                                                                          ;} Ausgabe iFOPT Test (Stuhl auf Sanguis)
:R*:stix::3652-                                                                          	;{ Urinstix
:R*:streifen::3652-                                                                     	;  Urinstix
:R*:urin::3652-                                                                         	;  Urinstix
:R*:ustix::3652-                                                                         	;  Urinstix
:R*:urinstix::3652-					                                                             	;} Urinstix

; Labor - Abstriche
:R*:Abstrich1::298-                                                                      	; Entnahme u. gegebenenfalls Aufbereitung v. Abstrichmaterial z. mikrobiologischen Untersuchung
:R*:AbstrichRa::298(ltext:Abstrich Rachenraum)-                                         	; Entnahme u. gegebenenfalls Aufbereitung v. Abstrichmaterial z. mikrobiologischen Untersuchung
:R*:Rachenab::298(ltext:Abstrich Rachenraum)-                                           	; Entnahme u. gegebenenfalls Aufbereitung v. Abstrichmaterial z. mikrobiologischen Untersuchung
:R*:AbstrichNa::298(ltext:Abstrich Nase)-                                               	; Entnahme u. gegebenenfalls Aufbereitung v. Abstrichmaterial z. mikrobiologischen Untersuchung
:R*:Nasenab::298(ltext:Abstrich Nase)-                                                   	; Entnahme u. gegebenenfalls Aufbereitung v. Abstrichmaterial z. mikrobiologischen Untersuchung

; Untersuchung
:*:neuro::800-                                                                           	; eingehende neurologische Untersuchung
:*:psy U::801-                                                                           	; psychiatrische Untersuchung
:*:psych U::801-                                                                         	; psychiatrische Untersuchung
:*:psychiatrische::801-                                                                  	; psychiatrische Untersuchung
:*:psyU::801-                                                                            	; psychiatrische Untersuchung
:*:rekt::11-                                                                             	; rektale Untersuchung
:*:Unters1::7(fakt:2,3)-                                                                 	; Vollständige Untersuchung – ein Organsystem
:*:Unters2::7(fakt:3,5)-                                                                 	; Vollständige Untersuchung – mehrere Organsysteme
:*:Geruchs::825-                                                                         	; neurologische Geruchs- und Geschmacksprüfung
:*:Geschm::825-                                                                          	; neurologische Geruchs- und Geschmacksprüfung
:*:Gleichg::826-                                                                         	; neurologische Gleichgewichtsprüfung
:*:Aphasie::830-                                                                         	; Eingehende Prüfung auf Aphasie
:*:Apraxie::830-                                                                         	; Eingehende Prüfung auf Apraxie
:*:Alexie::830-                                                                          	; Eingehende Prüfung auf Alexie

; Behandlung
:*:psych B1::804-                                                                       	; psychiatrische Behandlung [Punktzahl 150]
:*:psychB1::804-                                                                         	; psychiatrische Behandlung [Punktzahl 150]
:*:psych B2::806-                                                                        	; psychiatrische Behandlung [Punktzahl 250]
:*:psychB2::806-                                                                        	; psychiatrische Behandlung [Punktzahl 250]
:*:psy E::808-                                                                          	; Psychotherapie, Einleitung/Verlängerung
:*:psyE::808-                                                                            	; Psychotherapie, Einleitung/Verlängerung
:*:psy T::808-                                                                           	; Psychotherapie, Einleitung/Verlängerung
:*:psyT::808-                                                                           	; Psychotherapie, Einleitung/Verlängerung
:*:verbale::849-                                                                        	; verbale Intervention - Psychosomatik (20 Minuten)
:*:psyI::849-                                                                            	; verbale Intervention - Psychosomatik (20 Minuten)
:*:psy1::849-                                                                            	; verbale Intervention - Psychosomatik (20 Minuten)
:*:Einl::15-                                                                             	; Einleitung/Koordination
:*:Koord::15-                                                                            	; Einleitung/Koordination
:*:verband::200-                                                                        	; Verband

; Medikamentengabe / Infektionen / Infusionen
:R*:infil1::267-                                                                         	; Medikamentöse Infiltrationsbehandlung, je Sitzung
:R*:infil2::268-                                                                         	; Medikamentöse Infiltrationsbehandlung im Bereich mehrerer Körperregionen
:R*:infusion k::271-                                                                    	; Infusion < 30 min
:R*:infuk::271-                                                                          	; Infusion < 30 min
:R*:infu1::271-                                                                          	; Infusion < 30 min
:R*:infusion l::272-                                                                     	; Infusion > 30 min
:R*:inful::272-                                                                         	; Infusion > 30 min
:R*:infu2::272-                                                                         	; Infusion > 30 min
:R*:inj sc::252-                                                                         	; Injektion s.c., i.m., i.c.
:R*:injsc::252-                                                                          	; Injektion s.c., i.m., i.c.
:R*:inj ic::252-                                                                         	; Injektion s.c., i.m., i.c.
:R*:injic::252-                                                                          	; Injektion s.c., i.m., i.c.
:R*:inj im::252-                                                                         	; Injektion s.c., i.m., i.c.
:R*:injim::252-                                                                          	; Injektion s.c., i.m., i.c.
:R*:inj iv::253-                                                                         	; Injektion i.v.
:R*:injiv::253-                                                                         	; Injektion i.v.
:R*:iv::253-                                                                             	; Injektion i.v.
:R*:Medik::76(ltext:Erstellung Medikationsplan)                                       		; Medikationsplan
:R*:Medpl::76(ltext:Erstellung Medikationsplan)                                       		; Medikationsplan

; Impfungen
:R*:impf::375-                                                                           	; Impfung GOÄ Ziffer
:R*:influvac::(sach:Influenzaimpfstoff Influvac Tetra: 9.56)                            	; Impfung Grippe Ziffern + Influvacpreis
:R*:grlp::1-5-375-(sach:Influenzaimpfstoff Influvac: 9.56)                              	; Impfung Grippe Ziffern + Influvacpreis

; sonstige Gebühren
:*:kopie::                                                                               	;{ Berechnung Kopiekosten
	AlbisKopiekosten("0.50 ab 50 Seiten 0.15€")                                            	; bis zu 50 Seiten 0,50 € darüber hinaus 0,15 € je Seite
return ;}
:*:reiseunf::(sach:Reiseunfähigkeitsbescheinigung:20.00)                                 	;{ Reiseunfähigkeitsbescheinigung
:*:reiserüc::(sach:Reiseunfähigkeitsbescheinigung:20.00)                                 	;} Reiseunfähigkeitsbescheinigung
:R*:schreib::(sach:Schreibgebühr:3.50)                                                   	; Schreibgebühr

; Früherkennungsuntersuchungen, Erwachsene
:R*:gvu::29-                                                                             	; Früherkennungsuntersuchung, Erwachsener
:R*:kvu::28-                                                                             	; Krebsvorsorge, Mann
; ### ----- ----- ----- ----- ;}
;}
; --- Diagnosen                                                                        	;{
#If HotS_Diagnosis_IF()
HotS_Diagnosis_IF()                                                                          	{                 	;-- kontextsensitive Auslösung Diagnosenfelder

	; letzte Änderung 26.06.2022

									;, "Muster 12a ahk_class #32770"                    	: "(1|2|3|4)"                	; Häusliche Krankenpflege
									; "Muster 1a ahk_class #32770"                       	: "(4|5|6|7|8|9|10)"	; Arbeitsunfähigkeitsbescheinigung

	static diaControls := {"Muster 2 ahk_class #32770"                       	: "(1)"                            ; Krankenhauseinweisung
									, "Muster 6 ahk_class #32770"                      	: "(5|6|7)"                    	; Überweisung
									, "Muster 56 ahk_class #32770"                    	: "(1)"                        	; Antrag auf Kostenübernahme
									, "Muster G1204 ahk_class #32770"                 	: "(12|16|20|24)"        	; Befundbericht Rehabilitationsantrag
									, "Ärztliches Gutachten für ahk_class #32770"	: "(15|18|21|24)"        	;
									, "Schein von ahk_class #32770"                     	: "(40)"}                      	; Abrechnungsschein - Überweiserdaten
	static dbg := false

	If WinActive("ahk_class OptoAppClass") {
		If IsObject(afocus := AlbisGetFocus()) && dbg
			SciTEOutput(cJSON.Dump(afocus, 1))
		If (afocus.childs.KUERZEL.text= "dia")
			return true
	}
	else if WinActive("Akutdiagnose ahk_class #32770") || WinActive("Dauerdiagnosen ahk_class #32770")
		return true
	else {
		If (A_ThisHotkey = ":*?:?")
			return false
		For WinTitle, nn in diaControls
			If WinActive(WinTitle) {
				ControlGetFocus, cfocus, % WinTitle
				If RegExMatch(cFocus, "i)^Edit" nn "$")
					return true
			}
	}

return false
}
HotS_View_Diagnosis()                                                                    	{                 	;-- für kontextsensitive Auslösung

	global hsUser

  ; Schreibe ein Sternchen (*) um alle Hotstrings des Kontext anzuzeigen.
  ; Die Hotstring-Suche wird sofort nach Eingabe des Sternchen ausgeführt!

	fn_IFPAbr 	:= Func("HotS_Diagnosis_IF")

	Hotkey, If, % fn_IFPAbr
	Hotstring(":*?:*", Func("HotS_View").Bind(A_LineNumber+29, "Diagnosen"))
	Hotstring(":*?:?", Func("HotS_View").Bind(A_LineNumber+28, "Diagnosen"))

  ; Instanthotstring Objekt anlegen
	If !IsObject(hsUser)
		hsUser      	:= Object()
	If !IsObject(hsUser.Path)
		hsUser.Path 	:= Object()
	If !IsObject(hsUser[2])
		hsUser[2]   	:= Object()

  ; Pfad festlegen
	If !FilePathCreate(hsUser.Path.dia := Addendum.DBPath "\Dictionary\NutzerHotstrings_Diagnosen.hsd") {
		; ### Protokoll schreiben
		return
	}

  ; Hotstringdaten laden und Hotstrings erstellen
	If FileExist(hsUser.Path.dia) {
		If IsObject((hsUser[2] := FileOpen(hsUser.Path.dia, "r", "UTF-8").Read()))
			For dia, hsd in hsUser[2]
				Hotstring(":X" hsd.opt ":" hsd.hs, Func("SendDiagnosis").Bind(dia))
	}

return
}
; ----- ---- Hotstrings ---- ----- ;{

; Meniskopathie - Auswahlbox


/*  ▪ ▪ Hotstrings für die Eingabe von ICD-Diagnosen ▪ ▪

	Allgemeines:
	▪	Die hier verwalteten ICD-Diagnosen sind in meiner Praxis häufiger anzutreffen.
	▪	Die Liste ist nicht vollständig und wird daher fortwährend erweitert.

	Krankheitsbezeichnungen:
	▪	Die Hotstrings basieren einerseits auf fachlichen Krankheitsbezeichnungen, sollen aber auch für das nicht-medizinische Personal hilfreich sein.
		Daher werden, soweit vorhanden, auch die deutschen Krankheitsbezeichnungen verwendet.

	Hotstringeingabe:
	▪	Die Eingabe ist kontextsensitiv. D.h. das sich Diagnosen-Hotstrings nur innerhalb von Albis und auch nur an den von Albis vorgesehenen Stellen
		auslösen lassen.
	▪	Eingeleitet wird die Eingabe eines Hotstrings immer mit einem Punkt '.', gefolgt von der Eingabe einer Abkürzung/Wortverkürzung (Hotstring).
	▪	Die Abkürzungen sind eindeutig (es gibt keine Doppelung) und trotzdem so kurz wie möglich um die Eingabe der Diagnosen zu beschleunigen.

	Mit Sternchen zum ICD-Thesaurus:
	▪ 	Nach Eingabe eines Sterns (*) erscheint ein Dialog zur Suche einer Diagnose im aktuellen ICD-XX-GM-Katalog (German Modifaction).
	▪ 	Im Gegensatz zum vergleichbaren Dialog von Albis können mehrere Diagnosen auf einmal übertragen werden. Die Suche ist wesentlich
		schneller als man es von Albis gewohnt ist, kann also inkrementell erfolgen (Liste wird nach jeder Eingabe eines Buchstabens aktualisiert).
		Die Anzeige ist in zwei Bereiche unterteilt. Im oberen Bereich werden die passenden Hotstring-Diagnosen angezeigt, darunter die
		Diagnosen aus dem offiziellen ICD-GM-Katalog.

	Konzept/Design:
	▪	Da es trotz ihrer Häufigkeit immer noch viele Diagnosen sind, werden verschiedene Schreibweisen für ein und denselben Hotstring verwendet.
		Dies soll ein auswendig lernen von Abkürzungen überflüssig machen und ist das grundlegende Konzept hinter dieser Hotstringsammlung.
		Die von Albis über die Eingabe von Shift+F3 erreichbare 'kleine Diagnoseliste' nötigt ein auswendig lernen ab und kann aufgrund des starren
		Konzeptes nicht in auf unterschiedlichen Vorlieben und Fähigkeiten von Personen eingehen.


*/

:X*:.grdia::SendDiagnosis("Infektausschluß {J06.9A}; Impfung gegen Grippe, " ImpfJahr "/" ImpfJahr+1 " (Influvac Tetra Ch: X#) {Z25.1G}")    ; Influvac Preis
:C*:U07.1:: ;{
:C*:U07.2::

	OAppClass := "ahk_class OptoAppClass"

	RegExMatch(A_ThisHotkey, "\d$", HotSEnd)
	SendDiagnosis("COVID-19, Virus nachgewiesen {!U07." HotSEnd "V};")   ; Schnellschussdiagnose
	Sleep 300

	SendInput, {TAB}
	ControlGetFocus, cfocus, % OAppClass
	while (cfocus~="RichEdit" && A_Index < 30) {  ; 1.5s
		Sleep 50
		If (hwnd := WinExist("ahk_class #32770", "meldepflichtig"))
			If !VerifiedClick("OK", hwnd,,, true)
				SendInputF("ahk_id " hwnd, "{Escape}", 200)
		ControlGetFocus, cfocus, % OAppClass
	}

	Sleep 100
	If (hwnd := WinExist("ahk_class #32770", "meldepflichtig"))
		If !VerifiedClick("OK", hwnd,,, true)
			SendInputF("ahk_id " hwnd, "{Escape}", 200)

	SendInputF(OAppClass, "{Escape}", "focuschange")
	SendInputF(OAppClass, "{Escape}", 200)

	SendInputF(OAppClass, "{Insert}", "focuschange")
	Sleep 200

	SendRaw, % "lko"
	Sleep 100

	SendInputF(OAppClass, "{TAB}", "focuschange")
	Sleep 400

	Send, % "{Empty}"
	Send, % " "
	Sleep 100

	ControlGetFocus, cfocus, % OAClass
	ControlGetText, cfText, % cfocus, % OAppClass
	If !cText || !RegExMatch(cfText, "(^|\D)88240(\D|$)") {
		If cText {
			SendInput, % "{LControl Down}{a}{LControl Up}"
			Sleep 200
			SendInput, % "{Backspace}"
			Sleep 100
			SendInput, % "{Delete}"
			Sleep 100
		}
		If !VerifiedSetText(cfocus, "88240", OAppClass) {
			VerifiedSetText(cfocus, "88240", OAppClass)
			ControlGetText, cfText, % cfocus, % OAppClass
			If (cfText!="88240")
				VerifiedSetText(cfocus, "88240", OAppClass)
		}
		SendInputF(OAppClass, "{TAB}", "focuschange")
		Sleep 2ßß
	}

	SendInputF(OAppClass, "{Escape}", "focuschange")
	SendInputF(OAppClass, "{Escape}")

return ;}
SendInputF(win, keys, wait:=0) {
	ControlGetFocus, cfocus, % win
	ControlSend, % cFocus, % keys, % win
	EL := ErrorLevel
	If (wait~="(\d+)") {
		Sleep % wait
	}
	else if (wait="focuschange"){
		lFocus := cFocus
		while (lFocus = cFocus && A_Index < 30) {
			Sleep 50
			ControlGetFocus, cfocus, % win
		}
	}
return EL
}

; Allergien                                          	;{
:X*:.Hausst::SendDiagnosis("Hausstauballergie {T78.4#}")
:X*:.Kontaktder::                                                             	;{ [Auswahlbox] Kontaktdermatitis
:X*:.Kontaktex::
:X*:.Kontaktall::
Auswahlbox(["title=Kontaktexanthem/-dermatitis border=on maxRows=30"
				, 	"Nickelallergie {L23.0}"                                                               	, "Chromallergie {L23.0}"
				,  	"Allergische Kontaktdermatitis durch Metalle {L23.0}"  		    		, "Allergische Kontaktdermatitis durch Nickel {L23.0}"
				, 	"Allergische Kontaktdermatitis durch Bichromat {L23.0}"					, "Allergische Kontaktdermatitis durch Chrom {L23.0}"
				,  	"Allergische Kontaktdermatitis durch Klebstoff {L23.1}"            		, "Heftpflasterekzem {L23.1}"
				, 	"Allergische Kontaktdermatitis durch Heftpflaster {L23.1}"       		, "Allergische Kontaktdermatitis durch Zugpflaster {L23.1}"
				,  	"Allergische Kontaktdermatitis durch Katharidenpflaster {L23.1}"
				, 	"Allergie gegen Kosmetika {L23.2}"
				,  	"Allergische Kontaktdermatitis durch Kosmetika {L23.2}"         		, "Allergische Kontaktdermatitis durch Kölnisch Wasser {L23.2}"
				,  	"Allergische Dermatitis durch Haarfärbemittel {L23.2}"             		, "Allergische Kontaktdermatitis durch Arzneimittel {L23.3}"
				,  	"Allergisches Kontaktexanthem durch Arzneimittel {L23.3}"
				, 	"Allergische Kontaktdermatitis durch Farbstoff {L23.4}"
				,  	"Gummiallergie {L23.5}", "Allergie gegen Pflaster {L23.5}"          	, "Pflasterallergie {L23.5}", "Insektizidallergie {L23.5}"
				, 	"Allergische Dermatitis durch Terpentine {L23.5}"                         	, "Konservierungsmittelallergie {L23.5}", "Kunststoffallergie {L23.5}"
				,  	"Zementallergie {L23.5}", "Chemikalienallergie {L23.5}"             	, "Detergenzienallergie {L23.5}"	, "Waschmittelallergie {L23.5}"
				,	"Allergie gegen organisches Lösungsmittel {L23.5}"                       	, "Allergische Kontaktdermatitis durch Konservierungsstoffe {L23.5}"
				,  	"Allergische Dermatitis durch chemische Produkte {L23.5}"       		, "Allergische Kontaktdermatitis durch chemische Produkte {L23.5}"
				,  	"Allergische Kontaktdermatitis durch Kunststoff {L23.5}"          		, "Allergische Kontaktdermatitis durch Insektizide {L23.5}"
				,  	"Allergische Kontaktdermatitis durch Gummi {L23.5}"             		, "Allergische Kontaktdermatitis durch Zement {L23.5}"
				,  	"Allergische Kontaktdermatitis durch Plastik {L23.5}"						, "Allergische Kontaktdermatitis durch Nylon {L23.5}"
				,  	"Allergie durch industrielle Fette {L23.5}"										, "Seifenallergie {L23.5}"
				, 	"Allergische Kontaktdermatitis durch Hautkontakt mit Nahrungsmitteln {L23.6}"
				,  	"Allergische Kontaktdermatitis durch Mehl {L23.6}"							, "Allergisches Bäckerekzem {L23.6}"
				,  	"Allergische Kontaktdermatitis durch Fisch {L23.6}"							, "Allergische Kontaktdermatitis durch Obst {L23.6}"
				,  	"Allergische Kontaktdermatitis durch Milch {L23.6}"						, "Allergische Kontaktdermatitis durch Gemüse {L23.6}"
				,  	"Allergische Kontaktdermatitis durch Fleisch {L23.6}"
				, 	"Dermatitis venenata {L23.7}"
				,  	"Allergische Kontaktdermatitis durch Schlüsselblumen {L23.7}"		, "Allergische Kontaktdermatitis durch Primeln {L23.7}"
				,  	"Allergische Kontaktdermatitis durch Jakobskreuzkraut {L23.7}"		, "Allergische Kontaktdermatitis durch Gräser {L23.7}"
				,  	"Allergische Kontaktdermatitis durch Brennnesseln {L23.7}"				, "Allergische Kontaktdermatitis durch Ambrosiagewächse {L23.7}"
				,  	"Allergische Kontaktdermatitis durch Giftsumach {L23.7}"				, "Allergische Kontaktdermatitis durch Gifteiche {L23.7}"
				,  	"Allergische Kontaktdermatitis durch Giftefeu {L23.7}"					, "Wiesengräser-Dermatitis {L23.7}"
				,  	"Rhus-Dermatitis {L23.7}", "Pflanzenallergie {L23.7}"                  	, "Dermatitis pratensis {L23.7}"
				,  	"Allergische Dermatitis durch äußeren Reizstoff a.n.k. {L23.8}"    	, "Allergische Kontaktdermatitis durch Pelz {L23.8}"
				, 	"Hautallergie durch Kontakt {L23.9}"
				, 	"Allergisches Ekzem {L23.9}", "Kontaktallergie {L23.9}"	        		, "Hautallergie {L23.9}", "Allergische Kontaktdermatitis {L23.9}"])
return ;}
;}

; Augenheilkunde/Opthalmologie      	;{
:X*:.akute Konj::SendDiagnosis("Akute Konjunktivitis, nicht näher bezeichnet {H10.3#G}")
:X*:.Katarakt::SendDiagnosis("Katarakt bds. {H25.9BG}")
:X*:.Makulade::SendDiagnosis("Degeneration der Makula und des hinteren Poles {H35.39#G}")
:X*:.Glaukom W::                                                              ;{ Glaukom
:X*:.Weitwi::SendDiagnosis("Primäres Weitwinkelglaukom {H40.1#G}") ;}
:X*:.Glaukom nnbz::SendDiagnosis("Glaukom, n.n.bz. {H40.9G}")
;}

; Chirurgie                                         	;{
:X*:.Bride::                                                                     	;{ peritoneale Adhäsionen
:X*:.Adhä::SendDiagnosis("Peritoneale Adhäsionen {K66.0G}") ;}
:X*:.Aortenek::                                                               	;{ Quardurchmesser der Aorta <30mm
SendDiagnosis("Aortenektasie {I71.9G}")
return ;}
:X*:.Aneurysma Aorta ab::                                                	;{ abd. Aortenaneurysma
:X*:.abdominelles Ao::
:X*:.Abdomenane::
:X*:.Bauchao::SendDiagnosis("Aneurysma der Aorta abdominalis ohne Ruptur {I71.4G}") ;}
:X*:.Atheros::                                                                   	;{ Becken-Bein-Arterossklerose [Auswahlbox]
Auswahlbox(["title=Arterossklerose der Beingefäße border=on "
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ, ohne Beschwerden {I70.20#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ, Stadium I nach Fontaine {I70.20#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ mit einer Gehstrecke von 200 m und mehr {I70.21#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ, Stadium IIa nach Fontaine {I70.21#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ mit einer Gehstrecke weniger als 200 m {I70.22#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ, Stadium IIb nach Fontaine {I70.22#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ mit Ruheschmerzen {I70.23#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ, Stadium III nach Fontaine {I70.23#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ mit Ulzeration {I70.24#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ, Stadium IV nach Fontaine mit Ulzeration {I70.24#G}"
				,	"Atherosklerose der Extremitätenarterien vom Becken-Bein-Typ mit Gangrän {I70.25#G}"])
return ;}
:X*:.CHE::                                        	                            	;{ Cholezystolithiasis/Cholezystektomie
:X*:.Cholezyste::SendDiagnosis("Cholezystektomie ## {K80.00Z}")
:X*:.Gallenblasenst::
:X*:.Cholezysto::SendDiagnosis("Cholezystolithiasis {K80.20G}") ;}
:X*:.Erysi::SendDiagnosis("Erysipel {A46G}")
:X*:.Meningeom::                                                            	;{ Meningeome / andere gutartige Neubildungen der Hirnhaut
Auswahlbox(["title=Meningeome und andere gutartige Neubildungen der Hirnhaut border=on"
				, 	"Gutartige Neubildung der Arachnoidea cerebralis {D32.0}"	,	"Gutartige Neubildung der Dura mater cranialis {D32.0}"
				,	"Gutartige Neubildung der Dura mater encephali {D32.0}" 	,	"Gutartige Neubildung der Falx cerebelli {D32.0}"
				,	"Gutartige Neubildung der Falx cerebri {D32.0}"		     		,	"Gutartige Neubildung der Gehirnmeningen {D32.0}"
				,	"Gutartige Neubildung der Hirnhaut {D32.0}"		         		,	"Gutartige Neubildung der intrakraniellen Meningen {D32.0}"
				,	"Gutartige Neubildung der kranialen Pia mater {D32.0}"		,	"Gutartige Neubildung der kraniellen Meningen {D32.0}"
				,	"Gutartige Neubildung der zerebralen Meningen {D32.0}"		,	"Gutartige Neubildung der zerebralen Pia mater {D32.0}"
				,	"Gutartige Neubildung des Tentorium cerebelli {D32.0}"			,	"Intrakranielles Meningeom {D32.0}"
				,	"Klivusmeningeom {D32.0}"				,	"Meningeom des Frontallappens {D32.0}"				,	"Meningeom des Nervus opticus {D32.0}"
				,	"Optikusmeningeom {D32.0}"			,	"Orbitameningeom {D32.0}"			                    	,	"Sehnervmeningeom {D32.0}"
				,	"Tentoriummeningeom {D32.0}"                                    		,	"Zerebrales Meningeomutartige Neubildung der Arachnoidea {D32.9}"
				,	"Gutartige Neubildung der Dura mater {D32.9}"		    		,	"Gutartige Neubildung der Meningen {D32.9}"
				,	"Gutartige Neubildung der Pia mater {D32.9}"		        		,	"Hämangioblastisches Meningeom {D32.9}"
				,	"Hämangioperizytöses Meningeom {D32.9}"	            			,	"Hämangioperizytom der Meningen {D32.9}"
				,	"Leptomeningeom {D32.9}"		   		,	"Leptomeningiom {D32.9}"	                    			,	"Meningenangiom {D32.9}"
				,	"Meningeom {D32.9}"		        		,	"Meningiom - s.a. Meningeom {D32.9}"				,	"Meningotheliomatöses Meningeom {D32.9}"
				,	"Psammöses Meningeom {D32.9}"	,	"Psammom {D32.9}"		                                		,	"Subdurale gutartige Neubildung {D32.9}"
				,	"Synzytiales Meningeom {D32.9}"		,	"Transitionalzelliges Meningeom {D32.9}"])
return ;}
:X*:.Gangrän Fu::SendDiagnosis("Gangrän des Fußes und der Zehen {R02.07#G}")
:X*:.hiatus h::                                                                 	;{ Hiatushernie
:X*:.zwerchf::
:X*:.Hernia d::
:X*:.Herniad::SendDiagnosis("Hernia diaphragmatica {K44.9G}") ;}
:X*:.Leisten::SendDiagnosis("(Einseitige) Hernia inguinalis ohne Einklemmung u. Gangrän, kein Rezidiv {K40.90#}")
:X*:.Kopfplatz::                                                              	;{ Platzwunden
:X*:.KPW::SendDiagnosis("Kopfplatzwunde # {S01.9G}")
:X*:.Lippenplatz::
:X*:.Platzwunde Lippe::
:X*:.Lippenplatz::SendDiagnosis("Lippenplatzwunde {S01.51}") 	;}
:X*:Pilonidal::SendDiagnosis("Pilonidalzyste mit Abszess {L05.0#}")
:X*:.Prellung::                                                                 	;{ Prellungen
Auswahlbox(["title=Prellungen diverser Körperteile  border=on"
	, "Thoraxprellung {S20.2G}"	    	, "Prellung der behaarten Kopfhaut {S00.05}"	, "Prellung der Regio parietalis {S00.05}"   	, "Schädelprellung {S00.05}"
	, "Augenbrauenprellung {S00.1}"	, "Augenlidprellung {S00.1}"	, "Nasenbeinprellung {S00.35}"	, "Gesichtsprellung {S00.85}", "Schläfenprellung {S00.85}"
	, "Stirnprellung {S00.85}"          	, "Larynxprellung {S10.0}" 	, "Pharynxprellung {S10.0}"    	, "Rückenprellung {T09.05}"	, "HWS-Prellung {S20.2}"
	, "BWS-Prellung {S20.2}"           	, "LWS-Prellung {S30.0}"    	, "Thoraxprellung {S20.2}"     	, "Brustbeinprellung {S20.2}"	, "Mammaprellung {S20.0}"
	, "Beckenprellung {S30.0}"        	, "Gesäßprellung {S30.0}" 	, "Prellung der Kreuzbeingegend {S30.0}"                          	, "Steißbeinprellung {S30.0}"
	, "Bauchdeckenprellung {S30.1}"
	, "Nierenprellung {S37.01}"        	, "Leberprellung {S36.11}" 	, "Flankenprellung {S30.1}"    	, "Hodenprellung {S30.2}" 	, "Penisprellung {S30.2}"
	, "Oberarmprellung {S40.0}"     	, "Schulterprellung {S40.0}"	, "Ellenbogenprellung {S50.0}"	, "Unterarmprellung {S50.1}"	, "Handgelenkprellung {S60.2}"
	, "Handprellung {S60.2}"            	, "Fingerprellung {S60.0}" 	, "Fingerprellung mit Nagelschädigung {S60.1}"	, "Daumenprellung {S60.0}"
	, "Daumenprellung mit Nagelschädigung {S60.1}"                 	, "Multiple Prellungen an Schulter und Oberarm {S40.7}"
	, "Multiple Prellungen der Regio scapularis {S40.7}"                	, "Hüftprellung {S70.0}"	, "Oberschenkelprellung {S70.1}"
	, "Prellung des Musculus quadriceps femoris {S70.1}"             	, "Kniegelenkprellung {S80.0}"	, "Knieprellung {S80.0}"    	,"Unterschenkelprellung {S80.1}"
	, "Prellung des oberen Sprunggelenks {S90.0}"          	, "Prellung der Kleinzehe {S90.1}"	, "Zehenprellung {S90.1}"	, "Zehenprellung mit Nagelschädigung {S90.2}"
	, "Fersenprellung {S90.3}"	, "Fußprellung {S90.3}" 	, "Fußrückenprellung {S90.3}"   	, "Vorfußprellung {S90.3}"	, "Multiple Unterschenkelprellungen {S80.7}"
	, "Multiple Prellungen von Knöchel und Fuß {S90.7}"	, "Multiple Prellungen {T00.9}"   	, "Muskelprellung {T14.6}"])
return ;}
:X*:.pavk::	                                                                  	;{ periphere Gefäßkrankheit
:X*:.arterielle Ver::SendDiagnosis("Periphere Gefäßkrankheit {I73.9#G}")
:X*:.periphere arterielle Ver::SendDiagnosis("Periphere Gefäßkrankheit {I73.9#G}") ;}
:X*:.Sigmad::SendDiagnosis("Sigmadivertikulitis {K57.32#}")
:X*:.Ungius in::                                                               	;{ Unguis incarnatus
:X*:.Onychocryp::SendDiagnosis("Unguis incarnatus #. Finger ## {L60.0#G}") ;}
:X*:.Pneumot::SendDiagnosis("Pneumothorax {J93.9#}")  	;{ Pneumothorax
:X*:.Spannungsp::SendDiagnosis("Spontaner Spannungspneumothorax {J93.0#}") ;}
:X*:.Varizen::                                                                   	;{ Varikosis [Auswahlbox]
Auswahlbox(["title=Varikosis border=on"
           		,	"Varizen der unteren Extremitäten mit Ulzeration {I83.0#G}"
           		,	"Varizen der unteren Extremitäten mit Entzündung {I83.1#G}"
           		,	"Varizen der unteren Extremitäten mit Ulzeration und Entzündung {I83.2#G}"
           		,	"Varizen der unteren Extremitäten ohne Ulzeration oder Entzündung {I83.9#G}"
           		, 	"ausgeprägte Stammvarikosis d. V.saphena magna {I83.9#G}"])
return ;}
:X*:.sekundäre Wundheil::                                                	;{ sekundäre Wundheilungsstörung
:X*:.verzögerte Wundheil::SendDiagnosis("Sekundäre Wundheilung {T89.03G}")
:X*:.Wundheil::SendDiagnosis("Postoperative infektiöse Wundheilungsstörung {T81.4G}")
:X*:.Wundinf::SendDiagnosis("Wundinfektion nach einem Eingriff {T81.4G}") ;}
;}

; Diabetologie                                   	;{
:X*:.Diabetes neur::                                                           	;{ Diabetes mellitus Typ 2 mit neurologischen Komplikationen
:X*:.diabetische Poly::SendDiagnosis("Diabetes mellitus Typ 2 mit neurologischen Komplikationen {+E11.40G}; Diabetische Polyneuropathie {*G63.2BG}") ;}
:X*:.Dm2::                                                                          	;{ Diabetes mellitus Typ 2 E11.90
:X*:.Dm 2::
:X*:.Diab2::
:X*:.Diabetes 2::
:X*:.Diabetes Typ2::
:X*:.Diabetes Typ 2::SendDiagnosis("Diabetes mellitus Typ 2 {E11.90G}") ;}
:X*:.Hypoglyk::SendDiagnosis("Hypoglykämie {E16.2G}")
:X*:.Hyperglyk::SendDiagnosis("Hyperglykämie {R73.9}")
:X*:.Retinop::SendDiagnosis("Retinopathia diabetica (E10-E14+, vierte Stelle .3) {*H36.0G}")
:X*:.Diabetes F::                                                                	;{ [Auswahlbox] + ICD diabetischer Fuß + den Zustand beschreibende Diagnosen
:X*:.diabetisches F::
:X*:.diabetischer F::
:X*:.Fußsyn::
	Auswahlbox(["title=ICD diabetisches Fußsyndrom E10.7X, E11.7X,L89.X,M,Z,H,N,I border=on  maxRows=24"
						, "Typ-2-Diabetes mit diabetischem Fußsyndrom, nicht entgleist {E11.74G}"
						, "Typ-1-Diabetes mellitus mit diabetischem Fußsyndrom, nicht entgleist {E10.74G}"
						, "Typ-2-Diabetes mit diabetischem Fußsyndrom, entgleist {E11.75G}"
						, "Typ-1-Diabetes mellitus mit diabetischem Fußsyndrom, entgleist {E10.75G}"
						, "Diabetisches Ulcus, Wagner 0 {L89.08#G}", "Diabetisches Ulcus, Wagner 1 {L89.18#G}"
						, "Diabetisches Ulcus, Wagner 2 {L89.28#G}", "Diabetisches Ulcus, Wagner 3 {L89.38#G}"
						, "Diabetisches Fersenulcus, Wagner 0 {L89.07#G}", "Diabetisches Fersenulcus, Wagner 1 {L89.17#G}"
						, "Diabetisches Fersenulcus, Wagner 2 {L89.2#G}", "Diabetisches Fersenulcus, Wagner 3 {L89.37#G}"
						, "Fußdeformität {M21.6#G}", "Krallenzehen (erworben) {M20.5#G}", "Hallux valgus {M20.1#G}", "Charcot‘sche Osteoarthropathie {M14.6#G}"
						, "Zustand nach Zehenamputation {Z89.4}", "Zustand nach US-Amputation {Z89.5}", "Zustand nach OS-Amputationen {Z89.6G}"
						, "Rollstuhlpflichtigkeit {Z99.3G}", "MRSA-Infektion {U80.0G}"
						, "Retinopathia diabetica {H36.0*}", "Diabetische Nephropathie {N08.3*}", "Periphere diabetische Angiopathie {I79.2*G}"
						, "Diabetische Polyneuropathie {*G63.2BG}"])
return ;}
:X*:.Diabetes K::                                                                	;{ [Auswahlbox] + ICD Diabetes mit Komplikationen
:X*:.Komplikation Diabetes::
:X*:.diabetische Nephrop::
:X*:.diabetische Nieren::
:X*:.diabetische Angio::
:X*:.Retinopathia dia::
	Auswahlbox(["title=ICD Diabetes mit Komplikationen E1X border=on  BgColor1=FFDEAD BgColor2=FFF2DB "
						, "Typ-2-Diabetes mit diabetischem Fußsyndrom, nicht entgleist {E11.74G}", "Typ-1-Diabetes mellitus mit diabetischem Fußsyndrom, nicht entgleist {E10.74G}"
						, "Typ-2-Diabetes mit diabetischem Fußsyndrom, entgleist {E11.75G}", "Typ-1-Diabetes mellitus mit diabetischem Fußsyndrom, entgleist {E10.75G}"
						, "Typ-2-Diabetes mit Nierenkomplikationen, nicht entgleist [N08.3*] {+E11.20}", "Typ-1-Diabetes mit Nierenkomplikationen, nicht entgleist [N08.3*] {+E10.20}"
						, "Typ-2-Diabetes mit Nierenkomplikationen, entgleist [N08.3*] {+E11.21G}", "Typ-1-Diabetes mit Nierenkomplikationen, entgleist [N08.3*] {+E10.21G}"
						, "Retinopathia diabetica {H36.0*}", "Diabetische Nephropathie {N08.3*}", "Periphere diabetische Angiopathie {I79.2*}","Diabetische Polyneuropathie {*G63.2BG}"])
return ;}
;}

; Dermatologie                                  	;{
:X*:.Aktin::SendDiagnosis("Aktinische Keratose {L57.0G}")
:X*:.Basaliom::                                                                  	;{ Basaliom
:X*:.Basalzell::
Auswahlbox(["title=Basalzellkarzinom border=on"
				,	"Basaliom, Gesicht {C44.3G}"	, "Basaliom, behaarte Kopfhaut {C44.4G}", "Basaliom, Hals {C44.4G}", "Basaliom, Rumpf {C44.5G}"
				,	"Basaliom, Arm {C44.6G}"    	, "Basaliom, Bein {C44.7G}"     				, "Basaliom, ___ {C44.9G}"])
return ;}
:X*:.Dermatofibrom::SendDiagnosis("Dermatofibrom {D23.6G}")
:X*:.Rosa::SendDiagnosis("Rosacea {L71.9G}")
:X*:.Psori::                                                                        	;{ Psoriasis vulgaris
:X*:.Schup::SendDiagnosis("Psoriasis vulgaris {L40.0G}") ;}
:X*:.Hidra::                                                                       	;{ Akne inversa
:X*:.Akne in::
:X*:.Acne in::SendDiagnosis("Acne inversa {L73.2G}") ;}
;}

; Gastroenterologie/Hepatologie        	;{
:X*:.Barret::SendDiagnosis("Barrett-Ösophagus {K22.7G}")
:XC*:.COUL::                                                                      	;{ Colitis ukcerosa
:X*:.Colitis u::SendDiagnosis("Colitis ulcerosa {K51.9G}") ;}
:X*:.Darmb::                                                                      	;{ gastrointestinale Blutung
:X*:.Blutung Gast::
:X*:.Gastrointestinale B::
:X*:.GI-Blutung::
:X*:.GI Blutung::SendDiagnosis("Gastrointestinale Blutung {K92.2#}") ;}
:X*:.Melä::                                                                        	;{ Teerstuhl
:X*:.Teer::SendDiagnosis("Meläna {K92.1G}") ;}
:X*:.Hämat::SendDiagnosis("Hämatemesis {K92.0G}")
:X*:.Hepatomt::SendDiagnosis("Hepatomegalie {R16.0G}")
:XC*:.HepA::                                                                       	;{ akute Hepatitis A
:X*:.Hepatitis A::SendDiagnosis("Virushepatitis A ohne Coma hepaticum {B15.9G}") ;}
:XC*:.HepB::                                                                       	;{ akute Hepatitis B
:X*:.Hepatitis B::SendDiagnosis("Akute Hepatitis B {B16.9G}") ;}
:X*:.Hiatush::SendDiagnosis("Hiatushernie {K22.9G}")
:X*:.Jejunumdiv::SendDiagnosis("Jejunumdivertikulose {K57.10G}")
:X*:.Kolonpo::                                                                     	;{ Kolonpolyp
:X*:.Dickdarmpo::
:X*:.Colonpo::SendDiagnosis("Polyp des Kolons {K63.5#}") ;}
:X*:.Lakto::SendDiagnosis("Laktoseintoleranz {E73.9#}")
:X*:.Leberhäm::SendDiagnosis("Leberhämangiom {D18.03#}")
:X*:.Leberci::                                                                       	;{ Leberzirrhose Auswahlbox
:X*:.Leberzi::
:X*:.alkoholische L::
:X*:.ethyltoxische L::
:X*:.C2-bedingte L::
:X*:.C2-toxische L::
:X*:.toxische L::
:X*:.Child-P::
:X*:.Child P::
Auswahlbox(["title=Leberzirrhose toxisch/nicht toxisch border=on maxRows=12"
				, 	"Chronisches alkoholisches Leberversagen {K70.41G}"                            	, "dekompensierte Leberzirrhose {K70.41G}"
				,	"Alkoholische Leberzirrhose mit Ösophagusvarizen (I98.2*) {K70.3G}; Ösophagus- und Magenvarizen b.ao.k.K {*I98.2G}"
				,	"C2-toxische Leberkrankheit {K70.9G}"   	                                            	, "Aszites {R18G}"
				,	"Alkoholische Laënnec-Leberzirrhose {K70.3}"	                                     	, "Alkoholtoxische Leberzirrhose {K70.3}"
				, 	"Alkoholische Leberzirrhose mit Magenvarizen (I98.2*) {+K70.3}"           	, "Alkoholische Leberzirrhose mit Ösophagusvarizen (I98.2*) {+K70.3}"
				,	"Alkoholische Leberzirrhose mit Magenvarizenblutung (I98.3*) {+K70.3}"	, "Alkoholische Leberzirrhose mit Ösophagusvarizenblutung (I98.3*) {+K70.3}"
				, 	"Medikamentös-toxische Leberzirrhose {K71.7}"                                        	, "Toxische Leberzirrhose {K71.7}"
				,	"Toxische Leberzirrhose mit Ösophagusvarizen (I98.2*) {+K71.7}"            	, "Toxische Leberzirrhose mit Ösophagusvarizenblutung (I98.3*) {+K71.7}"
				,	"Autoimmune Leberzirrhose {K74.6}"                                                     	, "Dekompensierte Leberzirrhose {K74.6}"
				,	"Leberzirrhose, Stadium Child-Pugh A {K74.6}"                                        	, "Leberzirrhose, Stadium Child-Pugh B {K74.6}"
				,	"Leberzirrhose, Stadium Child-Pugh C {K74.6}"                                        	, "Nichtalkoholische Laënnec-Leberzirrhose {K74.6}"
				,	"Leberzirrhose mit Ösophagusvarizen (I98.2*) {+K74.6}"                        	, "Leberzirrhose mit Ösophagusvarizenblutung (I98.3*) {+K74.6}"
				,	"Kardiale Leberzirrhose {K76.1}"])
return ;}
:X*:.Aszites::SendDiagnosis("Aszites {R18G}")
:X*:.Ösophagusvar::                                                          	;{ Ösophagus- oder/und Magenvarizen
:X*:.Magenvar::SendDiagnosis("Alkoholische Leberzirrhose mit Ösophagusvarizen (I98.2*) {K70.3G}; Ösophagus- und Magenvarizen b.ao.k.K {*I98.2G}") ;}
:X*:.dekompensierte Leberz::                                               	;{ dekomp. Leberzirrhose
:X*:.Lebervers::
:X*:.Chronisches Leberversagen::SendDiagnosis("Chronisches alkoholisches Leberversagen {K70.41G}") ;}
:X*:.Mallory::SendDiagnosis("Mallory-Weiss-Syndrom {K22.6G}")
:X*:.portale Hyper::                                                           	;{ Pfortaderhypertonie
:X*:.Pfortaderhoch::
:X*:.Pfortaderhy::SendDiagnosis("Portale Hypertonie {K76.6G}") ;}
:X*:.Reflux::SendDiagnosis("Refluxoesophagitis mit Ösophagitis {K21.0G}")
:X*:.Reizd::SendDiagnosis("Reizdarmsyndrom {K58.2G}")
:X*:.Sod::SendDiagnosis("Sodbrennen {R12G}")
:X*:.Steatosis::                                                                     	;{ Steatosis hepatis
:X*:.Fettleber::SendDiagnosis("Steatosis hepatis {K76.0G}") ;}
;}

; Gynäkologie                                   	;{
:X*:.Dysme::SendDiagnosis("Dysmenorrhoe, nicht näher bezeichnet {N94.6G}")
:X*:.Salpingitis::SendDiagnosis("Akute Salpingitis {N70.0G}")
:X*:.Endometritis::SendDiagnosis("Akute Endometritis {N71.0G}")
:X*:.PCO:: ;{
:X*:.polyzystisches O::SendDiagnosis("PCO-[Polyzystisches-Ovar-] Syndrom {E28.2G};") ;}

;}

; HNO                                              	;{
:X*:.Epist::                                                                                 	;{ Epistaxis
:X*:.Nasenbl::SendDiagnosis("Epistaxis {R04.0#G}") ;}
:X*:.Lagerungss::                                                                       	;{ Lagerungsnystagmus
:X*:.paroxysmaler Schw::SendDiagnosis("Benigner paroxysmaler Schwindel {H81.1G}") ;}
:X*:.Hörsturz::SendDiagnosis("Idiopathischer Hörsturz {H91.2#}")
:X*:.Neuropathia:: ;{
:X*:.Vestib::SendDiagnosis("Neuropathia vestibularis Ausfall {H81.2#G}") ;}
:X*:.Tinit::SendDiagnosis("Tinnitus aurium {H93.1#G}")
:XC*:.OE:: ;{
:X*:.Otitis e::SendDiagnosis("Otitis externa {H60.9#G}") ;}
;}

; Infektionen                                      	;{
:X*:.Angina t::                                                                            	;{ Angina tonsillaris
:X*:.Tonsi::
:X*:.akute Ton::SendDiagnosis("Akute Tonsillitis {J03.9#G}") ;}
:X*:.Askar::                                                                               	;{ Ascaridose - Spulwürmer
:X*:.Ascar::
:X*:.Spulw::SendDiagnosis("Askaridose, nicht näher bezeichnet {B77.9G}") ;}
;B
:X*:.Borrel::                                                                                	;{ Borreliose
:X*:.Lyme::SendDiagnosis("Lyme-Krankheit {A69.2#}") ;}
:XC*:.BR::SendDiagnosis("Akute Bronchitis {J20.9G}")
;C
:X*:.cov19::                                                                                 	;{ Sars-CoV-19 Untersuchung + AU Diagnosen
:X*:.covid::
:X*:.SARS::
	; Zufallsdiagnose, da die U-ICDs laut Definition keine Krankheiten kodieren, die eine Arbeitsunfähigkeit verursachen.
	covidDiag := ["Akute Bronchitis {J20.9G}", "Virusinfektion {B34.9G}"]
	Random, dIndex, 1, 2
	SendDiagnosis("Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G};"
						. " Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G};"
						. " COVID-19, Virus nachgewiesen {!U07.1#};" covidDiag[dIndex])
return
;}
:X*:.covid19V::                                                                          	;{ COVID nachgewiesen V.a.
:X*:.covv::
:X*:.covidV::SendDiagnosis("COVID-19, Virus nachgewiesen {!U07.1V}") ;}
:X*:.covid19G::                                                                         	;{ COVID nachgewiesen G.
:X*:.covid19na::
:X*:.covg::
:X*:.covidG::SendDiagnosis("COVID-19, Virus nachgewiesen {!U07.1G}") ;}
:X*:.covid19A:: 	                                                                        	;{ COVID nicht nachgewiesen
:X*:.covid19ni::
:X*:.covn::
:X*:.covidni::SendDiagnosis("COVID-19, Virus nicht nachgewiesen {!U07.2G}") ;}
:X*:.multisys::SendDiagnosis("Multisystemisches Entzündungssyndrom in Verbindung mit COVID-19 {U10.9G}")
:X*:.Post-Covid::                                                                           	;{ Post-COVID
:X*:.Post Covid::
:X*:.PostCovid::SendDiagnosis("Post-COVID-19-Zustand {!U09.9G}") ;}
:RC*:.iC::Impfung gegen COVID-19 {U11.9G};
:X*:.Untersuchung C::                                                                  	;{ COVID Untersuchnungsdiangosen
:X*:.UCOVID::
:X*:.USARS::SendDiagnosis("Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G}"
									. 	 "Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G}") ;}
:X*:.ldsars::LDSars()
:X*:.corona::                                                                                	;{ [Auswahlbox] Covid-19 und Zustände
:X*:.covid3::
:X*:.SARS-C::
Auswahlbox(["title=SARS-CoV-2 border=on BgColor1=9FC4E8 BgColor2=D7E6F4"
				, 	"Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G}"
				,	"Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G}"
				,	"Kontakt mit und Exposition gegenüber COVID-19 {Z20.8G}"
				,	"COVID-19, Viruserkrankung {!U07.1V}"
				,	"COVID-19, Virus nachgewiesen {!U07.1G}"
				,	"COVID-19, Virus nicht nachgewiesen {!U07.2G}"
				,	"COVID-19 in der Eigenanamnese {U08.9G}"
				,	"Negativer COVID-19-Test {Z03.8}"
				,	"COVID-19-Infektion, klinisch-epidemiologisch bestätigt {Z20.8}"
				,	"COVID-19-Infektion, Virus nachgewiesen, ohne klinische Symptome {Z22.8}"
				,	"Post-COVID-19-Zustand {!U09.9G}"
				,	"Multisystemisches Entzündungssyndrom in Verbindung mit COVID-19 {U10.9G}"
				,	"Kawasaki-like-Syndrom zeitlich assoziiert mit COVID-19 {U10.9}"
				,	"PIMS zeitlich assoziiert mit COVID-19 {U10.9}"
				,	"Schweres akutes respiratorisches Syndrom [SARS] {U04.9#}"
				,	"Pneumonie durch SARS-CoV-2 {J12.8BG}"
				,	"Respiratorische Insuffizienz vom Typ I [hypoxisch] {J96.90G}"
				,	"Thrombozytopenie durch SARS-CoV-2 {D69.59G}"
				,	"Infektanämie {D64.9G}"
				,	"Myokarditis bei anderenorts klassifizierten Viruskrankheiten (+U07.1) {*I41.1}"
				, 	"Schwangerschaft mit COVID-19-Infektion {O98.5}"
				,	"Notwendigkeit der Impfung gegen COVID-19 {U11.9G}"
				,	"1. Impfung gegen COVID-19 {U11.9G}"
				,	"2. Impfung gegen COVID-19 {U11.9G}"
				,	"#. Impfung (Auffrischung) gegen COVID-19 {U11.9G}"
				,	"Unerwünschte Nebenwirkungen bei der Anwendung von COVID-19-Impfstoffen {!U12.9G}"])
return ;}
:X*:.covAU::                                                                               	;{ Corona AU begründende Diagnosen
:X*:.covidAU::
:X*:.covid-19 AU::
:X*:.covid-19AU::
:X*:.covid19 AU::
:X*:.covid19AU::
:X*:.covid AU::
:X*:.cov AU::SendDiagnosis("Virusinfektion, G. {B34.9G}; Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G};"
									. "Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G};"
									. "COVID-19, Virus nachgewiesen {!U07.1V}")  ;}
:XC*:.VI::SendDiagnosis("Virusinfektion, G. {B34.9G}")
;E
:X*:.EBV::                                                                                   	;{ EBV
:X*:.Ebstein::
:X*:.Mononu::
:X*:.Gamma-Her::
:X*:.Pfeiff::SendDiagnosis("Mononukleose durch Gamma-Herpesviren {B27.0G}")  ;}
:X*:.Enterob::                                                                               	;{ Fadenwürmer
:X*:.Fadenw::
:X*:.Madenw::
:X*:.Oxyu::SendDiagnosis("Enterobiasis [Oxyuriasis] {B80G}")  ;}
:X*:.Erysi::                                                                                 	;{ Erysipel
:X*:.Wundrose::SendDiagnosis("Erysipel [Wundrose]{A46#G}")    	;}
:X*:.Erythema::                                                                          	;{ Erythema migrans
:XC*:.ERMI::SendDiagnosis("Erythema migrans {A69.2G}") ;}
;G
:XC*:.GE::                                                                                    	;{ Gastroenteritis
:X*:.Gastroent::SendDiagnosis("Gastroenteritis {A09.0G}") ;}
:X*:.Gür::                                                                                  	;{ Gürtelrose ohne Komplikationen
:XC*:.HZ::
:X*:.Herpes Z::
:X*:.HerpesZ::
:X*:.Zoster::SendDiagnosis("Herpes zoster ohne Komplikation {B02.9#G}") ;}
:X*:.Grippe::                                                                            		;{ Grippe, Influenza
:X*:.Virusgr::
:X*:.Influe::SendDiagnosis("Virusgrippe  {J11.1#}") ;}
:X*:.Infektanä::SendDiagnosis("Infektanämie {D64.9#}")
:X*:.Neuralgie G::                                                                      	;{ Postzosterneuralgie
:X*:.Neuralgie Z::
:X*:.Postzos::SendDiagnosis("Neuralgie nach Herpes Zoster {*G53.0}; Herpes zoster mit Beteiligung anderer Abschnitte des Nervensystems {+B02.2G}") ;}
:X*:.Zoster En::SendDiagnosis("Zoster-Enzephalitis (G05.1*) {+B02.0}; Virusenzephalitis {*G05.1}")
;H
:X*:.HepE::                                                                                   	;{ akute Hepatitis E
:X*:.HepatitisE::
:X*:.VirushepE::
:X*:.VirushepatitisE::
:X*:.Virushepatitis E::
:X*:.akute HepatitisE::
:X*:.akute Hepatitis E::
:X*:.Hepatitis E::SendDiagnosis("Akute Virushepatitis E {B17.2#}") ;}
;K
:X*:.Keuch::                                                                               	;{ Pertussis
:X*:.Pertu::SendDiagnosis("Keuchhusten durch Bordetella pertussis {A37.0#}") ;}
;N
:X*:.Noro::SendDiagnosis("Akute Gastroenteritis durch Norovirus {A08.1G}")
;P
:X*:.Pedikul::                                                                               	;{ Pedikulose/Läuse
:X*:.Läuse::
:X*:.Pedicul::SendDiagnosis("Pedikulose {B85.2G}") ;}
:X*:.Phlegm::SendDiagnosis("Phlegmone an der unteren Extremität {L03.11G}")
;S
:X*:.Scharlach::                                                                           	;{ Scharlach
:X*:.Scarlatina::SendDiagnosis("Scharlach {A38G}") ;}
:X*:.Schnupf::SendDiagnosis("Erkältungsschnupfen {J00RG}")
:X*:.Sinusitis f::                                                                             	;{ Sinusitis frontalis
:X*:.Stirnhö::
:XC*:.SF::
:X*:.Akute Sinusitis f::SendDiagnosis("Akute Sinusitis frontalis {J01.1#G}") ;}
:X*:.Sinusitis m::                                                                         	;{ Sinusitis maxillaris
:X*:.Kieferhö::
:XC*:.SM::
:X*:.Akute Sinusitis m::SendDiagnosis("Akute Sinusitis maxillaris {J01.0#G}") ;}
;V
:XC*:.VI::                                                                                    	;{ Virusinfekt, allgemein
:X*:.Virusi::SendDiagnosis("Virusinfektion {B34.9G}") ;}
;}

; Kardiologie                                     	;{
:X*:.ACE::SendDiagnosis("ACE Hemmer Husten {T88.7G}")
:X*:.AICD::                                   	;{
:X*:.Defibr::
:X*:.Vorhandensein Defi::SendDiagnosis("Vorhandensein AICD {Z95.0G}") ;}
:X*:.Anasarka::                          	;{
:X*:.generalisiertes Ö::SendDiagnosis("Anasarka {R60.1G}") ;}
:X*:.Aortenklappenst::                   	;{
:XC*:.AS::
:X*:.Aortenste::
:X*:.Stenose Aorten::SendDiagnosis("Aortenklappenstenose {I35.0G}") ;}
:XC*:.Aortenklappenin::               	;{
:XC*:.AI::
:X*:.Insuffizienz Aorten::SendDiagnosis("Aortenklappeninsuffizienz {I35.1G}") ;}
:X*:.AV Bl::                                	;{
:X*:.AV-Bl::SendDiagnosis("Atrioventrikulärer Block #. Grades {I44.#G}") ;}
:X*:.AVB1::                                  	;{
:X*:.AVB2::
:X*:.AVB3::SendDiagnosis("Atrioventrikulärer Block $. Grades {I44.$G}") ;}
;B
:X*:.Bypass::                              	;{ Vorhandensein eines Bypass
:X*:.Vorhandensein By::SendDiagnosis("Vorhandensein eines #-Bypass {Z95.5G}") ;}
;F
:X*:.Foramen ovale::SendDiagnosis("Foramen ovale {Q21.1G}")
;H
:X*:.akzidentelles Herz::               	;{ akzidentelles Herzgeräusch
:X*:.Benignes Herz::
:X*:.Benigne Herz::SendDiagnosis("Benigne und akzidentelle Herzgeräusche {R01.0#}") ;}
:X*:.Herzinf::                              	;{ Herzinfarkt [Auswahlbox]
:X*:.Myokardinf::
Auswahlbox(["title=Herzinfarkt border=on BgColor1=FB6A6A BgColor2=FDBDBD"
				,	"Akuter Herzinfarkt {I21.9#}"
				,	"Alter Myokardinfarkt {I25.29#}"
				,	"Vorderwandinfarkt des Herzens {I21.0}"
				,	"Hinterwandinfarkt des Herzens {I21.1}"
				,	"Seitenwandinfarkt des Herzens {I21.2}"
				,	"Nicht-ST-Hebungsinfarkt {I21.4}"
				,	"Non-Q-Wave-Myokardinfarkt {I21.4}"
				,	"NSTEMI  {I21.4}"
				,	"Akuter transmuraler anteriorer Myokardinfarkt {I21.0}"
				,	"Akuter transmuraler anteroapikaler Myokardinfarkt {I21.0}"
				,	"Akuter transmuraler anterolateraler Myokardinfarkt {I21.0}"
				,	"Akuter transmuraler anteroseptaler Myokardinfarkt {I21.0}"
				,	"Akuter transmuraler Myokardinfarkt der Vorderwand {I21.0}"
				,	"Akuter Hinterwandinfarkt des Herzens {I21.1}"
				,	"Akuter HWI des Herzens {I21.1}"
				,	"Akuter Myokardinfarkt der Hinterwand {I21.1}"
				,	"Akuter transmuraler diaphragmaler Myokardinfarkt {I21.1}"
				,	"Akuter transmuraler inferiorer Myokardinfarkt {I21.1}"
				,	"Akuter transmuraler inferolateraler Myokardinfarkt {I21.1}"
				,	"Akuter transmuraler inferoposteriorer Myokardinfarkt {I21.1}"
				,	"Akuter transmuraler Myokardinfarkt der Hinterwand {I21.1}"
				,	"Akuter transmuraler apikolateraler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler basolateraler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler hochlateraler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler lateraler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler Myokardinfarkt der Seitenwand {I21.2}"
				,	"Akuter transmuraler posteriorer Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler posterobasaler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler posterolateraler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler posteroseptaler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler septaler Myokardinfarkt {I21.2}"
				,	"Akuter transmuraler Myokardinfarkt {I21.3}"
				,	"Transmuraler Myokardinfarkt {I21.3}"
				,	"Akute subendokardiale Nekrose {I21.4}"
				,	"Akuter nichttransmuraler Myokardinfarkt {I21.4}"
				,	"Akuter nichttraumatischer subendokardialer Infarkt {I21.4}"
				,	"Akuter subendokardialer Myokardinfarkt {I21.4}"
				,	"Infarkt der Myokardinnenschicht {I21.4}"
				,	"Nichttransmuraler Myokardinfarkt {I21.4}"])
return ;}
:X*:.akuter Herzin::                     	;{ akuter Herzinfarkt
:X*:.akuter Myokardin::SendDiagnosis("Akuter Herzinfarkt {I21.9#}") ;}
:X*:.alter Myoka::                        	;{ alter Herzinfarkt
:X*:.alter Herzin::SendDiagnosis("Alter Myokardinfarkt {I25.29G}") ;}
:X*:.Herzklappenin::                    	;{	Herzklappeninsuffizienzen Auswahlbox
Auswahlbox(["title=Herzklappeninsuffizienz border=on BgColor1=FB6A6A BgColor2=FDBDBD"
				,	"Aortenklappeninsuffizienz {I35.1G}"
				,	"Mitralinsuffizienz {I34.0G}"
				,	"Pulmonalklappeninsuffizienz {I37.1G}"
				,	"Trikuspidalklappeninsuffizienz {I07.1G}"
				,	"leichte Aortenklappeninsuffizienz {I35.1G}"
				,	"leicht- bis mittelgradige Aortenklappeninsuffizienz {I35.1G}"
				,	"mittelgradige Aortenklappeninsuffizienz {I35.1G}"
				,	"schwere Aortenklappeninsuffizienz {I35.1G}"
				,	"leichte Mitralinsuffizienz {I34.0G}"
				,	"leicht- bis mittelgradige Mitralinsuffizienz {I34.0G}"
				,	"mittelgradige Mitralinsuffizienz {I34.0G}"
				,	"schwere Mitralinsuffizienz {I34.0G}"
				,	"leichte Pulmonalklappeninsuffizienz {I37.1G}"
				,	"leicht- bis mittelgradige Pulmonalklappeninsuffizienz {I37.1G}"
				,	"mittelgradige Pulmonalklappeninsuffizienz {I37.1G}"
				,	"schwere Pulmonalklappeninsuffizienz {I37.1G}"
				,	"leichte Trikuspidalklappeninsuffizienz {I07.1G}"
				,	"leicht- bis mittelgradige Trikuspidalklappeninsuffizienz {I07.1G}"
				,	"mittelgradige Trikuspidalklappeninsuffizienz {I07.1G}"
				,	"schwere Trikuspidalklappeninsuffizienz {I07.1G}"])
return ;}
:X*:.Herzschr::SendDiagnosis("implantierter Herzschrittmacher {Z95.0G}")
:X*:.NYHA1::                            	;{ NYHA Stadien
:X*:.NYHA 1::
:X*:.NYHA I::
:X*:.NYHA2::
:X*:.NYHA 2::
:X*:.NYHA II::
:X*:.NYHA3::
:X*:.NYHA 3::
:X*:.NYHA III::
:X*:.NYHA4::
:X*:.NYHA 4::
:X*:.NYHA IV::SendDiagnosis("Linksherzinsuffizienz NYHA-Stadium $CMB {I50.1$G}") ;}
:X*:.Hypotonie::                          	;{ Hypotonie [Auswahlbox]
Auswahlbox(["title=Hypotonie border=on BgColor1=FB6A6A BgColor2=FDBDBD"
				,	"Orthostatische Hypotonie {I95.1#}"
				,	"Idiopathische Hypotonie {I95.0#}"
				,	"Hypotonie nach langem Stehen {I95.1#}"
				,	"Orthostatische Hypotonie {I95.1#}"
				,	"Arzneimittelinduzierte Hypotonie {I95.2#}"
				,	"Chronische Hypotonie {I95.8#}"
				,	"Hypotonie {I95.9#}"])
return ;}
:X*:.Hypertens::SendDiagnosis("Benigne essentielle Hypertonie mit hypertensiver Krise {I10.01G}")
;K
:X*:.KHK1::                                	;{ KHK-1
:X*:.KHK-1::
:X*:.KHK2::
:X*:.KHK-2::
:X*:.KHK3::
:X*:.KHK-3::SendDiagnosis("KHK-$ {I25.1$G}") ;}
:XC*:.KHKA::                             	;{ KHK G,A,V
:XC*:.KHKG::
:XC*:.KHKV::
:XC*:.KHKnnbz::
:XC*:.koronare H::SendDiagnosis("KHK {I25.19$#}") ;}
;L
:XC*:.LAHB::                                	;{ Linksanteriorer Hemiblock
:X*:.Linksan::SendDiagnosis("Linksanteriorer Hemiblock {I44.4G}") ;}
:XC*:.Linksschenkel::                 	;{ Linksschenkelblock
:XC*:.LSB::SendDiagnosis("Linksschenkelblock {I44.7G}") ;}
:X:.Linksherzi::                            	;{ Linksherzinsuffizienz
:XC*:.LHI::
Auswahlbox(["title=Linksherzinsuffizienz border=on BgColor1=FB6A6A BgColor2=FDBDBD"
					,	"Linksherzinsuffizienz NYHA-Stadium I {I50.11G}"
					,	"Linksherzinsuffizienz NYHA-Stadium II {I50.12G}"
					,	"Linksherzinsuffizienz NYHA-Stadium III {I50.13G}"
					,	"Linksherzinsuffizienz NYHA-Stadium IV {I50.14G}"
					,	"Linksherzinsuffizienz nnbz. {I50.19G}"])
return  ;}
:X*:.Lungenö::                            	;{ Lungenödem
:XC*:.LUÖD::
:X*:.Ödem L::SendDiagnosis("Lungenödem {J81#}") ;}
:X*:.Lungenst::SendDiagnosis("Lungenstauung {J81#")
;M
:X*:.Myokarditis a::                      	;{ akute Myokarditis
:X*:.AKMY::
:X*:.Akute Myok::SendDiagnosis("Akute Myokarditis {I40.9#}") ;}
:X*:.Myokarditis r::                     	;{ rezidivierende Myokarditis
:X*:.rheumatische Myok::SendDiagnosis("Rheumatische Myokarditis {I09.0#}") ;}
:X*:.Mitralklappenin::                  	;{ Mitralklappeninsuffizienz
:X*:.Mitrali::SendDiagnosis("Mitralinsuffizienz {I34.0G}") ;}
;O
:X*:.Ödem U::                             	;{ Unterschenkelödeme
:X*:.Unterschenkelö::SendDiagnosis("Unterschenkelödeme {R60.0#G}") ;}
;P
:XC*:.PFO::                                	;{ persistierendes Foramen ovale
:X*:.Foramen ovale::
:X*:.persistierendes F::SendDiagnosis("Persistierendes Foramen ovale {Q21.1#}") ;}
:X*:.Pulmonalin::                       	;{ Pulmonalklappeninsuffizienz
:X*:.Pulmonalisin::
:X*:.Pulmonalklappenin::SendDiagnosis("Pulmonalklappeninsuffizienz {I37.1#}") ;}
;R
:X*:.Rechtsh::                             	;{ Rechtsherzinsuffizienz
Auswahlbox(["title=Rechtsherzinsuffizienz border=on BgColor1=FB6A6A BgColor2=FDBDBD"
				,	"Primäre Rechtsherzinsuffizienz {I50.00}"
				,	"Rechtsherzinsuffizienz {!I50.01}"
				,	"sekundäre Rechtsherzinsuffizienz {!I50.01}"
				,	"Rechtsherzinsuffizienz ohne Beschwerden {!I50.02}"
				, 	"Rechtsherzinsuffizienz mit Beschwerden bei stärkerer Belastung {!I50.03}"
				,	"Rechtsherzinsuffizienz mit Beschwerden bei leichterer Belastung {!I50.04G}"])
return ;}
:XC*:.Rechtsschenkel::                  	;{ Rechtsschenkelblock
:XC*:.RSB::SendDiagnosis("Rechtsschenkelblock {I45.1G}") ;}
;S
:X*:.SA-Bl::                                	;{ SA-Block
:X*:.SABl::SendDiagnosis("intermittierende SA-Blöcke {I45.5G}") ;}
:X*:.Supraventrikuläre Ta::           	;{ Supraventrikuläre Tachykardie
:X*:.SVT::SendDiagnosis("Supraventrikuläre Tachykardie {I47.1G}") ;}
;T
:X*:.TAA::SendDiagnosis("Tachyarrhythmie bei Vorhofflimmern {I48.9}")
:X*:.Herzra::                              	;{ Tachykardie
:X*:.Tachyk::SendDiagnosis("Tachykardie {R00.0G}") ;}
:X*:.Trikuspidalklappenin::         	;{ Trikuspidalklappeninsuffizienz
:X*:.Trikuspidalisin::
:X*:.Trikuspidalin::SendDiagnosis("Trikuspidalklappeninsuffizienz {I07.1G}") ;}
;V
:XC*:.VES::                                 	;{ Ventrikuläre Extrasystolen
:XC*:.VESA::
:XC*:.VESG::
:XC*:.VESK::
:XC*:.VESZ::
:X*:.ventrikuläre Ex::SendDiagnosis("Ventrikuläre Extrasystolie {I49.3$#}") ;}
:XC*:.VT::                                  	;{ Ventrikuläre Tachykardie
:X*:.ventrikuläre Ta::SendDiagnosis("Ventrikuläre Tachykardie {I47.2G}") ;}
:X*:.VHF pa::                              	;{ Vorhofflimmern oder -flattern
:X*:.paroxysmal::
SendDiagnosis("Vorhofflimmern, paroxysmal {I48.0G}")
return
:X*:.VHF pers::
:X*:.persistierend::
SendDiagnosis("Vorhofflimmern, persistiernd {I48.1G}")
return
:X*:.VHF chr::
:X*:.VHF perm::
:X*:.permanent::
:X*:.chronisches Vor::
SendDiagnosis("Vorhofflimmern, permanentes {I48.2G}")
return
:X*:.VHF typ::
:X*:.typisches::SendDiagnosis("typisches Vorhofflattern {I48.3G}")
:X*:.VHF atyp::
:X*:.atypisches::SendDiagnosis("atypisches Vorhofflattern {I48.4G}") ;}

;W
:X*:.WPW::                                	;{ Wolff-Parkinson-White Syndrom
:X*:.WPWA::
:X*:.WPWG::
:X*:.WPWV::
:X*:.Wolf-Park::
:X*:.Wolff-Park::
:X*:.WPWZ::SendDiagnosis("Wolff-Parkinson-White-Syndrom {I45.6$#}") ;}
;}

; Morbus ...                                       	;{
:X*:.Based::               	;{ M. Basedow
:X*:.M. Ba::
:X*:.Morbus Ba::SendDiagnosis("M.Basedow {E05.0V}") ;}
:X*:.Behc::                  	;{ M. Behçet
:X*:.M. Be::
:X*:.Morbus Be::SendDiagnosis("Morbus Behçet {M35.2G}") ;}
:X*:.Ledder::               	;{ M.Ledderhose
:X*:.Fibromatose F::
:X*:.Fibromatose Pl::
:X*:.Fibromatose der Pl::
:X*:.M. Le::
:X*:.Morbus Le::SendDiagnosis("M.Ledderhose {M72.2#G}") ;}
:X*:.Dupu::                	;{ M.Dupuytren
:X*:.Fibromatose H::
:X*:.Fibromatose Pa::
:X*:.Fibromatose der Pa::
:X*:.M. Du::
:X*:.Morbus Du::SendDiagnosis("M.Dupuytren Kontraktur {M72.0#G}")  ;}
;}

; Nephrologie                                    	;{
:X*:.niereninsuffizienz a::   	    	;{ akute Niereninsuffizienz
:X*:.nierenkrankheit a::
:X*:.aNI::
:X*:.NIa::
:X*:.akute nieren::SendDiagnosis("Akute Nebenniereninsuffizienz {E27.2G}") ;}
:X*:.niereninsuffizienz c::            	;{ chronische Nierenkrankheit
:X*:.nierenkrankheit c::
:X*:.cNI::
:X*:.NIc::
:X*:.Chronische nieren::SendDiagnosis("Chronische Nierenkrankheit, Stadium X {N18.#G}")
:X*:.ni1::SendDiagnosis("Chronische Nierenkrankheit, Stadium 1 {N18.1G}")
:X*:.ni2::SendDiagnosis("Chronische Nierenkrankheit, Stadium 2 {N18.2G}")
:X*:.ni3::SendDiagnosis("Chronische Nierenkrankheit, Stadium 3 {N18.3G}")
:X*:.ni4::SendDiagnosis("Chronische Nierenkrankheit, Stadium 4 {N18.4G}") ;}


;}

; Neurologie                                     	;{
:X*:.Aphasie::SendDiagnosis("Aphasie {R47.0#}")
:X*:.Apraxie::SendDiagnosis("Apraxie {R48.2#}")
:X*:.Artikulations::SendDiagnosis("Artikulationsstörung {F80.0G}; Dysarthrie und Anarthrie {R47.1G}")
:X*:.Ataktisch:: ;{
:X*:.Ataxie::SendDiagnosis("Ataktischer Gang {R26.0G}") ;}
:X*:.Cluster::SendDiagnosis("Cluster-Kopfschmerz {G44.0G}")
:X*:.Critical-i::SendDiagnosis("Critical-illness-Polyneuropathie {G62.80G}; Critical-illness-Myopathie {G72.80G}")
:X*:.Dysar:: ;{
:X*:.Anar::SendDiagnosis("Dysarthrie und Anarthrie {R47.1G}") ;}
:X*:.Fatique::SendDiagnosis("Fatique Syndrom {G93.3G}")
:X*:.Fibularis:: ;{
:X*:.Fußhe::
:X*:.Peronaeus::SendDiagnosis("Läsion des Nervus fibularis communis {G57.3#G}") ;}
:X*:.Hemine::SendDiagnosis("Neurologischer Neglect {R29.5#G}")
:X*:.Folgen Apoplex::                          	;{
:X*:.Apoplex F::
:X*:.Folgen Schlaganfall::
:X*:.Schlaganfallf::
:X*:.Schlaganfall F::  ;Folgeerkrankungen
:X*:.Hemiparese::
Auswahlbox(["title=Folgen eines Schlaganfall border=on BgColor1=FB6A6A BgColor2=FDBDBD"
					,	"Hemiparese und Hemiplegie {G81.9#G}"
					,	"Schlaffe Hemiparese {G81.0#G}"
					,	"Linksherzinsuffizienz NYHA-Stadium III {I50.13G}"
					,	"Neurologischer Neglect {R29.5#G}"
					,	"Ataktischer Gang {R26.0G}"
					,	"Artikulationsstörung {F80.0G}"
					,	"Dysarthrie und Anarthrie {R47.1G}"])
return ;}
:X*:.Narkol::                                      	;{
:X*:.Katapl::
:X*:.Schlafkrankheit::SendDiagnosis("Narkolepsie und Kataplexie {G47.4G}") ;}
:X*:.Paräst::
:X*:.Paraes::SendDiagnosis("Parästhesie der Haut {R20.2G}")
:X*:.Pinealis::SendDiagnosis("Pinealiszyste {E34.8G}")
:XC*:.PNP:: 	                                    	;{ Polyneuropathien
:X*:.Polyneuro::Auswahlbox(["title=Polyneuropathien maxRows=14"
										, "Idiopathische progressive Polyneuropathie {G60.3}"
										, "Rezessiv vererbte sensible Neuropathie {G60.8}", "Dominant vererbte sensible Neuropathie {G60.8}"
										, "Hereditäre sensorische Neuropathie {G60.8}", "Hereditäre sensorische Polyneuropathie {G60.8}"
										, "Idiopathische Neuropathie {G60.9}", "Idiopathische periphere Neuropathie {G60.9}"
										, "Arzneimittelinduzierte Polyneuropathie {G62.0}"
										, "Alkoholpolyneuropathie {G62.1}"
										, "Toxische Neuropathie a.n.k. {G62.2}"
										, "Hereditäre Polyneuropathie {G60.9}"
										, "Neuropathische Beschwerden {G62.9}", "Polyneuropathie {G62.9}", "Akute multiple Neuropathien {G62.9}"
										, "Guillain-Barré-Syndrom {G61.0}", "Miller-Fisher-Syndrom {G61.0}"
										, "Serumpolyneuropathie {G61.1}"
										, "Polyradikulitis {G61.9}"
										, "Neuropathie durch Schnüffeln von Klebstoffen {G62.2}", "Bleipolyneuropathie {G62.2}"
										, "Critical-illness-Polyneuropathie {G62.80}"
										, "Motorische und sensorische Neuropathie {G62.88}", "Polyneuropathie durch Strahlung {G62.88}"
										, "Hereditäre sensomotorische Neuropathie, Typ I-IV {G60.0}", "Charcot-Marie-Tooth-Hoffmann-Syndrom {G60.0}"
										, "Refsum-Syndrom {G60.1}"
										, "Neuropathie mit hereditärer Ataxie {G60.2}", "Spinozerebelläre Ataxie mit axonaler Neuropathie Typ 1 {G60.2}"
										, "Morvan-Krankheit {G60.8}", "Nelaton-Syndrom {G60.8}"]) ;}
:X*:.Folgen Apo::                               	;{  Folgen Apoplex, Schlaganfall
:X*:.Folgen Schl::SendDiagnosis("Folgen eines Schlaganfalls {I69.4G}") ;}
:X*:.Tremor::SendDiagnosis("Essentieller Tremor {G25.0G}")
:X*:.TIA::                                            	;{
:X*:.transit::SendDiagnosis("TIA {G45.92#}") ;}
:X*:.Trigem::SendDiagnosis("Trigeminusneuralgie {G50.0G}")
:X*:.REM::SendDiagnosis("REM-Schlafstörung {G47.9G}")
:X*:.Restless::                                      	;{
:X*:.RLS::SendDiagnosis("Restless-Legs-Syndrom {G25.81G}") ;}
:X*:.Mediaversch::SendDiagnosis("Verschluss und Stenose der Arteria cerebri media {I66.0#G}")
;}

; Onkologie                                      	;{
:X*:.Familienan::SendDiagnosis("Bösartige Neubildung der Verdauungsorgane in der Familienanamnese {Z80.0G}")
:X*:.Hypophysena::SendDiagnosis("Hypophysenadenom {D35.2G}")
:X*:.lymphogene M::SendDiagnosis("lymphogene Metastasierung {C79.9#}")
:X*:.Adenom Neb::            	;{ Nebennierenadenom
:X*:.Nebennierenad::SendDiagnosis("Nebennierenadenom {E27.8LG}") ;}
:XC*:.CLL::                        	;{ Chronische lymphatische Leukämie
:X*:.chronisch ly::
:X*:.Chronische ly::SendDiagnosis("Chronische lymphatische Leukämie ohne Angabe einer (kompletten) Remission {C91.10#}") ;}
:X*:.CLL::SendDiagnosis("Chronische lymphatische Leukämie (CLL) ED #/## {C91.10G}") ; cll klein
:X*:.Pankreaskopf-Ca::       	;{ Pankreas-Carcinom
:X*:.Bauchspeicheldrüsenk::
:X*:.Pankreas-K::
:X*:.Pankreas K::
:X*:.Pankreas-C::
:X*:.Pankreas C::
:X*:.Pankreaskopf-K::
:X*:.Pankreaskopf K::
:X*:.Pankreaskopf Ca::SendDiagnosis("Bösartige Neubildung des Pankreaskopfes {C25.0#}") ;}
:X*:.Colon Ca::                  	;{ intestinaler Krebs ab Dickdarm
:X*:.Colon-Ca::
:X*:.Dickdarmk::
:X*:.Rektumk::
:X*:.Rektum C::
:X*:.Rektum-C::
:X*:.Sigma C::
:X*:.Sigma-C::
:X*:.Sigmak::
:XC*:.COCA::
Auswahlbox(["title=intestinaler Krebs ab Dickdarm border=on "
				, 	"Karzinom der Appendix {C18.1#}"
				, 	"Karzinom des Colon ascendens {C18.2#}"
				,	"Karzinom der Flexura coli dextra {C18.3#}"
				,	"Karzinom des Colon transversum {C18.4#}"
				,	"Karzinom der Flexura coli sinistra {C18.5#}"
				,	"Karzinom des Colon descendens {C18.6#}"
				,	"Karzinom des Colon sigmoideum {C18.7#}"
				,	"Sigmakarzinom {C18.7#}"
				,	"Karzinom des Colon und des Sigmas {C18.8#}"
				,	"Adenokarzinom des Kolons {C18.9#}"
				, 	"Lynch-Syndrom {C18.9#}"
				, 	"Kolorektales Karzinom {C19#}"
				, 	"Rektumkarzinom {C20#}"
				, 	"Karzinom des Anus {C21.0#}"])
return
;}
:X*:.Mamma-Ca::               	;{ Mamma-Carcinom
:X*:.Mamma Ca::
:X*:.Brustdrüsen-Ca::
:X*:.Brustdrüsen Ca::
SendDiagnosis("Mamma-Carcinom {C50.9#}")
return ;}
:X*:.Myelodys::SendDiagnosis("Myelodysplastisches Syndrom {D46.9#}")
;}

; Orthopädie / Unfallchirurgie            	;{
 ; --- obere Extremitäten ---
:X*:.Beckenschi::                                 	;{ Beckenschiefstand
Auswahlbox(["title=Beckenschiefstand border=on maxRows=3"
				, "Beckenschiefstand bei Beinverkürzung {M21.79#}", "Beckenschiefstand {M95.5#}", "Erworbener Beckenschiefstand {M95.5#}"])
return ;}
:X*:.Carpalt::                                     	;{ Carpaltunnelsyndrom
:X*:.Karpal::
:XC*:.CTS::SendDiagnosis("Karpaltunnelsyndrom {G56.0#G}") ;}
:X*:.Claudicatio s::SendDiagnosis("Claudicatio spinalis {M43.16G}")
:X*:.Degeneration Supraspi::SendDiagnosis("Degeneration d. M.supraspinatus und Infraspinatusmuskulatur (M66.59#G)")
:X*:.Epikondylitis ul::                             	;{ Epikondylitis des Armes
:X*:.Golf::SendDiagnosis("Epicondylitis ulnaris humeri {M77.0#G}")
:X*:.Epikondylitis ra::
:X*:.Tennis::SendDiagnosis("Epicondylitis radialis humeri {M77.1#G}")    	;}
:X*:.Fersensp::                                       ;{ plantarer und hinterer Fersensporn
:X*:.unterer F::
:X*:.plantarer F::SendDiagnosis("plantarer Fersensporn {M77.3#}")
:X*:.hinterer F::SendDiagnosis("hinterer Fersensporn {M77.3#}")              	;}
:X*:.Handgelenkfra::SendDiagnosis("Handgelenkfraktur {S62.8#}")
:X*:.BWS Kyph::                                  	;{ Kyphosen
:X*:.Kyphose BWS::SendDiagnosis("BWS-Kyphose {M40.24G}") ;}
:X*:.Rhiz::SendDiagnosis("Rhizarthrose {M18.9#G}")
; ----- Schulter
:X*:.Bizepsehnenentz::                         	;{ Tendinitis der Bizepssehne
:X*:.Tendinitis Bizeps::SendDiagnosis("Tendinitis des Musculus biceps brachii {M75.2#G}") ;}
:X*:.Bizepsehnenrup::                         	;{ Ruptur der Bizepssehne
:X*:.Ruptur Bizeps::SendDiagnosis("Abriß des Musculus biceps brachii {S46.2#G}") ;}
:X*:.Bursitis Sch::                                 	;{ Bursitis subacromialis
:X*:.Schulter Bursitis::SendDiagnosis("Bursitis im Schulterbereich {M75.5#G}") ;}
:X*:.Luxation lange Bizepss::               	;{ Luxation der langen Bizepssehne
:X*:.Luxation Bizepssehne lang::
:X*:.Luxation der langen Bizepss::SendDiagnosis("Luxation der langen Bizepssehne, re. (S46.1#G)") ;}
:X*:.Impingement-Syndrom S:: 	        	;{ Impingementsyndrom der Schulter
:X*:.ImpingementSyndrom S::
:X*:.Impingement Sch::
:XC*:.ISS::
:X*:.Schulter Imping::SendDiagnosis("Impingement-Syndrom der Schulter {M75.4#G}") ;}
:X*:.Rota::SendDiagnosis("Läsionen der Rotatorenmanschette {M75.1#G}")
:X*:.Schulterlux::SendDiagnosis("Luxation des Schultergelenkes nach # {S43.01#G}")
:X*:.Suprasp::SendDiagnosis("Supraspinatussehnenruptur {M75.1#G}")
 ; --- Wirbelsäule ---
:RC*:.BB::                                         	;{ Blockierung der Wirbelsäule (HWS,BWS,LWS)
:RC*:.Blockierung BWS::SendDiagnosis("Blockierung BWS {M99.82G}")
:RC*:.BH::SendDiagnosis("Blockierung HWS {M99.81G}")
:RC*:.BL::SendDiagnosis("Blockierung LWS {M99.83G}") ;}
:X*:.Bandscheibenvorfall H::               	;{ Bandscheibenvorfall der HWS
:X*:.Bandscheibe H::SendDiagnosis("zervikaler Bandscheibenschaden mit Radikulopathie {M50.1#}") ;}
:X*:.Bandscheibenvorfall L::                	;{ Bandscheibenvorfall der LWS
:X*:.Bandscheibe L::SendDiagnosis("lumbale Bandscheibenschäden mit Radikulopathie (G55.1*) {+M51.1G}; "
						. "Kompression v. Nervenwurzeln u. -plexus bei Bandscheibenschäden (M50-M51+) {*G55.1G}") ;}
:X*:.Neurofor::                                    	;{ Neuroforamenstenosen
:X*:.Stenose Foram L::SendDiagnosis("Knöcherne Stenose der Foramina intervertebralia {M99.69G};") ;}
:X*:.Skoliose L::SendDiagnosis("Skoliose, nicht näher bezeichnet: Lumbalbereich {M41.96#G}")
:X*:.Spondylose L::SendDiagnosis("Sonstige Spondylose: Lumbalbereich {M47.86G}")
:X*:.Wirbelkanalstenose H::                  	;{ Spinalkanalstenosen
:X*:.Spinalkanalstenose H::SendDiagnosis("Spinalkanalstenose HWS {M48.02#G}")
:X*:.Wirbelkanalstenose L::
:X*:.Spinalkanalstenose L::SendDiagnosis("Spinalkanalstenose: Lumbalbereich {M48.06#G}") ;}
:X*:.Wirbelgleiten L::                          	;{ Spondylolistesis LWS
:X*:.Spondylolistesis L::SendDiagnosis("Spondylolisthesis: Lumbosakralbereich {M43.17G}") ;}
:X*:.Trigger::				       						;{ Myogelosen, Triggerpunkte
:X*:.Trigg::SendDiagnosis("Triggerpunkt #WS {M62.88#}")
:X*:.Myog::SendDiagnosis("Myogelosen {M62.89G}")
:X*:.MyoTrig::SendDiagnosis("Triggerpunkt BWS {M62.88G}; Myogelosen {M62.89G}") ;}
:X*:.Muskelatr::                                    	;{ Muskelatrophie
:X*:.Muskelverlust::SendDiagnosis("Muskelatrophie der Extremitäten {M62.50#}")
:X*:.Trig::SendDiagnosis("Triggerpunkt BWS {M62.88G}") ;}
 ; --- untere Extremitäten ---
 ; ----- Hüfte
:X*:.Coxa::SendDiagnosis("Primäre Koxarthrose, beidseitig {M16.0G}")
:XC*:.HTEP::                                        	;{ Hüft-Endoprothese
:X*:.Endoprothese Hüft::SendDiagnosis("HTEP {Z96.68LG}") ;}
 ; ----- Knie
:X*:.Baker::SendDiagnosis("Baker-Zyste {M71.2#G}")
:X*:.Genu var::	                                 	;{ Genu varum
:X*:.O-Bein::
:X*:.O Bein::SendDiagnosis("Genu varum {M21.16BG}}") ;}
:X*:.Gonar::SendDiagnosis("Primäre Gonarthrose {M17.0#G}")
:X*:.Menisikopathie::                          	;{ Meniskopathien des Knie Auswahlbox
Auswahlbox([ "title=nicht akuter Meniskuserkrankungen border=on"
				, 	"Innenmeniskopathie {M23.33}"
				,	"Innenmeniskusvorderhornläsion {M23.31}"
				, 	"Innenmeniskushinterhornläsion {M23.32}"
				,	"Außenmeniskusläsion {M23.36}"
				,	"Alter Meniskusschaden {M23.29}"
				, 	"Meniskopathie {M23.39}"
				, 	"Degenerative Meniskopathie {M23.39}"
				,	"Meniskusganglion {M23.09}"
				, 	"Meniskuszyste {M23.09}"
				,  	"Gutartige Neubildung des Meniskus {D16.2}"])
return ;}
:X*:.Kniesc::                                       	;{ Knieschmerzen
:X*:.Gonal::SendDiagnosis("Gelenkschmerz: Kniegelenk {M25.56#G}") ;}
:X*:.Kniegelenker::                              	;{ Kniegelenkserguß
:X*:.Erguß Knie::SendDiagnosis("Kniegelenkerguss {M25.46#G} ") ;}
:X*:.Knie TEP::SendDiagnosis("Knie TEP {Z96.6#G}")
:X*:.Hallux::SendDiagnosis("Hallux valgus {M20.1#G}")
:X*:.Patella F::                                      	;{ Patella Fraktur
:XC*:.FXPA::SendDiagnosis("Fraktur der Patella {S82.0#G}") ;}
:X*:.Knick::                                          	;{ Knick-Senk-Spreizfuß
:X*:.Senk::
:X*:.Spreiz::SendDiagnosis("Knick-Senk-Spreizfuß {M21.07BG}") ;}
 ; --- ohne Lokalisation ---
:X*:.Polyarthro::SendDiagnosis("Polyarthrose {M15.0G}")
:X*:.Postmenop::                                	;{ Postmenopausale Osteoporose
:X*:.Osteoporose::SendDiagnosis("Postmenopausale Osteoporose {M80.08G}") ;}
;}

; Pädiatrie                                            	;{
:X*:.Sprachent::SendDiagnosis("Rückstand in der Sprachentwicklung {F80.9V}")
;}

; Pulmologie                                      	;{
:X*:.Asthma ni::SendDiagnosis("Nichtallergisches Asthma bronchiale {J45.1G}")
:X*:.Asthma br::                                   	;{  Asthma bronchiale
:X*:.Lungenast::
:XC*:.AB::SendDiagnosis("allerg. Asthma bronchiale {J45.0G}") ;}
:X*:.COPD::                                         	;{ COPD
:X*:.chronische Br::SendDiagnosis("COPD {J44.89G}") ;}
:X*:.pneumonie::SendDiagnosis("Pneumonie {J15.8#G}")
:X*:.pulmonale H::                            	;{ Pulmonale arterielle Hypertonie
:X*:.PAH::
:X*:.pulmonale arteriel::SendDiagnosis("Pulmonale arterielle Hypertonie {I27.28G}; ") ;}
:X*:.zentrales S::                                 	;{ zentrales Schlafapnoesyndrom
:X*:.zentrale S::SendDiagnosis("Zentrales Schlafapnoesyndrom {G47.30#}") ;}
:X*:.OSAS::                                       	;{ obstruktives Schlafapnoesyndrom
:X*:.obstruktives S::
:X*:.Schlafapnoe::
:X*:.obstruktive S::SendDiagnosis("Obstruktives Schlafapnoesyndrom {G47.31#}") ;}
:X*:.Sarkoi::SendDiagnosis("Sarkoidose der Lunge {D86.0#}")
:X*:.spastische Br::SendDiagnosis("Spastische Bronchitis {J40G}")
:X*:.Tracheost::                                    	;{ Tracheostoma
:X*:.Vorhandensein Tra::SendDiagnosis("Vorhandensein eines Tracheostomas {Z93.0#}") ;}
;}

; Psychiatrie                                       	;{
:X*:.Agor::SendDiagnosis("Agoraphobie mit Panikstörung {F40.00G}")
:X*:.ADHS::                                     	;{ ADHS/ADS
:X*:.Aufmerksamkeit Hyper::
:X*:.Aufmerksamkeit Defizit::
:X*:.ADHS::SendDiagnosis("ADHS {F90.0#}")
:X*:.ADS::SendDiagnosis("ADS {F98.80#}") ;}
:X*:.Alb::SendDiagnosis("Albträume [Angstträume] {F51.5G}")
:XC*:.C2::                                      	;{ Abhängigkeitssyndrom durch Alkoholgebrauch
:X*:.Alkoholabhängigkeitssy::SendDiagnosis("Abhängigkeitssyndrom durch Alkoholgebrauch {F10.2G}") ;}
:X*:.Anpass::SendDiagnosis("Anpassungsstörung {F43.2G}")
:X*:.akute Bel::                                	;{ akute Belastungsreaktion
:X*:.Belast::SendDiagnosis("Akute Belastungsreaktion {F43.0G}") ;}
:X*:.Benzo::SendDiagnosis("Abhängigkeitssyndrom durch Gebrauch von Sedativa oder Hypnotika {F13.2G}")
:X*:.Depression lr::                          	;{ Depression, leichtgradig rezidivierend
:X*:.leichtg::SendDiagnosis("leichtgradige Depression {F32.0G}") ;}
:X*:.Depression mr::                         	;{ Depression, mittelgradig rezidivierend
:X*:.mittelg::SendDiagnosis("mittelgradige Depression {F32.1G}") ;}
:X*:.Depression sr::                          	;{ Depression, schwergradig rezidivierend
:X*:.schwerg::SendDiagnosis("schwergradige Depression {F32.2G}") ;}
:X*:.disso::SendDiagnosis("Dissoziative Störungen [Konversionsstörungen], gemischt {F44.7G}")
:X*:.Dysthym::SendDiagnosis("Dysthymia {F34.1G}")
:X*:.Hypocho::SendDiagnosis("Hypochondrische Störung {F45.2G}")
:X*:.Lese::                                        	;{ Lese- und Rechtschreibstörung
:X*:.Rechtschr::SendDiagnosis("Lese- und Rechtschreibstörung {F81.0G}") ;}
:X*:.Neuras::SendDiagnosis("Neurasthenie {F48.0G}")
:X*:.Angst::                                   	;{ Panikstörung
:X*:.Panik::SendDiagnosis("Panikstörung [episodisch paroxysmale Angst] {F41.0G}") ;}
:XC*:.soziale P::		                    	;{ soziale Phobien
:XC*:.soziale Ä::SendDiagnosis("Soziale Phobien {F40.1#}") ;}
:XC*:.PTBS::                                   	;{ posttraumatische Belastungsstörung
:X*:.posttrau::SendDiagnosis("Posttraumatische Belastungsstörung {F43.1#}") ;}
:X*:.Psy::                                        	;{ Somatisierungsstörung
:X*:.Soma::SendDiagnosis("Somatisierungsstörung {F45.0G}") ;}
:X*:.Kontaktan::                             	;{ Kontaktanlässe mit Bezug auf die soziale Umgebung
:X*:.soziale Um::SendDiagnosis("Kontaktanlässe mit Bezug auf die soziale Umgebung {Z60G}") ;}
:X*:.Schlaf::SendDiagnosis("Ein- und Durchschlafstörungen {G47.0G}")
:X*:.Zwangss::SendDiagnosis("Zwangsstörungen {F42.8G}")
;}

; Rheumatologie/Autoimmune Erkr.      	;{
:X*:.Alopezia::
:X*:.kreisrund::SendDiagnosis("Alopecia areata {L63.9G}")
:X*:.Fibromy::SendDiagnosis("Fibromyalgie {M79.70G}")
:X*:.Löfgren::SendDiagnosis("Löfgren Syndrom  {D86.8#}")
:X*:.Morbus Bech::                           		;{ Morbus Bechterew
:X*:.Spondylitis ank::
:X*:.Bechterew::SendDiagnosis("Morbus von Bechterew {M45.09#}") ;}
:X*:.PCP::                                           	;{ chronische Polyarthritis
:X*:.Chronische Polyar::SendDiagnosis("Chronische Polyarthritis, mehrere Lokalisationen {M06.90G}") ;}
:X*:.Polymy::SendDiagnosis("Polymyalgia rheumatica {M35.3#}")
:X*:.Psoriasis Art::                              	;{ Psoriasis Arthropathie
:X*:.Psoriasis-Art::SendDiagnosis("Psoriasis-Arthropathie (M07.0-M07.3*, M09.0-*) {+L40.5G}") ;}
;}

; Urologie                                         	;{
:X*:.Balani::                           	;{ Eichel- und/oder Vorhautentzündungen
:X*:.Eichel::SendDiagnosis("Balanitis {N48.1G}")
:X*:.Balano::SendDiagnosis("Balanoposthitis {N48.1G}")
:X*:.Vorhaut::
:X*:.Posthitis::SendDiagnosis("Posthitis {N48.1G}") ;}
:X*:.Bakteriu::SendDiagnosis("Asymptomatische Bakteriurie {N39.0G}")
:X*:.Hämatu::SendDiagnosis("Hämaturie {R31G}")
:X*:.Harnst::SendDiagnosis("Harnstauungsniere #. Grad {N13.3#G}")
:X*:.Harnv::SendDiagnosis("Harnverhaltung {R33G}")
:X*:.Harnröhrenstr::               	;{ Harnröhrenstriktur
:X*:.Striktur Ha::Auswahlbox(["title=Harnröhrenstriktur border=on ",	"Harnröhrenstriktur {N35.9G}"
										, 	"Harnröhrenstriktur nach medizinischen Maßnahmen {N99.1G}"]) ;}
:X*:.harnw::                            	;{ Harnwegsinfekt
:X*:.hti::SendDiagnosis("Harnwegsinfektion {N39.0G}") ;}
:X*:.Hyposp::SendDiagnosis("Hypospadie {Q54.9G}")
:X*:.Nierenko::SendDiagnosis("Nierenkolik {N23#}")
:X*:.nierenz::SendDiagnosis("Erworbene Zyste der Niere {N28.1#G}")
:XC*:.PRCA::                         	;{ Prostata-Ca
:X*:.Carcinom Pros::
:X*:.Prostatak::
:X*:.Prostata-K::
:X*:.Prostata-C::
:X*:.Bösartige Neubildung Pros::
:X*:.Prostata Ca::SendDiagnosis("Bösartige Neubildung der Prostata {C61G}") ;}
:X*:.BPH::                              	;{ benigne Prostatahyperplasie
:X*:Benigne Prostatah::
:X*:.Prostatah::SendDiagnosis("Prostatahyperplasie {N40G}") ;}
:X*:.Urothel::                         	;{ Urothel-/Harnblasen-Ca
:X*:.Harnblasenk::SendDiagnosis("Urothelkarzinom {C68.9G}") ;}    ;
:X*:.Zystoz::SendDiagnosis("Zystozele {N81.1G}")
;}

; was der Internist so wichtig findet     	;{
:X*:.Adipo::SendDiagnosis("Adipositas {E66.0#G}")
:X*:.adi1::SendDiagnosis("Adipositas Erwachsene Grad I (WHO) {E66.00G}")
:X*:.adi2::SendDiagnosis("Adipositas Erwachsene Grad II (WHO) {E66.01G}")
:X*:.adi3::                                 	;{
Auswahlbox(["title=Adipositas durch übermäßige Kalorienzufuhr border=on BgColor1=FB6A6A BgColor2=FDBDBD"
				,	"Adipositas Erwachsene Grad III (WHO) [BMI 40-50] {E66.06G}"
				,	"Adipositas Erwachsene Grad III (WHO) [BMI 50-60] {E66.07G}"
				,	"Adipositas Erwachsene Grad III (WHO) [BMI >60] {E66.08G}"
				,  	"Adipositas Erwachsene Grad II (WHO) {E66.01G}"
				,  	"Adipositas Erwachsene Grad I (WHO) {E66.00G}"
				,  	"Adipositas Grad III (WHO) durch übermäßige Kalorienzufuhr {E66.09G}"])
return ;}
:X*:.Hyperto::                            	;{
:X*:.arterielle H::
:cX*:.HY::SendDiagnosis("essentielle Hypertonie {I10.00G}") ;}
:X*:.Hyperch::SendDiagnosis("Hypercholesterinämie {E78.0G}")
:X*:.Hypertri::SendDiagnosis("Hypertriglyzeridämie, G. {E78.1G}")
:X*:.Hypernat::SendDiagnosis("Hypernatriämie {E87.0}")
:X*:.Hyponat::SendDiagnosis("Hyponatriämie {E87.1}")
:X*:.Hyperka::SendDiagnosis("Hyperkaliämie {E87.5}")
:X*:.Hypoka::SendDiagnosis("Hypokaliämie {E87.6G}")
;}

; ohne Kategorie                               	;{
:X*:.Anämie Blutung::                 	;{ Blutungsanämie
:X*:.Blutungs::SendDiagnosis("Akute Blutungsanämie {D62G}") ;}
:X*:.Anämie Eisen::                     	;{ Eisenmangelanämie
:X*:.Eisenm::SendDiagnosis("Eisenmangelanämie {D50.9G}") ;}
:X*:.Angioödem::SendDiagnosis("Allergisches Angioödem {T78.3#}")
:X*:.Agn::SendDiagnosis("Agnosie {R48.1#}")	;
:X*:.Anos::                                  	;{ Anosmie
:X*:.Geruchsver::
:X*:.Geruchssinn::SendDiagnosis("Anosmie {R43.0G}") ;}
:X*:.Aortenskl::SendDiagnosis("Atherosklerose der Aorta {I70.0#}")
:X*:.Antikoa::                             	;{ Dauertherapie mit Antikoagulanzien/Falithrom/NOAK
:X*:.Falit::
:X*:.NOAK::
:X*:.Marcu::SendDiagnosis("Dauertherapie mit Antikoagulanzien{Z92.1G}") ;}
:X*:.Allgemeinunter::           	      	;{ Ärztliche Allgemeinuntersuchung
:X*:.Ärztliche All::SendDiagnosis("Ärztliche Allgemeinuntersuchung {Z00.0#}") ;}
:X*:.Aspira::SendDiagnosis("Aspirationspneumonie {J69.8#}")
:X*:.Autoimmunthy::SendDiagnosis("Autoimmunthyreoiditis {E06.3#}")
;B
:X*:.Bauchs::SendDiagnosis("Bauchschmerzen {R10.4}")  	;{ Bauchschmerzen aller Ebenen
:X*:.Oberb::SendDiagnosis("Oberbauchschmerz {R10.1G}")
:X*:.Unterb::SendDiagnosis("Unterbauchschmerz {R10.3G}")
:X*:.Oberba::
:X*:.Bauchschmerz1::SendDiagnosis("Schmerzen im Bereich des Oberbauches {R10.1G}") ;}
:X*:.Brusts::                               	;{ Brustschmerzen
:X*:.Thoraxs::SendDiagnosis("Brustschmerzen {R07.4G}; Brustschmerzen bei der Atmung {R07.1G}") ;}
;C
:XC*:.CE::                                 	;{ Cephalgien
:XC*:.Cephalgie::
:X*:.Kopfs::SendDiagnosis("Kopfschmerz {G44.8G}") ;}
;D
:X*:.Dyspn::                                 	;{ Dyspnoe - unspez.
:X*:.Luftnot::SendDiagnosis("Dyspnoe {R06.0G}") ;}
:X*:.Dyspho::                             	;{ Dysphonie - unspez.
:X*:.Heiser::SendDiagnosis("Dysphonie {R49.0G}") ;}
;E
;F
:X*:.FaktorII::                               	;{  Faktor-II-Mangel
:X*:.Faktor II::
:X*:.Faktor-II::
:X*:.Faktor 2::
:X*:.Faktor-2::
:X*:.Faktor2::SendDiagnosis("Hereditärer Faktor-II-Mangel {D68.21G}") ;}
:X*:.Fieber unbe::                       	;{	FUO - fever of unknown origin
:X*:.FUO::SendDiagnosis("Fieber unbekannter Ursache {R50.80G}") ;}
:X*:.Fruktose::SendDiagnosis("Störungen des Fruktosestoffwechsels {E74.1G}")
;G
:X*:.Gastrosto::                          	;{ Gastrostoma
:X*:.PEG::SendDiagnosis("Vorhandensein einer PEG (Gastrostoma) {Z93.1G}") ;}
:X*:.Gehen::SendDiagnosis("Probleme beim Gehen {R26.2G}")
:X*:.geriat::                                  	;{ [Auswahlbox] geriatrische ICDs
:X*:.gb::
params := ["title=geriatrische Diagnosen maxRows=30"]
For each, icd in Addendum.GeriatrieICD
	params.Push(icd)
Auswahlbox(params)
return	;}
:X*:.Gewichtsa::SendDiagnosis("abnorme Gewichtsabnahme {R63.4G}")
:X*:.Gewichtsz::SendDiagnosis("abnorme Gewichtszunahme {R63.5G}")
:X*:.Gicht::SendDiagnosis("Idiopathische Gicht, {M10.09#G}")
:X*:.gicht2::SendDiagnosis("Idiopathische Gicht {M10.00G}")
:X*:.Poda::SendDiagnosis("Podagra {M10.97#G}")
;H
:X*:.Harnink::SendDiagnosis("Harninkontinenz {R32G}")
:X*:.Hashi::SendDiagnosis("Hashimoto-Thyreoiditis {E06.3#}")
:X*:.Hyperhi::SendDiagnosis("Hyperhidrose, umschrieben {R61.0G}")
:X*:.Hyperlipi::SendDiagnosis("Hyperlipidämie {E78.5G}")
:X*:.Hyperlipo::SendDiagnosis("Hyperlipoproteinämie, G. {E78.8G}")
:X*:.hyperuri::SendDiagnosis("Hyperurikämie ohne Gelenkentzündung oder tophische Gicht {E79.0G}")
:X*:.Hypothy::SendDiagnosis("Hypothyreose {E03.9G}")
;I
:X*:.Insekt::SendDiagnosis("Insektenstich/-biß {T14.03G}")
:X*:.Immobil::                         	;{ Immobilität
:X*:.bettlägerig::SendDiagnosis("Immobilität {R26.3G}") ;}
;K
:X*:.Kache::SendDiagnosis("Kachexie {R64G}")
;L
:X*:.Laktose::SendDiagnosis("Laktoseintoleranz {E73.9G}")
:X*:.Leukozytose::SendDiagnosis("Leukozytose {D72.8G}")
:X*:.Lipom::SendDiagnosis("Lipom der Unterhaut {D17.3G}")
:X*:.Lip1::                              	;{ Lipödem und Stadien
:X*:.Lip I::
:X*:.Lip2::
:X*:.Lip II::
:X*:.Lip3::
:X*:.Lip III::SendDiagnosis("Lipödem, Stadium $CMB {E88.2$-1G}")
:X*:.Lipö::SendDiagnosis("nicht näher bezeichnetes Lipödem {E88.28G}") ;}
:XC*:.LAE::								;{ Lungenembolie n.n.bz.
:XC*:.Lungenarterienembolie::
:XC*:.Lungenembolie::SendDiagnosis("# Lungenembolie {I26.9#}") ;}
:X*:.Lungenembolie cor::        	;{ Lungenembolie mit akutem Cor pulmonale
:X*:.LE cor::SendDiagnosis("Lungenembolie mit akutem Cor pulmonale {I26.0#}") ;}
:X*:.LE ohne::                        	;{ Lungenembolie ohne akutes Cor pulmonale
:X*:.Lungenembolie ohne::SendDiagnosis("Lungenembolie ohne akutes Cor pulmonale {I26.9#}") ;}
;M
:X*:.Mara::SendDiagnosis("Alimentärer Marasmus {E41G}")
:X*:.Meteo::SendDiagnosis("Meteorismus {R14G}")
;N
:X*:.Niko::SendDiagnosis("Nikotinabusus {F17.2G}")
;O
:X*:.Obsti::SendDiagnosis("Obstipation {K59.00G}")
;P
:X*:.Parosm::SendDiagnosis("Parosmie {R43.1G}")
:X*:.Pulpitis::                         	;{ Wurzelkanalentzündung
:X*:.Wurzelka::SendDiagnosis("Pulpitis {K04.0G}") ;}
;R
:X*:.Raucherleu::SendDiagnosis("Raucherleukozytose {D72.8G}")
:X*:.Morbus Ray::                   	;{ Morbus Raynold (Rey)
:X*:.Morbus Rey::                    	; 	hilft bei abweichenden Rechtschreibungskenntnissen
:X*:.Ray::SendDiagnosis("Raynaud-Syndrom {I73.0G}") ;}
:X*:.Rhinopath::                       	;{ allergische Rhinopathie
:X*:.allergische Rhin::
:X*:.Pollinosis::
:X*:.Pollenallergie::SendDiagnosis("Allergische Rhinopathie durch Pollen {J30.1G}") ;}
:X*:.Schmerzen HWS::           	;{ Rückenschmerzen nach Lokalisation
:X*:.Zervikalneuralgie::SendDiagnosis("Zervikalneuralgie {M54.2G}")
:X*:.Schmerzen BWS::SendDiagnosis("Schmerzen im Bereich der Brustwirbelsäule {M54.6G}")
:X*:.Schmerzen LWS::
:X*:.Lumboi::SendDiagnosis("Lumboischialgie {M54.4G}")  ;}
;S
:X*:.Schlafst::SendDiagnosis("Ein- und Durchschlafstörungen {G47.0G}")
:X*:.Schwind::                        	;{ Schwindel und Taumel
:X*:.Taumel::SendDiagnosis("Schwindel und Taumel {R42G}") ;}
:X*:.chronischer S::                 	;{ chronisches Schmerzsyndrom
:X*:.chronisches S::SendDiagnosis("Chronisches Schmerzsyndrom {R52.2G}") ;}
:X*:.SIADH::SendDiagnosis("SIADH-Syndrom {E22.2G}")
:X*:.Singul::                           	;{ Schluckauf
:X*:.Schluckauf::SendDiagnosis("Schluckauf {R06.6}") ;}
:X*:.Standunsicherheit::SendDiagnosis("Standunsicherheit {R26.8G}")
:X*:.Struma m::SendDiagnosis("Struma multinodosa {E04.2G}")
:X*:.Stuhlink::SendDiagnosis("Stuhlinkontinenz {R15#}")
:X*:.Sturznei::SendDiagnosis("Sturzneigung {R29.6G}")
:X*:.Synkope::                         	;{ Synkope und Kollaps
:X*:.Bewußtlos::
:X*:.Kreislaufk::
:X*:.Kollap::SendDiagnosis("Synkope und Kollaps {R55G}") ;}
;T
:X*:.Testosteronm::SendDiagnosis("Testosteronmangel {E29.1#}")
:X*:.TVT::SendDiagnosis("Tiefe Venenthrombose {I80.1#G}")
:X*:.Thrombose V.pop::SendDiagnosis("Thrombose V.poplitea {I80.1#G}")
:X*:.Thrombose V.fib::SendDiagnosis("Thrombose V.fibularis {I80.1#G}")
;U/Ü
:X*:.Übel::                              	;{ Übelkeit und Erbrechen
:X*:.Erbre::SendDiagnosis("Übelkeit und Erbrechen {R11G}") ;}
:X*:.Unguis::SendDiagnosis("Unguis incarnatus {L60.0#}")
;V
:X*:.Verti::SendDiagnosis("Vertigo {R42G}")
:X*:.Urti::                                 	;{ [Auswahlbox] Urtikariaformen
:X*:.Kälteu::
:X*:.Wärmeu::
:X*:.Chronische Ur::
:X*:.Cholinergische::
:X*:.Urtikarielles::
:X*:.Thermale::
:X*:.HitzeU::
	Auswahlbox(["title=Urtikaria L50.X Border=on"
					, 	"Allergische Urtikaria {L50.0G}"     	, 	"Idiopathische Urtikaria {L50.1G}"	, 	"Kälteurtikaria {L50.2G}"
					, 	"Wärmeurtikaria {L50.2G}"       		, 	"Urticaria factitia {L50.3G}"   		, 	"Urticaria mechanica {L50.4G}"
					, 	"Cholinergische Urtikaria {L50.5G}"	, 	"Chronische Urtikaria {L50.8G}"	, 	"Urtikarielles Exanthem {L50.9}"])
return
:X*:.AllergischeU::
:X*:.Allergische U::SendDiagnosis("Allergische Urtikaria {L50.0G}")
;}
:X*:.VitD::                             	;{ Vitamin D Mangel
:X*:.VitaminD::
:X*:.Vitamin D::SendDiagnosis("Vitamin D Mangel {E64.8G}")  ;}
:X*:.Wespe::SendDiagnosis("Wespengiftallergie {T78.4G}")
; Z
:X*:.Zeru::                              	;{Cerumen
:X*:.Ohrenschm::
:X*:.Ceru::SendDiagnosis("Zeruminalpfropf {H61.2#G}") ;}
:X*:.Zöliakie::
:X*:.Sprue::SendDiagnosis("Zöliakie ED2012, G. {K90.0G};")
;}

; Impfungen ;
:X*:.igx::SendDiagnosis("E22 {Z25.1G}")

; Rechtschreibung
:*:Ortop::Orthop

; ### ----- ----- ----- ----- ;}
;}
; --- bef | anam | info                                                                	;{
#IF WinActive("ahk_class OptoAppClass") && RegExMatch(AlbisGetActiveControl("contraction"), "(bef|anam|info)")
:*:cafe-au::Café-au-lait-Flecken{Space}
;~ :*?X::"::SendAlternately("„|“")
#IF ;}
; --- Impf                                                                            	;{
#If (WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "Impf") )
:R*:1.C::1.COVID-19 Impfung
:R*:2.C::2.COVID-19 Impfung
:R*:3.C::2.COVID-19 Impfung
:R*:Abrech::KEINE ABRECHNUNG MÖGLICH da Abrechnungsausschluß wegen 88332 (Impfberatung)
:R*:.HZ::Herpes Zoster
#If
;}
; --- Info                                                                             	;{
#If (WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "info") )
:R*:letztes::letztes Quartal nicht da (keine Chronikerziffer möglich)
#If
;}
; --- aem                                                                              	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "aem")
:R*:ENife::Einnahmebeschreibung für Nifedipin-Notfalltropfen mitgegeben
:R*:ETram::Einnahmebeschreibung für Tramadol mitgegeben
:R*:EVitD3::Einnahmebeschreibung für Dekristol mitgegeben
:R*:ETry::Einnahmebeschreibung für Tryasol mitgegeben
:R*:EMac::Einnahmebeschreibung für Macrogol mitgegeben
#If
;}
; --- Rezept                                                                           	;{
#If WinActive("Muster 16 ahk_class #32770") || WinActive("Privatrezept ahk_class #32770")
:*:#Trya::                                           	;{ Tryasol 30 ml
	IfapVerordnung("Tryasol 30 ml")
return ;}
#If
;}
; --- scan/pdf                                                                         	;{
#If HotS_Scan_IF()
HotS_Scan_IF() {                                                                 	;-- kontextsensitive Auslösung Diagnosenfelder
	; letzte Änderung 30.09.2021
	If (WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), Addendum.PDF.ScanKuerzel))
		return true
return false
}
HotS_Scan_Hotstrings() {                                                     	;-- ausgelagerte Hotstringdatei

	If FileExist(Addendum.DBPath "\Dictionary\Hotstrings_scan.json") {
		fn_IFPAbr := Func("HotS_Scan_IF")
		Hotkey, If, % fn_IFPAbr
		For replacement, hotstrings in  JSONData.Load(Addendum.DBPath "\Dictionary\Hotstrings_scan.json", "", "UTF-8")
			For each, string in hotstrings
				Hotstring(string, replacement)
	}

}
;}
; --- Familiendaten                                                                   	;{
#IfWinActive, Familiendaten ahk_class #32770
:C*:G::Groß
:C*:Gm::Großmutter
:C*:Oma::Großmutter
:C*:Opa::Großvater
:C*:On::Onkel
:C*:V::Vater
:C*:v::vater
:C*:Mu::Mutter
:C*:m::mutter
:C*:Ma::Mann
:C*:T::Tochter
:C*:So::Sohn
:C*:Sc::Schw
:C*:a::ager
:C*:ä::ägerin
:C*:e::ester
:C*:i::ieger
:C*:B::Bruder
:C*:Ta::Tante
:C*:C::Cousin
:C*:Fr::Frau
:C*:E::Enkel
:C*:s::sohn
:C*:t::tochter
:C*:L::Lebenspartner
#IfWinActive
;}
; --- Kuerzelfeld                                                                      	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("identify"), "Edit3")
:*:GVU::                                                                                            	;{ GVU,HKS Automatisierung

	ControlGetText, EditDatum, % KK.Datum, ahk_class OptoAppClass
	If !RegExMatch(EditDatum, "\d{2}\.\d{2}\.\d{4}")	{
			PraxTT("Das Datumfeld konnte nicht ausgelesen werden", "3 3")
			return
	}

	DatumZuvor:= AlbisSetzeProgrammDatum(EditDatum)
	ControlFocus, % KK.Kuerzel, ahk_class OptoAppClass
	VerifiedSetText(KK.Kuerzel, "GVU", "ahk_class OptoAppClass")
	SendInput, {Tab}

return ;}
:*b0:AISC::                                                                                       	;{ AISC - Laborwebinterface
	SendInput, {Tab}
	Sleep 50
	SendInput, {Tab}
return ;}
:*:ci::                                                                                                 	;{ Impfung gegen Corona Biontech


return
CoronaImpfung() {

	global CI, hCI
	static hCtrlActive

	ControlGetFocus, hCtrlActive, ahk_class OptoAppClass
	Gui, CI: New 	, -DPIScale +HWNDhCI
	Gui, CI: Font		, % "s12", % "Futura Mk Md"
	Gui, CI: Add  	, Text, % "xm ym w800 Center"
	Gui, CI: Font		, % "s10", % "Futura Bk Bt"
	Gui, CI: Add  	, Text	, % "xm y+15 w150 Right"	, % "Biontech [COMIRNATY Ch.:"
	Gui, CI: Add  	, Edit		, % "x+5 vCIBTE"              	, % "ET5885"
	Gui, CI: Add  	, Text	, % "x+2"                      	, % "]"
	Gui, CI: Add  	, Button	, % "x+5 vCIBT1"              	, % "1.Impfdosis"
	Gui, CI: Add  	, Button	, % "x+5 vCIBT2"              	, % "2.Impfdosis"
	Gui, CI: Add  	, Text	, % "xm y+15 w150 Right"	, % "AstraZeneca [Astrazid Ch.:"
	Gui, CI: Add  	, Edit		, % "x+5 vCIAZE"              	, % "ET5885"
	Gui, CI: Add  	, Text	, % "x+2"                       	, % "]"
	Gui, CI: Add  	, Button	, % "x+5 vCIAZ1"             	, % "1.Impfdosis"
	Gui, CI: Add  	, Button	, % "x+5 vCIAZ2"             	, % "2.Impfdosis"


return
CIGuiClose:
CIGuiEscape:

return

CIMakro:

	ControlFocus, % KK.Kuerzel, ahk_class OptoAppClass
	VerifiedSetText(KK.Kuerzel, "dia", "ahk_class OptoAppClass")
	ControlFocus, % KK.Inhalt, ahk_class OptoAppClass
	VerifiedSetText(KK.Inhalt, , "ahk_class OptoAppClass")
	SendInput, {Tab}

return
}
;}
#If
;}
; --- AUTODOC ---                                                                      	;{
#IfWinActive, ahk_class OptoAppClass
:*:+::
	If !InStr(AlbisGetActiveControl("identify"), "Dokument|Edit3") {
		Send, {+}
		return
	}
	If !Addendum.Flag.AutoDocCalled {
		;MsgBox, % "identify`n" AlbisGetActiveControl("identify")
		Addendum.Flag.AutoDocCalled := true
		AutoDoc()
	}
return
;}
; --- Ausnahmeindikationen                                                             	;{
#If IsFocusedAusnahmeIndikation()
IsFocusedAusnahmeIndikation() {

	; das Fenster "weitere Information" trägt nur den Namen des Patienten
		WinGetText       	, activeWinText	, A
		ControlGetFocus	, activeControl	, A

	If InStr(activeWinText, "Ausnahmeindikation") && InStr(activeControl, "Edit14")
		return true

return false
}

:*:#::
	AutoFillAusnahmeindikation()
return
:*:1-::03000-03040-
:*:2-::03220-
:*:3-::03360-03362-
:*:a-::03000-03040-03220-03360-03362
:*:aok-::95051
#If

AutoFillAusnahmeindikation() {

		;~ static Chroniker, Geriatrisch
		static toAdd   	:= ["03000", "03040", "03220", "03360", "03362"]
		static toRemove := ["95001", "95002", "95003"]

	; Chroniker und Geriatrie Index werden eingelesen ;{  nicht mehr benötigt --- TEST
		;~ If FileExist(Addendum.DbPath "\PatData\DB_Chroniker.txt") && (StrLen(Chroniker) = 0) {
				;~ FileRead, tmpChroniker, % Addendum.DbPath "\PatData\DB_Chroniker.txt"
				;~ Loop, Parse, tmpChroniker, `n, `r
					;~ Chroniker .= Trim(StrSplit(A_LoopField, ";").1) ","
				;~ Chroniker := RTrim(Chroniker, ",")
		;~ }

		;~ If FileExist(Addendum.DBPath "\PatData\DB_DB_Geriatrisch.txt") && (StrLen(Geriatrisch) = 0) {
				;~ FileRead, tmpGeriatrisch, % Addendum.DbPath "\PatData\DB_Geriatrisch.txt"
				;~ Loop, Parse, tmpGeriatrisch, `n, `r
					;~ Geriatrisch .= Trim(StrSplit(A_LoopField, ";").1) ","
				;~ Geriatrisch := RTrim(Geriatrisch, ",")
		;~ }

	;}

	; Steuerelement auslesen
		ControlGetText, ctext, Edit14, ahk_class #32770 ahk_exe albis, Ausnahmeindikationen
		aktPatID := AlbisAktuellePatID()

	; hinzufügende Leistungskomplexe mit Steuerelementinhalt vergleichen
		Loop % toAdd.Count()	{
			If !InStr(ctext, toAdd[A_Index])				{
				If InStr(toAdd[A_Index], "03220") && isPatInGroup(aktPatID, "Chroniker")
					cText .= "-" toAdd[A_Index]
				else if InList(toAdd[A_Index], "03660,03362") && isPatInGroup(aktPatID, "Geriatrisch")
					cText .= "-" toAdd[A_Index]
				else
					cText .= "-" toAdd[A_Index]
			}
		}

		cText := LTrim(ctext, "-")

	; Leistungskomplexe sortieren
		Sort, ctext, N U D-

	; Nachfragen ob alles korrekt ist ;{
		MsgBox, 0x1024, Addendum für Albis on Windows, % ctext "`n`nSind die gemachten Änderungen korrekt?"
		IfMsgBox, Yes
		{
			; Ausnahmenindikationen eintragen
				VerifiedSetText("Edit14", cText, "ahk_class #32770", 200, "Ausnahmeindikation")
				ControlFocus	, % "Edit14", % "ahk_class #32770", % "Ausnahmeindikation"
				ControlSend	, % "Edit14", % "{Enter}" ,% "ahk_class #32770", % "Ausnahmeindikation"
				Sleep 1000

			; Fenster schliessen falls noch geöffnet
				If WinExist("ahk_class #32770", "Ausnahmeindikation")
					VerifiedClick("Ok", "ahk_class #32770", "Ausnahmeindikation",,2)

				Sleep 2000

			; Fenster "Daten von " schliessen
				If WinExist("Daten von", "Anrede")
					VerifiedClick("Button30", "Daten von")
		}
	;}

}
InList(var, list) {
	If var in %list%
		return true
return false
}
isPatInGroup(PatID, dgroup) {

	; dgroup - entweder "Chroniker" oder "geriatrisch"
	; benötigt die über Addendum_Ini.ahk eingelesenen Daten aus \PatData\DB_Chroniker.txt oder DB_geriatrisch.txt

	If (!IsObject(Addendum[dgroup]) || !Addendum[dgroup].Count()) {
		SciTEOutput("Die übergebene Gruppe (" dgroup ") enthält keine Patienten-IDs.")
		return false
	}

	For each, dID in Addendum[dgroup]
		If (PatID = dID)
			return true

return false
}

;}
; --- manuelle Eingabe der Versichertendaten                                          	;{
#If WinActive("Ersatzverfahren ahk_class #32770")                            	; Hilfen im Fenster Ersatzverfahren
:*:AOKN::                                                                                      	;{ Ersatzverfahren Hilfe beim eintragen von AOK Nordost Daten
	ControlSetText, Edit1, % "72101"        	, % "Ersatzverfahren ahk_class #32770"
	ControlSetText, Edit3, % "109519005"	, % "Ersatzverfahren ahk_class #32770"
	ControlSetText, Edit5, % "AOK Nordost"	, % "Ersatzverfahren ahk_class #32770"
	Sleep 100
	ControlFocus, Edit6, % "Ersatzverfahren ahk_class #32770"
return ;}
#If
;}
; --- Rezepte                                                                          	;{
#If WinActive("Muster 16 ahk_class #32770")
;:*:Ibu600::Ibuprofen 600{F3}
:*:Ibu 600::Ibuprofen 600{F3}
:*:Ibu 800::Ibuprofen 800{F3}
:*:Novamin::Novaminsulfon{F3}
#If
;}
; --- mehr                                                                            	;{
#If WinActive("Termin Memo ahk_class #32770")
:*:1I::1. Impfung
:*:1.I::1. Impfung
:*:1. I::1. Impfung
:*:2I::2. Impfung
:*:2.I::2. Impfung
:*:2. I::2. Impfung
:*:3I::3. Impfung
:*:3.I::3. Impfung
:*:3. I::3. Impfung
:*:4I::4. Impfung
:*:4.I::4. Impfung
:*:4. I::4. Impfung
:*C:Ko::Kontrolle{Space}
:*C:Spr::zur Spritze{Space}
#If
;}
;}

;{ 	6.2 -- TELEGRAM
#IfWinActive Telegram ahk_class Qt5QWindowIcon
:C1:HASH::                                                                                								;{
	SendHashTagNachricht()
return
SendHashTagNachricht() {

	Iniread, hashtagNachricht, % Addendum.ini, % "Addendum", % "HASHtagNachricht"

	DA_PatID := AlbisAktuellePatID()
	If DA_PatID is not number
		DA_PatID:=""

	DA_PatName:= AlbisCurrentPatient()
	IB_text =
	(
	Wenn im Eingabefeld eine Nummer steht, hat das Skript hat die PatientenID
	der geöffneten Akte von: %DA_PatName% ausgelesen.
	Falls der Telegram-Patient und die ID zusammen passen, drücke `"Ok`" oder
	einfach nur die Enter-Taste. Den Rest übernimmt das kleine Script.
	)

	InputBox, HASHtag, Albis ID Nachfrage, % IB_text,, 500, 250,,,,, % "#" DA_PatID

	rautepos := InStr(HASHtag, "#"), mrautepos := InStr(HASHtag, "#", 1, 2)
	If (rautepos = 0)
		HASHtag := "#" HASHtag
	else if (rautepos > 1) || (mrautepos <> 0)
		HASHtag := "#" StrReplace(HASHtag, "#")

	WinActivate, Telegram ahk_class Qt5QWindowIcon
	WinWaitActive, Telegram ahk_class Qt5QWindowIcon,, 2
	SendRaw, % HASHtag
	Send, {Enter}

	SendRaw, % StrReplace(hashtagNachricht, "##", "`n")

	VarSetCapacity(bestelltext, 0)
	SendInput, {LControl Down}{Enter}{LControl UP}


}
;}
Enter::                                                                                										;{
if (A_ThisHotkey = A_PriorHotkey && A_TimeSincePriorHotkey < 200)
		SendInput, {LControl Down}{Enter}{LControl UP}
else
		SendInput, {Enter}
return
;}
:*:---::- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
#IfWinActive ;}

;{ 	6.3 -- SciTE4AutoHotkey
#IfWinActive, ahk_class SciTEWindow
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
;                                                                  	     optische Trennung
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣;{
:*:.t1::	                                                                                                                    	;{ oberer einzelner Trenner 	------
SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------"
return ;}
:*:.t2::	                                                                                                                    	;{ mittlerer Trenner            	------
SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------;{"
return ;}
:*:.t3::                                                                                                                       	;{ alle 3                              	------
	SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------"
	SendInput, {Enter}
	SendInput, % "{Raw};"
	SendInput, {Space}
	SendInput, {Enter}
	SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------;{"
	SendInput, {Enter}
	SendInput, {Up 2}
	SendInput, {End}
return ;}
:*:.t4::; ⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽⎽
:*:.t5::; ⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺⎺
;}
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
;                                                                          	 SciteOutput
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣;{
; __ SciteOutPut __
:*:.so::SciTEOutput(){Left 1}
:*:.sa::SciTEOutput(A_ThisFunc " (" A_LineNumber "):`n" ){Left 1}
:*:.scout::SciTEOutput(A_ScriptName "(" A_LineNumber "): "){Left 2}
:R*:.sout::SciTEOutput(A_LineNumber ": "){Left 2}
:R*:.sciteout::SciTEOutput(A_LineNumber ": "){Left 2}
:R*:.scstd::SciTEOutput(A_LineNumber ": "){Left 2}
:*:.sot::
	SciteWrite("SciteOutPut(qq<clipText>: qq <clipText>, 0, 1)", true)
return ;}
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
;                                                                            	 Tasten
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣;{
:R*:.kl::{Left}
:R*:.kr::{Right}
:R*:.kcd::{LControl Down}
:R*:.kcu::{LControl Up}
:R*:.hcdu:: ;{
	Send, % "{Text}{LControl Down}{LControl Up}"
	Send, {Left 13}
return ;}
:R*:.ksd::{LShift Down}
:R*:.ksu::{LShift Up}
:R*:.ksdu:: ;{
	Send, % "{Text}{LShift Down}{LShift Up}"
	Send, {Left 11}
return ;}
:R*:.kad::{LAlt Down}
:R*:.kau::{LAlt Up}
:R*:.kadu:: ;{
	Send, % "{Text}{LAlt Down}{LAlt Up}"
	Send, {Left 9}
return ;}
:R*:.kwd::{Win Down}
:R*:.kwuk::{Win Up}
:*:.kwdu:: ;{
	Send, % "{Text}{Win Down}{Win Up}"
	Send, {Left 8}
return ;}
;}
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
;                                                                          	     MsgBox
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣;{
:*:.msg::                                                                                                                    	;{ Addendum Messagebox
	Send, % "{Text}MsgBox, 1, Addendum für Albis on Windows, % """""
return ;}
:*:.mxb::
msg := "MsgBox, 0x1000, % " q "Addendum - " q " StrReplace(A_ScriptName, " q ".ahk" q ") , % " q q
SendRaw, % msg
SendInput, {Left}
return
;}

:*:.aid::                                                                                                                     	;{ % "ahk_id "
	SendInput, % "{Raw}% " q "ahk_id " q " "
return ;}
:R*:%id::% "ahk_id "

:*:.as::A_ScriptDir
:*:.as1::                                                                                                                     	;{  A_ScriptDir "\"
	SendInput, % "{Raw}A_ScriptDir\"
	SendInput, {Left}
return ;}

:*:.afa::Addendum für Albis on Windows
:*:.addendum::Addendum für Albis on Windows

:*:.fo::FileOpen(, "r", "UTF-8").Read(){Left 22}
:*:.fw::FileOpen(, "w", "UTF-8").Write(){Left 22}
:*:.jl::cJSON.Load(FileOpen(filepath, "r", "UTF-8").Read()){Left 23}{LShift Down}{Left 8}{LShift Up}
:*C:.jw::FileOpen(filepath, "w", "UTF-8").Write(cJSON.Dump(oJSON, 1)){Left 43}{LShift Down}{Left 8}{LShift Up}
:*C:.JW::FileOpen(A_Temp "\tmp.json", "w", "UTF-8").Write(cJSON.Dump(oJSON, 1)){Enter}Run, % A_Temp "\tmp.json"{Up}{End}{Left 5}{LShift Down}{Left 5}{LShift Up}
:*:.SendD::SendDiagnosis(`"
:*R:\pLu::\p{Lu}
:*R:.pLu::\p{Lu}
:*R:\pLl::\p{Ll}
:*R:.pLl::\p{Ll}


#IfWinActive

SciteWrite(text, ReplaceFromClipBoard:= false) {

   	clipText := Clipboard
	clipText := StrReplace(clipText, "`n`r")
	text   	:= StrReplace(text, "qq", q )

	If StrLen(clipText) = 0 {
		PraxTT("Das Clipboard ist leer.", "3 3")
	}

	ControlGetFocus, fcontrol, % hwnd:= WinExist("A")
	ControlGetText, cText, % fcontrol, % "ahk_id " hwnd

	If StrLen(cText) = 0 {
		PraxTT("Der Inhalt des SciteFenster`nkonnte nicht ausgelesen werden.", "1 3")
		return
	}

	If ReplaceFromClipBoard && RegExMatch(cText, "i)\s+" clipText "\s*\:\=") ; checkt das es eine Variable ist
		text := StrReplace(text, "<clipText>", clipText)

	SendInput, % "{Raw} " text
}
;}

;{ 	6.4 -- ÜBERALL

;{     6.4.1 -- Bestellungen
:*:.Bestellungen::
;~ IniRead, bestelltext, % Addendum.Ini, Addendum, Bestellung
For each, line in StrSplit(IniReadExt("Addendum", "Bestellung"), "##") {
	If (StrLen(line) > 0)
		Send % "{Raw}" line
	else
		Send {Enter}
	Sleep 25
}
bestelltext := line := ""
return
;}

;{    	6.4.2 -- sonstige Kürzel

; - - - - - - - - - - - - - - - - -
; Programmieren
:*:#ahk::Autohotkey
:*:#lahk::language:Autohotkey
:*:#sahk::site:Autohotkey

; - - - - - - - - - - - - - - - - -
; Mail/Briefunterschriften
:*:#KV::                                                                     	;{ Unterschrift für Mail oder Briefe an KV
	If Addendum.Praxis.KVStempel
	  For sendNr, sendLine in StrSplit(Addendum.Praxis.KVStempel, "##")
		SendRawFast(sendLine, "Enter")
return ;}
:*:#Mail::                                                                   	;{ Unterschrift für Mail oder Briefe an Firmen, Patienten (ohne BSNR und LANR)
	If Addendum.Praxis.MailStempel
	  For sendNr, sendLine in StrSplit(Addendum.Praxis.MailStempel, "##")
		SendRawFast(sendLine, "Enter")
return ;}
:X*:#BSNR::SendRawFast(Addendum.Praxis.Arzt[Addendum.Praxis.StandardArzt].BSNR)
:X*:#LANR::SendRawFast(Addendum.Praxis.Arzt[Addendum.Praxis.StandardArzt].LANR)
:X*:#Tel1::
:X*:#Telefon1::SendRawFast(Addendum.Praxis.Telefon[1])
:X*:#Tel2::
:X*:#Telefon2::SendRawFast(Addendum.Praxis.Telefon[2])
:X*:#Fax1::SendRawFast(Addendum.Praxis.Fax[1])

; - - - - - - - - - - - - - - - - -
; Datum
:*:#vorgestern::
:*:#gestern::
:X*:#heute::
:*:#morgen::
:*:#übermorgen::  ;{
	RegExMatch(A_ThisHotkey, ":(?<Option>.*?):(?<String>.*?)::", thisHot)
	datestamp 	:= A_YYYY A_MM A_DD
	daysAndAdd	:= {"#vorgestern": -2, "#gestern": -1, "#heute": 0, "#morgen": 1, "#übermorgen": 2}
	daysToAdd 	:=  daysAndAdd[thisHotString]
	datestamp += %daysToAdd%, days
	SendRawFast(ConvertDBASEDate(datestamp))
return
:X*:#dd::SendRawFast(A_DD)
:X*:#mm::SendRawFast(A_MM)
:X*:#yy::SendRawFast(A_YYYY)
:X*:#dm::SendRawFast(A_DD "." A_MM)
:X*:#d.m::SendRawFast(A_DD "." A_MM ".")
;}

:*:akutell::aktuell

; - - - - - - - - - - - - - - - - -  #brief
; Sonderzeichen
ChatHotstrings() {
	static vDummy := ChatHotstrings()
	Hotstring("EndChars", "'-{}[]'/\`n `t")
	Hotstring(":X*::)", "HotS_Writer")
	Hotstring(":X*::)", "HotS_Writer")
	Hotstring(":X*::(", "HotS_Writer")
	Hotstring(":X*:8)", "HotS_Writer")
	Hotstring(":X*:;)", "HotS_Writer")
	Hotstring(":X*::)", "HotS_Writer")
	Hotstring(":X*::|", "HotS_Writer")
	Hotstring(":X*::(", "HotS_Writer")
	Hotstring(":X*::o", "HotS_Writer")
	Hotstring(":X*:??", "HotS_Writer")
	Hotstring(":X*:!!", "HotS_Writer")
	Hotstring(":X*:#kreisvoll", "HotS_Writer")
	Hotstring(":X*:#kreisleer", "HotS_Writer")
	Hotstring(":X*:#kreisstern", "HotS_Writer")
	Hotstring(":X*:#brief", "HotS_Writer")
	Hotstring(":X*:#smile", "HotS_Writer")
	Hotstring(":X*:#smily", "HotS_Writer")
	Hotstring(":X*:#stethoskop", "HotS_Writer")
	Hotstring(":X*:#pflaster", "HotS_Writer")
	Hotstring(":X*:#Herz", "HotS_Writer")
	Hotstring(":X*:#Kuss", "HotS_Writer")
	Hotstring(":X*:#Stern", "HotS_Writer")
	Hotstring(":X*:#Sonne", "HotS_Writer")
}
HotS_Writer() {

	static	sonderzeichen := {"kreisvoll":"●", "kv":"●", "kreisleer":"⚬", "kreisstern":"⊛", "brief":"✉", "smile":"☻", "smily":"☻", "Herz":"♡", "Stern":"★","Sonne":"🌞", "Kuss":"👄"
			    						, ":(":"☹", ":|":"😐", ":o":"😮", ";)":"😉","8)":"😎", ":)":"😁", "kreiszahl":"①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲"
				    					, "&^&":"🙵", "Stethoskop":"🩺", "pflaster":"🩹", "??":"❔", "!!":"❕"}


	; ⚬⊛✉☻☻♡★🌞👄☹😐😮😉😎😁🩹🩺🙵
	hotstring := A_ThisHotkey
	RegExMatch(hotstring, "O)\:(?<Option>.*)#*(?<String>.*)", match)
	SciTEOutput(match.String ": " sonderzeichen[match.String])
	If (StrLen(sz:=sonderzeichen[match.String]) = 1)
		ordnum := Ord(sZ)
	SendRaw, % (ordnum ? chr(ordnum) : sz)

}
;}

;{ 	6.4.3. -- Brüche
::1/2::½
::1/3::⅓
::2/3::⅔
::1/4::¼
::3/4::¾
::1/5::⅕
::2/5::⅖
::3/5::⅗
::4/5::⅘
::1/6::⅙
::5/6::⅚
::1/7::⅐
::1/8::⅛
::3/8::⅜
::5/8::⅝
::7/8::⅞
::1/9::⅑
::1/10::⅒
;}

;{     6.4.4. -- Autokorrektur

:*:Syndorm::Syndrom
:*:akutell::aktuell


;}

;}

;}

;}

;----------------------------------- End of Autoexecute ----------------------------------------------------------------------------------


;{10. Hotkey/Hotstring Labels und Funktionen

;{ ###############   	HOTKEY-FUNCTIONS   ###############
; - - - - - - - - - - - - - - - Hotkey, IF, ****
HoveredControl(WinNN, CtrlNN)                                                    	{              	;-- true wenn Maus über benanntem Steuerelement steht

	MouseGetPos, mx, my, hWin, hCtrl, 2
	WinGetClass , WinClass, % "ahk_id " hWin
	If !InStr(WinClass, WinNN)
		return false
	If RegExMatch(Control_GetClassNN(hWin, hCtrl), CtrlNN)
		return true

return false
hCtrlOff:
	ToolTip, % "classnn: " fCtrl, % mx-10, % my+30, 12
	SetTimer, hCtrlOff, -1000
	ToolTip,,,,12
return
}
IsDatumsfeld()                                                                               	{                 	;-- ist der Eingabefokus in einem Datumsfeld

	; zur Identifikation des fokussierten Steuerelementes (individuelle Positionierung des Kalenders) werden einige Daten eingesammelt)
	; letzte Änderung: 13.06.2022

		global CalFormulare
		If !IsObject(CalFormulare) {
			CalFormulare	:= {"Muster 1a"     	: {"cnt":1	, "ctrls"  	: "1|2|3"
																			    	, "pairs" 	: ":1,2:"}
										, "Muster 12a"    	: {"cnt":2	, "ctrls"  	: "6|7|12|13|17|18|22|23|27|28|32|33|37|38|"
																						    		. "43|44|53|54|58|59|64|65|70|71|77|78|82|83"
																					, "pairs" 	: ":6,7:12,13:17,18:22,23:27,28:32,33:37,38:43,44:53,54:58,59:64,65:70,71:77,78:82,83:"
																					, "group"	: [{"from":[6,12,17,22,27,32,37,43,53,58,64,70,77,82], "to":[7,13,18,23,33,38,54,59,65,71,78,83]}]}
										, "Muster 17"     	: {"cnt":3	, "ctrls"   	: "1|2|3|4|5|6|7"
																					, "pairs" 	: ":1,2:6,7:"}
										, "Muster 20"     	: {"cnt":4	, "ctrls"    	: "6|7"}
										, "Muster 21"     	: {"cnt":6	, "ctrls"    	: "3|4|7|8|11|12|15|16"}
										, "Privat-Au"       	: {"cnt":7	, "ctrls"    	: "3"}
										, "Daten von"     	: {"cnt":8	, "ctrls"    	: "29|31"}
										, "System-Daten"	: {"cnt":9	, "ctrls"    	: "9|10|11|12|13"}}
		}

		feldOk    		:= 0
		If !(WinTitle  	:= WinGetTitle(hactive := WinExist("A")))
			WinTitle	:= WinGetText(hactive)
		cContent 		:= ControlGetText("", "ahk_id " (hfocused := GetFocusedControlHwnd()))
		focusedNN 	:= Control_GetClassNN(hactive, hfocused)

	  ; kein Edit-Steuerelement - Abbruch
		If !RegExMatch(focusedNN, "i)(?<=Edit)\d+", classnn)
			return 0

	  ; weitere Daten auslesen
		RegExMatch(cContent, "^\d{2}\.\d{2}\.\d{4}$", Datum)
		If !RegExMatch(WinTitle, "i)Muster\s+\w+", wtitle)
			RegExMatch(WinTitle, "i)(?<title>Daten\s+von|System\-Daten)", w)

	  ; kein Datumsfeld dann wird hier abgebrochen
		If !(feldOk := CalFormulare.Haskey(wtitle) && RegExMatch(classnn, "(" CalFormulare[wtitle].ctrls ")") ? CalFormulare[wtitle].cnt : Datum ? 5 : 0)
			return 0

	  ; ist ein (von - bis) Datumfeld angewählt
		RegExMatch(CalFormulare[wtitle].pairs, "O):(\d+)," classnn "|" classnn ",(\d+):", pair)
		If pair.Count()
			to := pair.2 ? pair.2 : classnn, from := pair.1 ? pair.1 : classnn

	 ; Datenobjekt erstellen und zurückgeben
		return {	"Ok"      	: feldOk
				, 	"hwnd"  	: hfocused
				, 	"hfocused"	: hfocused
				, 	"hWin"   	: hactive
				,	"to"        	: to
				,  	"from"		: (from    	? from    	: classnn)
				, 	"Datum"	: (Datum 	? Datum  	: FormatTime(A_Now, "ddMMyyyy"))
				, 	"WinTitle"	: (wtitle  	? wtitle   	: WinTitle)
				, 	"focusNN"	: (classnn 	? classnn 	: focusedNN)
				,	"focusFD"	: (pair.1 	? "to"      	: pair.2 ? "from" : "other")}

}
IFActive_Lko()                                                                                	{              	;-- #IF Hotkey/Hotstring Funktion - lko/Karteikarte aktiv?
	global KK
	If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveWindowType(), "Karteikarte")
		If RegExMatch(AlbisGetActiveControl("contraction"), "i)^lk\w")
			return true
return false
}
IFActive_UpDown()                                                                         	{              	;-- #IF Hotkey/Hotstring Funktion - Listviewfenster mit Höher/Tiefer
	static UpDownWindows := ["Dauerdiagnosen von", "Dauermedikamente", "Cave! von", "Familiendaten von", "Abrechnungsassistent"]
	For each, wtitle in UpDownWindows
		If WinActive(wtitle " ahk_class #32770")
			return true
return false
}
ItemUpDown()                                                                               	{              	;-- für alle Albisfenster mit verschiebbaren Listvieweinträgen

	static UDWinTitle

	WinGetActiveTitle	, atitle
	WinGetClass     	, aclass, % "ahk_id " (whwnd := WinExist("A"))
	RegExMatch(A_ThisHotkey, "i)(?<key>Up|Down)", h)
	PostMessage, 0x00F5,,, % "Button" (InStr(atitle, "Familiendaten von") ? (hkey="Up" ? "7" : "8") : hkey="Up" ? "1" : "2"), % "ahk_id " whwnd
	UDWinTitle := atitle " ahk_class " aclass
	SetTimer, ItemUpDown_Refocus, -50

return

ItemUpDown_Refocus:
	If WinExist(UDWinTitle)
		ControlFocus, SysListView321, % UDWinTitle
	UDWinTitle := ""
return
}

; - - - - - - - - - - - - - - - Hotkey-Funktion
EnableAllControls(ask=true)                                                             	{              	;-- Zugriff für alle Steuerelemente eines Fenster ermöglichen

	hwnd := WinExist("A")

	if ask {
		WinGetTitle, wtitle, % "ahk_id " hwnd
		MsgBox, 4, % "Addendum für Albis on Windows"
						, % "Sollen im Fenster mit dem Titel: `n" "<< '" wtitle " >>`nalle Felder editierbar und sichtbar gemacht werden?"
		IfMsgBox, No
			return
	}

	For key, control in GetControls(hwnd,"","", "Enabled,hwnd,Visible") {
		If !control.enabled
			Control, Enable,,,	% "ahk_id " control.hwnd
		If !control.Visible
			Control, Show,,,	% "ahk_id " control.hwnd
	}

}
FoxitReaderSignieren()                                                                     	{                	;-- macht ein Backup der PDF Datei bevor diese signiert wird

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		If InStr(process.name, "FoxitReader") && RegExMatch(process.CommandLine, "\s\" q "(.*)?\" q, cmdline)
			break

	SplitPath, cmdline1, PDFName
	FileCopy, % cmdline1, % Addendum.BefundOrdner "\Backup\" PdfName, 1

	FoxitReader_SignaturSetzen()

}
HotkeyViewer()                                                                              	{               	;-- zeigt Hotkeys

	; Hotkey, IfWinActive, % "ahk_class " Addendum.PDF.ReaderWinClass
	rxhotkey 	:= "Oi)^\s*Hotkey\s*,\s*(?<Case>If.+)*(?<Key>[a-z\d\+\#\!\^]+)*\s*?,\s*(?<Func>.+)\s*?(;--\s*)(?<Description>.+)*$"
	WinNN 	:= {"ahk_class OptoAppClass":"Albis", "Dauerdiagnosen von ahk_class #32770":"Dauerdiagnosen"
						, "Dauermedikamente von ahk_class #32770":"Dauermedikamente", "Cave! von ahk_class #32770":"Cave" }
	allHotK 	:= ""

	For lineNr, line in FileOpen(A_ScriptFullPath, "r", "UTF-8").Read() {

		If RegExMatch(line, rxhotkey, Hot) {

				If Hot.Case {

					Switch Hot.case {

						Case "IfWinActive":
							IfCase := "bei aktiviertem ## Fenster"
						Case "IfWinExist":
							IfCase := "bei existierendem ## Fenster"
						Case "If":
							IfCase := "wenn Sonderbedingung"

					}
				}
		}

	}
}
UnRedo(Do)                                                                                   	{                	;-- Undo und Redo in den Textfeldern im Albisfenster

	dbg := false

	;WM_UNDO = 0x0304
	ControlGetFocus, focused, % "ahk_id " (AlbisWinID := AlbisWinID())
	If RegExMatch(focused, "i)Edit|RichEdit|ComboBox") {

			ControlGet       	, hfocused , Hwnd,, % focused, % "ahk_id " AlbisWinID
			ControlGetText	, text                 	, % focused, % "ahk_id " AlbisWinID

			If (Do = "Undo") {
				If ( Addendum.UndoRedo.OnUnRedo <> GetHex(hfocused) ) {
					Addendum.UndoRedo.OnUnRedo := GetHex(hfocused)
					Addendum.UndoRedo.List := Array()
				}
				Addendum.UndoRedo.List.InsertAt(1, text)
				PostMessage, 0x0304,,, % focused, % "ahk_id " AlbisWinID
			}
			else if (Do = "Redo") {
				If Addendum.UndoRedo.MaxIndex()
					ControlSetText,, % (text := Addendum.UndoRedo.List.Pop()), % "ahk_id " Addendum.UndoRedo.OnUnRedo
			}

			dbg ? SciTEOutput("UnRedo: " text) : ""

	}

}
DeInkrementer()                                                                            	{              	;-- Steuerelementinhalt um 1 erhöhen oder erniedrigen

	MouseGetPos,,, hWin, hCtrl, 2
	ControlFocus,, % "ahk_id " hCtrl
	ControlGetText, value,, % "ahk_id " hCtrl
	If !RegExMatch(value, "^\d+$") {
		ControlSend,, (A_ThisHotkey="WheelUp" ? {Up}:{Down}), % "ahk_id " hCtrl
		return
	}

	;value := A_ThisHotkey = "WheelUp" ? value + 1 : value > 1 ? value - 1 : value
	If (A_ThisHotkey = "WheelUp")
		value ++
	else If (A_ThisHotkey = "WheelDown") && (value > 1)
		value --
	else
		return

	ControlSetText  	,, % value	, % "ahk_id " hCtrl
	ControlSend     	,, {Enter}	, % "ahk_id " hCtrl
	ControlFocus 	,           	, % "ahk_id " hCtrl
	ControlClick     	,           	, % "ahk_id " hCtrl,, Left, 1
	Sleep 60
	ControlClick     	,           	, % "ahk_id " hCtrl,, Left, 1

}
PatAusweisDruck()                                                                          	{                	;-- Patientenausweis drucken
  ; ruft Menu Formular/Patientenausweis auf. Die Dialogeinstellungen sind derzeit noch über Albis vorzunehmen
  ; der Druck erfolgt über einen Hookaufruf bei Erscheinen des Patientenausweis Dialoges
	Addendum.Flags.PatAusweisDruck := true
	AlbisMenu(33003)
}
AusIndisKorrigieren()                                                                     	{                 	;-- lädt Daten für AlbisAusIndisKorrigieren()

 ; letzte Änderung: 03.04.2022

 ; die Einstellungen werden beim ersten Aufruf der Funktion gelesen
	If !IsObject(Addendum.AusIndis)
		Addendum.AusIndis := Object()
	If !Addendum.AusIndis.ToRemove {
		ToRemove := IniReadExt("Abrechnungshelfer", "AusIndis_ToRemove")
		ToRemove := InStr(ToRemove, "ERROR") ? "" : ToRemove
		ToRemove := RegExReplace(ToRemove, "^\s\(")
		ToRemove := RegExReplace(ToRemove, "^\s\)")
		ToRemove := RegExReplace(ToRemove, "[\,\.\s\;\:\-\+]", "|")
		Addendum.AusIndis.ToRemove := ToRemove ? ("(" RegExReplace(ToRemove, "[\\]{2,}", "|") ")") : ""
	}
	If !IsObject(Addendum.AusIndis.ToHave) {
		If (Addendum.AusIndis.iHateBMPs := !IniReadExt("Abrechnungshelfer", "AusIndis_Ich_mag_BMPs", "Ja")) {
			Addendum.AusIndis.BMP := IniReadExt("Abrechnungshelfer", "AusIndis_BMP", "Ja")
			Addendum.AusIndis.BMP := InStr(Addendum.AusIndis.BMP, "Error") && Addendum.AusIndis.iHateBMPs ? false : true
		}
		ToHave := IniReadExt("Abrechnungshelfer", "AusIndis_ToHave")
		If RegExMatch(ToHave, "\d{5}(\||$)") {
			Addendum.AusIndis.ToHave := []
			For idx, EBMZiffer in StrSplit(ToHave, "|")
				If RegExMatch(EBMZiffer, "\d{5}")
					Addendum.AusIndis.ToHave.Push(EBMZiffer)
		}
	}

  ; keine Daten um Korrektur durchzuführen
	If (!IsObject(Addendum.AusIndis.ToHave) && !Addendum.AusIndis.ToRemove) {
		TrayTip, % A_ScriptName, % "Es sind keine Daten für eine Korrektur von Ausnahmeindikationen`n"
												. "in der Addendum.ini hinterlegt worden.`n"
												. "Die Funktion wurde abgebrochen werden", 4
		return 0
	}

 ; AusIndis-Funktion ausführen
	PatID := AlbisAktuellePatID()
	AusIndis := AlbisAusIndisKorrigieren(	Addendum.AusIndis.ToRemove
														, 	Addendum.AusIndis.ToHave
														, 	"ChronischKrank=" 	Chroniker
														. 	" Geriatsch="        	GB
														. 	" KeinBMP="         	Addendum.AusIndis.iHateBMPs)

  ; Hinweis zu durchgeführten oder nicht durchgeführten Änderungen anzeigen
	If AusIndis.neu {
		RegExReplace(AusIndis.Alt  	, "\-",, AltC)
		RegExReplace(AusIndis.Neu	, "\-",, NeuC)
		inter := AusIndis.Alt ? (AltC > 0 ? "n" : "") " ersetzt mit" : "n erstmalig eingetragen"
		msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" AusIndis.Alt "`nwurde" inter "`n" AusIndis.Neu
	}
	else
		msg := "Ausnahmeindikation" (AltC > 0 ? "en" : "") ":`n" AusIndis.alt "`nwurde" (AltC > 0 ? "n" : "") " nicht geändert"

	PraxTT("Bitte überprüfen!`n⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘`n"  msg "`n⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘ ⁘", "6 0")

}


;}
;{ ###############  HOTSTRING-FUNCTIONS  ###############
HotS_IF(context)                                                                     	{              	;-- Hotkey/Hotstring - kontextabhängige Überprüfung

	; ◈ ᗅᏞВІЅ ◈
	If RegExMatch(context, "(Diagnosen|EBM|GOÄ)") {
		If !WinActive("ahk_class OptoAppClass")
			return false
	}
	else {

		contraction := AlbisGetActiveControl("contraction")
		Switch context {

			case "Diagnosen":
				If InStr(contraction, "dia") || WinActive("Akutdiagnose ahk_class #32770") || WinActive("Dauerdiagnosen ahk_class #32770")
					return true
				If WinActive("Muster 2 ahk_class #32770") {
					ControlGetFocus, cfocus, % "Muster 2 ahk_class #32770"
					If (cFocus = "Edit1")
						return true
				}

			case "EBM":
				If WinActive("ahk_class OptoAppClass") && InStr(contraction, "lk")
					return true

			case "GOÄ":
				If WinActive("ahk_class OptoAppClass") && (RegExMatch(contraction, "lp|lbg") || InStr(AlbisGetActiveWindowType(), "Privatabrechnung"))
					return true

		}

	}

}
HotS_View(ALineNr, context:="")                                                     	{              	;-- kontextsensitive Hotstringhilfe

	; 	⚫ letzte Änderung: 25.11.2022
	; 	Funktion liest die Hotstringdaten bei jedem Funktionsaufruf neu aus.

	;{ 🔴 Variablen 🔶🟨

		global hHSV, HSV_LV, HSV, HSV_B1, HSV_NHS, HSV_NRP, HSV_NOP, HSV_PM, HSV_NDS, NHSLast, NRPLast, NOPLast
		global hsUser

		static hsContext, hsBase, hotstrings, cFocused, wFocused, hsStrings, hsSearched, lastSelected, ICDAll, hsFileTimeLast
		static icdpath
		static dbg           	:= false

		; GOÄ / EBM
		static rxFee        	:= "O)^\s*\:(?<options>.*?):(?<abbreviation>.*?)\:\:(?<replacement>.*?)\s*\t*;[\s\{\}\=\~]*(?<description>.*)$"
		static rxSendRpl    	:= "Oi)^\s*(?<send>Send.*?(\{[a-z]+\})*)(?<replacement>[\d\-\(\):\pL]+)"
		static hsRemover 	:= {	"[\(\[\{].*[\}\]\)].*"          	: " "      	; z.B. [Punktzahl 200]
							    	, 		"\s\d+\s.*"                      	: " "    	; z.B. 30 Minuten
									, 		"i)([a-zäöüß]+\.)([\s;,]*)"   	: " "		; z.B. mind. 10 min
									, 		"i)[^a-zäöüß\s\-]"            	: " "     	;
									, 		"\s\-\s"                            	: " "
									, 		"\s{2,}"                         	: " "}

		If !icdpath {
			Pathstr := Addendum.DataPath "\icd10gm"
			Loop 3 {
				ICDYear := A_YYYY - A_Index + 1
				If FileExist(Pathstr . ICDYear ".txt") {
					icdPath := Pathstr . ICDYear ".txt"
					break
				}
			}
			ICDYear := icdPath ? ICDYear : ""
			If dbg
				SciTEOutput("ICDPAth: " icdPath "`nICDJAhr: " ICDYear)
		}

		hotstrings	:= {}, replacements := {}, hsNr := 0, hsBase := context
		hsContext	:= hsBase ~= "i)(EBM|GOÄ)" 	? 1
						:	 hsBase ~= "i)(Diagnose)" 	? 2 : 0

	  ; Daten des fokussierten Editsteuerelementes auslesen
		ControlGetFocus	, cFocused                             		, % "ahk_id " (wFocused := WinExist("A"))
		ControlGetText  	, cText       	, % cFocused         	, % "ahk_id " wFocused
		ControlGet         	, hcFocused	, hwnd,, % cFocused	, % "ahk_id " wFocused
		FPos 	:= GetWindowSpot(hcFocused)                         	; Bildschirmposition des fokussierten Steuerelements
		CaretP 	:= CaretPos(hcFocused)                                   	; Posiiton des Zeichens hinter dem das Caret steht
		acx    	:= A_CaretX, acy := A_CaretY                           	; Bildschirmposition des Caret

	  ; Hotstrings aus dem Skript lesen
		scriptlines	:= StrSplit(FileOpen(A_ScriptFullPath, "r", "UTF-8").Read(), "`n", "`r")
		scriptlines.RemoveAt(1, ALineNr)                                   	; entfernte alle Skriptzeilen bis ALineNr

	  ; Diagnosen: löschen der Eingabe
		If (hsContext = 2)  {                     	; Diagnosen

		  ; nur Text bis zur Position des Caret untersuchen
			CaretText 		:= SubStr(cText, 1, CaretP)
			SPos         	:= RegExMatch(CaretText, "i)(?<Start>^|\.|[\};\s*]+)\s*(?<Searched>[\pL\s\+\-,]+?)\s*(\*|\?|$)", hs)
			clen          	:= StrLen(hsSearched)
			hsSearched 	:= StrReplace(hsSearched, ".")

		  ; Eingabe entfernen
			Loop % clen + (hsStart="." ? 1:0) {
				ControlSend,, {BS}, % "ahk_id" hcFocused
				Sleep 5
			}

			If dbg {
				SciTEOutput(	"`nwFocused:   `t"        	wFocused
								. 	"`ncFocused:`t"            	cFocused
								. 	"`nhcFocused:`t"         	hcFocused
								. 	"`nacx:         `t"           	acx
								. 	"`nacy:         `t"            	acy
								. 	"`nCaretP:     `t"         	CaretP
								. 	"`ncText:      `t"            	cText
								. 	"`nCaretText:`t"           	CaretText
								. 	"`nhsSearched:`t"      	hsSearched
								. 	"`nhsPos:   `t"             	SPos
								.	"  clen: "   	                	clen)
			}

		}

	;}

	;{ 🟣 Hotstrings aus dem Skriptcode parsen
		AuswahlboxSection := false
		For lineNr, line in scriptLines {

			If RegExMatch(line, "i)(\s*;\s*###|\s*#If)")
				break

		; Hotstrings aus dem Skriptcode lesen
		; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
		;    GOÄ / EBM   - -   GOÄ / EBM   - -   GOÄ / EBM   - -   GOÄ / EBM
		; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
			If (hsContext = 1) {

				If RegExMatch(line, rxFee, hs) {

				  ; RegEx - ein Objekt läßt sich nicht bearbeiten - warum?
					hs.abbreviation	:= Trim(hs.abbreviation)
					hs.replacement	:= Trim(RegExReplace(hs.replacement, "\s{2,}"))
					hs.options        	:= StrReplace(hs.options, "*")
					hsDescription 	:= hs.Description

				  ; entfernt Sonderzeichen und anderes überflüssiges
					For rxStr, rxRpl in hsRemover
						hsDescription := RegExReplace(hsDescription, rxStr, rxRpl)

				  ; Hotstringbeschreibungen kürzen (thematische Sortierung)
					hsKey := "", upperChar := false
					For wIndex, word in StrSplit(Trim(hsDescription), A_Space)
						If RegExMatch(word, "^und") {
							hsKey .= word " "
						} else If RegExMatch(word, "^[a-zäöü]") {
							If upperChar
								break
							hsKey .= word " "
						} else if RegExMatch(word, "^[A-ZÄÖÜ]") {
							upperChar := wIndex = 1 ? false : true
							hsKey .= word " "
						}
					hsKeySaved := hsKey := Trim(hsKey)

				  ; alle Gruppierungen und Ersatztexte sammeln
					If (!replacements.haskey(hsKeySaved) && hs.replacement)
						replacements[hsKeySaved] := hs.replacement

				 ; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
				 ; Hotstrings nach Beschreibung sortieren
					; ▹ - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					; erste Gruppe anlegen oder einer Gruppe hinzufügen
					If (!hotstrings[hsKey].Count() || hotstrings.haskey(hsKey)) {

						If !hotstrings[hsKey].Count()
							hotstrings[hsKey] := {}

						hotstrings[hsKey][hs.abbreviation] := {"description"		: hs.description
																				,"replacement"	: hs.replacement
																				,"options"	    	: hs.options}
					}

					; ▹ - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					; Gruppierung suchen, Gruppierungen anpassen
					else {

					  ; 1. Wort für Wortvergleich der bekannten Beschreibungen mit neuen Beschreibung.
					  ; 2. Der Index des letzten übereinstimmenden Wortes innerhalb der Wortkette wird gespeichert.
					  ; 3. Die längste gemeinsame Wortkette wird zur Gruppierung (hskey) verwendet.
					  ; letzte Änderung: 30.10.2021
						HSIndex := []
						newhskey := hskey
						newhskeyWords := StrSplit(newhskey, A_Space)
						For hskey, hotstring in hotstrings {
							lastMatchedPos := 0
							For wIndex, word in StrSplit(hsKey, A_Space)
								If (word = newhsKeyWords[wIndex])
									lastMatchedPos := wIndex, HSIndex.InsertAt(lastMatchedPos, hsKey)
								else
									break
						}

					  ; Index der längsten passenden Wortkette
						MaxIndex := HSIndex.Count() = 0 ? 0 : HSIndex.MaxIndex()

					  ; neuen Key anlegen bei fehlender Übereinstimmung
						If !MaxIndex {
							hotstrings[newhsKey] := {}
							hotstrings[newhsKey][hs.abbreviation] := {	"description"		: hs.description
																						,	"replacement"	: hs.replacement
																						,	"options"	   		: hs.options}
						}
						else {
							newhskey  	:= ""
							replacedKey	:= hotstrings.Delete(HSIndex[MaxIndex])
							For each, hskey in StrSplit(HSIndex[MaxIndex], A_Space)
								newhskey .= hsKey " "
							newhskey := Trim(newhskey)
							hotstrings[newhskey] := replacedkey
							hotstrings[newhskey][hs.abbreviation] := {	"description"		: hs.description
																						,	"replacement"	: hs.replacement
																						,	"options"		   	: hs.options}
						}

					}

				}
				else if RegExMatch(line, "i)\bSend\b") {
					RegExMatch(line, rxSendRpl, match)
					If !replacements.haskey(hsKeySaved) && match.replacement
						replacements[hsKeySaved] := match.replacement
				}

			}

		; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
		;    Diagnosen    - -    Diagnosen    - -    Diagnosen    - -    Diagnosen
		; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
			else if (hsContext = 2) {

			  ; reguläre Hotstrings
			  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				If RegExMatch(line, "O)\s*\:(?<options>.*):(?<abbreviation>.*)::(?<diagnosis>.*?)(;\{|$)", hs) && !(line~="i)^\s*Auswahlbox\(") {

					AuswahhlboxSection := ""
					hs.diagnosis  	:= Trim(hs.diagnosis)
					hs.diagnosis      := RegExReplace(hs.diagnosis, "\s{2,}", " ")
					hs.diagnosis    	:= RegExReplace(hs.diagnosis, !RegExMatch(hs.diagnosis, ";.*\{.*\}.*$") ? "\s*;*.*$" : "\s*;*\s*$")
					hs.abbreviation	:= Trim(hs.abbreviation)

				  ; Hostringsammler
				   ; - - - - - - - - - - - - - -  -   -    -     -      -       -        -         -          -           -
					If hs.diagnosis && (Ord(hs.diagnosis)<>32) {

						diagnosis      	:= RegExMatch(hs.diagnosis, "Oi)(?<name>[\w\$_]+)\(\" q "(?<diagnosis>.+\{[A-Z].+?\})\" q, fncall)
												? 	 fncall.diagnosis : hs.diagnosis
						abbreviation   	:= hs.abbreviation . abbreviations  	; gesammelte Hotstrings und aktuellen Hotstring zusammenfügen

						If !IsObject(hotstrings[diagnosis])
							hotstrings[diagnosis] := {"abbreviation":abbreviation, "options":hs.options, "func":fncall.name}
						else
							hotstrings[diagnosis].abbreviation .= "●" . abbreviation

						abbreviations  	:= abbreviation := hs := ""

					}
				  ; Abkürzungssammler
				  ; - - - - - - - - - - - - - -  -   -    -     -      -       -        -         -          -           -
					else {

						last_options   	:= hs.options ? hs.options : last_options 	; für späteres Zuweisen der Hotstrings zu den Ersatztexten
						abbreviations 	.= StrLen(hs.abbreviation)>0 ? "●" hs.abbreviation : ""

					}

				}

			  ; Auswahlboxfunktion
			  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				else if RegExMatch(line, "Oi)\s*title\s*\=\s*(?<title>.+?)(" q "|\s\w+\=)", hs)  {

					If !IsObject(hotstrings[hs.title])
						hotstrings[hs.title] := {"abbreviation":LTrim(abbreviations, "●"), "options":last_options, "func":"Auswahlbox"}
					else
						hotstrings[hs.title].abbreviation .= abbreviations

					AuswahlboxSection := hs.title
					abbreviations := hs := ""

				}

			  ; nur innerhalb des Auswahlboxcodes
			  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				else if AuswahlboxSection {

					If (line~="^\s*return") {
						AuswahlboxSection := ""
						hs := ""
						continue
					}

				  ; pro Zeile können im Auswahlboxcode mehrere Diagnosen untergebracht sein, die alle erfasst werden müssen
					xline := RegExReplace(line, "(^|)" q "*\s*,\s*" q, "§§") 	; (^|)"*\s*,\s*" - entfernt alle umrandenden Tabulatoren und Leerzeichen aus dem Code
					For each, diag in StrSplit(xline, "§§") {

						;~ If RegExMatch(diag, "Oi)(?<diagnosis>[\pL\s\*\+\-.,;_#\$\{\}\[\]\(\)\d]+\{[A-Z].+?\})", hs) {
						If RegExMatch(diag, "Oi)(?<diagnosis>.+\{[A-Z].+?\})", hs) {

							;~ SciTEOutput(A_ThisFunc ": " hs.diagnosis)
							If (hs.diagnosis ~= "^\s*$")
								continue
							hs.diagnosis := Trim(hs.diagnosis)

							If !IsObject(hotstrings[hs.diagnosis])
								hotstrings[hs.diagnosis] := {"abbreviation"	: hotstrings[AuswahlboxSection].abbreviation
																		, "options"     	: hotstrings[AuswahlboxSection].options
																		, "func"            	: "Abx#" AuswahlboxSection}

						}

					}

					If RegExMatch(line, q "\s*\]\s*\)")
						AuswahhlboxSection := ""

				}       ; **Ende** 'else if AuswahlboxSection'

			}

		}

	  ; fehlende Ersatztexte zuordnen
		If (hsContext = 1) {
			For hsKey, abbreviations in hotstrings
				For abbreviation, hsdata in abbreviations
					If !hsdata.replacement
						hotstrings[hskey][abbreviation].replacement := replacements[hsKey]
		}

   	  ; ermittelt die Zahl der verwendeten Abkürzungen für alle Diagnosen
		else if (hsContext = 2) {
			diaCnt := 0
			hsabbreviations := Object()
			For hsDia, hs in hotstrings {

				For each, abbrev in StrSplit(hs.abbreviation, "●")
					If !hsabbreviations.haskey(abbrev)
						hsabbreviations[abbrev] := 1

			}
			abbrevCnt := hsabbreviations.Count()
			hsabbreviations := ""

		}


		;~ FileOpen("C:\tmp\hotstrings.json", "w", "UTF-8").Write(cJSON.Dump(hotstrings, 1))
		;~ Run, % "C:\tmp\hotstrings.json"

	;}

	;{ 🔵 Gui

		Gui, HSV: new, +Border +ToolWindow +AlwaysOnTop +HWNDhHSV
		Gui, HSV: Margin, 0, 0
		Gui, HSV: Font, s7 q5

	  ; Listview erstellen
		If (hsContext = 1)                 	{        	; GOÄ/EBM
			Gui, HSV: Add, Listview, % "xm ym h400 vHSV_LV hwndHSV_hLV gHSVHandler Grid ReadOnly AltSubmit -Multi "
												, % "Beschreibung|Ersatztext|Option|Abkürzung"
			For category, hsStrings in hotstrings
				For abbreviation, hsd in hsStrings
					LV_Add("", hsd.description, hsd.replacement, hsd.options, abbreviation) ;
			LV_ModifyCol()
		}
		else If (hsContext = 2)         	{			; Diagnosen

			Gui, HSV: Add, Listview, % "xm ym h400 vHSV_LV hwndHSV_hLV gHSVHandler ReadOnly AltSubmit ", % "Diagnose|Abkürzung"  ; -Hdr
			For dia, hsd in hotstrings
				LV_Add("", dia, hsd.abbreviation)
			LV_ModifyCol()
			LV_ModifyCol(1, 600)
			LVwidth := DllCall("SendMessage", "uint", HSV_hLV, "uint", 4125, "uint", 1, "uint", 0)
			LVwidth := LVwidth>=500 ? 500 : LVwidth
			LV_ModifyCol(2, LVwidth)

		}

		Gui, HSV: Add, Button	, % "xm y+14   	vHSV_B1   	gHSVHandler"	                                 	, % "Auswahl übernehmen"
		Gui, HSV: Add, Edit   	, % "x+5 w50  	vHSV_NOP 	gHSVHandler AltSubmit"                    	, % "*"                                          	; Options
		Gui, HSV: Add, Edit   	, % "x+5 w200 	vHSV_NHS 	gHSVHandler AltSubmit"                   	, % hsSearched                             	; Hotstring
		Gui, HSV: Add, Edit   	, % "x+5 w420 	vHSV_NRP 	gHSVHandler AltSubmit"                   	, % ""                                            	; Ersatztext
		Gui, HSV: Add, Button  	, % "x+5       	vHSV_PM  	gHSVHandler"                                  	, % "☘"

		GuiControlGet, cp, HSV: Pos, HSV_LV
		GuiControlGet, dp, HSV: Pos, HSV_NOP
		Gui, HSV: Add, Text    	, % "x" dpX " y" cpY+cpH " w" dpW " Backgroundtrans Center"        	, % "Optionen"

		GuiControlGet, dp, HSV: Pos, HSV_NHS
		Gui, HSV: Add, Text    	, % "x" dpX " y" cpY+cpH " w" dpW " Backgroundtrans Center"        	, % "Suchen oder Hotstring hinzufügen"

		GuiControlGet, dp, HSV: Pos, HSV_NRP
		dpW := cpW - dpX
		Gui, HSV: Add, Text    	, % "x" dpX " y" cpY+cpH " w" dpW " Backgroundtrans Center"        	, %  (hsContext = 2 ? "Diagnose" : "Ersatztext")

		If (hsContext = 1)                 	{
			GuiControlGet, cp, HSV: Pos, HSV_NOP
			GuiControlGet, dp, HSV: Pos, HSV_NRP
			Y := cpY+cpH, W := dpX+dpW-cpX
			Gui, HSV: Add, Edit  	, % "x" cpX " y" Y+13 " w" W " vHSV_NDS gHSVHandler AltSubmit"	, % ""                                           	; Beschreibung
			Gui, HSV: Add, Text	, % "x" cpX " y" Y " w" W " Backgroundtrans Center"                        	, % "Beschreibung"
		}

		GuiControl, HSV: Enabled0, HSV_B1
		GuiControl, HSV: Enabled0, HSV_PM

	  ; Listview und Guigröße an den Inhalt automatisch anpassen
		LVwidth	:= 0, APos := GetWindowSpot(AlbisWinID())
		SysGet, VScrollWidth   	, 2
		SysGet, CursorHeight	, 14

	 ; Listview Spaltenbreite zusammenrechnen und Listview-/Guibreite für komplette Ansicht ändern
		Gui, HSV: Listview, HSV_LV
		LV_Width := 0
		Loop, % LV_GetCount("Col")
			LVwidth += DllCall("SendMessage", "uint", HSV_hLV, "uint", 4125, "uint", A_Index-1, "uint", 0)
		LVwidth := LVwidth>APos.W ? APos.W : (LVwidth<600 ? 600 : LVwidth)
		GuiControl, HSV: Move, HSV_LV, % "w" LVwidth+2*VScrollWidth

	  ; Anzeigen
		GuiControlGet, Gui, HSV: Pos, HSV_B1
		Gui, HSV: Show	, % "x" acx-5 " y" acy-GuiY-GuiH-CursorHeight  " Hide NoActivate"
								, % "Addendum Hotstringviewer [" hsBase ": " LV_GetCount() (hsContext=2 ? "   Abkürzungen: " abbrevCnt : "") "] "

		GuiControlGet, cp, HSV: Pos 	, HSV_LV
		GuiControlGet, ep, HSV: Pos 	, HSV_PM
		GuiControlGet, fp , HSV: Pos 	, HSV_NRP
		GuiControl          	, HSV: Move	, HSV_NRP, % "w" cpW-(fpX+epW+10)
		GuiControlGet, fp	, HSV: Pos 	, HSV_NRP
		GuiControl          	, HSV: Move	, HSV_PM	, % "x" fpX+fpW+5

	  ; Anzeigeposition korrigieren
		hsvp 	:= GetWindowSpot(hHSV)
		hsvp.Y	:= hsvp.Y < 1 ? acy+CursorHeight : hsvp.Y
		hsvp.X	:= hsvp.X+hsvp.W > FPos.X+FPos.W ? (FPos.X+FPos.W//2)-hsvp.W//2 : hsvp.X
		SetWindowPos(hHSV, hsvp.X, hsvp.Y, hsvp.W, hsvp.H)

	 ; Gui sichtbar machen
		Gui, HSV: Show, % " NoActivate"

	  ; Suchfeld fokussieren
		GuiControl, HSV: Focus, HSV_NHS

	  ; gleich die Suche durchführen, wenn die Eingabe des Sternchen (*) einer Texteingabe folgte
		If hsSearched
			GuiControl, HSV:, HSV_NHS, % hsSearched

		Hotkey, IfWinActive, % "Addendum Hotstringviewer ahk_class AutoHotkeyGUI1"
		Hotkey, ENTER, HSVHandlerEnter
		Hotkey, If

	;}

	;  🟡 ;  ⚪

return

HSVHandlerEnter:  	;{
	Gui, HSV: Submit, NoHide
;}
HSVHandler:		    	;{

	Critical
	hsvp := GetWindowSpot(hHSV)

  ; Doppelklick, Enter oder ein Klick auf das Steuerelement ('Übernehmen' = HSV_B1) sendet Hotstring
	If (A_GuiEvent = "DoubleClick" && A_GuiControl = "HSV_LV") || (A_GuiControl = "HSV_B1") 	{

		diag := ""
		Gui, HSV: Listview, HSV_LV

		If (A_GuiControl = "HSV_B1") {         	; Mehrfachauswahl überprüfen
			row := 0
			while (row := LV_GetNext(row)) {
				LV_GetText(rText, row, hsContext=1 ? 2:1)
				diag .= RegExReplace(rText, ";\s*$") . "; "
			}
		}
		else
			LV_GetText(diag, A_EventInfo, hsContext=1 ? 2:1)

		If diag
			GuiControl, HSV: Enabled0, HSV_B1

	;
		If (hsContext = 2 && diag) {
			If !RegExMatch(diag, "\{.*[A-Z]\d+\.[\d#]+[A-Z]*\}") {         ; ICD Codes finden {E11.90}
				Splash("Kein Ersatztext gefunden.", 8, hsvp.X, Floor(hsvp.Y/2), hsvp.W, hHSV)
				Sleep 2000
				For dia, hsd in hotstrings
					If (dia = diag) {
						diag := hsd.Diagnose
						break
					}
			}
		}

		Gui, HSV: Destroy
		Sleep 50
		ControlFocus, % cFocused, % "ahk_id " wFocused
		Sleep 100
		SendDiagnosis(diag, true)

		hotstrings := hsContext := cFocused := wFocused := hsSearched := diag:= ICDAll := ""


	}

  ; einfacher Klick ins Hotstring Listviewelement
	else If (A_GuiControl = "HSV_LV") 	                                                                                		{

		If  (A_GuiEvent = "Normal") {

		  ; gLabel Events ausschalten
			HotS_gSwitch("Off")

			Gui, HSV: Submit, NoHide

			LV_GetText(NRP, row:=A_EventInfo, (hsContext=1 ? 2 : 1))
			LV_GetText(NHS	, row, (hsContext=1 ? 4 : 2))

			If NHS
				GuiControl	, HSV:, HSV_NHS	, % NHS                                                        	; Hotstring
			If NRP
				GuiControl  	, HSV:, HSV_NRP 	, % NRP                                                        	; Replacement

			For each, guictrl in ["B1", "PM"]
				GuiControl, % "HSV: Enabled" ((NHS||HSV_NHS) && NRP ? 1:0), % "HSV_" guictrl
			GuiControl, % "HSV:", HSV_PM, % (NHS && NRP ? "🗑" : HSV_NHS && NRP ? "🖫" :  "☘")

			If (hsContext=1) {
				LV_GetText(NDsc	, row, 1)
				LV_GetText(NOP	, row, 3)
				If NOP
					GuiControl, HSV:, HSV_NOP	, % NOP                                                            	; Optionen
				If NDsc
					GuiControl, HSV:, HSV_NDS	, % NDsc
			}
			else if (hsContext=2) {
				selRows := "", row := 0
				Gui, HSV: Listview, HSV_LV
				If (NOP := hotstrings.haskey(NRP) ? hotstrings[NRP].options : hsUser[NRP].opt)
					GuiControl, HSV:, HSV_NOP	, % NOP
			}

		; gLabel Events wiederherstellen
			HotS_gSwitch("On")

		}

	}

  ; Hotstring anlegen oder entfernen
  ; ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐ ≝ ≐
	else If (A_GuiControl = "HSV_PM")                                                                                 		{

			Gui, HSV: Submit, NoHide

		; ⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤
		; Eingaben prüfen
		; ⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤
			If (StrLen(HSV_NHS)  = 0) {
				Splash("Das Feld Hotstring darf nicht leer sein!", 4, hsvp.X, Floor(hsvp.Y/2), hsvp.W, hHSV)
				return
			} else if (StrLen(HSV_NRP) = 0) {
				Splash("Das Feld Ersatztext darf nicht leer sein!", 4, hsvp.X, Floor(hsvp.Y/2), hsvp.W, hHSV)
				return
			} else If (hsContext = 2 && !RegExMatch(HSV_NRP, "\{.*[A-Z]\d+\.[\d#]+[A-Z]*\}")) {
				Splash("Das Feld Ersatztext muss einen ICD Code enthalten!", 4, hsvp.X, Floor(hsvp.Y/2), hsvp.W, hHSV)
				return
			}

		; ⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤
		; Hotstringzeichenketten anpassen
		; ⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤
			NHS  	:= Trim(HSV_NHS)                           	; Hotstring
			NRP  	:= RTrim(Trim(HSV_NRP), ";")            	; Ersatztext (replacement)
			NOP  	:= Trim(HSV_NOP)                         	; Options
			Nhse 	:= xstring.EscapeStrRegEx(Nhs)         	; Hotstring escaped für RegEx-Suche
			NRPe 	:= xstring.EscapeStrRegEx(NRP)         	; Ersatzttext escaped für RegEx-Suche
			found 	:= 0

		; ⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤
		; Hotstring hinzufügen und speichern
		; ⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤⏤
			If (hsContext = 2) {

				For dia, hsd in hsUser[hsContext]
					If RegExMatch(hsd.hs, "\b" Nhse "\b") {

						found := 1
						If (dia = NRP) {
							Splash("Hotstring und Ersetzungstext sind beide bekannt! Es wird nichts übernommen", 4, hsvp.X, Floor(hsvp.Y/2), hsvp.W, hHSV)
							return
						}

						Msgbox, 0x1004,% "Addendum Hotstring Manager", % "Diesem Hotstring ist bereits eine Diagnose zugeordnet.`n"
												                                             			. "Wollen Sie das komplette Kontstrukt ersetzen?`n"
												                                            			. "Falls ja wird sicherheitshalber ein BackupUp angelegt."
						IfMsgBox, No
							return

						;hsUser[hsContext][dia] := {"opt":Nop, "hs":"_backup" NhS}

						break

					}

				If !found {
						nil := 0
				}

				HotS_gSwitch("Off")
				GuiControl, HSV:, HSV_NOP	, % "*"
				GuiControl, HSV:, HSV_NHS	;, % " "
				GuiControl, HSV:, HSV_NRP 	;, % " "

				If (hscontext = 1) {
					GuiControl, HSV:, HSV_NDS 	;, % " "
					hsUser[hsContext][dia] := {"opt":Nop, "hs":NhS}
				}
				else if (hscontext = 2) {
					hsUser[hsContext][dia] := {"opt":Nop, "hs":NhS}
				}

				Hotstring(":" Nop ":" Nhs, Func("SendDiagnosis").Bind(NRP))
				HotS_gSwitch("On")

				Gui, HSV: ListView, NHS_LV
				LV_Add("", NRP, Nhs)

				t := ""
				For dia, hsd in hsUser[hsContext]
					t .= hsd.opt "|" hsd.hs "|" dia "`n"

				FileOpen(hsUser.Path.dia, "w", "UTF-8").Write(RTrim(t, "`n"))

			}

	}

  ; Suchen oder neuen Hotstring anlegen
  ; 🔎 🔍 🔎 🔍 🔎 🔍 🔎 🔍 🔎 🔍 🔎 🔍
	else If RegExMatch(A_GuiControl, "i)HSV_(NHS|NOP|NRP|DSC)", hsEdit)    	                    	{

		Gui, HSV: Submit, NoHide
		allWithText := HSV_NHS && HSV_NRP

		Critical Off

	  ; Hotstring wurde geändert
		If (NHSLast <> HSV_NHS || NRPLast <> HSV_NRP || NOPLast <> HSV_NOP) {

			If (NHSLast <> HSV_NHS) {

				abbrevMatched := false
				LV_Delete()

			; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
			;    GOÄ / EBM   - -   GOÄ / EBM   - -   GOÄ / EBM   - -   GOÄ / EBM
			; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
				If (hsContext = 1) {

					For category, hsStrings in hotstrings
						For abbreviation, hsd in hsStrings
							If InStr(hsd.description . hsd.replacement . abbreviation, HSV_NHS) {
								LV_Add("", hsd.description, hsd.replacement, hsd.options, abbreviation)
								abbrevMatched := true
							}

				}

			; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
			;    Diagnosen    - -    Diagnosen    - -    Diagnosen    - -    Diagnosen
			; ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦ ◦
				else if (hsContext = 2) {

					Gui, HSV: Listview, HSV_LV
					Gui, HSV: Listview, -Redraw

				  ; Auflisten passender Hotstring-Diagnosen
				  ; - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -
					words := RegExReplace(HSV_NHS, "\s{2,}", " ")
					words := StrSplit(StrReplace(words, "."), " ")
					For dia, hsd in hotstrings {
						If (hds.func ~="Auswahlbox" || !dia) ; || hsd.Abbreviation ~= "")
							continue
						hits := 0
						For each, word in words
							If InStr(dia " " hsd.abbreviation " " hsd.description, word)                   	; alle Suchbegriffe müssen vorkommen
								hits ++
						If (hits = words.Count())
							LV_Add("", (SubStr(hsd.func,1,4)="Abx#" ? "[ABox] ": "") dia, hsd.Abbreviation)
						If (hsd.abbreviation = HSV_NHS)
							abbrevMatched := true
					}

					LV_Add(""	, "--------------------------------------------------------------------------------------------------------"
								    	, "--------------------------------------------------------------------------------------------------------")

				  ; ICD Katalog laden
				  ; - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -
					If !IsObject(ICDAll) && FileExist(icdpath)
						ICDAll := StrSplit(FileOpen(icdPath, "r", "UTF-8").Read(), "`n", "`r")

				  ; Suchen weiterer ICD-Codes aus dem ICD-10-GM Katalog
				  ; - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -
					If IsObject(ICDAll)
						For idx, icd in ICDAll {
							hits := 0
							For each, word in words
								If InStr(icd, word)
									hits ++
							If (hits = words.Count())                                                                                       	; alle Suchbegriffe müssen vorkommen
								LV_Add("", icd, "" )
							If (LV_GetCount() > 400)                                                                                      	; max. 400 Diagnosen pro Suche
								break
						}

				  ; Timer um RAM nach 10min freizugeben
				  ; - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -
					SetTimer, HotSViewEmptyObjects, % -1*10*60*1000

					Gui, HSV: Listview, HSV_LV
					Gui, HSV: Listview, +Redraw
				}
			}

		  ; Hotstring-Buttonsymbol ändern
			GuiControl, % "HSV:", HSV_PM, % (!abbrevMatched && HSV_NRP && allWithText) ? "⚙" : abbrevMatched ? "🗑" : ""

		  ; Felder leeren wenn Hotstrings leer ist
			Gui, HSV: Submit, NoHide
			HotS_gSwitch("Off")
			GuiControl, HSV:, HSV_NRP	, % (HSV_NHS ? HSV_NRP 	: "")
			GuiControl, HSV:, HSV_NOP	, % (HSV_NHS ? HSV_NOP	: "*")
			GuiControl, % "HSV: Enabled" (HSV_NHS && HSV_NRP ? "1" : "0"), HSV_PM
			GuiControl, % "HSV: Enabled" (HSV_NHS && HSV_NRP ? "1" : "0"), HSV_B1
			HotS_gSwitch("On")

			NRPLast := HSV_NRP, NHSLast := HSV_NHS

		}

	}

return	;}
HSVGuiClose:        	;{
HSVGuiEscape:
	Gui, HSV: Destroy
	hsContext := hsStrings := ""
return
HotSViewEmptyObjects:
	ICDAll := ""
return ;}
}
HotS_gSwitch(gStatus="On")                                                           	{              	;-- glabel an / aus

		global HSV, HSV_NHS, HSV_NRP, HSV_NOP, HSV_NDS

	; gLabel Events abschalten
		GuiControl, % "HSV: " (gStatus="On" ? "+gHSVHandler" : "-g"), HSV_NHS
		GuiControl, % "HSV: " (gStatus="On" ? "+gHSVHandler" : "-g"), HSV_NRP
		GuiControl, % "HSV: " (gStatus="On" ? "+gHSVHandler" : "-g"), HSV_NOP
		GuiControl, % "HSV: " (gStatus="On" ? "+gHSVHandler" : "-g"), HSV_NDS

}
HotS_Privatabrechung()                                                               	{              	;-- weitere Hotstrings: Behörden, Portokosten

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	;	Behörden

		gLASVk 	:= "(sach:Kurzauskunft LA für Soziales u. Versorgung:5.00)"
		gJVEG1	:= "(sach:Anfrage Sozialgericht gem. JVEG: "                  	Addendum.Abrechnung.JVEG.1	")"
		gJVEG2	:= "(sach:Anfrage Sozialgericht gem. JVEG: "                  	Addendum.Abrechnung.JVEG.2	")"
		gJVEG3	:= "(sach:Anfrage Sozialgericht gem. JVEG: "                  	Addendum.Abrechnung.JVEG.3	")"
		gJVEG4	:= "(sach:Anfrage Sozialgericht gem. JVEG: "                  	Addendum.Abrechnung.JVEG.4	")"
		gDRV1     	:= "(sach:Anfrage Rentenversicherung:"                          	Addendum.Abrechnung.DRV.1  	")"
		gDRV2     	:= "(sach:Anfrage Rentenversicherung:"                          	Addendum.Abrechnung.DRV.2  	")"
		gLASV   	:= "(sach:Landesamt für Soziales u. Versorgung:"            	Addendum.Abrechnung.LASV 	")"
		gBAfA   	:= "(sach:Anfrage Bundesagentur für Arbeit gem. JVEG:" 	Addendum.Abrechnung.BAfA 	")"

		HotString(":b0*?X:DRV"         	, "HotS_Info")                                              	; Hotstring Info
		HotString(":b0*?X:RLV"         	, "HotS_Info")
		HotString(":b0*?X:lasv"         	, "HotS_Info")
		HotString(":b0*?X:lageso"     	, "HotS_Info")
		HotString(":b0*?X:JVEG"     	, "HotS_Info")
		HotString(":b0*?X:jveg"      	, "HotS_Info")
		HotString(":b0*?X:Sozial"     	, "HotS_Info")
		HotString(":b0*?X:Bundes"     	, "HotS_Info")
		HotString(":b0*?X:Agentur"    	, "HotS_Info")
		HotString(":*X:JVEG1"           	, Func("HotS_Info").Bind(gJVEG1))               	; Anfrage Sozialgericht
		HotString(":*X:JVEG2"           	, Func("HotS_Info").Bind(gJVEG2))               	; Anfrage Sozialgericht
		HotString(":*X:JVEG3"           	, Func("HotS_Info").Bind(gJVEG3))               	; Anfrage Sozialgericht
		HotString(":*X:JVEG4"           	, Func("HotS_Info").Bind(gJVEG4))               	; Anfrage Sozialgericht
		HotString(":*X:Sozial1"       	, Func("HotS_Info").Bind(gJVEG1))               	; Anfrage Sozialgericht
		HotString(":*X:Sozial2"       	, Func("HotS_Info").Bind(gJVEG2))               	; Anfrage Sozialgericht
		HotString(":*X:Sozial3"       	, Func("HotS_Info").Bind(gJVEG3))               	; Anfrage Sozialgericht
		HotString(":*X:Sozial4"       	, Func("HotS_Info").Bind(gJVEG4))               	; Anfrage Sozialgericht
		HotString(":*X:lasv1"            	, Func("HotS_Info").Bind(gLASV))                    	; Anfrage LaGeSo Schwerbehinderung Normal
		HotString(":*X:lasv2"            	, Func("HotS_Info").Bind(gLASVk))                  	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
		HotString(":*X:lasvk"             	, Func("HotS_Info").Bind(gLASVk))                  	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
		HotString(":*X:lageso1"        	, Func("HotS_Info").Bind(gLASV))                   	; Anfrage LaGeSo Schwerbehinderung Normal
		HotString(":*X:lageso2"        	, Func("HotS_Info").Bind(gLASVk))                  	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
		HotString(":*X:lagesokurz"    	, Func("HotS_Info").Bind(gLASVk))                	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
		HotString(":*X:RLV1"              	, Func("HotS_Info").Bind(gDRV1))                   	; Anfrage Rentenversicherung
		HotString(":*X:RLV2"              	, Func("HotS_Info").Bind(gDRV2))                	; Anfrage Rentenversicherung
		HotString(":*X:DRV1"            	, Func("HotS_Info").Bind(gDRV1))                   	; Anfrage Rentenversicherung
		HotString(":*X:DRV2"            	, Func("HotS_Info").Bind(gDRV2))                   	; Anfrage Rentenversicherung
		HotString(":*X:Rent"            	, Func("HotS_Info").Bind(gDRV1))                	; Anfrage Rentenversicherung
		HotString(":*X:S0051"         	, Func("HotS_Info").Bind(gDRV2))                	; Anfrage Rentenversicherung
		HotString(":*X:Bundes1"        	, Func("HotS_Info").Bind(gBAfA))                 	; Anfrage Bundesagentur für Arbeit
		HotString(":*X:Agentur1"       	, Func("HotS_Info").Bind(gBAfA))                  	; Anfrage Bundesagentur für Arbeit

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Porto
		Postkarte    	:= "(sach:Porto Standard:"	Addendum.Preise.Porto.Postkarte ")"
		Standard   	:= "(sach:Porto Standard:"	Addendum.Preise.Porto.Standard ")"
		Kompakt    	:= "(sach:Porto Kompakt:" 	Addendum.Preise.Porto.Kompakt ")"
		Gross        	:= "(sach:Porto Groß:"    	Addendum.Preise.Porto.Gross ")"
		Maxi   	    	:= "(sach:Porto Maxi:"     	Addendum.Preise.Porto.Maxi ")"
		Paeckchen 	:= "(sach:Porto Päckchen:" Addendum.Preise.Porto.Paeckchen ")"
		Einschreiben	:= "(sach:Porto Einschreiben:2.65)"
		Einwurf     	:= "(sach:Porto Einwurf:2.35)"

		HotString(":b0*?X:porto"     	, "HotS_Info")
		HotString(":b0*?X:Standard"  	, "HotS_Info")
		HotString(":b0*?X:Kompakt"  	, "HotS_Info")
		HotString(":b0*?X:Gross"     	, "HotS_Info")
		HotString(":b0*?X:Maxi"     	, "HotS_Info")
		HotString(":b0*?X:Päck"      	, "HotS_Info")
		HotString(":*:porto0"       	, Func("HotS_Info").Bind(Postkarte))                   	; Postkarte
		HotString(":*:Postk"          	, Func("HotS_Info").Bind(Postkarte))                 	; Postkarte
		HotString(":*:karte"          	, Func("HotS_Info").Bind(Postkarte))                 	; Postkarte
		HotString(":*:porto1"        	, Func("HotS_Info").Bind(Standard))                 	; Porto bis 20g
		HotString(":*:Standard"   	, Func("HotS_Info").Bind(Standard))                 	; Porto bis 20g
		HotString(":*:porto2"      	, Func("HotS_Info").Bind(Kompakt))                 	; Porto bis 50g
		HotString(":*:Kompakt"   	, Func("HotS_Info").Bind(Kompakt))                   	; Porto bis 50g
		HotString(":*:porto3"      	, Func("HotS_Info").Bind(Gross))                        	; Porto bis 500g
		HotString(":*:Groß"        	, Func("HotS_Info").Bind(Gross))                       	; Porto bis 500g
		HotString(":*:porto4"      	, Func("HotS_Info").Bind(Maxi))                          	; Porto bis 1000g
		HotString(":*:Maxi"         	, Func("HotS_Info").Bind(Maxi))                         	; Porto bis 1000g
		HotString(":*:porto5"       	, Func("HotS_Info").Bind(Paeckchen))                	; DHL-Päckchen
		HotString(":*:Päck"         	, Func("HotS_Info").Bind(Paeckchen))                  	; DHL-Päckchen
		HotString(":*:Pack"         	, Func("HotS_Info").Bind(Paeckchen))                	; DHL-Päckchen

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; IF Bedingung - ENDE
		Hotkey, If

return
}
HotS_Info(replacement:="")                                                        		{              	;-- zeigt Hinweise zu Hotstrings

	static hs
	static Bericht := [	"Bericht ohne nähere gutachtliche Äußerung [Nr. 200 der Anlage 2 zu 8 10 Abs. 1 JVEG]"
						,		"Bericht mit außergewöhnlich umfangreicher Tätigkeit [Nr. 201 der Anlage 2 zu 8 10 Abs. 1 JVEG]"
						,		"Bericht mit kurzer gutachtlicher Äußerung [Nr. 202 der Anlage 2 zu 8 10 Abs. 1 JVEG]"
						,		"Bericht mit außergewöhnlich umfangreichen gutachterlichen Äußerung [Nr. 203 der Anlage 2 zu 8 10 Abs. 1 JVEG]"]


	If !IsObject(hs) {

		PK := Addendum.Preise.Porto.Postkarte, GB := Addendum.Preise.Porto.Gross, KB := Addendum.Preise.Porto.Kompakt, MB := Addendum.Preise.Porto.Maxi
		PN := Addendum.Preise.Porto.Paeckchen, SB := Addendum.Preise.Porto.Standard

		hs := {	"DRV"             	: "Rentenversicherung:`nDRV1  (" Addendum.Abrechnung.DRV.1 " €) "
								        	. "und DRV2  (" Addendum.Abrechnung.DRV.2 " €)"
				, 	"RLV"             	: "Rentenversicherung:`nRLV1  (" Addendum.Abrechnung.DRV.1 " €) "
								        	. "und RLV2  (" Addendum.Abrechnung.DRV.2 " €)"

				, 	"LASV"           	: "Erstfeststellungsverfahren nach dem Schwerbehindertenrecht:`n"
								        	. "LASV1  (" Addendum.Abrechnung.LASV " €) und LASV2/LASVk (5.00 €)"
				, 	"lageso"         	: "Erstfeststellungsverfahren nach dem Schwerbehindertenrecht:`n"
								        	. "lageso1  (" Addendum.Abrechnung.LASV " €) und lageso2/lagesokurz (5.00 €)"

				,	"JVEG"           	: "z.B. Sozialgericht:`n"
								        	. "JVEG1  ("  Addendum.Abrechnung.JVEG.1 ") " Bericht.1 "`n"
								        	. "JVEG2  ("  Addendum.Abrechnung.JVEG.2 ") " Bericht.2 "`n"
								        	. "JVEG3  ("  Addendum.Abrechnung.JVEG.3 ") " Bericht.3 "`n"
								        	. "JVEG4  ("  Addendum.Abrechnung.JVEG.4 ") " Bericht.4 "`n"
								        	. "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `n"
								        	. "Ansetzen nicht vergessen!: Portogebühren mit 'porto' und Kopiegebühren mit 'Kopie'"

				,	"Sozial"           	: "z.B. Sozialgericht:`n"
								        	. "Sozial1  ("  Addendum.Abrechnung.JVEG.1 ") " Bericht.1 "`n"
								        	. "Sozial2  ("  Addendum.Abrechnung.JVEG.2 ") " Bericht.2 "`n"
								        	. "Sozial3  ("  Addendum.Abrechnung.JVEG.3 ") " Bericht.3 "`n"
								        	. "Sozial4  ("  Addendum.Abrechnung.JVEG.4 ") " Bericht.4 "`n"
								        	. "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `n"
								        	. "Ansetzen nicht vergessen!: Portogebühren mit 'porto' und Kopiegebühren mit 'Kopie'"

				,	"Bundes"			: "Bundesagentur für Arbeit: Bundes1/Arbeit1`n"
											. "Entschädigungssatz in Anlehnung an das JVEG in Höhe von " StrReplace(Addendum.Abrechnung.BAfA, ".", ",") " Euro.`n"
				                        	. "Portokosten sind zusätzlich vergütungsfähig.`n"
											. "Kopien: 0,50 Euro pro Seite für die ersten 50 Seiten und von 0,15 Euro für jede weitere Seite."

				, 	"porto"          	: "P O R T O K O S T E N`n`n"
								        	. "porto0:  Postkarte " 	PK	" € `n"
								        	. "porto1:  Standard " 	SB	" €`t(Gewicht bis 20 g. max. LxBxH 23,5 x 12,5 x 0,5 cm [z.B. DIN lang oder C6 Umschlag])`n"
								        	. "porto2:  Kompakt "  	KB	" €`t(Gewicht bis 50 g. max. LxBxH 23,5 x 12,5 x 1 cm [z.B. DIN lang oder C6 Umschlag])`n"
								        	. "porto3:  Groß "       	GB	" €      `t(Gewicht bis 500 g. max. LxBxH 35,3 x 25 x 2 cm [z.B. Umschläge bis zu DIN B4])`n"
								        	. "porto4:  Maxi "       	MB	" €     `t(Gewicht bis 1000 g. max. LxBxH 35,3 x 25 x 5 cm [z.B. Umschläge oder Versandboxen bis zu DIN B4])`n"
								        	. "porto5:  Päckchen "	PN	" €`t(Gewicht bis 2 kg. max. LxBxH 60 x 30 x 15 cm`n`n"
								        	. "alternativ:  Postkarte,Karte,Standard,Kompakt,Groß,Maxi,Päckchen"}

			hs.Arbeit := hs.Bundes

	}

	ControlGetFocus, cFocus, A

	If !replacement {
		cp := GetWindowSpot(GetFocusedControlHwnd(WinExist("A")))
		RegExMatch(A_ThisHotkey, "\:.*\:(?<HotString>.*)$", this)
		RegExReplace(hs[thisHotString], "\n", "", CR)
		ToolTip, % hs[thisHotString], % A_CaretX, % cp.Y-cp.H-((CR+1)*14)-2 , 12
		SetTimer, HotS_InfoOff, -16000
	}
	else {
		SendRaw, % replacement
		ToolTip,,,, 12
	}

return

HotS_InfoOff:
	SetTimer, HotS_InfoOff, Off
	ToolTip,,,, 12
return
}

Auswahlbox(selection, selectionTitle:="")                                            	{                	;-- Listbox für Diagnosenauswahl

	; letzte Änderung: 27.02.2022

	;{ Variablen
	  ; Farbkombinationen für die Auswahlbox
		static ICDColor := {"A"	: ["EEC591"  	, "F3DCBE"]  	, "B"	: ["EEC591"	, "F3DCBE"]  	, "C"	: ["D0AE7B"	, "E6D3BB"]   	, "D"	: ["FEDC60"	, "FFEC93"]
									, "E"	: ["FEDC60"	, "FFEC93"]   	, "F"	: ["7AA7C5"  , "9BBBCE"] 	, "G"	: ["FFC04D"	, "FFD280"]    	, "H"	: ["CCB3CA"	, "E1D5E0"]
									, "I"	: ["FD9595"	, "FDCECE"]  	, "J"	: ["9FC4E8"   , "D7E6F4"]  	, "K"	: ["E0C5B1"	, "EEDBD2"]   	, "L"	: ["E0C5B1"	, "EEDBD2"]
									, "M": ["AABCA9"	, "CED5CC"] 	, "N"	: ["DEDE9A"	, "EBEBC7"]   	, "O"	: ["74CD74"	, "A9DEA9"]   	, "P"	: ["99C699"	, "B9D9B9"]
									, "Q"	: ["B0D7BB"	, "CEE3D3"]  	, "R"	: ["9EDAD1"	, "BEE9E6"]   	, "S"	: ["AABCA9" , "CED5CC"]   	, "T"	: ["E0BCA9"	, "EBD5C7"]
									, "U"	: ["BFBFBF"	, "DDDDDD"]	, "V"	: ["BFBFBF"	, "DDDDDD"]	, "X"	: ["BFBFBF"	, "DDDDDD"]	, "Y"	: ["BFBFBF"	, "DDDDDD"]
									, "Z"	: ["DEB4DE"	, "EACCEA"]	}
	  ; Farbe der Zeilenauswahl
		static rC := [{T: 0x0, B: 0xFFDEAD, TS: 0xFFFFFF, BS: 0x1389FF}, {T: 0x0, B: 0xFFF2DB, TS: 0xFFFFFF, BS: 0x1389FF}]
	  ; Hintergrundfarben
		static BgColor	:= ["c6A99C8", "c6A99C8"]
		static YDelta  	:= 2
		static AWBOk, AWBCancel, AWBLBh, AWBLb, hAWB, AWB, hAWBLb
		static FontName, FontSize, FontStyle, cy, ch, guiX, guiY
		static CaretX, CaretY, ab, monNr, Mon, Lbselection, conStored, ACControl, ACHwnd, ACWin, guititle

	  ; aktuellen Eingabefocus merken
		ControlGetFocus, ACControl, A
		ControlGet, ACHwnd, HWND,, % ACControl, % "ahk_id " (ACWin := GetHex(WinExist("A")))
		Addendum.AWBCheckFocus := GetHex(ACHwnd)

	 ; Wartet bis der selbstersetzende Hotstring entfernt wurde
		Loop % StrLen(A_ThisHotkey)*2
			Sleep 15

	  ; Steuerelementpositionen ermittlen
		CaretX    	:= A_CaretX, CaretY := A_CaretY
		monNr    	:= GetMonitorAt(CaretX, CaretY)
		Mon     	:= ScreenDims(monNr)
		TBHeight	:= TaskbarHeight(MonNr)
		Mon.H  	:= Mon.H - TBHeight
		cp         	:= GetWindowSpot(ACHwnd)

	;}

	;{ Auswahl(selection) parsen
		If !IsObject(selection) {
			PraxTT("Auswahlbox`n...es wurden keine Diagnosen übergeben.", "2 2")
			return
		}

	  ; flexible Option nutzen, falls übergeben
		If selectionTitle
			selection[1] := selectionTitle
		conStored := selection, Lbselection := "", ColAlteration := "Sub"
		MainICDs := Object(), rowcolors := []

	  ; Optionen und Einträge parsen
		For idx, val in selection
			If (idx=1) {                          	; Gui Optionen
				RegExMatch(val, "i)border\s*=\s*(?<Caption>[a-z]+)"                                     	, Show)
				RegExMatch(val, "i)title\s*=\s*(?<title>[\pL\d\-\_,+~#\s\.]+)(\s\w+=|\s*$)" 	, gui)
				RegExMatch(val, "i)Size\s*=\s*(?<Size>\d+)"                                                    	, Font)
				RegExMatch(val, "i)Style\s*=\s*(?<Style>[\pL\_\-\+\~\#\s]+)"                          	, Font)
				RegExMatch(val, "i)Name\s*=\s*(?<Name>[a-z\s]+)"                                     	, Font)
				RegExMatch(val, "i)BgColor1\s*=\s*(0x)*(?<Color1>[A-F\d]+)"                         	, RBg)
				RegExMatch(val, "i)TxtColor1\s*=\s*(0x)*(?<1>[A-F\d]+)"                                	, text)
				RegExMatch(val, "i)BgColor2\s*=\s*(0x)*(?<Color2>[A-F\d]+)"                         	, RBg)
				RegExMatch(val, "i)TxtColor2\s*=\s*(0x)*(?<2>[A-F\d]+)"                               	, text)
				RegExMatch(val, "i)GuiColor1\s*=\s*(0x)*(?<1>[A-F\d]+)"                                	, gui)
				RegExMatch(val, "i)GuiColor2\s*=\s*(0x)*(?<2>[A-F\d]+)"                                	, gui)
				RegExMatch(val, "i)maxRows\s*=\s*(?<maxRows>\d+)"                                  	, LB)
				LBmaxRows 	:= !LBmaxRows	? 20                                        	: LBmaxRows
				guititle      	:= !guititle     	? "Auswahlbox"                      	: guititle
				FontSize    	:= !FontSize  	? "10"                                    	: FontSize
				FontStyle    	:= !FontStyle 	? "Normal"                            	: FontStyle
				FontName 	:= !FontName	? "Arial"                                 	: FontName
				BgColor.1    	:= gui1          	? "c"   LTrim(gui1   	, "0x")      	: BgColor.1
				BgColor.2    	:= gui2          	? "c"   LTrim(gui2   	, "0x")      	: BgColor.2
				rC.1.B       	:= RBgColor1  	? "0x" LTrim(RBgColor1, "0x")    	: rC.1.B
				rC.1.T       	:= text1        	? "0x" LTrim(text1   	, "0x")      	: rC.1.T
				rC.2.B       	:= RBgColor2  	? "0x" LTrim(RBgColor2, "0x")    	: rC.2.B
				rC.2.T       	:= text2        	? "0x" LTrim(text2   	, "0x")       	: rC.2.T
			}
			else {                                	; Diagnosen

			  ; Liste zusammenstellen
				RegExMatch(val, "i)(?<Text>.*)?\{(?<Code>[\!\+\*]*(?<Main>(?<Chapter>[A-Z])[\d]+)\.*(?<Sub>[\d]+)*)?[\<\/\>\*RLBAVGZ#]*\}", ICD)
				If (StrLen(Trim(ICDCode . ICDText)) > 0)
					Lbselection .= ICDCode . (StrLen(ICDCode)<4 ? "        `t`t" : StrLen(ICDCode)<6 ? "     `t`t" : "   `t`t") . RTrim(ICDText) "|"

			  ; dies ist für die Farbalteration
				If (MainICDs.Count() = 0)
					MainICDs[ICDMain] := 1
				If (ColAlteration = "Sub" && !MainICDs.haskey(ICDMain))
					ColAlteration := "Main"

			}
	;}

	;{ Gui
		LBItems := selection.MaxIndex()-1 > LBmaxRows ? LBmaxRows : selection.MaxIndex()-1
		LBOptions := " w700 HWNDhAWBLb +0xD0 T4 T8 T32 T64 T128 Multi AltSubmit gAWBLBHandler" ; LBS_OWNERDRAWFIXED=0x0010, LBS_HASSTRINGS=0x0040

		Gui, AWB: new		, % (ShowCaption="off" ? "-Caption":"") " -DPIScale +AlwaysOnTop +HWNDhAWB"  ; +ToolWindow
		Gui, AWB: Color  	, % BgColor1, % BgColor2
		Gui, AWB: Margin	, 1, 1

		OD_Colors.SetItemHeight("s" FontSize " " FontStyle, FontName)
		Gui, AWB: Font		, % "s" FontSize " " FontStyle, % FontName
		Gui, AWB: Add		, Listbox, % "r" LBItems "            	vAWBLB " LBOptions                            	, % RTrim(Lbselection, "|")

		Gui, AWB: Font		, % "s" (FontSize-2 < 7 ? 7 : FontSize-2)
		Gui, AWB: Add		, Button, % "xm+1 y+1           	vAWBOk    	gAWBButtons "                 	, % "Diagnose(n) übernehmen"
		Gui, AWB: Add		, Button, % "x+30   	             	vAWBCancel	gAWBButtons "               	, % "Abbruch"

	  ; ideale Breite berechnen
		lbW := LBEX_CalcIdealWidthEx(0, RTrim(Lbselection, "|"), "|", "s" FontSize " " FontStyle, FontName, LBItems)
		GuiControl, AWB: Move, AWBLB, % "w" lbW

	  ; versteckt Anzeigen
		GuiControlGet, dp, AWB: Pos, AWBOk
		Gui, AWB: Show	, % " h" dpY+dpH+2 " NA Hide"                                                         	, % guititle

	  ; Gui Position entsprechend ihrer Größe anpassen
		ab  	:= GetWindowSpot(hAWb)
		guiX 	:= CaretX	+ ab.W     	> Mon.W	? Mon.X+Mon.W-ab.W-13 	: CaretX - 3
		guiY 	:= cp.Y+cp.H+ab.H 	> Mon.H	? cp.Y-YDelta-ab.H           	: cp.Y+cp.H
		guiY 	:= guiY                     	< 1      	? cp.Y+cp.H+ab.H+2      	: guiY
		guiY	:= guiY+ab.H             	> Mon.H	? Mon.H-ab.H-2                	: guiY

	  ; Gui positionieren
		SetWindowPos(hAWb, guiX, guiY, ab.W, ab.H)

	  ; Hintergrundfarben der Zeilen festlegen
		rA         	:= 1
		rowcolors 	:= []
		For idx, val in selection {
			If (idx>1) {

				RegExMatch(val, "Oi)(?<Text>.*)?\{(?<Code>[\!\+\*]*(?<Main>(?<Chapter>[A-Z])[\d]+)\.*(?<Sub>[\d]+)*)?[\<\/\>\*RLBAVGZ#]*\}", ICD)
				chCol := ICDColor[ICD.Chapter]
				If !rowcolors.Count() {
					rowcolors.Push({"T":0x0, "B":"0x" . chCol[rA], "TS":0xFFFFFF, "BS":0x1389FF})
				}
				else {
					; alternierende Farbwahl (ICD CoAlteration gleich ist und letztes ICD Kapitel und aktuelles auch dann bleibt die bisherige Farbe,
					; im Falle eine Kapitelwechsel rA wieder 1 stellen ansonsten zwischen 1 und 2 hin und herschalten)
					rA := lastICD=ICD[ColAlteration] && lastChapter=ICD.Chapter ? rA : lastChapter<>ICD.Chapter ? 1 : (rA = 2 ? 1 : 2)
					rowcolors.Push({"T":0x0, "B":"0x" . chCol[rA], "TS":0xFFFFFF, "BS":0x1389FF})
				}

			}
			lastICD := ICD[ColAlteration], lastChapter := ICD.Chapter
		}
		OD_Colors.Attach(hAWBLb, rowcolors)

	  ; Gui zeigen
		Gui, AWB: Show, % " NA"

		GuiControl, AWB: Choose, AWBLB, 1
		GuiControl, AWB: Focus, AWBLB

	;}

	;{ Gui Hotkeys
		HotKey, IfWinExist, % guititle " ahk_class AutoHotkeyGUI"
		HotKey, Enter	, AWBTLB
		Hotkey, Up    	, AWBUp
		Hotkey, Down  	, AWBDown
		HotKey, Esc   	, AWBGuiEscape
		Hotkey, IfWinExist
	;}


Return
AWBLBHandler:   	;{ gLabel

	If (A_GuiEvent = "DoubleClick")
		gosub AWBTLB

return
;}
AWBDown:        	;{ Zeile der Listboxauswahl verschieben
AWBUp:
	; Tasten an anderes Fenster senden, falls der Eingabefocus nicht bei Albis oder der Gui liegt
	If (WinActive("ahk_class OptoAppClass") || WinActive(guititle " ahk_class AutoHotkeyGUI"))
		ControlSend,, % "{LControl Down}{" A_ThisHotkey "}{LControl Up}", % "ahk_id " hAWBLb
	else
		SendInput % "{" A_ThisHotkey "}"
return ;}
AWBButtons:     	;{ Übernehmen oder Abbrechen
	If (A_GuiControl = "AWBCancel")
		goto AWBGuiClose
;}
AWBTLB:             	;{	ausgewähltes Senden

	Gui, AWB: Default
	Gui, AWB: Submit, Hide

	;~ WinActivate    	, % "ahk_id " ACWin
	;~ WinWaitActive	, % "ahk_id " ACWin,, 2
	ControlFocus	,,% "ahk_id " Addendum.AWBCheckFocus
	;~ ControlFocus	, % ACControl, % "ahk_id " ACWin

	Karteikartenauswahl := false
	If IsObject(conStored)
		For idx, item in StrSplit(AWBLb, "|") {
			dia := RegExReplace(conStored[item+1], "\s{2,}", " ")
			Karteikartenauswahl := !Karteikartenauswahl && RegExMatch(dia, "\<.*\>") ? true : false
			SendDiagnosis(dia)
		}
	If Karteikartenauswahl {
		SendInput, {Home}
		SendInput, {LControl Down}{Right}{LControl Up}
	}

	Addendum.AWBCheckFocus := ""
	OD_Colors.Detach(hAWBLb)
	Gui, AWB: Destroy

;}
AWBGuiClose:  	;{
AWBGuiEscape:
	If (WinActive("ahk_class OptoAppClass") || WinActive(guititle " ahk_class AutoHotkeyGUI")) {
		Addendum.AWBCheckFocus := ""
		OD_Colors.Detach(hAWBLb)
		Gui, AWB: Destroy
	} else {
		SendInput % "{" A_ThisHotkey "}"
		return
	}
	CaretX := CaretY := ab := monNr := Mon := Lbselection := conStored := ACControl := ACHwnd := ACWin := guititle := ""
	FontName := FontSize := FontStyle := cy := ch := guiX := guiY := hAWB := hAWBLb := ""
return ;}
AWBCheckFocus:	;{



return ;}
}
SendDiagnosis(dia, noreplace:=false, opt:="")                                        	{                 	;-- für automatische Ersetzung von Diagnosestrings

	/*
		 ·	selektiert ein '#' Zeichen im übergebenen String als Position für Nutzereingaben z.B. eine Seitenangabe: Pneumonie {J15.8#G};
		 · 	geplant sind flexiblere Ersetzungen für Seitenangabe, Diagnosesicherheit und Diagnosentext
		 ·  in eckige Klammern [] gesetzte Bereiche sind für flexible RegEx-Ersetzungen gedacht

			Einschränkung: BlockInput ist nur wirksam, wenn dieses Skript als Administrator gestartet wurde

			letzte Änderung: 27.02.2022
	*/

	static CMB := ["I","II","III","IV","V","VI"]

	;SendMode                      	, Input
	SetKeyDelay                   	, 20, 20

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; es lassen sich bestimmte Teile durch RegEx ersetzen
  ; aus : :X*:.NYHA IV::SendDiagnosis("Linksherzinsuffizienz NYHA-Stadium $CMB {I50.1$G}")
  ; wird: Linksherzinsuffizienz NYHA-Stadium IV {I50.14G}
	If RegExMatch(dia, "\$CMB") {

		If !RegExMatch(A_ThisHotkey, "[^\d](?<rad>\d)\s*$", g)
			If RegExMatch(A_ThisHotkey, "\b(?<ahl>I|II|III|IV|V|VI)\b$", Z)
				For grad, LZ in CMB
					If (Zahl = LZ)
						break

		dia := RegExReplace(dia, "\$CMB"	, CMB[grad])
		RegExMatch(dia, "\$\s*(?<pm>[\-\+])\s*(?<diff>\d)", g)

	  ; es gibt ICD Diagnosen bei denen die Grade mit 00 anstatt 01 beginnen
		dia := RegExReplace(dia, "[\$]\s*(?=[^\-\+])"	, grad + (gpm ? ((gpm="-" ? -1 : 1) * gdiff) : 0))
		dia := RegExReplace(dia, "[\$]\s*[\-\+]\d"      	, grad + (gpm ? ((gpm="-" ? -1 : 1) * gdiff) : 0))

	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; direkte Übernahme von Ausschluß von, Verdacht auf, Gesichert und Zustand nach aus dem Hotstring
  ; :XC*:.KHKG::SendDiagnosis("KHK {I25.19$#}")
	If InStr(dia, "$#") {
		RegExMatch(A_ThisHotkey, "[AVGZ]$", gesichert)
		dia := StrReplace(dia, "$#", gesichert ? gesichert : "#")
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; einzeln stehende $ Zeichen kodieren für eine Zahl, diese befindet sich am Ende des Hotstrings
  ; aus:  	:X*:.KHK-3::SendDiagnosis("KHK-$ {I25.1$G}")
  ; wird: 	KHK-3 {I25.13G}
	If InStr(dia, "$") {
		If RegExMatch(A_ThisHotkey, "[^\d](?<ahl>\d+)\s*$", Z)
			dia := StrReplace(dia, "$", Zahl)
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; weitere Ersetzungen
	dia := RegExReplace(dia, "\[.*?\]", DA)                  	; Ersetzen aus Eingabe oder Entfernen wenn leer
	dia := RegExReplace(dia, "[#]+", "#", rplCount)  	; nur einzelne # stehen lassen
	dia .= !RegExMatch(dia, "\s*;\s*$") ? ";" : ""         	; fügt ein Semikolon ans Ende des Strings an
	SendRaw, % dia

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Caret/Cursor auf das #-Zeichen versetzen und umschliessen
	If !noreplace && (rplPos := InStr(dia, "#"))
		SendInput, % "{Left " (StrLen(dia)-rplPos+1) "}{LShift Down}{Right 1}{LShift Up}"

}
SendGB(HotstringGB)                                                                  	{                 	;-- sendet Gebühren

	RegExMatch(HotstringGB, "Oi)(?<GB>[\dA-Z]+)\(*(?<Zusatz>(?<Bezeichnung>[A-Z]+)\:(?<Wert>[\dA-Z,]+))*\)*", hs)

	;~ If hs.

}

;}
;{ ###############   SCITE4AUTOHOTKEY    ###############

SendRawBzu:                   	;{	 	Strg (links) + Ziffernblock 2                                             	(für Scite4Autohotkey)
	SendRawFast(";}")
return ;}
SendRawBAuf:               	;{	 	Strg (links) + Ziffernblock 1                                             	(für Scite4Autohotkey)
	SendRawFast(";{ ")
return ;}
SendRawProzentAll: 	    	;{	 	Strg (links) + 5                                                                	(für Scite4Autohotkey)
	SendRawFast("%%", "L")
return ;}
SendStrgRight: 	            	;{ 	Ziffernblock 6                                                                 	(für Scite4Autohotkey)
	SendInput, {LControl Down}{Right}{LControl Up}
return ;}
SendStrgLeft:                 	;{ 	Ziffernblock 4                                                                 	(für Scite4Autohotkey)
	SendInput, {LControl Down}{Left}{LControl Up}
return ;}
SendRawR:                    	;{	 	Strg (links) + ´                                                                	(für Scite4Autohotkey) | `r
	SendRawFast("``r", "")
return ;}
SendRawN:                   	;{	 	Strg (links) + `                                                                	(für Scite4Autohotkey) | `n
	SendRawFast("``n", "")
Return ;}
SendRawTab:                	;{	 	Strg (links) + Alt (links) + `                                               	(für Scite4Autohotkey) | `t
	SendRawFast("``t", "")
return ;}
SendFourSpace:            	;{	 	Strg (links) + Pfeil rechts                                                   	(für Scite4Autohotkey)
	Send, {Space 4}
Return ;}
SendFourBackspace:      	;{	 	Strg (links) + Delete                                                         	(für Scite4Autohotkey)
	Send, {BackSpace 4}
Return ;}
SciteDescriptionBlock:    	;{	 	Alt (links)   + v                                                               	(für Scite4Autohotkey)
	SendRawFast(";---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "Enter")
	SendRawFast("; " Clipboard , "Enter")
	SendRawFast(";---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "Enter")
return
;}
SciteCommentBlock:      	;{	 	Strg (links) + ,                                                                  	(für Scite4Autohotkey)
	SendAlternately("*/|/*")            ; sendet umgekehrt, weil sonst alles darunter liegende als Kommentar gilt und die Codefaltung aufgehoben wird
return
;}
SciteBracket:                   	;{ 	Strg (links) + .                                                                  	(für Scite4Autohotkey)
	SendAlternately("{|}")
return
;}
SciteSmartCaretL:           	;{
	SciteSmartCaret(-1)
return
SciteSmartCaretR:
	SciteSmartCaret(1)
return

SciteSmartCaret(direction) {

	global curpos, curline, curlen, endpos, bufLen

	sci            	:= Sci_GetCurTextLine()
	Txt            	:= direction = - 1 ? xstring.Flip(sci.curText) : sci.curText
	LTxt           	:= sci.strL
	RTxt          	:= sci.strR
	SLen         	:= StrLen(Txt)
	curpos       	:= sci.Pos
	curColumn	:= sci.Column
	curline       	:= sci.curline
	posfromline	:= sci.PosFromLine
	LineStartPos	:= curpos - curColumn                                                                                			;lineendpos ist immer +2 wegen CR + LF Zeichen am Ende
	EndPos        	:= LineStartPos + SLen
	CPosInTxt  	:= curpos - LineStartPos
	newPos	    	:= curpos + ( direction * InStr(Txt, ",", false, CPosInTxt+1, 1) )

	SendMessage, 2025, % newPos, 0,, % "ahk_id " sci.hScintilla
	gosub MsgBoxer
	sscDir := 0

return

MsgBoxer:
	t =
	(
direction    	:  %direction%
CaretPosition:  %curpos%
curline      	:  %curline%
posfromline	:  %posfromline%
curColumn	:  %curColumn%
LineStartPos	:  %LineStartPos%
LineEndPos	:  %EndPos%
newPos     	:  %newPos%
Textlänge  	:  %SLen%
Text         	:  %Txt%
Text Caret li.	:  %LTxt%
Caret re.     	:  %RTxt%
	))

	SciTEOutput(t "`n")

return
}

;}

;{ ###############        ALBIS          ###############
AlbisMarkieren:              	;{ 	Strg	+ a
	AlbisSelectAll()
return ;}
AlbisKarteikarteSchliessen:	;{ 	Shift	+	Pfeil nach unten
	IfWinNotActive, ahk_class OptoAppClass
			WinActivate, ahk_class OptoAppClass
	Send, {LControl Down}{F4}{LControl Up}
return
;}
AlbisNaechsteKarteikarte:	;{ 	Shift	+	Pfeil nach rechts
	SendMessage, 0x224,,,, % "ahk_id " AlbisMDIClientHandle()	; WM_MDINEXT
return
;}
Schulzettel:                    	;{ 	Strg 	+ F11
	Run, %AddendumDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%AddendumDir%\Module\Albis_Funktionen\Schulzettel.ahk"
return ;}
Hausbesuche:               	;{ 	Strg 	+ F12
	Run, %AddendumDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%AddendumDir%\Module\Albis_Funktionen\Hausbesuche.ahk"
return ;}
DauermedBezeichner:   	;{ 	nur # eingeben im Dauermedikamentenfenster
	ControlGetText, fT,, % "ahk_id " (hfc := GetFocusedControl())
	If (!fT && InStr(GetClassName(hfc), "Edit"))
        MedTrenner(hfc)
	else
		SendRaw, % "#"
return
;}
EnableAllControls:	        	;{
	EnableAllControls(1)
return
;}
PlusEineWoche:             	;{ 	F10                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
		AddToDate(Feld, 7, "days")
	else
		SendInput, {F10}
Return
;}
MinusEineWoche:          	;{ 	F9                                                                                       - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
		AddToDate(Feld, -7, "days")
	else
		SendInput, {F9}
Return
;}
PlusEinMonat:                	;{ 	F12                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
		AddToDate(Feld, 4*7+2, "days")
	else
		SendInput, {F12}
Return
;}
MinusEinMonat:            	;{ 	F11                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
		AddToDate(Feld, -4*7-2, "days")
	else
		SendInput, {F11}
Return
;}
Kalender:                      	;{ 	Shift	+ F3
	If IsObject(feld := IsDatumsfeld())
		If InStr(GetProcessNameFromID(feld.hwin), "albis") {
			Calendar(feld)  ; Kalender-Gui anzeigen
			return
		}
	SendInput, {Shift Down}{F3}{Shift Up}
return
;}
DruckeLabor:                	;{ 	Alt 	+ F8
	AlbisLaborblattDrucken(1, "Standard", "")
return
;}
Laborblatt:                    	;{ 	Alt 	+ Rechts
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") || InStr(AlbisGetActiveWindowType(), "Laborblatt")
		return
	AlbisLaborBlattZeigen()
return
;}
Karteikarte:                   	;{ 	Alt 	+ Links
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") or InStr(AlbisGetActiveWindowType(), "Karteikarte")
		return
	AlbisKarteikarteZeigen()
return
;}
GVUListe_ProgDatum:   	;{ 	Strg	+	F7
	GVUListe(1)
Return
;}
GVUListe_ZeilenDatum: 	;{ 	Alt	+	F7
	GVUListe(2)	; Funktion ist im Addendumskript
return
;}
KarteikartenMenu:          	;{		Strg	+	Shift	+ M
	MouseGetPos, mx, my
	; Context menu opened –> Get handle:
	MN_GETHMENU := 0x1E1 ; Shell Constant: "Menu_GetHandleMenu"
	SendMessage, MN_GETHMENU, False, False
	hCM := ErrorLevel ; Return Handle in ErrorLevel
	ToolTip, % "Hab Dich Menu: " GetHex(hCM) , % mx , % my - 60, 6
	SetTimer, HDM, -3000
return

HDM:
	ToolTip,,,,6
return
;}
StarteAutoDiagnosen:       	;{   	Strg	+ Alt + d
	If !ScriptIsRunning("AutoDiagnosen3")
		ADPID :=  RunSkript(A_ScriptDir "\include\AutoDiagnosen3.ahk")
return ;}
COVIDImpfA:                 	;{
Send, % "{Raw}88331A(charge:" Addendum.COVID.BTChnr ")"
return
COVIDImpfB:
Send, % "{Raw}88331B(charge:" Addendum.COVID.BTChnr ")"
return
COVIDImpfR:
Send, % "{Raw}88331R(charge:" Addendum.COVID.BTChnr ")"
return ;}
Scheinloeschen:            	;{ 	Strg 	+ Alt + Shift + L
	PostMessage, 0x111, % 32868,,, % "ahk_class OptoAppClass" 	; !NUR Menuaufruf, kein sofortiges löschen!
return ;}
Scheinaendern:             	;{ 	Strg 	+ Alt + Shift + Ä
	PostMessage, 0x111, % 32789,,, % "ahk_class OptoAppClass"
return ;}
PrivatAU:                       	;{		Strg	+ Alt + J
	AlbisMenu("Formular/Privat-AU", "Privat-AU ahk_class #32770")
return
;}

Privatabrechnung_Begruendungen(cmd="zeigen", bgrWahl="")       	{         	;-- Begründungstexte aus eigener Datei verwenden

  ; letzte Änderung: 03.02.2022

	static hfocused, hparent, hfocused_last, hparent_last

	switch cmd {

		Case "zeigen":
			hfocused_last 	:= hfocused 	:= GetFocusedControlHwnd(AlbisWinID())
			hparent_last  	:= hparent	:= GetParent(hfocused)
			caretpos           	:= CaretPos(hfocused)
			ctrltxt            	:= SubStr(ControlGetText("", "ahk_id " hfocused), 1, caretpos - 1)

			If FileExist(Addendum.DataPath "\GOÄTexte.json") {
				cJSON.EscapeUnicode := "UTF-8"
				jsonstr := FileOpen(Addendum.DataPath "\GOÄTexte.json", "r", "UTF-8").Read()
				bgr := cJSON.Load(jsonstr)
			}

			RegExMatch(ctrltxt, "(?<Ziffer>\d+)[\s\-]*$", lp)
			Privatabrechnung_BgrGui(bgr[lpZiffer], lpZiffer)


		Case "einfuegen":
			If !hfocused_last
				return

			ControlFocus,, % "ahk_id " hfocused_last
			If !ErrorLevel
				SendRaw, % bgrWahl
			else
				PraxTT("Addendum konnte den Text nicht einfügen", "1 2")

			hfocused_last := hparent_last := ""

	}

}
Privatabrechnung_BgrGui(selection, ziffer)                                        	{          ;-- Gui für die Auswahl von EBM Begründungen

  ; letzte Änderung: 03.02.2022

	global bgr, hbgr
	static bgrLV, bgrUse, bgrCancel

	acx := A_CaretX, acy := A_CaretY

	SysGet, CursorHeight, 14
	SysGet, VScrollWidth	, 2

	Gui, Bgr: New, HWNDhbgr
	Gui, Bgr: Add, ListView	, % "xm ym w700 r10 vbgrLV "                          	, % "Texte"
	For idx, bgrtxt in selection
		LV_Add("", bgrtxt)
	Gui, Bgr: Add, Button	, % "xm y+5	vbgrUse    	gbgrButtonHandler"	, % "Begründungstext übernehmen"
	Gui, Bgr: Add, Button	, % "x+30 	vbgrCancel 	gbgrButtonHandler"	, % "Abbruch"
	Gui, Bgr: Show, % "x " acx " y " acy " AutoSize NA Hide", % "Begründungstexte für GOÄ Ziffer [" ziffer "]"
	win := GetWindowSpot(hbgr)
	LV_ModifyCol(1, 700-VScrollWidth-20)
	Gui, Bgr: Show, % "x " acx " y " acy-win.H " NA"

return

bgrButtonHandler:

	if (A_GuiControl = "bgrUse") {
		Gui, bgr: ListView, bgLV
		If !(cRow  := LV_GetNext(0))
			return
		LV_GetText(bgrWahl, cRow)
		Privatabrechnung_Begruendungen("einfuegen", bgrWahl)
	}

bgrGuiClose:
bgrGuiEscape:
	Gui, bgr: Destroy
return
}
ScheinCOVIDPrivat()                                                                      	{          ;-- Funktionsaufruf für "Neuen Schein anlegen"
	Quartal := QuartalTage("aktuell")
	res := AlbisAbrScheinCOVIDPrivat(Quartal)
}
Schnellziffer(Datum:="")                                                                 	{       	;-- Hotkeys für Zifferneingaben in der Abrechnung

		static SZiffern := {"a":"03220", "s":"03221", "d":"03230(x:2)", "f":"03360", "g":"03362", "h":"35100", "i":"35110"}

		ziffer := SZiffern["#" A_ThisHotkey]
	;~ SciTEOutput(A_ThisFunc ": [" A_ThisHotkey "] " ziffer)

	; Eingaben ermöglichen
		ControlGetFocus, cFocus, ahk_class OptoAppClass
		while (!RegExMatch(cFocus, "^Edit(?<nr>\d+)", Edit) && A_Index < 4) {
			If (A_Index > 1)
				Sleep 500
			SendInput, {Insert}
			Sleep 100
			ControlGetFocus, cFocus, ahk_class OptoAppClass
		}

	; auf Eingabefeld prüfen
		If !RegExMatch(EditNr, "\d+") {
			PraxTT("Eingabefeld konnte nicht erstellt werden", "1 1")
			return 0
		}

	; EBM Ziffer senden
		If !VerifiedSetText(cFocus, ziffer, "ahk_class OptoAppClass") {
			PraxTT("EBM Ziffer [" ziffer "]: ließ sich nicht einfügen", "1 1")
			return 0
		}

	; Eingabe durch TAB abschließen
		SendInput, {TAB}
		Sleep, 300
		ControlGetFocus, cFocus, ahk_class OptoAppClass
		while (RegExMatch(cFocus, "^Edit(?<nr>\d+)", Edit) && A_Index < 4) {
			If (A_Index > 1)
				Sleep 500
			SendInput, {TAB}
			Sleep 100
			ControlGetFocus, cFocus, ahk_class OptoAppClass
		}
		If RegExMatch(cFocus, "^Edit(?<nr>\d+)", Edit) {
			PraxTT("Die Eingabe konnte nicht abgeschlossen werden", "1 1")
			return 0
		}

return 1
}
EBM_COVIDImpfung()                                                                     	{       	;-- Abrechnungshilfe COVID Ziffern

		static EBMSkript := A_ScriptDir "\include\EBM-COVIDImpfung.ahk"

		If !Addendum.Thread["COVIDEBM"].ahkReady()  {

			If !FileExist(EBMSkript) {
				throw Exception("Kann EBM-Skript nicht im Ordner " A_ScriptDir "\include\ finden.")
				return
			}

			Addendum.Thread["COVIDEBM"] := AHKThread(FileOpen(EBMSkript, "r", "UTF-8").Read())

		}

	; Funktion im Thread aufrufen
		If Addendum.Thread["COVIDEBM"].ahkReady()  {

		  ; warten bis die Abrechnungsdaten zusammengestellt sind
			If Addendum.Thread["COVIDEBM"].ahkgetvar.EBMInitRunning {
				SetTimer, EBMInitCheck, 200
				return
			}

		 ; Funktion im Threadskript wird per Postmessage (ahkPostFunction) aufgerufen.
		 ; Das Hauptskript wartet nicht auf Antwort.
			Addendum.Thread["COVIDEBM"].ahkPostFunction("Chargennummern")

		}

return
EBMInitCheck:                                ; SetTimer Label (sleep)

	; im Anschluß die EBM Abrechnungsfunktion aufrufen
		If !Addendum.Thread["COVIDEBM"].ahkgetvar.EBMInitRunning {
			SetTimer, EBMInitCheck, Off
			If EBMCNRunning                 ; Hotkey wurde zu früh gefeuert, Thread arbeitet noch
				return
			Addendum.Thread["COVIDEBM"].ahkPostFunction("Chargennummern")
		}

return
}
LDSars()                                                                                        	{        	;-- Ziffern und Diagnose (EBM) Corona Abstrich

	;~ hparent := AlbisSendText("RichEdit20A1", "Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G}"
										  ;~ . ";Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G};")
	;SendMode, Event
	;SetKeyDelay, 10, 10

	Linedate	:= AlbisZeilenDatumLesen(60, false)
	PrgDate	:= AlbisLeseProgrammDatum()
	AlbisSetzeProgrammDatum(LineDate)

	Sleep 300
	Send, % "{Blind}"
	Sleep 100
	Send, % "{Text}Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G};Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G};"
	Send, {Tab}
	Sleep 200
	WinWait, % "ALBIS ahk_class #32770", % "meldepflichtig", 2
	VerifiedClick("Button1", "ALBIS ahk_class #32770", "meldepflichtig",, true)
	Send, {Esc}
	Sleep 300
	Send, {Down}
	Sleep 300
	Send, % "{Blind}"
	Sleep 100
	Send, % "{Text}lko"
	Send, {Tab}
	Sleep 300
	Send, % "{Blind}"
	Sleep 100
	Send, % "{Text}32006-88240"
	Send, {Tab}
	PraxTT("Circa 4s braucht Albis für Berechnungen!")
	;WinWait
	Sleep 4000
	Send, {Esc}
	AlbisSetzeProgrammDatum(PrgDate)

	;SetKeyDelay, 10, 10

}
AgeDependFee(fee_designation, PatAge:="")                                   	{          	;-- altersabhängige Ziffernausgabe

	static agerules := {"OP-Vorbereitung":{31010:{"min":0, "max":12}, 31011:{"min":13, "max":40}, 31012:{"min":41, "max":60}, 31013:{"min":61, "max":999}}}

	PatAge := PatAge ? PatAge : AlbisPatientAlter("Zeilendatum")
	If (PatAge = -1) {
		MsgBox, 0x1000, Addendum für Albis on Windows, Das Datum der aktuellen Zeile konnte nicht gelesen werden.
		return
	}

	For fee, age in agerules[fee_designation]
		If (PatAge >= age.min && PatAge <= age.max) {
			SciTEOutput("fee: " fee ", " PatAge)
			return fee
		}

return
}

;}
;{ ###############      ALLGEMEIN      	 ###############
KopierenUeberall:         	;{	 	Strg	+ c
	CopyToClipboard := true
AusschneidenUeberall:   	;	 	Strg (links) + x
	;SetKeyDelay, 10, 10
	Clipboard := ""
	while (clip = "")	{
		SendEvent, {CTRL Down}c{CTRL Up}
		ClipWait, 1
		clip := Clipboard
		If (A_Index > 6)
			break
	}
	If (clip != "")	{
		clip := SubStr(ClipBoard, 1, 30) . "`.`.`."
		GDISplash(clip,1)
	}
	else
		GDISplash("Clipboard war leer!",1)
	If !CopyToClipboard
		Send, {CTRL down}x{CTRL up}

	CopyToClipboard	:= false
	clip						:= ""
	;SetKeyDelay, 10, 10

return ;}
DoNothing:                   	;{
return
;}
BereiteDichVor:					;{
	ExitApp
return
;}
AddendumPausieren:     	;{ 	Strg + Alt + Pause
	PraxTT("Addendum ist pausiert", "5 0")
	Pause, Toggle
return ;}
PatientenkarteiOeffnen: 	;{     ;~ Hotkey, ^!p                  	, PatientenkarteiOeffnen           	;= Überall: selektierte Zahlen zum Öffnen einer Patientenkarteikarte nutzen

	PatientenID := Clip()
	If !RegExMatch(PatientenID, "\d+") {
		PraxTT("Der ausgewählte Text (" PatientenID ")`nkann keine Patienten ID sein!", "3 3")
		return
	}
	else if !cPat.Exist(PatientenID)	{
		MsgBox, 0x1004, Addendum für Albis on Windows, % "Die ausgewählte Patienten ID (" PatientenID ")`n ist nicht bekannt. Soll diese ID dennoch benutzt werden?"
		IfMsgBox, No
			return
	}
	AlbisAkteOeffnen(cPat.NAME(PatientenID), PatientenID)

return
;}
SendClipBoardText:       	;{ 	Strg + Alt + F10

	If (StrLen(TextToSend) = 0) {
		TextToSend 	:= ClipBoard
		ClipBoard 	:= ""
	}

	PraxTT("Sende ClipBoard Inhalt.", "3 3")
	Loop, % StrLen(TextToSend) {
		Send, % "{Text}" SubStr(TextToSend, A_Index, 1)
		Sleep, 25
	}
	;Send, % "{Text}" TextToSend
	SetTimer, emptyTextToSend, -180000 ; nach 3min leeren

return
emptyTextToSend:
	TextToSend := ""
return
;}
ScanBeenden:               	;{		Enter
 If (hwnd := WinExist("ScanSnap Manager - Bild erfassen und Datei speichern ahk_class #32770"))
	res := VerifiedClick("Button2", hwnd)
return ;}
;}

;}

;{11. Labels - TrayMenu
SkriptReload(msg:="")                     	{                                                          					;-- besseres Reload eines Skriptes

  ; es wird ein RestartScript gestartet, welches zuerst das laufende Skript komplett beendet und
  ; dann das Addendum-Skript per run Befehl erneut gestartet. Dies verhindert einige Probleme
  ; mit datumsabhängigen Handles bestimmter Prozesse im System.
  ; +Albisprogrammdatum wird auf den nächsten Tag gesetzt

	; Abbruch des Neustart falls Addendum gerade Befunde importiert oder ein OCR Vorgang läuft
		If (Addendum.Importing || Addendum.ImportRunning || Addendum.tessOCRRunning || Addendum.Thread["tessOCR"].ahkReady()) {

			; Protokolltext
				FormatTime	, Time      	, % A_Now    	, dd.MM.yyyy HH:mm:ss
				FormatTime	, TimeIdle	, % A_TimeIdle	, HH:mm:ss
				FileAppend	, % Time ", " A_ScriptName ", " A_ComputerName ", " msg ", terminated due to a running process, " TimeIdle "`n"
									, % Addendum.LogPath "\OnExit-" A_YYYY "-Protokoll.txt"

			; Nutzer fragen
				MsgBox, 4	, % RegExReplace(A_ScriptName, "\.[a-z]+$")
								, % 	"Addendum ist gerade beschäftigt.`n"
									. 	"Ein Abbruch könnte zu fehlerhaften Daten führen.`n"
									.  	"Möchten Sie dennoch einen Neustart durchführen?"
								, 10
				IfMsgBox, No
					Return
				IfMsgBox, Timeout
					Return

		}

	; Programmdatum auf aktuelles Datum setzen
		If (msg = "AutoRestart") && WinExist("ahk_class OptoAppClass")
			AlbisSetzeProgrammDatum()

		admScript	:= Addendum.Dir "\Module\Addendum\Addendum.ahk"
		scriptPID 	:= DllCall("GetCurrentProcessId")
		__          	:= q " " q                                             	; für bessere Lesbarkeit des Codes - ergibt > " " <

	; AutohotkeyH_U64.exe /f "Z:\Addendum für AlbisOnWindows\Addendum_Reload.ahk" "Z:\Addendum für AlbisOnWindows\Module\Addendum\Addendum.ahk" "1" "0x34A567" "9876" " 1"
		Run, % Addendum.AHK_HPath " /f " q Addendum.Dir "\include\Addendum_Reload.ahk" __ admScript __ "1" __ A_ScriptHwnd __ scriptPID __ "1" q
		ExitApp  ; damit wird die OnExit Funktion ausgeführt

return
}
ShowTextProtocol(FilePath)             	{                                                                          	;-- zeigt Textdateien im eingestellten Texteditor an
	Run % FilePath
	If RegExMatch(FilePath, "i)Fehlerprotokoll") && (hSciteWin := WinExist("ahk_class SciTEWindow")) {

		while (!(SciteExist := WinGetTitle(hSciteWin) ~= "i)Fehlerprotokoll") && A_Index < 1000)
			Sleep 10

		If SciteExist {
			WinMenuSelectItem, % "ahk_id " hSciteWin,, % "Sprache", % "Autohotkey"
		}

	}
}
ZeigePatDB()                                    	{                                            		                   			;-- zeigt in einer eigenen Listview die Patienten.txt
	For PatID, Pat in cPat.Patients
		AutoListview("Pat.Nr.|Nachname, Vorname|Geburtsdatum|Geschlecht|LASTBEH|SEIT"
						, PatID "," Pat.name "," Pat.Vorname "," Pat.GEBURT "," Pat.GESCHLECHT "," Pat.LASTBEH "," Pat.SEIT ",")
}
ZeigeFehlerProtokoll()                     	{                                                             					;-- das Skriptfehlerprotokoll wird angezeigt

	filebasename    	:= Addendum.LogPath "\ErrorLogs\Fehlerprotokoll-"
	thisMonthProtocol	:=  filebasename . A_MM . A_YYYY ".txt"
	lastMonthProtocol	:=  filebasename (A_MM-1 = 0 ? "12" A_YYYY-1 : A_MM-1 . A_YYYY) ".txt"

	If FileExist(thisMonthProtocol)
		Run % thisMonthProtocol
	else if FileExist(lastMonthProtocol)
		Run % lastMonthProtocol
	else
		PraxTT(StrReplace(A_ScriptName, ".ahk") "`nEs wurden keine Fehler protokolliert!", "5 1")

return
}
NoNo()                                         	{                             	                                 				;-- Addendum "Über" Fenster

	static UberPic, UberSN, UberEdit, UberOk, hUberText, hUber

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Info Text
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		D1 := "Addendum V" Version "  - The managing script -"
		D2 =
		(LTrim
		Addendum für Albis on Windows
		written by Ixiko - this version is from %vom%
		---------------------------------------------------------------------------------
		please report errors and suggestions to me: Ixiko@mailbox.org
		use subject: 'Addendum' so that you don't end up in the spam folder
		GNU Lizenz can be found in Docs directory  - 2017
		)

	UberW	:= 526
	ul     	:= "__________________________________________________________________________________________________________________"

	Betriebssystem := StrSplit(A_OsVersion, ".").1 " " (A_Is64bitOS=1 ? "64" : "32")
	C .= "installierte Autohotkeyversion: " 	A_AhkVersion                                  	"`n"
	C .= "Computer Name: "                   	A_ComputerName                        	"`n"
	C .= "IP Adresse des Computer: "       	A_IPAddress1                                 	"`n"
	C .= "Login Name: "                          	A_UserName                                 	"`n"
	C .= "Nutzer hat Adminrechte: "         	(A_IsAdmin=1 ? "Ja" : "Nein")         	"`n"
	C .= "Addendum Hauptverzeichnis: " 	Addendum.Dir                               	"`n"
	C .= "Skript ist compiliert: "                	(A_IsCompiled=1 ? "Ja" : "Nein")    	"`n"
	C .= "Skripthandle: "                         	A_ScriptHwnd                                	"`n"
	C .= "Betriebssystem: Windows "         	Betriebssystem "-bit"                         	"`n"
	C .= "Bildschirm-DPI: "                      	A_ScreenDPI                                  	"`n"
	C .= "Bildschirmgröße in Pixel: "         	A_ScreenWidth "x" A_ScreenHeight 	"`n"
	C .= "letzte Fehlernummer: "              	A_LastError

	MaxLen := 0
	Loop, Parse, C, `n
		MaxLen := StrLen(A_LoopField) > MaxLen ? StrLen(A_LoopField) : MaxLen

	B := "`n" D2 "`n" SubStr(ul, 1, MaxLen - 20) "`n" C

	Gui Uber: New, -Caption +ToolWindow  0x94400000 HWNDhUber
	Gui Uber: Color, c202842
	Gui Uber: Margin, 0, 10

	Gui Uber: Add, Picture, % "x" (Floor(UberW/2) - 250) " ym vUberPic", % Addendum.Dir "\assets\AddendumBigLogo.jpg"
	GuiControlGet, UberPic, Uber: Pos

	Gui Uber: Add, Progress, % "x13 y" (UberPicY+UberPicH+3) " w" (UberW-26) " -E0x00020000 h36 c5669ad" , 100

	Gui Uber: Font, s18 q5 cWhite, Futura Bk Bt
	Gui Uber: Add, Text, % "x0 y" ( UberPicY+UberPicH+5 ) " w" UberW " vUberSN BackgroundTrans Center", % D1
	GuiControlGet, UberSN, Uber: Pos

	Gui Uber: Font, s9 q5 cWhite, Futura Bk Bt
	Gui Uber: Add, Edit, % "x0 y" ( UberPicY+UberPicH+40 ) " w" UberW " vUberEdit ReadOnly -WantCtrlA -WantReturn -Wrap -0x002000C0 +0x00000001 -E0x00000200 HWNDhUberText", % B
	GuiControlGet, UberEdit, Uber: Pos

	Gui Uber: Add, Progress, % "x13 y" ( UberEditY+UberEditH ) " w" (UberW-26) " -E0x00020000 h5 c5669ad" , 100

	Gui Uber: Font, s10 q5 cWhite, Futura Bk Bt
	Gui Uber: Add, Button, % "xs y" ( UberEditY+UberEditH+12 ) " vUberOk gUberEnde", Schließen
	GuiControlGet, UberOk, Uber: Pos
	GuiControl, Move, UberOk, % "x" ( (UberW//2) - (UberOkW//2) )

	ControlFocus, %hUberText%
	SendInput, ^{HOME}
	ControlFocus, UberOk, % "ahk_id " hUber
	SendInput, {Left}

	Gui Uber: Show, AutoSize, Addendum für Albis on Windows
	DllCall("HideCaret","Int", hUberText)

return

UberEnde:
	If (A_GuiControl="UberOk")
			Gui Uber: Destroy
return
}
AddendumObjektSpeichern()         	{                                                             					;-- speichert und zeigt das Addendum Objekt an

	If IsObject(Addendum) {
		cJSON.EscapeUnicode := "UTF-8"
		PraxTT("Speichere Addendum Objekt", "0 1")
		FileOpen(A_Temp "\AddendumObjekt.json", "w", "UTF-8").Write(cJSON.Dump(Addendum, 1))
		PraxTT("Zeige das Addendum Objekt an.", "0 1")
		Run, % A_Temp "\AddendumObjekt.json"
		PraxTT("", "Off")
	}

}
ShowPDFObjects()                          	{

	global ScanPool, PatDocs

	If IsObject(Addendum) {
		cJSON.EscapeUnicode := "UTF-8"
		PraxTT("Speichere PDF Pooling Objekte", "0 1")
		PatDocsPrint  	:= PatDocs.Count() 	? cJSON.Dump(PatDocs, 1) 	: "{" q "PatDocs Objekt" 	q ":" q "ist leer" q "}"
		ScanPoolPrint 	:= ScanPool.Count()	? cJSON.Dump(ScanPool, 1)	: "{" q "ScanPool Objekt" 	q ":" q "ist leer" q "}"
		FileOpen(A_Temp "\PDFObjekte.json", "w", "UTF-8").Write(ScanPoolPrint "`n`n" PatDocsPrint)
		PraxTT("Zeige PDF Pooling Objekte an", "0 1")
		Run, % A_Temp "\PDFObjekte.json"
		PraxTT("", "Off")
	}

}
scriptVars:                                       	;{                                                                            	;-- zeigt die Variablen des AddendumSkriptes
	ListVars
return ;}

; ----------------------------------------- Tray Menu Programmstarter
ModulStarter(mName, mPath, ItemPos, MenuName) 	{                                                    	;-- startet Module vom Tray-Menu
	PraxTT("Starte: " mName, "5 1")
	If FileExist(mPath)
		Ires := RegExMatch(mPath, "\.exe$") ? RunExe(mPath) : RunSkript(mPath)
return lres
}
ToolStarter(mName, mPath, ItemPos, MenuName) 	{                                                       	;-- startet Tools vom Tray-Menu
	If InStr(mName, "Windows komplett neu starten") {
		PraxTT("Windows wird gleich herunter gefahren und neu gestartet!", "5 1")
		Run, % "Shutdown -g -t 0"
		return
	}
	If FileExist(mPath) {
		PraxTT("Starte Tool: " mName, "5 1")
		Ires := RegExMatch(mPath, "\.exe$") ? RunExe(mPath) : RunSkript(mPath)
	}
return lres
}

; ----------------------------------------- Funktionen ein- oder ausstellen
Menu_Autologin:                             	;{ ermöglicht automatisiertes Einloggen in Albis
	If (Addendum.AutoLogin = " (Logindaten fehlen)")
		return
	If (!AlbisDefaultLogin("User") || !AlbisDefaultLogin("Password")) {
		IniWrite, "-|-", % Addendum.ini, Computer, % compname
		Addendum.AutoLogin := " (Logindaten fehlen)"
		Menu, SubMenu36, UnCheck, % "Albis Autologin" Addendum.AutoLogin
	} else {
		Addendum.AutoLogin := !Addendum.AutoLogin
		Menu, SubMenu36, % (Addendum.Autologin ? "Check":"UnCheck") , % "Albis Autologin"
	}
	Addendum.AutoLoginCanChanged := (!Addendum.AutoLogin || Addendum.AutoLogin = " (Logindaten fehlen)") ? false : true
	IniWrite, % (!Addendum.Autologin ? "nein" : Addendum.Autologin = 1 ? "ja" : Addendum.AutoLogin), % Addendum.ini, % compname, % "Albis_Autologin"
return
;}
Menu_AlbisAutoPosition:                 	;{
	Addendum.AlbisLocationChange := !Addendum.AlbisLocationChange
	Menu, SubMenu31, ToggleCheck, % "Albis AutoSize"
	IniWrite, % (Addendum.AlbisLocationChange ? "ja":"nein"), % Addendum.Ini, % compname, % "Albis_AutoGroesse"
return
;}
Menu_Wartezimmer_AutoDelete:    	;{ Entfernen ohne Nachfrage
	Addendum.AutoDelete.WZ := !Addendum.AutoDelete.WZ
	Menu, SubMenu366, ToggleCheck, % "schnelles Löschen im Wartezimmer"
	IniWrite, % (Addendum.AutoDelete.WZ ? "ja":"nein")        	, % Addendum.Ini, % compname, % "Wartezimmer_schnell_loeschen"
return ;}
Menu_Dauermed_AutoDelete:        	;{ Entfernen ohne Nachfrage
	Addendum.AutoDelete.Dauermed := !Addendum.AutoDelete.Dauermed
	Menu, SubMenu366, ToggleCheck, % "schnelles Löschen von Dauermedikamenten"
	IniWrite, % (Addendum.AutoDelete.Dauermed ? "ja":"nein")	, % Addendum.ini, % compname, % "Dauermedikamente_schnell_loeschen"
return ;}
Menu_Dauerdia_AutoDelete:         	;{ Entfernen ohne Nachfrage
	Addendum.AutoDelete.Dauerdia := !Addendum.AutoDelete.Dauerdia
	Menu, SubMenu366, ToggleCheck, % "schnelles Löschen von Dauerdiagnosen"
	IniWrite, % (Addendum.AutoDelete.Dauerdia ? "ja":"nein")	, % Addendum.ini, % compname, % "Dauerdiagnosen_schnell_loeschen"
return ;}
Menu_Schnellrezept:                      	;{
	Addendum.Schnellrezept := !Addendum.Schnellrezept
	Menu, SubMenu36, ToggleCheck, % "Schnellrezept"
	IniWrite, % (Addendum.Schnellrezept ? "ja":"nein")  	, % Addendum.Ini, % "Addendum", % "Schnellrezept"
return ;}
Menu_AddendumGui:                    	;{ Infofenster
	Addendum.AddendumGui := !Addendum.AddendumGui
	Menu, SubMenu32, ToggleCheck, % "Infofenster"
	IniWrite, % (Addendum.AddendumGui ? "ja":"nein"), % Addendum.Ini, % compname, % "Infofenster_anzeigen"
	If Addendum.AddendumGui {
		Menu, SubMenu32, Add, % "Einzelbestätigung für Importvorgänge"		     			, Menu_Einzelbestaetigung
		Menu, SubMenu32, % (Addendum.iWin.ConfirmImport	? "Check":"UnCheck")	, % "Einzelbestätigung für Importvorgänge"
		Menu, SubMenu32, Add, % "Autozuweisung Wartezimmer"	                        		, Menu_AutoWZ
		Menu, SubMenu32, % (Addendum.iWin.AutoWZ         	? "Check":"UnCheck")	, % "Autozuweisung Wartezimmer"
		Menu, SubMenu32, Add, % "Abrechnungshelfer"		     	                             		, Menu_Abrechnungshelfer
		Menu, SubMenu32, % (Addendum.iWin.AbrHelfer       	? "Check":"UnCheck")	, % "Abrechnungshelfer"
		If RegExMatch((Addendum.AktiveAnzeige := AlbisGetActiveWindowType()), "i)(Karteikarte|Laborblatt|Biometriedaten)")
			AddendumGui()
	} else {
		Menu, SubMenu32, Delete, % "Einzelbestätigung für Importvorgänge"
		Menu, SubMenu32, Delete, % "Autozuweisung Wartezimmer"
		Menu, SubMenu32, Delete, % "Abrechnungshelfer"
		Controls("", "reset", "")
		If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {
			admGui_SaveStatus()
			admGui_Destroy()
		}
	}
return ;}
Menu_Einzelbestaetigung:               	;{ Infofenster
	Addendum.iWin.ConfirmImport := !Addendum.iWin.ConfirmImport
	Menu, SubMenu32, ToggleCheck, % "Einzelbestätigung für Importvorgänge"
	IniWrite, % (Addendum.iWin.ConfirmImport 	? "ja":"nein"), % Addendum.Ini, % compname	, % "Infofenster_Import_Einzelbestaetigung"
return ;}
Menu_AutoWZ:                             	;{ Infofenster
	Addendum.iWin.AutoWZ := !Addendum.iWin.AutoWZ
	Menu, SubMenu32, ToggleCheck, % "Autozuweisung Wartezimmer"
	IniWrite, % (Addendum.iWin.AutoWZ	? "ja":"nein"), % Addendum.Ini, % "Infofenster"	, % "AutoWZ"
return ;}
Menu_Abrechnungshelfer:              	;{ Infofenster
	Addendum.iWin.AbrHelfer := !Addendum.iWin.AbrHelfer
	Menu, SubMenu32, ToggleCheck, % "Abrechnungshelfer "
	IniWrite, % (Addendum.iWin.AbrHelfer ? "ja":"nein"), % Addendum.Ini, % compname, % "Infofenster_Abrechnungshelfer"
return ;}
Menu_GVUAutomation:                   	;{
	Addendum.GVUAutomation := !Addendum.GVUAutomation
	Menu, SubMenu36, ToggleCheck, % "GVU automatisieren"
	IniWrite, % (Addendum.GVUAutomation ? "ja":"nein")        	, % Addendum.Ini, % compname, % "GVU_automatisieren"
return
;}
Menu_PDFSignatureAutomation:     	;{ FoxitReader
	; Menupunkt umschalten, Einstellung speichern
		Addendum.PDFSignieren := !Addendum.PDFSignieren
		Menu, SubMenu37, ToggleCheck, % "Signaturhilfe"
		IniWrite, % (Addendum.PDFSignieren ? "ja":"nein")          	, % Addendum.Ini, % compname, % "FoxitPdfSignature_automatisieren"
	; weitere Funktionen an- oder abschalten
		If Addendum.PdfSignieren {

			; Hotkey Funktion einschalten
				;func_FRSignature := Func("FoxitReader_SignaturSetzen")
				Hotkey, IfWinActive, % "ahk_class classFoxitReader"
				Hotkey, % Addendum.PDFSignieren_Kuerzel, % Func("FoxitReader_SignaturSetzen")
				Hotkey, IfWinActive

			; AutoCloseTab - Tray Menu - anschalten
				Menu, SubMenu37, Enable, % "signiertes Dokument schliessen"
				Menu, SubMenu37, % (Addendum.AutoCloseFoxitTab ? "Check" : "UnCheck")  , % "signiertes Dokument schliessen"

		} else {

			; Hotkey Funktion ausschalten
				Hotkey, % Addendum.PDFSignieren_Kuerzel, Off

			; AutoCloseTab - Tray Menu - ausschalten
				Menu, SubMenu37, Disable, % "signiertes Dokument schliessen"

		}
return ;}
Menu_PDFSignatureAutoTabClose:	;{ FoxitReader
	Addendum.AutoCloseFoxitTab := !Addendum.AutoCloseFoxitTab
	Menu, SubMenu37, ToggleCheck, % "signiertes Dokument schliessen"
	IniWrite, % (Addendum.AutoCloseFoxitTab ? "ja":"nein") 	, % Addendum.Ini, % compname, % "FoxitTabSchliessenNachSignierung"
return ;}
Menu_LaborAbrufManuell:              	;{
	Addendum.Labor.AutoAbruf := !Addendum.Labor.AutoAbruf
	Menu, SubMenu34, % (Addendum.Labor.AutoAbruf ? "Check" : "UnCheck"), % "Laborabruf"
	IniWrite, % (Addendum.Labor.AutoAbruf ? "ja":"nein")   	, % Addendum.Ini, % compname, % "Laborabruf_automatisieren"
return ;}
Menu_LaborImportManuell:            	;{
	Addendum.Labor.AutoImport := !Addendum.Labor.AutoImport
	Menu, SubMenu34, ToggleCheck, % "Laborimport"
	IniWrite, % (Addendum.Labor.AutoImport ? "ja":"nein")   	, % Addendum.Ini, % compname, % "Laborimport_automatisieren"
return ;}
Menu_LaborAbrufTimer:                  	;{
	Addendum.Labor.AbrufTimer := !Addendum.Labor.AbrufTimer
	Menu, SubMenu34, ToggleCheck, % "zeitgesteuerter Abruf"
	IniWrite, % (Addendum.Labor.AbrufTimer ? "Ja":"Nein")   	, % Addendum.Ini, % compname, % "Laborabruf_Timer"
	If Addendum.Labor.AbrufTimer && !Addendum.Labor.AutoAbruf  {
		Addendum.Labor.AutoAbruf := !Addendum.Labor.AutoAbruf
		Menu, SubMenu34, ToggleCheck, % "Laborabruf"
		IniWrite, % "ja", % Addendum.Ini, % compname, % "Laborabruf_automatisieren"
	}
	If Addendum.AddendumGui && admGui_Exist()
		admGui_CountDown()
return ;}
Menu_AddendumToolbar:             	;{
	Addendum.ToolbarThread := !Addendum.ToolbarThread
	Menu, SubMenu36, ToggleCheck, % "Addendum Toolbar"
	IniWrite, % (Addendum.ToolbarThread ? "ja":"nein"), % Addendum.Ini, % compname, % "Addendum_Toolbar"
	admThreads()                                                                             ; startet oder stoppt die Ausführung von Threadcode
return ;}
Menu_mDicom:                              	;{
	Addendum.mDicom := !Addendum.mDicom
	Menu, SubMenu3, ToggleCheck, % "MicroDicom Export automatisieren"
	IniWrite, % (Addendum.mDicom ? "Ja":"Nein")         	, % Addendum.Ini, % compname, % "MicroDicomExport_automatisieren"
return ;}
Menu_PopUpMenu:                         	;{
	Addendum.PopUpMenu := !Addendum.PopUpMenu
	Menu, SubMenu36, ToggleCheck, % "PopUpMenu"
	IniWrite, % (Addendum.PopUpMenu ? "ja":"nein") 	, % Addendum.Ini, % compname, % "PopUpMenu"
  ; PopUpMenu wird neugestartet
	If Addendum.PopUpMenu {
		Addendum.Hooks[4] := AlbisPopUpMenuHook(AlbisPID())
		PraxTT("PopUpMenu Funktion - gestartet", "3 0")
	}
	else { ; oder auch beendetm wenn Nutzer diese nicht braucht
		If IsObject(Addendum.Hooks[4]) {
			UnhookWinEvent(Addendum.Hooks[4].hEvH, Addendum.Hooks[4].HPA)
			Addendum.Hooks[4] := ""
			Addendum.PopUpMenuCallback := ""
			PraxTT("PopUpMenu Funktion - beendet", "3 0")
		}
		else
			PraxTT("PopUpMenu Funktion - war nicht gestartet`nund konnte deshalb nicht beendet werden.", "3 0")
	}
return ;}
Menu_AutoSize:                             	;{
	; 0 = keines , 1 = 2k , 2 = 4k , 3 = 2k & 4k ..... "2k:nein,4k:nein"
	Menu, SubMenu31, ToggleCheck, % A_ThisMenuItem
	IniWrite, % (xset := GetSetAutoPosString(A_ThisMenuItem)), % Addendum.Ini, % compname, % "AutoPos_" A_ThisMenu
return
GetSetAutoPosString(appname) {                                ;-- wandelt die binär gespeicherte AutoPos Einstellung in lesbaren Text um

	ap := Addendum.Windows[appname].AutoPos
	ms := Addendum.MonSize
	If (ms = 2) {
		do1 := 0x1	& ap
		do2 := ms	& ap ? 0x0 : 0x2
	} else {
		do1 := ms	& ap ? 0x0 : 0x1
		do2 := 0x2	& ap
	}

	Addendum.Windows[appname].AutoPos := do1 + do2

return "2k:" (do1 ? "ja":"nein") ",4k:" (do2 ? "ja":"nein")
}  ;}
Menu_AutoOCR:                            	;{
	Addendum.OCR.AutoOCR := !Addendum.OCR.AutoOCR
	Menu, SubMenu33, ToggleCheck, % "AutoOCR"
	IniWrite, % (Addendum.OCR.AutoOCR ? "ja":"nein"), % Addendum.Ini, % "OCR", % "AutoOCR"
return ;}
Menu_WatchFolder:                        	;{
	Addendum.OCR.WatchFolder := !Addendum.OCR.WatchFolder
	Menu, SubMenu33, ToggleCheck, % "Befundordner überwachen"
	IniWrite, % (Addendum.OCR.WatchFolder ? "ja":"nein"), % Addendum.Ini, % compname, % "BefundOrdner_ueberwachen"
	If !Addendum.OCR.WatchFolder {
		WatchFolder("**End", 0)
		Addendum.OCR.WatchFolderStatus := "off"
		PraxTT("Überwachung des Befundordner angehalten", "2 1")
	}
	else {
		WatchFolder(Addendum.BefundOrdner, "admGui_FolderWatch", False, (1+2+8+64))
		Addendum.OCR.WatchFolderStatus := "running"
		PraxTT("Überwachung des Befundordner gestartet", "2 1")
	}
return ;}
Menu_Laborjournalanzeige:           	;{
	Addendum.Laborjournal.AutoView := !Addendum.Laborjournal.AutoView
	Menu, SubMenu35, ToggleCheck, % "Täglich anzeigen"
	IniWrite, % (Addendum.Laborjournal.AutoView ? "Ja":"Nein")   	, % Addendum.Ini, % compname, % "Laborjournal_AutoAnzeige"
	If Addendum.Laborjournal.AutoView && !Addendum.Laborjournal.StartTime
		gosub Menu_LaborjournalStartzeit
return ;}
Menu_LaborjournalStartzeit:            	;{

return ;}
Menu_FaxMailManager:                 	;{
	Addendum.Mail.FaxMail := !Addendum.Mail.FaxMail
	Menu, SubMenu38, ToggleCheck, % "Outlook FaxMailManager"
	IniWrite, % (Addendum.OCR.AutoOCR ? "ja":"nein"), % Addendum.Ini, % "OCR", % "AutoOCR"
return ;}
Menu_EBMKRW:                         		;{
	Addendum.EBMKRW := !Addendum.EBMKRW
	Menu, SubMenu39, ToggleCheck, % "EBM/KRW Hinweise"
	IniWrite, % (Addendum.EBMKRW ? "ja":"nein"), % Addendum.Ini, % compname, % "Pruefung_EBM_KRW_AutoClose"
return ;}
Menu_EBMKRWTimeout:                	;{
	InputBox, Timeout, % StrReplace(A_ScriptName, ".ahk"), % "Hier können Sie den Timeout für das automatische Schließen`n"
																					. 	  "des 'Prüfung EBM/KRW Hinweise' Fensters festlegen.`n "
																					. 	  "eingestelltes Timeout: " Addendum.EBMKRWTimeout "s",, 500, 200,,,, 20, % Addendum.EBMKRWTimeout
	RegExMatch(Timeout, "(?<out>\d+)", t)
	If (!tout || tout=Addendum.EBMKRWTime || ErrorLevel) {
		PraxTT("Sie haben die Timeout-Einstellung nicht geändert", "3 1")
		return
	}
  ; umbenennen des Menu um die eingestellte Sekundzahl anzuzeigen
	Menu, SubMenu39, Rename, % "Timeout EBM/KRW Hinweise (" Addendum.EBMKRWTimeout "s)", % "Timeout EBM/KRW Hinweise (" tout "s)"
	IniWrite, % (Addendum.EBMKRWTimeout := tout), % Addendum.ini, % compname, % "Pruefung_EBM_KRW_Timeout"
return ;}
Menu_DICOM_AutoCopy:              	;{
	Addendum.Dicom.AutoCopy := !Addendum.Dicom.AutoCopy
	Menu, SubMenu40, % (Addendum.Dicom.AutoCopy ? "Check":"UnCheck"), % "Kopie anlegen automatisieren"
	IniWrite, % (Addendum.Dicom.AutoCopy ? "ja":"nein"), % Addendum.Ini, % compname, % "DicomCD_AutoCopy"
return  ;}
Menu_DICOM_Dir:                        	;{

return  ;}
Menu_DICOM_DirPattern:               	;{

return  ;}
Menu_TrayTips:                             	;{
	Addendum.ShowTrayTips := !Addendum.ShowTrayTips
	Menu, SubMenu3, % (Addendum.ShowTrayTips ? "Check":"UnCheck"), % "TrayTips zeigen"
	IniWrite, % (Addendum.ShowTrayTips ? "ja":"nein"), % Addendum.Ini, % compname, % "TrayTips_zeigen"
return ;}
Menu_Faxbox:                              	;{
	Addendum.PDF.Faxbox.State := Addendum.PDF.Faxbox.State = "An" ? "Aus" : "An"
	Menu, SubMenu3, % (Addendum.PDF.Faxbox.State = "An" ? "Check":"UnCheck"), % "Faxe abrufen"
	IniWrite, % (Addendum.PDF.Faxbox.State), % Addendum.Ini, % "ScanPool", % "Faxe abrufen"
	If (Addendum.PDF.Faxbox.State = "An" && compname=Addendum.PDF.Faxbox.Client) {
		PraxTT("Faxe werden abgerufen", "2 0")
		SetTimer, % fn_Faxbox, % Addendum.PDF.Faxbox.Interval * 60 * 1000 ; Angabe in Minuten
		Faxbox(false)     ;  erster Abruf sofort
	} else {
		PraxTT("Faxabruf abgeschaltet", "2 0")
		SetTimer, % fn_Faxbox, % "Off"
	}
return ;}
Menu_ClipWatchPatients:               	;{
	Addendum.Clip.WatchPatients := !Addendum.Clip.WatchPatients
	Menu, SubMenu45, ToggleCheck, % "Patientendaten erkennen"
	IniWrite, % (Addendum.Clip.WatchPatients ? "ja":"nein")        	, % Addendum.Ini, % compname, % "Clipboard_auf_Patientendaten_überwachen"
return
;}
;}

;{12. Allgemeine und spezielle Funktionen

;Infofenster
Splash(string, time=4, x="", y="", w="", ParentHWnd=0) {                        	            	;-- Splash Gui - to show text for some seconds

	static hsp

	CoordMode, Mouse, Screen
	CoordMode, Caret, Screen

	If !x {
		Lmx:= A_CaretX, Lmy:= A_CaretY
		If !Lmx or !Lmy
			MouseGetPos, Lmx, Lmy
		Lmy -= 10
	}	else {
		Lmx := x, Lmy := y, Lmw := w
	}

	Lc := StrSplit(string, "`n").MaxIndex()
	Lc := Lc > 4 ? 4 : !Lc ? 1 : Lc
	FSize := 10
	mx := 8
	my := 5

	; build gui
	Gui sp: New  	, % "-SysMenu -Caption +ToolWindow " (ParentHWnd = 0 ? "AlwaysOnTop" : "+Parent" ParentHwnd) " +HWNDhsp"
	Gui sp: Font  	, % "s" FSize "  q5 cWhite bold", Futura Bk Md
	Gui sp: Color	, cC57773
	Gui sp: Margin	, % mx, % my
	Gui sp: Add   	, Text, % "xm ym" (Lmw ? " w" Lmw-2*mx :"") " R" Lc " Center BackgroundTrans hSplashText" , % string
	Gui sp: Show	, % "x" Lmx " y" (Lmy - FSize*(Lc) - 5 ) . (Lmw ? " w" Lmw :"") " Hide NA" , Addendum Splash

	; form a rounded rectangle
	SP := GetWindowSpot(hsp)
	WinSet, Region, % "0-0 w" (SP.W) " h" (SP.H) " R20-20", % "ahk_id " hsp
	DllCall("AnimateWindow", "UInt",hSp, "Int",500, "UInt",0x80000)

	; close with delay
	SetTimer, SPClosePopup, % "-" (time*1000)

return

SPClosePopup:
	DllCall("AnimateWindow", "UInt",hSp, "Int",500, "UInt",0x90000)
	Gui sp: Destroy
return

}
GDISplash(text, time=2, SplashPos="", SplashOpt="") {                                          	;-- Hinweisfenster

		static hGdiSp, GdiSp

		If !IsObject(SplashOpt)
			hFont := CreateFont("y5 Centre cff000000 q5 s16, " Addendum.Standard.Font), bgColor := "0x778fA2ff"
		else
			hFont := CreateFont(SplashOpt.FontStyle ", " SplashOpt.Font), bgColor := SplashOpt.bgColor

		If !IsObject(SplashPos) {
			cx := A_CaretX , cy := A_CaretY - 35
			If (!cx or !cy) {
				MouseGetPos, cx, cy
				cy += 20
			}
		}
		else
			cx := SplashPos.X, cy := SplashPos.Y

		Gui, GdiSp: New, -Caption -SysMenu -DPIScale +ToolWindow +E0x80000 +HwndhGdiSp

		GetFontTextDimension(hFont, Text, Width, Height ,1)
		Width	:= Floor(Width // 1.35)
		Height	:= Floor(Height // 1.15)
		hbm 	:= CreateDIBSection(Width, Height)
		hdc 		:= CreateCompatibleDC()
		obm 	:= SelectObject(hdc, hbm)
		G 		:= Gdip_GraphicsFromHDC(hdc)

		Gdip_SetSmoothingMode(G, 4)
		pBrush 	:= Gdip_BrushCreateSolid(bgColor)

		Gdip_FillRoundedRectangle(G, pBrush, 0, 0, Width, Height, 5)
		Gdip_TextToGraphics(G, text, Options, Font, Width, Height)
		UpdateLayeredWindow(hGdiSp, hdc, cx, cy, Width, Height)

		Gdip_DeleteBrush(pBrush), DeleteObject(hbm), DeleteDC(hdc), Gdip_DeleteGraphics(G)
		Gui, GdiSp: Show, % "x" cx " y" cy " NA", GdiSp

		SetTimer, GDISplashOff, % -1 * (time*1000)

return hGdiSp

GDISplashOff:
		Gui, GdiSp: Destroy
return
}
Tooltip(content, wait:=2) {
	Splash(content, wait)
}
TTipOff(nr, wait:=2000) {
static ttipnr
ttipnr := nr
SetTimer, TTipClose, %  -1*wait
return
TTipClose:
ToolTip,,,, % ttipnr
return
}

;Text senden
SendRawFast(string, cmd="") {                                                                               	;-- sendet eine String

	; Parameter: go - L, R, U, D - stehen für Left, Right, Up, Down, nach diesen kann eine Zahl für die Anzahl der Tastendrücke stehen

	msg := string
	SendRaw	, % msg

	RegExMatch(cmd, "i)(?<go>[LRUD]|(Enter))\s*(?<do>\d+)*", k)
	SendInput	, % (	kgo="L"    	? "{Left "
		            : 	kgo="R"    	? "{Right "
		            : 	kgo="U"   	? "{Up "
		            : 	kgo="D"   	? "{Down "
		            : 	kgo="Enter"	? "{Enter " :  	"")	. 	kdo "}"

	Splash(string (cmd ? " [" cmd "]" : ""), 1)
}
SendAlternately(chars:="*/|/*") {                                                                              	;-- sendet alternierend Zeichen

	static alternating:= []

	SendRaw, % ""
	For alterIndex, alterChars in alternating
		If (alterChars = chars) {
			msg := StrSplit(chars, "|").2
			SendRaw, %  msg
			alternating.RemoveAt(alterIndex)
			return
		}
	msg := StrSplit(chars, "|").1
	SendRaw, % msg
	alternating.Push(chars)

}

;verschiedenes
SciTEFoldAll() {                                                                                                        	;-- gesamten Code zusammenfalten
	ControlGet, hScintilla1, hwnd,, Scintilla1, % "ahk_class SciTEWindow"
	Sleep 300
	SendMessage, 2662,,,, % "ahk_id " hScintilla1 ; SCI_FOLDALL
}
RunExe(exePath) {                                                                                                    	;-- startet eine ausführbare Datei

	SplitPath, exePath, filename, filepath
	If !FileExist(exepath) {
		PraxTT("Das Modul: " filename "`nim Pfad:`n" filepath "`nist nicht vorhanden!", "6 3")
		return 0
	}

	Run, % exepath, % filepath, ModulPID

return ModulPID
}
RunSkript(modulpath) {                                                                                          	;-- startet oder beendet Autohotkeyskripte

	SplitPath, modulpath, filename, filepath
	If !FileExist(modulpath) {
		PraxTT("Das Modul: " filename "`nim Pfad:`n" filepath "`nist nicht vorhanden!", "6 3")
		return 0
	}

	If (PID := ScriptIsRunning(filename := StrReplace(filename, ".ahk")))  {
		MsgBox, 0x1004, Addendum für Albis on Windows, % "Das Modul '" filename "' ist schon gestartet!`nMöchten Sie das Modul stattdessen beenden?"
		IfMsgBox, Yes
		{
		   	; sendet exitapp-Befehl als 1.Versuch ; Exitapp 65405
				PostMessage, 0x111, 65405,,, % "ahk_pid " PID
				Process, WaitClose, % PID, 6
				Process, Exist , % PID
				If ErrorLevel {
					PostMessage, 0x111, 65405,,, % "ahk_pid " PID
					Process, WaitClose, % PID, 6
				}

		   	; gnadenloses Beenden, weil sich der Prozeß aufgehängt hat. Das war auch der Grund den Modulstarter um diese Fähigkeit zu erweitern!
				If !ErrorLevel	{
					Process, Close		, % PID
					Process, WaitClose, % PID, 4
					If !ErrorLevel
                           	MsgBox, 1, Addendum für Albis on Windows, % "Das Modul '" filename "' konnte nicht beendet werden!"
				}
		}
		return
	}

	Run, % A_AhkPath " " q (!RegExMatch(modulpath, "^[A-Z]\:\\") ? Addendum.Dir "\" : "") modulpath q,, ModulPID

return ModulPID
}
GetPatNameFromExplorer() {                                                                                 	;-- ermittelt aus pdf Dateibezeichnung den Namen, öffnet die KK
	; Voraussetzung: der Patient ist in der Addendum Patientendatenbank gelistet, ansonsten passiert nichts
		SplitPath, % Explorer_GetSelected(WinExist("A")),,,, fname
		FuzzyKarteikarte(fname)
}
CorrectWinPos(hWin, wPos) {                                                                                    	;-- Fenster an eine bestimmte Position verschieben

	; prüft das ein Object übergeben wurde und gültige Werte/Schlüsselnamen enthält
		If !IsObject(wPos)
			return

	; Einstellungen prüfen
		w := GetWindowSpot(hWin)

		switch wPos.X		{

			case "-":
				wPos.X := w.X

			case "L":
				wPos.X := 0
				wPos.Y := 0

			case "R":
				wPos.X := A_ScreenWidth	- wPos.W

			case "B":
				wPos.X := 0
				wPos.Y := A_ScreenHeight	- wPos.H - 30 ; genaue Höhe der Taskbar und ob überhaupt am unteren Bildschirmrand!

			case "T":
				wPos.X := 0
				wPos.Y := 0idx
		}

		switch wPos.Y		{
			case "-":
				wPos.Y := w.Y
		}

		switch wPos.W		{
			case "F":
				wPos.W := A_ScreenWidth
		}

		switch wPos.H		{
			case "F":
				wPos.H := A_ScreenHeight
		}

	; Fenster versetzen
		If (w.X != wPos.X) || (w.Y != wPos.Y) || (w.W != wPos.W) || (w.H != wPos.H)
			SetWindowPos(hWin, wPos.X, wPos.Y, wPos.W, wPos.H)

}
TimedButtonClick(hWin, ControlName) {                                                                  	;-- SetTimer Call für einen Click auf ein Steuerelement
	; per Timer aufgerufene Funktion kommt ohne sleep aus und bremst solange das Skript nicht aus
return VerifiedClick(ControlName, hWin)
}
Clip(Text:="", Reselect:="") {                                                                                   	;-- Clip() - Send and Retrieve Text Using the Clipboard

	; by berban - updated February 18, 2019
	; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=62156

	Static BackUpClip, Stored, LastClip

	If (A_ThisLabel = A_ThisFunc) {

		If (Clipboard == LastClip)
			Clipboard := BackUpClip
		BackUpClip := LastClip := Stored := ""

	}
	Else	{

		If !Stored{
			Stored := True
			BackUpClip := ClipboardAll                         	; ClipboardAll must be on its own line
		}
		Else
			SetTimer, % A_ThisFunc, Off

	; LongCopy gauges the amount of time it takes to empty the clipboard which can predict how long the subsequent clipwait will need
		LongCopy := A_TickCount
		Clipboard := ""
		LongCopy -= A_TickCount

		If (Text = "") {
			SendInput, % "{Shift Down}c{Shift Up}"
			ClipWait, LongCopy ? 0.6 : 0.2, True
		} Else {
			Clipboard := LastClip := Text
			ClipWait, 10
			SendInput, % "{Shift Down}v{Shift Up}"
		}

		SetTimer, % A_ThisFunc, -700
		Sleep 20                                                            	; Short sleep in case Clip() is followed by more keystrokes such as {Enter}

		If (StrLen(Text) = 0)
			Return (LastClip := Clipboard)
		Else If ReSelect && ((ReSelect = True) || (StrLen(Text) < 3000))
			SendInput, % "{Shift Down}{Left " StrLen(StrReplace(Text, "`r")) "}{Shift Up}"
	}

Return

Clip:
Return Clip()
}
ClipChanged(Typ) {                                                                                               	;-- Clipboardüberwachung

	static lastclipText

	Critical

	If (Typ=0 || Typ=2) {
		Critical, Off
		return
	}

	clipText := Clipboard
	If (lastclipText=clipText) {
		Critical, Off
		return
	}

	Critical, Off

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Namen von Patienten im Clipboard erkennen und dann die Karteikarte öffnen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ; weiter wenn Funktion abgeschaltet ist
	If Addendum.Clip.WatchPatients {

		excludeIDs		:= "2"
		RegExReplace(clipText, "\D",, notInt)
		If (RegExMatch(clipText, "^\s*[\[\(]*(?<ID>\d+)[\)\]]*\s*$", Clip) && notInt > 5) {   ; enthält >5 Zeichen die keine Zahl sind

			If cPat.Exist(ClipID)  {
					MsgBox	, 0x1003	, % StrReplace(A_ScriptName, ".ahk") "- mögliche Patienten-ID erkannt"
												, % "Das Clipboard enthält die ID [" ClipID "] von:`n"
												.	 "[" ClipID "] " cPat.NAME(ClipID, true) "`n`n"
												. 	 "Möchten Sie die Karteikarte geöffnet bekommen?"
												, 8
					IfMsgBox, Cancel
						return
					IfMsgBox, Yes
						AlbisAkteOeffnen(cPat.Name(ClipID, false), ClipID)
			}

		}
		else {

			IDs           	:= Object()
			ClipDates  	:= xstring.get.AllDates(clipText)
			ClipNames 	:= xstring.get.AllNames(clipText)

			If (ClipNames.Count()+ClipDates.Count()>0) {
				sstartT := A_TickCount
				If (ClipNames.Count()>2)
					TrayTip	, % "Suche nach Namen im Clipboard"
								, % "Der Text im Clipboard könnte " ClipNames.Count()+ClipDates.Count() " Patientenamen enthalten. "
								.	"Es werden gerade die Namen/Geburtstage mit der Albis-Datenbank abgeglichen.", 8
				else
					SciTEOutput("Untersuche Clipboard nach Patientennamen.")
			}

		  ; über Namen hinzufügen
			If (ClipNames.Count()>0) {

			  ; aussortieren von Patientenaufrufen über 5
				If (ClipNames.Count() > 5) {
					tmp := ClipNames.Clone()
					i := 0
					ClipNames := Object()
					For clipname, clipdata in tmp {
						i ++
						If (i>5)
							break
						ClipNames[clipname] := clipdata
					}
				}

				cPat.dbg := false
				For clipname, clipdata in ClipNames
					If clipname && IsObject(bestmatch := cPat.StringSimilarityPlus(clipname))
						For each, bmatch in bestmatch
							If bmatch.PatID && !(bmatch.PatID ~= "^(" excludeIDs ")$") && !IsObject(IDs[bmatch.PatID])
								IDs[bmatch.PatID] := bmatch.Pat
				cPat.dbg := false
			}

		  ; über Datum hinzufügen
			If (ClipDates.Count()>0) {

				bdates := Object()                                                                                            ; enthält konvertiertes Datumszahlen
				For each, clipdate in ClipDates {
					convDate := ConvertToDBASEDate(clipdate)
					bdates[convDate] :=  !bdates.haskey(convDate) ? 1 :  bdates[convDate]+1
				}

			  ; fügt Namen bei erkanntem Datum hinzu
				For bdate, bdcount in bdates
					If bdate && IsObject(Births := cPat.PatID([{"key":"GEBURT", "value":bdate}]))
						For each, PatID in Births
							If !(PatID ~= "^(" excludeIDs ")$") && !IsObject(IDs[PatID])
								IDs[PatID] := cPat.Patients[PatID]

			}

		  ; Vorschläge machen
			If (IDs.Count() > 0) {

				SciTEOutput(IDs.Count() " Patientenname" (IDs.Count()>1 ? "n":"") " ausgehend von " ClipNames.Count()+ClipDates.Count()
									. " möglichen Namen/Geburtstagen wurden in " Round((A_TickCount-sstartT)/1000, 1) " s ermittelt." )
				lastclipText := clipText
				fn := Func("ClipboardPatienten").Bind(IDs)
				SetTimer, % fn, -100


			return
			}

		}

	}

	lastclipText := ""

return
}

; Laborabruf
LaborAbrufCheck() {                                                                                                	;-- prüft Bedingungen vor der Ausführung des Timers

	If !Addendum.Labor.AbrufTimer
		return false

	LabCallClients := Addendum.Labor.ExecuteOn
	If !InStr(LabCallClients, compname)
		return false

return true
}
LaborAbrufTimer() {                                                                                                 	;-- berechnet die Zeit bis zum nächsten eingestellten Laborabruf

	; Todo: Laborabruf an freien Tagen nur starten, wenn noch Laborbefunde ausstehen.

		global func_LaborAbruf

	; Konditionen für den Aufruf prüfen
		If !LaborAbrufCheck()
			return

	; Variablen
		func_LaborAbruf := Func("LaborAbruf")
		ANow := (A_Hour A_Min A_Sec)

	; nächsten Abrufzeitpunkt des Tages finden
		For timerNr, TimeLabCall in Addendum.Labor.Abrufzeiten {
			TimeLabCall := LTrim(TimeLabCall, "T")
			If (ANow < TimeLabCall) {
				AbrufTag	:= A_DD "." A_MM "." A_YYYY
				diffTime	:= GetSeconds(TimeLabCall) - GetSeconds(ANow)
				break
			}
		}

	; nächste Abruf fällt auf den folgenden Tag
		If !diffTime {
			For timerNr, TimeLabCall in Addendum.Labor.Abrufzeiten {
				TimeLabCall	:= LTrim(TimeLabCall, "T")
				AbrufTag  	:= A_YYYY A_MM A_DD "000000"
				AbrufTag     += 1, days
				Abruftag   	:= SubStr(Abruftag, 7, 2) "." SubStr(Abruftag, 5, 2) "." SubStr(Abruftag, 1, 4)
				diffTime    	:= (GetSeconds(235959) - GetSeconds(ANow)) + GetSeconds(TimeLabCall)
				break
			}
		}

	; Timer starten
		SetTimer, % func_LaborAbruf, % -1*(diffTime*1000)

	; Ausgabe der Startzeit und Restzeit
		Uhrzeit := SubStr(TimeLabCall, 1, 2) ":" SubStr(TimeLabCall, 3, 2)
		Restzeit := TimeFormatEx(diffTime)
		Addendum.Labor.Clock	 	:= Uhrzeit
		Addendum.Labor.Call		:= TimeLabCall
		IniWrite, % AbrufTag ", " Uhrzeit, % Addendum.Ini, % "LaborAbruf", % "naechster_Abruf"
		outtext =
		(LTrim
			nächster Laborabruf
			Uhrzeit: %Uhrzeit%
			Start in: %Restzeit%
		)
		PraxTT(outtext, "3 1")

}
LaborAbruf() {                                                                                                         	;-- startet externes Skript für den Laborabruf

	; Überprüfung des Aufrufes
		If !FileExist(Addendum.Labor.ExecuteScript) {
			FileAppend, % datestamp() "  " A_ThisFunc "():  ianovaWebClient.exe nicht vorhanden!: " Addendum.Labor.ExecuteScript "`n", % Addendum.DBPath "\logs\LaborabrufLog.txt"
			return
		}

	; Albis ausführen, damit nicht nur Daten geholt sondern auch einsortiert werden können
		If AlbisWinID() &&

	; Laborabrufskript ausführen
		Run, % q Addendum.Labor.ExecuteScript q

	; LaborabrufTimer erneut starten
		LaborAbrufTimer()
}

; Laborjournal
LaborjournalTimer(nextstart:="") {                                                                            	;-- geplanter Start des Laborjournal zu einer beliebigen Uhrzeit

	; den LaborjournalTimer an Sprechstunden oder den Laborabruf binden
	; Festlegen einer Uhrzeit für den Start des Laborjournal zu einer bestimmten Uhrzeit
	; mit "Sprechstundenbeginn" wird das Journal zu Beginn der Sprechstunde angezeigt
	;
	; letzte Änderung 14.09.2021

	; Variablen
		static func_Lbjrnal	:= Func("LaborjournalStart")

		ANow 	:= A_Hour A_Min A_Sec
		today   	:= A_YYYY . A_MM . A_DD

	; ⌚                                        Startzeitpunkt berechnen                                   	⌚
	; ⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚	⌚ 	⌚
		If (!nextstart || nextstart = "Sprechstundenbeginn") {

			daysadded := 0
			wt := DayOfWeek(today, "full", "yyyyMMdd")
			RegExMatch(Addendum.Praxis.Sprechstunde[wt], "(?<H>\d\d):(?<M>\d\d)", talk)
			talkTime := talkH talkM "00"
			If (talkTime < ANow) {
				nextday := today
				Loop {
					nextday := DateAddEx(nextday, "0y 0m 1d")
					daysadded  ++
					If vacation.DateIsHoliday(nextday)
						continue
					wt := DayOfWeek(ConvertToDBASEDate(nextday), "full", "yyyyMMdd")
					If RegExMatch(Addendum.Praxis.Sprechstunde[wt], "(?<H>\d\d):(?<M>\d\d)", talk) {
						talkTime:= talkH talkM "00"
						break
					}
				}
			}
			nextstart := talkTime
		}
		else if RegExMatch(nextstart, "(\d{6})|(\d{2}:\d{2}:\d{2})")
			nextstart := StrReplace(nextstart, ":")

	; Startzeitpunkt prüfen, liegt dieses in Urlaubstagen, wird das Laborjournalskript nicht ausgeführt
		If vacation.DateIsHoliday(today) {
			SciTEOutput("(" A_LineNumber "): Heute ist ein Urlaubstag")
			return
		}

	; Millisekunden bis zum Ausführungszeitpunkt berechnen
		If (ANow < nextstart)
			diffTime := GetSeconds(nextstart) - GetSeconds(ANow) + (daysadded * 3600 * 24)
		else
			diffTime := (GetSeconds(235959) - GetSeconds(ANow)) + GetSeconds(nextstart)

	; Func-Timer starten
		;SciTEOutput("Start d. Laborjournal in: " diffTime "s")
		SetTimer, % func_Lbjrnal, % -1*(diffTime*1000)

}
LaborjournalStart() {

	; Laborjounalskript ausführen
		Workpath :=  Addendum.Dir "\Module\Extensions"
		If FileExist(WorkPath "\Laborjournal.ahk")
			Run, % WorkPath "\Laborjournal.ahk", % WorkPath

	; LaborjournalTimer erneut starten
		LaborjournalTimer()

}

; Faxbox
Faxbox(dbg:=false) {                                                                                            	;-- FritzFaxBox: kopiert Faxdokumente von einem Fritzlaufwerk in den Befundordner

  ; der TrayTip kann erst nach anzeigen, wenn Text mit angezeigt werden soll, dann muss die Texterkennung abgeschlossen sein

	cmd 	:= Addendum.PDF.Faxbox.cmd              	; kopieren oder verschieben
	If !(cmd ~= "i)(kopieren|verschieben)")
		return

	If dbg {
		SciTEOutput("> Faxboxfunktion läuft! " A_Now )
		SciTEOutput( " -----------------------------------------------------------")
	}

	tel     	:= Addendum.PDF.Faxbox.Telefon        		; das Addendum-Telefonbuch
	today	:= SubStr(A_YYYY, 3, 2) A_MM A_DD		;

  ; alle Dateien im FritzOrdner aufnehmen
	files	:= []
	Loop, Files, % Addendum.PDF.Faxbox.Path "\*.pdf"
		If RegExMatch(A_LoopFileName, "i)(?<Date>\d{2}\.\d{2}\.\d{2,4})_(?<Time>\d+\.\d+)_Telefax\.(?<TelNr>[\w\d]+)(?<Ext>\.pdf)*", fax_)
			files.Push({"name"      	: A_LoopFileName
						, 	"telnr"         	: fax_TelNr                     								;RegExReplace(A_LoopFileName, "[\d\._]+Telefax\.(\d+)(\.pdf)*", "$1")
						,	"sendDate"   	: fax_Date
						,	"sendTime"   	: fax_Time
						, 	"created"   	: SubStr(A_LoopFileTimeCreated, 1, 8)
						, 	"attrib"       	: A_LoopFileAttrib
						, 	"size"          	: A_LoopFileSize})

  ; WerbeMails entfernen (im Telefonbuch "WerbeFax" anstatt Fax vermerken)
  ; neue Faxnummern werden registiert und der Nutzer wird nach dem Absender gefragt
	tocopy := removed := allsize := 0
	For fileindex, file in files {

      ; Nutzerabfrage                                                                                       	;{
		Absender := Addendum.PDF.Faxbox.Telefon["#" file.telnr].sdr
		AbsenderKurz := RegExReplace(Absender, "\s*\(*(Werbe)*Fax\)*")
		If (!RegExMatch(file.telnr, "i)unbekannt") && !Absenderkurz)   {
			Run, % Addendum.PDF.Faxbox.Path "\" file.name
			SetTimer, FaxboxFocus, -100
			InputBox, Absender, Fritz!FaxBox, % "Bitte geben Sie den Namen des Absender ein: " file.name "`n`n"
															.	  "der Faxfilter benötigt eine Unterscheidung für zu löschende oder weiterzuleitende Dokumente`n"
															.  	  "fügen sie am Ende des Dateinamens in Klammern. `n"
															. 	  "  (Fax) für weiterzuleitende und`n"
															.     "  (WerbeFax) oder (Werbung) dafür an.",, 800, 360

			Gui, FxBx: New, hwndhFxBx +AlwaysOnTop +ToolWindow
			Gui, FxBx: Color, cFF763C, cFF763C
			Gui, FxBx: Add, Progress, xm ym, w600 h300


		  ; neue Faxnummer speichern
			If (!ErrorLevel && StrLen(Absender)>0 && !RegExMatch(Absender, "i)^\s*\(*(Werbung|(Werbe)*Fax)\)*"))
				FileAppend,	% (Addendum.PDF.Faxbox.Telefon["#" file.telnr].sdr := Absender " (Fax)") " | " file.telnr "`n"
								,	% Addendum.PDF.Faxbox.Telefonbuch, UTF-8
		}
		;}

	  ; Werbung entfernen                                                                               	;{
		If RegExMatch(Absender, "i)WerbeFax") {
			FileDelete, % Addendum.PDF.Faxbox.Path "\" file.name
			If !ErrorLevel
				removed ++
			continue
		}
		;}

	  ; kopieren/verschieben der Faxe vom Fritzbox Ordner in den Befundordner 	;{
		newfilepath := newfilename := ""
		If RegExMatch(file.name, "[\d\._]+Telefax\.(\d+)(\.pdf)*")  {
			newfilename := RegExReplace(file.name, "([\d\.]+)_(\d+)\.(\d+)_[\w\.]+\.pdf", "$1 $2.$3") "  "
								.	Format("{:12}", Trim(file.telnr)) . "  "
								.	(Absender ? RegExReplace(Absender, "i)\(*\s*Fax\s*\)*\s*")  : Trim(file.telnr)) ".pdf"
			FileCopy, % Addendum.PDF.Faxbox.Path "\" file.name, % (newfilepath := Addendum.Befundordner "\" newfilename), 1
		 ; vor dem Löschen des Original-Faxes werden die Dateigrößen miteinander verglichen
			FileGetSize, FSize, % newfilepath
			If (FSize=0) {
				PraxTT("Eine Faxdatei konnte nicht in den Befundordner verschoben werden:`n" file.name " / " newfilename, "5 1")
				newfilepath := ""
			}
			else If (cmd~="i)verschieben" && FSize = file.size) {
				FileDelete, % Addendum.PDF.Faxbox.Path "\" file.name
				If ErrorLevel
					PraxTT("Original Faxdatei konnte nicht aus dem Faxspeicher entfernt werden:`n" file.name, "5 1")
			}
			tocopy ++
			allsize += FSize
		}
		;}

	}

  ; Debugausgabe
	If (dbg || files.Count > 0) {
		restxt := "Fax-Dokumente gesamt: " files.Count() "`n"
		restxt .= "                       gelöscht: " removed  	"`n"
		restxt .= "                    abgerufen: " tocopy     	"  (" (Round(allsize/1024/1024, 1)	>1.5 ? Round(allsize/1024/1024, 1) " Mb"
																	                : Round(allsize/1024, 1)        	>1.5 ? Round(allsize/1024, 1)        	" kb"
																	                : allsize " b") 	")`n"
		resttxt .= "> Faxboxfunktion beendet! " A_Now
		TrayTip, FaxBox, % restxt, 5
	}

return
FaxboxFocus:   ;{

  ; maximal eine Sekunde auf das Handle warten
	while !((hwnd := WinExist("Fritz!FaxBox ahk_class #32770", "Bitte geben Sie den Namen")) && A_Index<50)
		Sleep 20
	If !hwnd
		return
	ControlFocus	, Edit1, % "ahk_id " hwnd
	ControlSend	, Edit1, {Home}, % "ahk_id " hwnd
	ControlSend	, Edit1, {LShift Down}{End}{LShift Up}, % "ahk_id " hwnd
	WinSet, AlwaysOnTop, On, % "ahk_id " hwnd
return ;}
}
;}

;{13. Gesundheitsvorsorge, Hautkrebsscreening, Karteikartenfunktionen
GVUListe(modus) {

		PatID            	:= AlbisAktuellePatID()
		GVUVorhanden	:= false
		GVUZaehler   	:= 0

		If (modus=2)
			QLDatum:= AlbisZeilenDatumLesen(300)
		else
			QLDatum:= AlbisLeseProgrammDatum()						    				;abhängig vom Tagesdatum lassen sich die Daten in die GVUListe eintragen

		If (QLDatum = 0) {
			PraxTT("Es konnte kein Datum der Untersuchung`naus der Karteikarte ermittelt werden", "3 0")
			return
		}

		QListe   	:= GetQuartal(QLDatum)
		GVUFile	:= AddendumDir "\Tagesprotokolle\" QListe "-GVU.txt"

	;wird immer komplett durchsucht, um die zum Quartal zugehörigen Untersuchungen zählen zu können
		FileRead, gfile, % GVUFile
		 Loop, Parse, gfile, `n, `r
		{
				If (StrLen(A_LoopField) > 0)
					GVUZaehler ++
				If InStr(StrSplit(A_LoopField, ";").3, PatID)
					GVUVorhanden := true
		}

		If GVUVorhanden {
				PraxTT("Patient mit der ID: " PatID " wurde schon übertragen.", "2 0")
		} else {
				FileAppend, % QLDatum ";" SubStr(QListe, 1, 2) "/" SubStr(QListe, 3, 2) ";" PatID "`n", % GVUFile
				PraxTT("Die GVU Liste " SubStr(QListe, 1, 2) "/" SubStr(QListe, 3, 2) "`nenthält " (GVUZaehler + 1) " Untersuchungen.", "3 0")
				Addendum.GVUListe := GVUZaehler + 1
		}

return
}

Hausbesuchskomplexe() {		; Hausbesuchsziffern können gescrollt werden

		static HBZiffern, AktiveKlasse, AktivZuletzt, AktiveID, UpDownS := 0

		KDly	:= A_KeyDelay
		KDur	:= A_KeyDuration
		;SetKeyDelay, 0, 0

		If !IsObject(HBZiffern) {
				HBZiffern  	  := Object()
				HBZiffern.Push([01410,01411,01412])
				HBZiffern.Push([97234,97235,97236,97237,97238,97239])
				HBZiffern.Push([1,2,3,4,5,6,7,8,9])
		}

		Send, % "{Text}01410-97234-03230(x:3)"

		AktiveID:= AktivZuletzt	:= GetHex(GetFocusedControlHwnd())
		AktiveKlasse              	:= GetClassName(AktivZuletzt)
		ControlGetPos	, cx, cy,,,, % "ahk_id " AktiveID

		ToolTip, % "Pfeiltaste 'Hoch' oder 'Runter' zum ändern der Ziffern", % cx - 40, % cy - 40, 11

		while (!GetKeyState("TAB") || !InStr(GetHex(GetFocusedControlHwnd()), AktiveID)) {

				if GetKeyState("Up")
						UpDownS := 1
				else if GetKeyState("Down")
						UpDownS := -1

				If (UpDownS= 1) || (UpDownS = -1) {

					ControlGet, row, Currentline	,			,, % "ahk_id " AktivZuletzt
					ControlGet, col	, CurrentCol	,			,, % "ahk_id " AktivZuletzt
					ControlGet, line, Line			, % row	,, % "ahk_id " AktivZuletzt

					hbStr			:= RegExMatch(line, "014\d+-972\d+-03230\(x:\d\)", hbregx)
					col			:= col - hbStr
					word 		:= (col <= 6) ? 1 : (col >= 7 && col <= 12) ? 2 : (col >= 13) ? 3 : 0

					If word in 1,2
					{
						For Index, element in HBZiffern[word]
							If InStr(line, element) {
								SubIndex:= A_Index
								break
							}

						newSIdx	:= subIndex + UpDownS
						Index 		:= newSIdx > HBZiffern[word].MaxIndex() ? 1 : ( ( newSIdx < 1 ) ? HBZiffern[word].MaxIndex() : (newSIdx) )
						newline		:= StrReplace(line, HBZiffern[word, SubIndex], HBZiffern[word, Index] )

						ToolTip, % "Pfeiltaste 'Hoch' oder 'Runter' zum ändern der Ziffern. " hbStr " , " HBZiffern[word, SubIndex], % cx - 40, % cy - 40, 11

					}
					else {

							RegExMatch(line, "(?<=03230\(x:)\d", faktor)
							faktor	:= ( faktor + UpDownS > 9 ) ? 1 : ( faktor + UpDownS < 1 ) ? 9 : (faktor + UpDownS)
							newline := RegExReplace(line, "(?<=03230\(x:)\d", faktor)

					}

					ControlSetText,, % Trim(newline, "`n`r"), % "ahk_id " AktivZuletzt
					Sleep, 50
					ControlSend	 ,, % "{Right " col "}"		  , % "ahk_id " AktivZuletzt

					line:= newline:= col:= hbStr:= word:= SubIndex:= lko:= UpDownS:= ""
				}

				Sleep, 60
				UpDownS := 0
				;~ ControlGetFocus, Aktiv, % "ahk_id " AktiveID
				;~ If !InStr(AktiveKlasse, Aktiv)
							;~ break

				If !InStr(GetHex(GetFocusedControlHwnd()), AktiveID)
					break
		}

		ToolTip,,,, 11
; SetKeyDelay, % KDly, % KDur

return
}

;}

;{14. EventHooks

;                         ---------------------------------------------------------------------------------------------------------------------------------
;                       	 AUF DIESEN FUNKTIONEN BERUHEN DIE WICHTIGSTEN VORGENOMMENEN AUTOMATISIERUNGEN
;                         ---------------------------------------------------------------------------------------------------------------------------------


; ----------------------- Hook-Initialisierung
InitializeWinEventHooks() {                                                                 	; Robotic Process Automation (RPA) fängt hiermit an!

	/* https://docs.microsoft.com/en-us/windows/desktop/winauto/event-constants

		; EVENT_OBJECT_CREATE                 	:= 0x8000
		; EVENT_OBJECT_DESTROY              	:= 0x8001
		; EVENT_OBJECT_SHOW                   	:= 0x8002
		; EVENT_OBJECT_HIDE                     	:= 0x8003
		; EVENT_OBJECT_REORDER              	:= 0x8004	; 	A container object has added,removed,or reordered its children.
																							(header control, list-view, toolbar control, and window object - !z-order!)
		; EVENT_OBJECT_FOCUS                 	:= 0x8005	; 	An object has received the keyboard focus. elements: listview,menubar, popup menu,switch window,
																							tabcontrol, treeview, and window object.
		; EVENT_OBJECT_INVOKED              	:= 0x8013	; 	An object has been invoked; for example, the user has clicked a button.
		; EVENT_OBJECT_NAMECHANGE     	:= 0x800C   ; 	An object's name has changed - like a title bar
		; EVENT_OBJECT_VALUECHANGE     	:= 0x800E	; 	An object's Value property has changed.  edit, header, hot key, progress bar, slider, and up-down, scrollbar

		; EVENT_SYSTEM_CAPTUREEND        	:= 0x0009	; 	A window has lost mouse capture.
		; EVENT_SYSTEM_CAPTURESTART      	:= 0x0008	; 	A window has received mouse capture. This event is sent by the system, never by servers.
		; EVENT_SYSTEM_DIALOGEND          	:= 0x0011	; 	A dialog box has been closed.
		; EVENT_SYSTEM_DIALOGSTART        	:= 0x0010	; 	A dialog box has been displayed.
		; EVENT_SYSTEM_FOREGROUND      	:= 0x0003	; 	The foreground window has changed. The system sends this event even if the foreground window has
																							changed to another window in the same

		; EVENT_SYSTEM_MENUSTART           	:= 0x0004
		; EVENT_SYSTEM_MENUPOPUPSTART 	:= 0x0006	; context menu was opened
		; EVENT_SYSTEM_MENUPOPUPEND  	:= 0x0007	; context menu was closed

	 */

		EVENT_SKIPOWNTHREAD         	:= 0x0001
		EVENT_SKIPOWNPROCESS       	:= 0x0002
		EVENT_OUTOFCONTEXT          	:= 0x0000

	; creating the hooks - allgemein zur Erfassung von neuen Fenstern/Dialogen in diversen Programmen
		HookProcAdr         	:= RegisterCallback("WinEventProc"    	, "F")
		hWinEventHook     	:= SetWinEventHook( 0x0003, 0x0003, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x0010, 0x0010, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8000, 0x8000, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8004, 0x8004, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x800C, 0x800C, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		Addendum.Hooks[1]	:= {"HPA": HookProcAdr, "hWinEventHook": hWinEventHook}

	; der ShellHook benötigt ein dummy Gui an das die Ereignisse (Events) geschickt werden
		Gui, AMsg: New	, +HWNDhMsgGui +ToolWindow
		Gui, AMsg: Add	, Edit, xm ym w100 h100 HWNDhEditMsgGui
		Gui, AMsg: Show	, AutoSize NA Hide, % "Addendum Message Gui"
		Addendum.MsgGui.hMsgGui     	:= hMsgGui
		Addendum.MsgGui.hEditMsgGui	:= hEditMsgGui

	; Shellhook wird gestartet
		DllCall("RegisterShellHookWindow", UInt, hMsgGui)
		MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(MsgNum, "ShellHookProc")

	; startet weitere Hooks die nur auf Ereignisse des Albisprozess reagieren
		AlbisStartHooks()

	; Eingabe Focus Hook für andere Programme (Autovervollständigung)
		;Addendum.Hooks[6]         	:= GetActiveFocusHook()

return
}

AlbisStartHooks() {                                                                        		; startet Hooks die Ereignisse des Albisprozeß registrierem

	global AlbisID_old := 0

	If (Addendum.AlbisPID := AlbisPID()) {

		Addendum.AlbisWinID	:= AlbisWinID()
		AlbisID_old               	:= AlbisWinID() ; für ShellhookProc

		; Addendum Toolbar neu starten oder Toolbar wieder anzeigen lassen
			If    	  ( Addendum.ToolbarThread && !IsObject(Addendum.Thread["Toolbar"]) ) {
				If (StrLen(Addendum.Threads.Toolbar) = 0)            ; Thread-Code muss zunächst geladen werden - Funktion aus Addendum_Ini.ahk
					admThreads()
				Addendum.Thread["Toolbar"] := AHKThread(Addendum.Threads.Toolbar)
			}
			else if ( Addendum.ToolbarThread && IsObject(Addendum.Thread["Toolbar"]) )
				Addendum.Thread["Toolbar"].ahkFunction("ToolbarShowHide", "Show")

		; startet den Hook bei Wechsel des Eingabefocus (eingabefähige Steuerelemente)
			Addendum.Hooks[2]  	:= AlbisFocusHook(AlbisPID)

		; feste Größe des Albisfenster wenn gewünscht
			If Addendum.AlbisLocationChange
				Addendum.Hooks[3]	:= AlbisLocationChangeHook(AlbisPID)

		; Test Hook - PopUpMenu abfangen und eigene Einträge integrieren
			If Addendum.PopUpMenu
				Addendum.Hooks[4]	:= AlbisPopUpMenuHook(AlbisPID)

	}

}

AlbisStopHooks() {                                                                           	; beendet ALBIS Hooks, wenn kein Albis geschlossen wurde

		local idx

		;Gui, AMsg: 	Show, Hide
		Gui, adm: 	Destroy
		Gui, adm2: 	Destroy

	; Toolbar-Thread nicht beenden, Toolbar nur Ausblenden
		If Addendum.ToolbarThread && IsObject(Addendum.Thread["Toolbar"]) {
			Addendum.Thread["Toolbar"].ahkFunction("ToolbarShowHide", "Hide")
			PraxTT("Albis is closed.`nAddendum Toolbar: " (Addendum.Thread[1].ahkReady() ? "is still running":"is terminated"), "2 2")
		}

	; ein Teil der Hook's beenden
		Loop % (Addendum.Hooks.MaxIndex() -1) { 	; Hook[1] wird nicht beendet da dieser auch andere Programme untersucht
			idx := A_Index + 1
			if idx in 1,5
				continue
			If IsObject(Addendum.Hooks[idx]) {
				UnhookWinEvent(Addendum.Hooks[idx].hEvH, Addendum.Hooks[idx].HPA)
				Addendum.Hooks[idx] := ""
			}
	}


}


; -----------------------  Hooks auf Albis
AlbisPopUpMenuHook(AlbisPID) {                                                              	; Hook zum Abfangen eines Rechtsklick Menu

	HookProcAdr   	:= RegisterCallback("AlbisPopUpMenuProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x0006, 0x0007, 0, HookProcAdr, AlbisPID, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisFocusHook(AlbisPID) {                                                                   	; Hook für EVENT_OBJECT_FOCUS (nur Albis)

	; benötigt um Änderung des Eingabefocus innerhalb von Albis zu erkennen
	; z.B. für eine leider noch funktionierende Autovervollständigungsfunktion

	HookProcAdr   	:= RegisterCallback("AlbisFocusEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, AlbisPID, 0, 0x3)
	hWinEventHook	:= SetWinEventHook( 0x800C, 0x800C, 0, HookProcAdr, AlbisPID, 0, 0x3)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisLocationChangeHook(AlbisPID) {                                                          	; Hook für EVENT_OBJECT_LOCATIONCHANGE (nur Albis)

	static EVENT_OBJECT_LOCATIONCHANGE	:= 0x800B
	static EVENT_OBJECT_REORDER                	:= 0x8004
	HookProcAdr	   	:= RegisterCallback("AlbisLocationChangeEventProc", "F")
	hWinEventHook	:= SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_LOCATIONCHANGE, 0, HookProcAdr, AlbisPID, 0, 0x0003)
	hWinEventHook	:= SetWinEventHook(EVENT_OBJECT_REORDER, EVENT_OBJECT_REORDER, 0, HookProcAdr, AlbisPID, 0, 0x0003)

	; Albisfenster ist minimiert oder nicht - flag für den EventProc erstellen
	a := GetWindowSpot( Addendum.AlbisWinID := AlbisWinID() )
	If (a.X <= -20) && (a.Y <= -20)
		Addendum.AlbisWasMinimizedLast := true
	else
		Addendum.AlbisWasMinimizedLast := false

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

GetActiveFocusHook() {                                                                       	; Hook für EVENT_OBJECT_FOCUS (alle anderen Prozesse)

	HookProcAdr   	:= RegisterCallback("ActiveFocusEventProc", "F")
	;~ hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, 0, 0, 0x0003)   	; keyboard focus
	hWinEventHook	:= SetWinEventHook( 0x0000, 0x0005, 0, HookProcAdr, 0, 0, 0x0003)    	; mouse focus

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}


; ----------------------- Verarbeitung abgefangener Systemmeldungen (EventProcs)
WinEventProc(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {               	; Hook für bestimmte Fenster/Prozesse

	/*  Beschreibung

		die Geschwindigkeit von Autohotkey reicht nicht aus um wirklich auch alle Events abfangen zu können. Daher sind die Zahl der abzufangenden
		Eventnummern möglichst klein zu halten. Um auf jedes abgefangene Event reagieren zu können, verwende ich eine Bearbeitungsstapel (EHStack).
		Dies ist ein Objekt, welches immer mit den neuen Event-Daten befüllt wird. Der Bearbeitungsprozeß (Label EventHook_WinHandler) signalisiert
		durch Setzen eines globalen Flags (EHWHState), wenn der komplette Stapel abgearbeitet wurde, damit die Bearbeitung auch wieder gestartet wird.
		Der EventHook_WinHandler arbeitet also so lange bis kein Element mehr zu bearbeiten ist.
		Es kommt hier unbedingt auf schnell zu ermittelnde Eventdaten an, um nur möglichst selten ein Event zu verpassen. Der WinEventProc Code ist im
		Laufe von Jahren von mir optimiert worden und ist ausreichend schnell jetzt.

	*/

	; Filter für Fensterklassen
		static class_filter	:= [ "OptoAppClass", "#32770", "Teamviewer", "WindowsForms10.Window", "Qt5QWindowIcon", "Qt5152QWindowIcon"
	                            	, "AutohotkeyGui", "SciTEWindow", "classFoxitReader", "IEFrame", "SUMATRA_PDF_FRAME"]

	; Filter für App-Namen (# Filter wird nicht verwendet)
		static app_filter 		:= [	"albis"	,  "albiscs"     	, "wkflsr32"	,  "AOWDump2"	, "Teamviewer"	, "infoBoxWebClient"	, "ifap"          	, "Autohotkey"
									,  	 	"SciTE"	,  "FoxitReader"	, "word"     	,  "mDicom"    		, "iexplore"    	,  "EXPLORER"        	,  "OUTLOOK"	,  "InternalAHK"]

		static classes := Object(), thistick := 1

		Critical

	; Ausschluß
		If !hwnd || (Addendum.LastHookHwnd = (EHookHwnd := Format("0x{:x}", hwnd)))
			return 0

	; Fensterklassen Filter
		If !(cClass := WinGetClass(hwnd))
			return 0

		known := false
		For fnr, filterclass in class_filter
			If (cClass = filterclass) {
				For snr, item in EHStack
					If (item.Hwnd = EHookHwnd)
						known := true, break

				If !known {
					event := Format("0x{:x}", event)
					eventNr := "#" event
					If (StrLen((HTitle := WinGetTitle(EHookHwnd)) . (HText := WinGetText(EHookHwnd))) > 0) {
						EHStack.InsertAt(1, {"Hwnd"	: EHookHwnd
													, "Event"	: event
													, "Title"   	: HTitle
													, "Text"   	: HText := StrReplace(HText, "`r`n", " ")
													, "Class"	: cClass})
						If (EHStack.1.Text ~= "i)ADT\-Abrechnungsdatei" && EHStack.1.Class ~= "OptoAppClass")
							SciTEOutput(SubStr("000" EHStack.Count(), -2) ": event: " GetHex(event)  ", hwnd: " EHookHwnd "; Txt: " EHStack.1.Text)
						If !EHWHState
							SetTimer, EventHook_WinHandler, -1
					}
				}
			}


	; Callback für PDF Funktionabilität
		If (StrLen(Addendum.FuncCallback) > 0) {
			For idx, viewerclass in PDFViewer
				If InStr(cClass, viewerclass) {
					If InStr(Addendum.FuncCallback, "|") {
						callbackParam := StrSplit(Addendum.FuncCallback, "|")
						If IsFunc(callbackParam[1])
							func_Call 	:= Func(callbackParam[1]).Bind(callbackParam[2], cClass, SHookHwnd)
					} else {
						If IsFunc(fnName := Addendum.FuncCallback)
							func_Call 	:= Func(fnName).Bind(cClass, SHookHwnd)
					}
					SetTimer, % func_Call, -200
					break
				}
		}

	; Order & Entry Programm läßt sich nicht filtern
		WinGet, EHproc, ProcessName, % "ahk_id " EHookHwnd
		If Instr(EHProc, "infoBoxWebClient") {
			EHStack.InsertAt(1	, {"Hwnd"	: EHookHwnd
										, 	"Event"	: event
										, 	"Title"	: HTitle
										, 	"Text"	: HText
										, 	"Class"	:cClass})
			If !EHWHState
				SetTimer, EventHook_WinHandler, -1
		}

return 0
}

AlbisPopUpMenuProc(hHook, event, hwnd) {                                                    	; weitere Menupunkte für das Albiskontextmenu

	; alle anderen Funktionen für Aufrufe finden sich in %AddendumDir%\Module\Addendum\include\Addendum_PopUpMenu.ahk

	static pm

	If (event = 6)	                                                                             	; Kontextmenu wird aufgerufen
		pm := Addendum_PopUpMenu(hwnd, GetMenuHandle(hwnd))
	else if (event = 7) && IsObject(pm)                                                	; es wurde ein Menupunkt per Mausklick ausgewählt
		Addendum_PopUpMenuItem(pm)

}

AlbisFocusEventProc(hHook, event, hwnd) {                                                    	; Handler EVENT_OJECT_FOCUS Nachrichten Albis

	If !GetDec(hwnd)
		return 0

	Critical

  ; Karteikarten-/Patientenprotokoll
	fn_PatChange := Func("CurrPatientChange").Bind(AlbisGetActiveWinTitle())
	SetTimer, % fn_PatChange, -0

  ; Auswahlbox wird bei Focusverlust geschlossen
	If Addendum.AWBCheckFocus && WinActive("ahk_class OptoAppClass") {
		If (Addendum.AWBCheckFocus <> GetFocusedControl()) {   ; schnellste Variante bisher
			Addendum.AWBCheckFocus := ""
			If WinExist("Auswahlbox ahk_class AutoHotkeyGUI")
				Gui, AWB: Destroy
		}
	}

  ; Abrechnungsassistenten
	If WinActive("ahk_class OptoAppClass") && (habrAssist := WinExist("Abrechnungsassistent ahk_class AutoHotkeyGUI"))
		If GetNextWindow(habrAssist, 3)
			WinActivate, % "ahk_id " habrAssist

return 0
}

AlbisLocationChangeEventProc(hHook, event, hwnd) {                                           	; Handler EVENT_OBJECT_LOCATIONCHANGE Albis

	; nur für Monitore mit einer Auflösung > 2k
	; last change: 27.01.2021

		global adm
		static SizeInit := false
		static AlbisX, AlbisY, AlbisW, AlbisH

		Addendum.AlbisWinID := AlbisWinID()
		monIndex	:= GetMonitorIndexFromWindow(Addendum.AlbisWinID)
		AlbisMon 	:= ScreenDims(monIndex)
		If (AlbisMon.W <= 1920)
			return 0

	; AlbisX, AlbisY, AlbisW, AlbisH werden erstellt
		If !SizeInit {

			SizeInit := true
			ShellTray := GetWindowSpot(WinExist("ahk_class Shell_TrayWnd")) ; Größe der Taskbar

			RegExMatch(Addendum.AlbisGroesse, "i)y\s*(?<Y>\d+)"        	, A)
			AlbisY := AY

			RegExMatch(Addendum.AlbisGroesse, "i)w\s*(?<W>\d+)"      	, A)
			AlbisW := AW

			RegExMatch(Addendum.AlbisGroesse, "i)h\s*(?<H>\d+|Max)"	, A)
			AlbisH := (AH = "Max") ? (AlbisMon.H - ShellTray.H + 12) : AH

			RegExMatch(Addendum.AlbisGroesse, "i)x\s*(?<X>\d+|R|L)" 	, A)
			If 	(AX = "R")        	; rechter Bildschirmrand
				AlbisX := AlbisMon.W - AlbisW
			else if (AX = "L") 	; linker Bildschirmrand
				AlbisX := 0
			else
				AlbisX := AX
		}

		If (hwnd = 0) || (GetHex(hwnd) != Addendum.AlbisWinID)
			return 0

		Addendum.MonSize  	:= AlbisMon.W > 1920 ? 2 : 1
		Addendum.Resolution	:= AlbisMon.W > 1920 ? "4k" : "2k"

	; wenn Albis minimiert wird dann wird die AddendumToolbar versteckt (diese wird sonst auf dem Desktop angezeigt)
		APos := GetWindowSpot(Addendum.AlbisWinID)                    	; akutelle Albisfenstergröße wird unter a abgelegt
		If (APos.X <= -20) && (APos.Y <= -20) && !Addendum.AlbisWasMinimizedLast {
			Addendum.AlbisWasMinimizedLast := true
			If IsObject(Addendum.Thread[1])
				Addendum.Thread[1].ahkPostFunction("ToolbarShowHide" , "Hide")
			return 0
		} else If (APos.X > -20) && (APos.Y > -20) {
			Addendum.AlbisWasMinimizedLast := false
			If IsObject(Addendum.Thread[1])
				Addendum.Thread[1].ahkPostFunction("ToolbarShowHide" , "") ; Toolbar wieder anzeigen
		}

	; Albisfenstergröße herstellen, Albis kann aber weiterhin auf dem Bildschirm verschoben werden
		If ((APos.W <> AlbisW) || (APos.H <> AlbisH)) {   ; ein minimiertes Albis hat eine x, y Position von -32000

			X := StrLen(AlbisX) > 0 ? (AlbisX + 10)	: APos.X
			Y := StrLen(AlbisY) > 0 ? (AlbisY -  6)	: APos.Y

			If IsRemoteSession() {

				W	:= Floor(AlbisW * (factor := A_ScreenDPI / 96))
				H	:= Floor(AlbisH *  factor)

			}
			else
				W := AlbisW, H := AlbisH

			SetWindowPos(Addendum.AlbisWinID, X, Y, W, H)

		}

	; Addendum Infofenster neu zeichnen
		If WinExist("AddendumGui ahk_class AutoHotkeyGUI") {
			WinSet, Style, 0x50020000 , % "ahk_id " hadm
			WinSet, Top                   	,, % "ahk_id " hadm
			RedrawWindow()	; zeichnet das Infofenster neu
		}

return 0
}

ActiveFocusEventProc(hHook, event, hwnd) {                                                   	; Handler EVENT_OJECT_FOCUS Nachrichten (alle anderen)

	static excludeFilter := ["OptoAppClass"]
	local idx

	If !hwnd
		return 0

	Critical
	cFocus	:= GetClassName((hwnd := Format("0x{:x}", hwnd)))
	SciTEOutput(GetHex(event) ": " cFocus)
	;~ If !(classList:= GetParenClasstList(hwnd))
		;~ return

	For idx, exclass in excludeFilter
		If InStr(classList, exclass)
			return 0


}

ShellHookProc(lParam, wParam) {                                                              	; Startet oder entfernt den Albis WinEventHook

		; WM_EXITSIZEMOVE := 0x232
		; Sammlung ignorierter Fensterklassen zur Funktionsbeschleunigung, da Autohotkey als Skriptsprache zu langsam ist,
		; um den Message-Queue schnell genug abarbeiten zu können
		static excludeFilter	:= [ "Windows.UI"              , "NVOpen"                   , "QTrayIcon"                , "SunAwt"
														,  "qtopen"                  , "IME"                      , "OfficePowerManagerWindow" , "ApplicationFrameWindow"
		                        ,  "Scite"                   , "ThunderRT"                , "QT5Q"                     , "Shell Embedding"
						  			       	,  "TForm"                   , "TApplication"             , "DDEM"                     , "COMTASKSWINDOWCLASS"
						  			       	,  "WICOMCLASS"              , "OleDdeWndClass"           , "GlassWndClass-"           , "Shell_TrayWnd"
						  			       	,  "DAXParkingWindow"        , "EdgeUiInputTopWndClass"   , "DummyDWMListenerWindow"   , "Progman"
						  			       	,  "ThumbnailDeviceHelperWnd", "tooltips_class32"         , "MsiDialogCloseClass"      , "MsiHiddenWindow"
							  		       	,  "HwndWrapper"             , ".NET-BroadcastEventWindow", "CtrlNotifySink"           , "OperationStatusWindow"
										        ,  "ThorConnWndClass"        , "Shell Preview Extension"  , "TForm_ipcMain"            , "GDI+"
														,  "Ghost"]
		static includeFilter 	:= [ "OptoAppClass"            , "#32770"                   , "AutohotkeyGui"            , "OpusApp" ]
		static PDFViewer 	   	:= [ "classFoxitReader"        , "SUMATRA_PDF_FRAME" ]
		static func_winpos, lastlparam, lastclass

		global AlbisID_old, hadm, hadmRN, oPV ; hadm = hwnd des Infofenster

		Critical	;, 50

	; return on empty callback parameters                                                                  	;{
		If !wParam
			return 0

		If !(wclass := WinGetClass(wparam))
			wclass := GetClassName(wParam)
		If !(title := WinGetTitle(wparam))
			WinGetTitle, title, % "ahk_id " wParam

		If (StrLen(StrReplace(wclass . title, A_Space)) = 0)
			return 0


		; filtert hier die Fensterklassen raus welche im Array excludeFilter vorhanden sind
		For i, exclude in excludeFilter
			If InStr(wclass, exclude)
				return 0
	;}

	; SHELL_WINDOWCREATED,WINDOWACTIVATED,REDRAW,RUDEAPPACTIVATED	;{
		SHookHwnd := Format("0x{:X}", wparam)
		If RegExMatch(lParam, "^(1|4|6|32772)$")                         	{                                               	; ALBIS wurde gestartet

		  ; Albisereignisse haben Vorrang
			If (AlbisWinID() != AlbisID_old) {
				AlbisID_old := AlbisWinID()
				AlbisStartHooks()                                                                            	; startet Hooks
				PraxTT("Ein neuer Albisprozeß wurde erkannt.`nDie Fensterhooks wurden gesetzt!", "2 0")
			}

		  ; InfoFenster
			If Addendum.AddendumGui {
				ZeigeInfofenster := RegExMatch(AlbisGetActiveWindowType(), "i)(iWin|Laborblatt|Biometriedaten)") ? true : false
				If (AlbisWinID()=SHookHwnd && ZeigeInfofenster) {
					If (!admGui_Exist() || Addendum.iWin.lastPatID != AlbisAktuellePatID())
						AddendumGui()
				}
				else if (admGui_Exist() && !ZeigeInfofenster)
					admGui_Destroy()
			}

		  ; Apps: Fensterposition
		  ; andere Programme (MS Word, Ifap, FoxitReader) - Einstellungen sollten in der Addendum.ini stehen
			for classnn, appName in Addendum.Windows.Proc {
			  if (appname ~= "i)^\s*Albis")
           continue
				else if InStr(classnn, wclass)                                 ; Fensterklassse ist bekannt und das Fenster soll in dieser Auflösung positioniert werden
					If (Addendum.Windows[appname].AutoPos & Addendum.MonSize)  {
						fn_winpos := Func("CorrectWinPos").Bind(SHookHwnd, Addendum.Windows[appname][Addendum.Resolution])
						SetTimer, % fn_winpos, -400    ; Zeit fürs Neuzeichnen
						break
					}

			}

		}
		else if (lParam = 2) && !WinExist("ahk_class OptoAppClass")	{	                                               	; ALBIS wurde beendet
			AlbisID_old := 0
			AlbisStopHooks()
			PraxTT("Albis wurde beendet. Hooks wurden entfernt.", "1 0")
			If (Addendum.AddendumGui && admGui_Exist())
				admGui_Destroy()
		}
	;}

	; RenameGui synchron mit SumatraPDF minimieren oder maximieren   					;{
		If (Addendum.Flag.RenameGui && lParam = 0x5 && SHookHwnd = oPV.Hwnd) {
				If (wParam=1)       	; is Minimized
					WinMinimize, % "ahk_id " hadmRN   ; Sumatra PDF auch minimieren
				else if (wParam=0)	; maximiert
					WinMaximize, % "ahk_id " hadmRN   ; Sumatra PDF ebenso maximieren

		}

	;}

	; --------------------------- PopUpMenuCallback -----------------------------                	;{
	; für Addendum_PopUpMenu - ruft die als String in der Variable .PopUpMenuCallback Funktion auf und übergibt zusätzliche Parameter
		If Addendum.PopUpMenuCallback {

			For each, viewerclass in PDFViewer
				If InStr(wclass, viewerclass) {

					If InStr(Addendum.PopUpMenuCallback, "|") {
						callbackParam := StrSplit(Addendum.PopUpMenuCallback, "|")
						If IsFunc(callbackParam.1)
							func_Call 	:= Func(callbackParam.1).Bind(callbackParam.2, wclass, SHookHwnd)
					}
					else {
						fnName := Addendum.PopUpMenuCallback
						If IsFunc(fnName)
							func_Call := Func(fnName).Bind(wclass, SHookHwnd)
					}

					Addendum.PopUpMenuCallback := ""
					SetTimer, % func_Call, -300
					break

				}

	}
	;}

return
}

WaitEH_WinHandler:                                                                           	;{

	If (EHStack.Count() > 0 && !EHWHState) {
		SetTimer, WaitEH_WinHandler, Off
		gosub EventHook_WinHandler
	} else if (EHStack.Count() = 0)
		SetTimer, WaitEH_WinHandler, Off

return ;}


; ----------------------- Automatisierungsroutinen / Fensterhandler
;
; Vorsicht mit Änderungen des Skriptcodes vor und nach dem großen If-Block ! Mehr als 3 Jahre ständiger Optimierung und Fehlersuche stecken in diesen Zeilen!
; Der Code ist jetzt zuverlässig und ausreichend schnell in der Erkennung und Abarbeitung der Hook-Events (Fenster, Applikationen).
;
EventHook_WinHandler:                                                                       	;{ Eventhookhandler - Popupblocker/Fensterhandler

	StackEntry:
	Critical 20
	EHWHState :=true, thisWin := EHStack.Pop()
	EHWT := thisWin.title, EHWText := thisWin.Text, EHWClass := thisWin.Class
	Addendum.LastHookHwnd := hHookedWin := thisWin.Hwnd

	If (StrLen(EHWT . EHWText) = 0) || InStr(EHWText, "SkinLoader")
		If (EHStack.Count() = 0) {
			EHWHState := Addendum.LastHookHwnd := false
			return
		} else
			goto StackEntry

	WinGet, EHproc1, ProcessName, % "ahk_id " hHookedWin
	If (EHWT && !EHproc1) {
		WinGet, EHproc1, ProcessName, % "ahk_id " GetParent(hHookedWin)
		If !EHproc1 && !EHStack.Count() {
			EHWHState := Addendum.LastHookHwnd := false
			return
		} else if !EHproc1 && (EHStack.Count() > 0)
			goto StackEntry
	}

	If      InStr(EHproc1, "albis")                                                                                         	{        	; ALBIS

		If   	  InStr(EHWText	, "ist keine Ausnahmeindikation")                                                                               	{
		  ; Fenster wird geschlossen
			while (A_Index < 10 && WinExist("ALBIS ahk_class #32770", "ist keine Ausnahmeindikation"))  {        ; #code changed
				If (A_Index > 1)
					Sleep 50
				VerifiedClick("Button2", hHookedWin)
			}
			Addendum.GNRChanged 	:= true
		}
		else if Instr(EHWText	, "ADT-Abrechnungsdatei")                                                                                       	{	; SetWinEventHook pausieren
			; pausieren oder weiter wenn dieses Fenster offen ist
			; Addendum reagiert nicht mehr
		}
		else If InStr(EHWText	, "Patient aus dem Wartezimmer entfernen")                                                                      	{
		  ; Wartezimmer entfernen nicht nachfragen
			If Addendum.AutoDelete.WZ && RegExMatch(AlbisWZTabAktuell(), "i)(" Addendum.AutoDelete.WZFilter ")")
				VerifiedClick("Ja", hHookedWin)
		}
		else If InStr(EHWText	, "Möchten Sie diesen Eintrag wirklich")       	                                                                 	{
			pTitle := StrSplit(WinGetTitle(GetParent(hHookedWin)), A_Space).1
			pressOK := 	(pTitle~="i)Wartezimmer"          	&& Addendum.AutoDelete.WZ)          	? true
							:	(pTitle~="i)Dauermedikamente" 	&& Addendum.AutoDelete.Dauermed) 	? true
							:	(pTitle~="i)Dauerdiagnosen"      	&& Addendum.AutoDelete.Dauerdia) 	? true : false
			If pressOk
				If !VerifiedClick("Ja", hHookedWin,,,1)
					  VerifiedClick("Ja", "ALBIS ahk_class #32770", "Möchten Sie diesen Eintrag wirklich",,1)
		}
		else If InStr(EHWText	, "ALBIS wartet auf Rückgabedatei")                                                                             	{
		  ; Laborhinweis schliessen
			BlockInput, On
			VerifiedClick("Button1", hHookedWin)
			AlbisActivate(1)
			SendInput, {Tab}
			WinActivate, % "ahk_class #32770", % "ALBIS wartet auf Rückgabedatei"
			BlockInput, Off
		}
		else If InStr(EHWText	, "Patient hat in diesem")                              	&& Addendum.noChippie                                  	{
		  ; Dialog wird sofort oder mit Verzögerung geschlossen
			SetTimer(Func("TimedButtonClick").Bind(hHookedWin, "Button1"), -1*(Addendum.noChippie=-1 ? 1:Addendum.noChippie*1000))
		}
		else if Instr(EHWText	, "Der Parameter ist bereits in dieser Gruppe")                                                                 	{
		  ; Fenster wird geschlossen
			VerifiedClick("Button1", hHookedWin)
		}
		else if Instr(EHWText	, "Behandlungszeitraum")                                                                                        	{
		  ; Behandlungszeitraum überschritten, neue Rechnung erstellen
			VerifiedClick("Button2", "ALBIS", "Behandlungszeitraum")
		}
		else if InStr(EHWText	, "folgende Diagnosen übernommen")                                                                               	{
		  ; Fenster wird geschlossen
			VerifiedClick("Button1", hHookedWin)
		}
		else if InStr(EHWT   	, "ALBIS - Login")                                       	&& Addendum.AutoLogin=1                                	{
		  ; Client spezifisches automatisches Einloggen in Albis
			AlbisAutoLogin(10)
		}
		else If InStr(EHWText	, "Sie haben eine Besuchsziffer ohne Wege")                                                                      	{
		  ; Fenster wird geschlossen
			VerifiedClick("Nein", "ALBIS ahk_class #32770", "Sie haben eine Besuchsziffer ohne Wege")
		}
		else If InStr(EHWText	, "Wollen Sie wirklich die GNR ändern")                                                                       		{
		  ; es wird automatisch mit "Ja" bestätigt
			VerifiedClick("Button1", "ALBIS ahk_class #32770", "Wollen Sie wirklich")
		}
		else If InStr(EHWText	, "Der Patient wurde bereits als verstorben")             && (Addendum.Flags.ImportFromJrnl || Addendum.Flags.ImportFromPat) 	{
		  ; es wird automatisch mit "Ja" bestätigt bei automatisiertem Befundimport
			VerifiedClick("Button1", hHookedWin)
		}
		else If InStr(EHWText	, "Folgender Pfad existiert nicht")                                                                              	{
		  ; Fenster wird geschlossen
			VerifiedClick("Button1", hHookedWin)
		}
		; Hinweise ------------------------------------------------------------------------------------------------------------------------------
		else If InStr(EHWText	, "Fehler beim Aufruf dppivassist")                                                                              	{
		  ; Fenster wird geschlossen
			VerifiedClick("Button1", hHookedWin)
		}
		else if Instr(EHWT   	, "Herzlichen Glückwunsch")                                                                                     	{
          ; Fenster "Herzlichen Glückwunsch" nur unter bestimmten Bedingungen schliessen
			If (Addendum.ImportFromJrnl || Addendum.ImportFromPat || WinExist("Gesundheitsvorsorgeliste"))
				VerifiedClick("Button1", hHookedWin)
		}
		else If InStr(EHWText	, "Datum darf nicht vordatiert werden")                                                                          	{
		  ; Datum wird verändert, Daten eingetragen, Datum wird zurückgesetzt
			AlbisVordatierer()
		}
		else If InStr(EHWT  	, "Daten von")                                           	&& Addendum.GNRChanged                                 	{
		  ; schließt nach Änderung der GNR das Fenster "Daten von"
			while (A_Index < 50 && !WinExist("ALBIS ahk_class #32770", "ist keine Ausnahmeindikation"))
				Sleep 10
			VerifiedClick("Button30", "Daten von ahk_class #32770")
			Addendum.GNRChanged 	:= false
		}
  	; Fehler --------------------------------------------------------------------------------------------------------------------------------
		else if InStr(EHWT  	, "Dauerdiagnosen von")                                                                                          	{
		  ; Autopositionierung des Dauerdiagnosenfensters
			AlbisResizeDauerdiagnosen()
		}
		else if InStr(EHWT  	, "ICD-10 Thesaurus")                                                                                           	{
		; vergrößert den Diagnosenauswahlbereich
			UpSizeControl("ICD-10 Thesaurus", "#32770", "Listbox1", 200, 150, AlbisWinID())
		}
		else if InStr(EHWText	, "in der Regel meldepflichtig")                                                                                	{
		; schließt Dialog Diagnosen des Kodes .. in der Regel meldepflichtig
			VerifiedClick("Button1", hHookedWin)
		}
		else if InStr(EHWT  	, "Labor - Anzeigegruppen")                                                                                     	{
		; die Anzeige der Listen
			AlbisResizeLaborAnzeigegruppen()
		}
		else If InStr(EHWText	, "Fehler im ifap praxisCENTER")	                                                                                {
		; Fehlerbenachrichtungen Java Runtime des ifap praxisCENTER
			VerifiedClick("Button1", hHookedWin)
			If !ifaptimer
				SetTimer, ifapActive, 30
			ifaptimer:= A_TickCount
		}
		else If InStr(EHWText	, "Bitte die Straße")                                                                                           	{
		; weitere Information Dialog (Patientendaten) und Adressproblematik
			; bei Adressen ohne Straßenangabe, kann das Fenster 'weitere Informationen' nicht ohne weiteres geschlossen werden
			AlbisLoescheEmpfaenger()
		}
		; Formulare -----------------------------------------------------------------------------------------------------------------------------
		else if InStr(EHWT  	, "Muster 1a") || InStr(EHWText, "Der Patient ist noch AU")                                                       {
		; Arbeitsunfähigkeitsbescheinigung - zusätzliche Informationen
			AlbisFristenGui()
		}
		else If InStr(EHWT   	, "Muster G1204")                                                                                               	{
		; Reha-Antrag Dialog an Hochkantmonitore anpassen
			SetTimer(Func("AlbisRehaDialog").Bind(hHookedWin), -100)
		}
		else if InStr(EHWT   	, "Muster 13")                                             	&& Addendum.HMV                                     	{

		}
		else if      (EHWT ~= "i)(Muster\s16|Privatrezept)")                             	&& Addendum.Schnellrezept                           	{
		; Kassen-/Privatrezept - Schnellrezept + Auto autIdem Kreuz
			AlbisRezeptHelferGui(Addendum.DataPath "\RezepthelferDB.json")
		}
		else if Instr(EHWT   	, "Patientenausweis")                                     	&& Addendum.Flags.PatAusweisDruck	                   	{
		  ; automatisch drucken
			VerifiedClick("Drucken", hHookedWin)
			If WinExist("Patientenausweis ahk_class #32770")
				VerifiedClick("Button23", "Patientenausweis ahk_class #32770")
			Addendum.Flags.PatAusweisDruck := false
			PraxTT("Patientenausweis gedruckt", "2 1")
		}
		; ---------------------------------------------------------------------------------------------------------------------------------------
		else If InStr(EHWText	, "elektronisch über KIM")                                                                                      	{
		; eAU Hinweistext
			VerifiedClick("Button1", hHookedWin,,, true)                                                               	; klickt auf: 'Ich bin mir dessen bewusst'
		}
		else If InStr(EHWText	, "Folgende eAU-Tabelle konnte nicht")                                                                          	{
		; Folgende eAU-Tabelle konnte nicht gefunden werden
			if !VerifiedClick("Button1", hHookedWin,,, true)
				VerifiedClick("OK", "ahk_class #32770", "Folgende eAU-Tabelle konnte nicht",, true)
		}
		else If InStr(EHWText	, "Adresse des Rechnungsempfängers")                                                                            	{
		; Schaltfläche zum Löschen der Adressfelder (weitere Informationen....)
			AlbisAdressfelderZusatz()
		}
		;
		; Druckausgabe --------------------------------------------------------------------------------------------------------------------------
		else If InStr(EHWT  	, "Druckausgabe speichern unter")                                                                               	{
		; bei Druck von Privatabrechnungen, Blutwerten, Formularen einen Dateinamen vorschlagen
			AlbisDruckAusgabe(hDruckausgabe := hHookedWin)
		}
		else If InStr(EHWText	, "Beim Laden der Druckereinstellungen für das Dokument")                                                       	{
		; bei Druck über RDP ohne VPN Verbindung werden lokale (Client-) Drucker zwar erkannt aber letztendlich nicht verwendet,
		; es funktioniert dennoch, indem man die Einstellungen für den Druck jedesmal manuell neu einstellt.
			VerifiedClick("OK", hHookedWin)
		}
		else if InStr(EHWText	, "Der Druckauftrag konnte nicht gestartet werden")                                                             	{
		  ; Fenster wird geschlossen
			VerifiedClick("Button1", hHookedWin)
		}
		; ---------------------------------------------------------------------------------------------------------------------------------------
		else If InStr(EHWT  	, "Suchen")                                                                                                      	{
			AlbisKarteikarteDokumentsuche(hHookedWin)
		}
		else If InStr(EHWText	, "Um ein Genesenen-zertifikat zu erstellen")                                                                   	{
		 ; wer will das alle Nase lang sehen, ein Hinweis lasse ich zu, danach automatisch wegklicken
			Addendum.Flags.TIWarner := !Addendum.Flags.TIWarner ? 1 : Addendum.Flags.TIWarner+1
			If (Addendum.Flags.TIWarner > 1)
				VerifiedClick("OK", hHookedWin,,, true)
				;~ VerifiedClick("OK", "ALBIS hk_class #32770", "",, true)
		}
		else If InStr(EHWT  	, "Arbeitsplätze")                                                                                               	{
			WinMove, % "ahk_id " hHookedWin,,,, 725
			ControlGet, hwnd, hwnd,, SysListView321, % "ahk_id " hHookedWin
			WinMove, % "ahk_id " hwnd,, 10, 10, 690
		}
		else If InStr(EHWT  	, "Arztwahl")                                                                                                  		{   ; ##
			VerifiedChoose("Listbox1", "ahk_id " hHookedWin, 1)
			VerifiedClick("OK",,, hHookedWin, true )
		}
		else If InStr(EHWT  	, "COVID-19 Zertifikat für Genesene")                                                                         		{   ; ##
			VerifiedCheck("Button1", "ahk_id " hHookedWin, 1)
			lastWin := EHWT
		}
		else If InStr(EHWT  	, "GNR-Vorschlag zur Befundung")                                                                               		{   ; ##
			AlbisGNRVorschlag()
		}
		else If InStr(EHWText	, "Bei manueller Eintragung oder Änderung")                                                                    		{
		  ; Rezeptänderungen wie ein Ausrufezeichen vor dem Medikamentennamen
			VerifiedClick("Ja", hWinEventHook)
		}
		else If InStr(EHWT  	, "Leistungskette bestätigen")                                                                                		{	; Autogröße
			SetTimer(Func("AlbisResizeLeistungskette"), -100)
		}
		else If InStr(EHWT  	, "CGM HEILMITTELKATALOG")                                                                                    		{	; ##
			If (Addendum.Windows["Albis_CGMHeilmittelkatalog"].AutoPos & Addendum.MonSize)
				SetTimer(Func("AlbisFormularPosition").Bind(hHookedWin, Addendum.Windows["Albis_CGMHeilmittelkatalog"][Addendum.Resolution]), -400)
		}
		else If InStr(EHWT  	, "Karteikartendaten von")                           	                                                           	{
			If Addendum.Karteikartenexport && InStr(FileExist(Addendum.Karteikartenexport "\"), "D") {
				VerifiedClick("Button21", "Karteikartendaten von ahk_class #32770")
				VerifiedSetText("Edit3", Addendum.Karteikartenexport, "Karteikartendaten von ahk_class #32770")
			}
		}
		; Laborabruf Automatisierung ------------------------------------------------------------------------------------------------------------
		else If InStr(EHWT  	, "Anforderung zuordnen")                                                                                        	{
			AlbisLabBuchZuordnen()
		}
		else If InStr(EHWText	, "Anforderungen in die Laborblätter")                                                                          	{
			; Sind Sie sicher, dass alle Anforderungen in die Laborblätter der zugewiesenen Patienten übertragen werden sollen? Ja bin ich!
			If !VerifiedClick("Alle übertragen", "Laborbuch ahk_class #32770", "alle Anforderungen in die Laborblätter")
				If !VerifiedClick("Alle übertragen", "ahk_id " hHookedWin)
					VerifiedClick("Button1", "ahk_id " hHookedWin)
		}
		else If Addendum.Labor.AutoAbruf                                                                                                       	{

			If     	  InStr(EHWText	, "Anforderungen ins Laborblatt")                                                                                       	{
			  ; es wird automatisch mit "Ja" bestätigt
					If !InStr(Addendum.Labor.AbrufStatus, "Failure")
						VerifiedClick("Button1", hHookedWin)
			}
			else If InStr(EHWText	, "Sind Sie sicher, dass alle Anforderungen")                                                                      	{
				; Dialogtext: "Sind Sie sicher, dass alle Anforderungen in die Laborblätter der zugewiesenen Patienten übertragen werden sollen?"
				; erreichbar im Laborbuch per Kontextmenu oder F8
					If !VerifiedClick("Alle übertragen", hHookedWin)
						VerifiedClick("Button1", hHookedWin)
			}
			else If InStr(EHWText	, "alle markierten Laboranforderungen")                                                                            	{
				If !InStr(Addendum.Labor.AbrufStatus, "Failure")
					VerifiedClick("Button1", hHookedWin)                          ; mit ja bestätigen
			}
			else If InStr(EHWText	, "Keine Datei(en) im Pfad") 	                                                                                         	{
			; Laborabrufvorgang wird beendet
				ResetLaborAbrufStatus()
				fn_tmp := Func("AlbisLaborKeineDateien").Bind(hHookedWin)
				SetTimer, % fn_tmp, -2000
				fn_tmp := ""
			}
			else if Instr(EHWT  	, "Labor auswählen")                              	&& !WinExist("elektronischer Labordatenabruf") 	{
			  ; wählt das eingetragene Labor und bestätigt mit 'Ok'
				AlbisLaborAuswählen(Addendum.Labor.LbName)
			}
			else If InStr(EHWT  	, "Labordaten")                                        	&& !WinExist("elektronischer Labordatenabruf") 	{
				AlbisLaborDaten()
			}
			else If InStr(EHWT  	, "GNR der Anford")             	                                                                                          	{
				If !Addendum.Labor.AbrufStatus
					AlbisLaborGNRHandler(1, hHookedWin)       	         	; Listbox1 enthält Rechnungsdaten
			}
			else If InStr(EHWText	, "Übertrage Anforderungen")                                                                                         	{
				; Fortschrittschrittsanzeige nach vorne holen und mittig unterhalb des Dialoges "GNR der Anforderung..." positionieren
					WinSet, AlwaysOnTop, On, % "ahk_id " hHookedWin
					If (hGNRAnf := WinExist("GNR der Anford. ahk_class #32770")) {
						wGNR 	:= GetWindowSpot(hGNRAnf)
						wUAnf	:= GetWindowSpot(hHookedWin)
						SetWindowPos(hHookedWin, Floor((wGNR.X+wGNR.W)/2)+Floor((wAnf.X+wAnf.W)/2), wGNR.Y+wGNR.H+5, wUAnf.W, wUAnf.H)
					}
			}

		}
		; GVU und HKS Formular ------------------------------------------------------------------------------------------------------------------
		else If Addendum.GVUAutomation                                                                                                         	{
			   If InStr(EHWT  	, "Muster 30")                                             	&& (ClickReady > 0)    		{
				RegExMatch(EHWT, "(?<=Seite\s)\d", FormularSeite)
				If (FormularSeite = 1)     				{
					If (ClickReady = 1)							{
						ClickReady := 2
					  ; überprüft auf welche Art das GVU Formular aufgerufen wurde, verhindert die Automatisierung schon angelegter Formulare
					  ; funktioniert nur wenn man sich ein Albismakro mit dem Namen GVU anlegt! In Edit3 steht dann GVU
						EditDatum := EditKuerzel := ""
						ControlGetText, EditDatum , Edit2, ahk_class OptoAppClass
						ControlGetText, EditKuerzel, Edit3, ahk_class OptoAppClass
						If (StrLen(EditDatum) = 0) || !InStr(EditKuerzel, "GVU")
								DatumZuvor := "", ClickReady := 1
					  ; Quartal, Untersuchungsmonat und Jahr ermitteln
						AbrQuartal := LTrim(GetQuartal(EditDatum, "/"), "0")
						RegExMatch(EditDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
						GVU_AlteDatenStartZeit := A_TickCount
						PraxTT("Warte auf alte Daten", "0 3")
						SetTimer, GVU_WarteAufAlteDaten, 250
						If !VerifiedClick("Alte Daten", "Muster 30 ahk_class #32770")
							VerifiedClick("Button63", "Muster 30 ahk_class #32770")	        	; Button63 = [Alte Daten] - GVU Seite 1
					}
					else if (ClickReady = 3)							{
						SetTimer, GVU_WarteAufAlteDaten, Off
						PraxTT("Abbruch", "2 3")
						VerifiedClick("Button62", "Muster 30 ahk_class #32770")	        	; Button62 = [Abbruch]	 - GVU Seite 1
						If RegExMatch(DatumZuvor, "\d{2}\.\d{2}\.\d{4}")
							AlbisSetzeProgrammDatum(DatumZuvor)
						DatumZuvor := ""
						SetTimer, SetClickReadyBack, -4000
					}
					else if (ClickReady = 4)							{
						SetTimer, GVU_WarteAufAlteDaten, Off
						PraxTT("Weiter", "2 3")
						VerifiedClick("Button61", "Muster 30 ahk_class #32770")	        	; Button61 = [Weiter]   	 ->	GVU Seite 2
					}
				}
				else if (FormularSeite = 2)				{
					ClickReady := 5
					SetTimer, GVU_WarteAufAlteDaten, Off
					PraxTT("GVU Formular Seite 2 schließen", "2 3")
					VerifiedClick("Button61", "Muster 30 ahk_class #32770", "", "", 3)      	; Button61 = [Weiter]   	 ->	GVU Seite 2
				}
			}
		else If InStr(EHWText	, "keine alten Daten vorhanden")                	&& (ClickReady = 2)			{
			ClickReady := 0
			PraxTT("", "off")
			SetTimer, GVU_WarteAufAlteDaten, Off
			GVU_KeineAlteDatenVorhanden()
			ClickReady := 5
		}
		else If InStr(EHWT  	, "Alte Formulardaten übernehmen")           	&& (ClickReady = 2)			{
			ClickReady := 0 ; auf 0 gesetzt damit nichts anderes den Ablauf hier behindern kann
			SetTimer, GVU_WarteAufAlteDaten, Off
			ClickReady := GVU_AlteFormulardatenUebernehmen()
		}
		else If InStr(EHWText	, "Soll das Makro abgebrochen werden")      	&& (ClickReady > 1)			{
			VerifiedClick("Button1", "ALBIS", "Soll das Makro abgebrochen werden")
			ClickReady	:= 1
			SetTimer, SetClickReadyBack, off
		}
		else If InStr(EHWT  	, "GNR-Vorschlag zur Befundung")              	&& (ClickReady = 5)			{

			entries := ""
			while (StrLen(entries) = 0)		{

				entries := AlbisReadFromListbox("GNR-Vorschlag zur Befundung", 1, 0)
				If (StrLen(entries) > 0)
					break
				;SciteOutPut("entrieTry: " A_Index "`tQuartal: " AbrQuartal "`tEinträge: " entries)
				Sleep 300
				If (A_Index > 5) 	{
					PraxTT("Die Erkennung des Abrechnungsquartals im Dialog hat nicht funktioniert`nBitte wählen Sie das richtige Abrechnungsquartal manuell aus!", "8 4")
					ClickReady:= 6
				}

			}

			Loop, Parse, entries, `n
				If InStr(A_LoopField, AbrQuartal)	{
					PostMessage, 0x186, % A_Index - 1, 0, ListBox1, % "ahk_id " WinExist("GNR-Vorschlag zur Befundung ahk_class #32770") ;LB_SetCursel
					entrieIndex:= A_Index - 1
					break
				}

			;ToolTip, % "entrieId: " entrieIndex "`nQuartal: " AbrQuartal "`nEinträge: " entries, 1250, 1, 5
			If !VerifiedClick("Button1", "GNR-Vorschlag zur Befundung ahk_class #32770", "", "", 3) ;GNR-Vorschlag schliessen
					MsgBox, Bitte schließen Sie den Dialog`nGNR-Vorschlag zur Befundung

			ClickReady := 6
			WinWait, Hautkrebsscreening - Nichtdermatologe ahk_class #32770,, 8
			If ErrorLevel
				WinActivate, Hautkrebsscreening - Nichtdermatologe ahk_class #32770
			else
				Albismenu(34505, "Hautkrebsscreening - Nichtdermatologe")     	;"eHautkrebs-Screening Nicht-Dermatologe": "34505" wird aufgerufen

		}
		else If InStr(EHWT  	, "Hautkrebsscreening - Nichtdermatologe")	&& (ClickReady = 6)			{
			GVU_HKSFormularBefuellen()
			ClickReady:= 7
		}
		else If InStr(EHWT  	, "Leistungskette bestätigen")                        	&& (ClickReady = 7)			{
			ClickReady:= 1
			GVU_LeistungsketteBestaetigen()
			GVU_CaveVonEintragen(EditDatum)
			If RegExMatch(DatumZuvor, "\d{2}\.\d{2}\.\d{4}")
				AlbisSetzeProgrammDatum(DatumZuvor)
			DatumZuvor := ""
		}
		else If InStr(EHWText	, "Übertrage Gebühren-Nummer(n)...")		                                        	{
			VerifiedClick("Button1", "ALBIS", "Übertrage Gebühren-Nummer")
		}
		else If InStr(EHWText	, "außerhalb des Quartals in dem der Schein gültig ist")		                	{ 	;<-- überarbeiten
			VerifiedClick("Button1", hHookedWin)
		}
		}

	; --------------------------------------------------------------------------------------------------------------------------------------------------------
	}
	else if InStr(EHproc1, "wkflsr32")                                                                                    		{
		if InStr(EHWT  	, "CGM HEILMITTELKATALOG")                                                                                        {
		  ; Fensterposition wird innerhalb des Albisfenster zentriert
			AlbisHeilmittelKatologPosition()
		}
	}
	else if InStr(EHproc1, "AoWDump2")                                                                                      	{        	; Absturzfenster
		If InStr(EHWT, "ALBIS(TM)") {   ;-- bearbeitet das Fenster Absturzbericht automatisch
			If !vacation.isConsultationTime("", "", whynotwork)
				AlbisAbsturzbericht(hHookedWin, A_Now, A_TimeIdlePhysical)
		}
	}
	else if InStr(EHproc1, "mDicom")                     && (Addendum.mDicom)                                               	{        	; mDicom Viewer
		If InStr(EHWT, "Export to video")
			MicroDicom_VideoExport()
	}
	else if InStr(EHproc1, "ipc")                                                                                           	{       	; ifap Programm
		If InStr(EHWText, "Fehler im ifap praxisCENTER") {
			VerifiedClick("Button1", hHookedWin)
			If !ifaptimer
				SetTimer, ifapActive, 50
			ifaptimer:= A_TickCount
		}
	}
	else if InStr(EHproc1, "infoBoxWebClient")           && (Addendum.Labor.AutoAbruf)                                      	{        	; fängt das WebFenster meines Labors ab
		If !WinExist("elektronischer Labordatenabruf") {
			IniWrite, % A_DD "." A_MM "." A_YYYY " (" A_Hour ":" A_Min ":" A_Sec ")", % Addendum.Ini, % "Labor", % "letzter_Laborabruf"
			If !Addendum.Labor.AbrufStatus  {
				Addendum.Labor.AbrufStatus := "Init"
				fn_tmp := Addendum.Labor.Reset := Func("ResetLaborAbrufStatus")
				SetTimer, % fn_tmp, -600000   ; 10 min
				fn_tmp := ""
				PraxTT("Das Laborabruf Fenster wurde detektiert!`n#3Addendum übernimmt den weiteren Vorgang!", "10 1")
				vianovaInfoBoxWebClient(hHookedWin)
			}
		}
	}
	else if Instr(EHproc1, "Scan2Folder")	                                                                                  	{        	; Fujitsu Scansnap Software
		If Instr(EHWText, "Die Dateien wurden erfolgreich gespeichert")
			VerifiedClick("Button1", hHookedWin)
	}
	else if Instr(EHproc1, "Autohotkey")                                                                                     	{        	; WinSpy (Autohotkey Tool) und Debuggerfenster
		If InStr(EHWT, "WinSpy") {
			;~ cp := GetWindowSpot(hHookedWin)
			;~ WinMoveZ(hHookedWin, 20, cp.X, cp.Y, cp.W, cp.H, Redraw:=0)
			;~ MoveWinToCenterScreen(hHookedWin)
		}
		;~ else if InStr(EHWT, "Variable list")
			;~ WinMoveZ(hHookedWin, 0, Floor(A_ScreenWidth/2), 300, 600, 1200)
	}
	else if Instr(EHproc1, "InternalAHK")                                                                                    	{       	; SciTE Editor - Variablenfenster verschieben
    	If InStr(EHWT, "Variable list")
			WinMoveZ(hHookedWin, 0, Floor(A_ScreenWidth/2), 300, 600, 1200)
	}
	else if Instr(EHproc1, "SciTE")           	                                                                             	{        	; SciTE Editor
		If RegExMatch(EHWT, "Addendum\.ini\s+in.*SciTE4AutoHotkey") && Instr(WinGetClass(hHookedWin), "#32770") {
			; nach dem Laden einer veränderten Addendum.ini Datei - wird diese von SciTE immer entfaltet.
			; Das macht alles unübersichtlich. Deshalb wird hier automatisches Codefolding durchgeführt
				VerifiedClick("Button1", "SciTE4AutoHotkey ahk_class #32770")
				SciTeFoldAll()
		}
	}
	else if InStr(EHproc1, "FoxitReader")                                                                                    	{        	; Foxitreader signieren vereinfachen (und zwar tatsächlich!)
		If Addendum.PDF.RecentlySigned {

			If (Addendum.PDF.RecentlySigned && InStr(EHWT, "Speichern unter bestätigen")) {
				VerifiedClick("Ja", "Speichern unter bestätigen ahk_class #32770",,, true)    	; mit 'Ja' bestätigen
				Addendum.PDF.RecentlySigned := false
				TTipOff(10, 8000)
			}
			If InStr(EHWT, "Speichern unter") {  ; && RegExMatch(EHWText, "i)(Speichern|Save)")
				Addendum.PDF.RecentlySigned := 2
				If !WinExist("Sepichern unter bestätigen ahk_class #32770")
					If VerifiedClick("Speichern", "Speichern unter ahk_class #32770")                 	; Speichern Button drücken
						Addendum.SaveAsInvoked := true
				WinWait, % "Speichern unter bestätigen ahk_class #32770" ,, 5
				If WinExist("Speichern unter bestätigen ahk_class #32770") {
					VerifiedClick("Ja", "Speichern unter bestätigen ahk_class #32770",,, true)    	; mit 'Ja' bestätigen
					Addendum.PDF.RecentlySigned := false
					TTipOff(10, 8000)
				}
			}

		}
		else If (Addendum.PDFSignieren && RegExMatch(EHWT, "i)(Sign\sDocument|Dokument\ssignieren)"))          	; für die englische und deutsche FoxitReader Version
			FoxitReader_SignDoc(hHookedWin)
	}
	else if InStr(EHproc1, "iexplore")                                                                                       	{       	; Windows Internetexplorer
		If (StrLen(Addendum.Labor.Kennwort) > 0) && InStr(EHWT, "CGM CHANNEL: Login")	 ; CGM CHANNEL: Login
			LabLogin(Addendum.Labor.Kennwort)
	}
	else if InStr(EHproc1, "Explorer")                                                                                       	{       	; Windows Dateiexplorer
		If InStr(EHWT  	, "Datei löschen")                                                                             	{
			hookedPID := WinGet(hHookedWin, "PID")
			WinGet, hParent, ID, % "ahk_pid " hookedPID
		}
	}
	else if InStr(EHproc1, "dbview")                                                                                         	{       	; dbviewer.exe
		If !VerifiedClick("Button1", hHookedWin)
			VerifiedClick("Button1", "Vielen Dank ahk_class #32770")
	}
	else if InStr(EHproc1, "OUTLOOK")                                                                                       	{       	; dbviewer.exe
		If InStr(EHWT, "Regeln und Benachrichtigungen") {
			ControlGet, cRBhwnd, hwnd,,  SysTreeView321, % "ahk_id " (hWin := hHookedWin)
			wp := GetWindowSpot(hWin), cp := GetWindowSpot(cRBhwnd)
			newH := A_ScreenHeight-(cp.Y-wp.Y)
			newY := A_ScreenHeight//2 - newH//2
			newY := newY<1 ? 1 : newY
			SetWindowPos(hWin, wp.X, newY, wp.W, newH)
			ControlMove,,,,, % newH - (DeltaH := wp.H-cp.H), % "ahk_id " cRBhwnd
			;~ ControlClick,, % "ahk_id " hWin,, Left, 1
		}
	}

	If (EHStack.Count() > 0)
		goto StackEntry

	Addendum.LastHookHwnd := 0
	EHWHState := false

return

ifapActive:                  	;{ holt das ifap Fenster (nach der hoffentlich letzten Fehlermeldung) in den Vordergrund
	if (A_TickCount - ifapTimer > 200) {
		ifaptimer := 0
		WinActivate, praxisCENTER ahk_class TForm_ipcMain
		SetTimer, ifapActive, off
	}
return ;}


;}

AlbisFormularPosition(hwnd, Settings) {

	mon := GetMonitorIndexFromWindow(hwnd, true)
	SciTEOutput(A_ThisFunc "() | " cJSON.Dump(Settings, 1))
	SetWindowPos(hwnd, 880, 10, 1020, 1020 )
	;~ SetWindowPos(WinExist("Muster 13 ahk_class #32770"), 230, 10, 660, 993)


}

SetTimer(boundFunc, delay)                                                                   	{	;-- SetTimer Wrapper für BoundFunc-Aufrufe
	SetTimer, % boundFunc, % delay
}

LabLogin(passwd)                                                                            	{	;-- einfache Passwortübergabe für das Login im Laborfenster

	; 4.3.4.1.4.3.4.1.4.1 - Pfad für URL-Eingabe
	; KennwortBez := "4.5.4.3.4.1.4.1.1.2.4" ;8
	; Kennwortfeld := "4.5.4.3.4.1.4.1.1.2.4.9"
	; KennwortEnter := "4.5.4.3.4.1.4.1.1.2.4.10"

	SendRaw, % passwd
	Sleep 500
	SendInput, {Enter}
	PraxTT("CGM Channel Login wurde durchgeführt.", "6 2")

}

FEGui(hwnd)                                                                                 	{	;-- Fokushook Testgui

	;global
	static FEText1, FE, hFEText1, hFEGui

	cpos  	:= GetWindowSpot(hwnd)
	ControlGetFocus, AlbisFocus, % "A"
	WinGetActiveStats, AlbisTitle, winW, winH, winX, winY
	popInfo	:= "title:`t" 	AlbisTitle                                                     	"`n"
	popInfo	.= "class:`t" 	GetClassName(hwnd)                                	"`n"
	popInfo	.= "focus:`t"	AlbisFocus                                                 	"`n"
	popInfo	.= "hwnd:`t"	Format("0x{:x}", hwnd)                               	"`n"
	popInfo	.= "pos:`t"   	"x" cpos.X " y" cpos.Y " w" cpos.W " h" cpos.H


	If !WinExist("FEGui ahk_class AutohotkeyGUI") {

		Gui, FE: New	, +ToolWindow -Caption -DPIScale +AlwaysOnTop HWNDhFEGui
		Gui, FE: Margin, 3, 3
		Gui, FE: Color	, FDE7CE
		Gui, FE: Font 	, s7 q5
		Gui, FE: Add 	, Text, xm ym vFEText1 HWNDhFEText1, % popInfo
		;WinSet, Trans, 100, % "ahk_id " hFEGui
		Gui, FE: Show, x800 y0 NoActivate, FEGui

	}
	else {

		GuiControl, FE:, FEText1
		Gui, FE: Show, NoActivate

	}

	SetTimer, FEGuiOff, -8000

return

FEGuiOff:
	Gui, FE: Show, Hide
return
}

CurrPatientChange(AlbisTitle)                                                                	{	;-- behandelt Änderung des Albisfenstertitels

	; letzte Änderung: 28.11.2021

		global 	hadm
		static 	AlbisTitleO, AktiveAnzeigeO, lastPatID

		Critical ; ,50 ??

		AktiveAnzeige := AlbisGetActiveWindowType()
		If (AlbisTitleO=AlbisTitle && AktiveAnzeigeO=AktiveAnzeige)
			return
		 AlbisTitleO := AlbisTitle, AktiveAnzeigeO := AktiveAnzeige

		Critical Off

	; Zurücksetzen des Laborabrufstatus bei Wechsel auf eine andere Ansicht
		Addendum.Labor.AbrufStatus := ""
		PatID := AlbisAktuellePatID()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Patientendatenbank und Corona-Impfhelfer
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (InStr(AktiveAnzeige, "Patientenakte") && PatID != lastPatID) {

			lastPatID 	:= PatID
			PatData 	:= AlbisTitle2Data(AlbisTitle)	            	; aktuelle Patientendaten
			If (PatID && !(res:=TProtMatchID(PatID))) {		    	; PatID wird dem Tagesprotokoll hinzugefügt und gespeichert
				TProtData := PatID "(" A_Hour ":" A_Min ")"
				Addendum.TProtocol.Push(TProtData)
				IniAppend( "<" TProtData ">" , Addendum.TPFullPath, Addendum.TPSectionDate, compname)
				TrayTip("Neuer Tagesprotokolleintrag", "[" PatID "] " cPat.Name(PatID, true), 1)
			}

			If Addendum.CImpf.Helfer && (scriptID := GetScriptID(Addendum.CImpf.ScriptName))
				Send_WM_COPYDATA("NewCase|" PatID "|" Addendum.hMsgGui, scriptID)
		}
		else if InStr(AktiveAnzeige, "Laborbuch") && (Addendum.Laborabruf_Voll)
			Albismenu(34157, "", 6, 1)			;34157 - alle übertragen (entspricht Tastaturkürzel: F8)

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Infofenster
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If admGui_Exist() {                       	; Verbergen wenn eine Bedingung nicht erfüllt ist
			If 	(	!Addendum.AddendumGui                                                                	; Flag ob das Infofenster angezeigt werden soll (Einstellungen)
				|| !WinExist("ahk_class OptoAppClass")                                                  	; Albisfenster ist nicht vorhanden
				|| !RegExMatch(AktiveAnzeige, "i)(iWin|Laborblatt|Biometriedaten)")			; falscher Kontext in Albis geöffnet
				|| (PatID && Addendum.iWin.lastPatID <> PatID)) {									; neuer Patient wird angezeigt (Infofenster wird neu erstellt)
					admGui_Destroy()
				}
		}
		If Addendum.AddendumGui {     	; Anzeigen oder Inhalte auffrischen
			If ((PatID && Addendum.iWin.lastPatID<>PatID) || RegExMatch(AktiveAnzeige, "i)(iWin|Laborblatt|Biometriedaten)"))
				AddendumGui()
			If admGui_Exist() && admGui_ActiveTab("Protokoll")
				admGui_TProtokoll(Addendum.iWin.TProtDate, 0, compname)
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Abrechnungshinweise schließen (kann per Traymenu ausgestellt werden)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If Addendum.EBMKRW && InStr(AlbisTitle, "Prüfung EBM/KRW") && !InStr(AlbisTitle, "Abrechnung vorbereiten") {
			 ; die Quartalsabrechnung wird erstellt, hier das Fenster nicht schliessen
			If !AlbisMDIChildHandle("Prüfung EBM 2000plus Komplexe") || !AlbisMDIChildHandle("Fehlerliste") {
				PraxTT("Die Anzeige Prüfung EBM/KRW wird in`n" (Addendum.EBMKRWTimeout ? Addendum.EBMKRWTimeout : "6") " Sekunden ausgeblendet!")
				SetTimer, EBMKRWOff, % -1 * (Addendum.EBMKRWTimeout ? Addendum.EBMKRWTimeout : 6) * 1000
			}
		}

Return

EBMKRWOff:      ;{

	If InStr(AlbisGetActiveWinTitle(), "Prüfung EBM/KRW") && !InStr(AlbisTitle, "Abrechnung vorbereiten") {
		AlbisActivate(1)
		SendInput, {Esc}
		while (InStr(AlbisGetActiveWinTitle(), "Prüfung EBM/KRW") && A_Index < 20)
			Sleep 20
		If InStr(AlbisGetActiveWinTitle(), "Prüfung EBM/KRW")
			AlbisMDIChildWindowClose("Prüfung EBM/KRW")
	}

return   ;}
}



;}

;{16. Interskriptkommunikation / OnMessage
MessageWorker(InComing)                        	{                                                 	;-- verarbeitet die eingegangen Nachrichten

		static func_AutoOCR	:= func("admGui_OCRAllFiles")

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"      	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}

	; Kommunikation mit dem Abrechnungshelfer Skript
		If 			RegExMatch(recv.txtmsg, "^AutoLogin\s+Off")                        	{
			result := Send_WM_COPYDATA("AutoLogin disabled", recv.opt)
			If Addendum.AutoLoginCanChanged
				Addendum.AutoLogin := false
		}
		else if 	RegExMatch(recv.txtmsg, "^AutoLogin\s+On")	                    	{
			result := Send_WM_COPYDATA("AutoLogin enabled", recv.opt)
			If Addendum.AutoLoginCanChanged
				Addendum.AutoLogin := true
		}

	; Tesseract OCR - Thread Kommunikation
		else if 	InStr(recv.txtmsg, "OCR_processed")                                          	{            	; Texterkennung einer Datei ist abgeschlossen
			admGui_CheckJournal(recv.opt, recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "OCR_ready")                                               	{	          	; der OCR-Thread hat alle Dateien abgearbeitet

			; Anzeige auffrischen
				admGui_Journal()                                              ;## todo: class Journal nutzen
				OCR := Addendum.OCR
				Addendum.tessOCRRunning := false
				If !OCR.RestartAutoOCR
					journal.OCR("-OCR ausführen")

			; den Thread beenden, Speicher freigeben
				If IsObject(Addendum.Thread["tessOCR"]) {
					PraxTT("Tesseract Texterkennung abgeschlossen.", "1 2")
					ahkthread_free(Addendum.Thread["tessOCR"])
					Addendum.Thread["tessOCR"] := ""
				}

			; ist AutoOCR eingeschaltet und sind weitere Dateien vorhanden wird
			; der Texterkennungsvorgang erneut ausgeführt
				If OCR.RestartAutoOCR && OCR.AutoOCR && (OCR.Client = compname)
					SetTimer, % func_AutoOCR, % "-" (OCR.AutoOCRDelay*1000)

		}

	; Laborabruf_iBWC.ahk
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "Status")	{
			Send_WM_COPYDATA("AutoLaborAbruf Status|" Addendum.Labor.AutoAbruf  "|" GetAddendumID(), recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "aus")    	{
			If Addendum.Labor.AutoAbruf {
				Addendum.Labor.AutoAbruf_Last := true
				gosub Menu_LaborAbrufManuell                    ; Traymenueinstellung ändern
			}
			Send_WM_COPYDATA("AutoLaborAbruf angehalten||" GetAddendumID(), recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "an")     	{
			If Addendum.Labor.AutoAbruf_Last {
				Addendum.Labor.AutoAbruf_Last := false
				gosub Menu_LaborAbrufManuell                    ; Traymenueinstellung ändern
			}
			Send_WM_COPYDATA("AutoLaborAbruf fortgesetzt||" GetAddendumID(), recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "restart")	{            	; Laborabruf Skript in ein paar Minuten neu starten
			func_LaborAbruf := Func("LaborAbruf")
			SetTimer, % func_LaborAbruf, % -1*(15*60*1000)  	;  start 15 Minuten
			Send_WM_COPYDATA("AutoLaborAbruf wiederholen|1|" GetAddendumID(), recv.fromID)
		}

	; Addendum Statusabfragen
		else if 	InStr(recv.txtmsg, "Status")                                                        	{
			Send_WM_COPYDATA("Status|okay|" GetAddendumID(), recv.fromID)
		}

return
}

WM_DISPLAYCHANGE(wParam, lParam)   	{                                                	;-- WinPos-Arrange: Scite4AHK - AutoZoom und Autoteilung

	; Funktion die auf eine Änderung der Bildschirmauflösung (z.B. Wechsel der Auflösung bei Nutzung des Windows Remotedesktop) reagiert
	;
	; - passt die Schriftgröße der Monitorauflösung an, damit die Schrift lesbar bleibt
	; - automatische Einstellung der vertikalen oder horizontalen Teilung beider Scite Fenster je nach Orientierung des Monitor (Quer oder Hochkant)
	; - verschiebt WinSpy in die Mitte des Hauptmonitors, falls es außerhalb des sichtbaren Bildschirmbereiches liegen sollte
	;
	; letzte Änderung: 16.11.2021

	static lastScreenSize, SCI_GETZOOM := 2374, SCI_SETZOOM := 2373, SCI_TOGGLE_OUTPANE_ORIENTATION := 401
	static ZoomLevel := {"4k":{"Sci1":"-1", "Sci2":"-1"},  "2k":{"Sci1":"-2", "Sci2":"-2"}} ;Sci1 = Hauptfenster, Sci2 = Ausgabefenster
	static Sci1Height := 250, Sci1Width := 380

	If (hSpy := WinExist("WinSpy"))
		MoveWinToCenterScreen(hSpy)

	If (hScite := WinExist("ahk_class SciTEWindow")) {

	  ; Scite auf maximale Größe bringen, aber nur wenn es aktiviert ist
		If WinActive("ahk_class SciTEWindow") {
			If DllCall("IsIconic"	, "UInt", hScite)                     	; nur wenn minimiert
				WinMaximize, % "ahk_class SciTEWindow",, 0, 0

		  ; zurück wenn Maximierung nicht funktionierte
			If DllCall("IsIconic"	, "UInt", hScite)
				return
		}

	  ; Fenster und Monitordimensionen ermitteln
		SPos      	:= GetWindowSpot(hScite)
		SctM     	:= ScreenDims(GetMonitorAt(SPos.X+SPos.W//2, SPos.Y+SPos.H//2))
		monDim 	:= SctM.W "x" SctM.H

		If (lastScreenSize <> monDim) {

		  ; ZoomLevels festlegen
			IsHighResolution	:= (SctM.H > 1080 && SctM.W > 1920) || (SctM.H > 1920 && SctM.W > 1080) ? true : false
			ZoomLevel_Sci1	:= IsHighResolution ? ZoomLevel.4k.Sci1 : ZoomLevel.2k.Sci1
			ZoomLevel_Sci2 	:= IsHighResolution ? ZoomLevel.4k.Sci2 : ZoomLevel.2k.Sci2

		  ; hwnds der Scite Fenster ermitteln
			ControlGet	, hScintilla1	, hwnd,, Scintilla1	, % "ahk_id " hScite
			ControlGet	, hScintilla2	, hwnd,, Scintilla2	, % "ahk_id " hScite

		  ; ZoomLevels auslesen
			SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla1
			ZoomScintilla1O := ErrorLevel
			SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla2
			ZoomScintilla2O := ErrorLevel

		; ZoomLevels anpassen
			If (ZoomScintilla1O <> ZoomLevel_Sci1)
				SendMessage, SCI_SETZOOM	, % ZoomLevel_Sci1, 0,, % "ahk_id " hScintilla1
			If (ZoomScintilla2O <> ZoomLevel_Sci2)
				SendMessage, SCI_SETZOOM	, % ZoomLevel_Sci2, 0,, % "ahk_id " hScintilla2

		; Output pane: horizontal oder vertikal je nach Orientierung des Monitors einstellen
			If hScintilla2 {
				Sci1Pos          	:= GetWindowSpot(hScintilla1)
				Sci2Pos          	:= GetWindowSpot(hScintilla2)
				isVerticalPane   	:= (Sci1Pos.X = Sci2Pos.X) 	? false : true
				isVerticalMon	:= (SctM.W > SctM.H)       	? false : true
				If (isVerticalMon && isVerticalPane) || (!isVerticalMon && !isVerticalPane)
					SendMessage, 0x111, % SCI_TOGGLE_OUTPANE_ORIENTATION, 0,, % "ahk_id " hScite
			}

		; Output pane: Größe anpassen (funktioniert nur wenn Scite4Autohotkey im Vordergrund ist)
			If hScintilla2 {

				;~ coordMouse  	:= A_CoordModeMouse
				;~ mousedelay   	:= A_MouseDelay

				;~ CoordMode, Mouse, Screen
				;~ MouseGetPos, X, Y
				;~ SetMouseDelay, -1

				;~ Sci1Pos          	:= GetWindowSpot(hScintilla1)
				;~ isVerticalPane   	:= (Sci1Pos.X = Sci2Pos.X) ? false : true

				;~ If !isVerticalPane {
					;~ delta := Sci1Pos.H - Sci1Height, X1 := Sci1Pos.X, Y1 := Sci1Pos.Y-5, X2 := X1+Sci1Pos.W-50, X1 += 30
					;~ DllCall("SetCursorPos", "int", X1, "int", Y1)
					;~ MouseClickDrag, Left, X<X1 ? X1 : X>X2 ? X2 : X, Y1, X1, Y1+delta, 0
					;~ SciTEOutput("delta: " delta " (" Sci1Pos.H "), Sci1X: " Sci1Pos.X ", Sci2X: " Sci2Pos.X )
				;~ }
				;~ else {
					;~ delta := Sci1Pos.W - Sci1Width, X1 := Sci1Pos.X-5, Y1 := Sci1Pos.Y+10, Y2 := Y1+Sci1Pos.H-50
					;~ DllCall("SetCursorPos", "int", X1, "int", Y1)
					;~ MouseClickDrag, Left, X1, Y1, X1+delta, Y1, 0
					;~ SciTEOutput("delta: " delta " (" Sci1Pos.W "), Sci1Y: " Sci1Pos.Y ", Sci2Y: " Sci2Pos.Y )
				;~ }

				;~ DllCall("SetCursorPos", "int", X, "int", Y)
				;~ CoordMode, Mouse	, % coordMouse
				;~ SetMouseDelay     	, % mousedelay
			}

			;~ SciTEOutput(Sci1Pos.X "<" Sci2Pos.X "=" isVerticalPane ":" isVerticalMon)

			lastScreenSize := monDim

		}

	}


}

ScreenEdgeDetection(wParam, lParam, msg) 	{                                               	;-- läßt den Mauszeiger nicht an einem Monitorrand kleben

	static Firstrun := true

  ; erster Start checkt ob eine Überwachung der Mausposition überhaipt notwendig ist.
	If Firstrun {


	}

}

;}

;{17. Automatisierungsfunktionen

;Funktion AutoDiagnose findet sich in include\AutoDiagnosen.ahk

AddAutoComplete(hWin, ControlName, TextList, TWidth:= 350) {

	If (TextList = "")
		Return

	ControlGetPos, cpX, cpY, cpW,, % ControlName, % "ahk_id " hWin
	CpY += 20

	Gui, AutoC: new		, -Caption +ToolWindow +AlwaysOnTop HWNDhAutoC
	Gui, AutoC: Margin	, 0, 0
	Gui, AutoC: Add		, Listbox, % "r" 10 " w" TWidth " vLBAutoC" , % TextList
	Gui, AutoC: Show		, % "x" CpX " y" CpY, Addendum AutoComplete

	Hotkey, IfWinExist, Addendum AutoComplete
	Hotkey, Enter, MedTLB
	Hotkey, Esc, AutoCGuiEscape

return

LBAutoC:
	Gui, AutoC: Submit, NoHide
	ControlSetText, % ControlName, % LBAutoC, ahk_id %lCHwnd%
	SendInput, {Enter}

AutoCGuiClose:
AutoCGuiEscape:
		Gui, AutoC: Destroy
return

}

AutoDoc() {

	static ADoc:= Object(), ADocList:= Object(), actrl:= Object(), LBList:="", ADocInit:= 0, matched:=""

	saX:= A_CaretX, saY:=A_CaretY

	If !ADocInit 	{
		ADoc:= GetFromIni("MaxResults|OffsetX|OffsetY|BoxHeight|ShowLength|Font|FontSize|MinScore", "AutoDoc")
		ADocList["Impferrinnerung"]    		:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "impf"	, "do": "Choose"	, "RichE": "< /TdPP>< /, Pneumovax>< /, GSI>< /, MMR>"                                     	 }
		ADocList["Impfausweis"]     	    	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "impf"	, "do": "Set"		, "RichE": "Impfausweis mitbringen"                                                                                	 }
		ADocList["lko - Gespräch"]	        	:= {"Edit2": "="							, "Edit3": "lko"	, "do": "Input" 	, "RichE": "03230(x:*)"                                                                                		        	 }
		ADocList["VorsorgeColo"]    	    	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "info"	, "do": "Set"		, "RichE": "an die Vorsorge Coloskopie errinnern"                                        						 }
		ADocList["In die Sprechstunde"]   	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "info"	, "do": "Set"		, "RichE": "Es wäre schön wenn der Patient in diesem Jahr in mein Sprechzimmer vorstößt!"}
		For key, val in ADocList
			LBList .= key "|"
		Gui, Suggestions: -Caption +ToolWindow +AlwaysOnTop HWNDhSugg
		Gui, Suggestions: Font, % "s" ADoc.FontSize, % ADoc.Font
		Gui, Suggestions: Add, ListBox, % "x0 y0 h" ADoc.BoxHeight " 0x100 Sort HWNDhSuggLB1 vMatched", % LBList
		ADocInit := 1
	}

	Gui, Suggestions: Show, % "x" (saX+ADoc.OffsetX) " y" (saY - ADoc.OffsetY - ADoc.BoxHeight) " h" ADoc.BoxHeight, AutoDoc
	WinActivate	, % "ahk_id " hSugg
	GuiControl	, Choose, matched, 1
	ControlFocus,, % "ahk_id " hSuggLB1

	HotKey, IfWinExist      	, AutoDoc
	HotKey, Enter            	, FillDocLine
	HotKey, NumPadEnter	, FillDocLine
	HotKey, Esc               	, SuggestionsGuiClose
	Hotkey, IfWinExist

Return

FillDocLine: ;{

	Gui, Suggestions: Submit, Hide
	actrl:= AlbisGetActiveControl("content")

	If (ADocList[matched].Edit2 <> "=")	{
	  ; <-später ein RegExMatch oder Replace für mehr Optionen
		VerifiedSetText("Edit2", StrReplace(ADocList[matched].Edit2, "%A_YYYY%", A_YYYY + (A_MM=12 ? 1:0)) , "ahk_class OptoAppClass", 100)
	}

	VerifiedSetText("Edit3", ADocList[matched].Edit3 , "ahk_class OptoAppClass", 100)

	If (ADocList[matched].Do = "Choose")	{
		VerifiedSetText("RichEdit20A1", ADocList[matched].RichE , "ahk_class OptoAppClass", 100)
		ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
		Sleep, 200
		SendInput, {Home}
		Sleep, 200
		SendInput, {LControl Down}{Right}{LControl Up}
	}

	If (ADocList[matched].Do = "Set")	{
		VerifiedSetText("RichEdit20A1", ADocList[matched].RichE , "ahk_class OptoAppClass", 100)
		ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
		Sleep, 100
		SendInput, {Tab}
	}

	If (ADocList[matched].Do = "Input")	{
		InputBox, faktor, Addendum für AlbisOnWindows, Bitte tragen Sie hier den Faktor für die Gesprächsziffer ein.,,,,,,,, 2
		VerifiedSetText("RichEdit20A1", StrReplace(ADocList[matched].RichE, "*", faktor), "ahk_class OptoAppClass", 100)
		ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
		Sleep, 100
		SendInput, {Tab}
	}

	gosub SuggestionsGuiEscape

;return
;}

SuggestionsGuiEscape: ;{
SuggestionsGuiClose:
	Gui, Suggestions: Hide
	Addendum.Flag.AutoDocCalled:=0
Return ;}

}

GetFromIni(Settings, Section) {

	RegVal       	:= Object()
	registrykeys	:= StrSplit(Settings, "|")
	For index, registrykey in registrykeys {
		IniRead, val, % AddendumDir "\Addendum.ini", % Section, % registrykey
		RegVal[tmpArr[index]] := val
	}

return RegVal
}

MedTrenner(f) {

	global 	MedT
	static 	MedTrenner, MedTTrenner, x, y

  ; Standardschrift im Albisprogramm zur Anzeige der Dauermedikamente ist Arial Standard 11 - am besten mit dieser Schriftart editieren und hier einfügen
	If !MedTrenner
		MedTrenner =
				(LTrim Join|
				#####################
				###   Problemmedikamente    ###
				###            Asthma                ###
				### Bedarfsmedikamente ###
				###           Blutdruck              ###
				###             COPD                ###
				###           Diabetes               ###
				###          Cholesterin            ###
				### Fremdmedikamente ###
				###          Gerinnung             ###
				###             Gicht                  ###
				###              Herz                  ###
				###          Hilfsmittel              ###
				###             Magen               ###
				###             Niere                 ###
				###         Osteoporose         ###
				###            Prostata              ###
				###             Psyche               ###
				###           Rheuma               ###
				###          Schilddrüse           ###
				###          Schmerzen            ###
				###          Stuhlgang            ###
				###          Sonstiges            ###
				Penicillinallergie
				Amoxicillinallergie
				)

	x := A_CaretX, y := A_CaretY + 15

	Gui, MedT: new		, -Caption +ToolWindow +AlwaysOnTop
	Gui, MedT: Margin	, 0, 0
	Gui, MedT: Add		, Listbox, r20 w350 vMedTTrenner , % MedTrenner
	Gui, MedT: Show		, % "x" x " y" y, Dauermedikamten-Trenner

	HotKey, IfWinExist, % "Dauermedikamten-Trenner ahk_class AutoHotkeyGUI"
	HotKey, Enter, MedTLB
	HotKey, Esc	, MedTGuiEscape
	Hotkey, IfWinExist

Return

MedTLB: ;{

	Gui, MedT: Submit, NoHide
	WinActivate   	, Dauermedikamente ahk_class #32770
	WinWaitActive	, Dauermedikamente ahk_class #32770,, 2
	y -= 15
	Click, %x%, %y%, 2
	Sleep, 200
	SendRaw, % MedTTrenner
	SendInput, {Enter}

;}

MedTGuiClose: ;{
MedTGuiEscape:
	Gui, MedT: Destroy
return ;}
}

InChronicList(PatID) {                                                                                    	;-- Neuaufnahme in Chroniker Liste

		GruppenName := "Chroniker"

		For key, ChronikerID in Addendum.Chroniker
			If (PatID = ChronikerID)
				return

	; Nutzer abfragen ob Pat. aufgenommen werden soll
		hinweis := "Pat: " cPat.NAME(PatID, true)  "`n"
		hinweis .= "ist nicht als Chroniker vermerkt.`nMöchten Sie automatisch alle Eintragungen`ninnerhalb von Albis vornehmen lassen?"

		MsgBox, 0x1024, Addendum für Albis on Windows, % hinweis, 10
		IfMsgBox, Yes
		{
				ChronCb	:= Indikation	:= PatGruppe:= false

			; Ziffer der Liste hinzufügen und in die Datei speichern
				Addendum.Chroniker.Push(PatID)
				FileAppend, % PatID "`n", % Addendum.DBPath "\PatData\DB_Chroniker.txt"

			; Patientengruppierung vornehmen  ;{

				failed := false

				Albismenu("34362", "Patientengruppen für ahk_class #32770")
				ControlGet, result, List, , % "SysListView321", % "Patientengruppen für ahk_class #32770"

				If !InStr(result, GruppenName) {

							PraxTT("Patient ist nicht der Chronikergruppe zugeordnet", "3 2")

						; click auf NEU
							VerifiedClick("Button2", "Patientengruppen für ahk_class #32770")
							Sleep, 300
							WinWait, % "Patientengruppen ahk_class #32770",, 2

						; Click hat funktioniert
							If WinExist("Patientengruppen ahk_class #32770") {

								; vorhandene Gruppeneinträge auslesen und Position der Chronikergruppe finden
									hwnd := WinExist("Patientengruppen ahk_class #32770")
									ControlGet, result, List, , % "Listbox1", % "ahk_id " hwnd

									Loop, Parse, result, `n
										If InStr(A_LoopField, GruppenName) {
											ListboxRow := A_Index
											break
										}

									If (ListBoxRow > 0) 	{

										; den Listboxeintrag mit der Chronikergruppe auswählen
											VerifiedChoose("Listbox1", "ahk_id " hwnd, ListBoxRow)
											VerifiedClick("Button1",  "ahk_id " hwnd)

											PatGruppe := true

									} else
											failed := true

							} else
    								failed := true

				}

			; Fenster 'Patientengruppe für' schließen
				If WinExist("Patientengruppen für ahk_class #32770") {

						If failed {

						; drückt auf Abbrechen
							VerifiedClick("Button7", "Patientengruppen für ahk_class #32770", "", "", true)
							Sleep 200
							If WinExist("Patientengruppen für ahk_class #32770")
								WinClose, % "Patientengruppen für ahk_class #32770"

						} else {

						; OK - Button
							VerifiedClick("Button1", "Patientengruppen für ahk_class #32770", "", "", true)

						}

				}

			;}

			; Chroniker Häkchen setzen und bei Ausnahmeindikation (weitere Informationen) diese Ziffer hinzufügen ;{

				failed        	:= false
				hPersonalien := Albismenu("32774", "Daten von ahk_class #32770") ; Menu 'Personalien'

				If hPersonalien {

						If VerifiedCheck("Chroniker"                      	, "ahk_id " hPersonalien)
							ChronCb := true

					; weitere Information für nächsten Dialog drücken
						VerifiedClick("Weitere In&formationen..."	, "ahk_id " hPersonalien)
						WinWait, % "ahk_class #32770", % "Adresse des", 2

						If (hInformationen := WinExist("ahk_class #32770 ahk_exe albis", "Ausnahmeindikation")) {

							; Feldinhalt wird ausgelesen, bei fehlender Eintragung wird die Ziffer hinzugefügt
								ControlGetText, ctext, Edit14, % "ahk_id " hInformationen
								If !InStr(ctext, "03220") {

										ctext := LTrim(ctext "-03220", "-")

										If VerifiedSetText("Edit14", cText, "ahk_class #32770", 200, "Ausnahmeindikation")
											indikation := true

										ControlFocus	, 					  % "Edit14", % "ahk_id " hInformationen
										ControlSend	, % "{Enter}"	, % "Edit14", % "ahk_id " hInformationen

								}

							; Fenster schliessen falls noch geöffnet
								If WinExist("ahk_class #32770", "Ausnahmeindikation") {
									ControlClick, Ok	, % "ahk_id " hInformationen
									WinWaitClose		, % "ahk_id " hInformationen,, 2
								}

						}

						; Fenster "Daten von " schliessen
							If WinExist("Daten von", "Anrede")
								VerifiedClick("Button30", "ahk_id " hPersonalien)
				}

			;}

				processed := "1. Chronikergruppe     : " (PatGruppe 	? "hinzugefügt" 	: "nicht hinzugefügt") "`n"
				processed .= "2. Chronikerhäkchen   : " (ChronCb 	? "gesetzt"     	: "nicht gesetzt") "`n"
				processed .= "2. Ausnahmeindikation: " (indikation 	? "Lk eingefügt"	: "Lk nicht eingefügt") "`n"
				PraxTT("Eingruppierung abgeschlossen.`n" processed, "6 3")

		}

}

;}

;{18. Gui's
ClipboardPatienten(PatListe := "") {                                                                           	;-- Patienten aus dem Clipboard anzeigen

	static hCPL, CPLLv, CPLClose, PatListeStored

	lvRows := (PatListe.Count()<=20 && PatListe.Count()>=3) ? PatListe.Count() : (PatListe.Count()>=20) ? 20 : 3
	PatListe := !IsObject(PatListe) && IsObject(PatListeStored) ? PatListeStored : PatListe

	Gui, CPL: New, +HWNDhCPL
	Gui, CPL: Add, Text   	, % "xm ym "                                                          	, % "Clipboard: " PatListe.Count() " Patient" (PatListe.Count()>1 ? "en":"") " erkannt"
	Gui, CPL: Add, ListView	, % "xm y+5 w300 r" lvRows " vCPLLv gCPLHandler"	,  % "PatID|Patientenname"
	Gui, CPL: Add, Button	, % "xm y+10 vCPLClose gCPLHandler"                  	, % "Patientenliste schließen"

	LV_ModifyCol(1, 80)
	LV_ModifyCol(2, 200)
	For PatID, Pat in PatListe
		LV_Add("", PatID, cPat.NAME(PatID, true))

	Gui, CPL: Show, AutoSize, % StrReplace(A_ScriptName, ".ahk") " - Clipboardhandler"

	PatListeStored := PatListe
	PatID := Pat := ""

return

CPLHandler:  ;{

	Critical
	If (A_GuiControl="CPLClose") {
		Gui, CPL: Destroy
	}
	else if (GuiControl="CPLLv") {

		If (A_GuiControlEvent = "DoubleClick") {
			LV_GetText(PatID, A_EventInfo, 1)
			If (PatID && cPat.Exist(PatID))
				AlbisAkteOeffnen(cPat.Name(PatID, false), PatID)
		}

	}

return ;}

CPLGuiClose:
	Gui, CPL: Destroy
return
}


; nicht mehr Infofenster
AutoListview(ColStr, newline, ColSize:="", Separator:=",")                 	{	;-- eine Listview Gui zum schnellen Anzeigen von Daten

	static Init, colmax, DasLV, hALV, ALV, hDasLV
	local idx

	If !Init {

			Init	:= true
			Sysget, Mon1, Monitor, 1

			If (Mon1Right>1920)			{
					gx         	:= 2500
					gy         	:= 50
					gw        	:= 1000
					gh        	:= 1024
					GMinW  	:= 400
					GMinh   	:= 400
					gmaxw  	:= Floor(Mon1Right/3)
					gmaxh  	:= (Mon1Bottom - 100)
					gfontSize	:= 11
					grows   	:= 40
					If ColSize	:= ""
					    	ColSize:= 30,80,120,120,350,120,120
			}
			else			{
					gx         	:= 1200
					gy         	:= 50
					gw        	:= 600
					gh        	:= 400
					GMinW 	:= 400
					GMinh  	:= 400
					gmaxw  	:= Floor(Mon1Right/3)
					gmaxh  	:= (Mon1Bottom - 100)
					gfontSize	:= 8
					grows   	:= 20
		    		If ColSize	:= ""
				        	ColSize:= 20,50,70,60,250,60,60
			}

		; Größe des ListView-Controls
			LVw := gw - 30
			LVh := gh	 - 10

		; Gui zeichnen
			Gui, ALV:NEW
			Gui, ALV:Font, % "s" gFontSize , Futura Bk Bt
			Gui, ALV:Margin, 0, 0
			Gui, ALV:+LastFound +AlwaysOnTop +Resize +HwndhALV MinSize%Gminw%x%Gminh% MaxSize%Gmaxw%x%Gmaxh%
			Gui, ALV:Add, Listview, % "xm ym w" LVw " h" LVh " r" gRows " vDasLV HWNDhDasLV BackgroundAAAAFF Grid" , % ColStr 					;
			Gui, ALV:Show, % "x" gx " y" gy , Addendum für Albis on Windows
			Gui, ALV:Default

			Loop, Parse, ColSize, `,
				LV_ModifyCol(A_Index, A_LoopField)

	}

	;Gui, ALV:Default
	;Gui, ALV: Listview, DasLV
	;GuiControl, ALV: -Redraw, DasLV
	If IsObject(newline)	{
		For idx, value in newline
			LV_Add("", idx, value)
	}
	else	{
		data:= StrSplit(newline, Separator)
		LV_Add("", data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8])
	}
	GuiControl, ALV: +Redraw, DasLV

Return

ALVGuiSize:
	Critical, Off
	if A_EventInfo = 1  ; Das Fenster wurde minimiert.  Keine Aktion notwendig.
		return
	Critical
	wNew:= A_GuiWidth, hNew:= A_GuiHeight
	GuiControl, Move, %hDasLV%, % "w" . wNew . " h" . hNew
	Critical, Off
return

}

Calendar(feld)                                                                                  	{	;-- Ersatz für den Kalender der sich mit Shift+F3 in Albis öffnet

	; letzte Änderung 13.06.2022

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		; Objekt enthält Information über die ClassNN von Datumfeldern in Albisdialogen
		; und das zugehörige Partnerfeld eines Datumfeldes, so es eines gibt (von ... bis ...)
		global CalFormulare, newCal, hmyCal
		static feldB, Cok, Can, Hinweis, MyCalendar, hCal, CalBorder, FormularT, AUTagesWarner, lasthWin, CalDay
		If !IsObject(CalFormulare)
			return

		feldB            	:= feld
		FormularT	   	:= ""

	  ; wird auf null gesetzt wenn ein neues Dialogfenster aufgerufen wurde
		If (feldB.hWin <> lasthWin) {
			lasthWin           	:= feldB.hWin
			AUTagesWarner 	:= 0
			SetTimer, CloseOnChangeOrTime, Off
			ToolTip,,,, 10
		}
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnungen für Datumsfokus (ein Datum wird immer übermittelt und sei es das aktuelle)
	;--------------------------------------------------------------------------------------------------------------------;{
		datum	:= feldB.datum
		day   	:= SubStr(datum, 1, 2), month:= SubStr(datum, 3, 2), year:= SubStr(datum, 5, 4)
		month	:= SubStr( "00" . ((month-1<1 ? 12:0) + month-1), -1)
		year   	:= month-1<1 ? year-1 : year
		FormatTime, begindate	, % year month day	, yyyyMMdd
		FormatTime, datum   	, % datum               	, yyyyMMdd
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender Gui
	;--------------------------------------------------------------------------------------------------------------------;{
		Gui, newCal: New, +Owner +AlwaysOnTop HWNDhCal
		Gui, newCal: Add, Text, vHinweis Center                                                         	, % "(" (feldB.to	?	 "halte Shift während der Auswahl des Anfang- und Enddatums"
																																						: 	 "Datum durch Mausklick/Tastatur auswählen") ")"
		Gui, newCal: Add, MonthCal	, % "vMyCalendar R3 W-3 4 HWNDhmyCal " (feldB.to ? "Multi" : "")	, % begindate
		Gui, newCal: Add, Button   	, % "default        	gCalendarOK    	vCok "                               	, % "Datum übernehmen"
		Gui, newCal: Add, Button   	, % "xp+200 yp 	gCalendarCancel	vCan "                               	, % "Abbrechen"

		GuiControl, newCal:, MyCalendar, % datum
		Gui, newCal: Show, AutoSize Hide , % "Addendum für Albis on Windows - erweiterter Kalender"

	;-:Größen und Positionen ermitteln
		WinA 	:= GetWindowSpot(hCal)
		WinB	:= GetWindowSpot(feldB.hWin)
		feldP  	:= GetWindowSpot(feldB.hwnd)

		GuiControlGet, MyCal	, newCal: Pos, MyCalendar
		GuiControlGet, CokPos	, newCal: Pos, Cok
		GuiControlGet, CanPos	, newCal: Pos, Can
		SendMessage 0x101F, 0, 0,, % "ahk_id " hmyCal ; erhalte d. Breite des Kalenderrandes (MCM_GETCALENDARBORDER)
		CalBorder := ErrorLevel

	;-:Position der Gui-Elemente anpassen
		GuiControl, newCal: MoveDraw, Hinweis	, % "x" MyCalX " w" MyCalW
		GuiControl, newCal: MoveDraw, Cok    	, % "x" (MyCalX + 2*CalBorder)
		GuiControl, newCal: MoveDraw, Can    	, % "x" (MyCalX + MyCalW - CanPosW - 2*CalBorder)

	;-:Berechnung der letztendlichen Position des Kalenders
		If RegExMatch(feldB.WinTitle, "i)Muster\s+(1a|12a|17|20|21)") || InStr(feldB.WinTitle, "Programmdatum")
			nCalX := WinB.X+WinB.W, nCalY := feldP.Y-20, tree := 1
		else
			nCalX := feldP.X+feldP.W+10, nCalY := feldP.Y-10, tree := 2

		;~ SciTEOutput("`nfeldB.WinTitle: " feldB.WinTitle "`ntree: " tree "`nclass(" feldB.hWin "): " WinGetClass(feldB.hWin))

	;-:y-Position anpassen, wenn der Kalender nach unten aus dem Bildschirm herausragt
		nCalY := nCalY+WinA.H > A_ScreenHeight ? A_ScreenHeight-WinA.H-30 : nCalY

	;-:x-Position von Kalender und Elternfenster anpassen, wenn der Kalender über den Monitor rechts hinausragt
		If !(isinva := IsInsideVisibleArea(nCalX+WinA.W, 1, 1, 1, CoordInjury)) {
			nCalX := 	WinB.X - WinA.W - 5 >= 0 ? WinB.X-WinA.W-5  : "np"
			If (nCalX = "np") {
				hMon	:= MonitorFromWindow(feldB.hWin)
				mon  	:= GetMonitorInfo(hMon)
				npX  	:= Floor((mon.R - WinB.W - WinA.W)/2)
				nCalX	:= npX + WinB.W + 5
				SetWindowPos(feldB.hWin, npX, WinB.Y, WinB.W, WinB.H, 0x1C)
			}
		}

		Gui, newCal: Show, % "x" nCalX " y" nCalY , Addendum für Albis on Windows - erweiterter Kalender

Return ;}

CalendarOK: 				                                                                                                              	 ;{

		Critical

		Gui, newCal: Submit
		Gui, newCal: Destroy

	  ; behandelt ein einzelnes Datum oder einen ein Anfanggs- und Enddatum
		 CalDay := StrSplit(MyCalendar, "-")
		 CalendarM1 := FormatTime(CalDay.1, "dd.MM.yyyy")
		 CalendarM2 := CalDay.2 ? FormatTime(CalDay.2, "dd.MM.yyyy") : "" ;CalendarM1

		 ;~ SciTEOutput("CalDay.1: " CalDay.1 ", M1: " CalendarM1 "`nCalDay.2: " CalDay.2 ", M2: " CalendarM2 "`nfocused: " feldB.FocusFD)

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; wenn 2 Kalendertage ausgewählt wurden (von - bis), werden diese jetzt eingetragen
	  ; CalendarM1 ins "von" oder "seit" Datumsfeld und
	  ; CalendarM2 ins "bis" oder "Vor. bis einschl." Datumsfeld
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (CalendarM1 && CalendarM2) && (CalendarM1 <> CalendarM2) {

		  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; Sonderbehandlung des Muster 12a (Formular häusliche Pflege)
			If RegExMatch(feldB.WinTitle, "i)Muster\s+12a") {
				editgroup	:= [[3,4], [7,8], [11,12], [15,16]]
				Loop % editgroup.MaxIndex()
					If (feldB.classnn=editgroup[A_Index, 1] || feldB.classnn=editgroup[A_Index, 2])	{
						VerifiedSetText("Edit" editgroup[A_Index   	, 1], CalendarM1, feldB.hWin)
						VerifiedSetText("Edit" editgroup[A_Index   	, 2], CalendarM2, feldB.hWin)
						VerifiedSetText("Edit" editgroup[A_Index+1	, 1], CalendarM1, feldB.hWin)
						break
					}
			}

		  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; bei allen anderen Formularen
			else  {

			  ; überträgt die Datumzahlen in ihr zugehöriges Steuerelement
				If CalendarM1
					res1 := VerifiedSetText("Edit" feldB.from	, CalendarM1, "ahk_id " feldB.hWin)
				If CalendarM2
					res2 := VerifiedSetText("Edit" feldB.to  	, CalendarM2, "ahk_id " feldB.hWin)

			  ; gibt eine Warnung aus wenn eine Arbeitsunfähigkeitsnbescheinigung für mehr 30 Kalendertage ausgestellt wurde
				DaysBetween := 0
				If RegExMatch(feldB.WinTitle, "i)Muster\s*1a")  && (CalDay.1 && CalDay.2) {

					;~ SciTEOutput(A_ThisFunc " (" A_LineNumber  "): Bin hier")
					DaysBetween := DaysBetween(CalDay.1, CalDay.2)
					If (DaysBetween > 30) {
						ControlGet, hwnd, hwnd,, % "Edit" feldB.To, % "ahk_id " feldB.hWin
						cp 	:= GetWindowSpot(hwnd)
						wp 	:= GetWindowSpot(feldB.hWin)
						ToolTip, % "WARNUNG: die AU ist für mehr als 30 Tage ausgestellt!`n"
									. 	DaysBetween " Tag" (DaysBetween>1 ? "e liegen" : "liegt") " zwischen dem " CalendarM1 " und dem " CalendarM2 "!"
									, % wp.X + wp.W - wp.BW, % cp.Y, 10
						AUTagesWarner := 30000                    	; 12sec anlassen
						SetTimer, CloseOnChangeOrTime, 200
					}

				}

			}

		}
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; ansonsten nur das 1.Kalenderdatum in das aktuelle Feld eintragen
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		else
			res1 := VerifiedSetText("", CalendarM1, "ahk_id " feldB.hfocused)


CalendarCancel:
newCalGuiClose:
newCalGuiEscape:
	Sleep 150
	Gui, newCal: Destroy
	WinActivate, % "ahk_id " feldB.hWin
Return

; schließt einen Warnhinweis ToolTip, wenn der "Muster 1a" Dialog geschlossen wurde
CloseOnChangeOrTime:

	AUTagesWarner -= 200
	If (AUTagesWarner <= 0 || !WinExist("Muster 1a ahk_class #32770")) {
		SetTimer, CloseOnChangeOrTime, Off
		AUTagesWarner := 0
		ToolTip,,,, 10
	}


return
;}

}
;}

;{19. Icons
Create_Addendum_ico()                         	{                                                 	;-- erstellt das Trayicon

VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGQAAABkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB/h3lni0pSjiNIkA9BkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNIkA9TjiRnikt/h3kAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACEhoNii0JDkQZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkQZji0OFhoQAAAAAAAAAAAAAAAAAAACBh31SjiJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBRjSOBh30AAAAAAAAAAACEhoNRjSNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBRjiCFhoMAAAAAAABii0JAkQBLlhG807S5062506250625062707JVnCBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJlg240q+506250625062506281LNopjxAkQBji0MAAAB/h3lDkQZAkQBAkQDA17f///////////////////+nyZZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCYwID+/v/////////////////W5NRGlAlAkQBDkAZ/h3lni0pAkQBAkQBAkQBwqkfz9vT////////////////Z5tlGlApAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQHN38n////////////////+/v+ZwIFAkQBAkQBAkQBni0tSjiJAkQBAkQBAkQBAkQDC2br////////////////+/v+TvXlAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB7sFb5+fr////////////////W49RGlAlAkQBAkQBAkQBSjiNIkA9AkQBAkQBAkQBAkQB0rU709/X////////////////M3sdBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC91bP////////////////+/v+ZwYFAkQBAkQBAkQBAkQBHkBBBkQNAkQBAkQBAkQBAkQBAkQDE2bv////////////////5+fp5r1RAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBYniTq8Ov////////////////W5NVGlApAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQB5r1T1+Pf///////////////+91bNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCszJ3////////////////+/v+awoNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDE2r3////////////////t8u9eoC5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGlArb59r////////////////X5dZGlApAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB8sVj3+fj///////////////+uzZ9AkQBAkQBAkQBAkQBAkQBAkQBAkQBIlQzF2b7M3sbM3sbG2sBJlg5AkQBAkQBAkQBAkQBAkQBAkQBAkQCVvnv+/v/////////////+/v+bwoRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQDH28D////////////////f6d9Klg9AkQBAkQBAkQBAkQBAkQBAkQBZnif1+Pn////////3+PtbnypAkQBAkQBAkQBAkQBAkQBAkQBAkQDL3sb////////////////Y5ddIlQxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB/s1z4+fn////////////+/v+awYNAkQBAkQBAkQBAkQBAkQBAkQBZnif1+Pn////////3+PtbnypAkQBAkQBAkQBAkQBAkQBAkQB3rlH2+Pj////////////+/v+cwoVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQHI3ML////////////////X5dRJlg5AkQBAkQBAkQBAkQBAkQBfoS/4+fr////////5+fxiozNAkQBAkQBAkQBAkQBAkQBEkwXA17n////////////////Z5tdHlQtAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCBs2D4+vv////////////+/v/i6+TV5NLV5NLV5NLV5NLV5NLf6t7////////////////g6t/V5NLV5NLV5NLV5NLV5NLg6uH9/v7////////////+/v+dw4ZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQHJ3ML////////////////////////////////////////////////////////////////////////////////////////////////////////a5tpIlQxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCEtmX5+vv////////////////////////////////////////////////////////////////////////////////////////////////+/v+cw4ZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgLM3sb////////////////////8/P/6+v/6+v/6+v/8/P/////////////////8/P/6+v/6+v/6+v/7+//////////////////////b59pIlQxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCHt2j7/Pz////////////9/f+RvHhhojNhojNhojOGtmn5+vz////////6+/2Jt25hojNhojNhojNqqEDl7eb////////////+/v+ew4lAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgLN3sj///////////////+lx5JAkQBAkQBAkQBZnif1+Pn////////4+PxcoCtAkQBAkQBAkQB6sFX4+fn////////////a5tlIlQxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCLuW77/P3////////////Z5dhHlQtAkQBAkQBZnif1+Pn////////4+PxcoCtAkQBAkQBAkQDC2Lr////////////+/v+fxIlAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkgPO38r////////////+/v+XwH9AkQBAkQBMlxPZ5dvm7ejm7ejb5txOmBZAkQBAkQBuqUTz9vT////////////b59lJlg1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCMuXD7/P3////////////R4M1DkwRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC81LH////////////+/v+fxIpAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkgPP4Mv////////////7/P2JuGpAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBhojHu8+/////////////b59pJlg1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCOu3L7/P3////////////I28NAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC10Kr////////////+/v+fxItAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkwXR4c7////////////3+Pl3rlJAkQBAkQBAkQBAkQBAkQBAkQBWnCHo8On////////////c6NtLlxBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCSvXj8/f7///////////+/17ZAkQBAkQBAkQBAkQBAkQBAkQCvzaD////////////+/v+gxYxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkwXT4tD////////////w9PJnpjpAkQBAkQBAkQBAkQBPmRjj7OL////////////c6NtLlxBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCUvnv9/v7///////////+20alAkQBAkQBAkQBAkQCoyZb////////////+/v+hxY5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkwfU49L////////////o7+pYnSNAkQBAkQBJlg7d6d3////////////f6d5MlxFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCYwIH+/v////////////+ry5tAkQBAkQCgxIz+/v/////////+/v+hxY1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFkwjW5NP////////////h6uFNlxRFkwjY5dX////////////e6d5MlhJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCawYL+/v/////////+/v+nyJadw4j+/v/////////+/v+jxpBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGlArX5Nb////////////2+ff2+fb////////////d6d1MlxFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCdw4n+/v////////////////////////////+ixo9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBHlQva5tf////////////////////////f6d5MlhJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCfw4r+/v////////////////////+kxpFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNGkA5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBHlQva5tj////////////////f6d9MlxNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBIkQ9SjyJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCixo/+/v////////////+kx5JAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBRjSNmi0lAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJlg7b59n////////g6eBOmBVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBni0p/h3lCkQVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCkx5H+/v////+myJRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkQZ/h3kAAABii0FAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBKlg/c6Nzf6eBPmBZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBii0IAAAAAAACEhoNSjiJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQClyJOnyJVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBRjiCFhoMAAAAAAAAAAACCh31RjiFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBKlg5Klg9AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBSjiKBh30AAAAAAAAAAAAAAAAAAACEhoNii0FCkQVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkQZii0KEhoMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB/h3lmi0pSjiJHkA5BkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQNHkA5SjiJmi0p/h3kAAAAAAAAAAAAAAAAAAAD4AAAAAB8AAOAAAAAABwAAwAAAAAADAACAAAAAAAEAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAgAAAAAABAADAAAAAAAMAAOAAAAAABwAA+AAAAAAfAAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -1
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return -2
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)

DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)


Return hIcon
}

Create_BefundImport_ico()                     	{                                                 	;-- Icon für das InfoWindow
VarSetCapacity(B64, 788 << !!A_IsUnicode)
B64 := "iVBORw0KGgoAAAANSUhEUgAAABQAAAAYCAYAAAD6S912AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAACtQAAArUBrcsL2AAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAHLSURBVDiNzZM9ixpRFIaPcU68gkHRhAUhBAbtLLYbG6spbWysFMRmy2kt/QE2tmkE21TWUZRAWjsbBZlK0WSsLgyjszPvFlmXbHb8yhrIA7c559znnvsVIiLQFVGIiIQQVK/XqVqtXiwAQK1Wi4bDIUkpf8V6vR5SqRQajQYuwXEcVCoVqKqKfr+P/W5hmiYmkwnS6TQMw4Dv+ydllmWhUCggn89jvV7DNM3nQgCYz+dQVRW1Wg2u6x6UzWYzZLNZlMtl2LYNAMFCAFgul8jlciiVSnAc54VsMBggkUjAMAx4nvcUPygEgM1mA03ToOs6pJRP8U6ng1gshm63+2Kho0IAkFJC13VomgbLstBsNpFMJjEajQKP4aQQAGzbRrFYRDweRyaTwXQ6Daw7WwgAu90O7XYblmUdrLlIeC574Zu/+l9HuLpQISIaj8e0WCxeJVqtVkREFFIUZRGJRN6+vjei7Xa7u4bnGaEjuRtmXgUlXNe9IaIfQbn//5b/zbP5gw9CiHee570/NImZP4XD4ZjjOJKIfp7q8N73/Tsi+nakke+e5zXOanlPNBr9yMyfmfmemfE4fGb+IoRQL5L9DjPfMvPXx3F7qv4BASqDqpfUXLIAAAAASUVORK5CYII="
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

Create_PDF_ico(NewHandle := False)     	{                                               	;-- PDF Icon für Befundfenster
VarSetCapacity(B64, 1192 << !!A_IsUnicode)
B64 := "AAABAAEAEBAAAAEAGABoAwAAFgAAACgAAAAQAAAAIAAAAAEAGAAAAAAAAAAAAAkAAAAJAAAAAAAAAAAAAAAFBUAKCoEEBDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKCoAHB2ENDacCAhsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBDANDa0QEMsKCoYBAQYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAhoKCnoQENAKCoAEBDIBAQcAAAAAAAAAAAAAAAAAAAAAAAADAyYFBUcCAh4AAAAAAAAAAAIHB2APD8cMDKENDa8ICGsGBkkDAyoBAQ0CAhoJCXANDa0KCoUJCXsAAAAAAAAAAAACAiAPD8EBAQ4DAyIGBkoJCXANDa0PD8YPD8UQENEPD8EPD74GBlAAAAAAAAAAAAAAAAALC44EBDUAAAAAAAAAAAAHB2APD74GBkoFBT4EBC0CAhoAAAAAAAAAAAAAAAAAAAAHB1gHB2EAAAAAAAAHB1gNDbUDAyoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBDANDaEAAAMGBk4NDbADAyYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQwPD8AICGcPD7sDAygAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALC44QEM0EBDkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQoNDawICGcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBkcQENEHB1wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQMDKENDasHB1cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAh0PD74JCXkFBUQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMHB1sJCXEBAQ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf/wAAD/8AAAf/AACB+AAAwAAAAOAAAADzgQAA8x8AAPA/AADwfwAA+P8AAPH/AADx/wAA4f8AAOH/AADh/wAA"
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

Create_Image_ico(NewHandle := False) 	{                                              	;-- Bilder Icon für Befundfenster
VarSetCapacity(B64, 1124 << !!A_IsUnicode)
B64 := "AAABAAEAEA8AAAEAGAA0AwAAFgAAACgAAAAQAAAAHgAAAAEAGAAAAAAAAAAAAAgAAAAIAAAAAAAAAAAAAABNTU0REREREREREREEBAQAAAAAAAAAAAAAAAAAAAALCwsRERERERERERERERFQUFBPT08AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACmpqapqakAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABERET////6+vofHx8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEBDd3d3///////+VlZUAAAAAAAAAAAAAAAAAAAAAAAAAAACSkpKampoKCgoEBASxsbH////////////39/caGhoAAAAAAAAAAAAAAAAAAABycnL////////f39/W1tb///////////////////+Tk5MAAAAAAAAAAAAAAABHR0f8/Pz////////////////////////////////////5+fkfHx8AAAAAAAAfHx/r6+v///////////////////////////////////////////+ampoAAAAJCQnOzs7////////////////d3d3a2tr////////////////////////9/f1paWm0tLT////////////5+flSUlIAAAAAAABGRkb19fX///////////////////////////////////////+bm5sAAAAAAAAAAAAAAAB8fHz///////////////////////////////////////9zc3MAAAAAAAAAAAAAAABMTEz///////////////////////////////////////+2trYAAAAAAAAAAAAAAACRkZH///////////////////////////////////////////+YmJgXFxcNDQ11dXX8/Pz///////////////////////////////////////////////////////////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------

;{20. Das Ende - OnExit Sprunglabel und OnError-Funktion
DasEnde(ExitReason, ExitCode) {

	global oPV

	OnExit("DasEnde", 0)

  ; Anzeigestatus des Infofenster sichern
	If Addendum.AddendumGui {
		admGui_SaveStatus()
		admGui_Destroy()
	}

 ; Drucken Steuerelement in geöffnetem AU Dialog wieder aktivieren, falls es inaktiviert wurde
	If WinExist("Muster 1a ahk_class #32770")
		WinSetStyle("Muster 1a ahk_class #32770", "Button5", 0x8000000, "off")

 ; Rename Dialog und zugehöriges Sumatra PDF Reader Fenster schliessen
	If (oPV.Previewer = "Sumatra") {
		oPV := Sumatra_Close(oPV.hwnd)
		oPV := ""
	}
	If Addendum.Flags.RenameGui {
		Gui, RN: 		Destroy
		Gui, RNP: 	Destroy
		Gui, adm2: 	Destroy
		OnMessage(0x5, "")
	}

  ; Hooks beenden
	For i, hook in Addendum.Hooks
		UnhookWinEvent(hook.hEvH, hook.HPA)

	FormatTime	, Time      	, % A_Now    	, dd.MM.yyyy HH:mm:ss
	FormatTime	, TimeIdle	, % A_TimeIdle	, HH:mm:ss
	FileAppend	, % Time ", " A_ScriptName ", " A_ComputerName ", " ExitReason ", " ExitCode ", " TimeIdle "`n"
						, % Addendum.LogPath "\OnExit-" A_YYYY "-Protokoll.txt", UTF-8

	Progress	, % "B2 zH25 w600 WM400 WS500"
					.  	 " 	cW" 	Addendum.Default.BgColor1
					. 	 " 	cB" 	Addendum.Default.PRGColor
					. 	 " 	cT" 	Addendum.Default.FntColor
					, % "Addendum wird beendet."
					, % "Addendum für AlbisOnWindows"
					, % "Version vom " DatumVom
					, % Addendum.Default.Font

	Loop 50 {
		Progress % (100 - (A_Index * 2))
		Sleep 1
	}

	Progress, Off

	Gdip_Shutdown(pToken)


ExitApp
}

;}

;{21. #Include(s)

;------------ Addendumbibliotheken für Albis on Windows --------
#Include %A_ScriptDir%\..\..\include\Addendum_Ini.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Autonaming.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Device.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gui.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_GVU.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_LanCom.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_MicroDicom.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Mining.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfReader.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PopUpMenu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Telegram.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

;------------ allgemeine Bibliotheken ---------------------------------
#Include %A_ScriptDir%\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_CtlColors.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#Include %A_ScriptDir%\..\..\lib\class_ImageButton.ahk
#Include %A_ScriptDir%\..\..\lib\class_IPHelper.ahk
;~ #Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_Loaderbar.ahk
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#Include %A_ScriptDir%\..\..\lib\class_OD_Colors.ahk
#Include %A_ScriptDir%\..\..\lib\class_TelegramBot.ahk
#Include %A_ScriptDir%\..\..\lib\class_socket.ahk
#Include %A_ScriptDir%\..\..\lib\Explorer_Get.ahk
#Include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#include %A_ScriptDir%\..\..\lib\MenuSearch.ahk
#Include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
#Include %A_ScriptDir%\..\..\lib\Sci.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\WatchFolder.ahk


;------------ Adddendum.ahk zusätzliche Funktionen /Labels -----
#include %A_ScriptDir%\include\ClientLabels.ahk

;------------ HOTSTRINGS ---------------------------------------------
;#Include %A_ScriptDir%\include\AutoDiagnosen.ahk

;}



; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . .                                                                                       	ADDENDUM HAUPTSKRIPT
global                                                                               AddendumVersion:= "1.46" , DatumVom:= "19.11.2020"
; . . . . . . . . . .
; . . . . . . . . . .                                    ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . . . . . .                                           BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . . . . . .                                                        RUNS ONLY WITH AUTOHOTKEY_H IN 32 OR 64 BIT UNICODE VERSION
; . . . . . . . . . .                                                                 THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .                    !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !!
; . . . . . . . . . .                    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; . . . . . . . . . .                                                                    THIS SCRIPT ONLY WORKS WITH AUTOHOTKEY_H V1.1.32.00
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

/*               	A DIARY OF CHANGES
Beispiel:| **08.12.2018** | **F+** | IPC - Inter Process Communication - zwischen dem Addendum und Praxomat Skript steht, das Addendumskript kann somit weitere Funktionen übernehmen (V0.98b) |
| **31.12.2020** | **M+**	| **Auto-ICD**               	- 	Diagnosen eingeben ohne Albissuche. V1.47 |
| **19.11.2020** | **M+**	| **Addendum_DBASE**   	- 	native DBASE Klasse für Albis .dbf Files, geeignet für: V1.46<br>
																						-	Analyse der von Albis verwendeten DBASE-Datei Strukturen <br>
																						-	Portierung von Daten <br>
																						-	Suche nach Daten |
| **19.11.2020** | **M+**	| **Befundmanagement** 	- 	alle Befunde eines Patienten (pdf, Bilder) in einen Ordner exportieren oder <br>
																					 	in allen PDF-Dateien eines Patienten nach einem Begriff suchen lassen und die Dateien und Fundstellen anzeigen lassen <br>
																						hierfür ist unter Umständen zunächst ein Texterkennungsvorgang notwendig, der allerdings automatisch durchgeführt wird <br>
| **19.11.2020** | **F+**	| **Infofenster**             	- 	Beim Umbenennen einer PDF Datei wird zur Erleichterung die 1.Seite als Vorschau eingeblendet<br>
																     				-	**automatische Dateibenennung**: Namen von Patienten werden erkannt, vorausgesetzt Addendum hat sie in der Datenbank  |
| **11.11.2020** | **F~**	| **Addendum_Albis**    	- 	**AlbisRezeptHelferGui()**: prüft auf Vorhandensein des Rezeptfenster vor Erstellung der Gui <br> |




 ** Registry Pfade für Albis
	Computer\HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\
	Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows
		Installationspfad
		LocalPath-1

	Ifap
	Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{3BA8167C-E0D3-4020-AFAF-BD0AF2BCB984}\1.1\0\win32
	Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{3BA8167C-E0D3-4020-AFAF-BD0AF2BCB984}\1.1\HELPDIR


** Ideen
	OrderEntry Labor - automatische Eintragung der richtigen Ausnahmekennziffer
	Laborabruf - weitermachen , wichtig bei Timergesteuertem Abruf ein Protokoll schreiben

*/

;{01. Skriptablaufeinstellungen / Vorbereitungen

		#Persistent
		#InstallKeybdHook
		#SingleInstance               	, Force
		#MaxThreads                  	, 100
		#MaxThreadsBuffer       	, On
		#MaxHotkeysPerInterval	, 99000000
		#HotkeyInterval              	, 99000000
		#MaxThreadsPerHotkey 	, 12
		#KeyHistory                  	, 0
		#MaxMem                    	, 128

		;#Warn ;, All

		ListLines                        	, Off

		SetTitleMatchMode        	, 2	              	;Fast is default
		SetTitleMatchMode        	, Fast        		;Fast is default

		CoordMode, Mouse         	, Screen
		CoordMode, Pixel            	, Screen
		CoordMode, ToolTip       	, Screen
		CoordMode, Caret        	, Screen
		CoordMode, Menu        	, Screen

		SetKeyDelay                  	, 30, 40
		SetBatchLines                	, -1
		SetWinDelay                    	, -1
		SetControlDelay            	, -1
		SendMode                    	, Input
		AutoTrim                       	, On
		FileEncoding                 	, UTF-8

		DetectHiddenWindows   	, On

	; Schutz gegen Doppelstart , das Skript mit einer zufälligen Verzögerung ausgeführt wird
		Random, slTime, 1, 5
		Sleep % slTime * 400
		If WinExist("Addendum Message Gui ahk_class AutohotkeyGUI")
			ExitApp
		DetectHiddenWindows, Off	;Off is default

	; Scriptende und -fehlerbehandlungsfunktionen festlegen
		OnExit("DasEnde")
		OnError("FehlerProtokoll")

	; Client Namen feststellen
		global compname := StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft

	; Tray Icon Leichen entfernen
		NoTrayOrphansWin10()

	; auf hohe Priorität setzen aufgrund der Hook-Prozesse
		Process, Priority,, High

	; startet die Windows Gdip Funktion
		If !(pToken:=Gdip_Startup()) {
				MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
				ExitApp
		}

	; DPI-Skalierung ausschalten
		Result := DllCall("User32\SetProcessDpiAwarenessContext", "UInt" , -1)

	; Tray Icon erstellen
		If (hIconAddendum := Create_Addendum_ico())
			Menu, Tray, Icon, % "hIcon: " hIconAddendum

	; wichtige Hotkeys sollten an dieser Stelle bleiben
		Hotkey, $^!ä   	, BereiteDichVor                                                         	;= Überall: beendet das Addendumskript
		Hotkey, $^!n   	, SkriptReload                                                            	;= Überall: Addendum.ahk wird neu gestartet

;}

;{02. Variablendefinitonen , individuelles Client-Traymenu wird hier erstelltc

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; allgemeine globale Variablen
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global AlbisWinID      	:= AlbisWinID()                      	; für alle Funktionen muss die ID des Albis Hauptfenster ohne weitere Aufrufe als Variable vorhanden sein
		global AlbisPID          	:= AlbisPID()                         	; mal sehen vielleicht benötige ich die Process ID von Albis für irgendwas
		global Albis               	:= Object()                           	; enthält aktuelle Fenstergrößen und andere Daten aus den Albisfenstern
		global AddendumDir                                                 	; Skriptverzeichnis
		global AutoDocCalled	:= 0                                      	; flag für AutoDoc Funktion welche nicht mehrfach gestartet werden kann (liest mehrfach aus der Ini aber nie richtig)
		global GVU                	:= Object()                           	; enthält Daten für die Vorsorgeautomatisierung
		;global ScanPool        	:= Object()                            	; enthält die Dateinnamen aller im BefundOrdner vorhandenen Dateien
		global ScanPool        	:= Array()                              	; enthält die Dateinnamen aller im BefundOrdner vorhandenen Dateien
		global Script                	:= Object()
		global JSortDesc, JSortCol                                           	; Sortierung des Journals im Addendum Infofenster
		global PageSum        	:= 0                                     	; Gesamtseitenzahl aller Pdf-Dateien im BefundOrdner
		global DatumZuvor
		global q                     	:= Chr(0x22)                           	; ist dieses Zeichen -> "
		global admServer                                                       	; für Interskriptkommuniktion zwischen verschiedenen Computern im LAN

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		Script.Gui         	:= Object()
		Script.User       	:= {"Interaction" : false, "Interface" : ""}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; globale und andere Variablen für die Hook's
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global EHProc		    		:= ""                                  	; der Name des Prozesses der gehookt wurde
		global EHStack := Array()

		global SHookHwnd	    	:= 0                                   	; hook handle des ShellHookProc
		global HmvHookHwnd  	:= 0                                   	; hook handle des für das Heilmittelverordnungsfensters (Physio)
		global HmvEvent		    	:= 0                                 	; event Nummer für Heilmittelverordnungsfenster

		global ifaptimer               	:= 0                                    	; ifap produziert eine Kaskade an Fehlermeldungen, flag wird bei erster Fehlermeldung gesetzt
		global AlbisHookTitle                                                 	; enthält den aktuellen Fenstertitel des Albisfenster
		global EHWHStatus        	:= false                               	; flag falls gerade noch der Hookhandler läuft
		global ClickReady            	:= 1                             		; für den GVU Formularausfüller
		global AutoLogin            	:= true                                 ; der Autologin Vorgang für Albis kann per Interskriptkommunikation ausgestellt werden
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; globale Variablen für die Addendum Patientendatenbank
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global oPat                    	:= Object()                       	; Patientendatenbank Objekt für Addendum
		global TProtokoll				:= Array()                         	; alternatives Tagesprotokoll als Absicherung falls es wieder Probleme mit der Albisdatenbank gibt
		global SectionDate					                                	; [Section]-Bezeichnung in der Tagesprotokoll Datenbank

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; globale Variablen (Objekte) und andere Variablen für Hotstringfunktionen
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		;global activeControl   	:= Object()                              	; enthält den Inhalt des aktuell fokusierten Controls bevor manuelle Eingaben erfolgen  #### verwende ich das noch irgendwo?

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; ** EINSTELLUNGEN AUS DER ADDENDUM.INI EINLESEN / ADDENDUM OBJEKT MIT DATEN BEFÜLLEN **
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

	/*                    	Beschreibung aller enthaltenen key : value - Paare

	             ~~~~~~~ Struktur des Addendum-Objektes ~~~~~~~
		Addendum.BefundOrdner  	= Scan-Ordner für neue Befundzugänge
		Addendum.InfoWindow     	= Posteingang Einblendung im Albis Patientenfenster
	                                                     .Y,.W,.H     	- der Einblendungsbereich
	                                                	 .LVScanPool - Objekt enthält die größe der ScanPool Listview (W,H)
		Addendum.Telegram          	= enumeriertes Objekt (Bsp: Addendum.Telegram[1].TBotID)
	                                                	 .BotName 	- Name des Bot's
			                                           	 .Token         - Telegram Bot ID
		                                               	 .ChatID     	- Telegram Chat ID
		Addendum.GVUListe          	= Anzahl der vorhandenen Untersuchungen in der GVU-Liste

	 */

		global Addendum	:= Object()
		global SPData      	:= Object()

		Addendum.Debug := false                                   	; gibt Daten in meine Standardausgabe (Scite) aus wenn = true

	;	ein Aufruf nur mit dem Verzeichnisnamen initialisiert den Ini-Pfad - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		workini := IniReadExt(AddendumDir "\Addendum.ini")
	;	- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	; Einstellungen - Funktionen aus \include\Addendum_Ini.ahk - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		admObjekte()                                                    	; Definition Subobjekte + diverse Variablen
		admVerzeichnisse()                                               	; Programm- und Datenverzeichnisse
		admAlbisMenuAufrufe()                                      	; Albis Menubezeichnungen und wm_command Befehle
		admDruckerStandards()            	                         	; Standarddrucker A4/PDF/FAX - Einstellungen werden vom PopUpMenu Skript benötigt
		admFensterPositionen()                                       	; verschiedene Fensterpositionen
		admFunktionen()                                               	; zu- und abschaltbare Funktionen
		admGesundheitsvorsorge()                                	; Abstände zwischen Untersuchungen minimales Alter
		admInfoWindowSettings()                                    	; Infofenster Einstellungen
		admLaborDaten()                                              	; Laborabruf, Verarbeitung Laborwerte
		admLanProperties()                                             	; LAN - Kommunikation Addendumskripte
		admMailAndHolidays()                                       	; Mailadressen, Urlaubszeiten
		admMonitorInfo()                                                	; ermittelt die derzeitige Anzahl der Monitore und deren Größen
		admPIDHandles()                     	                        	; Prozess-ID's
		admPDFSettings()					                            	; Pdf Verarbeitung
		admPraxisDaten()                                               	; Praxisdaten   -   für Outlook, Telegram Nachrichtenhandling, Sprechstundenzeiten, Urlaub
		admShutDown()                                                	; Einstellungen für automatisches Herunterfahren des PC
		admSicherheit()                                                   	; Lockdown Funktion - Nutzer hat den Arbeitsplatz verlassen
		admStandard()                                                    	; Standard-Einstellungen für Gui's
		admTagesProtokoll()                                           	; Tagesprotokoll laden
		admTelegramBots()				                            	; Telegram Bot Daten
		admThreads()                                                    	; weiteren Skriptcode laden und wenn eingestellt sofortiges starten als Thread

		ChronikerListe()                                                    	; Chroniker Index
		GeriatrischListe()                                                   	; Geriatrie Index

		GeriatrieICD := [	"Ataktischer Gang {R26.0G};"    	, "Spastischer Gang {R26.1G};" 	, "Probleme beim Gehen {R26.2G};"	, "Bettlägerigkeit {R26.3G};"              	, "Immobilität {R26.3G};"
								, 	"Standunsicherheit {R26.8G};"   	, "Multifaktoriell bedingte Mobilitätsstörung {R26.8G};"                        	, "Ataxie {R27.0G};"      			    		, "Koordinationsstörung {R27.8G};"
								,	"Hemineglect {R29.5G};"         		, "Sturzneigung a.n.k. {R29.6G};" 	, "Harninkontinenz {R32G};"          	, "Überlaufinkontinenz {N39.41G};"   	, "n.n. bz. Harninkontinenz {N39.48G};"
								,	"Stuhlinkontinenz {R15G};"
								,	"Dysphagie {R13.9G};"
								, 	"Dysphagie mit Beaufsichtigungspflicht während der Nahrungsaufnahme {R13.0G};", "Dysphagie bei absaugpflichtigem Tracheostoma mit geblockter Trachealkanüle {R13.1G};"
								,	"Orientierungsstörung {R41.0G};"	, "Gedächtnisstörung {R41.3G};"	, "Vertigo {R42G};"                        	, "chron. Schmerzsyndrom {R52.2G};"	, "chron. unbeeinflußbarer Schmerz {R51.1G};"
								,	"Chronische Schmerzstörung mit somatischen und psychischen Faktoren {F45.41G};"
								,	"Multiinfarktdemenz {F01.1G};"    	, "Subkortikale vaskuläre Demenz {F01.2G};"                                  		, "Gemischte kortikale und subkortikale vaskuläre Demenz {F01.3G};"
								,	"Vaskuläre Demenz {F01.9G};"    	, "Dementia senilis {F03G};"
								,	"Vorliegen eines Pflegegrades {Z74.9G};"
								,	"Demenz bei Alzheimer-Krankheit mit frühem Beginn (Typ 2) [F00.0*] {G30.0+G};Alzheimer-Krankheit, früher Beginn {G30.0G};"
								,	"Demenz bei Alzheimer-Krankheit mit spätem Beginn (Typ 1) [F00.1*] {G30.1+G};Alzheimer-Krankheit, später Beginn {G30.1G};"
								,	"Primäres Parkinson-Syndrom mit mäßiger Beeinträchtigung ohne Wirkungsfluktuation {G20.10G};"	, "Primäres Parkinson-Syndrom mit mäßiger Beeinträchtigung mit Wirkungsfluktuation {G20.11G};"
								, 	"Primäres Parkinson-Syndrom mit schwerster Beeinträchtigung ohne Wirkungsfluktuation {G20.20G};"	, "Primäres Parkinson-Syndrom mit schwerster Beeinträchtigung mit Wirkungsfluktuation {G20.21G};"]

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Tray Menu erstellen
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		func_NoNo                     	:= Func("NoNo")
		func_ZeigePatDB             	:= Func("ZeigePatDB")
		func_ZeigeFehlerProtokoll	:= Func("ZeigeFehlerProtokoll")

		Menu, Tray, NoStandard
		Menu, Tray, Tip					, % StrReplace(A_ScriptName, ".ahk") " V." AddendumVersion " vom " DatumVom
													. "`nClient: " compName SubStr("                   ", 1, Floor((20 - StrLen(compName))/2))
													. "`nPID: " DllCall("GetCurrentProcessId")
													. "`nAutohotkey.exe: " A_AhkVersion

		Menu, Tray, Color            	, % "c" Addendum.DefaultBGColor3
		Menu, Tray, Add				, Addendum, % func_NoNo
		Menu, Tray, Icon				, Addendum, % "hIcon: " hIconAddendum

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Module laden
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		modul 	:= Object()
		tool   	:= Object()

	; Clientabhängige Authorisierung laden
		CompAuth 	:= IniReadExt(compname, "Module")
		If InStr(CompAuth, "All") 	{

			Loop
			  If InStr(IniReadExt("Module", "Modul" SubStr("0" A_Index, -1)), "ERROR")
				break
			else
				CompAuth .= A_Index ","

			CompAuth := RTrim(CompAuth, ",")

		}

	; authorisierte Module ins Traymenu laden
		Loop {

			Modultemp:= IniReadExt("Module", "Modul" SubStr("0" A_Index, -1))
			Tooltemp	 := IniReadExt("Module", "Tool" 	 SubStr("0" A_Index, -1))

			If !InStr(ModulTemp, "ERROR") {

				tmp              	:= StrSplit(Modultemp, "|")
				iconpath       	:= tmp[4]
				modul[tmp[2]]	:= tmp[3]

				If !RegExMatch(iconpath, "i)\s*[A-Z]\:\\")
					iconpath := Addendum.AddendumDir "\" iconpath

				If InStr(tmp[1], "NoAuth") {
					Menu, SubMenu1, Add, % tmp[2], ModulStarter
					If FileExist(iconpath)
						Menu, SubMenu1, Icon, % tmp[2], % iconpath
				}
				else if A_Index in %CompAuth%
				{
					Menu, SubMenu1, Add, % tmp[2], ModulStarter
					If FileExist(iconpath)
						Menu, SubMenu1, Icon, % tmp[2], % iconpath
				}

			}

			If !InStr(ToolTemp, "ERROR")	{

				tmp          	:=StrSplit(StrReplace(ToolTemp, "Ã¼", "ü"), "|")
				iconpath   	:= tmp[4]
				tool[tmp[2]]	:= tmp[3]

				If !RegExMatch(iconpath, "i)\s*[A-Z]\:\\")
					iconpath := Addendum.AddendumDir "\" iconpath

				If InStr(tmp[1], "NoAuth") {
					Menu, SubMenu2, Add, % tmp[2], ToolStarter
					If FileExist(iconpath)
						Menu, SubMenu2, Icon, % tmp[2], % iconpath
				}
				else if A_Index in %CompAuth%
				{
					Menu, SubMenu2, Add, % tmp[2], ToolStarter
					If FileExist(iconpath)
						Menu, SubMenu2, Icon, % tmp[2], % iconpath
				}

			}

			If InStr(ModulTemp, "ERROR") && InStr(ToolTemp, "ERROR")
					break
		}

		Menu, SubMenu2, Add	, % "Windows komplett neu starten", ToolStarter
		Menu, SubMenu2, Icon	, % "Windows komplett neu starten", % Addendum.AddendumDir "\assets\ico\Windows.ico"
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; per Checkbox zu- oder abschaltbare Funktionen (SubMenu3)
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

		Menu, SubMenu3, Add, % "Pause Addendum"	, Menu_PauseAddendum

	; AutoSizer für Ifap, MS Word, FoxitReader, Sumatra PDF          	;{

		For classnn, appname in Addendum.Windows.Proc {
			Menu, SubMenu31, Add, % appname, Menu_AutoSize
			Menu, SubMenu31, Icon, % appname, % A_ScriptDir "\res\" appname ".ico"
			If (Addendum.Windows[appname].AutoPos & Addendum.MonSize)
				Menu, SubMenu31, Check	  	, % appname
			else
				Menu, SubMenu31, UnCheck	, % appname
		}
		Menu, SubMenu3, Add, % "AutoSizer", :SubMenu31

	;}

	; automatische Anpassung von Albis an die Monitorauflösung 	;{
		Menu, SubMenu3, Add, % "Albis AutoSize"     	, Menu_AlbisAutoPosition
		If Addendum.AlbisLocationChange
			Menu, SubMenu3, Check	  , % "Albis AutoSize"
		else
			Menu, SubMenu3, UnCheck, % "Albis AutoSize"
		;}

	; Automatisierung der GVU Formulare                                      	;{
		Menu, SubMenu3, Add, % "Albis GVU automatisieren", Menu_GVUAutomation
		If Addendum.GVUAutomation
			Menu, SubMenu3, Check	  , % "Albis GVU automatisieren"
		else
			Menu, SubMenu3, UnCheck, % "Albis GVU automatisieren"
		;}

	; Automatisierung der PDF Signierung mit dem FoxitReader        	;{
		Menu, SubMenu3, Add, FoxitReader Signaturhilfe                            	, Menu_PDFSignatureAutomation
		Menu, SubMenu3, Add, FoxitReader Dokument automatisch schliessen, Menu_PDFSignatureAutoTabClose

		If Addendum.AutoCloseFoxitTab
			Menu, SubMenu3, Check	  , FoxitReader Dokument automatisch schliessen
		else
			Menu, SubMenu3, UnCheck, FoxitReader Dokument automatisch schliessen

		If Addendum.PDFSignieren {
			Menu, SubMenu3, Check	, FoxitReader Signaturhilfe
			Menu, SubMenu3, Enable, FoxitReader Dokument automatisch schliessen
		}
		else {
			Menu, SubMenu3, UnCheck	, FoxitReader Signaturhilfe
			Menu, SubMenu3, Disable  	, FoxitReader Dokument automatisch schliessen
		}

		;}

	; Laborabruf automatisieren                                                   	;{

		Menu, SubMenu3, Add, % "Laborabruf automatisieren", Menu_LaborAbrufManuell
		If Addendum.Labor.AutoAbruf
			Menu, SubMenu3, Check	  , % "Laborabruf automatisieren"
		else
			Menu, SubMenu3, UnCheck, % "Laborabruf automatisieren"

		; zeitgesteuert
		Menu, SubMenu3, Add, % "zeitgesteuerter Laborabruf", Menu_LaborAbrufTimer
    	If Addendum.Labor.AbrufTimer {
			Addendum.Labor.AutoAbruf := true                                       ; Timer ist an dann auch  AutoAbruf
			Menu, SubMenu3, Check	  , % "zeitgesteuerter Laborabruf"
		} else
			Menu, SubMenu3, UnCheck, % "zeitgesteuerter Laborabruf"

	;}

	; Addendum Toolbar                                                             	;{

		Menu, SubMenu3, Add, % "Addendum Toolbar", Menu_AddendumToolbar
		If Addendum.ToolbarThread
			Menu, SubMenu3, Check	  , % "Addendum Toolbar"
		else
			Menu, SubMenu3, UnCheck, % "Addendum Toolbar"
	;}

	; AddendumGui - Infofenster                                                    	;{
		Menu, SubMenu3, Add, % "Addendum Infofenster", Menu_AddendumGui
		If Addendum.AddendumGui
			Menu, SubMenu3, Check	  , % "Addendum Infofenster"
		else
			Menu, SubMenu3, UnCheck, % "Addendum Infofenster"
	  ;}

	; Schnellrezept                                                                       	;{
		Menu, SubMenu3, Add, % "Schnellrezept", Menu_Schnellrezept
		If Addendum.Schnellrezept
			Menu, SubMenu3, Check	  , % "Schnellrezept"
		else
			Menu, SubMenu3, UnCheck, % "Schnellrezept"
	;}

	; MicroDicom Dateiexport automatisieren                                 	;{
		Menu, SubMenu3, Add, % "MicroDicom Export", Menu_MDicom
		Menu, SubMenu3, % (Addendum.mDicom ? "Check":"UnCheck"), % "MicroDicom Export"
	;}

	; PopUpMenu                                                                          	;{
		Menu, SubMenu3, Add, % "PopUpMenu", Menu_PopUpMenu
		Menu, SubMenu3, % (Addendum.PopUpMenu ? "Check":"UnCheck"), % "PopUpMenu"
	;}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; weitere Menupunkte
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; PROTOKOLLE anzeigen
		Menu, SubMenu4, Add, % "Patienten Datenbank anzeigen"	, % func_ZeigePatDB

		If FileExist(protokoll := Addendum.BefundOrdner "\PDfImportLog.txt") {
			func_PDFImportLog  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "PDF Importprotokoll"            	, % func_PDFImportLog
		}

		If FileExist(protokoll := Addendum.DBPath "\OCRTime_Log.txt") {
			func_tessOCRLog  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "PDF OCR Protokoll"            	, % func_tessOCRLog
		}

		Menu, SubMenu4, Add, % "aktuelles Fehlerprotokoll"        	, % func_ZeigeFehlerProtokoll

		If FileExist(protokoll := Addendum.AddendumDir "\logs'n'data\OnExit-Protokoll.txt") {
			func_ZeigeOnExitProtokoll := Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "OnExit Protokoll"                	, % func_ZeigeOnExitProtokoll
		}

	; FARBEN festlegen
		Menu, SubMenu1, Color                            	, % "c" Addendum.DefaultBGColor3
		Menu, SubMenu2, Color                            	, % "c" Addendum.DefaultBGColor3
		Menu, SubMenu3, Color                            	, % "c" Addendum.DefaultBGColor3
		Menu, SubMenu4, Color                            	, % "c" Addendum.DefaultBGColor3

	; MENU erstellen
		Menu, Tray, Add, Module starten/beenden   	, :SubMenu1
		Menu, Tray, Add, Daten / Protokolle             	, :SubMenu4
		Menu, Tray, Add, andere Tools                   	, :SubMenu2
		Menu, Tray, Add, Einstellungen                     	, :SubMenu3

		Menu, Tray, Icon, Module starten/beenden 	, % Addendum.AddendumDir "\assets\ModulIcons\Module.ico"
		Menu, Tray, Icon, Einstellungen                   	, imageres.dll, 52

		Menu, Tray, Add,
		Menu, Tray, Add, Zeige Skript Variablen       	, scriptVars
		Menu, Tray, Add, Skript Neu Starten             	, SkriptReload
		Menu, Tray, Add, Beenden                            	, DasEnde

	;}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Vorgang zur Ersterstellung des Addendum Datenbankverzeichnis dessen Speicherort frei gewählt werden kann
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		;die Default Einstellung beim ersten Start liegt im Addendum-Hauptverzeichnis \logs'n'data\_DB
		;ACHTUNG: Beenden Sie Addendum vor dem Kopieren des Verzeichnisses! Haben Sie das Verzeichnis verschoben, starten Sie Addendum.ahk.
		;                  	das Skript wird ein Ordnerauswahlfenster öffnen. Navigieren Sie zum neuen Verzeichnispfad. Der neue Pfad wird in der Addendum.ini gespeichert!

	; ist der Datenbankpfad noch nicht angelegt, wird es jetzt gemacht
		If InStr(Addendum.DBPath, "Error")	                                                	{

				Addendum.DBPath := AddendumDir "\logs'n'data\_DB"
				IniWrite, % "%AddendumDir%\logs'n'data\_DB", % AddendumDir "\Addendum.ini", % "Addendum", % "AddendumDBPath"
				Addendum.FirstDBAcess := true

		}
		If !InStr(FileExist(Addendum.DBPath), "D")    	&& (Addendum.FirstDBAcess)	{

				FileCreateDir, % Addendum.DBPath
				If ErrorLevel {
						DBPath:= Addendum.DBPath
						MsgBox,, Addendum für AlbisOnWindows,
						(LTrim
						Das Verzeichnis für die Datenbank:
						%DBPath%
						konnte nicht angelegt werden.
						Das Skript muss jetzt beendet werden!
						)

						ExitApp
				}

		}
		else if !InStr(FileExist(Addendum.DBPath), "D") && (!Addendum.FirstDBAcess) {

				;dieser Abschnitt ermittelt den neuen Pfad zur Addendumdatenbank
					ADBtmp:= StrReplace(Addendum.DBPath, "\_DB")
					SelectDBPath:
					FileSelectFolder, ADBtmp, *ADBtmp,, Wählen Sie den neuen Pfad zur Addendum Datenbank (..\_DB)
					If !ErrorLevel		{

							ADBtmp:= RegExReplace(ADBtmp, "\\$")
							If !InStr(AddendumDBPath, "_DB")		{

									MsgBox,, Addendum für AlbisOnWindows, % "Ihre Pfadangabe ist nicht richtig!`nWählen Sie den Ordner mit folgender Bezeichnung "  q "_DB" q "!"
									goto SelectDBPath
							}
							else if !FileExist(ADBtmp "\Patienten.txt") && !(FirstDBAcess=1)	{

									MsgBox, 4, Addendum für AlbisOnWindows,
									(LTrim
									Ihr ausgewählter Ordner enthält keine `"Patienten.txt`" Datei!
									Soll dies wirklich der Addendum Datenbankordner sein?
									)
									IfMsgBox, Yes
									{
											Addendum.DBPath:= ADBtmp
											IniWrite, % Addendum.DBPath, % AddendumDir "\Addendum.ini", % "Addendum", % "AddendumDBPath"

									}
							}
							else	{

									Addendum.DBPath:= ADBtmp
									IniWrite, % Addendum.DBPath, % AddendumDir "\Addendum.ini", % "Addendum", % "AddendumDBPath"

							}

							VarSetCapacity(ADBtmp, 0)
					}
		}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Einlesen der Patientendaten aus der Patienten.txt Datei im Addendum Datenbankordner
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		If FileExist(Addendum.DBPath "\Patienten.txt")
			oPat := ReadPatientDatabase(Addendum.DBPath) ; oPat - ist ein Objekt das die aktuell Daten zu registriertem Patientienten enthät
	;}

;}


;{04. Timer, Threads, Initiliasierung WinEventHooks, OnMessage, TCP-Kommunikation, eigene Dialoge anzeigen

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Starten clientunabhängiger Timer Labels
		SetTimer, UserAway, 300000                                                                     	; zum Ausführen von Berechnungen wenn ein Computer unbenutzt ist (5min)
		If IsLabel(nopopup_%compname%)
			SetTimer, nopopup_%compname%, 2000                                              	; spezielle Labels die Funktionen nur auf einzelnen Arbeitsplätzen ausführen
		If IsLabel(seldom_%compname%)
			SetTimer, seldom_%compname%, 900000		                                        	; Ausführung alle 15min

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Skriptkommunikation / Anpassen von Fenstern bei Änderung der Bildschirmaufläsung
		OnMessage(0x4A	, "Receive_WM_COPYDATA")                                         	; Interskriptkommunikation
		OnMessage(0x7E	, "WM_DISPLAYCHANGE")                                             	; Änderung der Bildschirmauflösung

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  TCP Server starten - für LAN Kommunikation
		;~ global admServer
		;~ admStartServer()

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;	SetWinEventHooks /Shellhook initialisieren
		InitializeWinEventHooks()

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; kurzzeitige Anzeige von Informationen bei Skriptneustart
		;~ PraxTT( 	oPat.Count() " Patienten sind in der Addendum-Datenbank bekannt.`n"
					;~ . 	(TProtokoll.MaxIndex()="" ? 0 : (TProtokoll.MaxIndex() = 1 ? "1 Patient ist im Tagesprotokoll." : TProtokoll.MaxIndex() " Patienten sind im Tagesprotokoll.") "`n")
					;~ . 	"Signaturen: " Addendum.SignatureCount ", Seitenanzahl: " Addendum.SignaturePages "`n"
					;~ . 	"laufende Hooks: " Addendum.Hooks.Count(), "2 0" )

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Infofenster anzeigen wenn eine Akte/Karteikarteikarte geöffnet ist
		Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If InStr(Addendum.AktiveAnzeige, "Patientenakte")         	{
			PatDb(AlbisTitle2Data(AlbisTitle), "exist")
			If Addendum.AddendumGui
				If RegExMatch(Addendum.AktiveAnzeige, "i)Karteikarte|Laborblatt|Biometriedaten")
					AddendumGui()
		}

;}

;{05. Hotkey command's

HotkeyLabel:
	;=--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Hotkey's die überall funktionieren
	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	Hotkey, $^Ins               	, RunSomething                                 	;= Überall: für's programmieren gedacht, im Label das Skript eintragen das per Hotkey gestartet werden soll
	Hotkey, !m                    	, MenuSearch                                    	;= Überall: Menu Suche - funktioniert in allen Standart Windows Programmen
	Hotkey, !q                     	, ScreenShot                                      	;= Überall: erstellt einen Screenshot per Mausauswahl und legt diesen im Clipboard ab
	Hotkey, CapsLock          	, DoNothing                                      	;= Überall: Capslock kann nicht mehr versehentlich gedrückt werden
	Hotkey, ^!e                  	, EnableAllControls                            	;= Überall: inaktivierte Bedienfelder in externen Fenstern aktivieren
	Hotkey, ^!p                  	, PatientenkarteiOeffnen                       	;= Überall: eine selektierte Zahl wird zum Öffnen einer Patientenkarteikarte genutzt
	Hotkey, Pause               	, AddendumPausieren
	Hotkey, ^!F10               	, SendClipBoardText                           	;= Überall: sendet den Inhalt des Clipboards als simulierte Tasteneingabe (um z.B. wiederholt ein Passwort zu senden)

	;Hotkey, ~LButton         	, CaptureDoubleClick                        	;= Überall: wenn Text unter Maus einen Patientennamen oder eine ID enthält wird die Akte dazu geöffnet

	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Scite4AutoHotkey
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	Hotkey, IfWinActive, ahk_class SciTEWindow
    Hotkey, ^NumPad1                  	, SendRawBAuf                      	;= SciTE:  '; {'
    Hotkey, ^NumPad2                  	, SendRawBzu                        	;= SciTE:  '; }'
    Hotkey, ^Numpad3                  	, SendRawElse                       	;= SciTE:  } else {
    Hotkey, Numpad6                     	, SendStrgRight                      	;= SciTE:  Strg+Pfeil rechts
    Hotkey, Numpad4                     	, SendStrgLeft                        	;= SciTE:  Strg+Pfeil links
    Hotkey, LControl & 5                 	, SendRawProzentAll              	;= SciTE: %%, setzt den Cursor zwischen beide Zeichen
    Hotkey, LControl & `.                	, SciteBracket                         	;= SciTE: 'LCtrl + .' bringt '{' und nach nochmaligem drücken '}'
    Hotkey, LControl & `,                	, SciteCommentBlock             	;= SciTE: 'LCtrl + ,' bringt '/*' und nach nochmaligem drücken '*/'
    Hotkey, LControl & `´                	, SendRawN                          	;= SciTE: ein Linefeed (`n)
    Hotkey, LAlt & `´                       	, SendRawR                           	;= SciTE: carriage return (`r)
    Hotkey, ^!`´                            	, SendRawTab                       	;= SciTE: Tab (`t)
    Hotkey, LAlt & Right                   	, SendFourSpace                   	;= SciTE: sends 4 times space
    Hotkey, LAlt & Left                        	, SendFourBackspace             	;= SciTE: sends 4 times backspace
    Hotkey, LAlt & v                         	, SciteDescriptionBlock, I2      	;= SciTe: Abschnittbezeichner für Addendumskripte
    Hotkey, LControl & Numpad4    	, SciteSmartCaretL                 	;= SciTe: for language depend moving of caret
    Hotkey, LControl & Numpad6    	, SciteSmartCaretR                 	;= SciTe: for language depend moving of caret
    ;Hotkey, ^+h                            	, SciteHotKeyWriter                	;= SciTe: opens a gui for writing hotkey's because I ever forget the names
	Hotkey, IfWinActive

	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; andere Fenster
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	Hotkey, IfWinActive, ScanSnap Manager - Bild erfassen und Datei speichern ahk_class #32770
	Hotkey, Enter                  	, ScanBeenden
	Hotkey, IfWinActive

	func_GetPatNameFromExplorer := Func("GetPatNameFromExplorer")
	Hotkey, IfWinActive, % "ahk_class CabinetWClass"
	Hotkey, ~F6                    	, % func_GetPatNameFromExplorer
	Hotkey, ~^RButton          	, % func_GetPatNameFromExplorer
	Hotkey, IfWinActive
	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Albis
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	fn_VhK     	:= Func("AlbisAusfuellhilfe").Bind("Muster12", "Hotkey")
	fn_Undo 	:= Func("UnRedo").Bind("Undo")
	fn_Redo 	:= Func("UnRedo").Bind("Redo")
	Hotkey, IfWinActive, ahk_class OptoAppClass			            			;~~ ALBIS Hotkey - bei aktiviertem Fenster
	Hotkey, $^c			    		, AlbisKopieren					        			;= Albis: selektierten Text ins Clipboard kopieren
	Hotkey, $^x			    		, AlbisAusschneiden			        		   	;= Albis: selektierten Text ausschneiden
	Hotkey, $^v			    		, AlbisEinfuegen                                 	;= Albis: Clipboard Inhalt (nur Text) in ein editierbares Albisfeld einfügen
	Hotkey, $^a		        		, AlbisMarkieren                                	;= Albis: gesamten Textinhalt markieren
	Hotkey, ^F9 		    		, Abrechnungshelfer			        			;= Albis:Abrechnungshelfer.ahk starten
	Hotkey, ^F11		    		, Schulzettel	         			        			;= Albis: Schulzettel.ahk starten
	Hotkey, ^F12		    		, Hausbesuche         			        			;= Albis: Hausbesuche.ahk starten
	Hotkey, !F5		        		, AlbisHeuteDatum			             		;= Albis: Einstellen des Programmdatum auf den aktuellen Tag
	Hotkey, ^F7                 	, GVUListe_ProgDatum                       	;= Albis: akt. Patienten in die GVU-Liste eintragen (Datum der Programmeinstellung)
	Hotkey,  !F7                  	, GVUListe_ZeilenDatum                       	;= Albis: akt. Patienten in die GVU-Liste eintragen (Zeilendatum wird genutzt )
	;Hotkey, ^!F1					, AlbisFunktionen									;= Albis: Funktionen für Albis - Auswahl-Gui
	;Hotkey, If, InStr(AlbisGetActiveWindowType(), "Patientenakte")        	;~~ ALBIS Hotkey - nur auslösbar wenn eine Akte geöffnet ist
	Hotkey, !F8		    			, DruckeLabor                                    	;= Albis: Laborblatt als pdf oder auf Papier drucken
	Hotkey, !Down      			, AlbisKarteiKarteSchliessen        			;= Albis: per Kombination die aktuelle Karteikarte schliessen
	Hotkey, !Up               		, AlbisNaechsteKarteiKarte	        			;= Albis: per Kombination eine geöffnete Kartekarte vorwärts
	Hotkey, !Right    				, Laborblatt    										;= Albis: Laborblatt des Patienten anzeigen
	Hotkey, !Left  					, Karteikarte      									;= Albis: Karteikarte des Patienten anzeigen
	Hotkey, $^z	                 	, % fn_Undo                                      	;= Albis: Eingaben rückgängig machen in Steuerelementen (Edit, RichEdit)
	Hotkey, $^d	                 	, % fn_Redo                                       	;= Albis: Eingaben wiederherstellen
	Hotkey, $^!d                	, StarteAutoDiagnosen                        	;= Albis: Skript AutoDiagnosen3 starten
	Hotkey, $^!e                	, exportiereAlles                                  	;= Albis: exportiert alle Befunde des Patienten
	Hotkey, IfWinActive

	Hotkey, IfWinExist, ahk_class OptoAppClass                              		;~~ mehrmonatiges Kalender-Addon
	Hotkey, $F9			    		, MinusEineWoche                                	;= Albis: addiert eine Woche einem Datum hinzu, z.B. im AU Formular
	Hotkey, $F10		        	, PlusEineWoche                                	;= Albis: subtrahiert eine Woche von einem Datum, z.B. im AU Formular
	Hotkey, $F11		    		, MinusEinMonat                                	;= Albis: addiert 4x7 Tage (4 Wochen) einem Datum hinzu, z.B. im AU Formular
	Hotkey, $F12		        	, PlusEinMonat                                    	;= Albis: subtrahiert 4x7 Tage (4 Wochen) von einem Datum, z.B. im AU Formular
	Hotkey, $+F3					, Kalender	                                        	;= Albis: Ersatz für den Shift+F3 Kalender von Albis der nur einen Monat zeigt
	Hotkey, IfWinExist

	Hotkey, IfWinActive, Dauermedikamente ahk_class #32770    		;~~ Sortieren im Dauermedikamentenfenster
	Hotkey, #				    		, DauermedBezeichner		        			;= Albis: zum Kategorisieren der Medikamente (übersichtlichere Darstellung, besser als mit dem BMP)
	Hotkey, ^Up                	, DauermedRauf                                   	;= Albis: schiebt die ausgewählte Zeile um eine Zeile nach oben
	Hotkey, ^Down              	, DauermedRunter                         		;= Albis: schiebt die ausgewählte Zeile um eine Zeile nach unten
	Hotkey, IfWinActive

	Hotkey, IfWinActive, Dauerdiagnosen von ahk_class #32770         	;~~ Sortieren im Dauerdiagnosenfenster
	Hotkey, ^Up                	, DiagnoseRauf                                    	;= Albis: schiebt eine Diagnose höher
	Hotkey, ^Down              	, DiagnoseRunter                                 	;= Albis: zieht eine Diagnose eine Zeile tiefer
	Hotkey, IfWinActive

	Hotkey, IfWinActive, Cave! von ahk_class #32770                         	;~~ Sortieren im Cave! von Dialog
    Hotkey, ^Up                	, CaveVonRauf                                    	;= Albis: schiebt eine Zeile höher
    Hotkey, ^Down              	, CaveVonRunter                                  	;= Albis: schiebt eine Zeile tiefer
	Hotkey, IfWinActive

	Hotkey, IfWinActive, Muster 12a ahk_class #32770                       	;~~ Ausfüllhilfe für das Formular Verordnung häusliche Krankenpflege
	Hotkey, Up                    	, % fn_Vhk                                         	;= Albis: löscht ein Datumsfeld
    Hotkey, Down               	, % fn_Vhk                                         	;= Albis: übernimmt vorhandene Datumseinträge
	Hotkey, IfWinActive

	;~ HotKey, IfWinExist, LbAutoComplete                                            	;~~ Spezielle Listbox erreichbar durch ''+''
	;~ Hotkey, Enter, AutoCBLBres                                                        	;= Albis: übernimmt die Auswahl
	;~ Hotkey, IfWinExist

	Hotkey, IfWinActive, Dokumentationsbögen ahk_class #32770   	;~~ Korrekturhilfe bei falsch gesetzer Auswahl bei Hautkrebsverdacht
	Hotkey, Numpad0		    	, FastSets                                             	;= Albis: ruft Dokumentationsbogen auf, setzt Auswahl um,speichert Dokument u. rückt e. Patienten nach unten
	Hotkey, IfWinActive


AlbisAusfuellhilfe(Formular, cmd) {

	/* AlbisAusfuellhilfe - Beschreibung

			geplant als universell einsetzbare Funktion für schnelleres Ausfüllen von editierbaren Feldern in Albis

	*/

		static Formulare := {"Muster12" : {"from"	: [6,12,17,22,27,32,37,43,52,59,65,71,78,83]
														,	"to"   	: [7,13,18,23,28,33,38,44,53,60,66,72,79,84]}}
		static fcLast, awLast, wt, entryNr

		hk 	:= A_ThisHotkey
		aw	:= WinExist("A")
		fc  	:= GetFocusedControlClassNN(aw)
		sf		:= true                                                      	; flag bleibt true wenn der Eingabefokus noch im gleichen Steuerelement ist (sf = static focus)

	; Zurücksetzen von Variablen wenn ein anderes Formular bearbeitet wird
		If !(aw = awLast) {
			wt 		:= WinGetTitle(aw)
			awLast 	:= aw
			fcLast	:= ""
		}

	; anderer Eingabefokus
		If !(fcLast = fc) {
			sf      	:= false
			fcLast 	:= fc
		}

	; fokussiertes Steuerelement ist nicht in der Ausfuellhilfsliste (Formulare) dann zurück
		If !(ctrlLabel := ControlFilter(Formulare[Formular], fc))
			return

	switch Formular
	{
		case "Muster12":

			If (hk = "UP") {



			}

		; ermittelt alle schon vorhanden Datumsangaben in den "vom" und "bis" Feldern
			If !sf {

				Formulare["Muster12"].Dates := Array()
				Loop, % Formulare["Muster12"]["from"].Count() {

					from := Formulare["Muster12"]["from"](A_Index)
					to 	:= Formulare["Muster12"]["to"](A_Index)
					date := Array()
					ControlGetText, date1, % "Edit" from	, % "ahk_id " aw
					ControlGetText, date2, % "Edit" to   	, % "ahk_id " aw
					date[1] := date1, date[2] := date2

					Loop, 2
						If !(RegExMatch(date[A_Index], "\d{2}\.\d{2}\.\d{2,4}") || RegExMatch(date[A_Index], "\d{6,8}"))
							date[A_Index] := ""

					If (StrLen((Dates := date.1 "," date.2)) = 1)
						continue

					If !DatesInArr(Dates, Formulare["Muster12"].Dates)
						Formulare["Muster12"].Dates.Push(Dates)

					}
				}


			SciTEOutput(Formulare["Muster12"]["fromDates"])
			SciTEOutput(Formulare["Muster12"]["toDates"])

		}


}
;{ Hilfsfunktionen Ausfuellhilfe
ControlFilter(CtrlObj, classnn) { 	; Hilfsfunktion für AlbisAusfuellhilfe - gibt Steuerelementbezeichnung zurück

	; überprüft im Moment nur Edit Steuerelement

	For ctrlLabel, class in CtrlObj
		Loop, % class.Count()
			If (classnn = "Edit" class[A_Index])
				return ctrlLabel

return false
}

DatesInArr(Dates, Arr) {	; Hilfsfunktion für AlbisAusfuellhilfe

	For idx, val in Arr
		If (val = Dates)
			return true

return false
}

IndexInArr(s, Arr) {	; Hilfsfunktion für AlbisAusfuellhilfe

	For idx, val in Arr
		If (val = s)
			return idx

return 0
}

;}

	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; flexible Hotkey's und Fensterkombinationen, je nach Einstellung in der Addendum.ini
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	If Addendum.PDFSignieren 	{ ; Signatur setzen im Foxit Reader nach Hotkey

		func_FRSignature := func("FoxitReaderSignieren")
		Hotkey, IfWinActive, % "ahk_class " Addendum.PDFReaderWinClass
		Hotkey, % Addendum.PDFSignieren_Kuerzel, % func_FRSignature
		Hotkey, IfWinActive

	}

	GetPatNameFromExplorer() {                                                             	;-- ermittelt in einer pdf Dateibezeichnung den Patientennamen und öffnet die Patientenakte
	; Voraussetzung: der Patient ist in der Addendum Patientendatenbank gelistet, ansonsten passiert nichts
		SplitPath, % Explorer_GetSelected(WinExist("A")),,,, fname
		FuzzyKarteikarte(fname)
	}

	;}

return

;}

;{06. Hotstrings

:*:#at::@

;{ 	6.1. -- ALBIS
; --- Leistungskomplexe                                                                  	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "lk")
:*:bz::32025                                                                             	;{ Blutzuckermessung
:*:glucose::32025
:*:zucker::32025                        ;}
:*:gs::                                                                                         	;{ Gesprächziffer
	AlbisSchreibeLkMitFaktor("03230", "2")
return ;}
:*:ps1::                                                                                       	;{ Psychosomatikziffer 1
:X*:psy1::
	AlbisSchreibeLkMitFaktor("35100", "2")
return ;}
:*:ps2::                                                                                       	;{ Psychosomatikziffer 2
:X*:psy2::
	AlbisSchreibeLkMitFaktor("35110", "2")
return ;}
:*:tel::                                                                                          	;{ Telefongebühr mit dem Krankenhaus
	AlbisSchreibeLkMitFaktor("80230", "2")
return ;}
:*:fax::                                                                                         	;{ Faxgebühr
	AlbisSchreibeLkMitFaktor("40120", "2")
return ;}
:*:gvu::                                                                                      	;{ GVU/HKS inkl. eHKS-Formular

	GVUDatum := AlbisLeseZeilenDatum(300, false)
	Sleep 300
	SendMode, Input
	SetKeyDelay, -1, -1
	Send, % "{Raw}01732-01746"
	Send, % "{TAB}"
	Sleep 300
	ProgrammDatum := AlbisLeseProgrammDatum()
	AlbisSetzeProgrammDatum(GVUDatum)
	Sleep 200
	AlbisFormular("eHKS_ND")
	Sleep 200
	AlbisHautkrebsScreening("opb", true, true )
	Sleep 200
	AlbisSetzeProgrammDatum(ProgrammDatum)
	Sleep 200
	Albismenu(32778, "Cave! von")
	Sleep 300
	ControlSend, SysListView321, {Escape}	, Cave! von ahk_class #32770
	Sleep 50
	ControlSend, SysListView321, 9            	, Cave! von ahk_class #32770
	Sleep 50
	ControlSend, SysListView321, {Space}	, Cave! von ahk_class #32770
	Sleep 50
	ControlSend, SysListView321, {Tab 2}	, Cave! von ahk_class #32770
	Sleep 50
	ControlSend, SysListView321, {Right}		, Cave! von ahk_class #32770


return ;}
:*:kur::                                                                                        	;{ Kurplan, Gutachten, Stellungnahme
:X*:gut::
	AlbisSchreibeLkMitFaktor("01622", "")
return ;}
:*:stix::                                                                                        	;{ Urinstix
:X*:urin::
:X*:ustix::
:X*:urinstix::
	AlbisSchreibeLkMitFaktor("32030", "")
return ;}
:*:inr::                                                                                         	;{ INR Messstreifen
:*:gerinnung::
	AlbisSchreibeLkMitFaktor("32026", "")
return ;}
:*:Kompr::                                                                                   	;{ Kompressionsverband
	AlbisSchreibeLkMitFaktor("02313", "")
return ;}
:*:Verwe::                                                                                    	;{ Verweilen außerhalb der Praxis
	AlbisSchreibeLkMitFaktor("01440", "2")
return ;}
:*:hb0::                                                                                       	;{ Hausbesuch - alle Ziffern
	Send, % "{Text}01410-97234-03230(x:3)"
	;result := BetterBox("Addendum - Hausbesuche", "Bitte die Uhrzeit eingeben", "", -1)

	;MsgBox, % "result: " result
	;Send, % "{Text}01410-97234-03230(x:3)"
	;HBdatum :=  InputBoxEx("Bitte Datum und Uhrzeit des Hausbesuch eingeben")
	;RegExMatch(HBdatum, "(?<Tag>\d{1,2}\.\d{1,2})\.*(?<Jahr>\d{2,4})*\s*\,*\s*(?<Uhrzeit>\d{0,2}\:*\d{0,2})", "Rx"
	;SciTEOutput(HBdatum)
return
:*:hb1::
	Send, % "{Text}01411(um:19:00)-97234-03230(x:3)"
return
:*:hb2::
:*:hbWE::
:*:hbFei::
:*:hbAbe::
:*:hbSpre::
	Send, % "{Text}01412(um:19:00)-97234-03230(x:3)"
return
;}
:*:c1::                                                                                          	;{ chronisch krank Kennzeichnung - automatisierte Überprüfung
	SendInput, % "{RAW}03220"
	InChronicList()
return
:*:c2::03221
InChronicList() {

		GruppenName := "Chroniker"

		PatID := AlbisAktuellePatID()

		For key, ChronikerID in Addendum.Chroniker
			If (PatID = ChronikerID)
				return

	; Nutzer abfragen ob Pat. aufgenommen werden soll
		hinweis := "Pat: " oPat[PatID].Nn ", " oPat[PatID].Vn ", geb. am: " oPc[PatID].Gd "`n"
		hinweis .= "ist nicht als Chroniker vermerkt.`nMöchten Sie automatisch alle Eintragungen`ninnerhalb von Albis vornehmen lassen?"

		MsgBox, 0x1024, Addendum für Albis on Windows, % hinweis, 10
		IfMsgBox, Yes
		If (PatID = AlbisAktuellePatID())
		{
				ChronCb	:= false
				Indikation	:= false
				PatGruppe:= false

			; Ziffer der Liste hinzufügen und in die Datei speichern
				Addendum.Chroniker.Push(PatID)
				FileAppend, % PatID "`n", % Addendum.DBPath "\DB_Chroniker.txt"

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

	; PraxTT("Pat. konnte keiner Gruppe zugeordnet werden.`nSetze Funktion mit anderen Eintragungen fort!", "3 3")
	;PraxTT("Pat. konnte keiner Gruppe zugeordnet werden.`nSetze Funktion mit anderen Eintragungen fort!", "3 3")
}

;}
:*:sang::                                                                                      	;{ Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:Stuhl::
:*:iFop::
:*:haemof::
	AlbisSchreibeLkMitFaktor("01737", "")
return ;}
:*:UUnt::                                                                                      	;{ Vorsorgeuntersuchung Kinder inkl. J1
:*:Kind::
	Send, % "{Text}01711-01712-01713-01714-01715-01716-01717-01718-01719-01723"
return ;}
:*:haus::                                                                                      	;{ Hausbesuchsziffern umfangreicher

	;Dringender Besuch wegen der Erkrankung, unverzüglich nach Bestellung ausgeführt zwischen 19:00 und 22:00 Uhr,
	; oder an Samstagen, Sonntagen und gesetzlichen Feiertagen, am 24.12. und 31.12. zwischen 07:00 und 19:00 Uhr

	HBKomplexe := { 	"K1"      	: ["01410", "01411", "01412", "01413", "01414", "01415", "01416", "01418"]            	; Hausbesuchsziffern
			    			, 	"K2"      	: ["97234", "97235", "97236", "97237", "97238", "97239"]            	; Wegegebühren
							, 	"K3"      	: ["03370", "03372"]                                                                    	; Palliativziffern
					    	,	"Zusatz" 	: {"01415" :	["37102", "37113", "03370", "03372", "03373", "37306", "01102", "03030"]
												  ,	"01410" :  ["03370", "03372(x:2)", "37305", "01102", "01205", "01210", "01207"]
												  ,	"01411" :  ["03370", "03373", "37306", "01102", "03030"]
												  ,	"01412" :  ["03370", "03373", "37306", "01102", "03030"]
												  ,	"01413" :  ["03370", "03372(x:2)", "37305", "01102"]
												  ,	"01418" :  ["03370", "03372", "03373", "37305", "01102", "01207"]}
				    		,	"Regel"  	: {"01413" : 	"-Wege"}
							,	"Texte"		: {"01410" :	 	"Besuch eines Kranken, wegen der Erkrankung ausgeführt."
												,   "01411" :	 	"Dringender Besuch wegen der Erkrankung, unverzüglich nach Bestellung ausgeführt."
												,   "01412" :		"Dringender Besuch wegen der Erkrankung, unverzüglich nach Bestellung ausgeführt"
														   		   . 	" zwischen 22 und 7 Uhr oder an Samstagen, Sonntagen und gesetzlichen Feiertagen"
														    	   . 	", am 24.12. und 31.12., zwischen 19 und 7 Uhr oder bei Unterbrechen der"
																   . 	" Sprechstundentätigkeit mit Verlassen der Praxisräume"
												,   "01413" :	 	"Besuch eines weiteren Kranken in derselben sozialen Gemeinschaft (z.B. Familie) "
																   . 	"und/oder in beschützenden Wohnheimen bzw. Einrichtungen bzw. Pflege- oder Altenheimen mit Pflegepersonal"
												,	"03372" :	 	"Zuschlag zu den Gebührenordnungspositionen 01410 oder 01413 für die palliativmedizinische Betreuung in der Häuslichkeit"
																   .	"`t- mehrfach abrechenbar je vollendete 15min"
												,	"03373" :	 	"Zuschlag zu den Gebührenordnungspositionen 01411, 01412 oder 01415 für die palliativmedizinische Betreuung in der Häuslichkeit"
																   .	"`t- je Besuch nur einem abrechenbar"
												,	"01102" :	 	"Inanspruchnahme des Arztes nach Vorbestellung an Samstagen zwischen 7 und 14 Uhr, 10,93 Euro"
												,	"01205" :		"Notfallpauschale im organisierten Not(-fall)dienst und für nicht an der vertragsärztlichen Versorgung teilnehmende Ärzte,"
																   .	" Institute und Krankenhäuser für die Abklärung der Behandlungsnotwendigkeit bei Inanspruchnahme`n"
																   .	"`t- zwischen 07:00 und 19:00 Uhr (außer an Samstagen, Sonntagen, gesetzlichen Feiertagen und am 24.12. und 31.12.)"
												,	"01207" :	  	"Notfallpauschale im organisierten Not(-fall)dienst und für nicht an der vertragsärztlichen Versorgung"
																   .  	" teilnehmende Ärzte, Institute und Krankenhäuser für die Abklärung der Behandlungsnotwendigkeit bei Inanspruchnahme"
																   .  	"`t- zwischen 19:00 und 07:00 Uhr des Folgetages`n`t- ganztägig an Samstagen, Sonntagen, gesetzlichen Feiertagen"
																   .  	" und am 24.12. und 31.12."}}


return ;}
:*:term::                                                                                       	;{ Zuschläge Terminvermittlung (Terminvermittlung Facharzt, TSS-Terminvermittlung)
:*:Vermit::
	Send, % "{Raw}03008-03010"
return ;}
:*:opvor::                                                                                  	;{ OP-Vorbereitung
:*:op vor::
:*:op-vor::
	Send, % "{Raw}31010-31011-31012-31013"
return ;}
:*R:pall::03370-03371-03372-03373                                          	; 	Palliativziffern
:*R:postop::31600                                                                        	; 	postoperative Nachbehandlung (nur bei Überweisung vom Facharzt)
:*R:aneu::01747                                                                          	; 	Aufklärungsgespräch Ultraschall-Screening Bauchaortenaneurysmen - nur Männer ab 65 Jahren einmalig berechnungsfähig!
:*R:entw::03350-03351                                                               	; 	Orientierende entwicklungsneurologische Untersuchung eines Neugeborenen, Säuglings, Kleinkindes oder Kindes
:*:DAK::93550-93555-93560-93565-93570                             	;	HAP DAK
:*R:1-::03000-03040                                                                   	; 	03000-03040
:*R:2-::03220                                                                              	;	03220
:*R:3-::03221                                                                            	;	03221
#If

BetterBox(Title="", Prompt="", Default="", Pos=-1, ReadDate=true) {               	;--	custom input box allows to choose the position of the text insertion point

	;-------------------------------------------------------------------------------
    ; custom input box allows to choose the position of the text insertion point
    ; return the entered text
    ;
    ; Title is the title for the GUI
    ; Prompt is the text to display
    ; Default is the text initially shown in the edit control
    ; Pos is the position of the text insertion point
    ;   Pos =  0  => at the start of the string
    ;   Pos =  1  => after the first character of the string
    ;   Pos = -1  => at the end of the string
    ;---------------------------------------------------------------------------

	global
    static BBResult ; used as a GUI control variable
	x:= A_CaretX - 20, y:= A_CaretY + 30
	If ReadDate
		Default := AlbisLeseZeilenDatum(40, false) ", "

    ; create GUI
    Gui, BetterBox: New, -DPIScale -MinimizeBox +HWNDhBBox +LastFound	;, % Title
    Gui, BetterBox: Margin, 30, 18
    Gui, BetterBox: Add, Text,, % Prompt
    Gui, BetterBox: Add, Edit, w290 vBBResult, % Default
    Gui, BetterBox: Add, Button, x80 w80 Default gBBOk, &OK
    Gui, BetterBox: Add, Button, x+m wp gBbCancel, &Cancel

    ; main loop
    Gui, BetterBox: Show, % "x" x " y" y , % title
    SendMessage, 0xB1, %Pos%, %Pos%, Edit1, A ; EM_SETSEL
	WinWaitClose, % title,, 15
    ;~ while WinExist(title)
		;~ Sleep, 10

return BBResult

    ;-----------------------------------
    ; event handlers
	;-----------------------------------
BBOK: ; "OK" button
    Gui, BetterBox: Submit ; get Result from GUI
    Gui, BetterBox: Destroy
Return

BBCancel: ; "Cancel" button
BetterBoxGuiClose:     ; {Alt+F4} pressed, [X] clicked
BetterBoxGuiEscape:    ; {Esc} pressed
        BBResult := "BetterBoxCancel"
        Gui, BetterBox: Destroy
Return
}
InputBoxEx(prompt, Default="", ReadDate=true) {

	ibX := A_CaretX - 15, ibY := A_CaretY + 30

	If ReadDate
		Default := AlbisLeseZeilenDatum(40, false) ", "

	#IfWinExist, Addendum InputBoxEx ahk_class #32770
	Hotkey, $Escape, ibXClose
	#IfWinExist

	SetTimer, ibXMoveToEnd, -50

	If !WinExist("Addendum InputBoxEx ahk_class #32770")
		InputBox, result, Addendum InputBoxEx, % prompt,, 350, 130, % ibX, % ibY, Locale,, % Default

	Hotkey, $Escape, Off

	;Send, % "{Text}01410-97234-03230(x:3)"
	;Hausbesuchskomplexe()
return result

ibXMoveToEnd:
	hwnd := WinExist("Addendum InputBoxEx ahk_class #32770")
	WinSet, Style 	, 0x94040A4C	, % "ahk_id " hwnd
	WinSet, ExStyle	, 0x00010200	, % "ahk_id " hwnd
	WinMove, % "ahk_id " hwnd,,, % ibY + 4,, 100
	ControlFocus	, Edit1              	, % "ahk_id " hwnd
	ControlSend	, Edit1, {End}   	, % "ahk_id " hwnd
return

ibXClose:
	hwnd := WinExist("Addendum InputBoxEx ahk_class #32770")
	ControlFocus	, Edit1              	, % "ahk_id " hwnd
	VerifiedSetText("Edit1", "", hwnd)
	VerifiedClick("Button2", hwnd)
return
}
ChooseFromList(arr) {

	static hCFL

	Script.User.Interaction	:= true
	Script.User.Interface   	:= A_ThisFunc

	hfCtrl := GetHex(GetFocusedControl())

	Gui, CFL: New, -SysMenu -Caption +AlwaysonTop +ToolWindow +HWNDhCFL ;0x98200000
	Gui, CFL: Margin	, 0, 0
	Gui, CFL: Color, c172842
	Gui, CFL: Add, ListView, % "x2 y2 w300 r" arr.Count() " CheckBox vLvCFL "


}

;}
; --- Kuerzelfeld                                                                                	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("identify"), "Edit3")
:*:GVU::                                                                                        	;{ GVU,HKS Automatisierung

	ControlGetText, EditDatum, Edit2, ahk_class OptoAppClass
	If !RegExMatch(EditDatum, "\d{2}\.\d{2}\.\d{4}")
	{
			PraxTT("Das Datumfeld konnte nicht ausgelesen werden", "3 3")
			return
	}

	DatumZuvor:= AlbisSetzeProgrammDatum(EditDatum)
	ControlFocus, Edit3, ahk_class OptoAppClass
	VerifiedSetText("Edit3", "GVU", "ahk_class OptoAppClass")
	SendInput, {Tab}

return ;}
:*b0:AISC:: ;{
	SendInput, {Tab}
	Sleep 50
	SendInput, {Tab}
return ;}

#If
;}
; --- Diagnosen                                                                              	;{
; bei Includes jetzt
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "dia")
:*:gb::                 ;{
	Auswahlbox(GeriatrieICD, "Diagnosenliste Geriatrie")
return ;}
#If
;}
; --- Info                                                                                          	;{
#If ( InStr(AlbisGetActiveControl("contraction"), "info") && WinActive("ahk_class OptoAppClass") )
:R*:letztes::letztes Quartal nicht da (keine Chronikerziffer möglich)
#If
;}
; --- aem                                                                                         	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "aem")
:R*:ENife::Einnahmebeschreibung für Nifedipin-Notfalltropfen mitgegeben
:R*:ETram::Einnahmebeschreibung für Tramadol mitgegeben
:R*:EVitD3::Einnahmebeschreibung für Dekristol mitgegeben
:R*:ETry::Einnahmebeschreibung für Tryasol mitgegeben
:R*:EMac::Einnahmebeschreibung für Macrogol mitgegeben
;}
; --- Rezept                                                                                      	;{
#If WinActive("Muster 16 ahk_class #32770") || WinActive("Privatrezept ahk_class #32770")
:*:#Trya::                                           	;{ Tryasol 30 ml
	IfapVerordnung("Tryasol 30 ml")
return ;}

;}
; --- AUTODOC ---                                                                          	;{
#IfWinActive, ahk_class OptoAppClass
:*:+::
	If !InStr(AlbisGetActiveControl("identify"), "Dokument|Edit3") {
		Send, {+}
		return
	}

	If !AutoDocCalled {
			;MsgBox, % "identify`n" AlbisGetActiveControl("identify")
			AutoDocCalled:=1
			AutoDoc()
	}
return
;}
; --- AppStarter                                                                                	;{
:R*:...::
HotstringComboBox(AlbisGetActiveControl("content"))
return
#IfWinActive
;}
; --- Ausnahmeindikationen                                                              	;{
#If IsFocusedAusnahmeIndikation()
:*:#::
	AutoFillAusnahmeindikation()
return
:*:1-::03000-03040-
:*:2-::03220-
:*:3-::03360-03362-
:*:a-::03000-03040-03220-03360-03362

AutoFillAusnahmeindikation() {

		static Chroniker, Geriatrisch
		static toAdd   	:= ["03000", "03040", "03220", "03360", "03362"]
		static toRemove := ["95001", "95002", "95003"]

	; Chroniker und Geriatrie Index werden eingelesen ;{
		If FileExist(Addendum.DbPath . "\DB_Chroniker.txt") && (StrLen(Chroniker) = 0) {
				FileRead, tmpChroniker, % Addendum.DbPath . "\DB_Chroniker.txt"
				Loop, Parse, tmpChroniker, `n, `r
					Chroniker .= Trim(StrSplit(A_LoopField, ";").1) ","
				Chroniker := RTrim(Chroniker, ",")
		}

		If FileExist(Addendum.DbPath . "\DB_DB_Geriatrisch.txt") && (StrLen(Geriatrisch) = 0) {
				FileRead, tmpGeriatrisch, % Addendum.DbPath . "\DB_Geriatrisch.txt"
				Loop, Parse, tmpGeriatrisch, `n, `r
					Geriatrisch .= Trim(StrSplit(A_LoopField, ";").1) ","
				Geriatrisch := RTrim(Geriatrisch, ",")
		}

	;}

	; Steuerelement auslesen
		ControlGetText, ctext, Edit14, ahk_class #32770 ahk_exe albis, Ausnahmeindikationen
		aktPatID := AlbisAktuellePatID()

		 ;~ SciTEOutput("Chroniker: " SubStr(Chroniker, 1, 30), 0, 1)
		 ;~ SciTEOutput("toAdd.MaxIndex(): " toAdd.MaxIndex(), 0, 1)
		 ;~ SciTEOutput("ctext: " ctext, 0, 1)

	; hinzufügende Leistungskomplexe mit Steuerelementinhalt vergleichen
		Loop, % toAdd.MaxIndex()
		{
				If !InStr(ctext, toAdd[A_Index])
				{
						If InStr(toAdd[A_Index], "03220") && InList(aktPatID, Chroniker)
							cText .= "-" toAdd[A_Index]
						else if InList(toAdd[A_Index], "03660,03362") && InList(aktPatID, Geriatrisch)
					    	cText .= "-" toAdd[A_Index]
						else
		     				cText .= "-" toAdd[A_Index]
				}
		}

		cText := LTrim(ctext, "-")
		;~ SciTEOutput("ctext: " ctext, 0, 1)
		;~ SciTEOutput("`n", 0, 1)

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
				If WinExist("ahk_class #32770", "Ausnahmeindikation") {
					ControlClick, Ok, % "ahk_class #32770", % "Ausnahmeindikation"
					WinWaitClose, % "ahk_class #32770", % "Ausnahmeindikation", 2
				}

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

IsFocusedAusnahmeIndikation() {

	; das Fenster "weitere Information" trägt nur den Namen des Patienten
		WinGetText       	, activeWinText	, A
		ControlGetFocus	, activeControl	, A

	If InStr(activeWinText, "Ausnahmeindikation") && InStr(activeControl, "Edit14")
		return true

return false
}
;}
; --- Privatabrechnung                                                                    	;{
#If WinActive("ahk_class OptoAppClass") && (RegExMatch(AlbisGetActiveControl("contraction"), "lp|lbg") || InStr(AlbisGetActiveWindowType(), "Privatabrechnung"))
:*R:Abstri::298-                                                                                    	; Entnahme und gegebenenfalls Aufbereitung von Abstrichmaterial zur mikrobiologischen Untersuchung
:*R:Bera1::1(fak:3,5:Beratung < 10 Minuten)-                                     	; Beratung < 10 Minuten
:*R:Bera2::3(fakt:2,3:Beratung mind. 10 Minuten)-                               	; Beratung mind. 10 Minuten
:*R:Bera3::3(fakt:3,5:Beratung < 20 Minuten)-                                      	; Beratung < 20 Minuten
:*R:Bera4::34(fakt:2,3:Beratung, eingehend mindestens 20 Minuten)-  	; Beratung, eingehend mindestens 20 Minuten
:*R:BA::250-                                                                                       	; Blutabnahme
:*R:BSG::3501-                                                                                    	; BSG
:*R:BZ::                                                                                             	;{ Blutzucker
:*R:Blutzuc::
:*R:glucose::
:*R:glukose::3560-                                                                                ;}
:*R:gleichg::826-                                                                               	; neurologische Gleichgewichtsprüfung
:*R:hb::50(dkm:4)-                                                                               	; Hausbesuch
:*R:impf::375-                                                                                    	; Impfung
:*R:infusion k::                                                                                     	; Infusion < 30 min
:*R:infu1::271-
:*R:infusion l::                                                                                     	; Infusion > 30 min
:*R:infu2::
:*R:infil1::267-                                                                                     	; Medikamentöse Infiltrationsbehandlung, je Sitzung
:*R:infil2::268-                                                                                     	; Medikamentöse Infiltrationsbehandlung im Bereich mehrerer Körperregionen
:*R:inj sc::                                                                                         	; Injection s.c., i.m., i.c.
:*R:inj ic::
:*R:inj im::252-
:*R:inj iv::253-                                                                                    	; Injection i.v.
:*R:inr::3530-                                                                                  		; INR mit Gerät
:*R:lzRR::654                                                                                    		; Langzeitblutdruckmessung
:*R:lufu::605                                                                                    		; Lungenfunktion
:*R:Lungenf::605                                                                              		; Lungenfunktion
:*R:neuro::800-                                                                                    	; eingehende neurologische Untersuchung
:*R:oxy::602-                                                                                    	; Pulsoxymetrie
:*R:Pulsoxy::602-                                                                                  	; Pulsoxymetrie
:*R:rekt::11-                                                                                      	; rektale Untersuchung
:*:sang::                                                                                              	;{ Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:Stuhl::
:*:iFop::
:*:haemof::3500-                                                                                 ;}
:*R:sono::410-420                                                                               	; Ultraschall (Ultraschalluntersuchung von bis zu drei weiteren Organen im Anschluss an Nummern 410 bis 418, je Organ. Organe sind anzugeben.)
:*R:Unters1::7(fakt:2,3)-                                                                    	; Vollständige Untersuchung – ein Organsystem
:*R:Unters2::7(fakt:3,5)-                                                                    	; Vollständige Untersuchung – mehrere Organsysteme
:*R:verb::200-                                                                                      	; Verband
:*R:verweil::56-                                                                                  	; Verweilen außerhalb der Praxis
:*R:zuschl::A-B-C-D-K1                                                                         	; Zuschläge
:*R:lageso::(sach:Landesamt für Gesundheit und Soziales:21.00)
:*R:lagesokurz::(sach:Landesamt für Gesundheit und Soziales:5.00)
:*R:porto1::(sach:Porto Standard:0.80)	; bis 20g
:*R:porto2::(sach:Porto Kompakt:0.95) 	; bis 50g
:*R:porto3::(sach:Porto Groß:1.55)      	; bis 500g
:*R:porto4::(sach:Porto Maxi:2.70)       	; bis 1000g
:*:kopie:: ;{
	AlbisKopiekosten(0.5)
return ;}
:*:reiseunf:: ;{
:*:reiserüc::(sach:Reiseunfähigkeitsbescheinigung:20.00) ;}
:*R:stix::                                                                                            	;{ Urinstix
:*R:streifen::
:*R:urin::
:*R:ustix::
:*R:urinstix::3652-					        ;}
#If
;}
; --- Blockeingaben                                                                         	;{
#If WinActive("ahk_class OptoAppClass") && AlbisGetActiveControl("contraction")
:*:corona::                                                                                     	;{ Corona Ziffern + Diagnose
:*:covid::

	KKT := ["lko|88240-32006", "dia|Bronchitis {J06.9G}; Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G}; COVID-19, Virus nicht nachgewiesen {!U07.2G};"]
	;AlbisSchreibeSequenz(KKT)


return ;}
#If




;}
; --- manuelle Eingabe der Versichertendaten                                   	;{
#If WinActive("Ersatzverfahren ahk_class #32770")
:*:AOKN::                                                                                 	;{ AOK Nordost
	ControlSetText, Edit1, % "72101"        	, % "Ersatzverfahren ahk_class #32770"
	ControlSetText, Edit3, % "109519005"	, % "Ersatzverfahren ahk_class #32770"
	ControlSetText, Edit5, % "AOK Nordost"	, % "Ersatzverfahren ahk_class #32770"
	Sleep 100
	ControlFocus, Edit6, % "Ersatzverfahren ahk_class #32770"

return ;}
#If
;}
;}

;{ 	6.2 -- TELEGRAM
#IfWinActive Telegram ahk_class Qt5QWindowIcon
:C1:HASH::                                                                                								;{
	DA_PatID:= AlbisAktuellePatID()

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

	InputBox, HASHtag, Albis ID Nachfrage, %IB_text%,,500,250,,,,,`#%DA_PatID%

	rautepos:=InStr(HASHtag, "`#")
	mrautepos:= InStr(HASHtag, "`#",1, 2)
	If (rautepos = 0) {
			HASHtag:= "`#" . HASHtag
		} else if (rautepos > 1) or (mrautepos <> 0) {
			StringReplace, HASHtag, HASHtag, `#, , All
			HASHtag:= "`#" . HASHtag
		}

	WinActivate, Telegram ahk_class Qt5QWindowIcon
	WinWaitActive, Telegram ahk_class Qt5QWindowIcon
	Sleep, 150
	SendRaw, % HASHtag
	Send, {Enter}

	SendRaw, % StrReplace(Addendum.hashtagNachricht, "##", "`n")

	VarSetCapacity(bestelltext, 0)
	SendInput, {LControl Down}{Enter}{LControl UP}
return
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
:*:.aid::                                                                                                                     	;{ % "ahk_id "
	SendInput, % "{Raw}% " q "ahk_id " q " "
return ;}
:R*:.as::A_ScriptDir
:*:.as1::                                                                                                                     	;{  A_ScriptDir "\"
	SendInput, % "{Raw}A_ScriptDir\"
	SendInput, {Left}
return ;}
:*:.t1::	                                                                                                                    	;{ oberer einzelner Trenner
SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------"
return ;}
:*:.t2::	                                                                                                                    	;{ mittlerer Trenner
SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------;{"
return ;}
:*:.t3::                                                                                                                      	;{ alle 3
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
:*:.msg::                                                                                                                    	;{ Addendum Messagebox
	Send, % "{Text}MsgBox, 1, Addendum für Albis on Windows, % """""
return ;}
; __ SciteOutPut __
:R*:.sot:: ;{
	SciteWrite("SciteOutPut(qq<clipText>: qq <clipText>, 0, 1)", true)
return ;}
; __ Tasten __
:R*:.lk::{Left}
:R*:.rk::{Right}
:R*:.cdk::{LControl Down}
:R*:.cuk::{LControl Up}
:R*:.cduk:: ;{
	Send, % "{Text}{LControl Down}{LControl Up}"
	Send, {Left 13}
return ;}
:R*:.sdk::{LShift Down}
:R*:.suk::{LShift Up}
:R*:.sduk:: ;{
	Send, % "{Text}{LShift Down}{LShift Up}"
	Send, {Left 11}
return ;}
:R*:.adk::{LAlt Down}
:R*:.auk::{LAlt Up}
:R*:.aduk:: ;{
	Send, % "{Text}{LAlt Down}{LAlt Up}"
	Send, {Left 9}
return ;}
:R*:.adk::{Win Down}
:R*:.auk::{Win Up}
:*:.aduk:: ;{
	Send, % "{Text}{Win Down}{Win Up}"
	Send, {Left 8}
return ;}
:*:.sout::                                                                                                                         ;{ Sciteoutput standard
	SendInput, % "{Text}SciTEOutput(A_LineNumber " q ": " q " )"
	SendInput, % "{Left}"
return ;}

:R*:.afa::Addendum für Albis on Windows
:R*:.addendum::Addendum für Albis on Windows
:R*:.ambx::MsgBox, 1, Addendum für Albis on Windows -

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

;{     	5.a.3.1 -- Bestellungen
::.Bestellungen:: ;{
IniRead, bestelltext, % AddendumDir "\Addendum.ini", Addendum, Bestellung
Loop, Parse, bestelltext, ##
{
	If (StrLen(A_LoopField) > 0)
		SendInput, % "{Raw}" A_LoopField
	else
		SendInput, {Enter}
	Sleep, 25
}

VarSetCapacity(bestelltext, 0)
return ;}
;}

;{     	5.a.3.2 -- sonstige Kürzel
; Programmieren
:*:.Heute::`%A_DD`%`.`%A_MM`%`.`%A_YYYY`%
:*:#ahk::Autohotkey
:*:#lahk::language:Autohotkey
:*:#sahk::site:Autohotkey
; Mail/Briefunterschriften
:*:KVStemp::                                                                     	;{ Unterschrift für Mail oder Briefe an KV
	If Addendum.Praxis.KVStempel
		Send, % "{Text}" Addendum.Praxis.KVStempel
return ;}
:*:MailStemp::                                                                   	;{ Unterschrift für Mail oder Briefe an Firmen, Patienten (ohne BSNR und LANR)
	If Addendum.Praxis.MailStempel
		Send, % "{Text}" Addendum.Praxis.MailStempel
return ;}
;}

;}

;{ 	6.5. -- TYPORA
#IfWinActive Typora ahk_class Chrome_WidgetWin_1
:*R:.br::<br>
:*R:.cbox::[Codebox=autohotkey file=Untitled.ahk][/Codebox]
:*R:.keyStrg::![](M:/Praxis/_Praxisblog/Bilder/Icons/Key-White_Strg-Links.png){Right}+
:*R:.keyhoch::![](M:/Praxis/_Praxisblog/Bilder/Icons/hoch.png)
:*R:.keyrunter::![](M:/Praxis/_Praxisblog/Bilder/Icons/runter.png)
:*R:.keyrechts::![](M:/Praxis/_Praxisblog/Bilder/Icons/KEy-White-Arrow-right.png)
:*R:.keylinks::![](M:/Praxis/_Praxisblog/Bilder/Icons/links.png)
:*R:.keyc::![](M:/Praxis/_Praxisblog/Bilder/Icons/Key-White-c.png)
:*R:.keyx::![](M:/Praxis/_Praxisblog/Bilder/Icons/Key-White-x.png)
:*R:.keyv::![](M:/Praxis/_Praxisblog/Bilder/Icons/Key-White-v.png)
:*R:.keym::![](M:/Praxis/_Praxisblog/Bilder/Icons/Key-White-M.png)
:*R:.keyL::![](M:/Praxis/_Praxisblog/Bilder/Icons/KeyL.png)
:*R:.keyK::![](M:/Praxis/_Praxisblog/Bilder/Icons/KeyK.png)
:*R:.keyAlt::![](M:/Praxis/_Praxisblog/Bilder/Icons/Alt.png){Right}+
:*R:.keyF01::![](M:/Praxis/_Praxisblog/Bilder/Icons/F1.png)
:*R:.keyF2::![](M:/Praxis/_Praxisblog/Bilder/Icons/F2.png)
:*R:.keyF3::![](M:/Praxis/_Praxisblog/Bilder/Icons/F3.png)
:*R:.keyF4::![](M:/Praxis/_Praxisblog/Bilder/Icons/F4.png)
:*R:.keyF5::![](M:/Praxis/_Praxisblog/Bilder/Icons/F5.png)
:*R:.keyF6::![](M:/Praxis/_Praxisblog/Bilder/Icons/F6.png)
:*R:.keyF7::![](M:/Praxis/_Praxisblog/Bilder/Icons/F7.png)
:*R:.keyF8::![](M:/Praxis/_Praxisblog/Bilder/Icons/F8.png)
:*R:.keyF9::![](M:/Praxis/_Praxisblog/Bilder/Icons/F9.png)
:*R:.keyF10::![](M:/Praxis/_Praxisblog/Bilder/Icons/F10.png)
:*R:.keyF11::![](M:/Praxis/_Praxisblog/Bilder/Icons/F11.png)
:*R:.keyF12::![](M:/Praxis/_Praxisblog/Bilder/Icons/F12.png)
:*R:.keyRaute::![](M:/Praxis/_Praxisblog/Bilder/Icons/Raute.png)
:*R:.keyDicom2Albis::![](M:/Praxis/_Praxisblog/Bilder/Icons/Dicom2Albis.png)
:*R:.keyMonet::![](M:/Praxis/_Praxisblog/Bilder/Icons/Monet.png)
:*R:.keyAddendum::![](M:/Praxis/_Praxisblog/Bilder/Icons/Addendum.png)
:*R:.keyLaborabruf::![](M:/Praxis/_Praxisblog/Bilder/Icons/Laborabruf.png)
:*R:.keyPraxomat::![](M:/Praxis/_Praxisblog/Bilder/Icons/Praxomat.png)
:*R:.keyGVU::![](M:/Praxis/_Praxisblog/Bilder/Icons/Gesundheitsvorsorgeliste.png)
:*R:.keyTPParser::![](M:/Praxis/_Praxisblog/Bilder/Icons/Tagesprotokollparser.png)
:*R:.keyScanPool::![](M:/Praxis/_Praxisblog/Bilder/Icons/ScanPool.png)
:*R:.keySonoCapture::![](M:/Praxis/_Praxisblog/Bilder/Icons/SonoCapture.png)
:*R:.keyAchtung::![](M:/Praxis/_Praxisblog/Bilder/Icons/Achtung.png)
:*R:.keyMedikamente::![](M:/Praxis/_Praxisblog/Bilder/Icons/Medikamente.png)
:*R:.keyMenu::![](M:/Praxis/_Praxisblog/Bilder/Icons/Menu.png)
:*R:.keyImpfung::![](M:/Praxis/_Praxisblog/Bilder/Icons/Impfung.png)
:*R:.keyCave::![](M:/Praxis/_Praxisblog/Bilder/Icons/Cave.png)
:*R:.TrennerKlein::![](Docs\Trenner_klein.png)
:*R:.TrennerGroß::![](Docs\Trenner.png)
#IfWinActive

;}

;}


;----------------------------------- End of Autoexecute ----------------------------------------------------------------------------------

;{10. Hotkey Labels, Hotstring Functions

;{ ############### 	HOTKEY-FUNCTIONS	###############

IsDatumsfeld() {                                                                                	; stellt fest ob sich der Eingabefocus in einem editierbaren Feld befindet

	feldOk    	:= 0
	cfHwnd 	:= GetFocusedControlHwnd()
	cContent 	:= ControlGetText("",  "ahk_id " cfHwnd, "", "", "")
	hactive		:= WinExist("A")
	RegExMatch(Control_GetClassNN(hactive, cfHwnd), "(?<=Edit)\d+", classnn)
	RegExMatch(WinGetTitle(hactive), "Muster\s\d+\w*", wtitle)
	classnn += 0

	If 			InStr(wtitle, "Muster 1a")
		If classnn in 1,2
			feldOk:= 1
	else if 	InStr(wtitle, "Muster 12a")
		If classnn in 6,7,12,13,17,18,22,23,27,28,32,33,37,38,43,44,52,53,59,60,65,66,71,72,78,79,83,84
			feldOk:= 2
	else if 	InStr(wtitle, "Muster 21")
		If classnn in 6,7
			feldOk:= 3
	else if 	InStr(wtitle, "Muster 20")
		If classnn in 3,4,7,8,11,12,15,16
			feldOk:= 4
	else If 	RegExMatch(cContent, "\d\d\.\d\d\.\d\d\d\d", datum)
	        feldOk:= 5

	;ToolTip, % "wtitle:" wtitle "`nclassnn: '" classnn "'`nFeldOk: " feldOk, 2000, 100, 3

	If (datum = "" && feldOk)
		If !RegExMatch(cContent, "\d\d\.\d\d\.\d\d\d\d", datum)
 			FormatTime, datum, %A_Now%, ddMMyyyy

	If feldOk
			return {"txt":cContent, "datum":datum, "hwnd":cfHwnd, "hWin":hactive, "WinTitle":wtitle, "classnn":classnn}
	else
			return 0
}

AddToDate(Feld, val, timeunits) {                                                        	; addiert Tage bzw. eine Anzahl von Monaten zu einem Datum hinzu

	calcdate:= SubStr(Feld.Datum, 7, 4) . SubStr(Feld.Datum, 4, 2) . SubStr(Feld.Datum, 1, 2)
	calcdate += val, %timeunits%
	FormatTime, newdate, % calcdate, dd.MM.yyyy
	ControlSetText,, % newdate, % "ahk_id " Feld.hwnd

}

EnableAllControls(ask=1) {                                                                	; Zugriff für alle Steuerelemente eines Fenster ermöglichen

	hwnd:= WinExist("A")

	if ask
	{
			WinGetTitle, wtitle, ahk_id %hwnd%
			MsgBox, 4, Addendum für Albis on Windows, Wollen Sie im Fenster: `n%hwnd%`n`"%wtitle%`"`nwirklich alle Felder freischalten und auch anzeigen?
			If MsgBox, No
					return
	}

	For key, control in GetControls(hwnd,"","", "Enabled,hwnd,Visible")
	{
			If !(control.enabled)
					Control, Enable,,, % "ahk_id " control.hwnd
			If !(control.Visible)
					Control, Show,,, % "ahk_id " control.hwnd
	}

	return
}

UnRedo(Do) {                                                                                    	; Undo und Redo in den Textfeldern im Albisfenster

	;WM_UNDO = 0x0304
	ControlGetFocus	, focused	, % "ahk_id " (AlbisWinID := AlbisWinID())
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
					SciTEOutput(text)

			}
			else if (Do = "Redo") {

					If (Addendum.UndoRedo.MaxIndex() > 0) {
							text := Addendum.UndoRedo.List.Pop()
							ControlSetText,, % text, % "ahk_id " Addendum.UndoRedo.OnUnRedo
					}

					SciTEOutput(text)

			}

	}

}

FoxitReaderSignieren() {                                                                        	; macht ein Backup der PDF Datei bevor diese signiert wird

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process") {
		If InStr(process.name, "FoxitReader") {
			RegExMatch(process.CommandLine, "\s\" q "(.*)?\" q, cmdline)
			break
		}
	}

	SplitPath, cmdline1, PDFName
	FileCopy, % cmdline1, % Addendum.BefundOrdner "\Backup\" PdfName, 1

	FoxitReader_SignaturSetzen()

}

;}
;{ ############### 	SCITE4AUTOHOTKEY	###############
RunSomething:              	;{
	Run, %AddendumDir%\Module\Albis_Funktionen\PraxTT.ahk
Return
;}
scriptVars:                     	;{     zeigt die Variablen des AddendumSkriptes                         	(für Scite4Autohotkey)
	ListVars
return
;}
FuerTests:							;{		Strg (links) + Alt (links) + t
	;MsgBox, % AlbisGetActiveControl("contraction")

return
;}
SendRawBzu:                   	;{	 	Strg (links) + Ziffernblock 2                                             	(für Scite4Autohotkey)
	SendRawFast(";}")
return
;}
SendRawBAuf:               	;{	 	Strg (links) + Ziffernblock 1                                             	(für Scite4Autohotkey)
	SendRawFast(";{ ")
return
;}
SendRawElse:                	;{	 	Strg (links) + Ziffernblock 3                                              	(für Scite4Autohotkey)
	SendRawFast(elseSEND)
return
;}
SendRawProzentAll: 	    	;{	 	Strg (links) + 5                                                                	(für Scite4Autohotkey)
	SendRawFast("%%", "L")
return
;}
SendStrgRight: 	            	;{ 	Ziffernblock 6                                                                 	(für Scite4Autohotkey)
	SendInput, {LControl Down}{Right}{LControl Up}
return
;}
SendStrgLeft:                 	;{ 	Ziffernblock 4                                                                 	(für Scite4Autohotkey)
	SendInput, {LControl Down}{Left}{LControl Up}
return
;}
SendRawBAll:                	;{     Alt (links) 	+ 7                                                                 	(für Scite4Autohotkey)
	SendInput, % "{Raw}{}"
	SendInput, {Left}
return
;}
SendRawRAll:                	;{     Strg (links) + 8                                                                 	(für Scite4Autohotkey)
	SendRawFast("()", "L")
return
;}
SendRawKAll:                	;{     Alt (links) 	+ 8                                                                 	(für Scite4Autohotkey)
	SendRawFast("[]", "L")
return
;}
SendRawKauf:               	;{     Win (links) + 8                                                                 	(für Scite4Autohotkey)
	SendRawFast("[", "R")
return
;}
SendRawKzu:                	;{     Win (links) + 9                                                                 	(für Scite4Autohotkey)
	SendRawFast("]", "L")
return
;}
SendRawR:                    	;{	 	Strg (links) + ´                                                                	(für Scite4Autohotkey) | `r
	SendRawFast("``r", "")
return
;}
SendRawN:                   	;{	 	Strg (links) + `                                                                	(für Scite4Autohotkey) | `n
	SendRawFast("``n", "")
Return
;}
SendRawTab:                	;{	 	Strg (links) + Alt (links) + `                                               	(für Scite4Autohotkey) | `t
	SendRawFast("``t", "")
return
;}
SendFourSpace:            	;{	 	Strg (links) + Pfeil rechts                                                   	(für Scite4Autohotkey)
	SendInput, {Space 4}
Return
;}
SendFourBackspace:      	;{	 	Strg (links) + Delete                                                         	(für Scite4Autohotkey)
	SendInput, {BackSpace 4}
Return
;}
SciteDescriptionBlock:    	;{	 	Alt (links)   + v                                                               	(für Scite4Autohotkey)
	Send, `;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Send, {Enter}
	SendRaw, % "; " Clipboard
	Send, {Enter}
	Send, `;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Send, {Enter}
 return
;}
SciteCommentBlock:      	;{	 	Strg (links) + ,                                                                  	(für Scite4Autohotkey)
	SendAlternately("/* | */")
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
	Txt				:= direction = - 1 ? StrFlip(sci.curText) : sci.curText
	SLen				:= StrLen(Txt)
	LineStartPos	:= endpos - SLen                                                                                			;lineendpos ist immer +2 wegen CR + LF Zeichen am Ende
	CPosInTxt		:= curpos - LineStartPos
	newPos			:= curpos + ( direction * InStr(Txt, ",", false, CPosInTxt + 1, 1) )

	SendMessage, 2025, %newPos%, 0,, ahk_id %hSci%
	gosub MsgBoxer
sscDir := 0

return

MsgBoxer:
	MsgBox,                                                                                                                        ;" ,`ncurLine: " curline
	(LTrim
			direction    	:  %direction%
			CaretPosition	:  %curpos%
			Linelength 	:  %curlen%
			Bufferlänge 	:  %buflen%
			LineStartPos	:  %LineStartPos%
			LineEndPos	:  %EndPos%
			newPos     	:  %newPos%
			Textlänge  	:  %SLen%
			Text          	:  %Txt%
	)
return
}

;}

;}
;{ ###############                 ALBIS            	###############
AlbisHeuteDatum:         	;{ 	das Albis Programmdatum auf heute setzen
	AlbisSetzeProgrammDatum(A_DD "." A_MM "." A_YYYY)
Return
;}
AlbisKopieren:               	;{ 	Strg	+	c                                                                	(wie sonst überall in Windows nur nicht in Albis)

	AlbisCText:= AlbisCopyCut()
	If StrLen(AlbisCText) > 0
	{
			clipboard := AlbisCText
			GdiSplash(SubStr(AlbisCText, 1, 30) . "`.`.`.", 1)
	}
	else
			Tooltip("!Nichts kopiert!", 1)
return
;}
AlbisAusschneiden:        	;{ 	Strg	+	x                                                                	(wie sonst überall in Windows nur nicht in Albis)

	AlbisCText:= AlbisCopyCut()
	If StrLen(AlbisCText) > 0
	{
			clipboard := AlbisCText
			GdiSplash(SubStr(AlbisCText, 1, 30) . "`.`.`.", 1)
			SendInput, {DEL}
	}
	else
			Tooltip("!Nichts ausgeschnitten!", 1)

return
;}
AlbisEinfuegen:             	;{ 	Strg	+	v                                                                	(wie sonst überall in Windows nur nicht in Albis)
	If (AlbisCText <> clipboard) {
        FileAppend, %clipboard% `n, logs'n'data\CopyAndPaste.log
	}
	SendRaw, % Clipboard
return
;}
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
	SendMessage, 0x224,,,, % "ahk_id " AlbisGetMDIClientHandle()	; WM_MDINEXT
return
;}
Abrechnungshelfer:       	;{ 	Strg	+	F9
	IfWinNotExist, Abrechnungshelfer ahk_class AutoHotkeyGUI
	{
        PraxTT("Das Modul Abrechnungshelfer`nwird gestart!", "4 2")
		Run, %AddendumDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%AddendumDir%\Module\Abrechnung\Abrechnungshelfer.ahk"
	}
return
;}
ScanPoolStarten:           	;{ 	Strg	+	F10
	IfWinNotExist, ScanPool ahk_class AutoHotkeyGUI
	{
        PraxTT("Das Modul ScanPool`nwird gestart!", "4 2")
		Run, %AddendumDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%AddendumDir%\Module\Albis_Funktionen\ScanPool\ScanPool.ahk"
	}
return
;}
Schulzettel:                    	;{ 	Strg 	+ F11
	Run, %AddendumDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%AddendumDir%\Module\Albis_Funktionen\Schulzettel.ahk"
return ;}
Hausbesuche:               	;{ 	Strg 	+ F12
	Run, %AddendumDir%\include\AHK_H\x64w\AutohotkeyH_U64.exe /f "%AddendumDir%\Module\Albis_Funktionen\Hausbesuche.ahk"
return ;}
DauermedBezeichner:   	;{ 	nur # eingeben im Dauermedikamentenfenster
	f:= GetFocusedControl()
	ControlGetText, fT,, ahk_id %f%
	If (fT="") {
        	c:= GetClassName(f)
        	If InStr(c, "Edit")
                	MedTrenner(f)
	} else {
        	SendInput, #
	}
return
;}
DauermedRauf:             	;{ 	Strg	+ Pfeil nach oben
	PostMessage, 0x00F5,,, Button1,  Dauermedikamente ahk_class #32770
	SetTimer, DauermedFokus, -50
Return
;}
DauermedRunter:          	;{ 	Strg	+ Pfeil nach unten
	PostMessage, 0x00F5,,, Button2,  Dauermedikamente ahk_class #32770
	SetTimer, DauermedFokus, -50
return
;}
DiagnoseRauf:              	;{ 	Strg	+ Pfeil nach oben
	PostMessage, 0x00F5,,, Button1,  Dauerdiagnosen von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
Return ;}
DiagnoseRunter:            	;{ 	Strg	+ Pfeil nach unten
	PostMessage, 0x00F5,,, Button2,  Dauerdiagnosen von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
return ;}
CaveVonRauf:               	;{ 	Strg	+ Pfeil nach oben  stark Verbesserungswürdig
	PostMessage, 0x00F5,,, Button1,  Cave! von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
Return ;}
CaveVonRunter:            	;{ 	Strg	+ Pfeil nach unten
	PostMessage, 0x00F5,,, Button2,  Cave! von ahk_class #32770
	SetTimer, DauerDiagnoseFokus, -50
return ;}
TTipOff:                        	;{
	ToolTip,,,, % TT
return
;}
DauermedFokus:           	;{ 	Eingabefokus auf das Listview im Dauermedikamentenfenster
	If WinExist("Dauermedikamente ahk_class #32770")
			ControlFocus, SysListView321, Dauermedikamente ahk_class #32770
return
;}
DauerDiagnoseFokus:   	;{ 	Eingabefokus auf das Listview im Dauerdiagnosenfenster
	If WinExist("Dauerdiagnosen von ahk_class #32770")
			ControlFocus, SysListView321, Dauerdiagnosen von ahk_class #32770
return
;}
CaveVonFokus:             	;{ 	Eingabefokus auf das Listview im Cave! von fenster
	If WinExist("Dauerdiagnosen von ahk_class #32770")
		ControlFocus, SysListView321, Cave! von ahk_class #32770
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
			AddToDate(Feld, 4*7, "days")
	else
			SendInput, {F12}
Return
;}
MinusEinMonat:            	;{ 	F11                                                                                     - in allen Datumsfeldern in Albis
	If IsObject(Feld:= IsDatumsfeld())
			AddToDate(Feld, -4*7, "days")
	else
			SendInput, {F11}
Return
;}
Kalender:                      	;{ 	Shift	+ F3
	If IsObject(feld:= IsDatumsfeld())
	{
				If !InStr(GetProcessNameFromID(feld.hwin), "albis")
				{
						SendInput, {Shift Down}{F3}{Shift Up}
						return
				}

			;Kalender Gui erstellen
				Calendar(feld)
	}
	else
	{
				SendInput, {Shift Down}{F3}{Shift Up}
	}
return
;}
DruckeLabor:                	;{ 	Alt 	+ F8
	AlbisDruckeLaborblatt(1, "Standard", "")
return
;}
Laborblatt:                    	;{ 	Alt 	+ Rechts
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") or InStr(AlbisGetActiveWindowType(), "Laborblatt")
			return
	AlbisZeigeLaborBlatt()
return
;}
Karteikarte:                   	;{ 	Alt 	+ Links
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") or InStr(AlbisGetActiveWindowType(), "Karteikarte")
			return
	AlbisZeigeKarteikarte()
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
KarteikartenMenu:          	;{
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
exportiereAlles:               	;{  	Strg	+ Alt + e
	exportiereAlles()
return ;}

;}
;{ ###############	          ALLGEMEIN        	###############
CaptureDoubleClick:       	;{
	if (A_PriorHotkey = A_ThisHotkey && A_TimeSincePriorHotkey < 500) {
			SendInput, {LControl Down}
    }
return
;}
KopierenUeberall:         	;{	 	Strg	+ c
	CopyToClipboard := true
AusschneidenUeberall:   	;	 	Strg (links) + x

	SetKeyDelay, 20, 50

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
	SetKeyDelay, 10, 20

return
;}
ScreenShot:                   	;{    	Alt	+ q
	ScreenShotArea()
Return
;}
DoNothing:                   	;{
return
;}
BereiteDichVor:					;{
	ExitApp
return
;}
FastSets:							;{
	SendInput, {F3}
	sleep, 500
	ControlClick, Button27, Hautkrebsscreening - Nichtdermatologe ahk_class #32770
	sleep, 500
	ControlClick, Button18, Hautkrebsscreening - Nichtdermatologe ahk_class #32770
	WinWaitActive, Elektronischer Export Dokumentationsbögen zum eHautkrebs-Screening ahk_class #32770,, 2
	SendInput, {Down}
return
;}
AddendumPausieren:     	;{ 	Strg + Alt + Pause
		PraxTT("Addendum ist pausiert", "5 0")
		Pause, Toggle
return ;}
PatientenkarteiOeffnen: 	;{

	PatientenID:= Clip()
	If !RegExMatch(PatientenID, "\d+") {
		PraxTT("Der ausgewählte Text (" PatientenID ")`nkann keine Patienten ID sein!", "3 3")
		return
	}
	else if !oPat.Haskey(PatientenID)	{
		MsgBox, 1, Addendum für Albis on Windows, % "Die ausgewählte Patienten ID (" PatientenID ")`n ist nicht bekannt. Soll diese ID dennoch benutzt werden?"
		IfMsgBox, No
			return
	}

	;PraxTT("Geöffnet wird die Karteikarte des Patienten:`n#2" oPat[PatientenID].Nn ", " oPat[PatientenID].Vn ", geb.am " oPat[PatientenID].Gd "(" PatientenID ")", "6 3")
	AlbisAkteOeffnen(PatientenID, PatientenID)

return ;}
SendClipBoardText:       	;{ 	Strg + Alt + F10

	If (StrLen(TextToSend) = 0) {
		TextToSend 	:= ClipBoard
		ClipBoard 	:= ""
	}

	PraxTT("Sende ClipBoard Inhalt.", "3 3")
	Loop, % StrLen(TextToSend) {
		Send, % "{Text}" SubStr(TextToSend, A_Index, 1)
		Sleep, 75
	}
	;Send, % "{Text}" TextToSend
	SetTimer, emptyTextToSend, -180000 ; nach 3min leeren

return
emptyTextToSend:
	TextToSend := ""
return
;}
ScanBeenden:               	;{
 If (hwnd := WinExist("ScanSnap Manager - Bild erfassen und Datei speichern ahk_class #32770"))
	res := VerifiedClick("Button2", hwnd)
return ;}
;}

;}

;{11. Labels - TrayMenu

UserAway:                                      	;{	                                                            				;-- macht verschiedene Dinge wenn niemand den Computer eine zeitlang benutzt

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; pausiert den Timer der nopopup_Labels nach 30sek. wenn niemand am Computer ist
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		If   	 	(A_TimeIdle > 60000) &&  !Addendum.useraway {

			Addendum.useraway := true

				If IsLabel(nopopup_%compname%)
					SetTimer, nopopup_%compname%, Off

				SetTimer, UserAway, 500                                        	;jetzt schneller nachsehen damit es nicht 5min. dauert bis die Überwachungsroutine nopopup_compname wieder anspringt

		}
		else if	(A_TimeIdle < 300)  	&&  Addendum.useraway {

				Addendum.useraway := true

				If IsLabel(nopopup_%compname%)
					SetTimer, nopopup_%compname%, On

				SetTimer, UserAway, 300000                                    	;die Überwachungsroutine wird wieder seltener aufgerufen werden

		}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; wenn ein neuer Tag angebrochen ist, zuerst das Albis Programmdatum aktualisieren und dann das Skript neu starten (verhindert Programmfehler, gleichzeitig
	; werden bei einem Skriptneustart die Addendum eigene Patientendatenbank sortiert und das Tagesprotokoll neu begonnen)
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		If !(A_DD = Addendum.AktuellerTag) {
				AlbisSetzeProgrammDatum(A_DD . "." . A_MM . "." . A_YYYY)
				gosub SkriptReload
		}

return
;}
DoingStatisticalThings:                     	;{

return
;}
SkriptReload:                                 	;{                                                            					;-- besseres Reload eines Skriptes

  ; es wird ein RestartScript gestartet, welches zuerst das laufende Skript komplett beendet und dann das Addendum-Skript per run Befehl erneut startet
  ; dies verhindert einige Probleme mit datumsabhängigen Handles bestimmter Prozesse im System, die Hook und Hotkey Routinen sind zuverlässiger dadurch geworden

	Script    	:= Addendum.AddendumDir "\Module\Addendum\Addendum.ahk"
	scriptPID 	:= DllCall("GetCurrentProcessId")
	__          	:= q " " q                                             	; für bessere Lesbarkeit des Codes in der unteren Zeile

	cmdline := "Autohotkey.exe /f " q Addendum.AddendumDir "\include\Addendum_Reload.ahk" __ Script __ "1" __ A_ScriptHwnd __ scriptPID __ " 1" q
	Run, % cmdline
	ExitApp

return ;}

ShowTextProtocol(FilePath) {                                                                                            	;-- zeigt Textdateien im eingestellten Texteditor an
	Run % FilePath
}
ZeigePatDB() {                                                                              		                   			;-- zeigt in einer eigenen Listview die Patienten.txt
	For PatID, PatData in oPat
		AutoListview("Pat.Nr.|Nachname|Vorname|Geburtsdatum|Geschlecht|Krankenkasse|letzte GVU", PatID "," PatData.Nn "," PatData.Vn "," PatData.Gt "," PatData.Gd "," PatData.KK "," PatData.letzteGVU, ",")
}
ZeigeFehlerProtokoll() {                                                                                  					;-- das Skriptfehlerprotokoll wird angezeigt

	filebasename    	:= Addendum.AddendumDir "\logs'n'data\ErrorLogs\Fehlerprotokoll-"
	thisMonth          	:= A_MM A_YYYY
	lastMonth          	:= (A_MM - 1 = 0) ? ("12" A_YYYY-1)
	thisMonthProtocol	:=  filebasename thisMonth ".txt"
	lastMonthProtocol	:=  filebasename lastMonth ".txt"

	If FileExist(thisMonthProtocol)
		Run % thisMonthProtocol
	else if FileExist(lastMonthProtocol)
		Run % lastMonthProtocol
	else
		MsgBox, 1, % A_ScriptName, Es wurden keine Fehler erfasst!

return
}
NoNo() {                                                                         	                                 				;-- Addendum "Über" Fenster

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

	C .= "installierte Autohotkeyversion: " 	A_AhkVersion "`n"
	C .= "Computer Name: "                   	A_ComputerName "`n"
	C .= "IP Adresse des Computer: "       	A_IPAddress1 "`n"
	C .= "Login Name: "                         	A_UserName "`n"
	C .= "Nutzer hat Adminrechte: "         	(A_IsAdmin=1 ? "Ja" : "Nein") "`n"
	C .= "Addendum Hauptverzeichnis: " 	Addendum.AddendumDir "`n"
	C .= "Skript ist compiliert: "                	(A_IsCompiled=1 ? "Ja" : "Nein") "`n"
	C .= "Skripthandle: "                         	A_ScriptHwnd "`n"
	C .= "Betriebssystem: Windows " StrSplit(A_OsVersion, ".").1 " " (A_Is64bitOS=1 ? "64" : "32") "-bit`n"
	C .= "Bildschirm-DPI: "                      	A_ScreenDPI "`n"
	C .= "Bildschirmgröße in Pixel: "         	A_ScreenWidth "x" A_ScreenHeight "`n"
	C .= "letzte Fehlernummer: "              	A_LastError

	MaxLen := 0
	Loop, Parse, C, `n
		MaxLen := StrLen(A_LoopField) > MaxLen ? StrLen(A_LoopField) : MaxLen

	B := "`n" D2 "`n" SubStr(ul, 1, MaxLen - 20) "`n" C

	Gui Uber: New, -Caption +ToolWindow  0x94400000 HWNDhUber
	Gui Uber: Color, c202842
	Gui Uber: Margin, 0, 10

	Gui Uber: Add, Picture, % "x" (Floor(UberW/2) - 250) " ym vUberPic", % Addendum.AddendumDir "\assets\AddendumBigLogo.jpg"
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
	ControlFocus, UberOk, ahk_id %hUber%
	SendInput, {Left}

	Gui Uber: Show, AutoSize, Addendum für Albis on Windows
	DllCall("HideCaret","Int", hUberText)

return

UberEnde:
	If (A_GuiControl="UberOk")
			Gui Uber: Destroy
return
}

Wartezeit:                                     	;{

	FileReadLine, WZeit, %WZStatistikFile%, 1
	FileReadLine, WPat, %WZStatistikFile%, 2
		if (errorlevel = 0) {
			x = %A_ScreenWidth% - 400
			y = %A_ScreenHeight% - 120
			SplashPraxID = PraxomatInfo("Addendum für Albis on Windows", "konnte die Wartezeitstatistikdatei nicht öffnen", %x% , %y%, 1000)
			;SetTimer, SplashTextAus, %showtime%
		}

return
;}
ModulStarter:                                 	;{                                                                            	;-- startet Module vom Tray-Menu
	If RegExMatch(modul[A_ThisMenuItem], "\.exe$")
		RunExe(modul[A_ThisMenuItem])
	Else if RegExMatch(modul[A_ThisMenuItem], "\.ahk$")
		RunSkript(modul[A_ThisMenuItem])
return ;}
ToolStarter:                                   	;{                                                                            	;-- startet Tools vom Tray-Menu
	If InStr(A_ThisMenuItem, "Windows komplett neu starten")
		Run, Shutdown -g -t 0
	else if RegExMatch(tool[A_ThisMenuItem], "\.exe")
		Run, % tool[A_ThisMenuItem]
	else
		RunSkript(tool[A_ThisMenuItem])
return ;}

; ----------------------------------------- Funktionen ein- oder ausstellen
Menu_PauseAddendum:                  	;{
	Pause, Toggle, 1
return ;}
Menu_AlbisAutoPosition:                   	;{
	Addendum.AlbisLocationChange := !Addendum.AlbisLocationChange
	Menu, SubMenu3, ToggleCheck, % "Albis AutoSize"
	IniWrite, % (Addendum.AlbisLocationChange ? "ja":"nein")	, % Addendum.AddendumIni, % compname, % "Albis_AutoGroesse"
return
;}
Menu_AddendumGui:                    	;{
	Addendum.AddendumGui := !Addendum.AddendumGui
	Menu, SubMenu3, ToggleCheck, % "Addendum Infofenster"
	IniWrite, % (Addendum.AddendumGui ? "1":"0")	        	, % Addendum.AddendumIni, % compname, % "Addendum_Infofenster"
	If Addendum.AddendumGui {
		Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If RegExMatch(Addendum.AktiveAnzeige, "i).*\|(Karteikarte)|(Laborblatt)")
				AddendumGui()
	} else If WinExist("ahk_id " hadm) {
		Gui, adm: 	Destroy
		Gui, adm2:	Destroy
	}
return ;}
Menu_GVUAutomation:                    	;{
	Addendum.GVUAutomation := !Addendum.GVUAutomation
	Menu, SubMenu3, ToggleCheck, % "Albis GVU automatisieren"
	IniWrite, % (Addendum.GVUAutomation ? "1":"0")        	, % Addendum.AddendumIni, % compname, % "GVU_automatisieren"
return
;}
Menu_PDFSignatureAutomation:      	;{

	; Menupunkt umschalten, Einstellung speichern
		Addendum.PDFSignieren := !Addendum.PDFSignieren
		Menu, SubMenu3, ToggleCheck, % "FoxitReader Signaturhilfe"
		IniWrite, % (Addendum.PDFSignieren ? "ja":"nein")          	, % Addendum.AddendumIni, % compname, % "FoxitPdfSignature_automatisieren"

	; weitere Funktionen an- oder abschalten
		If Addendum.PdfSignieren {

			; Hotkey Funktion einschalten
				;func_FRSignature := Func("FoxitReader_SignaturSetzen")
				Hotkey, IfWinActive, % "ahk_class classFoxitReader"
				Hotkey, % Addendum.PDFSignieren_Kuerzel, % Func("FoxitReader_SignaturSetzen")
				Hotkey, IfWinActive

			; AutoCloseTab - Tray Menu - anschalten
				Menu, SubMenu3, Enable, FoxitReader Dokument automatisch schliessen
				If Addendum.AutoCloseFoxitTab
					Menu, SubMenu3, Check	  , FoxitReader Dokument automatisch schliessen
				else
					Menu, SubMenu3, UnCheck, FoxitReader Dokument automatisch schliessen

		} else {

			; Hotkey Funktion ausschalten
				Hotkey, % Addendum.PDFSignieren_Kuerzel, Off

			; AutoCloseTab - Tray Menu - ausschalten
				Menu, SubMenu3, Disable, FoxitReader Dokument automatisch schliessen

		}

return ;}
Menu_PDFSignatureAutoTabClose:	;{
	Addendum.AutoCloseFoxitTab := !Addendum.AutoCloseFoxitTab
	Menu, SubMenu3, ToggleCheck, % "FoxitReader Dokument automatisch schliessen"
	IniWrite, % (Addendum.AutoCloseFoxitTab ? "ja":"nein") 	, % Addendum.AddendumIni, % compname, % "FoxitTabSchliessenNachSignierung"
return ;}
Menu_LaborAbrufManuell:               	;{

	Addendum.Labor.AutoAbruf := !Addendum.Labor.AutoAbruf
	Menu, SubMenu3, ToggleCheck, % "Laborabruf automatisieren"
	IniWrite, % (Addendum.Labor.AutoAbruf ? "ja":"nein")   	, % Addendum.AddendumIni, % compname, % "Laborabruf_automatisieren"

return ;}
Menu_LaborAbrufTimer:                  	;{

	Addendum.Labor.AbrufTimer := !Addendum.Labor.AbrufTimer
	Menu, SubMenu3, ToggleCheck, % "zeitgesteuerter Laborabruf"
	If Addendum.Labor-AbrufTimer {
		Addendum.Labor.AutoAbruf := true    ; muss dann auch an sein
		Menu, SubMenu3, ToggleCheck, % "Laborabruf automatisieren"
		IniWrite, % (Addendum.Labor.AutoAbruf ? "ja":"nein")   	, % Addendum.AddendumIni, % compname, % "Laborabruf_automatisieren"
	}
	IniWrite, % (Addendum.Labor.AbrufTimer ? "ja":"nein")   	, % Addendum.AddendumIni, % compname, % "Laborabruf_Timer"
	IniWrite, % "ja"                                                            	, % Addendum.AddendumIni, % compname, % "Laborabruf_automatisieren" ; dieser muss dann automatisch "An=Ja" sein!

return ;}
Menu_AddendumToolbar:             	;{
	Addendum.ToolbarThread := !Addendum.ToolbarThread
	Menu, SubMenu3, ToggleCheck, % "Addendum Toolbar"
	IniWrite, % (Addendum.ToolbarThread ? "ja":"nein"), % Addendum.AddendumIni, % compname, % "Toolbar_anzeigen"
	admThreads()                                                                             ; startet oder stoppt die Ausführung von Threadcode
return ;}
Menu_Schnellrezept:                      	;{
	Addendum.Schnellrezept := !Addendum.Schnellrezept
	Menu, SubMenu3, ToggleCheck, % "Schnellrezept"
	IniWrite, % (Addendum.Schnellrezept ? "ja":"nein")       	, % Addendum.AddendumIni, % "Addendum", % "Schnellrezept"
return ;}
Menu_mDicom:                              	;{
	Addendum.mDicom := !Addendum.mDicom
	Menu, SubMenu3, ToggleCheck, % "MicroDicom Export"
	IniWrite, % (Addendum.mDicom ? "ja":"nein")              	, % Addendum.AddendumIni, % compname, % "MicroDicomExport_automatisieren"
return ;}
Menu_PopUpMenu:                          	;{
	Addendum.PopUpMenu := !Addendum.PopUpMenu
	Menu, SubMenu3, ToggleCheck, % "PopUpMenu"
	IniWrite, % (Addendum.PopUpMenu ? "ja":"nein")           	, % Addendum.AddendumIni, % compname, % "PopUpMenu"

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
	IniWrite, % (xset := GetSetAutoPosString(A_ThisMenuItem)), % Addendum.AddendumIni, % compname, % "AutoPos_" A_ThisMenu

	;SciteOutPut(xset)
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

;}

;{12. Allgemeine Funktionen

;Screenshot
ScreenShotArea() {                                                                                                    	;-- Screenshot subfunktion

	ListLines, Off
	DetectHiddenWindows, On
	InputRect(vWinX, vWinY, vWinR, vWinB)
	vWinW := vWinR-vWinX, vWinH := vWinB-vWinY
	if (vScrInputRectState = -1)
		return

	vScreen := vWinX "|" vWinY "|" vWinW "|" vWinH
	pScrShotToken := Gdip_Startup()
	pBitmap := Gdip_BitmapFromScreen(vScreen, 0x40CC0020)
	DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", Ptr,pBitmap, PtrP,hBitmap, Int,0xffffffff)
	Gdip_SetBitmapToClipboard(pBitmap)

	SplashImage, % "HBITMAP:" hBitmap, B1
	Sleep, 3000
	SplashImage, Off

	DeleteObject(hBitmap)
	Gdip_DisposeImage(pBitmap)
	Gdip_Shutdown(pScrShotToken)
return
}

InputRect(ByRef vX1, ByRef vY1, ByRef vX2, ByRef vY2) {                                                	;-- Screenshot subfunktion

	ListLines, Off
	global vScrInputRectState := 0
	DetectHiddenWindows, On
	Gui, Scr: -Caption -DPIScale +ToolWindow +AlwaysOnTop +hWndhGuiSel
	Gui, Scr: Color, Yellow
	WinSet, Transparent, 100, % "ahk_id " hGuiSel
	Hotkey, *LButton, InputRect_Return, On
	Hotkey, *RButton, InputRect_End, On
	Hotkey, Esc, InputRect_End, On
	KeyWait, LButton, D
	MouseGetPos, vX0, vY0
	SetTimer, InputRect_Update, 10
	KeyWait, LButton
	Hotkey, *LButton, Off
	Hotkey, Esc, InputRect_End, Off
	SetTimer, InputRect_Update, Off
	Gui, Scr: Destroy
	return

	InputRect_Update:
	if !vScrInputRectState
	{
		MouseGetPos, vX, vY
		(vX < vX0) ? (vX1 := vX, vX2 := vX0) : (vX1 := vX0, vX2 := vX)
		(vY < vY0) ? (vY1 := vY, vY2 := vY0) : (vY1 := vY0, vY2 := vY)
		Gui, 1:Show, % "NA x" vX1 " y" vY1 " w" (vX2-vX1) " h" (vY2-vY1)
		return
	}
	vScrInputRectState := 1

	InputRect_End:
	if !ScrvInputRectState
		vScrInputRectState := -1
	Hotkey, *LButton, Off
	Hotkey, *RButton, Off
	Hotkey, Esc, Off
	SetTimer, InputRect_Update, Off
	Gui, Scr: Destroy

	InputRect_Return:
	return
}

;Infofenster
Splash(string, time=4) {                                                                                					;-- Splash Gui - to show text for some seconds

	CoordMode, Mouse, Screen
	CoordMode, Caret, Screen

	Lmx:= A_CaretX, Lmy:= A_CaretY
	If !Lmx or !Lmy
		MouseGetPos, Lmx, Lmy

	Lc := StrSplit(string, "`n").MaxIndex()
	Lc := Lc > 4 ? 4 : Lc

	; build gui
	Gui sp: New, +HWNDhsp AlwaysOnTop -SysMenu -Caption +ToolWindow
	Gui sp: Font, s10 q5 cWhite, Futura Bk Md
	Gui sp: Color, c3F627F
	Gui sp: Margin, 8, 5
	Gui sp: Add, Text, % "xm ym R" Lc " HWNDhSplashText" , % string
	Gui sp: Show, % "x" (Lmx) " y" (Lmy - 30 - ((Lc - 1) *10)) " NA" , Addendum Splash
	WinSet, Transparent, 120, % "ahk_id " hsp

	; form a rounded rectangle
	SP := GetWindowSpot(hsp)
	WinSet, Region, % "0-0 w" (SP.W) " h" (SP.H) " R20-20", % "ahk_id " hsp

	; close with delay
	SetTimer, SPClosePopup, % "-" (time*1000)

return

SPClosePopup:
      Gui sp: Destroy
return

}

GDISplash(text, time=2) {                                                                                          	;-- Hinweisfenster

		ListLines, Off

		static hGdiSp, GdiSp
		static Options:= "y5 Centre cff000000 q5 s16"
		static Font 	:= "Futura Bk Bt"
		static hFont

		If !hFont
			hFont 	:= CreateFont(Options "`, " Font)

		Gui, GdiSp: Destroy

		cx := A_CaretX , cy := A_CaretY - 35
		If (!cx or !cy) {
			MouseGetPos, cx, cy
			cy += 20
		}

		Gui, GdiSp: New, -Caption -SysMenu -DPIScale +ToolWindow +E0x80000 +HwndhGdiSp

		GetFontTextDimension(hFont, Text, Width, Height ,1)
		Width	:= Width // 1.35
		Height	:= Height // 1.15
		hbm 	:= CreateDIBSection(Width, Height)
		hdc 		:= CreateCompatibleDC()
		obm 	:= SelectObject(hdc, hbm)
		G 		:= Gdip_GraphicsFromHDC(hdc)

		Gdip_SetSmoothingMode(G, 4)
		pBrush 	:= Gdip_BrushCreateSolid(0x778fA2ff)

		Gdip_FillRoundedRectangle(G, pBrush, 0, 0, Width, Height, 5)
		Gdip_TextToGraphics(G, text, Options, Font, Width, Height)
		UpdateLayeredWindow(hGdiSp, hdc, cx, cy, Width, Height)

		Gdip_DeleteBrush(pBrush), DeleteObject(hbm), DeleteDC(hdc), Gdip_DeleteGraphics(G)
		Gui, GdiSp: Show, % "NA x" cx " y" cy, GdiSp

		SetTimer, GDISplashOff, % "-" (time*1000)
		return hGdiSp

GDISplashOff:
		Gui, GdiSp: Destroy
return
}

Tooltip(content, wait:=2) {
	Splash(content, wait)
}

;Text senden
SendRawFast(string, cmd="") {                                                                                   	;-- sendet eine String ohne das es zum Senden von Hotkey's kommen kann

	ListLines, Off
	;Parameter: go - L, R, U, D - stehen für Left, Right, Up, Down, nach diesen kann

	Splash(string " " cmd, 3)

	RegExMatch(cmd, "\d+", do)
	RegExMatch(cmd, "i)\w", go)

	SetKeyDelay, -1, -1

	SendRaw, % string
	If InStr(go, "L")
		Send, {Left %do%}
	else If InStr(go, "R")
		Send, {Right %do%}
	else If InStr(go,"U")
		Send, {Up %do%}
	else If Instr(go, "D")
		Send, {Down %do%}

	SetKeyDelay, 20, 50

}

SendAlternately(chars:="/* | */") {                                                                              	;-- sendet entweder das eine Zeichen und beim nächsten Aufruf das andere Zeichen

	static alternating:= []

	Loop, % alternating.MaxIndex()
		If (alternating[A_Index] = chars) 	{
			SendRaw, % StrSplit(chars, "|").2
			alternating.RemoveAt(A_Index)
			return
		}

	SendRaw, % StrSplit(chars, "|").1
	alternating.Push(chars)

}

SciTEFoldAll() {                                                                                                            	;-- gesamten Code zusammenfalten

	ControlGet, hScintilla1, hwnd,, Scintilla1, % "ahk_class SciTEWindow"
	Sleep 300
	SendMessage, 2662,,,, % "ahk_id " hScintilla1 ; SCI_FOLDALL
	TrayTip("SciTEFoldAll", "hScintilla1: " hScintilla1 "`nErrorLevel: " ErrorLevel, 2 )

}

RunExe(exePath) {                                                                                                         	;-- startet eine ausführbare Datei

	SplitPath, exePath, filename, filepath
	If !FileExist(exepath) {
		PraxTT("Das Modul: " filename "`nim Pfad:`n" filepath "`nist nicht vorhanden!", "6 3")
		return 0
	}

	Run, % exepath, % filepath, ModulPID

return ModulPID
}

RunSkript(modulpath) {                                                                                              	;-- startet oder beendet Autohotkeyskripte

	AddendumDir := Addendum.AddendumDir
	SplitPath, modulpath, filename, filepath
	If !FileExist(modulpath) {
		PraxTT("Das Modul: " filename "`nim Pfad:`n" filepath "`nist nicht vorhanden!", "6 3")
		return 0
	}

	filename := StrReplace(filename, ".ahk")

	If (PID := ScriptIsRunning(filename))  {
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

	If !RegExMatch(modulpath, "^[A-Z]\:\\")
		Run, Autohotkey.exe "%AddendumDir%\%modulpath%",, ModulPID
	else
		Run, Autohotkey.exe "%modulpath%",, ModulPID

return ModulPID
}

;Testfunktionen
exportiereAlles() {

	PatID := AlbisAktuellePatID()
	Loop, Files, % "M:\albiswin\Briefe\*Bild", D
		t .= A_LoopFileDir "`n"

}

;}

;{13. Gesundheitsvorsorge, Hautkrebsscreening, Karteikartenfunktionen
GVUListe(modus) {

		PatID            	:= AlbisAktuellePatID()
		GVUVorhanden	:= false
		GVUZaehler   	:= 0

		If (modus=2)
			QLDatum:= AlbisLeseZeilenDatum(300)
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

		SendMode, Event
		SetKeyDelay, 50, 50

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
		SetKeyDelay, -1, -1
		SendMode, Input

return
}

;}

;{14. WinEventHooks - alle Funktionen und Labels

;                         ---------------------------------------------------------------------------------------------------------------------------------
;                       	 AUF DIESEN FUNKTIONEN BERUHEN DIE WICHTIGSTEN VORGENOMMENEN AUTOMATISIERUNGEN
;                         ---------------------------------------------------------------------------------------------------------------------------------


; ------------------------------------ Initialisieren der Hooks
InitializeWinEventHooks() {                                                                                             	; Robotic Process Automation (RPA) fängt hiermit an!

	/* https://docs.microsoft.com/en-us/windows/desktop/winauto/event-constants

		; EVENT_OBJECT_CREATE                 	:= 0x8000
		; EVENT_OBJECT_DESTROY              	:= 0x8001
		; EVENT_OBJECT_SHOW                   	:= 0x8002
		; EVENT_OBJECT_HIDE                     	:= 0x8003
		; EVENT_OBJECT_REORDER              	:= 0x8004	; A container object has added,removed,or reordered its children. (header control, list-view, toolbar control, and window object - !z-order!)
		; EVENT_OBJECT_FOCUS                 	:= 0x8005	; An object has received the keyboard focus. elements: listview,menubar, popup menu,switch window, tabcontrol, treeview, and window object.
		; EVENT_OBJECT_INVOKED              	:= 0x8013	; An object has been invoked; for example, the user has clicked a button.
		; EVENT_OBJECT_NAMECHANGE     	:= 0x800C   ; An object's name has changed - like a title bar
		; EVENT_OBJECT_VALUECHANGE     	:= 0x800E	; An object's Value property has changed.  edit, header, hot key, progress bar, slider, and up-down, scrollbar

		; EVENT_SYSTEM_CAPTUREEND        	:= 0x0009	; A window has lost mouse capture.
		; EVENT_SYSTEM_CAPTURESTART      	:= 0x0008	; A window has received mouse capture. This event is sent by the system, never by servers.
		; EVENT_SYSTEM_DIALOGEND          	:= 0x0011	; A dialog box has been closed.
		; EVENT_SYSTEM_DIALOGSTART        	:= 0x0010	; A dialog box has been displayed.
		; EVENT_SYSTEM_FOREGROUND      	:= 0x0003	; The foreground window has changed. The system sends this event even if the foreground window has changed to another window in the same

		; EVENT_SYSTEM_MENUSTART           	:= 0x0004
		; EVENT_SYSTEM_MENUPOPUPSTART 	:= 0x0006	; context menu was opened
		; EVENT_SYSTEM_MENUPOPUPEND  	:= 0x0007	; context menu was closed

	 */

		EVENT_SKIPOWNTHREAD         	:= 0x0001
		EVENT_SKIPOWNPROCESS       	:= 0x0002
		EVENT_OUTOFCONTEXT          	:= 0x0000

	; creating the hooks - allgemein zur Erfassung von neuen Fenstern/Dialogen in diversen Programmen
		HookProcAdr         	:= RegisterCallback("WinEventProc", "F")
		hWinEventHook     	:= SetWinEventHook( 0x0003, 0x0003, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x0010, 0x0010, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8000, 0x8000, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8002, 0x8002, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8004, 0x8004, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		Addendum.Hooks[1]	:= {"HPA": HookProcAdr, "hWinEventHook": hWinEventHook}

	/* Shellhook für neue Programmfenster - es werden nur Fenster der obersten Ebene (top-level-windows) erkannt ...

			Damit kann ich WinEvent Hooks, die nur Albis überwachen, realisieren. Wird Albis geschlossen, weil es z.B. neugestartet werden muss, wird der momentan aktive WinEvent Hook beendet.
			Der Shell Hook wartet auf einen neuen Albisprozeß und startet nach Erscheinen des Albisfenster einen neuen Hook. Ohne diese kleinen Umweg würde der alte Hook weiter existieren und
			die Automatisierungsfunktionen würden nicht aufgerufen.

			Warum nur einen Hook auf Albis?
			Autohotkey ist eine Skriptsprache. Diese ist schnell, aber bei weitem nicht so schnell wie eine kompilierbare Programmiersprache wie C++.
			Können Eventnachrichten nicht schnell genug abgearbeitet werden, entgehen sie der Erkennung und es die programmierten Automatisierungen starten nicht.

	*/

	; der ShellHook benötigt ein dummy Gui an das die Ereignisse (Events) geschickt werden
		Gui, AMsg: New, +HWNDhMsgGui +ToolWindow
		Gui, AMsg: Add, Edit, xm ym w300 h300 HWNDhEditMsgGui
		Gui, AMsg: Show, AutoSize NA Hide, % "Addendum Message Gui"
		Addendum.MsgGui.hMsgGui     	:= hMsgGui
		Addendum.MsgGui.hEditMsgGui	:= hEditMsgGui

	; Shellhook wird gestartet
		DllCall("RegisterShellHookWindow", UInt, hMsgGui)
		MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(MsgNum, "ShellHookProc")

	; startet weitere Hooks die nur auf Ereignisse vom Albisprozess reagieren
		AlbisStartHooks()

	; Windowfixer für alle Programmfenster, wenn hinterlegt

	; Eingabe Focus Hook für andere Programme (Autovervollständigung)
		;Addendum.Hooks[6]         	:= GetActiveFocusHook()

return
}

AlbisStartHooks() {                                                                                                     		; startet Hooks die Ereignisse des Albisprozeß registrierem

	global AlbisID_old := 0

	If (AlbisPID := AlbisPID()) {

		Addendum.AlbisWinID	:= AlbisWinID()
		Addendum.AlbisPID   	:= AlbisPID
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

AlbisStopHooks() {                                                                                                         	; beendet alles bestehenden ALBIS Hooks, wenn kein Albisprozeß mehr vorhanden ist

		local idx

		Gui, AMsg: 	Show, Hide
		Gui, adm: 	Destroy
		Gui, adm2: 	Destroy
		SetTimer, toptop, off

	; Toolbar-Thread nicht beenden, Toolbar nur Ausblenden
		If IsObject(Addendum.Thread["Toolbar"]) {
			Addendum.Thread["Toolbar"].ahkFunction("ToolbarShowHide", "Hide")
			PraxTT("Albis is closed.`nAddendum Toolbar: " (Addendum.Thread[1].ahkReady() ? "is still running":"is terminated"), "2 2")
		}

	; ein Teil der Hook's beenden
		Loop % (Addendum.Hooks.MaxIndex() -1) { 	; Hook[1] wird nicht beendet da dieser auch andere Programme untersucht
			idx := A_Index + 1
			if idx in 1,5
				continue
			If IsObject(Addendum.Hooks[idx]) {
				SciTEOutput("Hook[" idx "] - hEvent: " Addendum.Hooks[idx].hEvH ", HookProcAdress: " Addendum.Hooks[idx].HPA " - stopped" )
				UnhookWinEvent(Addendum.Hooks[idx].hEvH, Addendum.Hooks[idx].HPA)
				Addendum.Hooks[idx] := ""
			}
	}


}


; ------------------------------------  spezielle Hooks auf den Albis-Prozeß
AlbisPopUpMenuHook(AlbisPID) {                                                                                  	; Hook zum Abfangen eines Rechtsklick Menu

	HookProcAdr   	:= RegisterCallback("AlbisPopUpMenuProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x0006, 0x0007, 0, HookProcAdr, AlbisPID, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisFocusHook(AlbisPID) {                                                                                            	; Hook ausschließlich für EVENT_OBJECT_FOCUS Nachrichten wird gestartet

	; benötigt um Änderung des Eingabefocus innerhalb von Albis zu erkennen
	; z.B. für eine leider noch funktionierende Autovervollständigungsfunktion

	HookProcAdr   	:= RegisterCallback("AlbisFocusEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, AlbisPID, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisLocationChangeHook(AlbisPID) {                                                                            	; Hook ausschließlich für EVENT_OBJECT_LOCATIONCHANGE Nachrichten wird gestartet

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

GetActiveFocusHook() {                                                                                                 	; Hook der bei jeder Änderung des Eingabefokus aktiv wird

	HookProcAdr   	:= RegisterCallback("ActiveFocusEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, 0, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}

}


; ------------------------------------ Verarbeitung abgefangener Systemmeldungen (EventProcs)
WinEventProc(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {              	; zum Abfangen bestimmter Fensterklassen

		static class_filter	:= [ "OptoAppClass", "#32770", "Teamviewer", "WindowsForms10.Window"
									,  "Qt5QWindowIcon", "AutohotkeyGui", "SciTEWindow"
									,  "classFoxitReader", "IEFrame", "SUMATRA_PDF_FRAME"]
		static app_filter 	:= { 1: "albis"
									,  	2: "albiscs"
									,  	3: "Teamviewer"
									,  	4: "infoBoxWebClient"
									,  	5: "ifap"
									,  	6: "Autohotkey"
									,  	7: "SciTE"
									,  	8: "FoxitReader"
									, 	9: "word"
									, 10: "mDicom"
									, 11: "iexplore"}

		Critical

		If (GetDec(hwnd) = 0)    	; || (StrLen(cClass:=WinGetClass(hwnd)) = 0)
			return 0

		EHookHwnd	:= Format("0x{:x}", hwnd)
		If (Addendum.LastHookHwnd = EHookHwnd)
			return 0

		cClass:=WinGetClass(hwnd)

	;
		If (StrLen(Addendum.FuncCallback) > 0) {

				For idx, viewerclass in PDFViewer
					If InStr(class, viewerclass) {

						If InStr(Addendum.FuncCallback, "|") {

								callbackParam := StrSplit(Addendum.FuncCallback, "|")
								If IsFunc(callbackParam[1])
									func_Call 	:= Func(callbackParam[1]).Bind(callbackParam[2], class, SHookHwnd)

						}
						else {

								fnName := Addendum.FuncCallback
								If IsFunc(fnName)
									func_Call 	:= Func(fnName).Bind(class, SHookHwnd)

						}

						SetTimer, % func_Call, -300
						break

					}

		}

	; nur Fensterklassen Filter
		For fnr, filterclass in class_filter
			If (cClass = filterclass) {

				For snr, item in EHStack
					If (item.Hwnd = EHookHwnd)
						return 0

				WinGetTitle, HTitle, % "ahk_id " EHookHwnd
				WinGetText, HText, % "ahk_id " EHookHwnd
				If (StrLen(HTitle . HText) = 0)
					return 0

				EHStack.InsertAt(1, {"Hwnd":GetHex(EHookHwnd), "Event":GetHex(Event), "Title":HTitle, "Text":StrReplace(StrReplace(HText, "`n", " "), "`r", ""), "Class":cClass})
				;~ If InStr(EHStack.1.Text, "Chipkarte")
					;~ SciTEOutput("Fenster detektiert:   " EHStack.1.Text)
				If !EHWHStatus
					SetTimer, EventHook_WinHandler, -1


				return 0
			}

	; Order & Entry Prozess
		WinGet, EHproc, ProcessName, % "ahk_id " EHookHwnd
		If Instr(EHProc, "infoBoxWebClient") {
			WinGetTitle, Title, % "ahk_id " EHookHwnd
			WinGetText, Text, % "ahk_id " EHookHwnd
			EHStack.InsertAt(1,  {"Hwnd":GetHex(EHookHwnd), "Event":GetHex(Event), "Title":Title, "Text":StrReplace(Text, "`r`n", " "), "Class":cClass})
			If !EHWHStatus
				SetTimer, EventHook_WinHandler, -1
		}



return 0
}

AlbisPopUpMenuProc(hHook, event, hwnd) {                                                                   	; eigene Menupunkte in Albiskontextmenu's unterbringen

	; alle anderen Funktionen für Aufrufe finden sich in %AddendumDir%\Module\Addendum\include\Addendum_PopUpMenu.ahk

	static pm

	If (event = 6)	                                                                             	; Kontextmenu wird aufgerufen
		pm := Addendum_PopUpMenu(hwnd, GetMenuHandle(hwnd))
	else if (event = 7) && IsObject(pm)                                                	; es wurde ein Menupunkt per Mausklick ausgewählt
		Addendum_PopUpMenuItem(pm)

}

AlbisFocusEventProc(hHook, event, hwnd) {                                                                    	; behandelt EVENT_OJECT_FOCUS Nachrichten von Albis

	static AlbisTitleO
	global hFEGui, FE, FEText1

	If (GetDec(hwnd) = 0)
		return 0

	cFocus	:= GetClassName((hwnd := Format("0x{:x}", hwnd)))
	Addendum.AktiveAnzeige := AlbisHookTitle := AlbisGetActiveWinTitle()
	If (AlbisTitleO = AlbisHookTitle) && RegExMatch(AlbisHookTitle, "^\s*\d+\s*\/") ;&& !RegExMatch)|(Laborbuch)|(Wartezimmer)|(EBM)(ToDo)")
		return 0

	AlbisTitleO := AlbisHookTitle
	fn_PatChange := Func("CurrPatientChange").Bind(AlbisHookTitle)
	SetTimer, % fn_PatChange, -0

return 0
}

AlbisLocationChangeEventProc(hHook, event, hwnd) {                                                    	; behandelt EVENT_OBJECT_LOCATIONCHANGE Nachrichten von Albis

	; nur für Monitore mit einer Auflösung > 2k
	; last change: 27.08.2020

		global adm
		static SizeInit := false
		static AlbisX, AlbisY, AlbisW, AlbisH

	; AlbisX, AlbisY, AlbisW, AlbisH werden erstellt
		If !SizeInit {

			SizeInit := true
			ShellTray := GetWindowSpot(WinExist("ahk_class Shell_TrayWnd")) ; Größe der Taskbar

			RegExMatch(Addendum.AlbisGroesse, "i)y\s*(?<Y>\d+)"        	, A)
			AlbisY := AY

			RegExMatch(Addendum.AlbisGroesse, "i)w\s*(?<W>\d+)"      	, A)
			AlbisW := AW

			RegExMatch(Addendum.AlbisGroesse, "i)h\s*(?<H>\d+|Max)"	, A)
			AlbisH := (AH = "Max") ? (A_ScreenHeight - ShellTray.H + 12) : AH

			RegExMatch(Addendum.AlbisGroesse, "i)x\s*(?<X>\d+|R|L)" 	, A)
			If (AX = "R")        	; rechter Bildschirmrand
				AlbisX := A_ScreenWidth - AlbisW
			else if (AX = "L") 	; linker Bildschirmrand
				AlbisX := 0
			else
				AlbisX := AX
		}

		Addendum.AlbisWinID := AlbisWinID()

		If (hwnd = 0) || (GetHex(hwnd) <> Addendum.AlbisWinID)
			return 0

		Addendum.MonSize  	:= A_ScreenWidth > 1920 ? 2 : 1
		Addendum.Resolution  	:= A_ScreenWidth > 1920 ? "4k" : "2k"

		If (A_ScreenWidth <= 1920)
			return 0

	; wenn Albis minimiert wird dann wird die AddendumToolbar versteckt (diese wird sonst auf dem Desktop angezeigt)
		a := GetWindowSpot( Addendum.AlbisWinID := AlbisWinID() )                    	; akutelle Albisfenstergröße wird unter a abgelegt
		If (a.X <= -20) && (a.Y <= -20) && !Addendum.AlbisWasMinimizedLast {
			Addendum.AlbisWasMinimizedLast := true
			If IsObject(Addendum.Thread[1])
				Addendum.Thread[1].ahkPostFunction("ToolbarShowHide" , "Hide")
			return 0
		} else If (a.X > -20) && (a.Y > -20) {
			Addendum.AlbisWasMinimizedLast := false
			If IsObject(Addendum.Thread[1])
				Addendum.Thread[1].ahkPostFunction("ToolbarShowHide" , "") ; Toolbar wieder anzeigen
		}

	; Albisfenstergröße herstellen, Albis kann aber weiterhin auf dem Bildschirm verschoben werden
		If ((a.W <> AlbisW) || (a.H <> AlbisH)) {   ; ein minimiertes Albis hat eine x, y Position von -32000

			X := StrLen(AlbisX) > 0 ? (AlbisX + 10)	: a.X
			Y := StrLen(AlbisY) > 0 ? (AlbisY - 6)  	: a.Y

			If IsRemoteSession() {
				factor 	:= A_ScreenDPI / 96
				W	:= Round(AlbisW * factor)
				H	:= Round(AlbisH * factor)
			}
			else
				W := AlbisW, H := AlbisH

			SetWindowPos(Addendum.AlbisWinID, X, Y, W, H)

		}

	; Addendum Infofenster neu zeichnen
		If WinExist("ahk_id " hadm) {
			WinSet, Style, 0x50020000, % "ahk_id " hadm
			WinSet, Top                   	,, % "ahk_id " hadm
			WinSet, Redraw             	,, % "ahk_id " hadm
			WinActivate, % "ahk_id " hadm
			WinActivate, % "ahk_id " Addendum.AlbisWinID
		}

return 0
}

ActiveFocusEventProc(hHook, event, hwnd) {

	static excludeFilter := ["OptoAppClass"]

	If (GetDec(hwnd) = 0)
		return 0
	cFocus	:= GetClassName((hwnd := Format("0x{:x}", hwnd)))
	;wclass	:= GetClassName((hWin := WinExist("A")))
	;~ If !(classList:= GetParenClasstList(hwnd))
		;~ return

	For idx, exclass in excludeFilter
		If InStr(classList, exclass)
			return 0

	;~ SciTEOutput("classList: " classList)

}

ShellHookProc(lParam, wParam) {                                                                                  	; Startet oder entfernt einen WinEventHook bei Erscheinen oder Schließen des Albisprogramms

		; diese Fensterklassen werden ignoriert. Dies beschleunigt die Funktion, da Autohotkey als Skriptsprache zu langsam ist, um den Message-Queue schnell genug abzuarbeiten
		static excludeFilter	:= [ "Windows.UI", "NVOpen", "QTrayIcon", "SunAwt", "qtopen", "Scite", "ThunderRT", "QT5Q", "Shell Embedding", "IME", "TForm"
										 , "TApplication", "DDEM", "COMTASKSWINDOWCLASS", "OfficePowerManagerWindow", "GDI+", "WICOMCLASS", "OleDdeWndClass"
										 , "GlassWndClass-", "Shell_TrayWnd", "Ghost", "ApplicationFrameWindow", "DAXParkingWindow", "EdgeUiInputTopWndClass"
										 , "DummyDWMListenerWindow", "Progman", "ThumbnailDeviceHelperWnd", "tooltips_class32", "MsiDialogCloseClass", "MsiHiddenWindow"
										 , "HwndWrapper", ".NET-BroadcastEventWindow", "CtrlNotifySink", "OperationStatusWindow", "ThorConnWndClass", "Shell Preview Extension" ]
		static includeFilter 	:= [ "OptoAppClass", "#32770", "AutohotkeyGui", "OpusApp" ]
		static PDFViewer 	:= [ "classFoxitReader", "SUMATRA_PDF_FRAME" ]
		static func_winpos

		global AlbisID_old

		Critical	;, 50


	; return on empty callback parameters ;{
		If (wParam = 0)
			return 0

		If (StrLen(class := WinGetClass(wparam)) = 0)
			class := GetClassName(wParam)

		Title	:= WinGetTitle(wparam)
		If (StrLen(Class . Title) = 0)
			return 0

		; filtert hier die Fensterklassen raus welche im Array excludeFilter vorhanden sind
		For i, exclude in excludeFilter
			If InStr(class, exclude)
				return 0
	;}

		SHookHwnd := Format("0x{:x}", wparam)

	; SHELL_WINDOWCREATED,WINDOWACTIVATED,REDRAW,RUDEAPPACTIVATED
		If  lParam in 1,4,6,32772           ; ACHTUNG: "{" darf nicht auf gleicher Zeile sein bei 'var in comma-separated lists'
		{

			; Albis hat Vorrang
				AlbisID := Format("0x{:x}", WinExist("ahk_class OptoAppClass"))
				If (AlbisID <> AlbisID_old) {
					AlbisID_old := AlbisID
					AlbisStartHooks()                                                                            ; startet Hooks
					PraxTT("Ein neuer Albisprozeß wurde erkannt.`nDie Fensterhooks wurden gesetzt!", "2 0")
				}

			; InfoFenster prüfen
				If Addendum.AddendumGui
					If (AlbisID = SHookHwnd) && RegExMatch((Addendum.AktiveAnzeige	:= AlbisGetActiveWindowType()), "i)Karteikarte|Laborblatt|Biometriedaten")
						If !WinExist("ahk_id " hadm)
							AddendumGui()
						else
							RedrawWindow(hadm)

			; andere Programme (MS Word, Ifap, FoxitReader) - Einstellungen sollten in der Addendum.ini stehen
				For classnn, appName in Addendum.Windows.Proc {
					; Fensterklassse ist bekannt und das Fenster soll in dieser Auflösung positioniert werden
						If InStr(classnn, class) {
							If (Addendum.Windows[appname].AutoPos & Addendum.MonSize)  {
								func_winpos := Func("CorrectWinPos").Bind(SHookHwnd, Addendum.Windows[appname][Addendum.Resolution])
								SetTimer, % func_winpos, -400    ; dem Fenster etwas Zeit geben um sich zu zeichnen, erst danach verschieben
								break
							}
						}
				}

		}
		else if (lParam = 2) && !WinExist("ahk_class OptoAppClass")	{	                 	; ALBIS wurde beendet

			admGui_Destroy()
			AlbisStopHooks()
			PraxTT("Albis wurde beendet. Hooks wurden entfernt.", "2 0")
			AlbisID_old := 0

		}

	; --------------------------- PopUpMenuCallback -----------------------------
	; für Addendum_PopUpMenu - ruft die als String in der Variable .PopUpMenuCallback Funktion auf und übergibt zusätzliche Parameter
		If (StrLen(Addendum.PopUpMenuCallback) > 0) {

			For idx, viewerclass in PDFViewer
				If InStr(class, viewerclass) {

					If InStr(Addendum.PopUpMenuCallback, "|") {

							callbackParam := StrSplit(Addendum.PopUpMenuCallback, "|")
							If IsFunc(callbackParam[1])
								func_Call 	:= Func(callbackParam[1]).Bind(callbackParam[2], class, SHookHwnd)

					}
					else {

							fnName := Addendum.PopUpMenuCallback
							If IsFunc(fnName)
								func_Call 	:= Func(fnName).Bind(class, SHookHwnd)

					}

					Addendum.PopUpMenuCallback := ""
					SetTimer, % func_Call, -300
					break

				}

	}

return
}

WaitEH_WinHandler:                                                                                                     	;{

	If (EHStack.Count() > 0) && !EHWStatus {
		SetTimer, WaitEH_WinHandler, Off
		gosub EventHook_WinHandler
	} else if (EHStack.Count() = 0)
		SetTimer, WaitEH_WinHandler, Off

return ;}

; ------------------------------------ Automatisierungsroutinen / Fensterhandler
EventHook_WinHandler:                                                                                                	;{ Eventhookhandler - Popupblocker/Fensterhandler - für diverse Fenster verschiedener Programme

	;~ If EHWHStatus
		;~ return

	StackEntry:
	EHWHStatus :=true
	thisWin := EHStack.Pop()
	EHWT := thisWin.title, EHWText := StrReplace(thisWin.Text, "`n", " "), EHWClass := thisWin.Class
	Addendum.LastHookHwnd := hHookedWin := GetHex(thisWin.Hwnd)
	;SciTEOutput(EHWText)

	If ((StrLen(EHWT . EHWText) = 0) || InStr(EHWText, "SkinLoader") || !WinExist("ahk_id " hHookedWin))
		If (EHStack.Count() = 0) {
			EHWHStatus := false
			Addendum.LastHookHwnd := 0
			return
		}
		else
			goto StackEntry

	;~ SciTEOutput(EHStack.Count()
			;~ .	"`t#E: " 	thisWin.Event         	 SubStr("       ", 1, 7 - StrLen(thisWin.Event))
			;~ .	"`t#H: " 	thisWin.hwnd         	 SubStr("                                ", 1, 10 - StrLen(thisWin.hwnd)) "    "
			;~ . 	"`t#T: " 	SubStr(EHWT, 1, 20)	 SubStr("                                ", 1, 20 - (StrLen(EHWT) > 20      	? 20 : StrLen(EHWT)*0.5)) (StrLen(EHWT) >= 20 ? " " : "    ")
			;~ . 	"`t#C: "	EHWClass                 	 SubStr("                                ", 1, 20 - (StrLen(EHWClass) > 20	? 20 : StrLen(EHWClass)*0.7)) "    "
			;~ .	"`t#X: " 	StrReplace(SubStr(EHWText, 1,80), "`n", " "))

	WinGet, EHproc1, ProcessName, % "ahk_id " hHookedWin
	If (EHWT && !EHproc1) {
		WinGet, EHproc1, ProcessName, % "ahk_id " (GetParent(hHookedWin))
		If !EHproc1 && (EHStack.Count() = 0) {
			EHWHStatus := false
			return
		} else
			goto StackEntry
	}

	If        InStr(EHproc1, "albis")                                                                      	{

			If   	  InStr(EHWText	, "ist keine Ausnahmeindikation")                                                              	 {  	; Fenster wird geschlossen
					VerifiedClick("Button2", hHookedWin)
					GNRChanged 	:= true
			}
			else If InStr(EHWT  	, "Daten von") && GNRChanged                                                              	 {  	; schließt nach Änderung der GNR das Fenster "Daten von " für schnelleres Handling
					while (A_Index < 30)
						Sleep 10
					If !WinExist("", "ist keine Ausnahmeindikation")  {
						VerifiedClick("Button30", "Daten von ahk_class #32770")
						GNRChanged 	:= false
					}
			}
			else If InStr(EHWText	, "ALBIS wartet auf Rückgabedatei")                                                          	 {  	; Laborhinweis schliessen
					BlockInput, On
					VerifiedClick("Button1", hHookedWin)
					AlbisActivate(1)
					SendInput, {Tab}
					WinActivate, % "ahk_class #32770", % "ALBIS wartet auf Rückgabedatei"
					BlockInput, Off
			}
			else If InStr(EHWText	, "Chipkarte")                                                                                          	 {  	; in Abhängigkeit des Client wird das Fenster sofort oder verzögert geschlossen
					If !InStr(compname, "SP1")
						Sleep, 5000
					VerifiedClick("Button1", hHookedWin)
			}
			else if Instr(EHWText	, "Der Parameter ist bereits in dieser Gruppe")                                         	 {  	; Fenster wird geschlossen
					VerifiedClick("Button1", hHookedWin)
			}
			else if Instr(EHWText	, "Behandlungszeitraum")                                                                         	 {  	; Behandlungszeitraum überschritten, neue Rechnung erstellen
					VerifiedClick("Button2", "ALBIS", "Behandlungszeitraum")
			}
			else if Instr(EHWT   	, "Patientenausweis")                                                                                 	 {  	; automatisch drucken
					VerifiedClick("Button23", hHookedWin)
			}
			else if InStr(EHWText	, "folgende Diagnosen übernommen")                                                       	 {  	; Fenster wird geschlossen
					VerifiedClick("Button1", hHookedWin)
			}
			else if InStr(EHWT   	, "ALBIS - Login") && AutoLogin                                                               	 {  	; Client spezifisches automatisches Einloggen in Albis
					AlbisAutoLogin(10)
			}
			else If InStr(EHWText	, "Sie haben eine Besuchsziffer ohne Wege")                                             	 {  	; Fenster wird geschlossen
					VerifiedClick("Button2", "Sie haben eine Besuchsziffer ohne Wege")
			}
			else If InStr(EHWText	, "Datum darf nicht vordatiert werden")                                                     	 {  	; Datum wird verändert, Daten werden eingetragen, Datum wird zurückgesetzt
					VerifiedClick("Button1", "ALBIS ahk_class #32770", "Datum darf nicht")
					fControlHwnd:= GetFocusedControlHwnd()
					ControlGetText, VorDatierDatum, Edit2, ahk_class OptoAppClass
					AlbisSetzeProgrammDatum(VorDatierDatum)
					VerifiedSetFocus("", "", "", fControlHwnd)
					SendInput, {Tab}
			}
			else If InStr(EHWText	, "Wollen Sie wirklich die GNR ändern")                                              		 {  	; es wird automatisch mit "Ja" bestätigt
					VerifiedClick("Button1", "ALBIS ahk_class #32770", "Wollen Sie wirklich")
			}
			else If InStr(EHWText	, "Der Patient wurde bereits als verstorben")                                               	 {  	; es wird automatisch mit "Ja" bestätigt
					VerifiedClick("Button1", hHookedWin)
			}
			else If InStr(EHWText	, "Fehler beim Aufruf dppivassist")                                                             	 { 	; Fenster wird geschlossen
					VerifiedClick("Button1", hHookedWin)
			}
			else If InStr(EHWText	, "Folgender Pfad existiert nicht")                                                               	 { 	; Fenster wird geschlossen
					VerifiedClick("Button1", hHookedWin)
			}
			else if Instr(EHWT   	, "Herzlichen Glückwunsch") && WinExist("Gesundheitsvorsorgeliste")        	 { 	; Fenster "Herzlichen Glückwunsch" schliessen
					VerifiedClick("Button1", hHookedWin)
			}
			else if InStr(EHWText	, "Der Druckauftrag konnte nicht gestartet werden")                                    	 { 	; Fenster wird geschlossen
					VerifiedClick("Button1", hHookedWin)
			}
			else if InStr(EHWT   	, "CGM LIFE")	                                                                                    	 {  	; Verbindungsfehler (gesperrt durch Windows Firewall)
					WinClose, % "CGM LIFE ahk_class Afx:00400000:0:00010005:86101803"
			}
			else if InStr(EHWT  	, "CGM HEILMITTELKATALOG")                                                              	 { 	; Fensterposition wird innerhalb des Albisfenster zentriert
						AlbisHeilmittelKatologPosition()
			}
			else if InStr(EHWT  	, "Dauerdiagnosen von")                                                                        	 { 	; automatisches Sortieren von anamnestischen und Behandlungsdiagnosen
					res:= AlbisResizeDauerdiagnosen()
			}
			else if InStr(EHWT  	, "ICD-10 Thesaurus")                                                                            	 { 	; vergrößert den Diagnosenauswahlbereich
					UpSizeControl("ICD-10 Thesaurus", "#32770", "Listbox1", 200, 150, AlbisWinID())
			}
			else if InStr(EHWT  	, "Muster 1a") || InStr(EHWText, "Der Patient ist noch AU")                       	 { 	; Arbeitsunfähigkeitsbescheinigung - Einblendung von zusätzlichen Informationen
					AlbisFristenGui()
			}
			else If InStr(EHWText	, "Fehler im ifap praxisCENTER")	                                                            	 {    	; Fehlerbenachrichtungen Java Runtime des ifap praxisCENTER

					VerifiedClick("Button1", hHookedWin)
					If !ifaptimer
						SetTimer, ifapActive, 30
					ifaptimer:= A_TickCount

			}
			else If InStr(EHWText	, "Bitte die Straße")                                                                                  	 {  	; weitere Information Dialog (Patientendaten) und Adressproblematik

					; bei Adressen mit Postleitzahlen, kann das Fenster 'weitere Informationen' nicht einfach geschlossen werden
					MsgBox, 0x1024, Addendum für Albis on Windows, Rechnungsempfänger löschen?
					IfMsgBox, Yes
					{
						Loop 8
							VerifiedSetText("Edit" A_Index, "", "ahk_class #32770", 200, "Adresse des Rechnungs")
					}
					else
						VerifiedSetText("Edit5", " ", "ahk_class #32770", 200, "Adresse des Rechnungs")

					VerifiedClick("Button2", "ahk_class #32770", "Adresse des Rechnungs")

			}
			else if Addendum.Schnellrezept && InStr(EHWT, "Muster 16")                                                     	 {  	; Kassenrezept - Schnellrezept + Auto autIdem Kreuz
					AlbisRezeptHelferGui(Addendum.AdditionalData_Path "\RezepthelferDB.json")
			}
			;-------------------- Laborabruf Automation --------------------------------------------------------
			else If Addendum.Labor.AutoAbruf && InStr(EHWText	, "Anforderungen ins Laborblatt übertragen"){ 	; es wird automatisch mit "Ja" bestätigt
					VerifiedClick("Button1", hHookedWin)
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWText	, "Keine Datei(en) im Pfad") 	                   	 { 	; Laborabrufvorgang wird beendet
					ResetLaborAbrufStatus()
					VerifiedClick("Button1", hHookedWin)
			}
			else if Addendum.Labor.AutoAbruf && Instr(EHWT  	, "Labor auswählen")                                 	 {
					AlbisLaborAuswählen(Addendum.Labor.LaborName)
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWT  	, "Labordaten")                                        	 {
					AlbisLaborDaten()
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWT  	, "GNR der Anford")             	                 	 {
					AlbisRestrainLabWindow(1, hHookedWin)       	;Listbox1 enthält Rechnungsdaten
			}

	}
	else if InStr(EHproc1, "mDicom")             	&& (Addendum.mDicom = true)	{

			If InStr(EHWT, "Export to video")
				MicroDicom_VideoExport()

	}
	else if InStr(EHproc1, "ipc")                                                                          	{      	; ifap Programm

			If InStr(EHWText, "Fehler im ifap praxisCENTER") {
				VerifiedClick("Button1", hHookedWin)
				If !ifaptimer
					SetTimer, ifapActive, 50
				ifaptimer:= A_TickCount
			}

	}
	else if InStr(EHproc1, "infoBoxWebClient") 	&& (Addendum.Labor.AutoAbruf) 	{        	; fängt das WebFenster meines Labors ab

			IniWrite, % A_DD "." A_MM "." A_YYYY " (" A_Hour ":" A_Min ":" A_Sec ")", % Addendum.AddendumIni, Labor, letzter_Laborabruf
			If !Addendum.Laborabruf.Status {

				Addendum.Laborabruf.Status := "Init"
				fcRLStat := Addendum.LaborAbruf.Reset := Func("ResetLaborAbrufStatus")
				SetTimer, % fcRLStat, -600000   ; 10 min
				PraxTT("Das Laborabruf Fenster wurde detektiert!`n#3Addendum übernimmt den weiteren Vorgang!", "40 4")
				vianovaInfoBoxWebClient(hHookedWin)

			}

	}
	else if Instr(EHproc1, "Scan2Folder")	                                                        	{         	; Fujitsu Scansnap Software

			If Instr(EHWText, "Die Dateien wurden erfolgreich gespeichert")
					VerifiedClick("Button1", hHookedWin)

	}
	else if Instr(EHproc1, "Autohotkey")                                                             	{         	; WinSpy (Autohotkey Tool) und Debuggerfenster

			If InStr(EHWT, "WinSpy")
				MoveWinToCenterScreen(hHookedWin)
			else if InStr(EHWT, "Variable list")
				WinMoveZ(hHookedWin, 0, Floor(A_ScreenWidth/2), 300, 600, 1200)

	}
	else if Instr(EHproc1, "InternalAHK")                                                            	{       	; SciTE Editor - Variablenfenster verschieben

	    	if InStr(EHWT, "Variable list")
				WinMoveZ(hHookedWin, 0, Floor(A_ScreenWidth/2), 300, 600, 1200)

	}
	else if Instr(EHproc1, "SciTE")           	                                                        	{         	; SciTE Editor

			If InStr(EHWT, "SciTE4AutoHotkey") && Instr(WinGetClass(hHookedWin), "#32770") && RegExMatch(EHWText, "Datei.*Addendum\.ini")			{
				; nach dem Laden einer veränderten Addendum.ini Datei - wird diese von SciTE immer entfaltet.
				; Das macht alles unübersichtlich. Deshalb wird hier automatisches Codefolding durchgeführt
					VerifiedClick("Button1", "SciTE4AutoHotkey ahk_class #32770")
					SciTeFoldAll()
			}

	}
	else if InStr(EHproc1, "FoxitReader")                                                             	{        	; Foxitreader signieren vereinfachen (und zwar tatsächlich!)

			If Addendum.PDFRecentlySigned {

					if InStr(EHWT, "Speichern unter") && RegExMatch(EHWText, "i)Speichern|Save", gotText) {

						Addendum.PDFRecentlySigned := false
						thisFoxitID := GetParent(hHookedWin)                          	; Handle des zugehörigen FoxitReader Fenster
						;SciTEOutput("foxitClass: " foxClass)
						If VerifiedClick("Speichern", hHookedWin,,,1)                 	; Speichern Button drücken
							Addendum.SaveAsInvoked := true
						WinWait, % "Speichern unter bestätigen" ,, 2
						If !ErrorLevel
							VerifiedClick("Ja", "Speichern unter bestätigen",,, 1)    	; mit 'Ja' bestätigen
						;Sleep, 200
						;If Addendum.AutoCloseFoxitTab                                	; zu früh ausgelöst stürzt FoxitReader ab und die PDF Datei ist defekt
						;	FoxitInvoke("Close", thisFoxitID)

					}

					;SciTEOutput("Hab Dich: " EHWT "`trxMatch: " gotText)
			}

			If Addendum.PDFSignieren && (InStr(EHWT, "Sign Document") || InStr(EHWT, "Dokument signieren"))          	; für die englische und deutsche FoxitReader Version
				FoxitReader_SignDoc(hHookedWin)

	}
	else if InStr(EHproc1, "iexplore")                                                                  	{       	; Windows Dateiexplorer
			If (StrLen(Addendum.Labor.Kennwort) > 0) && InStr(EHWT, "CGM CHANNEL: Login") { ; CGM CHANNEL: Login

					; 4.3.4.1.4.3.4.1.4.1 - Pfad für URL-Eingabe
					; KennwortBez := "4.5.4.3.4.1.4.1.1.2.4" ;8
					; Kennwortfeld := "4.5.4.3.4.1.4.1.1.2.4.9"
					; KennwortEnter := "4.5.4.3.4.1.4.1.1.2.4.10"

					SendEvent, % "{Text}" Addendum.Labor.Kennwort
					Sleep 500
					SendInput, {Enter}

					PraxTT("CGM Channel Login wurde durchgeführt.", "6 2")

			}
	}

	If (EHStack.Count() > 0)
		goto StackEntry

	Addendum.LastHookHwnd := 0
	EHWHStatus := false

return

ifapActive:                  	;{ holt das ifap Fenster (nach der hoffentlich letzten Fehlermeldung) in den Vordergrund
	if (A_TickCount - ifapTimer > 200) {
		ifaptimer := 0
		WinActivate, praxisCENTER ahk_class TForm_ipcMain
		SetTimer, ifapActive, off
	}
return ;}


;}

WinEventHook_Helper_Thread:                                                                                      	;{ Labordaten - Fenster abfangen

	hpop:= GetLastActivePopup(AlbisWinID())
	If (hpop != AlbisWinID()) && (hpop != hpop_old)	{

		EHWT	 := WinGetTitle(hpop)
		EHWText:= WinGetText(hpop)
		hpop_old:= hpop

		If InStr(EHWT, "Labordaten")		{

			SetTimer, WinEventHook_Helper_Thread, Off
			VerifiedCheck("Button5"	, hpop,,, 1)
			VerifiedClick("Button1" , hpop)
			AlbisOeffneLaborbuch()

		}

	}

return
;}

WinEvent_HMV(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {           	; Hookprocedure welche nur bei geöffneten Heilmittelverordnungsformularen aktiv ist
	Critical
	HmvHookHwnd:= GetHex(hwnd)
	HmvEvent:= GetHex(event)
	SetTimer, Heilmittelverordnung_WinHandler,  -0
return 0

Heilmittelverordnung_WinHandler:                                                                                 	;{ Eventhookhandler - kümmert sich nur um Heilmittelverordnungsfenster

return ;}
}

FEGui() {                                                                                                                       	; Fokushook Testgui

	global
	Gui, FE: New, +ToolWindow -Caption -DPIScale +AlwaysOnTop HWNDhFEGui
	Gui, FE: Add, Text, x2 y2 w500 vFEText1 HWNDhFEText1, FocusHook test
	;WinSet, Trans, 100, % "ahk_id " hFEGui
	Gui, FE: Show, x800 y0, FEGui

}

CorrectWinPos(hWin, wPos) {                                                                                         	; ein Fenster an eine bestimmte Position auf dem Bildschirm verschieben

	; prüft das ein Object übergeben wurde und gültige Werte/Schlüsselnamen enthält
		If !IsObject(wPos)
			return

	; Einstellungen prüfen
		w := GetWindowSpot(hWin)
		switch wPos.X
		{
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
					wPos.Y := 0
		}

		switch wPos.Y
		{
				case "-":
					wPos.Y := w.Y
		}

		switch wPos.W
		{
				case "F":
					wPos.W := A_ScreenWidth
		}

		switch wPos.H
		{
				case "F":
					wPos.H := A_ScreenHeight
		}


	; Fenster versetzen
		If (w.X != wPos.X) || (w.Y != wPos.Y) || (w.W != wPos.W) || (w.H != wPos.H) {
			;SciTEOutput(Addendum.MonSize "(" Addendum.Resolution ") :" wPos.X ", " wPos.Y ", " wPos.W ", " wPos.H)
			SetWindowPos(hWin, wPos.X, wPos.Y, wPos.W, wPos.H)
		}
}

CurrPatientChange(AlbisTitle) {                                                                                      	; behandelt Änderung des Albisfenstertitels

		global hadm, admHJournal
		static 	AlbisTitleO

		If (AlbisTitleO = AlbisTitle)
			return

		Addendum.AktiveAnzeige	:= AlbisGetActiveWindowType()
		                	 AlbisTitleO	:= AlbisTitle

		If InStr(AlbisTitle, "Prüfung EBM/KRW") && !InStr(AlbisTitle, "Abrechnung vorbereiten")	{
			AlbisActivate(1)
			Sleep 6000
			SendInput, {Esc}
			return
		}

	; AddendumGui verbergen wenn Bedingungen nicht erfüllt sind                                             ;!RegExMatch(Addendum.AktiveAnzeige, "i).*\|(Karteikarte)|(Laborblatt)")
		If !Addendum.AddendumGui || !WinExist("ahk_class OptoAppClass") {
			admGui_Destroy()
		}
		else if Addendum.AddendumGui {

			If RegExMatch(Addendum.AktiveAnzeige, "i)Karteikarte|Laborblatt|Biometriedaten")
				AddendumGui()
			else If RegExMatch(Addendum.AktiveAnzeige, "i)Abrechnung|Rechnungsliste") {
				admGui_Destroy()
			}
			else {
				admGui_Destroy()
			}

		}


	; Patientenakte geöffnet - dann Überprüfung des Patientennamen und Erstellen der AddendumGui
		If      	InStr(Addendum.AktiveAnzeige, "Patientenakte")     		{
			PatDb(AlbisTitle2Data(AlbisTitle), "exist")        ; aktuellen Patienten registrieren bei der Addendum Patientendatenbank
		}
		else If 	InStr(Addendum.AktiveAnzeige, "aPEK")	                 	{
			MsgBox, 4, Addendum für Albis on Windows, Hinweisdialog zur Prüfung EBM/KRW weiter ansehen?, 6
			IfMsgBox, Yes
				return
			AlbisCloseMDITab("Prüfung EBM/KRW")
		}
		else if 	InStr(Addendum.AktiveAnzeige, "Laborbuch") && (Addendum.Laborabruf_Voll)
			Albismenu(34157, "", 6, 1)			;34157 - alle übertragen

Return
}


; ------------------------------------ zugehörige Funktionen für die WinEventHandler Labels
SplashAus: ;{
	SplashTextOff
return ;}

;}

;{16. Interskriptkommunikation / OnMessage
Receive_WM_COPYDATA(wParam, lParam) {                                                             	;-- empfängt Nachrichten von anderen Skripten die auf demselben Client laufen

    StringAddress := NumGet(lParam + 2*A_PtrSize)
	fn_MsgWorker := Func("MessageWorker").Bind(StrGet(StringAddress))
	SetTimer, % fn_MsgWorker, -10
	;MessageWorker(StrGet(&StringAddress))

return
}

MessageWorker(InComing) {                                                                                     	;-- verarbeitet die eingegangen Nachrichten

		global AutoLogin
		;SciTEOutput(" - [Addendum] Message received: " InComing)

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}

	; Kommunikation mit dem Abrechnungshelfer Skript
		If RegExMatch(recv.txtmsg, "^AutoLogin\s+Off")	{
			AutoLogin := false
			result := Send_WM_COPYDATA("AutoLogin disabled", recv.opt)
		}
		else if RegExMatch(recv.txtmsg, "^AutoLogin\s+On")	{
			AutoLogin := true
			result := Send_WM_COPYDATA("AutoLogin enabled", recv.opt)
		}
	; Tesseract OCR Thread Kommunikation
		else if InStr(recv.txtmsg, "OCR_processed") {
			admGui_CheckJournal(recv.opt, recv.fromID)
		}
		else if InStr(recv.txtmsg, "OCR_ready") {
			admGui_Journal()
			tessOCRRunning := false
		}
return
}

WM_DISPLAYCHANGE(wParam, lParam) {                                                                 	;-- einsetzbar für Script-Editoren - AutoZoom bei Wechsel der Bildschirmauflösung

	; zoomt bei einer Auflösung > 1920 Scite4AHK (RDP Session)
	; letzte Änderung: 27.09.2020

	static lastScreenSize := ""
	static ZoomScintilla1O, ZoomScintilla2O
	static SCI_SETZOOM:=2373, SCI_GETZOOM:=2374
	static ZoomLevel := {"4k"	: {"Sci1":"1"	, "Sci2":"-2"}
								, 	"2k"	: {"Sci1":"-1"	, "Sci2":"-3"}}

	If (hScite := WinExist("ahk_class SciTEWindow")) {

		SPos := GetWindowSpot(hScite)
		SysGet SciteMon_, Monitor, % GetMonitorAt(SPos.X, SPos.Y)
		monDim := SciteMon_Right "x" SciteMon_Bottom

		If (lastScreenSize <> monDim) {

			lastScreenSize := monDim

			ControlGet	, hScintilla1	, hwnd,, Scintilla1	, % "ahk_id " hScite
			ControlGet	, hScintilla2	, hwnd,, Scintilla2	, % "ahk_id " hScite

			SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla1
			ZoomScintilla1O := ErrorLevel
			SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla2
			ZoomScintilla2O := ErrorLevel

			SendMessage, SCI_SETZOOM, % (SciteMon_Bottom > 1080 ? ZoomLevel.4k.Sci1 : ZoomLevel.2k.Sci1), 0,, % "ahk_id " hScintilla1
			SendMessage, SCI_SETZOOM, % (SciteMon_Bottom > 1080 ? ZoomLevel.4k.Sci2 : ZoomLevel.2k.Sci2), 0,, % "ahk_id " hScintilla2
			WinMaximize, ahk_class SciTEWindow,, 0, 0

		}

	}

	If (hSpy := WinExist("WinSpy"))
		MoveWinToCenterScreen(hSpy)

}

;}

;{17. Automatisierungsfunktionen

;Funktion AutoDiagnose findet sich in include\AutoDiagnosen.ahk

AddAutoComplete(lCHwnd, ControlName, TextList, TWidth:= 350) {

	If (TextList="")
			Return

	ControlGetPos, CpX, CpY,,, % ControlName, ahk_id %lCHwnd%
	CpY += 20

	Gui, AutoC: new		, -Caption +ToolWindow +AlwaysOnTop
	Gui, AutoC: Margin	, 0, 0
	Gui, AutoC: Add		, Listbox, % "r" 10 " w" TWidth " vLBAutoC" , % TextList
	Gui, AutoC: Show		, % "x" CpX " y" CpY, Addendum AutoComplete

	HotKey, IfWinExist, Addendum AutoComplete
	HotKey, Enter, MedTLB
	HotKey, Esc, AutoCGuiEscape

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

	If !ADocInit
	{
		ADoc:= GetFromIni("MaxResults|OffsetX|OffsetY|BoxHeight|ShowLength|Font|FontSize|MinScore", "AutoDoc")
		ADocList["Impferrinnerung"]    		:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "impf"	, "Do": "Choose"	, "RichE": "< /TdPP>< /, Pneumovax>< /, GSI>< /, MMR>"                                        	 }
		ADocList["Impfausweis"]     	    	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "impf"	, "Do": "Set"		, "RichE": "Impfausweis mitbringen"                                                                                	 }
		ADocList["lko - Gespräch"]	        	:= {"Edit2": "="							, "Edit3": "lko"	, "Do": "Input"	, "RichE": "03230(x:*)"                                                                                						 }
		ADocList["VorsorgeColo"]    	    	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "info"	, "Do": "Set"		, "RichE": "an die Vorsorge Coloskopie errinnern"                                        						 }
		ADocList["In die Sprechstunde"]   	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "info"	, "Do": "Set"		, "RichE": "Es wäre schön wenn der Patient in diesem Jahr in mein Sprechzimmer vorstößt!"}
		For key, val in ADocList
			LBList.= key "|"
		Gui, Suggestions: -Caption +ToolWindow +AlwaysOnTop HWNDhSugg
		Gui, Suggestions: Font, % "s" ADoc.FontSize, % ADoc.Font
		Gui, Suggestions: Add, ListBox, % "x0 y0 h" ADoc.BoxHeight " 0x100 Sort HWNDhSuggLB1 vMatched", % LBList
		ADocInit:=1
	}

	Gui, Suggestions: Show, % "x" (saX+ADoc.OffsetX) " y" (saY - ADoc.OffsetY - ADoc.BoxHeight) " h" ADoc.BoxHeight, AutoDoc
	WinActivate, ahk_id %hSugg%
	GuiControl, Choose, matched, 1
	ControlFocus,, ahk_id %hSuggLB1%

	HotKey, IfWinExist      	, AutoDoc
	HotKey, Enter            	, FillDocLine
	HotKey, NumPadEnter	, FillDocLine
	HotKey, Esc               	, SuggestionsGuiClose
	Hotkey, IfWinExist

Return

FillDocLine:
	Gui, Suggestions: Submit
	actrl:= AlbisGetActiveControl("content")

	If !(ADocList[(matched)]["Edit2"]="=")
	{
				n:= (A_MM=12) ? 1 : 0
				replacement:= StrReplace(ADocList[(matched)]["Edit2"], "%A_YYYY%", A_YYYY + n)		;<-später ein RegExMatch oder Replace für mehr Optionen
				VerifiedSetText("Edit2", replacement , "ahk_class OptoAppClass", 100)
	}

	VerifiedSetText("Edit3", ADocList[(matched)]["Edit3"] , "ahk_class OptoAppClass", 100)

	If (ADocList[(matched)]["Do"]="Choose")
	{
				VerifiedSetText("RichEdit20A1", ADocList[(matched)]["RichE"] , "ahk_class OptoAppClass", 100)
				ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
				Sleep, 200
				SendInput, {LControl Down}{Right}{LControl Up}
	}

	If (ADocList[(matched)]["Do"]="Set")
	{
				VerifiedSetText("RichEdit20A1", ADocList[(matched)]["RichE"] , "ahk_class OptoAppClass", 100)
				ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
				Sleep, 100
				SendInput, {Tab}
	}

	If (ADocList[(matched)]["Do"]="Input")
	{
				InputBox, faktor, Addendum für AlbisOnWindows, bitte tragen Sie hier den faktor für die Gesprächsziffer ein.,,,,,,,, 2
				VerifiedSetText("RichEdit20A1", StrReplace(ADocList[(matched)]["RichE"], "*", faktor), "ahk_class OptoAppClass", 100)
				ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
				Sleep, 100
				SendInput, {Tab}
	}

	gosub SuggestionsGuiEscape
return

SuggestionsGuiClose:
	Gui, Suggestions: Hide
SuggestionsGuiEscape:
	;HotKey, Enter            	, Off
	;HotKey, NumpadEnter	, Off
	;HotKey, Esc	             	, Off
	AutoDocCalled:=0
Return

}

GetFromIni(Settings, Section) {

	tempArr 	:= []
	tempObj 	:= Object()
	tempArr 	:= StrSplit(Settings, "|")
	Loop, % tempArr.Count()
	{
			IniRead, var, % AddendumDir "\Addendum.ini", % Section, % tempArr[A_Index]
			tempObj[(tempArr[A_Index])] := var
	}

return tempObj
}

MedTrenner(f) {
		static MedTrenner, MedTTrenner, x, y
		global MedT

	; Standardschrift im Albisprogramm zur Anzeige der Dauermedikamente ist Arial Standard 11 - am besten mit dieser Schriftart editieren und hier einfügen
		If (MedTrenner="") {
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
		}

		x:= A_CaretX, y:= A_CaretY + 15

		Gui, MedT: new		, -Caption +ToolWindow +AlwaysOnTop
		Gui, MedT: Margin	, 0, 0
		Gui, MedT: Add		, Listbox, r20 w350 vMedTTrenner , % MedTrenner
		Gui, MedT: Show		, % "x" x " y" y, Dauermedikamten-Trenner

		HotKey, IfWinExist, Dauermedikamten-Trenner
			HotKey, Enter, MedTLB
			HotKey, Esc	, MedTGuiEscape

Return

MedTLB:

		Gui, MedT: Submit, NoHide
		WinActivate, Dauermedikamente ahk_class #32770
			WinWaitActive, Dauermedikamente ahk_class #32770,, 2
		y-= 15
		Click, %x%, %y%, 2
				Sleep, 200
		SendRaw, %MedTTrenner%
		SendInput, {Enter}

MedTGuiClose:
MedTGuiEscape:

		Gui, MedT: Destroy

return
}

;}

;{18. PDF-Reader Automation (Sumatra, FoxitReader)

;Sumatra PDF Viewer
SumatraInvoke(command, SumatraID="") {                                                                     	;-- wm_command wrapper for Sumatra PDF Version: 3.1

	/* DESCRIPTION of FUNCTION:  SumatraInvoke() by Ixiko (version 12.07.2020)

		---------------------------------------------------------------------------------------------------
								   a WM_command wrapper for Sumatra Pdf-Reader V3.1* by Ixiko
																		...........................................................
													 Remark: maybe not all commands are listed at now!
		---------------------------------------------------------------------------------------------------
		        by use  of a valid SumatraID, this function will post your command to Sumatra
			                                             otherwise this function returns the command code
																		...........................................................
			Remark: You have to control the success of the postmessage command yourself!
		---------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		SumatraInvoke("Show_FullPage")        SumatraInvoke("Place_Signature", SumatraID)
		.................................................       ...................................................................
		this one only returns the Sumatra              sends the command "Place_Signature" to
        command-code                                             your specified Sumatra process using
																 parameter 2 (SumatraID) as window handle.
															 		          command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

	static SumatraCommands
	If !IsObject(SumatraCommands) {

		SumatraCommands := { 	"Open":                                 	400  	; File
											,	"Close":                                 	401  	; File
											,	"SaveAs":                               	402  	; File
											,	"Rename":                              	580  	; File
											,	"Print":                                   	403  	; File
											,	"SendMail":                           	408  	; File
											,	"Properties":                           	409  	; File
											,	"OpenLast1":                         	510  	; File
											,	"OpenLast2":                         	511  	; File
											,	"OpenLast3":                         	512  	; File
											,	"Exit":                                    	405  	; File
											,	"SinglePage":                         	410  	; View
											,	"DoublePage":                       	411  	; View
											,	"BookView":                           	412  	; View
											,	"ShowPagesContinuously":     	413  	; View
											,	"TurnCounterclockwise":         	415  	; View
											,	"TurnClockwise":                    	416  	; View
											,	"Presentation":                        	418  	; View
											,	"Fullscreen":                           	421  	; View
											,	"Bookmark":                          	000  	; View - do not use! empty call!
											,	"ShowToolbar":                      	419  	; View
											,	"SelectAll":                             	422  	; View
											,	"CopyAll":                              	420  	; View
											,	"NextPage":                           	430  	; GoTo
											,	"PreviousPage":                      	431  	; GoTo
											,	"FirstPage":                            	432  	; GoTo
											,	"LastPage":                            	433  	; GoTo
											,	"GotoPage":                          	434  	; GoTo
											,	"Back":                                  	558  	; GoTo
											,	"Forward":                             	559  	; GoTo
											,	"Find":                                   	435  	; GoTo
											,	"FitPage":                              	440  	; Zoom
											,	"ActualSize":                          	441  	; Zoom
											,	"FitWidth":                             	442  	; Zoom
											,	"FitContent":                          	456  	; Zoom
											,	"CustomZoom":                     	457  	; Zoom
											,	"Zoom6400":                        	443  	; Zoom
											,	"Zoom3200":                        	444  	; Zoom
											,	"Zoom1600":                        	445  	; Zoom
											,	"Zoom800":                          	446  	; Zoom
											,	"Zoom400":                          	447  	; Zoom
											,	"Zoom200":                          	448  	; Zoom
											,	"Zoom150":                          	449  	; Zoom
											,	"Zoom125":                          	450  	; Zoom
											,	"Zoom100":                          	451  	; Zoom
											,	"Zoom50":                            	452  	; Zoom
											,	"Zoom25":                            	453  	; Zoom
											,	"Zoom12.5":                          	454  	; Zoom
											,	"Zoom8.33":                          	455  	; Zoom
											,	"AddPageToFavorites":           	560  	; Favorites
											,	"RemovePageFromFavorites": 	561  	; Favorites
											,	"ShowFavorites":                    	562  	; Favorites
											,	"CloseFavorites":                    	1106  	; Favorites
											,	"CurrentFileFavorite1":           	600  	; Favorites
											,	"CurrentFileFavorite2":           	601  	; Favorites -> I think this will be increased with every page added to favorites
											,	"ChangeLanguage":               	553  	; Settings
											,	"Options":                             	552  	; Settings
											,	"AdvancedOptions":               	597  	; Settings
											,	"VisitWebsite":                        	550  	; Help
											,	"Manual":                              	555  	; Help
											,	"CheckForUpdates":               	554  	; Help
											,	"About":                                	551}  	; Help

	}

	If SumatraID
		PostMessage, 0x111, % SumatraCommands[command],,, % "ahk_id " SumatraID
	else
		return SumatraCommands[command]

}

Sumatra_GetPages(SumatraID="") {                                                                                	;-- aktuelle und maximale Seiten des aktuellen Dokumentes ermitteln

	If !SumatraID
		SumatraID := WinExist("ahk_class SUMATRA_PDF_FRAME")

	ControlGetText, PageDisp, Edit3 	, % "ahk_id " SumatraID
	ControlGetText, PageMax, Static3	, % "ahk_id " SumatraID
	RegExMatch(PageMax, "\s*(?<Max>\d+)", Page)

return {"disp":PageDisp, "max":PageMax}
}

Sumatra_ToPrint(SumatraID="", Printer="") {                                                                      	;-- Druck Dialoghandler - Ausdruck auf übergebenen Drucker

		; druckt das aktuell angezeigte Dokument
		; abhängige Biblitheken: LV_ExtListView.ahk

		static sumatraprint	:= "i)[(Print)|(Drucken)]+ ahk_class i)#32770 ahk_exe i)SumatraPDF.exe"

		rxPrinter:= StrReplace(Trim(Printer), " ", "\s")
		rxPrinter:= StrReplace(rxPrinter, "(", "\(")
		rxPrinter:= StrReplace(rxPrinter, ")", "\)")

		OldMatchMode := A_TitleMatchMode
		SetTitleMatchMode, RegEx                                                              	; RegEx Fenstervergleichsmodus einstellen

		SumatraInvoke("Print", SumatraID)                                                  	; Druckdialog wird aufgerufen
		WinWait, % sumatraprint,, 6                                                             	; wartet 6 Sekunden auf das Dialogfenster
		hSumatraPrint := GetHex(WinExist(sumatraprint))                               	; 'Drucken' - Dialog handle
		ControlGet, hLV, Hwnd,, SysListview321, % "ahk_id " hSumatraPrint    	; Handle der Druckerliste (Listview) ermittlen
		sleep 200                                                                                       	; Pause um Fensteraufbau abzuwarten
		ControlGet, Items	, List  , Col1 	,, % "ahk_id " hLV                             	; Auslesen der vorhandenen Drucker
		ItemNr := 0                                                                                    	; ItemNr auf 0 setzen
		Loop, Parse, Items, `n                                                                    	; Listview Position des Standarddrucker suchen
			If RegExMatch(A_LoopField, "i)^" rxPrinter) {                                	; Standarddrucker gefunden
				ItemNr := A_Index                                                                	; nullbasierende Zählung in Listview Steuerelementen
				break
			}
		If ItemNr {                                                                                    	; Drucker in der externen Listview auswählen
			objLV := ExtListView_Initialize(sumatraprint)                                	; Initialisieren eines externen Speicherzugriff auf den Sumatra-Prozeß
			ControlFocus,, % "ahk_id " objLV.hlv                                            	; Druckerauswahl fokussieren
			SciTEOutput(A_LineNumber)
			err	 := ExtListView_ToggleSelection(objLV, 1, ItemNr - 1)            	; gefundenes Listview-Element (Drucker) fokussieren und selektieren
			SciTEOutput("err1: " err)
			ExtListView_DeInitialize(objLV)                                                     	; externer Speicherzugriff muss freigegeben werden
			Sleep 200
			err	:= VerifiedClick("Button13", hSumatraPrint)                           	; 'Drucken' - Button wählen
			SciTEOutput("err: " err)
			WinWaitClose, % "ahk_id " hSumatraPrint,, 3                              	; wartet max. 3 Sek. bis der Dialog geschlossen wurde
		}

		SciTEOutput("ItemNr: " ItemNr ", ID: " hSumatraPrint)
		SetTitleMatchMode, % OldMatchMode                                            	; TitleMatchMode zurückstellen

return {"DialogID":hSumatraPrint, "ItemNr":ItemNr}                                 	; für Erfolgskontrolle und eventuelle weitere Abarbeitungen
}

;FoxitReader
FoxitInvoke(command, FoxitID="") {		                                                                        	;-- wm_command wrapper for FoxitReader Version:  9.1

		/* DESCRIPTION of FUNCTION:  FoxitInvoke() by Ixiko (version 11.07.2020)

		---------------------------------------------------------------------------------------------------
												a WM_command wrapper for FoxitReader V9.1 by Ixiko
																		...........................................................
													 Remark: maybe not all commands are listed at now!
		---------------------------------------------------------------------------------------------------
				by use  of a valid FoxitID, this function will post your command to FoxitReader
			                                             otherwise this function returns the command code
																		...........................................................
			Remark: You have to control the success of the postmessage command yourself!
		---------------------------------------------------------------------------------------------------
						I intentionally use a text first and then convert it to a -Key: Value- object,
                                                         so you can swap out the object to a file if needed
		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		FoxitInvoke("Show_FullPage")                       FoxitInvoke("Place_Signature", FoxitID)
		.................................................          ...............................................................
		this one only returns the Foxit                  sends the command "Place_Signature" to
        command-code                                      your specified FoxitReader process using
																	 parameter 2 (FoxitID) as window handle.
																	        command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

	static FoxitCommands
	If !IsObject(FoxitCommands) {

		FoxitCommands   := {	"SaveAs":                                                       	1299
										,	"Close":                                                         	57602
										,	"Hand":                                                         	1348        	; Home - Tools
										,	"Select_Text":                                                 	46178      	; Home - Tools
										,	"Select_Annotation":                                       	46017      	; Home - Tools
										,	"Snapshot":                                                    	46069      	; Home - Tools
										,	"Clipboard_SelectAll":                                    	57642      	; Home - Tools
										,	"Clipboard_Copy":                                         	57634      	; Home - Tools
										,	"Clipboard_Paste":                                         	57637      	; Home - Tools
										,	"Actual_Size":                                                 	1332        	; Home - View
										,	"Fit_Page":                                                     	1343        	; Home - View
										,	"Fit_Width":                                                   	1345        	; Home - View
										,	"Reflow":                                                        	32818      	; Home - View
										,	"Zoom_Field":                                                	1363        	; Home - View
										,	"Zoom_Plus":                                                 	1360        	; Home - View
										,	"Zoom_Minus":                                              	1362        	; Home - View
										,	"Rotate_Left":                                                 	1340        	; Home - View
										,	"Rotate_Right":                                               	1337        	; Home - View
										,	"Highlight":                                                    	46130      	; Home - Comment
										,	"Typewriter":                                                  	46096      	; Home - Comment, Comment - TypeWriter
										,	"Open_From_File":                                        	46140      	; Home - Create
										,	"Open_Blank":                                               	46141      	; Home - Create
										,	"Open_From_Scanner":                                  	46165      	; Home - Create
										,	"Open_From_Clipboard":                               	46142      	; Home - Create - new pdf from clipboard
										,	"PDF_Sign":                                                   	46157      	;Home - Protect
										,	"Create_Link":                                                	46080      	; Home - Links
										,	"Create_Bookmark":                                       	46070      	; Home - Links
										,	"File_Attachment":                                          	46094      	; Home - Insert
										,	"Image_Annotation":                                      	46081      	; Home - Insert
										,	"Audio_and_Video":                                       	46082      	; Home - Insert
										,	"Comments_Import":                                      	46083      	; Comments
										,	"Highlight":                                                    	46130      	; Comments - Text Markup
										,	"Squiggly_Underline":                                     	46131      	; Comments - Text Markup
										,	"Underline":                                                   	46132      	; Comments - Text Markup
										,	"Strikeout":                                                     	46133      	; Comments - Text Markup
										,	"Replace_Text":                                              	46134      	; Comments - Text Markup
										,	"Insert_Text":                                                  	46135      	; Comments - Text Markup
										,	"Note":                                                          	46137      	; Comments - Pin
										,	"File":                                                            	46095      	; Comments - Pin
										,	"Callout":                                                       	46097      	; Comments - Typewriter
										,	"Textbox":                                                      	46098      	; Comments - Typewriter
										,	"Rectangle":                                                   	46101      	; Comments - Drawing
										,	"Oval":                                                          	46102      	; Comments - Drawing
										,	"Polygon":                                                      	46103      	; Comments - Drawing
										,	"Cloud":                                                        	46104      	; Comments - Drawing
										,	"Arrow":                                                         	46105      	; Comments - Drawing
										,	"Line":                                                           	46106      	; Comments - Drawing
										,	"Polyline":                                                      	46107      	; Comments - Drawing
										,	"Pencil":                                                         	46108      	; Comments - Drawing
										,	"Eraser":                                                        	46109      	; Comments - Drawing
										,	"Area_Highligt":                                             	46136      	; Comments - Drawing
										,	"Distance":                                                     	46110      	; Comments - Measure
										,	"Perimeter":                                                   	46111      	; Comments - Measure
										,	"Area":                                                          	46112      	; Comments - Measure
										,	"Stamp":                                                        	46149      	; Comments - Stamps , opens only the dialog
										,	"Create_custom_stamp":                                 	46151      	; Comments - Stamps
										,	"Create_custom_dynamic_stop":                     	46152      	; Comments - Stamps
										,	"Summarize_Comments":                                	46188      	; Comments - Manage Comments
										,	"Import":                                                        	46083      	; Comments - Manage Comments
										,	"Export_All_Comments":                                  	46086      	; Comments - Manage Comments
										,	"Export_Highlighted_Texts":                            	46087      	; Comments - Manage Comments
										,	"FDF_via_Email":                                            	46084      	; Comments - Manage Comments
										,	"Comments":                                                 	46088      	; Comments - Manage Comments
										,	"Comments_Show_All":                                   	46089      	; Comments - Manage Comments
										,	"Comments_Hide_All":                                   	46090      	; Comments - Manage Comments
										,	"Popup_Notes":                                               	46091      	; Comments - Manage Comments
										,	"Popup_Notes_Open_All":                                	46092      	; Comments - Manage Comments
										,	"Popup_Notes_Close_All":                               	46093 }     	; Comments - Manage Comments
		FoxitCommands1 := { 	"firstPage":                                                      	1286        	; View - Go To
										,	"lastPage":                                                      	1288        	; View - Go To
										,	"nextPage":                                                     	1289        	; View - Go To
										,	"previousPage":                                               	1290        	; View - Go To
										,	"previousView":                                               	1335        	; View - Go To
										,	"nextView":                                                     	1346        	; View - Go To
										,	"ReadMode":                                                 	1351        	; View - Document Views
										,	"ReverseView":                                               	1353        	; View - Document Views
										,	"TextViewer":                                                  	46180      	; View - Document Views
										,	"Reflow":                                                        	32818      	; View - Document Views
										,	"turnPage_left":                                              	1340        	; View - Page Display
										,	"turnPage_right":                                            	1337        	; View - Page Display
										,	"SinglePage":                                                 	1357        	; View - Page Display
										,	"Continuous":                                                	1338        	; View - Page Display
										,	"Facing":                                                       	1356        	; View - Page Display - two pages side by side
										,	"Continuous_Facing":                                     	1339        	; View - Page Display - two pages side by side with scrolling enabled
										,	"Separate_CoverPage":                                  	1341        	; View - Page Display
										,	"Horizontally_Split":                                        	1364        	; View - Page Display
										,	"Vertically_Split":                                            	1365        	; View - Page Display
										,	"Spreadsheet_Split":                                       	1368        	; View - Page Display
										,	"Guides":                                                       	1354        	; View - Page Display
										,	"Rulers":                                                        	1355        	; View - Page Display
										,	"LineWeights":                                               	1350        	; View - Page Display
										,	"AutoScroll":                                                  	1334        	; View - Assistant
										,	"Marquee":                                                    	1361        	; View - Assistant
										,	"Loupe":                                                        	46138      	; View - Assistant
										,	"Magnifier":                                                   	46139      	; View - Assistant
										,	"Read_Activate":                                             	46198      	; View - Read
										,	"Read_CurrentPage":                                      	46199      	; View - Read
										,	"Read_from_CurrentPage":                             	46200      	; View - Read
										,	"Read_Stop":                                                  	46201      	; View - Read
										,	"Read_Pause":                                               	46206      	; View - Read
										,	"Navigation_Panels":                                      	46010      	; View - View Setting
										,	"Navigation_Bookmark":                                	45401      	; View - View Setting
										,	"Navigation_Pages":                                      	45402      	; View - View Setting
										,	"Navigation_Layers":                                      	45403      	; View - View Setting
										,	"Navigation_Comments":                               	45404      	; View - View Setting
										,	"Navigation_Appends":                                  	45405      	; View - View Setting
										,	"Navigation_Security":                                    	45406      	; View - View Setting
										,	"Navigation_Signatures":                                	45408      	; View - View Setting
										,	"Navigation_WinOff":                                    	1318        	; View - View Setting
										,	"Navigation_ResetAllWins":                             	1316        	; View - View Setting
										,	"Status_Bar":                                                  	46008        	; View - View Setting
										,	"Status_Show":                                               	1358        	; View - View Setting
										,	"Status_Auto_Hide":                                       	1333        	; View - View Setting
										,	"Status_Hide":                                                	1349        	;View - View Setting
										,	"WordCount":                                                	46179      	;View - Review
										,	"Form_to_sheet":                                            	46072      	;Form - Form Data
										,	"Combine_Forms_to_a_sheet":                        	46074      	;Form - Form Data
										,	"DocuSign":                                                   	46189      	;Protect
										,	"Login_to_DocuSign":                                     	46190      	;Protect
										,	"Sign_with_DocuSign":                                   	46191      	;Protect
										,	"Send_via_DocuSign":                                    	46192      	;Protect
										,	"Sign_and_Certify":                                        	46181      	;Protect
										,	"-----_-------------":                                         	46182      	;Protect
										,	"Place_Signature":                                          	46183      	;Protect
										,	"Validate":                                                     	46185      	;Protect
										,	"Time_Stamp_Document":                              	46184      	;Protect
										,	"Digital_IDs":                                                 	46186      	;Protect
										,	"Trusted_Certificates":                                     	46187      	;Protect
										,	"Email":                                                         	1296        	;Share - Send To - same like Email current tab
										,	"Email_All_Open_Tabs":                                 	46012      	;Share - Send To
										,	"Tracker":                                                      	46207      	;Share - Tracker
										,	"User_Manual":                                              	1277        	;Help - Help
										,	"Help_Center":                                               	558          	;Help - Help
										,	"Command_Line_Help":                                 	32768      	;Help - Help
										,	"Post_Your_Idea":                                           	1279        	;Help - Help
										,	"Check_for_Updates":                                    	46209      	;Help - Product
										,	"Install_Update":                                            	46210      	;Help - Product
										,	"Set_to_Default_Reader":                                	32770      	;Help - Product
										,	"Foxit_Plug-Ins":                                             	1312        	;Help - Product
										,	"About_Foxit_Reader":                                    	57664      	;Help - Product
										,	"Register":                                                      	1280        	;Help - Register
										,	"Open_from_Foxit_Drive":                              	1024        	;Extras - maybe this is not correct!
										,	"Add_to_Foxit_Drive":                                     	1025        	;Extras - maybe this is not correct!
										,	"Delete_from_Foxit_Drive":                             	1026        	;Extras - maybe this is not correct!
										,	"Options":                                                     	243          	;the following one's are to set directly any options
										,	"Use_single-key_accelerators_to_access_tools":  	128          	;Options/General
										,	"Use_fixed_resolution_for_snapshots":             	126          	;Options/General
										,	"Create_links_from_URLs":                              	133          	;Options/General
										,	"Minimize_to_system_tray":                             	138          	;Options/General
										,	"Screen_word-capturing":                               	127          	;Options/General
										,	"Make_Hand_Tool_select_text":                       	129          	;Options/General
										,	"Double-click_to_close_a_tab":                       	91            	;Options/General
										,	"Auto-hide_status_bar":                                  	162          	;Options/General
										,	"Show_scroll_lock_button":                             	89            	;Options/General
										,	"Automatically_expand_notification_message":	1725        	;Options/General - only 1 can be set from these 3
										,	"Dont_automatically_expand_notification":      	1726        	;Options/General - only 1 can be set from these 3
										,	"Dont_show_notification_messages_again":     	1727        	;Options/General - only 1 can be set from these 3
										,	"Collect_data_to_improve_user_experience":   	111          	;Options/General
										,	"Disable_all_features_which_require_internet":	562          	;Options/General
										,	"Show_Start_Page":                                        	160          	;Options/General
										,	"Change_Skin":                                             	46004
										,	"Filter_Options":                                            	46167      	;the following are searchfilter options
										,	"Whole_words_only":                                     	46168      	;searchfilter option
										,	"Case-Sensitive":                                            	46169      	;searchfilter option
										,	"Include_Bookmarks":                                    	46170      	;searchfilter option
										,	"Include_Comments":                                     	46171      	;searchfilter option
										,	"Include_Form_Data":                                    	46172      	;searchfilter option
										,	"Highlight_All_Text":                                       	46173      	;searchfilter option
										,	"Filter_Properties":                                          	46174      	;searchfilter option
										,	"Print":                                                           	57607
										,	"Properties":                                                   	1302        	;opens the PDF file properties dialog
										,	"Mouse_Mode":                                             	1311
										,	"Touch_Mode":                                              	1174
										,	"predifined_Text":                                           	46099
										,	"set_predefined_Text":                                    	46100
										,	"Create_Signature":                                        	26885      	;Signature
										,	"Draw_Signature":                                          	26902      	;Signature
										,	"Import_Signature":                                        	26886      	;Signature
										,	"Paste_Signature":                                          	26884      	;Signature
										,	"Type_Signature":                                           	27005      	;Signature
										,	"Pdf_Sign_Close":                                          	46164}    	;Pdf-Sign

		For key, val in FoxitCommands1
			FoxitCommands[key] := val

	}

	If FoxitID
		PostMessage, 0x111, % FoxitCommands[command],,, % "ahk_id " FoxitID
	else
		return FoxitCommands[command]
}

FoxitReader_GetPages(FoxitID="") {                                                                                	;-- aktuelle und maximale Seiten des aktuellen Dokumentes ermitteln

	; letzte Änderung 18.10.2020

	; nachsehen ob korrekte FoxitID übergeben wurde
		While (!FoxitID || !WinExist("ahk_id " FoxitID)) {
			If (FoxitID := WinExist("ahk_class classFoxitReader"))
				break
			else if (A_Index > 20)
				return {"disp":1, "max":1}
			Sleep 50
		}

	; Handle der Statusbar ermitteln
		WinGet, hCtrl, ControlList, % "ahk_id " FoxitID
		Loop, Parse, hCtrl, `n
			If InStr(A_LoopField, "BCGPRibbonStatusBar") {
				ControlGet, StatusbarHwnd, Hwnd,, % A_LoopField, % "ahk_id " FoxitID
				break
			}

	; Text der Steuerelemente nach Seitenanzeige durchsuchen
		WinGet, hCtrl, ControlList, % "ahk_id " StatusbarHwnd
		;SciTEOutput("Statusbarhwnd: " StatusbarHwnd "`nhCtrls: " hCtrl)
		Loop, Parse, hCtrl, `n
		{
			ControlGetText, Pages, % A_LoopField, % "ahk_id " StatusbarHwnd
			If RegExMatch(Pages, "(?<Disp>\d+)\s*\/\s*(?<Max>\d+)", Page) {
				PageDisp	:= StrLen(PageDisp) = 0	? 1 : PageDisp
				PageMax	:= StrLen(PageMax) = 0	? 1 : PageMax
				return {"disp":PageDisp, "max":PageMax}
			}
		}

return {"disp":1, "max":1} ; wenigsten eine 1 zurückgeben wenn nichts ermittelt werden konnte
}

FoxitReader_ToPrint(FoxitID="", Printer="") {                                                                     	;-- Druck Dialoghandler - Ausdruck auf übergebenen Drucker

		static foxitprint    	:= "i)[(Print)|(Drucken)]+ ahk_class i)#32770 ahk_exe i)FoxitReader.exe"

		If !FoxitID
			FoxitID := WinExist("ahk_class classFoxitReader")

		OldMatchMode := A_TitleMatchMode
		SetTitleMatchMode, RegEx                                                              	; RegEx Fenstervergleichsmodus einstellen

		FoxitInvoke("Print", FoxitID)                                                               	; 'Drucken' - Dialog wird aufgerufen
		WinWait, % foxitPrint,, 6                                                                 	; wartet 6 Sekunden auf das Dialogfenster
		hfoxitPrint	:= GetHex(WinExist(foxitPrint))                                        	; 'Drucken' - Dialog handle
		ItemNr  	:= VerifiedChoose("ComboBox1", hfoxitPrint, Printer)          	; Drucker auswählen
		If (ItemNr <> 0) {
			VerifiedClick("Button44", hfoxitPrint,,, true)                                    	; OK Button drücken
			WinWaitClose, % "ahk_id " hfoxitPrint,, 3                                    	; wartet max. 3 Sek. bis der Dialog geschlossen wurde
		}

		SetTitleMatchMode, % OldMatchMode                                            	; TitleMatchMode zurückstellen

return {"DialogID":hFoxitPrint, "ItemNr":ItemNr}                                   	; für Erfolgskontrolle und eventuelle weitere Abarbeitungen
}

FoxitReader_SignaturSetzen(FoxitID="") {	                                                            			;-- ruft Signatur setzen auf und zeichnet eine Signatur in die linke obere Ecke des Dokumentes

		; letzte Änderung: 19.09.2020

			CoordModeMouse_before :=  A_CoordModeMouse
			CoordMode, Mouse, Screen

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Variablen
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			; ! NICHT ÄNDERN! dieser String wird für 'feiyus' FindText() Funktion benötigt, er entspricht der linken oberen Ecke des PDF-Frame (AfxWnd100su4)
			;~ static TopLeft :=	"|<>*210$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000003zzzzs0000007"
			        	        	;~ . 	"zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000T"
			                		;~ . 	"zzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001"

			static TopLeft :="|<>ABABAB-000000$47.000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzk0007zzzU000Dzzy0000Tzzw0000zzzs0001zzzk0003zzzU0007zzz0000Dzzy0000Tzzw0000zzzs0001"
			static basetolerance := 0.5

			Addendum.PDFRecentlySigned := true
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Abbruch wenn kein FoxitReaderfenster
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			If !FoxitID
				If !(FoxitID := WinExist("ahk_class classFoxitReader"))
					return 0

			; PDF Backup!
			for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process") {
				If InStr(process.name, "FoxitReader") {
					RegExMatch(process.CommandLine, "\s\" q "(.*)?\" q, cmdline)
					break
				}
			}

			SplitPath, cmdline1, PDFName
			FileCopy, % cmdline1, % Addendum.BefundOrdner "\Backup\" PdfName, 1
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Ermitteln des Childhandles (Docwindow) im Foxitreader (Bildschirmposition des Dokuments)
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			ActivateAndWait("ahk_id " FoxitID, 1)
			res	        	:= Controls("", "Reset", "")
			hDocTab      	:= Controls("", "GetActiveMDIChild, return hwnd", FoxitID)
			hDocWnd 	:= Controls("FoxitDocWnd", "hwnd", hDocTab)
			DocWnd      	:= GetWindowSpot(hDocWnd)

			SciTEOutput("DocWnd handle: " hDocWnd " x" DocWnd.X  " y" DocWnd.Y " w" DocWnd.W " h" DocWnd.H )
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; FoxitReader vorbereiten für das Platzieren der Signatur
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			PraxTT("sende Befehl: 'Signatur platzieren'", "12 2")

			ActivateAndWait("ahk_id " FoxitID, 1)

			FoxitInvoke("SinglePage", FoxitID)
			FoxitInvoke("Fit_Page"  	, FoxitID)
			FoxitInvoke("FirstPage" 	, FoxitID)

			Sleep, 500
			MouseMove	, % DocWnd.X + 50, % DocWnd.Y + 50, 0
			MouseClick	, Left

		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Signatur setzen Menupunkt aufrufen
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			ActivateAndWait("ahk_id " FoxitID, 1)
			FoxitInvoke("Place_Signature", FoxitID)
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sucht nach der linken oberen Ecke der Pdf Seite der 1.Seite, um dort im Anschluss die Signatur zu erstellen
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			PraxTT(" suche nach dem Signierbereich des Dokumentes.", "12 2")

			tryCount := 1
			FindSignatureRange:
			; Funktion FindText macht eine Bildsuche (entspricht der linken oberen Ecke des PDF Preview Bereiches)
			if (Ok := FindText(DocWnd.X, DocWnd.Y, DocWnd.W, DocWnd.H, basetolerance, 0, TopLeft)) {

				PraxTT("Fläche zum Signieren gefunden.", "4 0")
				X	:= ok.1.1 ;+ 5
				Y	:= ok.1.2 ;+ 5
				W	:= ok.1.3
				H	:= ok.1.4
				;X += W//2
				;Y += H//2
				Comment := ok.1.5
				;MouseMove, % X, % Y, 0
				SciTEOutput("SBereich: x" X " y" Y " w" (X + Addendum.SignatureWidth) " h" (Y + Addendum.SignatureHeight))
				MouseMove, % X, % Y
				;MouseClickDrag, Left, % X, % Y, % (X + Addendum.SignatureWidth), % (Y + Addendum.SignatureHeight), 0
				SciTEOutput(">" tryCount ":->" X ", " Y)

			} else {

				sleep, 100
				tryCount ++
				basetolerance += 0.1
				If (tryCount < 30)
					goto FindSignatureRange
				else {

					PraxTT("", "off")
					return 0

				}
			}
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sichern der akutellen FoxitID
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
			Addendum.PDFSignaturedID   	:= FoxitID
			Addendum.PDFRecentlySigned 	:= true
		;}

			CoordMode, Mouse, % CoordModeMouse_before

return 1
}

FoxitReader_SignDoc(hDokSig) {		                                     	                      		        	;-- Bearbeiten des 'Dokument signieren' (Sign Document) Dialoges

		; letzte Änderung: 18.10.2020

		static appendix := "ahk_class #32770 ahk_exe FoxitReader.exe"

		Addendum.PDFRecentlySigned := true
		FoxitID := GetParent(hDokSig)
		PraxTT("Das Fenster 'Dokument signieren' wird gerade bearbeitet....!'", "30 2")   ; {"timeout":30, "zoom":2, "position":"Bottom", "parent":FoxitID}) 	;"30 2")

	;{ Felder im Signierfenster auf die in der INI festgelegten Werte einstellen

			ActivateAndWait("ahk_id " hDokSig, 5)

		; Signieren als -------------------------------------------------------------------------------------------------------
			ControlFocus 	, ComboBox1                                                          	, % "ahk_id " hDokSig
			Control, ChooseString, % Addendum.SignierenAls , ComboBox1     	, % "ahk_id " hDokSig
			sleep, 50

		; Darstellungstyp ----------------------------------------------------------------------------------------------------
			ControlFocus 	, ComboBox4                                                        	, % "ahk_id " hDokSig
			ControlGet, entryNr, FindString, % Addendum.Darstellungstyp
									, ComboBox4                                                        	, % "ahk_id " hDokSig  	; prüft das Feld Signaturvorschau auf die in der ini hinterlegte Signatur
			If !entryNr
				MsgBox, 4144, Addendum für AlbisOnWindows, % "Der gewünschte Darstellungstyp "
																						. 	Addendum.Darstellungstyp "`n"
																						. 	"ist nicht vorhanden"
			else
				Control, ChooseString, % Addendum.Darstellungstyp, ComboBox4	, % "ahk_id " hDokSig
			Sleep, 50

		; Ort: -----------------------------------------------------------------------------------------------------------------
			;VerifiedSetText("Edit2", JEE_StrUtf8BytesToText(Addendum.Ort), "ahk_id " hDokSig)
			ControlFocus 	, Edit2			                                                    		, % "ahk_id " hDokSig
			ControlSetText	, Edit2,  % JEE_StrUtf8BytesToText(Addendum.Ort)    	, % "ahk_id " hDokSig
			ControlSend		, Edit2,  {Tab}	                                                    	, % "ahk_id " hDokSig
			sleep, 100

		; Grund: -------------------------------------------------------------------------------------------------------------
			;VerifiedSetText("Edit3", JEE_StrUtf8BytesToText(Addendum.Grund), "ahk_id " hDokSig)
			ControlFocus 	, Edit3		                                                     			, % "ahk_id " hDokSig
			ControlSetText	, Edit3, % JEE_StrUtf8BytesToText(Addendum.Grund)	, % "ahk_id " hDokSig
			ControlSend		, Edit3, {Tab}	                                                     	, % "ahk_id " hDokSig
			sleep, 50

		; nach der Signierung sperren: -----------------------------------------------------------------------------------
			If Addendum.DokumentSperren
				VerifiedCheck("Button4","","", hDokSig)
			sleep, 50

	;}

	;{ Signatur-Dialog schließen

		; Signaturfenster schliessen
			while WinExist("ahk_id " hDokSig) {

				If VerifiedClick("Button5", hDokSig)
					break
				else If (A_Index > 10)
					MsgBox, 4144, % "Addendum für AlbisOnWindows",  %	"Das Eintragen des Kennwortes hat nicht funktioniert.`n"
											                                                		.	"Bitte tragen Sie es bitte manuell ein!`n"
																									.	"Drücken Sie danach bitte erst auf Ok."
				sleep, 50

			}

	;}

	;{ statistische Daten erfassen und speichern

		; Ermitteln der Seitenzahl
			PdfPages := FoxitReader_GetPages(FoxitID)                                     ; Seitenzahl des Dokumentes ermitteln (Statistik)
			If !RegExMatch(PdfPages.Max, "\d+")
				PdfPages.Max := 1

		; Signaturzähler erhöhen, Zähler sichern und kurz anzeigen
			Addendum.SignatureCount ++
			Addendum.SignaturePages += PdfPages.Max
			IniWrite, % Addendum.SignatureCount	, % AddendumDir "\Addendum.ini", % "ScanPool", % "SignatureCount"
			IniWrite, % Addendum.SignaturePages	, % AddendumDir "\Addendum.ini", % "ScanPool", % "SignaturePages"
			ToolTip, %	"Signature Nr: "     	Addendum.SignatureCount
						. 	"`nSeitenzahl: "         	                PdfPages.Max
						. 	"`nges. Seiten: "     	Addendum.SignaturePages , 1000, 1, 10

	;}

	; Dateidialog Routinen starten
		SetTimer, PDFNotRecentlySigned		, -10000
		PraxTT("'", "off")

return

PDFHelpCloseSaveDialogs:            	;{ - Notfalllösung für die immer noch unsichere Dialogerkennung

		If (hwnd := WinExist("Speichern unter " appendix, "Speichern")) {

			SetTimer, PDFHelpCloseSaveDialogs,  Off
			VerifiedClick("Speichern", hwnd)                                                                 	; Speichern Button drücken
			WinWait, % "Speichern unter bestätigen " appendix,, 2
			If !ErrorLevel
				return VerifiedClick("Ja", "Speichern unter bestätigen " appendix,,, true)    	; mit 'Ja' bestätigen
			Addendum.PDFRecentlySigned := false

		}

return ;}

PDFNotRecentlySigned: ;{
		Addendum.PDFRecentlySigned := false
		Addendum.SaveAsInvoked := false
		SetTimer, PDFHelpCloseSaveDialogs,  Off
		Tooltip,,,, 10
return ;}
}

FoxitReader_GetPDFPath() {                                                                                            	;-- den aktuellen Dokumentenpfad im 'Speichern unter' Dialog auslesen

	foxitSaveAs := "Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe"
	If WinExist(foxitSaveAs) {

		WinGetText, allText, % foxitSaveAs
		RegExMatch(allText, "(?<Name>[\w+\s_\-\,]+\.pdf)\n.*Adresse\:\s*(?<Path>[A-Z]\:[\\\w\s_\-]+)\n", File)
		return FilePath "\" FileName

	}

return ""
}

JEE_StrUtf8BytesToText(vUtf8) {                                                                                                	;-- wandelt UTF8Bytes in Text (ini Dateien)
	if A_IsUnicode	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
	return StrGet(&vUtf8X, "UTF-8")
	} 	else
		return StrGet(&vUtf8, "UTF-8")
}

;}

;{19. Automatisierung anderer Programme (MicroDicom)
MicroDicom_VideoExport() {                                                                                          	;-- automatischer Export in ein Videoformat mit Dateinamenerstellung anhand von Dicom Tags

	; MicroDicom - kostenloser DICOM Viewer für Windows
	; Download Seite: https://www.microdicom.com/downloads.html
	; optimiert für die Version 3.4+

		StringCaseSense, Off

	; window description for MicroDicom 64bit V2 'Export to video' dialog
		MDExport  	:= {	"Title"              	: {"class":"Export to video ahk_class #32770", "Type":"window"}
								, 	"Dest"             	: {"class":"Edit1"         	, "Type":"edit"}
								,	"Name"           	: {"class":"Edit2"         	, "Type":"edit"}
								, 	"Source"          	: {"class":"ComboBox1"	, "Type":"combobox"}
								, 	"Planes"           	: {"class":"Button4"     	, "type":"checkbox"}              	; export each planes
								, 	"SpFiles"          	: {"class":"Button5"     	, "type":"checkbox"}               	; Separate files for &every
								,	"WMV"             	: {"class":"Button7"     	, "type":"button"}                	; WMV Export (Checkbox)
								,	"AVI"                	: {"class":"Button8"     	, "type":"button"}                  	; AVI Export (Checkbox)
								,	"FRateDefault"  	: {"class":"Button9"     	, "type":"button"}                  	; Frame rate custom
								,	"FRateCustom"  	: {"class":"Button10"     	, "type":"button"}                  	; Frame rate custom
								,	"FPS"              	: {"class":"Edit3"         	, "Type":"edit"}                     	; frames per second
								,	"WMVQuality"   	: {"class":"Edit4"         	, "Type":"edit"}                     	; Exportquality
								,	"ComprYes"      	: {"class":"Button11"     	, "type":"button"}                	; AVI Compression Yes
								,	"ComprNo"      	: {"class":"Button12"    	, "type":"button"}                  	; AVI Compression No
								,	"SizeOrig"      	: {"class":"Button14"    	, "type":"button"}                  	; VideoSize Original
								,	"SizeAsScreen"  	: {"class":"Button15"    	, "type":"button"}                  	; Video Size As on screen
								,	"SizeOther"     	: {"class":"Button16"    	, "type":"button"}                  	; Video Size As on screen
								,	"SizeWidth"      	: {"class":"Edit5"         	, "Type":"edit"}                     	; Video Size Other Width
								,	"SizeHeight"      	: {"class":"Edit6"         	, "Type":"edit"}                     	; Video Size Other Height
								,	"Annotation"     	: {"class":"Button18"    	, "type":"checkbox"}              	; show annotations
								,	"AllOverlay"     	: {"class":"Button19"    	, "type":"button"}                  	; all overlay
								,	"Anonymous"    	: {"class":"Button20"    	, "type":"button"}                  	; anonymous overlay
								,	"NoOverlay"    	: {"class":"Button21"    	, "type":"button"}                  	; Without overlay
								,	"Export"         	: {"class":"Button22"    	, "type":"button"}                	; Export button
								,	"Cancel"         	: {"class":"Button23"    	, "type":"button"}}                	; Cancel button

	; smart RPA working object for function RPA() - sets all needed options in 'Export to video' dialog
		ExportRPA := ["Dest|" Addendum.VideoOrdner
							, "Name|%PatName%"		; wird ersetzt
							, "Source|i)All\spatients" 	; RegEx string
							, "Planes|check"
							, "SpFiles|uncheck"
							, "WMV|click"
							, "FRateCustom|click"
							, "FPS|10"
							, "WMVQuality|100"
							, "SizeOrig|click"
							, "Annotation|check"
							, "NoOverlay|click"
							, "Export|click"]

	; Handle des MicroDicom Hauptfensters
		If !(hMDicom := WinExist("MicroDicom viewer")) {
			PraxTT("Der automatische Videoexport ist fehlgeschlagen.", "4 0")
			return
		}

	; den Exportdialog erstmal wieder schließen
		If (hExport := WinExist(MDExport.title.class))
			VerifiedClick(MDExport.Cancel.class, hExport)

	; die Dicom Tags anzeigen
		If !(hDicomTags := MicroDicom_Invoke("DicomTags", hMDicom, "DICOM Tags ahk_class #32770", 6)) {
			SciTEOutput("Der automatische Videoexport ist fehlgeschlagen.")
			PraxTT("Der automatische Videoexport ist fehlgeschlagen.`nDer Dialog Dicom Tags konnte nicht aufgerufen werden", "4 0")
			return
		}

	; Patientenname, Geburtstag,Untersuchungsart und Untersuchungstag ermitteln (Auslesen der SysListview321 im Dicom Tags Dialog)
		ExportRPA[2] := MicroDicom_BuildFileName(hDicomTags)

	; Dicom Tags Dialog schliessen
		VerifiedClick("Button1", hDicomTags)
		WinWaitClose, % "ahk_id " hDicomTasgs,, 3

	; den Exportvorgang mit den ausgelesenen Daten starten
		If !MicroDicom_Invoke("ExportToVideo", hMDicom, MDExport.title.class, 6) {
			PraxTT("Der automatische Videoexport ist fehlgeschlagen.", "4 0")
			return
		}

	; den Inhalt der Dicom CD als Video exportieren
		RPA(MDExport, ExportRPA)

	; wartet bis zu 3min auf den Abschluss des Exportvorganges - ein Explorerfenster wird nach Abschluss angezeigt
		;WinWait, % "ahk_class CabinetWClass", % Addendum.VideoOrdner, 300  ; werden andere Threads noch ausgeführt wenn WinWait aktiv ist?

	; Schliessen und Auswerfen der CD nach Export
		;MsgBox, 4, Addendum für Albis on Windows, % ""

return
}

MicroDicom_Invoke(command, MicroDicomID="", WinToWait="", TimeToWait=3) {          	;-- Menuaufrufe für MicroDicom 64bit

		static MicroDicom := { 	"DicomTags":           	40018  	; View
										,	"ExportToImage":      	32790  	; File/Export/To a picture file
										,	"ExportToVideo":      	32806}	; File/Export/To a video file

	; wm_command code wird zurück gegeben
		If !MicroDicomID
			return MicroDicom[command]

	; Fenster ist schon geöffnet, Rückgabe des Handle
		If (StrLen(WinToWait) > 0) && (hwnd := WinExist(WinToWait))
			return hwnd

	; Senden des wm_command Befehls
		PostMessage, 0x111, % MicroDicom[command],,, % "ahk_id " MicroDicomID

	; wartet auf den sich öffnenden Dialog und gibt das Fensterhandle zurück
		If (StrLen(WinToWait) > 0) {
			WinWait, % WinToWait,, % TimeToWait
			return WinExist(WinToWait) 	; handle ist 0 bei nicht vorhandenem Fenster
		}


return ; gibt nichts zurück!
}

MicroDicom_BuildFileName(hDicomTags) {                                                                    	;-- erstellt einen sinnvollen Dateinamen für den Export

		ControlGet, tags, List,, SysListView321, % "ahk_id " hDicomTags
		RegExMatch(tags, "StudyDescription\s*\w+\s*\w+\s*\w+\s*([\w\d\s]+)", description)
		RegExMatch(tags, "i)PatientName\s*\w+\s*\w+\s*\w+\s*(?<Name>[\p{L}\s^]+)", Patient)
		RegExMatch(tags, "i)PatientBirthDate\s*\w+\s*\w+\s*\w+\s*(?<BD>\d+)", Pat)
		RegExMatch(tags, "i)StudyDate\s*\w+\s*\w+\s*\w+\s*(?<D>\d+)", St)
		PatientName := RegExReplace(PatientName, "(\p{L})\^(\p{L})", "$1, $2" )
		PatientName := RegExReplace(PatientName, "(\p{L})\s+(\p{L})", "$1, $2" )
		PatientName := StrReplace(PatientName, "^")
		PatientName := Trim(StrReplace(PatientName, "`n"))
		PatientBirthDate := SubStr(PatBD, 7,2) "." SubStr(PatBD, 5,2) "." SubStr(PatBD, 1,4)
		StudyDate := SubStr(StD, 7,2) "." SubStr(StD, 5,2) "." SubStr(StD, 1,4)

return "Name|" PatientName " (" PatientBirthdate ") - " StrReplace(description1, "`n") " vom " StudyDate " - "
}

RPA(WinObj, steps) {                                                                                                      	;-- universal prototype: 'macro like' robot process automation for system controls in windows

	; universal function to check or uncheck checkboxes, radiobuttons or to select entries in combo/listboxes and to set text entries in edit controls
	; function is made to be used on standard windows controls
	; the main goal is to give users the possibility to have their own settings for windows
	;
	; dependancies: Addendum_Control.ahk (VerifiedSetText/Click/Check/Choose)

	If !IsObject(WinObj)
		throw Exception("parameter 'WinObject': is not an object", "RPA")

	If !IsObject(steps)
		throw Exception("parameter 'steps': is not an object", "RPA")

	If !(hRPAWin := WinExist(WinObj.title.class))
		throw Exception("Could not find window:`n" WinObj.title.class, "RPA")

	For idx, step in steps
	{
			on := StrSplit(step, "|").1
			do := StrSplit(step, "|").2

			switch WinObj[on].type
			{
					case "edit":
						VerifiedSetText(WinObj[on].class, do, hRPAWin)
					case "button":
						VerifiedClick(WinObj[on].class, hRPAWin)
					case "checkbox":
						VerifiedCheck(WinObj[on].class, hRPAWin,,, InStr(do, "uncheck") ? false:true)
					case "combobox":
						VerifiedChoose(WinObj[on].class, hRPAWin, do)
			}

	}

}
;}
;----------------------------------------------------------------------------------------------------------------------------------------------

;{20. Gui's, Icons

Auswahlbox(content, guiname="Auswahlbox") {

		global

		static YDelta := 15
		static CaretX, CaretY, awbG, monNr, Mon, LbContent, AWBLb, hAWBLb, conStored
		static FontName, FontSize, FontStyle, cy, ch, GuiY

		SetWinDelay, -1
		SetControlDelay, -1
		SetKeyDelay, -1, -1

		ControlGetFocus, ACControl, A
		ACWin          	:= WinExist("A")
		CaretX          	:= A_CaretX
		CaretY          	:= A_CaretY
		monNr	        	:= GetMonitorAt(CaretX, CaretY)
		Mon  	        	:= ScreenDims(monNr)
		LbContent     	:= ""
		conStored       	:= content

		ControlGet, ACHwnd, HWND,, % ACControl, % "ahk_id " ACWin
		ControlGetPos,,cy,,ch,, % "ahk_id " ACHwnd
		ControlGetFont(ACHwnd, FontName, FontSize, FontStyle)
		;SciTEOutput(">Style: " FontStyle "`ny: " cy " ,h: " ch)

		If IsObject(content)
			If RegExMatch(content[1], "\{.*\}")                                     ; für Diagnosen
				For idx, val in content {
					RegExMatch(val, "(?<Text>.*)?\{(?<Code>.*)?G\}", ICD)
					LbContent .= SubStr(ICDCode . "          ", 1, FontSize-StrLen(ICDCode)) . "`t" . ICDText . "|"
				}

		Gui, AWB: new		, -Caption +ToolWindow +AlwaysOnTop +HWNDhAWB
		Gui, AWB: Margin	, 0, 0
		Gui, AWB: Font		, % "s" FontSize, % FontName                                                                ; " " FontStyle
		Gui, AWB: Add		, Listbox, % "r20 w1000 vAWBLB HWNDhAWBLb T12 " (IsObject(content) ? "AltSubmit":"")  " Choose1 Multi" , % RTrim(LbContent, "|")
		Gui, AWB: Show	, % "x" CaretX " y" CaretY + YDelta " Hide NA", % "Auswahlbox"

		awbG  	:= GetWindowSpot(hAWb)

		If (awbG.Y + awbG.H > Mon.H)
			Gui, AWB: Show	, % "x" CaretX " y" CaretY - awbG.H - YDelta " NA", % "Auswahlbox"

		HotKey, IfWinExist, % "Auswahlbox"
		HotKey, Enter	, AWBTLB
		Hotkey, Down  	, AWBDown
		Hotkey, Up    	, AWBUp
		HotKey, Esc   	, AWBGuiEscape
		Hotkey, IfWinExist

Return

AWBDown:
AWBUp:            	;{
	SetKeyDelay, -1, -1
	ControlSend,, % "{" A_ThisHotkey "}", % "ahk_id " hAWBLb
return ;}
AWBTLB:             	;{

		Gui, AWB: Submit, Hide

		WinActivate    	, % "ahk_id " ACWin
		WinWaitActive	, % "ahk_id " ACWin,, 2

		ControlFocus, % ACControl, % "ahk_id " ACWin

		If IsObject(conStored)
			For idx, item in StrSplit(AWBLb, "|") {
				Send, % "{Text}" conStored[item]
				Sleep 100
			}

;}
AWBGuiClose:
AWBGuiEscape: 	;{

		Gui, AWB: Destroy

return ;}
}

; nicht mehr Infofenster
HotstringComboBox(con) {

	FileRead, thisScript, % A_ScriptFullPath
	Loop, Parse, thisScript, `n, `r
	{
			If (collecting = 1) && RegExMatch(A_LoopField, "^#If$")
					collecting:= 0
			If collecting = 1
			{
					rpos:= RegExMatch(A_LoopField, "(?<=.:)\w+", HStr)
					If RegExMatch(A_LoopField, "(?<=\w.:)[\w\d\{\}\s\,\.\;]+(?=\s*;)", TextToSend)
					{

					}

			}
			If RegExMatch(A_LoopField, "^#If.*""contraction"".*" con)
					collecting:= 1


	}

return
}

AutoListview(ColStr, newline, ColSize:="", Separator:=",") {                                             	;-- eine Listview Gui zum schnellen Anzeigen von Daten

	static Init, colmax, DasLV, hALV, ALV, hDasLV

	If !Init {

			Init	:= true
			Sysget, Mon1, Monitor, 1

			If (Mon1Right>1920)
			{
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
			else
			{
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
	If IsObject(newline)
	{
			For idx, value in newline
				LV_Add("", idx, value)
	}
	else
	{
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

GuiLaborDruck() {


}

Calendar(feld) {                                                                                                            	;-- Ersatz für den Kalender der sich mit Shift+F3 in Albis öffnet

	;--------------------------------------------------------------------------------------------------------------------
	; Variablen
	;--------------------------------------------------------------------------------------------------------------------;{
		static Formular	:= ["Muster 1a", "Muster 12a", "Muster 20" , "Muster 21", "Privat-Au"]	; Formulare mit Anfangs- und Enddatum
		static Cok, Can, Hinweis, MyCalendar, hCal, CalBorder, FormularT
		static feldB:= Object()
		global newCal, hmyCal
		FormularT	:= ""
		feldB:= feld

	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Berechnungen für Datumsfokus
	;--------------------------------------------------------------------------------------------------------------------;{
		datum:= feldB.datum

		day:= SubStr(datum, 1, 2), month:= SubStr(datum, 3, 2), year:= SubStr(datum, 5, 4)
		if (month-1 < 1)
				month:= SubStr( "00" . (12 + month - 1), -1), year -= 1
		else
				month := SubStr("00" . (month - 1), -1)
		FormatTime, begindate, % year . month . day, yyyyMMdd
		FormatTime, datum, % datum, yyyyMMdd
	;}

	;--------------------------------------------------------------------------------------------------------------------
	;	Kalender Gui
	;--------------------------------------------------------------------------------------------------------------------;{
		Gui, newCal: Destroy
		Gui, newCal: +Owner +AlwaysOnTop HWNDhCal

	;-:verschiedene Darstellungen je nach geöffnetem Formular
		Loop, % Formular.MaxIndex()
			If InStr(feldB.WinTitle, Formular[A_Index])							;der Hinweis wird nur angezeigt, wenn Anfangs- und Enddatum gewählt werden können
			{
				Gui, newCal: Add, Text, vHinweis Center, (halte Shift um Anfangs-und Enddatum zu wählen)
				Gui, newCal: Add, MonthCal, % "vMyCalendar Multi R3 W-3 4 HWNDhmyCal"	, % begindate
				FormularT:= Formular[A_Index]
				break
			}

		If FormularT = ""
			Gui, newCal: Add, MonthCal, % "vMyCalendar R3 W-3 4 HWNDhmyCal"        	, % begindate

		GuiControl, newCal:, MyCalendar, % datum

		Gui, newCal: Add, Button, default      	gCalendarOK    	vCok, Datum übernehmen
		Gui, newCal: Add, Button, xp+200 yp 	gCalendarCancel	vCan, Abbrechen

		Gui, newCal: Show, AutoSize Hide , Addendum für Albis on Windows - erweiterter Kalender

	;-:Größen ermitteln
		WinA       	:= GetWindowInfo(hCal)
		WinB       	:= GetWindowInfo(feldB.hWin)
		GuiControlGet, MyCal, newCal: Pos, MyCalendar
		GuiControlGet, Cok, newCal: Pos
		GuiControlGet, Can, newCal: Pos
		ControlGetPos, feldX, feldY, feldW, feldH,, % "ahk_id " feld.hwnd
		feldX += WinB.WindowX
		SendMessage 0x101F, 0, 0,, % "ahk_id " hmyCal ; MCM_GETCALENDARBORDER
		CalBorder:= ErrorLevel

	;-:Position von Gui-Elementen anpassen
		GuiControl, newCal: MoveDraw, Hinweis, % "x" MyCalX " w" MyCalW
		GuiControl, newCal: MoveDraw, Cok, % "x" (MyCalX + 2*CalBorder)
		GuiControl, newCal: MoveDraw, Can, % "x" (MyCalX + MyCalW - CanW - 2*CalBorder)

		If InStr(FormularT, "Muster 12a")
				newCalX:= feldX + feldW//2 - WinA.windowW//2, newCalY:= feldY + feldH + 10
		else if InStr(FormularT, "Muster 1a")
				newCalX:= WinB.WindowX + WinB.windowW + 5	, newCalY:= WinB.WindowY + WinB.windowH//2 - WinA.windowH//2
		else if InStr(FormularT, "Muster 20") || InStr(FormularT, "Muster 1a")
				newCalX:= WinB.WindowX - WinA.windowW - 5 	, newCalY:= WinB.WindowY + WinB.windowH//2 - WinA.windowH//2
		else
				newCalX := feldX + feldW + 10, newCalY:= feldY - 10

	;-:wenn Gui die Bildschirmhöhe übersteigt, muss die y-Position angepasst werden
		newCalY := newCalY + WinA.windowH > A_ScreenHeight ? A_ScreenHeight - WinA.windowH - 30 : newCalY

		Gui, newCal: Show, % "x" newCalX " y" newCalY , Addendum für Albis on Windows - erweiterter Kalender

Return ;}

CalendarOK: 				                                                                                                              	 ;{
	Gui, newCal: Submit
	Gui, newCal: Destroy

	If !InStr(MyCalendar,"-")                                                                                                 	; just in case we remove MULTI option
	{
			FormatTime, CalendarM1, %MyCalendar%, dd.MM.yyyy
			CalendarM2:= CalendarM1
	}
	Else
	{
			 FormatTime, CalendarM1, % StrSplit(MyCalendar,"-").1, dd.MM.yyyy
			 FormatTime, CalendarM2, % StrSplit(MyCalendar,"-").2, dd.MM.yyyy
	}

	If InStr(FormularT, "Muster 1a") && (CalendarM1 != CalendarM2)
	{
			ControlSetText, Edit1, %CalendarM1%, % "ahk_id " feldB.hWin
			ControlSetText, Edit2, %CalendarM2%, % "ahk_id " feldB.hWin
	}
	else If (InStr(FormularT, "Muster 12a") || InStr(FormularT, "Muster 21")) && (CalendarM1 != CalendarM2)
	{
			ControlSetText, Edit6, %CalendarM1%, % "ahk_id " feldB.hWin
			ControlSetText, Edit7, %CalendarM2%, % "ahk_id " feldB.hWin
	}
	else if InStr(FormularT, "Muster 20") && (CalendarM1 != CalendarM2)
	{
			editgroup   	:= []
			editgroup[1]	:= [3,4]
			editgroup[2]	:= [7,8]
			editgroup[3]	:= [11,12]
			editgroup[4]	:= [15,16]
			Loop, % editgroup.MaxIndex()
				If (feldB.classnn = editgroup[A_Index, 1]) || (feldB.classnn = editgroup[A_Index, 2])
				{
						ControlSetText, % "Edit" editgroup[A_Index, 1] 	 , %CalendarM1%, % "ahk_id " feldB.hWin
						ControlSetText, % "Edit" editgroup[A_Index, 2] 	 , %CalendarM2%, % "ahk_id " feldB.hWin
						ControlSetText, % "Edit" editgroup[A_Index+1, 1] , %CalendarM1%, % "ahk_id " feldB.hWin
						break
				}
	}
	else
			ControlSetText,, %CalendarM1%, % "ahk_id " feldB.hwnd

	CalendarM1:=CalendarM2:=0

CalendarCancel:
newCalGuiClose:
newCalGuiEscape:
	Sleep 150
	Gui, newCal: Destroy
	WinActivate, % "ahk_id " feldB.hactive
Return
;}

}

; - - - - - - - - - - - - - - ICONS
Create_Addendum_ico()                         	{                                                                  	;-- erstellt das Trayicon

VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAD/////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBf/oo//oo//oo//oo/tl4VTMyJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/6noxCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBeJVkT+oo7/oo//oo//oo//oo+kZ1VCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd7TTr/oo//oo//oo//oo//oo9CKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdDKRj/oo//oo//oo//oo//oo//oo9GKhpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo/+oY7/oo//oo//oo//oo+PWUdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRTP/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo/+oY7/oo//oo//oo/7oI11SDZCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdnQC/0nIn/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9aNyZCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdeOyn/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9LLh1CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdYNiX/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo+QWkiQWkiQWkiQWkiQWkiRW0n/oo//oo//oo//oo//oo+QWkiQWkiQWkiQWkiQWkiQWkj/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo+1cmC1cmC1cmC3c2H/oo//oo//oo//oo//oo+1cmC1cmC1cmC2c2H/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+0cWBCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdILBrkkX7/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo9UNCNCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBeeY1H/oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+rbFpCKBdCKBdCKBf/oo//oo//oo/+oY7/oo9CKBdCKBdGKhr/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo9NLx5CKBdCKBdlPy3/oo//oo//oo9DKRhCKBdCKBeWXk3/oo//oo//oo/7oI13SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+gZVJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhn/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdRMiH/oo//oo//oo//oo//oo9GKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBeRWkn+oY7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+WXkxCKBdCKBdCKBdCKBdCKBdCKBdDKRj/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdVNCP/oo//oo//oo//oo//oo9EKhlCKBdCKBdCKBdCKBdCKBeJVkT+oY7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+MWEZCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdYNiX/oo//oo//oo//oo//oo9DKRhCKBdCKBdCKBeEUkD9oI7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/9oI2BUD9CKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdcOSf/oo//oo//oo//oo//oo9CKBdCKBdCKBf7n43/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/7oI13SjlCKBd3Sjn/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdhOyr/oo//oo//oo//oo//oo+HVEP7n4z/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo/7oIz/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPy3/oo//oo//oo//oo//oo//oo//oo/7oYx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdoQS//oo//oo//oo//oo//oo/7oYx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRTP/oo//oo//oo/7oIx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdxRzX5n4r7oIx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
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

Create_BefundImport_ico()                     	{                                                                     	;-- Icon für das InfoWindow
VarSetCapacity(B64, 788 << !!A_IsUnicode)
B64 := "iVBORw0KGgoAAAANSUhEUgAAABQAAAAYCAYAAAD6S912AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAACtQAAArUBrcsL2AAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAHLSURBVDiNzZM9ixpRFIaPcU68gkHRhAUhBAbtLLYbG6spbWysFMRmy2kt/QE2tmkE21TWUZRAWjsbBZlK0WSsLgyjszPvFlmXbHb8yhrIA7c559znnvsVIiLQFVGIiIQQVK/XqVqtXiwAQK1Wi4bDIUkpf8V6vR5SqRQajQYuwXEcVCoVqKqKfr+P/W5hmiYmkwnS6TQMw4Dv+ydllmWhUCggn89jvV7DNM3nQgCYz+dQVRW1Wg2u6x6UzWYzZLNZlMtl2LYNAMFCAFgul8jlciiVSnAc54VsMBggkUjAMAx4nvcUPygEgM1mA03ToOs6pJRP8U6ng1gshm63+2Kho0IAkFJC13VomgbLstBsNpFMJjEajQKP4aQQAGzbRrFYRDweRyaTwXQ6Daw7WwgAu90O7XYblmUdrLlIeC574Zu/+l9HuLpQISIaj8e0WCxeJVqtVkREFFIUZRGJRN6+vjei7Xa7u4bnGaEjuRtmXgUlXNe9IaIfQbn//5b/zbP5gw9CiHee570/NImZP4XD4ZjjOJKIfp7q8N73/Tsi+nakke+e5zXOanlPNBr9yMyfmfmemfE4fGb+IoRQL5L9DjPfMvPXx3F7qv4BASqDqpfUXLIAAAAASUVORK5CYII="
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

Create_PDF_ico(NewHandle := False)     	{                                                                   	;-- PDF Icon für Befundfenster
VarSetCapacity(B64, 1192 << !!A_IsUnicode)
B64 := "AAABAAEAEBAAAAEAGABoAwAAFgAAACgAAAAQAAAAIAAAAAEAGAAAAAAAAAAAAAkAAAAJAAAAAAAAAAAAAAAFBUAKCoEEBDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKCoAHB2ENDacCAhsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBDANDa0QEMsKCoYBAQYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAhoKCnoQENAKCoAEBDIBAQcAAAAAAAAAAAAAAAAAAAAAAAADAyYFBUcCAh4AAAAAAAAAAAIHB2APD8cMDKENDa8ICGsGBkkDAyoBAQ0CAhoJCXANDa0KCoUJCXsAAAAAAAAAAAACAiAPD8EBAQ4DAyIGBkoJCXANDa0PD8YPD8UQENEPD8EPD74GBlAAAAAAAAAAAAAAAAALC44EBDUAAAAAAAAAAAAHB2APD74GBkoFBT4EBC0CAhoAAAAAAAAAAAAAAAAAAAAHB1gHB2EAAAAAAAAHB1gNDbUDAyoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBDANDaEAAAMGBk4NDbADAyYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQwPD8AICGcPD7sDAygAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALC44QEM0EBDkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQoNDawICGcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBkcQENEHB1wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQMDKENDasHB1cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAh0PD74JCXkFBUQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMHB1sJCXEBAQ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf/wAAD/8AAAf/AACB+AAAwAAAAOAAAADzgQAA8x8AAPA/AADwfwAA+P8AAPH/AADx/wAA4f8AAOH/AADh/wAA"
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

Create_Image_ico(NewHandle := False) 	{                                                                  	;-- Bilder Icon für Befundfenster
VarSetCapacity(B64, 1124 << !!A_IsUnicode)
B64 := "AAABAAEAEA8AAAEAGAA0AwAAFgAAACgAAAAQAAAAHgAAAAEAGAAAAAAAAAAAAAgAAAAIAAAAAAAAAAAAAABNTU0REREREREREREEBAQAAAAAAAAAAAAAAAAAAAALCwsRERERERERERERERFQUFBPT08AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACmpqapqakAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABERET////6+vofHx8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEBDd3d3///////+VlZUAAAAAAAAAAAAAAAAAAAAAAAAAAACSkpKampoKCgoEBASxsbH////////////39/caGhoAAAAAAAAAAAAAAAAAAABycnL////////f39/W1tb///////////////////+Tk5MAAAAAAAAAAAAAAABHR0f8/Pz////////////////////////////////////5+fkfHx8AAAAAAAAfHx/r6+v///////////////////////////////////////////+ampoAAAAJCQnOzs7////////////////d3d3a2tr////////////////////////9/f1paWm0tLT////////////5+flSUlIAAAAAAABGRkb19fX///////////////////////////////////////+bm5sAAAAAAAAAAAAAAAB8fHz///////////////////////////////////////9zc3MAAAAAAAAAAAAAAABMTEz///////////////////////////////////////+2trYAAAAAAAAAAAAAAACRkZH///////////////////////////////////////////+YmJgXFxcNDQ11dXX8/Pz///////////////////////////////////////////////////////////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

;}

;{21. Das Ende - OnExit Sprunglabel und OnError-Funktion

DasEnde(ExitReason, ExitCode) {

	OnExit("DasEnde", 0)

  ; Anzeigestatus des Infofenster sichern
	If Addendum.AddendumGui {
		admGui_SaveStatus()
		Gui, adm: Destroy
		Gui, adm2: Destroy
	}

	FormatTime	, Time      	, % A_Now    	, dd.MM.yyyy HH:mm:ss
	FormatTime	, TimeIdle	, % A_TimeIdle	, HH:mm:ss
	FileAppend	, % Time ", " A_ScriptName ", " A_ComputerName ", " ExitReason ", " ExitCode ", " TimeIdle "`n", % AddendumDir "\logs'n'data\OnExit-Protokoll.txt"

	For i, hook in Addendum.Hooks
		UnhookWinEvent(hook.hEvH, hook.HPA)

	gosub istda

}

istda: ;{

	Progress	, % "B2 B2 cW202842 cBFFFFFF cTFFFFFF zH25 w400 WM400 WS500"
					, % (onlyReload ? "Addendum wird neu gestartet...." : "Addendum wird beendet.")
					, % "Addendum für AlbisOnWindows"
					, % "by Ixiko vom " DatumVom
					, % "Futura Bk Bt"

	Loop 50 {
		Progress % (100 - (A_Index * 2))
		Sleep 1
	}

	Progress, Off

	Gdip_Shutdown(pToken)

	If onlyReload
		Reload

ExitApp
;}

;}

;{22. #Include(s)

;------------ allgemeine Bibliotheken ---------------------------------
#Include %A_ScriptDir%\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\lib\class_CtlColors.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_socket.ahk
#Include %A_ScriptDir%\..\..\lib\crypt.ahk
#Include %A_ScriptDir%\..\..\lib\Explorer_Get.ahk
#Include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
#Include %A_ScriptDir%\..\..\lib\TrayIcon.ahk
#Include %A_ScriptDir%\..\..\lib\Sci.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\Vis2.ahk
;#Include %A_ScriptDir%\..\..\lib\WatchFolder.ahk

;------------ Addendumbibliotheken für Albis on Windows --------
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gui.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Ini.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_LanCom.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Misc.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PopUpMenu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

;------------ Adddendum.ahk zusätzliche Funktionen /Labels -----
#include %A_ScriptDir%\include\ClientLabels.ahk
#include %A_ScriptDir%\include\MenuSearch.ahk
#include %A_ScriptDir%\include\NoTrayOrphans.ahk

;------------ HOTSTRINGS ---------------------------------------------
;#Include %A_ScriptDir%\include\AutoDiagnosen.ahk

;}





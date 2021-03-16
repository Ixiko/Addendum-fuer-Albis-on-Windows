; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . .                                                                                       	ADDENDUM HAUPTSKRIPT
global                                                                               AddendumVersion:= "1.53" , DatumVom:= "14.03.2021"
; . . . . . . . . . .
; . . . . . . . . . .                                    ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . . . . . .                                           BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . . . . . .                                                        RUNS ONLY WITH AUTOHOTKEY_H IN 32 OR 64 BIT UNICODE VERSION
; . . . . . . . . . .                                                                 THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .   !!!!!!!!!    !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !!     !!!!!!!!!
; . . . . . . . . . .                                                                    THIS SCRIPT ONLY WORKS WITH AUTOHOTKEY_H V1.1.32.00+
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

/*               	A DIARY OF CHANGES
| **14.03.2020** | **F+~** | **Laborjournal**        	- 	es wurden mehr Parameter im Labor angezeigt als eingestellt war <br>
																				-	Zeigt die Laborwerte der letzten n Werktage an. Es kommt somit immer eine feste Anzahl an Tagen zur Darstellung. <br>**Addendum V1.53**|
| **13.03.2020** | **F~**    | **Addendum**           	- 	der Auto-Neustart des Skriptes setzt das Albis-Programmdatum wieder auf das aktuelle Tagesdatum
																					-	der Dialog 'Labor Anzeigegruppen' wird abgefangen um die Steuerelemente zu vergrößern
																					-	|
| **10.03.2020** | **F+**    | **Laborabruf_iBWC**   	- 	Laborabruf wartet das Öffnen des Laborbuches ab und löst dann 'alle übertragen' aus. Addendum übernimmt dann die
									    												Übertragung ins Laborblatt.
										     											Voraussetzung ist, das bei Albis unter Optionen/Labor im Reiter Import bei 'Laborbuch nach Import automatisch öffnen'
											      										ein Häkchen gesetzt ist und das in Addendum die Einstellung Laborabruf automatisieren ebenso abgehakt ist.
												     									Anmerkung: 	Das Übertragen der Werte in das Laborblatt wird in einigen Fällen nicht komplett erfolgen können.
													    													Einige seltene Dialogfenster sind noch nicht erfasst und automatisiert worden durch mich.  |
| **10.03.2020** | **F~** | **Infofenster**           	- 	RPA des Sumatra Readers funktioniert jetzt tadellos. Funktionen optimiert und fehlerhafte Listviewzugriff behoben. |

| **14.03.2020** | **F+** | **Addendum_Datum**  	- 	hat mehrere neue Funktionen erhalten |
| **14.03.2020** | **F+** | **Addendum_Mining** 	- 	Verbesserte Datumauswertung. Daten die Patienten Geburtstage sind werden aussortiert. |
| **13.03.2020** | **F~** | **Addendum_DBase**  	- 	**SearchExt()** - hat Callback Funktionalität erhalten, Fehlerbehebung beim Stringvergleich |
| **11.03.2020** | **F~** | **Addendum_Albis**     	- 	**AlbisKeineChipkarte()**<br>zum schnellen Schließen des Dialoges: "Patient hat in diesem Quartal seine Chipkarte noch
																					nicht vorgezeigt"<br>
																					**AlbisNeuerSchein()**<br>1. Teil von Funktionen zum Anlegen eines neuen Abrechnungsscheines. AlbisNeuerSchein öffnet
																					und schließt das Fenster
																					**AlbisResizeLaborAnzeigegruppen()**<br>Funktion erweitert die Listboxsteuerelemente im Dialog 'Labor Anzeigegruppen' |

	CGM_ALBIS DIENST Service SID:                    	S-1-5-80-845206254-3503829181-3941749774-3351807599-4094003504
	CGM_ALBIS_BACKGROUND_SERVICE_(6002) 	S-1-5-80-4257249827-193045864-994999254-1414716813-2431842843
	CGM_ALBIS_SMARTUPDATE_SERVICE				S-1-5-80-4281495583-391623409-1399029959-4115513306-2324107004

	Ifap
	Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{3BA8167C-E0D3-4020-AFAF-BD0AF2BCB984}\1.1\0\win32
	Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{3BA8167C-E0D3-4020-AFAF-BD0AF2BCB984}\1.1\HELPDIR


** Ideen
	OrderEntry Labor - automatische Eintragung der richtigen Ausnahmekennziffer


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
		#MaxMem                    	, 512
		#WarnContinuableException, Off

		;#Warn ;, All

		;ListLines                        	, % (ListLinesFlag := false)
		ListLines                        	, Off

		DetectHiddenWindows   	, On
		SetTitleMatchMode        	, 2	              	;Fast is default
		SetTitleMatchMode        	, Fast        		;Fast is default

		CoordMode, Mouse         	, Screen
		CoordMode, Pixel            	, Screen
		CoordMode, ToolTip       	, Screen
		CoordMode, Caret        	, Screen
		CoordMode, Menu        	, Screen

		SetBatchLines                	, -1
		SetWinDelay                    	, -1
		SetControlDelay            	, -1
		AutoTrim                       	, On
		FileEncoding                 	, UTF-8


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

	; Tray Icon erstellen
		If (hIconAddendum := Create_Addendum_ico())
			Menu, Tray, Icon, % "hIcon: " hIconAddendum

	; wichtige Hotkeys sollten an dieser Stelle bleiben
		Hotkey, $^!ä  	, BereiteDichVor                                                         	;= Überall: beendet das Addendumskript
		Hotkey, $^!n   	, SkriptReload                                                            	;= Überall: Addendum.ahk wird neu gestartet

;}

;{02. Variablendefinitonen , individuelles Client-Traymenu wird hier erstelltc
	;----------------------------------------------------------------------------------------------------------------------------------------------
	; globale Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	  ; allgemeine globale Variablen
		global AlbisWinID      	:= AlbisWinID()                      	; für alle Funktionen muss die ID des Albis Hauptfenster ohne weitere Aufrufe als Variable vorhanden sein
		global AlbisPID          	:= AlbisPID()                         	; mal sehen vielleicht benötige ich die Process ID von Albis für irgendwas
		global Albis               	:= Object()                           	; enthält aktuelle Fenstergrößen und andere Daten aus den Albisfenstern
		global AutoDocCalled	:= 0                                      	; flag für AutoDoc Funktion welche nicht mehrfach gestartet werden kann (liest mehrfach aus der Ini aber nie richtig)
		global JSortDesc, JSortCol                                           	; Sortierung des Journals im Addendum Infofenster
		global DatumZuvor
		global q                     	:= Chr(0x22)                           	; ist dieses Zeichen -> "
		global admServer                                                       	; für Interskriptkommuniktion zwischen verschiedenen Computern im LAN
		global AddendumDir                                                 	; Skriptverzeichnis

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)

	 ; Hook's
		global EHProc		    		:= ""                                  	; der Name des Prozesses der gehookt wurde
		global EHStack              	:= Array()
		global SHookHwnd	    	:= 0                                   	; hook handle des ShellHookProc
		global HmvHookHwnd  	:= 0                                   	; hook handle des für das Heilmittelverordnungsfensters (Physio)
		global HmvEvent		    	:= 0                                 	; event Nummer für Heilmittelverordnungsfenster
		global ifaptimer               	:= 0                                    	; ifap produziert eine Kaskade an Fehlermeldungen, flag wird bei erster Fehlermeldung gesetzt
		global AlbisHookTitle                                                 	; enthält den aktuellen Fenstertitel des Albisfenster
		global EHWHStatus        	:= false                               	; flag falls gerade noch der Hookhandler läuft
		global ClickReady            	:= 1                             		; für den GVU Formularausfüller
		global AutoLogin            	:= true                                 ; der Autologin Vorgang für Albis kann per Interskriptkommunikation ausgestellt werden

	  ; Daten sammelnde Objekte
		global oPat                    	:= Object()                       	; Patientendatenbank Objekt für Addendum
		global TProtokoll				:= Array()                         	; alternatives Tagesprotokoll als Absicherung falls es wieder Probleme mit der Albisdatenbank gibt
		global ScanPool             	:= Array()                           	; enthält die Dateinnamen aller im BefundOrdner vorhandenen Dateien
		global GVU                   	:= Object()                         	; enthält Daten für die Vorsorgeautomatisierung
		global SectionDate					                                	; [Section]-Bezeichnung in der Tagesprotokoll Datenbank

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; ** EINSTELLUNGEN AUS DER ADDENDUM.INI EINLESEN / ADDENDUM OBJEKT MIT DATEN BEFÜLLEN **
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	/*                    	Beschreibung aller enthaltenen key : value - Paare

	             ~~~~~~~ Struktur des Addendum-Objektes ~~~~~~~
		Addendum.BefundOrdner  	= Scan-Ordner für neue Befundzugänge
		Addendum.iWin                	= Infofenster im Albis Patientenfenster
	                                                     .Y,.W,.H     	- der Einblendungsbereich
	                                                	 .LVScanPool - Objekt enthält die größe der ScanPool Listview (W,H)
		Addendum.Telegram          	= enumeriertes Objekt (Bsp: Addendum.Telegram[1].TBotID)
	                                                	 .BotName 	- Name des Bot's
			                                           	 .Token         - Telegram Bot ID
		                                               	 .ChatID     	- Telegram Chat ID
		Addendum.GVUListe          	= Anzahl der vorhandenen Untersuchungen in der GVU-Liste

	 */

		global Addendum	:= Object()
		Addendum.Debug := false                                   	; gibt Daten in meine Standardausgabe (Scite) aus wenn = true

	;	- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	;	ein Aufruf nur mit dem Verzeichnisnamen initialisiert den Ini-Pfad der Funktion
		workini := IniReadExt(AddendumDir "\Addendum.ini")
	;	- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	; Einstellungen - Funktionen aus \include\Addendum_Ini.ahk  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		admObjekte()                                                    	; Definition Subobjekte + diverse Variablen
		admVerzeichnisse()                                               	; Programm- und Datenverzeichnisse
		admDruckerStandards()            	                         	; Standarddrucker A4/PDF/FAX - Einstellungen werden vom PopUpMenu Skript benötigt
		admFensterPositionen()                                       	; verschiedene Fensterpositionen
		admFunktionen()                                               	; zu- und abschaltbare Funktionen
		admGesundheitsvorsorge()                                	; Abstände zwischen Untersuchungen minimales Alter
		admInfoWindowSettings()                                    	; Infofenster Einstellungen
		admModule()                                                     	; Addendum Skriptmodule
		admTools()                                                   			; externe Programme die z.B. über das Infofenster gestartet werden können
		admLaborDaten()                                              	; Laborabruf, Verarbeitung Laborwerte
		admLanProperties()                                             	; LAN - Kommunikation Addendumskripte
		admMailAndHolidays()                                       	; Mailadressen, Urlaubszeiten
		admMonitorInfo()                                                	; ermittelt die derzeitige Anzahl der Monitore und deren Größen
		admPIDHandles()                     	                        	; Prozess-ID's
		admPDFSettings()					                            	; PDF anzeigen, signieren
		admOCRSettings()                                              	; OCR - cmdline Programme
		admPraxisDaten()                                               	; Praxisdaten   -   für Outlook, Telegram Nachrichtenhandling, Sprechstundenzeiten, Urlaub
		admShutDown()                                                	; Einstellungen für automatisches Herunterfahren des PC
		admSicherheit()                                                   	; Lockdown Funktion - Nutzer hat den Arbeitsplatz verlassen
		admSonstiges()
		admStandard()                                                    	; Standard-Einstellungen für Gui's
		admTagesProtokoll()                                           	; Tagesprotokoll laden
		admTelegramBots()				                            	; Telegram Bot Daten
		admThreads()                                                    	; Skriptcode für Multithreading (z.B. Tesseract OCR)

		ChronikerListe()                                                    	; Chroniker Index
		GeriatrischListe()                                                   	; Geriatrie Index

		GeriatrieICD := [	"Ataktischer Gang {R26.0G};"        	, "Spastischer Gang {R26.1G};"                         	, "Probleme beim Gehen {R26.2G};"
								, 	"Bettlägerigkeit {R26.3G};"               	, "Immobilität {R26.3G};"                                  	,"Standunsicherheit {R26.8G};"
								, 	"Ataxie {R27.0G};"                        	, "Koordinationsstörung {R27.8G};"                    	, "Multifaktoriell bedingte Mobilitätsstörung {R26.8G};"
								,	"Hemineglect {R29.5G};"         	    	, "Sturzneigung a.n.k. {R29.6G};"                      	, "Harninkontinenz {R32G};"
								,	"Stuhlinkontinenz {R15G};"            	, "Überlaufinkontinenz {N39.41G};"                   	, "n.n. bz. Harninkontinenz {N39.48G};"
								,	"Dysphagie {R13.9G};"                                                                                              	, "Dysphagie mit Beaufsichtigungspflicht während der Nahrungsaufnahme {R13.0G};"
								,                                                                                                                                   	  "Dysphagie bei absaugpflichtigem Tracheostoma mit geblockter Trachealkanüle {R13.1G};"
								,	"Orientierungsstörung {R41.0G};"  	, "Gedächtnisstörung {R41.3G};"                       	, "Vertigo {R42G};"
								, 	"chron. Schmerzsyndrom {R52.2G};"	, "chron. unbeeinflußbarer Schmerz {R51.1G};"  	, "Chronische Schmerzstörung mit somatischen und psychischen Faktoren {F45.41G};"
								,	"Multiinfarktdemenz {F01.1G};"      	, "Subkortikale vaskuläre Demenz {F01.2G};"    		, "Gemischte kortikale und subkortikale vaskuläre Demenz {F01.3G};"
								,	"Vaskuläre Demenz {F01.9G};"      	, "Dementia senilis {F03G};"
								,	"Vorliegen eines Pflegegrades {Z74.9G};"
								,	"Demenz bei Alzheimer-Krankheit mit frühem Beginn (Typ 2) [F00.0*] {G30.0+G};Alzheimer-Krankheit, früher Beginn {G30.0G};"
								,	"Demenz bei Alzheimer-Krankheit mit spätem Beginn (Typ 1) [F00.1*] {G30.1+G};Alzheimer-Krankheit, später Beginn {G30.1G};"
								,	"Primäres Parkinson-Syndrom mit mäßiger Beeinträchtigung ohne Wirkungsfluktuation {G20.10G};"
								, 	"Primäres Parkinson-Syndrom mit mäßiger Beeinträchtigung mit Wirkungsfluktuation {G20.11G};"
								, 	"Primäres Parkinson-Syndrom mit schwerster Beeinträchtigung ohne Wirkungsfluktuation {G20.20G};"
								, 	"Primäres Parkinson-Syndrom mit schwerster Beeinträchtigung mit Wirkungsfluktuation {G20.21G};"]


	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tray Menu erstellen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		func_NoNo                     	:= Func("NoNo")
		func_ZeigePatDB             	:= Func("ZeigePatDB")
		func_ZeigeFehlerProtokoll	:= Func("ZeigeFehlerProtokoll")
		func_AddendumObjekt		:= Func("AddendumObjektSpeichern")

		Menu, Tray, NoStandard
		Menu, Tray, Tip					, % StrReplace(A_ScriptName, ".ahk") " V." AddendumVersion " vom " DatumVom
													. "`nClient: " compName SubStr("                   ", 1, Floor((20 - StrLen(compName))/2))
													. "`nPID: " DllCall("GetCurrentProcessId")
													. "`nAutohotkey.exe: " A_AhkVersion

		Menu, Tray, Color            	, % "c" Addendum.Default.BGColor3
		Menu, Tray, Add				, Addendum, % func_NoNo
		Menu, Tray, Icon				, Addendum, % "hIcon: " hIconAddendum
	;}

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
					iconpath := Addendum.Dir "\" iconpath

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
					iconpath := Addendum.Dir "\" iconpath

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

		Menu, SubMenu2, Add	, % "Albis reanimieren"                 	, ToolStarter
		Menu, SubMenu2, Icon	, % "Albis reanimieren"                 	, % Addendum.Dir "\assets\ico\Windows.ico"
		Menu, SubMenu2, Add	, % "Windows komplett neu starten", ToolStarter
		Menu, SubMenu2, Icon	, % "Windows komplett neu starten", % Addendum.Dir "\assets\ico\Windows.ico"
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; per Checkbox zu- oder abschaltbare Funktionen (SubMenu3)
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; AutoSizer für Ifap, MS Word, FoxitReader, Sumatra PDF          	;{
		For classnn, appname in Addendum.Windows.Proc {
			Menu, SubMenu31, Add, % appname, Menu_AutoSize
			Menu, SubMenu31, Icon, % appname, % A_ScriptDir "\res\" appname ".ico"
			Menu, SubMenu31, % (Addendum.Windows[appname].AutoPos & Addendum.MonSize ? "Check":"UnCheck"), % appname
		}
		Menu, SubMenu3, Add, % "AutoSizer", :SubMenu31
	;}

	; automatische Anpassung von Albis an die Monitorauflösung 	;{
		Menu, SubMenu3, Add, % "Albis AutoSize"     	, Menu_AlbisAutoPosition
		Menu, SubMenu3, % (Addendum.AlbisLocationChange ? "Check":"UnCheck")	, % "Albis AutoSize"
	;}

	; Automatisierung von GVU Formulare                                      	;{
		Menu, SubMenu3, Add, % "Albis GVU automatisieren", Menu_GVUAutomation
		Menu, SubMenu3, % (Addendum.GVUAutomation ? "Check":"UnCheck")        	, % "Albis GVU automatisieren"
	;}

	; Automatisierung der PDF Signierung mit dem FoxitReader        	;{
		Menu, SubMenu3, Add, % "FoxitReader Signaturhilfe"                                  	, Menu_PDFSignatureAutomation
		Menu, SubMenu3, Add, % "FoxitReader Dokument automatisch schliessen"    	, Menu_PDFSignatureAutoTabClose
		Menu, SubMenu3, % (Addendum.AutoCloseFoxitTab ? "Check":"UnCheck")      , % "FoxitReader Dokument automatisch schliessen"
		Menu, SubMenu3, % (Addendum.PDFSignieren ? "Check":"UnCheck")          	, % "FoxitReader Signaturhilfe"
		Menu, SubMenu3, % (Addendum.PDFSignieren ? "Enable":"Disable")             	, % "FoxitReader Dokument automatisch schliessen"
	;}

	; Laborabruf automatisieren                                                   	;{
		Menu, SubMenu3, Add, % "Laborabruf automatisieren", Menu_LaborAbrufManuell
		Menu, SubMenu3, % (Addendum.Labor.AutoAbruf ? "Check":"UnCheck")       	, % "Laborabruf automatisieren"

		; zeitgesteuerter Abruf
		;  dieses Menu ist nur auf Clients sichtbar, welche in der Addendum.ini
		;  unter Labor/RunOnlyClients hinterlegt sind (Komma getrennte Liste)
		LabCallClients := StrReplace(Addendum.Labor.ExecuteOn, " ")
		If compname in %LabCallClients%
		{
			Menu, SubMenu3, Add, % "zeitgesteuerter Laborabruf", Menu_LaborAbrufTimer
			Menu, SubMenu3, % (Addendum.Labor.AbrufTimer ? "Check":"UnCheck")   	, % "zeitgesteuerter Laborabruf"
			If Addendum.Labor.AbrufTimer {
				If !Addendum.Labor.AutoAbruf
					IniWrite, % "ja"  	, % Addendum.Ini, % compname, % Laborabruf_automatisieren
				Addendum.Labor.AutoAbruf	:= true                                       ; Timer ist an dann auch  AutoAbruf
			}
		}
		else If Addendum.Labor.AbrufTimer {
				Addendum.Labor.AbrufTimer	:= false
				IniWrite, % "nein"	, % Addendum.Ini, % compname, % Laborabruf_automatisieren
			}
	;}

	; Addendum Toolbar                                                             	;{
		Menu, SubMenu3, Add, % "Addendum Toolbar", Menu_AddendumToolbar
		Menu, SubMenu3, % (Addendum.ToolbarThread ? "Check":"UnCheck")        	, % "Addendum Toolbar"
	;}

	; AddendumGui - Infofenster                                                    	;{
		Menu, SubMenu3, Add, % "Addendum Infofenster", Menu_AddendumGui
		Menu, SubMenu3, % (Addendum.AddendumGui ? "Check":"UnCheck")        	, % "Addendum Infofenster"
	  ;}

	; WatchFolder                                                                         	;{
		Menu, SubMenu3, Add, % "Befundordner überwachen", Menu_WatchFolder
		Menu, SubMenu3, % (Addendum.OCR.WatchFolder ? "Check":"UnCheck")     	, % "Befundordner überwachen"
	;}

	; Schnellrezept                                                                       	;{
		Menu, SubMenu3, Add, % "Schnellrezept", Menu_Schnellrezept
		Menu, SubMenu3, % (Addendum.Schnellrezept ? "Check":"UnCheck")           	, % "Schnellrezept"
	;}

	; MicroDicom Dateiexport automatisieren                                 	;{
		If GetAppImagePath("mDicom.exe") {   	; prüft ob MicroDicom installiert ist
			Menu, SubMenu3, Add, % "MicroDicom Export", Menu_MDicom
			Menu, SubMenu3, % (Addendum.mDicom ? "Check":"UnCheck"), % "MicroDicom Export"
		} else {
			Addendum.mDicom := false
		}
	;}

	; PopUpMenu                                                                          	;{
		Menu, SubMenu3, Add, % "PopUpMenu", Menu_PopUpMenu
		Menu, SubMenu3, % (Addendum.PopUpMenu ? "Check":"UnCheck"), % "PopUpMenu"
	;}

	; AutoOCR                                                                            	;{
		If (Addendum.OCR.Client = compname) {
			Menu, SubMenu3, Add, % "AutoOCR", Menu_AutoOCR
			Menu, SubMenu3, % (Addendum.OCR.AutoOCR ? "Check":"UnCheck"), % "AutoOCR"
		}
	;}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; weitere Menupunkte
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; DATEN / PROTOKOLLE

		Menu, SubMenu4, Add, % "Patienten Datenbank"              	, % func_ZeigePatDB
		Menu, SubMenu4, Add
		Menu, SubMenu4, Add, % "aktuelles Fehlerprotokoll"        	, % func_ZeigeFehlerProtokoll
		Menu, SubMenu4, Add, % "Addendum Objekt"                	, % func_AddendumObjekt

	; Protokolle/Datenbanken anzeigen
		If FileExist(protokoll := Addendum.DBPath           	"\Labordaten\LaborabrufLog.txt")            	{
			func_LaborAbrufLog  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Laborabruf Protokoll"           	, % func_LaborAbrufLog
		}
		If FileExist(protokoll := Addendum.DBPath          	"\Labordaten\LaborimportLog.txt")          	{
			func_LaborImportLog  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Laborimport Protokoll"           	, % func_LaborImportLog
		}
		If FileExist(protokoll := Addendum.Dir                 	"\logs'n'data\OnExit-Protokoll.txt")           	{
			func_ZeigeOnExitProtokoll := Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "OnExit Protokoll"                	, % func_ZeigeOnExitProtokoll
		}
		If FileExist(protokoll := Addendum.DBPath           	"\sonstiges\PDfImportLog.txt")                   	{
			func_PDFImportLog  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "PDF Importprotokoll"            	, % func_PDFImportLog
		}
		If FileExist(protokoll := Addendum.DBPath          	"\OCRTime_Log.txt")                              	{
			func_tessOCRLog  	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "PDF OCR Protokoll"            	, % func_tessOCRLog
		}
		If FileExist(protokoll := Addendum.DBPath          	"\sonstiges\AddendumMonitorLog.txt")     	{
			func_tessMonLog   	:= Func("ShowTextProtocol").Bind(protokoll)
			Menu, SubMenu4, Add, % "Addendum Monitor"            	, % func_tessMonLog
		}

	; FARBEN festlegen
		Menu, SubMenu1, Color                            	, % "c" Addendum.Default.BGColor3
		Menu, SubMenu2, Color                            	, % "c" Addendum.Default.BGColor3
		Menu, SubMenu3, Color                            	, % "c" Addendum.Default.BGColor3
		Menu, SubMenu4, Color                            	, % "c" Addendum.Default.BGColor3

	; MENU erstellen
		Menu, Tray, Add, Module starten/beenden   	, :SubMenu1
		Menu, Tray, Add, Daten / Protokolle             	, :SubMenu4
		Menu, Tray, Add, andere Tools                   	, :SubMenu2
		Menu, Tray, Add, Einstellungen                     	, :SubMenu3

		Menu, Tray, Icon, Module starten/beenden 	, % Addendum.Dir "\assets\ModulIcons\Tools2.ico"
		Menu, Tray, Icon, Daten / Protokolle             	, % Addendum.Dir "\assets\ModulIcons\Daten.ico"
		Menu, Tray, Icon, andere Tools                   	, % Addendum.Dir "\assets\ModulIcons\Tools1.ico"
		Menu, Tray, Icon, Einstellungen                   	, % Addendum.Dir "\assets\ModulIcons\Einstellungen.ico"

		Menu, Tray, Add,
		Menu, Tray, Add, ListLines On/Off                	, scriptLine
		Menu, Tray, Add, Zeige Skript Variablen       	, scriptVars
		Menu, Tray, Add, Skript Neu Starten             	, SkriptReload
		Menu, Tray, Add, Beenden                            	, DasEnde

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Einlesen der Patientendaten aus der Patienten.txt Datei im Addendum Datenbankordner
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		oPat := admDB.ReadPatDB() ; oPat - ist ein Objekt das die aktuell Daten zu registriertem Patientienten enthält
	;}


;}


;{04. Timer, Threads, Initiliasierung WinEventHooks, OnMessage, TCP-Kommunikation, eigene Dialoge anzeigen

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Starten clientunabhängiger Timer Labels
		;SetTimer, UserAway, 300000                                         	; zum Ausführen von Berechnungen wenn ein Computer unbenutzt ist (5min)
		If IsLabel(nopopup_%compname%)
			SetTimer, nopopup_%compname%, 2000                   	; spezielle Labels die Funktionen nur auf einzelnen Arbeitsplätzen ausführen
		If IsLabel(seldom_%compname%)
			SetTimer, seldom_%compname%, 900000		            	; Ausführung alle 15min

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Skript neu starten um 0 Uhr
		func_AutoReload := Func("SkriptReload").Bind("AutoRestart")
		diffTime := GetSeconds("240000") - GetSeconds(A_Hour A_Min A_Sec )
		SetTimer, % func_AutoReload, % -1*(diffTime*1000)

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Laborabruf Timer
		LaborAbrufTimerStarten()

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; SetWinEventHooks /Shellhook initialisieren
		InitializeWinEventHooks()

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Skriptkommunikation / Anpassen von Fenstern bei Änderung der Bildschirmaufläsung
		OnMessage(0x4A	, "Receive_WM_COPYDATA")                	; Interskriptkommunikation
		OnMessage(0x7E	, "WM_DISPLAYCHANGE")                    	; Änderung der Bildschirmauflösung

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; TCP Server starten - für LAN Kommunikation
		;~ global admServer
		;~ admStartServer()

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Infofenster sofort anzeigen wenn eine Akte/Karteikarteikarte geöffnet ist
	; Überwachung des Befundordners einschalten
		Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If InStr(Addendum.AktiveAnzeige, "Patientenakte")         	{
			PatDb(AlbisTitle2Data(AlbisTitle), "exist")
			If Addendum.AddendumGui {
				If Addendum.OCR.WatchFolder {
					WatchFolder(Addendum.BefundOrdner, "admGui_FolderWatch", False, (1+2+8+64))
					Addendum.OCR.WatchFolderStatus := "running"
				}
				If RegExMatch(Addendum.AktiveAnzeige, "i)Karteikarte|Laborblatt|Biometriedaten")				; WinActive("ahk_class OptoAppClass") &&
					AddendumGui()
			}
		}

	; WorkingSet freigeben nach dem ersten Start
		ScriptMem()

;}

;{05. Hotkey command's

HotkeyLabel:
	;=--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Hotkey's die überall funktionieren
	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	fn_LVCopy	:= Func("ListviewClipboard")

	Hotkey, $^Ins               	, RunSomething                                 	;= Überall: für's programmieren gedacht, im Label das Skript eintragen das per Hotkey gestartet werden soll
	Hotkey, !m                    	, MenuSearch                                    	;= Überall: Menu Suche - funktioniert in allen Standart Windows Programmen
	Hotkey, !q                     	, ScreenShot                                      	;= Überall: erstellt einen Screenshot per Mausauswahl und legt diesen im Clipboard ab
	Hotkey, CapsLock          	, DoNothing                                      	;= Überall: Capslock kann nicht mehr versehentlich gedrückt werden
	Hotkey, ^!e                  	, EnableAllControls                            	;= Überall: inaktivierte Bedienfelder in externen Fenstern aktivieren
	Hotkey, ^!p                  	, PatientenkarteiOeffnen                       	;= Überall: eine selektierte Zahl wird zum Öffnen einer Patientenkarteikarte genutzt
	Hotkey, Pause               	, AddendumPausieren
	Hotkey, ^!F10               	, SendClipBoardText                           	;= Überall: sendet den Inhalt des Clipboards als simulierte Tasteneingabe (um z.B. wiederholt ein Passwort zu senden)
	Hotkey, $^+c		    		, % fn_LVCopy
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

	fn_FoxitFontSize    		:= Func("HoveredControl").Bind("classFoxitReader", "Edit4")
	func_DeInkrementer	:= Func("DeInkrementer")
	Hotkey, If, % fn_FoxitFontSize
	Hotkey, WheelUp 	         	, % func_DeInkrementer
	Hotkey, WheelDown         	, % func_DeInkrementer
	Hotkey, If

	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Albis
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	fn_VhK         	:= Func("AlbisAusfuellhilfe").Bind("Muster12", "Hotkey")
	fn_Undo     	:= Func("UnRedo").Bind("Undo")
	fn_Redo    	:= Func("UnRedo").Bind("Redo")
	fn_Redo    	:= Func("UnRedo").Bind("Redo")
	fn_cLP       	:= Func("isActiveFocus").Bind("contraction=lp|bg")
	fn_LPTexte		:= Func("Privatabrechnung_Begruendungen")


	Hotkey, IfWinActive, ahk_class OptoAppClass			            			;~~ ALBIS Hotkey - bei aktiviertem Fenster
	Hotkey, $^c			    		, AlbisKopieren					        			;= Albis: selektierten Text ins Clipboard kopieren
	Hotkey, $^x			    		, AlbisAusschneiden			        		   	;= Albis: selektierten Text ausschneiden
	Hotkey, $^v			    		, AlbisEinfuegen                                 	;= Albis: Clipboard Inhalt (nur Text) in ein editierbares Albisfeld einfügen
	Hotkey, $^a		        		, AlbisMarkieren                                	;= Albis: gesamten Textinhalt markieren
	;Hotkey, ^!F1					, AlbisFunktionen									;= Albis: Funktionen für Albis - Auswahl-Gui
	;Hotkey, If, InStr(AlbisGetActiveWindowType(), "Patientenakte")        	;~~ ALBIS Hotkey - nur auslösbar wenn eine Akte geöffnet ist
	Hotkey, !F5		        		, AlbisHeuteDatum			             		;= Albis: Einstellen des Programmdatum auf den aktuellen Tag
	Hotkey, ^F7                 	, GVUListe_ProgDatum                       	;= Albis: akt. Patienten in die GVU-Liste eintragen (Datum der Programmeinstellung)
	Hotkey,  !F7                  	, GVUListe_ZeilenDatum                       	;= Albis: akt. Patienten in die GVU-Liste eintragen (Zeilendatum wird genutzt )
	Hotkey, !F8		    			, DruckeLabor                                    	;= Albis: Laborblatt als pdf oder auf Papier drucken
	Hotkey, ^F9 		    		, Abrechnungshelfer			        			;= Albis:Abrechnungshelfer.ahk starten
	Hotkey, ^F11		    		, Schulzettel	         			        			;= Albis: Schulzettel.ahk starten
	Hotkey, ^F12		    		, Hausbesuche         			        			;= Albis: Hausbesuche.ahk starten
	Hotkey, !Down      			, AlbisKarteiKarteSchliessen        			;= Albis: per Kombination die aktuelle Karteikarte schliessen
	Hotkey, !Up               		, AlbisNaechsteKarteiKarte	        			;= Albis: per Kombination eine geöffnete Kartekarte vorwärts
	Hotkey, !Right    				, Laborblatt    										;= Albis: Laborblatt des Patienten anzeigen
	Hotkey, !Left  					, Karteikarte      									;= Albis: Karteikarte des Patienten anzeigen
	Hotkey, $^z	                 	, % fn_Undo                                      	;= Albis: Eingaben rückgängig machen in Steuerelementen (Edit, RichEdit)
	Hotkey, $^d	                 	, % fn_Redo                                       	;= Albis: Eingaben wiederherstellen
	Hotkey, $^!d                	, StarteAutoDiagnosen                        	;= Albis: Skript AutoDiagnosen3 starten
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

	Hotkey, If, % fn_cLP
	Hotkey, +F6                    	, % fn_LPTexte                                    	;= Begründungstexte bei Abrechnung nach GOÄ anzeigen, bearbeiten, verwenden
	Hotkey, If


	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; flexible Hotkey's und Fensterkombinationen, je nach Einstellung in der Addendum.ini
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	If Addendum.PDFSignieren 	{ ; Signatur setzen im Foxit Reader nach Hotkey

		func_FRSignature := func("FoxitReaderSignieren")
		Hotkey, IfWinActive, % "ahk_class " Addendum.PDF.ReaderWinClass
		Hotkey, % Addendum.PDFSignieren_Kuerzel, % func_FRSignature
		Hotkey, IfWinActive

	}

	;}

return

HoveredControl(WinNN, CtrlNN) {                  ;-- gibt true zurück wenn der Mauszeiger über einem bestimmten Steuerelement steht

	MouseGetPos, mx, my, hWin, hCtrl, 2
	WinGetClass , WinClass	, % "ahk_id " hWin
	If !InStr(WinClass, WinNN)
		return false
	fCtrl := Control_GetClassNN(hWin, hCtrl)
	If (fCtrl = CtrlNN)
		return true

return false
hCtrlOff:
	ToolTip, % "classnn: " fCtrl, % mx-10, % my+30, 12
	SetTimer, hCtrlOff, -1000
	ToolTip,,,,12
return
}

DeInkrementer() {

	MouseGetPos,,, hWin, hCtrl, 2
	ControlFocus,, % "ahk_id " hCtrl
	ControlGetText, value,, % "ahk_id " hCtrl
	If !RegExMatch(value, "^\d+$") {
		ControlSend,, (A_ThisHotkey="WheelUp" ? {Up}:{Down}), % "ahk_id " hWin
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

;}

;{06. Hotstrings

:*:#at::@

;{ 	6.1. -- ALBIS
; --- Leistungskomplexe                                                                  	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "lk")
; Stix
:*:sang::01737-                                                                           	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:Stuhl::01737-                                                                           	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:iFop::01737-                                                                           	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:haemof::01737-                                                                    	; Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:bz::32025                                                                             	; Blutzuckermessung
:*:glucose::32025                                                                      	; Blutzuckermessung
:*:zucker::32025                                                                       	; Blutzuckermessung
:*:stix::32030                                                                             	; Urinstix
:*:urin::32030                                                                           	; Urinstix
:*:ustix::32030                                                                           	; Urinstix
:*:urinstix::32030                                                                         	; Urinstix
:*:inr::32026-                                                                              	; INR Messstreifen
:*:gerinnung::32026-                                                                	; INR Messstreifen

; Gespräch
:*:gs::03230(x:2)                                                                         	; Gesprächziffer
:*:ps1::35110(x:2)                                                                       	; Psychosomatikziffer 1
:*:psy1::35110(x:2)                                                                      	; Psychosomatikziffer 1
:*:ps2::35110(x:2)                                                                       	; Psychosomatikziffer 2
:*:psy2::35110(x:2))                                                                  	; Psychosomatikziffer 2

; Kommunikation
:*:tel::80230(x:2)                                                                        	; Telefongebühr mit dem Krankenhaus
:*:fax::40120(x:2)                                                                        	; Faxgebühr
:*:term::03008-03010-                                                                	; Zuschläge Terminvermittlung (Terminvermittlung Facharzt, TSS-Terminvermittlung)
:*:Vermit::03008-03010-                                                             	; Zuschläge Terminvermittlung (Terminvermittlung Facharzt, TSS-Terminvermittlung)

; Vorsorgen
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
:*:aneu::01747                                                                          	; Aufklärungsgespräch Ultraschall-Screening Bauchaortenaneurysmen - nur Männer ab 65 Jahren einmalig berechnungsfähig!
:*:UUnt::01711-01712-01713-01714-01715-01716-01717-01718-01719-01723 	; Vorsorgeuntersuchung Kinder inkl. J1
:*:U1::01711                                                                             	; U1
:*:U2::01712                                                                             	; U2
:*:U3::01713                                                                             	; U3
:*:U4::01714                                                                             	; U4
:*:U5::01715                                                                             	; U5
:*:U6::01716                                                                             	; U6
:*:U7::01717                                                                             	; U7
:*:U8::01718                                                                             	; U8
:*:U9::01719                                                                             	; U9
:*:J1::01723                                                                             	; J1
:*:Kind::01711-01712-01713-01714-01715-01716-01717-01718-01719-01723 	; Vorsorgeuntersuchung Kinder inkl. J1
:*R:entw::03350-03351                                                               	; Orientierende entwicklungsneurologische Untersuchung eines Neugeborenen, Säuglings, Kleinkindes oder Kindes

; Formulare
:*:kur::01622                                                                               	; Kurplan, Gutachten, Stellungnahme
:*:gut::01622                                                                               	; Kurplan, Gutachten, Stellungnahme
:*:mdk::01622                                                                            	; Kurplan, Gutachten, Stellungnahme
:*:Muster52::01621                                                                     	; Muster 52

; Labor und Sonderziffern
:*:sars::32006-88240                                                                	; COVID19 - Laborbudget und Sonderziffer
:*:covid19::32006-88240                                                            	; COVID19 - Laborbudget und Sonderziffer

; Wunden/Verbände
:*:Kompr::02313                                                                         	; Kompressionsverband

; Hausbesuche
:*:Verwe::01440(x:2)                                                                    	; Verweilen außerhalb der Praxis
:*:hb0::01410-97234-03230(x:3)                                               	; Hausbesuch (bestellt)
:*:hb1::01411(um:19:00)-97234-03230(x:3)						     		; Hausbesuch nach Feierabend
:*:hb2::01412(um:19:00)-97234-03230(x:3)                             	; Hausbesuch dringend, sofort, Feiertage
:*:hbWE::01412(um:19:00)-97234-03230(x:3)                          	; Hausbesuch Wochenende
:*:hbFei::01412(um:19:00)-97234-03230(x:3)                          	; Hausbesuch Feiertag
:*:hbAbe::01412(um:19:00)-97237-03230(x:3)                          	; Hausbesuch dringend ausserhalb Sprechzeit
:*:hbSpre::01412(um:11:00)-97234-03230(x:3)                          	; Hausbesuch dringend, aus der Sprechstunde heraus

; OP-Vorbereitung/Nachbehandlung
:*:opvor::31010-31011-31012-31013                                         	; OP-Vorbereitung
:*:op vor::31010-31011-31012-31013                                       	; OP-Vorbereitung
:*:op-vor::31010-31011-31012-31013                                       	; OP-Vorbereitung
:*:kata::31012-31013-88115(OPS:5-144.a:RLB){Left}{LShift Down}{Left 3}{LShift Up} ; Vorbereitung Katarakt-OP auf Überweisung
:*R:postop::31600                                                                        	; postoperative Nachbehandlung (nur bei Überweisung vom Facharzt)

; Quartalsziffern
:*:DAK::93550-93555-93560-93565-93570                             	; HAP DAK
:*:1-::03000-03040                                                                   	; 03000-03040
:*:2-::03220                                                                              	; 03220
:*:3-::03221                                                                             	; 03221
:*:4-::03360                                                                             	; 03360
:*:5-::03362                                                                             	; 03362
:*:pall::03370-03371-03372-03373                                          	; Palliativziffern
:*:c1::03220                                                                                	;{ chronisch krank Kennzeichnung - automatisierte Überprüfung
	SendInput, % "{RAW}03220"
	InChronicList()
return
:*:c2::03221

;}
#If

;}
; --- Kuerzelfeld                                                                                	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("identify"), "Edit3")
:*:GVU::                                                                                        	;{ GVU,HKS Automatisierung

	ControlGetText, EditDatum, Edit2, ahk_class OptoAppClass
	If !RegExMatch(EditDatum, "\d{2}\.\d{2}\.\d{4}")	{
			PraxTT("Das Datumfeld konnte nicht ausgelesen werden", "3 3")
			return
	}

	DatumZuvor:= AlbisSetzeProgrammDatum(EditDatum)
	ControlFocus, Edit3, ahk_class OptoAppClass
	VerifiedSetText("Edit3", "GVU", "ahk_class OptoAppClass")
	SendInput, {Tab}

return ;}
:*b0:AISC::                                                                                   	;{ AISC - Laborwebinterface
	SendInput, {Tab}
	Sleep 50
	SendInput, {Tab}
return ;}

#If
;}
; --- Diagnosen                                                                              	;{
; bei Includes jetzt
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "dia")
:*:gb::                                                                                      	;{ geriatrischer Basiskomplex
	Auswahlbox(GeriatrieICD, "Diagnosenliste Geriatrie")
return ;}
:*:igpr::                                                                                    	;{ Impfung Grippe Privatpatienten

	SendMode, Event
	SetKeyDelay, 10, 10

	lpig := "1-5-375-(sach:Influenzaimpfstoff Influvac: 10.94)"
	lpdia := "Infektausschluß, A.v. {J06.9A};Impfung gegen Grippe, 2020/2021 (Influvac Tetra Ch: X17) {Z25.1G};"

	Linedate	:= AlbisLeseZeilenDatum(40, false)
	PrgDate	:= AlbisLeseProgrammDatum()
	AlbisSetzeProgrammDatum(LineDate)

	Sleep 300
	Send, % lpdia
	Send, {Tab}
	Sleep 3000
	Send, {Esc}
	Sleep 300
	Send, {Insert}
	Sleep 300
	Send, {Text}lp
	Send, {Tab}
	Sleep 300
	Send, % lpig
	Send, {Tab}
	PraxTT("Noch circa 4s braucht Albis für Berechnungen!")
	;WinWait
	Sleep 4000
	Send, {Esc}
	AlbisSetzeProgrammDatum(PrgDate)

	SetKeyDelay, 20, 40

return ;}
:*R:grdia::Infektausschluß, A.v. {J06.9A};Impfung gegen Grippe, 2020/2021 (Influvac Tetra Ch: X17) {Z25.1G};                      	; Influvac Preis
:*:covid-19na::COVID-19, Virus nachgewiesen {!U07.1G};{Left 2}{LShift Down}{Left}{LShift Up}
:*:covid-19ni::COVID-19, Virus nicht nachgewiesen {!U07.2G};{Left 2}{LShift Down}{Left}{LShift Up}
:*:covid19na::COVID-19, Virus nachgewiesen {!U07.1G};{Left 2}{LShift Down}{Left}{LShift Up}
:*:covid19ni:: COVID-19, Virus nicht nachgewiesen {!U07.2G};{Left 2}{LShift Down}{Left}{LShift Up}
:*:covid1::COVID-19, Virus nachgewiesen {!U07.1G};{Left 2}{LShift Down}{Left}{LShift Up}
:*:covid2::COVID-19, Virus nicht nachgewiesen {!U07.2G};{Left 2}{LShift Down}{Left}{LShift Up}
:*:sars::Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G};Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G};
:*X:sars2::LDSars()
:*:#ledder::M.Ledderhose, li. {M72.2LG};
:*:#hiatus h::Hernia diaphragmatica {K44.9G};
:*:#zwerchf::Hernia diaphragmatica {K44.9G};
:*:#harnw::Harnwegsinfektion {N39.0G};
:*:#hti::Harnwegsinfektion {N39.0G};
:*:#reflux::Refluxösophagitis {K21.0G};
:*:#akute Bel::Akute Belastungsreaktion {F43.0G};
:*:#Belast::Akute Belastungsreaktion {F43.0G};
:*:#TIA::TIA {G45.92G};
:*:#Hyperto::Benigne essentielle Hypertonie {I10.00G};
:*:#Diabetes neur::Diabetes mellitus mit neurologischen Komplikationen {+E14.40G}; Diabetische Polyneuropathie {*G63.2BG};
:*:#Diabetische Poly::Diabetes mellitus mit neurologischen Komplikationen {+E14.40G}; Diabetische Polyneuropathie {*G63.2BG};
:*:#Diab2::Diabetes mellitus Typ 2 {E11.90G}
:*:#Diabetes 2::Diabetes mellitus Typ 2 {E11.90G}
:*:#Diabetes Typ2::Diabetes mellitus Typ 2 {E11.90G}
:*:#Diabetes Typ 2::Diabetes mellitus Typ 2 {E11.90G}

#If
LDSars() {

	hparent := AlbisSendText("RichEdit20A1", "Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G}"
										  . ";Spezielle Verfahren zur Untersuchung auf infektiöse und parasitäre Krankheiten {Z11G};")

	WinWait, % "ALBIS ahk_class #32770", % "meldepflichtig", 2
	VerifiedClick("Button1", "ALBIS ahk_class #32770", "meldepflichtig",, true)

	SciTEOutput("2:" hParent)
	hParent := AlbisSendText("Edit6", "lko", hParent)

	SciTEOutput("3:" hParent)
		;return
	hParent := AlbisSendText("RichEdit20A1", "88240-32006", hParent)
	SciTEOutput("4:" hParent)
		;return
	hParent := AlbisSendText("Edit6", "**End", hParent)
	SciTEOutput("5:" hParent "`n")
		;return

}
;}
; --- Info                                                                                          	;{
#If (WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "info") )
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
#If
;}
; --- Rezept                                                                                      	;{
#If WinActive("Muster 16 ahk_class #32770") || WinActive("Privatrezept ahk_class #32770")
:*:#Trya::                                           	;{ Tryasol 30 ml
	IfapVerordnung("Tryasol 30 ml")
return ;}
#If
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
; --- Ausnahmeindikationen                                                              	;{
#If IsFocusedAusnahmeIndikation()                                                  	;
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

	; hinzufügende Leistungskomplexe mit Steuerelementinhalt vergleichen
		Loop, % toAdd.MaxIndex()	{
			If !InStr(ctext, toAdd[A_Index])				{
				If InStr(toAdd[A_Index], "03220") && InList(aktPatID, Chroniker)
					cText .= "-" toAdd[A_Index]
				else if InList(toAdd[A_Index], "03660,03362") && InList(aktPatID, Geriatrisch)
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
IsFocusedAusnahmeIndikation() {

	; das Fenster "weitere Information" trägt nur den Namen des Patienten
		WinGetText       	, activeWinText	, A
		ControlGetFocus	, activeControl	, A

	If InStr(activeWinText, "Ausnahmeindikation") && InStr(activeControl, "Edit14")
		return true

return false
}
InList(var, list) {

	If var in %list%
		return true

return false
}

;}
; --- Privatabrechnung                                                                   	;{
#If WinActive("ahk_class OptoAppClass") && (RegExMatch(AlbisGetActiveControl("contraction"), "lp|lbg") || InStr(AlbisGetActiveWindowType(), "Privatabrechnung"))  ; GOÄ [lp, lpg oder Privatabrechnung]
; Abstrichentnahme
:*R:Abstrich1::298-                                                                              	; Entnahme und gegebenenfalls Aufbereitung von Abstrichmaterial zur mikrobiologischen Untersuchung
:*R:AbstrichRa::298(ltext:Abstrich Rachenraum)-                                     	; Entnahme und gegebenenfalls Aufbereitung von Abstrichmaterial zur mikrobiologischen Untersuchung
:*R:AbstrichNa::298(ltext:Abstrich Nase)-                                              	; Entnahme und gegebenenfalls Aufbereitung von Abstrichmaterial zur mikrobiologischen Untersuchung

; Beratung
:*R:Bera1::1(fak:3,5:Beratung < 10 Minuten)-                                     	; Beratung < 10 Minuten
:*R:Bera2::3(fakt:2,3:Beratung mind. 10 Minuten)-                               	; Beratung mind. 10 Minuten
:*R:Bera3::3(fakt:3,5:Beratung < 20 Minuten)-                                      	; Beratung < 20 Minuten
:*R:Bera4::34(fakt:2,3:Beratung, eingehend mindestens 20 Minuten)-  	; Beratung, eingehend mindestens 20 Minuten
:*R:hb::50(dkm:4)-                                                                               	; Hausbesuch
:*R:verweil::56-                                                                                  	; Verweilen außerhalb der Praxis
:*R:zuschl::A-B-C-D-K1                                                                         	; Zuschläge
:*R:verbale::849-                                                                                 	; verbale Intervention - psychosomatik
:*R:verband::200-                                                                                	; Verband

; Diagnostik
:*R:lzRR::654                                                                                    		; Langzeitblutdruckmessung
:*R:lufu::605                                                                                    		; Lungenfunktion
:*R:Lungenf::605                                                                              		; Lungenfunktion
:*R:oxy::602-                                                                                    	; Pulsoxymetrie
:*R:Pulsoxy::602-                                                                                  	; Pulsoxymetrie
:*R:sono1::410-420                                                                             	; Ultraschall von bis zu drei weiteren Organen im Anschluss an Nummern 410 bis 418, je Organ. Organe sind anzugeben.
:*R:sono2::410(organ:Leber)-420(organ:Gb)-420(organ:Niere bds.)     	; Ultraschall mehrere Untersuchungen
:*R:sonoknie::410(organ:Kniegelenk li.)                                                	; Utraschall Knie
:*R:sono knie::410(organ:Kniegelenk li.)                                               	; Utraschall Knie

; Labor
:*R:BA::250-                                                                                       	; Blutabnahme
:*R:BSG::3501-                                                                                    	; BSG
:*R:BZ::3560-                                                                                    	; Blutzucker
:*R:Blutzuc::3560-                                                                                	; Blutzucker
:*R:glucose::3560-                                                                            	; Blutzucker
:*R:glukose::3560-                                                                                ; Blutzucker
:*:sang::3500-                                                                                    	;{ Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:Stuhl::3500-                                                                                   	;  Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:iFop::3500-                                                                                   	;  Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*:haemof::3500-                                                                                 ;} Ausgabe iFOPT Test (Stuhl auf Sanguis)
:*R:stix::3652-                                                                                     	;{ Urinstix
:*R:streifen::3652-                                                                               	; Urinstix
:*R:urin::3652-                                                                                   	; Urinstix
:*R:ustix::3652-                                                                                   	; Urinstix
:*R:urinstix::3652-					                                                             	;} Urinstix

; körperliche Untersuchung
:*R:gleichg::826-                                                                               	; neurologische Gleichgewichtsprüfung
:*R:neuro::800-                                                                                    	; eingehende neurologische Untersuchung
:*R:rekt::11-                                                                                      	; rektale Untersuchung
:*R:Unters1::7(fakt:2,3)-                                                                    	; Vollständige Untersuchung – ein Organsystem
:*R:Unters2::7(fakt:3,5)-                                                                    	; Vollständige Untersuchung – mehrere Organsysteme

; Medikamentengabe
:*R:infil1::267-                                                                                     	; Medikamentöse Infiltrationsbehandlung, je Sitzung
:*R:infil2::268-                                                                                     	; Medikamentöse Infiltrationsbehandlung im Bereich mehrerer Körperregionen
:*R:infusion k::271-                                                                              	; Infusion < 30 min
:*R:infu1::271-                                                                                    	; Infusion < 30 min
:*R:infusion l::                                                                                     	; Infusion > 30 min
:*R:infu2::                                                                                           	; Infusion > 30 min
:*R:inj sc::252-                                                                                     	; Injection s.c., i.m., i.c.
:*R:inj ic::252-                                                                                     	; Injection s.c., i.m., i.c.
:*R:inj im::252-                                                                                    	; Injection s.c., i.m., i.c.
:*R:inj iv::253-                                                                                    	; Injection i.v.
:*R:inr::3530-                                                                                  		; INR mit Gerät
:*R:Medik::76(ltext:Erstellung Medikationsplan)                                  		; Medikationsplan
:*R:Medpl::76(ltext:Erstellung Medikationsplan)                                  		; Medikationsplan

; Impfungen
:*R:impf::375-                                                                                    	; Impfung GOÄ Ziffer
:*R:influvac::(sach:Influenzaimpfstoff Influvac: 10.94)                            	; Impfung gegen Grippe Ziffern + Preis für Influvac
:*R:grlp::1-5-375-(sach:Influenzaimpfstoff Influvac: 10.94)                    	; Impfung gegen Grippe Ziffern + Preis für Influvac

; Anfragen
:*R:JVEG::(sach:Anfrage Sozialgericht gem. JVEG:21.00)                     	; Anfrage Sozialgericht
:*R:sozialgericht::(sach:Anfrage Sozialgericht gem. JVEG:21.00)          	; Anfrage Sozialgericht
:*R:lasv1::(sach:Landesamt für Soziales u. Versorgung:21.00)               	; Anfrage LaGeSo Schwerbehinderung Normal
:*R:lasv2::(sach:Kurzauskunft LA für Soziales u. Versorgung:5.00)          	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
:*R:lageso1::(sach:Landesamt für Soziales u. Versorgung:21.00)         	; Anfrage LaGeSo Schwerbehinderung Normal
:*R:lageso2::(sach:Kurzauskunft LA für Soziales u. Versorgung:5.00)      	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
:*R:lagesokurz::(sach:Kurzauskunft LA für Soziales u. Versorgung:5.00)  	; Anfrage LaGeSo Schwerbehinderung Kurzbefund
:*R:Rentenversich::(sach:Anfrage Rentenversicherung:28.20)                 	; Anfrage Rentenversicherung
:*R:RLV::(sach:Anfrage Rentenversicherung:28.20)                              	; Anfrage Rentenversicherung
:*R:DRV::(sach:Anfrage Rentenversicherung:28.20)                              	; Anfrage Rentenversicherung
:*R:Bundesa::(sach:Anfrage Bundesagentur für Arbeit gem. JVEG:32.50)	; Anfrage Rentenversicherung
:*R:Agentur::(sach:Anfrage Bundesagentur für Arbeit gem. JVEG:32.50) 	; Anfrage Rentenversicherung

; Porto
:*R:porto0::(sach:Porto Standard:0.60)                                                	; Postkarte
:*R:Postk::(sach:Porto Standard:0.60)                                                 	; Postkarte
:*R:porto1::(sach:Porto Standard:0.80)                                                	; Porto bis 20g
:*R:Standard::(sach:Porto Standard:0.80)                                             	; Porto bis 20g
:*R:porto2::(sach:Porto Kompakt:0.95)                                                 	; Porto bis 50g
:*R:Kompakt::(sach:Porto Kompakt:0.95)                                               	; Porto bis 50g
:*R:porto3::(sach:Porto Groß:1.55)                                                      	; Porto bis 500g
:*R:Groß::(sach:Porto Groß:1.55)                                                      	; Porto bis 500g
:*R:porto4::(sach:Porto Maxi:2.70)                                                       	; Porto bis 1000g
:*R:Maxi::(sach:Porto Maxi:2.70)                                                        	; Porto bis 1000g
:*R:porto5::(sach:Porto Maxi:4.50)                                                       	; DHL-Päckchen
:*R:Päck::(sach:Porto Maxi:4.50)                                                        	; DHL-Päckchen
:*:kopie::                                                                                           	;{ Berechnung Kopiekosten
	AlbisKopiekosten(0.5)
return ;}
:*:reiseunf::(sach:Reiseunfähigkeitsbescheinigung:20.00)                     	;{ Reiseunfähigkeitsbescheinigung
:*:reiserüc::(sach:Reiseunfähigkeitsbescheinigung:20.00)                    	;} Reiseunfähigkeitsbescheinigung
:*R:schreib::(sach:Schreibgebühr:3.50)                                                 	; Schreibgebühr

#If
;}
; --- Blockeingaben                                                                         	;{
;~ #If WinActive("ahk_class OptoAppClass") && AlbisGetActiveControl("contraction")
;~ :*:CORONA::                                                                              	;{ Corona Ziffern + Diagnose
;~ :*:COVID::

	;~ KKT := ["lko|88240-32006", "dia|Bronchitis {J06.9G}; Spezielle Verfahren zur Untersuchung auf SARS-CoV-2 {!U99.0G}; COVID-19, Virus nicht nachgewiesen {!U07.2G};"]
	;AlbisSchreibeSequenz(KKT)


;~ return ;}
;~ #If




;}
; --- manuelle Eingabe der Versichertendaten                                   	;{
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
:*:.t3::                                                                                                                       	;{ alle 3
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

:*:.so::SciTEOutput(){Left 1}
:R*:.afa::Addendum für Albis on Windows
:R*:.addendum::Addendum für Albis on Windows
:R*:.ambx::MsgBox, 1, Addendum für Albis on Windows -
:*:.fo::FileOpen(, "r", "UTF-8").Read(){Left 22}
:*:.FileOpen::FileOpen(, "r", "UTF-8").Read(){Left 22}
:*:.fopen::FileOpen(, "r", "UTF-8").Read(){Left 22}
:*:.jl::JSONData.Load(filepath, "", "UTF-8"){Left 14}{LShift Down}{Left 8}{LShift Up}
:*:.jw::JSONData.Save(filepath, oJSON, true,, 1, "UTF-8"){Left 27}{LShift Down}{Left 8}{LShift Up}

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
::.Bestellungen:: ;{
IniRead, bestelltext, % AddendumDir "\Addendum.ini", Addendum, Bestellung
Loop, Parse, bestelltext, ##
{
	If (StrLen(A_LoopField) > 0)
		Send, % "{Raw}" A_LoopField
	else
		Send, {Enter}
	Sleep, 25
}

VarSetCapacity(bestelltext, 0)
return ;}
;}

;{    	6.4.2 -- sonstige Kürzel
; Programmieren
:*:#ahk::Autohotkey
:*:#lahk::language:Autohotkey
:*:#sahk::site:Autohotkey
; Mail/Briefunterschriften
:*:#KVStemp::                                                                     	;{ Unterschrift für Mail oder Briefe an KV
	If Addendum.Praxis.KVStempel
		Send, % "{Text}" Addendum.Praxis.KVStempel
return ;}
:*:#MailStemp::                                                                   	;{ Unterschrift für Mail oder Briefe an Firmen, Patienten (ohne BSNR und LANR)
	If Addendum.Praxis.MailStempel
		Send, % "{Text}" Addendum.Praxis.MailStempel
return ;}
;}

;}

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

	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		If InStr(process.name, "FoxitReader") {
			RegExMatch(process.CommandLine, "\s\" q "(.*)?\" q, cmdline)
			break
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
	Send, {Space 4}
Return
;}
SendFourBackspace:      	;{	 	Strg (links) + Delete                                                         	(für Scite4Autohotkey)
	Send, {BackSpace 4}
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
	Txt				:= direction = - 1 ? string.Flip(sci.curText) : sci.curText
	SLen				:= StrLen(Txt)
	LineStartPos	:= endpos - SLen                                                                                			;lineendpos ist immer +2 wegen CR + LF Zeichen am Ende
	CPosInTxt		:= curpos - LineStartPos
	newPos			:= curpos + ( direction * InStr(Txt, ",", false, CPosInTxt + 1, 1) )

	SendMessage, 2025, % newPos, 0,, % "ahk_id " hSci
	;gosub MsgBoxer
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

	If InStr(AlbisGetActiveWinTitle(), "Wartezimmer") {
		ListviewClipboard("Wartezimmer")
		return
	}
	else if RegExMatch(WinGetClass(DLLCall("GetLastActivePopup", "uint", AlbisWinID())), "#32770") {
		ListviewClipboard("")
		return
	}

	AlbisCText := AlbisCopyCut()
	If (StrLen(AlbisCText) > 0) 	{
		clipboard := AlbisCText
		GdiSplash(SubStr(AlbisCText, 1, 30) "...", 1)
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
	SendMessage, 0x224,,,, % "ahk_id " AlbisMDIClientHandle()	; WM_MDINEXT
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
	AlbisLaborblattDrucken(1, "Standard", "")
return
;}
Laborblatt:                    	;{ 	Alt 	+ Rechts
	If !InStr(AlbisGetActiveWindowType(), "Patientenakte") or InStr(AlbisGetActiveWindowType(), "Laborblatt")
			return
	AlbisLaborBlattZeigen()
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

Privatabrechnung_Begruendungen(cmd="zeigen", bgrWahl="") {

	static hfocused, hparent, hfocused_last, hparent_last

	hfocused 	:= GetFocusedControlHwnd(AlbisWinID())
	hparent		:= GetParent(hfocused)

	If (cmd = "zeigen") {

		hfocused_last 	:= hfocused
		hparent_last  	:= hparent
		caretpos           	:= CaretPos(hfocused)
		ctrltxt            	:= SubStr(ControlGetText("", "ahk_id " hfocused), 1, caretpos - 1)

		If FileExist(Addendum.AdditionalData_Path "\GOÄTexte.json")
			bgr := JSON.Load(FileOpen(Addendum.AdditionalData_Path "\GOÄTexte.json", "r").Read(),, "UTF-8")

		RegExMatch(ctrltxt, "\d+$", ziffer)
		Privatabrechnung_BgrGui(bgr[ziffer], ziffer)
		;FileOpen(Addendum.AdditionalData_Path "\GOÄTexte.json", "w", "UTF-8").Write(JSON.Dump(bgr,, 2, "UTF-8"))

	}
	else if (cmd = "einfuegen") {

		ControlFocus,, % "ahk_id " hfocused_last
		if !ErrorLevel
			Send, % bgrWahl	;, % "ahk_id " hfocused_last
		else
			PraxTT("Addendum konnte den Text nicht einfügen", "1 2")

	}
	;SciTEOutput("so geht es nicht")

}
Privatabrechnung_BgrGui(content, ziffer)  {

	global
	local acx, acy

	acx        	:= A_CaretX
	acy       	:= A_CaretY

	Gui, Bgr: New, HWNDbgrHgui
	Gui, Bgr: Add, ListView	, % "xm ym w500 r10 vbgrLV ", % "Texte"
	For idx, bgrtxt in content
		LV_Add("", bgrtxt)
	Gui, Bgr: Add, Button	, % "xm y+5	vbgrUse    	gbgrButtonHandler", % "Begründungstext übernehmen"
	Gui, Bgr: Add, Button	, % "x+30 	vbgrCancel 	gbgrButtonHandler", % "Abbruch"
	Gui, Bgr: Show, AutoSize, % "Begründungstexte für GOÄ Ziffer [" ziffer "]"

return

bgrButtonHandler:

	if (A_GuiControl = "bgrUse") {
		Gui, bgr: ListView, bgLV
		If !(crow  := LV_GetNext(0))
			return
		LV_GetText(bgrWahl, crow)
		Privatabrechnung_Begruendungen("einfuegen", bgrWahl)
	}

bgrGuiClose:
bgrGuiEscape:
	Gui, bgr: Destroy
return
}


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

	PatientenID := Clip()
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

return
Clip(Text:="", Reselect:="") {                                                                        	;-- Clip() - Send and Retrieve Text Using the Clipboard

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
ScanBeenden:               	;{
 If (hwnd := WinExist("ScanSnap Manager - Bild erfassen und Datei speichern ahk_class #32770"))
	res := VerifiedClick("Button2", hwnd)
return ;}
;}

;}

;{11. Labels - TrayMenu

scriptLine:                                        	;{
	;ListLines, % (ListLinesFlag := !ListLinesFlag)
	ListLines, On
return ;}

SkriptReload(msg:="") {                                                                                   					;-- besseres Reload eines Skriptes

  ; es wird ein RestartScript gestartet, welches zuerst das laufende Skript komplett beendet und dann das Addendum-Skript per run Befehl erneut startet
  ; dies verhindert einige Probleme mit datumsabhängigen Handles bestimmter Prozesse im System, die Hook und Hotkey Routinen sind zuverlässiger dadurch geworden
  ; +Albisprogrammdatum wird auf den nächsten Tag gesetzt


	; Abbruch des Neustart falls Addendum gerade Befunde importiert oder ein OCR Vorgang läuft
		If (adm.Importing || adm.ImportRunning || Addendum.tessOCRRunning || Addendum.Thread["tessOCR"].ahkReady()) {

			; Protokolltext
				FormatTime	, Time      	, % A_Now    	, dd.MM.yyyy HH:mm:ss
				FormatTime	, TimeIdle	, % A_TimeIdle	, HH:mm:ss
				FileAppend	, % Time ", " A_ScriptName ", " A_ComputerName ", " msg ", terminated due to a running process, " TimeIdle "`n"
									, % Addendum.AddendumDir "\logs'n'data\OnExit-Protokoll.txt"

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

	; Status des Infofenster sichern
		If Addendum.AddendumGui
			admGui_SaveStatus()

	; Programmdatum auf aktuelles Datum setzen
		If (msg = "AutoRestart") && WinExist("ahk_class OptoAppClass")
			AlbisSetzeProgrammDatum()

		admScript	:= Addendum.Dir "\Module\Addendum\Addendum.ahk"
		scriptPID 	:= DllCall("GetCurrentProcessId")
		__          	:= q " " q                                             	; für bessere Lesbarkeit des Codes - ergibt > " " <

	; Autohotkey.exe /f "Z:\Addendum für AlbisOnWindows\Addendum_Reload.ahk" "Z:\Addendum für AlbisOnWindows\Module\Addendum\Addendum.ahk" "1" "0x34A567" "9876" " 1"
		cmdline := "Autohotkey.exe /f " q Addendum.Dir "\include\Addendum_Reload.ahk" __ admScript __ "1" __ A_ScriptHwnd __ scriptPID __ " 1" q
		Run, % cmdline
		ExitApp

return
}
ShowTextProtocol(FilePath) {                                                                                            	;-- zeigt Textdateien im eingestellten Texteditor an
	Run % FilePath
}
ZeigePatDB() {                                                                              		                   			;-- zeigt in einer eigenen Listview die Patienten.txt
	For PatID, Pat in oPat
		AutoListview("Pat.Nr.|Nachname|Vorname|Geburtsdatum|Geschlecht|Krankenkasse|letzte GVU", PatID "," Pat.Nn "," Pat.Vn "," Pat.Gt "," Pat.Gd "," Pat.KK "," Pat.letzteGVU, ",")
}
ZeigeFehlerProtokoll() {                                                                                  					;-- das Skriptfehlerprotokoll wird angezeigt

	filebasename    	:= Addendum.Dir "\logs'n'data\ErrorLogs\Fehlerprotokoll-"
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
AddendumObjektSpeichern() {                                                                         					;-- speichert und zeigt das Addendum Objekt an

	PraxTT("Speichere Addendum Objekt", "0 1")
	JSONData.Save(A_ScriptDir "\AddendumObjekt.json", Addendum, true,, 1, "UTF-8")
	Run, % A_ScriptDir "\AddendumObjekt.json"
	PraxTT("", "Off")

}

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
Menu_AlbisAutoPosition:                   	;{
	Addendum.AlbisLocationChange := !Addendum.AlbisLocationChange
	Menu, SubMenu3, ToggleCheck, % "Albis AutoSize"
	IniWrite, % (Addendum.AlbisLocationChange ? "ja":"nein")	, % Addendum.Ini, % compname, % "Albis_AutoGroesse"
return
;}
Menu_AddendumGui:                    	;{
	Addendum.AddendumGui := !Addendum.AddendumGui
	Menu, SubMenu3, ToggleCheck, % "Addendum Infofenster"
	IniWrite, % (Addendum.AddendumGui ? "ja":"nein")	       	, % Addendum.Ini, % compname, % "Infofenster_anzeigen"
	If Addendum.AddendumGui {
		Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If RegExMatch(Addendum.AktiveAnzeige, "i).*\|(Karteikarte)|(Laborblatt)")
			AddendumGui()
	} else If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {
		Controls("", "reset", "")
		admGui_SaveStatus()
		admGui_Destroy()
	}
return ;}
Menu_GVUAutomation:                    	;{
	Addendum.GVUAutomation := !Addendum.GVUAutomation
	Menu, SubMenu3, ToggleCheck, % "Albis GVU automatisieren"
	IniWrite, % (Addendum.GVUAutomation ? "ja":"nein")        	, % Addendum.Ini, % compname, % "GVU_automatisieren"
return
;}
Menu_PDFSignatureAutomation:      	;{

	; Menupunkt umschalten, Einstellung speichern
		Addendum.PDFSignieren := !Addendum.PDFSignieren
		Menu, SubMenu3, ToggleCheck, % "FoxitReader Signaturhilfe"
		IniWrite, % (Addendum.PDFSignieren ? "ja":"nein")          	, % Addendum.Ini, % compname, % "FoxitPdfSignature_automatisieren"

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
	IniWrite, % (Addendum.AutoCloseFoxitTab ? "ja":"nein") 	, % Addendum.Ini, % compname, % "FoxitTabSchliessenNachSignierung"
return ;}
Menu_LaborAbrufManuell:               	;{

	Addendum.Labor.AutoAbruf := !Addendum.Labor.AutoAbruf
	Menu, SubMenu3, ToggleCheck, % "Laborabruf automatisieren"
	IniWrite, % (Addendum.Labor.AutoAbruf ? "ja":"nein")   	, % Addendum.Ini, % compname, % "Laborabruf_automatisieren"

return ;}
Menu_LaborAbrufTimer:                  	;{

	Addendum.Labor.AbrufTimer := !Addendum.Labor.AbrufTimer
	Menu, SubMenu3, ToggleCheck, % "zeitgesteuerter Laborabruf"
	If Addendum.Labor-AbrufTimer {
		Addendum.Labor.AutoAbruf := true    ; muss dann auch an sein
		Menu, SubMenu3, ToggleCheck, % "Laborabruf automatisieren"
		IniWrite, % (Addendum.Labor.AutoAbruf ? "ja":"nein")	, % Addendum.Ini, % compname, % "Laborabruf_automatisieren"
	}
	IniWrite, % (Addendum.Labor.AbrufTimer ? "ja":"nein")   	, % Addendum.Ini, % compname, % "Laborabruf_Timer"
	IniWrite, % "ja"                                                            	, % Addendum.Ini, % compname, % "Laborabruf_automatisieren" ; dieser muss dann automatisch "An=Ja" sein!

return ;}
Menu_AddendumToolbar:             	;{
	Addendum.ToolbarThread := !Addendum.ToolbarThread
	Menu, SubMenu3, ToggleCheck, % "Addendum Toolbar"
	IniWrite, % (Addendum.ToolbarThread ? "ja":"nein"), % Addendum.Ini, % compname, % "Toolbar_anzeigen"
	admThreads()                                                                             ; startet oder stoppt die Ausführung von Threadcode
return ;}
Menu_Schnellrezept:                      	;{
	Addendum.Schnellrezept := !Addendum.Schnellrezept
	Menu, SubMenu3, ToggleCheck, % "Schnellrezept"
	IniWrite, % (Addendum.Schnellrezept ? "ja":"nein")  	, % Addendum.Ini, % "Addendum", % "Schnellrezept"
return ;}
Menu_mDicom:                              	;{
	Addendum.mDicom := !Addendum.mDicom
	Menu, SubMenu3, ToggleCheck, % "MicroDicom Export"
	IniWrite, % (Addendum.mDicom ? "ja":"nein")         	, % Addendum.Ini, % compname, % "MicroDicomExport_automatisieren"
return ;}
Menu_PopUpMenu:                          	;{
	Addendum.PopUpMenu := !Addendum.PopUpMenu
	Menu, SubMenu3, ToggleCheck, % "PopUpMenu"
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
Menu_AutoOCR:                            	;{
	Addendum.OCR.AutoOCR := !Addendum.OCR.AutoOCR
	Menu, SubMenu3, ToggleCheck, % "AutoOCR"
	IniWrite, % (Addendum.OCR.AutoOCR ? "ja":"nein"), % Addendum.Ini, % "OCR", % "AutoOCR"
;}
Menu_WatchFolder:                        	;{
	Addendum.OCR.WatchFolder := !Addendum.OCR.WatchFolder
	Menu, SubMenu3, ToggleCheck, % "Befundordner überwachen"
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
;}
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
	if !vScrInputRectState	{
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
Splash(string, time=4, x="", y="", w="", ParentHWnd=0) {                         					;-- Splash Gui - to show text for some seconds

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
	Lc := Lc > 4 ? 4 : Lc
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
GDISplash(text, time=2, SplashPos="", SplashOpt="") {                                              	;-- Hinweisfenster

		static hGdiSp, GdiSp

		If !IsObject(SplashOpt) {
			hFont    	:= CreateFont("y5 Centre cff000000 q5 s16, " Addendum.Standard.Font)
			bgColor	:= "0x778fA2ff"
		}
		else {
			hFont    	:= CreateFont(SplashOpt.FontStyle ", " SplashOpt.Font)
			bgColor 	:= SplashOpt.bgColor
		}

		If !IsObject(SplashPos) {
			cx := A_CaretX , cy := A_CaretY - 35
			If (!cx or !cy) {
				MouseGetPos, cx, cy
				cy += 20
			}
		}
		else {
			cx := SplashPos.X
			cy := SplashPos.Y
		}

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
		;DllCall("AnimateWindow", "UInt",hGdiSp, "Int",300, "UInt",0x90000)

		SetTimer, GDISplashOff, % -1 * (time*1000)

return hGdiSp

GDISplashOff:
		;DllCall("AnimateWindow", "UInt",hGdiSp, "Int",300, "UInt",0x90000)
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

	Splash(string " " cmd, 1)

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

	SendMode, Event
	SetKeyDelay, 10, 10
	Send, % ""

	Loop, % alternating.MaxIndex()
		If (alternating[A_Index] = chars) 	{
			SendInput, % StrSplit(chars, "|").2
			alternating.RemoveAt(A_Index)
			return
		}

	SendInput, % StrSplit(chars, "|").1
	alternating.Push(chars)

}

SciTEFoldAll() {                                                                                                            	;-- gesamten Code zusammenfalten

	ControlGet, hScintilla1, hwnd,, Scintilla1, % "ahk_class SciTEWindow"
	Sleep 300
	SendMessage, 2662,,,, % "ahk_id " hScintilla1 ; SCI_FOLDALL
	;TrayTip("SciTEFoldAll", "hScintilla1: " hScintilla1 "`nErrorLevel: " ErrorLevel, 2 )

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

	AddendumDir := Addendum.Dir
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
ListviewClipboard(WinTitle:="") {                                                                                	;-- kopiert den Inhalt einer Listview,Listbox,DDL ins Clipboard

	; letzte Änderung: 22.02.2021

	; Parser Einstellungen für Albis Listviewsteuerelemente
		static CopyOpt :=
		( Join
			{
			"Dauerdiagnosen":    	{
				"SysListView321": {
					   "num"         	: 1,
					   "numFormat"	: ".`t",
					   "Rxrpl"           	: "^\w+\s+\w+\s+[\d.]+\s+(?=\w)",
					   "rxrplWith"  	: ""
				  },
				"SysListView322": {
						"num"         	: 1,
						"numFormat"	: ".`t",
						"Rxrpl"       	: {
							"(\d+)mg"                 	: "$1 mg",
							"\d+\s+St\**"            	: " ",
							"\sN\d\s"                    	: " ",
							"^\s*(.*)\s+\d+\s*$"	: "$1",
							"Pharma"                    	: "",
							"\s{2,}"                       	: " "
					   },
				   "rxrplWith": "$1"
				  }
			 },
			"Dauermedikamente":	{
				"SysListView321": {
					   "num"         	: 1,
					   "numFormat"	: ".`t",
					   "Rxrpl"        	: {
							"(\d+)mg"                  	: "$1 mg",
							"\d+\s+St\**"            	: " ",
							"\s+[\d.]+\s+\d+\s*$"	: "",
							"\sN\d\s"                   	: " ",
							"Pharma"                   	: "",
							"\s{2,}"                     	: " "
					   },
				   "rxrplWith"   	: "$1"
				  }
			 },
			 "Zusatzdaten":     	{
				"SysListView322": {
					"num"         	: 0,
					"numFormat"	: ".`t",
					"Rxrpl"        	: {
						"([\w\s]+)\s\.\.\.+\s+([\>\,\.\w\s\/]+)" : "`t$1  $2",
						"([\w\(\),.+\s]+)\s(\.+|\s)\s*([ISR])" : "`t$3`t$1",
						"([\w\(\).\s]+)\s\.+\s+([#,.\w\s\/]+)" : "`t$2`t$1"
					}
				  },
				"rxrplWith": "$1"
				}
			 }
		)

	; Dauermedikamente
		RemoveThatShit := {	"Aaa"                     	: ""
									, 	"#\-*1\s*A"              	: " "
									, 	"Ab[Zz]"                 	: ""
									, 	"Acino"                  	: ""
									, 	"Actavis"                  	: ""
									, 	"AL"                       	: ""
									, 	"Ari"                      	: ""
									, 	"Arzneim"              	: ""
									, 	"Aristo"                  	: ""
									, 	"Allomedic"            	: ""
									,	"Augentropfen"      	: ""
									, 	"Aurobindo"           	: ""
									, 	"Arzne"                  	: ""
									, 	"AWD"                   	: ""
									,	"Axico"                  	: ""
									,	"Axicorp\s*P"          	: ""
									, 	"BAYER|Bayer"       	: ""
									, 	"Becat"                  	: ""
									, 	"beta"                    	: ""
									,	"Br\W"                   	: ""
									,	"\s*\-\s+Ct(\d+)"     	: " $1"
									,	"\-*\s*CT"              	: ""
									,	"Deutsch"              	: ""
									,	"Dosiergel"            	: ""
									,	"Dosieraeros"        	: ""
									,	"ED"                      	: ""
									,	"Emra.Med"           	: ""
									,	"Eurimpharm"        	: ""
									,	"Fair.Med"             	: ""
									,	"Fertigspritze"        	: ""
									,	"Filmtabl*e*t*t*e*n*"	: ""
									,	"GALEN"                	: ""
									,	"Glenmark"           	: ""
									,	"GmbH*"               	: ""
									,	"Hartka"                	: ""
									,	"Henning"              	: ""
									,	"Hennig"                	: ""
									,	"Heum*a*n*n*"         	: ""
									,	"HEXAL"                 	: ""
									,	"Hkp"                    	: ""
									,	"Huebe"                	: ""
									,	"Inject"                     	: ""
									,	"Inte"                     	: ""
									,	"Isis"                      	: ""
									,	"kohlpharma*"       	: ""
									,	"KwikPen"              	: ""
									,	"Lomapharm"         	: ""
									,	"Lichten(\d+)"           	: " $1"
									,	"Mili"                     	: ""
									,	"Milinda"               	: ""
									,	"Msr"                     	: ""
									,	"Myl"                     	: ""
									,	"Nachfuell"            	: ""
									,	"Net"                     	: ""
									,	"Novolet"               	: ""
									,	"Orifarm"              	: ""
									,	"Pe\s"                    	: ""
									,	"Pen\s"                  	: ""
									,	"Penf*i*l*l*"               	: ""
									,	"Phar*m*a*"          	: ""
									,	"Protect(\d+)"        	: " $1"
									,	"[Rr]atiop*h*a*r*m*"  	: ""
									,	"Retard"                 	: ""
									,	"Retardtable*t*t*e*n*"	: ""
									,	"SANDOZ"               	: ""
									,	"Sta(\d)"                 	: " $1"
									,	"STADA|Stada"      	: ""
									,	"\ssto"                   	: ""
									,	"Tabl\W"                  	: ""
									,	"Tabletten"             	: ""
									,	"T[Aa][Hh]"            	: ""
									,	"TEI"                      	: ""
									,	"TEVA"                   	: ""
									,	"Tro"                     	: ""
									,	"Vital"                    	: ""
									,	"Weichkaps*e*l*n*"	: ""
									,	"Winthrop"	            	: ""
									,	"Zentiva"               	: ""
									,	"4Wochen"            	: "" }

	; aktiver Fenstertitel
		WinTitle := AlbisGetActiveWinTitle()

	; beim Wartezimmer muss das Listviewhandle anders ermittelt werden
		If (WinTitle = "Wartezimmer") {
			hWZ := AlbisMDIWartezimmerID()
			ControlGet, hChild    	, HWND,,                      	, % "ahk_id " hWZ
			ControlGet, outcontrol	, HWND,, SysListView321	, % "ahk_id " hChild
		}	else {
			MouseGetPos,,, outwin, hControl	, 2
			MouseGetPos,,, outwin, classnn 	, 1
			WinGetTitle, Wintitle, % "ahk_id " outwin
		}

	; Steuerelement ist eines der unteren dann wird der Inhalt ausgelesen und eventuell geparsed (CopyOpt)
		If RegExMatch(classnn, "(SysListView32|Listbox|ComboBox|DDL)") {

			ControlGet, content, List,,, % "ahk_id" hControl
			RegExReplace(content "`n`r", "[\n\r]", "", cLinesCount)

			nopts := false
			For OptTitle, OptControls in CopyOpt {
				If InStr(WinTitle, OptTitle) {

					rxrpl      	:= OptControls[classnn].Rxrpl
					rxrplWith 	:= OptControls[classnn].rxrplWith
					numbers 	:= OptControls[classnn].num ? true : false
					numF		:= OptControls[classnn].numFormat

					If !rxrpl {
						nopts := true
						break
					}

				; ---
					cLines := StrSplit(RegExReplace(content, "[\n\r]+", "`n"), "`n", "`r")
					newcontent := ""
					For cIdx, cLine in cLines {

						SubNR := (StrLen(cIdx) = 1 ? -3 : -2)
						leading := StrLen(cIdx) = StrLen(cLines.MaxIndex()) ? "" : SubStr("     ", 1, SubNr)

						If !IsObject(rxrpl)
							cLine := RegExReplace(cLine, rxrpl, rxrplWith)
						else
							For rxrplStr, rxrplWithStr in rxrpl {
								If (rxrplStr <> "Pharma")  ; entfernt Herstellernamen
									cLine := RegExReplace(cLine, rxrplStr, rxrplWithStr)
								else
									For rxrplStr2, rxrplWithStr2 in RemoveThatShit {
										rxrplStr2 := RegExReplace(rxrplStr2, "^#", "", noWhiteSpace)
										cLine := RegExReplace(cLine, (noWhiteSpace ? "" : "\s") . rxrplStr2, rxrplWithStr2)
									}
							}

						newcontent .= (numbers ? leading . cIdx . numF : "") cLine "`n`r"
					}

				}

				If nopts
					break
			}

			If nopts {
				ClipBoard := content ? RegExReplace(content, "[\n\r]+", "`n") : newcontent ? RegExReplace(newcontent, "[\n\r]+", "`n") : ""
				ClipWait, 1
				PraxTT("Inhalt des Steuerelementes kopiert.`n[" cLinesCount " Zeilen]`n[" StrLen(content) " Zeichen]", "2 1")
			} else {
				ClipBoard := newcontent ? RegExReplace(newcontent, "[\n\r]+", "`n") : content ? RegExReplace(content, "[\n\r]+", "`n") : ""
				ClipWait, 1
				PraxTT("Inhalt des Steuerelementes kopiert.`n[" cLinesCount " Zeilen]`n[" StrLen(newcontent) " Zeichen]`nInhalt wurde nach Vorgabe formatiert!", "6 1")
			}

		}

}
GetPatNameFromExplorer() {                                                                                     	;-- ermittelt in einer pdf Dateibezeichnung den Patientennamen und öffnet die Patientenakte
	; Voraussetzung: der Patient ist in der Addendum Patientendatenbank gelistet, ansonsten passiert nichts
		SplitPath, % Explorer_GetSelected(WinExist("A")),,,, fname
		FuzzyKarteikarte(fname)
}


; Labor
LaborAbrufCheck() {                                                                                                   	;-- prüft Bedingungen vor der Ausführung des Timers

	If !Addendum.Labor.AbrufTimer
		return false

	LabCallClients := Addendum.Labor.ExecuteOn
	If !InStr(LabCallClients, compname)
		return false

return true
}
LaborAbrufTimerStarten() {                                                                                         	;-- berechnet die Millisekunden bis zum nächsten eingestellten Laborabruf

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

	; nächster Abruf ist am nächsten Tag
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
		Addendum.Laborabruf.Clock	 	:= Uhrzeit
		Addendum.Laborabruf.Call		:= TimeLabCall
		IniWrite, % AbrufTag ", " Uhrzeit, % Addendum.Ini, % "LaborAbruf", % "naechster_Abruf"
		outtext =
		(LTrim
			nächster Laborabruf
			Uhrzeit: %Uhrzeit%
			Start in: %Restzeit%
		)
		PraxTT(outtext, "3 1")

}
LaborAbruf() {                                                                                                             	;-- startet externes Skript für den Laborabruf

	; Überprüfung des Aufrufes
		If !FileExist(Addendum.Labor.ExecuteScript) {
			FileAppend, % datestamp() "  " A_ThisFunc "():  ianovaWebClient.exe nicht vorhanden!: " Addendum.Labor.ExecuteScript "`n", % Addendum.DBPath "\Labordaten\LaborabrufLog.txt"
			return
		}

	; Laborabrufskript ausführen
		Run, % q Addendum.Labor.ExecuteScript q

	; LaborabrufTimer erneut starten
		LaborAbrufTimerStarten()
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


; ----------------------- Hook-Initialisierung
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
		HookProcAdr         	:= RegisterCallback("WinEventProc"    	, "F")
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
		Gui, AMsg: New	, +HWNDhMsgGui +ToolWindow
		Gui, AMsg: Add	, Edit, xm ym w100 h100 HWNDhEditMsgGui
		Gui, AMsg: Show	, AutoSize NA Hide, % "Addendum Message Gui"
		Addendum.MsgGui.hMsgGui     	:= hMsgGui
		Addendum.MsgGui.hEditMsgGui	:= hEditMsgGui

	; Shellhook wird gestartet
		DllCall("RegisterShellHookWindow", UInt, hMsgGui)
		MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(MsgNum, "ShellHookProc")

	; startet weitere Hooks die nur auf Ereignisse vom Albisprozess reagieren
		AlbisStartHooks()

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
AlbisPopUpMenuHook(AlbisPID) {                                                                                  	; Hook zum Abfangen eines Rechtsklick Menu

	HookProcAdr   	:= RegisterCallback("AlbisPopUpMenuProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x0006, 0x0007, 0, HookProcAdr, AlbisPID, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisFocusHook(AlbisPID) {                                                                                            	; Hook ausschließlich für EVENT_OBJECT_FOCUS Nachrichten wird gestartet

	; benötigt um Änderung des Eingabefocus innerhalb von Albis zu erkennen
	; z.B. für eine leider noch funktionierende Autovervollständigungsfunktion

	HookProcAdr   	:= RegisterCallback("AlbisFocusEventProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x8005, 0x8005, 0, HookProcAdr, AlbisPID, 0, 0x3)
	hWinEventHook	:= SetWinEventHook( 0x800C, 0x800C, 0, HookProcAdr, AlbisPID, 0, 0x3)

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


; ----------------------- Verarbeitung abgefangener Systemmeldungen (EventProcs)
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

		;If (GetDec(hwnd) = 0)
		If !hwnd
			return 0

		EHookHwnd := Format("0x{:x}", hwnd)
		If (Addendum.LastHookHwnd = EHookHwnd)
			return 0

		cClass := WinGetClass(hwnd)

	; Fensterklassen Filter
		For fnr, filterclass in class_filter
			If (cClass = filterclass) {

				known := false
				For snr, item in EHStack
					If (item.Hwnd = EHookHwnd) {
						known := true
						break
					}
				If !known {
					WinGetTitle, HTitle, % "ahk_id " EHookHwnd
					WinGetText, HText, % "ahk_id " EHookHwnd
					If (StrLen(HTitle . HText) > 0) {
						EHStack.InsertAt(1, {"Hwnd"	: EHookHwnd
													, "Event"	: Format("0x{:x}", Event)
													, "Title"   	: HTitle
													, "Text"   	: StrReplace(HText, "`r`n", " ")
													, "Class"	: cClass})
						If !EHWHStatus
							SetTimer, EventHook_WinHandler, -1
					}
				}

			}

	; Callback für
		If (StrLen(Addendum.FuncCallback) > 0) {
			For idx, viewerclass in PDFViewer
				If InStr(cClass, viewerclass) {
					If InStr(Addendum.FuncCallback, "|") {
						callbackParam := StrSplit(Addendum.FuncCallback, "|")
						If IsFunc(callbackParam[1])
							func_Call 	:= Func(callbackParam[1]).Bind(callbackParam[2], cClass, SHookHwnd)
					} else {
						fnName := Addendum.FuncCallback
						If IsFunc(fnName)
							func_Call 	:= Func(fnName).Bind(cClass, SHookHwnd)
					}
					SetTimer, % func_Call, -200
					break
					}
		}

	; Order & Entry Programm läßt sich nicht filtern
		WinGet, EHproc, ProcessName, % "ahk_id " EHookHwnd
		If Instr(EHProc, "infoBoxWebClient") {
			;~ WinGetTitle, Title, % "ahk_id " EHookHwnd
			;~ WinGetText, Text, % "ahk_id " EHookHwnd
			EHStack.InsertAt(1,  {"Hwnd":EHookHwnd, "Event":Format("0x{:x}", Event), "Title":HTitle, "Text":StrReplace(HText, "`r`n", " "), "Class":cClass})
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

	;SciTEOutput(GetHex(event) ", " GetHex(hwnd) )

	If (GetDec(hwnd) = 0)
		return 0

	AlbisHookTitle := AlbisGetActiveWinTitle()
	;cFocus	:= GetClassName((hwnd := Format("0x{:x}", hwnd)))
	;~ If (AlbisTitleO = AlbisHookTitle) && RegExMatch(AlbisHookTitle, "^\s*\d+\s*\/") ;&& !RegExMatch)|(Laborbuch)|(Wartezimmer)|(EBM)(ToDo)")
		;~ return 0

	;AlbisTitleO := AlbisHookTitle
	fn_PatChange := Func("CurrPatientChange").Bind(AlbisHookTitle)
	SetTimer, % fn_PatChange, -0

return 0
}

AlbisLocationChangeEventProc(hHook, event, hwnd) {                                                    	; behandelt EVENT_OBJECT_LOCATIONCHANGE Nachrichten von Albis

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

		If (hwnd = 0) || (GetHex(hwnd) <> Addendum.AlbisWinID)
			return 0

		Addendum.MonSize  	:= AlbisMon.W > 1920 ? 2 : 1
		Addendum.Resolution  	:= AlbisMon.W > 1920 ? "4k" : "2k"

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
			Y := StrLen(AlbisY) > 0 ? (AlbisY -    6) 	: APos.Y

			If IsRemoteSession() {
				factor 	:= A_ScreenDPI / 96
				W	:= Floor(AlbisW * factor)
				H	:= Floor(AlbisH * factor)
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
			WinActivate, % "ahk_id " AlbisWinID()
		}

return 0
}

ActiveFocusEventProc(hHook, event, hwnd) {                                                                    	; behandelt EVENT_OJECT_FOCUS Nachrichten alle Programme

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

		global AlbisID_old, hadm ; hadm = hwnd des Infofenster

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
				If (AlbisWinID() <> AlbisID_old) {
					AlbisID_old := AlbisWinID()
					AlbisStartHooks()                                                                            ; startet Hooks
					PraxTT("Ein neuer Albisprozeß wurde erkannt.`nDie Fensterhooks wurden gesetzt!", "2 0")
				}

			; InfoFenster prüfen
				If Addendum.AddendumGui
					If (AlbisWinID() = SHookHwnd) && RegExMatch(AlbisGetActiveWindowType(), "i)Karteikarte|Laborblatt|Biometriedaten") {
						If !WinExist("AddendumGui ahk_class AutohotkeyGui")
							AddendumGui()
						else
							RedrawWindow(hadm)
					}

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
		else if (lParam = 2) && !WinExist("ahk_class OptoAppClass")	{	                                               	; ALBIS wurde beendet

			If Addendum.AddendumGui
				admGui_Destroy()

			AlbisStopHooks()
			PraxTT("Albis wurde beendet. Hooks wurden entfernt.", "2 0")
			AlbisID_old := 0

		}
		;~ else if (lParam = 2) && WinExist("Addendum (Datei umbenennen) ahk_class AutohotkeyGUI")	{	; SUMATRA Pdf Reader bei geöffnetem Dialog beendet

		;~ }

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

	If (EHStack.Count() > 0 && !EHWHStatus) {
		SetTimer, WaitEH_WinHandler, Off
		gosub EventHook_WinHandler
	} else if (EHStack.Count() = 0)
		SetTimer, WaitEH_WinHandler, Off

return ;}


; ----------------------- Automatisierungsroutinen / Fensterhandler
EventHook_WinHandler:                                                                                                	;{ Eventhookhandler - Popupblocker/Fensterhandler - für diverse Fenster verschiedener Programme

	StackEntry:
	EHWHStatus :=true, thisWin := EHStack.Pop()
	EHWT := thisWin.title, EHWText := thisWin.Text, EHWClass := thisWin.Class
	Addendum.LastHookHwnd := hHookedWin := thisWin.Hwnd

	If (StrLen(EHWT . EHWText) = 0) || InStr(EHWText, "SkinLoader")
		If (EHStack.Count() = 0) {
			EHWHStatus := Addendum.LastHookHwnd := false
			return
		} else
			goto StackEntry

	WinGet, EHproc1, ProcessName, % "ahk_id " hHookedWin
	If (EHWT && !EHproc1) {
		WinGet, EHproc1, ProcessName, % "ahk_id " GetParent(hHookedWin)
		If !EHproc1 && !EHStack.Count() {
			EHWHStatus := Addendum.LastHookHwnd := false
			return
		} else if !EHproc1 && (EHStack.Count() > 0)
			goto StackEntry
	}

	If        InStr(EHproc1, "albis")                                                                      	{        	; ALBIS

			If   	  InStr(EHWText	, "ist keine Ausnahmeindikation")                                                              	 {  	; Fenster wird geschlossen
				VerifiedClick("Button2", hHookedWin)
				GNRChanged 	:= true
			}
			else If InStr(EHWT  	, "Daten von") && GNRChanged                                                              	 {  	; schließt nach Änderung der GNR das Fenster "Daten von " für schnelleres Handling
				while (A_Index < 30) || WinExist("ALBIS ahk_class #32770", "ist keine Ausnahmeindikation")
					Sleep 10
				VerifiedClick("Button30", "Daten von ahk_class #32770")
				GNRChanged 	:= false
			}
			else If InStr(EHWText	, "ALBIS wartet auf Rückgabedatei")                                                          	 {  	; Laborhinweis schliessen
				BlockInput, On
				VerifiedClick("Button1", hHookedWin)
				AlbisActivate(1)
				SendInput, {Tab}
				WinActivate, % "ahk_class #32770", % "ALBIS wartet auf Rückgabedatei"
				BlockInput, Off
			}
			else If InStr(EHWText	, "Patient hat in diesem")                                                                           	 {  	; in Abhängigkeit des Client wird das Fenster sofort oder verzögert geschlossen
				hNoChippie := hHookedWin
				If (Addendum.noChippie > 0)
					Sleep % Addendum.noChippie*1000
				VerifiedClick("Button1", hNoChippie)
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
			else if InStr(EHWT  	, "CGM HEILMITTELKATALOG")                                                              	 { 	; Fensterposition wird innerhalb des Albisfenster zentriert
				AlbisHeilmittelKatologPosition()
			}
			else if InStr(EHWT  	, "Dauerdiagnosen von")                                                                        	 { 	; Autopositionierung des Dauerdiagnosenfensters
				res:= AlbisResizeDauerdiagnosen()
			}
			else if InStr(EHWT  	, "ICD-10 Thesaurus")                                                                            	 { 	; vergrößert den Diagnosenauswahlbereich
				UpSizeControl("ICD-10 Thesaurus", "#32770", "Listbox1", 200, 150, AlbisWinID())
			}
			else if InStr(EHWT  	, "Labor - Anzeigegruppen")                                                                      	 { 	; die Anzeige der Listen
				res:= AlbisResizeLaborAnzeigegruppen()
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
				; bei Adressen ohne Straßenangabe, kann das Fenster 'weitere Informationen' nicht ohne weiteres geschlossen werden
				MsgBox, 0x1024, Addendum für Albis on Windows, Rechnungsempfänger löschen?
				IfMsgBox, Yes
				{
					Loop 8
						VerifiedSetText("Edit" A_Index, "", "ahk_class #32770", 200, "Adresse des Rechnungs")
				}
				VerifiedClick("OK", "ahk_class #32770", "Adresse des Rechnungs")
			}
			else if Addendum.Schnellrezept && InStr(EHWT, "Muster 16")                                                     	 {  	; Kassenrezept - Schnellrezept + Auto autIdem Kreuz
					AlbisRezeptHelferGui(Addendum.AdditionalData_Path "\RezepthelferDB.json")
			}
			;-------------------- Laborabruf Automation --------------------------------------------------------
			else If Addendum.Labor.AutoAbruf && InStr(EHWText	, "Anforderungen ins Laborblatt übertragen"){ 	; es wird automatisch mit "Ja" bestätigt
				If !InStr(Addendum.LaborAbruf.Status, "Failure")
					VerifiedClick("Button1", hHookedWin)
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWText	, "alle markierten Laboranforderungen")    	 {
				If !InStr(Addendum.LaborAbruf.Status, "Failure")
					VerifiedClick("Button1", hHookedWin)                          ; mit ja bestätigen
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWText	, "Keine Datei(en) im Pfad") 	                   	 { 	; Laborabrufvorgang wird beendet
				ResetLaborAbrufStatus()
				VerifiedClick("Button1", hHookedWin)
			}
			else if Addendum.Labor.AutoAbruf && Instr(EHWT  	, "Labor auswählen")                                 	 {    	; wählt das eingetragene Labor und bestätigt mit 'Ok'
				AlbisLaborAuswählen(Addendum.Labor.LaborName)
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWT  	, "Labordaten")                                        	 {
				AlbisLaborDaten()
			}
			else If Addendum.Labor.AutoAbruf && InStr(EHWT  	, "GNR der Anford")             	                 	 {		;
				If !Addendum.LaborAbruf.Status
					AlbisLaborGNRHandler(1, hHookedWin)       	         	; Listbox1 enthält Rechnungsdaten
			}

	}
	else if InStr(EHproc1, "mDicom")             	&& (Addendum.mDicom)             	{        	; mDicom Viewer

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

		IniWrite, % A_DD "." A_MM "." A_YYYY " (" A_Hour ":" A_Min ":" A_Sec ")", % Addendum.Ini, Labor, letzter_Laborabruf
		If !Addendum.Laborabruf.Status {
			Addendum.Laborabruf.Status := "Init"
			func_RLStat := Addendum.LaborAbruf.Reset := Func("ResetLaborAbrufStatus")
			SetTimer, % func_RLStat, -600000   ; 10 min
			PraxTT("Das Laborabruf Fenster wurde detektiert!`n#3Addendum übernimmt den weiteren Vorgang!", "10 1")
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
		If Addendum.PDF.RecentlySigned {
			if InStr(EHWT, "Speichern unter") && RegExMatch(EHWText, "i)Speichern|Save", gotText) {
				Addendum.PDF.RecentlySigned := false
				thisFoxitID := GetParent(hHookedWin)                          	; Handle des zugehörigen FoxitReader Fenster
				If VerifiedClick("Speichern", hHookedWin,,,true)                 	; Speichern Button drücken
					Addendum.SaveAsInvoked := true
				WinWait, % "Speichern unter bestätigen" ,, 4
				If !ErrorLevel
					VerifiedClick("Ja", "Speichern unter bestätigen",,, true)    	; mit 'Ja' bestätigen
				}
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

				ClipBoard := Addendum.Labor.Kennwort
				SendEvent, % "{Text}" Addendum.Labor.Kennwort
				Sleep 500
				SendInput, {Enter}

				PraxTT("CGM Channel Login wurde durchgeführt.", "6 2")

		}
	}
	else if InStr(EHproc1, "dbview")                                                                   	{       	; dbviewer.exe
		If !VerifiedClick("Button1", hHookedWin)
			VerifiedClick("Button1", "Vielen Dank ahk_class #32770")
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
				wPos.Y := 0
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
		If (w.X != wPos.X) || (w.Y != wPos.Y) || (w.W != wPos.W) || (w.H != wPos.H) {
			;SciTEOutput(Addendum.MonSize "(" Addendum.Resolution ") :" wPos.X ", " wPos.Y ", " wPos.W ", " wPos.H)
			SetWindowPos(hWin, wPos.X, wPos.Y, wPos.W, wPos.H)
		}
}

CurrPatientChange(AlbisTitle) {                                                                                      	; behandelt Änderung des Albisfenstertitels

		global hadm, admHJournal
		static 	AlbisTitleO

		Critical, 50

		If (AlbisTitleO = AlbisTitle)
			return
		 AlbisTitleO	:= AlbisTitle
		 AktiveAnzeige := AlbisGetActiveWindowType()

	; Zurücksetzen des Status des Laborabrufes ist bei Wechsel auf eine andere Ansicht möglich
		Addendum.LaborAbruf.Status := ""

	; Patienten registrieren bei der Addendum Patientendatenbank
		If 	InStr(AktiveAnzeige, "Patientenakte") {
			If PatDb(AlbisTitle2Data(AlbisTitle), "exist")
				; ist das Infofenster verfügbar, wird das Protokoll dort aufgefrischt
					If Addendum.AddendumGui && WinExist("AddendumGui ahk_class AutoHotkeyGUI")
						admGui_TProtokoll(Addendum.TProtDate, 0, compname)
		}

	; AddendumGui verbergen wenn Bedingungen nicht erfüllt sind
		If 	!Addendum.AddendumGui || !WinExist("ahk_class OptoAppClass") {
			admGui_Destroy()
		}
		else if Addendum.AddendumGui {

			If RegExMatch(AktiveAnzeige, "i)Karteikarte|Laborblatt|Biometriedaten")
				AddendumGui()
			else If (RegExMatch(AktiveAnzeige, "i)Abrechnung|Rechnungsliste") || RegExMatch(AlbisTitle, "i)Laborbuch|Terminkalender"))
				admGui_Destroy()
			else
				admGui_Destroy()

		}

	; Hinweisfenster entfernen
		If InStr(AlbisTitle, "Prüfung EBM/KRW") && !InStr(AlbisTitle, "Abrechnung vorbereiten")	{
			AlbisActivate(1)
			Sleep 6000
			SendInput, {Esc}
		}

	; Patientenakte geöffnet - dann Überprüfung des Patientennamen und Erstellen der AddendumGui
		If 	InStr(AktiveAnzeige, "aPEK")	                 	{
			MsgBox, 4, Addendum für Albis on Windows, Hinweisdialog zur Prüfung EBM/KRW weiter ansehen?, 12
			IfMsgBox, Yes
				return
			AlbisMDIChildWindowClose("Prüfung EBM/KRW")
		}
		else if 	InStr(AktiveAnzeige, "Laborbuch") && (Addendum.Laborabruf_Voll)
			Albismenu(34157, "", 6, 1)			;34157 - alle übertragen

Return
}



;}

;{16. Interskriptkommunikation / OnMessage
MessageWorker(InComing) {                                                                                     	;-- verarbeitet die eingegangen Nachrichten

		global AutoLogin
		static func_AutoOCR	:= func("admGui_OCRAllFiles")

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}

		;SciTEOutput("Addendum Incoming: `nmsg: " recv.txtmsg "`nfrom: " recv.fromID)
	; Kommunikation mit dem Abrechnungshelfer Skript
		If 			RegExMatch(recv.txtmsg, "^AutoLogin\s+Off")                    	{
			AutoLogin := false
			result := Send_WM_COPYDATA("AutoLogin disabled", recv.opt)
		}
		else if 	RegExMatch(recv.txtmsg, "^AutoLogin\s+On")	                	{
			AutoLogin := true
			result := Send_WM_COPYDATA("AutoLogin enabled", recv.opt)
		}
	; Tesseract OCR Thread Kommunikation
		else if 	InStr(recv.txtmsg, "OCR_processed")                                      	{
			admGui_CheckJournal(recv.opt, recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "OCR_ready")                                           	{

			; Anzeige auffrischen
				admGui_Journal()
				OCR := Addendum.OCR
				Addendum.tessOCRRunning := false
				If !OCR.RestartAutoOCR
					admGui_OCRButton("-OCR ausführen")

			; den Thread beenden, Speicher freigeben
				If IsObject(Addendum.Thread["tessOCR"]) {
					PraxTT("tesseract OCR-Thread beendet", "2 1")
					ahkthread_free(Addendum.Thread["tessOCR"])
					Addendum.Thread["tessOCR"] := ""
				}

			; ist AutoOCR eingeschaltet und sind weitere Dateien vorhanden wird
			; der Texterkennungsvorgang nochmals gestartet
				If OCR.RestartAutoOCR && OCR.AutoOCR && (OCR.Client = compname)
					SetTimer, % func_AutoOCR, % "-" (OCR.AutoOCRDelay*1000)

		}
	; Laborabruf_iBWC.ahk
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "Status"){
			Send_WM_COPYDATA("AutoLaborAbruf Status|" Addendum.Labor.AutoAbruf  "|" GetAddendumID(), recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "aus")	{
			If Addendum.Labor.AutoAbruf {
				Addendum.Labor.AutoAbruf_Last := true
				gosub Menu_LaborAbrufManuell                    ; Traymenueinstellung ändern
			}
			Send_WM_COPYDATA("AutoLaborAbruf angehalten||" GetAddendumID(), recv.fromID)
		}
		else if 	InStr(recv.txtmsg, "AutoLaborAbruf") && InStr(recv.opt, "an") 	{
			If Addendum.Labor.AutoAbruf_Last {
				Addendum.Labor.AutoAbruf_Last := false
				gosub Menu_LaborAbrufManuell                    ; Traymenueinstellung ändern
			}
			Send_WM_COPYDATA("AutoLaborAbruf fortgesetzt||" GetAddendumID(), recv.fromID)
		}
	; Addendum Statusabfragen
		else if 	InStr(recv.txtmsg, "Status")                                                    	{
			Send_WM_COPYDATA("Status|okay|" GetAddendumID(), recv.fromID)
		}

return
}

WM_DISPLAYCHANGE(wParam, lParam) {                                                                 	;-- einsetzbar für Script-Editoren - AutoZoom bei Wechsel der Bildschirmauflösung

	; zoomt bei einer Auflösung > 1920x1080 Scite4AHK
	; verschiebt WinSpy in die Mitte des Hauptmonitors
	; letzte Änderung: 06.02.2021

	static lastScreenSize, SCI_GETZOOM := 2374, SCI_SETZOOM := 2373
	static ZoomLevel := {"4k":{"Sci1":"2", "Sci2":"-2"},  "2k":{"Sci1":"0", "Sci2":"-3"}}

	If (hScite := WinExist("ahk_class SciTEWindow")) {

		SPos := GetWindowSpot(hScite)
		SysGet SciteMon_, Monitor, % GetMonitorAt(SPos.X, SPos.Y)
		monDim := SciteMon_Right "x" SciteMon_Bottom

		ZoomLevel_Sci1 := SciteMon_Bottom > 1080 ? ZoomLevel.4k.Sci1 : ZoomLevel.2k.Sci1
		ZoomLevel_Sci2 := SciteMon_Bottom > 1080 ? ZoomLevel.4k.Sci2 : ZoomLevel.2k.Sci2

		If (lastScreenSize <> monDim) {

			lastScreenSize := monDim

			ControlGet	, hScintilla1	, hwnd,, Scintilla1	, % "ahk_id " hScite
			ControlGet	, hScintilla2	, hwnd,, Scintilla2	, % "ahk_id " hScite

			SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla1
			ZoomScintilla1O := ErrorLevel
			SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla2
			ZoomScintilla2O := ErrorLevel

			If (ZoomScintilla1O <> ZoomLevel_Sci1)
				SendMessage, SCI_SETZOOM	, % ZoomLevel_Sci1, 0,, % "ahk_id " hScintilla1
			If (ZoomScintilla2O <> ZoomLevel_Sci2)
				SendMessage, SCI_SETZOOM	, % ZoomLevel_Sci2, 0,, % "ahk_id " hScintilla2

			WinMaximize, % "ahk_class SciTEWindow",, 0, 0

		}

	}

	If (hSpy := WinExist("WinSpy"))
		MoveWinToCenterScreen(hSpy)

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
		ADocList["Impferrinnerung"]    		:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "impf"	, "Do": "Choose"	, "RichE": "< /TdPP>< /, Pneumovax>< /, GSI>< /, MMR>"                                    	 }
		ADocList["Impfausweis"]     	    	:= {"Edit2": "31.12.%A_YYYY%"	, "Edit3": "impf"	, "Do": "Set"		, "RichE": "Impfausweis mitbringen"                                                                                	 }
		ADocList["lko - Gespräch"]	        	:= {"Edit2": "="							, "Edit3": "lko"	, "Do": "Input"	, "RichE": "03230(x:*)"                                                                                		        	 }
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

FillDocLine: ;{

	Gui, Suggestions: Submit
	actrl:= AlbisGetActiveControl("content")

	If !(ADocList[(matched)]["Edit2"]="=")	{
		n:= (A_MM=12) ? 1 : 0
		replacement:= StrReplace(ADocList[(matched)]["Edit2"], "%A_YYYY%", A_YYYY + n)		;<-später ein RegExMatch oder Replace für mehr Optionen
		VerifiedSetText("Edit2", replacement , "ahk_class OptoAppClass", 100)
	}

	VerifiedSetText("Edit3", ADocList[(matched)]["Edit3"] , "ahk_class OptoAppClass", 100)

	If (ADocList[(matched)]["Do"]="Choose")	{
		VerifiedSetText("RichEdit20A1", ADocList[(matched)]["RichE"] , "ahk_class OptoAppClass", 100)
		ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
		Sleep, 200
		SendInput, {LControl Down}{Right}{LControl Up}
	}

	If (ADocList[(matched)]["Do"]="Set")	{
		VerifiedSetText("RichEdit20A1", ADocList[(matched)]["RichE"] , "ahk_class OptoAppClass", 100)
		ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
		Sleep, 100
		SendInput, {Tab}
	}

	If (ADocList[(matched)]["Do"]="Input")	{
		InputBox, faktor, Addendum für AlbisOnWindows, bitte tragen Sie hier den faktor für die Gesprächsziffer ein.,,,,,,,, 2
		VerifiedSetText("RichEdit20A1", StrReplace(ADocList[(matched)]["RichE"], "*", faktor), "ahk_class OptoAppClass", 100)
		ControlFocus, RichEdit20A1, % "ahk_class OptoAppClass"
		Sleep, 100
		SendInput, {Tab}
	}

	gosub SuggestionsGuiEscape

return ;}

SuggestionsGuiClose:
	Gui, Suggestions: Hide
SuggestionsGuiEscape: ;{
	;HotKey, Enter            	, Off
	;HotKey, NumpadEnter	, Off
	;HotKey, Esc	             	, Off
	AutoDocCalled:=0
Return ;}

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

		global 	MedT
		static 	MedTrenner, MedTTrenner, x, y

	; Standardschrift im Albisprogramm zur Anzeige der Dauermedikamente ist Arial Standard 11 - am besten mit dieser Schriftart editieren und hier einfügen
		If (MedTrenner = "") {
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

		x := A_CaretX, y := A_CaretY + 15

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
		WinActivate   	, Dauermedikamente ahk_class #32770
		WinWaitActive	, Dauermedikamente ahk_class #32770,, 2
		y -= 15
		Click, %x%, %y%, 2
		Sleep, 200
		SendRaw, % MedTTrenner
		SendInput, {Enter}

MedTGuiClose:
MedTGuiEscape:
		Gui, MedT: Destroy
return
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
		Gui, adm: 	Destroy
		Gui, adm2: 	Destroy
	}

	FormatTime	, Time      	, % A_Now    	, dd.MM.yyyy HH:mm:ss
	FormatTime	, TimeIdle	, % A_TimeIdle	, HH:mm:ss
	FileAppend	, % Time ", " A_ScriptName ", " A_ComputerName ", " ExitReason ", " ExitCode ", " TimeIdle "`n", % Addendum.AddendumDir "\logs'n'data\OnExit-Protokoll.txt"

	For i, hook in Addendum.Hooks
		UnhookWinEvent(hook.hEvH, hook.HPA)

	Progress	, % "B2 B2 cW202842 cBFFFFFF cTFFFFFF zH25 w400 WM400 WS500"
					, % "Addendum wird beendet."
					, % "Addendum für AlbisOnWindows"
					, % "by Ixiko vom " DatumVom
					, % "Futura Bk Bt"

	Loop 50 {
		Progress % (100 - (A_Index * 2))
		Sleep 1
	}

	Progress, Off
	Gdip_Shutdown(pToken)

ExitApp
}

istda: ;{

	Progress	, % "B2 B2 cW202842 cBFFFFFF cTFFFFFF zH25 w400 WM400 WS500"
					, % "Addendum wird beendet."
					, % "Addendum für AlbisOnWindows"
					, % "by Ixiko vom " DatumVom
					, % "Futura Bk Bt"

	Loop 50 {
		Progress % (100 - (A_Index * 2))
		Sleep 1
	}

	Progress, Off
	Gdip_Shutdown(pToken)

ExitApp
;}

;}

;{22. #Include(s)

;------------ Addendumbibliotheken für Albis on Windows --------
#Include %A_ScriptDir%\..\..\include\Addendum_Ini.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gui.ahk
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
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk

;------------ allgemeine Bibliotheken ---------------------------------
#Include %A_ScriptDir%\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\lib\class_CtlColors.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#Include %A_ScriptDir%\..\..\lib\class_ImageButton.ahk
#Include %A_ScriptDir%\..\..\lib\class_IPHelper.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_socket.ahk
#Include %A_ScriptDir%\..\..\lib\crypt.ahk
#Include %A_ScriptDir%\..\..\lib\Explorer_Get.ahk
#Include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#include %A_ScriptDir%\..\..\lib\MenuSearch.ahk
#include %A_ScriptDir%\..\..\lib\NoTrayOrphans.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\lib\InputBoxEx.ahk
#Include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
#Include %A_ScriptDir%\..\..\lib\TrayIcon.ahk
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





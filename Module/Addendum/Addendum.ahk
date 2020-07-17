; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . .                                                                                       	ADDENDUM HAUPTSKRIPT
global                                                                               AddendumVersion:= "1.33" , DatumVom:= "17.07.2020"
; . . . . . . . . . .
; . . . . . . . . . .                                    ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . . . . . .                                           BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . . . . . .                                                        RUNS ONLY WITH AUTOHOTKEY_H IN 32 OR 64 BIT UNICODE VERSION
; . . . . . . . . . .                                                                 THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .                    !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !! ATTENTION !!
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . .                                                                    THIS SCRIPT ONLY WORKS WITH AUTOHOTKEY_H V1.X
; . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

/*               	A DIARY OF CHANGES
Beispiel:| **08.12.2018** | **F+** | IPC - Inter Process Communication - zwischen dem Addendum und Praxomat Skript steht, das Addendumskript kann somit weitere Funktionen übernehmen (V0.98b) |



 ** Registry Pfad für Albis
	Computer\HKEY_CURRENT_USER\Software\ALBIS

** Ideen
	OrderEntry Labor - automatische Eintragung der richtigen Ausnahmekennziffer

*/

	SciTEOutput(";------------------------------------------------------------------------------------------------------------------------------------------------`n")

;{01. Skriptablaufeinstellungen / Vorbereitungen

		#NoEnv
		#Persistent
		#InstallKeybdHook
		#SingleInstance               	, Force
		#MaxThreads                  	, 100
		#MaxThreadsBuffer       	, On
		#MaxHotkeysPerInterval	, 99000000
		#HotkeyInterval              	, 99000000
		#MaxThreadsPerHotkey 	, 2
		#KeyHistory                  	, 0
		;#Warn ;, All

		ListLines Off

		SetTitleMatchMode        	, 2	              	;Fast is default
		SetTitleMatchMode        	, Fast        		;Fast is default

		CoordMode, Mouse         	, Screen
		CoordMode, Pixel            	, Screen
		CoordMode, ToolTip       	, Screen
		CoordMode, Caret        	, Screen
		CoordMode, Menu        	, Screen

		SetKeyDelay                  	, -1, -1
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
		TTipCN := compName SubStr("                   ", 1, Floor((20 - StrLen(compName))/2))

	; Trayicon-Leichen entfernen
		NoTrayOrphansWin10()

	; auf hohe Priorität setzen aufgrund der Hook-Prozesse
		Process, Priority,, High

	; startet die Windows Gdip Funktion
		If !(pToken:=Gdip_Startup()) {
				MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
				ExitApp
		}

;}

;{02. Variablendefinitonen , individuelles Client-Traymenu wird hier erstellt

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; allgemeine globale Variablen
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global AlbisWinID      	:= AlbisWinID()                      	; für alle Funktionen muss die ID des Albis Hauptfenster ohne weitere Aufrufe als Variable vorhanden sein
		global Albis               	:= Object()                           	; enthält aktuelle Fenstergrößen und andere Daten aus den Albisfenstern
		global AddendumDir 	:= FileOpen("C:\albiswin.loc\AddendumDir","r").Read()
		global AlbisPID          	:= AlbisPID()                         	; mal sehen vielleicht benötige ich die Process ID von Albis für irgendwas

		global AutoDocCalled	:= 0                                      	; flag für AutoDoc Funktion welche nicht mehrfach gestartet werden kann (liest mehrfach aus der Ini aber nie richtig)
		global GVU                	:= Object()                           	; enthält Daten für die Vorsorgeautomatisierung
		global ScanPool        	:= Object()                            	; enthält die Dateinnamen aller im BefundOrdner vorhandenen Dateien

		global JSortDesc, JSortCol                                           	; Sortierung des Journals im Addendum Infofenster

		global PageSum        	:= 0                                     	; Gesamtseitenzahl aller Pdf-Dateien im BefundOrdner
		global DatumZuvor
		global q                     	:= Chr(0x22)                           	; ist dieses Zeichen -> "

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; globale und andere Variablen für die Hook's
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global EHProc		    		:= ""                                  	; der Name des Prozesses der gehookt wurde
		global EHEvent                                                          	; Event ID
		global EHookHwnd	    	:= 0                                   	; hook handle des allgemeinen WinEvent hooks
		global SHookHwnd	    	:= 0                                   	; hook handle des ShellHookProc
		global HmvHookHwnd  	:= 0                                   	; hook handle des für das Heilmittelverordnungsfensters (Physio)
		global HmvEvent		    	:= 0                                 	; event Nummer für Heilmittelverordnungsfenster
		global ifaptimer               	:= 0                                    	; ifap produziert eine Kaskade an Fehlermeldungen, flag wird bei erster Fehlermeldung gesetzt
		global Laborabruf	    		:= 0			          					; flag wird gesetzt wenn der Laborabruf statt findet, hat mir scheinbar mehrfach die Routine aufgerufen??warum??
		global AlbisHookTitle                                                 	; enthält den aktuellen Fenstertitel des Albisfenster
		global Laborabruf_Voll		:= 0                                    	; ist der Wert gleich 1 - wird der Laborabruf von Anfang bis zum kompletten übertragen aller Werte in die Patientenakte ausgeführt
		global Laborabruf_Status	:= 0                                 	; flag wird auf 1 gesetzt wenn die Routine läuft
		global EHWHStatus        	:= false                               	; flag falls gerade noch der Hookhandler läuft
		global ClickReady            	:= 1                             		; für den GVU Formularausfüller
		global AutoLogin            	:= true                                 ; der Autologin Vorgang für Albis kann per Interskriptkommunikation ausgestellt werden
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; globale Variablen für die Addendum Patientendatenbank
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global oPat                    	:= Object()                       	; Patientendatenbank Objekt für Addendum
		global TProtokoll				:= []	                                	; alternatives Tagesprotokoll als Absicherung falls es wieder Probleme mit der Albisdatenbank gibt
		global SectionDate					                                	; [Section]-Bezeichnung in der Tagesprotokoll Datenbank
		FirstDBAcess                		:= 0

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; globale Variablen (Objekte) und andere Variablen für Hotstringfunktionen
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global AutoDiagnosen	:= Object()
		global hotstring           	:= Object()                           	; Daten des aktuell einzufügenden Hotstring
		global activeControl   	:= Object()                              	; enthält den Inhalt des aktuell fokusierten Controls bevor manuelle Eingaben erfolgen
		global ADInterception	:= false                                	; zum Unterbrechen einer Funktionsroutine innerhalb der Hotstringabfragen
		oDia				            	:= Object()                              	; Autodiagnose Objekt - um Abzweigungen von Diagnosen bei gleichem ersten Wort zu ermöglichen
		If FileExist(AddendumDir "\include\Addendum_Diagnosen.json")
		{
				JDia	:= Object()
				JDia	:= new JSONFile(AddendumDir "\include\Daten\Addendum_Diagnosen.json")
		}
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; ** EINSTELLUNGEN AUS DER ADDENDUM.INI EINLESEN / ADDENDUM OBJEKT MIT DATEN BEFÜLLEN **
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
	;                   	Beschreibung aller enthaltenen key:value Paare
	;            ~~~~~~~ Struktur des Addendum-Objektes ~~~~~~~
	;	Addendum.BefundOrdner  	:= Scan-Ordner für neue Befundzugänge
	;	Addendum.InfoWindow     	:= Posteingang Einblendung im Albis Patientenfenster
	;                                                    .Y,.W,.H     	- der Einblendungsbereich
	;                                               	 .LVScanPool - Objekt enthält die größe der ScanPool Listview (W,H)
	;	Addendum.Telegram          	:= enumeriertes Objekt (Bsp: Addendum.Telegram[1].TBotID)
	;                                               	 .BotName 	- Name des Bot's
	;		                                           	 .Token         - Telegram Bot ID
	;	                                               	 .ChatID     	- Telegram Chat ID
	;	Addendum.GVUListe          	:= Anzahl der vorhandenen Untersuchungen in der GVU-Liste

	;	ein Aufruf nur mit dem Verzeichnisnamen initialisiert den Ini-Pfad - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		IniReadExt(AddendumDir "\Addendum.ini")
	;	- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		global Addendum	:= Object()
		global SPData      	:= Object()

		Addendum.Debug := false                                                	; gibt Daten in meine Standardausgabe (Scite) aus wenn = true

	; Vordefinition für Subobjekte + diverse Variablen	;{

		Addendum.AlbisMenu            	:= Object()                        	; wm_command Daten für Dialogaufrufe in Albis
		Addendum.Chroniker            	:= Array()                         	; Patient ID's der Chroniker
		Addendum.Geriatrisch           	:= Array()                         	; Patient ID's für geriatrisches Basisassement
		Addendum.Hooks                  	:= Object()                       	; enthält Adressen der Callbacks der Hookprozeduren und anderes
		Addendum.InfoWindow            	:= Object()                        	; InfoFenster/AddendumGui - Einstellungen
		Addendum.Labor	    	        	:= Object()							; Laborabruf und andere Daten für Laborprozesse
		Addendum.Laborabruf            	:= Object()
		Addendum.LAN	        	        	:= Object()							; LAN Kommunikationseinstellungen
		Addendum.Praxis                     	:= Object()                        	; Praxisadressdaten
		Addendum.Praxis.Sprechstunde	:= Object()                       	; Daten zu Öffnungszeiten, Urlaubstagen
		Addendum.Praxis.Email          	:= Array()                         	; Praxis Email Adressen
		Addendum.Praxis.Urlaub         	:= Array()                          	; Praxis Urlaubsdaten
		Addendum.Telegram             	:= Object()                        	; Datenobject für einen oder mehrere Telegram-Bots
		Addendum.Thread                  	:= Array()                          	; enthält die PID's gestarteter Threads (z.B. AddendumToolbar)
		Addendum.MsgGui               	:= Object()                        	; Gui für Interskript Kommunikation
		Addendum.Drucker               	:= Object()                          	; verschiedene Druckereinstellungen

		Addendum.AktuellerTag	    	:= A_DD                            	; der aktuelle Tag

		; Hotkey Tips für die Statusbar von Albis
		Addendum.Help                     	:= " | Strg+Pfeil runter - Akte schliessen | Strg+Alt+F5 - Datum(HEUTE) | Strg+F7/Alt+F7 - GVU | Strg+F10 - ScanPool | Alt+m - Menusuche |"

	;}

	; Verzeichnisse und Dateipfade                   	;{

		Addendum.AddendumDir                	:= AddendumDir
		Addendum.AddendumIni               	:= AddendumDir "\Addendum.ini"
		Addendum.AdditionalData_Path       	:= IniReadExt("Addendum"           	, "AdditionalData_Path"  	, AddendumDir "\include\Daten")
		Addendum.DosisDokumentenPfad   	:= IniReadExt("Addendum"            	, "DosisDokumentenPfad")                 	; Pfad zu MS Word Dokumenten mit eigenen Hinweisen zu Medikamentendosierungen
		Addendum.DBPath                          	:= IniReadExt("Addendum"          	, "AddendumDBPath")                       	; Datenbankverzeichnis
		Addendum.BefundOrdner                	:= IniReadExt("ScanPool"             	, "BefundOrdner")                             	; BefundOrdner = Scan-Ordner für neue Befundzugänge
		Addendum.ExportOrdner                	:= IniReadExt("ScanPool"             	, "ExportOrdner", Addendum.BefundOrdner)
		Addendum.TPPath	                        	:= Addendum.DBPath "\Tagesprotokolle\" A_YYYY                                	; Tagesprotokollverzeichnis
		Addendum.TPFullPath                      	:= Addendum.TPPath "\" A_MM "-" A_MMMM "_TP.txt"	                        	; Name des aktuellen Tagesprotokolls
		Addendum.AlbisExe                         	:= IniReadExt("Albis"                  	, "AlbisExe" )                                  	; Pfad zum Albis-Stammverzeichnis --- es wäre besser diesen aus der Registry auszulesen!
		;Addendum.7ZipDir                        	:= RegRead64("HKEY_CURRENT_USER\Software\7-Zip", "Path64")

	;}

	; Standard-Einstellungen für Gui's                  	;{

		; sie können die Schriftart und das Farbdesign der Addendum-Dialogfenstern ändern (in der Addendum.ini Datei!)
		Addendum.StandardFont                	:= IniReadExt("Addendum"         	, "StandardFont")
		Addendum.StandardBoldFont          	:= IniReadExt("Addendum"         	, "StandardBoldFont")
		Addendum.StandardFontSize          	:= IniReadExt("Addendum"         	, "StandardFontSize")
		Addendum.DefaultFntColor            	:= IniReadExt("Addendum"         	, "DefaultFntColor")
		Addendum.DefaultBGColor            	:= IniReadExt("Addendum"         	, "DefaultBGColor")
		Addendum.DefaultBGColor1          	:= IniReadExt("Addendum"         	, "DefaultBGColor1")
		Addendum.DefaultBGColor2          	:= IniReadExt("Addendum"         	, "DefaultBGColor2")
		Addendum.DefaultBGColor3          	:= IniReadExt("Addendum"         	, "DefaultBGColor3")
		Addendum.hashtagNachricht         	:= IniReadExt("Addendum"         	, "HASHtagNachricht"	, "")
		Addendum.dpiF                               	:= screenDims().DPI / 96                                                                   	; DPI-Faktor

	;}

	; Gesundheitsvorsorgeautomatisierung       	;{

		Addendum.GVUAbstand                 	:= IniReadExt("Abrechnungshelfer"	, "minGVU_Abstand")
		Addendum.GVUminAlter                	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlter")
		Addendum.GVUminAlterEinmalig   	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterEinmalig")
		Addendum.GVUminAlterMehrmalig	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterMehrmalig")

	;}

	; Leistungskomplexautomatisierung            	;{

		; Chroniker und Geriatrie Index werden eingelesen
		ChronikerListe()
		GeriatrischListe()

	;}

	; Hook-Flags                                                	;{

		Addendum.Laborabruf.Status	:= false
		Addendum.Laborabruf.Daten:= false
		Addendum.Laborabruf.Voll  	:= false

		Addendum.useraway         	:= false        ; true wenn Nutzer einen bestimmten Zeitraum keine Eingaben gemacht hat

	;}

	; Sicherheitsfunktionseinstellungen              	;{

		Addendum.STOut                            	:= IniReadExt(compname          	, "Verstecke_Alle_Daten"    	     , 60	)
		Addendum.STEnd                           	:= IniReadExt(compname           	, "Zeige_Alle_Daten_Hotstring", "xq")

	;}

	; diverse handles                                         	;{

		Addendum.AlbisPID                        	:= AlbisPID()
		Addendum.scriptPID                       	:= DllCall("GetCurrentProcessId")                                          	; die ScriptPID wird für das SkriptReload und die Interskript-Kommunikation benötigt

	;}

	; Skripte die als Threads geladen werden    	;{

		Addendum.ToolbarSkript               	:= FileOpen(AddendumDir "\Module\Addendum\threads\AddendumToolbar.ahk","r").Read()
		If AlbisExist()
			Addendum.Thread[1] := AHKThread(Addendum.ToolbarSkript)

	;}

	; zu- und abschaltbare Funktionen             	;{

		Addendum.AlbisLocationChange  	:= IniReadExt(compname, "Albis_AutoGroesse"                          , "ja")      	; automatische Größe von Albis ist an
		Addendum.WordAutoSize              := IniReadExt(compname, "Microsoft_Word_AutoGroesse"          , "nein")    	; automatische Größenanpassung eines Wordfenster
		Addendum.GVUAutomation			:= IniReadExt(compname, "GVU_automatisieren"                       , "nein")    	; Automatisierung der GVU Formulare
		Addendum.PDFSignieren   	    	:= IniReadExt(compname, "FoxitPdfSignature_automatisieren"  	, "nein")    	; Automatisierung PDF Signierung
		Addendum.PDFSignieren_Kuerzel	:= IniReadExt(compname, "FoxitPdfSignature__Kuerzel" 	        	, "Insert")  	; Tastaturkürzel zum Auslösen des Signaturvorganges
		Addendum.AutoCloseFoxitTab   	:= IniReadExt(compname, "FoxitTabSchliessenNachSignierung"  	, "Ja")    	; signiertes Dokument automatisch schliessen
		Addendum.Labor.AbrufTimer         := IniReadExt(compname, "Laborabruf_Timer"           	   	    		, "nein")    	; zeitgesteuerter Laborabruf
		Addendum.Labor.AutoAbruf 		    := IniReadExt(compname, "Laborabruf_automatisieren"              , "nein")    	; Laborabruf bei manuellem Start automatisieren
		Addendum.AddendumGui          	:= IniReadExt(compname, "Infofenster_anzeigen"                     	, "ja")      	; Infofenster das den Inhalt des Bildvorlagen-Ordners anzeigt

	;}

	; Labor abrufen - manuell oder zeitgesteuert	;{

		Addendum.Labor.LDTDirectory     	:= IniReadExt("Labor", "LDTDirectory"           	    , "C:\Labor")
		Addendum.Labor.LaborName       	:= IniReadExt("Labor", "LaborName"             	    , "")              	; falls sie mehrere Laboreinträge haben (z.B. nach Wechsel) tragen Sie das aktuelle hier ein
		Addendum.Labor.Laborkuerzel     	:= IniReadExt("Labor", "Aktenkuerzel"           	    , "labor")      	; Karteikartenkürzel um Informationen ablegen zu können
		Addendum.Labor.Alarmgrenze     	:= IniReadExt("Labor", "Alarmgrenze"           	    , "30%")	    	; Alarmierungsgrenze in Prozent oberhalb der Normgrenzen
		Addendum.Labor.AbrufZeiten      	:= IniReadExt("Labor", "LaborAbruf_Zeiten"   	    , "")	    		; z.B. "06:00, 15:00, 19:00, 21:00"
		Addendum.Labor.Kennwort       	:= IniReadExt("Labor", "LaborKennwort"       	    , "")	    		; für order&entry per CGM-Channel
		If InStr(Addendum.Labor.Kennwort, "Error")
			Addendum.Labor.Kennwort := ""

		; ACHTUNG:
		; hier werden die Einstellungen für eine Alarmierung über Ihren eigenen Telegram-Bot eingelesen! Denken Sie immer daran die Telegram Token und ChatID's vor dem
		; Zugriff Fremder oder auch dem eigenen Personal zu schützen
		Addendum.Labor.TGramOpt        	:= IniReadExt("Labor", "TGramOpt"           	    , "nein")	        	; die Nummer ihres Telegram Bots und + oder - Tel z.b. "1+Tel"
		If !((Addendum.Labor.TGramOpt = false) || RegExMatch(Addendum.Labor.TGramOpt, "^\d+\+|\-Tel")) 	; die Werte müssen einen exakten Syntax haben, sonst wird die Telegram Option gelöscht!
			IniWrite, % "nein", % Addendum.AddendumIni, % "Labor", % "TGramOpt"

	;}

	; Pdf Verarbeitung                                       	;{

		Addendum.xpdfPath				       	:= IniReadExt("ScanPool"                	, "xpdfPath"                                 	, AddendumDir "\include\xpdf")
		Addendum.PDFReaderFullPath      	:= IniReadExt("ScanPool"                	, "PDFReaderFullPath"                 	, AddendumDir "\include\FoxitReader\FoxitReaderPortable\FoxitReaderPortable.exe")
		Addendum.PDFReaderName        	:= IniReadExt("ScanPool"                	, "PDFReaderName"     	            	, "FoxitReader")
		Addendum.PDFReaderWinClass    	:= IniReadExt("ScanPool"                	, "PDFReaderWinClass"               	, "classFoxitReader")
		Addendum.SignatureCount        	:= IniReadExt("ScanPool"                	, "SignatureCount")
		Addendum.Ort                             	:= IniReadExt("ScanPool"                	, "Ort")
		Addendum.Grund                        	:= IniReadExt("ScanPool"                	, "Grund")
		Addendum.SignierenAls             	:= IniReadExt("ScanPool"                	, "SignierenAls")
		Addendum.DokumentSperren     	:= IniReadExt("ScanPool"                	, "DokumentNachDerSignierungSperren", "ja")
		Addendum.Darstellungstyp         	:= IniReadExt("ScanPool"                	, "DarstellungsTyp")
		Addendum.PasswortOn               	:= IniReadExt("ScanPool"                	, "PasswortOn", "nein")
		Addendum.PatAkteSofortOeffnen  	:= IniReadExt("ScanPool"                	, "Patientenakte_sofort_oeffnen"   	, "ja")
		Addendum.SignatureWidth          	:= IniReadExt("ScanPool"                	, "Signature_Breite"                    	, 50)
		Addendum.SignatureHeight          	:= IniReadExt("ScanPool"                	, "Signature_Hoehe"                    	, 25)

		SPData.xpdfPath				           	:= IniReadExt("ScanPool"                   	, "xpdfPath"                               	, AddendumDir "\include\xpdf")
		SPData.PDFReaderFullPath           	:= IniReadExt("ScanPool"                  	, "PDFReaderFullPath"                	, AddendumDir "\include\FoxitReader\FoxitReaderPortable\FoxitReaderPortable.exe")
		SPData.PDFReaderName              	:= IniReadExt("ScanPool"                 	, "PDFReaderName"                   	, "FoxitReader")
		SPData.PDFReaderWinClass         	:= IniReadExt("ScanPool"                 	, "PDFReaderWinClass"              	, "classFoxitReader")
		SPData.BefundOrdner                  	:= Addendum.BefundOrdner
		SPData.MaxFiles                           	:= ScanPoolArray("Load", Addendum.BefundOrdner "\PdfIndex.txt")

	;}

	; Telegram Bot Daten                                   	;{

		BotNr := 1
		Loop {

				BotName	:= IniReadExt("Telegram", "Bot" BotNr)
				If InStr(BotName, "Error") {
						If (A_Index = 1)
								BotNr := 0
						break
				}

				Addendum.Telegram.Push(ReadBots(BotName))
				BotNr ++

		}

		;}

	; Einstellungen des Infofensters einlesen      	;{

		ReadInfoWindowSettings()

	;}

	; Praxisdaten                                             	;{       für Outlook, Telegram Nachrichtenhandling

		Addendum.Praxis.Name            	:= IniReadExt("Allgemeines", "PraxisName")
		Addendum.Praxis.Strasse           	:= IniReadExt("Allgemeines", "Strasse")
		Addendum.Praxis.PLZ                	:= IniReadExt("Allgemeines", "PLZ")
		Addendum.Praxis.Ort                	:= IniReadExt("Allgemeines", "Ort")
		Addendum.Praxis.Sprechstunde    	:= Sprechstunde(IniReadExt("Allgemeines", "Sprechstunde"))

		Addendum.ShotDown_Leerlaufzeit	:= IniReadExt("Addendum", "ShotDown_Leerlaufzeit", "60")
		Addendum.AutoShotDown         	:= IniReadExt(compname, "AutoShotDown")      	; automatisches Herunterfahren des PC nach Feierabend

		;If Addendum.AutoShotDown
			;RunTask(weekday, time)
		;~ If !InStr(Addendum.Sprechstunde, "Error") && InStr(Addendum.AutoShotDown, "Error") {
					;~ AutoShutDownEinstellung(Addendum.ShotDown_Leerlaufzeit)
		;~ }
		;~ AutoShutDownEinstellung(Leerlaufzeit) {

				;~ hinweis =
				;~ (LTrim
				;~ Für diesen PC können Sie jetzt das automatische Herunterfahren nach Feierabend aktivieren.
				;~ Das Herunterfahren des PC wird nach bestimmter
				;~ )

		;~


		If Addendum.Debug {
			;SciTEOutput(": "  (	= true ? "true": "false"))
			SciTEOutput("Addendum.AddendumGui: "      	 (Addendum.AddendumGui       	= true ? "true": "false"))
			SciTEOutput("Addendum.Labor.AutoAbruf: "        (Addendum.Labor.AutoAbruf     	= true ? "true": "false"))
			SciTEOutput("Addendum.Labor.AbrufTimer: "       (Addendum.Labor.AbrufTimer    	= true ? "true": "false"))
			SciTEOutput("Addendum.PDFSignieren: "         	 (Addendum.PDFSignieren         	= true ? "true": "false"))
			SciTEOutput("Addendum.WordAutoSize: "        	 (Addendum.WordAutoSize        	= true ? "true": "false"))
			SciTEOutput("Addendum.AlbisLocationChange: " (Addendum.AlbisLocationChange	= true ? "true": "false"))
			SciTEOutput("Addendum.Labor.TGramOpt: "  	 (Addendum.Labor.TGramOpt    	= true ? "true": "false"))
			SciTEOutput("Addendum.PatAkteSofortOeffnen: " (Addendum.PatAkteSofortOeffnen = true ? "true": "false"))
			SciTEOutput("Addendum.AutoCloseFoxitTab: " 	 (Addendum.AutoCloseFoxitTab  	= true ? "true": "false"))
			SciTEOutput("Addendum.DokumentSperren: " 	 (Addendum.DokumentSperren     	= true ? "true": "false"))
			SciTEOutput("Addendum.PasswortOn: " 		      	 (Addendum.PasswortOn             	= true ? "true": "false"))
			SciTEOutput("Addendum.AutoShotDown: "     		 (Addendum.AutoShotDown         	= true ? "true": "false"))
		}

		AddendumGetMailAndHolidays()

	;}

	; LAN - Kommunikation Addendumskripte  	;{



		;~ AddendumGetKommunikationProperties()
		;~ AddendumGetKommunikationProperties() {

				;~ Addendum.LAN.MasterName := IniReadExt(Addendum, "MasterName")
				;~ If InStr(Addendum.LAN.MasterName, "Error") {


				;~ }

		;~ }


	;}

	; Albis Menu                                                 	;{

		Addendum.AlbisMenu.Privatliquidation:= { "Auswahlliste": "33023"
			, "Behandlungsliste": "32891"
			, "Ausgangsbuch": "33125"
			, "Offene_Posten": "32892"
			, "Mahnliste": { "Alle": "33145"
				, "Mahnstufe_1": "33142"
				, "Mahnstufe_2": "33143"
				, "Mahnstufe_3": "33144"
				, "Mahnbescheid": "34192"
				, "Fällige": "33850"}
			, "Quittungsliste": "32893"
			, "Stornierte_Restbeträge": "34864"
			, "Journal": "32894"
			, "Stornierte": "33851"
			, "Buchungsliste": "33841"
			, "Rechnungen_Buchungen": "33842"
			, "Rechnungen_Mahnungen": "33844"
			, "Kostenplan": "34067"
			, "Faktorzuordnungen": "34155"
			, "Quittungsliste_löschen": "32895"
			, "Sachkostenaufstellung": "34080"
			, "KH-Abschlag-_und_Vorteilsausgleich": "34156"}

	;}

	; Karteikarte - Rechtsklickmenu                   	;{

		Addendum.Drucker.Standard          	:= IniReadExt(compname, "Drucker_Standard")
		Addendum.Drucker.StandardA4     	:= IniReadExt(compname, "Drucker_Standard_A4")
		Addendum.Drucker.StandardA4Tray 	:= IniReadExt(compname, "Drucker_Standard_A4_Tray")
		Addendum.Drucker.PDF                   	:= IniReadExt(compname, "Drucker_PDF")
		Addendum.Drucker.FAX                   	:= IniReadExt(compname, "Drucker_FAX")

	;}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Tray Menu erstellen
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{

		;fnCall := Object()
		func_NoNo                     	:= Func("NoNo")
		func_ZeigePatDB             	:= Func("ZeigePatDB")
		func_ZeigeFehlerProtokoll	:= Func("ZeigeFehlerProtokoll")
		func_ZeigeOnExitProtokoll	:= Func("ZeigeOnExitProtokoll")
		func_ZeigeGVUs              	:= Func("GVU_Gui")

		If (hIconAddendum:= Create_Addendum_ico())
			Menu, Tray, Icon				, % "hIcon: " hIconAddendum
		else
			Menu, Tray, Icon				, % A_ScriptDir "\Addendum.ico"

		Menu, Tray, NoStandard

		Menu, Tray, Tip					, % StrReplace(A_ScriptName, ".ahk") " V." AddendumVersion " vom " DatumVom
													. "`nClient: " compName SubStr("                   ", 1, Floor((20 - StrLen(compName))/2))
													. "`nPID: " DllCall("GetCurrentProcessId")
													. "`nAutohotkey.exe: " A_AhkVersion

		Menu, Tray, Add				, Addendum, % func_NoNo

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Module laden
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		modul 	:= Object()
		tool   	:= Object()

	; Clientabhängige Authorisierung laden
		Auth 	:= IniReadExt(compname, "Module")
		If InStr(Auth, "All") 	{

				Auth := ""
				Loop
					If InStr(IniReadExt("Module", "Modul" SubStr("0" . A_Index, -1)), "ERROR")
							break
					else
							Auth .= A_Index ","

				Auth := RTrim(Auth, ",")

		}

	; authorisierte Module ins Traymenu laden
		Loop {

				Modultemp:= IniReadExt("Module", "Modul" SubStr("0" A_Index, -1))
				Tooltemp	 := IniReadExt("Module", "Tool" 	 SubStr("0" A_Index, -1))

				If !InStr(ModulTemp, "ERROR") {
						tmp := StrSplit(Modultemp, "|")
						modul[(tmp[2])] := tmp[3]

						If InStr(tmp[1], "NoAuth")
								Menu, SubMenu1, Add, % tmp[2], ModulStarter
						else if A_Index in %Auth%
								Menu, SubMenu1, Add, % tmp[2], ModulStarter
				}

				If !InStr(ToolTemp, "ERROR")	{
						tmp :=StrSplit(StrReplace(ToolTemp, "Ã¼", "ü"), "|")
						tool[(tmp[2])] := tmp[3]

						If InStr(tmp[1], "NoAuth")
								Menu, SubMenu2, Add, % tmp[2], ToolStarter
						else if A_Index in %Auth%
								Menu, SubMenu2, Add, % tmp[2], ToolStarter
				}

				If InStr(ModulTemp, "ERROR") && InStr(ToolTemp, "ERROR")
						break
		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; per Checkbox zu- oder abschaltbare Funktionen (SubMenu3)
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	  ; automatische Anpassung von Albis an die Monitorauflösung 	;{
		If InStr(compname, "SP1") {
			Menu, SubMenu3, Add, % "Pause Addendum"	, Menu_PauseAddendum
			Menu, SubMenu3, Add, % "Albis AutoSize"      	, Menu_AlbisAutoPosition
			If Addendum.AlbisLocationChange
				Menu, SubMenu3, Check	  , % "Albis AutoSize"
			else
				Menu, SubMenu3, UnCheck, % "Albis AutoSize"
		}
		;}

	  ; AddendumGui - Infofenster                                                	;{
		Menu, SubMenu3, Add, % "Addendum Infofenster", Menu_AddendumGui
		If Addendum.AddendumGui
			Menu, SubMenu3, Check	  , % "Addendum Infofenster"
		else
			Menu, SubMenu3, UnCheck, % "Addendum Infofenster"
	  ;}

	  ; Automatisierung der GVU Formulare                                  	;{
		Menu, SubMenu3, Add, % "Albis GVU automatisieren", Menu_GVUAutomation
		If Addendum.GVUAutomation
			Menu, SubMenu3, Check	  , % "Albis GVU automatisieren"
		else
			Menu, SubMenu3, UnCheck, % "Albis GVU automatisieren"
		;}

	  ; Automatisierung der PDF Signierung mit dem FoxitReader    	;{
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

	  ; Laborabruf automatisieren für manuelle Vorgänge, das Menu für den Client der regelmäßig die Labordatenabrufen soll wird hier auch erstellt ;{

		If !InStr(Addendum.Labor.AutoAbruf, "Error") {

			Menu, SubMenu3, Add, % "Laborabruf automatisieren", Menu_LaborAbrufManuell
			If Addendum.Labor.AutoAbruf
				Menu, SubMenu3, Check	  , % "Laborabruf automatisieren"
			else
				Menu, SubMenu3, UnCheck, % "Laborabruf automatisieren"

		}

		If !InStr(Addendum.Labor.AbrufTimer, "Error") {

			; zeitgesteuert
				Addendum.Labor.AutoAbruf := true
				Menu, SubMenu3, Add, % "zeitgesteuerter Laborabruf", Menu_LaborAbrufTimer
				If Addendum.Labor.AbrufTimer
					Menu, SubMenu3, Check	  , % "zeitgesteuerter Laborabruf"
				else
					Menu, SubMenu3, UnCheck, % "zeitgesteuerter Laborabruf"

		} ;}

	  ; MS Word mit fester Größe ;{
		Menu, SubMenu3, Add, % "Microsoft Word AutoSize", Menu_MSWordAutoPosition
		If Addendum.WordAutoSize
			Menu, SubMenu3, Check	  , % "Microsoft Word AutoSize"
		else
			Menu, SubMenu3, UnCheck, % "Microsoft Word AutoSize"
	;}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; weitere Menupunkte
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

		Menu, SubMenu4, Add, % "Patienten Datenbank anzeigen"	, % func_ZeigePatDB
		Menu, SubMenu4, Add, % "Vorsorgeuntersuchungen"         	, % func_ZeigeGVUs
		Menu, SubMenu4, Add, % "aktuelles Fehlerprotokoll"        	, % func_ZeigeFehlerProtokoll
		Menu, SubMenu4, Add, % "OnExit Protokoll"                    	, % func_ZeigeOnExitProtokoll

		Menu, Tray, Add, Module starten/beenden   	, :SubMenu1
		Menu, Tray, Add, Daten / Protokolle             	, :SubMenu4
		Menu, Tray, Add, andere Tools                   	, :SubMenu2
		Menu, Tray, Add, Einstellungen                     	, :SubMenu3

		Menu, Tray, Add,
		;Menu, Tray, Add, Zeige HotKey's                  	, Hotkeys
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
				FirstDBAcess := 1
		}
		If !InStr(FileExist(Addendum.DBPath), "D")    	&& (FirstDBAcess = 1)	{
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
		else if !InStr(FileExist(Addendum.DBPath), "D") && (FirstDBAcess = 0)	{

				;dieser Abschnitt ermittelt den neuen Pfad zur Addendumdatenbank
					ADBtmp:= StrReplace(Addendum.DBPath, "\_DB")
					SelectDBPath:
					FileSelectFolder, ADBtmp, *ADBtmp,, Wählen Sie den neuen Pfad zur Addendum Datenbank (..\_DB)
					If !ErrorLevel
					{
							ADBtmp:= RegExReplace(ADBtmp, "\\$")
							If !InStr(AddendumDBPath, "_DB")
							{
									MsgBox,, Addendum für AlbisOnWindows, % "Ihre Pfadangabe ist nicht richtig!`nWählen Sie den Ordner mit folgender Bezeichnung "  q "_DB" q "!"
									goto SelectDBPath
							}
							else if !FileExist(ADBtmp "\Patienten.txt") and !(FirstDBAcess=1)
							{
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
							else
							{
									Addendum.DBPath:= ADBtmp
									IniWrite, % Addendum.DBPath, % AddendumDir "\Addendum.ini", % "Addendum", % "AddendumDBPath"
							}

							VarSetCapacity(ADBtmp, 0)
					}
		}

		VarSetCapacity(FirstDBAcess, 0)
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Einlesen der Patientendaten aus der Patienten.txt Datei im Addendum Datenbankordner
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		If FileExist(Addendum.DBPath "\Patienten.txt")
			oPat:= ReadPatientDatabase(Addendum.DBPath)
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Einlesen des heutigen Tagesprotokoll bei Skript Neu- oder Restart , Anlegen eines neuen Ordner zum Jahreswechsel
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		TProtokoll := LeseTagesProtokoll()
	;}

;}


;{04. Timer, Threads, Initiliasierung WinEventHooks, OnMessage

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Starten clientunabhängiger Timer Labels
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		If IsLabel(nopopup_%compname%)
			SetTimer, nopopup_%compname%, 2000                      	;spezielle Labels die Funktionen nur auf einzelnen Arbeitsplätzen ausführen
		If IsLabel(seldom_%compname%)
			SetTimer, seldom_%compname%, 900000		            	; 900000ms = 1000 * 900s= 15 * 60s = 15 * 1min = 15min -> Ausführungszeit

		SetTimer, UserAway, 300000                                         	; zum Ausführen von Berechnungen wenn ein Computer unbenutzt ist (5min)
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;  Skriptkommunikation / Anpassen von Fenstern bei Änderung der Bildschirmaufläsung
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		OnMessage(0x4a, "Receive_WM_COPYDATA")
		OnMessage(0x7E, "WM_DISPLAYCHANGE")
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	;	SetWinEventHooks /Shellhook initialisieren
	;---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		InitializeWinEventHooks()
	;}

		PraxTT(oPat.Count() " Patienten gefunden.`n"
					. (TProtokoll.MaxIndex() = "" ? 0 : TProtokoll.MaxIndex() = 1 ? "1 Patient ist im Tagesprotokoll" : TProtokoll.MaxIndex() " Patienten sind im Tagesprotokoll.")
					. "`nSignaturen: " Addendum.SignatureCount, "3 0")

		If InStr(Addendum.AktiveAnzeige, "Patientenakte")         	{

				PatDb(AlbisTitle2Data(AlbisTitle), "exist")
				If Addendum.AddendumGui
						AddendumGui()

		}

;}

;{05. Hotkey command's

HotkeyLabel:
	;=--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Hotkey's die überall funktionieren
	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	Hotkey, $^!ä                	, BereiteDichVor                                 	;= Überall: beendet das Addendumskript
	Hotkey, $^!n                	, SkriptReload                                    	;= Überall: Addendum.ahk wird neu gestartet
	Hotkey, $^Ins               	, RunSomething                                 	;= Überall: für's programmieren gedacht, im Label das Skript eintragen das per Hotkey gestartet werden soll
	Hotkey, !m                    	, MenuSearch                                    	;= Überall: Menu Suche - funktioniert in allen Standart Windows Programmen
	Hotkey, !q                     	, ScreenShot                                      	;= Überall: erstellt einen Screenshot per Mausauswahl und legt diesen im Clipboard ab
	Hotkey, CapsLock          	, DoNothing                                      	;= Überall: Capslock kann nicht mehr versehentlich gedrückt werden
	Hotkey, ^!e                  	, EnableAllControls                            	;= Überall: inaktivierte Bedienfelder in externen Fenstern aktivieren
	Hotkey, ^#!t                 	, FuerTests                                         	;= Überall: für kurzen Programmcode der nur zu Testzwecken ausgeführt werden soll
	;Hotkey, ^#!o                 	, ImmerOnTop                                     	;= Überall: macht das ein Fenster immer vor den anderen liegt oder setzt den Zustand zurück
	Hotkey, ^#!i                  	, AddICDHotstring                                	;= Überall: eine ICD Diagnose zu den AutoDiagnosen hinzufügen
	Hotkey, ^!p                  	, PatientenkarteiOeffnen                       	;= Überall: eine selektierte Zahl wird zum Öffnen einer Patientenkarteikarte genutzt
	Hotkey, Pause               	, AddendumPausieren
	Hotkey, ^!F10               	, SendClipBoardText                           	;= Überall: sendet den Inhalt des Clipboards als simulierte Tasteneingabe (um z.B. wiederholt ein Passwort zu senden)

	;Hotkey, ~LButton         	, CaptureDoubleClick                        	;= Überall: wenn Text unter Maus einen Patientennamen oder eine ID enthält wird die Akte dazu geöffnet
	;Hotkey, ^!ScrollLock     	, ShowAll                                           	;= Überall: Sicherheitsfunktion - Hotkey um alles wieder anzuzeigen
	;Hotkey, ^!Pause           	, HideAll                                            	;= Überall: Sicherheitsfunktion - Drücken um alle Fenster die Patientendaten enthalten zu verstecken

	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Windows Datei Explorer
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	func_GetPatNameFromExplorer := Func("GetPatNameFromExplorer")
	Hotkey, IfWinActive, % "ahk_class CabinetWClass"
	Hotkey, ~F6                    	, % func_GetPatNameFromExplorer
	Hotkey, ~^RButton          	, % func_GetPatNameFromExplorer
	Hotkey, IfWinActive
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
    Hotkey, ^+h                            	, SciteHotKeyWriter                	;= SciTe: opens a gui for writing hotkey's because I ever forget the names
	Hotkey, IfWinActive
	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Albis
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
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
	Hotkey, #F7  					, % func_ZeigeGVUs                  			;= Albis: GVUListe anzeigen, bearbeiten und ausführen lassen
	;Hotkey, ^!F1					, AlbisFunktionen									;= Albis: Funktionen für Albis - Auswahl-Gui
	;Hotkey, If, InStr(AlbisGetActiveWindowType(), "Patientenakte")        	;~~ ALBIS Hotkey - nur auslösbar wenn eine Akte geöffnet ist
	Hotkey, !F8		    			, DruckeLabor                                    	;= Albis: Laborblatt als pdf oder auf Papier drucken
	Hotkey, !Down      			, AlbisKarteiKarteSchliessen        			;= Albis: per Kombination die aktuelle Karteikarte schliessen
	Hotkey, !Up               		, AlbisNaechsteKarteiKarte	        			;= Albis: per Kombination eine geöffnete Kartekarte vorwärts
	Hotkey, !Right    				, Laborblatt    										;= Albis: Laborblatt des Patienten anzeigen
	Hotkey, !Left  					, Karteikarte      									;= Albis: Karteikarte des Patienten anzeigen
	Hotkey, IfWinActive

	;~ InKarteikarte := Func("AlbisKarteikarteAktiv") ; das war ein Test
	;~ Hotkey, If, % InKarteikarte                                                           	;~~ ALBIS Hotkey - nur auslösbar wenn eine Akte geöffnet ist
	;~ Hotkey, ~RButton          	, KarteikartenMenu
	;~ Hotkey, If

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

	;HotKey, IfWinExist, LbAutoComplete                                            	;~~ Spezielle Listbox erreichbar durch ''+''
	;Hotkey, Enter, AutoCBLBres                                                        	;= Albis: übernimmt die Auswahl
	;Hotkey, IfWinExist

	Hotkey, IfWinActive, Dokumentationsbögen ahk_class #32770   	;~~ Korrekturhilfe bei falsch gesetzer Auswahl bei Hautkrebsverdacht
	Hotkey, Numpad0		    	, FastSets                                             	;= Albis: ruft Dokumentationsbogen auf, setzt Auswahl um,speichert Dokument u. rückt e. Patienten nach unten
	Hotkey, IfWinActive
	;}

	;=-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; flexible Hotkey's und Fensterkombinationen
	;=------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- ;{
	If Addendum.PDFSignieren 	{

			func_FRSignature := func("FoxitReader_SignaturSetzen")
			Hotkey, IfWinActive, % "ahk_class " Addendum.PDFReaderWinClass
			Hotkey, % Addendum.PDFSignieren_Kuerzel, % func_FRSignature
			Hotkey, IfWinActive

	}
	;}

return

GetPatNameFromExplorer() {                                                             	;-- ermittelt in einer pdf Dateibezeichnung den Patientennamen und öffnet die Patientenakte

	; Voraussetzung: der Patient ist in der Addendum Patientendatenbank gelistet, ansonsten passiert nichts
		SplitPath, % Explorer_GetSelected(WinExist("A")),,,, fname
		FuzzyKarteikarte(fname)

}

FuzzyKarteikarte(StrContainingPatientName) {                                    	;-- fuzzy name matching function, öffnet eine Karteikarte

		DiffSafe 	:= Array()
		minDiff		:= 100
		PatFound 	:= true
		umlauts	:= {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}

		fname 	:= StrReplace(StrContainingPatientName, ",", " ")
		fname 	:= StrReplace(fname, ";", " ")
		fname 	:= RegExReplace(fname, "\d+\.", "")
		fname	:= RegExReplace(fname, "([a-zöäüß][A-ZÄÖÜ])([a-zäöüß])", "$1 $2")
		fname 	:= StrReplace(fname, "  ", " ")
		name 	:= StrSplit(fname, " ")

	; Stringdifference Methode mit Suche des Patienten in der Addendum Patientendatenbank     	;{
		For PatID, Pat in oPat 		{

				z:=0
				If InStr(PatID, "MaxPat")
					break

				DbName	:= Trim(StrReplace(Pat.Nn Pat.Vn, " "))                	; NachnameVorname
				NVname 	:= Trim(Name.1 Name.2)                                 	; NachnameVorname
				VNname 	:= Trim(Name.2 Name.1)	                                	; VornameNachname
				DiffA     	:= StrDiff(DBName, NVname)
				DiffB     	:= StrDiff(DbName, VNname)

				If (DiffA < minDiff) {
					minDiff     	:= DiffA
					nearestDiff	:= PatID
				}
				If (DiffB < minDiff) {
					minDiff     	:= DiffB
					nearestDiff	:= PatID
				}

				If ( DiffA <= 0.11 ) || ( DiffB <= 0.11 ) {
					AdmKarteikarte(PatID)
					return PatFound
				}

				If ( DiffA <= 0.12 ) || ( DiffB <= 0.12 ) 	{

						If (Abs(StrLen(DbName) - StrLen(NVname)) > 2)  { 				; enthält eine der beiden Zeichenketten 2 Zeichen mehr als die andere wird der Treffer zunächst zwischengespeichert
							DiffSafe.Push(PatID)
						}
						;else  {
							;AdmKarteikarte(PatID)
							;return
						; }

				}

				If ( DiffA <= 0.21 ) || ( DiffB <= 0.21 )
						SciTEOutput(Pat.Nn ", " Pat.Vn)

		}

		SciTEOutput(minDiff ", " nearestDiff)

		If (minDiff < 0.2) {
				AdmKarteikarte(nearestDiff)
				return PatFound
		}

		If RegExMatch(DbName, "^" VNname) || RegExMatch(DbName, "^" NVname) 	{

				MsgBox, 4, Patientenakte öffnen?, % "Möchte Sie die Akte des Patienten:`n(" PatID ") " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`nöffnen?"
				IfMsgBox, Yes
					AlbisAkteOeffnen(Pat.Nn ", " Pat.Vn, PatID)
				return PatFound

		}


		PraxTT("Patientenname wurde nicht gefunden!`nVersuche es mit einer Patientensuche über den Albisdialog.", "6 0")
		Sleep, 500
		;}

	; Suchen über die Albis Patientensuche                                                                                	;{

		; Code nimmt die ersten beiden Buchstaben vom Vor- und Zuname und übergibt sie dem Patient öffnen Dialog
		; stehen mehrere Personen zur Auswahl im Dialog "Patient auswählen" wird per String Differenz der dichteste Treffer gesucht
		; wenn Albis für diese Kombination nur einen Treffer hat, wird sofort die Akte geöffnet (unabhängig vom Skript)
		; --- die umlautveränderten Namen werden hier für die Suche noch nicht genutzt!!
		PatMatches 	:= Object()

	; öffnet den Dialog "Öffne Patient" und übergibt nur die ersten beiden Buchstaben vom Vor- und Zunamen
		AlbisDialogOeffnePatient("open", SubStr(name[1], 1, 2) ", " SubStr(name[2], 1, 2))

		while !(WinExist("Patient auswählen ahk_class #32770") || WinExist("ALBIS ahk_class #32770", "Patient") || (A_Index > 40))
			Sleep 50

		; Patient ist nicht vorhanden ;{
			while WinExist("ALBIS ahk_class #32770", "Patient") 	{
					PatFound := false
					VerifiedClick("Button1", "ALBIS ahk_class #32770", "Patient") ; OK
					WinWaitClose, % "ALBIS ahk_class #32770", % "Patient", 1
					WinWait, % "Patient öffnen ahk_class #32770",, 1
					If WinExist("Patient öffnen ahk_class #32770")
						AlbisDialogOeffnePatient("close")
					If (A_Index > 10)
						break
					If WinExist("ALBIS ahk_class #32770", "Patient")
						sleep, 200
			}
		;}

		;WinWait, % "Patient auswählen ahk_class #32770",, 3

		If (hwnd := WinExist("Patient auswählen ahk_class #32770")) {

				LVPats	:= AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")
				For row, Patient in LVPats
					If (StrDiff(Patient.2 Patient.3, word.1 word.2) <= 0.12 )
							PatMatches.Push(Patient)

				If (PatMatches.MaxIndex() = 1) {

					hinweis :=  "Ich denke dieser Patient könnte es sein:`n`n" PatMatches[1][2] "," PatMatches[1][3] ", geb. am " PatMatches[1][4]
					hinweis .= "`nletzter Behandlungstag: " PatMatches[1][6] "`n`nSoll ich diese Akte öffnen?"

					Msgbox, 0x1024, Addendum für Albis on Windows, % hinweis
					IfMsgBox, Yes
					{
							VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true)
							WinWait, % "Patient öffnen ahk_class #32770",, 3
							If ErrorLevel
							{
									MsgBox, Der Patient öffnen Dialog fehlt mir jetzt`num weiter zu machen!
									return false
							}

							VerifiedSetText("Edit1", PatMatches[1][5], "Patient öffnen ahk_class #32770", 200)

						; Akte wird jetzt geöffnet durch drücken von OK ;{
							while WinExist("Patient öffnen ahk_class #32770") 	{
									; Button OK drücken
										VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
										WinWaitClose, Patient öffnen ahk_class #32770,, 1
									; Fenster ist immer noch da? Dann sende ein ENTER.
										if WinExist("Patient öffnen ahk_class #32770") {
												WinActivate, Patient öffnen ahk_class #32770
												ControlFocus, Edit1, Patient öffnen ahk_class #32770
												SendInput, {Enter}
										}

									If (A_Index > 10) {
											PatFound := false
											break
									}

									sleep, 200
							}
							;}

					}

				}


				VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true) ; Abbruch
				AlbisDialogOeffnePatient("close")

		}

		;}

return PatFound
}

FuzzyKarteikarte_umlauts(StrContainingPatientName) {                          	;-- fuzzy name matching function, öffnet eine Karteikarte

		DiffSafe := Array()

		PatFound := false
		umlauts	:= {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}

		fname 	:= StrReplace(StrContainingPatientName, ",", " ")
		fname 	:= StrReplace(fname, ";", " ")
		fname 	:= RegExReplace(fname, "\d+\.", "")
		fname	:= RegExReplace(fname, "([a-zöäüß][A-ZÄÖÜ])([a-zäöüß])", "$1 $2")
		fname 	:= StrReplace(fname, "  ", " ")
		word		:= StrSplit(fname, " ")
		mxwrds	:= word.MaxIndex()

	; wandelt Umlaute um, damit genauere Treffer erzielt werden können
		Loop, % word.MaxIndex() {

				If RegExMatch(word[A_Index], "[äüöÄÜÖß]") {

					; Kopie eines Namens mit einem Umlaut anlegen
						tmp := word[A_Index]

						For umlaut, replacement in umlauts
							tmp := RegExReplace(tmp, umlaut, replacement)

						word.Push(tmp)

				}
		}

		minDiff	:= 100

	; Stringdifference Methode mit Suche des Patienten in der Addendum Patientendatenbank     	;{
		For PatID, Pat in oPat
		{
				z:=0
				If InStr(PatID, "MaxPat")
					break

				Loop, % (word.MaxIndex() - 1)
				{

						DbName	:=StrReplace(Pat.Nn Pat.Vn, " ")                	; NachnameVorname
						fname    	:= word[A_Index]  	word[A_Index+1]		; NachnameVorname
						rname    	:= word[A_Index+1]	word[A_Index]	    	; VornameNachname
						DiffA     	:= StrDiff(DBName, fname)
						DiffB     	:= StrDiff(DbName, rname)

						If InStr(DbName, word[A_Index]) || InStr(DbName, word[A_Index+1])
								SciTEOutput("gefunden Name: " oPat[PatID].Nn ", " oPat[PatID].Vn " (" word[A_Index] ", " word[A_Index+1] ")", 0,1)

						If ( DiffA <= 0.12 ) || ( DiffB <= 0.12 ) 	{

								If Abs(StrLen(DbName) - StrLen(fname)) > 2 { 				; enthält eine der beiden Zeichenketten mehr Zeichen als die andere wird der Treffer zwischengespeichert
										DiffSafe.Push(PatID)
										continue
								}
								AdmKarteikarte(PatID)

						}

						If (DiffA < minDiff)
							minDiff	:= DiffA, nearestDiff	:= PatID
						 else If (DiffB < minDiff)
							minDiff	:= DiffB, nearestDiff	:= PatID

				}
		}

		SciTEOutput("minDiff: " DiffA ", PatIndex: " nearestDiff ", Name: " oPat[nearestDiff].Nn ", " oPat[nearestDiff].Vn, 0, 1)
		return

		If RegExMatch(DbName, "^" rname) || RegExMatch(DbName, "^" fname) 	{

				MsgBox, 4, Patientenakte öffnen?, % "Möchte Sie die Akte des Patienten:`n(" PatID ") " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`nöffnen?"
				IfMsgBox, Yes
					AlbisAkteOeffnen(Pat.Nn ", " Pat.Vn, PatID)
				return

		}


		PraxTT("Patientenname wurde nicht gefunden!`nVersuche es per direkten Aufruf über den Albisdialog.", "6 0")
		Sleep, 500
		;}

	; Suchen über die Albis Patientensuche                                                                                	;{

		; Code nimmt die ersten beiden Buchstaben vom Vor- und Zuname und übergibt sie dem Patient öffnen Dialog
		; stehen mehrere Personen zur Auswahl im Dialog "Patient auswählen" wird per String Differenz der dichteste Treffer gesucht
		; wenn Albis für diese Kombination nur einen Treffer hat, wird sofort die Akte geöffnet (unabhängig vom Skript)
		; --- die umlautveränderten Namen werden hier für die Suche noch nicht genutzt!!


		PatMatches := Object()

	; öffnet den Dialog "Öffne Patient" und übergibt nur die ersten beiden Buchstaben vom Vor- und Zunamen
		AlbisDialogOeffnePatient("open", SubStr(word.1, 1, 2) ", " SubStr(word.2, 1, 2))

		WinWait, % "Patient auswählen ahk_class #32770",, 3

		If (hwnd := WinExist("Patient auswählen ahk_class #32770")) {

				LVPats	:= AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")

				;~ SciTEOutput("", 1)
				;~ SciTEOutput("used name: " word.1 word.2, 0, 1)
				;~ SciTEOutput("Compared with: " , 0, 1)

				For row, Patient in LVPats
				{
						;~ SciTEOutput(Patient.2 Patient.3 "(" StrDiff(Patient.2 Patient.3, word.1 word.2) ") | " , 0, 0)
						If (StrDiff(Patient.2 Patient.3, word.1 word.2) <= 0.12 )
							PatMatches.Push(Patient)
				}

				;~ SciTEOutput("LVPats.MaxIndex(): " LVPats.MaxIndex(), 0, 1)
				;~ SciTEOutput("PatMatches.MaxIndex(): " PatMatches.MaxIndex(), 0, 1)
				;~ SciTEOutput("-------------------------------------------------------------", 0, 1)
				;~ SciTEOutput("", 0, 1)

				If (PatMatches.MaxIndex() = 1) {

					hinweis :=  "Ich denke dieser Patient könnte es sein:`n`n" PatMatches[1][2] "," PatMatches[1][3] ", geb. am " PatMatches[1][4]
					hinweis .= "`nletzter Behandlungstag: " PatMatches[1][6] "`n`nSoll ich diese Akte öffnen?"

					Msgbox, 0x1024, Addendum für Albis on Windows, % hinweis
					IfMsgBox, Yes
					{
							VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true)
							WinWait, % "Patient öffnen ahk_class #32770",, 3
							If ErrorLevel
							{
									MsgBox, Der Patient öffnen Dialog fehlt mir jetzt`num weiter zu machen!
									return
							}

							VerifiedSetText("Edit1", PatMatches[1][5], "Patient öffnen ahk_class #32770", 200)

						; Akte wird jetzt geöffnet durch drücken von OK ;{
							while WinExist("Patient öffnen ahk_class #32770")
							{
									; Button OK drücken
										VerifiedClick("Button2", "Patient öffnen ahk_class #32770")
										WinWaitClose, Patient öffnen ahk_class #32770,, 1
									; Fenster ist immer noch da? Dann sende ein ENTER.
										if WinExist("Patient öffnen ahk_class #32770")
										{
												WinActivate, Patient öffnen ahk_class #32770
												ControlFocus, Edit1, Patient öffnen ahk_class #32770
												SendInput, {Enter}
										}

									If (A_Index > 10)
											return

									sleep, 200
							}
							;}
					}

				}

				VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true) ; Abbruch
				AlbisDialogOeffnePatient("close")

		}

		;}

return
}

AdmKarteikarte(PatID) {

	If !Addendum.PatAkteSofortOeffnen {

			MsgBox, 4, Patientenakte öffnen?, % "Möchte Sie die Akte des Patienten:`n(" PatID ") " oPat[PatID].Nn ", " oPat[PatID].Vn " geb.am " oPat[PatID].Gd "`nöffnen?"
			IfMsgBox, No
				return

	}

	AlbisAkteOeffnen(oPat[PatID].Nn ", " oPat[PatID].Vn, PatID)

return
}

;}

;{06. Hotstrings

:*:#at::@

;{ 	6.1. -- ALBIS
; --- Leistungskomplexe                                                                      	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "lk")
:X*:gs::                                                                                         	;{ Gesprächziffer
	AlbisSchreibeLkMitFaktor("03230", "2")
return ;}
:X*:ps1::                                                                                       	;{ Psychosomatikziffer 1
:X*:psy1::
	AlbisSchreibeLkMitFaktor("35100", "2")
return ;}
:X*:ps2::                                                                                       	;{ Psychosomatikziffer 2
:X*:psy2::
	AlbisSchreibeLkMitFaktor("35110", "2")
return ;}
:X*:tel::                                                                                         	;{ Telefongebühr mit dem Krankenhaus
	AlbisSchreibeLkMitFaktor("80230", "2")
return ;}
:X*:kur::                                                                                        	;{ Kurplan, Gutachten, Stellungnahme
:X*:gut::
	AlbisSchreibeLkMitFaktor("01622", "")
return ;}
:X*:stix::                                                                                        	;{ Urinstix
:X*:urin::
:X*:ustix::
:X*:urinstix::
	AlbisSchreibeLkMitFaktor("32030", "")
return ;}
:X*:Kompr::                                                                                   	;{ Kompressionsverband
	AlbisSchreibeLkMitFaktor("02313", "")
return ;}
:X*:hb::                                                                                           	;{ Hausbesuch - alle Ziffern
	Hausbesuchskomplexe()
return ;}
:*:c1::                                                                                           	;{ chronisch krank Kennzeichnung - automatisierte Überprüfung
	SendInput, % "{RAW}03220"
	InChronicList()
return
:*:c2::03221
:*:DAK::93550-93555-93560-93565-93570

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
#If
;}
; --- Kuerzelfeld                                                                                   	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("identify"), "Edit3")
:X*:GVU::                                                                                        	;{ GVU,HKS Automatisierung

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
#If
;}
; --- Diagnosen                                                                                 	;{
; bei Includes jetzt
;}
; --- Info                                                                                            	;{
#If ( InStr(AlbisGetActiveControl("contraction"), "info") && WinActive("ahk_class OptoAppClass") )
:R*:letztes::letztes Quartal nicht da (keine Chronikerziffer möglich)
#If
;}
; --- aem                                                                                           	;{
#If WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "aem")
:R*:ENife::Einnahmebeschreibung für Nifedipin-Notfalltropfen mitgegeben
:R*:ETram::Einnahmebeschreibung für Tramadol mitgegeben
:R*:EVitD3::Einnahmebeschreibung für Dekristol mitgegeben
:R*:ETry::Einnahmebeschreibung für Tryasol mitgegeben
:R*:EMac::Einnahmebeschreibung für Macrogol mitgegeben
;}
; --- Rezept                                                                                        	;{
#If WinActive("Muster 16 ahk_class #32770") || WinActive("Privatrezept ahk_class #32770")
:*:#Trya::                                           	;{ Tryasol 30 ml
	IfapVerordnung("Tryasol 30 ml")
return ;}

;}
; --- AUTODOC ---                                                                            	;{
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
; --- AppStarter                                                                                  	;{
:R*:...::
HotstringComboBox(AlbisGetActiveControl("content"))
return
#IfWinActive
;}
; --- Ausnahmeindikationen                                                                 	;{
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
:R*:.aid::                                                                                                                    	;{ % "ahk_id "
	SendInput, % "{Raw}% " q "ahk_id " q " "
return ;}
:R*:.as::A_ScriptDir
:R*:.as1::                                                                                                                     	;{  A_ScriptDir "\"
	SendInput, % "{Raw}A_ScriptDir\"
	SendInput, {Left}
return ;}
:R*:.t1::	                                                                                                                    	;{ oberer einzelner Trenner
SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------"
return ;}
:R*:.t2::	                                                                                                                    	;{ mittlerer Trenner
SendInput, % "{Raw};----------------------------------------------------------------------------------------------------------------------------------------------;{"
return ;}
:R*:.t3::                                                                                                                     	;{ alle 3
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
	SendInput, % "{Raw}MsgBox, 1, Addendum für Albis on Windows, % """""
return ;}
; __ SciteOutPut __
:R*:.sot::
SciteWrite("SciteOutPut(qq<clipText>: qq <clipText>, 0, 1)", true)
return
; __ Tasten __
:R*:.lk::{Left}
:R*:.rk::{Right}
:R*:.cdk::{LControl Down}
:R*:.cuk::{LControl Up}
:R*:.cduk:: ;{
	SendInput, % "{Raw}{LControl Down}{LControl Up}"
	SendInput, {Left 13}
return ;}
:R*:.sdk::{LShift Down}
:R*:.suk::{LShift Up}
:R*:.sduk:: ;{
SendInput, % "{Raw}{LShift Down}{LShift Up}"
SendInput, {Left 11}
return ;}
:R*:.adk::{LAlt Down}
:R*:.auk::{LAlt Up}
:R*:.aduk:: ;{
SendInput, % "{Raw}{LAlt Down}{LAlt Up}"
SendInput, {Left 9}
return ;}
:R*:.adk::{Win Down}
:R*:.auk::{Win Up}
:R*:.aduk:: ;{
	SendInput, % "{Raw}{Win Down}{Win Up}"
	SendInput, {Left 8}
return ;}
:R*:.adh::                                                                                                                    	;{ Hotstring für Autodiagnosen
	SendInput, % "{Raw}:XR*:::"
	Send, {Left 2}
return ;}
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
::Bestellungen::
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
return
;}

;{     	5.a.3.2 -- sonstige Kürzel
::.Heute::`%A_DD`%`.`%A_MM`%`.`%A_YYYY`%
::#AfA::Addendum für Albis on Windows
::#Addendum::Addendum für Albis on Windows
::#Msgbox::MsgBox, 1, Addendum für Albis on Windows -
::#ahk::Autohotkey
::#lahk::language:Autohotkey
::#sahk::site:Autohotkey
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

IsDatumsfeld() {

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

AddToDate(Feld, val, timeunits) {                                                             ; addiert Tage bzw. eine Anzahl von Monaten zu einem Datum hinzu

	calcdate:= SubStr(Feld.Datum, 7, 4) . SubStr(Feld.Datum, 4, 2) . SubStr(Feld.Datum, 1, 2)
	calcdate += val, %timeunits%
	FormatTime, newdate, % calcdate, dd.MM.yyyy
	ControlSetText,, % newdate, % "ahk_id " Feld.hwnd

}

EnableAllControls(ask=1) {

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
	;SendRawFast("{}", "L")
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
SciteHotKeyWriter:          	;{ 	Strg (links) + Alt links + h

		aX:= A_CaretX, aY:= A_CaretY
		ControlGetFocus, scitefocus						, % "ahk_id " scitehwnd:= WinExist("A")
		ControlGet, hScintilla, Hwnd,, Scintilla1  	, % "ahk_id " scitehwnd
		SendMessage, 2008, 0, 0,							, % "ahk_id " hScintilla
		SciteCaretPos:= ErrorLevel
		SendMessage, 2166, % SciteCaretPos, 0,	, % "ahk_id " hScintilla
		SciteCurrentLine:= ErrorLevel
		SendMessage, 4033, 0, 0,							, % "ahk_id " scitehwnd
		SciteCurrentColumn:= ErrorLevel

		Suspend, On

		HKeyW:= 300

		Gui, HKey: new
		Gui, HKey: Font, s16 q5 	, Futura Md Bt
		Gui, HKey: Add, Text			, % "w"HKeyW " Center "							, Hotkeykombination erstellen
		Gui, HKey: Font, s9 q5 		, Futura Bk Bt
		Gui, HKey: Add, Text			, % "w"HKeyW " xs	yp30 Center"				, % "Current position:  Line(" SciteCurrentLine ")  Column(" SciteCurrentColumn ")  Pos(" SciteCaretPos ")"
		Gui, HKey: Font, s10 q5 	, Futura Bk Bt
		Gui, HKey: Add, Hotkey		, % "w"HKeyW " xs 	vchoosedHotkeys"
		Gui, HKey: Add, Button		, xs Section gHkeyB vHKeyOk1				, als Tastennamen
		Gui, HKey: Add, Button		, ys gHkeyB vHKeyOk2							, als Hotkey-Modifikatoren
		Gui, HKey: Add, Button		, xs gHkeyB vHKeyCancel						, Abbruch
		Gui, HKey: Show				, % "x" aX " y" aY+50 " AutoSize"				, Hotkeykombination drücken
		GuiControl, HKey: Focus, choosedHotkeys

		Suspend, Off

return

HkeyB:

		If A_GuiControl = "HKeyCancel"
				goto HkeyBGuiClose

		Gui, HKey: Submit

		If InStr(A_GuiControl, "HKeyOk1")
		{
				b:= choosedHotkeys, b.= SubStr(choosedHotkeys, 1, StrLen(choosedHotkeys)-1)
				keyStat:= "Down"
				b:= RegExReplace(b, "\!(?=.*\w)"					, 		 "{LAlt " keyStat "}",,1)
				b:= RegExReplace(b, "\#(?=.*\w)"					, 		 "{Win " keyStat "}",,1)
				b:= RegExReplace(b, "\^(?=.*\w)"					, "{LControl " keyStat "}",,1)
				b:= RegExReplace(b, "\+(?=.*\w)"					, 	   "{LShift " keyStat "}",,1)
				keyStat:= "Up"
				b:= RegExReplace(b, "(?<=\w|\w.|\w$)\!"		, 		 "{LAlt " keyStat "}",,1)
				b:= RegExReplace(b, "(?<=\w|\w.|\w$)\#"	, 		 "{Win " keyStat "}",,1)
				b:= RegExReplace(b, "(?<=\w|\w.|\w$)\^"	, "{LControl " keyStat "}",,1)
				b:= RegExReplace(b, "(?<=\w|\w.|\w$)\+"	, 	   "{LShift " keyStat "}",,1)
		}
		else if InStr(A_GuiControl, "HKeyOk2")
		{
				b:= choosedHotkeys
		}

		PraxTT(b, "0 3")

		WinActivate		, 	% "ahk_id " scitehwnd
		WinWaitActive	, 	% "ahk_id " scitehwnd,, 2
		ControlFocus		, 	% "ahk_id " hScintilla
			Sleep, 100
		ControlClick		,, 	% "ahk_id " hScintilla
			Sleep, 100

	  ; SCI_GOTOPOS
		SendMessage, 2025, % SciteCaretPos, 0,	, % "ahk_id " hscintilla
		SendRaw, % b


HkeyBGuiClose:
HkeyBGuiEscape:

	Gui, HKey: Destroy

return
;}

;}
;{ ###############                 ALBIS            	###############
StartePraxomat:             	;{ 	Strg	+	ä
	If WinExist("ahk_class OptoAppClass")
        	Run, %AddendumDir%\Module\Albis_Funktionen\Praxomat.ahk                                                        ;-- Strg + eine beliebige sonstige Taste (Win, Alt, Shift) + ä -> startet Praxomat
return
;}
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
GVUListeAnzeigen:        	;{
	;gosub GVU_GUI
return ;}
KarteikartenMenu:          	;{
	MouseGetPos, mx, my
	; Context menu opened –> Get handle:
	MN_GETHMENU := 0x1E1 ; Shell Constant: "Menu_GetHandleMenu"
	SendMessage, MN_GETHMENU, False, False
	hCM := ErrorLevel ; Return Handle in ErrorLevel
	ToolTip, % "Hab Dich Menu: " GetHex(hCM) , % mx , % my - 60, 6
	SetTimer, HDM, -3000
return
Fct_Get_ClickedItem_MenuContext_AddedItems( TargetProgram_Classes, InsertItemsAt_Position, ArrayOf_Items* ) {
	; Get handle of the target program:
	For each, ProgramClass in TargetProgram_Classes
	{
		; Handle of active window:
		If Program_Handle := WinActive( "ahk_class " ProgramClass ) ; Handle return false if class it's not active
			Break ; Program is active, then OK
	} ; Or not if the loop ( for each ) is finished

	; Check we are really in the right active program:
	If not Program_Handle ; False or Handle
		Return "-> Not into the right program <-"

	; Class #32768 is for all standard windows context menu:
	Global MenuContext_AhkClass := "ahk_class #32768"

	; Wait context menu appears:
	WinWait, %MenuContext_AhkClass% ;

	; Context menu opened –> Get handle:
	MN_GETHMENU := 0x1E1 ; Shell Constant: "Menu_GetHandleMenu"
	SendMessage, MN_GETHMENU, False, False
	MenuContext_Handle := ErrorLevel ; Return Handle in ErrorLevel


	;***************************
	; 	Section: Add Items:
	;***************************

	; Constants for menus in User32.dll:
	Static MF_SEPARATOR := 0x800
	Static MF_STRING := 0x0
	Static MF_POPUP := 0x10
	Static MF_BYPOSITION := 0x400

	; Add each new item into the context menu:
	For each, ItemToAdd in ArrayToAdd_Items :=  ArrayOf_Items
	{
		; Save absolut position of this Item in the menu:
		ItemToAdd.Position := InsertItemsAt_Position-1 + A_Index-1 ; Zero based

		; Add Separator:
		If ( ItemToAdd.Type == MenuContext_Separator )
		{
			; Insert Separator: –> https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-InsertMenuA
			DllCall( "User32\InsertMenu"
					,"UPtr", MenuContext_Handle
					,"UInt", ItemToAdd.Position ; At the specified position
					,"UInt", MF_SEPARATOR + MF_BYPOSITION
					,"UPtr", False
					,"UInt", False )
		}

		; Add Classic Item:
		Else If ( ItemToAdd.Type == MenuContext_Item )
		{
			; Insert text of item: –> https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-InsertMenuA
			DllCall( "User32\InsertMenu"
					,"UPtr", MenuContext_Handle
					,"UInt", ItemToAdd.Position ; At the specified position
					,"UInt", MF_STRING + MF_BYPOSITION
					,"UPtr", False
					,"Str", ItemToAdd.Name ) ; Insert Value ( text )
		}

		; Add Submenu and its Subitems:
		Else If ( ItemToAdd.Type == MenuContext_Submenu )
		{
			AddSubmenu := ItemToAdd ; Renames to clarify
			; Create Submenu, and return handle:
			AddSubmenu.Handle := DllCall( "User32\CreatePopupMenu" )

			; Insert Submenu into the context menu: –> https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-InsertMenuA
			DllCall( "User32\InsertMenu"
					,"UPtr", MenuContext_Handle
					,"UInt", AddSubmenu.Position  ; At the specified position
					,"UInt", MF_STRING + MF_BYPOSITION + MF_POPUP
					,"UPtr", AddSubmenu.Handle
					,"Str", AddSubmenu.Name )

			; Now add each Item and Separator into this Submenu:
			For each, ItemOfSubmenu in AddSubmenu.Items
			{
				; AppendMenu –> https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-AppendMenuA

				; In case of Separator:
				If ( ItemOfSubmenu == MenuContext_Separator )
					DllCall( "User32\AppendMenu"
							,"UPtr", AddSubmenu.Handle
							,"UInt", MF_SEPARATOR
							,"UPtr", False, "UInt", False )

				Else  ; In case of Subintem: insert Value as text:
					DllCall( "User32\AppendMenu"
							,"UPtr", AddSubmenu.Handle
							,"UInt", MF_STRING
							,"UPtr", False
							,"Str", ItemOfSubmenu )
			}
		}
		Continue ; Continue to add next item...
	}
	; End for each add items


	;***************************
	; 	Section: Wait Click:
	;***************************

	Label_Wait_New_Click: ; When User clicks on Separator: wait new click

	; Wait User do a regular or right click:
	While not GetKeyState( "LButton" ) and not GetKeyState( "RButton" )
	{
		; Fix error if 2 right clicks are very too close: menu disappears sometimes for some apps:
		If not WinExist( MenuContext_AhkClass ) {
			Send {RButton} 	; Reopen context menu when disappears
			Return MenuContext_RightClickAgain ; Refill the menu
		}
	}

	; Is it a right click on another file ? This means that the menu is closed and reopening...
	If GetKeyState( "RButton" )
		Return MenuContext_RightClickAgain ; Then refill the menu
	;TODO: Fix error when the User right-clicks on context menu, it's filled up again

	; Else Yes: Click into an item, but witch item ?


	;***********************************
	; 	Section: Get Clicked Item:
	;***********************************

	; Get position of mouse into screen:
	CoordMode, Mouse, Screen
	MouseGetPos, MouseScreenX, MouseScreenY  ; Int vars: 4 octets

	; POINT –> https://docs.microsoft.com/fr-fr/previous-versions/dd162805(v=vs.85)
	; Create a generic C++ POINT{x,y} with a 'ULongLong' –> 'Int64'
	; X start at the 1st ULong (right), and Y start at 2nd ULong (32th bit on left):
	MousePointScreen := x := MouseScreenX  | y := ( MouseScreenY << 32 )

	;Calculate DPI of special screen:  1K, 2K, 4K, 8K, etc...
	WinDPIMultiplicator := A_ScreenDPI / 96  ; 96 is the standard DPI screen: 1K (1600x900)

	; Check if clicked item is into the new added items:
	For each, AddedItemInMenu in ArrayOfAdded_Items := ArrayOf_Items
	{
		; Click on Separator –> Disable:
		If ( AddedItemInMenu.Type == MenuContext_Separator )
		{
			; Get Rectangle of Separator:
			VarSetCapacity( ItemRectangle, 16, 0 ) ; Create Rectangle of 16 octets: 4 corners of Int (4 octets)

			; Fill Rectangle: –> https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-GetMenuItemRect
			isFilledRectangle := DLLCall( "User32\GetMenuItemRect"
									,"UPtr", Program_Handle
									,"UPtr", MenuContext_Handle
									,"UInt", AddedItemInMenu.Position ; Absolut position in the context menu
									,"UPtr", &ItemRectangle )
			; Is clicked on separator ?
			If isFilledRectangle
			and isPointIntoRectangle := DllCall( "User32\PtInRect", "UPtr", &ItemRectangle, "Int64", MousePointScreen )
				Goto Label_Wait_New_Click
		}

		; Click on added Item –> Return Value
		Else If ( AddedItemInMenu.Type == MenuContext_Item )
		{
			VarSetCapacity( ItemRectangle, 16, 0 )
			; https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-GetMenuItemRect
			isFilledRectangle := DLLCall( "User32\GetMenuItemRect"
									,"UPtr", Program_Handle
									,"UPtr", MenuContext_Handle
									,"UInt", AddedItemInMenu.Position
									,"UPtr", &ItemRectangle )
			If isFilledRectangle
			and isPointIntoRectangle := DllCall( "User32\PtInRect", "UPtr", &ItemRectangle, "Int64", MousePointScreen )
				Return AddedItemInMenu.Name
		}

		; Click on Item of Submenu –> Check each Subitem:
		Else If ( AddedItemInMenu.Type == MenuContext_Submenu )
		{
			Loop 3 ; 3 times, because sometimes this function does not work 1 or 2 times:
				For each, ItemInSubmenu in ( ItemIsSubmenu := AddedItemInMenu ).Items
				{
					ItemPositionInMenu := A_Index

					; Get Rectangle for Subitem
					VarSetCapacity( ItemRectangle, 16, 0 )
					; https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-GetMenuItemRect
					isFilledRectangle := DLLCall( "User32\GetMenuItemRect"
											,"UPtr", False ; Indicate Submenu, instead of window
											,"UPtr", ItemIsSubmenu.Handle ; Handle of the Submenu
											,"UInt", ItemPositionInMenu -1 ; Zero based
											,"UPtr", &ItemRectangle )

					; Get each corner of Rectangle:
					ItemRectangleX1 := NumGet( &ItemRectangle, 0, "Int" ) ; Int –> 4 octets
					ItemRectangleY1 := NumGet( &ItemRectangle, 4, "Int" )
					ItemRectangleX2 := NumGet( &ItemRectangle, 8, "Int" )
					ItemRectangleY2 := NumGet( &ItemRectangle, 12, "Int" )

					; Use DPI multiplicator for special screen (2K, 4K, ...):
					isMouseInto4Corners :=  ( MouseScreenX >= ItemRectangleX1 *WinDPIMultiplicator )
									and ( MouseScreenX <= ItemRectangleX2 *WinDPIMultiplicator )
									and ( MouseScreenY >= ItemRectangleY1 *WinDPIMultiplicator )
									and ( MouseScreenY <= ItemRectangleY2 *WinDPIMultiplicator )

					;ToolTip % "isFilledRectangle: " isFilledRectangle "`n" "isMouseInto4Corners: " isMouseInto4Corners "`nX: " ItemRectangleX1*WinDPIMultiplicator "<" MouseScreenX ">" ItemRectangleX2*WinDPIMultiplicator "`nY: " ItemRectangleY1*WinDPIMultiplicator "<" MouseScreenY ">" ItemRectangleY2*WinDPIMultiplicator, , , 5

					; Check if mouse is into this Rectangle:
					If isFilledRectangle and isMouseInto4Corners
					{
						; Click on Separator ?
						If ( ItemInSubmenu == MenuContext_Separator ) ; Disable Separator:
							Goto Label_Wait_New_Click ; Wait new click

						Else ; Standard text Subitem:
							Return ItemInSubmenu ; Return value of item
					}
				}
		}
		; End of one item: this is not this item
		Continue ; Then check next item...
	}
	; End: For all the new items, all checked without return

	; Then clicked Item is probably on another item of context menu ( like 'Open' ):
	Return MenuContext_NoClicked_AddedItem
}

HDM:
	ToolTip,,,,6
return
;}
;}
;{ ###############	          ALLGEMEIN        	###############
CaptureDoubleClick:       	;{
	if (A_PriorHotkey = A_ThisHotkey && A_TimeSincePriorHotkey < 500) {
			SendInput, {LControl Down}
    }
return
;}
KopierenUeberall:         	;{	 	Strg	+ c
	CopyToClipboard:= 1
AusschneidenUeberall:   	;	 	Strg (links) + x
	SetKeyDelay, 20, 50

	Clipboard := ""
	while clip = ""
	{
			SendEvent, {CTRL down}c{CTRL up}
			ClipWait, 1
			clip:= Clipboard
			If A_Index > 6
					break
	}

	If clip != ""
	{
		clip:= SubStr(ClipBoard, 1, 30) . "`.`.`."
		GDISplash(clip,1)
	}
	else
		GDISplash("Clipboard war`nleer!",1)

	If !CopyToClipboard
			Send, {CTRL down}x{CTRL up}

	CopyToClipboard	:= 0
	clip						:= ""
	SetKeyDelay, 10, 20
return
;}
ShowAll:                        	;{    	zeigt alle Fenster wieder an
WinShow, ahk_group HiddenWindows
PostMessage, 0x111, 29698, , SHELLDLL_DefView1, ahk_class Progman
return
;}
HideAll:                        	;{    	blendet alle Fenster aus
HideAllWindows(STEnd)
HAW:=1
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
AddICDHotstring:          	;{ 	Strg + Win + Alt + i
	;SciteAutoDiagnose(Clipboard)
return ;}
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

	SendInput, % "{Raw}" TextToSend
	SetTimer, emptyTextToSend, -180000 ; nach 3min leeren

return
emptyTextToSend:
	TextToSend := ""
return
;}
;}

;}

;{11. Alle - Labels - Clientunabhängige Ausführung

ACPatientOeffnen:                         	;{

		If !WinExist("ahk_id " . FcsWinHwnd) {
				FcsWinHwnd:=0
				SetTimer, ACPatientOeffnen, Off
				return
		}
		IfWinNotActive, ahk_id %FcsWinHwnd%
				return

		ControlGetText, CText, Edit1, ahk_id %POHwnd%
		If !(CText = CText_old) {
			CText_old:= CTText
			;PatList:= AutoCompleteList(CText)
			AddAutoComplete(POHwnd, "Edit1", PatList, 400)
		}

return
;}
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
Hotkeys:                                         	;{                                                                         	;-- wegen Fehler deaktiviert

	;--Anzeige aller benutzten Hotkeys mit Beschreibung durch auslesen des Skriptes
	;script		:= A_ScriptDir . "\" . A_ScriptName
	;activeApp := ActiveContext()

	; http://www.autohotkey.com/forum/topic215.html

	;ViewKeyList()




return

;~ HumanReadableHK( key ) {
;~ ;a function to take care of replacing symbols in hotkeys with their words
  ;~ key:= 	 StrReplace(key, "+"			, "Shift + "			)
  ;~ key:=  RegExReplace(key, "#(?!=\s)", "Win + "			)
  ;~ key:= 	 StrReplace(key, "!"				, "Alt links + "	)
  ;~ key:= 	 StrReplace(key, "LAlt"			, "Alt links + "	)
  ;~ key:= 	 StrReplace(key, "^"			, "Strg + "			)
  ;~ key:= 	 StrReplace(key, "LControl"	, "Strg + "			)
  ;~ key:= RegExReplace(key, "\s\&\s"	, " + ")
  ;~ key:= RegExReplace(key, "\*(?=\^|\!|\+|\#|\$|\~|\w)"	, "")
  ;~ key:= RegExReplace(key, "\$(?=\^|\!|\+|\#|\*|\~|\w)"	, "")
  ;~ key:= RegExReplace(key, "\~(?=\^|\!|\+|\#|\$|\*|\w)"	, "")
  ;~ return key
;~ }

;~ ViewKeyList() {

	;~ KeyList:= Object()
	;~ KeyList:= FormatKeyList(A_ScriptFullPath)

	;~ activeContext:= activeContext()

	;~ For idx, key in KeyList
		;~ If !KeyList.HasKey(key.cat)
			;~ KeyList[Key.cat] := 1
		;~ else
			;~ KeyList[Key.cat] += 1

	;~ ;MsgBox, % KeyList["Albis"] "`n" KeyList["Überall"] "`n" KeyList["SciTE"] "`n" KeyList[1]["cat"] ":" KeyList[1]["hk"] "/" KeyList[1]["des"] "/" KeyList[1]["test"]

	;~ ;Gui, KL: New
	;~ ;Gui, KL: Add, ListView,
;~ }

;~ FormatKeyList(file, srt:= False) {                                        	;this does the actual formatting. Pass the text to be formatted

	;~ fileread, scriptlines, % file
	;~ f:= FileOpen(A_ScriptDir "\Hotkey.txt", "w")
	;~ KeyList := Object(), idx:= 0

	;~ Loop, Parse, scriptlines, `n, `r
	;~ {
			;~ RegExMatch(A_LoopField, "i)(?<=\sHotkey\,).*(?=\s\,\s\w)", hk)
			;~ if hk
			;~ {
				;~ RegExMatch(A_LoopField, "i)(?<=\s\;\=\s).*(?=\:\s)", cat)
				;~ RegExMatch(A_LoopField, "i)(?<=\s\;\=\s).*", des)
				;~ hk:= Trim(hk), des:= Trim(des), cat:= Trim(cat)
			;~ }

			;~ If cat && des && hk
			;~ {
				;~ idx ++
				;~ keyname		:= HumanReadableHK(Trim(hk))
				;~ description	:= StrReplace(des				, cat	, "")
				;~ description  	:= StrReplace(description	, "`:"	, "")
				;~ f.Write(A_Index ": Nr." idx " - " cat "|" keyname "|" description "`n")
				;~ KeyList[(idx)]	:= {"cat": cat, "hk": keyname, "des": description, "test": "warum"}
			;~ }

			;~ hk:=des:=cat:=""
	;~ }

	;~ f.Close()

;~ return KeyList
;~ }

;~ activeContext() {                                        							;-- gibt einen Identifikator für das aktive Fenster zurück

	;~ If InStr(activeClass	:= WinGetClass(WinExist("A")), "OptoAppClass")
					;~ return "Albis"
	;~ else if InStr(activeClass, "Scite")
					;~ return "Scite"
	;~ else if InStr(activeClass, "Qt5QWindowIcon") and InStr(WinGetTitle(WinExist("A")), "Telegram")
					;~ return "Telegram"

;~ }

;}
SkriptReload:                                 	;{                                                            					;-- besseres Reload eines Skriptes

  ; es wird ein RestartScript gestartet, welches zuerst das laufende Skript komplett beendet und dann das Addendum-Skript per run Befehl erneut startet
  ; dies verhindert einige Probleme mit datumsabhängigen Handles bestimmter Prozesse im System, die Hook und Hotkey Routinen sind zuverlässiger dadurch geworden

	Script    	:= Addendum.AddendumDir "\Module\Addendum\Addendum.ahk"
	scriptPID 	:= DllCall("GetCurrentProcessId")
	__          	:= q " " q                                             	; nur für die Lesbarkeit des Codes in der unteren Zeile

	cmdline := "Autohotkey.exe /f " q Addendum.AddendumDir "\include\Addendum_Reload.ahk" __ Script __ "1" __ A_ScriptHwnd __ scriptPID __ " 1" q
	Run, % cmdline
	ExitApp

return ;}

ShowAddendumObjekt() {

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
ZeigeOnExitProtokoll() {                                                                                                  	;-- Protokoll der Skriptbeendigungen anzeigen
	Run % Addendum.AddendumDir "\logs'n'data\OnExit-Protokoll.txt"
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
	RunSkript(modul[A_ThisMenuItem])
return ;}
ToolStarter:                                   	;{                                                                            	;-- startet Tools vom Tray-Menu
	RunSkript(tool[A_ThisMenuItem])
return ;}

; ----------------------------------------- Einstellungen
Menu_PauseAddendum:                  	;{
	Pause
return ;}
Menu_AlbisAutoPosition:                   	;{
	Addendum.AlbisLocationChange := !Addendum.AlbisLocationChange
	Menu, SubMenu3, ToggleCheck, Albis AutoSize
	IniWrite, % (Addendum.AlbisLocationChange ? "1":"0")	, % Addendum.AddendumIni, % compname, % "Albis_AutoGroesse"
return
;}
Menu_AddendumGui:                    	;{
	Addendum.AddendumGui := !Addendum.AddendumGui
	Menu, SubMenu3, ToggleCheck, Addendum Infofenster
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
Menu_MSWordAutoPosition:             	;{
	Addendum.WordAutoSize:= !Addendum.WordAutoSize
	Menu, SubMenu3, ToggleCheck, Microsoft Word AutoSize
	IniWrite, % (Addendum.WordAutoSize ? "1":"0")           	, % Addendum.AddendumIni, % compname, % "Microsoft_Word_AutoGroesse"
return
;}
Menu_GVUAutomation:                    	;{
	Addendum.GVUAutomation := !Addendum.GVUAutomation
	Menu, SubMenu3, ToggleCheck, Albis GVU automatisieren
	IniWrite, % (Addendum.GVUAutomation ? "1":"0")        	, % Addendum.AddendumIni, % compname, % "GVU_automatisieren"
return
;}
Menu_PDFSignatureAutomation:      	;{

	; Menupunkt umschalten, Einstellung speichern
		Addendum.PDFSignieren := !Addendum.PDFSignieren
		Menu, SubMenu3, ToggleCheck, FoxitReader Signaturhilfe
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
	Menu, SubMenu3, ToggleCheck, FoxitReader Dokument automatisch schliessen
	IniWrite, % (Addendum.AutoCloseFoxitTab ? "ja":"nein") 	, % Addendum.AddendumIni, % compname, % "FoxitTabSchliessenNachSignierung"
return ;}
Menu_LaborAbrufManuell:               	;{
	Addendum.Labor.AutoAbruf := !Addendum.Labor.AutoAbruf
	Menu, SubMenu3, ToggleCheck, Laborabruf automatisieren
	IniWrite, % (Addendum.Labor.AutoAbruf ? "1":"0")        	, % Addendum.AddendumIni, % compname, % "Laborabruf_automatisieren"
return ;}
Menu_LaborAbrufTimer:                  	;{
	If RegExMatch(Addendum.Labor.AbrufTimer, "Nein")
		return
	Addendum.Labor.AbrufTimer := !Addendum.Labor.AbrufTimer
	Menu, SubMenu3, ToggleCheck, zeitgesteuerter Laborabruf
	IniWrite, % (Addendum.Labor.AbrufTimer ? "1":"0")        	, % Addendum.AddendumIni, % compname, % "Laborabruf_Timer"
	IniWrite, % "1"                                                           	, % Addendum.AddendumIni, % compname, % "Laborabruf_automatisieren" ; dieser muss hier immer "An" sein!
return ;}
;}

;----------------------------------------------------------------------------------------------------------------------------------------------

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
	SP:= Object()

	Lmx:= A_CaretX, Lmy:= A_CaretY
	If !Lmx or !Lmy
		MouseGetPos, Lmx, Lmy

	Loop, Parse, string, `n
			Lc:= A_Index
	Lc:= Lc>4 ? 4:Lc

		;building gui
	Gui sp: New, +HWNDhsp AlwaysOnTop -SysMenu -Caption +ToolWindow
	Gui sp: Font, s10 q5 cWhite, Futura Bk Md
	Gui sp: Color, c3F627F
	WinSet, Transparent, 120, ahk_id %hsp%
	Gui sp: Margin, 8, 5
	Gui sp: Add, Text, % "xm ym R" (Lc) " HWNDhSplashText" , % string
	;GuiControl, sp: Text, %hSplashText%, % string
	Gui sp: Show, % "NA x" (Lmx) " y" (Lmy - 30 - ((Lc-1)*10) ) , Addendum Splash

		;form a rounded rectangle
	SP:= GetWindowInfo(hsp)
	WinSet, Region, % "0-0 W" (SP.windowW) " H" (SP.windowH) " R20-20", ahk_id %hsp%

		;close with delay
	SetTimer, SPClosePopup, % "-" (time*1000)

return

SPClosePopup:
      Gui sp: Destroy
return

}

GDISplash(text, time:=2) {                                                                                          	;-- Hinweisfenster

		ListLines, Off

		static hGdiSp, GdiSp
		static Options:= "y5 Centre cff000000 q5 s16"
		static Font 	:= "Futura Bk Bt"
		static hFont

		If !hFont
			hFont 	:= CreateFont(Options "`, " Font)

		Gui, GdiSp: Destroy

		cx:= A_CaretX , cy:= A_CaretY - 35
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
		;WinSet, Redraw, ahk_id %hGdiSp%

		SetTimer, GDISplashOff, % "-" (time*1000)
		return hGdiSp

GDISplashOff:

		Gui, GdiSp: Destroy

return

}

Tooltip(content, wait:=2) {
	tooltip,% content,% A_CaretX,% A_CaretY, 11
    SetTimer,tOff,% "-" wait*1000
    return
    tOff:
    tooltip,,,,11
    return
}


;Text senden
SendRawFast(string, cmd="") {                                                                                   	;-- sendet eine String ohne das es zum Senden von Hotkey's kommen kann

	ListLines, Off
	;Parameter: go - L, R, U, D - stehen für Left, Right, Up, Down, nach diesen kann

	Splash(string . " " .  cmd, 3)

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
		If alternating[A_Index] = chars
		{
			SendRaw, % StrSplit(chars, "|").2
			alternating.RemoveAt(A_Index)
			return
		}

	SendRaw, % StrSplit(chars, "|").1
	alternating.Push(chars)

}

SciTEFoldAll() {                                                                                                            	;-- gesamten Code zusammenfalten

	SCI_FOLDALL := 2662
	ControlGet, hScintilla1, hwnd,, Scintilla1, % "ahk_class SciTEWindow"
	SendMessage, % SCI_FOLDALL,,,, % "ahk_id " hScintilla1
	TrayTip("SciTEFoldAll", "ErrorLevel: " ErrorLevel, 2 )
	;SendMessage, 2662, 0, 0,, % "ahk_id " ControlGet("hwnd", "", "Scintilla1", "ahk_class SciTEWindow") ;

}

RunSkript(modulpath) {                                                                                              	;-- startet oder beendet Autohotkeyskripte

	AddendumDir := Addendum.AddendumDir
	SplitPath, modulpath, modulname
	modulname := StrReplace(modulname, ".ahk", "")

	If (PID:= ScriptExist(modulname))  {
		MsgBox, 0x1004, Addendum für Albis on Windows, % "Das Modul '" tostartAHK "' ist schon gestartet!`nMöchten Sie das Modul stattdessen beenden?"
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
                           	MsgBox, 1, Addendum für Albis on Windows, % "Das Modul '" tostartAHK "' konnte nicht beendet werden!"
				}
		}
		return
	}


	If !RegExMatch(modulpath, "^[A-Z]\:\\")
		run, Autohotkey.exe "%AddendumDir%\%modulpath%",, ModulPID
	else
		run, Autohotkey.exe "%modulpath%",, ModulPID

return ModulPID
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

GVU_GUI() {

	; Variablen
		;global GVU, hGVu, hLV_GVU, LBQuartale, GVULV, LVCount
		global 	GVUTextProgress, GVUProgress, GVU_LV, GVURun, hGVU_LV, GVUInfo, ZText
		;local 	LBQuartale
		Quartale     	:= ""
		GVUListe   	:= Object()
		aktuelleListe	:= IniReadExt("GVUListe", "aktuelles_GVUFile")

	; Inhalt vorhandener Dateien in ein Objekt einlesen
		Loop, Files, % AddendumDir "\TagesProtokolle\*-GVU.txt"
		{
				SplitPath, A_LoopFileFullPath,,,, filename
				Quartal:= StrReplace(filename, "-GVU")
				Lnr:= A_Index
				GVUListe[Lnr]:= {"Quartal": Quartal, "Liste":[]}
				Quartale.= Quartal "|"
				FileRead, file, % A_LoopFileFullPath
				Loop, Parse, file, `n, `r
				{
					If StrLen(A_LoopField) > 0
						GVUListe[Lnr].Liste.Push(A_LoopField)
				}
		}

		Quartale        	:= RTrim(Quartale, "|")
		GVUPrg_Width	:= 680
		GVULV_Width	:= 800
		GVULVOptions	:= "xm w" GVULV_Width " r30 BackgroundAAAAFF Grid NoSort"

		Gui, GGVU: New  	, +AlwaysOnTop +HwndhGVU -DPIScale
		Gui, GGVU: Font  	, S10 Normal q5, % Addendum.StandardFont
		Gui, GGVU: Margin	, 5, 5
		Gui, GGVU: Add   	, Text    	, xm ym+3                                                                               	, % "angezeigtes Quartal:"
		Gui, GGVU: Add    	, ListBox	, x+5 ym+3 r1 gGVU_LBEvent vLBQuartale                              	, % Quartale
		Gui, GGVU: Add   	, Button  	, x+5 ym w200 gGVU_Ablaufstarten vGVU_Run +0x00000300  	, % "Formularerstellung starten"
		Gui, GGVU: Font  	, s8 Normal q5, % Addendum.StandardFont
		Gui, GGVU: Add   	, Text     	, x+10 y10 w250 vZText                                                          	, % "Vorsorgen unbearbeitet: " GVUListe[Lnr].Liste.MaxIndex()
		Gui, GGVU: Font  	, S10 Normal q5, % Addendum.StandardFont
		Gui, GGVU: Add   	, Progress	, % "xm w" GVUPrg_Width " h20 CBLime BackgroundEEEEEE cBlue vGVUProgress", % "0"
		Gui, GGVU: Add   	, Text    	, % "x+5 w" (GVULV_Width - GVUPrg_Width - 5) " vGVUTextProgress"                                                 	, % ""
		Gui, GGVU: Add   	, Listview	, % GVULVOptions " vGVU_LV HWNDhGVU_LV"                    	, % "U-Nr|bearbeitet|PatientenID|Name|Geburtsdatum|Untersuchungsdatum|Hinweis"
		GuiControl, GGVU: ChooseString, LBQuartale, % aktuelleListe

		gosub GVU_LVFuellen

		Gui, GGVU: Show 	, xCenter yCenter AutoSize   	, % "Anzeige der Patienten in der GVU Liste"


return

GGVUGuiEscape:     	;{
GGVUGuiClose:
	If InStr(Running, "Vorsorgeautomatisierung")
	{
			Running:= "Vorsorgeautomatisierung beendet"
			GuiControl, GGVU: Text, GVUInfo, % "Einen Augenblick bitte ...`n`nDas Fenster wird nach`nvollständiger Formularerstellung`ngeschlossen."
			GuiControl, GGVU: Show, GVUInfo
			while, !StrLen(Running) = 0
			{
					Sleep, 20
					If A_Index > 3000 ; 1min
						break
			}
	}
	Gui, GGVU: Destroy
return ;}

GVUListView:            	;{

return ;}

GVU_LBEvent:           	;{
	Gui, GGVU: Default
	LV_Delete()
	gosub GVU_LVFuellen
return ;}

GVU_LVFuellen:        	;{

	Gui, GGVU: Default
	Gui, GGVU: Submit, NoHide

	For Lnr, data in GVUListe
		If InStr(data.Quartal, LBQuartale)
			break

	Loop, % GVUListe[Lnr].Liste.MaxIndex()
	{
			arr:= StrSplit(GVUListe[Lnr].Liste[A_Index], ";")
			If StrLen(arr[4])>0
				bearbeitet:= "ja"
			else
				bearbeitet:= "nein"
			LV_Add("", A_Index, bearbeitet, arr[3], oPat[Arr[3]].Nn ", " oPat[Arr[3]].Vn, oPat[Arr[3]].Gd, arr[1])
	}

	LV_ModifyCol()
	LV_ModifyCol(7, 300)
	LV_ModifyCol(2, "SortDesc")

return ;}

GVU_Ablaufstarten:   	;{

	Gui, GGVU: Default
	GuiControlGet, GVURunStatus, GGVU:, GVURun, Text
	If InStr(GVURunStatus, "Starten")
	{
			Gui, GGVU: Default
			GuiControl , GGVU: Text, GVURun, % "Formularerstellung pausieren"
		; unbearbeitete Sammeln und einer Funktion übergeben
			IDListe 	:= Object()
			IDIndex	:= 0
			PraxTT("Formularlauf ist gestartet!", "5 3")

			Gui, GGVU: Default
			Gui, GGVU: ListView, GVU_LV
			Loop, % LV_GetCount()
			{
					LV_GetText(bearbeitet, A_Index, 2)
					If InStr(bearbeitet, "nein")
					{
							IDIndex ++
							LV_GetText(UNr	    	, A_Index, 1)
							LV_GetText(ID         	, A_Index, 3)
							LV_GetText(Gd         	, A_Index, 5)
							LV_GetText(UDatum	, A_Index, 6)
							IDListe[IDIndex] := Object()
							IDListe[IDIndex] := {"Unr": A_Index, "PatientenID": ID, "Geburtsdatum": Gd, "UDatum": UDatum}
					}
			}

			If IDIndex = 0
			{
					MsgBox, Es sind alle Untersuchungen angelegt und abgerechnet worden.
					return
			}

			GVU_Automat(IDListe)
	}
	else If InStr(GVURunStatus, "pausieren")
	{
			Gui, GGVU: Default
			GuiControl, GGVU: Text, GVURun, % "Formularerstellung fortsetzen"
			GuiControl, GGVU: Show, GVUInfo
			Running:= "Vorsorgeautomatisierung pausiert"
	}
	else If InStr(GVURunStatus, "fortsetzen")
	{
			Gui, GGVU: Default
			GuiControl, GGVU: Text, GVURun, % "Formularerstellung pausieren"
			GuiControl, GGVU: Hide, GVUInfo
			Running:= "Vorsorgeautomatisierung"
	}

return ;}

}

GVU_Automat(IDListe) {

	Running:= "Vorsorgeautomatisierung"
	z:= []
	PatToDo:= IDListe.MaxIndex()
	lenToDo:= StrLen(PatToDo)
	Sleep, 3000

	Gui, GGVU: Default
	Gui, GGVU: ListView, GVU_LV
	GuiControl, GGVU:, GVUTextProgress, % SubStr("0000", -1 * lenToDo) "/" PatToDo

	For Index, data in IDListe
	{
			PatID:= data.PatientenID
			z[1]:= "akt. Patient     : " oPat[PatID].Nn ", " oPat[PatID].Vn ", geb. am " oPat[PatID].Gd
			z[2]:= "Health-Status  : " Running
			z[3]:= "A_Index          : " A_Index
			t:=""
			Loop % z.MaxIndex()
				t.= z[A_Index] "`n"
			Gui, GGVU: Default
			GuiControl, GGVU: Text, ZText, % t

			; Überprüfen des Patientenalters und Abbruch falls notwendig
				age:= Floor(DateDiff( "YY", data.Geburtsdatum, data.UDatum))
				If (GVU[minAlter] > age)
				{
						PraxTT("Patientenalter: " age ", Mindestalter: " GVU[minAlter] "`nPatient unterschreitet das Mindesalter", "5 3")
						GVU_DatenSichern(data.PatientenID, data.UDatum, ">>>Das Alter des Patienten unterschreitet das Mindestalter<<<", "")
						Sleep, 4000
					; GUI Anzeige Fortschritt zeigen
						Gui, GGVU: Default
						Gui, GGVU: ListView, GVU_LV
						Loop, % LV_GetCount()
						{
								LV_GetText(ID, A_Index, 3)
								If ID = data.PatientenID
								{
										LV_Modify(A_Index,,, "ja")
										break
								}
						}
						Sleep, 200
						LV_ModifyCol(2, "SortDesc")
						GuiControl, GGVU:, GVUTextProgress, % SubStr("0000" Index, -1 * lenToDo) "/" PatToDo
						GuiControl, GGVU:, GVUProgress, % Floor(Index*100/PatToDo)
						continue
				}

				If InStr(Running, "Vorsorgeautomatisierung pausiert")
						GVU_Pause()
				else If InStr(Running, "Vorsorgeautomatisierung beendet")
						break

				AlbisAkteOeffnen(data.PatientenID, data.PatientenID)
				DatumZuvor:= AlbisSetzeProgrammDatum(data.UDatum)

				If InStr(Running, "Vorsorgeautomatisierung pausiert")
						GVU_Pause()

			; Albis GVU Makro starten (über ein Kürzel mit dem Namen GVU wird erst das GVU- dann das HKS Formular aufgerufen und automatisch die Ziffern eingetragen)
				AlbisKarteikartenFocusSetzen("Edit2")
				VerifiedSetText("Edit2", data.UDatum, "ahk_class OptoAppClass")
				SendInput, {Tab}
				Sleep, 100
				ControlFocus, Edit3, ahk_class OptoAppClass
				VerifiedSetText("Edit3", "GVU", "ahk_class OptoAppClass")
				SendInput, {Tab}

			; wartet auf Abschluß der gesamten Prozedur
				WinGetTitle, AlbisT1, ahk_class OptoAppClass
				while, !StrLen(DatumZuvor) = 0
				{
						If GetKeyState("Escape")
							return
						Sleep, 150
						WinGetTitle, AlbisT2, ahk_class OptoAppClass
						If AlbisT1 <> AlbisT2
							break
				}

			; GUI Anzeige Fortschritt zeigen
				Gui, GGVU: Default
				Gui, GGVU: ListView, GVU_LV
				LV_Modify(data.UNr,,, "ja")
				Sleep, 50
				LV_ModifyCol(2, "SortDesc")
				GuiControl, GGVU:, GVUTextProgress, % SubStr("0000" Index, -1 * lenToDo) "/" PatToDo
				GuiControl, GGVU:, GVUProgress, % Floor(Index*100/PatToDo)

				If InStr(Running, "Vorsorgeautomatisierung pausiert")
						GVU_Pause()

			; Fortfahren
				MsgBox, 4, Addendum, Möchten Sie mit dem nächsten Patienten fortfahren?
				IfMsgBox, No
					break

			; Akte schließen
				WinGetTitle, AlbisT1, ahk_class OptoAppClass
				Albismenu(57602)
				while, (AlbisT1 = AlbisT2)
				{
						WinGetTitle, AlbisT2, ahk_class OptoAppClass
						If AlbisT1 <> AlbisT2
							break
						Sleep, 50
				}

	}

	PraxTT("Der Formularlauf ist beendet!", "5 3")
	Running:= ""

return
}

GVU_Pause() {

	Loop
	{
			If RegExMatch(Running, "^Vorsorgeautomatisierung$") || InStr(Running, "beendet") || (StrLen(Running) = 0)
				break
			Sleep, 50
			If RegExMatch(Running, "^Vorsorgeautomatisierung$") || InStr(Running, "beendet") || (StrLen(Running) = 0)
				break
			Sleep, 50
	}

return
}

GVU_CaveVonEintragen(UDatum, SaveData:=true) {                                               	;-- zum ändern der Cave! von Zeile 9 (hier bewahre ich Daten der letzten Untersuchungen auf)

		RegExMatch(UDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
		GVUMonat	:= aGVU1
		GVUJahr   	:= aGVU2

	; Ändern des Kürzeltextes der letzten GVU
		GVUZeile:= AlbisGetCaveZeile("", "GVU", true)

	; Sichern der ausgelesenen Cave! von Zeile 9
		PatientenID:= AlbisAktuellePatID()
		FileAppend, % A_DD "." A_MM "." A_Year ";" PatientenID ";" AlbisCurrentPatient() ";" GVUZeile "`n", % AddendumDir . "\Tagesprotokolle\cave9.txt", UTF-8

	; Erstellen der neuen Zeile ;{
		If RegExMatch(GVUZeile, "[gvuGVU]+.[hksHKS]+\/*[kvuKVU]*\s\d+.\d+")
			neueZeile:= RegExReplace(GVUZeile, "[gvuGVU]+.[hksHKS]+\/*[kvuKVU]*\s*\d+.\d+", "GVU/HKS " GVUMonat "^" GVUJahr)
		else
		{
				If InStr(GVUZeile, ",")
				{
						GVUTeile:= StrSplit(GVUZeile, ",")
						MaxIndex:= GVUTeile.MaxIndex()
						If RegExMatch(GVUTeile[MaxIndex], "GB\s*\d+\.\d\s*[AB]")
						{
								GVUTeile[MaxIndex+1] := GVUTeile[MaxIndex]
								GVUTeile[MaxIndex]  	:= "GVU/HKS " GVUMonat "^" GVUJahr
								Loop, % MaxIndex + 1
									neueZeile .= GVUTeile[A_Index] ", "
						}
						else
								neueZeile := GVUZeile ", GVU/HKS " GVUMonat "^" GVUJahr
				}
				else
				{
						if StrLen(GVUZeile) = 0
								neueZeile := "GVU/HKS " GVUMonat "^" GVUJahr ", "
						else
								neueZeile := GVUZeile ", GVU/HKS " GVUMonat "^" GVUJahr
				}

		}

		neueZeile:= RTrim(neueZeile, ", ")
		;}

	; Schreiben der neuen Zeile
		AlbisSetCaveZeile(9, neueZeile)
		Clipboard:= neueZeile

		If SaveData
			GVU_DatenSichern(PatientenID, UDatum, GVUZeile, neueZeile)

return
}

GVU_DatenSichern(PatientenID, UDatum, GVUAlt:="", GVUNeu:="") {

		PatVorhanden:= false

	; Sichern der Daten in die GVU Liste
		RegExMatch(UDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
		QListe:= GetQuartal("01." aGVU1 "." aGVU2)
		GVUtoAddfile:= AddendumDir "\Tagesprotokolle\" QListe "-GVU.txt"

	; Anlegen einer Sicherheitskopie der *gvu.txt Datei
		;Loop
		; {
		; 		backupFile:= AddendumDir "\TagesProtokolle\" SubStr(GVUFile, 1, StrLen(GVUFile) - 4) . A_Index . ".txt"
		; } until !FileExist(backupFile)
		; FileCopy, % AddendumDir "\TagesProtokolle\" GVUFile, % backupFile

		FileRead, gfile, % GVUtoAddfile
		Loop, Parse, gfile, `n, `r
		{
				If StrLen(A_LoopField) = 0
					continue
				else If InStr(A_LoopField, PatientenID)
				{
					PatVorhanden := true
					If RegExMatch(GVUAlt, "^\>\>\>")
						newgFile.= A_LoopField ";" GVUAlt "`n"
					else
						newgFile.= A_LoopField ";Alt:" GVUAlt ";Neu:" GVUNeu "`n"
				}
				else
					newgFile.= A_LoopField "`n"
		}

		If !PatVorhanden
			newgFile.= UDatum ";" SubStr(QList, 1, 2) "/" SubStr(QList, 3, 2) ";" PatientenID ";Alt:" GVUAlt ";Neu:" GVUNeu "`n"

		newgFile:= RTrim(newgFile, "`n")

		FileDelete	, % GVUtoAddfile
		FileAppend, % newgFile, % GVUtoAddfile, UTF-8

		If oPat.HasKey(PatientenID)
		{
				oPat[PatientenID].letzteGVU := UDatum
				PatDBSave(AddendumDBPath)
		}

return
}

GVU_AlteFormulardatenUebernehmen() {

	; Auslesen des letzten Abrechnungsdatum einer GVU
		entry := AlbisReadFromListbox("Alte Formulardaten übernehmen ahk_class #32770", 1, 1)
		RegExMatch(entry, "\d+\.(\d+)\.\d\d(\d+)", lastGVU)
		RegExMatch(EditDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
	; Berechnung der Monate zwischen zwei Vorsorgeuntersuchungen
		Untersuchungsabstand := ((aGVU2*12)+aGVU1) - ((lastGVU2*12)+lastGVU1)

		If (Untersuchungsabstand < Addendum.GVUAbstand)
		{
				;MsgBox, % "Zeile: " A_LineNumber "`ndelta: " Untersuchungsabstand "`nAbstand: " GVU[Abstand] ; GVU[Abstand]
				ClickReady := 1
				PraxTT("Bei diesem Patienten kann im Moment noch`nkeine weitere GVU abgerechnet werden.`nEs fehlen noch " 24 - Untersuchungsabstand " Monate!", "6 3")
				Sleep, 2000
				VerifiedClick("Button2",  "Alte Formulardaten übernehmen ahk_class #32770")	; Button 2 - [Abbruch]
				GVU_DatenSichern(AlbisAktuellePatID(), EditDatum, ">>>Mindestabstand zwischen zwei Vorsorgeuntersuchung nicht errreicht<<<", "")
				GVU_CaveVonEintragen(lastGVU, false)
		}
		else
		{
				ClickReady := 4
				PraxTT("Eine neue GVU wird angelegt.", "3 3")
				GVU_FormularDatenAuswaehlen()
		}


return ClickReady
}

GVU_FormularDatenAuswaehlen() {

		VerifiedSetFocus("Listbox1", "Alte Formulardaten übernehmen ahk_class #32770", 1)
		SendInput, {Space}
		Sleep, 100
		VerifiedClick("Button1", "Alte Formulardaten übernehmen ahk_class #32770")
		WinWaitClose, % "Alte Formulardaten übernehmen ahk_class #32770", , 3

return
}

GVU_KeineAlteDatenVorhanden() {

		VerifiedClick("Button1", "ALBIS ahk_class #32770", "alten Daten vorhanden") ; schließen

		hSeite1:= WinExist("Muster 30 (01.2009), Gesundheitsuntersuchung (Seite 1)")
		VerifiedCheck("Button52", hSeite1) ; Beweg.apparat
		VerifiedCheck("Button57", hSeite1) ; 140/90
		  VerifiedClick("Button61", hSeite1)

		WinWait, % "Muster 30 (01.2009), Gesundheitsuntersuchung (Seite 2) ahk_class #32770",,3
		hSeite2:= WinExist("Muster 30 (01.2009), Gesundheitsuntersuchung (Seite 2)")
		VerifiedCheck("Button37", hSeite2) ; orthopädische Erkrankung
		VerifiedCheck("Button57", hSeite2) ; sonstiges

		VerifiedClick("Button61", hSeite2)	;Button Weiter

return
}

GVU_LeistungsketteBestaetigen() {

		VerifiedSetFocus("Listview321", "Leistungskette bestätigen ahk_class #32770")
		SendInput, {Space}
		Sleep, 100
		SendInput, {Space}
	; Dialog schliessen
		VerifiedClick("Button1", "Leistungskette bestätigen ahk_class #32770")
		WinWaitClose, Leistungskette bestätigen ahk_class #32770,, 3

return
}

GVU_HKSFormularBefuellen() {

	; Verdachtsdiagnose "nein" - warum Fehler wenn nicht ausgewählt
		VerifiedClick("Button28", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Malignes Melanom"
		VerifiedClick("Button10", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Basalzellkarzinom"
		VerifiedClick("Button12", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Spinozelluläres Karzinom"
		VerifiedClick("Button14", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Anderer Hautkrebs"
		VerifiedClick("Button24", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Sonstiger dermatologisch abklärungsbedürftiger Befund"
		VerifiedClick("Button26", "Hautkrebsscreening - Nichtdermatologe")
	; nein - "Screening-Teilnehmer wird an einen Dermatologen überwiesen:"
		VerifiedClick("Button30", "Hautkrebsscreening - Nichtdermatologe")
	; Häkchen bei "gleichzeitige Gesundheitsvorsorge setzen"
		VerifiedClick("Button16", "Hautkrebsscreening - Nichtdermatologe")
	; "Speichern" und damit auch Schließen des Formulares
		VerifiedClick("Button19", "Hautkrebsscreening - Nichtdermatologe")

return
}

Hausbesuchskomplexe() {		; Hausbesuchsziffern können gescrollt werden

		static HBZiffern, AktiveKlasse, AktivZuletzt, AktiveID, UpDown

		SetKeyDelay, 10, 20

		If !IsObject(HBZiffern) {
				HBZiffern  	  := Object()
				HBZiffern.Push([01410,01411,01412])
				HBZiffern.Push([97234,97235,97236,97237,97238,97239])
				HBZiffern.Push([1,2,3,4,5,6,7,8,9])
		}

		SendRaw, 01410-97234-03230`(x:3`)

		AktiveID:=AktivZuletzt	:= GetHex(GetFocusedControlHwnd())
		AktiveKlasse              	:= GetClassName(AktivZuletzt)
		ControlGetPos		, cx, cy,,,                                        	, % "ahk_id " AktiveID

		ToolTip, % "Pfeiltaste 'Hoch' oder 'Runter' zum ändern der Ziffern", % cx - 40, % cy - 40, 11
		SetKeyDelay, 30, 40

		while !GetKeyState("TAB")
		{
				if GetKeyState("Up")
						UpDown := 1
				else if GetKeyState("Down")
						UpDown := -1

				If (UpDown= 1) || (UpDown = -1)
				{
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

									Index 		:= SubIndex + UpDown > HBZiffern[word].MaxIndex() ? 1 : ( subIndex + UpDown < 1 ) ? HBZiffern[word].MaxIndex() : (subIndex + UpDown)
									newline		:= StrReplace(line, HBZiffern[word, SubIndex], HBZiffern[word, Index] )
									ToolTip, % "Pfeiltaste 'Hoch' oder 'Runter' zum ändern der Ziffern. " hbStr " , " HBZiffern[word, SubIndex], % cx - 40, % cy - 40, 11
							}
							else
							{
									RegExMatch(line, "(?<=03230\(x:)\d", faktor)
									faktor	:= ( faktor + UpDown > 9 ) ? 1 : ( faktor + UpDown < 1 ) ? 9 : (faktor + UpDown)
									newline := RegExReplace(line, "(?<=03230\(x:)\d", faktor)
							}

							ControlSetText,, % Trim(newline, "`n`r"), % "ahk_id " AktivZuletzt
							SetKeyDelay, -1, 0
							ControlSend	 ,, % "{Right " col "}"		  , % "ahk_id " AktivZuletzt
							SetKeyDelay, 30, 30

							; 14274

							line:= newline:= col:= hbStr:= word:= SubIndex:= lko:= UpDown:= ""
				}

				Sleep, 60
				UpDown:= 0
				;~ ControlGetFocus, Aktiv, % "ahk_id " AktiveID
				;~ If !InStr(AktiveKlasse, Aktiv)
							;~ break

				If !InStr(GetHex(GetFocusedControlHwnd()), AktiveID)
					break
		}

		ToolTip,,,, 11
		SetKeyDelay, 0, 0

return
}


;}

;{14. WinEventHooks - alle Funktionen und Labels
;   	 auf diesen Funktionen beruhen die wichtigsten vorgenommenen Automatisierungen


; ------------------------------------ Initialisieren der Hooks
InitializeWinEventHooks() {                                                                                             	; Robotic Process Automation (RPA) fängt hiermit an!

		global hMsgGui, hEditMsgGui

		;https://docs.microsoft.com/en-us/windows/desktop/winauto/event-constants ;{
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

		EVENT_SKIPOWNTHREAD         	:= 0x0001
		EVENT_SKIPOWNPROCESS       	:= 0x0002
		EVENT_OUTOFCONTEXT          	:= 0x0000
	;}


	; creating the hooks
		HookProcAdr         	:= RegisterCallback("WinEventProc", "F")
		hWinEventHook     	:= SetWinEventHook( 0x0003, 0x0003, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x0010, 0x0010, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8000, 0x8000, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8002, 0x8002, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook     	:= SetWinEventHook( 0x8001, 0x8001, 0, HookProcAdr, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )

		Addendum.Hooks[1]:= {"HPA": HookProcAdr, "hWinEventHook": hWinEventHook}

	/* Shellhook für neue Programmfenster - es werden nur Fenster der obersten Ebene (top-level-windows) erkannt ...

			Damit kann ich WinEvent Hooks, die nur Albis überwachen, realisieren. Wird Albis geschlossen, weil es z.B. neugestartet werden muss, wird der momentan aktive WinEvent Hook beendet.
			Der Shell Hook wartet auf einen neuen Albisprozeß und startet nach Erscheinen des Albisfenster einen neuen Hook. Ohne diese kleinen Umweg würde der alte Hook weiter existieren und
			die Automationsfunktionen würden nicht aufgerufen.

			Wozu hooke ich nur Albis?
			Autohotkey ist eine Skriptsprache. Diese ist schnell, aber bei weitem nicht so schnell wie eine reine Programmiersprache wie C++. Es dürfen aber keine Nachrichten aus der
			Nachrichtenschlange aufgrund eines Geschwindigkeitsproblem verloren gehen, sonst riskiere ich fehlerhafte Funktionsaufrufe!

	*/

		Gui, AMsg: New, HWNDhMsgGui +ToolWindow
		Gui, AMsg: Add, Edit, xm ym w300 h300 HWNDhEditMsgGui
		Gui, AMsg: Show, AutoSize Hide, % "Addendum Message Gui"

		DllCall("RegisterShellHookWindow", UInt, hMsgGui)
		MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
		OnMessage(MsgNum, "ShellHookProc")

		Addendum["MsgGui"]["hMsgGui"]     	:= hMsgGui
		Addendum["MsgGui"]["hEditMsgGui"]	:= hEditMsgGui

	; starte Hooks die nur für Albis sind
		AlbisStartHooks()

return
}

AlbisStartHooks() {                                                                                                     		; startet Hooks die nur aufgerufen wenn Albis Ereignisse produziert

	If (AlbisPID := AlbisPID()) {

			Addendum.AlbisWinID          	:= AlbisWinID()
			Addendum.AlbisPID           	:= AlbisPID

			; Startet Hook zur Erkennung eines Wechsel des Eingabefocus
			Addendum.Hooks[2]          	:= AlbisFocusHook(AlbisPID)

			; nur für's Programmieren bei einer Bildschirmauflösung über 2k
				If InStr(compname, "SP1")
					Addendum.Hooks[3]     	:= AlbisLocationChangeHook(AlbisPID)

			; Test Hook - PopUpMenu abfangen und eigene Einträge integrieren
				Addendum.Hooks[4]       	:= AlbisPopUpMenuHook(AlbisPID)

	}

}

AlbisStopHooks() {                                                                                                         	; beendet ALBIS Prozesshooks

	local idx

	Gui, AMsg: 	Show, Hide
	Gui, adm: 	Destroy
	Gui, adm2: 	Destroy

	Addendum.Thread[1].ahkTerminate()
	TrayTip("Albis is closed.", "Addendum Toolbar: " (Addendum.Thread[1].ahkReady() ? "is still running":"is terminated"), 5)

	Loop % (Addendum.Hooks.MaxIndex() -1) {
			idx := A_Index + 1
			UnhookWinEvent(Addendum.Hooks[idx].hEvH, Addendum.Hooks[idx].HPA)
			Addendum.Hooks[idx] := ""
	}


}


; ------------------------------------  spezielle Hooks auf den Albis-Prozeß
AlbisPopUpMenuHook(AlbisPID) {                                                                                  	; Hook zum Abfangen eines Rechtsklick Menu

	HookProcAdr   	:= RegisterCallback("AlbisPopUpMenuProc", "F")
	hWinEventHook	:= SetWinEventHook( 0x0006, 0x0007, 0, HookProcAdr, AlbisPID, 0, 0x0003)

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}

AlbisFocusHook(AlbisPID) {                                                                                            	; Hook ausschließlich für EVENT_OBJECT_FOCUS Nachrichten wird gestartet

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

return {"HPA": HookProcAdr, "hEvH": hWinEventHook}
}


; ------------------------------------ Verarbeitung abgefangener Systemmeldungen (EventProcs)
WinEventProc(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {              	; zum Abfangen bestimmter Fensterklassen

	static class_filter	:= [ "OptoAppClass", "#32770", "Teamviewer", "WindowsForms10.Window", "Qt5QWindowIcon", "AutohotkeyGui", "SciTEWindow"
								,  "classFoxitReader", "IEFrame"]
	static app_filter 	:= { 1: "albis"
								,  	2: "albiscs"
								,  	3: "Teamviewer"
								,  	4: "infoBoxWebClient"
								,  	5: "ifap"
								,  	6: "Autohotkey"
								,  	7: "SciTE"
								,  	8: "FoxitReader"
								, 	9: "word"
								, 10: "iexplore"}

	Critical

	If (GetDec(hwnd) = 0) || (StrLen(wClass:= WinGetClass(hwnd)) = 0)
		return 0

	EHookEvent	:= event
	EHookHwnd	:= Format("0x{:x}", hwnd)

	; nur Fensterklassen Filter
		For i, filterclass in class_filter
			If InStr(wclass, filterclass) {
					SetTimer, EventHook_WinHandler, -0
					Critical Off
					return 0
			}

	; Order & Entry Prozess
		WinGet, EHproc, ProcessName, % "ahk_id " hwnd
		If Instr(EHProc, "infoBoxWebClient")
				SetTimer, EventHook_WinHandler, -0

	Critical Off

return 0
}

AlbisPopUpMenuProc(hHook, event, hwnd) {                                                                   	; eigene Menupunkte in Albiskontextmenu's unterbringen

	static pm

	If (event = 6)
		pm := Addendum_PopUpMenu(hwnd, GetMenuHandle(hwnd))
	else if (event = 7) && IsObject(pm)
		Addendum_PopUpMenuItem(pm)

	;SciTEOutput("PopUpMenuHook: (" GetHex(hwnd) ") event: " event " ,pm is object: " (IsObject(pm) ? "true":"false"))

}

AlbisFocusEventProc(hHook, event, hwnd) {                                                                    	; behandelt EVENT_OJECT_FOCUS Nachrichten von Albis

	static AlbisTitleO
	global hFEGui, FE, FEText1

	;Critical

	If (GetDec(hwnd) = 0)
		return 0

	hwnd	:= Format("0x{:x}", hwnd)
	cFocus	:= GetClassName(hwnd)

	If WinExist("ahk_id " hFEGui)
		GuiControl, FE:, FEText1, % hwnd ": " cFocus

	If InStr(cFocus, "RichEdit")	{
		ControlGetText, activeControlText,, % "ahk_id " hwnd
		activeControl.diaText	:= activeControlText
		activeControl.hwnd   	:= hwnd
		If WinExist("ahk_id " hFEGui)
			GuiControl, FE:, FEText1, % activeControlText
		;fn1:= Func("ADControlGet").Bind(hwnd)
		;SetTimer, % fn1, -0
		return 0
	}

	Addendum.AktiveAnzeige := AlbisHookTitle:= AlbisGetActiveWinTitle()
	If (AlbisTitleO = AlbisHookTitle) && RegExMatch(AlbisHookTitle, "^\s*\d+\s*\/") ;&& !RegExMatch)|(Laborbuch)|(Wartezimmer)|(EBM)(ToDo)")
		return 0

	AlbisTitleO := AlbisHookTitle
	fn2:= Func("CurrPatientChange").Bind(AlbisHookTitle)
	SetTimer, % fn2, -0

return 0
}

AlbisLocationChangeEventProc(hHook, event, hwnd) {                                                    	; behandelt EVENT_OBJECT_LOCATIONCHANGE Nachrichten von Albis

	If !Addendum.AlbisLocationChange || (A_ScreenWidth <= 1920)
		return 0

	If WinExist("ahk_id" hadm)
		WinSet, Redraw,, % "ahk_id " hadm

	a := GetWindowSpot( Addendum.AlbisWinID := AlbisWinID() )
	If ((a.W <> 1922) || (a.H <> 1080)) && (a.X > -20) && (a.Y > -20)
		SetWindowPos(Addendum.AlbisWinID, a.X, a.Y, 1922, 1080)

	If (a.X <= -20) && (a.Y <= -20)
		Addendum.Thread[1].ahkPostFunction("ToolbarShowHide" , "Hide")
	else If (a.X > -20) && (a.Y > -20)
		Addendum.Thread[1].ahkPostFunction("ToolbarShowHide" , "")

return 0
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
		static func_ifap

		Critical, 50

	; return on empty ;{
		If (wParam = 0)
			return 0

		class := WinGetClass(wparam)
		If (StrLen(class) = 0)
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
		TbIsRunning := Addendum.Thread[1].ahkReady()

	; Albisprozess
		If   		( WinExist("ahk_class OptoAppClass") && (lParam in 1,6)	&& !Addendum.Thread[1].ahkReady() )  	{		; ALBIS wurde gestartet

				PraxTT("Ein neuer Albisprozeß wurde erkannt.`nDie Fensterhooks wurden gesetzt!", "2 0")
			; startet Hooks
				AlbisStartHooks()

			; Addendum Toolbar neu starten
				If !IsObject(Addendum.Thread[1])
						Addendum.Thread[1] := AHKThread(Addendum.ToolbarSkript)

			return 0

		}
		else if 	(!WinExist("ahk_class OptoAppClass") && (lParam = 2)  	&&  Addendum.Thread[1].ahkReady() )	{		; ALBIS wurde beendet

				PraxTT("Albis wurde beendet. Hooks wurden entfernt.", "2 0")
				AlbisStopHooks()
				return 0
		}

	; andere Programme (Microsoft Word, Ifap)
		If      	( InStr(class, "OpusApp")                 	&& (lParam in 1,6) )                                                          	{ 		; Wordfenster
				SetTimer, EventHook_WinHandler, -0
				return 0
		}
		else if 	( InStr(class, "TForm_ipcMain")         	&& (lParam in 1,6) )                                                       	{  	; ifap wurde aktiviert

			; das ifap Fenster kann manchmal nicht maximiert werden, x,y sind dann stark negativ und Größe und Breite sind unter 0
			; als Timer angelegt, da ein auflaufendes Event den Ablauf unterbrechen würde
				func_ifap := Func("ifapCorrectWinPos").Bind(SHookHwnd)
				SetTimer, % func_ifap, -300   ; dem ifap Fenster Zeit geben um sich neu zu zeichnen, erst danach verschieben
				return 0
		}

	; für Addendum_PopUpMenu - ruft die als String in der Variable .PopUpMenuCallback Funktion auf und übergibt zusätzliche Parameter
		If (StrLen(Addendum.PopUpMenuCallback) > 0)
			For idx, viewerclass in PDFViewer
				If InStr(class, viewerclass) {

					If InStr(Addendum.PopUpMenuCallback, "|") {

							callbackParam := StrSplit(Addendum.PopUpMenuCallback, "|")
							If !IsFunc(callbackParam[1])
								return 0
							func_Call 	:= Func(callbackParam[1]).Bind(callbackParam[2], class, SHookHwnd)

					}
					else {

							fnName := Addendum.PopUpMenuCallback
							If !IsFunc(fnName)
								return 0
							func_Call 	:= Func(fnName).Bind(class, SHookHwnd)

					}

					Addendum.PopUpMenuCallback := ""
					SetTimer, % func_Call, -300
					return 0

				}

return 0
}


; ------------------------------------ Automatisierungsroutinen / Fensterhandler
EventHook_WinHandler:                                                                                                	;{ Eventhookhandler - Popupblocker/Fensterhandler - für diverse Fenster verschiedener Programme

		If EHWHStatus
			return
		EHWT:= WinGetTitle(EHookHwnd), EHWText:= WinGetText(EHookHwnd)
		If (StrLen(EHWT . EHWText) = 0) || RegExMatch(EHWT, "^List1|OK") || RegExMatch(EHWText, "^SkinLoader")
			return
		EHWHStatus :=true

		WinGet, EHproc1, ProcessName, % "ahk_id " (hHookedWin:= EHookHwnd)
		If InStr(EHproc1, "albis") || InStr(EHproc1, "wkflsr32")
		{
				;EHWTS := StrReplace(EHWText	, "`n", ""), EHWTS := StrReplace(EHWTS	, "`r", ""), SciTEOutput("HH: " EHookHwnd ", EHWT: " EHWT ", EHWTEXT: " SubStr(EHWTS, 1, 40) )
				If InStr(EHWText, "ist keine Ausnahmeindikation")                                                                          	; Fenster wird geschlossen
				{
							VerifiedClick("Button2", hHookedWin)
							GNRChanged 	:= true
							EHWHStatus  	:= false
				}
				else If InStr(EHWTitle, "Daten von") || GNRChanged                                                                   	; schließt nach Änderung der GNR das Fenster "Daten von " für schnelleres Handling
				{
							If !WinExist("", "ist keine Ausnahmeindikation")  {
								VerifiedClick("Button30", "Daten von ahk_class #32770")
								GNRChanged 	:= false
							}
							EHWHStatus  	:= false
				}
				else If InStr(EHWText, "ALBIS wartet auf Rückgabedatei") ;&& Instr(compname, "Labor")                    	; in Abhängigkeit des Client wird das Fenster sofort oder verzögert geschlossen
				{
							VerifiedClick("Button1", "ahk_class #32770", "ALBIS wartet auf Rückgabedatei")
							EHWHStatus := false
				}
				else If InStr(EHWText, "Patient hat in diesem Quartal")                                                                   	; in Abhängigkeit des Client wird das Fenster sofort oder verzögert geschlossen
				{
							If !VerifiedClick("Button1", "ALBIS ahk_class #32770", "Patient hat in diesem Quartal")
								VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if Instr(EHWText, "Der Parameter ist bereits in dieser Gruppe")                                                  	; Fenster wird geschlossen
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if Instr(EHWT, "Patientenausweis")                                                                                           	; automatisch drucken
				{
							VerifiedClick("Button23", hHookedWin)
							EHWHStatus := false
				}
				else if InStr(EHWText, "folgende Diagnosen übernommen")                                                           	; Fenster wird geschlossen
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if InStr(EHWT, "ALBIS - Login") && AutoLogin                                                                         	; Client spezifisches automatisches Einloggen in Albis
				{
							AlbisAutoLogin(10)
							EHWHStatus := false
				}
				else If InStr(EHWText, "Sie haben eine Besuchsziffer ohne Wege")                                                 	; Fenster wird geschlossen
				{
							VerifiedClick("Button2", "Sie haben eine Besuchsziffer ohne Wege")
							EHWHStatus := false
				}
				else If InStr(EHWText, "Datum darf nicht vordatiert werden")                                                         	; Datum wird verändert, Daten werden eingetragen, Datum wird zurückgesetzt
				{
							VerifiedClick("Button1", "ALBIS ahk_class #32770", "Datum darf nicht")
							fControlHwnd:= GetFocusedControlHwnd()
							ControlGetText, VorDatierDatum, Edit2, ahk_class OptoAppClass
							AlbisSetzeProgrammDatum(VorDatierDatum)
							VerifiedSetFocus("", "", "", fControlHwnd)
							SendInput, {Tab}
							EHWHStatus := false
				}
				else If InStr(EHWText, "Wollen Sie wirklich die GNR ändern")                                                          	; es wird automatisch mit "Ja" bestätigt
				{
							VerifiedClick("Button1", "ALBIS ahk_class #32770", "Wollen Sie wirklich")
							EHWHStatus  	:= false
				}
				else If InStr(EHWText, "Der Patient wurde bereits als verstorben")                                                     	; es wird automatisch mit "Ja" bestätigt
				{
							VerifiedClick("Button1", "ALBIS ahk_class #32770", "Der Patient wurde bereits als verstorben")
							EHWHStatus := false
				}
				else If InStr(EHWText, "Fehler beim Aufruf dppivassist")                                                                  	; Fenster wird geschlossen
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else If InStr(EHWText, "Folgender Pfad existiert nicht")                                                                   	; Fenster wird geschlossen
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if WinExist("Gesundheitsvorsorgeliste") && Instr(EHWT, "Herzlichen Glückwunsch")                  	; Fenster "Herzlichen Glückwunsch" schliessen
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if InStr(EHWText, "Der Druckauftrag konnte nicht gestartet werden")                                         	; Fenster wird geschlossen
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if InStr(EHWT, "CGM LIFE")	                                                                                                	; Verbindungsfehler (gesperrt durch Windows Firewall)
				{
						WinClose, % "CGM LIFE ahk_class Afx:00400000:0:00010005:86101803"
				}
				else if InStr(EHWT, "CGM HEILMITTELKATALOG")                                                                        	; Fensterposition wird innerhalb des Albisfenster zentriert
				{
							AlbisHeilmittelKatologPosition()
							EHWHStatus := false
				}
				else if InStr(EHWT, "Dauerdiagnosen von")                                                                                   	; automatisches Sortieren von anamnestischen und Behandlungsdiagnosen
				{
						;AlbisSortiereDiagnosen() ; abgeschaltet da ich noch keine funktionierende LV_Select Funktion gefunden habe
						res:= AlbisResizeDauerdiagnosen()
						EHWHStatus := false
				}
				else if InStr(EHWT, "ICD-10 Thesaurus")                                                                                       	; vergrößert den Diagnosenauswahlbereich
				{
						UpSizeControl("ICD-10 Thesaurus", "#32770", "Listbox1", 200, 150, AlbisWinID())
						EHWHStatus := false
				}
				else if InStr(EHWT, "Muster 1a") || InStr(EHWText, "Der Patient ist noch AU")                                    	; Arbeitsunfähigkeitsbescheinigung - Einblendung von zusätzlichen Informationen
				{
						AlbisFristenGui()
						EHWHStatus := false
				}
				else if InStr(EHWT, "Muster 16")                                                                                                  	; Kassenrezept - Schnellrezept + Auto autIdem Kreuz
				{
						;TrayTip, % "Rezeptformular", % "Rezeptformular wurde geöffnet", 2
						AlbisRezeptHelferGui(Addendum.AdditionalData_Path "\RezepthelferDB.json")
						;result := AlbisRezeptHook(hHookedWin)
						;Menu, Tray, Tip, % StrReplace(A_ScriptName, ".ahk") " V." Version " vom " vom "`n" result
						;AlbisRezeptAutIdem()
						EHWHStatus := false
				}
				else If InStr(EHWText, "Bitte die Straße")                                                                                       	; weitere Information Dialog (Patientendaten) und Adressproblematik
				{
						; bei Adressen mit Postleitzahlen, kann das Fenster 'weitere Informationen' nicht einfach geschlossen werden
						MsgBox, 0x1024, Addendum für Albis on Windows, Rechnungsempfänger löschen?
						IfMsgBox, Yes
						{
								Loop 8
									VerifiedSetText("Edit" A_Index, "", "ahk_class #32770", 200, "Adresse des Rechnungs")
						}
						else {
								VerifiedSetText("Edit5", " ", "ahk_class #32770", 200, "Adresse des Rechnungs")
						}
						VerifiedClick("Button2", "ahk_class #32770", "Adresse des Rechnungs")
				}
				;-------------------- Fehlerbenachrichtungen Java Runtime des ifap praxisCENTER ----
				else If InStr(EHWText, "Fehler im ifap praxisCENTER")
				{
                            VerifiedClick("Button1", hHookedWin)
							If !ifaptimer
								SetTimer, ifapActive, 30
							ifaptimer:= A_TickCount
							EHWHStatus := false
				}
				;-------------------- Laborabruf Automation --------------------------------------------------------            	;{
				else If InStr(EHWText, "Anforderungen ins Laborblatt übertragen") && Addendum.Labor.AutoAbruf   	; es wird automatisch mit "Ja" bestätigt
				{
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
							;return
				}
				else If InStr(EHWText, "Keine Datei(en) im Pfad") 	&& Addendum.Labor.AutoAbruf                            	; Laborabrufvorgang wird beendet
				{
							Addendum.Laborabruf.Voll 	:= false
							Addendum.Laborabruf.Daten	:= false
							Addendum.Laborabruf.Status := false
							PraxTT("", "off")
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
				}
				else if Instr(EHWT, "Labor auswählen")            	&& Addendum.Labor.AutoAbruf
				{
							AlbisLaborAuswählen(Addendum.Labor.LaborName)
							EHWHStatus := false
				}
				else If InStr(EHWT, "Labordaten")                     	&& Addendum.Labor.AutoAbruf
				{
							AlbisLaborDaten()
							EHWHStatus := false
				}
				else If InStr(EHWT, "GNR der Anford")             	&& Addendum.Labor.AutoAbruf
				{
							AlbisRestrainLabWindow(1, hHookedWin)       	;Listbox1 enthält Rechnungsdaten
							EHWHStatus := false
				}  ;}
				;-------------------- GVU und HKS Formular Automatisierung -----------------------------------------      	;{
				else If Addendum.GVUAutomation && InStr(EHWT, "Muster 30") && (ClickReady > 0)
				{
						RegExMatch(EHWT, "(?<=Seite\s)\d", FormularSeite)
						If (FormularSeite = 1)
						{
								If (ClickReady = 1)
								{
										ClickReady := 2
									; überprüft auf welche Art das GVU Formular aufgerufen wurde, verhindert die Automatisierung schon angelegter Formulare
									; funktioniert nur wenn man sich ein Albismakro mit dem Namen GVU anlegt! In Edit3 steht dann GVU
										EditDatum:= EditKuerzel := ""
										ControlGetText, EditDatum , Edit2, ahk_class OptoAppClass
										ControlGetText, EditKuerzel, Edit3, ahk_class OptoAppClass
										If (StrLen(EditDatum) = 0) || !InStr(EditKuerzel, "GVU")
										{
												DatumZuvor	:= ""
												ClickReady	:= 1
												EHWHStatus := false
												;return
										}

									; Quartal, Untersuchungsmonat und Jahr ermitteln
										AbrQuartal := LTrim(GetQuartal(EditDatum, "/"), "0")
										RegExMatch(EditDatum, "\d+\.(\d+)\.\d\d(\d+)", aGVU)
										GVU_AlteDatenStartZeit := A_TickCount
										PraxTT("Warte auf alte Daten", "0 3")
										SetTimer, GVU_WarteAufAlteDaten, 250
										VerifiedClick("Button63", "Muster 30 ahk_class #32770")	        	; Button63 = [Alte Daten] - GVU Seite 1
										EHWHStatus := false
										;return
								}
								else if (ClickReady = 3)
								{
										SetTimer, GVU_WarteAufAlteDaten, Off
										PraxTT("Abbruch", "2 3")
										VerifiedClick("Button62", "Muster 30 ahk_class #32770")	        	; Button62 = [Abbruch]	 - GVU Seite 1
										If RegExMatch(DatumZuvor, "\d{2}\.\d{2}\.\d{4}")
											AlbisSetzeProgrammDatum(DatumZuvor)
										DatumZuvor := ""
										SetTimer, SetClickReadyBack, -4000
										EHWHStatus := false
										;return
								}
								else if (ClickReady = 4)
								{
										SetTimer, GVU_WarteAufAlteDaten, Off
										PraxTT("Weiter", "2 3")
										VerifiedClick("Button61", "Muster 30 ahk_class #32770")	        	; Button61 = [Weiter]   	 ->	GVU Seite 2
										EHWHStatus := false
										;return
								}
						}
						else if (FormularSeite = 2)
						{
								ClickReady := 5
								SetTimer, GVU_WarteAufAlteDaten, Off
								PraxTT("GVU Formular Seite 2 schließen", "2 3")
								VerifiedClick("Button61", "Muster 30 ahk_class #32770", "", "", 3)      	; Button61 = [Weiter]   	 ->	GVU Seite 2
								EHWHStatus := false
								;return
						}

						EHWHStatus := false
						;return
				}
				else If Addendum.GVUAutomation && InStr(EHWText, "keine alten Daten vorhanden")          	&& (ClickReady = 2)
				{
							ClickReady := 0
							PraxTT("", "off")
							SetTimer, GVU_WarteAufAlteDaten, Off
							GVU_KeineAlteDatenVorhanden()
							ClickReady := 5
				}
				else If Addendum.GVUAutomation && InStr(EHWT, "Alte Formulardaten übernehmen")         	&& (ClickReady = 2)
				{
							ClickReady := 0 ; auf 0 gesetzt damit nichts anderes den Ablauf hier behindern kann
							SetTimer, GVU_WarteAufAlteDaten, Off
							ClickReady := GVU_AlteFormulardatenUebernehmen()
							EHWHStatus := false
							;return
				}
				else If Addendum.GVUAutomation && InStr(EHWText, "Soll das Makro abgebrochen werden")	&& (ClickReady > 1)
				{
							VerifiedClick("Button1", "ALBIS", "Soll das Makro abgebrochen werden")
							ClickReady	:= 1
							SetTimer, SetClickReadyBack, off
							EHWHStatus := false
							;return
				}
				else If Addendum.GVUAutomation && InStr(EHWT, "GNR-Vorschlag zur Befundung")            	&& (ClickReady = 5)
				{
							entries := ""
							while, (StrLen(entries) = 0)
							{
									entries := AlbisReadFromListbox("GNR-Vorschlag zur Befundung", 1, 0)
									If StrLen(entries) > 0
										break
									;ToolTip, % "entrieTry: " A_Index "`nQuartal: " AbrQuartal "`nEinträge: " entries, 1250, 1, 5
									Sleep, 300
									If A_Index > 5
									{
											PraxTT("Die Erkennung des Abrechnungsquartals im Dialog hat nicht funktioniert`nBitte wählen Sie das richtige Abrechnungsquartal manuell aus!", "8 4")
											ClickReady:= 6
											;return
									}
							}

							Loop, Parse, entries, `n
								If InStr(A_LoopField, AbrQuartal)
								{
										PostMessage, 0x186, % A_Index - 1, 0, ListBox1, % "ahk_id " WinExist("GNR-Vorschlag zur Befundung ahk_class #32770") ;LB_SetCursel
										entrieIndex:= A_Index - 1
										break
								}

							;ToolTip, % "entrieId: " entrieIndex "`nQuartal: " AbrQuartal "`nEinträge: " entries, 1250, 1, 5
							If !VerifiedClick("Button1", "GNR-Vorschlag zur Befundung ahk_class #32770", "", "", 3) ;GNR-Vorschlag schliessen
									MsgBox, Bitte schließen Sie den Dialog`nGNR-Vorschlag zur Befundung

							ClickReady := 6
							WinWait, Hautkrebsscreening - Nichtdermatologe ahk_class #32770,, 8
							If WinExist("Hautkrebsscreening - Nichtdermatologe ahk_class #32770")
								WinActivate, Hautkrebsscreening - Nichtdermatologe ahk_class #32770
							else
								Albismenu(34505, "Hautkrebsscreening - Nichtdermatologe")     	;"eHautkrebs-Screening Nicht-Dermatologe": "34505" wird aufgerufen

							EHWHStatus := false

				}
				else If Addendum.GVUAutomation && InStr(EHWT, "Hautkrebsscreening - Nichtdermatologe")	&& (ClickReady = 6)
				{
							GVU_HKSFormularBefuellen()
							ClickReady:= 7
							EHWHStatus := false
				}
				else If Addendum.GVUAutomation && InStr(EHWT, "Leistungskette bestätigen")                     	&& (ClickReady = 7)
				{
						ClickReady:= 1
						GVU_LeistungsketteBestaetigen()
						GVU_CaveVonEintragen(EditDatum)
						If RegExMatch(DatumZuvor, "\d{2}\.\d{2}\.\d{4}")
							AlbisSetzeProgrammDatum(DatumZuvor)
						DatumZuvor := ""
						EHWHStatus := false
				}
				else If Addendum.GVUAutomation && InStr(EHWText, "Übertrage Gebühren-Nummer(n)...")
				{
							VerifiedClick("Button1", "ALBIS", "Übertrage Gebühren-Nummer")
							EHWHStatus := false
				}
				else If Addendum.GVUAutomation && InStr(EHWText, "außerhalb des Quartals in dem der Schein gültig ist")
				{        	;<-- überarbeiten - mehr Code notwendig
							VerifiedClick("Button1", hHookedWin)
							EHWHStatus := false
							;return
				}
				;}

				EHWHStatus := false

		}
        else if InStr(EHproc1, "ipc")                                                                                     	; ifap Programm
		{
					If InStr(EHWText, "Fehler im ifap praxisCENTER") {
							VerifiedClick("Button1", hHookedWin)
							If !ifaptimer
									SetTimer, ifapActive, 50
							ifaptimer:= A_TickCount
							EHWHStatus := false
							return

							ifapActive:    ; holt das ifap Fenster (nach der hoffentlich letzten Fehlermeldung) in den Vordergrund
								if (A_TickCount - ifapTimer > 200) {
										ifaptimer := 0
										WinActivate, praxisCENTER ahk_class TForm_ipcMain
										SetTimer, ifapActive, off
								}
							return
					}
		}
		else if ((Addendum.Labor.AutoAbruf = true) && InStr(EHproc1, "infoBoxWebClient") )	; fängt das WebFenster meines Labors ab
		{
				IniWrite, % A_DD "." A_MM "." A_YYYY " (" A_Hour ":" A_Min ":" A_Sec ")", % Addendum.AddendumIni, Labor, letzter_Laborabruf
				If !Addendum.Laborabruf.Status {
					PraxTT("Das Laborabruf Fenster wurde detektiert!`n#3Addendum übernimmt den weiteren Vorgang!", "40 4")
					MCSVianova_WebClient(hHookedWin)
				}

				EHWHStatus := false
				return
		}
		else if Instr(EHproc1, "Scan2Folder")                                                                       	; Fujitsu Scansnap Software
		{
				If Instr(EHWText, "Die Dateien wurden erfolgreich gespeichert")
						VerifiedClick("Button1", hHookedWin)

				EHWHStatus := false
		}
		else if Instr(EHproc1, "Autohotkey")                                                                         	; WinSpy (Autohotkey Tool) und Debuggerfenster
		{
				If InStr(EHWT, "WinSpy")
					MoveWinToCenterScreen(hHookedWin)
				else if InStr(EHWT, "Variable list")
					WinMoveZ(hHookedWin, 0, Floor(A_ScreenWidth/2), 300, 600, 1200)
					;SetWindowPos(hHookedWin, Floor(A_ScreenWidth//2), 300, 600, 1200)

				EHWHStatus := false
				return
		}
		else if Instr(EHproc1, "InternalAHK")                                                                        	; SciTE Editor - Variablenfenster verschieben
		{
				if InStr(EHWT, "Variable list")
					WinMoveZ(hHookedWin, 0, Floor(A_ScreenWidth/2), 300, 600, 1200)
					;SetWindowPos(hHookedWin, Floor(A_ScreenWidth//2), 300, 600, 1200)
				EHWHStatus := false
		}
		else if Instr(EHproc1, "SciTE")                                                                                  	; SciTE Editor
		{

				If InStr(EHWT, "SciTE4AutoHotkey") && Instr(WinGetClass(hHookedWin), "#32770") && RegExMatch(EHWText, "Datei.*Addendum\.ini")			{

					; nach dem Laden einer veränderten Addendum.ini Datei - wird diese von SciTE immer entfaltet. Das macht alles unübersichtlich. Deshalb wird hier automatisches Codefolding durchführen
						VerifiedClick("Button1", "SciTE4AutoHotkey ahk_class #32770")
						SciTeFoldAll()
						EHWStatus := false

				}
		}
		else if InStr(EHproc1, "FoxitReader")                                                                          	; Foxitreader signieren vereinfachen (und zwar tatsächlich!)
		{

				If Addendum.PDFSignieren	{

					If InStr(EHWT, "Sign Document") || InStr(EHWT, "Dokument signieren") {          	; für die englische und deutsche FoxitReader Version

							Addendum.PDFRecentlySigned := true
							WinGet, EHPID      	, PID	, % "ahk_id " 	hHookedWin
							FoxitID := GetParent(hHookedWin)
							FoxitReader_SignDoc(hHookedWin, EHWT)
							EHWHStatus := false
							return

					}

					/*

					If (InStr(EHWT, "Speichern unter bestätigen") && Addendum.PDFRecentlySigned) {

							FoxitID := GetParent(hHookedWin)
							SciTEOutput("hSaveAs: " hHookedWin ", FoxitID: " FoxitID)

						; Dialogfenster schließen und Schließen des signierten FoxitReader Tabs
							If VerifiedClick("Button1", hHookedWin,,, true)
								If Addendum.AutoCloseFoxitTab		{
									Sleep, 500                    	; kurze Pause um das Neuzeichnen des Foxitreaderfensters abzuwarten
									result := FoxitInvoke("Close", FoxitID)
								}

						;ein PdfReader-Fenster nach vorne holen, wenn noch eines da ist
							If WinExist("ahk_id " FoxitID)
								WinActivate, % "ahk_id " FoxitID

							EHWHStatus := false
							return

					}
					else if (InStr(EHWT, "Speichern unter") && Addendum.PDFRecentlySigned)      	{

						; Dialog schließen
							VerifiedClick("Button3", hHookedWin,,, true)

							EHWHStatus := false
							return
					}
				*/

				}

		}
		else if InStr(EHproc1, "WinWord")
		{
				If Addendum.WordAutoSize && (A_ScreenWidth > 1920)
					WinMoveZ(hHookedWin, 0, 500, 500, 1600, 800)
					;SetWindowPos(hHookedWin, 500, 500, 1600, 800)

				EHWHStatus := false
		}
		else if InStr(EHproc1, "iexplore")
		{

				If (StrLen(Addendum.Labor.Kennwort) > 0) && InStr(EHWT, "CGM CHANNEL: Login") {

						; 4.3.4.1.4.3.4.1.4.1 - Pfad für URL-Eingabe
						; 4.6.4.2.4.1.4.1.1.2.4

						WinGet, hWnd, ID, % "CGM CHANNEL: Login ahk_class IEFrame"
						WinActivate, % "ahk_id " hwnd
						try
							oAcc := Acc_Get("Object", "4.6.4.2.4.1.4.1.1.2.4", 8, "ahk_id " hWnd)
						catch {
							EHWStatus := false
							return
						}

						ControlHwnd := GetHex(Acc_WindowFromObject(oAcc))
						ControlGetText, cText,, % "ahk_id " ControlHWnd
						;ToolTip, % "Laborlogin detektiert! Text: " cText ",hwnd: " ControlHWnd , 300, 1, 15
						If ControlHWnd {

							try
								oAcc := Acc_Get("Object", "4.6.4.2.4.1.4.1.1.2.4", 9, "ahk_id " hWnd)
							catch {
								EHWStatus := false
								return
							}

							ControlHwnd := Acc_WindowFromObject(oAcc)
							ControlFocus,, % "ahk_id " ControlHWnd
							SendInput, % "{Raw}" Addendum.Labor.Kennwort
							Sleep 2000
							SendInput, % "{Enter}"

						}

						EHWHStatus := false
				}
		}


	EHWHStatus := false

return


SetClickReadyBack: ;{
	ClickReady:= 1
return ;}
;}

WinEventHook_Helper_Thread:                                                                                      	;{ Labordaten - Fenster abfangen

	hpop:= GetLastActivePopup(AlbisWinID())
	If (hpop != AlbisWinID()) && (hpop != hpop_old)
	{
				EHWT	 := WinGetTitle(hpop)
				EHWText:= WinGetText(hpop)
				hpop_old:= hpop

				If InStr(EHWT, "Labordaten")
				{
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

ifapCorrectWinPos(ifapHwnd) {                                                                                       	; das ifap Fenster an eine bestimmte Position auf dem Bildschirm verschieben

		ifapPos := GetWindowSpot(ifapHwnd)
		If (ifapPos.X < -20) || (ifapPos.Y < -20) || (ifapPos.W < 600) || (ifapPos.H < 600)
			SetWindowPos(ifapHwnd, 100, 100, 1200, 800)

}

CurrPatientChange(AlbisTitle) {                                                                                      	; behandelt Änderung des Albisfenstertitels

		static 	AlbisTitleO
		global 	Laborabruf_Voll

		If (AlbisTitleO = AlbisTitle)
			return

		Addendum.AktiveAnzeige	:= AlbisGetActiveWindowType()
		                	 AlbisTitleO	:= AlbisTitle

		If InStr(AlbisTitle, "Prüfung EBM/KRW") && !InStr(AlbisTitle, "Abrechnung vorbereiten")	{
				AlbisActivate(1)
				Sleep 2000
				SendInput, {Esc}
				return
		}

	; Patientenakte geöffnet - dann Überprüfung des Patientennamen und Erstellen der AddendumGui
		If InStr(Addendum.AktiveAnzeige, "Patientenakte")         		{
					PatDb(AlbisTitle2Data(AlbisTitle), "exist")
					AlbisActivate(1)
					If Addendum.AddendumGui
							AddendumGui()
		}
		else If InStr(Addendum.AktiveAnzeige, "aPEK")	                 	{
				MsgBox, 4, Addendum für Albis on Windows, Hinweisdialog zur Prüfung EBM/KRW weiter ansehen?, 6
				IfMsgBox, Yes
						return
				AlbisCloseMDITab("Prüfung EBM/KRW")
		}
		else if InStr(Addendum.AktiveAnzeige, "Laborbuch") && (Laborabruf_Voll)
				Albismenu(34157, "", 6, 1)			;34157 - alle übertragen

	; AddendumGui schließen wenn Bedingungen nicht erfüllt sind
		If !Addendum.AddendumGui || (!InStr(Addendum.AktiveAnzeige, "Karteikarte") && !InStr(Addendum.AktiveAnzeige, "Laborblatt")) { ;!RegExMatch(Addendum.AktiveAnzeige, "i).*\|(Karteikarte)|(Laborblatt)")
				SetTimer, toptop, Off
				Gui, adm: 	Destroy
				Gui, adm2:	Destroy
		}

Return
}


;----------------------------------------------------------------
; zugehörige Funktionen für die WinEventHandler Labels
;----------------------------------------------------------------;{
GVU_WarteAufAlteDaten:

	If (A_TickCount - GVU_AlteDatenStartZeit > 6000)
	{
			SetTimer, GVU_WarteAufAlteDaten, Off
			PraxTT("Alte Daten wurden eingetragen", "5 3")
			ClickReady := 4
			VerifiedClick("Button61", "Muster 30 ahk_class #32770")	; Button61 = [Weiter]   	 - GVU Seite 1
			return
	}

	If (ClickReady <> 2)
	{
			SetTimer, GVU_WarteAufAlteDaten, Off
			PraxTT("Alte Daten wurden übernommen.", "5 3")
	}
return

SplashAus:
	SplashTextOff
return

;}

;}

;{15. Sicherheitsfunktionen

HideAllWindows(STEnd) {
	EnumAddress := RegisterCallback("EnumWindowsProcHide", "Fast")
	DllCall("EnumWindows", UInt, EnumAddress, UInt, 0)
}

EnumWindowsProcHide(lhwnd, lParam) {
    global Output, HiddenWindows
	DetectHiddenWindows, Off
    WinGetTitle, ltitle, ahk_id %lhwnd%
    WinGetClass, lclass, ahk_id %lhwnd%
	WinGetText, ltext, ahk_id %lhwnd%
	if ( (ltitle)or(lclass="Shell_TrayWnd") )   {
				if !Instr(ltitle, "Teamviewer") {                                        ;Teamviewer ausgeschlossen um auch von außerhalb entsperren zu können
					Output .= "Title: " . ltitle . "`tClass: " . lclass . "`n"
					WinHide, %ltitle% ahk_class %lclass%
					GroupAdd, HiddenWindows, ahk_id %lhwnd%
				}
    }
	PostMessage, 0x111, 29698, , SHELLDLL_DefView1, ahk_class Progman
    return true  ; Tell EnumWindows() to continue until all windows have been enumerated.
}


;}

;{16. Interskriptkommunikation / OnMessage
Receive_WM_COPYDATA(wParam, lParam) {                                                             	;-- empfängt Nachrichten von anderen Skripten die auf demselben Client laufen
	global InComing
    StringAddress := NumGet(lParam + 2*A_PtrSize)
	InComing := StrGet(StringAddress)
	SetTimer, MessageWorker, -10
    return true
}

MessageWorker:                                                                                                       	;{  verarbeitet die eingegangen Nachrichten

		MWmsg:= StrSplit(InComing, "|")

	; Kommunikation mit dem Praxomat Skript
		If InStr(MWmsg[1], "PatDBCount")	{
				result := Send_WM_COPYDATA("PatDBCount|" oPat.Count(), MWmsg[2])					;Send_WM_Copydata findet sich in AddendumFunctions.ahk
		}
		else If InStr(MWmsg[1], "PatDB")		{
				If InStr(MWmsg[2], "GetPatID")
						result := Send_WM_COPYDATA("PatDB|PatID|" FindPatData(oPat, "PatID", MWmsg.3, MWmsg.4, MWmsg.5, MWmsg.6), MWmsg.MaxIndex())
		}
		else If InStr(MWmsg[1], "TPCount")	{
				result := Send_WM_COPYDATA("TPCount|" TProtokoll.Count(), MWmsg[2])
		}
	; Kommunikation mit dem Abrechnungshelfer Skript
		else If RegExMatch(MWmsg[1], "^AutoLogin\s+Off")	{
				AutoLogin:= false
				result := Send_WM_COPYDATA("AutoLogin disabled", MWmsg[2])
		}
		else if RegExMatch(MWmsg[1], "^AutoLogin\s+On")	{
				AutoLogin:= true
				result := Send_WM_COPYDATA("AutoLogin enabled", MWmsg[2])
		}
	; generell mögliche Kommunikation



return
;}

WM_DISPLAYCHANGE(wParam, lParam) {                                                                 	;-- einsetzbar für Script-Editoren - AutoZoom bei Wechsel der Bildschirmauflösung

	; zoomt bei einer Auflösung > 1920 Scite4AHK (RDP Session)

	static ZoomScintilla1O, ZoomScintilla2O

	SCI_SETZOOM:=2373, SCI_GETZOOM:=2374
	ControlGet	, hScintilla1	, hwnd,, Scintilla1	, ahk_class SciTEWindow
	ControlGet	, hScintilla2	, hwnd,, Scintilla2	, ahk_class SciTEWindow

	SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla1
	ZoomScintilla1O := ErrorLevel
	SendMessage, SCI_GETZOOM, 0, 0,, % "ahk_id " hScintilla2
	ZoomScintilla2O := ErrorLevel

	SysGet m1, Monitor, 1
	If m1Right > 1920
	{
			SendMessage, SCI_SETZOOM, % 4, 0,, % "ahk_id " hScintilla1
			SendMessage, SCI_SETZOOM, % 1, 0,, % "ahk_id " hScintilla2
	}
	else
	{
			If hscintilla1
			{
				SendMessage, SCI_SETZOOM, % -1, 0,, % "ahk_id " hScintilla1
				SendMessage, SCI_SETZOOM, % -3, 0,, % "ahk_id " hScintilla2
				WinMove, ahk_class SciTEWindow,, 0, 0
				WinMaximize, ahk_class SciTEWindow,, 0, 0
			}

	}

	If hSpy:= WinExist("WinSpy")
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
					#####################################
					###########   Problemmedikamente    ###########
					###########            Asthma                ###########
					###########           Blutdruck              ###########
					###########             COPD                ###########
					###########           Diabetes               ###########
					###########          Cholesterin            ###########
					###########          Gerinnung             ###########
					###########             Gicht                  ###########
					###########          Hilfsmittel              ###########
					###########             Magen               ###########
					###########             Niere                 ###########
					###########         Osteoporose         ###########
					###########            Prostata              ###########
					###########             Psyche               ###########
					###########           Rheuma               ###########
					###########         Schilddrüse            ###########
					###########         Schmerzen             ###########
					###########          Stuhlgang             ###########
					###########           Sonstiges             ###########
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
			HotKey, Esc, MedTGuiEscape

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

;{18. PDF-Reader Automation

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
											,	"Fullscreen":                          	421  	; View
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

FoxitReader_GetPages(FoxitID="") {

	If !FoxitID
		FoxitID := WinExist("ahk_class classFoxitReader")

	ControlGetText, Pages, Edit3, % "ahk_id " FoxitID
	RegExMatch(Pages, "(?<Disp>\d+)\s*\/\s*(?<Max>\d+)", Page)

return {"disp":PageDisp, "max":PageMax}
}

FoxitReader_SignaturSetzen() {						                                                           			;-- ruft Signatur setzen auf und zeichnet eine Signatur in die linke obere Ecke des Dokumentes

		; letzte Änderung: 15.07.2020

			CoordModeMouse_before :=  A_CoordModeMouse
			CoordMode, Mouse, Relative

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Variablen
		;----------------------------------------------------------------------------------------------------------------------------------------------
			; ! NICHT ÄNDERN! dieser String wird für 'feiyus' FindText() Funktion benötigt, er entspricht der linken oberen Ecke des PDF-Frame (AfxWnd100su4)
			static TopLeft:=	"|<>*210$71.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000003zzzzs0000007"
			        	    	. 	"zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000T"
			            		. 	"zzzz0000000zzzzy0000001zzzzw0000003zzzzs0000007zzzzk000000DzzzzU000000Tzzzz0000000zzzzy0000001"

			static tryCount := 0, basetolerance := 0.3

			Addendum.PDFRecentlySigned := true

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Sicherheitsabfrage, Abbruch wenn kein FoxitReaderfenster
		;----------------------------------------------------------------------------------------------------------------------------------------------
			WinGetClass, activeClass, A
			If !InStr(activeClass, "classFoxitReader")
				return 0

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Ermitteln des Childhandles (Docwindow) im Foxitreader (Bildschirmposition des Dokuments)
		;----------------------------------------------------------------------------------------------------------------------------------------------
			ReaderID	  	:= WinExist("A")
			hDocWnd		:= Controls("FoxitDocWnd1", "ID"           	, "ahk_id " ReaderID)
			DocWnd   	:= Controls("FoxitDocWnd1", "ControlPos"	, "ahk_id " ReaderID)
			Reader      	:= GetWindowSpot(ReaderID)
			APos           	:= GetWindowSpot(AlbisWinID())

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; FoxitReader vorbereiten für das Platzieren der Signatur
		;----------------------------------------------------------------------------------------------------------------------------------------------
			PraxTT("sende Befehl: 'Signatur platzieren'", "12 2")
			FoxitInvoke("SinglePage", ReaderID)
			FoxitInvoke("Fit_Page"  	, ReaderID)
			FoxitInvoke("FirstPage" 	, ReaderID)

			; Tastatur- und Mauseingriffe durch den Nutzer kurzzeitig sperren
			;hk(1, 1, "!ACHTUNG!`n`nMaus- und Tastatureingaben sind für maximal 10 Sekunden gesperrt", 10)

			; Maus zum Dokumentenfenster schieben, Linksklick simulieren
			Sleep, 300
			MouseMove	, % Reader.X + DocWnd.X + 2, % Reader.Y + DocWnd.Y + 2, 0
			MouseClick	, Left

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; Signatur setzen Menupunkt aufrufen
		;----------------------------------------------------------------------------------------------------------------------------------------------
			FoxitInvoke("Place_Signature", ReaderID)

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sucht nach der linken oberen Ecke der Pdf Seite der 1.Seite, um dort im Anschluss die Signatur zu erstellen
		;----------------------------------------------------------------------------------------------------------------------------------------------
			PraxTT(" suche nach dem Signierbereich des Dokumentes.", "12 2")

			FindSignatureRange:
			; Funktion macht eine Bildsuche (entspricht der linken oberen Ecke des PDF Preview Bereiches)
			if (Ok := FindText(DocWnd.X, DocWnd.Y, DocWnd.W, DocWnd.H, basetolerance, 0, TopLeft)) {

				MouseMove, % Ok.1.1, % Ok.1.2, 0
				PraxTT("Fläche zum Signieren gefunden.", "4 2")
				X := ok.1.1, Y := ok.1.2, W := ok.1.3, H := ok.1.4, Comment := ok.1.5, X += W//2, Y += H//2
				MouseClickDrag, Left, % X, % Y, % (X + Addendum.SignatureWidth), % (Y + Addendum.SignatureHeight), 0

			} else {

				sleep, 100
				tryCount ++
				basetolerance += 0.1
				If (tryCount < 20)
					goto FindSignatureRange
				else
					return 0

			}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; sichern der akutellen ReaderID
		;----------------------------------------------------------------------------------------------------------------------------------------------
			Addendum.PDFSignaturedID   	:= ReaderID
			Addendum.PDFRecentlySigned 	:= true

			; Tastatur- und Mauseingriffe wieder entsperren
			;hk(0, 0, "Tastatur- und Mausfunktion sind wieder verfügbar!", 2)

			CoordMode, Mouse, % CoordModeMouse_before

return 1
}

FoxitReader_SignDoc(hDokSig, FoxitTitle, FoxitText:="") {			                        		        	;-- Winhook-Handler zum Bearbeiten des Dokument signieren Dialoges

		; letzte Änderung: 15.07.2020 - Dateidialogbehandlung verbessert

		static FoxitID
		Addendum.PDFRecentlySigned := true

		PraxTT("Das Fenster 'Dokument signieren' wird gerade bearbeitet....!'", "0 2")

	;{ Auslesen der Fensterposition des Albis Fenster, erstellen zweier Objecte

			AlbisWinID	:= AlbisWinID()
			AWI				:= GetWindowSpot(AlbisWinID) 	; AlbisWindowInfo = AWI.WindowX
			CWI          	:= GetWindowSpot(hDokSig)     	; ChildWindowInfo or WindowOfInterest

		  ; Verschieben des Dokument signieren Fensters, es wird mittig über Albis abgelegt
			WinMove, % "ahk_id " hDokSig,, % AWI.X + (AWI.W-CWI.W)//2, % AWI.Y + (AWI.H-CWI.H)//2
			sleep, 100
	;}

	;{ Felder im Signierfenster auf die in der INI festgelegten Werte einstellen

			WinActivate		, % "ahk_id " hDokSig
			WinWaitActive	, % "ahk_id " hDokSig,, 5

		; Signieren als -------------------------------------------------------------------------------------------------------
			ControlFocus 	, ComboBox1                                                          	, % "ahk_id " hDokSig
			Control, ChooseString, % Addendum.SignierenAls , ComboBox1     	, % "ahk_id " hDokSig
			sleep, 50

		; Darstellungstyp ----------------------------------------------------------------------------------------------------
			ControlFocus 	, ComboBox4                                                        	, % "ahk_id " hDokSig
			; prüft das Feld Signaturvorschau auf die in der ini hinterlegte Signatur
			ControlGet, entryNr, FindString	, % Addendum.Darstellungstyp, ComboBox4, % "ahk_id " hDokSig
			If !entryNr
				MsgBox, 4144, Addendum für AlbisOnWindows - ScanPool, % Hinweis5
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

	;{ Signierfenster schließen und die folgenden Speicherdialage ebenso automatisch abschließen

		;jetzt das Signaturfenster schliessen
			while, WinExist("ahk_id " hDokSig) {
				VerifiedClick("Button5", "", "", hDokSig)
				sleep, 50
				If (A_Index > 10)
					MsgBox,,Addendum für AlbisOnWindows - ScanPool - Info, Das Eintragen des Kennwortes hat nicht funktioniert.`nBitte tragen Sie es bitte manuell ein!`nDrücken Sie danach bitte erst auf Ok.
			}

		;Signaturzähler erhöhen, anzeigen und Zähler sichern
			Addendum.SignatureCount ++
			IniWrite, % Addendum.SignatureCount, % AddendumDir "\Addendum.ini", % "ScanPool", % "SignatureCount"
			ToolTip, % "Signature Nr: " Addendum.SignatureCount, 1000, 1, 10

		; Dateidialog Routinen starten
			SetTimer, PDFHelpCloseSaveDialogs	, 100
			SetTimer, PDFNotRecentlySigned		, -10000
	;}

		PraxTT("'", "off 2")

return

PDFHelpCloseSaveDialogs:            	;{ - Notfalllösung für die immer noch unsichere Dialogerkennung

		If (hwnd:= WinExist("Speichern unter bestätigen ahk_class #32770 ahk_exe FoxitReader.exe" )) {

			VerifiedClick("Button1", hwnd,,, true)               	; Ja Button drücken

			Addendum.PDFRecentlySigned := false
			SetTimer, PDFHelpCloseSaveDialogs,  Off

			If Addendum.AutoCloseFoxitTab		{
				WinActivate   	, % "ahk_id " FoxitID
				WinWaitActive	, % "ahk_id " FoxitID,, 2
				PraxTT("signiertes Dokument wird geschlossen!'", "0 2")
				Sleep, 500                                                	; kurze Pause um das Neuzeichnen des Foxitreaderfensters abzuwarten
				result := FoxitInvoke("Close", FoxitID)
				PraxTT("", "off")
			}

			return FoxitID

		}
		else if (hwnd:= WinExist("Speichern unter ahk_class #32770 ahk_exe FoxitReader.exe")) {

			FoxitID := GetParent(hwnd)                            	; Handle des zugehörigen FoxitReader Fenster
			VerifiedClick("&Speichern", hwnd,,, true)               	; Speichern Button drücken
			return FoxitID

		}


return ;}

PDFNotRecentlySigned:
		Addendum.PDFRecentlySigned := false
		SetTimer, PDFHelpCloseSaveDialogs,  Off
		Tooltip,,,, 10
return
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

JEE_StrUtf8BytesToText(vUtf8) {
	if A_IsUnicode
	{
	VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
	StrPut(vUtf8, &vUtf8X, "CP0")
	return StrGet(&vUtf8X, "UTF-8")
	}
	else
	return StrGet(&vUtf8, "UTF-8")
}



;}

;----------------------------------------------------------------------------------------------------------------------------------------------

;{20. Gui's, Icons

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
Create_Addendum_ico() {                                                                                               	;-- erstellt das Trayicon
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAD/////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBf/oo//oo//oo//oo/tl4VTMyJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/6noxCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBeJVkT+oo7/oo//oo//oo//oo+kZ1VCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd7TTr/oo//oo//oo//oo//oo9CKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdDKRj/oo//oo//oo//oo//oo//oo9GKhpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo/+oY7/oo//oo//oo//oo+PWUdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRTP/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo/+oY7/oo//oo//oo/7oI11SDZCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdnQC/0nIn/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9aNyZCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdeOyn/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9LLh1CKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdYNiX/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo+QWkiQWkiQWkiQWkiQWkiRW0n/oo//oo//oo//oo//oo+QWkiQWkiQWkiQWkiQWkiQWkj/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo//oo+1cmC1cmC1cmC3c2H/oo//oo//oo//oo//oo+1cmC1cmC1cmC2c2H/oo//oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+0cWBCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdILBrkkX7/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo9UNCNCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBeeY1H/oo//oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+rbFpCKBdCKBdCKBf/oo//oo//oo/+oY7/oo9CKBdCKBdGKhr/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd3Sjn/oo//oo//oo//oo//oo9NLx5CKBdCKBdlPy3/oo//oo//oo9DKRhCKBdCKBeWXk3/oo//oo//oo/7oI13SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+gZVJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhn/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdRMiH/oo//oo//oo//oo//oo9GKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBeRWkn+oY7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+WXkxCKBdCKBdCKBdCKBdCKBdCKBdDKRj/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdVNCP/oo//oo//oo//oo//oo9EKhlCKBdCKBdCKBdCKBdCKBeJVkT+oY7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo+MWEZCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdYNiX/oo//oo//oo//oo//oo9DKRhCKBdCKBdCKBeEUkD9oI7/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/9oI2BUD9CKBdCKBdCKBf/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdcOSf/oo//oo//oo//oo//oo9CKBdCKBdCKBf7n43/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo/7oI13SjlCKBd3Sjn/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdhOyr/oo//oo//oo//oo//oo+HVEP7n4z/oo//oo//oo93SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo/7oIz/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPy3/oo//oo//oo//oo//oo//oo//oo/7oYx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdoQS//oo//oo//oo//oo//oo/7oYx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRTP/oo//oo//oo/7oIx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdxRzX5n4r7oIx3SjlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="

	If (!DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0))
		return False
	VarSetCapacity(Dec, DecLen, 0)
	If (!DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0))
		return False
	pStream := DllCall("Shlwapi.dll\SHCreateMemStream", "Ptr", &Dec, "UInt", DecLen, "UPtr")
	DllCall("Gdiplus.dll\GdipCreateBitmapFromStreamICM", "Ptr", pStream, "PtrP", pBitmap)
	DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
	ObjRelease(pStream)

	return hBitmap
}

Create_BefundImport_ico() {                                                                                          	;-- Icon für das InfoWindow
VarSetCapacity(B64, 788 << !!A_IsUnicode)
B64 := "iVBORw0KGgoAAAANSUhEUgAAABQAAAAYCAYAAAD6S912AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAACtQAAArUBrcsL2AAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAHLSURBVDiNzZM9ixpRFIaPcU68gkHRhAUhBAbtLLYbG6spbWysFMRmy2kt/QE2tmkE21TWUZRAWjsbBZlK0WSsLgyjszPvFlmXbHb8yhrIA7c559znnvsVIiLQFVGIiIQQVK/XqVqtXiwAQK1Wi4bDIUkpf8V6vR5SqRQajQYuwXEcVCoVqKqKfr+P/W5hmiYmkwnS6TQMw4Dv+ydllmWhUCggn89jvV7DNM3nQgCYz+dQVRW1Wg2u6x6UzWYzZLNZlMtl2LYNAMFCAFgul8jlciiVSnAc54VsMBggkUjAMAx4nvcUPygEgM1mA03ToOs6pJRP8U6ng1gshm63+2Kho0IAkFJC13VomgbLstBsNpFMJjEajQKP4aQQAGzbRrFYRDweRyaTwXQ6Daw7WwgAu90O7XYblmUdrLlIeC574Zu/+l9HuLpQISIaj8e0WCxeJVqtVkREFFIUZRGJRN6+vjei7Xa7u4bnGaEjuRtmXgUlXNe9IaIfQbn//5b/zbP5gw9CiHee570/NImZP4XD4ZjjOJKIfp7q8N73/Tsi+nakke+e5zXOanlPNBr9yMyfmfmemfE4fGb+IoRQL5L9DjPfMvPXx3F7qv4BASqDqpfUXLIAAAAASUVORK5CYII="
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

Create_PDF_ico(NewHandle := False) {                                                                         	;-- PDF Icon für Befundfenster
VarSetCapacity(B64, 1192 << !!A_IsUnicode)
B64 := "AAABAAEAEBAAAAEAGABoAwAAFgAAACgAAAAQAAAAIAAAAAEAGAAAAAAAAAAAAAkAAAAJAAAAAAAAAAAAAAAFBUAKCoEEBDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKCoAHB2ENDacCAhsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBDANDa0QEMsKCoYBAQYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAhoKCnoQENAKCoAEBDIBAQcAAAAAAAAAAAAAAAAAAAAAAAADAyYFBUcCAh4AAAAAAAAAAAIHB2APD8cMDKENDa8ICGsGBkkDAyoBAQ0CAhoJCXANDa0KCoUJCXsAAAAAAAAAAAACAiAPD8EBAQ4DAyIGBkoJCXANDa0PD8YPD8UQENEPD8EPD74GBlAAAAAAAAAAAAAAAAALC44EBDUAAAAAAAAAAAAHB2APD74GBkoFBT4EBC0CAhoAAAAAAAAAAAAAAAAAAAAHB1gHB2EAAAAAAAAHB1gNDbUDAyoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEBDANDaEAAAMGBk4NDbADAyYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQwPD8AICGcPD7sDAygAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALC44QEM0EBDkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQoNDawICGcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBkcQENEHB1wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQMDKENDasHB1cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAh0PD74JCXkFBUQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMHB1sJCXEBAQ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf/wAAD/8AAAf/AACB+AAAwAAAAOAAAADzgQAA8x8AAPA/AADwfwAA+P8AAPH/AADx/wAA4f8AAOH/AADh/wAA"
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

Create_Image_ico(NewHandle := False) {                                                                       	;-- Bilder Icon für Befundfenster
VarSetCapacity(B64, 1124 << !!A_IsUnicode)
B64 := "AAABAAEAEA8AAAEAGAA0AwAAFgAAACgAAAAQAAAAHgAAAAEAGAAAAAAAAAAAAAgAAAAIAAAAAAAAAAAAAABNTU0REREREREREREEBAQAAAAAAAAAAAAAAAAAAAALCwsRERERERERERERERFQUFBPT08AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACmpqapqakAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABERET////6+vofHx8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEBDd3d3///////+VlZUAAAAAAAAAAAAAAAAAAAAAAAAAAACSkpKampoKCgoEBASxsbH////////////39/caGhoAAAAAAAAAAAAAAAAAAABycnL////////f39/W1tb///////////////////+Tk5MAAAAAAAAAAAAAAABHR0f8/Pz////////////////////////////////////5+fkfHx8AAAAAAAAfHx/r6+v///////////////////////////////////////////+ampoAAAAJCQnOzs7////////////////d3d3a2tr////////////////////////9/f1paWm0tLT////////////5+flSUlIAAAAAAABGRkb19fX///////////////////////////////////////+bm5sAAAAAAAAAAAAAAAB8fHz///////////////////////////////////////9zc3MAAAAAAAAAAAAAAABMTEz///////////////////////////////////////+2trYAAAAAAAAAAAAAAACRkZH///////////////////////////////////////////+YmJgXFxcNDQ11dXX8/Pz///////////////////////////////////////////////////////////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
;return ImageFromBase64(false, B64)
return GdipCreateFromBase64(B64)
}

;}

;{21. Das Ende - OnExit Sprunglabel und OnError-Funktion

DasEnde(ExitReason, ExitCode) {

	OnExit("DasEnde", 0)

	FormatTime, time, % A_Now, dd.MM.yyyy HH:mm:ss
	FileAppend, % "Skript: " A_ScriptName ", time: " time  ", Client: " A_ComputerName ", ExitReason: " ExitReason ", ExitCode: " ExitCode "`n", % AddendumDir "\logs'n'data\OnExit-Protokoll.txt"

	For i, hook in Addendum.Hooks
		UnhookWinEvent(hook.hEvH, hook.HPA)

	SciTEOutput("ShowLinesOut")

	counter := Fehlerprotokoll("return counter", 0)
	SciTEOutput("Fehlermeldungen seit letztem Addendumstart: " counter.same + counter.different "`n")


	gosub istda

}

istda: ;{

	If onlyReload
		AnzeigeText:= "Addendum wird neu gestartet`.`.`.`."
	else
		AnzeigeText:= "Addendum wird beendet`."

	Progress, B2 B2 cW202842 cBFFFFFF cTFFFFFF zH25 w400 WM400 WS500, %AnzeigeText% , Addendum für AlbisOnWindows , Praxomat_st, Futura Bk Bt
	Loop 50 {
		Progress % (100 - (A_Index*2))
		Sleep, 1
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
#Include %A_ScriptDir%\..\..\lib\class_JSONFile.ahk
#Include %A_ScriptDir%\..\..\lib\class_TCP-UDP.ahk
#Include %A_ScriptDir%\..\..\lib\crypt.ahk
#Include %A_ScriptDir%\..\..\lib\Explorer_Get.ahk
#Include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\lib\TrayIcon.ahk
#Include %A_ScriptDir%\..\..\lib\Sci.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
;#Include %A_ScriptDir%\..\..\lib\WatchFolder.ahk

;------------ Addendumbibliotheken für Albis on Windows --------
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Ini.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Misc.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
#Include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk

;------------ Adddendum.ahk zusätzliche Funktionen /Labels -----
#Include %A_ScriptDir%\include\Addendum_Gui.ahk
#Include %A_ScriptDir%\include\Addendum_PopUpMenu.ahk
#include %A_ScriptDir%\include\ClientLabels.ahk
#include %A_ScriptDir%\include\MenuSearch.ahk
#include %A_ScriptDir%\include\NoTrayOrphans.ahk

;------------ HOTSTRINGS ---------------------------------------------
;#Include %A_ScriptDir%\include\AutoDiagnosen.ahk

;}





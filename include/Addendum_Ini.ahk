; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                 liest Informationen aus der Addendum.ini ein für das globale Objekt Addendum
;                                                                              	!diese Bibliothek enthält Funktionen für Einstellungen des Addendum Hauptskriptes!
;                                                            	by Ixiko started in September 2017 - last change 16.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

return
admObjekte() {                                                                                       	;-- Addendum-Objekt erweitern

 		Addendum.Chroniker            	:= Array()          	; Patient ID's der Chroniker
		Addendum.Geriatrisch           	:= Array()         	; Patient ID's für geriatrisches Basisassement
		Addendum.Default                  	:= Object()       	; Addendum Font/Farbeinstellungen
		Addendum.Hooks                  	:= Object()       	; enthält Adressen der Callbacks der Hookprozeduren und anderes
		Addendum.Labor	    	        	:= Object()			; Laborabruf und andere Daten für Laborprozesse
		Addendum.Laborabruf            	:= Object()         	; Statusinformationen zum Laborabruf
		Addendum.OCR                     	:= Object()          	; Einstellungen für Texterkennung und Autonaming
		Addendum.PDF                      	:= Object()          	; Einstellungen/Pfade für die Bearbeitung von PDF Dateien
		Addendum.LAN	                    	:= Object()			; LAN Kommunikationseinstellungen
		Addendum.LAN.Clients          	:= Object()          	; LAN Netzwerkgeräte Einstellungen
		Addendum.Praxis                     	:= Object()        	; Praxisadressdaten
		Addendum.Praxis.Sprechstunde	:= Object()       	; Daten zu Öffnungszeiten, Urlaubstagen
		Addendum.Praxis.Email          	:= Array()         	; Praxis Email Adressen
		Addendum.Praxis.Urlaub         	:= Array()          	; Praxis Urlaubsdaten
		Addendum.Telegram             	:= Object()        	; Datenobject für einen oder mehrere Telegram-Bots
		Addendum.MsgGui               	:= Object()        	; Gui für Interskript Kommunikation und den Shellhook
		Addendum.Drucker               	:= Object()          	; verschiedene Druckereinstellungen
		Addendum.AktuellerTag	    	:= A_DD            	; der aktuelle Tag
		Addendum.Kosten                  	:= Object()        	; hinterlegte Preise für Hotstringausgaben bei Privatabrechnungen
		Addendum.Thread                  	:= Object()          	; enthält die PID's gestarteter Threads (z.B. AddendumToolbar, Addendum_OCR.ahk)
		Addendum.Threads                 	:= Object()         	; Skriptcode für Prozess-Threads
		Addendum.Windows                	:= Object()        	; Fenstereinstellungen für andere Programme
		Addendum.UndoRedo             	:= Object()          	; Undo/Redo Textbuffer
		Addendum.UndoRedo.List        	:= Array()           	; Undo/Redo Textbuffer
		Addendum.Module                  	:= Object()        	; Addendum Skriptmodule
		Addendum.Tools                     	:= Object()			; externe Programme die z.B. über das Infofenster gestartet werden können
		Addendum.PatExtra                	:= Object()          	; zusätzliche Informationen über Patienten


		Addendum.PraxTTDebug := false

}

admVerzeichnisse() {                                                                             	;-- Programm- und Datenverzeichnisse (!ALS ERSTES AUFRUFEN! nach admObjekte)

		Addendum.Dir                              	:= AddendumDir
		Addendum.Ini                                 	:= AddendumDir "\Addendum.ini"

		Addendum.DBPath                          	:= IniReadExt("Addendum" 	, "AddendumDBPath"     	)              	; Datenbankverzeichnis
		Addendum.BefundOrdner                	:= IniReadExt("ScanPool"     	, "BefundOrdner"            	)             	; BefundOrdner = Scan-Ordner für neue Befundzugänge
		Addendum.VideoOrdner                	:= IniReadExt("ScanPool"     	, "VideoOrdner"             	)             	; Video's z.B. aus Dicom CD's (CT Bilder u.a.)
		Addendum.AlbisExe                         	:= IniReadExt("Albis"           	, "AlbisExe"                    	)             	; Pfad zum Albis-Stammverzeichnis
		Addendum.DosisDokumentenPfad   	:= IniReadExt("Addendum" 	, "DosisDokumentenPfad"	)              	; MS Word Dateien mit eigenen Hinweisen zu Medikamentendosierungen

		Addendum.TPPath	                        	:= Addendum.DBPath "\Tagesprotokolle\" A_YYYY                     	; Tagesprotokollverzeichnis
		Addendum.TPFullPath                      	:= Addendum.TPPath "\" A_MM "-" A_MMMM "_TP.txt" 	               	; Name des aktuellen Tagesprotokolls

		Addendum.ExportOrdner                	:= IniReadExt("ScanPool"   	, "ExportOrdner"        	, Addendum.BefundOrdner "\Export")  ;### ÜBERPRÜFUNG ANLEGEN
		Addendum.AdditionalData_Path       	:= AddendumDir "\include\Daten"


	; Ordnerstruktur anlegen für die Verwaltung eingegangener Befunde, wenn ein BefundOrder in der ini angegeben wurde
		If (!InStr(Addendum.BefundOrdner, "Error") && StrLen(Addendum.BefundOrdner) > 0) && isFullFilePath(Addendum.BefundOrdner) {

			If !InStr(FileExist(Addendum.BefundOrdner "\Text"), "D")
				FileCreateDir, % Addendum.BefundOrdner "\Text"

			If !InStr(FileExist(Addendum.BefundOrdner "\Backup"), "D")
				FileCreateDir, % Addendum.BefundOrdner "\Backup"

			If !InStr(FileExist(Addendum.BefundOrdner "\Backup\Importiert"), "D")
				FileCreateDir, % Addendum.BefundOrdner "\Backup\Importiert"

		}

	; Addendum Datenverzeichnis prüfen
		If InStr(Addendum.DBPath, "Error")	|| !isFullFilePath(Addendum.DBPath)                  	{

		; Anlegen des Verzeichnisses falls nicht vorhanden
			If !FilePathExist(Addendum.Dir "\logs'n'data\_DB")
				If !FilePathCreate(Addendum.DBPath := Addendum.Dir "\logs'n'data") {
					MsgBox,, Addendum für AlbisOnWindows, %	"Das Verzeichnis für die Datenbank:`n" Addendum.DBPath
																					. 	"`nkonnte nicht angelegt werden.`nDas Skript wird beendet!"

					ExitApp
				}
				else if !FilePathCreate(Addendum.DBPath := Addendum.Dir "\logs'n'data\_DB") {
					MsgBox,, Addendum für AlbisOnWindows, %	"Das Verzeichnis für die Datenbank:`n" Addendum.DBPath
																					. 	"`nkonnte nicht angelegt werden.`nDas Skript wird beendet!"

					ExitApp
				}

		; speichern des Verzeichnisses in der ini
			IniWrite, % "%AddendumDir%\logs'n'data\_DB", % Addendum.Dir "\Addendum.ini", % "Addendum", % "AddendumDBPath"
			Addendum.FirstDBAcess := true
		}

		;Addendum.7ZipDir                        	:= RegRead64("HKEY_CURRENT_USER\Software\7-Zip", "Path64")

}

ChronikerListe() {                                                                                  	;-- Chroniker Leistungskomplexe 03220/03221

	; Addendum.Chroniker := Array() wurde im Addendum Skript definiert

	filepath := Addendum.DbPath "\DB_Chroniker.txt"

	If FileExist(filepath)
		Loop, Parse, % FileOpen(filepath, "r", "UTF-8").Read(), `n, `r
			If RegExMatch(A_LoopField, "^\d+", PatID)
				Addendum.Chroniker.Push(PatID)

}

GeriatrischListe() {                                                                                  	;-- geriatrischer Basiskomplex 03360/03362

	; Addendum.Geriatrisch := Array() wurde im Addendum Skript definiert

	filepath := Addendum.DbPath "\DB_Geriatrisch.txt"

	If FileExist(filepath)
		Loop, Parse, % FileOpen(filepath, "r", "UTF-8").Read(), `n, `r
			If RegExMatch(A_LoopField, "^\d+", PatID)
				Addendum.Geriatrisch.Push(PatID)

}

admDruckerStandards() {                                                                       	;-- Standarddrucker A4/PDF/FAX

		Addendum.Drucker.Standard          	:= IniReadExt(compname, "Drucker_Standard")
		Addendum.Drucker.StandardA4     	:= IniReadExt(compname, "Drucker_Standard_A4")
		Addendum.Drucker.StandardA4Tray 	:= IniReadExt(compname, "Drucker_Standard_A4_Tray")
		Addendum.Drucker.PDF                   	:= IniReadExt(compname, "Drucker_PDF")
		Addendum.Drucker.FAX                   	:= IniReadExt(compname, "Drucker_FAX")

}

admFensterPositionen() {                                                                         	;-- Fensterposition anderer Prozesse (ifap...)

	; lädt individuelle Fensterpositionierungen aus der Addendum.ini für Programmfenster
	; Funktion prüft ob Einstellungen vorhanden sind, wenn für ein Programm keine Daten vorhanden sind, wird die Fensterklasse
	; im des ShellhookProc (Addendum.ahk) nicht überprüft durch Entfernung des key:value Paares aus dem .Windows.Proc Objekt.
	; es lassen sich prinzipiell für jedes von Ihnen genutzte Programm Einstellungen anlegen. Im Moment muss dies aber noch manuell erfolgen.

	; Monitordimension
		Addendum.MonSize  	:= A_ScreenWidth >= 1920 ? 2 : 1
		Addendum.Resolution  	:= A_ScreenWidth >= 1920 ? "4k" : "2k"

	; ein Objekt mit den Bezeichnungen der Fensterklassen die automatisch positioniert werden sollen
		Addendum.Windows.Proc  	:= {	"TForm_ipcMain"            	: "Ifap"                ; key=classNN : value=key der Einstellungen
														,	"classFoxitReader"             	: "Foxit"
														,	"SUMATRA_PDF_FRAME" 	: "Sumatra"
														,	"OpusApp"                    	: "MSWord" }

	; for in loop zum Auslesen der Einstellungen aus der ini
		For classnn, appname in Addendum.Windows.Proc {

			tmp1 := Trim(IniReadExt(compname, "AutoPos_" appname, "2k:nein,4k:nein"))
			If !InStr(tmp1, "ERROR") {

				tmp2 := Trim(IniReadExt(compname, "Position_" appname, "2k[1,100,100,1600,1000],4k[1,100,100,1600,1000]"))
				Addendum.Windows[appname]             	:= Object()
				Addendum.Windows[appname]["AutoPos"]	:= RegExAutoPos(tmp1)
				Addendum.Windows[appname]["2k"]         	:= RegExScreenSize(tmp2, "2k")
				Addendum.Windows[appname]["4k"]         	:= RegExScreenSize(tmp2, "4k")

			}
			else
				Addendum.Windows.Proc.Delete(classnn)

		}

}

admFunktionen() {                                                                                 	;-- zuschaltbare Funktionen

	; Addendum.OptF. ???

	; automatische Positionierung von Albis ist an
		tmp  := Trim(IniReadExt(compname, "AutoSize_Albis", "2k:nein,4k:nein" ))
		Addendum.AlbisLocationChange  	:= RegExAutoPos(tmp)
	; Größeneinstellung
		Addendum.AlbisGroesse            	:= IniReadExt(compname	, "Position_Albis"                               	, "w1922 h1080"	)
	; Automatisierung der GVU Formulare
		Addendum.GVUAutomation			:= IniReadExt(compname	, "GVU_automatisieren"                    	, "nein"              	)
	; Automatisierung PDF Signierung
		Addendum.PDFSignieren   	    	:= IniReadExt(compname	, "FoxitPdfSignature_automatisieren" 	, "nein"              	)
	; Tastaturkürzel zum Auslösen des Signaturvorganges
		Addendum.PDFSignieren_Kuerzel	:= IniReadExt(compname	, "FoxitPdfSignature__Kuerzel" 	      	, "^Insert"            	)
	; signiertes Dokument automatisch schliessen
		Addendum.AutoCloseFoxitTab   	:= IniReadExt(compname	, "FoxitTabSchliessenNachSignierung"	, "Ja"                 	)
	; zeitgesteuerter Laborabruf
		Addendum.Labor.AbrufTimer         := IniReadExt(compname	, "Laborabruf_Timer"           	   	    	, "nein"              	)
		If InStr(Addendum.Labor.AbrufTimer, "Error") || (StrLen(Addendum.Labor.AbrufTimer) = 0)
			Addendum.Labor.AbrufTimer	:= false
	; Laborabruf bei manuellem Start automatisieren
		Addendum.Labor.AutoAbruf 		    := IniReadExt(compname	, "Laborabruf_automatisieren"         	, "nein"              	)
		If InStr(Addendum.Labor.AutoAbruf, "Error") || (StrLen(Addendum.Labor.AutoAbruf) = 0)
			Addendum.Labor.AutoAbruf	:= false
	; Infofenster das den Inhalt des Bildvorlagen-Ordners anzeigt
		Addendum.AddendumGui          	:= IniReadExt(compname	, "Infofenster_anzeigen"                  	, "ja"                 	)
	; Logbuch für Aktionen in der Karteikarte
		Addendum.PatLog                      	:= IniReadExt(compname	, "Logbuch_Patienten"                     	, "ja"                 	)
	; Hilfe beim Export von Dicom CD's im MicroDicom Programm
		Addendum.mDicom                     	:= IniReadExt(compname	, "MicroDicomExport_automatisieren"	, "nein"                 	)
	; PopUpMenu integrieren
		Addendum.PopUpMenu              	:= IniReadExt(compname	, "PopUpMenu"                              	, "ja"                 	)
	; die Addendum Toolbar starten
		Addendum.ToolbarThread            	:= IniReadExt(compname	, "Toolbar_anzeigen"                         	, "nein"                 	)
	; ermöglicht die sofortige Bearbeitung/Anzeige neuer Dateien
		Addendum.WatchFolder            	:= IniReadExt(compname	, "BefundOrdner_ueberwachen"       	, "ja"                 	)
	; Schnellrezept integrieren
		Addendum.Schnellrezept           	:= IniReadExt("Addendum"	, "Schnellrezept"                              	, "ja"                 	)
	; zeigt TrayTips
		Addendum.ShowTrayTips           	:= IniReadExt("Addendum"	, "TrayTips_zeigen"                            	, "nein"                 	)
	; AutoOCR - Eintrag wird nur auf dem dazu berechtigten Client angezeigt
		Addendum.OCR.AutoOCR            	:= IniReadExt("OCR"      	, "AutoOCR"                                  	, "nein"                 	)


}

admGesundheitsvorsorge() {                                                                    	;-- Abstände zwischen Untersuchungen minimales Alter

	Addendum.GVUAbstand                 	:= IniReadExt("Abrechnungshelfer"	, "minGVUAbstand")
	Addendum.GVUminAlter                	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlter")
	Addendum.GVUminAlterEinmalig   	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterEinmalig")
	Addendum.GVUminAlterMehrmalig	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterMehrmalig")

}

admPIDHandles() {                                                                                   	;-- Prozess-ID's

		Addendum.AlbisPID                        	:= AlbisPID()
		Addendum.scriptPID                       	:= DllCall("GetCurrentProcessId")          	; die ScriptPID wird für das SkriptReload und die Interskript-Kommunikation benötigt

}

admInfoWindowSettings() {                                                                      	;-- AddendumGui/Infofenster

		tmp1 := IniReadExt(compname, "InfoFenster_Position", "y2 w400 h340")
		RegExMatch(tmp1, "y\s*(\d+)\s*w\s*(\d+)\s*h\s*(\d+)", match)

		Addendum.iWin := {	"Y"                     	: match1
									, 	"W"                    	: match2
									,	"ReIndex"            	: false
									,	"rowprcs"           	: 0
									,	"Init"                    	: false
									,	"ConfirmImport"   	: IniReadExt(compname, "Infofenster_Import_Einzelbestaetigung"  	, "Ja")
									,	"firstTab"           	: IniReadExt(compname, "Infofenster_aktuelles_Tab"                     	, "Patient")
									,	"TProtDate"         	: IniReadExt(compname, "Infofenster_Tagesprotokoll_Datum"       	, "Heute")
									,	"JournalSort"      	: IniReadExt(compname, "Infofenster_JournalSortierung"               	, "3 1 1")
									,	"LBDrucker"       	: IniReadExt(compname, "Infofenster_Laborblatt_Drucker"               	, "")
									,	"LVScanPool"        	: {"W"	: (match2 - 10)
																				,	"R"	: IniReadExt(compname, "InfoFenster_BefundAnzahl", 7)}}

		Addendum.iWin.LISTENERS := {	"WM_CREATE"                          	: 0x01
														, 	"WM_DESTROY"                     	: 0x02
														, 	"WM_SIZE"                              	: 0x05
														, 	"WM_ACTIVATE"                      	: 0x06
														,	"WM_SETREDRAW"                  	: 0x0B
														,	"WM_PAINT"                          	: 0x0F
														, 	"WM_CLOSE"                          	: 0x10
														, 	"WM_ERASEBKGND"              	: 0x14
														,	"WM_SHOWWINDOW"          	: 0x18
														, 	"WM_CHILDACTIVATE"          	: 0x22
														,	"WM_NOTIFY"                          	: 0x4E
														,	"WM_STYLECHANGED"          	: 0x7D
														,	"WM_NCPAINT"                      	: 0x85
														, 	"WM_NCCREATE"                  	: 0x81
														,	"WM_SYNCPAINT"                  	: 0x88
														, 	"WM_NCUAHDRAWCAPTION"	: 0xAE
														, 	"WM_NCUAHDRAWFRAME"      	: 0xAF
														, 	"WM_HSCROLL"                      	: 0x114
														,	"WM_VSCROLL"                      	: 0x115
														, 	"WM_CHANGEUISTATE"          	: 0x127
														,	"WM_UPDATEUISTATE"          	: 0x128
														, 	"WM_MOVING"                      	: 0x216
														,	"WM_PARENTNOTIFY"              	: 0x210
														, 	"WM_MDIACTIVATE"              	: 0x222
														, 	"WM_CLEAR"                          	: 0x303}

	; zuletzt angezeigter Tab
		thisT := Addendum.iWin.firstTab
		If InStr(thisT, "ERROR") || (StrLen(thisT) = 0)
			Addendum.iWin.firstTab := "Patient"
		If RegExMatch(thisT, "(\w+)\s+(\d+)", tmp) {
			Addendum.iWin.firstTab     	:= tmp1
			Addendum.iWin.firstTabPos	:= tmp2
		}
		else
			Addendum.iWin.firstTabPos	:= 0

	; Tab - Tagesprotokoll
		thisT := Addendum.iWin.TProtDate
		If InStr(thisT, "ERROR") || (StrLen(thisT) = 0)
			Addendum.iWin.TProtDate := "Heute"

	; Tab - Patient
		Addendum.iWin.AbrHelfer	:= IniReadExt(compname, "Infofenster_Abrechnungshelfer", "ja"	)  	; Abrechnungshelfer anzeigen

	Addendum.ImportRunning 	:= false
	Addendum.iWin.Init	:= 0

}

admModule() {                                                                                       	;-- Liste der verfügbaren Module

	; Einstellungen welche Module auf diesem Client benutzt werden dürfen
		admModule := IniReadExt(compname, "Module")
		For index, modulNr in StrSplit(admModule, ",") {
			iniModul := IniReadExt("Module", "Modul" SubStr("00" modulNr, -1))
			If InStr(iniModul, "Error")
				continue
			m 	:= StrSplit(iniModul, "|")
			cmd	:= !RegExMatch(m.3, "^[A-Z]\:") ? AddendumDir "\" m.3 : m.3
			ico	:= !RegExMatch(m.4, "^[A-Z]\:") ? AddendumDir "\" m.4 : m.4
			Addendum.Module.Push({"name":m.2, "command":cmd, "ico":ico})
		}

}

admTools() {                                                                                          	;-- Liste der verfügbaren externen Programme

		Loop {
			iniTool := IniReadExt("Module", "Tool" SubStr("00" A_Index, -1))
			If InStr(iniTool, "Error")
				break
			m 	:= StrSplit(iniTool, "|")
			cmd	:= !RegExMatch(m.3, "^[A-Z]\:") ? AddendumDir "\" m.3 : m.3
			ico	:= !RegExMatch(m.4, "^[A-Z]\:") ? AddendumDir "\" m.4 : m.4
			Addendum.Tools.Push({"name":m.2, "command":cmd, "ico":ico})
		}

}

admLaborDaten() {                                                                                	;-- Laborabruf, Verarbeitung Laborwerte

		Addendum.Laborabruf.Status	:= false
		Addendum.Laborabruf.Daten:= false
		Addendum.Laborabruf.Voll  	:= false

		Addendum.Labor.LDTDirectory     	:= IniReadExt("LaborAbruf"	, "LDTDirectory"                        , "C:\Labor"	)
		Addendum.Labor.ExecuteOn       	:= IniReadExt("LaborAbruf"	, "OnlyRunOnClient"                     	    , ""    	)  	; auf welchem/n Client/s der Laborabruf ausgeführt werden darf
		Addendum.Labor.ExecuteFile       	:= IniReadExt("LaborAbruf"	, "Laborabruf_Extern"                    	    , ""    	)  	; externes Programm das für den Abruf ausgeführt werden muss
		Addendum.Labor.ExecuteScript     	:= IniReadExt("LaborAbruf"	, "Laborabruf_Skript"                    	    , ""    	)  	; Skript das den Abruf übernimmt
		Addendum.Labor.LaborName       	:= IniReadExt("LaborAbruf"	, "LaborName"                             	    , ""    	)  	; falls sie mehrere Labore haben, tragen Sie das aktuelle hier ein
		Addendum.Labor.Laborkuerzel     	:= IniReadExt("LaborAbruf"	, "Aktenkuerzel"                            	    , "labor")  	; Karteikartenkürzel um Informationen ablegen zu können
		Addendum.Labor.Alarmgrenze     	:= IniReadExt("LaborAbruf"	, "Alarmgrenze"                            	    , "30%"	)  	; Alarmierungsgrenze in Prozent oberhalb der Normgrenzen
		                           AbrufZeiten      	:= IniReadExt("LaborAbruf"	, "LaborAbruf_Zeiten"                    	    , ""    	)	; z.B. "06:00, 15:00, 19:00, 21:00"
		Addendum.Labor.Kennwort       	:= IniReadExt("LaborAbruf"	, "LaborKennwort"                        	    , ""    	)	; für order&entry per CGM-Channel
		If InStr(Addendum.Labor.Kennwort, "Error")
			Addendum.Labor.Kennwort := ""

	; Laborabrufzeiten umwandeln in HHMM Format
		Addendum.Labor.AbrufZeiten := Array()
		spos := 1
		while (spos := RegExMatch(AbrufZeiten, "(?<H>\d+)\:(?<M>\d+)", timer, spos)) {
			spos += StrLen(timer)
			Addendum.Labor.AbrufZeiten.Push("T" SubStr("00" timerH, -1) SubStr("00" timerM, -1) "00") ; eine Null als erste Zahl wird von AHK leider entfernt
		}

		; ACHTUNG:
		; hier werden die Einstellungen für eine Alarmierung über Ihren eigenen Telegram-Bot eingelesen! Denken Sie immer daran die Telegram Token und ChatID's vor dem
		; Zugriff Fremder oder auch dem eigenen Personal zu schützen
		Addendum.Labor.TGramOpt        	:= IniReadExt("LaborAbruf", "TGramOpt"           	    , "nein")	        	; die Nummer ihres Telegram Bots und + oder - Tel z.b. "1+Tel"
		If !((Addendum.Labor.TGramOpt = false) || RegExMatch(Addendum.Labor.TGramOpt, "^\d+\+|\-Tel")) 	; die Werte müssen einen exakten Syntax haben, sonst wird die Telegram Option gelöscht!
			IniWrite, % "nein", % Addendum.Ini, % "LaborAbruf", % "TGramOpt"

}

admLanProperties() {                                                                              	;-- Netzwerkeinstellungen / Client PC Namen / IP-Adressen

	; prüft auf den korrekt gesetzten Computernamen
		CName		:= IniReadExt(compname, "ComputerName")
		If (CName <> A_ComputerName)
			IniWrite, % A_ComputerName, % Addendum.Ini, % compname, % "ComputerName"

	; hält die in der ini gespeicherte IP zu jedem Client aktuell
		CompIP 	:= A_IPAddress1
		storedIP    	:= IniReadExt(compname, "IP")
		If (CompIP <> storedIP)
			IniWrite, % CompIP, % Addendum.Ini, % compname, % "IP"

	; liest die Namen aller Clients ein welche überwacht werden sollen oder welche miteinander kommunizieren dürfen
		For idx, clientName in StrSplit(IniReadExt("Computer", "Computer"), "|")
			Addendum.LAN.Clients[clientName] := {	"computername"	: IniReadExt(clientName, "ComputerName")
													                	,	"remoteShotDown"	: IniReadExt(clientName, "AutoShutDown")
													                	,	"rdpPath"            	: IniReadExt(clientName, "rdpPath")
													                	,	"IP"                    	: IniReadExt(clientName, "IP") }

}

admMailAndHolidays() {                                                                           	;-- eigene Mailadresse und Urlaubsdaten

	Loop
		If !InStr((tmp :=  IniReadExt("Allgemeines", "EMail" A_Index)), "ERROR")
			Addendum.Praxis.Email.Push(tmp)
		else
			break

	Loop
		If !InStr((tmp :=  IniReadExt("Allgemeines", "Urlaub" A_Index)), "ERROR")
			Addendum.Praxis.Urlaub.Push(tmp)
		else
			break

}

admMonitorInfo() {                                                                                  	;-- ermittelt die Anzahl der angeschlossenen Monitore und schreibt die Werte in die ini

	; Werte werden unter info des Infofenster angezeigt
	; benötigt \include\Addendum_Screens.ahk

	Addendum.Monitor := Array()

	SysGet, monitorCount, MonitorCount
	Loop, % monitorCount {

		mon := screenDims(A_Index)
		Addendum.Monitor.Push("x" mon.X " y" mon.Y " w" mon.W " h" mon.H) ; ", Orientation: " mon.orient ", yEcke: " mon.yEdge ", yRand: " mon.yBorder
		monDim := IniReadExt(compname, "Monitor" A_Index, Addendum.Monitor[A_Index])
		If !IsRemoteSession() && (monDim <> Addendum.Monitor[A_Index])
			IniWrite, % Addendum.Monitor[A_Index], % Addendum.Ini, % compname, % "Monitor" A_Index

	}

}

admPDFSettings()  {                                                                                	;-- PDF anzeigen/signieren

	; PDF Reader: 2 Programme sind möglich. Der FoxitReader wird nur für's Signieren verwendet (Startvorgang dauert zu lange)
		Addendum.PDF.Reader                       	:= IniReadExt("ScanPool"     	, "PDFReader"                             	, AddendumDir "\include\FoxitReader\FoxitReaderPortable\FoxitReaderPortable.exe")
		Addendum.PDF.ReaderName               	:= IniReadExt("ScanPool"     	, "PDFReaderName"     	            	, "FoxitReader")
		Addendum.PDF.ReaderWinClass           	:= IniReadExt("ScanPool"     	, "PDFReaderWinClass"               	, "classFoxitReader")
		Addendum.PDF.ReaderAlternative         	:= IniReadExt("ScanPool"     	, "PDFReaderAlternative"            	, "")
		Addendum.PDF.ReaderAlternativeName	:= IniReadExt("ScanPool"     	, "PDFReaderAlternativeName"      	, "")

		altReader := Addendum.PDF.ReaderAlternative, altName := Addendum.PDF.ReaderAlternativeName
		If (altName = "ERROR") || (altReader = "ERROR") || (StrLen(altName) = 0)	|| (StrLen(altReader) = 0){
				Addendum.PDF.ReaderAlternativeName	:= ""
				Addendum.PDF.ReaderAlternative       	:= ""
		}

	; PDF signieren und Statistik
		Addendum.PDF.SignatureCount          	:= IniReadExt("ScanPool"     	, "SignatureCount")
		Addendum.PDF.SignaturePages          	:= IniReadExt("ScanPool"     	, "SignaturePages")

		If InStr(Addendum.PDF.SignatureCount, "Error") || (StrLen(Trim(Addendum.PDF.SignatureCount)) = 0)
			Addendum.PDF.SignatureCount := 0
		If  InStr(Addendum.PDF.SignaturePages, "Error") || (StrLen(Trim(Addendum.PDF.SignaturePages)) = 0)
			Addendum.PDF.SignaturePages := 0

		Addendum.PDF.SignatureWidth          	:= IniReadExt("ScanPool"     	, "Signature_Breite"                            	, 50)
		Addendum.PDF.SignatureHeight        	:= IniReadExt("ScanPool"     	, "Signature_Hoehe"                            	, 25)
		Addendum.PDF.Ort                              	:= IniReadExt("ScanPool"     	, "Ort"                                               	,, true)
		Addendum.PDF.Grund                       	:= IniReadExt("ScanPool"     	, "Grund"                                           	,, true)
		Addendum.PDF.SignierenAls                 	:= IniReadExt("ScanPool"     	, "SignierenAls"                                  	,, true)
		Addendum.PDF.Darstellungstyp          	:= IniReadExt("ScanPool"     	, "DarstellungsTyp"                              	,, false)
		Addendum.PDF.PasswortOn                	:= IniReadExt("ScanPool"     	, "PasswortOn"                                   	, "nein")
		Addendum.PDF.PatAkteSofortOeffnen  	:= IniReadExt("ScanPool"     	, "Patientenakte_sofort_oeffnen"            	, "ja")
		Addendum.PDF.DokumentSperren        	:= IniReadExt("ScanPool"     	, "DokumentNachDerSignierungSperren", "ja")

}

admOCRSettings() {                                                                                	;-- PDF Bearbeitung

		Addendum.PDF.xpdfPath	                 	:= IniReadExt("OCR" , "xpdfPath"          	, AddendumDir "\include\OCR\xpdf")
		Addendum.PDF.qpdfPath	                	:= IniReadExt("OCR" , "qpdfPath"         	, AddendumDir "\include\OCR\qpdf")
		Addendum.PDF.tessPath	                	:= IniReadExt("OCR" , "tessPath"          	, AddendumDir "\include\OCR\tesseract")
		Addendum.PDF.imagickPath                	:= IniReadExt("OCR" , "imagickPath"      	, AddendumDir "\include\OCR\imagick")

	; welcher Client im Netzwerk die Texterkennung ausführen soll
	; ist kein Client angegeben ist die Funktion ausgeschaltet, ein- und ausschalten läßt sich AutoOCR über das Traymenu
		Addendum.OCR.Client                        	:= IniReadExt("OCR" , "OCRClient")
		If InStr(Addendum.OCR.Client, "Error")                                	; ## Abfrage integrieren
			Addendum.OCR.Client := ""

	; Zeitverzögerung in Sekunden bis zum Start der Texterkennung,
		Addendum.OCR.AutoOCRDelay           	:= IniReadExt("OCR" , "AutoOCR_Startverzoegerung", 30)
		If InStr(Addendum.OCR.AutoOCRDelay, "Error")                   	; ## Abfrage integrieren
			Addendum.OCR.AutoOCRDelay := 30

}

admPraxisDaten() {                                                                                  	;-- Daten der Praxis

	; Sprechstunden werden mit Tagesbezeichnung z.B. .["Montag"] := "07:00-19:00" gespeichert

	Addendum.Praxis.Name            	:= IniReadExt("Allgemeines", "PraxisName")
	Addendum.Praxis.Strasse           	:= IniReadExt("Allgemeines", "Strasse")
	Addendum.Praxis.PLZ                	:= IniReadExt("Allgemeines", "PLZ")
	Addendum.Praxis.Ort                	:= IniReadExt("Allgemeines", "Ort")
	Addendum.Praxis.KVStempel         	:= IniReadExt("Allgemeines", "KVStempel")            ; genutzt für Textersetzung
	Addendum.Praxis.MailStempel      	:= IniReadExt("Allgemeines", "MailStempel")          ; genutzt für Textersetzung
	Addendum.Praxis.Sprechstunde    	:= Sprechstunde(IniReadExt("Allgemeines", "Sprechstunde"))
	RegExMatch(Addendum.Praxis.Sprechstunde[A_DDDD], "(\d+)\:(\d+)$", Uhr)
	Addendum.Praxis.TagesEnde     	:= ((Uhr1*3600) + (Uhr2*60)) * 1000	; ms von 0 Uhr bis zum Ende der Sprechstunde - für Shutdown Timer

	Addendum.Praxis.Arzt := Array()
	Loop {

		Arzt := IniReadExt("Allgemeines", "Arzt" A_Index "Name")
		Fach := IniReadExt("Allgemeines", "Arzt" A_Index "Abrechnung")

		If InStr(Arzt, "Error") || (Arzt = "") {
			Arzt := Addendum.Praxis.Name
			ArztSucheBeenden := true
		}

		If InStr(Abr, "Error") || (Abr = "") {
			Fach := "Hausarzt"
			ArztSucheBeenden := true
		}

		Addendum.Praxis.Arzt.Push({"Name":Arzt, "Fach":Fach})

		If ArztSucheBeenden
			break

	}

	StandardArzt := IniReadExt("Allgemeines", "StandardArzt", 1)
	If (StandardArzt > 0)
		Addendum.Praxis.StandardArzt := StandardArzt


}

admShutDown() {                                                                                    	;-- AutoShutDown Einstellungen

	Addendum.ShutDown_Leerlaufzeit	:= IniReadExt("Addendum", "ShutDown_Leerlaufzeit", "60")
	Addendum.AutoShutDown         	:= IniReadExt(compname, "AutoShutDown", "nein")      	; automatisches Herunterfahren des PC nach Feierabend

}

admSicherheit() {                                                                                    	;-- Lockdown Funktion

		Addendum.STOut     	:= IniReadExt(compname          	, "Verstecke_Alle_Daten"    	     , 60	)
		Addendum.STEnd    	:= IniReadExt(compname           	, "Zeige_Alle_Daten_Hotstring", "xq")

}

admSonstiges() {                                                                                    	;-- sonstige Einstellungen oder Variablen setzen

		local PatID, leistungen, tmp, abrdatum, ziffer

		Addendum.useraway     	:= false                                                    	; true wenn Nutzer einen bestimmten Zeitraum keine Eingaben gemacht hat
		Addendum.FirstDBAcess	:= false                                                    	; flag nur notwendig für den erstmaligen Aufruf von Addendum.ahk

	; lädt Daten für den  Abrechnungshelfer
		Path_PatExtra := Addendum.DBPath "\PatData\PatExtra.json"
		If FileExist(Path_PatExtra) {

			tmp := JSONData.Load(Path_PatExtra,, "UTF-8")
			For PatID, leistungen in tmp
				For ziffer, abrdatum in leistungen
					If (StrLen(Trim(abrdatum)) > 0) {
							If !IsObject(Addendum.PatExtra[PatID])
								Addendum.PatExtra[PatID] := Object()
							Addendum.PatExtra[PatID][ziffer] := abrdatum
					}

		; Daten werden nur bei Änderungen des PatExtra.json Files auf die Festplatte gesichert
			lastPatExtra := IniReadExt("Abrechnungshelfer", "PatExtra_FileTime")
			lastPatExtra := (lastPatExtra = "ERROR" || StrLen(lastPatExtra) = 0) ? "" : Trim(lastPatExtra)
			FileGetTime, PatExtraFileTime, % Path_PatExtra, M
			If (lastPatExtra <> PatExtraFileTime) {
				JSONData.Save(Addendum.DBPath "\PatData\PatExtraXX.json", Addendum.PatExtra, true,, 1, "UTF-8")
				IniWrite, % PatExtraFileTime, % Addendum.Ini, Abrechnungshelfer, PatExtra_FileTime
			}

		}
		else
			PraxTT("File not exist: " Addendum.DBPath "\PatData\PatExtra.json", "4 1")

}

admStandard() {                                                                                    	;-- Einstellungen für Addendum - Fenster und Dialoge

		Addendum.Default.Font                  	:= IniReadExt("Addendum"	, "StandardFont"        	)
		Addendum.Default.BoldFont          	:= IniReadExt("Addendum"	, "StandardBoldFont"    	)
		Addendum.Default.FontSize           	:= IniReadExt("Addendum"	, "StandardFontSize"    	)
		Addendum.Default.FntColor            	:= IniReadExt("Addendum"	, "DefaultFntColor"     	)
		Addendum.Default.BGColor            	:= IniReadExt("Addendum"	, "DefaultBGColor"     	)
		Addendum.Default.BGColor1          	:= IniReadExt("Addendum"	, "DefaultBGColor1"   	)
		Addendum.Default.BGColor2          	:= IniReadExt("Addendum"	, "DefaultBGColor2"   	)
		Addendum.Default.BGColor3          	:= IniReadExt("Addendum"	, "DefaultBGColor3"   	)

}

admTagesProtokoll() {                                                                            	;-- Anlegen/Lesen des aktuellen Tagesprotokolls

	; Anlegen eines neuen Tagesprotollordners falls dieser noch nicht vorhanden
		If !InStr(FileExist(Addendum.TPPath), "D") {
			FileCreateDir, % Addendum.TPPath
			If ErrorLevel 	{
				MsgBox,, Addendum für AlbisOnWindows, Ein neues Verzeichnis für die Tagesprotokolldateien konnte nicht angelegt werden.`nDas Skript muss jetzt beendet werden!
				ExitApp
			}
		}

	; Ermitteln ob ein für den aktuellen Monat ein Tagesprotokoll existiert und wenn ja dann auch Einlesen der Tagesprotokolldaten
		SectionDate  	:= A_DDDD "|" A_DD "." A_MM                                         					;z.B. [Montag|15.10.]
		If FileExist(Addendum.TPFullPath)  {

			IniRead, TPtmp, % Addendum.TPFullPath, % SectionDate, % compname
			If !InStr(TPtmp, "ERROR") {
				For Index, Value in TProtokoll
					If (StrLen(Value) = 0) {
						TProtokoll.RemoveAt(Index)
						continue
					}
				Loop, Parse, TPtmp, >
					If RegExMatch(A_LoopField, "\d+", PatID)
						TProtokoll.Push(PatID)
				}
			}
			else
				FileAppend, % ";Tagesprotokoll Monat " A_MMMM " " A_YYYY " für Albis on Windows.`n", % Addendum.TPFullPath, UTF-8

}

admTelegramBots() {                                                                              	;-- Telegram Bot Einstellungen

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

}

admThreads() {                                                                                       	;-- lädt und startet Threadcode

	; Skriptcode kann bei Änderungen der Flag-Einstellungen dynamisch hinzugeladen werden oder entladen werden
	; geschrieben um RAM-Belegung zu sparen
	; letzte Änderung: 11.10.2020

	; Toolbar-Thread starten oder beenden
		If Addendum.ToolbarThread {

			threadFilePath := Addendum.Dir "\Module\Addendum\threads\AddendumToolbar.ahk"
			If FileExist(threadFilePath) {                                                                                                                	; Skriptdatei vorhanden:

				Addendum.Threads.Toolbar  	:= FileOpen(threadFilePath,"r").Read()                                                    ; Skriptdatei laden
				If !Addendum.Thread["Toolbar"].ahkReady() {                                                                                  	; Toolbar-Thread ist nicht gestartet?
					If !(Addendum.Thread["Toolbar"] := AHKThread(Addendum.Threads.Toolbar)) {                            	; Toolbar-Thread starten, wenn nicht startbar dann
						Addendum.ToolbarThread := false                                                                                        	; das erneute Aufrufen des Threads unterbinden
						PraxTT("Die Addendum Toolbar konnte nicht gestartet werden.", "3 3")			    			    	    	; Information an den Nutzer
						throw Exception("Der Toolbar-Thread konnte nicht gestartet werden.",	-1, threadFilePath)           	; Fehlerausgabe
					}
				}

			} else {                                                                                                                                                	; Skriptdatei nicht vorhanden:

				Addendum.ToolbarThread := false                                                                                                	; das erneute Aufrufen des Threads unterbinden
				failpath := StrReplace(threadFilePath, Addendum.Dir)
				failpath	:= StrReplace(failPath, "\AddendumToolbar.ahk")
				PraxTT(	"Die Datei: AddendumToolbar.ahk ist nicht vorhanden.`n"                                                    	; Information zum Fehler für den Nutzer
						. 	"Verzeichnis: " failpath, "3 3")
				throw Exception("Die Datei: AddendumToolbar.ahk ist nicht vorhanden.",	-1, failPath)                        	; Fehlerausgabe

			}

		}

		If !Addendum.ToolbarThread {

			If Addendum.Thread["Toolbar"].ahkReady() {                                                                                    	; prüft ob der Thread noch läuft
				Addendum.Thread["Toolbar"].ahkTerminate()
				Addendum.Thread["Toolbar"] := ""
			}
			If (StrLen(Addendum.Threads.Toolbar) > 0)                                                                                      	; Variable die den Skriptcode enthält leeren
				Addendum.Threads.Toolbar := ""

		}

	; Funktionsbibliothek für PDF OCR mit Tesseract
		If (StrLen(Addendum.Threads.OCR) =0) {

			threadFilePath := Addendum.Dir "\Include\Addendum_OCR.ahk"
			If FileExist(threadFilePath)
				Addendum.Threads.OCR := FileOpen(threadFilePath, "r").Read()
			else
				SciTEOutput("\Include\Addendum_OCR.ahk konnte nicht geladen werden.")

		}

}


; ---- HILFSFUNKTIONEN ----
Sprechstunde(IniString) {                                                                          	;-- Sprechstundenzeiten

		obj := Object()
		Tage := StrSplit(IniString, "|")
		For i, Tag in Tage {
			RegExMatch(Tag, "(\w+)[\s\,]+([\d\:\-\,\s]+)", part)
			obj[part1] := part2
		}

return obj
}

ReadBots(BotName) {                                                                              	;-- Telegram-Bot Daten

		BotToken           	:= IniReadExt("Telegram", BotName "_Token")
		BotChatID         	:= IniReadExt("Telegram", BotName "_ChatID")
		BotActive           	:= IniReadExt("Telegram", BotName "_Active", 0)
		BotLastMsg        	:= IniReadExt("Telegram", BotName "_LastMsg", "--")
		BotLastMsgTime	:= IniReadExt("Telegram", BotName "_LastMsgTime", "000000")

return {"BotName":BotName, "Token": BotToken, "ChatID": BotChatID, "Active": BotActive, "BotLastMsg":BotLastMsg, "BotLastMsgTime":BotLastMsgTime}
}

RegExAutoPos(iniStr) {
	; 0 = keines , 1 = 2k , 2 = 4k , 3 = 2k & 4k
	wp := 0
	wp += RegExMatch(iniStr, "i)2k\s*\:\s*ja") ? 1 : 0
	wp += RegExMatch(iniStr, "i)4k\s*\:\s*ja") ? 2 : 0
return wp
}

RegExScreenSize(iniStr, screenDim) {
	If !RegExMatch(iniStr, screenDim "\[(?<Mon>\d+)\s*\,\s*(?<X>\d+|[LRBT-])\s*\,\s*(?<Y>\d+|\-)\s*\,\s*(?<W>\d+|S)\s*\,\s*(?<H>\d+|S)", w)
		throw Exception("ungültiger ini Eintrag: " iniStr, -1, "gültiges Format ist: 2k[MonNr,X,Y,W,H],4k[MonNr,X,Y,W,H]")
return {"Mon":wMon, "X":wX, "Y":wY, "W":wW, "H":wH}
}





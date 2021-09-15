; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                      	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                     liest Informationen aus der Addendum.ini ein für das globale Objekt Addendum
;                                                  	!diese Bibliothek enthält Funktionen für Einstellungen des Addendum Hauptskriptes!
;                                  	   by Ixiko started in September 2017 - last change 09.09.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

return
admObjekte() {                                                                                       	;-- Addendum-Objekt erweitern

		Addendum.AktuellerTag	    	:= A_DD            	; der aktuelle Tag
		Addendum.Abrechnung	        	:= Object()          	; Daten zu Abrechnungsgebühren
 		Addendum.Chroniker            	:= Array()          	; Patient ID's der Chroniker
		Addendum.Default                  	:= Object()       	; Addendum Font/Farbeinstellungen
		Addendum.Drucker               	:= Object()          	; verschiedene Druckereinstellungen
		Addendum.Geriatrisch           	:= Array()         	; Patient ID's für geriatrisches Basisassement
		Addendum.Hooks                  	:= Object()       	; enthält Adressen der Callbacks der Hookprozeduren und anderes
		Addendum.Kosten                  	:= Object()        	; hinterlegte Preise für Hotstringausgaben bei Privatabrechnungen
		Addendum.Labor	    	        	:= Object()			; Laborabruf und andere Daten für Laborprozesse
		Addendum.Laborjournal        	:= Object()			; Laborjournal AutoAnzeige
		Addendum.LAN	                    	:= Object()			; LAN Kommunikationseinstellungen
		Addendum.LAN.Clients          	:= Object()          	; LAN Netzwerkgeräte Einstellungen
		Addendum.Module                  	:= Object()        	; Addendum Skriptmodule
		Addendum.MsgGui               	:= Object()        	; Gui für Interskript Kommunikation und den Shellhook
		Addendum.OCR                     	:= Object()          	; Einstellungen für Texterkennung und Autonaming
		Addendum.PatExtra                	:= Object()          	; zusätzliche Informationen über Patienten
		Addendum.PDF                      	:= Object()          	; Einstellungen/Pfade für die Bearbeitung von PDF Dateien
		Addendum.Praxis                     	:= Object()        	; Praxisadressdaten
		Addendum.Praxis.Sprechstunde	:= Object()       	; Daten zu Öffnungszeiten, Urlaubstagen
		Addendum.Praxis.Email          	:= Array()         	; Praxis Email Adressen
		Addendum.Praxis.Urlaub         	:= Array()          	; Praxis Urlaubsdaten
		Addendum.Telegram             	:= Object()        	; Datenobject für einen oder mehrere Telegram-Bots
		Addendum.Thread                  	:= Object()          	; enthält die PID's gestarteter Threads (z.B. AddendumToolbar, Addendum_OCR.ahk)
		Addendum.Threads                 	:= Object()         	; Skriptcode für Prozess-Threads
		Addendum.Tools                     	:= Object()			; externe Programme die z.B. über das Infofenster gestartet werden können
		Addendum.UndoRedo             	:= Object()          	; Undo/Redo Textbuffer
		Addendum.UndoRedo.List        	:= Array()           	; Undo/Redo Textbuffer
		Addendum.Windows                	:= Object()        	; Fenstereinstellungen für andere Programme
		Addendum.WZ                      	:= Object()       	; Wartezimmer Einstellungen
		Addendum.CImpf                  	:= Object()       	; Corona Impfhelfer
		Addendum.PraxTTDebug := false

}

admVerzeichnisse() {                                                                             	;-- Programm- und Datenverzeichnisse (!nach admObjekte aufrufen!)

		Addendum.Dir                              	:= AddendumDir
		Addendum.Ini                                 	:= AddendumDir "\Addendum.ini"

		Addendum.DBPath                          	:= IniReadExt("Addendum" 	, "AddendumDBPath"     	)              	; Datenbankverzeichnis
		Addendum.LogPath                          	:= IniReadExt("Addendum" 	, "AddendumLogPath"     	)              	; Logbücher-Verzeichnis
		Addendum.BefundOrdner                	:= IniReadExt("ScanPool"     	, "BefundOrdner"            	)             	; BefundOrdner = Scan-Ordner für neue Befundzugänge
		Addendum.ExportOrdner                	:= IniReadExt("ScanPool"    	, "ExportOrdner", Addendum.BefundOrdner "\Export")  ;### ÜBERPRÜFUNG ANLEGEN
		Addendum.VideoOrdner                	:= IniReadExt("ScanPool"     	, "VideoOrdner"             	)             	; Video's z.B. aus Dicom CD's (CT Bilder u.a.)
		Addendum.AlbisExe                         	:= IniReadExt("Albis"           	, "AlbisExe"                    	)             	; Pfad zum Albis-Stammverzeichnis
		Addendum.DosisDokumentenPfad   	:= IniReadExt("Addendum" 	, "DosisDokumentenPfad"	)              	; MS Word Dateien mit eigenen Hinweisen zu Medikamentendosierungen

		Addendum.TPPath	                        	:= Addendum.DBPath "\Tagesprotokolle\" A_YYYY                     	; Tagesprotokollverzeichnis
		Addendum.TPFullPath                      	:= Addendum.TPPath "\" A_MM "-" A_MMMM "_TP.txt" 	               	; Name des aktuellen Tagesprotokolls

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

		Addendum.Drucker.Standard          	:= IniReadExt(compname, "Drucker_Standard"             	, "")
		Addendum.Drucker.StandardA4     	:= IniReadExt(compname, "Drucker_Standard_A4"       	, "")
		Addendum.Drucker.StandardA4Tray 	:= IniReadExt(compname, "Drucker_Standard_A4_Tray"	, "")
		Addendum.Drucker.PDF                   	:= IniReadExt(compname, "Drucker_PDF"                   	, "")
		Addendum.Drucker.FAX                   	:= IniReadExt(compname, "Drucker_FAX"                    	, "")

	; ERROR löschen wenn die Einstellung für den Standarddrucker nicht vorhanden ist
		For standard, printer in Addendum.Drucker
			If InStr(printer, "ERROR")
				Addendum.Drucker[standard] := ""

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

	; Laborabruf bei manuellem Start automatisieren
		Addendum.Labor.AutoImport		    := IniReadExt(compname	, "Laborimport_automatisieren"         	, "nein"              	)
		If InStr(Addendum.Labor.AutoImport, "Error") || (StrLen(Addendum.Labor.AutoImport) = 0)
			Addendum.Labor.AutoImport	:= false

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

	; Schnellrezept integrieren
		Addendum.Schnellrezept           	:= IniReadExt(compname	, "Schnellrezept"                              	, "ja"                 	)

	; zeigt TrayTips
		Addendum.ShowTrayTips           	:= IniReadExt(compname	, "TrayTips_zeigen"                            	, "nein"                 	)

	; AutoOCR - Eintrag wird nur auf dem dazu berechtigten Client angezeigt
		Addendum.OCR.AutoOCR            	:= IniReadExt("OCR"      	, "AutoOCR"                                  	, "nein"                 	)

	; ermöglicht die sofortige Bearbeitung/Anzeige neuer Dateien
		Addendum.OCR.WatchFolder    	:= IniReadExt(compname	, "BefundOrdner_ueberwachen"       	, "ja"                 	)

	; automatische Bestätigung bei "Möchten Sie diesen Eintrag wirklich löschen?"
		Addendum.AutoDelete                	:= IniReadExt(compname	, "Eintrag_wirklich-loeschen"            	, "nein"                 	)

	; Boss PC Name (debugging features nur auf einem PC darstellen)
		Addendum.IamTheBoss                	:= IniReadExt(compname	, "IamTheBoss"                                                           	)
		If InStr(Addendum.IamTheBoss, "Error") || (StrLen(Addendum.IamTheBoss) = 0)
			Addendum.IamTheBoss	:= false

	; Excel Helferlein
		Addendum.CImpf.Helfer             	:= IniReadExt(compname	, "CoronaImpfhelfer"                      	, "nein"                 	)
		If Addendum.CImpf.Helfer {
			Addendum.CImpf.ScriptPath     	:= AddendumDir "\Module\Corona-Impfung"
			Addendum.CImpf.ScriptName  	:= "Corona-ExcelDB.ahk"
		}
}

admGesundheitsvorsorge() {                                                                    	;-- Abstände zwischen Untersuchungen minimales Alter

	Addendum.GVUAbstand                 	:= IniReadExt("Abrechnungshelfer"	, "minGVUAbstand")
	Addendum.GVUminAlter                	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlter")
	Addendum.GVUminAlterEinmalig   	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterEinmalig")
	Addendum.GVUminAlterMehrmalig	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterMehrmalig")

}

admPIDHandles() {                                                                                  	;-- Prozess-ID's

		Addendum.AlbisPID                        	:= AlbisPID()
		Addendum.scriptPID                       	:= DllCall("GetCurrentProcessId")          	; die ScriptPID wird für das SkriptReload und die Interskript-Kommunikation benötigt

}

admInfoWindowSettings() {                                                                      	;-- AddendumGui/Infofenster

	; Position des Addendumfenster in Albis
		tmp1 := IniReadExt(compname, "InfoFenster_Position", "y2 w400 h340")
		RegExMatch(tmp1, "y\s*(\d+)\s*w\s*(\d+)\s*h\s*(\d+)", match)

	; nutze hier oder in der ini den Syntax [Trigger1,Trigger2](Name des Wartezimmer,Kommentar,Anwesenheit)
	; wenn bei Kommentar "Dateiname" steht nutzt das Infofenster die Dokumentbezeichnung
		DefaultTrigger := "[Anfrage,Antrag](Anfragen,Dateiname,Ohne Status)"

		Addendum.iWin := {	"Y"                     	: match1
									, 	"W"                    	: match2
									,	"ReIndex"            	: false
									,	"rowprcs"           	: 0
									,	"Init"                    	: false
									,	"AbrHelfer"        	: IniReadExt(compname, "Infofenster_Abrechnungshelfer"                	, "Ja")    ; Abrechnungshelfer anzeigen
									,	"WZKommentar"   	: IniReadExt(compname, "Infofenster_AutoWZ_Kommentar"              	, "Ja")
									,	"AutoWZTrigger" 	: IniReadExt(compname, "Infofenster_AutoWZ_Trigger"                     	, DefaultTrigger)
									,	"ConfirmImport"   	: IniReadExt(compname, "Infofenster_Import_Einzelbestaetigung"      	, "Ja")
									,	"firstTab"           	: IniReadExt(compname, "Infofenster_aktuelles_Tab"                         	, "Patient")
									,	"TProtDate"         	: IniReadExt(compname, "Infofenster_Tagesprotokoll_Datum"            	, "Heute")
									,	"JournalSort"      	: IniReadExt(compname, "Infofenster_JournalSortierung"                     	, "3 1 1")
									,	"Debug"            	: IniReadExt(compname, "Infofenster_Debug"                                  	, "Nein")
									,	"LBDrucker"       	: IniReadExt(compname, "Infofenster_Laborblatt_Drucker"                  	, "")
									,	"LBAnm"             	: IniReadExt(compname, "Infofenster_Laborblatt_Anmerkung_drucken"	, "Nein")
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

	; Laborabruf Prozeßschlüssel anlegen
		Addendum.Labor.AbrufStatus 	:= false
		Addendum.Labor.AbrufDaten	:= false
		Addendum.Labor.AbrufVoll    	:= false

	; Laborabrufzeiten z.B. "06:00, 15:00, 19:00, 21:00" , werden in ein Array umgewandelt
		spos := 1
		Addendum.Labor.AbrufZeiten    	:= Array()
		while (spos := RegExMatch(IniReadExt("LaborAbruf", "LaborAbruf_Zeiten", ""), "(?<H>\d+)\:(?<M>\d+)", timer, spos)) {
			spos += StrLen(timer)
			Addendum.Labor.AbrufZeiten.Push("T" SubStr("00" timerH, -1) SubStr("00" timerM, -1) "00")
		}

	; auf welchem/n Client/s der Laborabruf ausgeführt werden darf
		Addendum.Labor.ExecuteOn       	:= IniReadExt("LaborAbruf"	, "OnlyRunOnClient"         	    , ""            	)

	; externes Programm das für den Abruf ausgeführt werden muss
		Addendum.Labor.ExecuteFile       	:= IniReadExt("LaborAbruf"	, "Laborabruf_Extern"      	    , ""            	)

	; Skript das den Abruf übernimmt
		Addendum.Labor.ExecuteScript     	:= IniReadExt("LaborAbruf"	, "Laborabruf_Skript"             	, ""            	)

	; nach dem Ende des Laborabrufes das Laborbuch anzeigen lassen
		Addendum.Labor.ZeigeLabJournal	:= IniReadExt("LaborAbruf"	, "Laborabruf_ZeigeJournal"  	, "nein"        	)

	; falls sie mehrere Labore haben, tragen Sie das aktuelle hier ein
		Addendum.Labor.LbName         	:= IniReadExt("LaborAbruf"	, "LaborName"               	    , ""            	)

	; Alarmierungsgrenze in Prozent oberhalb der Normgrenzen
		Addendum.Labor.Alarmgrenze     	:= IniReadExt("LaborAbruf"	, "Alarmgrenze"             	    , "30%"      	)

	; Karteikartenkürzel um Informationen ablegen zu können
		Addendum.Labor.Laborkuerzel     	:= IniReadExt("LaborAbruf"	, "Aktenkuerzel"              	    , "labor"    	)

	; Verzeichnis in dem die heruntergeladenen LDT-Dateien zwischengespeichert werden
		Addendum.Labor.LDTDirectory     	:= IniReadExt("LaborAbruf"	, "LDTDirectory"                  	, "C:\Labor"	)

	; für order&entry per CGM-Channel
		Addendum.Labor.Kennwort       	:= IniReadExt("LaborAbruf"	, "LaborKennwort"            	    , ""            	)
		If InStr(Addendum.Labor.Kennwort, "Error")
			Addendum.Labor.Kennwort := ""


	; -----------------------------------------------------------------------------------------------------------------------------------------------------------------
	; 	ACHTUNG
	; -----------------------------------------------------------------------------------------------------------------------------------------------------------------
	; 	hier werden die Einstellungen für eine Alarmierung über Ihren eigenen Telegram-Bot eingelesen!
	; 	Denken Sie immer daran die Telegram Token und ChatID's vor dem Zugriff Fremder oder auch dem eigenen Personal zu schützen
	; -----------------------------------------------------------------------------------------------------------------------------------------------------------------

	; die Nummer ihres Telegram Bots und + oder - Tel z.b. "1+Tel"
		If (Addendum.Labor.TGramMsg     	:= IniReadExt("LaborAbruf"	, "TGramVersand"                  	, "Nein"	        	)) {

				Addendum.Labor.TGramOpt    	:= IniReadExt("LaborAbruf"	, "TGramOpt"                	    	, "+Tel -Crypt"  	)
				Addendum.Labor.TGramBotNr 	:= IniReadExt("LaborAbruf"	, "TGramBotNr"             	    	, "1"      	    	)

			; die Werte müssen einen exakten Syntax haben, sonst wird die Telegram Option gelöscht!
				If !RegExMatch(Addendum.Labor.TGramOpt, "i)([\+\-][a-z]+)\V*([\+\-][a-z]+)") {
					Addendum.Labor.TGramMsg := false
					IniWrite, % "nein", % Addendum.Ini, % "LaborAbruf", % "TGramVersand"
				}
		}

}

admLaborJournal() {                                                                              	;-- geplantes Anzeigen des Laborjournal

	Addendum.Laborjournal.SheduledView 	:= IniReadExt(compname, "Laborjournal_AutoAnzeige"   	,  "Nein"            	)
	Addendum.Laborjournal.StartTime     	:= IniReadExt(compname, "Laborjournal_Startzeit"           	,  "---"                	)

	For key, val in Addendum.Laborjournal
		If (val = "ERROR" || val = "---")
			Addendum.Laborjournal[key] := ""

}

admLanProperties() {                                                                              	;-- Netzwerkeinstellungen / Client PC Namen / IP-Adressen

	; prüft auf den korrekt gesetzten Computernamen
		CName		:= IniReadExt(compname, "ComputerName")
		If (CName <> A_ComputerName)
			IniWrite, % A_ComputerName, % Addendum.Ini, % compname, % "ComputerName"

	; hält die in der ini gespeicherte IP zu jedem Client aktuell
		CompIP 	:= A_IPAddress1
		storedIP   	:= IniReadExt(compname, "IP")
		If (CompIP <> storedIP)
			IniWrite, % CompIP, % Addendum.Ini, % compname, % "IP"

	; liest die Namen aller Clients, die überwacht werden sollen oder die miteinander kommunizieren dürfen
		For idx, clientName in StrSplit(IniReadExt("Computer", "Computer"), "|")
			Addendum.LAN.Clients[clientName] := {	"computername"	: IniReadExt(clientName, "ComputerName")
													                	,	"remoteShutdown"	: IniReadExt(clientName, "AutoShutDown")
													                	,	"rdpPath"            	: IniReadExt(clientName, "rdpPath")
													                	,	"IP"                    	: IniReadExt(clientName, "IP") }

	; Addendum erhält die Fähigkeit im LAN Daten senden und empfangen zu können

		ServerPC   	:= IniReadExt("Addendum", "Addendum_ServerPC"     	, "")
		ServerPC   	:=  InStr(ServerPC, "ERROR") ? "" 	: ServerPC
		If ServerPC && (ServerPC = compname)
			Addendum.LAN.IamServer := true
		else
			Addendum.LAN.IamServer := false

		Server	:= IniReadExt("Addendum", "Addendum_ServerAdresse", 0)
		Server  	:= InStr(admServer, "ERROR") ? ""	: Server
		RegExMatch(Server, "(?<adr>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s*\:\s*(?<port>\d+)", tcp_)
		If tcp_adr && tcp_port
			Addendum.LAN.admServer := {"IP":tcp_adr, "Port":tcp_port}

}

admMailAndHolidays() {                                                                           	;-- eigene Mailadresse und Urlaubsdaten

	Loop
		If !InStr((tmp :=  IniReadExt("Allgemeines", "EMail" A_Index)), "ERROR")
			Addendum.Praxis.Email.Push(tmp)
		else
			break

}

admMonitorInfo() {                                                                                  	;-- angeschlossenen Monitore

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
		Addendum.OCR.Client                        	:= IniReadExt("OCR" , "AutoOCRClient")
		If InStr(Addendum.OCR.Client, "Error")                                	; ## Abfrage integrieren
			Addendum.OCR.Client := ""

	; Zeitverzögerung in Sekunden bis zum Start der Texterkennung,
		Addendum.OCR.AutoOCRDelay           	:= IniReadExt("OCR" , "AutoOCR_Startverzoegerung", 30)
		If InStr(Addendum.OCR.AutoOCRDelay, "Error")                   	; ## Abfrage integrieren
			Addendum.OCR.AutoOCRDelay := 30

	; wenn true dann ignoriert Watchfolder alle Dateiveränderungen solange nicht auf false gegangen wird
		Addendum.OCR.PauseWF := false
		Addendum.OCR.WFIgnore := Array()

	; Zähler der Texterkennung, aktuell und seit Programmstart
		Addendum.OCR.staticFileCount	:= 0
		Addendum.OCR.filecount         	:= 0

}

admPraxisDaten() {                                                                                  	;-- Daten der Praxis

	; Sprechstunden werden mit Tagesbezeichnung z.B. .["Montag"] := "07:00-19:00" gespeichert

	Addendum.Praxis.Name            	:= IniReadExt("Allgemeines", "PraxisName")
	Addendum.Praxis.Strasse           	:= IniReadExt("Allgemeines", "Strasse")
	Addendum.Praxis.PLZ                	:= IniReadExt("Allgemeines", "PLZ")
	Addendum.Praxis.Ort                	:= IniReadExt("Allgemeines", "Ort")
	Addendum.Praxis.KVStempel         	:= IniReadExt("Allgemeines", "KVStempel")                            	; genutzt für Textersetzung
	Addendum.Praxis.MailStempel      	:= IniReadExt("Allgemeines", "MailStempel")                           	; genutzt für Textersetzung
	Addendum.Praxis.Sprechstunde    	:= Sprechstunde(IniReadExt("Allgemeines", "Sprechstunde")) 	; wandelt den Ini-String um
	RegExMatch(Addendum.Praxis.Sprechstunde[A_DDDD], "(\d+)\:(\d+)$", Uhr)
	Addendum.Praxis.TagesEnde     	:= ((Uhr1*3600) + (Uhr2*60)) * 1000	; ms von 0 Uhr bis zum Ende der Sprechstunde - für Shutdown Timer

  ; freie Tage und Praxisurlaub werden aus der ini gelesen
	Urlaub	:= IniReadExt("Allgemeines", "Urlaub")
	If (StrLen(Urlaub) > 0 && Urlaub <> "ERROR")
		Addendum.Praxis.Urlaub := vacation.ConvertDateString(Urlaub)

  ; Praxisärzte einlesen (!noch ohne Funktion!)
	Addendum.Praxis.Arzt := Array()
	Loop {

		Arzt	:= IniReadExt("Allgemeines", "Arzt" A_Index "Name")
		Fach	:= IniReadExt("Allgemeines", "Arzt" A_Index "Abrechnung")

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

class vacation {                     ;-- eine Funktionsklasse zur Abfrage des Praxisurlaub

	; benötigt Addendum.Praxis.Urlaub als globales Objekt
	; vollständiger Funktionsumfang ist noch nicht programmiert
	; bisherige Funktionen:    1. 	Ini-String mit den Daten wird in ein Array überführt, die Urlaubstage können mit relativ freier Schreibweise in der Ini-Datei eingetragen sein
	;                                       	einzelne Tage    	:  	in der Form 01.01.2022 auch als 1.1.22
	;                                           Datumsbereiche	:	05.06.2022-06.06.2022 oder 05.06.-06.06.2022 oder 5.-6.6.22
	;                                        	zwingend zur Unterscheidung sind die Punkte und das Minuszeichen
	;
	; letzte Änderung: 08.09.2021

	ConvertDateString(holidays:="") 	{                                                   	;-- Ini-String

		  ; Urlaubszeiten mit einem Datum in der Vergangenheit werden aussortiert. Die Daten werden umgewandelt.
			rxUrlaub := "((?<StartDD>\d{1,2})\.(?<StartMM>\d{1,2})*(\.*(?<StartYY>\d{2,4})*))"
						.   "-*((?<EndDD>\d{1,2})*\.*(?<EndMM>\d{1,2})*\.*(?<EndYY>\d{2,4})*)"	; 12.08.-19.08.2020 oder nur 12.08.-19.08.
			spos := 1, AToday := A_YYYY . A_MM . A_DD, AYearT := SubStr(A_YYYY, 1, 2)
			vacations := Array()

		  ; Abbruch wenn, nichts oder ein falsche String übergeben wurde
			If !holidays || !RegExMatch(holidays, rxUrlaub)
				return

		 ; Datum für Datum extrahieren
			while (spos := RegExMatch(holidays, rxUrlaub, PX, spos)) {

			  ; Stringposition weiterrücken
				spos  	+= StrLen(PX)

			 ; eine zweistellige Jahreszahl auf eine vierstellige Zahl ändern
				PXStartYY	:= StrLen(PXStartYY) = 2 	? AYearT . PXStartYY 	: PXStartYY
				PXEndYY	:= StrLen(PXEndYY) = 2	? AYearT . PXEndYY 	: PXEndYY

			  ; Monate und Tage 2-stellig auffüllen
				PXStartMM	:= SubStr("00" PXStartMM, -1)
				PXStartDD 	:= SubStr("00" PXStartDD, -1)
				PXEndMM   	:= SubStr("00" PXEndMM, -1)
				PXEndDD   	:= SubStr("00" PXEndDD, -1)

			  ; die Formatierung wird auf YYYYMMDD geändert
				AYearS	:= (!PXStartYY && !PXEndYY) ? A_YYYY : PXStartYY ? PXStartYY : PXEndYY ? PXEndYY   ; Jahr
				ADayS	:= AYearS . PXStartMM . PXStartDD
				AYearE	:= (!PXStartYY && !PXEndYY) ? A_YYYY : PXStartYY ? PXStartYY : PXEndYY ? PXEndYY   ; Jahr
				ADayE	:= (PXEndMM && PXEndDD) ? AYearE . PXEndMM . PXEndDD : ""

			  ; ein Datumsbereich wird anders hinterlegt, als ein einzelner freier Tag
				If (ADayS && ADayE && !this.PeriodExists(ADayS, ADayE))
					vacations.Push({"IsPeriod":true	, "firstday":ADayS, "lastday":ADayE})
				else If (ADayS && !ADayE && !this.PeriodExists(ADayS))
					vacations.Push({"IsPeriod":false	, "day":ADayS})

				;~ SciTEOutput("(" A_LineNumber "): " "(" PX ") " ADayS "; " ADayE)
			}

			;~ SciTEOutput("(" A_LineNumber "): Anzahl eingetragener Urlaube: " Addendum.Praxis.Urlaub.Count())

	return vacations
	}

	PeriodExists(firstday, lastday:="")	{                                                   	;-- Datum oder Datumsbereich suchen

		If !firstday && !lastday
			return false

		For hdNR, holidays in Addendum.Praxis.Urlaub
			If (firstday && lastDate) {
				If (firstday = holidays.firstday && lastday = holidays.lastday)
					return hdNR
			}
			else {
				If (!holiday.IsPeriod && firstday = holidays.day)
					return hdNR
			}

	return false
	}

	DateIsHoliday(datestring)           	{                                                   	;-- ermittelt ob ein übergebenes Datum innerhalb eines Praxisurlaub liegt

	  ; automatisch 4-stelliges Jahresformat (funktioniert bis 2099)
		If RegExMatch(datestring, "(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>(\d{2}|\d{4}))", t)
			datestring := SubStr(SubStr(A_YYYY, 1, 2) . tY, -2) . tM . tD

		For hdNR, holidays in Addendum.Praxis.Urlaub
			If holidays.IsPeriod && (datestring >= holidays.firstday && datestring <= holidays.lastday)
				return hdNR
			else if (!holidays.IsPeriod && datestring = holidays.Day)
				return hdNR

	return false
	}

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

	; Verzögerung bis zum Schliessen des Dialoges "Patient hat in diesem Quartal seine Chipkarte..." (0 = keine Verzögerung)
		Addendum.noChippie := IniReadExt(compname, "keineChipkarte", 3)

	; Behördengebühren werden über die ini Datei verwaltet - besser anpassbar bei Änderung
		Behoerdengebuehr := {"JVEG":"25.00", "DRV1":"28.20", "DRV2": "35.00", "BAfA":"32.50"}
		Addendum.Abrechnung.JVEG 	:= IniReadExt("Addendum", "JVEG")
		Addendum.Abrechnung.DRV1  	:= IniReadExt("Addendum", "DRV1")
		Addendum.Abrechnung.DRV2  	:= IniReadExt("Addendum", "DRV2")
		Addendum.Abrechnung.BAfA  	:= IniReadExt("Addendum", "BAfA")
		For Behoerde, Gebuehr in Addendum.Abrechnung
			If !RegExMatch(Gebuehr, "\d+\.\d+")
				Addendum.Abrechnung[Behoerde] := Behoerdengebuehr[Behoerde]

	; Gebührenberechnung Coronaimpfung im Addendumskript
		If ((Addendum.CImpf.CNRRun   	:= IniReadExt(compname, "COVID19_ImpfGebuehr_Skript", "Ja")) = true) {
			Addendum.CImpf.CNRWin    	:= IniReadExt(compname, "COVID19_ImpfGebuehrfenster_Position")
			If InStr(Addendum.CImpf.CNRWin, "ERROR")
				Addendum.CImpf.CNRWin := ""
		}

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
	; letzte Änderung: 05.06.2021

	; Toolbar-Thread starten oder beenden
		If Addendum.ToolbarThread {

			threadFilePath := Addendum.Dir "\Module\Addendum\threads\AddendumToolbar.ahk"
			If FileExist(threadFilePath) {                                                                                                           	; Skriptdatei vorhanden:

				Addendum.Threads.Toolbar  	:= FileOpen(threadFilePath,"r").Read()                                       	; Skriptdatei laden
				If !Addendum.Thread["Toolbar"].ahkReady() {                                                                            	; Toolbar-Thread ist nicht gestartet?
					If !(Addendum.Thread["Toolbar"] := AHKThread(Addendum.Threads.Toolbar)) {                     	; Toolbar-Thread starten
						Addendum.ToolbarThread := false                                                                                 	; erneutes Aufrufen des Threads unterbinden
						PraxTT("Die Addendum Toolbar konnte nicht gestartet werden.", "3 3")			    			    	; Information an den Nutzer
						throw Exception("Der Toolbar-Thread konnte nicht gestartet werden.",	-1, threadFilePath)    	; Fehlerausgabe
					}
				}

			} else {                                                                                                                                       	; Skriptdatei nicht vorhanden:

				Addendum.ToolbarThread := false                                                                                        	; erneutes Aufrufen des Threads unterbinden
				failpath := StrReplace(threadFilePath, Addendum.Dir)
				failpath	:= StrReplace(failPath, "\AddendumToolbar.ahk")
				PraxTT(	"Die Datei: AddendumToolbar.ahk ist nicht vorhanden.`n"                                            	; Information zum Fehler für den Nutzer
						. 	"Verzeichnis: " failpath, "3 3")
				throw Exception("Die Datei: AddendumToolbar.ahk ist nicht vorhanden.",	-1, failPath)                  	; Fehlerausgabe

			}

		}

		If !Addendum.ToolbarThread {

			If Addendum.Thread["Toolbar"].ahkReady() {                                                                                	; prüft ob der Thread noch läuft
				Addendum.Thread["Toolbar"].ahkTerminate()
				Addendum.Thread["Toolbar"] := ""
			}
			If (StrLen(Addendum.Threads.Toolbar) > 0)                                                                                 	; Variable die den Skriptcode enthält leeren
				Addendum.Threads.Toolbar := ""

		}

	; Funktionsbibliothek für PDF OCR mit Tesseract
		If (StrLen(Addendum.Threads.OCR) =0) {

			OCRThreadPath1 := Addendum.Dir "\Include\Addendum_OCR.ahk"
			OCRThreadPath2 := Addendum.Dir "\Include\Addendum_Mining.ahk"

			If FileExist(OCRThreadPath1)
				Addendum.Threads.OCR :=	FileOpen(OCRThreadPath1, "r").Read()  "`n"
			else
				SciTEOutput("\Include\Addendum_OCR.ahk konnte nicht geladen werden.")

			If FileExist(OCRThreadPath2)
				Addendum.Threads.OCR .=	FileOpen(OCRThreadPath2, "r").Read()
			else
				SciTEOutput("\Include\Addendum_Mining.ahk konnte nicht geladen werden.")

		}

}

admWartezimmer() {

	Addendum.WZ.RemoveFast 	:= IniReadExt("Wartezimmer", "Schnell_Entfernen", "")

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





; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                      	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                     liest Informationen aus der Addendum.ini ein für das globale Objekt Addendum
;                                                  	!diese Bibliothek enthält Funktionen für Einstellungen des Addendum Hauptskriptes!
;                                  	   by Ixiko started in September 2017 - last change 24.05.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

return

AddendumProperties() {

	; cJSON Unicode-Ausgabe ab hier
		cJSON.EscapeUnicode := "UTF-8"

	; !! die Reihenfolge der Funktionsaufrufe nicht ändern !!

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
		admLaborJournal()                                              	; das Laborjournal zum Dienstbeginn anzeigen
		admLanProperties()                                             	; LAN - Kommunikation Addendumskripte
		admMailAndHolidays()                                       	; Mailadressen, Urlaubszeiten
		admMonitorInfo()                                                	; ermittelt die derzeitige Anzahl der Monitore und deren Größen
		admPIDHandles()                     	                        	; Prozess-ID's
		admPDFSettings()					                            	; PDF anzeigen, signieren
		admOCRSettings()                                              	; OCR - cmdline Programme
		admPraxisDaten()                                               	; Praxisdaten   -   für Outlook, Telegram Nachrichtenhandling, Sprechstundenzeiten, Urlaub
		admPreise()                                                        	; Preise für dyn. Ersetzung in Hotstrings
		admShutDown()                                                	; Einstellungen für automatisches Herunterfahren des PC
		admSicherheit()                                                   	; Lockdown Funktion - Nutzer hat den Arbeitsplatz verlassen
		admSonstiges()                                                    	; Einstellungen die in keine der anderen Kategorie gehören
		admStandard()                                                    	; Standard-Einstellungen für Gui's
		admWartezimmer()                                            	; Wartezimmer Einstellungen
		admTagesProtokoll()                                           	; Tagesprotokoll laden
		admTelegramBots()				                            	; Telegram Bot Daten
		admThreads()                                                    	; Skriptcode für Multithreading (z.B. Tesseract OCR)

		ChronikerListe()                                                    	; Chroniker Index
		GeriatrischListe()                                                   	; Geriatrie Index

}

admObjekte() {                                                                                       	;-- Addendum-Objekt erweitern

		Addendum.AktuellerTag	    	:= A_DD            	; der aktuelle Tag
		Addendum.Abrechnung	        	:= Object()          	; Daten zu Abrechnungsgebühren
		Addendum.AutoDelete	        	:= Object()          	; in welchem Kontext darf ohne Nachfrage gelöscht werden
 		Addendum.Chroniker            	:= Array()          	; Patient ID's der Chroniker
		Addendum.Default                  	:= Object()       	; Addendum Font/Farbeinstellungen
		Addendum.Drucker               	:= Object()          	; verschiedene Druckereinstellungen
		Addendum.Flags                    	:= Object()       	; für Informationsaustausch mit automatischen Funktionen
		Addendum.Geriatrisch           	:= Array()         	; Patient ID's für geriatrisches Basisassement
		Addendum.Hooks                  	:= Object()       	; enthält Adressen der Callbacks der Hookprozeduren und anderes
		Addendum.Kosten                  	:= Object()        	; hinterlegte Preise für Hotstringausgaben bei Privatabrechnungen
		Addendum.Labor	    	        	:= Object()			; Laborabruf und andere Daten für Laborprozesse
		Addendum.Laborjournal        	:= Object()			; Laborjournal AutoAnzeige
		Addendum.LAN	                    	:= Object()			; LAN Kommunikationseinstellungen
		Addendum.LAN.Clients          	:= Object()          	; LAN Netzwerkgeräte Einstellungen
		Addendum.Mail                     	:= Object()        	; Mail Manager Einstellungen
		Addendum.Module                  	:= Object()        	; Addendum Skriptmodule
		Addendum.MsgGui               	:= Object()        	; Gui für Interskript Kommunikation und den Shellhook
		Addendum.OCR                     	:= Object()          	; Einstellungen für Texterkennung und Autonaming
		Addendum.PatExtra                	:= Object()          	; zusätzliche Informationen über Patienten
		Addendum.PDF                      	:= Object()          	; Einstellungen/Pfade für die Bearbeitung von PDF Dateien
		Addendum.PDF.Faxbox         	:= Object()
		Addendum.Praxis                     	:= Object()        	; Praxisadressdaten
		Addendum.Praxis.Sprechstunde	:= Object()       	; Daten zu Öffnungszeiten, Urlaubstagen
		Addendum.Praxis.Email          	:= Array()         	; Praxis Email Adressen
		Addendum.Praxis.Urlaub         	:= Array()          	; Praxis Urlaubsdaten
		Addendum.Preis                     	:= Object()        	; Preise für dyn. Ersetzung in Hotstrings
		Addendum.Telegram             	:= Object()        	; Datenobject für einen oder mehrere Telegram-Bots
		Addendum.Thread                  	:= Object()          	; enthält die PID's gestarteter Threads (z.B. AddendumToolbar, Addendum_OCR.ahk)
		Addendum.Threads                 	:= Object()         	; Skriptcode für Prozess-Threads
		Addendum.Tools                     	:= Object()			; externe Programme die z.B. über das Infofenster gestartet werden können
		Addendum.UndoRedo             	:= Object()          	; Undo/Redo Textbuffer
		Addendum.UndoRedo.List        	:= Array()           	; Undo/Redo Textbuffer
		Addendum.Windows                	:= Object()        	; Fenstereinstellungen für andere Programme
		Addendum.WZ                      	:= Object()       	; Wartezimmer Einstellungen
		Addendum.CImpf                  	:= Object()       	; Corona Impfhelfer

		Addendum.PraxTTDebug        	:= false

}

admVerzeichnisse() {                                                                             	;-- Programm- und Datenverzeichnisse (!nach admObjekte aufrufen!)

		Addendum.Dir                              	:= AddendumDir
		Addendum.Ini                                 	:= AddendumDir "\Addendum.ini"

		Addendum.DBPath                          	:= IniReadExt("Addendum" 	, "AddendumDBPath"     	)              	; Datenbankverzeichnis
		Addendum.LogPath                          	:= IniReadExt("Addendum" 	, "AddendumLogPath"     	)              	; Logbücher-Verzeichnis
		Addendum.BefundOrdner                	:= IniReadExt("ScanPool"     	, "BefundOrdner"            	)             	; BefundOrdner = Scan-Ordner für neue Befundzugänge
		Addendum.ExportOrdner                	:= IniReadExt("ScanPool"    	, "ExportOrdner", Addendum.BefundOrdner "\Export")  ;### ÜBERPRÜFUNG ANLEGEN
		Addendum.VideoOrdner                	:= IniReadExt("ScanPool"     	, "VideoOrdner"             	)             	; Video's z.B. aus Dicom CD's (CT Bilder u.a.)
		Addendum.DosisDokumentenPfad   	:= IniReadExt("Addendum" 	, "DosisDokumentenPfad"	)              	; Word Dateien mit eigenen Hinweisen zu
																																									; Medikamentendosierungen

		Addendum.TPPath	                        	:= Addendum.DBPath "\Tagesprotokolle\" A_YYYY                     	; Tagesprotokollverzeichnis
		Addendum.TPFullPath                      	:= Addendum.TPPath "\" A_MM "-" A_MMMM "_TP.txt" 	               	; Name des aktuellen Tagesprotokolls
		Addendum.DataPath                     	:= AddendumDir "\include\Daten"                                             	; Verzeichnis mit vorbereiteten Daten

	; Albisverzeichnisse
		Addendum.Albis     	:= Object()
		Addendum.Albis     	:= GetAlbisPaths()
		Addendum.AlbisExe  	:= Addendum.Albis.Exe

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

	If FileExist(filepath := Addendum.DbPath "\DB_Chroniker.txt")
		For idx, line in StrSplit(FileOpen(filepath, "r", "UTF-8").Read(), "`n", "`r")
			If RegExMatch(A_LoopField, "^\d+", PatID)
				Addendum.Chroniker.Push(PatID)

}

GeriatrischListe(LoadPatients:=true) {                                                      	;-- ICD Ziffern der Geriatrie und Patienten mit abgerechnetem GB

	;-- Pat. mit jemals abgerechnetem geriatrischen Basiskomplex 03360/03362

	; geriatrischer Basiskomplex - notwendige Diagnosen
		Addendum.GeriatrieICD := [	"Ataktischer Gang {R26.0G}"         	, "Spastischer Gang {R26.1G}"                         	, "Probleme beim Gehen {R26.2G}"
						                		, 	"Bettlägerigkeit {R26.3G}"             	, "Immobilität {R26.3G}"                                   	, "Standunsicherheit {R26.8G}"
						                		, 	"Ataxie {R27.0G}"                          	, "Koordinationsstörung {R27.8G}"                    	, "Multifaktoriell bedingte Mobilitätsstörung {R26.8G}"
						                		,	"Hemineglect {R29.5G}"  	            	, "Sturzneigung a.n.k. {R29.6G}"                       	, "Harninkontinenz {R32G}"
						                		,	"Stuhlinkontinenz {R15G}"             	, "Überlaufinkontinenz {N39.41G}"                   	, "n.n. bz. Harninkontinenz {N39.48G}"
						                		,	"Orientierungsstörung {R41.0G}"   	, "Gedächtnisstörung {R41.3G}"                       	, "Vertigo {R42G}"
						                		, 	"chron. Schmerzsyndrom {R52.2G}"	, "chron. unbeeinflußbarer Schmerz {R51.1G}"  	, "Dementia senilis {F03G}"
						                		,	"Multiinfarktdemenz {F01.1G}"      	, "Subkortikale vaskuläre Demenz {F01.2G}" 		, "Vorliegen eines Pflegegrades {Z74.9G}"
						                		,	"Vaskuläre Demenz {F01.9G}"        	, "Chronische Schmerzstörung mit somatischen und psychischen Faktoren {F45.41G}"
						                		,	"Dysphagie {R13.9G}"                   	, "Dysphagie mit Beaufsichtigungspflicht während der Nahrungsaufnahme {R13.0G}"
						                		, 	"Gemischte kortikale und subkortikale vaskuläre Demenz {F01.3G}"
						                		,	"Dysphagie bei absaugpflichtigem Tracheostoma mit geblockter Trachealkanüle {R13.1G}"
						                		,	"Demenz bei Alzheimer-Krankheit mit frühem Beginn (Typ 2) [F00.0*] {G30.0+G}; Alzheimer-Krankheit, früher Beginn {G30.0G}"
						                		,	"Demenz bei Alzheimer-Krankheit mit spätem Beginn (Typ 1) [F00.1*] {G30.1+G}; Alzheimer-Krankheit, später Beginn {G30.1G}"
						                		,	"Primäres Parkinson-Syndrom mit mäßiger Beeinträchtigung ohne Wirkungsfluktuation {G20.10G}"
						                		, 	"Primäres Parkinson-Syndrom mit mäßiger Beeinträchtigung mit Wirkungsfluktuation {G20.11G}"
						                		, 	"Primäres Parkinson-Syndrom mit schwerster Beeinträchtigung ohne Wirkungsfluktuation {G20.20G}"
						                		, 	"Primäres Parkinson-Syndrom mit schwerster Beeinträchtigung mit Wirkungsfluktuation {G20.21G}"]


	If LoadPatients
		If FileExist(filepath := Addendum.DbPath "\DB_Geriatrisch.txt")
			For idx, line in StrSplit(FileOpen(filepath, "r", "UTF-8").Read(), "`n", "`r")
				If RegExMatch(A_LoopField, "^\d+", PatID)
					Addendum.geriatrisch.Push(PatID)

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

	; ein Objekt mit den Bezeichnungen der Fensterklassen welche automatisch positioniert werden sollen
		Addendum.Windows.Proc  	:= {	"TForm_ipcMain"            	: "Ifap"                ; key=classNN : value=key der Einstellungen
														,	"classFoxitReader"             	: "Foxit"
														,	"SUMATRA_PDF_FRAME" 	: "Sumatra"
														,	"OpusApp"                    	: "MSWord" }

	; Auslesen der Einstellungen
		For classnn, appname in Addendum.Windows.Proc {

			tmp1 := Trim(IniReadExt(compname, "AutoPos_" appname, "2k:nein,4k:nein"))
			If (!InStr(tmp1, "ERROR") && StrLen(tmp1) > 0) {

				tmp2 := Trim(IniReadExt(compname, "Position_" appname, "2k[1,100,100,1600,1000],4k[1,100,100,1600,1000]"))
				Addendum.Windows[appname]             	:= Object()
				Addendum.Windows[appname]["AutoPos"]	:= RegExAutoPos(tmp1)
				Addendum.Windows[appname]["2k"]         	:= RegExScreenSize(tmp2, "2k")
				Addendum.Windows[appname]["4k"]         	:= RegExScreenSize(tmp2, "4k")

			}
			else
				Addendum.Windows.Proc.Delete(classnn)

		}

	; Scite_ZoomLevel_Anpassen=ja
	; Scite_ZoomLevels=4k[1,-1],2k[-2,-2]


}

admFunktionen() {                                                                                 	;-- zuschaltbare Funktionen

	; Addendum.OptF. ???

	; automatische Positionierung von Albis ist an
		tmp  := Trim(IniReadExt(compname, "AutoSize_Albis", "2k:nein,4k:nein" ))
		Addendum.AlbisLocationChange  	:= RegExAutoPos(tmp)

	; Größeneinstellung
		Addendum.AlbisGroesse            	:= IniReadExt(compname	, "Position_Albis"                                 	, "w1922 h1080"	)

	; Automatisierung der GVU Formulare
		Addendum.GVUAutomation			:= IniReadExt(compname	, "GVU_automatisieren"                     	, "nein"              	)

	; Automatisierung PDF Signierung
		Addendum.PDFSignieren   	    	:= IniReadExt(compname	, "FoxitPdfSignature_automatisieren"    	, "nein"              	)

	; Tastaturkürzel zum Auslösen des Signaturvorganges
		Addendum.PDFSignieren_Kuerzel	:= IniReadExt(compname	, "FoxitPdfSignature__Kuerzel" 	         	, "^Insert"            	)

	; signiertes Dokument automatisch schliessen
		Addendum.AutoCloseFoxitTab   	:= IniReadExt(compname	, "FoxitTabSchliessenNachSignierung" 	, "Ja"                 	)

	; zeitgesteuerter Laborabruf
		Addendum.Labor.AbrufTimer         := IniReadExt(compname	, "Laborabruf_Timer"           	   	        	, "nein"              	)

		If InStr(Addendum.Labor.AbrufTimer, "Error") || (StrLen(Addendum.Labor.AbrufTimer) = 0)
			Addendum.Labor.AbrufTimer	:= false

	; Laborabruf bei manuellem Start automatisieren
		Addendum.Labor.AutoAbruf 		    := IniReadExt(compname	, "Laborabruf_automatisieren"             	, "nein"              	)
		If InStr(Addendum.Labor.AutoAbruf, "Error") || (StrLen(Addendum.Labor.AutoAbruf) = 0)
			Addendum.Labor.AutoAbruf	:= false

	; Laborabruf bei manuellem Start automatisieren
		Addendum.Labor.AutoImport		    := IniReadExt(compname	, "Laborimport_automatisieren"             	, "nein"              	)
		If InStr(Addendum.Labor.AutoImport, "Error") || (StrLen(Addendum.Labor.AutoImport) = 0)
			Addendum.Labor.AutoImport	:= false

	; Infofenster das den Inhalt des Bildvorlagen-Ordners anzeigt
		Addendum.AddendumGui          	:= IniReadExt(compname	, "Infofenster_anzeigen"                      	, "ja"                 	)

	; Logbuch für Aktionen in der Karteikarte
		Addendum.PatLog                      	:= IniReadExt(compname	, "Logbuch_Patienten"                        	, "ja"                 	)

	; Hilfe beim Export von Dicom CD's im MicroDicom Programm
		Addendum.mDicom                     	:= IniReadExt(compname	, "MicroDicomExport_automatisieren"   	, "nein"                 	)

	; PopUpMenu integrieren
		Addendum.PopUpMenu              	:= IniReadExt(compname	, "PopUpMenu"                                  	, "ja"                 	)

	; die Addendum Toolbar starten
		Addendum.ToolbarThread            	:= IniReadExt(compname	, "Addendum_Toolbar"                         	, "nein"                 	)

	; Schnellrezept integrieren
		Addendum.Schnellrezept           	:= IniReadExt(compname	, "Schnellrezept"                                  	, "ja"                 	)

	; zeigt TrayTips
		Addendum.ShowTrayTips           	:= IniReadExt(compname	, "TrayTips_zeigen"                                	, "nein"                 	)

	; AutoOCR - Eintrag wird nur auf dem dazu berechtigten Client angezeigt
		Addendum.OCR.AutoOCR           	:= IniReadExt("OCR"      	, "AutoOCR"                                      	, "nein"                 	)

	; ermöglicht die sofortige Bearbeitung/Anzeige neuer Dateien
		Addendum.OCR.WatchFolder    	:= IniReadExt(compname	, "BefundOrdner_ueberwachen"           	, "ja"                 	)

	; Dauermedikamente, Dauerdiagnosen, Wartezimmereinträge ohne Nachfrage löschen
		Addendum.AutoDelete.Dauermed	:= IniReadExt(compname	, "Dauermedikamente_schnell_loeschen", "nein"               	)
		Addendum.AutoDelete.Dauerdia  	:= IniReadExt(compname	, "Dauerdiagnosen_schnell_loeschen" 	, "nein"                	)
		Addendum.AutoDelete.WZ           	:= IniReadExt(compname	, "Wartezimmer_schnell_loeschen"       	, "nein"               	)
		Addendum.AutoDelete.WZFilter 	:= IniReadExt(compname	, "Wartezimmer_loeschen_Filter"         	, ""                    	)


	; Heilmittelverordnung automatisch positionieren
		Addendum.HMV                        	:= IniReadExt(compname	, "AutoPos_Heilmittelverordnung"    		, "nein"               	)
		If (InStr(Addendum.HMV, "Error") || !Addendum.HMV)
			Addendum.HMV := false

	; aus der Fritzbox versendete Fax-Dateien mit Outlook weiter verabeiten
		RegRead,OutlookPath,HKEY_LOCAL_MACHINE, SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE, Path
		If (Addendum.Mail.Outlook          	:= OutlookPath ? true : false)
			Addendum.Mail.FaxMail        	:= IniReadExt(compname	, "Outlook_FaxMail_Manager"             	, "nein"                	)

	; nach dem Ende des Laborabrufes das Laborbuch anzeigen lassen
		Addendum.Labor.ZeigeLabJournal	:= IniReadExt("LaborAbruf", "Laborabruf_ZeigeJournal"               	, "nein"              	)

	; Boss PC Name (debugging features nur auf einem PC darstellen)
		Addendum.IamTheBoss                	:= IniReadExt(compname	, "IamTheBoss"                                                                	)
		If InStr(Addendum.IamTheBoss, "Error") || (StrLen(Addendum.IamTheBoss)=0)
			Addendum.IamTheBoss	:= false

	; Excel Helferlein
		Addendum.CImpf.Helfer             	:= IniReadExt(compname	, "CoronaImpfhelfer"                          	, "nein"                 	)
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

	Addendum.AlbisPID  := AlbisPID()
	Addendum.scriptPID := DllCall("GetCurrentProcessId")          	; die ScriptPID wird für das SkriptReload und die Interskript-Kommunikation benötigt

}

admInfoWindowSettings() {                                                                      	;-- AddendumGui/Infofenster

	; letzte Änderung: 07.11.2021

	; __________________________________________________________________________________________________________________________________________
	; Position des Addendumfenster in Albis
		tmp1 := IniReadExt(compname, "InfoFenster_Position", "y2 w400 h340")
		RegExMatch(tmp1, "y\s*(?<Y>\d+)\s*w\s*(?<W>\d+)\s*h\s*(?<H>\d+)", iWin)

	; __________________________________________________________________________________________________________________________________________
	;{ Einstellungen für die automatische Aufnahme eines Patienten in ein bestimmtes Wartezimmer
	; die Triggerung erfolgt anhand der Dokumentbezeichnung auf Basis bestimmter Stichwörter, die Bedingung doctitle und docstate müssen beide wahr sein
	; über die Zuweisungen können Patienten auf unterschiedliche Wartezimmer verteilt werden

	; Standardeinstellungen
		AutoWZData := [], AutoWZDataMismatch := 0
		std_doctitles 	:= "Antrag|Anforderung|Befundanforderung|Rentenantrag|LASV|Lageso|Lebensversicherung|MDK|DRV|Reha|Rehaantrag|"
							. 	 "Bundesagentur|Arbeitsagentur|Jobcenter|Polizei|Kriminalpolizei|Kripo|Rentenversicherung|Sozialgericht|"
							.  	 "Krankenkasse|Kasse|Privatversicherung|Ärztekammer|Kassenärztliche Vereinigung|KV|Hauskrankenpflege|Sozialdienst"
		std_docstates	:= "unausgefüllt|unbeantwortet|nicht beantwortet|unerledigt|nicht erledigt|unbearbeitet|zu bearbeiten|nicht bearbeitet|Nachfrage|Abfrage"
		std_assigns  	:=  "Ärztekammer|Kassenärztliche Vereinigung|KV:Anfragen;"
								. "Bundesagentur|Arbeitsagentur|Jobcenter:Anfragen;"
								. "LASV|Lageso:Anfragen;"
								. "Lebensversicherung:Anfragen;"
								. "MDK:Anfragen;"
								. "Hauskrankenpflege|Sozialdienst:Funktion;"
								. "Krankenkasse|Kasse|Privatversicherung:Anfragen;"
								. "Polizei|Krimalpolizei|Kripo:Anfragen;"
								. "Reha|Rehaantrag|Rentenversicherung|DRV:Anfragen;"
								. "Sozialgericht:Anfragen;"
								. "Antrag|Anforderung|Befundanforderung:Anfragen;"

	; Trigger erstellen oder laden
		Loop 2 {

			inikey 	:= A_Index = 1 ? "AutoWZ_Dokumenttitel" : "AutoWZ_Dokumentstatus"
			stddata	:= A_Index = 1 ? std_doctitles : std_docstates
			inival 	:= IniReadExt("InfoFenster", inikey , stddata)

		  ; Fehler in den Daten erkennen
			datamismatch := (!RegExMatch(inival, "^\s*\w.*\w\s*$") || InStr(inival, "ERROR") || StrLen(inival)=0
									|| (RegExMatch(inival, "[A-ZÄÖÜ].*[A-ZÄÖÜ]") && !InStr(inival, "|")))
			AutoWZData[A_Index] := datamismatch ? stddata : inival
			If (datamismatch && (!InStr(inival, "ERROR") || !StrLen(inival)=0)) {
				IniWrite, StrUtf8BytesToText(inival), % Addendum.Ini, % "InfoFenster", % inikey "_Backup"
				AutoWZDataMismatch += A_Index
			}

		}

	; AutoWZ Funktionsstatus
		AutoWZ := IniReadExt("InfoFenster", "AutoWZ", "Ja")
		If (InStr(AutoWZ, "ERROR") || StrLen(AutoWZ)=0) {
			AutoWZ := true
			inival := "Ja"
			IniWrite, StrUtf8BytesToText(inival), % Addendum.Ini, % "InfoFenster", % AutoWZ
		}

	; AutoWZ Trigger und Zuweisung zu Wartezimmern laden
		AWZAssigns := IniReadExt("InfoFenster", "AutoWZ_Wartezimmer_Zuweisung", std_assigns)
		If (InStr(AWZAssigns, "ERROR") || StrLen(AWZAssigns)=0) {
			AutoWZ := false
			AutoWZDataMismatch += 0x4
			IniWrite, StrUtf8BytesToText(std_assigns), % Addendum.Ini, % "InfoFenster", % "AutoWZ_Wartezimmer_Zuweisung"
		}
		AWZAssigns := RTrim(AWZAssigns, ";")

	; aktuelle Wartezimmer aus der local.ini holen
		IniRead, wartezimmer, % Addendum.Albis.LocalPath "\local.ini", Wartezimmer, Räume
		wartezimmer := "(" StrReplace(wartezimmer, ",", "|") ")"

	; Zuweisung nur übernehmen wenn das Wartezimmer in Albis existiert
		AutoWZAssigns := Object()
		keinWZ := Object()
		For each, item in StrSplit(AWZAssigns, ";") {
			RegExMatch(item, "^(?<Title>.*?):(?<toWZ>.*)$", assign)
			If RegExMatch(assigntoWZ, "i)" wartezimmer)
				AutoWZAssigns[assignTitle] := assignToWZ
			else
				If !keinWZ.Haskey(assignToWZ)
					keinWZ[assignToWZ] := 1
		}

	; Hinweis ausgeben wenn Wartezimmer nicht mehr vorhanden sind, so daß automatische Zuweisungen nicht statt finden werden
		If keinWZ.Count() {
			For WZ, n in keinWZ
				noassigns .= WZ ", "
			PraxTT("{Autowartezimmer}`n"
					.  "Es bestehen Verknüpfungen zu nicht mehr vorhandenen Wartezimmern (" RTrim(noassigns, ", ") ").`n"
					.  "Die Verknüpfungen können in der Addendum.ini [Infofenster] korrigiert werden.", "20 1")
		}


	;}

		Addendum.iWin := {"Y"                             	: iWinY
									, 	"W"                            	: iWinW
									,	"LVScanPool"              	: {"W"	: (iWinW - 10)
																			,	"R"	: IniReadExt(compname, "InfoFenster_BefundAnzahl", 7)}
									,	"ReIndex"                    	: false
									,	"rowprcs"                   	: 0
									,	"Init"                            	: false
									,	"AutoWZ"                  	: AutoWZ
									,	"AutoWZTitles"           	: AutoWZData[1]
									,	"AutoWZStates"           	: AutoWZData[2]
									,	"AutoWZDataMismatch"	: AutoWZDataMismatch                                                                                   		; Abfrage vor Dokumentimport?
									,	"AutoWZAssigns"          	: AutoWZAssigns
									,	"Impfstatistikdatum"     	: IniReadExt("Infofenster", "Impfstatistikdatum"                                      	, ""            	)
									; Spalte "3" , Sortierrichtung 1 = Aufsteigend : 0 = Absteigend, Listview scrollen zur Reihe Nr.
									,	"JournalSort"              	: IniReadExt(compname	, "Infofenster_JournalSortierung"                     	, "3 1 1"     	)
									,	"JournalPos"              	: IniReadExt(compname	, "Infofenster_JournalPosition"                            	, "0"           	)
									,	"ConfirmImport"           	: IniReadExt(compname	, "Infofenster_Import_Einzelbestaetigung"          	, "Ja"          	)
									,	"firstTab"                   	: IniReadExt(compname	, "Infofenster_aktuelles_Tab"                             	, "Patient"   	)
									,	"TProtDate"                 	: IniReadExt(compname	, "Infofenster_Tagesprotokoll_Datum"                	, "Heute"      	)
									, 	"TPClient"                     	: IniReadExt(compname	, "Infofenster_Tagesprotokoll_Client"               	, compname	)
									,	"Debug"                       	: IniReadExt(compname	, "Infofenster_Debug"                                     	, "Nein"     	)
									,	"ShowTools"                 	: IniReadExt(compname	, "Infofenster_Tools_zeigen"                            	, "Nein"     	)
									,	"LBDrucker"                  	: IniReadExt(compname	, "Infofenster_Laborblatt_Drucker"                   	, ""               	)
									,	"LBAnm"                       	: IniReadExt(compname	, "Infofenster_Laborblatt_Anmerkung_drucken"	, "Nein"     	)
									,	"LDTSearch"                	: IniReadExt(compname	, "Infofenster_Suche_in_LDT_Dateien"              	, "Nein"     	)
									,	"AbrHelfer"                	: IniReadExt(compname	, "Infofenster_Abrechnungshelfer"                    	, "Ja"         	)} ; Abrechnungshilfe einblenden

	; Korrekturen
		iw := Addendum.iWin
		Addendum.iWin.TProtDate     	:= StrLen(iw.TProtDate)=0         	|| InStr(iw.TProtDate       	, "Error")	? "Heute"   	: iw.TProtDate
		Addendum.iWin.TPClient        	:= StrLen(iw.TPClient)=0           	|| InStr(iw.TPClient         	, "Error")	? compname	: iw.TPClient
		Addendum.iWin.firstTab         	:= StrLen(iw.firstTab)=0      	      	|| InStr(iw.firstTab          	, "Error")	? "Journal"	: iw.firstTab
		Addendum.iWin.Debug         	:= StrLen(iw.Debug)=0      	      	|| InStr(iw.Debug          	, "Error")	? false       	: iw.Debug
		Addendum.iWin.AbrHelfer        	:= StrLen(iw.AbrHelfer)=0         	|| InStr(iw.AbrHelfer         	, "Error")	? true        	: iw.AbrHelfer
		Addendum.iWin.ConfirmImport	:= StrLen(iw.ConfirmImport)=0   	|| InStr(iw.ConfirmImport 	, "Error")	? true        	: iw.ConfirmImport

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

		Addendum.ImportRunning	:= false
		Addendum.iWin.Init        	:= 0

}

admModule() {                                                                                       	;-- Liste der verfügbaren Module

	; Einstellungen für Module welche auf diesem Client benutzt werden dürfen
		IniReadC   	:= 0
		admModule	:= IniReadExt(compname, "Module")
		Loop {

			modulNr := A_Index
			iniModul := IniReadExt("Module", "Modul" SubStr("00" modulNr, -1))

			If InStr(iniModul, "Error") {           ; Abbruch wenn keine weiteren Module vorhanden sind
				IniReadC ++
				If (IniReadC <=5)
					continue
				break
			} else
				IniReadC := 0

			m := StrSplit(iniModul, "|")
			If (m.1 = "Auth")
				If modulNr not in %admModule%
					continue

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

	Addendum.Laborjournal.AutoView	:= IniReadExt(compname, "Laborjournal_AutoAnzeige"   	,  "Nein"            	)
	Addendum.Laborjournal.StartTime	:= IniReadExt(compname, "Laborjournal_Startzeit"           	,  "---"                	)

	For key, val in Addendum.Laborjournal
		If (val = "ERROR" || val = "---")
			Addendum.Laborjournal[key] := ""

}

admLanProperties() {                                                                              	;-- Netzwerkeinstellungen / Client PC Namen / IP-Adressen

	; prüft auf den korrekt gesetzten Computernamen
		CName := IniReadExt(compname, "ComputerName")
		If (CName <> A_ComputerName)
			IniWrite, % StrUtf8BytesToText(A_ComputerName), % Addendum.Ini, % compname, % "ComputerName"

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
		ServerPC   	:= InStr(ServerPC, "ERROR") ? "" : ServerPC

		Addendum.LAN.IamServer := ServerPC=compname ? true : false

		Server	:= IniReadExt("Addendum", "Addendum_ServerAdresse", 0)
		Server  	:= !InStr(Server, "ERROR") ? Server : ""
		RegExMatch(Server, "(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s*\:\s*(?<port>\d+)", tcp_)
		If (tcp_ip && tcp_port)
			Addendum.LAN.Server := {"ip":tcp_ip, "port":tcp_port}

}

admMailAndHolidays() {                                                                          	;-- eigene Mailadresse und Urlaubsdaten

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
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; PDF Reader: 2 Programme sind möglich. Der FoxitReader wird nur für's Signieren verwendet (Startvorgang dauert zu lange)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Addendum.PDF.Reader                       	:= IniReadExt("ScanPool"     	, "PDFReader"                           	, AddendumDir "\include\FoxitReader\FoxitReaderPortable\FoxitReaderPortable.exe")
		Addendum.PDF.ReaderName               	:= IniReadExt("ScanPool"     	, "PDFReaderName"     	           	, "FoxitReader")
		Addendum.PDF.ReaderWinClass           	:= IniReadExt("ScanPool"     	, "PDFReaderWinClass"             	, "classFoxitReader")
		Addendum.PDF.ReaderAlternative         	:= IniReadExt("ScanPool"     	, "PDFReaderAlternative"           	, "")
		Addendum.PDF.ReaderAlternativeName	:= IniReadExt("ScanPool"     	, "PDFReaderAlternativeName"  	, "")

		altReader := Addendum.PDF.ReaderAlternative, altName := Addendum.PDF.ReaderAlternativeName
		If (altName = "ERROR") || (altReader = "ERROR") || (StrLen(altName) = 0)	|| (StrLen(altReader) = 0){
				Addendum.PDF.ReaderAlternativeName	:= ""
				Addendum.PDF.ReaderAlternative       	:= ""
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; PDF signieren und Statistik
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Addendum.PDF.SignatureCount          	:= IniReadExt("ScanPool"     	, "SignatureCount")
		Addendum.PDF.SignaturePages          	:= IniReadExt("ScanPool"     	, "SignaturePages")

		If InStr(Addendum.PDF.SignatureCount, "Error") || (StrLen(Trim(Addendum.PDF.SignatureCount)) = 0)
			Addendum.PDF.SignatureCount := 0
		If  InStr(Addendum.PDF.SignaturePages, "Error") || (StrLen(Trim(Addendum.PDF.SignaturePages)) = 0)
			Addendum.PDF.SignaturePages := 0

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; FoxitReader Signatureinstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Addendum.PDF.SignatureWidth          	:= IniReadExt("ScanPool"     	, "Signature_Breite"                            	, 50)
		Addendum.PDF.SignatureHeight        	:= IniReadExt("ScanPool"     	, "Signature_Hoehe"                            	, 25)
		Addendum.PDF.Ort                              	:= IniReadExt("ScanPool"     	, "Ort"                                               	,, true)
		Addendum.PDF.Grund                       	:= IniReadExt("ScanPool"     	, "Grund"                                           	,, true)
		Addendum.PDF.SignierenAls                 	:= IniReadExt("ScanPool"     	, "SignierenAls"                                  	,, true)
		Addendum.PDF.Darstellungstyp          	:= IniReadExt("ScanPool"     	, "DarstellungsTyp"                              	,, false)
		Addendum.PDF.PasswortOn                	:= IniReadExt("ScanPool"     	, "PasswortOn"                                   	, "nein")
		Addendum.PDF.PatAkteSofortOeffnen  	:= IniReadExt("ScanPool"     	, "Patientenakte_sofort_oeffnen"            	, "ja")
		Addendum.PDF.DokumentSperren        	:= IniReadExt("ScanPool"     	, "DokumentNachDerSignierungSperren", "ja")

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Albis Karteikartenkürzel für das Einfügen von PDF Dokumenten in die Akte
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Addendum.PDF.ScanKuerzel               	:= IniReadExt("ScanPool"     	, "Scan"                                             	, "pdf")
		Addendum.PDF.ScanKuerzel               	:= InStr(Addendum.PDF.ScanKuerzel, "ERROR") ? "pdf" : Addendum.PDF.ScanKuerzel

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Fritzfaxbox - PDF Faxe in den Befundordner kopieren oder verschieben (state = Aus [/Faxe kopieren/Faxe verschieben])
	; wenn auf der Fritzbox gespeicherte Faxbefunde automatisch in den Befundordner kopiert werden sollen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Addendum.PDF.Faxbox.State               	:= IniReadExt("ScanPool", "FritzFaxbox_Status", "Aus")
		Addendum.PDF.Faxbox.Path                	:= IniReadExt("ScanPool", "FritzFaxbox_Verzeichnis")
		Addendum.PDF.Faxbox.Interval             	:= IniReadExt("ScanPool", "FritzFaxbox_Abrufintervall", 10)
		Addendum.PDF.Faxbox.Client               	:= IniReadExt("ScanPool", "FritzFaxbox_Client")


		If FileExist(Addendum.PDF.Faxbox.Telefonbuch := IniReadExt("ScanPool", "FritzFaxbox_Telefonnummern")) {
			Addendum.PDF.Faxbox.Telefon := Object()
			temp := FileOpen(Addendum.PDF.Faxbox.Telefonbuch, "r", "UTF-8").Read()
			For each, line in StrSplit(temp, "`n", "`r")
				If (StrSplit(line, "|").1 ~= "i)\(.*Fax.*?\)")
					Addendum.PDF.Faxbox.Telefon["#" StrSplit(line, "|").2] := StrSplit(line, "|").1
		}

		If FileExist(path := IniReadExt("ScanPool", "FritzFaxbox_Filterliste")) {
			Addendum.PDF.Faxbox.Blacklist	:= Object()
			temp := FileOpen(path, "r", "UTF-8").Read()
			For each, line in StrSplit(temp, "`n", "`r") {
				If line && !RegExMatch(line, "^#") {
					line 	:= RegExReplace(line, ";.*")
					tel 	:= RegExReplace(line, "[^\d]")  	; die Telefonnummer
					rm 	:= RegExReplace(line, "[^\*]") 	; das Löschzeichen
					Addendum.PDF.Faxbox.Blacklist[tel] := rm ? 1 : 0
				}
			}
		}
		else {
			filecontent =
			(LTrim
			# Legen Sie in diese Datei zeilenweise die Telefonnummer (auch mit Sonderzeichen/Leerzeichen) fest von den kein Fax in den Befundordner kopiert werden soll.
			# Schreiben hinter die Telefonnummer ein Sternchen wenn Faxdateien von dieser Telefonnummer sofort gelöscht werden sollen
			# Kommentare zur Nummer können überall nach einem Semikolon geschrieben werden
			# Beispiel
			# 040 233 234 343
			# 030/3614581*    	; ausschließlich Werbung
			)
			FileOpen(path, "w", "UTF-8").Write(filecontent)
		}

}

admOCRSettings() {                                                                                	;-- PDF Bearbeitung (OCR / AutoNaming / Watchfolder)

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
		If !RegExMatch(Addendum.OCR.AutoOCRDelay, "^\d+$") {                  	; ## Abfrage integrieren
			Addendum.OCR.AutoOCRDelay := 30
			IniWrite, % Addendum.OCR.AutoOCRDelay, % Addendum.Ini, % "OCR", % "AutoOCR_Startverzoegerung"
		}

	; Protokolle
		Addendum.OCR.WFLog	:= IniReadExt(compname, "Infofenster_WatchFolder_Protokoll", "Ja"	)
		Addendum.OCR.TimeLog	:= Addendum.OCR.Client = compname ? IniReadExt(compname, "Infofenster_OCRZeit_Protokoll", "Ja") : false

	; ----------------------------------------------------
	; OCR Variablen werden vorbereitet
	; ----------------------------------------------------
	; wenn true dann ignoriert Watchfolder alle Dateiveränderungen solange nicht auf false gegangen wird
		Addendum.OCR.PauseWF 	:= false
		Addendum.OCR.WFIgnore  	:= Array()

	; Zähler der Texterkennung, aktuell und seit Programmstart
		Addendum.OCR.staticFileCount	:= 0
		Addendum.OCR.filecount         	:= 0

}

admPreise() {                                                                                        	;-- Preise für Portokosten und anderes

	lastencoding := A_FileEncoding
	FileEncoding, UTF-16
	IniRead, Preise, % Addendum.Dir "\Addendum.ini", % "Preise"
	FileEncoding, % lastencoding

	For line, preis in StrSplit(Preise, "`n", "`r") {
		If RegExMatch(preis, "i)^\s*Preis_(?<Bezeichnung>[\pL\d\-\_]+)\s*\=\s*(?<Euro>\d+)[\.\,]*(?<Cent>\d+)*", K)
			Addendum.Preise[KBezeichnung] := KEuro "." KCent . (StrLen(KCent) < 2 ? SubStr("00", 1, 2-StrLen(KCent)) : "")
		else if RegExMatch(preis, "i)\s*Porto_(?<Bezeichnung>[\pLßäöü\d\-\_]+)\s*\=\s*(?<Euro>\d+)[\.\,]*(?<Cent>\d+)*", K) {
			If !IsObject(Addendum.Preise.Porto)
				Addendum.Preise.Porto := Object()
			Addendum.Preise.Porto[KBezeichnung] := KEuro "." KCent . (StrLen(KCent) < 2 ? SubStr("00", 1, 2-StrLen(KCent)) : "")
		}
	}

}

admPraxisDaten() {                                                                                  	;-- Daten Ihrer Praxis

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
		Addendum.Praxis.Urlaub := vacation.ConvertHolidaysString(Urlaub)

  ; Telefone und Faxgeräte
	Addendum.Praxis.Telefon 	:= Array()
	Addendum.Praxis.Fax     	:= Array()
	Loop {
		Tel 	:= IniReadExt("Allgemeines", "Telefon" A_Index)
		Fax 	:= IniReadExt("Allgemeines", "Fax" 	A_Index)
		Tel 	:= InStr(Tel, "ERROR")	? "" : Tel
		Fax 	:= InStr(Fax, "ERROR") 	? "" : Fax
		If (!Fax && !Tel)
			break
		If Tel
			Addendum.Praxis.Telefon[A_Index] := Tel
		If Fax
			Addendum.Praxis.Fax[A_Index]    	:= Fax
	}

  ; Praxisärzte einlesen (!noch ohne Funktion!)
	Addendum.Praxis.Arzt := Array()
	Loop {

		Arzt   	:= IniReadExt("Allgemeines", "Arzt" A_Index "Name")
		BSNR	:= IniReadExt("Allgemeines", "Arzt" A_Index "BSNR")
		LANR	:= IniReadExt("Allgemeines", "Arzt" A_Index "LANR")
		Fach		:= IniReadExt("Allgemeines", "Arzt" A_Index "Abrechnung")

		If InStr(Arzt, "Error") || (Arzt = "") {
			Arzt := Addendum.Praxis.Name
			ArztSucheBeenden := true
		}

		If InStr(BSNR, "Error") || (BSNR = "") {
			BSNR := "00000000"
			ArztSucheBeenden := true
		}

		If InStr(LANR, "Error") || (LANR = "") {
			LANR := "000000000"
			ArztSucheBeenden := true
		}

		If InStr(Fach, "Error") || (Fach = "") {
			Fach := "Hausarzt"
			ArztSucheBeenden := true
		}

		Addendum.Praxis.Arzt.Push({"Name":Arzt, "Fach":Fach, "BSNR":BSNR, "LANR":LANR})

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

	; Daten für Abrechnungshelfer (Infofenster)
	; ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ;{
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

		; ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ
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
	;}

	; Verzögerung bis zum Schliessen des Dialoges "Patient hat in diesem Quartal seine Chipkarte..." (0 = keine Verzögerung)
	; ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ;{
		Addendum.noChippie := IniReadExt(compname, "keineChipkarte", 3)
	;}

	; Behördengebühren werden über die ini Datei verwaltet
	; ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ;{
		Behoerdengebuehr := {"LASV":"25.00", "JVEG":"25.00|55.00|45.00|90.00", "DRV":"28.20|35.00", "BAfA":"32.50"}
		tmp1 := IniReadExt("Addendum", "JVEG")
		tmp2 := IniReadExt("Addendum", "DRV")
		Addendum.Abrechnung.JVEG 	:= StrSplit(tmp1, "|")
		Addendum.Abrechnung.DRV  	:= StrSplit(tmp2, "|")
		Addendum.Abrechnung.LASV  	:= IniReadExt("Addendum", "LASV")
		Addendum.Abrechnung.BAfA  	:= IniReadExt("Addendum", "BAfA")
		For Behoerde, Gebuehr in Addendum.Abrechnung
			If !RegExMatch(Gebuehr, "\d+\.\d+") && !IsObject(Gebuehr)
				Addendum.Abrechnung[Behoerde] := Behoerdengebuehr[Behoerde]
			else if IsObject(Gebuehr)
				For sub, subGebuehr in Gebuehr
					If !RegExMatch(subGebuehr, "\d+\.\d+") {
						tmp := StrSplit(Behoerdengebuehr[Behoerde], "|")
						Addendum.Abrechnung[Behoerde] := tmp
					}
	;}

	; Gebührenberechnung Coronaimpfung im Addendumskript
	; ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ᚔ ;{
		If ((Addendum.CImpf.CNRRun 	:= IniReadExt(compname, "COVID19_ImpfGebuehr_Skript", "Ja")) = true) {
			Addendum.CImpf.CNRWin  	:= IniReadExt(compname, "COVID19_ImpfGebuehrfenster_Position")
			If InStr(Addendum.CImpf.CNRWin, "ERROR")
				Addendum.CImpf.CNRWin := ""
		}
	;}


	; BMP Häkchen


}

admStandard() {                                                                                    	;-- Einstellungen für Addendum - Fenster und Dialoge

	Addendum.Default.Font                  	:= IniReadExt("Addendum"	, "StandardFont"                        	)
	Addendum.Default.BoldFont          	:= IniReadExt("Addendum"	, "StandardBoldFont"                    	)
	Addendum.Default.FontSize           	:= IniReadExt("Addendum"	, "StandardFontSize"                    	)
	Addendum.Default.FntColor            	:= IniReadExt("Addendum"	, "DefaultFntColor"                     	)
	Addendum.Default.BGColor            	:= IniReadExt("Addendum"	, "DefaultBGColor"                     	)
	Addendum.Default.BGColor1          	:= IniReadExt("Addendum"	, "DefaultBGColor1"                   	)
	Addendum.Default.BGColor2          	:= IniReadExt("Addendum"	, "DefaultBGColor2"                   	)
	Addendum.Default.BGColor3          	:= IniReadExt("Addendum"	, "DefaultBGColor3"                   	)
	Addendum.Default.PRGColor          	:= IniReadExt("Addendum"	, "DefaultProgressColor", "FFFFFF"	)

}

admTagesProtokoll() {                                                                            	;-- Anlegen/Lesen des aktuellen Tagesprotokolls

	; Anlegen eines neuen Tagesprotollordners falls dieser noch nicht vorhanden
		If !InStr(FileExist(Addendum.TPPath), "D") {
			FileCreateDir, % Addendum.TPPath
			If ErrorLevel 	{
				MsgBox,,	% "Addendum für AlbisOnWindows"
							,	% "Ein neues Verzeichnis für die Tagesprotokolldateien konnte nicht angelegt werden.`n"
							. 		"Das Skript muss jetzt beendet werden!"
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

admWartezimmer() {                                                                             	;-- Wartezimmer-Automatisierungen

	Addendum.WZ.AutoWZ 	:= IniReadExt(compname, "WarteZimmer_schnell_entfernen")
	Addendum.WZ.AutoDelete	:= IniReadExt(compname, "Wartezimmer_schnell_loeschen", "Nein")
	If (!Addendum.WZ.AutoDelete || InStr(Addendum.WZ.AutoDelete, "ERROR"))
		Addendum.WZ.AutoDelete := false

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
return {	"BotName"        	: BotName
		, 	"Token"             	: IniReadExt("Telegram", BotName "_Token")
		, 	"ChatID"            	: IniReadExt("Telegram", BotName "_ChatID")BotChatID
		, 	"Active"              	: IniReadExt("Telegram", BotName "_Active", 0)
		, 	"BotLastMsg"      	: IniReadExt("Telegram", BotName "_LastMsg", "--")
		, 	"BotLastMsgTime"	: IniReadExt("Telegram", BotName "_LastMsgTime", "000000")}
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

class vacation {                                                                                        	;-- Funktionsklasse Datumsberechnungen (Urlaub)

	; ACHTUNG: benötigt Addendum.Praxis.Urlaub/Sprechstunde als globales Objekt
	;
	; Funktionen:    1. 	Daten aus Addendum.ini werden geparst. die Urlaubstage können mit relativ freier Schreibweise in der Ini-Datei eingetragen sein
	;                          	einzelne Tage 	:  	in der Form 01.01.2022 auch als 1.1.22
	;                           Datumbereiche	:	05.06.2022-06.06.2022 oder 05.06.-06.06.2022 oder 5.-6.6.22 sowie Jahreswechsel
	;                       	Achtung        	:	zwingend zur Unterscheidung sind Punkte und Minuszeichen. Leerzeichen, andere Zeichen können zwischen den Datumzahlen
	;
	; letzte Änderung: 18.12.2021

	ConvertHolidaysString(holidays:="")             	{                                     	;-- Ini-String

		/*  parst einen String welcher Urlaubstage enthält

			Beispielformat: 01.03.2021, 08.10.-14.10.2021, 25.-26.12.2021, 29.-30.12.

			- die einzelnen Tage oder Bereiche müssen durch ein Komma getrennt sein
			- fehlt eine Jahreszahl wird diese mit dem aktuellen Jahr gleichgesetzt
			- ACHTUNG: auch wenn eine recht flexible Schreibweise des Datum möglich ist,
								 werden fehlerhafte Datumzahlen nicht erkannt

		*/

			Debug := false

		  ; Urlaubszeiten mit einem Datum in der Vergangenheit werden aussortiert.
			static rxUrlaub

			If !rxUrlaub {
				rxUrlaub1	:= "(?<StartD>\d{1,2})\.(?<StartM>\d{1,2})*\.*(?<StartY>\d{2,4})*"
				rxUrlaub2	:= "\s*-\s*(?<EndD>\d{1,2})\.(?<EndM>\d{1,2})\.*(?<EndY>\d{2,4})*"
				rxUrlaub 	:= "(" rxUrlaub1 rxUrlaub2 ")|(" StrReplace(rxUrlaub1, "Start") ")"
			}

			spos := 1, AToday := A_YYYY . A_MM . A_DD, AYearNow := SubStr(A_YYYY, 1, 2)
			vacations := Array()

		  ; Abbruch wenn, nichts oder ein falsche String übergeben wurde
			If !holidays
				return

		 ; Datum für Datum extrahieren
			while (spos := RegExMatch(holidays, rxUrlaub, Px, spos)) {

			  ; Stringposition weiterrücken
				spos  	+= StrLen(Px)

			  ; nur ein Datum
				PxStartD  	:= PxD	? PxD	: PxStartD
				PxStartM  	:= PxM	? PxM 	: PxStartM
				PxStartY  	:= PxY	? PxY 	: PxStartY

			  ; Debug
				Debug && PxD && PxM ? SciTEOutput("PxDMY: " PxD "." PxM "." PxY)

			 ; eine zweistellige Jahreszahl auf eine vierstellige Jahreszahl ändern, leere Jahreszahlen ersetzen
				PxStartY	:= StrLen(PxStartY)	= 2 	? AYearNow . PxStartY 	: PxStartY	? PxStartY 	: PxEndY 	? PxEndY 	: A_YYYY
				PxEndY 	:= StrLen(PxEndY)	= 2 	? AYearNow . PxEndY  	: PxEndY	? PxEndY	: PxEndD 	? A_YYYY	: ""

			  ; fehlenden Monat ersetzen
				PxStartM 	:= PxStartM ? PxStartM : PxEndM

			  ; Debug
				Debug ? SciTEOutput("Px: " px "`nPxStart: " PxStartD "." PxStartM "." PxStartY "`nPxEnd: " PxEndD "." PxEndM "." PxEndY )

			  ; Monate und Tage 2-stellig auffüllen
				PxStartM 	:= SubStr("00" PxStartM	, -1)
				PxStartD 	:= SubStr("00" PxStartD	, -1)
				PxEndM   	:= SubStr("00" PxEndM	, -1)
				PxEndD   	:= SubStr("00" PxEndD	, -1)

			  ; die Formatierung wird auf YYYYMMDD geändert
				AYearS 	:= (!PxStartY && !PxEndY)	? A_YYYY : PxStartY ? PxStartY 	: PxEndY 	? PxEndY   ; Jahr
				AYearE 	:= (!PxStartY && !PxEndY)	? A_YYYY : PxEndY ? PxEndY 	: PxStartY 	? PxStartY   ; Jahr
				ADayE  	:= (PxEndM 	&& PxEndD)	? AYearE . PxEndM . PxEndD : ""
				;~ AYearS   	:= PxStartM < PxEndM ? AYears + 1
				ADayS  	:= AYearS . PxStartM . PxStartD

			  ; ein Datumsbereich wird anders hinterlegt, als ein einzelner freier Tag
				If (ADayS && ADayE && !this.PeriodExists(ADayS, ADayE))
					vacations.Push({"IsPeriod":true	, "firstday":ADayS, "lastday":ADayE})
				else If (ADayS && !ADayE && !this.PeriodExists(ADayS))
					vacations.Push({"IsPeriod":false	, "day":ADayS})

				PxD := PxM := PxY := PxStartD := PxStartM := PxStartY := PxEndD := PxEndM := PxEndY := ""

			  ; Debug
				Debug ? SciTEOutput("`n")

			}

	return vacations
	}

	PeriodExists(firstday, lastday:="")                	{                                    	;-- Datum oder Datumsbereich suchen

		If !firstday && !lastday
			return false

		For hdNR, holidays in Addendum.Praxis.Urlaub
			If (firstday && lastDate) {
				If (firstday=holidays.firstday && lastday=holidays.lastday)
					return hdNR
			}
			else {
				If (!holiday.IsPeriod && firstday=holidays.day)
					return hdNR
			}

	return false
	}

	DateIsHoliday(datestring)                            	{                                     	;-- ermittelt ob ein übergebenes Datum innerhalb eines Praxisurlaub liegt

	  ; automatisch 4-stelliges Jahresformat (funktioniert bis 2099)
		If RegExMatch(datestring, "(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>(\d{2}|\d{4}))", t)
			datestring := SubStr(SubStr(A_YYYY, 1, 2) . tY, -2) . tM . tD

		For hdNR, holidays in Addendum.Praxis.Urlaub
			If holidays.IsPeriod && (datestring>=holidays.firstday && datestring<=holidays.lastday)
				return hdNR
			else if (!holidays.IsPeriod && datestring=holidays.day)
				return hdNR

	return false
	}

	isConsultationTime(date:="", time:="", ByRef whynotwork:="") {            	;-- prüft ob zu diesem Zeitpunkt Sprechstunde (Konsultationszeit) ist

	  ; benötigt wird das durch .ConvertHolidaysString() erstellte Objekt in Addendum.Praxis.Urlaub

	  ; über die ByRef Variable "whynotwork" wird zurückgegeben welcher Status am untersuchten Zeitpunkt vorliegt,
	  ; wie 	1= regulärer arbeitsfreier Tag
	  ; 		2= Urlaubstag/Feiertag
	  ;			3= außerhalb der Sprechzeiten  ###(a ist vor, b ist nach der Sprechstunde und c ist zwischen den Sprechstunden)

	  ; letzte Änderung: 15.01.2022

		static reasons := {"nw1a":"regulär freier Arbeitstag", "nw1b":"arbeitsfreies Wochende", "nw2":"Urlaub", "nw3a":"vor der Sprechstunde", "nw3b":"nach der Sprechstunde", "nw3c":"zwischen den Sprechstunden" }

	  ; aktueller Tag oder/und Uhrzeit einstellen wenn sie leer sind
		time	:= time	? time	: A_Hour ":" A_Min
		date	:= date	? date	: A_DD "." A_MM "." A_YYYY

	  ; Datum/Zeitstring konvertieren
		RegExMatch(time, "(?<Hour>\d{1,2})\s*[:\.\-]\s*(?<Min>\d{1,2})\s*[:\.\-]*\s*(?<Sek>\d{1,2})*", t)
		time := SubStr("00" tHour, -1) . SubStr("00" tMin, -1)
		time := SubStr(time, 1, 4)
		If !RegExMatch(date, "\d+.\d+.\d+")
			date := A_DD "." A_MM "." A_YYYY
		date := ConvertToDBASEDate(date)


	  ; Wochentag berechnen
		dow	:= DayOfWeek(date, "full", "yyyyMMdd")

	 ; an diesem ist Wochentag ist keine Sprechstunde
		If !Addendum.Praxis.Sprechstunde.haskey(dow) {
			whynotwork := {"reason":reasons["nw1" (RegExMatch(dow, "(Samstag|Sonntag)") ? "b":"a")], "level":1, "weekday":dow}
			return false
		}
		else If this.DateIsHoliday(date) {
			whynotwork := {"reason":reasons["nw2"], "level":2, "weekday":dow}
			return false
		}

	  ; Zeitpunkt liegt während oder außerhalb der Sprechstunde
		isduring  	:= false
		conTimes	:= RegExReplace(Addendum.Praxis.Sprechstunde[dow], "(\s*,\s*)|(\d)\s+(\d)", "$2|$3)")
		For conNr, timestr in StrSplit(conTimes, "|") {
			If RegExMatch(timestr, "(?<Hour1>\d{1,2})\s*:\s*(?<Min1>\d{1,2})\s*\-\s*(?<Hour2>\d{1,2})\s*:\s*(?<Min2>\d{1,2})", con) {
				conStart:= SubStr("00" conHour1, -1) . SubStr("00" conMin1, -1)
				conEnd	:= SubStr("00" conHour2, -1) . SubStr("00" conMin2, -1)
				If (time>=conStart && time<=conEnd)                                          	; liegt in der Sprechstunde
					isduring := true
			}
		}

	  ; außerhalb der Sprechzeiten whynotwork
		If !isduring {
			RegExMatch(conTimes, "^\s*(?<Hour1>\d{1,2})\s*:\s*(?<Min1>\d{1,2}).*?(?<Hour2>\d{1,2})\s*:\s*(?<Min2>\d{1,2})\s*$", con)
			conDayS 	:= SubStr("00" conHour1, -1) . SubStr("00" conMin1, -1)
			conDayE	:= SubStr("00" conHour2, -1) . SubStr("00" conMin2, -1)
			sublevel  	:= time<conDayS ? "a" : time>conDayE ? "b" : "c"
			whynotwork := {"reason":reasons["nw3" sublevel], "level":2, "sublevel": sublevel, "weekday":dow}
		}


	return isduring
	}

}




; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                 liest Informationen aus der Addendum.ini ein für das globale Objekt Addendum
;                                                                              	!diese Bibliothek enthält Funktionen für Einstellungen des Addendum Hauptskriptes!
;                                                            	by Ixiko started in September 2017 - last change 11.11.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

admObjekte() {                                                                                       	;-- Addendum-Objekt erweitern

 		Addendum.Chroniker            	:= Array()          	; Patient ID's der Chroniker
		Addendum.Geriatrisch           	:= Array()         	; Patient ID's für geriatrisches Basisassement
		Addendum.Hooks                  	:= Object()       	; enthält Adressen der Callbacks der Hookprozeduren und anderes
		Addendum.Labor	    	        	:= Object()			; Laborabruf und andere Daten für Laborprozesse
		Addendum.Laborabruf            	:= Object()
		Addendum.LAN	       	        	:= Object()			; LAN Kommunikationseinstellungen
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
		Addendum.Thread                  	:= Object()          	; enthält die PID's gestarteter Threads (z.B. AddendumToolbar)
		Addendum.Threads                 	:= Object()         	; Skriptcode für Prozess-Threads
		Addendum.Windows                	:= Object()        	; Fenstereinstellungen für andere Programme
		Addendum.UndoRedo             	:= Object()          	; Undo/Redo Textbuffer
		Addendum.UndoRedo.List        	:= Array()           	; Undo/Redo Textbuffer

}

admVerzeichnisse() {                                                                             	;-- Programm- und Datenverzeichnisse (!ALS ERSTES AUFRUFEN! nach admObjekte)

		Addendum.AddendumDir                	:= AddendumDir
		Addendum.AddendumIni               	:= AddendumDir "\Addendum.ini"

		Addendum.DosisDokumentenPfad   	:= IniReadExt("Addendum" 	, "DosisDokumentenPfad"	)              	; MS Word Dateien mit eigenen Hinweisen zu Medikamentendosierungen
		Addendum.DBPath                          	:= IniReadExt("Addendum" 	, "AddendumDBPath"     	)              	; Datenbankverzeichnis
		Addendum.BefundOrdner                	:= IniReadExt("ScanPool"     	, "BefundOrdner"            	)             	; BefundOrdner = Scan-Ordner für neue Befundzugänge
		Addendum.VideoOrdner                	:= IniReadExt("ScanPool"     	, "VideoOrdner"             	)             	; Video's z.B. aus Dicom CD's (CT Bilder u.a.)
		Addendum.AlbisExe                         	:= IniReadExt("Albis"           	, "AlbisExe"                    	)             	; Pfad zum Albis-Stammverzeichnis

		Addendum.TPPath	                        	:= Addendum.DBPath "\Tagesprotokolle\" A_YYYY                     	; Tagesprotokollverzeichnis
		Addendum.TPFullPath                      	:= Addendum.TPPath "\" A_MM "-" A_MMMM "_TP.txt" 	               	; Name des aktuellen Tagesprotokolls

		Addendum.ExportOrdner                	:= IniReadExt("ScanPool"   	, "ExportOrdner"        	, Addendum.BefundOrdner "\Export")  ;### ÜBERPRÜFUNG ANLEGEN
		Addendum.AdditionalData_Path       	:= IniReadExt("Addendum"	, "AdditionalData_Path" 	, AddendumDir "\include\Daten")

	; Ordnerstruktur anlegen für die Verwaltung eingegangener Befunde, wenn ein BefundOrder in der ini angegeben wurde
		If !InStr(Addendum.BefundOrdner, "Error") && RegExMatch(Addendum.BefundOrdner, "[A-Z]\:\\") {

			If !InStr(FileExist(Addendum.BefundOrdner "\Text"), "D")
				FileCreateDir, % Addendum.BefundOrdner "\Text"

			If !InStr(FileExist(Addendum.BefundOrdner "\Backup"), "D")
				FileCreateDir, % Addendum.BefundOrdner "\Backup"

			If !InStr(FileExist(Addendum.BefundOrdner "\Backup\Importiert"), "D")
				FileCreateDir, % Addendum.BefundOrdner "\Backup\Importiert"

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

admAlbisMenuAufrufe() {                                                                       	;-- Menubezeichnungen und wm_command Befehle

		Addendum.AlbisMenu := Object()        	; wm_command Daten für Dialogaufrufe in Albis
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

		tmp  := Trim(IniReadExt(compname, "AutoSize_Albis", "2k:nein,4k:nein" ))
		Addendum.AlbisLocationChange  	:= RegExAutoPos(tmp)                                                                                           	; automatische Positionierung von Albis ist an
		Addendum.AlbisGroesse            	:= IniReadExt(compname	, "Position_Albis"                               	, "w1922 h1080"	)  	; Größeneinstellung
		Addendum.GVUAutomation			:= IniReadExt(compname	, "GVU_automatisieren"                    	, "nein"              	)  	; Automatisierung der GVU Formulare
		Addendum.PDFSignieren   	    	:= IniReadExt(compname	, "FoxitPdfSignature_automatisieren" 	, "nein"              	)  	; Automatisierung PDF Signierung
		Addendum.PDFSignieren_Kuerzel	:= IniReadExt(compname	, "FoxitPdfSignature__Kuerzel" 	      	, "Insert"            	)  	; Tastaturkürzel zum Auslösen des Signaturvorganges
		Addendum.AutoCloseFoxitTab   	:= IniReadExt(compname	, "FoxitTabSchliessenNachSignierung"	, "Ja"                 	)  	; signiertes Dokument automatisch schliessen
		Addendum.Labor.AbrufTimer         := IniReadExt(compname	, "Laborabruf_Timer"           	   	    	, "nein"              	) 	; zeitgesteuerter Laborabruf
		Addendum.Labor.AutoAbruf 		    := IniReadExt(compname	, "Laborabruf_automatisieren"         	, "nein"              	) 	; Laborabruf bei manuellem Start automatisieren
		Addendum.AddendumGui          	:= IniReadExt(compname	, "Infofenster_anzeigen"                  	, "ja"                 	)  	; Infofenster das den Inhalt des Bildvorlagen-Ordners anzeigt
		Addendum.PatLog                      	:= IniReadExt(compname	, "Logbuch_Patienten"                     	, "ja"                 	)  	; Logbuch für Aktionen in der Karteikarte
		Addendum.mDicom                     	:= IniReadExt(compname	, "MicroDicomExport_automatisieren"	, "ja"                 	)  	; Hilfe beim Export von Dicom CD's im MicroDicom Programm
		Addendum.PopUpMenu              	:= IniReadExt(compname	, "PopUpMenu"                              	, "ja"                 	)  	; PopUpMenu integrieren
		Addendum.ToolbarThread            	:= IniReadExt(compname	, "Toolbar_anzeigen"                         	, "nein"                 	)  	; die Addendum Toolbar starten
		Addendum.Schnellrezept           	:= IniReadExt("Addendum"	, "Schnellrezept"                              	, "ja"                 	)  	; Schnellrezept integrieren
		Addendum.ShowTrayTips           	:= IniReadExt("Addendum"	, "TrayTips_zeigen"                            	, "nein"                 	)  	; zeigt TrayTips

		If InStr(Addendum.Labor.AbrufTimer, "Error")
			Addendum.Labor.AbrufTimer	:= false
		If InStr(Addendum.Labor.AutoAbruf, "Error")
			Addendum.Labor.AutoAbruf	:= false

}

admGesundheitsvorsorge() {                                                                    	;-- Abstände zwischen Untersuchungen minimales Alter

	Addendum.GVUAbstand                 	:= IniReadExt("Abrechnungshelfer"	, "minGVU_Abstand")
	Addendum.GVUminAlter                	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlter")
	Addendum.GVUminAlterEinmalig   	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterEinmalig")
	Addendum.GVUminAlterMehrmalig	:= IniReadExt("Abrechnungshelfer"	, "minGVUAlterMehrmalig")

}

admPIDHandles() {                                                                                   	;-- Prozess-ID's

		Addendum.AlbisPID                        	:= AlbisPID()
		Addendum.scriptPID                       	:= DllCall("GetCurrentProcessId")          	; die ScriptPID wird für das SkriptReload und die Interskript-Kommunikation benötigt

}

admInfoWindowSettings() {                                                                      	;-- AddendumGui/Infofenster

	aufteilung =
	(LTrim
		Daten1=2, 1, 151, 132
		Typ1=0
		Daten2=152, 1, 318, 132
		Typ2=0
		Diagnosen=320, -1, 616, 255
		Typ3=0
		Medikamente=617, 0, 946, 256
		Typ4=0
		Cave=2, 132, 318, 256
		Typ5=0
		Termine=1368, 1, 1, 1
		Typ6=0
		Patientengruppen=946, -1, 1085, 111
		Typ7=0
		Anamnese=946, 111, 1367, 257
		TypAnamnese=1
		Allergie=9, 168, 184, 255
		TypAllergien=1
		Operation=628, 365, 727, 410
		TypOperation=1
		Familie=1086, 0, 1366, 111
		AbrechAssist=947, 80, 1228, 163
		Info=946, 111, 1368, 255
	)

	tmp1 := IniReadExt(compname, "InfoFenster_Position", "y2 w400 h340")
	RegExMatch(tmp1, "y\s*(\d+)\s*w\s*(\d+)\s*h\s*(\d+)", match)

	Addendum.InfoWindow := {	"Y"                     	: match1
											, 	"W"                    	: match2
											,	"RefreshMinTime"	: 1500
											,	"RefreshTime"    	: 5000
											,	"ReIndex"            	: false
											,	"rowprcs"           	: 0
											,	"Init"                    	: false
											,	"ConfirmImport"   	: IniReadExt(compname, "Infofenster_Import_Einzelbestaetigung"  	, "Ja")
											,	"firstTab"           	: IniReadExt(compname, "Infofenster_aktuelles_Tab"                     	, "Patient")
											,	"TProtDate"         	: IniReadExt(compname, "Infofenster_Tagesprotokoll_Datum"       	, "Heute")
											,	"JournalSort"      	: IniReadExt(compname, "Infofenster_JournalSortierung"               	, "2 1")
											,	"LVScanPool"        	: {"W"	: (match2 - 10)
															            	,	"R"	: IniReadExt(compname, "InfoFenster_BefundAnzahl", 7)}}

	this := Addendum.InfoWindow.firstTab
	If InStr(this, "ERROR") || (StrLen(this) = 0)
		Addendum.InfoWindow.firstTab := "Patient"
	If RegExMatch(this, "(\w+)\s+(\d+)", tmp) {
		Addendum.InfoWindow.firstTab     	:= tmp1
		Addendum.InfoWindow.firstTabPos	:= tmp2
	}
	else
		Addendum.InfoWindow.firstTabPos	:= 0

	this := Addendum.InfoWindow.TProtDate
	If InStr(this, "ERROR") || (StrLen(this) = 0)
		Addendum.InfoWindow.TProtDate := "Heute"

}

admLaborDaten() {                                                                                	;-- Laborabruf, Verarbeitung Laborwerte

		Addendum.Laborabruf.Status	:= false
		Addendum.Laborabruf.Daten:= false
		Addendum.Laborabruf.Voll  	:= false

		Addendum.Labor.LDTDirectory     	:= IniReadExt("LaborAbruf"	, "LDTDirectory"                        , "C:\Labor"	)
		Addendum.Labor.LaborName       	:= IniReadExt("LaborAbruf"	, "LaborName"                             	    , ""    	)  	; falls sie mehrere Labore haben, tragen Sie das aktuelle hier ein
		Addendum.Labor.Laborkuerzel     	:= IniReadExt("LaborAbruf"	, "Aktenkuerzel"                            	    , "labor")  	; Karteikartenkürzel um Informationen ablegen zu können
		Addendum.Labor.Alarmgrenze     	:= IniReadExt("LaborAbruf"	, "Alarmgrenze"                            	    , "30%"	)  	; Alarmierungsgrenze in Prozent oberhalb der Normgrenzen
		Addendum.Labor.AbrufZeiten      	:= IniReadExt("LaborAbruf"	, "LaborAbruf_Zeiten"                    	    , ""    	)	; z.B. "06:00, 15:00, 19:00, 21:00"
		Addendum.Labor.Kennwort       	:= IniReadExt("LaborAbruf"	, "LaborKennwort"                        	    , ""    	)	; für order&entry per CGM-Channel
		If InStr(Addendum.Labor.Kennwort, "Error")
			Addendum.Labor.Kennwort := ""

		; ACHTUNG:
		; hier werden die Einstellungen für eine Alarmierung über Ihren eigenen Telegram-Bot eingelesen! Denken Sie immer daran die Telegram Token und ChatID's vor dem
		; Zugriff Fremder oder auch dem eigenen Personal zu schützen
		Addendum.Labor.TGramOpt        	:= IniReadExt("LaborAbruf", "TGramOpt"           	    , "nein")	        	; die Nummer ihres Telegram Bots und + oder - Tel z.b. "1+Tel"
		If !((Addendum.Labor.TGramOpt = false) || RegExMatch(Addendum.Labor.TGramOpt, "^\d+\+|\-Tel")) 	; die Werte müssen einen exakten Syntax haben, sonst wird die Telegram Option gelöscht!
			IniWrite, % "nein", % Addendum.AddendumIni, % "LaborAbruf", % "TGramOpt"

}

admLanProperties() {                                                                              	;-- Netzwerkeinstellungen / Client PC Namen / IP-Adressen

	; prüft auf den korrekt gesetzten Computernamen
		CName		:= IniReadExt(compname, "ComputerName")
		If (CName <> A_ComputerName)
			IniWrite, % A_ComputerName, % Addendum.AddendumIni, % compname, % "ComputerName"

	; hält die in der ini gespeicherte IP zu jedem Client aktuell
		CompIP 	:= A_IPAddress1
		storedIP    	:= IniReadExt(compname, "IP")
		If (CompIP <> storedIP)
			IniWrite, % CompIP, % Addendum.AddendumIni, % compname, % "IP"

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
			IniWrite, % Addendum.Monitor[A_Index], % Addendum.AddendumIni, % compname, % "Monitor" A_Index

	}

}

admPDFSettings()  {                                                                                	;-- Arbeitsverzeichnis, Einstellungen PDF Dokumentverarbeitung

		Addendum.xpdfPath				       	:= IniReadExt("ScanPool"     	, "xpdfPath"                                 	, AddendumDir "\include\xpdf")
		Addendum.PDFReaderFullPath      	:= IniReadExt("ScanPool"     	, "PDFReaderFullPath"                 	, AddendumDir "\include\FoxitReader\FoxitReaderPortable\FoxitReaderPortable.exe")
		Addendum.PDFReaderName        	:= IniReadExt("ScanPool"     	, "PDFReaderName"     	            	, "FoxitReader")
		Addendum.PDFReaderWinClass    	:= IniReadExt("ScanPool"     	, "PDFReaderWinClass"               	, "classFoxitReader")
		Addendum.SignatureCount        	:= IniReadExt("ScanPool"     	, "SignatureCount")
		Addendum.SignaturePages        	:= IniReadExt("ScanPool"     	, "SignaturePages")
		Addendum.SignatureWidth          	:= IniReadExt("ScanPool"     	, "Signature_Breite"                    	, 50)
		Addendum.SignatureHeight          	:= IniReadExt("ScanPool"     	, "Signature_Hoehe"                    	, 25)
		Addendum.Ort                             	:= IniReadExt("ScanPool"     	, "Ort")
		Addendum.Grund                        	:= IniReadExt("ScanPool"     	, "Grund")
		Addendum.SignierenAls             	:= IniReadExt("ScanPool"     	, "SignierenAls")
		Addendum.DokumentSperren     	:= IniReadExt("ScanPool"     	, "DokumentNachDerSignierungSperren", "ja")
		Addendum.Darstellungstyp         	:= IniReadExt("ScanPool"     	, "DarstellungsTyp")
		Addendum.PasswortOn               	:= IniReadExt("ScanPool"     	, "PasswortOn", "nein")
		Addendum.PatAkteSofortOeffnen  	:= IniReadExt("ScanPool"     	, "Patientenakte_sofort_oeffnen"   	, "ja")
		Addendum.OCRclient                	:= IniReadExt("ScanPool"     	, "tessOCRClient")


		SPData.xpdfPath				           	:= Addendum.xpdfPath
		SPData.PDFReaderFullPath           	:= Addendum.PDFReaderFullPath
		SPData.PDFReaderName              	:= Addendum.PDFReaderName
		SPData.PDFReaderWinClass         	:= Addendum.PDFReaderWinClass
		SPData.BefundOrdner                  	:= Addendum.BefundOrdner

		If (StrLen(Trim(Addendum.SignaturePages)) = 0)
			Addendum.SignaturePages := 0

		If InStr(Addendum.OCRclient, "Error")
			Addendum.OCRclient := ""

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

		Addendum.Help             	:= " | Strg+Pfeil runter - Akte schliessen |" 	; Hotkey Tips für die Statusbar von Albis
										    	. 	 " Strg+Alt+F5 - Datum(HEUTE) |"
										    	. 	 " Strg+F7/Alt+F7 - GVU |"
										    	. 	 " Alt+m - Menusuche |"
		Addendum.useraway     	:= false                                                    	; true wenn Nutzer einen bestimmten Zeitraum keine Eingaben gemacht hat
		Addendum.FirstDBAcess	:= false                                                    	; flag nur notwendig für den erstmaligen Aufruf von Addendum.ahk

}

admStandard() {                                                                                    	;-- Einstellungen für Addendum - Fenster und Dialoge

		Addendum.StandardFont                	:= IniReadExt("Addendum"	, "StandardFont"        	)
		Addendum.StandardBoldFont          	:= IniReadExt("Addendum"	, "StandardBoldFont"    	)
		Addendum.StandardFontSize          	:= IniReadExt("Addendum"	, "StandardFontSize"    	)
		Addendum.DefaultFntColor            	:= IniReadExt("Addendum"	, "DefaultFntColor"     	)
		Addendum.DefaultBGColor            	:= IniReadExt("Addendum"	, "DefaultBGColor"     	)
		Addendum.DefaultBGColor1          	:= IniReadExt("Addendum"	, "DefaultBGColor1"   	)
		Addendum.DefaultBGColor2          	:= IniReadExt("Addendum"	, "DefaultBGColor2"   	)
		Addendum.DefaultBGColor3          	:= IniReadExt("Addendum"	, "DefaultBGColor3"   	)
		Addendum.hashtagNachricht         	:= IniReadExt("Addendum"	, "HASHtagNachricht", "")
		Addendum.dpiF                               	:= screenDims().DPI / 96                                                                   	; DPI-Faktor

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

			threadFilePath := Addendum.AddendumDir "\Module\Addendum\threads\AddendumToolbar.ahk"
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
				failpath := StrReplace(threadFilePath, Addendum.AddendumDir)
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

			threadFilePath := Addendum.AddendumDir "\Include\Addendum_OCR.ahk"
			If FileExist(threadFilePath) {
				Addendum.Threads.OCR := FileOpen(threadFilePath, "r").Read()
			}
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



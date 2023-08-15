; _____________________________________________________________________________________________________________________________________________________________________________________________
; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;                                                                              	-- ADDENDUM QUICKEXPORTER --
; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; ‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾
/*
;         Entwickelt für einen deutlich weniger aufwendigen Karteikartenexport. Übersichtlicheres Tagesprotokoll.
;
;        ⯈ 	stellt Patientendaten in einem Durchlauf zusammen. Exportiert die Daten in ein Festplattenverzeichnis. (QUICK&DIRTY&OUT!)
;			  		▸ Tagesprotokoll (anpassbare Filter: ganze Karteikartenkürzel oder Teile des Inhaltes lassen sich entfernen)
;			   		▸ Labordaten
;			   		▸ eigene Befunddaten (Briefe  .doc, .docx; Bilder .jpg, .png, .bmp)
;			   		▸ externe Befunde
;				⯈ 	führt ein Erstellungsprotokoll des Mediums und Versandwegs und des Ausgangsdatums
;				⯈ 	Ärzteadressbruch
;				⯈ 	automatisiertes Anschreiben an den Empfänger (Pdf Format)
;				⯈		exportierte Verzeichnisse lassen sich per 7zip aus der Anwendung heraus packen und verschlüsseln
;
;			- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
;
;			  	Es wird nicht empfohlen Dateien/Verzeichnisse innerhalb des Stammverzeichnis manuell zu verschieben. Das Skript arbeitet mit den unterschiedlichen
; 				Verzeichnissen und den enthaltenen Daten um eine logische Sortierung nach Empfänger zu erstellen. In jedem Verzeichnis eines Empfängers werden
;			  	zunächst die noch nicht versandten Karteikarten gelegt. Vergibt man in der grafischen Oberfläche einer Akte eine Versanddatum wird diese innerhalb
;		  		des Empfängerverzeichnisses in ein Subverzeichnis mit Namen 'versendet' verschoben. Insgesamt soll dies den Massenexport von Daten übersichtlich
;		  	 	und damit verwechslungsarm halten.
;
;			- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
;
;			 	▸ Das Skript ist für meine individuellen Bedürfnisse hergestellt.
;				▸ Ein letztes Skript der Addendum für AoW Tools ist für den schnellen Export von Karteikarten als "Massenware" möglichst ohne weiteres Eingreifen gedacht.
;				▸ Zunächst werden die Daten in HTML und später als PDF ausgegeben.
;				▸ Die Umwandlung der HTML Dateien läuft über einen den Microsoft Print to PDF Druckertreiber und benötigt deshalb die Automatisierung des MS Edge Browser.
;				▸ Im HTML Code sind keine Javascript Codes enthalten. Die Ausgabe (z.B. um ein eigenes Logo integrieren) läßt durch Anpassung der HTML Vorlage im
;				  Skriptverzeichnis \resourcen vornehmen.
;				▸ Labordaten werden nicht direkt der Datenbank entnommen, sondern durch Automatisierung der Albis Oberfläche und ebenso durch PDF Druck gewonnen.
;				▸ Die kompletten Befunddaten inkl. der Texte (Konvertierung von DBASE Datenbanken) liegen als .json Dateien getrennt nach Patienten vor.
;
;
;                               	──────────────────────────────────────────────────────────────────────────────────────────────────────────
;
;		Abhängigkeiten:	⯈ 	siehe includes + Microsoft Edge Browser (Html zu Pdf Konvertierung) + SumatraPDF (Pdf Druck)
;
;
*/
;
; _____________________________________________________________________________________________________________________________________________________________________________________________
; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;		            																			aus Addendum für Albis on Windows - by Ixiko - started in September 2017
;		            																script birth: 24.04.2023 - last change 01.08.2023 - this file runs under Lexiko's GNU Licence
;
; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; ‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾


;{ ⚞──────────────────     SKRIPTEINSTELLUNGEN   	──────────────────⚟

	#NoEnv
	#Persistent
	#MaxThreads                  	, 50
	#MaxThreadsBuffer           	, On
	#MaxThreadsPerHotkey 	       	, 18
	#MaxHotkeysPerInterval		  	, 99000000
	#HotkeyInterval              	, 99000000
	#KeyHistory      					  	, Off
	#MaxMem                     	,

	SetBatchLines                	, -1
	;~ ListLines                         	, Off
	AutoTrim                     	, On
	CoordMode, Mouse             	, Screen
	CoordMode, Pixel             	, Screen
	CoordMode, ToolTip           	, Screen
	CoordMode, Caret            	, Screen
	CoordMode, Menu             	, Screen
	FileEncoding                 	, UTF-8

	DetectHiddenWindows          	, Off
	SetTitleMatchMode           	, 2
	SetTitleMatchMode           	, Fast

	SetWinDelay                  	, -1
	SetControlDelay              	, -1

	cJSON.EscapeUnicode := "UTF-8"

	OnExit, QEGuiClose

	Menu, Tray, Icon     	, % "hIcon:" Create_QuickExporter_ico()

;}

;{ ⚞──────────────────        VARIABLEN          	──────────────────⚟
	start := A_TickCount
	global AddendumDir
	global q                  	:= Chr(0x22)                           	; ist dieses Zeichen -> "
	global afxMDI, afxView, oPat, cPat, PatDb
	global Addendum           	:= Object()
	global Mdi                	:= Object()
	global compname           	:= StrReplace(A_ComputerName, "-")
	global scriptTitle         	:= "Addendum - Albis Quickexporter"
	global dbg
	global noSQLWrite
	global stdExportPath, stdExportDrive
	global settings := Object()
	global dpiFactor := 96/A_ScreenDPI                              	; z.b. dpiFactor := 96/144  ; width := guiW * dpiFactor

	global TExports    	:= []					     	               	        	; Spaltennamen der db.Tabelle Exports
	global DBAccess    	:= "W"                                      	; Open db ReadWrite - this is set to default
	global vCtrls		:= Array()					                        			; Dokumentationsinterface, Steuerelemente Einstellungen, Regeln
	global vCtrlRef    	:= Object() 				                    			; Referenzen zu den vCtrls Daten
	global dCtrls				                   	                     			; Arztadressen, Steuerelemente Einstellungen, Regeln
	global sql   								            			; SQLite Snippets
	global ctree								            			; Regeln
	global hQE						         		        			; Gui hwnd
	global QEDB                                                        	; SQLite Datenbank Objekt (class_SQLite --> QExport.db)
	global dab						         		        			; Addressbuchhandler
	global ExportsCols  			         	            			;
	global ExportsTbl                         	             			;
	global cPat
	global PatDB
	global hSBar
	global Recipients   	:= Object()

  ; UIA für HTML nach PDF Umwandlung mittels Edge Browser
	global DeepSearchFromPoint := False ; When set to True, UIAViewer iterates through the whole UIA tree to find the smallest element from mouse point. This might be very slow with large trees.
	global UIA := UIA_Interface(), IsCapturing := False, Stored := {}, Acc, EnableAccTree := False
	global cUIA

 	RegExStrings := ["i)\{?<Code>([A-Z]\d+\.*\d*)(?<BodySide>[RLB]{0,1})(?<Status>[VGA]{0,1})\s*\}"]		; ICD Code aufteilen nach

; ------------------------------------------------------------------------------------------------------------------------------
; Verzeichnisse anlegen
; ----------------------------------------------------------------------------------------------------------------------------;{
	settings.paths := Object()
	settings.paths.SQLProtPath  	:= A_ScriptDir "\logs"
	settings.paths.resources     	:= A_ScriptDir "\resources"
	settings.paths.sqlitefiles		:= A_ScriptDir "\sqlite_db"
	settings.paths.QEExporterINI 	:= A_ScriptDir "\QExporter.ini"
	settings.paths.DBLocation  		:= A_ScriptDir "\QExporter.db"
	settings.paths.DBBupLocation	:= settings.paths.sqlitefiles "\QExport-Backup_"

	; Verzeichnis prüfen / anlegen
	If !FilePathCreate(settings.paths.SQLProtPath)
		throw "Konnte Verzeichnis zum Schreiben von Daten nicht anlegen.`n" settings.paths.SQLProtPath
	If !FilePathCreate(settings.paths.resources)
		throw "Konnte Verzeichnis zum Laden und Schreiben von Daten nicht finden.`n" settings.paths.resources
	If !FilePathCreate(settings.paths.sqlitefiles)
		throw "Konnte Verzeichnis zum Laden und Schreiben von Daten nicht finden.`n" settings.paths.sqlitefiles

	; Dateien prüfen
	If !FileExist(settings.paths.resources "\Karteikarte.html") {
		MsgBox, 0x1004, % scriptTitle, % "Es fehlt eine HTML/CSS-Datei im Skript '\resources' Verzeichnis.`n"
						                			.	 "Ein Export von Patientendaten ist ohne diese Datei nicht möglich`n"
						                			.	 "Das Skript wird nach Drücken von 'Ok' beendet."
		ExitApp
	}


	;}


; ------------------------------------------------------------------------------------------------------------------------------
; Skript-Einstellungen laden
; ----------------------------------------------------------------------------------------------------------------------------;{
	workini := IniReadExt(settings.paths.QEExporterINI)
	dbg                       	:= IniReadExt("Script"        	, "Debug"                                      	, 0)
	GuiPos                    	:= IniReadExt("Script"        	, "GuiPosition_" compname                      	, "x=Center y=Center")
	If !(stdExportPath        	:= IniReadExt("Export"        	, "standard_exportpfad_" compname))
		stdExportPath           	:= IniReadExt("Export"        	, "standard_exportpfad")
	employeeshorts            	:= IniReadExt("Export"        	, "Namenskuerzel") " "
	settings.noLabData        	:= IniReadExt("Export"        	, "keine_Labordaten_exportieren"              	, 0)
	settings.LVExpFilter      	:= IniReadExt("Script"        	, "fehlerfreie_Exporte_ausblenden"            	, 0)
	settings.LVExpFilter2     	:= IniReadExt("Script"        	, "versandte_Exporte_ausblenden"              	, 0)
	settings.reNewData        	:= IniReadExt("Script"        	, "Exportdaten_neu_erstellen"                  	, 0)
	settings.PWDVisibility    	:= IniReadExt("Script"        	, "UserPWD_Sichtbarkeit"	                    	, 0)
	settings.DontShowSmtraMsg 	:= IniReadExt("Script"        	, "Sumatra_Hinweis_nicht_zeigen"               	   )
	settings.SumatraCMD         := IniReadExt("Script"        	, "Sumatrapdf_Pfad")
	settings.QESdate           	:= IniReadExt("SQLite"        	, "Tagesprotokoll_ab"                          	, "")
	settings.sendLetter        	:= IniReadExt("Versandbrief" 		, "Brief_erstellen"                           	, "")
	settings.printLetter       	:= IniReadExt("Versandbrief" 		, "Brief_drucken"                             	, "")
	settings.pdfprinter       	:= IniReadExt("Versandbrief" 		, "Brief_Drucker"                             	    )
	settings.ZipWithLetter    	:= IniReadExt("Export"       		, "Briefe_in_7z_inkludieren"                  	, 0)
	settings.paths.networkpath 	:= IniReadExt("Export"       		, "networkpath_" compname                      	, "")
	settings.noSQLWrite       	:= IniReadExt("SQLite"        	, "SQL_write_protection"	                    	, 0)

	RegExMatch(GuiPos, "i)x\s*=\s*(?<X>\-*\d+|Center)", gui)	; x Position
	RegExMatch(GuiPos, "i)y\s*=\s*(?<Y>\-*\d+|Center)", gui)	; y Position

	stdExportPath             	:= stdExportPath              	= "Error" ? "" : stdExportPath
	settings.dbg := dbg       	:= dbg			                  	= "Error" ? 0  : dbg
	settings.noSQLWrite       	:= settings.noSQLWrite        	= "Error" ? 0  : settings.noSQLWrite
	settings.LVExpFilter      	:= settings.LVExpFilter       	= "Error" ? 0  : settings.LVExpFilter
	settings.LVExpFilter2     	:= settings.LVExpFilter2      	= "Error" ? 0  : settings.LVExpFilter2
	settings.LVExpFilter      	:= settings.LVExpFilter2      	= true && settings.LVExpFilter = true ? false : settings.LVExpFilter
	settings.reNewData        	:= settings.reNewData         	= "Error" ? 0  : settings.reNewData
	settings.PWDVisibility    	:= settings.PWDVisibility     	= "Error" ? 0  : settings.PWDVisibility
	settings.noLabData        	:= settings.noLabData         	= "Error" ? 0  : settings.noLabData
	settings.QESdate          	:= settings.QESdate           	= "Error" ? "" : settings.QESdate
	settings.SumatraCMD       	:= settings.SumatraCMD        	= "Error" ? "" : settings.SumatraCMD
	settings.pdfprinter       	:= settings.pdfprinter        	= "Error" ? "" : settings.pdfprinter
	settings.DontShowSmtraMsg 	:= settings.DontShowSmtraMsg   	= "Error" ? "" : settings.DontShowSmtraMsg
	settings.paths.networkpath	:= settings.paths.networkpath  	= "Error" ? "" : settings.paths.networkpath
	settings.CheckAlbis       	:= false
	settings.folderToPath     	:= {"shipped":"bereits versendet", "misc":"andere Gründe", "mortal":"verstorben"}
	settings.physicans	      	:= {}

	; ----------------------------------------------------------------------------------------------------------
  ; Brennen auf CD/DVD Laufwerk
	; CD Laufwerke werden erkannt und hinzugefügt
	; rdp Sitzung: Brennen per remote Skript auf einem Netzwerkclient
	settings.burn               := Object()
	settings.burn.cdbxpPath   	:= IniReadExt("Export"       		, "CDBurnerXP_Verzeichnis_" compname          	, "")
	settings.burn.usedrive    	:= IniReadExt("Export"       		, "DVD_Laufwerk_" compname                     	, "")

  ; bisherige Einstellungen Empfänger -> Versandmedium und Versandweg laden
	IniRead, tmp, % settings.paths.QEExporterINI, % "Ärzte"
	For each, iniline in StrSplit(tmp, "`n", "`r") {
		bsnr  	:= StrSplit(iniline, "=").1
		Medium 	:= StrSplit(StrSplit(iniline, "=").2, "|").1
		Route 	:= StrSplit(StrSplit(iniline, "=").2, "|").2
		settings.physicans[bsnr] := {"MediumO":Medium, "Medium":Medium, "RouteO":Route, "Route":Route}
	}


	; Absenderdaten (bisher nur für einen Arzt)
	settings.sender := Object()
	IniRead, tmp, % settings.paths.QEExporterINI, % "Versandbrief"
	For each, iniline in StrSplit(tmp, "`n", "`r")
		settings.sender[StrReplace(StrSplit(iniline, "=").1, "Absender_")] := StrSplit(iniline, "=").2
	tmp	:= ""


	; Druckereinstellungen
	settings.printer := Array()
	IniRead, tmp, % settings.paths.QEExporterINI, % "Drucker"
	For each, iniline in StrSplit(tmp, "`n", "`r") {
		RegExMatch(iniline, "i)^\s*(?<Client>.+?)_Drucker(?<Index>\d+)_(?<Key>.+?)=(?<Value>.+)", Prn)
		If (PrnClient != compname)
			continue
		If (!settings.printer.Count() || !IsObject(settings.printer[PrnIndex]))
			settings.printer[PrnIndex] := Object()
		settings.printer[PrnIndex][Trim(PrnKey)] := Trim(PrnValue)
	}
	tmp	:= ""
	; zum Drucken von PDF Dateien macht sich SumatraPDF per cmdline Befehl am besten
	If !settings.SumatraCMD
		settings.SumatraCMD := GetAppImagePath("SumatraPDF")


	; AddendumDir
	RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
	SplitPath, A_ScriptDir,,,,, AddendumDrive


	; ----------------------------------------------------------------------------------------------------
	; Exportverzeichnis (versucht den Pfad über die Verwendung eines anderen Laufwerkes zu finden)
	; -------------------------------------------------------------------------------------------------;{
	if !stdExportPath {
		settings.noExportPathMsg := "Willkommen beim Quickexporter, für das Setup wird nur die Angabe eines Arbeitsverzeichnisses benötigt."
	}
	else {

		if RegExMatch(stdExportPath, "^(?<Drive>(\w:|\\\\\w+))", stdExport) {

			; Laufwerk ist nicht bekannt
			if !FileExistZ(stdExportDrive "\") {

				settings.noExportPathMsg :=  "Das Standardexportverzeichnis [" stdExportPath "] in den Einstellungen verweist auf ein nicht vorhandenes Laufwerk [" stdExportDrive "]."
				settings.ExportPath := RegExReplace(stdExportPath, "^\w:", AddendumDrive)
				stdExportpfad := stdExportDrive := ""

				if FileExistZ(settings.ExportPath "\") {
					stdExportpfad 	:= settings.ExportPath
					stdExportDrive	:= AddendumDrive
					settings.noExportPathMsg  := AddendumDrive := ""
				}
				else
					settings.noExportPathMsg .= " Auch auf dem Skriptlaufwerk [" AddendumDrive "] findet sich das Verzeichnis nicht."

			}

			; Verzeichnis wird nicht gefunden
			else if !FileExistZ(stdExportPath) {

				settings.noExportPathMsg := "Das Standardexportverzeichnis [" stdExportPath "] ist nicht vorhanden"
				if FileExistZ(stdExportPath := RegExReplace(stdExportPath, "^\w:", AddendumDrive))
					stdExportDrive := AddendumDrive, settings.noExportPathMsg := ""
				else   ; leeren wenn kein Verzeichnis vorhanden ist
					stdExportpfad := stdExportDrive := AddendumDrive := ""

			}

		}
		; keine gültige Verzeichnisangabe
		else {

			settings.noExportPathMsg :=  "Das ausgelesene Standardexportverzeichnis enthält ist keine gültige Pfadangabe. [" stdExportpfad "]"
			stdExportpfad := stdExportDrive := AddendumDrive := ""

		}

	}

	if settings.noExportPathMsg {

		SelectStdExportPath:
		settings.ssep_msg := "Wählen Sie ein Festplattenverzeichnis, in welchem Quickexporter die exportierten Patientendaten verwalten kann.`n`n"
		settings.ssep_msg .= "Es ist ratsam ein neues (leeres) Verzeichnis anzulegen oder sie geben das bereits bestehende Arbeitsverzeichnis an."

		MsgBox, 0x1000, % scriptname, % (settings.noExportPathMsg ? settings.noExportPathMsg "`n`n" : "") settings.ssep_msg
		stdExportPath := SelectFolder(0, "c:")
		settings.noExportPathMsg := ""

		if stdExportPath && !FileExistZ(stdExportPath)
			settings.noExportPathMsg := "Das übergebene Verzeichnis [" settings.noExportPathMsg "] ist nicht vorhanden."
		else if !stdExportPath
			settings.noExportPathMsg := "Sie haben kein Verzeichnis gewählt."

		If settings.noExportPathMsg {

			MsgBox, 0x1004, % scriptname, % settings.noExportPathMsg "`nDrücken Sie 'Ja' für eine neue Abfrage oder drücken Sie 'Nein' um das Skript zu beenden?"
			IfMsgBox, Yes
				goto SelectStdExportPath
			ExitApp

		}

		SplitPath, stdExportPath,,,,, stdExportDrive
		settings.stdIniPath := IniReadExt("Export", "Standard_Exportpfad")
		IniWrite, % stdExportPath, % settings.paths.QEExporterINI, % "Export", % "Standard_Exportpfad_" compname
		If (!settings.stdiniPath || InStr(settings.stdiniPath, "Error"))
			IniWrite, % stdExportPath, % settings.paths.QEExporterINI, % "Export", % "Standard_Exportpfad"

		settings.noExportPathMsg := settings.ssep_msg := settings.stdiniPath := ""
	}

	;}


	;}


; ------------------------------------------------------------------------------------------------------------------------------
; TrayMenu erstellen
; ----------------------------------------------------------------------------------------------------------------------------;{
	func_toggle1          	:= Func("ToogleMenu").Bind("dbg")
	func_toggle2          	:= Func("ToogleMenu").Bind("noSQLWrite")
	func_showObjects      	:= Func("ShowObjects")
	Menu, TMenu	, Add	, % "Debug"                  	, % func_toggle1
	Menu, TMenu	, Add	, % "SQL write protection"		, % func_toggle2
	Menu, Tray	, Add	, Einstellungen								, :TMenu
	Menu, Tray	, Add	, Skriptobjekte zeigen				, % func_showObjects

	Menu, TMenu, % (dbg ? "Check" : "Uncheck")           	, % "Debug"
	Menu, TMenu, % (settings.noSQLWrite ? "Check" : "Uncheck")	, % "SQL write protection"

	;}


; ------------------------------------------------------------------------------------------------------------------------------
; Addendum Daten
; ----------------------------------------------------------------------------------------------------------------------------;{
	workini := IniReadExt(AddendumDir "\Addendum.ini")
	Addendum.Default := {}
	Addendum.Default.Font     	:= IniReadExt("Addendum"	, "StandardFont"      	)
	Addendum.Default.BoldFont  	:= IniReadExt("Addendum"	, "StandardBoldFont"  	)
	Addendum.Default.FontSize 	:= IniReadExt("Addendum"	, "StandardFontSize"  	)
	Addendum.Default.FntColor  	:= IniReadExt("Addendum"	, "DefaultFntColor"   	)
	Addendum.Default.BGColor  	:= IniReadExt("Addendum"	, "DefaultBGColor"     	)
	Addendum.Default.BGColor1  	:= IniReadExt("Addendum"	, "DefaultBGColor1"   	)
	Addendum.Default.BGColor2  	:= IniReadExt("Addendum"	, "DefaultBGColor2"   	)
	Addendum.Default.BGColor3 	:= IniReadExt("Addendum"	, "DefaultBGColor3"   	)
	Addendum.qpdfPath          	:= IniReadExt("OCR"     	, "qpdfPath"           	)

	Addendum.Dir             	:= AddendumDir
	Addendum.AddendumDir   		:= AddendumDir
	Addendum.DBPath 	     		:= AddendumDir    	"\logs'n'data\_DB"
	Addendum.dataPath 		  	:= AddendumDir    	"\include\Daten"
	Addendum.AlbisDBPath 	  	:= GetAlbisPath() 	"\db"
	Addendum.AlbisPath      	:= AddendumDrive  	"\albiswin"
	Addendum.AlbisBriefe    	:= AddendumDrive  	"\albiswin\Briefe"
	Addendum.AutoLogin 		   	:= true

	Addendum.Albis     	  		:= Object()
	Addendum.Albis     	   		:= GetAlbisPaths()
	Addendum.AlbisExe  		  	:= Addendum.Albis.Exe

	Addendum.Praxis 	     		:= Object()
	Addendum.PersonRx	     		:= "(Herr\s*|Frau\s*|Hr\.\s*|Fr\.\s*|Prof.\s*|Priv.\s*|Doz\.\s*|Dr\.\s*|med\.\s*)"
	admPraxisDaten()

	if Addendum.qpdfPath && !FileExist(Addendum.qpdfPath)
		Addendum.qpdfPath := ""

;}


; ------------------------------------------------------------------------------------------------------------------------------
; vorgefertigte SQLite Snippets für den wiederholten Gebrauch
; ----------------------------------------------------------------------------------------------------------------------------;{
; INSERT INTO EXPORTS Values ('$1','$2','$3','$4','$5','$6','$7','$8','$9','$10','$11','$12');$13
; (.+),(.+),(.+),(.+),(.+),(.+),(.+),(.+),(.+),(.+),(.+),(.+)?([\n\r]+)
					;			vorhandene Tabelle anzeigen
	sql := {	"TableLookup"	: "select name from sqlite_master where type='table';"	 ; This will produce a list of all tables in the Database

					;			SQLite Exports
					,		"Exports"        	: {"CreateTable": "CREATE TABLE EXPORTS (PatID INTEGER PRIMARY KEY,"
																				    							. 	"SName TEXT NOT NULL,"
																				    							.	"PName TEXT NOT NULL,"
																				    							.	"Birth DATE NOT NULL,"
																				    							. 	"SDate DATE NOT NULL,"
																				    							.	"Recipient INTEGER,"
																				    							.	"Medium TEXT,"
																				    							.	"Route TEXT,"
																				    							.	"Shipment DATE,"
																				    							.	"More1 TEXT,"
																				    							.	"ExportPath TEXT,"
																				    							.	"More3 TEXT,"
																				    							.	"FOREIGN KEY (Recipient) REFERENCES Recipients(BSNR));"
													, "Count"        		: "SELECT COUNT(*) FROM EXPORTS;"
													, "GetRecords"		: {"WHEREPatID"	: "SELECT * FROM EXPORTS WHERE PatID = "}
													, "GetTable"			: "SELECT * FROM EXPORTS;"}

					; 			SQLite Empfänger
					,		"Recipients"		: {"CreateTable": "CREATE TABLE RECIPIENTS (ANR TEXT,"
																											.	"RTitel TEXT,"
																											.	"RSName TEXT,"
																											.	"RPName TEXT,"
																											.	"KIM TEXT,"
																											.	"LANR INTEGER,"
																											.	"BSNR INTEGER PRIMARY KEY,"
																											.	"Strasse TEXT,"
																											.	"Ort TEXT,"
																											.	"PLZ TEXT,"
																											.	"Tel1 TEXT,"
																											.	"Tel2 TEXT,"
																											.	"Fax1 TEXT,"
																											.	"Fax2 TEXT,"
																											.	"EMail Text,"
																											.	"More1 TEXT,"
																											.	"More2 TEXT);"

												, "GetTable"			: "SELECT * FROM RECIPIENTS;"
												, "CreateIndex"    	: "CREATE INDEX RECIPIENTS_RNR ON RECIPIENTS (RNR,LANR,BSNR);"
												, "GetTable"			: "SELECT * FROM RECIPIENTS;"}

					;			SQLite Moved
					,		"Moved"        	: {"CreateTable": "CREATE TABLE MOVED (PatID INTEGER PRIMARY KEY,"
																				    							. 	"SName TEXT NOT NULL,"
																				    							.	"PName TEXT NOT NULL,"
																				    							.	"Birth DATE NOT NULL,"
																				    							. 	"SDate DATE NOT NULL,"
																				    							.	"Recipient INTEGER,"
																				    							.	"Medium TEXT,"
																				    							.	"Route TEXT,"
																				    							.	"Shipment DATE,"
																				    							.	"More1 TEXT,"
																				    							.	"ExportPath TEXT,"
																				    							.	"Folder TEXT,"
																				    							.	"FOREIGN KEY (Recipient) REFERENCES Recipients(BSNR));"
													, "Count"        		: "SELECT COUNT(*) FROM MOVED;"
													, "GetRecords"		: {"WHEREPatID"	: "SELECT * FROM MOVED WHERE PatID = "}
													, "GetTable"			: "SELECT * FROM MOVED;"}

					, 		"GetTable"			: {"ExportsReci"    	: "SELECT e.PatID, e.SName, e.PName, e.Birth, e.SDate, e.Medium, e.Route, e.Shipment, e.More1, e.More2, e.More3"
																				.	", r.ANR, r.RTitel, r.RSName, r.RPName FROM Exports e JOIN Recipients r ON e.Recipient = r.RNR;"}
					,		"CreateIndex"	: {"Recipients"    	: "CREATE UNIQUE INDEX RECIPIENTS_RNR_IDX ON RECIPIENTS (RNR);"}
					,		"ALTERColumn"	:	"ALTER TABLE EXPORTS RENAME COLUMN More2 TO ExportPath;"
					, 		"dummy"        	: "dummy"}
;}


; ------------------------------------------------------------------------------------------------------------------------------
; Regelsystem	=>	Empfänger <-> Medium <-> Routen
; ----------------------------------------------------------------------------------------------------------------------------;{
	ctree := {"BSNR"   	: {"1"                             	: {"Medium"	: "|Einzel-CD|Papierausdruck|EMail verschlüsselt|ungeklärt"
																					, 	"Route"  	: "|Briefkasten|Post|persönliche Übergabe|Passwort per #|ungeklärt"}
											,  "2"        	                    	: {"Medium"	: "|Einzel-CD|Papierausdruck|ungeklärt"
																					, 	"Route"  	: "|Briefkasten|Post|persönliche Übergabe|ungeklärt"}
											,  "([3-9]|\d{1,})"           	: {"Medium"	: "|ePost|Einzel-CD|Sammel-CD|Papierausdruck|ungeklärt"
																					, 	"Route"  	: "|Telematik|Briefkasten|Post|persönliche Übergabe|ungeklärt"}}

						,	"Medium"	: {"ePost"		                	: {"Recipient"	: "(Hausarzt|Dr.*)"
																					, 	"Medium"	: "|ePost|Einzel-CD|Sammel-CD|Papierausdruck|ungeklärt"
																					, 	"Route"  	: "||Telematik|Briefkasten|Post|persönliche Übergabe|ungeklärt"}
											,	"Einzel-CD"               	: {"Recipient"	: ".*{1}"
																					, 	"Medium"	: "|ePost|Einzel-CD|Sammel-CD|Papierausdruck|ungeklärt"
																					, 	"Route"  	: "|Briefkasten|Post|persönliche Übergabe|ungeklärt"}
											,	"Sammel-CD"            	: {"Recipient"	: "(Hausarzt|Dr.*)"
																					, 	"Medium"	: "|ePost|Einzel-CD|Sammel-CD|Papierausdruck|ungeklärt"
																					, 	"Route"  	: "|Briefkasten|Post|persönliche Übergabe|ungeklärt"}
											,	"Papierausdruck"   		: {"Recipient"	: ".*{1}"
																					, 	"Medium"	: "|ePost|Einzel-CD|Sammel-CD|Papierausdruck|ungeklärt"
																					, 	"Route"  	: "|Briefkasten|Post|persönliche Übergabe|ungeklärt"}}

						,	"Route"		: {"Telematik"            		: {"Medium"	: "|ePost|Einzel-CD|Sammel-CD|Papierausdruck|ungeklärt", "Recipient" : "Hausarzt|Dr.*"}
											,	"Briefkasten"          		: {"Recipient" : {"(Hausarzt|Dr.*)"  	: "Einzel-CD|Sammel-CD|Papierausdruck"
																							        	,	"Patient"            	: "Einzel-CD|Papierausdruck"
																										,	"Angehöriger"     	: "Einzel-CD|Papierausdruck"}}
											,	"Post"			           		: {"Recipient"	: {"(Hausarzt|Dr.*)"	: "Einzel-CD|Sammel-CD|Papierausdruck"
																				        				,	"Patient"            	: "Einzel-CD|Papierausdruck"
																						        		,	"Angehöriger"     	: "Einzel-CD|Papierausdruck"}}
											,	"persönliche Übergabe"	: {"Recipient"	: {"(Hausarzt|Dr.*)"	: "Einzel-CD|Sammel-CD|Papierausdruck"
																					     	    		,	"Patient"            	: "Einzel-CD|Papierausdruck"
																						        		,	"Angehöriger"     	: "Einzel-CD|Papierausdruck"}}}}

;}


; ------------------------------------------------------------------------------------------------------------------------------
; Standard Dokumentationsausgabe -3 Jahre vom aktuellem Datum
; ----------------------------------------------------------------------------------------------------------------------------;{
	minusYYY := 	A_YYYY A_MM A_DD 000000
	minusYYY += 	0, days
	minusYYY += 	-3*365, days
	minusYYY := 	ConvertDBASEDate(minusYYY)
;}


; ------------------------------------------------------------------------------------------------------------------------------
; Gui Steuerelemente, Einstellungen, Regeln, Default Werte
; ----------------------------------------------------------------------------------------------------------------------------;{
	vCtrls.Push({	"vName"	: "QEPatID"
				    	,	"user"	  	: true
				    	,	"col"	    	: 1
							, "colwidth" 	: 80
				    	,	"ctype"	  	: "Edit"
				    	,	"opt"	    	: "w55"
				    	,	"txt"	    	: "PatID"
				    	,	"topt"	   	: "Center h17 +0x200"
				    	,	"rx"	     	: "(?<ID>\d+)"
				    	,	"lv"    		: "PatID"
				    	,	"db"	     	: {"Exports":"PatID"}
				    	,	"standard"  : ""
				    	,	"default" 	: ""
				    	,	"value"	  	: ""}) ; 1

	vCtrls.Push({	"vName" : "QEPatient"
				    	,	"user"   		: true
		  	    	,	"col"	    	: {2:"Name",3:"Vorname",4:"Geburt", 13:"Birth"}
             	, "colwidth" 	: 180
				    	,	"ctype"	  	: "Edit"
				    	,	"opt"    		: "w230"
				    	,	"copt"   		: "hwndQEhPatient"
				    	,	"txt"	    	: "Nachname, Vorname"
				    	,	"topt"	  	: "Center h17 +0x200"
				    	,	"rx"	    	: "(?<surname>[\pL\-\s]+)\s*\,\s*(?<prename>[^\d\*\.]+))[\*\s\,]*((?<birth>\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})))"
				    	,	"lv"     		:  ["Name", "Vorname", "Geburt"]
				    	,	"db"	    	: {"Exports":["SName","PName","Birth"]}
				    	,	"standard" 	: ""
				    	,	"default" 	: ""
				    	,	"value"	  	: ""}) ; 2

	vCtrls.Push({	"vName"	: "QESDate"
			    		,	"user"   		: true
			       	,	"col"	    	: 5
            	, "colwidth" 	: 180
			    		,	"lvcol"	  	: 14
			    		,	"opt"	     	: "w110"
			    		,	"ctype"  		: "Combobox"
			    		,	"copt"   		: "gQEBHandler hwndQEhSDate"
			    		,	"txt"	    	: "von Datum"
			    		,	"topt"  		: "Center h20 +0x200"
			    		,	"rx"      	: "\d{2}+\.\d{2}\.(\d{4}|\d{2})"
			    		,	"lv"  	  	: "SDatum"
			    		,	"db"    		: {"Exports":"SDate", "isDate":true}
			    		,	"standard"  : minusYYY
			    		,	"default" 	: "|" minusYYY
			    		,	"value"  		: ""}) ; 3

	vCtrls.Push({	"vName"	: "QERecipient"
			    		,	"user"   		: true
			       	,	"col"	    	: 6
            	, "colwidth" 	: 1
			    		,	"lvcol"	  	: 12
			    		,	"opt"     	: "w290"
			    		,	"ctype"	  	: "Combobox"
			    		,	"copt"	   	: "gQEBHandler"
			    		,	"txt"	    	: "Empfänger"
			    		,	"topt"  		: "Center h20 +0x200"
			    		,	"rx"	    	: "[\pL\s\-\,\.]+\s*(\[\d+\])*"
			    		,	"lv"	    	: "Empfänger"
			    		,	"lvr"	    	: ["Recipient>Recipients:ANR Recipients:RTitel Recipients:RSName, Recipients:RPName ", "bsnr>Exports:Recipient"]
			    		,	"db"	    	: {"Recipients":"bsnr", "Exports":"Recipient", "lv":"bsnr"}
			    		,	"standard"	: ""
			    		,	"default" 	: "|Patient|Angehöriger|Hausarzt"
			    		,	"value"	   	: ""}) ; 4

	vCtrls.Push({	"vName"	: "QEMedium"
					,	"user"		: true
					,	"vNR"		: 5
			    	,	"col"	    	: 7
        	, "colwidth"  	: 1
					,	"lvcol"		: 7
					, 	"opt"		: "w160"
					,	"ctype"		: "Combobox"
					,	"copt"		: "gQEBHandler"
					,	"txt"		: "Versandmedium"
					,	"topt"		: "Center h20 +0x200"
					, 	"rx"		: ".{1,}"
					,	"lv"		: "VMedium"
					,	"db"		: {"Exports":"Medium"}
					,	"standard" 	: ""
					, 	"default"	: "|🖄 ePost|💿 Einzel-CD|💽 Sammel-CD|📚 Papierausdruck|⚿ Datei verschlüsselt|❔ ungeklärt"
					, 	"rule    " 	: [{"vName"	: "QERecipient", "match":"i)^\s*(Patient|Angehöriger)"
										, 	 "dvalue"	: "|💿 Einzel-CD|📚 Papierausdruck|⚿ Datei verschlüsselt|❔ ungeklärt"}
										,  {"vName"	: "QERecipient", "match":"i)^\s*(.*{1,})"
										, 	"dvalue"	: "|🖄 ePost|💿 Einzel-CD|💽 Sammel-CD|📚 Papierausdruck|⚿ Datei verschlüsselt|❔ ungeklärt"}]
					,	"value"		: ""}) ; 5

	vCtrls.Push({	"vName"	: "QERoute"
					,	"user"		: true
					,	"vNR"		: 6
			    	,	"col"	    	: 8
        	, "colwidth"  	: 1
					,	"lvcol"		: 8
					, 	"opt"		: "w175"
					,	"ctype"		: "Combobox"
					,	"copt"		: "gQEBHandler"
					,	"topt"		: "Center h20 +0x200"
					, 	"rx"		: ".{1,}"
					,	"txt"		: "Versandweg"
					,	"lv"		: "VRoute"
					,	"db"		: {"Exports":"Route"}
					,	"standard" 	: ""
					, 	"default"	: "|📡 Telematik|✉ Postbrief|📪 Briefkasten|🤝 persönliche Übergabe|📧 EMail 🔐 Passwort per #|❔ ungeklärt"
					, 	"rule"     	: [{	"vName"	: "QEMedium", "match":"i)^\s*(..ePost)"
										,		"dvalue" 	: "📡 Telematik"}
										, { 	"vName"	: "QERoute", "match"	: "i)^\s*(.*{1,})"
										,		"dvalue" 	: "|✉ Postbrief|📪 Briefkasten|🤝 persönliche Übergabe|📧 EMail 🔐 Passwort per #|❔ ungeklärt"}]
					,	"value"		: ""}) ; 6

	vCtrls.Push({	"vName"	: "QEShipment"
					,	"user"	  	: true
					,	"vNR"	    	: 7
					,	"col"	    	: 9
        	, "colwidth" 	: 1
					,	"lvcol"	  	: 15
					, "opt"	    	: "w120"
					,	"ctype"	  	: "Combobox"
					,	"copt"	  	: "gQEBHandler hwndQEhShipment"
					,	"txt"	    	: "Versanddatum"
					,	"topt"	  	: "Center h20 +0x200"
					, "rx"	    	: "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})"
					,	"lv"	     	: "VDatum"
					,	"db"     		: {"Exports":"Shipment", "isDate":true}
					,	"standard" 	: ""
					,	"default" 	: ""
					,	"value"	  	: ""}) ; 7

	vCtrls.Push({	"vName"	: "QEMore1"
					,	"user"		: true
					,	"vNR"		: 8
			    	,	"col"	    	: 10
        	, "colwidth"  	: 1
					,	"lvcol"		: 10
					,	"ctype"		: "Edit"
					, 	"opt"			: "w600"
					,	"topt"	    	: "Center h17 +0x200"
					, 	"rx"			: ".{1,}"
					,	"txt"			: "Hinweise, Anmerkungen, sonstiges"
					,	"lv"			: "Anmerkungen"
					,	"lvtxt"			: "More1"
					,	"db"			: {"Exports":"More1"}
					,	"standard" : ""
					, 	"default" 	: ""
					,	"value"		: ""}) ; 8

	vCtrls.Push({	"vName"	: "QEExportPath"
					,	"user"		: false
			    	,	"col"	    	: 11
        	, "colwidth"  	: 1
					,	"lvcol"		: 11
					, 	"opt" 		: "w0"
					, 	"rx"			: ".{1,}"
					,	"lv"			: "Export"
					,	"db"			: {"Exports":"ExportPath"}
					,	"standard" : ""
					, 	"list" 			: ""
					,	"value"		: ""}) ; 9


				; im Moment nicht benötigt 🦺
	vCtrls.Push({	"vName"	: "QEMore3"
					,	"user"		: false
			    	,	"col"	    	: 0
        	, "colwidth"  	: 1
					,	"lvcol"		: 0
					, 	"opt"			: "w0"
					, 	"rx"			: ".*"
					,	"db"			: {"Exports":"More3"}
					,	"standard" : ""
					, 	"default" 	: ""
					,	"value"		: ""}) ; 10

  ; welcher Spaltemnnummer in der Datenbank
	For colNr, ctrl in vCtrls
		vCtrlRef[ctrl.vName] := colNr

;}


; ------------------------------------------------------------------------------------------------------------------------------
; Adressbuchklasse und Quickexporterklasse
; ----------------------------------------------------------------------------------------------------------------------------;{
	DAB := new DocAddressBook()
	qexporter := new Quickexporter(stdExportPath, Addendum.AlbisPath, "callback=ZeigsMir, employeeshorts=" employeeshorts)

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Patientendatenbank laden
	cPat      :=  ReadPatDB()
;}

;}

;{ ⚞──────────────────      VORBEREITUNGEN       	──────────────────⚟

; Basispfad  	- Unterverzeichnisse anlegen                            	;{

  ; Vorhandensein des Basis- oder Standard-Exportpfad prüfen
	If !FileExistZ(stdExportPath "\")
		throw "Das Standard-Exportverzeichnis ist nicht vorhanden.`n" stdExportPath "\"

  ; xData - Unterverzeichnis für die ungefilterten Daten aus den Datenbanken (als .json Dateien gespeichert)
	If !FilePathCreate(stdExportPath "\xData")
		throw "das Backupverzeichnis für entfernte Daten konnte nicht angelegt werden.`n" stdExportPath "\xData\"

  ; Unterverzeichnisse für das Verschieben von Daten anlegen
	If !FilePathCreate(stdExportPath "\verschoben")
		throw "das Backupverzeichnis für entfernte Daten konnte nicht angelegt werden.`n" stdExportPath "\verschoben\"

  ; die Unterverzeichnisse für die 3 Kategorien für verschobene (nicht zu exportierende Daten) anlegen
	For category, path in settings.folderToPath
		If !FilePathCreate(stdExportPath "\verschoben\" path)
			throw "das Backupverzeichnis für entfernte Daten der Kategorie '" category "' konnte nicht angelegt werden.`n" stdExportPath "\verschoben\" path "\"

;}

; SQLite     	- Backup der sqlite Datenbank                          	;{
	If FileExist(settings.paths.DBLocation) && !FileExist(fpBup := settings.paths.DBBupLocation . A_YYYY A_MM A_DD A_Hour ".db")
		FileCopy, % settings.paths.DBLocation, % fpBup
;}

; SQLite     	- Datenbank anlegen / lesen                           	;{
	QEDB := new SQLiteDB
	If !QEDB.OpenDB(settings.paths.DBLocation, DBAccess) {
		SQLiteErrorWriter("Öffnen der Datenbank fehlgeschlagen", QEDB.ErrorMsg, QEDB.Errorcode, true)
		ExitApp
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Prüfen ob die SQLite Datenbank angelegt ist
	If !FileExist(settings.paths.DBLocation) {
		MsgBox, 16, SQLite Error, % "Die notwendige Datenbank konnte nicht angelegt werden"
		ExitApp
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Überprüfung der Datenbanktabellen
	tablefound := 0 ;{
	If !QEDB.GetTable(sql.TableLookup, Result)
		SQLiteErrorWriter("Tabellen konnten nicht gelesen werden", QEDB.ErrorMsg, QEDB.Errorcode, true)
	For each, row in Result.Rows
		For each, table in row
			tablefound += (table = "EXPORTS") ? 0x1 : (table = "RECIPIENTS") ? 0x2 : (table = "MOVED") ? 0x4 : 0
	Result := ""
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Tabelle Exports anlegen, falls nicht gemacht
	If !(tablefound & 0x1) {
		If !QEDB.Exec(sql.Exports.CreateTable) {  ; Exportierte Daten
			SQLiteErrorWriter("'EXPORTS' Tabelle - Fehler bei Erstellung", QEDB.ErrorMsg, QEDB.Errorcode, true)
			ExitApp
		}
		else
			SciTEOutput("'EXPORTS' Tabelle angelegt")
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Tabelle Recipients anlegen, falls nicht gemacht
	If !(tablefound & 0x2) {
		If !QEDB.Exec(sql.Recipients.CreateTable)  { ; Empfänger
			SQLiteErrorWriter("'RECIPIENTS' Tabelle - Fehler bei Erstellung", QEDB.ErrorMsg, QEDB.Errorcode, true)
			ExitApp
		}
		else
			SciTEOutput("'RECIPIENTS' Tabelle angelegt")
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Tabelle Moved anlegen, falls nicht gemacht
	If !(tablefound & 0x4) {
		If !QEDB.Exec(sql.Moved.CreateTable) {  ; Exportierte Daten
			SQLiteErrorWriter("'MOVED' Tabelle - Fehler bei Erstellung", QEDB.ErrorMsg, QEDB.Errorcode, true)
			ExitApp
		}
		else
			SciTEOutput("'MOVED' Tabelle angelegt")
	}
	;~ else
		;~ SciTEOutput("'MOVED' Tabelle ist vorhanden")

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Datensätze hinzufügen es eine .sql Datei gibt, welche noch nicht eingelesen wurde #### noch nicht programmiert
	If !(tablefound & 0x1) &&  !(tablefound & 0x2) {
		If (_SQL := FileOpen(settings.paths.sqlitefiles "\SQLDaten_Exports Tabelle.sql", "r", "UTF-8").Read()) {
			;~ If !QEDB.Exec(_SQL)
				;~ SQLiteErrorWriter("abelle - Fehler beim Schreiben von Datensätzen", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
		}
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Spaltenbezeichnungen von Exports für Listview Gui
	If !QEDB.Query(sql.Exports.GetTable, RecordSet) ;{
		SQLiteErrorWriter("Tabelle konnte nicht gelesen werden", QEDB.ErrorMsg, QEDB.Errorcode, true, sql.GetTable.Exports)
	If RecordSet.ColumnCount {
		Loop, % RecordSet.ColumnCount
			TExports.Push(RecordSet.ColumnNames[A_Index])
	}
	RecordSet := ""

	sqlExportsTable := {"PatID"	    	: {"col": 1	, "lvname":"PatID"       	, "lvcol":1	   	, "isDate":""}
					     			, "SName"	    	: {"col": 2	, "lvname":"Name"        	, "lvcol":2	   	, "isDate":""}
					     			, "PName"	    	: {"col": 3	, "lvname":"Vorname"   	 	, "lvcol":3  		, "isDate":""}
					     			, "Birth"	    	: {"col": 4	, "lvname":"Birth"	     	, "lvcol":13		, "isDate":["Geburt" 	, 3]}
					     			, "SDate"	    	: {"col": 5	, "lvname":"DateS"	     	, "lvcol":14		, "isDate":["SDatum"	, 5]}
					     			, "Recipient" 	: {"col": 6	, "lvname":"BSNR"		      , "lvcol":12		, "isDate":""}
					     			, "Medium"	  	: {"col": 7	, "lvname":"VMedium"      , "lvcol":7	   	, "isDate":""}
					     			, "Route" 	  	: {"col": 8	, "lvname":"VRoute"	      , "lvcol":8	   	, "isDate":""}
					     			, "Shipment"  	: {"col": 9	, "lvname":"ShipD"	      , "lvcol":15		, "isDate":["VDatum"	, 9]}
					     			, "More1"	    	: {"col": 10, "lvname":"Anmerkungen"  , "lvcol":10		, "isDate":""}
					     			, "ExportPath"	: {"col": 11, "lvname":"Export"			  , "lvcol":11		, "isDate":""}
					     			, "More3"	    	: {"col": 12, "lvname":""						  , "lvcol":0	   	, "isDate":""}}

	;						            	 1		2			3			   4			5				6				7			 8     	  9				10			11	  12	   13		14	  15	    16
	ExportsCols	:=          "PatID|Name|Vorname|Geburt|SDatum|Empfänger|VMedium|VRoute|VDatum|Anmerkungen|Export|BSNR|Birth|DateS|ShipD|More3"
	;						            	 1			2			3		  4		5			   6			    	7			 8     	  9		10		  11	           12	       13     14	     15		 16		# -- wird nicht gespeichert
	ExportsTbl 	:= StrSplit("PatID|SName|PName|Birth|SDate|RecipientName|Medium|Route|Shipment|More1|ExportPath|Recipient|Birth|SDate|Shipment|More3", "|")	; Spaltennamen in der SQLite Datenbank

	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; SDates auslesen für Gui ComboBox
	If !QEDB.GetTable(_SQL := "SELECT SDate FROM EXPORTS;", Result) ;{
			SQLiteErrorWriter("Export Tabelle - Get SDates Fehler", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL )
	vCtrls.3.value := vCtrls.3.default
	If IsObject(Result)
		For each, row in result.Rows {
			rdate := ConvertDBASEDate(row.1)
			If !InStr(vCtrls.3.value, rdate)
				vCtrls.3.value .= (vCtrls.3.value ? "|" : "") rdate
		}
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Empfänger auslesen
	If !QEDB.GetTable(_SQL := "SELECT * FROM RECIPIENTS;", Result) ;{
			SQLiteErrorWriter("Export Tabelle - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, false, _SQL)
	vCtrls.4.value := vCtrls.4.default
	If IsObject(Result)
		For each, row in result.Rows {

			bsnr  	:= row.7
			anr   	:= RegExReplace(row.1, "i)\s*Herr\s*"          , "Hr.")
			anr   	:= RegExReplace(anr  , "i)\s*Frau\s*"          , "Fr.")
			btitel	:= RegExReplace(row.2, "i)\s*Dr\.*\s*med\.*\s*", "Dr.")

			Recip := bsnr>3 ? ANR " " btitel " " row.3 ", " row.4 " [" bsnr "]" : bsnr=1 ? "Patient" : "Angehöriger"

		  ; für Dateiverzeichnisse aufheben
			If (bsnr > 3 && !Recipients.haskey("#" bsnr))
				Recipients["#" bsnr] := Recip

		  ; der ComboBox hinzufügen
			If !(vCtrls.4.value ~= "\|" Recip "")
				vCtrls.4.value .= (vCtrls.4.value ? "|" : "") Recip

		}
	;}

	vCtrls.5.value := vCtrls.5.Default
	vCtrls.6.value := vCtrls.6.Default

;}

;}

;{ ⚞──────────────────           GUI             	──────────────────⚟

; GUI        	- Dokumentationsinterface                              	;{
	global QE, QELVExpFilter, QERNEWF, QEhLvExp, QECM
	bgt := " Backgroundtrans "
	settings.prgControls := ["QETblChgPrg"	, "QELVExpFilterPrg", "QEStatsPrg", "QECryptPrg"	, "QEMarkerPrg", "QEVBPrg"]
	settings.txtControls := ["QEsqlChanges"	, "QELVFilterTxt" 	, "QEStatsTxt", "QECryptText"	, "QEMarkerTxt", ["QEVBTxt1", "QEVBTxt2"]]
	settings.VBControls  := ["QEVBData"   	, "QEVBIprint"    	, "QEVBIprinter"]

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Kontextmenu vorbereiten
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 	func_QEMoveS      	:= Func("QEMoveRows").Bind("shipped")
 	func_QEMoveU      	:= Func("QEMoveRows").Bind("misc")
 	func_QEMoveM      	:= Func("QEMoveRows").Bind("mortal")
 	func_QE7zip        	:= Func("QE7zip")
 	func_QEPatAddress 	:= Func("QEPatAddress")
 	func_QEKarteikarte 	:= Func("QEKarteikarte")
 	func_QEExplorer   	:= Func("QEFileExplorer")
	Menu, QECMSub1	, Add, % "bereits verschickt"                    	, % func_QEMoveS
	Menu, QECMSub1	, Add, % "andere Gründe"                         	, % func_QEMoveU
	Menu, QECMSub1	, Add, % "verstorben"                            	, % func_QEMoveM
	Menu, QECM     	, Add, % "Verschieben nach"                      	, :QECMSub1
	Menu, QECM     	, Add, % "mit 7zip packen"                       	, % func_QE7zip
	Menu, QECM     	, Add
	Menu, QECM     	, Add, % "Patientenadresse ins Clipboard"        	, % func_QEPatAddress
	Menu, QECM     	, Add, % "in Albis anzeigen"                     	, % func_QEKarteikarte
	Menu, QECM    	, Add, % "Verzeichnis im Dateiexplorer anzeigen"	, % func_QEExplorer

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Gui startet hier
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
	scale := A_ScreenDPI > 96 ? "+DPIScale" : ""
	dpiF := Round(A_ScreenDPI//96)
	settings.fSize := {"s7"   : ("s" Floor(7*dpiF) 		" q5")
						 	  	,	 "s8" 	: ("s" Floor(8*dpiF) 		" q5")
									,	 "s9" 	: ("s" Floor(9*dpiF) 		" q5")
								  ,	 "s10"	: ("s" Floor(10*dpiF)  	" q5")
									,	 "s12"	: ("s" Floor(12*dpiF)  	" q5")
									,	 "s18"	: ("s" Floor(18*dpiF)  	" q5")}

	if scale
		Gui, QE: new, % "-DPIScale"

	Gui, QE: new, % "+DPIScale"
	Gui, QE: Default

	Gui, % "+hwndhQE +OwnDialogs"
	Gui, Color, cFFFFE5, cDDFFFF

	dpiW := DllCall("GetDpiForWindow", "ptr", hQE)

	Gui Font, % settings.fSize.s7, Arial
	Gui Add, Text    		, % "xm ym -Theme"                                                                                             	, % "Schnelleingabefeld z.B. 'Mu, Klaus' o.'12345' o. 13.04.1976"
	Gui Font, % settings.fSize.s10, Arial
	Gui Add, Edit    		, % "xm y+1	  	  w300  r1  vQEFEdit  hwndQEhFEdit  "
	func_call := Func("QELVScroll").Bind("UserInput")
	GuiControl, QE: +g     	, QEFEdit	, % func_call
	;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Eingabesteuerelemente mit Beschriftungen hinzufügen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	Gui Font, % settings.fSize.s10, Arial
	cp := GuiControlGet("QE", "Pos", "QEFEdit")
	For each, ctrl in vCtrls {
		If ctrl.user {
			Gui Add, Text         	, % (ctrl.col = 1 ? "xm" : "x+5") " y" cp.Y+cp.H+5 " " ctrl.opt " " ctrl.topt	, % ctrl.txt
			Gui Add, % ctrl.ctype	, % "y+1 " ctrl.opt " " ctrl.copt " v" ctrl.vName  " HWNDhCtrl"               	, % ctrl.value
			If (ctrl.cType = "Edit") {
				settings.debug ? SciTEOutput("hEdit (" ctrl.name "): " hCtrl) : ""
				EM_SetMargins(hCtrl, 0, 0)
			}
		}
	}

	Gui Font, % settings.fSize.s8, Arial
	RegExMatch(settings.fSize.s8, "i)s(\d+)", fSize)
	Btntxt 	:= "Patient suchen"
	Gui Add, Button		, % "xm y+5 w" Floor(fSize1*StrLen(Btntxt)//2) " vQESearch 	 	gQEBHandler "                                           	, % Btntxt
	Btntxt := "Datensatz anlegen"
	Gui Add, Button		, % "x+10 	w" Floor(fSize1*StrLen(Btntxt)//2) " vQENewItem  	gQEBHandler "                                            	, % Btntxt
	GuiControl, QE: Disable, QENewItem
	;}


  ; PATIENTENAKTEN EXPORTIEREN       	;{
	Gui Add, Button		, % "x+10            	vQEExport 	gQEBHandler hwndQEhExport"                                                     	, % "Patientenakten exportieren"
	cp := GuiControlGet("QE", "Pos", "QEExport")
	Gui Add, Checkbox   , % "xp+3 y+3       vQERNEWF    gQEBHandler hwndHwnd1" (settings.reNewData ? "Checked" : "")                   	, % "Daten neu erstellen"
	Gui Add, Checkbox   , % "y+5            vQENoLData  gQEBHandler hwndHwnd2" (settings.noLabData ? "Checked" : "")                   	, % "Export ohne Labor"
	Gui Add, Button	  	, % "x" cp.X " y+12 w" cp.W " h" cp.H-2 "	vQEExplorer hwndQEhExplorer"	                                      	, % "Pfad im Explorer öffnen"
	ep := GetWindowSpot(Hwnd1)
	GuiControl, QE: MoveDraw, % hwnd1, % "h" ep.H+1
	GuiControl, QE: MoveDraw, % hwnd2, % "h" ep.H+1
	GuiControl, QE: +g     	, QEExplorer	, % func_QEExplorer
	;}


  ; TABELLENÄNDERUNGEN              	;{
	noStEdg := " -E0x00020000 "
	cp := GuiControlGet("QE", "Pos", "QEExport")
	gBx := cp.X+cp.W+10, tblCw := 120 , tblPrgC := "DFFF00", tblPrgB := "E7EBC7"
	Gui Add, GroupBox	, % "x" gBx " y" cp.Y-6 " w" tblCw " h200 " noStEdg " vQETblChgBox"
	Gui Add, Progress	, % "xp+0     y" cp.Y+1 " h17 w" tblCw " vQETblChgPrg c" tblprgC " Background" tblPrgB
	Gui Add, Text   	, % "xp+0     yp+1      w" tblCw     " Center vQEsqlChanges" bgt                                                	, % "Tabellenänderungen"
	Gui Add, Button		, % "x" gBx+7 " y+9 h16 w" tblCw-14  "    vQEBulkCh 	 	gQEBHandler "                                           	, % "übernehmen"
	Gui Add, Button		, % "y+4 h16 	w" Floor((tblCw-14)/2) "   	vQEsqlSave   	gQEBHandler hwndQEhsqlSave"                             	, % "sichern"
	Gui Add, Button		, % "x+3 h16 	w" Floor((tblCw-14)/2) "  	vQEsqlCancel 	gQEBHandler hwndQEhsqlCancel"                            	, % "verwerfen"
	cp 	:= GuiControlGet("QE", "Pos", "QEsqlChanges")
	dp	:= GuiControlGet("QE", "Pos", "QEsqlCancel")
	cp	:= GuiControlGet("QE", "Pos", "QEBulkCh")
	GuiControl, QE: MoveDraw, QEsqlCancel, % "x" cp.X+cp.w-dp.w
	GuiControl, QE: Disable , QEsqlSave
	GuiControl, QE: Disable , QEsqlCancel
	;}


  ; EXPORTE AUSBLENDEN              	;{
	cp := GuiControlGet("QE", "Pos", "QETblChgBox")
	gBx := cp.X+cp.W+10 , gBy := cp.Y, LVFw := 160
	Gui Add, GroupBox	, % "x" gBx " y" gBy " w" LVFw " h200 " noStEdg " vQELVExpFilterBox"
	Gui Add, Progress	, % "xp+0 y" gBy+7 " w" LVFw " h17 vQELVExpFilterPrg cFFA24D BackgroundF1DBC9"
	Gui Add, Text      	, % "xp+0 y" gBy+9 " w" LVFw-2 " Center vQELVFilterTxt   "         	bgt                                        	, % "Exporte ausblenden"
	Gui Add, Checkbox  	, % "xp+4 y+7   	vQELVExpFilter  	gQEBHandler " 	bgt " "       	(settings.LVExpFilter	? "Checked" : "")  	, % "fehlerfreie und Versendete"
	Gui Add, Checkbox  	, % "y+5           	vQELVExpFilter2  	gQEBHandler " 	bgt " "     	(settings.LVExpFilter2 ? "Checked" : "")   	, % "Versendete"
	Gui Add, Checkbox  	, % "y+5           	vQELVExpFilter3  	gQEBHandler " 	bgt " "     	                                           	, % "nach Empfänger"
	GuiControl, QE: MoveDraw, QELVExpFilter	, % "h" ep.H+1
	GuiControl, QE: MoveDraw, QELVExpFilter2, % "h" ep.H+1
	GuiControl, QE: MoveDraw, QELVExpFilter3, % "h" ep.H+1
	cp 	:= GuiControlGet("QE", "Pos", "QELVExpFilterPrg")
	dp	:= GuiControlGet("QE", "Pos", "QELVExpFilter3")
	fbxh := dp.Y + dp.H - cp.Y + 12
	GuiControl, QE: MoveDraw, QELVExpFilterBox	, % "h" fbxh
	GuiControl, QE: MoveDraw, QETblChgBox    	, % "h" fbxh
	;}


  ; RELOAD             	             	;{
	Gui Font, % settings.fSize.s8, Arial
	cp := GuiControlGet("QE", "Pos", "QEMore1")
	Gui Add, Button		, % "x" cp.X+cp.W-71 " y" cp.y+cp.H+3 "	vQEReload 	 gQEBHandler "                                               	, % "Reload Script"
	;}


  ; EXPORTSTATISTIKEN               	;{
	cp := GuiControlGet("QE", "Pos", "QELVExpFilterBox")
	dp := GuiControlGet("QE", "Pos", "QENoLData")
	gBx	:= cp.X + cp.W + 10, gBy := cp.Y, statSw := 130
	Gui, Add, GroupBox	, % "x" gBx " y" gBy " w" statSw " h" fbxh " " noStEdg " vQECounterBox"                                       	, % ""
	Gui, Add, Progress	, % "xp+0 y" gBy+7 " w" statSw " h17 " noStEdg " vQEStatsPrg c00FF79 BackgroundC2D6CB"
	Gui Add, Text      	, % "xp+0 y" gBy+9 " w" statSw  bgt " Center vQEStatsTxt"                                                      	, % "Exportstatistiken"
	Gui Font, Normal
	Gui Add, Text     	, % "xp+5 y+9 w80"  bgt                                                                                        	, % "exportiert"
	Gui Font, bold
	Gui Add, Text      	, % "x+0    	vQELVExpCounter1  "  bgt                                                                         	, % "0              "
	Gui Font, Normal
	Gui Add, Text     	, % "x" gBx+5 " y+1  w80"                bgt                                                                   	, % "vorbereitet"
	Gui Font, bold
	Gui Add, Text     	, % "x+0    	vQELVExpCounter2  "  bgt                                                                         	, % "0              "
	Gui Font, Normal
	Gui Add, Text      	, % "x" gBx+5 " y+1  w80"                bgt                                                                   	, % "versendet"
	Gui Font, bold
	Gui Add, Text      	, % "x+0    	vQELVExpCounter3  "  bgt                                                                         	, % "0              "
	Gui Font, Normal
	Gui Add, Text      	, % "x" gBx+5 " y+1  w80"                bgt                                                                   	, % "ausstehend"
	Gui Font, bold
	Gui Add, Text      	, % "x+0    	vQELVExpCounter4  "  bgt                                                                         	, % "0              "
	cp 	:= GuiControlGet("QE", "Pos", "QECounterBox")
	Gui Add, Button		, % "x" cp.X " y" cp.Y+cp.H+5 " w" statSw  " 	vQEStatistic 	 	gQEBHandler hwndQEhStatic"                        	, % "Statistik erstellen"
	;}


  ; 7z EINSTELLUNGEN                 	;{
	gBx := cp.X+cp.w+10 , gBy := cp.Y, crptW:=200
	Gui, Add, GroupBox	, % "x" gBx " y" gBy " w" crptW " h" fbxh " " noStEdg " vQECryptBox"
	Gui, Add, Progress	, % "x" gBx+1 " y" gBy+7 " w" crptW-1 " h17 vQECryptPrg c8080FF BackgroundC9DBF1"
	Gui Add, Text     	, % "x" gBX+5 " y" gBy+9 " w" crptW " Center vQECryptText " bgt                                                	, % "7zip"
	Gui Add, Edit     	, % "y+7  vQEZPWDH w" crptW-40 " Password gQEFocus HWNDQEhZPWDH"          	                                   	, % StringDecoder()   ; stellt das letzte passwort wieder her
	Gui Add, Edit     	, % "xp+0 yp+0 vQEZPWDV w" crptW-40 " gQEFocus HWNDQEhZPWDV"                                                   	, % StringDecoder()   ; stellt das letzte passwort wieder her
	Gui Add, Button    	, % "x+2  vQEZPWDSHow  "                    	                                                                 	, % "  👀"
	Gui Add, Checkbox   , % "x" gBx+5 " y+0 vQEZCrypt   "                        (settings.NoZipEncryption ? "Checked":"")             	, % "kein Passwort"
	Gui Add, Checkbox   , % "x+3 yp+0       vQEZLEtter  "                        (settings.ZipWithLetter   ? "Checked":"")             	, % "Versandbriefe"
	Gui Add, Button    	, % "x" gBx+5 " y+6 vQEZCompress"                                                                             	, % "packen"
	Gui Font, % settings.fSize.s7
	Gui Add, Edit     	, % "x+5 yp+1  w126 h17   vQEZPath"                                                                            	, % settings.paths.networkpath
	Gui Add, Button    	, % "x+3 yp+0  h17   vQEZSelPath gQEBHandler"                                                                 	, % "..."
	Gui Font, % settings.fSize.s8
	dp := GuiControlGet("QE", "Pos", "QEZPWDSHow")
	GuiControl, QE: MoveDraw, QEDBG     	, % "h" ep.H+1   ; nur für WIndows Server 2019 - Checkboxen werden oben abgeschnitten (so nicht mehr)
	GuiControl, QE: MoveDraw, QEZCrypt  	, % "h" ep.H-2
	GuiControl, QE: MoveDraw, QEZLetter  	, % "h" ep.H-2
	GuiControl, QE: MoveDraw, QEZPWDH    	, % "h" ep.H-2 " w" crptW-10-dp.W-5
	GuiControl, QE: MoveDraw, QEZPWDV    	, % "h" ep.H-2 " w" crptW-10-dp.W-5
	GuiControl, QE: MoveDraw, QEZPWDSHow 	, % "h" ep.H-2 " x" gBX+crptW-dp.W-5
	func_call := Func("QETogglePWDVisibility").Bind("NoZipEncryption")
	GuiControl, QE: +g     	, QEZPWDShow	, % func_call
	func_call := Func("QETogglePWDVisibility").Bind("UserBUttonClick")
	GuiControl, QE: +g     	, QEZCrypt  	, % func_call
	GuiControl, QE: +g     	, QEZCompress	, % func_QE7zip
	QETogglePWDVisibility("noUserButtonClick")
	;}


  ; Brennen auf DVD/CD / Kopieren   	;{
	Gui Font, % settings.fSize.s8 " Normal"
	cp := GuiControlGet("QE", "Pos", "QECryptBox")
	dp := GuiControlGet("QE", "Pos", "QEShipment")
	gBX := cp.X+cp.W+10, gBy := cp.Y, gBw := dp.X - cp.X - cp.W - 16
	Gui Add, GroupBox	, % "x" gBx " y" gBy " w" gBw " h" fbxh+23 " " noStEdg " vQEBurnBox"
	Gui Add, Progress	, % "x" gBx " y" gBy+7 " w" gBw " h17 vQEBurnPrg c80D2FF BackgroundE2BBFF"
	Gui Add, Text    	, % "x" gBX " y" gBy+9 " w" gBw " vQEBurnTxt1 Center " bgt                                                        	, % "Brennen"
	Gui Font, % settings.fSize.s7 " Normal"
	Gui Add, Text    	, % "x" gBX+5 " y+7 " bgt                                                                                          	, % "Laufwerk"
	Gui Add, Edit	  	, % "x+2 vQEBurnDrive gQEBHandler"                                                                                 	, % "rdp | Q:"
	Gui Font, % settings.fSize.s8 " Normal"
	Gui Add, Button 	, % "x" gBX " y+3 vQECopyTo gQEBHandler "                                                                         	, % "Kopieren"
	func_call := Func("QECopyFilesToFolder")
	GuiControl, QE: +g     	, QECopyTo  	, % func_call

	;}


  ; MARKIEREN ALS                   	;{
	Gui Font, % settings.fSize.s8, Arial
	cp := GuiControlGet("QE", "Pos", "QEShipment")
	gBx := cp.X , gBy := cp.Y+cp.H-1, Mrkrw := cp.W
	Gui Add, GroupBox	, % "x" gBx " y" gBy " w" Mrkrw " h" fbxh " " noStEdg " vQEMarkerBox"
	Gui Add, Progress	, % "x" gBx+1 " y" gBy+7 " w" Mrkrw " h17 vQEMarkerPrg cBFFF80 BackgroundDBF1C9"
	Gui Add, Text  		, % "x" gBx+1 " y" gBy+9 " w" Mrkrw " Center vQEMarkerTxt " bgt                                	                  	, % "markieren als"
	Gui Font, % settings.fSize.s8, Arial
	Gui Add, Button		, % "x" gBx+2 " y+9" " w" Mrkrw-4 " Center vQEMarker gQEBHandler"                                                 	, % "versendet"
	Gui Add, Button		, % "x" gBx+2 " y+1" " w" Mrkrw-4 " Center vQEEraser gQEBHandler"                                             	   	, % "unversendet"
	;}


  ; VERSANDBRIEF                    	;{
	Gui Font, % settings.fSize.s8
	cp := GuiControlGet("QE", "Pos", "QEMarkerBox")
	gBX := cp.X+cp.W+5, gBy := cp.Y, VBw := 500, VBTw := 110
	Gui Add, GroupBox	, % "x" gBx " y" gBy " w" VBw " h" fbxh+23 " " noStEdg " vQEVBBox"
	Gui Add, Progress	, % "x" gBx " y" gBy+7 " w" VBw " h17 vQEVBPrg cC980FF BackgroundE2BBFF"
	Gui Add, Text    	, % "x" gBX " y" gBy+9 " w" VBw//2 " Center vQEVBTxt1 " bgt                                                       	, % "Absender"
	Gui Add, Text    	, % "x" gBX+VBw//2+1 " y" gBy+9 " w" VBw//2-1 " Center vQEVBTxt2" bgt                                             	, % "Versandbrief Einstellungen"
	Gui Font, % settings.fSize.s8
	Gui Add, Edit    	, % "x" gBX+1 " y+9 h" fbxh-32 " w" VBw//1.9 " vQEVBData "  bgt                                                    	, % settings.sender.name  		"`n"
																																																														  						.	settings.sender.Facharzt 	"`n"
																																																																	  			.	settings.sender.ORT				"`n"
																																																																  				.	settings.sender.Mail
	cp := GuiControlGet("QE", "Pos", "QEVBData")
	Gui Add, Checkbox	, % "x" cp.X+cp.W+5 " y" cp.Y+2 " vQEVBAct 	 "                     (settings.sendLetter 	? "Checked":"")        		, % "Versandbrief erstellen"
	Gui Add, Checkbox	, % "x" cp.X+cp.W+5 " y+5 vQEVBIprint	 "                           (settings.printLetter 	? "Checked":"")        		, % "Versandbrief drucken"
	Gui Font, % settings.fSize.s7
	Gui Add, DDL     	, % "x" cp.X+cp.W+5 " y+7 w" VBw//2.2 " vQEVBIprinter "                                                          		, % GetPrinters()
	Gui Font, % settings.fSize.s8
	Gui Add, ComboBox      , % "x" gBx+1 " y+4 w" Round(VBw//2.5) " vQEVBLBPrint1 "                                                       , % ""
	Gui Add, ComboBox      , % "x+3            w" Round(VBw//2.5) " AltSubmit vQEVBLBPrint2 "                                             , % ""
	Gui Add, Button 	, % "x+3 w" Round(VBw//5)-7  " vQEVBPrintNow gQEBHandler +0x1 Disabled"                                           	, % "Drucken"
	cp := GuiControlGet("QE", "Pos", "QEVBDDLPrint")
	GuiControl, QE: MoveDraw		, QEVBPrintNow	, % "h" ep.H
	GuiControl, QE: MoveDraw		, QEVBLBPrint1	, % "h" ep.H+1
	GuiControl, QE: MoveDraw		, QEVBLBPrint2	, % "h" ep.H+1
	GuiControl, QE: MoveDraw		, QEVBIprint   	, % "h" ep.H+1
	GuiControl, QE: MoveDraw		, QEVBAct     	, % "h" ep.H+1
	func_call := Func("QESetPropVersand")
	GuiControl, QE: +g     			, QEVBAct     	, % func_call
	func_call := Func("QELetterList")
	GuiControl, QE: +g     			, QEVBLBPrint1	, % func_call
	func_call := Func("QELetterChoosed")
	GuiControl, QE: +g     			, QEVBLBPrint2	, % func_call
	func_call := Func("QELetterPrint")
	GuiControl, QE: +g     			, QEVBPrintNow 	, % func_call

	GuiControl, QE: ChooseString, QEVBIprinter	, % settings.pdfprinter
	If !settings.sendletter
		QESetPropVersand()
	If (editHwnd := GetWindow(GuiControlGet("QE", "hwnd", "QEVBLBPrint2"), 5))
		SendMessage(0x00CF, 1,,, "ahk_id " editHwnd )

	;}


  ; OPTIONEN                        	;{ 🗄 🗃 🗊  ⛘ 📁 📂 🗂 🗁 🗀 🖿  ⚙ ❖ 🗃
	Gui Font, % settings.fSize.s8
	cp := GuiControlGet("QE", "Pos", "QEVBBox")
	cpX := cp.X+cp.W+10, cpW := 120
	Gui Add, Checkbox  	, % "x" cpX " y" cp.Y+cp.H-36 " w" cpW  " vQEDBG gQEBHandler "   (settings.dbg   	 ? "Checked":"")               	, % "Skript debugging"
	Gui Add, Checkbox   , % "y+5 w" cpW " vQESQLWP 	gQEBHandler "                  (settings.noSQLWrite    ? "Checked":"")               	, % "SQLite Schreibschutz"
	GuiControl, QE: MoveDraw, QEDBG     	, % "h" ep.H+1
	GuiControl, QE: MoveDraw, QESQLWP   	, % "h" ep.H+1
	;}


  ; EXPORTVERZEICHNIS               	;{
	Gui Font, % settings.fSize.s9 " Bold"
	cp 	:= GuiControlGet("QE", "Pos", "QEStatistic")
	gBX := cp.X+cp.W+10, gBy := cp.Y-5, EXPw:=450, EXPh:= cp.H-2
	Gui Add, Progress	, % "x" gBx+1         " y" gBy+6 " w" EXPw//2  " h" EXPh "    cFDD7D7 BackgroundFDD7D7  vQEXprg1 +E0x4000 "
	Gui Add, Progress	, % "x" gBx+1+EXPw//2 " y" gBy+6 " w" EXPw//2  " h" EXPh "    cF0C6C6 BackgroundF0C6C6  vQEXprg2 +E0x40000 "
	Gui Add, Text    	, % "x" gBx+6         " y" gBy+8               " h" Exph "    vQEstdExpPath    "  bgt                                           , % "Exportverzeichnis"
	dp 	:= GuiControlGet("QE", "Pos", "QEstdExpPath")
	Gui Font, % settings.fSize.s9 " Normal"
	Gui Add, Text    	, % "x+10 y" gBy+7 " w" EXPw-dp.W-12          " " bgt                                                                              	, % stdExportPath
	GuiControl, QE: MoveDraw, QEXprg1, % "w" dp.W+9
	GuiControl, QE: MoveDraw, QEXprg2, % "x" dp.X+dp.W+5 " w" EXPw-dp.W-6
	;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Sprungmarkenalphabet
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	cp 	:= GuiControlGet("QE", "Pos", "QELVExpFilterBox")
	cpY := cp.Y+cp.H+5
	Gui Font, % settings.fSize.s8, Futura Bk Bt
	for lpos, letterkey in StrSplit("abcdefghijklmnopqrstuvwxyz")  {
		fnkey := Func("QELVScroll").Bind(letterkey)
		Gui, Add, Button, % (lpos=1 ? "xm" : "x+2")  " "  (lpos=1 ? "y" cpY : "")  " h12 hwndQEhlkey" , % letterkey
		GuiControl +g, % QEhlkey, % fnkey
		GuiControl, QE: MoveDraw, % QEhlkey, % "h18"
	}
	Gui Font, Bold
	Gui Add, Text      	, % "x+3 yp+2   	vQELetters "  bgt                                                                              	, % "                           "
	Gui Font, Normal
	;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Standardeinstellung der Gui Steuerfelder
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	QEStandards()

	Gui Show, Hide, Addendum - Quickexporter ; h615
	GSize := GetWindowSpot(hQE)
	;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Listviewausgabe
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	Gui Font, % settings.fSize.s9, Arial
	LvExpOpt := "-ReadOnly Grid Checked BackgroundEFEED8 AltSubmit +LV0x8000000"
	cp 	:= GuiControlGet("QE", "Pos", "QELVExpFilterBox")
	dp 	:= GuiControlGet("QE", "Pos", "QEstdExpPath")
	cpY 	:= cp.Y+cp.H+25
	Gui Add, ListView  	, % "xm y" cpY " r42 w" Round(1830*dpiF) " " LvExpOpt " vQELvExp HWNDQEhLvExp gQEBHandler", % ExportsCols
	For colNr, colwidth in (settings.LVColWidths := [70,120,120,70,70,180,90,90,70,450,450,0,0,0,0,0])
		LV_ModifyCol(colNr, colwidth)

	;~ GuiControl, QE: MoveDraw, QEstdExpPath, % "y" cpY - dp.H - 5

	;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Statusbar für Fortschrittsanzeigen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	Gui, Font, % settings.fSize.s10 " Normal"
	Gui, Add, StatusBar,, % ""
	Gui Show, % "w" (guiW := 1800) " Hide"


	GSize := GetWindowSpot(hQE)
	If !IsInsideVisibleArea((guiX ~= "\-*\d+" ? guiX : GSize.X) , (guiY ~= "\-*\d+" ? guiY : GSize.Y), GSize.w, GSize.h, CoordInjury) {
		SciTEOutput("Gui wäre ausserhalb des sichtbaren Bereiches mit: x" guiX " y" guiY " w"  guiW " h" GSize.h "`nRückgabewert: " CoordInjury)
		guiX := "Center", guiY := "Center"
	}
	;}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Create a new instance of LV_InCellEdit for QEhLvExp with options HiddenCol1 and BlankSubItem
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	global ICELLvExp := New LV_InCellEdit(QEhLvExp, true, True, "QEInCellEditHandler")
	ICELLvExp.SetColumns(4,5,6,7,8,9,10,11)

	Gui Show, % "x" guiX " y" guiY " w" guiW " AutoSize"

	sb1Width := Floor(guiW*0.3)
	sb2Width := Floor(guiW*0.3)
	sb3Width := Floor(guiW*0.18)
	sb4Width := Floor(guiW*0.18)

	hSBar := SB_SetParts(sb1Width, sb2Width, sb3Width, sb4Width)
	SB_SetText("...Willkommen", 1)
	SB_SetProgress(0, 2, "+Smooth cFFBB00 Background0000FF", hSBar)
	SB_SetText("`t"  "00:00:00", 5)

  ; Zahl der sichtbaren Listviewzeilen
	SendMessage, 0x1028,,,, % "ahk_id " QEhLvExp
	settings.CounterPage := ErrorLevel
	;}


  ;}

; GUI        	- EXPORTS Tabelle, Daten laden u. anzeigen            	;{
	If !QEDB.GetTable(sql.Exports.GetTable, ExpTable)
		SQLiteErrorWriter("Tabelle konnte nicht gelesen werden", QEDB.ErrorMsg, QEDB.Errorcode, true, sql.Exports.GetTable)
	Gui, QE: ListView, QELvExp

	QEShowTable(ExpTable)


;}

; GUI        	- Hotkey Interaktionen                                	;{
	global RBRow

	Hotkey, !+ß               	, QEGuiClose

; __________________________________________________________________________
; --------------------------------------------------------------------------
	Hotkey, IfWinActive, % "Addendum - Quickexporter ahk_class AutoHotkeyGUI"
; --------------------------------------------------------------------------
; ‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾
	Hotkey, Enter	            	, QEFocus
	Hotkey, NumpadEnter       	, QEFocus
	Hotkey, Tab               	, QEFocus
	Hotkey, RButton            	, QEFocus

	func_call := Func("QELVScrollHk").Bind("PgUp")
	Hotkey, PgUp               	, % func_call
	func_call := Func("QELVScrollHk").Bind("PgDn")
	Hotkey, PgDn              	, % func_call

	Hotkey, IfWinActive


	isHovered_LV := Func("isHovered")
; __________________________________________________________________________
; --------------------------------------------------------------------------
	#If isHovered()
	ishovered()                                                       	{	;-- liefert true wenn Maus über Listview und kein Eingabefeld fokussiert wurde
		global hQE, QEhLvExp
		If !WinActive("ahk_id " hQE)
			return false
		MouseGetPos, mx, my, hwin, hctrl, 2
	return (hwin=hQE && hctrl=QEhLvExp ? true : false)
	}

; --------------------------------------------------------------------------
; ‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾
	Hotkey, If                	, % isHovered_LV
	func_call := Func("QELVScrollHk").Bind("End", QEhLvExp)
	Hotkey, End               	, % func_call
	func_call := Func("QELVScrollHk").Bind("Home", QEhLvExp)
	Hotkey, Home              	,  % func_call
	#If


;}

	;AddMorePatientsToDB()
	; xDataMoveDirs("Patient")

	counting := 0
	CheckAdressTextFiles()

	global critObj := CriticalObject()
	critObj.1 := stdExportPath
	global threadRWF := AHKThread(A_ScriptDir "\resources\RemoveWrongFiles.ahk", (&critObj), true)

	QELetterChoiceTS()

	If !settings.SumatraCMD && !settings.DontShowSmtraMsg {

		Gui Msg: New	, % "hwndhMsg +Owner" hQE
		Gui Msg: Font	, % settings.fSize.s12, Calibri
		Gui MSg: Add	, Link     	, % "xm ym w600", % "Für das Drucken von Versandbriefen wird der SumatraPDF Reader benötigt.`nDas Programm muss nicht installiert werden."
															        					. " Allerdings wird das Skript es ohne weitere Konfiguration verwenden können.`n"
															        					. "Möchten Sie das Programm nicht installieren, verwenden Sie die portable Version.Im Anschluss erstellen Sie in der QExporter.ini einen "
															        					. "Eintrag unter [Script]/sumatrapdf_Pfad=[Pfad zur exe].`n`n"
															        					. "<a href=" q "https://www.sumatrapdfreader.org/download-free-pdf-viewer" q ">SumatraPDF Downloadseite</a>"
		Gui Msg: Add	, Checkbox	, % "y+10 vMSGChk "           	, % "Nicht wieder anzeigen"
		Gui Msg: Add	, Button  	, % "X+50 vMSGOk gMSGHandler" 	, % "Fenster schliessen"
		Gui MSg: Show	, AutoSize 	, % "PDF Reader benötigt"

	}

	GuiControl, QE: Focus, QELvExp


	;~ SciTEOutput(cJSON.Dump(settings, 1))


;}

return


;{ ⚞──────────────────      EXPORT KLASSEN       	──────────────────⚟

class DocAddressBook                                                         	{	;-- Gui, SQLite Datenbank, alles drum und dran

	__New() {

	; um das Gui zeichnen zu können
	 this.Ctrls:=	[ {"db":"ANR"       	, "dbIndex":"1" 	,  "txt":"Anrede"   	, "input":"CB"  	, "vName": "NDInpANR"      	, "opt":"w60            gNDHandler"  	, "default": "|Frau|Herr"}
								,	{"db":"RTitel"     	, "dbIndex":"2" 	,  "txt":"Titel"    	, "input":"CB"  	, "vName": "NDInpTitle"   	, "opt":"w120           gNDHandler"  	, "default": "Dr. med.|Prof. Dr. med.", "choose":1}
								,	{"db":"RSName"    	, "dbIndex":"3" 	,  "txt":"Name"     	, "input":"Edit"	, "vName": "NDInpSName"   	, "opt":"w180 AltSubmit gNDHandler"}
								,	{"db":"RPName"    	, "dbIndex":"4" 	,  "txt":"Vorname"  	, "input":"Edit" 	, "vName": "NDInpPName"   	, "opt":"w180 AltSubmit gNDHandler"}
								,	{"db":"KIM"        	, "dbIndex":"5" 	,  "txt":"Kim"      	, "input":"Edit"	, "vName": "NDInpKIM"      	, "opt":"w300 AltSubmit gNDHandler"}
								,	{"db":"LANR"      	, "dbIndex":"6" 	,  "txt":"LANR"      	, "input":"Edit"	, "vName": "NDInpLANR"     	, "opt":"w100 gNDHandler"}
								,	{"db":"BSNR"      	, "dbIndex":"7" 	,  "txt":"BSNR"     	, "input":"Edit"	, "vName": "NDInpBSNR"    	, "opt":"w100 gNDHandler"}
								,	{"db":"More1"   		, "dbIndex":"16"	,  "txt":"Facharzt" 	, "input":"CB"  	, "vName": "NDInpFach"    	, "opt":"w180           gNDHandler" 	, "default":"|Allgemeinmedizin|Innere Medizin|Chirurgie"}
								,	{"db":"Strasse"    	, "dbIndex":"8" 	,  "txt":"Straße"   	, "input":"CB"  	, "vName": "NDInpStreet"   	, "opt":"w300"	                     	, "default":""}
								,	{"db":"PLZ"        	, "dbIndex":"10"	,  "txt":"PLZ"      	, "input":"Edit" 	, "vName": "NDInpPLZ"      	, "opt":"w80"   	                   	, "default":""}
								,	{"db":"ORT"        	, "dbIndex":"9" 	,  "txt":"Ort"      	, "input":"Edit" 	, "vName": "NDInpORt"      	, "opt":"w180 gNDHandler"	          	, "default":""}
								,	{"db":"Tel1"	    	, "dbIndex":"11"	,  "txt":"Tel1"     	, "input":"Edit"	, "vName": "NDInpTel1"    	, "opt":"w110 gNDHandler"}
								,	{"db":"Tel2"	    	, "dbIndex":"12"	,  "txt":"Tel2"     	, "input":"Edit"	, "vName": "NDInpTel2"     	, "opt":"w110 gNDHandler"}
								,	{"db":"Fax1"	    	, "dbIndex":"13"	,  "txt":"Fax1"     	, "input":"Edit"	, "vName": "NDInpFax1" 	    , "opt":"w110 gNDHandler"}
								,	{"db":"Fax2"	    	, "dbIndex":"14"	,  "txt":"Fax2"     	, "input":"Edit"	, "vName": "NDInpFax2"     	, "opt":"w110 gNDHandler"}
								,	{"db":"EMail"     	, "dbIndex":"15"	,  "txt":"eMail"    	, "input":"Edit"	, "vName": "NDInpEMail"   	, "opt":"w160 AltSubmit gNDHandler"}
								,	{"db":"More2"   		, "dbIndex":"17"	,  "txt":"Notizen"  	, "input":"Edit" 	, "vName": "NDInpNotes"    	, "opt":"w300 AltSubmit gNDHandler" 	, "default":""}]

	; SQLite
		this.table := "(ANR, RTitel, RSName, RPName, KIM, LANR, BSNR, Strasse, Ort, PLZ, Tel1, Tel2, Fax1, Fax2, EMail, More1, More2)"

	; Referenziert für ein zweites Objekt welches Beziehung vom
		this.vIndex := Object()
		For cIndex, ctrl in this.Ctrls
			this.vIndex[ctrl.db] := cIndex

	; gLabel Func
		this.gLabels := Object()



	}

	__Delete() {


		Gui, ND: Submit, NoHide
		this.GetNDInput(true)
		Gui, ND: Destroy
		this := ""

	}


	AddressBookDialog() {

		global
		static firstCall := true

		If firstCall {
			firstCall := false
			ctrlvalues := this.LoadLastNDInput()
		}
		else {
			Gui, ND: Show, AutoSize
			return
		}

		Gui, ND: new, hwndhND

		this.guiHandle           	:= hND
		this.gLabels	           	:= Object()
		this.gLabels["NDHandler"] := Object()

		Gui, ND: Font, % settings.fSize.s10, Arial

		For cIndex, ctrl in this.Ctrls  {

			vName := ctrl.vName

		; Steuerelement hat gLabel
			RegExMatch(ctrl.opt, "i)\sg(?<Label>[\pL_\-]+)", g)
			opt := StrReplace(ctrl.opt, " g" gLabel)

		; Bezeichnung
			Gui, ND: Add, Text, % "xm y" (cIndex=1?"m":"+3") " w100"                                                                        	, % ctrl.txt

		; Eingabefeld
			Gui, ND: Add, % (ctrl.input="CB"?"ComboBox":ctrl.input), % "x+0 " . opt . " v" ctrl.vName " " (gLabel ? " hwndh" ctrl.vName : "") , % ctrl.default
			if (ctrlvalues && ctrl.value)
				GuiControl, % "ND: " (ctrl.input="CB" ? "ChooseString" : ""), % ctrl.vName, % ctrl.value
			else if (!ctrl.value && ctrl.input="CB" && ctrl.choose)
				GuiControl, % "ND: " (ctrl.input="CB" && ctrl.choose ~= "^\d+$" ? "Choose" : "ChooseString"), % ctrl.vName, % ctrl.choose

	  ; gLabel zum Steuerelement hinzufügen
			If gLabel {

				If !IsObject(this.gLabels[gLabel])
					this.gLabels[gLabel] := Object()

				chwnd   	:= "h" ctrl.vName
				hwnd    	:= %chwnd%
				on_gCall 	:= ObjBindMethod(this, gLabel, ctrl.vName)

				this.gLabels[gLabel][ctrl.vName] := {"hwnd":hwnd, "gLabel":on_gCall}
				GuiControl, ND: +g, % hwnd, % on_gCall

			}

		}

	; gLabel festlegen
		gLabel := "NDHandler"

	; Übernehmen und Abbrechen
		Gui, ND: Add, Button, % "xm y+20    	vNDTake 	hwndhNDTake"  	, % "Versandadresse übernehmen"
		on_gCall 	:= ObjBindMethod(this, gLabel, "NDTake")
		this.gLabels[gLabel]["NDTake"]  := {"hwnd": hNDTake  	, "gLabel": on_gCall}
		GuiControl, ND: +g, % hNDTake 	, % on_gCall

		Gui, ND: Add, Button, % "x+50   		vNDCancel 	hwndhNDCancel"	, % "Abbrechen"
		on_gCall 	:= ObjBindMethod(this, gLabel, "NDCancel")
		this.gLabels[gLabel]["NDCancel"]:= {"hwnd": hNDCancel	, "gLabel": on_gCall}
		GuiControl, ND: +g, % hNDCancel	, % on_gCall

	; Gui zeigen
		Gui, ND: Show, AutoSize, Hausärzte

  return

  ; Fenster wurde manuell geschlossen, Eingaben sichern
  NDGuiClose:
  NDGuiEscape:

		Gui, ND: Default
		this.GetNDInput(true)
		Gui, ND: Show, Hide

  return
	}

	NDHandler(param*) {

		Critical

		;~ SciTEOutput(param.1 ", " param.2 ", " param.3 ", " param.4 "; " A_EventInfo)

		if (param.1 = "NDInpANR") 		{
			If (param.3 = "Normal")
				GuiControl, ND: Focus, NDInpSName
		}
		else if (param.1 = "NDInpTitle") 	{
			If (param.3 = "Normal")
				GuiControl, ND: Focus, NDInpSName
		}

   ; Daten übernehmen und in DB speichern, sowie Empfängerauswahl erweitern
		else if (param.1 = "NDTake" ) 		{

			Gui, ND: Default

			this.GetNDInput(false)

			elRSName  	:= this.Ctrls[this.vIndex["RSName"]]
			elRPName  	:= this.Ctrls[this.vIndex["RPName"]]
			IF !elRPName.valUe
				msgN .= "Ohne Vornamen"
			If !elRSName.value
				msgN := msg ? "Ohne Vor- und Nachnamen" : "Ohne Nachnamen"
			If msgN
				msgN .= " können die Daten nicht übernommen werden."

		; BSNR prüfen
			elBSNR  	:= this.Ctrls[this.vIndex["BSNR"]]
			If (bsnr := Trim(elBSNR.value))
				item   	:= SQLiteRecordExist("Recipients", "BSNR", bsnr)
			If (!bsnr || !bsnr ~= "^\d+$" || StrLen(bsnr) < 9 || item.Exist) {

				bsnrflat := RegExReplace(bsnr, "[\d\s]")
				msgB := ( !bsnr            	? "Ohne eine BSNR ist die Zuweisung eines Arztes zu einem Datensatz nicht möglich. Bitte eine BSNR angeben!"
							 :  item.Exist       	? "Ein Arzt (" item.record[3] ", " items.record[4]  ") mit der gleichen BSNR existiert bereits."
							 :	bsnrflat         	? "Die BSNR enthält außer Zahlen weitere Zeichen [" RegExReplace(bsnr, "[\d\s]") "]. Bitte geben Sie nur Zahlen ein."
							 :	StrLen(bsnr) < 9	?	"Die eingegebene BSNR hat weniger als 9 Zeichen. Sie können die BSNR dennoch verwenden, da diese eindeutig ist. Wählen Sie 'Ignorieren' um die BSNR für diesen Arzt zu übernehmen."
							 :                    	"")

				If (bsnr && !bsnr ~= "^\d+$")
					GuiControl, ND:, % elBSNR.vName, % RegExReplace(bsnr, "[^\d]")
			}

		; LANR prüfen
			elLANR := this.Ctrls[this.vIndex["LANR"]]
			lanr := RegExReplace(elLANR.value, "[^\d]")
			If (!lanr && StrLen(lanr) < 9) {
				MsgBox, 0x1051, % scriptTitle, % "Die LANR " (lanr ? "[" lanr "]" : "") " wird für die Zuweisung des Arztes nicht benötigt.`n`n"
																			.	 "Mit 'Ja' wird stattdessen eine Pseudo-Nummer erstellt.`n"
																			.	 "Mit 'Nein' wird der eingetragene Wert übernommen.`n"
																			.	 "Mit 'Abbruch' kommen Sie zum Eingabedialog zurück.`n"

				IfMsgBox, Cancel
					return

				IfMsgBox, Yes
				{
					dlanr := "111111111"

					while (exists := SQLiteRecordExist("Recipients", "LANR", dLANR, "returnOnlyExist") && A_Index < 100) {
						dLANR += 1
					}

					If !exists {
						GuiControl, ND:, % elLANR.vName, % dLANR
						elLANR.value := dLANR
						Sleep 3000
					}

				}


			}

		; andere prüfen
			For cIndex, ctrl in this.Ctrls
				If !(ctrl.db ~= "i)(LANR|BSNR|RSNAME|RPName|RTitel)") {
					el := this.Ctrls[this.vIndex[ctrl.db]]
					If !Trim(el.value)
						msgM .= (msgM ? ", " : "") el.Text
				}


		 ; eine Info per Messagebox anzeigen
			If (msgN || msgB || msgM)	{

				MsgBoxhex := ( msgN || msgB          	? 0x40030    		; okay
						     		: !msgN && !msgB && msgM 	? 0x40031				;
							    	: StrLen(bsnr) < 9				? 0x40034				; OK/YES
							     	: 0x40030)																	; OK

				MsgBoxhex += 0

				SciTEOutput( msgN "`n" msgB "`n" msgM)

				MsgBox, % MsgBoxhex	, % "Hausärzte - ACHTUNG", % msgN "`n" msgB "`n" msgM "`n"
																							 . ( msgN || msgB 		  	  		? "Es kann Datensatz angelegt werden, da wichtige Daten der Praxis fehlen." "`n"
								                 							 .												   		"-----------------------------------------------------------------------" "`n "
								                 							 . " 1. " (msgN	            	? msgN : msgB) "`n" 		  (msgN && msgB 						? " "
								                 							 . " 2. " (msgN && msgB     	? msgB "`n" : "") : "")  ((msgN || msgB) && msgM		? ""
								                 							 . "Und es" : "Es") 																	 ((!msgN || !msgB) && msgM	? " "
								                 							 . "fehlen noch ein paar Angaben (" msgM ").`n" : "")  (!msgN && !msgB && msgM 	? " "
								                 							 . "fehlen noch ein paar Angaben (" msgM ").`nDie Daten sind allerdings eindeutig und können mit 'Ja' übernommen werden.`nDruck auf 'Nein' führt zum Eingabedialog zurück.")
																							 :	msgN "`n" msgB "`n" msgM)

				IfMsgBox, No
					return
				IfMsgBox, OK
				{
					If (MsgBoxhex != 0x1049)
						return
				}
				IfMsgBox, Cancel
					return

			}

			Gui, ND: Default

		; Daten in SQLITE DB Speichern
		; SQL Statement erzeugen
			_SQLCol 		:= ""
			_SQLValues 	:= ""
			For cIndex, ctrl in this.Ctrls {
				If (ctrl.db ~= "i)^(Tel|Fax)") {
					ctrl.value := RegExReplace(ctrl.value, "[^\d]")
					GuiControl, QE:, % ctrl.vName, % ctrl.value
				}
				_SQLCol   	.= (_SQLCol 	 ? ", " : "") ctrl.db
				_SQLValues	.= (_SQLValues ? ", " : "") q . ctrl.value . q
			}

			_SQL := "INSERT INTO Recipients (" . _SQLCol . ") VALUES (" _SQLValues ");"
			SciTEOutput(A_ThisFunc "() | " _SQL )

			If noSQLWrite
				return

		; schreibt in die Datenbank
			QEDB.Exec("BEGIN TRANSACTION;")

			If !QEDB.Exec(_SQL)
				SQLiteErrorWriter("Recipients Tabelle - Fehler bei INSERT INTO", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
			else {
			}

			QEDB.Exec("COMMIT TRANSACTION;")


			GuiControl, QE:, QERecipient, % (recipientname := SQLiteGetRecipient(elBSNR.value) " [" elBSNR.value "]")

			MsgBox, 0x1000, % "Hausärzte", % "Empfänger hinzugefügt:`n" recipientname

			For cIndex, ctrl in this.Ctrls
				If (ctrl.input="CB" && ctrl.choose)
					GuiControl, % "ND: " (ctrl.input="CB" && ctrl.choose ~= "^\d+$" ? "Choose" : "ChooseString"), % ctrl.vName, % ctrl.choose
				else
					GuiControl, % "ND: ", % ctrl.vName, % ""

			this.GetNDInput(true)
			Gui, ND: Show, Hide

		return recipientname
		}
		else if (param.1 = "NDCancel" ) 	{
			Gui, ND: Default
			this.GetNDInput(true)
			Gui, ND: Show, Hide
		}


	}

	GetNDInput(SaveInput:=false) {

		Gui, ND: Default
		Gui, ND: Submit, NoHide

		For cIndex, ctrl in this.Ctrls {
			ctrlN 	:= ctrl.vName
			If (ctrl.input = "db")
				value := GuiControlGet("ND", "", ctrlN)
			else
				value := %ctrlN%
			ctrl.value := value
			If (SaveInput && value)
				IniWrite, % value, % settings.paths.QEExporterINI, % "Adressbuchdialog", % ctrl.vName
		}


	}

	LoadLastNDInput() {

		ctrlvalues := 0
		workini := IniReadExt(settings.paths.QEExporterINI)
		For cIndex, ctrl in this.Ctrls {
				value := IniReadExt("Adressbuchdialog", ctrl.vName)
				If (value && value != "ERROR") {
					ctrl.value := value
					If !(ctrl.vName ~= "i)(NDInpANR|NDInpTitle)")
						ctrlvalues += 1
				}
		}

	return ctrlvalues
	}

}

class Quickexporter                                                          	{	;-- Datenexportmethoden

  ; HTML Tagesprotokollexporter

  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	__New(basepath, AlbisPath, opt :="")               	{

		this.AlbisPath	     	:= AlbisPath
		this.AlbisBriefe   	:= AlbisPath "\Briefe"
		this.AlbisDBPath	:= AlbisPath "\db"

		cJSON.EscapeUnicode := "UTF-8"

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Optionen parsen
		this.reNewData := true
		RegExMatch(opt, "i)employeeshorts\s*=\s*(?<shorts>[\pL\-\_\|\*]+?)([\s,]|$)", opt)

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Basispfad prüfen
		If !FilePathCreate(this.basepath := basepath )
			throw "Konnte Pfad zum Schreiben der Daten nicht anlegen.`n" basepath

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Statistikendatei
		this.StatsIniFilePath := this.basePath "\QExporterStats.ini"

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; xData Path
		If !FilePathCreate(this.xDataPath := this.basepath "\xData" )
			throw "Konnte Pfad zum Schreiben der Daten nicht anlegen.`n" this.xDataPath

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; für HTML Ausgabe
		If !FileExist(A_ScriptDir "\resources\Karteikarte.html") {
			MsgBox, 0x1004, % scriptTitle, % "Es fehlt die CSS und HTML-Datei im \resources Verzeichnis.`n"
													.	 "Ein Export von Patientendaten ist ohne diese Datei nicht möglich`n"
													.	 "Das Skript nach drücken von 'Ok' beendet."
			ExitApp
		}
		this.html := FileOpen(A_ScriptDir "\resources\Karteikarte.html", "r", "UTF-8").Read()

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Dokument und Mediafilter
		this.mediatypes     	:= "(jpg|png|bmp|gif|wav|mpeg|mp4|mp3|pdf|doc|avi|docx)"
		this.externmedia    	:= "(jpg|png|bmp|gif|pdf)"
		this.internmedia    	:= "(doc|docx)"

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RegEx Filter um Karteikartentext vonz.B. Arbeitsanweisungen, wiederholten Begriffen oder was auich immer zu befreien
		this.employeeshorts	:= optshorts
		this.rxFRplStd         	:= "i)[\,\-\s]*\b(" this.employeeshorts ")\b[\,\-\s]*"
		this.rxHERpl           	:= "([THE\\]+\d+|\\+HK\d+.*\\H$)"
		this.rxfAU	            	:= "([THE\\]+\d+|\\+HK\d+.*\\H$|;\s\{KGZ.*$)"
		this.rxInfoReiniger  	:= [{"rxm" : "Rechnung bar bezahlt", "rpl":""}
											, {"rxm" : this.rxFRplStd, "rpl":""}
											, {"rxm" : "[\<\>]", "rpl":""}
											, {"rxm" : "^\pL{2,3}$", "rpl":""}
											, {"rxm"	: "(\p{Ll}{2,})(\p{Lu}\H)", "rpl":"$1 $2"}
											, {"rxm" : "i)\b(yGT|AP|ALAT|ASAT|grBB|IgA|IgG|IgM|CRP|NA|K|Krea|HST|Harnsäure|LDL|HDL|Triglyzeride|"
														  . "C1\-Esterase|GFR|Lipase|Amylase|BSG|TSH|fT3|fT4|Stuhl\s*auf\s*E\+R|Stuhl\s*auf\s*Helicobacter|"
														  . "BA\s*gesamte\s*Blutwerte\s*=\s*|Glucose|Bili ges.|Cholesterin|Kreatinin|Harnstoff|Gesamt\s*Eiweiß|"
														  . "Transferrin|Ferritin|Kalium|Natrium|HbA1c|gr.Blutbild|VitD25|Vitb12|Folsäure|D\-Dimer|INR"
														  . "Parathormon|Ostase|Ca|Phosphat|Calcium|VitD|Eiweiß\s*Elektrophorese)\b[\s,;]*", "rpl":""}] ; , {"rmx":"", "rpl":""}

		/* RegEx Filter für die Ausgabe
			    			;~ , 	"AU\-Schein" 	    	: [{"rmx" : "(\sT\d+[HTE\d\\\s]+\s)|(\\H[TE]\d+)|(\\HK\d+.+\\H(\s|$))|(\\H[;\s]*$)"             	, "rpl" : "; "}
			    			;~ , 	"Überweisung"	    	: [{"rxm"	: "((^|\s)T\d+[HTE\d\\\s]+\s)|(\\H[TE]\d+)|(\\HK\d+.+\\H(\s|$))|(\\HK\d+.+\\HV[\d\.]+(\s|$))|(\\H[;\s]*$)"                	, "rpl" : ";"}
		*/
		this.rxFilter := { "Anamnese"	         	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}
																 , {"rxm" 	: "(\p{Ll}{2,})(\p{Lu}\H)"                                                                                    	, "rpl" : "$1 $2"}]
			    			, 	"AU-Schein" 	        	: [{"rxm" 	: "\\*H*T\d+(\\H[TE]\d+\s*)+"                                                                             	, "rpl" : "; "}
														    	 , {"rxm"	: "\\H[TE]\d+"                                                                                                     	, "rpl" : "; "}
														    	 , {"rxm"	: "\\HK\d+.+\\H(\s|$)"                                                                                        	, "rpl" : "; "}
														    	 , {"rxm"	: "\\H[;\s]*$"                                                                                                      	, "rpl" : ""}
														    	 , {"rxm"	: "\{KGZ\svom\s[\d\.]+"                                                                                     	, "rpl" : ""}
														    	 , {"rxm"	: "(\s*;\s*)+"                                                                                                         	, "rpl" : "; "}
														    	 , {"rxm"	: "(^[\s;]+|[\s;]+$)"                                                                                            	, "rpl" : ""}]
			    			, 	"Befund"       		    	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}
																 , {"rxm" 	: "(\p{Ll}{2,})(\p{Lu}\H)"                                                                                    	, "rpl" : "$1 $2"}
																 , {"rxm" 	: "[\<\>]"                                                                                                             	, "rpl":""}]
			    			, 	"Bemerkung"		    	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}
																,  {"rxm" 	: "(\p{Ll}{2,})(\p{Lu}\H)"                                                                                    	, "rpl" : "$1 $2"}]
			    			, 	"Belastungsgrenze"  	: [{"rxm" 	: "(\s*E\d+[HTE\d\\\s]+|\\H[TE]\d+|\\HK\d+.+\\H(\s|$)|\\H;*)"                        	, "rpl" : "; "}
																 , {"rxm" 	: "(\p{Ll}{2,})(\p{Lu}\H)"                                                                                    	, "rpl" : "$1 $2"}
														    	,  {"rxm"	: "[;\s]+"                                                                                                            	, "rpl" : "; "}
														    	,  {"rxm"	: "(^[\s;]+|[\s;]+$)"                                                                                            	, "rpl" : ""}]
			    			, 	"Bericht"			    	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}]
			    			, 	"Bild"					    	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}]
			    			, 	"Blutdruck"	        	: [{"rxm"	: "-(\d+)\s*{\\\/}\s*(\d+)\s*({mmHg}*)"                                                              	, "rpl" : "$1/$2 mmHg "}]
			    			, 	"Diagnose"		    	: []
			    			, 	"Eigenanamnese"   	: [{"rxm" 	: "(\p{Ll}{2,})(\p{Lu}\H)"                                                                                    	, "rpl" : "$1 $2"}]
			    			, 	"Einweisung"		     	: [{"rxm" 	: "(T1\d+\\H.\d+|\\HK72.*?\\H\s*$|\\H\p{Lu}\d+)"                                        	, "rpl" : ";"}]
			    			, 	"HKPfl-Verord."      	: [{"rxm" 	: "(\s*T\d+[HTE\d\\\s]+|\\H[TE]\d+|\\HK\d+.+\\H(\s|$)|\\H;*)"                        	, "rpl" : "; "}
														    	 , {"rxm"	: "[;\s]+"                                                                                                            	, "rpl" : "; "}
														    	 , {"rxm"	: "(^[\s;]+|[\s;]+$)"                                                                                            	, "rpl" : ""}]
			    			, 	"Heilmittelverord."  	: [{"rxm"	: this.rxHERpl	                                                                                                  	, "rpl" : ""}]
			    			, 	"Information"		    	: this.rxInfoReiniger
			    			, 	"Kind-AU"	     	    	: [{"rxm" 	: "(^\s*T\d+[HTE\d\\\s]+\s|\\H[TE]\d+|\\HK\d+.+\\H(\s|$)|\\H;*)"                  	, "rpl" : "; "}
														    	 , {"rxm"	: "[;\s]+"                                                                                                            	, "rpl" : "; "}
														    	 , {"rxm"	: "(^[\s;]+|[\s;]+$)"                                                                                            	, "rpl" : ""}]
			    			, 	"Laborbefund"	    	: []
			    			, 	"Medhinweise"  	    	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}]
			    			, 	"Therapie"		         	: [{"rxm"	: this.rxFRplStd	                                                                                                  	, "rpl" : ""}]
			    			, 	"Überweisung"	    	: [{"rxm" 	: "T\d+(\\H[TE]\d+)+\s"                                                                                     	, "rpl" : "; "}
														    	 , {"rxm"	: "\\H[TE]\d+"                                                                                                     	, "rpl" : "; "}
														    	 , {"rxm"	: "\\HK\d+.+\\H(\s|$)"                                                                                        	, "rpl" : "; "}
														    	 , {"rxm"	: "\\H[;\s]*$"                                                                                                      	, "rpl" : ""}
																 , {"rxm"	: "(\s*;\s*)+"                                                                                                         	, "rpl" : "; "}
														    	 , {"rxm"	: "(^[\s;]+|[\s;]+$)"                                                                                            	, "rpl" : ""}]}

	  ; Kürzel und Alias-Zuordnung (Aliasse sind die Bezeichnungen in der HTML/PDF Ausgabe der Karteikarte)
		this.rxKZL := {"anam" 	: {"alias"	: "Anamnese"}
			    			, "aem"  	: {"alias"	: "Medhinweise"}
			    			, "au" 		: {"alias"	: "AU-Schein"}
			    			, "avi" 		: {"alias"	: "Film"}
			    			, "bef" 		: {"alias"	: "Untersuchung"}
			    			, "bem"  	: {"alias"	: "Bemerkung"}
			    			, "besch"   	: {"alias"	: "Bescheinigung"}
			    			, "BSG"  	: {"alias"	: "Laborbefund"}
			    			, "BZ"    	: {"alias"	: "Blutzucker"}
			    			, "bild" 		: {"alias"	: "Bild"}
			    			, "brief"		: {"alias"	: "Brief"}
			    			, "CovGe"	: {"alias"	: "CovGeZert"}
			    			, "CovZe"	: {"alias"	: "CovImpfZert"}
			    			, "dia" 		: {"alias"	: "Diagnose"}
			    			, "eAu" 		: {"alias"	: "AU-Schein"}
			    			, "eBrief"	: {"alias"	: "eArztbrief"}
			    			, "epi"    	: {"alias"	: "Bericht"}
			    			, "erp"    	: {"alias"	: "eRezept"}
			    			, "f0051"	: {"alias"	: "Befundbericht"}
			    			, "f1204"	: {"alias"	: "Befundbericht"}
			    			, "f2260"	: {"alias"	: "Kurzbericht PT"}
			    			, "f2402"	: {"alias"	: "VO Ergoth."}
			    			, "f2902"	: {"alias"	: "Überw. D-Arzt"}
			    			, "f3110"	: {"alias"	: "Belast.erprobung"}
			    			, "f0052"	: {"alias"	: "Zusatzbogen"}
			    			, "fau" 		: {"alias"	: "AU-Schein"}
			    			, "faküf" 	: {"alias"	: "VO Reha-Sport"}
			    			, "fakür"	: {"alias"	: "VO Reha-Sport"}
			    			, "faza"  	: {"alias"	: "Muster 53"}
			    			, "fbbck"	: {"alias"	: "Belastungsgrenze"}
			    			, "fbga" 	: {"alias"	: "BG Unfallmeldung"}
			    			, "fddia"	: {"alias"	: "Diagnose"}
			    			, "fkind"		: {"alias"	: "Kind-AU"}
			    			, "fkonp"	: {"alias"	: "Konsiliarbericht PT"}
			    			, "fpau"		: {"alias"	: "AU-Schein"}
			    			, "füb" 		: {"alias"	: "Überweisung"}
			    			, "fübg"		: {"alias"	: "Überw. D-Arzt"}
			    			, "fhp" 		: {"alias"	: "HKPFL-Verord."}
			    			, "fkh" 		: {"alias"	: "Einweisung"}
			    			, "fhv13" 	: {"alias"	: "Heilmittelrezept"}
			    			, "flogo" 	: {"alias"	: "VO Logopädie"}
			    			;~ , "frp"    	: {"alias"	: "Einnahmeverordnung"}
			    			, "frpgr"   	: {"alias"	: "grünes Rezept"}
			    			, "frppn"   	: {"alias"	: "Privatezept"}
			    			, "fvap"   	: {"alias"	: "VO SAPV"}}
			  tmp  := { "fvmv"   	: {"alias"	: "VO med.Vorsorge"}
			    			, "fvreh"   	: {"alias"	: "VO REHA"}
			    			, "fvwied"  	: {"alias"	: "Wiederein<br>gliederungsplan"}
			    			, "info" 		: {"alias"	: "Information"}
			    			, "labor" 	: {"alias"	: "Laborbefund"}
			    			, "medbg" 	: {"alias"	: "BG-Rezept"}
			    			, "medrp" 	: {"alias"	: "Kassenrezept"}
			    			, "medh"	: {"alias"	: "Heilmittelrezept"}
			    			, "medhm"	: {"alias"	: "Hilfsmittelrezept"}
			    			, "medp"  	: {"alias"	: "Privatrezept"}
			    			, "medpn" 	: {"alias"	: "Privatrezept"}
			    			, "medbm" : {"alias"	: "BtM-Rezept"}
			    			, "Micro"  	: {"alias"	: "Microalbuminurie"}
			    			, "scan" 	: {"alias"	: "Befund"}
			    			, "sono" 	: {"alias"	: "Sono-Befund"}
			    			, "ub" 		: {"alias"	: "Überweisung"}
			    			, "ther"   	: {"alias"	: "Therapie"}
			    			, "R R"    	: {"alias"	: "Blutdruck"}
			    			, "rö"     	: {"alias"	: "Ro-Befund"}
							, "taug"   	: {"alias"	: "Bescheinigung"}
							, "ton"   	: {"alias"	: "Tonaufnahme"}
							, "ub"    	: {"alias"	: "überweisung"}
							, "Urin"   	: {"alias"	: "Urin"}
			    			, "vopln"   	: {"alias"	: "Med.plan"}
							, "z"			: {"alias"	: "Eigenanamnese"}}
		for kzl, aliasKzl in tmp
			this.rxKZL[kzl] := aliasKzl

	  ; Zuordnung Rezeptdaten
		this.rzDPos := {	"151":"#rD1_1", "152":"#rD1_2", "153":"#rD1_3", "152":"#rD1_4"
			        			,	"155":"#rD2_1", "156":"#rD2_2", "157":"#rD2_3", "158":"#rD2_4"
			        			,	"163":"#rD3_1", "164":"#rD3_2", "165":"#rD3_3", "166":"#rD3_4"
			        			,	"167":"#rD4_1", "168":"#rD4_2", "169":"#rD4_3", "170":"#rD4_4"
			        			,	"171":"#rD5_1", "172":"#rD5_2", "173":"#rD5_3", "174":"#rD5_4"
			        			,	"177":"#rD6_1", "178":"#rD6_2", "179":"#rD6_3", "180":"#rD6_4" }

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RegEx Standardreplacer läuft immer bei bestimmten Kürzeln
		this.rxStandardsRunOn := "(anam|bef|info|ther|medpn|medrp|medbm|aem|ther)"
		this.rxStandards := {  "(\s)\d{0,1}\\H[KAET]\d+"	: "$1 "
					        			, "^\s*\|\s*"     	             	: ""
					        			, "\s*\d*[\\\/]*\d*\s*"	         	: ""
					        			, "\s*\d*\s"                         	: " "
					        			, "[\s\d]+"                       		: " "
					        			, "(\d+Bild\\.*?\.[a-z]+)"    	: " [$1]"}


	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	BulkExport(lvRows)                                 	{

		global fn_AlbisCheck

		If !IsObject(lvRows)
			return

		this.ExportRunning := true

		filepaths := []
		this.pat := Object()
		this.TimeSum := this.TSummedFiles := 0
		this.convertProblems := 0

		Gui, QE: Default
		GuiControl, QE: Disable, QEExplorer
		GuiControl, QE: Disable, QESearch
		Gui, QE: ListView, QELvExp

	 ; Anlegen von Dateiordnern, vorbereiten um gleichzeitig Daten zu speichern
		IDs := {}
		this.startTime := A_TickCount
		For index, element in lvRows {

			rowNr      	:= element.row
			rowData	:= element.data
			row        	:= element.data.names
			PatID    	:= row.PatID

			indexshow := "[" index "/" lvRows.Count() "] " cPat.NAME(PatID, false)
			Zeigsmir("bereite Daten vor" , index, lvRows.Count(),0,0, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))

			If !IsObject(this.pat[PatID])
				this.pat[PatID] := Object()

		; weiter damit man nicht unpassende Daten extrahiert
			If !this.CheckPatID(rowNr, rowData)
				continue

		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		; Karteikartenexportpfad - Anlegen auch der Unterverzeichnisse, Textdatei mit Adressdaten des Arztes
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{

		  ; Basispfad  - orientiert sich am Empfänger
			this.pat[PatID].basePath := this.basepath "\" (row.Recipient < 3 ? "Patient" : row.RecipientName " [" row.Recipient "]")
			If !FilePathCreate(this.pat[PatID].basePath)
				throw "Konnte sub Pfad zum Schreiben der Daten nicht anlegen.`n" this.pat[PatID].basePath
			If !FilePathCreate(this.pat[PatID].basePath "\versendet")
				throw "Konnte 'versendet' Pfad nicht anlegen.`n" this.pat[PatID].basePath "\versendet"
			If (row.Recipient > 3) {
				adressfilepath := this.pat[PatID].basePath "\Adresskopf " RegExReplace(row.RecipientName, "i)(Herr\s*|Frau\s*|Hr\.\s*|Fr\.\s*|Dr\.\s*|med\.\s*)") " [" row.Recipient "]"
				adressfilepath := RegExReplace(adressfilepath, "\s{2,}", " ")
				If !FileExist(adressfilepath)
					WriteAdressTextFile(adressfilepath, row.Recipient)
			}

		 ; Patientenpfad - Unterverzeichnis im Basispfad
			shipped := row.Shipment ~= "(\d{2}\.\d{2}\.\d{4}|\d{8})" ? "versendet" : ""
			SciTEOutput(A_ThisFunc "() | shipped: " shipped ", " row.Shipment)
			this.pat[PatID].basePatPath 	:= this.pat[PatID].basePath . (shipped ? "\" shipped : "") . "\(" PatID ") " (PatName := row.SName ", " row.PName)
			this.pat[PatID].basename     	:= "[" PatID "] " PatName " Tagesprotokoll ab " (row.SDate ~= "^\s*\d+\s*$" ? ConvertDBASEDate(Trim(row.SDate)) : row.SDate)
			this.pat[PatID].xDataPath		:= this.xDataPath "\(" PatID ") " PatName
			this.pat[PatID].jsonfile       	:= this.pat[PatID].xDataPath "\" this.pat[PatID].basename ".json"
			this.pat[PatID].jsonfile2     	:= this.pat[PatID].xDataPath "\" this.pat[PatID].basename " - obj KKarte.json"
			;}

		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		; Patientenname und Adresse vorbereiten
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			PData := cPat.GetAdditionalData(PatID, ["NAME", "VORNAME", "GEBURT", "STRASSE", "HAUSNUMMER", "PLZ", "ORT"], 1)
			this.pat[PatID].ExportPatName 	:= cPat.NAME(PatID, true)
			this.pat[PatID].ExportPatAddress	:= PData.1.STRASSE (PData.1.HAUSNUMMER ? " " PData.1.HAUSNUMMER : "") ", " PData.1.PLZ " " PData.1.ORT
		;}

		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  -
		; Legt einen Dateispeicherpfad für den Karteikartenexport an
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			If !FilePathCreate(this.pat[PatID].basePatPath)
				throw "Konnte Pfad zum Schreiben der Daten nicht anlegen.`n" this.pat[PatID].basePatPath
			If !FilePathCreate(This.pat[PatID].xDataPath)
				throw "Konnte Pfad zum Schreiben der Daten nicht anlegen.`n" This.pat[PatID].xDataPath
		;}

		; PatIDs sammeln, Tabellendaten verknüpfen
			this.pat[PatID].index 	:= index
			this.pat[PatID].rowNR	:= rowNr
			this.pat[PatID].row 	:= row

		}

		obj := []
		obj.Push({"masterDB"      	: "BEFUND.dbf"			, "StoreDB"    	:"BEFTEXT.dbf"
					,  "masterKey"    	: "TextDB"					, "StoreKey"    	:"LFDNR"
					, "masterFillKey"  	: "INHALT"					, "StoredData"	:"Text"
																				, "StoreSubKey"	: "POS"})
		;~ this.DataAggregation(obj)

		this.TimeSum := (A_TickCount-this.startTime)

	  ; Albis überwachen
		if settings.CheckAlbis {
			fn_AlbisCheck := Func("CheckAlbisProblems")
			SetTimer, % fn_AlbisCheck, -500
		}

	  ; Exportloop
		For index, element in lvRows {

			this.startTime := A_TickCount

			rowNr    	:= element.row
			rowData 	:= element.data
			row        	:= element.data.names
			PatID     	:= row.PatID

			indexshow := "[" index "/" lvRows.Count() "] " cPat.NAME(PatID, false)
			Zeigsmir("exportiere Daten",0,0,0,0, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))

			Gui, QE: Default
			Gui, QE: ListView, QELvExp

			LV_GetText(notes, rowNr, 10)
			notes := RegExReplace(notes, "\[.+\]\s*", "")
			LV_Modify(rowNr, "col10", notes)

		  ; weiter damit man nicht unpassende Daten extrahiert
			If !this.CheckPatID(rowNr, rowData)
				continue

		  ; beim ersten Aufruf der nächsten Methode werden alle anderen PatID's mit übertragen

		  ; die Methoden für die Datenextraktion
			this.pat[PatID].PatPath := this.export(row, indexshow)
			WinActivate, % "ahk_id " hQE

		  ; Verzeichnis erhalten, dann Exportverzeichnis in Datenbank speichen und in der Tabelle anzeigen
			If InStr(FileExist(this.pat[PatID].PatPath), "D") {

				Gui, QE: Default
				Gui, QE: ListView, QELvExp

				If rowNr {
					LV_Modify(rowNr, "-Check Vis")
					LV_Modify(rowNr, "Col11", StrReplace(this.pat[PatID].PatPath, this.basePath "\"))
					filepaths.Push({"row":rowNr, "fpath":this.pat[PatID].PatPath})
				}

				If !SQLUpdateItem("Exports", "ExportPath", PatID, this.pat[PatID].PatPath) {
					MsgBox, % (index < lvRows.Count() ? 0x1004: 0x1000)
								, % scriptTitle " - SQLite3"
								, % "Es gab ein Problem beim Schreiben in die Datenbank.`nSiehe Protokolldatei. "
								. 	(index < lvRows.Count() ? "`nDennoch mit nächster Datei fortfahren?" : "" )
					IfMsgBox, No
						breakExport := true
					IfMsgBox, Cancel
						breakExport := true
				}

			}
			else {
				MsgBox, % (index < lvRows.Count() ? 0x1004: 0x1000)
							, % scriptTitle
							, % "Es gab ein Problem beim Exportieren nach:`n" this.pat[PatID].PatPath
							. 	(index < lvRows.Count() ? "`nDennoch mit nächster Datei fortfahren?" : "" )
					IfMsgBox, No
						breakExport := true
					IfMsgBox, Cancel
						breakExport := true

			}

			duration := A_TickCount - this.startTime
			If (duration/1000 >= 10) {
				this.TimeSum := duration
				this.TSummedFiles ++
			}

			If breakExport || this.Stop
				break

		}

		If this.LabExportRunning {
			ToolTip, % "Labor wird noch exportiert",,, 5
			while this.LabExportRunning
				Sleep 20
			ToolTip,,,, 5
		}

	  ; Anzeige entsprechend den Einstellungen auffrischen
		QEShowTable()

		GuiControl, QE: Enable, QEExplorer
		GuiControl, QE: Enable, QESearch

		If this.TimeSum {
			timeStrings := GetTimestrings(this.TimeSum)
			this.bulkDur := timeStrings.min "min " timeStrings.sec "s " timeStrings.msec "ms"
			ZeigsMir("beendet", 0, 0, 0, 0, 0, TimeFormatEx(Round(this.TimeSum/1000), true))
			t := this.TSummedFiles " Patientenakten in " this.bulkDur ".  ~" Round(this.TimeSum/this.TSummedFiles/1000) "s pro Patient`n[" Round(this.TimeSum/1000) "s -> " this.TimeSum "ms" "]"
			;~ SciTEOutput(t)
			FileAppend, % sqlTimeStamp() "|**Zeitprotokoll**|" t "`n", % settings.paths.SQLProtPath "\resources\QE_SQL-Log.txt"
		}

		SetTitleMatchMode, RegEx
		WinMinimize, % "i)\[InPrivate\].*?Microsoft​\s+Edge ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe"

		if settings.CheckAlbis && IsObject(fn_AlbisCheck)
			SetTimer, % fn_AlbisCheck, Off
		this.ExportRunning := false

		If this.convertProblems {
			MsgBox, 0x1004, % scriptTitle, % "Es liegen " this.convertProblems " Hinweise zu möglichen Konvertierungsproblemen vor.`n"
														. "Möchten Sie die Datei ansehen?"
			IfMsgBox, Yes
				Run, % A_ScriptDir "\resources\QEHinweise.txt"
		}

		ZeigsMir("selfcall=off")
		ZeigsMir("STOP")

	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	Export(row, indexshow:="") 	                       	{

		PatID := row.PatID
		PatName := row.SName ", " row.PName

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Karteikarteneinträge einsammeln
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			If (settings.reNewData || !FileExist(this.pat[PatID].jsonfile)) {
				aDB	:= new AlbisDB(Addendum.AlbisDBPath, "callback=ZeigsMir debug=0")
				KK 	:= aDB.Karteikarte(PatID,,, A_YYYY A_MM A_DD, "savepath=" this.pat[PatID].jsonfile " ,reNewData=" settings.reNewData)  ; row.SDate
				aDB 	:= ""
				FileOpen(this.pat[PatID].jsonfile, "w", "UTF-8").Write(cJSON.Dump(KK, 1))
			}
			else {
				If !FileExist(this.pat[PatID].jsonfile2)
					KK := cJSON.Load(FileOpen(this.pat[PatID].jsonfile, "r", "UTF-8").Read())
			}
		;}

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Matchobjekt für die Ausgabe filtern - > KK gefiltert nach KKarte
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		FileGetSize, fHTML, % (this.pat[PatID].htmlpath := this.pat[PatID].basePatPath "\Karteikarte.html")
		 If (FileExist(this.pat[PatID].jsonfile2) && !settings.reNewData && !FileExist(this.pat[PatID].htmlpath) && !fHTML) {                    	; laden wenn bereits erstellt und noch keine HTML Datei erstellt wurde
			KKarte := cJSON.Load(FileOpen(this.pat[PatID].jsonfile2, "r", "UTF-8").Read())
		 }
		 else {

			KKarte := Object(), filescopied := filescopiedLast := 0
			For idx, item in KK {

					If (Mod(idx, 10)=0) || (filescopiedLast != filescopied) {
						ZeigsMir("Dokumentexport", idx, KKarte.Count(), filescopied,, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))
						filescopiedLast := filescopied
					}

				  ; kein Inhalt, kein Kürzel oder eigentlich entfernter Datensatz
					If (!item.inhalt || !item.KUERZEL  || item.removed)
						continue

				; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				; Dokumente: exportiert sofort BIlder oder PDF Dateien in den eigene Unterverzeichnisse
				; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					Nitem := ""
					docPath := docName := docExt := ""
					If (item.KUERZEL ~= "i)\b(scan|besch|brief|epi|bild|avi|ton|vopln|taug)\b") {

						RegExMatch(item.Inhalt, "i)(?<Path1>\d{3}Bild\\.*?\.\pL{3,4})|(?<Path2>\d{3}\\.*?\.\pL{3,4})$", doc)

						docName := item.inhaltN ? item.inhaltN : item.InhaltO ? item.InhaltO : ""
						docName := RegExReplace(docName, "[\n\r]+$")
						docPath := docPath1 ? docPath1 : docPath2 ? docPath2 : ""
						docPath := RegExReplace(docPath, "\\+", "\")

					  ; kein Verzeichnis, dann weiter
						If !docPath
							continue

						docExt		:= RegExMatch(docPath, "\." this.externmedia "$") ? true : false
						docName := RegExReplace(Trim(docName), "i)\b(" this.employeeshorts ")\b[\s\-]*$", "")
						docName := docName ? docName : A_Hour  A_Min  A_Sec "" A_MSec
						docName := xstring.KKTextToFilename(docName)
						;~ docName := xstring.MSFilename(docName)

						If FileExist(srcFilePath := this.AlbisBriefe "\" docPath) {

							SplitPath, srcFilePath, mfName, mfDir, mfExt

						  ; jetzt kopieren
							filescopied ++
							If !FileExist(destFilePath := this.pat[PAtID].basePatPath "\" (destFileName := item.DATUM " - " docName "." mfExt))
								FileCopy, % srcFilePath, % destFilePath, 1

						}

						Nitem := Object()
						Nitem.Kuerzel	:= item.KUERZEL
						Nitem.alias      	:= this.rxKZL[item.KUERZEL].alias
						Nitem.Inhalt 		:= docName " (" mfExt ")"
						Nitem.Ext     		:= mfExt
						Nitem.Link   		:= destFileName
						Nitem.docExt 	:= docExt
						Nitem.DATUM 	:= item.DATUM

					}

				; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				; KÜRZEL & INHALT: Text ersetzen um Arbeitsanweisungen oder anderes herauszunehmen
				; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					If !isObject(Nitem)  {						;(!docPath || !docName)

						KZLfound := false
						For rxKUERZEL, str in this.rxKZL
							If (item.KUERZEL ~= "i)\b" rxKUERZEL "\b") {
							  ; Kürzel mit Alias-Bezeichnung ersetzten
								item.alias := str.alias
								KZLfound := true
								break
							}

					  ; nicht gelistete Kürzel werden nicht übernommen
						If !KZLfound
							continue

					}

				; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				; Daten zum Objekt hinzufügen
					Datum := item.DATUM

				  ; Objekt erweitern
					If (!IsObject(KKarte[DATUM]) && (isObject(Nitem) || StrLen(item.Inhalt)>0))
						KKarte[DATUM] := Object()

				   ; Dokumentlinks werden gesondert abgearbeitet
					If isObject(Nitem)  {

					  ; Nitem.Alias und document Array
						If !IsObject(KKarte[DATUM][(Nalias :=Nitem.alias)])
							KKarte[DATUM][Nalias] := Object()
						If !IsObject(KKarte[DATUM][Nalias].document)
							KKarte[DATUM][Nalias].document := Array()

						KKarte[DATUM][Nalias].document.Push({	"link"  	: Nitem.Link
																					, 	"txt"   	: Nitem.inhalt
																					, 	"docExt"	: Nitem.docExt
																					, 	"Ext"   	: Nitem.Ext
																					,	"Kuerzel": Nitem.KUERZEL
																					, 	"alias"	: Nitem.alias})

					}
				  ; im Falle eines Dokumentes als Array sammeln
					else if (StrLen(item.Inhalt)>0) {

						Lalias := item.alias
						If !IsObject(KKarte[DATUM][Lalias])
							KKarte[DATUM][Lalias] := Object()

						KKarte[DATUM][Lalias].KUERZEL := item.KUERZEL
						lenInh := StrLen(KKarte[DATUM][Lalias].inhalt)
						If !InStr(KKarte[DATUM][Lalias].inhalt, item.inhalt)
							KKarte[DATUM][Lalias].inhalt .= (KKarte[DATUM][Lalias].inhalt ? " " : "") item.inhalt

					}

					Nitem := ""

				}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - shorts: pm|MN|sc|cle|cl|KB|YR|Nk|jz|nk
			; Bereinigungen 	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀	🛀
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			; KKarte - Entfernen überflüssiger Daten oder lesbare Darstellung von Rezeptdaten
				For KKDate, KKAliases in KKarte {

					For KKAlias, item in KKAliases {

						If (KKAlias = "Information") {
							item.inhalt := RegExReplace(item.inhalt,"^\pL{2,3}$")
							item.inhaltN := RegExReplace(item.inhaltN,"^\pL{2,3}$")
							item.inhaltO := RegExReplace(item.inhaltO,"^\pL{2,3}$")
						}

						if item.Inhalt {

							item.InhaltO := lastInh := item.Inhalt
							item.Inhalt := Trim(item.Inhalt)

					  ; Rezepttexte durch RegExMatch filtern
					  ; - - - - - - - - - - - - - - - - - - - - - - - - -
						If (KKAlias ~= "i)rezept") {

							rezept := StrSplit(item.inhalt, "\")
							If (rezept.Count() > 0) {
								rezText := ""
								For idx, rezLine in rezept {

									rezDgPos := rezDsPos := rezDrug := rezDose := ""
									RegExMatch(rezLine, "HE32(?<DgPos>\d)\s(?<Drug>.+)"  	, rez)
									RegExMatch(rezLine, "HE10(?<DsPos>\d+)\s(?<Dose>.+)"	, rez)
									If (StrLen(rezDrug := Trim(rezDrug)) > 0) {
										rezDgPos -= 6
										rezText .= (rezText ? "<br>" : "") rezDrug " (#rD" rezDgPos "_1-#rD" rezDgPos "_2-#rD" rezDgPos "_3-#rD" rezDgPos "_4)"
									}
									else if (rezDsPos ~= "\d{3}") {
										rezDose := StrLen(Trim(rezDose)) > 0 ? Trim(rezDose) : 0
										If this.rzDPos.haskey(rezDsPos)
											rezText := RegExReplace(rezText, this.rzDPos[rezDsPos], rezDose)
									}

								}

								rezText      	:= RegExReplace(rezText, "#rD\d_\d"      	, "0")
								item.inhalt	:= RegExReplace(rezText, "\(0\-0\-0\-0\)"	, "")
							}

						}

					  ; Entfernen von überflüssigen Daten
					  ; - - - - - - - - - - - - - - - - - - - - - - - - -
						else If IsObject(this.rxFilter[KKAlias]) {

							If (this.rxFilter[KKAlias].Count() > 0)
								For eachRx, rxObj in this.rxFilter[KKAlias]
									item.Inhalt := RegExReplace(item.Inhalt, rxObj.rxm, rxObj.rpl)



							item.inhalt := RegExReplace(" " item.Inhalt " ", this.rxFRplStd, "") ; Namenskürzel entfernen
							item.Inhalt := RegExReplace(Trim(item.Inhalt), "[\n\r]+"       	, "<br>")
							item.Inhalt := RegExReplace(item.Inhalt, "(\<br\>\s*)+"	, "<br>")
							item.Inhalt := RegExReplace(item.Inhalt, "\s{2,}"      		, " ")

						}

					  ; nicht aufgeführte Kürzel wurden nicht übernommen
					  ; - - - - - - - - - - - - - - - - - - - - - - - - -


						}

						If !item.inhalt && item.InhaltO {
							item.inhalt := Trim(item.InhaltO)
							item.inhalt := RegExReplace(" " item.Inhalt " ", this.rxFRplStd, "") ; Namenskürzel entfernen
						}

					 }
				}
				;}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Matchobjekt speichern
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
				If IsObject(KKarte)
					FileOpen(this.pat[PatID].jsonfile2, "w", "UTF-8").Write(cJSON.Dump(KKarte, 1))
				else {
					SciTEOutput("Es konnten keine Daten für diesen Patienten [" PatID "] " PatName "  ausgelesen werden.")
					return
				}
			;}

		}

		;}

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; HTML Datei erstellen
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		FileGetSize, fHTML, % (this.pat[PatID].htmlpath := this.pat[PatID].basePatPath "\Karteikarte.html")
		If (!FileExist(this.pat[PatID].htmlpath) || !fHTML || settings.reNewData) {

			ZeigsMir("Karteikarte als PDF erstellen",0,KKarte.Count(),0,0, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))
			If FileExist(this.pat[PatID].htmlpath)
				FileDelete, % this.pat[PatID].htmlpath

			htmlfile := FileOpen(this.pat[PatID].htmlpath, "w", "UTF-8")
			lastMonth := lastday := ""
			WTD := true
			EOL := "`r`n"
			KKIndex := 0
			expHTML := this.html

			For dayDate, data in KKarte {

				KKIndex ++
				If (Mod(KKIndex, 10) = 0)
					ZeigsMir("Karteikarte als PDF erstellen", KKIndex, KKarte.Count(),0,0, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))

				If (lastday != dayDate) {

					year := SubStr(Trim(dayDate), 1, 4)
					If !lastyear
						lastyear := year
					else if (year > lastyear)
						newyear := "A"

					WTD := true
					lastday := dayDate
					lastyear := year

				}

				expHTML .=  "<div class='DayBox'" (newYear ? " id='" Format("{:U}", newYear) "'>" : ">") 																. EOL
				If newYear {
					expHTML .= "`t<div class='rowKzlA'>"                                                                                                                                  	. EOL
					expHTML .= "`t`t<div class='colYear'>" year "</div>"																												. EOL
					expHTML .= "`t</div>"																																							. EOL
				}
				newYear := false

				For KZLAlias, item in data {

					If IsObject(item.document) && item.document.Count()>0 {
						For each, doc in item.document {
							expHTML .= "`t<div class='rowKzl'>"                                                                                                                             	. EOL
							expHTML .= "`t`t<div class='colDay'>" (WTD ? ConvertDBASEDate(dayDate) : "") "</div>"                                               	. EOL
							expHTML .= "`t`t<div class='colShort' id='" doc.alias "'>" doc.alias "</div>"                                                                   	. EOL
							expHTML .= "`t`t<div class='colContent' id='" doc.alias "'><a href='" doc.Link
												. "' target='_blank'>" EncodeHTML(doc.txt) "</a></div>" 																				. EOL
							expHTML .= "`t</div>"                                                                                                                                                  	. EOL . EOL
							WTD := false
						}
					}
					else {

						If !item.inhalt && item.inhaltO
							item.inhalt := Trim(item.inhaltO)

						If item.Inhalt {
							item.inhalt := RegExReplace(item.inhalt, "[\n\r]+"    	, "<br>")
							item.inhalt := RegExReplace(item.inhalt, "(\<br\>\s*)+", "<br>")
							item.inhalt := RegExReplace(item.inhalt, "\h{2,}"  		, " ")
							expHTML .= "`t<div class='rowKzl'>"                                                                                                                             	. EOL
							expHTML .= "`t`t<div class='colDay'>" (WTD ? ConvertDBASEDate(dayDate) : "") "</div>"            	                                 	. EOL
							expHTML .= "`t`t<div class='colShort' id='" KZLAlias "'>" EncodeHTML(KZLAlias) "</div>"                                                 	. EOL
							expHTML .= "`t`t<div class='colContent' id='" KZLAlias "'>" EncodeHTML(item.Inhalt) "</div>"                                          	. EOL
							expHTML .= "`t</div>"                                                                                                                                                  	. EOL . EOL

						}

					}

					;~ expHTML .= "`t</div>" . EOL

					WTD := false

				}

				expHTML .= "</div>" . EOL . EOL

			}

			IF (KKIndex <= 5) {
				If !FileExist(A_ScriptDir "\resources\QEHinweise.txt")
					FileAppend, % "### Datei mit Hinweisen auf mögliche Konvertierungsfehler ###", % A_ScriptDir "\resources\QEHinweise.txt", UTF-8
				msg1 := "`n- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
				FileAppend, % msg1 	"`nName:`t" 	(msg:=  this.pat[PatID].ExportPatName " (" PatID ") hat nur " KKIndex " Einträge im Tagesprotokoll. Fehlerhafte Konvertierung?")
											. 	"`nPfad:`t"   		this.pat[PatID].basePatPath , % A_ScriptDir "\resources\QEHinweise.txt", UTF-8
				this.dbg ? SciTEOutput(A_ThisFunc "() | " Trim(msg)) : ""
				this.convertProblems += 1
			}

			ZeigsMir("Karteikarte als PDF erstellen", KKIndex, KKarte.Count(),0,0, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))

		  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		  ; HTML Ausgaben kürzen
			expHTML := StrReplace(expHTML, "##PATIENTNAMEBIRTH##"	, this.pat[PatID].ExportPatName)
			expHTML := StrReplace(expHTML, "##PATIENTADDRESS##" 	, this.pat[PatID].ExportPatAddress)

			expHTML .= "</div></div>"         	. EOL . EOL
			expHTML .= "</body>" .  EOL . "</html>"		. EOL
			expHTML := RegExReplace(expHTML, "(\<br\>\s)+", "<br>")

		  ; Medikamentendosierung
			LStart := 20
			Loop % Lstart {

				rxrmatchstr := "(St|ml)"
				rxreplacestr := "$1 "
				ARounds := Lstart-A_Index + 1

				Loop % ARounds {
					rxrmatchstr .= "<br>([\d\pL\s]+|)"
					rxreplacestr .= (A_Index = 1 ? " " : A_Index = 2 ? "" : "-") "$" A_Index+1
				}

				rxrmatchstr .= "\<br\>"
				rxreplacestr .= " <br>"
				;~ SciTEOutput(A_ThisFunc " (" A_LineNumber "): `n  " SubStr("00" A_Index, -1) "  -  " rxrmatchstr "`n" rxreplacestr )

				If RegExMatch(expHTML, rxrmatchstr)
					expHTML := RegExReplace(expHTML, rxrmatchstr, rxreplacestr)

			}
		  ; leere	Tagesboxen entfernen
			expHTML := RegExReplace(expHTML, "i)\<div\h+class\=\'DayBox\'(\s+id=\'[A-D]\')*\>([\r\n])*\<\/div\>")
			expHTML := RegExReplace(expHTML, "\h+(\r\n)", "`r`n")
			expHTML := RegExReplace(expHTML, "(\r\n){2,}", "`r`n")
			expHTML := RegExReplace(expHTML, "i)(\s+\<div\s+class\=\'DayBox\'\>)", "`r`n$1")

			;~ expHTML := RegExReplace(expHTML, "(\p{Ll}{2,})(\p{Lu}\p{Ll}{2,})" 	, "$1 $2")
			expHTML := RegExReplace(expHTML, ";\h*"                            	, "; ")
			expHTML := RegExReplace(expHTML, ",\h*"                             	, ", ")
			expHTML := RegExReplace(expHTML, "\h{2,}"                        	, " ")
			expHTML := RegExReplace(expHTML, "(HEXAL AG)([\pL\d])"      	, "$1`n$2")
			expHTML := RegExReplace(expHTML, "(GmbH*)(\p{Lu})"           	, "$1`n$2")
			expHTML := RegExReplace(expHTML, "HNPN3"                         	, "HKP N3 ")
			expHTML := RegExReplace(expHTML, "\)([A-Z\d])"                   	, ")`n$1")
			expHTML := RegExReplace(expHTML, "\<\h+br\>\h*"           	, "<br>")


			htmlfile.Write(expHTML)
			htmlfile.Close()
		}
	  ;}

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; Patientenkarteikarte öffnen für Export der Labordaten
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		If !settings.noLabData {
			FileGetSize, fLabpdf, % this.pat[PatID].basePatPath "\Labordaten.pdf"
			If (settings.reNewData || !FileExist(this.pat[PatID].basePatPath "\Labordaten.txt"))                                                          	; diese Datei zeigt an das keine Labordaten im System waren
				If (settings.reNewData || !FileExist(this.pat[PatID].basePatPath "\Labordaten.pdf") || !fLabpdf ) {                                  	; eine vorhandene Datei wird nicht überschrieben

					res := AlbisAkteOeffnen(PatName :=cPat.Name(PatID, false), PatID, false)

					SetTitleMatchMode        	, RegEx
					WinWait, % "Herzlichen\s+Glückwunsch ahk_class #32700 ahk_exe " Adddendum.AlbisExe,, 2
					If (hwnd := WinExist("Herzlichen\s+Glückwunsch zum ahk_class #32700 ahk_exe " Adddendum.AlbisExe)) {
						dbg ? SciTEOutput("herzlichen Glückwunsch gefunden") : ""
						 If !VerifiedClick("OK",,, hwnd, 2)
							VerifiedClick("Button1",,, hwnd, 2)
						 WinWaitClose, % "Herzlichen Glückwunsch zum ahk_class #32700 ahk_exe " Adddendum.AlbisExe,, 2
					}

					WinWait, % "ALBIS.*?\[" PatID " ahk_class OptoAppClass",, 3

					currPatID := AlbisAktuellePatID()
					If (currPatID = PatID) {

						If FileExist(this.pat[PatID].basePatPath "\Labordaten.pdf")
							FileDelete, % this.pat[PatID].basePatPath "\Labordaten.pdf"

						ZeigsMir("Labordaten exportieren",,,,, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))
						ZeigsMir("selfcall=on")

						res := Laborblatt_Export(PatName, this.pat[PatID].basePatPath, "Microsoft Print to Pdf", "Alles")
						If (res = -1) {
							t := Addendum.Praxis.Name " (" Addendum.Praxis.Facharzt ")" " " Addendum.Praxis.Strasse ", " Addendum.Praxis.PLZ " " Addendum.Praxis.ORT "`n"
							t .= "Pat.: " cPat.Name(PatID, true) ", "  cPat.Get(PatID, "STRASSE") (cPat.Get(PatID, "HAUSNUMMER") ? " " cPat.Get(PatID, "HAUSNUMMER"):"") " in " cPat.Get(PatID, "PLZ") " " cPat.Get(PatID, "ORT") "`n"
							t .= "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `n"
							t .= "  Es lagen keine Labordaten für einen Datenexport vor.                ("  Addendum.Praxis.ORT ", den " A_DD "." A_MM "." A_YYYY  ")"
							t .= "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `n"
							FileOpen(this.pat[PatID].basePatPath "\Labordaten.txt", "w", "UTF-8").Write(t)
						  ; ersetzt das Feld mit dem Link zu Laborbefunden
							rplHTML := "<li><a href='Labordaten.pdf' target='_blank'>Laborbefunde</a> (Tagesbezogen) </li>"
							htmlText := FileOpen(this.pat[PatID].htmlpath, "r", "UTF-8").Read()
							htmlText := StrReplace(htmlText, rplHTML, "<span>[keine Laborbefunde erhoben]</span>")
							FileOpen(this.pat[PatID].htmlpath, "w", "UTF-8").Write(htmlText)
						}

						Sleep 1000

						firstChildsCount := AlbisMDIClientWindows().Count()
						ChildsCountDiff := 0, startCTime := A_TickCount, waitingTime := 0

						If (firstChildsCount>2) {
							while (!ChildsCountDiff && waitingTime <= 5){                               ; Abbruch wenn eine Akte geschlossen wurde oder wenn 5 Sekunden vergangen sind

								If (Mod(A_Index+9, 10) = 0 ) {
									AlbisActivate(1)
									ControlSend,, % "{LControl Down}{F4}{LControl Up}", %  "ahk_class OptoAppClass"
									If ErrorLevel {
										ControlSend, MDIClient1, % "{LControl Down}{F4}{LControl Up}", %  "ahk_class OptoAppClass"
										EL := ErrorLevel
									}
									Sleep 400
									ChildsCountDiff := firstChildsCount - AlbisMDIClientWindows().Count()
									dbg ? SciTEOutput(A_ThisFunc "() [ " A_Index ", " Round((A_TickCount-startCTime)/1000, 1) "s] ChildsDiff: " ChildsCountDiff ", EL: " EL) : ""
								}
								else
									Sleep 200

								ChildsCountDiff := firstChildsCount - AlbisMDIClientWindows().Count()
								waitingTime := Round((A_TickCount-startCTime)/1000, 1)

							}
						}

						ZeigsMir("selfcall=off")
						Sleep 2000

					}
				}
		}
		;}

	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; HTML in PDF konvertieren - nutzt den Edge Browser. Cmdline Tools wie Pandoc (wkhtmltopdf und andere PDF Renderer) zerstören das Layout
	  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		FileGetSize, fPDF, % (pdfFilePath := this.pat[PatID].basePatPath "\Karteikarte " PatName ".pdf")
		If FileExist(this.pat[PatID].htmlpath) && (!FileExist(pdfFilePath) || !fPDF || settings.reNewData) {
			ZeigsMir("MSEdge für PDF-Ausgabe nutzen",,,,, indexshow, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))
			ZeigsMir("selfcall=on")
			If FileExist(this.pat[PatID].pdfFilePath :=PdfFilePath)
				FileDelete, % this.pat[PatID].pdfFilePath

			; cmdline Befehl ist zuverlässiger und wesentlich schneller
			RunWait, % "msedge.exe --disable-pdf-tagging --no-pdf-header-footer --headless --print-to-pdf=" q this.pat[PatID].pdfFilePath q " " q this.pat[PatID].htmlpath q
			this.pat[PatID].pdfSaved := FileExist(this.pat[PatID].pdfFilePath) ? 2 : 1

			;~ this.pat[PatID].pdfSaved := msEdge_SaveAsPDF(this.pat[PatID].htmlpath, this.pat[PatID].pdfFilePath)
			If (!FileExist(this.pat[PatID].pdfFilePath) || this.pat[PatID].pdfSaved <= 1){
				SciTEOutput("[" PatID "] " PatName ": Die Erstellung der Karteikarte als PDF-Datei ist fehlgeschlagen. [Ursache: "
									.(this.pat[PatID].pdfSaved	=1  	? "Die Umwandlung per commandline Befehl ist fehlgeschlagen."
									. this.pat[PatID].pdfSaved	=-1  	? "Die zu konvertierende HTML Datei ist nicht vorhanden"
									:	this.pat[PatID].pdfSaved	=-2  	? "Das Verzeichnis in welches die PDF-Datei gespeichert werden soll ist nicht angelegt."
									:	this.pat[PatID].pdfSaved	=-3  	? "Der Edge Browser konnte nicht angesprochen werden (class_UIA_Browser.ahk)."
									:	this.pat[PatID].pdfSaved	=-4  	? "Der Druckdialog konnte nicht geöffnet werden."
									:	this.pat[PatID].pdfSaved	<=-77	? "Der Edge Browser wurde nicht gefunden (class_UIA_Browser.ahk). [" this.pat[PatID].pdfSaved "]"
									: this.pat[PatID].pdfSaved	<= 0 	? "Bei der Automatisierung des Edge Browser ist etwas schief gegangen. [error code: " this.pat[PatID].pdfSaved "]" : ""))
			}
			ZeigsMir("selfcall=off")
		}
		;}


		ZeigsMir("selfcall=off")

	return this.pat[PatID].basePatPath
	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	CheckPatID(rowNr, rowData)                         	{


		PatID			:= rowData.names.PatID
		NAME     	:= StrSplit(cPat.NAME(PatID, false) , ",", A_Space).1
		VORNAME 	:= StrSplit(cPat.NAME(PatID, false)	, ",", A_Space).2
		Birth   	:= cPat.GEBURT(PatID,false)

		;~ SciTEOutput(element.row " - " PAtID "," SName "," PName "," rBirth "," NAME "," VORNAME  "," Birth)

		If (!cPat.Exist(PatID))  {

			LV_GetText(note, rowNr, 10)
			If RegExMatch(note, "\[(.+)\]", match)
				note := RegExReplace(note, "\[.+\]\s*", "")
			If cPat.Exist(PatID)
				LV_Modify(rowNr, "Col10", "[!! PatID " PatID " inkorrekt: " NAME ", " VORNAME " *" cPat.GEBURT(PatID, true)  ", letzte Beh.: " cPat.LASTBEH(PatID, true) " !!] " note)
			else
				LV_Modify(rowNr, "Col10", "[!! PatID gehört zu keinem Patienten !!] " note)

			return false
		}

	return true
	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	DataAggregation(DatabaseLinks)                     	{ 	;-- sammelt Daten aus zwei Datenbanken in zwei Textdateien

		Delim := "`t|`t"

		dbase 	:= new DBASE(this.AlbisDBPath "\BEFUND.DBF")
		res    	:= dbf.OpenDB()

		For index, columnname in dbf.fields
			tblheader .= (!columnname ? "" : Delim) columnname

		;~ while !dbf.atEOF

	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	ExportStatistic()                                  	{

		this.TimeSum	:= this.TSummedFiles := 0
		this.startTime:= A_TickCount


		bPath        	:= {"folders"	: {}
										, "stats"  	: {	"filesC"          	: 0
																 , 	"filesS"          	: 0
																 , 	"pages_shipped"    	: 0
																 , 	"pages_unshipped"  	: 0
																 , 	"shipped"         	: 0
																 ,	"unshipped"       	: 0
																 , 	"fExt"            	: {}
																 , 	"Max"             	: {"Files"	: {"Pat":0, "Doc":0}
																						        			, "Size"	: {"Pat":0, "Doc":0}
																						        			, "Pat" 	: ""}}}

		tmp := this.GetFilesFromBasePath()

		For fExt, fExtC in tmp.fExt
			bPath.stats.fExt[fExt] := fExtC

	  ; Fortschrittsbalkenanzeige
		steps := tmp.files.Count()<=100 ? 10 : Round(tmp.files.Count()/100)

		For fIndex, file in tmp.files {

			fpath	:= file.1
			fSize	:= file.2
			fExt 	:= file.3

			If InStr(fpath, "xData")
					continue



			If 	RegExMatch(fpath, "Oi)^(?<docName>.+)\s*\[(?<bsnr>\d+)\]\s*(\\(?<shipped>versendet))*\\\s*\((?<PatID>\d+)\)\s*(?<Patname>.+)\\(?<fname>.+)\s*\.(?<fExt>\w+)$", m)
			||	RegExMatch(fpath, "Oi)^\s*(?<docName>Patient)\s*(\\(?<shipped>versendet))*\\\s*\((?<PatID>\d+)\)\s*(?<Patname>.+)(\\(?<xData>XData))*\\(?<fname>.+)\.(?<fExt>\w{1,4})$", m) {

				bsnr := !m.bsnr ? "Patient" : m.bsnr
				shipment := m.shipped = "versendet" ? "shipped" : "unshipped"
				If !bsnr
					SciTEOutput("m.docName: " m.fname "." fExt " (" bsnr " / " m.PatID " " m.Patname ")")
				If !IsObject(bPath.folders[bsnr]) {

					shData := {"name": m.Patname, "filesC":1, "filesS":fSize, "fAverage":fSize , "fExt":{(m.fExt):1}}

					bPath.folders[bsnr] := {	"1_name_doctor" 	: m.docName
										      				,	"2_bsnr_doctor"  	: bsnr
										      				, "unshipped"      	: {"filesC":0, "filesS":0, "Pages":0, "fExt":{}}
										      				, "shipped"        	: {"filesC":0, "filesS":0, "Pages":0, "fExt":{}}
										      				,	"filesC"         	: 1
										      				,	"filesS"					: fSize
										      				,	"fAverage"				: fSize
										      				,	"fExt"           	: {(m.fExt):1}}

					bPath.folders[bsnr][shipment][m.PatID] := shData

				}
				else {

					item := bPath.folders[bsnr]

					If !IsObject(item[shipment][m.PatID])
						item[shipment][m.PatID] := {"name": m.Patname, "filesC":1, "filesS":fSize, "fAverage":fSize, (m.fExt):1}
					else {

						item[shipment].filesC                	+= 	1
						item[shipment].filesS                	+= 	fSize
						item[shipment].fAverage              	:= 	FormatedFileSize(item[shipment].filesS / item[shipment].filesC)
						item[shipment].fExt[m.fExt] 	       	:= 	item[shipment].fExt[m.fExt] ? item[shipment].fExt[m.fExt]+1 : 1
						item[shipment][m.PatID].filesC       	+= 	1
						item[shipment][m.PatID].filesS       	+= 	fSize
						item[shipment][m.PatID].fAverage   		:= 	FormatedFileSize(item[shipment][m.PatID].filesS / item[shipment][m.PatID].filesC)
						item[shipment][m.PatID].fExt[m.fExt]	:= 	item[shipment][m.PatID].fExt[m.fExt] ? item[shipment][m.PatID].fExt[m.fExt]+1 : 1

					}

					item.filesC     	+= 1
					item.filesS     	+= fSize
					item.fAverage    	:= FormatedFileSize(item.filesS / item.filesC)
					item.fExt[m.fExt] := item.fExt[m.fExt] ? item.fExt[m.fExt] +1 : 1

				}

				if (fExt = "pdf" && Addendum.qpdfPath) {
					PDFPages := this.GetPDFPages(fpath)
					bPath.stats["pages_" shipment] += PDFPages
				}
				else if (fExt ~= "i)(doc|docX)")

				bPath.stats.filesC  	+= 1
				bPath.stats.filesS  	+= fSize
				bPath.stats.fAverage 	:= FormatedFileSize(bPath.stats.filesS / bPath.stats.filesC)
				bPath.stats[shipment]	+= 1

				If (bPath.stats.Max.Files.Doc < item.filesC)
					bPath.stats.Max.Files.Doc := item.filesC, bPath.stats.Max.Files.docName := m.docName " [" bsnr "]"

				If (bPath.stats.Max.Files.Pat < bPath.folders[bsnr][shipment][m.PatID].filesC)
					bPath.stats.Max.Files.Pat := bPath.folders[bsnr][shipment][m.PatID].filesC, bPath.stats.Max.Files.PatIDName 	:= "[" m.PatID "] " m.Patname

				If (bPath.stats.Max.Size.Doc 	< item.filesS)
					bPath.stats.Max.Size.Doc := item.filesS,  bPath.stats.Max.Size.docName := m.docName " [" bsnr "]"

				If (bPath.stats.Max.Size.Pat	< bPath.folders[bsnr][shipment][m.PatID].filesS)
					bPath.stats.Max.Size.Pat := bPath.folders[bsnr][shipment][m.PatID].filesS, bPath.stats.Max.Size.PatIDName 	:= "[" m.PatID "] " m.Patname

			  ; Fortschritt anzeigen
				If (Mod(fIndex+steps-1, steps) = 0 || fIndex = tmp.files.Count()) {

					DocFolders 	:= bPath.folders.Count()-1
					PatsinDocFolders := PatsinPatFolder := PatC :=  0
					For Folder, PatFolder in bPath.folders {
						isDocFolder := false
						If (Folder ~= "\[\d+\]")
							isDocFolder := true
						shipped 	:=  PatFolder.shipped.Count()   ? PatFolder.shipped.Count()-4   : 0
						unshipped :=  PatFolder.unshipped.Count() ? PatFolder.unshipped.Count()-4 : 0
						If isdocFolder
							PatsinDocFolders 	+= (shipped + unshipped)
						else
							PatsinPatFolder  	+= (shipped + unshipped)

						PatC += += (shipped + unshipped)
					}

					ZeigsMir("Dateistatistik läuft. Datei: "	, fIndex
																										, tmp.files.Count()
																										, FormatedFileSize(bPath.stats.filesS)
																										, [DocFolders, PatC]
																										, ""
																										, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 ), true))

				}

			}
			else {
				SciTEOutput("fpath failed: " fpath)
			}

			bsnr := ""

		}

		bPath.stats.Max.Size.Doc 	:= FormatedFileSize(bPath.stats.Max.Size.Doc)
		bPath.stats.Max.Size.Pat 	:= FormatedFileSize(bPath.stats.Max.Size.Pat)
		bPath.stats.filesS       	:= FormatedFileSize(bPath.stats.filesS)
		For bsnr, data in bPath.folders 		{
			shipped 	    	:= bPath.folders[bsnr].Delete("shipped")
			unshipped	    	:= bPath.folders[bsnr].Delete("unshipped")
			PatCs   	     	:= shipped.count()   	? shipped.count()-4 		: 0
			PatCu          	:= unshipped.count()  ? unshipped.count()-4 	: 0
			bPath.folders[bsnr].shipped 	:= {"PatC":PatCs 	, "filesC":shipped.filesC  	, "filesS":(shipped.filesS   	? FormatedFileSize(shipped.filesS)  	: 0)	, "fExt":shipped.fExt}
			bPath.folders[bsnr].unshipped	:= {"PatC":PatCu	, "filesC":unshipped.filesC	, "filesS":(unshipped.filesS 	? FormatedFileSize(unshipped.filesS)	: 0)	, "fExt":unshipped.fExt}
		}

		FileOpen(this.basePath "\Statistik " A_Now ".json", "w", "UTF-8").Write(cJSON.Dump(bPath, 2))
		SciTEOutput(cJSON.Dump(bPath, 2))

	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	class TaskGui extends Quickexporter               	{

		__New() {

			Gui, TG: new, hwndhTG
			this.hGui := hTG

			; ; \sSt(<br>(\d+x*|Dosis:\d+\*\d*\s*m\s*\d+xTgl\.|pro|Woche|.|))+

		}

	}


  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
  ; interne Methoden
  ; ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲ ▼▲
	GetFilesFromBasePath()                             	{

		tmp      	:= {"files":[], "fExt":{}, "stats":{"filesC":0, "filesS":0}}
		LenBPath 	:= StrLen(this.basepath)

	 ; gespeicherte Stastikdaten lesen
		If FileExist(this.StatsIniFilePath)
			IniRead, filesCountLast, % 	this.StatsIniFilePath, % "Verzeichnisse", % "Dateien_gesamt", 0
		else
			FileOpen(this.StatsIniFilePath, "w", "UTF-8").Write("[Quickexporter]`nStatistikdatei_angelegt_am=" A_DD "." A_MM "." A_YYYY "`n")

	  ; Fortschrittsbalkenanzeige
		steps := !filesCountLast ? 10 : Round(filesCountLast/100)

	  ; Dateien einsammeln
		Loop, Files, % this.basePath "\*.*", R
			If (A_LoopFileExt ~= "i)^" this.mediatypes "$") {
				tmp.files.Push([SubStr(A_LoopFileFullPath, -1*(StrLen(A_LoopFileFullPath)-LenBPath-2)), A_LoopFileSize, StrReplace(A_LoopFileExt, ".")])
				tmp.fExt[A_LoopFileExt] := tmp.fExt[A_LoopFileExt] ?  tmp.fExt[A_LoopFileExt]+1 : 1
				tmp.stats.filesC += 1
				tmp.stats.filesS += A_LoopFileSize
				If (Mod(A_Index+steps-1, steps) = 0)
					ZeigsMir("Verzeichnis auslesen. Datei: ", tmp.files.Count()
																									, (!filesCountLast ? tmp.files.Count() : filesCountLast)
																									, FormatedFileSize(tmp.stats.filesS)
																									, 0
																									,
																									, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 )
																									, true))
			}

		ZeigsMir("Verzeichnis auslesen. Datei: ", tmp.files.Count()
																						, (!filesCountLast ? tmp.files.Count(): filesCountLast)
																						, FormatedFileSize(tmp.stats.filesS)
																						, 0
																						,
																						, TimeFormatEx(Round((A_TickCount-this.startTime)/1000 )
																						, true))

		IniWrite, % tmp.files.Count(), % this.StatsIniFilePath, % "Verzeichnisse", % "Dateien_gesamt"

	return tmp
	}

	GetPDFPages(filepath)                             	{

		if !FileExist(filepath)
			return 0

	}

	Filexpro(sFile := "", property:="")               	{           ; v.90 By SKAN on D1CC @ goo.gl/jyXFo9
		; modified to return all properties with values
		; this function works only inside this class!
		Local
		Static xDetails, _DirLast

		 fex := {}
		 SplitPath, sFile, _FileExt, _Dir, _Ext, _File, _Drv

		objShl := ComObjCreate("Shell.Application")
		objDir := objShl.NameSpace(this.basePath)
		objItm := objDir.ParseName(_FileExt)

		 If ( VarSetCapacity(xDetails) = 0 )   {
			i:=-1,  xDetails:={},  xDetails.SetCapacity(309)
			While ( i++ < 309 )
				xDetails[ objDir.GetDetailsOf(0,i) ] := i
			xDetails.Delete("")
		 }

		 If property
			return ObjDir.GetDetailsOf(objItm, xDetails)

			i:=0
			for Prop, PropNum in xDetails {

				If ( (Dot:=InStr(Prop,".")) and (Prop:=(Dot=1 ? "System":"") . Prop) )
					value := objItm.ExtendedProperty(Prop)
				else If (PropNum > -1)
					value := ObjDir.GetDetailsOf(objItm,PropNum)

				if value
					fex[Prop] := value

			}
			fex.SetCapacity(-1)

		Return fex
	}

}

;}

;{ ⚞──────────────────       FUNKTIONEN          	──────────────────⚟

; Gui                                                              	;{

CPRGHandler:                                                                           	;{

	if (A_GuiControl = "CPRGCancel")
		settings.CPRGCancel := true

return
;}

MSGHandler:                                                                             ;{
MSGGuiClose:
MSGGuiEscape:
	If GuiControlGet("Msg", "", "MsgChk")
		IniWrite, 1, % settings.paths.QEExporterINI, % "Script", % "Sumatra_Hinweis_nicht_zeigen"
	Gui, Msg: Destroy
return ;}

QEGuiClose:                                                                           	;{	;-- speichert Einstellungen beim Beenden des Programmes

	OnExit

	Gui, QE: Default
	Gui, QE: Submit, NoHide


	res := RegistryWrite("REG_SZ", KeyName, assoc, enc := StringDecoder(QEZPWDH,, false))


	QESdate := !(QESdate ~= "\d{2}\.\d{2}\.\d{4}")	? GuiControlGet("QE", "", "QESDate") : QESdate
	If (QESdate := QESdate ~= "\d{2}\.\d{2}\.\d{4}" 	? QESdate : "")
		IniWrite, % QESdate, % settings.paths.QEExporterINI, % "SQLite", % "Tagesprotokoll_ab"


	For bsnr, physican in settings.physicans
		If (physican.MediumO != physican.Medium) || (physican.RouteO != physican.Route)
			IniWrite, % Trim(physican.Medium "|" physican.Route, "|"), % settings.paths.QEExporterINI, % "Ärzte", % bsnr



	settings.LVExpFilter  	:= GuiControlGet("QE", "", "QELVExpFilter")
	settings.LVExpFilter2 	:= GuiControlGet("QE", "", "QELVExpFilter2")
	If (settings.LVExpFilter2 && settings.LVExpFilter)
		GuiControl, QE:, QELVExpFilter, % (settings.LVExpFilter := false)


	;~ networkpath := GuiControlGet("QE",, "QEZPath")
	IF (QEZPath && QEZPath != settings.paths.networkpath)
		IniWrite, % QEZPath                             	, % settings.paths.QEExporterINI, % "Export"      	, % "networkpath_" compname

	GuiPos := GetWindowSpot(hQE)
	IniWrite, % "x=" GuiPos.X " y=" GuiPos.Y           	, % settings.paths.QEExporterINI, % "Script"      	, % "GuiPosition_" compname
	IniWrite, % settings.LVExpFilter                   	, % settings.paths.QEExporterINI, % "Script"      	, % "fehlerfreie_Exporte_ausblenden"
	IniWrite, % settings.LVExpFilter2                 	, % settings.paths.QEExporterINI, % "Script"      	, % "versandte_Exporte_ausblenden"
	IniWrite, % settings.PWDVisibility                	, % settings.paths.QEExporterINI, % "Script"      	, % "UserPWD_Sichtbarkeit"
	IniWrite, % GuiControlGet("QE",, "QESDBG")      		, % settings.paths.QEExporterINI, % "Script"      	, % "Debug"
	IniWrite, % GuiControlGet("QE",, "QESQLWP")		     	, % settings.paths.QEExporterINI, % "SQLite"      	, % "SQL_write_protection"
	IniWrite, % GuiControlGet("QE",, "QENoLData")   		, % settings.paths.QEExporterINI, % "Export"      	, % "keine_Labordaten_exportieren"
	IniWrite, % GuiControlGet("QE",, "QEVBAct")       	, % settings.paths.QEExporterINI, % "Versandbrief"	, % "Brief_erstellen"
	IniWrite, % GuiControlGet("QE",, "QEVBIPrint")    	, % settings.paths.QEExporterINI, % "Versandbrief"	, % "Brief_drucken"
	IniWrite, % GuiControlGet("QE",, "QEZLetter")     	, % settings.paths.QEExporterINI, % "Export"      	, % "Briefe_in_7z_inkludieren"

	;~ If QEVBIprinter
		;~ IniWrite, % QEVBIprinter                         	, % settings.paths.QEExporterINI, % "Versandbrief"	, % "Brief_Drucker"


	If (VBerstellen := GuiControlGet("QE",, "QEVBAct")) {

		Absender := StrSplit(GuiControlGet("QE",, "QEVBData"), "`n", "`r")
		writeToIni := true
		emptyLines := 0
		For lineIndex, part in Absender
			If (StrLen(Trim(part)) = 0) {
				emptylines |= lineIndex^2
			}
			else {

			}
		If emptyLines {
			MsgBox, 0x1000, % scriptTitle, % "Es fehlen Absenderangaben im Adressfeld (" GetHex(emptylines) ").`nSollen die bisherigen Daten dennoch mit diesen ersetzt werden?"
			IfMsgBox, No
				WriteToIni := false
		}
		If writeToIni {


			if RegExMatch(Absender[1], "i)[\pL\-\.]+\s+[\pL\-\.]+(\s+[\pL\-\.]+)*")
				IniWrite, % Absender[1]                     		, % settings.paths.QEExporterINI, % "Versandbrief" 	, % "Absender_Name"
			if RegExMatch(Absender[2], "i)(medizin|Innere|hausarzt|hausärztlich|Internist)")
				IniWrite, % Absender[2]                     		, % settings.paths.QEExporterINI, % "Versandbrief" 	, % "Absender_Facharzt"
			if RegExMatch(Absender[3], "^\s*\p{Lu}[\pL\-\.]+")
				IniWrite, % Absender[3]                      		, % settings.paths.QEExporterINI, % "Versandbrief" 	, % "Absender_Ort"
			If InStr(Absender[4], "@")
			if RegExMatch(Absender[3], "^\s*\w[\w\-_]+?@\w.+\.\w{2,4}")
				IniWrite, % Absender[4]                      		, % settings.paths.QEExporterINI, % "Versandbrief" 	, % "Absender_Mail"
		}

	}

	Gui, QE: Destroy

; Datenbankzugriff beenden
	QEDB.CloseDB()

; Instanzen beenden
	DAB := cPat :=PatDB := qexporter := ""

	If ReloadScript
		Reload

ExitApp   ;}

QEFocus:                                                                               	;{

	Gui, QE: Default
	HotkeyCall := A_ThisHotkey
	GuiControlGet, vfocus, QE: FocusV
	If (vfocus = "QELvExp" && InStr(A_ThisHotkey, "RButton")) {
		MouseGetPos, mx, my
		Sleep 50
		Menu, QECM, Show, % mx-20, % my+10
		return
	}

	Gui, QE: Submit, NoHide
	If (vfocus = "QEFEdit" && A_ThisHotkey ~= "i)(NumpadEnter|Enter|Tab)") {
		;~ SciTEOutput("vFocus: " vfocus ", hotkey: " A_ThisHotkey ", x" mx " y" my)
		QEFEditHandler(QEFEdit)
		return
	}
	else if (vFocus ~= "QEZPWD(H|V)")             	{
		PWD := %vFocus%
		;~ counting += 1
		;~ ToolTip, % "(" counting ") EvI: " A_EventInfo ", " A_GuiControlEvent "," vfocus "`nPWD: "  PWD, 2000, 300, 12
		If (lastPWD = PWD)
			return
		vis := settings.PWDVisibility
		GuiControl, % "QE: -g"   	 		, % "QEZPWDH"
		GuiControl, % "QE: -g"  	 		, % "QEZPWDV"
		GuiControl, % "QE: Enable" 		, % (vfocus="QEZPWDH" ? "QEZPWDV" : "QEZPWDH")
		GuiControl, % "QE: "     	 		, % (vfocus="QEZPWDH" ? "QEZPWDV" : "QEZPWDH"), % (lastPWD := PWD)
		GuiControl, % "QE: Hide"	 		, % (vis  ? "QEZPWDH" : "QEZPWDV")
		GuiControl, % "QE: Disable"		, % (vis  ? "QEZPWDH" : "QEZPWDV")
		GuiControl, % "QE: Show"   		, % (vis  ? "QEZPWDV" : "QEZPWDH")
		GuiControl, % "QE: Enable" 		, % (vis  ? "QEZPWDV" : "QEZPWDH")
		GuiControl, % "QE: +gQEFocus" , % (vis  ? "QEZPWDH" : "QEZPWDV")
		GuiControl, % "QE: +gQEFocus" , % (vis  ? "QEZPWDV" : "QEZPWDH")
		IF InStr(PWD, "***")
			StringDecoder("***", "", false) ; löscht das letzte Passwort
		PWD := ""

	  return
	}

	MouseGetPos, mx, my,
	if (vfocus ~= "QELvExp")
;}

QEBHandler:                                                                           	;{	;-- Steuerelementhandler

	Critical, On
	Thread, Priority, 2
	Gui, QE: Default
	Gui, QE: Submit, NoHide

  ; nur Steuerelement mit Patientennamenauslesen
	Pat := QEPatientTxt()

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  If (A_GuiEvent == "K") && (Chr(A_EventInfo) = "e") {
	Gui, QE: ListView, QELvExp
	dbg ? SciTEOutput(A_ThisLabel ": row: " LV_GetNext(0, "Focused") " A_EventInfo: " A_EventInfo ", cRow: " cRow) : ""
	ICELLvExp.EditCell(LV_GetNext(0, "Focused"))
}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Karteikartenexport - im Moment keine Mehrfachauswahl möglich
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			 If (A_GuiControl = "QEExport")     	      		{ ; Daten der ausgewählten Patienten aus den Datenbanken exportieren

	  ; Quickexport ist noch nicht fertig
		If qexporter.ExportRunning {
			qexporter.Stop := true
			GuiControl, QE:, QEExport, % "Export wird`nangehalten"
			cpos := GetWindowSpot(QEhExport)
			ToolTip, % "Der aktuelle Datenexport wird noch beendet.`n               . . . . . Bitte warten . . . . .        ", % cpos.X-30, % cpos.Y-60, 6
			return
		}

		Gui, QE: ListView, QELvExp

      ; Listviewzeilen lesen und Bulkexport starten
		If IsObject(lvRows := QELvGetCheckedRows()) {

			Thread, Interrupt, 0
			qexporter.Stop := false

			GuiControl, QE:, QEExport, % "STOP EXPORT!"
			cpos := GetWindowSpot(QEhExport)

			ToolTip, % "Drücke 'STOP' um den Vorgang `nbeim nächsten Patienten abzubrechen", % cpos.X-30, % cpos.Y-60, 6
			func := Func("ToolTipOff").Bind(6)
			SetTimer, % func, -5000

		  ;Export wird gestartet - ObjBindMethod läßt den neuen Thread unterbrechbar werden
			BulkExport := ObjBindMethod(qexporter, "BulkExport", lvRows)
			SetTimer, % BulkExport, -200

			func := Func("CheckQExporterState")
			SetTimer, % func, -3000

		}



	}
	else If (A_GuiControl = "QEExplorer") 	          	{ ; Verzeichnis des ausgewählten Patienten im Dateiexplorer anzeigen

		If IsObject(lvRows := QELvGetCheckedRows()) {
			If (lvRows.Count() > 0) {
				fpath := lvRows[1].data.names.ExportPath
				fpath := (fpath ~= "^\w:") ? RegExReplace(fpath, "^\w:", stdExportDrive) : stdExportPath "\" fpath
				If InStr(FileExist(fpath), "D")
					Run, % "explorer.exe " q fpath q
				else {
					GSize := GetWindowSpot(QEhExplorer)
					ToolTip, % "Das Verzeichnis für den Patienten (" lvRows[1].names.PatID ") " lvRows[1].names.Name ", " lvRows[1].names.Vorname " ist nicht vorhanden", % GSize.X, % G.Size.Y+GSize.H+5, 5
					func := Func("ToolTipOff").Bind(5)
					SetTimer, % func, -3000
				}

			}
			lvRows := fpath := GSize := func := ""
		}

	}
	else If (A_GuiControl = "QEReload") 		          	{ ; Skript neu laden / starten
		ReloadScript := 1
		gosub QEGuiClose
	}
	else If (A_GuiControl = "QELVExpFilter")           	{	; mit jeder Checkboxänderungen werden die Tabellendaten aufgefrischt
		settings.LVExpFilter 	:= GuiControlGet("QE", "", "QELVExpFilter")
		settings.LVExpFilter2 	:= GuiControlGet("QE", "", "QELVExpFilter2")
		If (settings.LVExpFilter2 && settings.LVExpFilter) {
			GuiControl, QE: -g	, QELVExpFilter2
			GuiControl, QE:	  	, QELVExpFilter2, 0
			GuiControl, QE: +g	, QELVExpFilter2, QEBHandler
			settings.LVExpFilter2 := false
		}
		QEShowTable()
	}
	else If (A_GuiControl = "QELVExpFilter2")         	{	; mit jeder Checkboxänderungen werden die Tabellendaten aufgefrischt
		settings.LVExpFilter 	:= GuiControlGet("QE", "", "QELVExpFilter")
		settings.LVExpFilter2 	:= GuiControlGet("QE", "", "QELVExpFilter2")
		If (settings.LVExpFilter2 && settings.LVExpFilter) {
			GuiControl, QE: -g	, QELVExpFilter
			GuiControl, QE:    	, QELVExpFilter, 0
			GuiControl, QE: +g	, QELVExpFilter, QEBHandler
			settings.LVExpFilter := false
		}
		QEShowTable()
	}
	else If (A_GuiControl = "QELVExpFilter3")         	{	; mit jeder Checkboxänderungen werden die Tabellendaten aufgefrischt
		settings.LVExpFilter 	:= GuiControlGet("QE", "", "QELVExpFilter3")
		QEShowTable()
	}
	else If (A_GuiControl = "QERNEWF")                 	{ ; Export: Daten neu erstellen
		settings.reNewData := GuiControlGet("QE", "", "QERNEWF")
	}
	else If (A_GuiControl = "QENoLData")              	{ ; Export: ohne Labordaten
		settings.noLabData := GuiControlGet("QE", "", "QENoLData")
	}
	else If (A_GuiControl = "QESQLWP")                	{	; Option: NoSQLWrite
		settings.noSQLWrite := GuiControlGet("QE", "", "QESQLWP")
	}
	else If (A_GuiControl = "QEZCrypt")                	{	; Option: Verschlüsselungen für 7zip Archive einschalten
		settings.NoZipEncryption := GuiControlGet("QE", "", "QEZCrypt")
	}
	else If (A_GuiControl = "QESDBG")                 	{	; Option: Debug
		settings.dbg := dbg := GuiControlGet("QE", "", "QESDBG")
	}
	else If (A_GuiControl = "QEStatistic")             	{ ; erstellt eine Exportstatistik
		qexporter.ExportStatistic()
	}
	else If (A_GuiControl = "QEsqlSave")               	{
		QEInCellEditHandler("SAVESQL")
	}
	else If (A_GuiControl = "QEsqlCancel")            	{
		QEInCellEditHandler("CANCELSQL")
	}
	else If (A_GuiControl = "QEBulkCh")               	{	; veränderte Zelldaten übernehmen
		QEBulkDataChange()
	}
	else If (A_GuiControl = "QEZPWDSHow")             	{ ; Verschlüsselungszeichenfolge im Klartext anzeigen
		settings.PWDVisibility := !settings.PWDVisibility
	}
	else If (A_GuiControl = "QEVBAct")                	{ ; automatische Versandbrieferstellung ist ein- bzw- ausschalten
		settings.printLetter := GuiControlGet("QE", "", "QEVBAct")
		settingss.VBControls := ["QEVBIprint", "QEVBIprinter", "QEVBData"]
		for each, ctrl in settings.VBControls
			GuiControl, % "QE: Enable" settings.printLetter, % ctrl
	}
	else If (A_GuiControl = "QEZSelPath")              	{ ;
		FileSelectFolder, ZPath, , 3
		If ZPath
			Guicontrol, QE:, QEZPath, % ZPath
	}
	else If (A_GuiControl = "QECopyTo")                	{ ; kopiert ausgewählte Verzeichnisse ins 7z Packverzeichnis

	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Listview Interaktionen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If (A_GuiControl = "QELvExp")		                  	{


		If (A_GuiEvent = "DoubleClick" && (cRow := A_EventInfo)) {
			;~ ICELLvExp.EditCell(cRow)
			;~ If Pat.ID && (QENewMode := !GuiControlGet("QE", "Enabled", "QENewItem") ? 0 : QENewItem = "Vorbereiten für`nExport" ? 1 : 2) {
				;~ PatSet := ""
				;~ MsgBox, 0x1004, Frage, % "Bestätige mit 'Ja' um den " (QENewMode = 1 ? "neuen" : "geänderten") " Datensatz zu sichern bevor die Eingaben ersetzt werden."
				;~ IfMsgBox, Yes
					;~ PatSet := QESaveItems(Pat)

			;~ ; Standards wiederherstellen, wenn Daten erfolgreich gesichert wurden oder wenn Daten nicht gesichert werden sollten
				;~ If (!IsObject(PatSet) || PatSet.RecordWritten = true)
					;~ QEStandards()  ; ändert Text in "Patient suchen"

			;~ }

		  ;~ ; liest die angeklickte Zeile aus und befüllt die Eingabefelder mit den Werten
			;~ LvExp	:= QEFillElements(cRow)

			;~ Gui, QE: Submit, NoHide
			;~ Pat := QEPatientTxt()
			;~ GuiControl, % "QE: "         	, QESearch			, % "Felder leeren" 	           	; Felder leeren ermöglichen


			;~ return

		}

	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Patienten suchen oder Felder leeren
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	QECtrlItemChange := false
	If (A_GuiControl = "QESearch") || (HotkeyCall ~= "i)(Enter|Tab)" && vfocus ~= "i)(QEPatID|QEPatient)") {

		Pat := QEPatientTxt()
		If (Pat.ID || Pat.Surname || Pat.Prename || Pat.Birth)
			PatSet := QEGetPatSet({"ID":Pat.ID, "Surname":Pat.Surname, "Prename":Pat.Prename, "Birth":Pat.Birth})

		QESearchTxt	:=  GuiControlGet("QE", ""      	, "QESearch")
		NewItem    	:=  GuiControlGet("QE", ""      	, "QENewItem")
		QENewMode 	:= !GuiControlGet("QE", "Enabled"	, "QENewItem") ? 0 : NewItem = "Datensatz anlegen" ? 1 : 2
		SciTEOutput(A_LineNumber ": QESearchTxt: " QESearchTxt "`n      NewItemTxt:  " newItem "`n      surname:     " Pat.surname "`n      prename:     " Pat.prename
										.  "`n      ID:          " Pat.ID ", Patienttxt: " Pat.txt ", NMode: " QENewMode)

	  ; Patientenfelder leeren
		If (QESearchTxt = "Felder leeren") {

		  ; Datensatz neu oder geändert - vor dem Leeren fragen
			If (IsObject(PatSet) && PatSet.Exist = 0) {
				MsgBox, 0x1004, Frage, % "Bestätige mit 'Ja' um " (QENewMode = 1 ? "einen neuen" : "den geänderten") " Datensatz zu sichern bevor die Eingaben entfernt werden."
				IfMsgBox, Yes
					PatSet := QESaveItems(Pat, QENewMode)
			}

			QEStandards()  ; ändert Text in "Patient suchen"
			GuiControl, QE: Enable , QEPatID
			GuiControl, QE: Enable , QEPatient
			GuiControl, QE: Disable, QENewItem
			GuiControl, QE:      	 , QESearch	 , % "Patient suchen"

		}

	  ; Patientendaten suchen
		else If (QESearchTxt = "Patient suchen") {

			if IsObject(PatSet) {

				Gui, QE: Default
				SciTEOutput(A_ThisLabel ": PatSet is object = " IsObject(PatSet) " and PatSet.Exist is " PatSet.RecordExist)

				if (PatSet.RecordExist = 1) {
					GuiControlToolTip("QEPatient", "Ein Datensatz für die PatID (" Pat.ID ")`nwurde bereits angelegt")
					SciTEOutput("Ein Datensatz für die PatID (" Pat.ID ")`nwurde bereits angelegt")
					TrayTip, % scriptTitle, % "Ein Datensatz für die PatID (" Pat.ID ")`nwurde bereits angelegt",  5
					GuiControl, QE: Disable, QENewItem
					GuiControl, QE: Enable , QEPatID
					GuiControl, QE: Enable , QEPatient
					GuiControl, QE: Focus  , QEPatient
				}
				else if (PatSet.RecordExist = 0) {
					GuiControlToolTip("QEPatient", "Ein Datensatz für die PatID (" Pat.ID ")`nkann angelegt werden.")
					SciTEOutput("Ein Datensatz für die PatID (" Pat.ID ")`nkann angelegt werden.")
					GuiControl, QE: Enable      , QENewItem
					GuiControl, QE: Disable    	, QEPatID
					GuiControl, QE: Disable   	, QEPatient
					GuiControl, QE: ChooseString, QERecipient, % "Patient"
					GuiControl, QE: ChooseString, QEMedium	 , % "❔ ungeklärt"
					GuiControl, QE: ChooseString, QERoute  	 , % "❔ ungeklärt"
				}
				else if (PatSet.RecordExist = -1) {
					GuiControlToolTip("QEPatient", "Es trat folgender SQLite-Lesefehler auf:`n" PatSet.ErrorMsg)
					GuiControl, QE: Disable, QENewItem
					GuiControl, QE: Enable , QEPatID
					GuiControl, QE: Enable , QEPatient
					GuiControl, QE: Focus  , QEPatient
				}

			}
			else if !IsObject(PatSet) {
				GuiControl, QE: Disable, QENewItem
				GuiControl, QE: Enable , QEPatID
				GuiControl, QE: Enable , QEPatient
				GuiControl, QE: Focus  , QEPatient
			}

			GuiControl, QE:, QESearch	, % "Felder leeren"

		}

	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Button Vorbereiten zum Exportieren oder Änderungen übernehmen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else if (A_GuiControl = "QENewItem")              	{

		Gui, QE: Default

		Pat     := QEPatientTxt()
		NewItem :=  GuiControlGet("QE", "", "QENewItem")

		If (QENewMode := !GuiControlGet("QE", "Enabled"	, "QENewItem") ? 0 : NewItem = "Datensatz anlegen" ? 1 : 2) {
			MsgBox, 0x1004, Frage, %  "Soll ein Datensatz für '" Pat.Surname ", " Pat.PreName " *" Pat.Birth " [" Pat.ID "]" "' die Daten übernommen werden sollen."
			IfMsgBox, No
				return
		}

		If (Pat.ID) {
			dbg>1 ? SciTEOutput(A_LineNumber ": " Pat.ID ", QENewMode: " QENewMode) : ""
			PatSet := QESaveItems(Pat, QENewMode)
			GuiControl, % "QE: ", QESearch, % (PatSet.EText ? "Felder leeren" : "Patient suchen")
			GuiControl, % "QE: " (PatSet.RecordWritten ? "Disable" : "Enable"), QENewItem
		}

	}

	else if (A_GuiControl = "QESDate")                	{
		QECtrlItemChange := true
		;~ If !RegExMatch(QESdate, "\d{2}\.\d{2}\.\d{4}")
			;~ fdate := Format()
	}


 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Hausarzt - öffnet Dialog um weitere Ärzte anzulegen
 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else if (A_GuiControl = "QERecipient")             	{

		QECtrlItemChange := true

		Gui, QE: Submit, NoHide

		RegExMatch(QERecipient, "\[(?<snr>\d+)\]", b)
		dbg ? SciTEOutput(A_GuiControl ", " QERecipient " = " bsnr " - " (Clipboard := settings.physicans[bsnr].Medium ": " settings.physicans[bsnr].Route)) : ""
		GuiControl, QE: ChooseString, QEMedium, % settings.physicans[bsnr].Medium
		If (QERecipient ~= "i)Hausarzt") {
			recip := dab.AddressBookDialog()
			GuiControl, QE: ChooseString, % "QERecipient", % Recip
		}
		else if RegExMatch(QERecipient, "\[(?<snr>\d+)\]", b) {

			If settings.physicans[bsnr].Medium
				GuiControl, QE: ChooseString, % "QEMedium", % settings.physicans[bsnr].Medium
			else
				GuiControl, QE:, % "QEMedium", % vCtrls.5.value
			If settings.physicans[bsnr].Route
				GuiControl, QE: ChooseString, % "QERoute", % settings.physicans[bsnr].Route
			else
				GuiControl, QE:, % "QERoute", % vCtrls.6.value
		}
		else {
			GuiControl, QE:, % "QEMedium"	, % StrReplace(vCtrls.5.value, "|🖄 ePost")
			GuiControl, QE:, % "QERoute"	, % StrReplace(vCtrls.6.value, "|📡 Telematik")
		}

	}
	else if (A_GuiControl = "QEMedium")               	{

		QECtrlItemChange := true

		Gui, QE: Submit, NoHide
		If (QEMedium ~= "ePost") {
			GuiControl, QE:            		, QERoute, % "|📡 Telematik"
			GuiControl, QE: ChooseString	, QERoute, % "📡 Telematik"
		}
		else {
			GuiControl, QE:          		, QERoute, % StrReplace(vCtrls.6.value, "|📡 Telematik")
			GuiControl, QE: Choose        	, QERoute, 0
		}

		Gui, QE: Submit, NoHide
		RegExMatch(QERecipient, "\[(?<snr>\d+)\]", b)
		If !IsObject(settings.physicans[bsnr])
			settings.physicans[bsnr] := {"MediumO":"", "RouteO":""}
		settings.physicans[bsnr].Medium := QEMedium 	? QEMedium	: settings.physicans[bsnr].Medium
		settings.physicans[bsnr].Route 	:= QERoute  	? QERoute  	: settings.physicans[bsnr].Route

	}
	else if (A_GuiControl = "QERoute")                	{

		QECtrlItemChange := true

		Gui, QE: Submit, NoHide
		If (QERoute ~= "i)Telematik") {
			;~ GuiControl, QE:                   		, QERoute 	, % "|📡 Telematik"
			GuiControl, QE:             	, QEMedium	, % vCtrls.5.value
			GuiControl, QE: ChooseString	, QEMedium	, % "🖄 ePost"
		}

		Gui, QE: Submit, NoHide
		RegExMatch(QERecipient, "\[(?<snr>\d+)", b)
		settings.physicans[bsnr].Medium := QEMedium 	? QEMedium	: settings.physicans[bsnr].Medium
		settings.physicans[bsnr].Route 	:= QERoute  	? QERoute  	: settings.physicans[bsnr].Route

	}
	else if (A_GuiControl = "QEShipment")             	{

		QECtrlItemChange := true

		Gui, QE: Submit, NoHide
		sdate	:= QEShipment
		sdate 	:= Trim(sDate)
		If (sdate ~= "^\d{8}$") {
			GuiControl, QE:, QEShipment, % "|"
			GuiControl, QE:, QEShipment, % "|" SubStr(sdate, 1, 2) "." SubStr(sdate, 3, 2) "." SubStr(sdate, 5, 4)
			GuiControl, QE: Choose, QEShipment, 1
		} else if (sdate ~= "^(\d{1}\.\d{2}|\d{2}\.\d{1})\.\d{4}$") {
			GuiControl, QE:, QEShipment, % "|"
			GuiControl, QE:, QEShipment, % "|" SubStr("00" StrSplit(sdate, ".").1, -1) "." SubStr("00" StrSplit(sdate, ".").2, -1) "." StrSplit(sdate, ".").3
			GuiControl, QE: Choose, QEShipment, 1
		}
	}
	else if (A_GuiControl ~= "i)(QEMarker|QEEraser)") 	{ ; als versendet markieren oder Versanddatum entfernen

		Gui, QE: Submit, NoHide
		QEMarkUnMark((A_GuiControl="QEMarker" ? "mark" : "unmark"), QEShipment)

	}


 ; wurden Daten wurden geändert
	If lastD.haskey(agcontrol := %A_GuiControl%) || QECtrlItemChange {

		Gui, QE: Submit, NoHide
		value := A_GuiControl

		If (lastD[agcontrol] != value) {
			lastD[agcontrol] := Value
			GuiControl, QE:, QENewItem, % "Änderungen`nübernehmen"
			GuiControl, QE: Enable, QENewItem
		}

	}



	HotkeyCall := vfocus := cExp := cpos := ""

return

;}
QECopyFilesToFolder()                                                                 	{ ;--

	; alle notwendigen Daten zusammenstellen
	lvRows    	:= QELvGetCheckedRows()
	pathToCopy 	:= settings.paths.networkpath

	; Netzwerkpfad oder Kopierpfad ist vorhanden
	If !Instr(FileExist(pathToCopy), "D") {
		MsgBox, 0x1000, % Scripttitle, % " Das Zielverzeichnis existiert nicht! `n" pathToCopy
		return
	}

	; Dateiliste erstellen
	If !lvRows.Count() {
		MsgBox, 0x1000, % scriptTitle, % "Sie haben keine Inhalte in der Tabelle ausgewählt."
		return
	}
	else {
		rc := lvRows.Count()
		MsgBox, 0x1004, % scriptTitle, % "Es " (rc=1 ? "wird":"werden") " jetzt die Daten " (rc=1 ? "eines" :"von " rc ) " Patienten nach " pathToCopy " kopiert.", 10
		IfMsgBox, No
		IfMsgBox, Cancel
			return
	}

	files := []

 ; -----------------------------------------------------------------------------
 ; durch die Verzeichnis gehen und Dateien einsammeln
	For each, lvrow in lvRows {

		rowNr  	:= lvRow.row
		rowID  	:= lvRow.rowID
		PatID  	:= lvRow.names.PatID
		bsnr   	:= lvRow.names.Recipient
		recip  	:= SQLiteGetRecipient(bsnr)
		shipment:= lvRow.names.shipment
		shipped := shipment ~= "(\d{8}|\d{2}\.\d{2}\.(\d{4}|\d{2})" ? "\versendet" : ""
		pathBSNR:= bsnr > 3 ? " [" bsnr "]" : ""

		PatPath  	:= recip . pathBSNR . shipped . "\(" PatID ") " cPat.Name(PatID)
		SciTEOutput( "PatID: " PatID " - " PatPath "`nist dir: " (Instr(FileExist(stdExportPath "\" PatPath), "D") ? "ja":"nein") )

	 ; -----------------------------------------------------------------------------
	 ; Versandbriefe erstellen und mit "einpacken"
		If settings.ZipWithLetter && !FileExist(stdExportPath "\" recip . pathBSNR . "\Versandschreiben_" recip " - " A_YYYY A_MM A_DD ".pdf") {
			If !IsObject(letters := CreateShippingLetter(bsnr, A_YYYY A_MM A_DD ))
				MsgBox, % 0x1000, % scriptTitle , % "Das Erstellen der/des Versandbriefe(s) wurde mit einem Fehler beendet: " settings.lastError
			}

	 ; -----------------------------------------------------------------------------
	 ; Verzeichnisse mit Inhalt kopieren
	 ; /R	Überschreibt schreibgeschützte Dateien ohne Rückfrage.
	 ; /C	Setzt den Kopiervorgang auch beim Auftreten eines Fehlers fort.
	 ; /Y	Überschreibt die Zieldatei ohne vorherige Nachfrage.
	 ; /I	Wenn mehrere Dateien kopiert werden und kein Ziel existiert, nimmt XCOPY an, dass das Ziel ein Verzeichnis (und keine einzelne Datei) ist.
	 ; /V	Überprüft jede neue Datei auf ihre Korrektheit.
	 ; /E	Kopiert sämtliche Unterverzeichnisse, unabhängig davon, ob diese leer oder nicht leer sind.
	 ; /L simuliert die Vorgänge

		cmd_XCOPY	:= "XCOPY " . q . stdExportPath "\" PatPath "\*.*" . q . " " . q . PathToCopy . "\" . PatPath . "\" . q . " /Y /C /R /K /I /E /V" (test ? " /L)" : "")
		stdout_XCOPY := StdoutToVar(cmd_XCOPY)

	}

	MsgBox, % 0x1000, % scriptTitle, % "Es wurden die Verzeichnisse von " lvRows.Count() " Patienten nach:`n<< " PathToCopy . "\" . PatPath " >>`nkopiert.", 10

	; Fragen ob der Versandbrief mitgedruckt werden soll
	;settings.ZipWithLetter
}

QEPatAddress()                                                                         	{	;-- Patientenadresse für Postversand anzeigen

	static guihwnd

	If IsObject(lvRows := QELvGetCheckedRows())
		If (lvRows.Count() > 0) {
			PatID 	:= lvRows[1].data.names.PatID
			street 	:= cPat.Get(PatID, "STRASSE")
			nr     	:= cPat.Get(PatID, "HAUSNUMMER")
			plz   	:= cPat.Get(PatID, "PLZ")
			place  	:= cPat.Get(PatID, "ORT")
			postal_txt := cPat.Name(PatID) "`n" street " " nr "`n" plz " " place
			Clipboard := postal_txt
			ClipWait, 2

			Gui, adr: New, HWNDguihwnd -DPIScale
			Gui, adr: Default
			Gui, Font, s10 q5 Normal
			Gui, Add, Edit	, % "xm ym w500 r4 ReadOnly "     	, % postal_txt
			Gui, Font, s9 q5 Normal
			Gui, Add, Text	, % "xm y+5 w400 hwndhCtrl "      	, % "Der Adresstext kann mit Strg+V`nin einer anderen Anwendung eingefügt werden."
			Gui, Font, s9 q5 Normal
			Gui, Add, Button, % "x+5 yp-0 gadrGuiClose"       	, % "Fenster`nschließen"
			Gui, Show, AutoSize, % "Adresse von (" PatID ") " cPat.Name(PatID, true)

			SendInput, {Tab} ; die automatische Textauswahl des letzten Editsteuerelementes verhindern
			ControlFocus, % "ahk_id " hCtrl

			;automatisch Schließen nach 5 min
			SetTimer, adrGuiClose, % -1*1*60*1000

		}


return
adrGuiClose:
	SetTimer, adrGuiClose, Delete
	Gui, adr: Destroy
return
}

QEFileExplorer()                                                                      	{	;-- Dateien:     Verzeichnis des ausgewählten Patienten im Dateiexplorer anzeigen

	global QEhExplorer

	If IsObject(lvRows := QELvGetCheckedRows()) {
		If (lvRows.Count() > 0) {
			fpath := lvRows[1].data.names.ExportPath
			fpath := (fpath ~= "^\w:") ? RegExReplace(fpath, "^\w:", stdExportDrive) : stdExportPath "\" fpath
			If InStr(FileExist(fpath), "D")
				Run, % "explorer.exe " q fpath q
			else {
				GSize := GetWindowSpot(QEhExplorer)
				ToolTip, % "Das Verzeichnis für den Patienten (" lvRows[1].names.PatID ") " lvRows[1].names.Name ", " lvRows[1].names.Vorname " ist nicht vorhanden", % GSize.X, % G.Size.Y+GSize.H+5, 5
				func := Func("ToolTipOff").Bind(5)
				SetTimer, % func, -3000
			}

		}

	}

}

QELetterChoosed(p*)                                                                   	{


	namesBSNR := GuiControlGet("QE", "", "QEVBLBPrint1")
	letterNr  := GuiControlGet("QE", "", "QEVBLBPrint2")

	letterPath := settings.Recips[namesBSNR].docs[letterNr]
	settings.activeLetterPath := InStr(letterPath, "keine") ? "" : letterPath
	;~ SciTEOutput("lP: " letterPath)
	return

}

QEPrinterDialog(filepath, reprints:=true)                                             	{ ;-- Gui:         Druckerauswahl Dialog

	static fileToPrint, printagain
	static PrnList, PrnNext, PrnCncl

	fileToPrint := filepath
	printagain := reprints

	Gui PRn: New	, % "hwndhMsg +AlwaysOnTop +Owner" hQE
	Gui PRn: Font	, % settings.fSize.s12, Calibri
	Gui PRn: Add	, text     	, % "xm ym w400", % "Es ist kein Drucker ausgewählt.`nWählen Sie einen Drucker aus der Liste."

	Gui PRn: Add	, DDL      	, % "y+10 w320 vPrnList gPRnHandler"  	, % GetPrinters()
	GuiControl, Prn: Choose, PrnList, 1
	Gui PRn: Add	, Button  	, % "y+20 vPrnNext gPrnHandler"  	, % "Weiter"
	Gui PRn: Add	, Button  	, % "X+50 vPrnCncl gPrnHandler" 	, % "Abbruch"
	Gui PRn: Show	, AutoSize 	, % "PDF Reader benötigt"

return

PrnHandler:
	Critical, Off
	Critical
	Gui, Prn: Submit, NoHide

	If (A_GuiControl = "PrnList") {
		GuiControl, QE: ChooseString, QEVBIprinter, % PrnList
	}
	else if(A_GuiControl = "PrnNext")  {
		GuiControl, QE: ChooseString, QEVBIprinter, % PrnList
		QEPrinterDialog(fileToPrint, printagain)
	}
	else if (A_GuiControl = "PrnCancel") {
		PrnGuiClose:
		PrnGuiEscape:
		Gui, Prb: Destroy
	}
return
}

QELetterPrint(filepath:="", reprints:=true)                                           	{	;-- Drucken:     Versandbrief auf Drucker ausgeben/drucken

	global QEVBIprinter, QE

	Gui, QE: Default
	Gui, QE: Submit, NoHide

	printer_name 	:= GuiControlGet("QE", "", "QEVBIprinter")
	printer_name 	:= !printer_name ? QEVBIPrinter : printer_name
	fileToPrint 	:= filepath ~= "[\\\/\:\pL\.\s]+" ? filepath : settings.activeLetterPath
	fileToPrint 	:= stdExportPath "\" StrReplace(fileToPrint, stdExportPath)

	If !printer_name {
		If !reprints
			return

		func_call := Func("QEPrinterDialog").Bind(fileToPrint, reprints)
		SetTimer, % func_call, -1
		Sleep 200
		while WinExist("PDF Reader benötigt ahk_class AutoHotkeyGUI1")
			Sleep 50
		return

	}


	Progress, m2 b fs11 zh0 w1200 cFFFF00, % "sende '" fileToPrint "' an Drucker '" printer_name "'"

	FileGetAttrib, fatrib, % ()
	If InStr(fatrib, "A") {

		IF !reprints
			return

		MsgBox, 0x1004, % scriptTitle, % "Diese Datei wurde bereits ausgedrckt. Nocheinmal drucken?"
		IfMsgBox, No
		{
			Progress, Off
			return
		}

	}

	For PrnIndex, prn in settings.printer
		 If prn.name = printer_name {
			pdfprinterOpts := prn.Options
			break
		}
	pdfprinterOpts := !pdfprinterOpts ? "fit,paper=A4" : pdfprinterOpts

	SplitPath, fileToPrint,, pdfPath, fExt, pdfFileOutName
	fPDF:= pdfPath "\" pdfFileOutName ".pdf"
	If (fext = "html" && !FileExist(fPDF)) {
		cmdline := "msedge.exe --headless --disable-pdf-tagging --no-pdf-header-footer --print-to-pdf=" q fPDF q " " q "file://" fileToPrint q
		RunWait, % cmdline
	}

	If FileExist(fPDF)
		fileToPrint := fPDF
	else
		return


	If InStr(GetPrinters(), printer_name) && FileExist(fileToPrint) {
		Smtra_PrintPDF(printer_name, pdfprinterOpts, fileToPrint)
		FileSetAttrib, +A, % fileToPrint
	}

	progress, Off

return
ProgressOFF:
Progress, Off
return
}

QELetterList(p*)                                                                      	{	;-- GuiControl:  Versandbriefliste für Steuerelement bereitstellen

	if IsObject(p)
	SciTEOutput("p: " cJSON.Dump(p, 1))

	Gui, QE: Default
	Gui, QE: Submit, NoHide

	namesBSNR := GuiControlGet("QE", "", "QEVBLBPrint1")
	;~ SciTEOutput("LB: " namesBSNR ", " settings.Recips[namesBSNR].LBList)

	If settings.Recips.Count() {

		GuiControl QE: Enable     	, QEVBPrintNow
		GuiControl QE: Enable     	, QEVBLBPrint2
		GuiControl QE:            	, QEVBLBPrint2, |
		GuiControl QE:             	, QEVBLBPrint2, % settings.Recips[namesBSNR].LBList
		GuiControl QE: Choose     	, QEVBLBPrint2, 1

		If InStr(settings.Recips[namesBSNR].LBList, "keine") {
			GuiControl QE: Disable     	, QEVBPrintNow
			GuiControl QE: Disable     	, QEVBLBPrint2
		}
	}


}

QELetterChoiceTS(p*)                                                             	     	{	;-- GuiControl:  startet QELetterChoice als SetTimerThread

		static recall:=0


	; Tabellenaufbau hat Vorrang
		If settings.QEShowTable || settings.QELetterChoice {
			wp := GetWindowspot(hQE)
			recall += 1
			Progress, m2 b fs18 zh0, % "[" recall "] Warte auf Abschluß der " (settings.QEShowTabelle ? "Tabellenbefüllung." : "Versandbrieferfassung.") "..."
			func_call := Func("QELetterChoiceTS")
			SetTimer, % func_call, -1000
			return
		}

		recall := 0

	; Empfänger Namen laden
		If !QEDB.GetTable(sql.Recipients.GetTable, RecipTable) {
			SQLiteErrorWriter("Tabelle konnte nicht gelesen werden", QEDB.ErrorMsg, QEDB.Errorcode, true, sql.Exports.GetTable)
			return
		}
		If !RecipTable.Rows.Count()
			return

  ; nach QELetterChoice muss er aber dennoch springen,
	;	da die Funktion das DDL Steuerelement auch wieder leert
		func_call := Func("QELetterChoice").Bind(RecipTable.Rows)
		SetTimer, % func_call, -1000

}

QELetterChoice(RecipTableRows)                                                         	{ ;-- GuiControl:  zeigt Versandbriefe in einer DDL an damit man diese auswählen und ausdrucken kann

	static ddlID

	settings[A_ThisFunc] := 1

	ZeigsMir("...erfasse Versandbriefe", 0, RecipTableRows.Count())
	Gui, QE: Default

	LBPhysican := ""
	settings.Recips := Object()

	For rNR, Recip in RecipTableRows {

		ZeigsMir("...erfasse Versandbriefe", rNr, RecipTableRows.Count())

		If (Recip.7 > 3) {

			ANRs  		:= RegExReplace(Recip.1	, "i)\s*(Frau)\s*"	, "Fr. ")
			ANRs  		:= RegExReplace(ANRs  , "i)\s*(Herr)\s*"	, "Hr. ")
			namesBSNR := Recip.3 ", " Recip.4 " [" Recip.7 "]"
			path			:= Trim(ANRs (Recip.2 ? RegExReplace(Recip.2, "i)\s*med.\s*") " " : "") namesBSNR)

			; Empfänger             	 := Festplattenverzeichnis
			settings.Recips[namesBSNR]  := {"path"	: path
																		, "docs"	: []
																		, "LBList": ""}
			LBPhysican .= (LBPhysican?"|":"") namesBSNR

		}
	}

	for namesBSNR, Recip in settings.Recips {

		If InStr(FileExist(stdExportPath "\" Recip.path), "D")
			Loop, Files, % stdExportPath "\" Recip.path "\*.*"
			{

			 ; Dateiausschlüsse
				If (!A_LoopFileSize || !InStr(A_LoopFileName, "Versand"))
					continue

				If A_LoopFileExt NOT IN pdf,html 		;~= "i)(pdf|html)"
					continue

				filepath := StrReplace(A_LoopFileFullPath, stdExportPath "\")
				filefound := false
				For each, fobj in Recip.docs
					If (fobj.path = filepath) {
						filefound := true
						break
					}
				If !filefound {
					RegExMatch(A_LoopFileName, "i)Versand.+?(?<Y>\d+)-(?<M>\d+)-(?<D>\d+)", d)
					Recip.docs.Push(filepath)
					Recip.LBList 	.= (Recip.LBList?"|":"") dD "." dM "." dY "`t(." A_LoopFileExt ")"
				}

			}
	}


	ZeigsMir("Erfassung v. Versandbriefen abgeschlossen")

; Versandbriefe das erstemal befüllen ;{

	GuiControl QE: Disable	, QEVBPrintNow
	GuiControl QE: Disable	, QEVBLBPrint1
	GuiControl QE: Disable	, QEVBLBPrint2


	If settings.Recips.Count() {


		GuiControl QE: Enable     	, QEVBLBPrint1
		GuiControl QE:            	, QEVBLBPrint1, |
		GuiControl QE:            	, QEVBLBPrint1, % LBPhysican

		RecipListShow:= 0

	 ; Zeige einen Empfänger mit Versandschreiben an
		For namesBSNR, Recip in settings.Recips {

			;~ SciTEOutput(namesBSNR ": " cJSON.Dump(Recip, 1))

			If Recip.LBList && !RecipListShow {

				RecipListShow:= 1
				GuiControl QE: ChooseString	, QEVBLBPrint1, % namesBSNR
				GuiControl QE: Enable     	, QEVBPrintNow
				GuiControl QE: Enable     	, QEVBLBPrint2
				GuiControl QE:            	, QEVBLBPrint2, |
				GuiControl QE:             	, QEVBLBPrint2, % Recip.LBList
				GuiControl QE: Choose     	, QEVBLBPrint2, 1
				settings.activeLetterPath := Recip.docs[1]
			}

			If !settings.Recips[namesBSNR].docs.Count() {
				settings.Recips[namesBSNR].LBList := "keine Versandbrief(e) vorhanden"
			}

		}

	}


	;}

	settings[A_ThisFunc] := 0

return
}

QESetPropVersand(p*)                                                                   	{	;-- GuiControl:  schaltet die Interaktion mit der Gruppe ein oder aus


	settings.printLetter := GuiControlGet("QE", "", "QEVBAct")
		for each, ctrl in settings.VBControls
			GuiControl, % "QE: " (settings.printLetter ? "Enable":"Disable"), % ctrl

}

QETogglePWDVisibility(p*)                                                              	{ ;-- GuiControl:  Passwortsichtbarkeit umschalten

	global QEhZPWDH, QEhZPWDV

	Gui, QE: Default

	If !(p.1 ~= "i)(noUserButtonClick|NoZipEncryption)")
		settings.PWDVisibility := !settings.PWDVisibility

	If (p.1 ~= "i)(noUserButtonClick|UserButtonClick)") {

		GuiControl, QE: Focus, QELvExp
		vis := settings.PWDVisibility

		isVisibleH  	:= ControlHasStyle(QEhZPWDH, "VISIBLE")
		isVisibleV  	:= ControlHasStyle(QEhZPWDV, "VISIBLE")
		isDisabledH 	:= ControlHasStyle(QEhZPWDH, "DISABLED")
		isDisabledV 	:= ControlHasStyle(QEhZPWDV, "DISABLED")
		;~ SciTEOutput("vis: " vis ", QEPWDH ist " (isVisibleH ? "Sichtbar" : "Unsichtbar") " und ist " (isDisabledH ? "Disabled" : "Enabled"))
		;~ SciTEOutput("vis: " vis ", QEPWDV ist " (isVisibleV ? "Sichtbar" : "Unsichtbar") " und ist " (isDisabledV ? "Disabled" : "Enabled"))

		GuiControl, QE: Text    , % "QEZPWDShow", % (vis ? "👀  " : "  👀")
		GuiControl, QE: Hide    , % (vis ? "QEZPWDH": "QEZPWDV")
		GuiControl, QE: Disable , % (vis ? "QEZPWDH": "QEZPWDV")
		GuiControl, QE: Show    , % (vis ? "QEZPWDV": "QEZPWDH")
		GuiControl, QE: Enable 	, % (vis ? "QEZPWDV": "QEZPWDH")
		GuiControl, QE: Focus  	, % (vis ? "QEZPWDV": "QEZPWDH")

		SendInput, {Left}{End}

	}

	If (p.1 = "NoZipEncryption") {
		noenc := settings.NoZipEncryption := GuiControlGet("QE", "", "QEZCrypt")
		GuiControl, % "QE: " (noenc ? "Disable" : !settings.PWDVisibility ? "Enable" : "Disable"), QEZPWDH
		GuiControl, % "QE: " (noenc ? "Show"    : !settings.PWDVisibility ? "Show"   : "Hide")	 , QEZPWDH
		GuiControl, % "QE: " (noenc ? "Disable" :  settings.PWDVisibility ? "Enable" : "Disable"), QEZPWDV
		GuiControl, % "QE: " (noenc ? "Hide"    :  settings.PWDVisibility ? "Show"   : "Hide")	 , QEZPWDV
		GuiControl, % "QE: " (noenc ? "Disable" : "Enable"), QEZPWDShow
		GuiControl, QE: Focus  	, % "QEZPWD" (settings.PWDVisibility ? "V" : "H")
	}

}

QEMarkUnMark(cmd, sdate:="", test:=0)                                                  	{	;-- als versendet/unversendet markieren

	global vTChanges

	If !IsObject(vTChanges)
		vTChanges := Object()

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

	if (cmd == "unmark")
		sDate := DBsDate := ""
	if (cmd == "mark") {

	 ; inkorrektes Datumsformat oder kein Datum
		If (!sdate ~= "(\d{2}\.\d{2}\.\d{4})" || StrLen(sdate)=0) {
			MsgBox, 0x1000, % scriptTitle, % "Das Versanddatum hat nicht das richtige Format", 4
			return
		}

		; DD.MM.yyyy -> yyyyMMDD
		DBsDate := ConvertToDBASEDate(sDate)

	}

 ; Auffordern zum Speichern von Tabellenänderungen
	If (vTChanges.Count() > 0) {
		MsgBox, 0x1004, % scriptTitle, % "Es gibt noch " vTChanges.Count() " nicht gespeicherte Änderungen in der Tabelle. "
																. 	 "Diese müssen zuvor in die Datenbank geschrieben werden.`n`n"
															  . 	 "Drücken Sie auf 'Ja' um die Änderungen zu speichern.`n"
																. 	 "Mit 'Nein' werden keine Änderungen vorgenommen."
		IfMsgBox, No
			return

		QEInCellEditHandler("SAVESQL")

	}

  ; ausgewählte Zeilen auslesen
	sameShipment := fmvDirEquals := 0
	lvRows := QELvGetCheckedRows()
	If !lvRows.Count() {
		MsgBox, 0x1000, % scriptTitle, % "Sie haben keine Inhalte in der Tabelle ausgewählt."
		return
	}

	GuiControl, QE: -ReDraw, QELvExp

	For each, lvrow in lvRows {

		rowNr  	:= lvRow.row
		rowID  	:= lvRow.rowID
		PatID  	:= lvRow.names.PatID
		bsnr   	:= lvRow.names.Recipient
		shipment:= lvRow.names.shipment

		dbg ? SciTEOutput(A_ThisFunc "() | row: " rowNr ", PatID: " PatID) : ""

	  ; - - - - - - - - - - - - - - - - - - - - - -
		; Versandbrief erstellen
	  ; - - - - - - - - - - - - - - - - - - - - - -
		If (cmd = "mark" && settings.sendletter && bsnr~="\d+" && bsnr>3 && DBsDate ~= "\d{8}")
			res := CreateShippingLetter(bsnr, DBsDate)


	  ; - - - - - - - - - - - - - - - - - - - - - -
	  ; Versanddatum auslesen
	  ; - - - - - - - - - - - - - - - - - - - - - -
		LV_GetText(old_sDate  	, rowNr, 9)          		; sichtbare Zelle
		LV_GetText(old_DBsDate	, rowNr, 15)           	; nicht sichtbar (Sortierungshilfe)

		; altes und neues Versanddatum sind identisch, dann kein Änderungsobjekt anlegen
		If (cmd == "mark" && old_DBsdate = DBsDate) {
			sameShipment += 1
		}
		; Änderungen zum Speichern im vTChanges Objekt ablegen
		else {

			If !IsObject(vTChanges[PatID])
				vTChanges[PatID] := Object()
			vTChanges[PatID]["Shipment"] 	:= {"PatID"    	: PatID
																			, "rowNR"    	: rowNr
																			, "rowID"   	: rowID
																			, "colNr"   	: 9
																			, "cellText"	: DBsDate
																			, "cellTextO"	: old_DBsdate
																			, "writeEmptyValues": true}

		}

		; die Zelleninhalte können dennoch geändert werden
		LV_Modify(rowNr, "Col9"	 , sDate)
		LV_Modify(rowNr, "Col15" , DBsDate)

	  ; - - - - - - - - - - - - - - - - - - - - - -
	  ; Exportverzeichnis festlegen
	  ; - - - - - - - - - - - - - - - - - - - - - -
		If (oldpath := SQLiteGetCell("EXPORTS", "Exportpath", "PatID",  PatID)) {

			; oldpath muss den kompletten Pfad darstellen
			oldpath := stdExportPath "\" StrReplace(oldpath, stdExportPath "\")

			; "\versendet" - Unterverzeichnis aus oldpath entfernen
			sqlpath := RegExReplace(oldpath, "i)(Patient|\])\\versendet\\\(", "$1\(")

			; falls jetzt als 'Unversendet' markiert, dann nimm den sqlpath string
			; im anderen Fall wird ein '\versendet' in den Verzeichnisstring eingefügt
			newpath := cmd = "unmark" ? sqlpath : RegExReplace(sqlpath, "i)(Patient|\])\\\(", "$1\versendet\(") ; versendet erneut Einfügen bei cmd="mark"

			; Die Gültigkeit des Pfade wird später beim Speichern überprüft.

		}
		; noch kein gespeichertes Verzeichnis
		else {
			bsnr := SQLiteGetCell("EXPORTS", "Recipient", "PatID",  PatID)
			recipient := bsnr>3 ? SQLiteGetRecipient(bsnr) : "Patient"
			newPath := stdExportPath "\" recipient (cmd == "unmark" ? "" : "\versendet") "\(" PatID ") " cPat.NAME(PatID, false)
		}

		dbg ? SciTEOutput(A_ThisFunc "() | newPath: " newPath ) : ""

		; - - - - - - - - - - - - - - - - - - - - - -
	  ; alter und neuer Pfad sind nicht identisch,
		; dann einen Eintrag im Änderungsobjekt anlegen
		; - - - - - - - - - - - - - - - - - - - - - -
		if (oldPath != newPath) {
			if !IsObject(vTChanges[PatID])
				vTChanges[PatID] := Object()
			vTChanges[PatID]["ExportPath"] 	:= {"PatID": PatID, "rowNR":rowNr, "rowID": rowID, "colNr":11, "cellText":newPath, "cellTextO":oldPath}
			; Exportpfad Zelle in der Listview ändern
			LV_Modify(rowNr, "Col11" , StrReplace(newPath, stdExportPath "\"))
			dbg ? SciTEOutput(A_ThisFunc "() | Col11 - Path row" rowNr " geändert"  ) : ""
		}
		else
			fmvDirEquals += 1

	}

	GuiControl, QE: +ReDraw, QELvExp

	; - - - - - - - - - - - - - - - - - - - - - -
	; Fehlerbehandlung und Anzeige
	; - - - - - - - - - - - - - - - - - - - - - -
	if fmvDirError
		msg :=  "Es traten Fehler beim Verschieben von Verzeichnissen auf:`n" RTrim(fmvDirError, "`n")
	if sameShipment
		msg .= (msg ? "`nzusätzlich wurden übereinstimmende Versandtage gefunden (" sameShipment ")":"" )
	if msg
		MsgBox, 0x1000, % scriptTitle, % msg

	settings.saveTblChanges := true
	res := QEInCellEditHandler("SAVESQL")


return [fmvDirEquals, res.1, res.2, res.3, sameShipment]
}

QELVScrollHk(param*)                                                                   	{	;-- Listview: 	 seitenweise nach oben oder unten scrollen


	row := 0
	lkey := param.1
	lkeyHwnd := param.2

	if (lkey ~= "i)(End|Home)") {
		ControlSend,, % "{LControl Down}{" lkey "}{LControl Up}", % "ahk_id " lkeyHwnd
	return
	}

	Gui, QE: Default
	Gui, QE: ListView,QELVExp

	maxRows	  	:= LV_GetCount()
	topIndex  	:= LV_EX_GetTopIndex(QEhLvExp)
	scrollVPos	:= topIndex + ((lkey="PgUp" ? -2 : 2) * settings.CounterPage)
	scrollVPos  := scrollVPos < 1 ? 1 : scrollVPos > maxRows ? maxRows : scrollVPos

	GuiControl, QE: -ReDraw	, QELvExp
	LV_Modify(scrollVPos	, "Vis")
	GuiControl, QE: +ReDraw, QELvExp

}

QELVScroll(param*)                                                                     	{	;-- Listview: 	 Schnellsuchhilfe über Steuerelemente
 ; thishotkey

	global QEhLvExp

	static letterfound := []
	static letterpos := ""
	static lastRow := 0

	CoordMode, ToolTip, Window

	Gui, QE: Default
	Gui, QE: ListView, QELVExp

	row := 0

	If (param.1 = "UserInput") {

		userinput := true

		If lastRow {
			LV_Modify(lastRow      	, "-Focus -Select")
			gosub machaus
		}

		ltext := GuiControlGet("QE", "", "QEFEdit")

		If (StrLen(ltext) > 1) {
			letterpos := SubStr(ltext, 1, StrLen(ltext)-1)
			lkey := SubStr(ltext, StrLen(ltext), 1)
		}
		else {
			letterpos := ""
			lkey := ltext
		}


	}
	else {

		lkey := param.1
		lkeyHwnd := param.2

		; Position des Steuerelementes finden
		wp := GetWindowSpot(hQE)
		kp := GetWindowSpot(lkeyHwnd)
		lkeyX := kp.X - wp.CX
		lkeyY := kp.Y - wp.CY

	}

	GuiControl, QE: -ReDraw, QELvExp
	LV_ModifyCol(2, "CaseLocale ") ; "Sort Asc")
	LV_ModifyCol(2, "Sort Asc ")

; es looped los
	Loop, % (LoopRows := LV_GetCount() - lastRow) {

		LV_GetText(cellText, (thisRow := lastRow + A_Index), 2)

		letterstoMatch := letterpos . lkey
		If (SubStr(cellText, 1, StrLen(letterpos)+1) = lettersToMatch) {

			lettersmatched 	:= true
			lastRow       	:= thisRow

		 ; wenn die Funktion nicht über das Schnelleingabefeld aufgerufen wurde
			If !userinput {

				letterpos	.= lkey

			; zeigt an was bisher gedrückt und gefunden wurde
				GuiControl, QE:, QELetters, % letterpos

			; gedrückten Buchstaben etwas Herausziehen
				AllreadyMoved := false
				For kIdx, k in letterfound
					If (k.lkey = lkey) {
						AllreadyMoved := true
						break
					}
				If !AllreadyMoved {
					letterfound.InsertAt(1, {"lkey":lkey, "row":thisRow, "val":celltext, "hwnd":GetHex(lkeyHwnd), "LastY":lkeyY, "LastX":lkeyX})
					SetWindowPos(lkeyHwnd, lkeyX, lkeyY-6, kp.W, kp.H)				; unabhängig von DPISkalierungen einsetzbar
					lkp := GuiControlGet("QE", "Pos", lkeyHwnd)
					kp 	:= GetWindowSpot(lkeyHwnd)
				}

				SetTimer, machaus, -8000

			}

			break

		}


	}

  ; bringt die Zeile vertikal in die Mitte der Ansicht
	If lettersmatched {
		scrollVPos  	:= LV_EX_GetTopIndex(QEhLvExp)              	; erste sichtbare Zeile in der Listview
		endscrollVPos := thisRow + ((scrollVPos-thisRow <= 0 ? 1 : -1) * settings.CounterPage//2)         	; gefundene Zeile + halte Anzahl Listviewzeilen
		LV_Modify(endscrollVPos	, "Vis")
		SendMessage, 0x1013, % endscrollVPos,,, % "ahk_id " QEhLvExp
		LV_Modify(endscrollVPos	, "-Focus -Select")
		LV_Modify(thisRow      	, "Focus Select")
		scrollVPos  	:= LV_EX_GetTopIndex(QEhLvExp)              	; erste sichtbare Zeile in der Listview
	}

	GuiControl, QE: +ReDraw, QELvExp


Return
machaus:

	SetTimer, machaus, Off
	letterpos := ""
	lastRow := 0
	while letterfound.Count() {
		k := letterfound.Pop()
		kp := GetWindowSpot(k.hwnd)
		SetWindowPos(k.hwnd, k.LastX, k.LastY, kp.W, kp.H)
	}
	GuiControl, QE:, QELetters, % ""

return

}

QE7Zip(outbasePath:="", mode:=1)                                                      	{	;-- 7zip:     	 packt ausgewählte Datenordner als 7zip Archiv zusammen

	; mode = 1 	- alle Verzeichnisse welche Ausgewählt wurden werden gezippt
	; 		 2	- nur versandfertige und noch nicht versendete Daten  werden gezippt
	/*
	RunWait "7z.exe a package.7z $DATA banner.png dialog.html -m0=BCJ2 -m1=LZMA:d25:fb255 -m2=LZMA:d19 -m3=LZMA:d19 -mb0:1 -mb0s1:2 -mb0s2:3 -mx", "$TEMP"
	RunWait "7z.exe a -tzip package.zip SciTE ReadMe.txt", "$TEMP"
	RunProgram="$DATA\InternalAHK.exe /CP65001 $DATA\$SETUP"
	*/

	If !FileExist(ziplibPath := Addendum.Dir "\lib\dll\64bit\7-zip64.dll") {
		MsgBox, 0x1000, % "Addendum - " StrReplace(A_ScriptName, ".ahk")
									, % "Zum Ausführen der Funktion ist eine 7zip Datei im Addendum Verzeichnis notwendig. Diese Datei konnte nicht gefunden werden. (\lib\dll\64bit\7-zip64.dll)."
		settings[A_ThisFunc] := 0
		return -99
	}

	If !InStr(FileExist(outbasePath), "D")
		outbasePath := stdExportPath

	settings.ZipWithLetter := GuiControlGet("QE",, "QEZLetter")

	If dbg {
		SciTEOutput(" ")
		SciTEOutput("**************************************************************************************************")
		SciTEOutput("**************************************************************************************************")
		SciTEOutput("[01] outbasePath: " outbasePath)
	}

  ; ---------------------------------------------------------------------------------------------------------
  ; Ermitteln der ausgewählten Zeilen
  ; ---------------------------------------------------------------------------------------------------------
	If !IsObject(lvRows := QELvGetCheckedRows("QEMoveRows")) {
		SciTEOutput(A_ThisFunc "(): keine Zeile gewählt" )
		settings[A_ThisFunc] := 0
		return -1
	}


  ; ---------------------------------------------------------------------------------------------------------
  ; Zustimmung des Nutzers zur Menge und Autfteilung der Daten einholen
  ; ---------------------------------------------------------------------------------------------------------;{
	confirm := "",
	For each, lvRow in lvRows
		If (PatID := lvRow.PatID) {
			confirm .= (confirm ? "`n":"") "[" PatID "] " SubStr(cPat.Name(PatID), 1, 40) (StrLen(cPat.Name(PatID))>40 ? "...":"") " *" cPat.GEBURT(PatID, true)
			If !IsObject(lRecipients[lvRow.names.Recipient])
				lRecipients[lvRow.names.Recipient] := [lvRow.names.RecipientName, 1]
			else
				lRecipients[lvRow.names.Recipient].2 += 1
		}
	If !confirm
		return

	MsgBox, 0x1004, % scriptTitle, % "Sollen die Daten von " (lvRows.Count()>1 ? lvRows.Count() :"einem") " Patienten`n"
																	. "in ein 7zip Archiv gepackt werden?`n"
	IfMsgBox, No
		return
	IfMsgBox, Cancel
		return

	/*
	If (lRecipients.Count()>1) {

		For lrecipient, lrecip in lRecipients
			lRecipientNames .= (lRecipientNames ? "`n" : "") . (lrecip.1=1 ? "Patient" : lrecip.1= 2 ? "Angehöriger" : lrecip.1 " [" lrecipient "]") . "  " . lrecip.2 "x"

		MsgBox, 0x1004, % scriptTitle, % "Es sind mehrere Empfänger in der Auswahl vorhanden:`n"
										. "-------------------------------------------------------------------------`n"
										. confirm
										. "`n-------------------------------------------------------------------------`n"
										. "Sollen mehrere 7zip Archive erstellt werden. Es würden die oben benannten Pakete`n"
										. "erstellt werden. Versehen jeweiligen mit dem Namen des Empfängers."
		IfMsgBox, No
			zipToOne := false
		IfMsgBox, Cancel
			zipToOne := false

	}
	*/
	;}


  ; ---------------------------------------------------------------------------------------------------------
  ; Verzeichnisse komprimieren
  ; ---------------------------------------------------------------------------------------------------------;{

		; 7zip Einstellungen
		; - - - - - - - - - - - - - - - - - - - - - - - - ;{
		; -spe	Eliminate the duplication of the root folder for extract archive command
		; -spf	Use the fully qualified file paths
		;	-scs	charset für Listfiles z.B. -scsUTF-8
		;	-scc	charset für die Console z.B. -scsUTF-8
		;	-stl	verwendet den Zeitstempel der zuletzt geänderten Datei
		; -mcp	codepage für das Archive Standard ist ANSI Unicode ist so -mcp=65001 ?
		; 7z a archive.7z @listfile.txt -scsUTF.8
			VolumeSize    		:= "100M"
			compressType     	:= "7z"
			compressLevel   	:= 9
			compressOpts     	:= "" ;"-m0=LZMA2 -mfb=64 -md=32m -ms=on -mmt=on" 		; LZMA2 Verschlüsselung,
			compressOpts     	.= (CompressType = "7z" ? " -mhe" : "")             	; EncryptFileNames - sollte bei Patientendaten immer an sein
			stream 		      	:= "p1"
			excludeFile     	:= "*.db" ;"*.db"
			workingdir  			:= A_Temp
			overwrite 	    	:= 1
			recurse   	    	:= 0
			hide  	   	    	:= 0
			password        	:= !settings.NoZipEncryption ? StringDecoder() : ""
			zip             	:= Object()
			initZipFile     	:= false
			zipFilePathLast 	:= ""
			zipexports       	:= 0
			zipToOne        	:= 1
			sOpts             := "-spf2" ;"-scsDOS -sccDOS -bb0"
			;}


		; Dateien erstellen, die 7zip dll sollte im
		; Arbeitsverzeichnis (outbasePath) liegen um relative Pfade zu ermöglichen
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
			fileList := FileOpen(outbasePath "\7zList.txt", "w", "UTF-8")
			If !FileExist(ziplipOutPath := outbasePath "\7-zip64.dll") {
				FileCopy, % ziplibPath, % ziplipOutPath
				If !FileExist(ziplipOutPath) {
					MsgBox, 0x1000, % "Addendum - " StrReplace(A_ScriptName, ".ahk")
												, % "Für eine korrekte Ausführung der Funktion muss die 7zip-64.dll im Arbeitsverzeichnis (" outbasepath ") vorliegen. "
												.	 "Die Datei konnte nicht in dieses Verzeichnis kopiert werden."
					return
				}
			}
		;}


		Gui, QE: ListView, QELvExp


	; -------------------------------------------------------------------------------
	; Zeilenweise auswerten
	; -------------------------------------------------------------------------------
		For each, lvRow in lvRows {


	; -----------------------------------------------------------------------------
	; Pfad festlegen
		zipFilePath    	:= StrReplace(lvRow.names.Exportpath, stdExportPath "\")
		zipFileFullpath := stdExportPath "\" zipFilePath
		RegExMatch(zipFileFullpath, "(?<Folder>\(\d+\).+)$", Pat)
		SplitPath, zipFilepath, , zipDir, zipExt, zipNameNoExt


	; -----------------------------------------------------------------------------
	; dieser Abnschnitt wird einmal im Falle eines getrennten 7zip File
	; oder für jede 7z Datei wird ein zip-Objekt angelegt

		; Empfänger als Key und Dateinamen erstellen
		If !zipToOne || (zipToOne && initZipFile) {
			recip := lvRow.names.Recipient <= 2 ? "Patient" : lvRow.names.Recipient
			physicanfolder:= lvRow.names.RecipientName " [" recip "]"
			physicanname	:= RegExReplace(lvRow.names.RecipientName, "^.+?\.\s*.+\.\s*")
			outFilePath 	:= outbasePath "\" physicanfolder "\" A_YYYY "-" A_MM "-" A_DD "_" physicanname
		}

		else if (zipToOne && !initZipFile) {

			recip := "Patientendaten"
			zip[recip] := { "pwd"          	: password
									,  "Files"         	: []
									,  "outfilePath"  	: (outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z")
									,  "outpassPath"  	: (outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten_passwort.txt")
									,  "receiver"     	: "7zip Archiv"
									,	 "LetterIncluded"	: 0	}

			If FileExist(outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z")
				FileDelete, % outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z"

			Loop {
				If FileExist(outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z." SubStr("000" A_Index, -2))
					FileDelete, % outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z." SubStr("000" A_Index, -2)
				else
					break
			}

			recipPat := lvRow.names.Recipient <= 2 ? "Patient" : lvRow.names.Recipient
			physicanfolder:= lvRow.names.RecipientName " [" recipPat "]"
			physicanname	:= RegExReplace(lvRow.names.RecipientName, "^.+?\.\s*.+\.\s*")
			zip[recipPat] := { "pwd"          	: ""
											,  "File"         	: ""
											,  "zipPaths"     	: ""
											,  "outfilePath"  	: (outbasePath "\" physicanfolder "\" A_YYYY "-" A_MM "-" A_DD "_" physicanname ".7z")
											,  "outpassPath"  	: (outbasePath "\" physicanfolder "\" A_YYYY "-" A_MM "-" A_DD "_" physicanname "_passwort.txt")
											,  "receiver"     	: (lvRow.names.RecipientName " [" recipPat "]")
											,	 "bsnr"         	: recipPat
											,	 "LetterIncluded"	: 0	}


		}


		; neues zip Objekt anlegen
		If (!zipToOne || (zipToOne && initZipFile)) && !IsObject(zip[recip]) {
			zip[recip] := {"pwd"           	: ""
									,  "File"         	: ""
									,  "zipPaths"     	: ""
									,  "outfilePath"  	: (outFilePath ".7z")
									,  "outpassPath"  	: (outFilePath "_passwort.txt")
									,  "receiver"     	: (lvRow.names.RecipientName " [" recip "]")
									,	 "bsnr"         	: recip
									,	 "LetterIncluded"	: 0	}
			dbg ? SciTEOutput(lvRow.names.RecipientName " [" recip "]") : ""
			If FileExist(zip[recip].outFilePath ".7z")
				FileDelete, % zip[recip].outFilePath ".7z"
		}


		; ---------------------------------------------------------------------------------------------------
		; erstellt Dateien je nach Empfänger, legt 7zip Instanzen an um mehrere Dateien bearbeiten zu können
		If !IsObject(zip[recip].File) && ((zipToOne && !initZipFile) || !zipToOne) {

		 ; 7zip Archiv erstellen/öffnen
			zip[recip].File := new 7Zip(zip[recip].outfilePath, outbasePath)

		; 7zip Einstellungen für die Datei festlegen
			zip[recip].File.SetUnicodeMode(true)
			zip[recip].File.option.VolumeSize 	:= VolumeSize 	; Gesamtarchiv wird dann gesplittet
			zip[recip].File.option.compressType := compressType
			zip[recip].File.option.compressLevel:= compressLevel
			zip[recip].File.option.compressOpts	:= compressOpts " " sOpts
			zip[recip].File.option.excludeFile	:= excludeFile
			zip[recip].File.option.workingdir 	:= workingdir
			zip[recip].File.option.stream 	  	:= stream
			zip[recip].File.option.overwrite  	:= overwrite
			zip[recip].File.option.recurse    	:= recurse
			zip[recip].File.option.Yes        	:= true
			zip[recip].File.option.extractPaths	:= 0
			zip[recip].File.option.hide      		:= Hide
			zip[recip].File.option.password   	:= !settings.NoZipEncryption ? zip[recip].pwd : ""

			If !settings.NoZipEncryption {
				tpLen := StrLen(zip[recip].outfilepath)

				tpath := tpLen > 82 ? ".." SubStr(zip[recip].outfilepath, -1*76) : zip[recip].outfilepath
				zipOpts := zip[recip].File.option
				Text  := "----------------------------------- !!! ACHTUNG !!! ----------------------------------`n"
				Text  .= "               diese Datei NICHT ZUSAMMEN mit der 7zip Datei VERSENDEN!               `n"
				Text  .= "--------------------------------------------------------------------------------------`n"
				Text  .= "  Zipdatei "                                                                         "`n"
				Text  .= "  " tpath                                                                            "`n"
				Text  .= "--------------------------------------------------------------------------------------`n"
				Text  .= "  7Zip-Optionen"                                                                     "`n"
				Text  .= "--------------------------------------------------------------------------------------`n"
				Text  .= "  VSize:      `t"    	zipOpts._volumeSize()                                       	 "`n"
				Text  .= "  compresslevel:`t"  	zipOpts._compressLevel()                                    	 "`n"
				Text  .= "  compresstype:`t"   	zipOpts._compressType()                                     	 "`n"
				Text  .= "  options:     `t"   	zipOpts._compressOpts()                                     	 "`n"
				Text  .= "  stream:     `t"    	zipOpts._stream()                                           	 "`n"
				Text  .= "  includeFile:`t"   	zipOpts._includeFile()                                      	 "`n"
				Text  .= "  excludeFile:`t"   	zipOpts._excludeFile()                                      	 "`n"
				Text  .= "  owrite:    `t"     	zipOpts._overwrite()                                        	 "`n"
				Text  .= "  pwd:       `t"     	zipOpts._password()                                         	 "`n"
				Text  .= "  hide:     `t"      	zipOpts._hide()                                             	 "`n"
				Text  .= "  recurse:   `t"    	zipOpts._recurse()                                          	 "`n"
				Text  .= "  workdir:   `t"    	zipOpts._workingdir()                                       	 "`n"
				Text  .= "--------------------------------------------------------------------------------------`n"
				Text  .= "  Passwort:  `t"     	zip[recip].pwd                                              	 "`n"
				Text  .= "  "
				If zipToOne
					FileOpen(zip["Patientendaten"].outpassPath, "w", "UTF-8").Write(Text)
			}

			IF !IsObject(zip[recip].File) {
				dbg ? SciTEOutput(" ### res: " zip[recip].File ", " res) : ""
				If (res = 0 || ErrorLevel = -1) {
					MsgBox, 0x1000, % "Addendum - " StrReplace(A_ScriptName, ".ahk") , % A_ThisFunc "() - Konnte die 7-zip64.dll nicht laden"
					return
				}
			}

			If zipToOne
				recip := recipPat

			initZipFile := true

		}


	; -----------------------------------------------------------------------------
	; Versandbriefe erstellen und mit "einpacken"
		If settings.ZipWithLetter && !zip[recip].LetterIncluded {

			zip[recip].LetterIncluded := true

			If IsObject(letters := CreateShippingLetter(zip[recip].bsnr, A_YYYY A_MM A_DD )) {

				dbg ? SciTEOutput("html: " StrReplace(letters.htmlpath, stdExportPath "\") (letters.pdfpath ? "`npdf:  " StrReplace(letters.pdfpath, stdExportPath "\") : "")) : ""

				; Briefe hinzufügen
				If zipToOne {
					zip["Patientendaten"].Files.Push(StrReplace(letters.htmlpath, stdExportPath "\"))
					fileList.WriteLine(q letters.htmlpath q)
				} else
					res := zip[recip].File.add(letters.htmlpath)

				If letters.pdfPath
					If zipToOne {
						zip["Patientendaten"].Files.Push(StrReplace(letters.pdfpath, stdExportPath "\"))
						fileList.WriteLine(q letters.pdfpath q)
					}
					else
					res := zip[recip].File.add(letters.pdfpath)

			}
			else {
				dbg ? SciTEOutput(A_ThisFunc "() | Das Erstellen der/des Versandbriefe(s) wurde mit einem Fehler beendet:`n  - " settings.lastError) : ""
			}


		}


	; -----------------------------------------------------------------------------
	; pfad zur Zipdatei hinzufügen
		LV_Modify(lvRow, "-Check -Focus -Select")
		If zipToOne {
			zip["Patientendaten"].Files.Push(StrReplace(zipFileFullpath, stdExportPath "\"))
			fileList.WriteLine(q zipFileFullpath "\" q)
		}
		else {
			res := zip[recip].File.add(zipFileFullPath)
			If (EL := ErrorLevel)
				If IsObject(res) {
					SciTEOutput("File: " zipFileFullPath)
					SciTEOutput("resO: " cJSON.Dump(res))
					SciTEOutput("ErrorLevel: " ErrorLevel)
				} else if res {
					SciTEOutput("File: " zipFileFullPath)
					SciTEOutput("res1: "res "`n")
					SciTEOutput("EL: " EL  "," A_LastError)
				}
		}

	}

		If zipToOne {

			fileList.Close()

			func_call := Func("CheckCompressing")
			SetTimer, % func_call, -1

			res := zip["Patientendaten"].File.add("@" q outbasePath "\7zList.txt" q)
			zip["Patientendaten"].Close()
			zip["Patientendaten"] := ""

			files := []
			Loop {
				If FileExist(outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z." SubStr("000" A_Index, -2))
					files.Push(outbasePath "\" A_YYYY "-" A_MM "-" A_DD "_Patientendaten.7z." SubStr("000" A_Index, -2))
				else
					break
			}

			Gui,QE: Submit, NoHide
			networkpath := QEZPath ? QEZPath : GuiControlGet("QE",, "QEZPath")

			If (networkpath ~= "^\s*(\\\\\w+|\w:)" && InStr(FileExist(networkpath), "D"))
				CopyFilesWithProgress(files, networkpath)
			else
				SciTEOutput(A_ThisFunc "() | Kein Zielpfad existent")

			return 1
		}


	;}

	; ---------------------------------------------------------------------------------------------------------
	; gemeinsame Zip Datei erstellen
	; ---------------------------------------------------------------------------------------------------------;{
	If !zipToOne {
		If isObject(zip) {
			zip._Exportierte_Dateien := zipexports
			FileOpen(outbasePath "\" A_YYYY "-" A_MM "-" A_DD " " A_Hour "-" A_Min " - 7zip Protokoll.json", "w", "UTF-8").Write(cJSON.Dump(zip, 1))
		}

		allzipsPath := outbasePath "\" A_YYYY "-" A_MM "-" A_DD " -allzips"
		allzipsPassPath := ""
		If FileExist(allzipsPath ".7z")
			FileDelete, % allzipsPath ".7z"

		allzips := new 7Zip(allzipsPath ".zip", outbasePath)
		allzips.option.VolumeSize     		:= "100M"    	;VolumeSize
		allzips.option.compressType     	:= "zip"
		allzips.option.compressLevel    	:= "5"       	;compressLevel
		allzips.option.compressOpts	    	:= "-mhe"    	;compressOpts " " sOpts
		allzips.option.excludeFile	    	:= ""
		allzips.option.workdir    	    	:= workdir
		allzips.option.stream 	  	    	:= stream
		allzips.option.overwrite  	    	:= overwrite
		allzips.option.extractPaths				:= 0
		allzips.option.hide      					:= Hide
		allzips.option.password        		:= settings.NoZipEncryption ? "" : StringDecoder()

		tpLen := StrLen(allzipsPath)
		tpath := tpLen > 82 ? ".." SubStr(allzipsPath, -1*76) : allzipsPathh

		Text  := "----------------------------------- !!! ACHTUNG !!! ----------------------------------`n"
		Text  .= "               diese Datei NICHT ZUSAMMEN mit der 7zip Datei VERSENDEN!               `n"
		Text  .= "--------------------------------------------------------------------------------------`n"
		Text  .= "  Zipdatei "                                                                         "`n"
		Text  .= "  " tpath ".7z"                                                                      "`n"
		Text  .= "--------------------------------------------------------------------------------------`n"
		Text  .= "  7Zip-Optionen"                                                                     "`n"
		Text  .= "--------------------------------------------------------------------------------------`n"
		Text  .= "  VSize:    `t"     	allzips.option._volumeSize()                                 	 "`n"
		Text  .= "  clevel:    `t"     	allzips.option._compressLevel()                              	 "`n"
		Text  .= "  ctype:    `t"     	allzips.option._compressType()                               	 "`n"
		Text  .= "  options:    `t"    	allzips.option._compressOpts()                               	 "`n"
		Text  .= "  stream:   `t"     	allzips.option._stream()                                     	 "`n"
		Text  .= "  includeFile:`t"   	allzips.option._includeFile()                                	 "`n"
		Text  .= "  excludeFile:`t"   	allzips.option._excludeFile()                                	 "`n"
		Text  .= "  owrite:   `t"     	allzips.option._overwrite()                                  	 "`n"
		Text  .= "  pwd:      `t"     	allzips.option._password()                                   	 "`n"
		Text  .= "  hide:    `t"      	allzips.option._hide()                                       	 "`n"
		Text  .= "  expaths:   `t"    	allzips.option._extractPaths()                               	 "`n"
		Text  .= "--------------------------------------------------------------------------------------`n"
		Text  .= "  Passwort:  `t"     	allzips.option.password                                      	 "`n"
		Text  .= "  "
		FileOpen(allzipsPath "_Passwort.txt", "w", "UTF-8").Write(Text)


		dbg ? SciTEOutput(A_ThisFunc "() | " StrReplace(allzipsPath, stdExportPath "\") " hinzugefügt wurden: ") : ""
		for recip, item in zip {

			if IsObject(item.File) {

				; Zuwachs an Daten ausgeben
				If (filecount := item.File.getFileCount())
					item.filecount := filecount


				item.File.Close()
				;~ SciTEOutput(item.receiver " - Datei geschlossen")
				item.File := ""
				if item.haskey("File")
					item["File"].Delete()
			}

			allzips.add(item.outfilepath)
			FileDelete, % item.outfilePath
			dbg ? SciTEOutput(A_ThisFunc "() | " StrReplace(item.outfilepath, stdExportPath "\") ": " FileExist(item.outfilepath) " (" item.filecount) " Dateien)" : ""
		}

		allzips.Close()

	}

	;}


}
CheckCompressing() {

	while (!(hwnd:=WinExist("Compressing ahk_class #32770")) && A_Index < 20)
		Sleep 100

	If hwnd {
		If IsObject(wp := GetWindowSpot(hwnd))
			SetWindowPos(hwnd, wp.X, wp.Y, wp.W+170, wp.H+440,, 0x42)

		SciTEOutput(wp.H)

		VerifiedCheck("More", "ahk_id " hwnd)

		while (!hLV && A_Index < 10) {
			Sleep 100
			ControlGet, hLV, hwnd,, SysListview321, % "ahk_id " hwnd
		}

		If hLV
			If IsObject(cp := GetWindowSpot(hLV))
				SetWindowPos(hLV, cp.X, cp.Y, cp.W+150, cp.H+20,, 0x42)

		ControlGet, hProgress, hwnd,, msctls_progress321, % "ahk_id" hwnd
		If hProgress
			If IsObject(cp := GetWindowSpot(hProgress))
				SetWindowPos(hProgress, cp.X, cp.Y, cp.W+150, cp.H,, 0x42)

		Loop, 5 {

			If A_Index = 1
				continue
			ControlGet, hStatic, hwnd,, % "Static" A_Index, % "ahk_id " hwnd
			If hStatic
				If IsObject(cp := GetWindowSpot(hStatic))
					SetWindowPos(hStatic, cp.X, cp.Y, cp.W+150, cp.H,, 0x42)

		}

	}

}

QEKarteikarte()                                                                       	{	;-- Contextmenu: per Kontextmenu alle abgehakten Karteikarten in ALBIS anzeigen lassen

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; ermitteln der ausgewählten Zeilen
  	If !IsObject(lvRows := QELvGetCheckedRows("QEMoveRows")) {
		MsgBox, 0x1000, % scriptTitle, % "Es wurden keine Zeilen ausgewählt!", 5
		return
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; läuft Albis überhaupt und ist es ansprechbar?
	dctht := A_DetectHiddenText, dcthw := A_DetectHiddenWindows
	DetectHiddenText       	, Off
	DetectHiddenWindows	, Off
	If !(WinExist("ahk_class OptoAppClass")) {
		MsgBox, 0x1000, % scriptTitle, % "Bitte Albis starten und die Anmeldung durchführen um die Funktion zu nutzen", 8
		mustReturn := true
	}
	DetectHiddenText % dctht
	DetectHiddenWindows % dcthw
	If mustReturn
		return

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Nutzerbestätigung einholen falls mehr als 10 Karteikarten ausgewählt wurden
	If (lvRows.Count() > 10) {
		MsgBox, 0x1004, % scriptTitle, % "Wollen Sie tatsächlich " lvRows.Count() " Karteikarten zur Ansicht in Albis öffnen?"
		IfMsgBox, No
			return
		IfMsgBox, Cancel
			return
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Karteikarten in Albis anzeigen
	For lvRowIndex, lvRow in lvRows {

		If (lvRows.Count()>2)
			ZeigsMir("öffne Karteikarte von ", lvRowIndex, lvRows.Count(),,, "[" lvRow.PatID "] " cPat.Name(lvRow.PatID, true))

		hAkte := AlbisAkteOeffnen(cPat.Name(lvRow.PatID), lvRow.PatID, false)
		;~ SciTEOutput(A_ThisFunc "() | hAkte: " hAkte " PatID: " lvRow.PatID "/" AlbisAktuellePatID())

		while (AlbisAktuellePatID() != lvRow.PatID && A_Index < 40) {
			If WinExist("Versicherungsverhältnis für Privatpatienten ahk_class #32770")
				If !VerifiedClick("Ab&brechen", "Versicherungsverhältnis für Privatpatienten ahk_class #32770",,, 2)
					VerifiedClick("Button11", "Versicherungsverhältnis für Privatpatienten ahk_class #32770",,, 2)
			Sleep 100
		}

		WinWait, % "ALBIS ahk_class #32770", % "nicht vorhanden", 2
		If (hNoPat := WinExist("ALBIS ahk_class #32770", "nicht vorhanden"))
			If !VerifiedClick("OK",,, hNoPat, 2)
				VerifiedClick("Button1",,, hNoPat, 2)

		Sleep 300
	}

	If WinExist("Patient öffnen ahk_class #32770")
		If !VerifiedClick("Abbruch"	, "Patient öffnen ahk_class #32770")
			VerifiedClick("Button3"	, "Patient öffnen ahk_class #32770")

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Fortschrittsanzeige zurückstellen
	If (lvRows.Count()>2) {
		ZeigsMir("selfcall=off")
		ZeigsMir("STOP")
	}

}

QEMoveRows(folder:="")                                                                 	{	;-- Dateien:  	 Daten werden verschoben nicht gelöscht

	; settings.folderToPath - muss global sein
	SciTEOutput(A_ThisFunc "() | - folder: " folder " = " settings.folderToPath[folder])

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Ermitteln der ausgewählten Zeilen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If !IsObject(lvRows := QELvGetCheckedRows("QEMoveRows")) {
		SciTEOutput(A_ThisFunc "(): keine Zeile gewählt" )
		return
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Zustimmung des Nutzers einholen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	confirm := ""
	For each, lvRow in lvRows
		If (PatID := lvRow.PatID)
			confirm .= (confirm ? "`n":"") "[" PatID "] " SubStr(cPat.Name(PatID), 1, 40) (StrLen(cPat.Name(PatID))>40 ? "...":"") " *" cPat.GEBURT(PatID, true)
	If !confirm
		return

	MsgBox, 0x1004, % scriptTitle, % "Sollen die Daten " (lvRows.Count()>1 ? "der":"des") " Patienten`n"
													. "-------------------------------------------------------------------------`n"
													. confirm
													. "`n-------------------------------------------------------------------------`n"
													. "ins Verzeichnis: '" settings.folderToPath[folder] "' verschoben werden?`n"
	IfMsgBox, No
		return
	IfMsgBox, Cancel
		return
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Daten aus der Listview entfernen (letzte ausgewählte Zeile wird zuerst entfernt)
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Loop % lvRows.Count() {

		Index        	:= lvRows.Count()-A_Index+1
		lvRow         	:= lvRows[Index]
		rowNr        	:= lvRow.row
		PatID        	:= lvRow.PatID
		rowData     	:= lvRow.data.indexed
		Export        	:= rowData[11]  ; Export-Spalte

	  ; Verzeichnisse
		origin        	:= Export ? stdExportPath "\" Export : ""
		If origin && InStr(FileExist(origin), "D")
			destination	:= stdExportPath "\verschoben\" settings.folderToPath[folder] "\(" PatID ") " cPat.Name(PatID)
		else {
			LV_ModifyCol(rowNr, "Col10", RegExReplace(rowData[10], "\[.*?\]", "") " [Verzeichnis nicht vorhanden: " StrReplace(origin, stdExportPath "\") "]" )
			origin := destination := ""
			continue
		}

      ; Originalpfad ist nicht vorhanden
		res := FileMovePath(origin, destination)
		If IsObject(res) && (res.1 < 0 || res.2 > 0) {
			LV_ModifyCol(rowNr, "Col10", RegExReplace(rowData[10], "\[.*?\]") " [Datenpfade konnte nicht angelegt werden]" )
			origin := destination := ""
			continue
		}

	  ; es hat funktioniert, dann kann die Zeile entfernt werden
		LV_Delete(rowNr := lvRow.row)
		If !ErrorLevel {
			out1 := out2 := ""
			Loop 7 {
				out1 .= rowData[A_Index] ", "
				out2 .= (out2 ? ", " : "") rowData[A_Index+7]
			}
		}

	  ; Exportpfad auffrischen
		If !SQLUpdateItem("EXPORTS", "Exportpath", PatID, destination)
			noSQLUpdate .= "More1 - " PatID

	  ; MORE3 speichert spezielle Attribute von Verzeichnissen
		more3 := SQLiteGetCell("EXPORTS", "MORE3", "PatID",  PatID)
		more3 := RegExReplace(more3, "i)\s*folder\s*=\s*.*?(\s|$)")
		more3 := Trim(more3)
		If !SQLUpdateItem("EXPORTS", "MORE3", PatID, "folder=" folder (more3 ? " " more3 : ""))
			noSQLUpdate := RTrim(noSQLUpdate, PatID) "More3 - " PatID

		msg := "["PatID "] " cPat.Name(PatID, true) " wurde aus den Exporten aussortiert nach '\verschoben\" settings.folderToPath[folder] "'`n               "  out1 "`n               " out2
		SQLiteErrorWriter(msg, "", "", false, "")

	}

	If noSQLUpdate && dbg
		SciTEOutput(A_ThisFunc "() | noSQLUpdates: " noSQLUpdate)


}

QEBulkDataChange()                                                                     	{	;-- Daten:    	 mehrere Tabellenspalten aufeinmal ändern

	global vTChanges

	If !IsObject(vTChanges)
		vTChanges := Object()

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

  ; Eingabewerte auslesen
	QEFields := QEGetLastInputs()

  ; ausgewählte Zeilen auslesen
	lvRows := QELvGetCheckedRows()

  ; Schlüssel mit leeren Werten entfernen
	fields := {}
	For QEName, value in QEFields {
		If (value && !(QEName ~= "i)(QEPatID|QEPatient|QEShipment)"))
			fields[QEName] := value
	}
	If (fields.Count() = 0)
		return

  ; Daten in der Tabelle ändern
	For each, lvRow in lvRows {                                 ; abgehackte Tabellenzeilen

		rowNr       	:= lvRow.row                               	; Zeilennummer
		PatID        	:= lvRow.names.PatID                       	; Patientennummer

		For QEName, value in fields {					                  	; Nutzereingabefelder

			QEControl 	:= vCtrls[vCtrlRef[QEName]]
			If (!value || !(colNr := QEControl.lvcol) || QEName = "QEShipment")
				continue

			cellText :=  lvRow.indexed[colNr]

		 ; der Empfänger soll geändert werden
			If (QEName = "QERecipient") {                         	; Empfänger hat 2 Spalten (Spalte 6 ist sichtbar und  Spalte 12 versteckt)

				If (lvRow.names.recipient = (bsnr:=value))
					continue

				shippedO    	:= lvRow.indexed.15
				oldpath      	:= stdExportPath "\" lvRow.indexed.11
				newPath     	:= stdExportPath "\" (bsnr > 3 ? Recipients["#" bsnr] : "Patient") "\" (shippedO ? "versendet\": "") "(" PatID ") " cPat.Name(PatID)

			 ; neuer Empfänger führt zu reichlich vielen Änderungen
				LV_Modify(rowNr, "Col6" , fields.RecipientName)					;  ist ohne [BSNR]
				LV_Modify(rowNr, "Col11", StrReplace(newPath, stdExportPath "\"))
				LV_Modify(rowNr, "Col12", bsnr)

			 ; vTChanges erweitern
				If !IsObject(vTChanges[PatID])
					vTChanges[PatID] := Object()

			 ; geänderten Exportpfad für die Speicherung in der Datenbank sichern
				vTChanges[PatID]["ExportPath"] := {"PatID": PatID, "rowNR":rowNr, "colNr":11, "cellText":newPath, "cellTextO":oldpath}

			  ; als letztes muss noch die BSNR gesichert werden
				dbColName := "Recipient"
				vTChanges[PatID]["Recipient"] := {"PatID": PatID, "rowNR":rowNr, "colNr":12, "cellText":bsnr, "cellTextO":lvRow.names.recipient}
				saveTblChanges := true

			}
			else  {   ; if (QEName != "QEExport")

				if (cellText = value)
					continue

				LV_Modify(rowNr, "Col" colNr, value)
				dbColName := QEControl.db.Exports
				If QEControl.db.isDate
					value := ConvertToDBASEDate(value)

			  ; Änderung hinzufügen
				If !IsObject(vTChanges[PatID])
					vTChanges[PatID] := Object()

				vTChanges[PatID][dbColName] := {"PatID": PatID, "rowNR":rowNr, "colNr":colNr, "cellText":value, "cellTextO":cellTextO}

				saveTblChanges := true
			}

			If (!settings.saveTblChanges && saveTblChanges = true) {
				GuiControl, QE: Enable, QEsqlSave
				GuiControl, QE: Enable, QEsqlCancel
				settings.saveTblChanges := true
			}

		}

	}

}
vTChangeExists(PatID, rowNr, colNr, value) {

	global vTChanges

	vTChangeExist := false

	If !IsObject(vTChanges[PatID])
		vTChanges[PatID] := Object()
	else {

		For vTPatID, vTCol in vTChanges
			For colname, vT in vTCol
				If (vTPatID=PatID && vT.rowNr=rowNr && vT.colNr=colNr && vT.cellText=value) {
					vTChangeExist := true
					break
				}

	}

return vTChangeExist
}

QEFEditHandler(sText)                                                                 	{	;-- GuiControl:	 Schnellsuchefeld

	qmatch := {}

	Gui, QE: Submit, NoHide
	Gui, QE: ListView, QELvExp

	If RegExMatch(sText, "O)^\s*(?<PatID>\d+)", qm) {

		qmatch.PatID := qm.PatID
		Loop % LV_GetCount() {
			LV_GetText(rowPatID, A_Index, 1)
			If (rowPatID = qmatch.PatID) {
				;~ SciTEOutput("qmatch.PatID: ("  A_Index ") " qmatch.PatID)
				LV_Modify(A_Index+10, "Vis")
				LV_Modify(A_Index, "Check Focus Select Vis")
				break
			}

		}
		QsText := SubStr(sText, 1, StrLen(qm[0]))
	}
	;~ If RegExMatch(sText " " , "O)(?<Name1>[\pL\-]+)\s*(?<Delim>[\s,])\s*(?<Name2>[\pL\-]+)*", qm) {
		;~ If () {
			;~ qmatch.surname  	:= qm.Delim="," && qm.Name2 ? qm.Name1 :
			;~ qmatch.prename 	:= qm.Name2
		;~ }
	;~ }


return
}

QEInCellEditHandler(EditState, hLV:="", hEdit:="", rowNr:="", colNr:="", cellText:="") 	{	;-- Listview: 	 Daten wie in einer Exceltabelle ändern, Funktion speichert auch in die DB

	global vTChanges
	static vTable := {"hEdit":{}}
	dbg := false

	If !IsObject(vTChanges)
		vTChanges := Object()

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

	emptyChanges := false

	dbg ? SciTEOutput("EditState: " EditState) : ""

	If    	(EditState = "REMOVE"             	&& !settings.saveTblChanges) 	{

		vTable := {"hEdit":{}}
		vTChanges := Object()
		GuiControl, QE: Disable, QEsqlSave
		GuiControl, QE: Disable, QEsqlCancel

		return "vTable, vTChanges items removed"

	}
	else if (EditState ~= "i)(BEGIN|END)")                                   	{

		If (EditState = "BEGIN") {

			LV_GetText(PatID, rowNr, 1)
			vPatID := "#" PatID
			If !IsObject(vTable[vPatID])
				vTable[vPatID] := {}
			vTable[vPatID][colNr] 	:= {"BEGIN": cellText, "hEdit":hEdit}
			vTable.hEdit[hEdit]    	:= {"PatID": PatID, "colNr":colNr, "rowNr":rowNr, "rowID":LV_GetRowMapID(rowNr, QEhLvExp)}

			If (colNr = 6) {
				LV_GetText(bsnr, rowNr, 6)
				vTable.hEdit[hEdit].bsnr := bsnr
			}


		}
		else If (EditState = "END") {

			vPatID 	:= "#" (PatID := vTable.hEdit[hEdit].PatID)
			colNr 	:= vTable.hEdit[hEdit].colNr
			;~ rowNr	 	:= vTable.hEdit[hEdit].rowID ? LV_GetRowIndex(vTable.hEdit[hEdit].rowID, QEhLvExp) : vTable.hEdit[hEdit].rowNr
			rowNr	 	:= vTable.hEdit[hEdit].rowNr

			If (vTable[vPatID][colNr].BEGIN != cellText) {

				settings.saveTblChanges := true
				vTable[vPatID][hEdit]["END"]	:= cellText

			  ; Geburt oder VDatum  - Umformen ins Datenbankformat
				If (colNr = 5 || colNr = 9) {
					If RegExMatch(cellText, "(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>\d{4})", m)			; z.B. 01.04.2023
						cellText := mY SubStr("00" mM, -1)  SubStr("00" mD, -1)
					else if RegExMatch(cellText, "(?<D>\d{2})(?<M>\d{2})(?<Y>\d{4})", m)				; z.B. 01042023
						cellText := mY mM mD
					else {                                                                     	; Ersetzen mit ursprünglichen Daten
						LV_Modify(rowNr, "Col" vTable.hEdit[hEdit].colNr, vTable[vPatID][colNr].BEGIN)
						return
					}

				}

				If !IsObject(vTChanges[PatID])
					vTChanges[PatID] := Object()

				dbColName := ExportsTbl[colNr]   ; Spaltenname erhalten
				vTChanges[PatID][dbColName] := {"PatID": PatID, "rowNR":rowNr, "colNr":colNr, "cellText":cellText, "rowID":vTable.hEdit[hEdit].rowID}
				;~ SciTEOutput(cJSON.Dump(vTChanges, 1))

				; Speichern und Abbruch Steuerelemente freischalten
				GuiControl, QE: Enable, QEsqlSave
				GuiControl, QE: Enable, QEsqlCancel
			}

		}

	}
	else if (EditState = "SAVESQL"            	&& settings.saveTblChanges) 	{

		notsaved := []
		RowsToHide := []

		Gui, QE: Submit, NoHide
		GuiControl, QE: -Redraw, QELvExp

		If QERecipient
			RegExMatch(QERecipient, "\[(?<BSNR>\d+)\]", "Fltr")

		For vTPatID, vTCol in vTChanges {   ; Änderungen erfolgen in der Reihenfolge der Eingabe

			ExportPathO := shipmentO := bsnrO := ""
			ExportPathChanged := false
			UnCheckRow := false

			For dbColname, vT in vTCol {

				rowNr := vT.rowNr
				;~ rowNr := vT.rowID ? LV_GetRowIndex(vT.rowID, QEhLvExp) : vT.rowNr

				If !UnCheckRow {
					LV_Modify(rowNr, "-Check -Select -Focus")
					UnCheckRow := true
				}

				; Empfänger, Versanddatum und Exportpfad werden zusammen geändert
				if !ExportPathChanged && (dbColName ~= "i)(Recipient|Shipment|ExportPath)") {

					ExportPathChanged := true
					isProcessingExport := true

					shipmentO	        		:= SQLiteGetCell("Exports", "Shipment"	, "PatID",  vT.PatID)
					bsnrO        	    		:= SQLiteGetCell("Exports", "Recipient"	, "PatID",  vT.PatID)
					oldPath     	    		:= SQLiteGetCell("Exports", "ExportPath", "PatID",  vT.PatID)
					oldPathShort        	:= StrReplace(oldPath, stdExportPath "\")

					recipient            	:= vTChanges[vT.PatID]["Recipient"].cellText
					shipment            	:= vTChanges[vT.PatID]["Shipment"].cellText
					writeEmptyValues    	:= vTChanges[vT.PatID]["Shipment"].writeEmptyValues
					newPath              	:= vTChanges[vT.PatID]["ExportPath"].cellText
					newPath             	:= stdExportPath "\" StrReplace(newPath, stdExportPath "\")
					newPathShort        	:= StrReplace(newPath, stdExportPath "\")

				  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				  ; Verschieben des kompletten Verzeichnis je nach Empfänger
					res := FileMovePath(oldPath, newPath)

				  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				  ; bei einem Dateifehler alles rückgängig machen
				  ;  res.1 - Fehlernummer oder Anzahl kopierter Dateien
				  ;  res.2 - Anzahl nicht kopierter Dateien
				  ;  res.3 - 0 wenn Originalverzeichnis nicht entfernt werden konnte
					If ((res.1 != -3 && res.1 < 0) || res.2 ) {  ; || !res.3

						dbg ? SciTEOutput(A_ThisFunc "() | res: " res.1 " | " res.2 " | " res.3 " | ") : ""

						If bsnrO {
							LV_Modify(rowNr, "Col6"	, SQLiteGetRecipient(bsnrO))
							LV_Modify(rowNr, "Col12", bsnrO)
						}

						LV_Modify(rowNr, "Col9" 	, ConvertDBASEDate(shipmentO))
						LV_Modify(rowNr, "Col15"	, shipmentO)

						LV_Modify(rowNr, "Col10"	, "[Pfadänderung ist fehlgeschlagen]")
						LV_Modify(rowNr, "Col11"	, oldPathShort)

						continue
					}

					; Filter: fehlerfreie und Versendete ausblenden
					else if (settings.LVExpFilter) {

						mf := MissedFiles(rowNr)
						If mf.errmsg {
							LV_GetText(anm, rowNr, 10)
							anm := "[" mf.errmsg " " RegExReplace(anm, "\s*\[.*\]+\s*", "")
							LV_Modify(rowNr, "Col10", anm)
						}
						else
							RowsToHide.InsertAt(1, rowNr)

					}

					; Filter: Versendete
					else if (settings.LVExpFilter2 && shipment ~= "^\d{8}$")
						RowsToHide.InsertAt(1, rowNr)

				  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				  ; Daten können in die SQLite Datenbank geschrieben werden
					If (StrLen(RegExReplace(recipient, "[^\pL\d]"))>0)								        	; Empfänger wenn übergeben
						If !SQLUpdateItem("Exports", "Recipient", vT.PatID, recipient)
							notsaved.Push(vTChanges[vT.PatID]["Recipient"])
					If (newPath != oldPath)
						If !SQLUpdateItem("Exports", "ExportPath", vT.PatID, newPath)							; Verzeichnis
							notsaved.Push(vTChanges[vT.PatID]["ExportPath"])
					If (shipment~="\d{8}" || writeEmptyValues)
						If !SQLUpdateItem("Exports", "Shipment", vT.PatID, shipment)
							notsaved.Push(vTChanges[vT.PatID]["Shipment"])

				}

				else if !(dbColName ~= "i)(Recipient|Shipment|ExportPath)") {

					If !SQLUpdateItem("Exports", dbColName, vT.PatID, vT.cellText)
						notsaved.Push(vT)

				}

			}

		}

		GuiControl, QE: +Redraw, QELvExp

	  ; herausgefilterte Zeilen jetzt entfernen
		If RowsToHide.Count() {
			Sleep 2000
			GuiControl, QE: -Redraw, QELvExp
			while RowsToHide.Count() {
				rowNr := RowsToHide.Pop()
				LV_Delete(rowNr)
			}
			GuiControl, QE: +Redraw, QELvExp
		}

		If isProcessingExport {
			GuiControl, QE: 			, QEFEdit, % "" ; leeren
			GuiControl, QE: Focus	, QEFEdit				; fokussieren
		}

		emptyChanges := true
		QELetterChoiceTS() ; Dokument-DDL erneuern

	}
	else if (EditState = "CANCELSQL"          	&& settings.saveTblChanges) 	{

		MsgBox, 0x1004, % scriptTitle, % "Achtung! Wollen Sie wirklich sämtliche Änderungen verwerfen?"
		IfMsgBox, No
			return
		IfMsgBox, Cancel
			return

		emptyChanges:= true
		func_QEShow	:= Func("QEShowTable")
		SetTimer, % func_QEShow, -2000

	; Nutzer wollte gerade Empfänger zuordnen. Dann lege den Tastaturfocus ins Schnelleingabefeld zurück
		For vTPatID, vTCol in vTChanges
			For dbColname, vT in vTCol
				If (dbColName ~= "i)(Recipient|Shipment|ExportPath)") {
					GuiControl, QE: Focus, QEFEdit
					break
				}

		QELetterChoiceTS() ; Dokument-DDL leeren

	}

  ; setzt Gui zurück und leert die Objekte
	If emptyChanges {

		vTChanges := {}
		vTable  	:= {"hEdit":{}}
		emptyChanges := false

		settings.saveTblChanges := false

		GuiControl, QE: Disable, QEsqlSave
		GuiControl, QE: Disable, QEsqlCancel

	}

}

CheckQExporterState()                                                                  	{

	global qexporter

	Critical

	If !qexporter.ExportRunning {
		qexporter.stop := false
		Thread, Interrupt, 15
		GuiControl, QE:, QEExport, % "Patientenakten`nexportieren"
		ToolTipOff(6)
		ZeigsMir("STOP")
		func := Func("CheckQExporterState")
		SetTimer, % func,  Off
		return
	}

	func := Func("CheckQExporterState")
	SetTimer, % func,  -200

}

QEGetLastInputs()                                                                      	{	;-- GuiControl:	 Eingabefelder zwischenspeichern

	global

	Gui, QE: Default
	Gui, QE: Submit, NoHide

	Pat := QEPatientTxt()
	If !RegExMatch(QERecipient, "\[(?<snr>\d+)\]", b)
		bsnr := QERecipient ~= "i)\s*Patient" ? 1 : QERecipient ~= "i)\s*Angehöriger" ? 2 : 0
	else if RegExMatch(QERecipient, "\[(?<snr>\d+)\]", b)
		QERecipient := RegExReplace(QERecipient, "\[.*\]")

	QESdateO := QESdate
	If RegExMatch(QESdate, "(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>\d{4})", m)
		QESdate := SubStr("00" mD, -1) "."  SubStr("00" mM, -1) "."  mY
	else if RegExMatch(QESdate, "(?<D>\d{2})(?<M>\d{2})(?<Y>\d{2})", m)
		QESdate := mD "."  mM "." SubStr(A_YYYY, 1, 2) mY
	else
		QESdate := ""

	If (QESdate != QESdateO) {
		GuiControl, QE:, QESdate, % "|" QESdate
		GuiControl, QE: ChooseString, QESdate, % QESdate
	}

	QEShipmentO := QEShipment
	If RegExMatch(QEShipment, "(?<D>\d{1,2})\.(?<M>\d{1,2})\.(?<Y>\d{4})", m)
		QEShipment := SubStr("00" mD, -1) "."  SubStr("00" mM, -1)  mY
	else if RegExMatch(QEShipment, "(?<D>\d{2})(?<M>\d{2})(?<Y>\d{4})", m)
		QEShipment := mD "."  mM "." SubStr(A_YYYY, 1, 2) "." mY
	else
		QEShipment := ""

	If (QEShipment != QEShipmentO) {
		GuiControl, QE:, QEShipment, % "|"
		GuiControl, QE:, QEShipment, % "|" QEShipment
		GuiControl, QE: ChooseString, QEShipment, % QEShipment
	}



return {"QEPatID":Pat.ID, "QEPatient":Pat.txt, "QESDate":QESdate, "QERecipient":bsnr, "QEMedium":QEMedium, "QERoute":QERoute, "QEShipment":QEShipment, "QEMore1":QEMore1, "RecipientName":QERecipient}
}

QEStandards()                                                                        		{	;-- GuiControl:	 Eingabefelder leeren, Buttonstandards wiederherstellen

	;global vCtrls

	Gui, QE: Default

	For each, ctrl in vCtrls {
		GuiControl, % "QE: " ,  % ctrl.vName, % (ctrl.cType = "ComboBox" ? (ctrl.value ? ctrl.value : "|") : ctrl.value)
		If (ctrl.cType = "ComboBox")
			GuiControl, % "QE: " (ctrl.Standard ? "ChooseString":"Choose"), % ctrl.vName, % (ctrl.standard ? ctrl.standard : 0)
		}

	GuiControl, QE:           	, QEShipment, % "|"
	GuiControl, QE:           	, QEShipment, % "|" (today := A_DD "." A_MM "." A_YYYY)
	GuiControl, QE: ChooseString, QEShipment, % today

	GuiControl, QE: ChooseString, QESdate  	, % settings.QESdate

	GuiControl, QE: Enable     	, QEPatID
	GuiControl, QE: Enable     	, QEPatient
	GuiControl, QE: Enable     	, QEMedium
	GuiControl, QE: Enable     	, QERoute
	GuiControl, QE: Enable     	, QESearch
	GuiControl, QE:            	, QESearch 	, % "Patient suchen"

	GuiControl, QE:           	, QENewItem	, % "Datensatz anlegen"
	GuiControl, QE: Disable    	, QENewItem

}

QEPatientTxt()                                                                        	{	;-- Daten:	     liest nur den Inhalt des bearbeiten Patientenfeldes aus

	Gui, QE: Default
	Gui, QE: Submit, NoHide
	If !(PatID 	:= GuiControlGet("QE", "", "QEPatID"))
		PatID := QEPatID ? QEPatID : 0
	If (QEPatientTxt := GuiControlGet("QE", "", "QEPatient")) {
		Gui, QE: Submit, NoHide
		QEPatientTxt := StrLen(RegExReplace(QEPatient, "[^\pL\d]")) > 0 ? QEPatient : QEPatientTxt
		RegExMatch(QEPatientTxt, "O)(?<surname>[\pL\-\s]+)*\s*(\,\s*(?<prename>[^\d\*\.]+))*[\*\s\,]*((?<birth>\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})))*", m)
		RegExMatch(QEPatientTxt, "O)^\s*(?<ID>\d+)[\s,]*$", n)
		return {"ID"     : (n.ID?n.ID : PatID?PatID : 0)
					, "surname": RegExReplace(m.surname, "[,\s]+$")
					, "prename": RegExReplace(m.prename, "[,\s]+$")
					, "birth"  : m.Birth
					, "txt"    : QEPatientTxt}
	}

}

QESaveItems(Pat, Mode:=0)                                                              	{	;-- Daten:     	 legt einen neuen Datensatz oder ändert einen vorhandenen in der SQLite Datenbank

	If (!Pat.ID || !IsObject(Pat))
		return

	PatSet := QEGetPatSet(Pat, Mode)
	dbg ? SciTEOutput(A_ThisFunc "() | " cJSON.Dump(PatSet, 1)) : ""


  ; legt einen neuen Datensatz an. fügt eine Zeile in der Tabelle hinzu
	If (PatSet.RecordExist = 0) {

		PatSet := SQLWriteRecord(PatSet, "EXPORTS", true)
		arr := [PatSet.ID, PatSet.Surname, PatSet.Prename, ConvertDBASEDate(PatSet.Birth)]
		For colNr, ctrl in vCtrls {
			If (colNr > 2 && PatSet.Columns.haskey(colname := ctrl.db.Exports)) {
					item  	:= PatSet.Columns[colname].gui
					lvCol 	:= PatSet.Columns[colname].lvCol
					arr[lvCol] := item
			}
		}
		Gui, QE: Default
		Gui, QE: ListView, QELvExp
		LV_Add("", arr*)

		QEShowFilesCount()

	}

 ; ändert einen vorhandenen Datensatz
	else If (PatSet.RecordExist=1) {  ;### nicht implementiert bisher

		_SQL := "UPDATE EXPORTS SET  SName='"	     	","
													.    " PName='"	     	","
													.  	 " Birth='"   	 	","
													. 	 " SDate='"	    	","
													. 	 " Recipient='"		","
													. 	 " Medium='"	  	","
													. 	 " Route='"   		","
													. 	 " Shipment='"	 	","
													. 	 " More1='"     	","
													. 	 " ExportPath='" 	","
													. 	 " More3='"    	 	""
													. 	 " WHERE PatID=" PatSet.ID ";"


	}

return PatSet
}

QELvGetCheckedRows(callfrom:="")                                                       	{	;-- Listview: 	 Daten aller abgehakten Zeilen erhalten

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

	lvRows 	:= [], cRow 	:= 0
	while (cRow := LV_GetNext(cRow, "C")) {
		LV_GetText(PatID, cRow, 1)
		lvr := QELvGetRow(cRow)
		lvRows.Push({"row"  	: cRow
							,  "PatID"	: PatID
							,  "rowID"	: LV_GetRowMapID(cRow, QEhLvExp)                              	; ID anstatt Index
							,  "data" 	: lvr
							,  "indexed":	lvr.indexed
							,  "names"	: lvr.names})
	}

return lvRows.Count()>0 ? lvRows : ""
}

QELvGetRow(cRow, retOnly:="")                                                          	{	;-- Listview: 	 liest eine Zeile der Listview in füllt die EIngabefelder damit

	; retOnly : entweder "indexed" oder "names", per Standard wird beides zurückgegeben

	;global TExports
	global QEhSDate, QEhPatient

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

 ; komplette LV-Zeile lesen
	lvr := {"indexed" : Array(), "names" : Object()}
	Loop % LV_GetCount("Column") {
		LV_GetText(rText, cRow, (colNr := A_Index))
		rText := colNr=10 ? RegExReplace(rText, "\s*\[.*\]\s*") : rText
		lvr.names[ExportsTbl[colNr]] := lvr.indexed[colNr] := rText
	}


return retOnly ? lvr[retOnly] : lvr
}

QEFillElements(cRow)                                                                   	{	;-- GuiControl:  füllt die Steuerelemente des Eingabebereiches

	;global TExports
	global QEhSDate, QEhShipment

	dataset_ready := true

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

 ; komplette LV-Zeile lesen
	lvr := QELvGetRow(cRow)

  ; editierbare Steuerelemente befüllen
	For colNr, ctrl in vCtrls {

	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
	  ; PatID
		If (ctrl.vName = "QEPatID") {
			If (lvr.indexed.1 ~= ctrl.rx) {
				GuiControl, QE:, QEPatID	, % lvr.indexed.1
				GuiControl, QE: Disable	, QEPatID
			}
			else {
				dataset_ready := false
				GuiControl, QE: Enable	, QEPatID
			}
		}

	  ; Namensfeld
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QEPatient") {
			If (lvr.indexed.2 && lvr.indexed.3 && lvr.indexed.4) {
				GuiControl, QE:, QEPatient, % (Name := lvr.indexed.2 ", " lvr.indexed.3 " *" lvr.indexed.4)
				GuiControl, QE: Disable, QEPatient
				Sleep 100
				SendMessage, 0x000A, 0,,, % "ahk_id " QEhPatient    ; WM_Enabled Message, wParam = 0 = Disabled
			}
			else {
				dataset_ready := false
				GuiControl, QE: Enable	, QEPatient
			}
		}

	  ; SDate
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QESDate") {
			sDate := lvr.indexed.5 ? lvr.indexed.5 : ctrl.Standard
			If !InStr(ctrl.value, "|" sDate)
				GuiControl, QE: , % "QESDate", % (ctrl.value .= "|" sDate)
			GuiControl, QE: ChooseString	, % "QESDate", % sDate
			GuiControl, QE: Enable         	, % "QESDate"
			If !SDates
				dataset_ready := false
		}

	  ; Recipient
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QERecipient") {
			bsnr 	:= lvr.indexed.12
			recip	:= lvr.indexed.6
			recipient := bsnr > 2 ? recip " [" bsnr "]" : bsnr=1 ? "Patient" : "Angehöriger"

			If !InStr(ctrl.value, "|" recipient)
				GuiControl, QE: , % "QERecipient", % (ctrl.value .= "|" recipient)

			GuiControl, QE: ChooseString	, % "QERecipient", % recipient
			GuiControl, QE: Enable         	, % "QERecipient"

			GuiControl, QE: , % "QEMedium"	, % (bsnr > 2 ? "|🖄 ePost|💽 Sammel-CD" 	: "") 	. "|💿 Einzel-CD|📚 Papierausdruck|⚿ Datei verschlüsselt|❔ ungeklärt"
			GuiControl, QE: , % "QERoute"  	, % (bsnr > 2 ? "|📡 Telematik" 	         			: "")	. "|✉ Postbrief|📪 Briefkasten|🤝 persönliche Übergabe|📧 EMail 🔐 Passwort per #|❔ ungeklärt"
			}

	  ; Versandmedium
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QEMedium") {
			Medium	:= lvr.indexed.7
			Medium	:= bsnr <= 2 && Medium ~= "i)ePost" ? "❔ ungeklärt" : Medium
			GuiControl, QE: ChooseString	, % "QEMedium", % Medium
			GuiControl, QE: Enable         	, % "QEMedium"

			If (Medium ~= "i)ePost")
				GuiControl, QE: ChooseString	, % "QERoute", % "📡 Telematik"

			If !Medium
				dataset_ready := false

		}

	  ; Versandweg
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QERoute") {
			Route	:= lvr.indexed.8
			GuiControl, QE: ChooseString	, % "QERoute", % Route
			GuiControl, QE: Enable         	, % "QERoute"

			If !Route
				dataset_ready := false
		}

	  ; Versanddatum
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QEShipment") {
			If (lvr.indexed.9 ~= ctrl.rx) {
				If !InStr(ctrl.value, "|" (Shipment := lvr.indexed.9))
					GuiControl, QE: , % "QEShipment", % (ctrl.value .= "|" Shipment)
			}
			GuiControl, % "QE: " (Shipment ? "ChooseString":"Choose")	, % ctrl.vName, % (Shipment ? Shipment : 0)
			GuiControl, QE: Enable         	, % ctrl.vName
		}

	  ; HInweise
	  ; - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (ctrl.vName = "QEMore1") {
			GuiControl, QE: , % "QEMore1", % lvr.indexed.10
			GuiControl, QE: Enable         	, % "QEMore1"
		}

	}


	;~ cExp := QECheckExport()
	lvr.cRow := cRow
	lvr.dataset_ready := dataset_ready

return lvr
}

QEGetPatSet(Pat, Mode:=0)                                                             	{	;-- Daten:       sucht den passenden Patienten, ändert die Eingabefelder, liest vorhandene weitere Daten aus

	global

	If !IsObject(Pat)
		return

  ; sucht den Patienten mit den angegebenen Daten
	If (!Pat.ID && (Pat.Surname || Pat.Prename || Pat.Birth)) {

		If (Pat.Surname && Pat.Prename)
			patmatches := cPat.StringSimilarityPlus(Pat.Surname " " Pat.Prename)

			If (!IsObject(patmatches) || patmatches.Count() != 1) {

				If Pat.Birth {

					searchobj := Array()
					If Pat.Surname
						searchobj.Push({"key": "NAME"    	, "value": Pat.Surname})
					If Pat.prename
						searchobj.Push({"key": "VORNAME", "value": Pat.Prename})
					If Pat.Birth
						searchobj.Push({"key": "GEBURT"	, "value": ConvertToDBASEDate(Pat.Birth)})
					PatIDs := cPat.PatID(searchobj)

					 If (PatIDs.Count() > 1) {
						For idx, PatID in PatIDs
							ids .= (ids ? "`n" : "")  idx  ": [" PatID "] " cPat.Name(PatID, true)
						MsgBox, 0x4, % "Nochmal", % "Es wurden mehrere passende Patienten gefunden:`n`n" ids
						GuiControl, QE: Enable	, QEPatID
						GuiControl, QE: Enable	, QEPatient
						GuiControl, QE: Focus	, QEPatient
						return
					}
					else  If (PatIDs.Count() =0 ) {
						MsgBox, 0x4, % "Nochmal", % "Es wurden keine passende Patienten gefunden."
						GuiControl, QE: Enable	, QEPatID
						GuiControl, QE: Enable	, QEPatient
						GuiControl, QE: Focus	, QEPatient
						return
					}
					else
						Pat.ID := PatIDs.1

				}

			}

			If !Pat.ID && patmatches.Count() != 1 {

				MsgBox, 0x4, Nochmal, % "Bitte die Eingabe prüfen.`n"
													. "Es " (!IsObject(patmatches) || !patmatches.Count() 	? " konnte kein passender Patient gefunden werden"
																							: patmatches.Count()>1 	? " gab mehr als ein Suchergebnis (" patmatches.Count() ")!" :"")
				GuiControl, QE: Enable	, QEPatID
				GuiControl, QE: Enable	, QEPatient
				GuiControl, QE: Focus	, QEPatient
				return
			}
			else
				Pat.ID := patmatches.1.PatID

	}

 ; Pat.ID ist vorhanden dann den Datensatz zusammenstellen
	If Pat.ID && cPat.Exist(Pat.ID) {

		Gui, QE: Submit, NoHide

	; Patientendatensatz (PatSet-Objekt) anlegen
		PatSet := {	"ID"			  	: Pat.ID
						, 	"Surname"   	: StrSplit(cPat.NAME(Pat.ID), ",", A_Space).1
						, 	"Prename"   	: StrSplit(cPat.NAME(Pat.ID), ",", A_Space).2
						, 	"Birth"	  		: cPat.GEBURT(Pat.ID)
						, 	"eText"      	: cPat.NAME(Pat.ID, true)
						, 	"RecordExist"	: ""
						, 	"Record"	  	: ""
						,		"columns"    	: {}}


	; fehlende Eingaben ersetzen
		if (QEPatID != PatSet.ID)
			GuiControl, QE:, QEPatID, % PatSet.ID
		if (QEPatient != PatSet.EText)
			GuiControl, QE:, QEPatient, % PatSet.EText

		If !Mode {
			PatSet.Columns	:= { "PatID" 	: {"lvcol":1, "gui": Pat.ID                                    	, "guictrl": "QEPatID"}
											, 	 "SName"	: {"lvcol":2, "gui": StrSplit(cPat.NAME(Pat.ID), ",", A_Space).1, "guictrl": "QEPatient"}
											,		 "PName"	: {"lvcol":3, "gui": StrSplit(cPat.NAME(Pat.ID), ",", A_Space).2, "guictrl": "QEPatient"}
											,		 "Birth" 	: {"lvcol":4, "gui": cPat.GEBURT(Pat.ID)                       	, "guictrl": "QEPatient"}}
		} else {
			Pat := QEPatientTxt()
			PatSet.columns	:= { "PatID" 	: {"lvcol":1, "gui": Pat.ID                                   	, "guictrl": "QEPatID"}
											, 	 "SName"	: {"lvcol":2, "gui": Pat.Surname			          								, "guictrl": "QEPatient"}
											,		 "PName"	: {"lvcol":3, "gui": Pat.Prename  							        				, "guictrl": "QEPatient"}
											,		 "Birth" 	: {"lvcol":4, "gui": Pat.Birth       								      			, "guictrl": "QEPatient"}}
		}


	 ; wenn Eingabefelder befüllt sind, werden die Inhaltehinzugefügt
		UserInputs := QEGetLastInputs()
		For colNr, ctrl in vCtrls
			If (colNr > 2 && !PatSet.Columns.haskey(colname:=ctrl.db.Exports)) {

				autoInput := ""
				uin := Trim(UserInputs[ctrl.vName])

				; prüft ob die Empfängerzuordnung oder Medium oder Route nicht angegeben wurden. Fehlende Daten werden ergänzt wie unten zu sehen.
				if (ctrl.vName = "QERecipient") {
					If (StrLen(RegExReplace(uin, "[^\pL\d]")) = 0)
						autoInput := "Patient"
				}
				else if (ctrl.vName = "i)(QEMedium|QERoute)") {
					If (StrLen(RegExReplace(uin, "[^\pL\d]")) = 0)
						autoInput := "❔ ungeklärt"
				}

				PatSet.columns[colname] := {"lvcol"		: !IsObject(ctrl.col) ? ctrl.col : cJSON.Dump(ctrl.col) ;vCtrlRef[ctrl.vName]
														  ,	    "gui"   	: (autoInput ? autoInput : UserInputs[ctrl.vName] )
														  ,	    "guictrl"	: ctrl.vName}

			}


	; Prüft ob ein Datensatz für die PatID bereits angelegt ist
		PatExists := SQLiteRecordExist("EXPORTS", "PatID", Pat.ID, "returnOnlyExist")

	  ; Patient wurde bereits angelegt
		If (PatExists = 1) {
			PatSet.RecordExist := PatExists
			if IsObject(result := SQLiteRecordExist("EXPORTS", "PatID", Pat.ID)) {
				SciTEOutput(A_ThisFunc "() | PatExists: " PatExists "`nobj: " cJSON.Dump(result))
				PatSet.Record := result.Record                        	; die passende Zeile mit allen Spalten
				For colname, column in result.Columns {               	; ## wozu brauche das??
					if !IsObject(PatSet.Columns[colname])
						PatSet.Columns[colname] := Object()
					PatSet.Columns[colname].db     	:= column.db		  		; geordnet nach Spaltenbezeichnung
					PatSet.Columns[colname].dbCol 	:= column.dbCol
				}
			}
		}

		; Patient ist noch nicht in der QEExporter.db
		else if (PatExists = 0) {
			PatSet.RecordExist 	:= PatExists

		}

		; ein SQLite-Fehler ist aufgetreten
		else {
			PatSet.RecordExist 	:= -1
			PatSet.ErrorMsg    	:= StrSplit(PatExists, ",").1
			PatSet.Errorcode   	:= StrSplit(PatExists, ",").2
		}

	return PatSet
	}
	else if (!Pat.ID || !cPat.Exist(Pat.ID)) {
		;~ TrayTip, % scriptTitle, % "Keinen passenden Patienten gefunden",  5
		cp  := GuiControlGet("QE", "Pos", "QEPatient")
		msg := QEPatient ? "Es wurde kein Patient mit dem Namen '" QEPatient "'`nin der Albisdatenbank gefunden"
				:  QEPatID   ? "Die PatID (" QEPatID ") existiert nicht"
				:              "QEPatient und QEPatID Variablen sind leer"
		ToolTip, % msg, % cp.X, % cp.Y-50, 16
		ToolTipOff(16, 5)
		return
	}

	;~ SciTEOutput("Somethings went wrong "  cJSON.Dump(Pat, 1))

}

QECheckExport()                                                                        	{	;-- GuiControl:  prüft den Inhalt der Eingabefelder auf korrekte Eingaben

	; Kombifunktion: prüfen und Guidaten auslesen
	; benötigt um Button "Vorbereiten zum Exportieren" ein- oder auszuschalten
	; sammelt den Inhalt der Gui-Steuerelemente (ComboBox, Edit) in einem Objekt für schnellere Auswertungen

	; vCtrls - sind global

	Gui, QE: Default
	Gui, QE: Submit, NoHide

	isfree := PatID := 0
	cols := {"names" : Object(), "indexed": Array()}

	For index, ctrl in vCtrls {

		vctrl 	:= ctrl.vName
		val	:= %vctrl%
		colName := ctrl.lv
		cols.names[colName] := val
		cols.indexed[index] := {"name":colName, "val":val}

		If (index < 6) {
			If RegExMatch(val, ctrl.rx, match){
				isfree += 1
				PatID := (index=1) && matchID ? matchID : PatID
			}
			else
				notfree .= (notfree ? "|" : "") ctrl
		}

	}

return {"isfree":isfree, "notfree":notfree, "cols": cols, "ready": (isfree = 5 ? true : false), "PatID":PatID}
}

QEShowTable(DataTable:="", ColLimit := 10)                                             	{	;-- Listview: 	 Dokumentationsdaten anzeigen

	global QEhLvExp
    static recis := ["Patient","Angehöriger"]

	settings[A_ThisFunc] := 1
	nofpath   := []

  Gui, QE: Default
  Gui, QE: ListView, QELvExp
  GuiControl, QE: Disable, QELVExpFilter
  GuiControl, QE: Disable, QELVExpFilter2
  GuiControl, QE: Disable, QELVExpFilter3

	filePaths := []
	expFilesCount :=expsend :=  0

   ;lädt Daten aus Exports Tabelle wenn keine Tabelle übergeben wurde
	If !IsObject(DataTable) || (!DataTable.HasNames && !DataTable.HasRows) {
		Zeigsmir("wird geladen" , 0,100,0,0, "Tabelle")
		If !QEDB.GetTable(sql.Exports.GetTable, DataTable)
			SQLiteErrorWriter("Tabelle konnte nicht gelesen werden", QEDB.ErrorMsg, QEDB.Errorcode, true, sql.Exports.GetTable)
	}
	Zeigsmir("ist geladen" , 0,0,0,0, "Tabelle")

   ; Filtereinstellungen laden
	settings.LVExpFilter  := GuiControlGet("QE", "", "QELvExpFilter")
	settings.LVExpFilter2	:= GuiControlGet("QE", "", "QELvExpFilter2") ; nur unversendete
	If (settings.LVExpFilter2 && settings.LVExpFilter) {
		GuiControl, QE: -g	, QELVExpFilter
		GuiControl, QE:		, QELVExpFilter, 0  ; uncheck
		GuiControl, QE: -g	, QELVExpFilter, QEBHandler
		settings.LVExpFilter := false
	}

	dbg ? SciTEOutput(A_ThisFunc "-  settings.LVExpFilter: " settings.LVExpFilter " -  .LVExpFilter2: " settings.LVExpFilter2  " -  .LVExpFilter3: " settings.LVExpFilter) : ""

	If (settings.LVExpFilter3 	:= GuiControlGet("QE", "", "QELvExpFilter3")) {
		RegExMatch(GuiControlGet("QE", "", "QERecipient"), "\[(?<recipient>\d+)\]", gui_)
	}

   LV_Delete()
   GuiControl, QE: -ReDraw, QELvExp
   GuiControl, QE: -g, QELvExp
   Gui, QE: ListView, QELvExp
   ColCount := LV_GetCount("Column")

  ; Zellen befüllen
	If (DataTable.HasNames && DataTable.HasRows)  {

		steps := Floor(DataTable.RowCount/50)

		Loop % DataTable.RowCount {

		  ; Fortschritt anzeigen
				If (Mod(A_Index, steps)=0)
					Zeigsmir("Die Tabelle wird befüllt" , A_Index, DataTable.RowCount, "Zeile: " A_Index "/" DataTable.RowCount,0)

			DataTable.Next(Row)

			; alle im Patienten die "verschoben" werden nie anzeigen
				If (Row[12] && InStr(Row[12], "folder=")) {
					expFilesCount ++
					continue
				}

			; aussortierte Exporte/Namen nicht anzeigen
			;~ If (settings.LVExpFilter3 && Row[6] != gui_recipient)  ; nächster Filter wird überschrieben, wenn der im Empfängerfeld eingestellte Arzt der Empfänger im aktuellen Datensatz ist
				If (settings.LVExpFilter2 && (StrLen(Trim(Row[9]))>0 || InStr(Row[11], "\versendet"))) {
					expFilesCount ++
					expSend ++
					continue
				}

		  ; fehlerfreie Exporte/Namen nicht anzeigen
			If (settings.LVExpFilter || settings.LVExpFilter2) {

 			  ; 0 wenn Daten als versendet gekennzeichnet wurden
				expFiles := MissedFiles(Row)
				If (!expFiles.missed || expFiles.shipped) {
					expFilesCount ++
					GuiControl, QE:, QELVExpCounter, % "[" expFilesCount "/" LV_GetCount() "]"
					If (settings.LVExpFilter2 && expFiles.shipped) || (settings.LVExpFilter )
						continue
				}

			}

		  ; fügt eine Zeile hinzu
			RowCount := LV_Add("", "")

		  ; Spalten füllen
			Loop % ColLimit+1 {                                                                                                                  	; DataTable.ColumnCount

				cell := Trim(Row[col := A_Index])
				If (col ~="^(4|5|9)$") {                                                                                                          	; Datum umwandeln
					If (StrLen(cell) > 0) {
						LV_Modify(RowCount, "Col" (col=4 ? "13" : col=5 ? "14" : "15"), cell )
						cell := ConvertDBASEDate(cell)
					}
				}
				else if (col = 11 )	{		                                                                                                        		; Exportpfad
					if (StrLen(cell) > 0) {                                                                                                         	; ein Verzeichniseintrag ist vorhanden
						rowID := LV_GetRowMapID(RowCount, QEhLvExp)                                                                                    	; ID bestimmen und anstatt des Indexwertes speichern
						filePaths.Push({"rowID"  	: rowID
													, "rowNr" 	: RowCount
													, "PatID" 	: Row[1]
													, "fpath"   : (RegExReplace(Row[11], "^\w:", stdExportDrive))
													, "shipment": Row[9]})                                                                                           	; die Dateien im Verzeichnis werden später auf Vollständigkeit geprüft
						cell := SubStr(cell, StrLen(stdExportPath)+2, StrLen(cell)-StrLen(stdExportPath)-1)                                            	; für die Darstellung wird der lange Pfad entfernt
					}
					else {                                                                                                                          	; PatID für Suche nach vermisstem Dateipfad sichern
						nofpath.Push(Row[1])
					}
				}
				else if (col = 6)                                	{	                                                                              	; Empfänger als Klarnamen
					LV_Modify(RowCount, "Col12", (bsnr := cell) )
					If (cell=1 || cell=2) {                                                                                                     			; Vorgaben	"PAtient" und "Angehöriger"
						cell := recis[bsnr]
					}
					else if (cell>2) {	                                                                                                    	       	; BSNR Nummern
						If !QEDB.GetTable(_SQL := "SELECT ANR, RTitel, RSName, RPName FROM RECIPIENTS WHERE BSNR=" bsnr ";", Result) {
							SQLiteErrorWriter("Recipient - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, false, _SQL)
						}
						else if IsObject(Result) {
							m := result.Rows.1
							m.1 := RegExReplace(m.1, "i)\s*(Frau)\s*", "Fr. ")
							m.1 := RegExReplace(m.1, "i)\s*(Herr)\s*", "Hr. ")
							cell := (m.1 ? m.1 "" : "") . (m.2 ? RegExReplace(m.2, "i)\s*med.\s*") " " : "") . m.3 . ", " . m.4
						}
					}
				}
				else if (col = 10 && expFiles.ErrMsg)	{                                                                                            	; Anmerkungen
					cell := expFiles.ErrMsg ("#" expFiles.ErrMsg ? " " : "") cell
				}

			  ; Zeile auffüllen
				LV_Modify(RowCount, "Col" col, cell)

			}

			expFiles := ""

		}

	}

	If settings.LVExpFilter2
		LV_ModifyCol(6, "Sort")

	GuiControl, QE: +gQEBHandler, QELvExp
	GuiControl, QE: +ReDraw   	, QELvExp

	dbg ? SciTEOutput(A_ThisFunc "() | " expsend " versendete Datensätze ausgeblendet") : ""

	ZeigsMir("selfcall=off")
	ZeigsMir("STOP")

 ; startet eine Funktion die
	dbg ? SciTEOutput(A_ThisFunc ": " filepaths.Count()) : ""
	If (!settings.LVExpFilter && filepaths.Count() > 0) {
		warnfunc := Func("QEMissedFilesDetector").Bind(filepaths)
		SetTimer, % warnfunc, -100
		;~ SciTEOutput(cJSON.Dump(filepaths, 1))
	}

	m := cell := bsnr := col := result := ""

	QEShowFilesCount()

	GuiControl, QE: Enable, QELVExpFilter
	GuiControl, QE: Enable, QELVExpFilter2
	GuiControl, QE: Enable, QELVExpFilter3

	settings[A_ThisFunc] := 0

	If (nofpath.Count()>0) {
		func_call := Func("FindMissedPaths").Bind(nofpath)
		SetTimer, % func_call, -3000
	}

	RowsCount := LV_GetCount()
	SB_SetText(RowsCount " Patienten", 4)

return RowsCount
}

QEShowFilesCount()                                                                     	{	;-- zählt Exportdaten

	static exportsCountSQL  	:= "SELECT PatID FROM EXPORTS WHERE ExportPath <> '' AND LENGTH(More3) = 0;"
	static preparedSQL     		:= "SELECT PatID FROM EXPORTS WHERE Route != '❔ ungeklärt' AND LENGTH(Shipment) = 0 AND LENGTH(More3) = 0;"
	static shipmentCountSQL 	:= "SELECT PatID FROM EXPORTS WHERE Shipment <> '';"

	Gui, QE: Default


	If !QEDB.GetTable(_SQL:=exportsCountSQL  	, result)
		SQLiteErrorWriter(table " Tabelle - Fehler beim Zählen von Spalteninhalten", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
	If IsObject(result) {
		GuiControl, QE:, QELVExpCounter1, % (exports := result.RowCount)
		result := ""
	}

	If !QEDB.GetTable(_SQL:=preparedSQL     	, result)
		SQLiteErrorWriter(table " Tabelle - Fehler beim Zählen von Spalteninhalten", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
	If IsObject(result) {
		GuiControl, QE:, QELVExpCounter2, % (prepared := result.RowCount)
		result := ""
	}

	If !QEDB.GetTable(_SQL:=shipmentCountSQL	, result)
		SQLiteErrorWriter(table " Tabelle - Fehler beim Zählen von Spalteninhalten", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
	If IsObject(result) {
		GuiControl, QE:, QELVExpCounter3, % (shipments := result.RowCount)
	}

	If (exports ~= "\d+" and shipments ~= "\d+")
		GuiControl, QE:, QELVExpCounter4, % (outstandings := exports-shipments)

return {"exports":exports, "prepared":prepared, "shipments":shipments, "outstandings":outstandings}
}

MissedFiles(Row)                                                                      	{	;-- Listview: 	 zeigt in der Tabelle an ob eine oder mehrere der generierten Dokumente nicht vorhanden sind

	dbg ? SciTEOutput("Missingswarner gestartet") : ""

	missed := 0
	nofpath   := []
	PatID    	:= Row[1]
	PatName 	:= Row[2] ", " Row[3]
	shipped 	:= Trim(Row[9]) ~= "^(\d{1,2}\.\d{1,2}\.\d{4}|\d{8})$" ? true : false
	note     	:= RegExReplace(Row[10], "\s*\[.+\]\s*", "")
	if Row[11] {
		fpath  	:= Row[11] ~= "^\w:" ? RegExReplace(Row[11], "^\w:", stdExportDrive) : stdExportPath "\" Row[11]
		fpath  	:= shipped ? RegExReplace(fpath, "(\\versendet)*(\\\(\d+\))", "\versendet$1") : fpath
	}
	else if (!Row[11] && cPat.Exist(PatID)) {
		nofpath.Push(PatID)
	}

	If !cPat.Exist(PatID)  {
		missed := 0xB
		x := "[!! PatID gehört zu keinem Patienten !! "
		note := ""
		Shipment := false
	}
	else If fpath && !InStr(FileExist(fpath), "D") {
		missed := 0xB
		x := "[Verzeichnis ist nicht vorhanden"
	}
	else {

		filepath := fpath "\Karteikarte"
		filepath2 := filepath "nExport " StrReplace(Patname, ",", "")

		FileGetSize, fKKhtml, % filepath ".html"
		If !fKKhtml
			FileGetSize, fKKhtml, % filepath " " Patname ".html"

		FileGetSize, fKKpdf, % filepath ".pdf"
		If !fKKpdf
			FileGetSize, fKKpdf, % filepath " " Patname ".pdf"
		If !fKKpdf
			FileGetSize, fKKpdf, % filepath "nExport.pdf"

		FileGetSize, fLabpdf, % fpath "\Labordaten.pdf"


	  ; Karteikarte.html, Karteikarte.pdf oder Labordaten.pdf Dateien fehlend?
		missed := 0
		If (!FileExist(filepath ".html") && !FileExist(filepath " " Patname ".html"))                              	|| (fKKhtml 	= 0) {
			x .= "[fehlend: Karteikarte.html"
			missed |= 0x1
		}
		If (!FileExist(filepath ".pdf") && !FileExist(filepath "nExport.pdf")
			&& !FileExist(filepath " " Patname ".pdf")	&& !FileExist(filepath2 ".pdf"))	        					|| (fKKpdf 	= 0) {
			x .= (!x ? "[fehlend: " : ", " ) "Karteikarte*.pdf"
			missed |= 0x2
		}
		If (!FileExist(fpath "\Labordaten.pdf") && !FileExist(fpath "\Labordaten.txt")                			|| fLabpdf		= 0) {
			x .= (!x ? "[fehlend: " : ", " ) "Labordaten.pdf"
			missed |= 0x3
		}

	}

 ; Fehler bei der SQL Speicherung oder Pfaderzeugung enden manchmal mit dem vollständigen löschen des Exportverzeichnisses
	If (nofpath.Count()>0) {
		func_call := Func("FindMissedPaths").Bind(nofpath)
		SetTimer, % func_call, -3000
	}

return {"missed":missed, "shipped":shipped , "ErrMsg":(x ? x "] ":"") (note ? " " note : "")}
}

QEMissedFilesDetector(filepaths)                                                      	{ ;-- Listview: 	 zeigt in der Tabelle an ob eine oder mehrere der generierten Dokumente nicht vorhanden sind

	global QEhLvExp

	Gui, QE: Default
	Gui, QE: ListView, QELvExp

	expFilesCount := 0
	sendFilesCount := 0

	dbg ? SciTEOutput("Missingswarner gestartet") : ""

	For each, element in filepaths {

		x := ""

		;~ If element.rowID
			;~ rowNr := LV_GetRowIndex(element.rowID, QEhLvExp)                                  	  ; MapID to Index
		;~ else
			rowNr	:= element.rowNr

		LV_GetText(PatID	     	, rowNr, 1)
		LV_GetText(shipment   	, rowNr, 9)
		LV_GetText(note        	, rowNr, 10)

		RegExMatch(element.fpath, "\)\s+(?<PatName>.+)\s*$", row)
		shipped 	:= shipment ~= "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})" ? true : false
		note     	:= RegExReplace(note, "\[.+\]\s*", "")
		fpath   	:= (element.fpath ~= "^\w:") ? RegExReplace(element.fpath, "^\w:", stdExportDrive) : stdExportPath "\" element.fpath
		fpath    	:= shipped ? RegExReplace(fpath, "(\\versendet)*(\\\(\d+\))", "\versendet$2") : fpath

	  ; Patienten Nummer ist unbekannt
		If !cPat.Exist(PatID)  {
			x := "[!! (" PatID ") PatID gehört zu keinem Patienten !!"
			note := ""
		}
	  ; Verzeichnispfad ist nicht vorhanden
		else If !InStr(FileExist(fpath), "D") {
			x := "[Verzeichnis fehlt: " SubStr(fpath,-1*(StrLen(fpath)-StrLen(stdExportPath)-2))
		}
	  ; Verzeichnispfad ist leer oder es fehlen wichtige Dateien
		else {

			; Karteikarte.pdf, Karteikarte.html, KarteikartenExport Unsinn, Reiner.pdf
			filepath := fpath "\Karteikarte", 	filepath2 := filepath "nExport " StrReplace(rowPatname, ",", "")

		 ; Dateigrößen
			FileGetSize, fKKhtml, % filepath ".html"
			If !fKKhtml
				FileGetSize, fKKhtml, % filepath " " rowPatname ".html"

			FileGetSize, fKKpdf, % filepath ".pdf"
			If !fKKpdf
				FileGetSize, fKKpdf, % filepath " " rowPatname ".pdf"
			If !fKKpdf
				FileGetSize, fKKpdf, % filepath "nExport.pdf"

			FileGetSize, fLabpdf, % fpath "\Labordaten.pdf"

		  ; Karteikarte.html, Karteikarte.pdf oder Labordaten.pdf Dateien fehlend?
			If (!FileExist(filepath ".html") && !FileExist(filepath " " rowPatname ".html"))          	|| (fKKhtml 	= 0)
				x .= "[fehlend: Karteikarte.html"
			else If (!FileExist(filepath ".pdf") && !FileExist(filepath "nExport.pdf")
				&& !FileExist(filepath " " rowPatname ".pdf")	&& !FileExist(filepath2 ".pdf"))		|| (fKKpdf 	= 0)
				x .= (!x ? "[fehlend: " : ", " ) "Karteikarte*.pdf"
			If (!FileExist(fpath "\Labordaten.pdf") && !FileExist(fpath "\Labordaten.txt"))       	|| (fLabpdf	= 0)
				x .= (!x ? "[fehlend: " : ", " ) "Labordaten.pdf"

		}

	; PatID und Personenangaben - Übereinstimmung prüfen
		Gui, QE: ListView, QELvExp

		LV_Modify(rowNr, "Col10", (x ? x "] ": "") (note ? note : ""))

		dbg ? SciTEOutput("[" rowNr "] " element.fpath " - " x) : ""

	}

return expFilesCount
}

QEDisEnableAll()                                                                       	{	;-- GuiControl:  Interaktion mit allen Steuerelementen ein- oder ausschalten

	global hQE
	static DisabledList, lastState := "enable"

; alle hwnds und classNN feststellen
	WinGet, ctrlClassList, ControlList    	, % "ahk_id " hQE
	WinGet, ctrlHwndList , ControlListHwnd	, % "ahk_id " hQE
	ctrlClassList := StrSplit(ctrlClassList, "`n")

; Objekt mit bereits "disabled controls"
	if (lastState == "enable")
		DisabledList := Object()

; nur bestimmte Steuerelementklassen werden abgeschaltet
	For ctrlIndex, ctrlHwnd in StrSplit(ctrlHwndList, "`n") {

		if !(ctrlClassList[ctrlIndex] ~= "i)(Edit|Button|Combobox)")
			continue

		if (lastState == "enable") {

			ControlGet, CtrlStyle 	, Style	 ,,, % "ahk_id " ctrlHwnd
			ControlGetText, CtrlText,, % "ahk_id " ctrlHwnd

			if ((CtrlStyle & 0x8000000) || (CtrlText == "Reload Script"))   ; 0x8000000 ist WS_DISABLED
				DisabledList["#" ctrlHwnd] := ctrlClassList[ctrlIndex]
			else
				Control, Style, +0x8000000,, % "ahk_id " ctrlHwnd

		}
		else {

			if !DisabledList.haskey("#" ctrlHwnd)
				Control, Style, -0x8000000,, % "ahk_id " ctrlHwnd

		}


	}

	lastState := lastState == "enable" ? "disable" : "enable"


return 1
}

_ShowRecordSet(RecordSet) {

   Global
   static recis := ["Patient","Angehöriger"]
   Local ColCount, RowCount, Row, RC


   Gui, QE: Default
   Gui, QE: ListView, QELvExp
   GuiControl, QE: -ReDraw, QELvExp

   LV_Delete()
   ColCount := LV_GetCount("Column")

   If (RecordSet.HasRows) {

      If (RecordSet.Next(Row) < 1) {
         MsgBox, 16, %A_ThisFunc%, % "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode
         Return
      }

      Loop {

		 RowCount := LV_Add("", "")

		 Loop % RecordSet.ColumnCount {

			cell := Trim(Row[A_Index])

		 ; Datum umwandeln
			If (A_Index ~="^(4|5|9)$"){
				LV_Modify(RowCount, "Col" (A_Index=4 ? "12" : A_Index=5 ? "13" : "14"), cell )
				cell := ConvertDBASEDate(cell)
			}
		; Empfänger als Klarnamen
			else if (A_Index = 6) {

				bsnr := cell
				LV_Modify(RowCount, "Col12", bsnr )

				If (cell=1 || cell=2)
					cell := recis[bsnr]

				else if (cell>2) {
					IF !QEDB.GetTable(_SQL := "SELECT ANR, RTitel, RSName, RPName FROM RECIPIENTS WHERE BSNR=" bsnr ";", Result)
						SQLiteErrorWriter("Recipient - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, false, _SQL)
					else If IsObject(Result) {
						;~ SciTEOutput("Recipient: " cJSON.Dump(result, 1))
						m := result.Rows.1
						cell := (m.1 ? m.1 " " : "") (m.2 ? RegExReplace(m.2, "i)\s*med.") " " :"") m.3 ", " m.4 "[" bsnr "]"
					}
				}
			}
			LV_Modify(RowCount, "Col" . A_Index, cell)
		}

            RC := RecordSet.Next(Row)
      } Until (RC < 1)

   }
   If (RC = 0)
      MsgBox, 16, %A_ThisFunc%, % "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode
   ;~ Loop, % RecordSet.ColumnCount
		;~ If (A_Index < 11)
			;~ LV_ModifyCol(A_Index, "AutoHdr");~ Loop, % RecordSet.ColumnCount
		;~ If (A_Index < 11)
			;~ LV_ModifyCol(A_Index, "AutoHdr")

 GuiControl, QE: +ReDraw, QELvExp
}


;}

; SQLite                                                           	;{
SQLiteRecordExist(table, colName, value, opt:="returnAllData")                         	{	;-- nur für eindeutige Ergebnisse

	; gibt nur das erste gefundene Ergebnis zurück

	if !QEDB.GetTable(_SQL :=  "SELECT * FROM " table " WHERE " colName " = " value ";", result) {
		SQLiteErrorWriter(table " Tabelle - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
		return QEDB.ErrorMsg "," QEDB.Errorcode
	}
	if IsObject(result) {
		ret := {"Exist": (result.HasNames && result.HasRows && result.RowCount>0 ? true : false)}
		if (opt = "returnOnlyExist")
			return ret.Exist
		if ret.Exist {
			ret.record  := result.Rows.1  	   ; es darf nur einen Datensatz je PatID geben
			ret.Columns	:= Object()
			For colNr, colname in result.ColumnNames
				ret.Columns[colname] 	:= {"db":result.Rows[1, colNr], "dbCol": colNr}
		}
	}


return ret
}

SQLiteGetRecipient(bsnr, adresstext:=false)                                            	{

	IF !QEDB.GetTable(_SQL := "SELECT ANR, RTitel, RSName, RPName, Strasse, PLZ, Ort FROM RECIPIENTS WHERE BSNR=" bsnr ";", Result)
		SQLiteErrorWriter("Recipient - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, false, _SQL)
	else If IsObject(Result) {

		m := result.Rows.1

		If (adresstext=1)
			return m[1] " " (m.2 ? RegExReplace(m.2, "i)Dr.m", "Dr. m") " " : "") .  m[4] " " m[3] "`n" m[5] "`n" m[6] " " m[7] 		; Frau Dr. med. Sieglinde Morgenrot \n An der aufgehenden Sonne 1 \n 99999 Sonnenstund
		; -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --

			ANR     	:= m[1]
			ANRs  		:= RegExReplace(m.1	, "i)\s*(Frau)\s*"	, "Fr. ")
			ANRs  		:= RegExReplace(ANRs, "i)\s*(Herr)\s*"	, "Hr. ")
			Titel   	:= RegExReplace(m[2], "i)Dr\.\s*m"    	, "Dr. m")
			Titels		:= RegExReplace(m[2], "i)\s*med\.*\s*"	, " ")
			SName    	:= m[3]
			PName    	:= m[4]
			strasseNr	:= m[5]
			PLZ	    	:= m[6]
			Ort     	:= m[7]
		; -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --


		if (adresstext=2)
			return {"Recipientname" : RegExReplace((ANRs " " (Titels ? Titels " " : "") SName ", " PName), "\s{2,}", " ")
						,	"Geschlecht"		: ANR
						,	"ANRName"      	: Titel " " PName " " SName
						,	"StrasseNr"   	: StrasseNr
						,	"PLZOrt"       	: PLZ " " Ort}
		; -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --


		if (adresstext=3)
			return {"Recipientname" : RegExReplace((ANRs " " (Titels ? Titels " " : "") SName ", " PName), "\s{2,}", " ")
						,	"Names"        	: SName ", " PName}
		; -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --


		return RegExReplace(ANRs Titels " " m.3 ", " m.4, "\s{2,}", " ")
		;~ return m.1 . (m.2 ? RegExReplace(m.2, "i)\s*med.\s*") " " : "") . m.3 . ", " . m.4

	}

}

SQLiteGetCell(table, SelectColumn, WHERE, rowID)                                      	{

	_SQL := "SELECT " SelectColumn " FROM " table " WHERE " WHERE " = " rowID ";"
	If !QEDB.GetTable(_SQL, Result) {
		SQLiteErrorWriter(table " - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, false, _SQL)
		return
	}

	;~ SciTEOutput(_SQL "`n" cJSON.dump(result, 1))

	item := result.Rows.1.1
	If (SelectColumn = "ExportPath")
		item := RegExReplace(item, "^\w:", stdExportDrive)

return item
}

SQLWriteRecord(PatSet, Table:="EXPORTS", InsertInto:=true)	                          	{	;-- einen Datensatz anfügen oder überschreiben

	;global vCtrls
	global dbg, noSQLWrite

	PatSet.RecordWritten := false

  ; fügt eine neue Zeile hinzu
	If InsertInto {

		_SQLStart  	:=
		_SQLCol 		:= ""
		_SQLValues 	:= ""

		dbg ? SciTEOutput(" - - - - - - - - - " A_ThisFunc " - - - - - - - - -" ) : ""

		For colname, cell in PatSet.Columns {

			If (StrLen(Trim(colname))=0 || colname ~= "i)Object.+")
				continue

			dbg ? SciTEOutput(colname ":" cell.gui) : ""

			_SQLCol   	.= (_SQLCol ? ", ": "") "'" colname "'"

				      item 	:= cell.gui ~= "^\s*\d+\.\d+\.\d{2,4}"    	? ConvertToDBASEDate(cell.gui)
								: colname = "SName"                            	? PatSet.surname
								: colname = "PName"                            	? PatSet.prename
								: colname = "Birth"                           	? PatSet.Birth : cell.gui
			_SQLValues 	.= (_SQLValues ? ", ": "") "'" item "'"

		}

		_SQL := "INSERT INTO " . Table . " (" . _SQLCol . ") VALUES (" _SQLValues ");"
		dbg ? SciTEOutput(A_ThisFunc "() | " _SQL ) : ""

	}

	If noSQLWrite
		return

	QEDB.Exec("BEGIN TRANSACTION;")

	If !QEDB.Exec(_SQL)
		SQLiteErrorWriter(table " Tabelle - Fehler bei " (InsertInto ? "INSERT INTO":"UPDATE"), QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
	else {
		PatSet.RecordExist := true
		PatSet.RecordWritten := true
	}

	QEDB.Exec("COMMIT TRANSACTION;")


return PatSet
}

SQLUpdateItem(Table, ColumnName, rowID, newValue)                                     	{	;-- überschreibt den Inhalt einer einzigen Zelle

	/* Beschreibung

		SQL-Syntax: 	   	UPDATE <table_name> SET <column_name> = <new_value> WHERE <condition>;

		SQL-Beispiel:    	UPDATE Exports SET ExportPath = 'c:\temp\Karteikartenexporte' WHERE PatID = 12154;

		Funktionsaufruf: 	SQLUpdateItem("Exports", "ExportPath", 12154, "c:\temp\Karteikartenexporte")  ; !!
		!! 			  	-> 		Funktion wählt anhand des übergegeben Tabellennamens automatisch den Spaltennamen für die Zeilenidentifikation (Bedingung)
							    		Exports   	> WHERE PatID = rowID;
								    	Recipients	> WHERE BSNR  = rowID;

	*/

	_SQL := "UPDATE " Table " SET " ColumnName "  = '" newValue "' WHERE " (Table="Exports" ? "PatID":"BSNR") " = " rowID ";"

	If !settings.noSQLWrite {
		dbg ? SciTEOutput(A_ThisFunc "() | " _SQL) : ""
		If !QEDB.Exec(_SQL) {
			SQLiteErrorWriter(table " Tabelle - Fehler beim UPDATE einer Zelle in Zeile, Spalte '" rowID "', '" ColumnName "'", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
			return false
		}
	}

return true
}

SQLDeleteItem(Table, rowID)                                                            	{	;--

}

SQLiteExecCallBack(DB, ColumnCount, ColumnValues, ColumnNames)                         	{
   This := Object(DB)
   MsgBox, 0, %A_ThisFunc%
      , % "SQLite version: " . This.Version . "`n"
      . "SQL statement: " . StrGet(A_EventInfo) . "`n"
      . "Number of columns: " . ColumnCount . "`n"
      . "Name of first column: " . StrGet(NumGet(ColumnNames + 0, "UInt"), "UTF-8") . "`n"
      . "Value of first column: " . StrGet(NumGet(ColumnValues + 0, "UInt"), "UTF-8")
   Return 0
}

SQLiteErrorWriter(statement, errmsg, errcode, showMsg:=true, info:="")                	{

	FileAppend, % sqlTimeStamp() "|" statement 	. (errmsg	? "|" errmsg	: "")
								                 	    				. (errcode	? "|" errcode	: "")
									                    				. (info 		? "|" info    	: "")
																		. "`n"
		                                                    			, % settings.paths.SQLProtPath "\QE_SQL-Log.txt"
	If showMsg
		MsgBox, 16, % "SQLite Error - " scriptTitle, % statement
																		. (errmsg 	? "`nMsg:`t"  	errmsg	: "")
																		. (errcode 	? "`nCode:`t" 	errcode	: "")
																		. (info 		? "`nInfo:`t"   	info   	: "")

}

sqlTimeStamp()                                                                        	{
return A_YYYY A_MM A_DD ", " A_Hour ":" A_Min ":" A_Sec ":" A_MSec
}
;}

; MS Edge                                                         	;{
msEdge_SaveAsPDF(url, pdfFilePath, dbg:=false)                                        	{ 	;-- automatisiert die Umwandlung der HTML Datei ins PDF-Format oder druckte eine geöffnete Seite aus


	; anstatt eines PDF-Dateipfades den Namen eines Druckers übergeben um eine HTML Seite ausdrucken zu können
	; pdfFilePath - wo die PDF Datei gespeichert werden soll
	; eine cmdline Zeile reicht zur Umwandlung nach PDF:
	; msedge.exe --headless --print-to-pdf="c:/tmp/test.pdf" "file://c:/tmp/test.html"


	global qexporter

	static loc := {	"cancel"       	: {"en":"Cancel"                     	, "de":"Abbrechen"}
			    			,	"settings"     	: {"en":"Settings and more (Alt+F)" 	, "de":"Einstellungen und mehr (ALT+F)"}
			    			,	"saveasPDF"    	: {"en":"Save as PDF"               	, "de":"Als PDF speichern"}
			    			,	"save"        	: {"en":"Save"                      	, "de":"Speichern"}
			    			,	"portrait"     	: {"en":"Portrait"                  	, "de":"Hochformat"}
			    			,	"printer"      	: {"en":"Printer"                    	, "de":"Drucker"}
			    			,	"copies"      	: {"en":"Copies"                    	, "de":"Kopien"}
			    			,	"print"        	: {"en":"Print"                      	, "de":"Drucken"} }

	url := Trim(StrReplace(url, "\", "/"))
	If (url ~= "i)(^|\W)\w:\/")
		If !FileExist(url) {
			settings.lastError := A_ThisFunc "() | url / Dateipfad ist nicht korrekt: " url
			return -1
		}

	SplitPath, pdfFilePath,, pdfPath
	If !(PdfFilePath~="^\s*\w:\\")
		printToPaper := true, printer := pdfFilePath
	else if !InStr(FileExist(pdfPath), "D") {
		settings.lastError := A_ThisFunc "() | übergebene Datei ist nicht vorhanden " PdfFilePath
		return -2
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 ; Edge starten falls noch nicht erfolgt
	res := MSEdge_Start()
	hEdge := GetHex(res.hwnd)
	newEdge := res.newEdge

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
   ; UIA Zugriff starten
	SetTitleMatchMode, 2
	wTitle 	:= RegExReplace(Trim(WinGetTitle(hEdge)), "i)\[InPrivate\].*$", "[InPrivate]")
	cUIA 	:= new UIA_Browser(wTitle " ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe")
	If !IsObject(cUIA) {
		settings.lastError := A_ThisFunc "() | Browserautomation konnte nicht initialisiert werden. " cUIA
		return -3
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Spracheinstellung des Edge Browser ermitteln
	lang := cUIA.JSGetBrowserLanguage()
	dbg ? SciTEOutput("[0] Browsersprache: " lang ) : ""


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; geöffnete Speichern unter Dialoge schliessen
	If (hSaveas := WinExist("Speichern unter ahk_class #32770 ahk_exe msedge.exe")) {
		WinClose, % "ahk_id " hSaveas
		WinWaitClose, % "ahk_id " hSaveas,, 3
		hSaveas := 0
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; evtl. noch geöffneten Druckdialog schliessen
	cDoc := cUIA.FindFirstBy("ControlType=Document AND Value=edge://print/", 0x7, 2, false)
	If InStr(wTitle, cDoc.GetCurrentPropertyValue("Name")) {
		dbg ? SciTEOutput("[1] evtl. geöffneten Druckdialog schliessen" ) : ""
		cUIA.FindFirstBy("ControlType=Button AND Name=" loc.cancel[lang], 0x7, 2, false).Click()
		cUIA.WaitElementNotExist("ControlType=Button AND AutomationId='selecttrigger-1'", 0x7, 2, false)
		cUIA.GetCurrentDocumentElement()
		cUIA.GetCurrentMainPaneElement()
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; urlbar lesen
	If !(urlbar := UrlDecode(cUIA.GetCurrentUrl())) {
		dbg ? SciTEOutput("Konnte Urlbar nicht auslesen") : ""
		settings.lastError := A_ThisFunc "() | Konnte Urlbar nicht auslesen "
		return -4
	}
	If (!newEdge && !urlbar ~= "i)Karteikarte\.html" && !StrLen(urlbar)=0 ) {
		cUIA.newTab()
		cUIA.WaitPageLoad("New Tab", 1000) ; Wait the New Tab page to load with a timeout of 5 seconds
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; url setzen und dorthin navigieren
	dbg ? SciTEOutput("[2] Setze url und navigiere: " (encUrl := Trim(UrlEncode(url))) ) : ""
	cUIA.Navigate(url, "Karteikartenexport", 2000, 500) ; Set the URL to google and navigate
	If !InStr(urlbar := UrlDecode(cUIA.GetCurrentUrl()), url) {
		dbg ? SciTEOutput("  Navigation ist fehlgeschlagen: " urlbar) : ""
		settings.lastError := A_ThisFunc "() |  url navigation problem`n     urlbar: " urlbar "`n     url:      " url
		return -5
	}
	Sleep 2000


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Menu öffnen und Drucken im Menu auswählen
	If !IsObject(cUIA.WaitElementExist("ControlType=Button AND Name='" loc.settings[lang] "'", 0x7, 2, false, 2).Click()) {
		dbg ? SciTEOutput("[3] Druckdialog per Tastenkobination aufrufen") : ""
	   ; alternativ durch Senden der Tastenkombination Strg+p
		WinActivate, % "[InPrivate] ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe"
		WinWaitActive, % "[InPrivate] ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe",, 3
		ControlSend,, {LCtrl up}{LAlt up}{LShift up}{RCtrl up}{RAlt up}{RShift up}{LControl Down}p{LControl Up} , % "ahk_id " hEdge
		Sleep 4000
	}
	else {
		dbg ? SciTEOutput("[3] Druckdialog über das Menu anfordern") : ""
		If !IsObject(cUIA.WaitElementExist("Name=" loc.print[lang] " AND ControlType=MenuItem AND AutomationID=view_1001", 0x7, 2, false).Click()) {
			settings.lastError := A_ThisFunc "() | Der Druckdialog konnte nicht geöffnet werden"
			return -6
		}
		Sleep 2000
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Druckdialog-Fenster abfangen
	dbg ? SciTEOutput("[4] Druckdialog abfangen") : ""
	If !IsObject(cBtn := cUIA.WaitElementExist("ControlType=Button AND AutomationId='selecttrigger-1'", 0x7, 2, false, 4)) {
		settings.lastError := A_ThisFunc "() | Der Druckdialog konnte nicht abgefangen werden."
		return -7
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Drucker auswählen
	prnName := cBtn.value "  " cBtn.GetCurrentPropertyValue("Name")
	If printToPaper {

		dbg ? SciTEOutput("[5] Druckerauswahl auf '"  printer "' setzen [" prnName "] ") : ""

	 ; Druckerauswahl Button ausrollen und Drucker wählen
		cBtn.Click()
		cUIA.WaitElementExist("ControlType=ListItem AND Name='" printer "'", 0x7, 2, false).Click()
		dbg ? SciTEOutput("[5] a) verschiedene Druckeinstellungen vornehmen (Kopien=1, Layout=Hochformat)") : ""
	; 1 Kopie einstellen
		cUIA.WaitElementExist("ControlType=Spinner AND Name='" loc.copies[lang] "'",0x7, 2, false).Value := "1"
	; Hochformat einstellen
		cUIA.WaitElementExist("ControlType=RadioButton AND Name='" loc.portrait[lang] "' AND AutomationId='optionPortrait'", 0x7, 2, false).Click()
	; Druck starten
		print_Button := "ControlType=Button AND Name='" loc.print[lang] "'"
		cUIA.FindFirstBy(print_Button, 0x7, 2, false)
		cUIA.WaitElementNotExist(print_Button, 0x7, 2, false)
		Sleep 500
		If IsObject(cUIA.FindFirstBy(print_Button, 0x7, 2, false)) {
			MsgBox, 0x1000, % scriptTitle, % "Der Druckdialog ist nicht geschlossen worden.`nMöglicherweise ist der Druck nicht ausgeführt worden.`nBitte kontrollieren."

		}

	return 1
	}


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Microsoft Print to PDF - Speichern der Datei im korrekten Verzeichnis
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	dbg ? SciTEOutput("[5] Druckerauswahl auf '"  loc.saveasPDF[lang] "' setzen [" prnName "] ") : ""
	If !RegExMatch(prnName, "i)" loc.printer[lang] "\s+" loc.saveasPDF[lang]) {
		dbg ? SciTEOutput("    eingestellt ist: " cBtn.GetCurrentPropertyValue("Name") " ersetzen mit '" loc.saveasPDF[lang] "'") : ""
		; Druckerauswahl Button ausrollen
		cBtn.Click()
		cUIA.WaitElementExist("ControlType=ListItem AND Name='" loc.saveasPDF[lang] "'", 0x7, 2, false).Click()
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; nach Speichern Button suchen und Click ausführen
  ; 12 Type: 50000 (Button) Name: "Speichern" LocalizedControlType: "Schaltfläche" ClassName: "c01123 c01153 c01124"
	dbg ? SciTEOutput("[6] Speichern klicken") : ""
		cUIA.FindFirstBy("ControlType=Button AND Name=" loc.save[lang], 0x7, 1, false).Click()  ;1 ist wichtig damit er von links vergleicht
	WinWait, % "Speichern unter ahk_class #32770 ahk_exe msedge.exe", % "ShellView", 10
	dbg ? SciTEOutput("[7] Auf Dialog 'Speichern unter' warten") : ""

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 ; Speichern unter Dialog mit Pfad befüllen und Speichern
	MSEdgeSaveAs:
	If !(hSaveas := WinExist("Speichern unter ahk_class #32770 ahk_exe msedge.exe", "ShellView")) {
		MsgBox, 0x1004, % scriptTitle, % "Bitte den Druckdialog öffnen und als Druckertreiber 'Speichern als PDF-Datei' einstellen.`n"
													.	 "Anschließend 'Speichern' drücken und den Dialog 'Speichern unter' abwarten.`n"
													.	 "'Ja' für weitermachen drücken oder jederzeit 'Nein' für Abbruch.'"
		IfMsgBox, No
			CancelPDF := true
		IfMsgBox, Cancel
			CancelPDF := true

		If CancelPDF {
			settings.lastError := A_ThisFunc "() | Nutzer hat 'Speichern als' abgebrochen."
			return -8
		}
		goto MSEdgeSaveAs

	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; SaveAs Dialog finden
	dialog := UIA.ElementFromHandle(hSaveas)
	;~ dbg ? SciTEOutput("[8] Eingabefeld finden und Speicherpfad setzen") : ""
	cEdit := dialog.FindFirstByNameAndType("Dateiname:", "Edit", 0x7, 1, false)
	;~ dbg ? SciTEOutput("    " cEdit.GetCurrentPropertyValue("Name") ", " cEdit.GetCurrentPropertyValue("ControlType")) : ""

  ; Dateipfad eintragen und prüfen
	Loop 10 {

		If (A_Index>1)
			Sleep 200
		cEdit.SetValue(pdfFilePath)
		Sleep 50
		savepath := cEdit.Value
		IF (SavePath = pdfFilePath) {
			SavePathIsSet := true
			break
		}

	}

  ; jetzt speichern
	;~ dbg ? SciTEOutput("[9] Speicherpfad wurde " (SavePathIsSet ? "gesetzt" : "nicht gesetzt" )) : ""
	If SavePathIsSet {
		dialog.FindFirstByNameAndType("Speichern", "Button").Click()
		WinWaitClose, % "Speichern unter ahk_class #32770 ahk_exe msedge.exe",, 20
		while !FileExist(pdfFilePath) && (A_Index < 100)
			Sleep 50
	}

	cUIA := ""

return FileExist(pdfFilePath) ? true : false
}

msEdge_Start()                                                                        	{

	browserExe := "msedge.exe"

	atmm := A_TitleMatchMode
	SetTitleMatchMode, RegEx
	titleRegEx := "i)\[InPrivate\].*?Microsoft​\s+Edge ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe"

	; Karteikartenexport - [InPrivate] – Microsoft​ Edge ahk_exe msedge.exe ahk_class Chrome_WidgetWin_1
	If !(hEdge := WinExist(titleRegEx)) {

	  ; Run in Incognito mode to avoid any extensions interfering. Force accessibility in case its disabled by default.
		try
			Run, % browserExe " -inprivate --force-renderer-accessibility"
		catch {
			MsgBox, 0x1000, % scriptTitle, % "Der Edge Browser wird für die Ausgabe als PDF-Dokument benötigt.`nLieß sich aber nicht starten."
			return
		}

		newEdge := true
		WinWait, % titleRegEx,, 5
		hEdge := WinExist(titleRegEx)
	}

	SetTitleMatchMode, % atmm

	WinActivate, % "ahk_id " hEdge
	WinWaitActive, % "ahk_id " hEdge,, 10

return {"newEdge":newEdge, "hwnd":hEdge}
}

EncodeHTML(String)                                                                     	{  ;-- nur UTF-8

	;~ string := RegExReplace(string, "&", "&amp;")
	;~ string := StrReplace(string, "<([^\/])", "&lt;") ; &amp;, &quot;, &#39;
	;~ string := StrReplace(string, """", "&quot;")
	;~ string := StrReplace(string, "'", "&#39;")

 return string
}

LC_UriEncode(Uri, RE="[0-9A-Za-z\<\>\pL]")                                            	{
	; Modified by GeekDude from http://goo.gl/0a0iJq
	VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0), StrPut(Uri, &Var, "UTF-8")
	While Code := NumGet(Var, A_Index - 1, "UChar")
		Res .= (Chr:=Chr(Code)) ~= RE ? Chr : Format("%{:02X}", Code)
	Return, Res
}

LC_UriDecode(Uri)                                                                      	{
	Pos := 1
	While Pos := RegExMatch(Uri, "i)(%[\da-f]{2})+", Code, Pos)	{
		VarSetCapacity(Var, StrLen(Code) // 3, 0), Code := SubStr(Code,2)
		Loop, Parse, Code, `%
			NumPut("0x" A_LoopField, Var, A_Index-1, "UChar")
		Decoded := StrGet(&Var, "UTF-8")
		Uri := SubStr(Uri, 1, Pos-1) . Decoded . SubStr(Uri, Pos+StrLen(Code)+1)
		Pos += StrLen(Decoded)+1
	}
	Return, Uri
}

UrlEncode(Url)                                                                         	{ ; keep ":/;?@,&=+$#."
	return LC_UriEncode(Url, "[0-9a-zA-Z:/;?@,&=+$#.]")
}

UrlDecode(url)                                                                         	{
	return LC_UriDecode(url)
}

ChromiumActivateUIA(wTitle)                                                           	{
	WinGet, cList, ControlList, %wTitle%
	if InStr(cList, "Chrome_RenderWidgetHostHWND1")
		SendMessage, WM_GETOBJECT := 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1, %wTitle%
	ControlClick, Chrome_RenderWidgetHostHWND1, % wTitle,, Left, 1, NA  ; removes an open downloads menu
}

;}

; ALBIS Automatisierung                                           	;{
Laborblatt_Export(name, PatPath, Printer="", Spalten="Alles") 	                      	{ 	;-- automatisiert den Laborblattexport

		global qexporter
		static AlbisViewType

		qexporter.LabExportRunning := true

		AlbisViewType	:= AlbisGetActiveWindowType(true)
		savePath       	:= PatPath "\Labordaten.pdf"
		Printer            	:= Printer  	? Printer 	: "Microsoft Print to PDF"
		Spalten         	:= Spalten	? Spalten 	: "Alles"

		If FileExist(savePath)
			FileDelete, % savePath

	; Aufruf der Automatisierungsfunktion
		AlbisActivate(3)
		res := AlbisLabblattExport(Spalten, savePath, Printer, 1)
		Sleep 1000

	; ursprüngliche Ansicht wiederherstellen
		If (AlbisViewType != AlbisGetActiveWindowType(false))
			AlbisKarteikartenAnsicht(AlbisViewType)

		Sleep 2000
		qexporter.LabExportRunning :=  false

	; Fehlercodeanzeige
		If (res != 1) {
			PraxTT("Der Export der Laborwerte ist fehlgeschlagen.`nUrsache: " res, "4 1")
			return res = "keine Labordaten" ? -1 : 0
		}


return 1
}

AlbisLabblattExport(PrintRange, SaveAs="", Printer="", PrintAnnotation=1)              	{ 	;-- PDF Export oder Druckausgabe des Laborblattes

	; spezielle Version für Quickexporter

	static AlbisWT := "ALBIS #32770 ahk_exe " Addendum.AlibsExe
	static noLbParamWT := "Keine Laborparameter zum Ausdruck gewählt"

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Parameter parsen und prüfen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		If settings.CheckAlbis
			CheckAlbisProblems()

	  ; Printer
		Printer := Printer ? Printer : "Microsoft Print to Pdf"

	  ; SAVEAS
		pdfFilePath :=  RegExReplace(SaveAs, "\.\w+$", ".pdf")
		SplitPath, savesAs, SaveName, SaveDir
		If !InStr(FileExist(SaveDir "\"), "D") {
			PraxTT(A_ThisFunc ": Der übergebene Dateiordner existiert nicht`n[" SaveDir "]", "3 1")
			return "missing path: " SaveDir
		}
		If FileExist(PdfFilePath)
			FileDelete, % PdfFilePath

	  ; PrintRange
		PrintRange := Trim(PrintRange)
		If !RegExMatch(PrintRange, "\d{2}\.\d{2}\.\d{4}\-\d{2}\.\d{2}\.\d{4}") && !RegExMatch(PrintRange, "^\d+$") && (PrintRange != "Alles") {
			PraxTT(A_ThisFunc ": PrintRange - Es muss ein Datumsbereich, eine Zahl oder das Wort 'Alles' übergeben werden!`n[" PrintRange "]", "3 1")
			sleep 3000
			return "Problem mit dem Datumsbereich"
		}

	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; prüft auf eine geöffnete Karteikarte
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		AlbisIsBlocked()       ; blockierende Fenster entfernen
		If !InStr(AlbisGetActiveWindowType(), "Laborblatt")
			If !InStr(AlbisGetActiveWindowType(), "Patientenakte")		{
				PraxTT("Der Laborblattdruck kann nur mit geöffneter`nKarteikarte durchgeführt werden!.", "3 1")
				sleep 3000
				return "no case"
			}

	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Laborblatt anzeigen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		count := 0
		while  !InStr(AlbisGetActiveWindowType(), "Laborblatt") {
			If !AlbisLaborblattZeigen(false) {
				count ++
				PraxTT("Das Anzeigen des Laborblattes ist fehlgeschlagen (" count "/5).", "3 1")
				If (count>4)
					return "problem switching to labview"
				Sleep 1000
			}
		}
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Dialog: Laborblatt Druck
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		AlbisActivate(2)
		If !(hLaborblattdruck := WinGetLaborblattDruck(false)) {
			Albismenu(57607, "Laborblatt Druck ahk_class #32770", 4, 1)
			WinWait, % "Laborblatt Druck ahk_class #32770",, 10
			If !(hLaborblattdruck := WinGetLaborblattDruck(false)) {
				PraxTT(A_ThisFunc ": Der Dialog 'Laborblatt Druck' konnte nicht aufgerufen werden!", "3 1")
				sleep 3000
				return " Der Dialog 'Laborblatt Druck' konnte nicht aufgerufen werden!"
			}
			dbg ? SciTEOutput("hLaborblattdruck: " hLaborblattdruck) : ""
		}

	; prüft durch Auslesen des Zeitraumes ob überhaupt Labordaten vorhanden sind (Albis bringt einen Fehler bei Druck von leeren Daten)
		ControlGetText	, LabblattVon			, Edit2, % "ahk_id " hLaborblattdruck
		ControlGetText	, LabblattBis				, Edit3, % "ahk_id " hLaborblattdruck
		ControlGet    	, LabblattParameter	, List,, ListBox1, % "ahk_id " hLaborblattdruck ; "Laborblatt Druck ahk_class #32770"

	; Parameter CHeckbox auswählen
		VerifiedClick("Button13", hLaborblattdruck)  ; Parameter
		If !(SendMessage(0x00F0,,, ListBox1, "ahk_id " hLaborblattdruck))  		; BM_GetCHECK ErrorLevel=1 wenn ausgewählt
			VerifiedClick("Parameter", hLaborblattdruck)  ; Parameter
		Sleep 200

	; Anzahl der Listboxzeilen erhalten
		LBCount := SendMessage(0x018B,,, ListBox1, "ahk_id " hLaborblattdruck)

		If (!LabblattVon || !LabblattBis || !LabblattParameter) {
			If !VerifiedClick("Button10", hLaborblattdruck)
				VerifiedClick("Abbrechen", hLaborblattdruck)
			Sleep 500
			return "keine Labordaten"
		}

	 ; sämtliche Listboxeinträge auswählen
		;~ LabParamAuswahl:
		;~ If (res := IsObject(LB_SetSelAll(hLaborblattdruck))) {
			;~ LabParamWahl:
			;~ MsgBox, 0x1004, % scriptTitle, % "Bitte alle Parameter in der Listbox auswählen.`n"
														;~ .	  res.notselected " Parameter " ( res.notselected=1 ? "ist" : "sind" ) " nicht ausgewählt.`n`n"
														;~ .	 "'Ja' für weitermachen drücken oder jederzeit 'Nein' für Abbruch.'"
				;~ IfMsgBox, No
					;~ CancelPDF := true
				;~ IfMsgBox, Cancel
					;~ CancelPDF := true
				;~ If CancelPDF	{
					;~ return "keine Auswahl an Labordaten getroffen"
				;~ }
				;~ LBCount := SendMessage(0x018B,,, LbName, hwin)
				;~ LBSelCount := SendMessage(0x0190,,, LbName, hwin)
				;~ If (LBCount-LBSelCount > 0)
					;~ goto LabParamWahl
			;~ }


	 ; Einstellungen eintragen
	 ; - - - - - - - - - - - - - - - - - - - -
	 ; Datumsbereich
		If RegExMatch(PrintRange, "(?<Von>\d\d\.\d\d\.\d+)\-(?<Bis>\d\d\.\d\d\.\d+)", rx) || (PrintRange = "Alles") {
			If !VerifiedClick("Button3", hLaborblattdruck)                    ; Zeitraum
				return "LBEx5"
			If (PrintRange != "Alles") {
				If !VerifiedSetText("Edit2", rxVon	, hLaborblattdruck)
					return "problem setting printrange from"
				If !VerifiedSetText("Edit3", rxBis	, hLaborblattdruck)
					return "problem setting printrange to"
				Sleep 250
			}
		}
	  ; Anzahl von Spalten
		else {
			If !VerifiedClick("Button2", hLaborblattdruck)                 	 ; letzte
				return "LBEx8"
			If !VerifiedSetText("Edit1", (PrintRange = 0 ? 1 : PrintRange), hLaborblattdruck)
				return "LBEx9"
			Sleep 250
		}
	 ; Anmerkungen und Probedaten
		If VerifiedCheck("Button5",,, hLaborblattdruck, PrintAnnotation)
			sleep 200
	; Druckschrift Normal
		VerifiedClick("Button7", hLaborblattdruck)
		sleep 200

	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; PDF Druck ausführen - UIA macht es schnell und sicher
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		albisUIA 	:= UIA.ElementFromHandle(AlbisWinID())
		lbprint   	:= UIA.ElementFromHandle(hLaborblattdruck)
		lbprint.FindFirstByNameAndType("Drucker...", "Button",,2,false).Click()

		printd	:= albisUIA.WaitElementExist("ControlType=Window AND Name=Drucken AND ClassName=#32770",,2,false)
		printer 	:= printd.FindFirstByNameAndType("Name:", "ComboBox",,2,false)
		If (printer.value != "Microsoft Print to Pdf") {
			printer.Click()
			printd.WaitElementExist("Name=Microsoft Print to PDF AND ControlType=ListItem",,2,false).Click()
		}

		If (printer.value != "Microsoft Print to Pdf") {
			Sleep 300
			printer.Click()
			printd.WaitElementExist("Name=Microsoft Print to PDF AND ControlType=ListItem",,2,false).Click()
			dbg ? SciTEOutput("Druckertreiber: " printer.value) : ""
			If (!printer.value || printer.value != "Microsoft Print to Pdf")
				return "wrong printer driver"
		}

		printd.FindFirstByNameAndType("OK", "Button",,2,false).Click()
		albisUIA.WaitElementNotExist("ControlType=Window AND Name=Drucken AND ClassName=#32770",,2,false, 5)
		WinWaitClose, % "Drucken #32770 ahk_exe " Addendum.AlibsExe,, 2

		lbprint.FindFirstByNameAndType("OK", "Button",,2,false).Click()
		lbprint.WaitElementNotExist("ControlType=Window AND Name=Laborblatt Druck AND ClassName=#32770",,2,false, 5)
		WinWaitClose, % "Laborblatt Druck #32770 ahk_exe " Addendum.AlibsExe,, 2

	  ; Ausnahme "keine Laborparameter ... gewählt" abfangen
		WinWait, % AlbisWT,  % noLbParamWT, 3
		If WinExist(AlbisWT, noLbParamWT) {
			If !VerifiedClick("Button1", AlbisWT, noLbParamWT,, true)
				VerifiedClick("OK", AlbisWT, noLbParamWT,, true)
			WinWaitClose, % AlbisWT, % noLbParamWT, 3
			If WinExist(AlbisWT, noLbParamWT) {
				WinActivate   	,  % AlbisWT, % noLbParamWT
				WinWaitActive	,  % AlbisWT, % noLbParamWT, 3
				ControlSend,, {Enter},  % AlbisWT,  % noLbParamWT
				WinWaitClose, % AlbisWT, % noLbParamWT, 3
			}
		}


		WinWait, % "Druckausgabe speichern unter ahk_class #32770 ahk_exe " Addendum.AlbisExe,, 10
	 ; Speichern unter Dialog mit Pfad befüllen und Speichern
		LabblattSaveAs:
		If !(hSaveas := WinExist("Druckausgabe speichern unter ahk_class #32770 ahk_exe " Addendum.AlbisExe)) {
			MsgBox, 0x1004, % scriptTitle, % "Bitte den Druckdialog öffnen und als Druckertreiber 'Speichern als PDF-Datei' einstellen.`n"
														.	 "Anschließend 'Speichern' drücken und den Dialog 'Speichern unter' abwarten.`n"
														.	 "'Ja' für weitermachen drücken oder jederzeit 'Nein' für Abbruch.'"
			IfMsgBox, No
				CancelPDF := true
			IfMsgBox, Cancel
				CancelPDF := true
			If CancelPDF	{
				return ""
			}
			goto LabblattSaveAs
		}

	  ; SaveAs Dialog finden
		dialog := UIA.ElementFromHandle(hSaveas)

	  ; Dateipfad eintragen und prüfen
		cEdit := dialog.FindFirstByNameAndType("Dateiname:", "Edit",,2,false)
		Sleep 200
		cEdit.SetValue(PdfFilePath)
		Sleep 100
		If (cEdit.Value != PdfFilePath) {
			cEdit.SetValue(PdfFilePath)
			Sleep 100
		}
		If (cEdit.Value = PdfFilePath)
			SavePathIsSet := true

	  ; jetzt speichern
		If SavePathIsSet {

			dialog.FindFirstByNameAndType("Speichern", "Button",,2,false).Click()
			WinWaitClose, % "Speichern\s+unter ahk_class #32770 ahk_exe " Addendum.AlbisExe,, 20
			WinWait, % "ALBIS ahk_class #32770 ahk_exe " Addendum.AlbisExe, % "Anschluss", 5

		  ; Die Erfassung des Ausgabedialogs von Albis PDF ist aufgrund der maximalen Systemauslastung während des Druckvorgangs nicht möglich.
		  ; Daher können keine Fenster abgefangen werden. Es läßt sich jedoch festzustellen, ob das Albisfenster durch einen geöffneten Dialog blockiert ist.
			;~ Process, Priority,, H
			slpTime := 50, startTime:= A_TickCount, everySecond := Floor(1000//slpTime), ScreenshotsDone := LastTickCount := 0
			hAlbis := AlbisWInID()
			If (hwin := WinExist("ALBIS ahk_class #32770 ahk_exe " Addendum.AlbisExe, "Anschluss")) {
				win := GetWindowSpot(hWin)
			}
			Loop {

				ticks := A_TickCount-startTime
				duration := Round((A_TickCount-startTime)/1000,1)

				If (ticks - lastticks >= 1000 ) {
					lastticks := ticks
					if hwin {
						win := GetWindowSpot(hWin)
						ToolTip, % "warte auf Beendigung des PDF-Laborbefundes`n"
									. 	"seit " duration "s (warte noch max. " 240-duration "s). "
									, % Floor(win.X+Win.W//2), % Floor(win.Y+Win.H//2), 6
					}
				}

				If WinIsBlocked(hAlbis) || (hwin := WinExist("ALBIS ahk_class #32770 ahk_exe " Addendum.AlbisExe, "Anschluss")) {
					LastTickCount := A_TickCount
					LastDuration := duration
					if (ScreenshotsDone<4) && (duration >= 60*(ScreenshotsDone+1)) {     ; 4min!!
						monNr := GetMonitorIndexFromWindow(AlbisWInID())
						MonitorScreenShot(MonNr, StrReplace(A_ScriptName, ".ahk"), settings.paths.SQLProtPath "\" A_YYYY A_MM A_DD A_Hour A_Min A_Sec "-Screenshot.bmp" )
						ScreenshotsDone ++
					}
				}
			  ; 500ms kein PDF-Ausgabedialog oder keine Blockierung des Albisfenster feststellbar dann raus hier
				else If (A_TickCount-LastTickCount >= 500) {
					;~ Process, Priority,, N
					break
				}

				Sleep % slpTime

			}
			ToolTip,,,, 6

		} else {
			;~ Process, Priority,, N
			return 0
		}
	;}


		;~ fn_AlbisCheck := Func("CheckAlbisProblems")
		;~ SetTimer, % fn_AlbisCheck, Off

return 1
}

CheckAlbisProblems()                                                                   	{

	global fn_AlbisCheck
	global qexporter

	static aCrashWT 	:= "Albis\(TM\) ahk_class #32770 ahk_exe AOWDump2.exe"
	static AOWDump 	:= "AOWDump2 ahk_class #32770 ahk_exe AOWDump2.exe"
	static aLoginWT 	:= "ALBIS\s+\-\s+Login ahk_class #32770 ahk_exe "	Addendum.AlbisExe
	static printWT     	:= "ALBIS ahk_class #32770 ahk_exe "				 Addendum.AlbisExe
	static ccount := 0
	static crash := 0
	static LbDruck := 0, LBName:=""

	ccount++

	SetTitleMatchMode, RegEx

	IF (hLBDruck := WinExist(printWT, "Anschluss")) {
		LbDruck += 1
		wTxt := WinGetText(hLBDruck)
		If RegExMatch(wTxt, "i)Dokument.*?(?<nn>\d+\s*\/.*?\/)", n)
			If (LBName != nnn) {
				LbName := nnn
				LbDruck ++
			}
	}

	If IsFunc(fn_AlbisCheck)
		SetTimer, % fn_AlbisCheck, Off

	If WinExist("ahk_class OptoAppClass")
		AlbisExist := true

	If (hCrash := WinExist(aCrashWT )) {
		crash += 1
		dbg ? SciTEOutput(_ThisFunc "() | Crashwindow gefunden" ) : ""
		AlbisAbsturz(hCrash, A_TimeIdlePhysical, "Zeit nach letzer physischer Eingabe: " (Round(idlep/1000/60) ? Round(idlep/1000/60) " Min. " : Round(idlep/1000) " Sek. "))
	}

	If (hCrReport := WinExist(AOWDump))  {
		dbg ? SciTEOutput(_ThisFunc "() | AOW Dump gefunden" ) : ""
		crreport +=1
		AlbisAbsturz(0)
	}

	if WinExist("ALBIS ahk_class #32770", "Online\-Update") || WinExist("Erinnerung\s+online\s+Update ahk_class #32770")
		AlbisAbsturz(0)

	If (hnoParam := WinExist(printWT, "Keine\s+Laborparameter\s+zum\s+Ausdruck")) {
		WinActivate   	, % printWT, % "ahk_id " hnoParam
		WinWaitActive	, % printWT, % "ahk_id " hnoParam, 2
		If !VerifiedClick("OK",,, hnoParam, 2)
			VerifiedClick("Button1",,, hnoParam, 2)
		WinWaitClose, % printWT, % "ahk_id " hnoParam, 2
		If (hLabBlattDruck := WinGetLaborblattDruck()) {
			WinActivate   	, % "ahk_id " hLaborblattdruck
			WinWaitActive	, % "ahk_id " hLaborblattdruck,, 2
			ControlFocus		, ListBox1, % "ahk_id " hLaborblattdruck
			ControlSend		, ListBox1, {Space}, % "ahk_id " hLaborblattdruck
			Sleep 200
			If !VerifiedClick("&OK"		,,, hLaborblattdruck, 2)
				VerifiedClick("Button9"	,,, hLaborblattdruck, 2)
			VerifiedClick("Ja", "Addendum - Albis Quickexporter ahk_class 32770",,, 2)
		}

	}

	If (hCrash || hCrReport) && !WinExist("ahk_class OptoAppClass")
		WinWait, % "ahk_class OptoAppClass",, 30

	if (crash || crreport || lbdruck)
		ToolTIp, % "count:`t" ccount "`ncrash:`t" crash "`ncrreport:`t" crreport "`nLBDruck:`t" Lbdruck , 1000, 1, 8

  ; Timer wieder starten
	If settings.CheckAlbis {
		If !IsObject(fn_AlbisCheck)
			fn_AlbisCheck := Func("CheckAlbisProblems")
		If qexporter.ExportRunning
			SetTimer, % fn_AlbisCheck, -2000
	}

}

AlbisAbsturz(params*)                     	                                           	{

	SetTitleMatchMode, RegEx

	crashsend := {"anonym"     	: "Button1"
						, "senden"      	: "Button2"
						, "speichern"    	: "Button3"
						, "beenden"    	: "Button4"
						, "starten"      	: "Button5"
						, "DSVGO"       	: "static8"
						, "Bemerkung"  	: "Edit1"}

	hwin 		:= param.1 ? param.1 : WinExist("ALBIS\(TM\) ahk_class #32770 ahk_exe AoWDump2.exe")
	crashtime	:= params.2
	idlemsg 	:= params.3

	If hwin {
		ControlSetText, % crashsend.Bemerkung, % idlemsg, % "ahk_id " hWin
		VerifiedCheck(crashsend.starten,,, hWin, 1)
		VerifiedCheck(crashsend.anonym,,, hWin, 1)
		If !VerifiedClick(crashsend.beenden,,, hWin, 2)
			VerifiedClick("&Beenden",,, hWin, 2)
	}

	If !(hAoW := WinExist("AoWDump2 ahk_class #32770 ahk_exe AoWDump2.exe"))
		WinWait, % "AoWDump2 ahk_class #32770 ahk_exe AoWDump2.exe",, 3
	If (hAoW := WinExist("AoWDump2 ahk_class #32770 ahk_exe AoWDump2.exe")) {
		If !VerifiedClick("OK" ,,, hAoW, 2)
			VerifiedClick("Button1" ,,, hAoW, 2)
		WinWaitClose, % "AoWDump2 ahk_class #32770 ahk_exe AoWDump2.exe",, 3
	}

	If !(hUpdate := WinExist("ALBIS ahk_class #32770", "Online\-Update"))
		WinWait, % "ALBIS ahk_class #32770", % "Online\-Update", 2
	If (hUpdate := WinExist("ALBIS ahk_class #32770", "Online\-Update")) {
		If !VerifiedClick("Nein",,, hUpdate, 2)
			VerifiedClick("Button2",,, hUpdate, 2)
		WinWaitClose, % "ALBIS ahk_class #32770", % "Online\-Update", 3
	}

	If !(hReminder := WinExist("Erinnerung\s+online\s+Update ahk_class #32770"))
		WinWait, % "Erinnerung\s+online\s+Update ahk_class #32770",, 2
	If (hReminder := WinExist("Erinnerung\s+online\s+Update ahk_class #32770")) {
		If !VerifiedClick("OK",,, hReminder, 2)
			VerifiedClick("Button1",,, hReminder, 2)
		WinWaitClose, % "Erinnerung\s+online\s+Update ahk_class #32770",, 3
	}

}

WinGetDrucken(dbg:=false)                                                              	{

	WinGet, wList, List, % "Drucken"
	Loop {
		If !(hwnd := wList%A_Index%)
			break
		wTitle 	:= WinGetTitle(hwnd := GetHex(hwnd))
		wClass	:= WinGetClass(hwnd)
		wText 	:= RegExReplace(WinGetText(hwnd), "[\n\r]", "|")
		t 			.= " [" hwnd "]`n" wTitle ", class: " wClass ", Text: `n" wText "`n"
		If (wTitle = "Drucken" && wClass = "#32770" && InStr(wText, "Standort") && InStr(wText, "Druckbereich")) {
			hDrucken := GetHex(hwnd)
			break
		}
	}

	dbg ? SciTEOutput(t) : ""

return hDrucken
}

WinGetLaborblattDruck(dbg:=false)                                                     	{

	WinGet, wList, List, % "Laborblatt Druck"
	Loop {
		If !(hwnd := wList%A_Index%)
			break
		wTitle 	:= WinGetTitle(hwnd := GetHex(hwnd))
		wClass	:= WinGetClass(hwnd)
		wText 	:= RegExReplace(WinGetText(hwnd), "[\n\r]", "|")
		t 			.= " [" hwnd "]`n" wTitle ", class: " wClass ", Text: `n" wText "`n"
		If (InStr(wTitle, "Laborblatt Druck") && wClass = "#32770" && InStr(wText, "Anmerkungen") && InStr(wText, "Probendaten")) {
			hLaborblattdruck := GetHex(hwnd)
			break
		}
	}

	dbg ? SciTEOutput(t) : ""

return hLaborblattdruck
}

;}

; sonstige                                                         	;{

; Dateihandling
; ---------------------------------------------------------------------------------------------
CopyFilesWithProgress(sfiles, destinationPath)                                         	{ ;-- kopiert 7z und zeigt dabei den Kopierfortschritt an

	global hQE, QEhLvExp, CPRG, CPRGPrg
	global WM_USER               := 0x00000400
	global PBM_SETMARQUEE        := WM_USER + 10
	global PBM_SETSTATE          := WM_USER + 16
	global PBST_NORMAL           := 0x00000001
	static dpiF := A_ScreenDPI/96
	static CPRGTxt1, CPRGTxt2, QEhCPRG, CPRGCancel

	If (!IsObject(sfiles) || sfiles.Count()=0)
		return

; Interaktionsmöglichkeiten mit dem Rest der Gui abstellen
	QEDisEnableAll()


; Fortschrittanzeige in die Listview einpassen
	gp  	:= GetWindowSpot(QEhLvExp)

	ltX 	:= gp.CX
	ltY 	:= gp.CY
	prgY	:= Round(gp.H//2)-25


	Gui CPRG: New, % "HWNDQEhCPRG +AlwaysOnTop -DPIScale -Caption +Owner" hQE "  +Parent" QEhLvExp
	Gui CPRG: Default
	Gui CPRG: Color, c5772B7

	Gui CPRG: Font, % settings.fSize.s18 " cWhite Normal", % Addendum.Default.Font
	Gui CPRG: Add	, Text   	, % "xm y" prgY-71 " w" gp.CW-20 " h40 Center Backgroundtrans vCPRGTxt1 "   	, % "["  "0/" sfiles.Count() "]"

	Gui CPRG: Add	, Progress, % "xm+80 y+20 w" gp.CW-240 " h50 vCPRGPrg hwndCPRGhPrg Background264521 c6AB757 E0x20001" , 0  ; -E0x20000

	Gui CPRG: Font, % settings.fSize.s24 " cWhite Normal", % Addendum.Default.BoldFont
	Gui CPRG: Add	, Text   	, % "xm y+20 w" gp.CW-20 " h60 Center Backgroundtrans vCPRGTxt2 "               	, % "kopiere Dateien nach " destinationPath

	Gui CPRG: Font, % settings.fSize.s18 " cWhite Normal", % Addendum.Default.Font
	Gui CPRG: Add , Button  , % "x" gp.CW-140 " y" gp.CH-50 " w100 vCPRGCancel gCPRGHandler"               	, % "Abbrechen"

	Gui CPRG: Show, % "x0 y0 w" gp.CW " h" gp.CH " Hide"                                                   	, % "QuickExporter Dateikopierer mit Fortschrittsanzeige"

	GuiControlGet, cp, CPRG: Pos, CPRGCancel
	GuiControl, CPRG: MoveDraw, CPRGCancel, % "x" gp.W-cpW-80 " y" gp.H-cpH-40

	Gui, CPRG: Show


 ;Callback für Progress einrichten
	address := RegisterCallback("CopyProgressUpdate", "Fast")

 ;Dateien kopieren
	settings.CPRGCancel := false
	For fIdx, file in sfiles 	{

		if settings.CPRGCancel {
			TrayTip, % scriptTitle, % "Kopiervorgang wurde abgebrochen", 8
			break
		}

		SplitPath, file, filename, dir

		GuiControl, CPRG: , CPRGTxt1,  % "[" SubStr("0000000000" fIdx, -1*(StrLen(sfiles.Count())-1)) "/" sfiles.Count() "] " filename


		if FileExist(dest := destinationPath "\" filename)
			FileDelete, % dest

		dllcall("CopyFileExW","wStr", file, "wStr", dest, "uint",address, "Uint",0, "int",0, "int",0)
		if settings.CPRGCancel {
			TrayTip, % scriptTitle, % "Kopiervorgang wurde abgebrochen", 8
			break
		}
		Sleep 500

	}

	settings.CPRGCancel := false
	Gui CPRG: Destroy

; Interaktion mit dem Rest der Gui wieder ermöglichen
	QEDisEnableAll()

}

CopyProgressUpdate(var1lo,var1hi,var2lo,var2hi,var3lo,var3hi,var4lo,var4hi,var5,var6,var7,var8,var9) {

	global CPRG, CPRGPrg

	Gui CPRG: Default
	copied := A_PtrSize=8 ? (var1hi/var1lo)*100 : (var2lo/var1lo)*100
	GuiControl, CPRG: , CPRGPrg, % copied
	;~ ToolTip, % copied, 3000, 300, 10

	if (settings.CPRGCancel = 1) {
		TrayTip, % scriptTitle, % "Kopiervorgang wird nach Abschluß der Übertragung der aktuellen Datei abgebrochen.", 8
		settings.CPRGCancel := 2
	}


return 0
}

Smtra_PrintPDF(printer_name, printer_options, PdfFilePath)                             	{ ;-- PDF Dateien drucken

	smtracmdline	:= "-print-to " q printer_name q " -print-settings " q printer_options q " -exit-when-done" ; Dateiname
	stdoutCMD    	:= settings.SumatraCMD " " smtracmdline " " q . PdfFilePath . q
	dbg ? SciTEOutput(A_ThisFunc "`n" stdoutCMD) : ""
	If dbg && (stdout := StdoutToVar(stdoutCMD))
		SciTEOutput(stdout)

}

FileMovePath(oldPath, newPath, test:=false)                                           	{ ;-- verschiebt ganze Dateiordner um exportierte Daten nach Vorgang zu sortieren

	If newpath && !(newpath ~= "i)^([A-Z]:\\|[\w+\-_]+\\\\)")
		newPath := stdExportPath "\" StrReplace(newPath, stdExportPath "\")
	If oldpath && !(oldpath ~= "i)^([A-Z]:\\|[\w+\-_]+\\\\)")
		oldpath := stdExportPath "\" StrReplace(oldpath, stdExportPath "\")

	; - - - - - - - - - - - - - - -
  ; Funktionsausführung abbrechen
	; -           wenn            -
	; oldPath - leer wäre
	if !oldpath
		return [-1,0,0]
	; oldpath und newpath auf keine Verzeichnisse verweisen
	else if !InStr(FileExist(oldpath), "D") && !InStr(FileExist(newpath), "D") {
		dbg ? SciTEOutput(A_ThisFunc "() | pfad fehlerhaft: " StrReplace(old, stdExportPath "\")) : ""
		return [-2,0,0]
	}
	; newpath - leer wäre
	else if !newPath
		return [-3,0,0]
	; newpath auf ein bereits bestehendes Verzeichnis verweist
	else if InStr(FileExist(newPath), "D") {
		dbg ? SciTEOutput(A_ThisFunc "() | neuer Pfad bereits vorhanden: " StrReplace(newPath, stdExportPath "\")) : ""
		return [1,0,0]			; das wäre wie kopiert
	}
	; oldpath und newpath auf dasselbe Verzeichnis verweisen
	else if (oldpath = newPath) {
		dbg ? SciTEOutput(A_ThisFunc "() | pfade identische: " StrReplace(newPath, stdExportPath "\")) : ""
		return [1,0,0]			; das wäre wie kopiert
	}

	; der neue Pfad ist noch nicht angelegt
	If !InStr(FileExist(newpath), "D") {

	  ; z.B. m.std: D:\....., m.sub: Hr. Dr. Holzler, Frank [00700700700], m.shp: leer oder versendet, , m.pat: (12345) PatNachname, PatVorname
		If !RegExMatch(newPath, "Oi)((?<std>\w:\\.+?)\\(?<sub>([^\\]+\]|Patient|verschoben)))(\\(?<shp>[^\\]+))*\\(?<pat>\(\d+\).+)", m) {
			msg := "Problem mit RegExMatch: " m.std ", " m.sub ", " m.shp ", " m.pat
			dbg ? SciTEOutput(A_ThisFunc "() | rxmatch: " msg) : ""
			return [-5,0,0]
		}

		If !FilePathCreate(dir1 := m.std "\" m.sub) {
			fpError .= "0x1 `n" dir1 "`n"
		}
		If (m.sub ~= "i)(\[\d+\]|Patient)" && !FilePathCreate(dir1 "\versendet")) {
			fpError .= "0x2 `n" dir1 "\versendet`n"
		}
		If (m.shp && !FilePathCreate(dir1 "\" m.shp)) {
			fpError .= "0x4 `n" dir1 "\" m.shp "`n"
		}
		If !FilePathCreate(dir1 (m.shp ? "\" m.shp : "") "\" m.pat) {
			fpError .= "0x8 " dir1 (m.shp ? "\" m.shp : "") "\" m.pat "`n"
		}

		If InStr(FileExist(newpath), "D") {
			if RegExMatch(newPath, "\[(?<snr>\d+)\]", b) {
				; Addendum.PersonRx			:= "(Herr\s*|Frau\s*|Hr\.\s*|Fr\.\s*|Prof.\s*|Priv.\s*|Doz\.\s*|Dr\.\s*|med\.\s*)"
				If !FileExist(adressfilepath := m.std "\" m.sub "\Adresskopf " RegExReplace(m.sub, "i)" Addendum.PersonRx) ".txt")
					WriteAdressTextFile(adressfilepath, bsnr)
			}
		}
		else {
			MsgBox, 0x1004, % scriptTitle, % "Die Datenverzeichnisse [" fpError "] konnten nicht angelegt werden`n. Drück ja für weiter. Nein für Abbruch"
			IfMsgBox, No
				return [-6, fpError, 0]
		}


	}


	cmd_XCOPY	:= "XCOPY " . q . oldpath "\*.*" . q . " " . q . newPath . "\" . q . " /Y /C /E /H /R /K /I" (test ? " /L)" : "")
	stdout_XCOPY := StdoutToVar(cmd_XCOPY)
	If !RegExMatch(stdout_XCOPY, "i)(\d+)\s+Datei.*kopiert", sXCopy)
		sXCopy1 := SXCopy
	If !sXCopy1
		sxCopy := 1

	filesmissed := 0
	If !test
		Loop, Files, % oldpath "\*.*"
			If !FileExist(newPath "\" A_LoopFileName) {
				fileList .= A_LoopFileName "`n"
				filesmissed += 1
			}

	If filesmissed {
		dbg ? SciTEOutput(A_ThisFunc "() | " filesmissed " Dateie(n) wurde nicht kopiert.`nVerzeichnis (" oldpath ") ") : ""
		dbg ? SciTEOutput(A_ThisFunc "() | Dateiliste:`n" fileList) : ""
		return [-7, fileList, 0]
	}
	else if !filesmissed {

		rmDIR := true
		cmd_RMDIR 	:= "RMDIR /S /Q " . q . oldpath "\" . q
		if !test {
			RunWait % ComSpec " /c " cmd_RMDIR " >"  A_Temp "\rmdirLog.txt"
			If (stdout_RMDIR := FileOpen(A_Temp "\rmdirLog.txt", "r", "UTF-8").Read())
				SciTEOutput(A_ThisFunc "() | RMDIR: " stdout_RMDIR)
			If InStr(FileExist(oldpath), "D") {
				SciTEOutput("Das bisherige Verzeichnis '" StrReplace(oldpath, stdExportPath "\") "' konnte nicht entfernt werden." )
				rmDIR := false
			}
		}

	}

return  [sXCopy1, filesmissed, rmDIr]
}

CreateShippingLetter(bsnr, shippingDate)                                               	{	;-- erstellt einen Versandbrief aus einer Vorlage

	; Funktion macht aus einem html/svg Dokument einen Begleitbrief zum Datenversand. Dieser wird zusätzlich ins PDF Format  konvertiert.
	; Man kann den Brief klassisch ausdrucken oder mit auf eine DVD brennen oder als HTML-EMail (natürlich nur mit verschlüsselten Daten versenden).

	static templateDir, templatePath

; Verzeichnisstruktur zusammenstellen
	shipmentYMD			:= SubStr(shippingDate, 1, 4) "-" SubStr(shippingDate, 5, 2) "-" SubStr(shippingDate, 7, 2)
	shipmentDMY			:= SubStr(shippingDate, 7, 2) "." SubStr(shippingDate, 5, 2) "." SubStr(shippingDate, 1, 4)
	recip           := SQLiteGetRecipient(bsnr, 2)
	LetterSavePath 	:= stdExportPath "\" (bsnr > 3 ? recip.RecipientName " [" bsnr "]" : "Patient")
	LetterFileName 	:= "Versandschreiben_" RegExReplace(recip.RecipientName , "^.*?\..*?\.\s*") "_" shipmentYMD
	If !templateDir {
		templateDir   	:= settings.paths.resources
		templatePath   	:= templateDir "\Vorlage_Datenversand.html"
	}

	;~ SciTEOutput(LetterSavePath)

; nicht vorhandene Verzeichnisse erkennen
	If !InStr(FileExist(LetterSavePath), "D") {
		settings.lastError := A_ThisFunc "() | kein Exportverzeichnis"
		return -1 		; kein Export Verzeichnis
	}
	else if !InStr(FileExist(templateDir), "D") || !InStr(FileExist(templateDir "\images"), "D") {
		settings.lastError := A_ThisFunc "() | kein Briefvorlage vorhanden in " templateDir
		return -2			; Vorlagen nicht vorhanden
	}

; Erstellen des HTML Dokumentes
	If !FileExist(fHTML:= LetterSavePath "\" LetterFileName ".html") {

		; XML-COM Object
			doc := ComObjCreate("Msxml2.DOMDocument.6.0")

		; Umgehung einer Sicherheitseinstellung.
		; Die HTML-Datei enthält externe DTD-Referenzen.
		; Diese können nicht aufgelöst werden.
		; Die Verarbeitung von DTD ist vollständig
		; deaktiviert, um den Fehler zu umgehen.
			doc.async := false
			doc.validateOnParse := false
			doc.resolveExternals := false
			doc.setProperty("ProhibitDTD") := False

		; ACHTUNG:
			doc.loadXML(xmlstr := FileOpen(templatePath, "r", "UTF-8").Read())

		; Überprüfen, ob das Laden der XML erfolgreich war
			if (doc.parseError.errorCode != 0) {
				settings.lastError := A_ThisFunc "() | Fehler beim Laden der XML-Datei:`n" doc.parseError.reason "[line: " doc.parseError.line " pos: " doc.parseError.linepos "]"
				return -3
			}

		; Bildpfade encodieren, sonst werden die Bilder später nicht angezeigt.
		; der vollständige Pfad ist notwendig damit die fertige HTML Datei "umziehen" kann z.b. in ein anderes Verzeichnis
			pic := [templateDir "/images/_Image_Unterschrift.png"
					 ,  templateDir "/images/_image_QRCode.png"
					 ,  templateDir "/images/_Image_PraxisLogo.svg"]
			For each, path in pic
				path := UrlEncode(path)

		; Werte ändern und Attribute setzen in der HTML Datei
			svgBase := "//*[name()='html']//*[name()='body']//*[name()='svg']"
			svgPathValues := [  {"xp" : "[name()='g' and @id='Unterzeichnerbereich']//*[name()='text' and @id='text_NameUnterzeichner']"     	, "value" : settings.sender.name}
								    		, {"xp" : "[name()='g' and @id='EMailfeld']//*[name()='text' and @id='Text_EMailadresse']"	                    , "value" : settings.sender.Mail}
								    		, {"xp" : "[name()='text' and @id='OrtDatum']//*[name()='tspan' and @id='tspan606']"	                          , "value" : settings.sender.Ort ", den " shipmentDMY}
								    		, {"xp" : "[name()='g' and @id='Empfängerfeld']//*[name()='text' and @id='text_AnrTitelName']"	                , "value" : recip.anrname}
								    		, {"xp" : "[name()='g' and @id='Empfängerfeld']//*[name()='text' and @id='text_StrasseNr']"	                    , "value" : recip.strasseNr}
								    		, {"xp" : "[name()='g' and @id='Empfängerfeld']//*[name()='text' and @id='text_PLZOrt']"	                      , "value" : recip.PLZOrt}
								    		, {"xp" : "[name()='g' and @id='Unterzeichnerbereich']//*[name()='image' and @id='_Image_Unterschrift']"        , "value" : pic.1}
								    		, {"xp" : "[name()='g' and @id='EMailfeld']//*[name()='image' and @id='_Image_QRCode']"	                        , "value" : pic.2}
								    		, {"xp" : "[name()='image' and @id='Logo']"	                                                                    , "value" : pic.3}]

			For each, data in svgPathValues {

				node := doc.selectSingleNode(xp := svgBase . "//*" . data.xp)
				oldvalue := node.text

				If IsObject(node) {
					if InStr(xp, "image") {
						result := node.getAttribute("xlink:href")
						node.setAttribute("xlink:href", data.value)
					}
					else
						node.text := data.value
				}
				else
					oldvalue := "noNode"

				t .= data.xp "`n  " oldvalue " ==> " data.value (result ? " (" result ")" : "") "`n`n"

			}

		; Debugausgabe in die Standardkonsole
			If dbg
				FileAppend, % t "`n", *

			doc.Save((LetterFilePath := LetterSavePath "\" LetterFileName ".html"))
			doc := ""

	}

	FileGetSize, fhtmlfileSize, % fHTML

; PDF-Version erstellen
	IF (FileExist(fHTML) && fhtmlfileSize > 0) {

		If !FileExist(fPDF:= LetterSavePath "\" LetterFileName ".pdf")
			RunWait, % "msedge.exe --headless --disable-pdf-tagging --no-pdf-header-footer --print-to-pdf=" q fPDF q " " q "file://" fHTml q
		If FileExist(fPDF)
			PdfFilePath := FileExist(fPDF) ? fPDF : ""
	}
	else {
		settings.lastError := A_ThisFunc "() | Der HTML Versandbrief konnte nicht erstellt werden."
		return -4
	}

; Ausdrucken des Briefes
	FileGetAttrib, fatrib, % fPDF
	IF FileExist(fPDF) && !InStr(fatrib, "A") {
		settings.printLetter := GuiControlGet("QE", "", "QEVBIprint")
		If settings.printLetter &&  {

			Gui, QE: Submit, NoHide
			printer_name 	:= QEVBIprinter
			printers    	:= " " GetPrinters() " "

			If Trim(printers) {

				MsgBox, 0x1004, % scriptTitle, %  "Sie können jetzt den Begleitbrief ausdrucken.`n`n"
																				. "Empfänger: " recip.Geschlecht " " recip.anrname "`n"
																				. (printer_name && printers ~= "\W" printer_name "\W" ? "Gerät: " printer_name : "")
				IfMsgBox, Yes
				{
					If !printer_name {
					; ### kommt später eine GUi
						settings.lastError := A_ThisFunc "() | Es wurde kein Drucker ausgewählt."
					}
				; Drucken
					else
						msEdge_SaveAsPDF(letterFilePath, printer_name)
				}

			}
			else
				settings.lastError := A_ThisFunc "() | Es stehen keine Drucker zur Auswahl zur Verfügung."

		}
	}

return {"htmlPath":fHTML, "pdfPath":(FileExist(fPDF) ? fPDF : "")}
}

CheckAdressTextFiles()                                                                 	{	;-- prüft ob die Adressdaten des Arztes im Verzeichnis liegen

	Loop, Files, % stdExportPath "\*.*", D
	If RegExMatch(A_LoopFileFullPath, "O)\\.+\\(?<doc>.+)?\[(?<bsnr>\d+)\]$", m)
		If !FileExist(adressfilepath := A_LoopFileFullPath "\Adresskopf " RegExReplace(m.doc, "i)(Herr\s*|Frau\s*|Hr\.\s*|Fr\.\s*|Dr\.\s*|med\.\s*)") "[" m.bsnr "].txt")
			WriteAdressTextFile(adressfilepath, m.bsnr)

}

WriteAdressTextFile(adressfilepath, bsnr) 					   						    			          	{	;-- hinterlegt eine Textdatei mit den Adressdaten des Arztes im Verzeichnis

	If (adresstext := SQLiteGetRecipient(bsnr, true))
		FileOpen(adressfilepath, "w", "UTF-8").Write(adresstext)

}

xDataMoveDirs(subdir:="")                                                             	{

	dbg ? SciTEOutput("path: " stdExportPath (subdir ? "\" subdir : "") "\*.*") : ""
	xDataPath     	:= stdExportPath "\xData"
	xDBasePatPath 	:= stdExportPath (subdir ? "\" subdir : "")
	xDPatPaths := Object()

	Loop, Files, % xDBasePatPath "\*.*", R
	{
		If (SubStr(A_LoopFileDir, -4) = "xData") {
			If !IsObject(xDPatPaths[A_LoopFileDir]) {
				xDPatPaths[A_LoopFileDir] := []
				xDPatPaths[A_LoopFileDir].1 := 1
				xDPatPaths[A_LoopFileDir].2 := xDPatPaths.Count()
			}
			else
				xDPatPaths[A_LoopFileDir].1	+= 1

		}
	}

	For xDPatPath, xcounts in xDPatPaths {

		shortendPath := SubStr(xDPatPath, StrLen(xDBasePatPath)+2, StrLen(xDPatPath)-StrLen(xDBasePatPath)-1)
		shortendPath := RegExReplace(shortendPath, "i)(\\versendet|\\xData)*")
		xDataShortendPath := xDataPath "\" shortendPath
		copyerror := ""
		cmd_xcopy := "xcopy " . q . xDPatPath . "\#" . q . " " . q . xDataShortendPath . "\#" . q . " /e"

		If (DirExists := InStr(FileExist(xDataShortendPath), "D") ? true : false) {

			Loop, Files, % xDPatPath "\*.*"
			{

				copyfile := true
				If FileExist(xDataShortendPath "\" A_LoopFileName) {
					cmd_filecompare := "fc " . q . xDPatPath "\" A_LoopFileName . q . " " . q . xDataShortendPath "\" A_LoopFileName . q
					res := StdoutToVar(cmd_filecompare)
					dbg ? SciTEOutput(A_ThisFunc "() | filecompare: `n" cmd_filecompare "`n   " StrReplace(res, "`n", "`n    ")  "`n") : ""
					If InStr(res, "Keine Unterschiede gefunden")
						copyfile := false
				}
				If copyfile {
					cmd_xcopy_rpl := StrReplace(cmd_xcopy, "#", A_LoopFileName)
					res := StdoutToVar(cmd_xcopy_rpl)
					dbg ? SciTEOutput(A_ThisFunc "() | " cmd_xcopy_rpl "`n   " res  "`n") : ""
					If !RegExMatch(res, "[1-9]\d*\s+Datei\(en\)\s+kopiert")
						copyerror .=  xDataShortendPath "\" A_LoopFileName "`n"
				}

			}

		}
		else {

			cmd_xcopy_rpl := StrReplace(cmd_xcopy, "#", "")
			res := StdoutToVar(cmd_xcopy_rpl)
			dbg ? SciTEOutput(A_ThisFunc "() | " cmd_xcopy_rpl "`n   " res  "`n") : ""
			If !RegExMatch(res, "[1-9]\d*\s+Datei\(en\)\s+kopiert")
				copyerror .=  res "`n"

		}

		If !copyerror {
			cmd_RMDIR := "RMDIR /S /Q " q . xDPatPath "\" . q
			;~ res := StdoutToVar(cmd_RMDIR)
			dbg ? SciTEOutput(A_ThisFunc "() | remove dir: " cmd_RMDIR) : ""
		}
		;~ SciTEOutput("cmdline (" DirExists ") :`n   " cmd_xcopy)
		;~ Run, % "xcopy "

	}


return
}

RemoveWrongFiles(stdPath)                                                             	{

	rmfFunc := "running"
	wrongfiles := 0
	Loop, Files, % stdPath "\*.*", R
	{
		If RegExMatch(A_LoopFileName, "i)^\s*Laborbefunde_\d+\s+\-.*\.pdf") {
			wrongfiles ++
			dbg ? SciTEOutput("[" SubStr("00000" wrongfiles, -4) "] " StrReplace(A_LoopFileFullPath, stdPath "\")) : ""
		}
	}

	rmfFunc := "ready"

} ;RWF-End

FileExistZ(File, C:=0, P*)                                                            	{ ;-- FileExistZ v0.91, By SKAN on D353/D355 @ tiny.cc/fileexistz

	/*	Description

			The variadic parameters of FileExistZ() are meant for changing the attributes of a file (or folder).
			Note: Attributes will be set only if Check parameter is omitted or 0
			FileExistZ() can Add, Remove, Toggle a file attribute or completely replace existing attributes (when possible).


			READONLY              :=      0x1  ;       1
			HIDDEN                :=      0x2  ;       2
			SYSTEM                :=      0x4  ;       4
			DIRECTORY             :=     0x10  ;      16
			ARCHIVE               :=     0x20  ;      32
			DEVICE                :=     0x40  ;      64
			NORMAL                :=     0x80  ;     128
			TEMPORARY             :=    0x100  ;     256
			SPARSE_FILE           :=    0x200  ;     512
			REPARSE_POINT         :=    0x400  ;    1024
			COMPRESSED            :=    0x800  ;    2048
			OFFLINE               :=   0x1000  ;    4096
			NOT_CONTENT_INDEXED   :=   0x2000  ;    8192
			ENCRYPTED             :=   0x4000  ;   16384
			INTEGRITY_STREAM      :=   0x8000  ;   32768
			VIRTUAL               :=  0x10000  ;   65536
			NO_SCRUB_DATA         :=  0x20000  ;  131072
			RECALL_ON_OPEN        :=  0x40000  ;  262144
			RECALL_ON_DATA_ACCESS := 0x400000  ; 4194304

			If FileExistZ("D:\")
				MsgBox, The drive exists.

	*/

Local K,V,N,A, M := A := DllCall("GetFileAttributes", "Str",File, "Int")


  If (P.Count() and A != -1 and not C)
      For K,V in P
          N := StrSplit(V,"="," ",2), K := N[1], V := Round(N[2])
        , M := K="+" ? M|V : K="-" ? M&~V : K="^" ? M^V : K=":" ? V : M|V
  A := (A != -1  and A != M)  ? DllCall("SetFileAttributes", "Str",File, "Int",M)
                              ? DllCall("GetFileAttributes", "Str",File, "Int") : -1  : A
Return Format("0x{2:06X}", A := A!=-1 and C ? A & C=C ? A : -1 : A, A=0 ? 8 : A=-1 ? 0 : A)
}

GetStdExportPaths()                                                                   	{ ;-- Verzeichnisstrukturen im Standardexportverzeichnis ermitteln

	;die Verzeichnisstruktur des Standardexportpfades wird ermittelt
	; stdExportPath ist global

	paths := {}
	Loop, Files, % stdExportPath "\*.*", RD
	{
		if (A_LoopFileDir ~= "i)xData")
			continue
		shortpath := StrReplace(RegExReplace(A_LoopFileDir, "\\\(\d+\)\s+.*$"), stdExportPath "\")
		if !paths.haskey(shortpath) {
			paths[shortpath] := 1
		}
		else
			paths[shortpath] += 1

	}

return paths
}

FindMissedPaths(nofpath)                                                               	{	;-- sucht nach "verschwundenen" Exportverzeichnissen

	; und repariert in der sqlite Datenbank automatisch die nicht verlinkten Verzeichnisse

	global hMF, MFEdit, MFClose, MFConfirm, MFInfo1, MFPrg, MFGBox, MFCDirs
	static tpaths := {}
	static Erklaertext
	If !Erklaertext
		Erklaertext =
		(LTrim
		Was ist passiert?
		Dieses Fenster wird dann angezeigt, wenn Fehler in der QuickExporter SQLite Datenbank festgestellt wurden. Die Funktion korrigiert z.B. nicht vorhandene oder falsch verknüpfte Verzeichnisse. Diese Fehler sollten selten auftreten, aber sie kommen dennoch vor. Diese Funktion hilft Datenbankfehler die aus fehlerhaften Programmcode entstanden sind zu korrigieren.
		)

	if (!IsObject(nofpath) || (npC := nofpath.Count()) = 0)
		return

	matchPaths  := {}
	paths     	:= GetStdExportPaths()	; vorhandene Verzeichnisse ermitteln

 ; versucht ein Verzeichnis mit dem Namen [PatID] Nachname, Vorname in den Unterverzeichnissen zu finden
	For each, PatID in nofpath {
		PatPath := "(" PatID ") " cPat.NAME(PatID, false)
		matchPaths[PatID] := {"PatPath" : PatPath, "PatName":cPat.NAME(PatID, false), "folders":[]}
		For basePath, counts in paths {
			if Instr(FileExist(stdExportPath "\" basePath "\" PatPath), "D")
				matchPaths[PatID].folders.Push(basePath "\" PatPath)
		}
	}

	SciTEOutput(cJSON.Dump(matchPaths, 1))

	; Anzeigetext erstellen
	x .= (npC = 1 ? "Für einen" : "Für " npc) " Patienten waren keine Verzeichnisdaten hinterlegt.`n"
	x	.=  StrReplace(Format("{:" StrLen(x) "}", ""), " ", "-") "`n`n"

	Gui, QE: Default
  Gui, QE: ListView, QELvExp

  ; weitere Ausgaben erstellen
	For PatID, data in matchPaths {

		; mehrere Verzeichnisse (Pfadkopien mgl.weise), das korrekte Verzeichnis muss manuell gewählt werden
		if (data.folders.count() > 1) {
			pathcopies += 1
			tPaths[PatID] := data
			t	.= "  " Format("{:-30}", data.PatName) "korrespondierende Verzeichnisse:`n"
			For fidx, folder in data.folders
				t.= "  " Format("{:-30}", " ") fidx ". " folder "`n"
		}

		; exakt ein Verzeichnis, dann kann die Verknüpfung gleich in die Datenbank geschrieben werden
		else if (data.folders.count() = 1) {
			pathcorrections += 1
			SQLUpdateItem("Exports", "ExportPath", PatID, stdExportPath "\" data.folders[1])
			Loop % LV_GetCount() {
				LV_GetText(lvPatID, A_Index, 1)
				If (lvPatID = PatID) {
					LV_Modify(A_Index, "Col11", data.folders[1])
					break
				}
			}
			p .=  "  " Format("{:-30}", data.PatName) " " data.folders[1] "`n"
		}

		; kein Verzeichnis vorhanden
		else if (data.folders.count() = 0) {

			pathslost += 1
			n .= "  " Format("{:-30}", data.PatName)

		; lese die Empfänger aus (Arzt/Patient/Angeöriger), ermittle das Versandmedium und -route
			recip  := SQLiteGetCell("EXPORTS", "Recipient", "PatID",  PatID)
			medium := SQLiteGetCell("EXPORTS", "Medium"   , "PatID",  PatID)
			route  := SQLiteGetCell("EXPORTS", "Route"    , "PatID",  PatID)

		; wenn der Empfänger 0, leer oder keinem bekanntem Empfänger entspricht, dann wird der Empfänger "Patient" gewählt und Medium und Route sind "ungeklärt"
			RecipExists := false
			If (recip ~= "^\s*\d+\s*$")
				RecipExists := SQLiteRecordExist("RECIPIENTS", "BSNR", Trim(recip), "returnOnlyExist")

			if (RecipExists = false) {

				; SQLUpdateItem(Table    , ColumnName , rowID, newValue)
				SQLUpdateItem("EXPORTS", "Recipient", PatID, 1)
				SQLUpdateItem("EXPORTS", "Medium" 	, PatID, "❔ ungeklärt")
				SQLUpdateItem("EXPORTS", "Route"	  , PatID, "❔ ungeklärt")

				; Daten auch in der Listview ändern
				Loop % LV_GetCount() {
					LV_GetText(lvPatID, A_Index, 1)
					If (lvPatID = PatID) {
						LV_Modify(A_Index, "Col6", "Patient")
						LV_Modify(A_Index, "Col7", "❔ ungeklärt")
						LV_Modify(A_Index, "Col8", "❔ ungeklärt")
						break
					}
				}

				; Textausgabe erweitern
				n .= "  - unbekannten Empfänger mit 'Patient' ersetzt"

			}

			n .= "`n"

		}

	}

	If pathcorrections {
		x .= (pathcorrections=1 ? " Bei einem " : "Für " pathcorrections) " Patienten konnte die Pfad-Verlinkung wiederhergestellt werden.`n"
		x .= p "`n"
	}
	if pathslost  {
		x .= "Für " (pathslost=1 ? "diesen Patienten konnte kein korrespondierendes Verzeichnis" : pathslost " Patienten konnten keine korrespondierenden Verzeichnisse" ) " ermittelt werden.`n"
		x .= n "`n"
	}

	xlen := StrSplit(x, "`n").Count()-1
	xlen := xlen > 30 ? 30 : xlen

	Gui, MF: New  	, % "+hwndhMF +Owner" hQE
	Gui, MF: Font 	, s9 q5 Normal cBlue, Arial
	Gui, MF: Add  	, Edit, xm ym w1000 r4 ReadOnly, % Erklaertext

	Gui, MF: Font 	, s11 q5 Normal cBlack, Consolas

	if pathcorrections || pathslost {
		Gui, MF: Add 	, Edit, % "xm y+10 w1000 r" xlen " ReadOnly vMFEdit hwndMFhEdit", % Rtrim(x, "`n")
		ControlSend,, {Down}, % "ahk_id " MFhEdit
	}

	if pathcopies  {
		Gui, MF: Add 	, Groupbox, % "xm y+5 w1000 vMFGBox"
		Gui, MF: Add 	, Text, % "xp+5 yp+10 w990 vMFCDirs", % (pathcopies=1 ? "Ein Patient hat" : pathcopies " Patienten haben") " mehrere korrespondierende Verzeichnisse."
		For PatID, data in tpaths {

			Gui, MF: Add, Text, % "xm+10 y+10"             	, % "  " Format("{:-30}", data.PatName) "korrespondierende Verzeichnisse:"

			folders := data.folders
			data.cbRef := Object()

			For fidx, folder in folders {
				Gui, MF: Add, Text   	, % "xm+10 y+" (fidx=1 ? 3: 1)        	, % "  " Format("{:-30}", " ") Format("{:2}", fidx) ". "
				Gui, MF: Add, Checkbox, % "x+5 hwndhCB"    	, % folder
				data.cbRef[hCB] := fidx
			}
		}

		color := "D0FBA6"
		Gui, MF: Add, Progress, % "xm y+10 w1000 h25 vMFPrg c" Color " Background" Color
		Gui, MF: Font, s10 q5 Normal, Arial
		Gui, MF: Add, Text   	, % "xp+5 yp+5 w1000 r2 vMFInfo1 Backgroundtrans", % "Wählen Sie das passende Verzeichnis durch Setzen eines Haken aus. "
																						              			. "Bestätigen Sie durch Druck auf 'Änderungen übernehmen'.`n"
																			              						. "Es werden keine Verzeichnisse auf der Festplatte verändert!`n"
    cp := GuiControlGet("MF", "Pos", "MFInfo1")
		dp := GuiControlGet("MF", "Pos", "MFGBox")
		GuiControl, MF: MoveDraw, MFPrg	, % "h" cp.H+10
		GuiControl, MF: MoveDraw, MFGBox, % "h" cp.Y - dp.Y + 5
	}

	Gui, MF: Font 	, s10 q5 Normal, Arial
	Gui, MF: Add  	, Button, xm y+20 vMFConfirm 	gMFGuiClose, % "Änderungen übernehmen"
	Gui, MF: Add  	, Button, x+50    vMFClose  	gMFGuiClose, % "Schließen ohne Änderungen zu übernehmen"

	Gui, MF: Show	, AutoSize, % "QuickExporter - Autoverlinkung"

	;~ SciTEOutput(x)
	;~ SciTEOutput(cJSON.Dump(matchPaths,1))

return

MFGuiClose:
	Gui, MF: Destroy
return
}


; Verschlüsselung
; ---------------------------------------------------------------------------------------------
generatePassword(pwdlength:=40) {

	static pwdchars 	:= " 01234567890|!§$%&/()=?{[]}\|+*#'~,.;:-_<>µ|@€µ|abcdefghijklmnopqrstuvwxyzäöüßáéíóúàèìòù|ÄÖÜÁÉÍÓÚÀÈÌÒÙABCDEFGHIJKLMNOPQRSTUVWXYZ|"

	password := ""
	char := StrSplit(pwdchars)

	Loop 30 {
	 Random, charNR1, 1, % char.Count()
	 Random, charNR2, 1, % char.Count()
	 charNr := Round((charNR1+charNR2)/2)
	 password .= char[CharNr]
	}

return Trim(password)
}

StringDecoder(string:="", phrase:="", decrypt:= true)                             	   	{	;-- Passwords are stored encrypted

		; string 		- 	text to be encoded,[use a single *** to delete registry key]
		; assoc  		- 	an association name for the name key in the Windows Registry
		; phrase 		- 	a ready-made string for encoding, leave empty if you want to generate a new encoding string
		; 							it is saved automatically. The passed encoding string (phrase) always takes precedence over
		;								the string stored in the Windows registry.
		;


	  ; checks
		static dbgtxt, dbgcount:=0, regVal := Object()
		static KeyName 	:= "HKEY_CURRENT_USER\SOFTWARE\QuickExporter"
		static assoc  	:= "UserPWD"
		;~ static phrase 	:= "getbackthisstring"


		If InStr(string, "***")  {
				res := RegistryWrite("REG_SZ", KeyName , assoc, "")
				string := StrReplace(string, "*")
				regVal[assoc] := ""
				If !String
					return
			}

		If !phrase {
			IniRead, phrase, % settings.paths.QEExporterINI, % "Script", % "UserPWD_Phrase"
			phrase := Trim(phrase) ~= "i)^Error$" ? "" : Trim(phrase)

		; generates a random phrase string
			If !phrase {
				enc := []
				Loop 3  {
					i := A_Index
					Loop 20 {
						Random, rndVal, 32, 126
						enc[i] .= Chr(rndVal)
					}
				}
				Loop 20 {
					Random, i, 1, 3
					phrase .= StrSplit(enc[i])[A_Index]
				}
			}


			If phrase
				IniWrite, % phrase, % settings.paths.QEExporterINI, % "Script", % "UserPWD_Phrase"

		}

		If !decrypt {   ; Encrypting

			if string
				regVal[assoc] := encstr := ""
			else
				String := generatePassword(40)

			regVal[assoc] := Crypt.Encrypt.StrEncrypt(string, phrase, 7, 6)
			res := RegistryWrite("REG_SZ", KeyName, assoc, regVal[assoc])

			return (regVal[assoc])

		}
		else if decrypt {	; decrypting

		 ; übergebener String so decodiert werden
			If String {
				encstr := String
				String := ""
			}

		; reads the encrypted string from registry
			if !encstr {
				If !regVal[assoc]
					regVal[assoc] := encstr := RegistryRead(KeyName, assoc)
				else if regVal[assoc]
					encstr := regVal[assoc]

				if !encstr                           ; no encoded string found and decode mode is set then return
					return

			}

			;~ SciTEOutput("ep: "  encstr ",  " phrase)
			return Crypt.Encrypt.StrDecrypt(encstr, phrase, 7, 6)  ; das entschlüsselte Passwort

		}


}


; Callback Funktionen
; ---------------------------------------------------------------------------------------------
ZeigsMir(param*)                                                                       	{

	global hQE
	global qexporter

	static paramLast, pretextLast, pretext, timestr, newcall:=1
	static mYtime, selfcall := false

	Gui, QE: Default

	If (param.Count() = 1 && !InStr(param.1, "selfcall=")) {
		ToolTip,,,, 1
		ToolTip,,,, 2
		thour 	:= StrSplit(timestr, ":").1
		tmin  	:= StrSplit(timestr, ":").2
		tsec		:= StrSplit(timestr, ":").3
		pretext 	:= "...Vorgang abgeschlossen"
		pretext 	.= (timestr ? (" nach " (thour ? thour " Stunde" (thour >1 ? "n " : "")  :"") . (tmin ? (!tsec ? "und ": (thour ? ", " : "")) tmin " Minuten " : "") .  (tsec ? (thour || tmin ? "und " : " ") tsec " Sekunden" : "")) : "")
		;~ SciteOutput("p: " timestr " / " pretext)
		SB_SetText(pretext, 1) ; nach 1 Stunde 12 Min. und 14 Sek.
		SB_SetProgress(0, 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText("", 3)
		SB_SetText("", 4)
		SB_SetText("`t00:00:00", 5)
		paramLast := timestr := ""
		mYtime := 0
		SetTimer, ZeigsMirOff, Off
		return
	}

	If !mYTime
		mYTime := A_TickCount

	If param.7 && !(param.7 = 0) {
		timestr := param.7
		SB_SetText("`t" timestr, 5)
	}
	else  {

		Thisduration := TimeFormatEx(Round((A_TickCount - mYTime)/1000 ), true)
		SB_SetText("`t" Thisduration, 5)

		If (param.1 = "selfcall=on" && qexporter.ExportRunning) {
			selfcall := true
			fn_ZeigsMir := Func("ZeigsMir").Bind("selfcall=on")
			SetTimer, % fn_ZeigsMir, -1000
			return
		}
		else If (param.1 = "selfcall=off" || (selfcall && !qexporter.ExportRunning)) {
			selfcall := false
			return
		}

	}


	If !(param.6~="^\d+$") && (param.6 != pretextLast) {
		pretext :=  !param.6 ? pretextLast : param.6
		param.6 := ""
	}
	If (paramLast != param.1 || pretextLast != pretext) {
		pretextLast := pretextLast != pretext 	? pretext 	: pretextLast
		paramLast := paramLast != param.1? param.1	: paramLast
	}

	SB_SetText(pretextLast " - " paramLast, 1)

	SetTimer, ZeigsMirOff, % -1*3*60*1000

	Gui, QE: Default

	If (param.1 ~= "i)exportiere\sDaten") {

		SB_SetProgress(0, 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText("", 3)
		SB_SetText("", 4)

	}
	else If (param.1 ~= "i)DBASE") {

		SB_SetProgress(Round(param.2*100/ param.3), 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText("Leseposition: " SubStr("00000000" param.2, param.4) "/" param.3, 3)
		SB_SetText("Datensätze gefunden: " param.5, 4)

	}
	else if (param.1 ~= "i)BEFTEXTE\sauslesen") {

		SB_SetProgress(Round(param.4*100/ param.5), 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText("Leseposition: " param.4 "/" param.5, 3)
		SB_SetText("Datensätze gefunden: " param.2, 4)

	}
	else if (param.1 ~= "i)BEFTEXTE") {

		SB_SetProgress(Round(param.4*100/ param.5), 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText("Leseposition: " param.4 "/" param.5, 3)
		SB_SetText("Datensatz: " SubStr("00000000" param.2, -1*StrLen(param.3)) "/" param.3, 4)

	}
	else if (param.1 ~= "i)Dokumentexport") {

		SB_SetProgress(Round(param.2*100/ param.3), 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText("Karteikartenposition: " param.2 "/" param.3, 3)
		SB_SetText("exportiere Dokumente: " (param.4 ? param.4 : 0), 4)

	}
	else if (param.1 ~= "i)Datei:") {

		SB_SetText(param.1 " " param.2 ( param.3 != param.2 ? "/" param.3 : "") , 1)
		If  ( param.3 != param.2)
			SB_SetProgress(Round(param.2*100/ param.3), 2, "+Smooth cFFBB00 Background0000FF", hSBar)

		SB_SetText("Datengröße: " param.4, 3)
		If (!param.5 || !param.6)
			SB_SetText("Ärzte: " param.5.1 ", Pat.: " param.5.2, 4)

	}
	else {

		SB_SetProgress(Round(param.2*100/param.3), 2, "+Smooth cFFBB00 Background0000FF", hSBar)
		SB_SetText(param.2 "/" param.3 	, 3)
		SB_SetText(param.4 "/" param.5 	, 4)

	}




return
ZeigsMirOff: ;{
	ToolTip,,,, 1
	ToolTip,,,, 2
	paramLast := ""
	timestr := ""
	SB_SetText("...Vorgang abgeschlossen", 1)
	SB_SetProgress(0, 2, "+Smooth cFFBB00 Background0000FF", hSBar)
	SB_SetText("", 3)
	SB_SetText("", 4)
	SB_SetText("`t00:00:00", 5)
return ;}
}

showThreadVar()                                                                       	{
	;~ global threadRWF

	threadRWF.ahkassign("fIdx", "fileIndex")
	fileIndex := threadRWF.ahkgetvar.fIdx
	ToolTip, % "wrongfiles: " threadRWF.ahkgetvar.wrongfiles "`fileIndex: " fileIndex, 1000, 1, 3

	;~ If (threadRWF.ahkgetvar.rmfFunc = "ready") {
	If !threadRWF.ahkReady() {
		;~ ahkthread_free(threadRWF)
		;~ threadRWF:=""
		ToolTip,,,, 3
		func_call := settings.funcs.showThreadVar
		SetTimer, % func_call, Off

	}


}


; Daten
; -----------------------------------------------------------------------------------------
ReadPatDB()                                                                            	{
	infilter 	:= ["NR", "VERSART", "PRIVAT", "NAME", "VORNAME", "GEBURT", "NAMENORM", "TELEFON", "TELEFON2", "FAX", "VERSART"
				,"HAUSARZT", "ARBEIT", "DEL_DATE", "RKENNZ", "STRASSE", "PLZ", "ORT", "HAUSNUMMER", "GESCHL", "SEIT", "LAST_BEH", "MORTAL"]
	Patients		:= new PatDBF(Addendum.AlbisDBPath, infilter, "alldata=true  moredata=true")
return Patients
}

GetPrinters()                                                                         	{	;--	gibt die Drucker so zurück das man diese in eine List- oder Combox laden kann

	for Item in ComObjGet( "winmgmts:" ).ExecQuery("Select * from Win32_Printer")
			DDLPrinter .= !InStr(DDLPrinter, Item.Name) ? Item.Name "|" : ""

return  RTrim(DDLPrinter, "|")
}

AddMorePatientsToDB(Letzte_Behandlung:="")                                             	{

	If !Letzte_Behandlung
		Letzte_Behandlung := A_YYYY-3 . SubStr("00" A_MM, -1) . SubStr("00" A_DD, -1)

	SciTEOutput("Patientenakten mit der letzten Behandlung ab dem  " ConvertDBASEDate(Letzte_Behandlung) ".")

	If !QEDB.GetTable(_SQL :=  "SELECT * FROM EXPORTS;", result) {
		SQLiteErrorWriter(table " Tabelle - GetTable Fehler", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)
		return
	}
	ids := Object()
	For index, row in result.Rows
		ids[Row.1] := index

	Existing := 0
	_SQL := "INSERT INTO EXPORTS (PatID,SName,PName,Birth,SDate,Recipient,Medium,Route,Shipment,More1,ExportPath,More3)`r`nVALUES "
	Patients := cPat.GetPatDB()
	For PatID, m in Patients
		If (m.LAST_BEH >= Letzte_Behandlung && !(m.Name ~= "(Clemenz|AOK|BARMER|Mustermann)") && !m.Mortal) {

			If !ids[PatID]
				_SQL .= "(" q PatID q "," q m.NAME q "," q m.VORNAME q "," q m.GEBURT q "," q "20100101" q "," q "1" q "," q "🖄 ePost" q "," q "📡 Telematik" q "," q q "," q q "," q q "," q q "),`r`n"
			else {
				Existing ++
				;~ SciTEOutput("[" SubStr("0000" existing, -3) "] (" PatID ") " m.Name ", " m.Vorname "`texistiert bereits")
			}

		}

	_SQL := RTrim(_SQL, ",`r`n") ";"

	SciTEOutput("Patientenakten für Export: " StrSplit(_SQL, "`n").Count() ", vorhanden waren: " existing)
	FileOpen(stdExportPath "\lBeh.sql", "w", "UTF-8").Write(_SQL)

	QEDB.Exec("BEGIN TRANSACTION;")

	If !QEDB.Exec(_SQL)
		SQLiteErrorWriter(table " Tabelle - Fehler bei INSERT INTO", QEDB.ErrorMsg, QEDB.Errorcode, true, _SQL)

	QEDB.Exec("COMMIT TRANSACTION;")


}

ShowObjects()                                                                         	{

	static hwnd

	obj := settings
	hwnd := ObjTree(obj,"settings Object()", "+ReadOnly +Resize,GuiShow=w800 h600")

return
}


; Wrapper
; ---------------------------------------------------------------------------------------------
SendMessage(msg, wParam:="", lParam:="", ControlName:="", wTitle:="", wText:="")      	{

	static BM_GetCHECK := 0x00F0

	SendMessage % msg, % wParam, % lParam, % ControlName, % wTitle, % wwText
return ErrorLevel
}

initRegistryRead()                                                                    	{
	SetRegView, % (A_PtrSize=8 ? 64 : 32)
}

RegistryRead(KeyName, ValueName:="")                                                  	{

	static _init := initRegistryRead()
	RegRead, RegValue, % KeyName, % ValueName

return RegValue
}

RegistryWrite(ValueType, KeyName , ValueName:="", Value:="")                          	{

	RegWrite, % ValueType, % KeyName, % ValueName, % Value

return ErrorLevel ? 0 : 1
}


; Steuerelemente
; ---------------------------------------------------------------------------------------------
GuiControlToolTip(vCtrlName, msg, delay:= 5, ttPosition := "above")                    	{ ;-- blendet ein Tooltip ober- oder unterhalb eines Gui Steuerelementes ein

	hwnd := GuiControlGet("QE", "HWND", "QEPatient")
	cp   := GetWindowSpot(hwnd)
	cpY  := cp.Y + (ttPosition="above" ? -30 - (StrSplit(msg, "`n").Count()*20) : cp.H + 3)
	ToolTip, % msg, % cp.X, % (cpY>=0 ? cpY : 0), 16
	ToolTipOff(16, delay*1000)
	SciTEOutput(A_ThisFunc "() | " hwnd ", " cpY)

}

ToolTipOff(TTNr, delay:="")                                                           	{ ;-- startet Timer zum Entfernen eines Tooltips

	func_call := Func("TTToff").Bind(TTNr)
	SetTimer, % func_call, % (delay ~= "^\d+$" ? -1*delay : -1)

return
}
TTToff(TTNr) {
	ToolTip,,,, % TTNr
}

ToogleMenu(var)                                                                       	{

	; Variables are set to super global!
	global

	If (var = "dbg") {
		dbg := !dbg
		Menu, TMenu, % (dbg ? "Check" : "Uncheck")           	, % "Debug"
		GuiControl, QE:, QEDBG, % (dbg ? "an":"aus")
		TrayTip, % StrReplace(A_ScriptName, ".ahk"), % "Script debugging wurde " (dgb ? "eingeschaltet":"ausgeschaltet") "."
		IniWrite, % dbg           	, % settings.paths.QEExporterINI, % "Script", % "Debug"
	}
	else if (var = "noSQLWrite") {
		noSQLWrite := !noSQLWrite
		Menu, TMenu, % (noSQLWrite ? "Check" : "Uncheck")	, % "SQL write protection"
		GuiControl, QE:, QESQLWP, % (noSQLWrite ? "an":"aus")
		TrayTip, % StrReplace(A_ScriptName, ".ahk"), % "Das Schreiben in die SQLite Datenbank ist " (noSQLWrite ? "ausgeschaltet":"eingeschaltet") "."
		IniWrite, % noSQLWrite	, % settings.paths.QEExporterINI,  % "SQLite", % "SQL_write_protection"
	}


}

WinToClient(hWnd, ByRef x, ByRef y)                                                    	{

	WinGetPos wX, wY,,, % "ahk_id " hWnd
	x += wX, y += wY
	VarSetCapacity(pt, 8), NumPut(y, NumPut(x, pt, "int"), "int")
	if !DllCall("ScreenToClient", "ptr", hWnd, "ptr", &pt)
		return false
	x := NumGet(pt, 0, "int"), y := NumGet(pt, 4, "int")

return true
}

ControlHasStyle(hwin, WS_StYLE, ControlName :="")                                     	{

	static WS := {"VISIBLE"	:0x10000000
						  ,	"DISABLED":0x08000000
						  ,	"CHILD"  	:0x40000000
						  ,	"POPUP" 	:0x80000000}

	COntrolGet, Style, Style,, % ControlName, % "ahk_id " hwin

return (Style & WS[WS_StYLE])
}

SB_SetProgress(Value=0,Seg=1,Ops="", hwndBar=0)                                        	{
	; SB_SetProgress
	; (w) by DerRaphael / Released under the Terms of EUPL 1.0
	; see http://ec.europa.eu/idabc/en/document/7330 for details

   ; Definition of Constants
   Static SB_GETRECT      := 0x40a      ; (WM_USER:=0x400) + 10
        , SB_GETPARTS     := 0x406
        , SB_PROGRESS                   ; Container for all used hwndBar:Seg:hProgress
        , PBM_SETPOS      := 0x402      ; (WM_USER:=0x400) + 2
        , PBM_SETRANGE32  := 0x406
        , PBM_SETBARCOLOR := 0x409
        , PBM_SETBKCOLOR  := 0x2001
        , dwStyle         := 0x50000001 ; forced dwStyle WS_CHILD|WS_VISIBLE|PBS_SMOOTH

   ; Find the hWnd of the currentGui's StatusbarControl
   Gui,+LastFound
   If !hwndBar
	ControlGet,hwndBar,hWnd,,msctls_statusbar321

   if (!StrLen(hwndBar)) {
      rErrorLevel := "FAIL: No StatusBar Control"     ; Drop ErrorLevel on Error
   } else If (Seg<=0) {
      rErrorLevel := "FAIL: Wrong Segment Parameter"  ; Drop ErrorLevel on Error
   } else if (Seg>0) {
      ; Segment count
      SendMessage, SB_GETPARTS, 0, 0,, ahk_id %hwndBar%
      SB_Parts :=  ErrorLevel - 1
      If ((SB_Parts!=0) && (SB_Parts<Seg)) {
         rErrorLevel := "FAIL: Wrong Segment Count"  ; Drop ErrorLevel on Error
      } else {
         ; Get Segment Dimensions in any case, so that the progress control
         ; can be readjusted in position if neccessary
         if (SB_Parts) {
            VarSetCapacity(RECT,16,0)     ; RECT = 4*4 Bytes / 4 Byte <=> Int
            ; Segment Size :: 0-base Index => 1. Element -> #0
            SendMessage,SB_GETRECT,Seg-1,&RECT,,ahk_id %hwndBar%
            If ErrorLevel
               Loop,4
                  n%A_index% := NumGet(RECT,(a_index-1)*4,"Int")
            else
               rErrorLevel := "FAIL: Segmentdimensions" ; Drop ErrorLevel on Error
         } else { ; We dont have any parts, so use the entire statusbar for our progress
            n1 := n2 := 0
            ControlGetPos,,,n3,n4,,ahk_id %hwndBar%
         } ; if SB_Parts

         If (InStr(SB_Progress,":" Seg ":")) {

            hWndProg := (RegExMatch(SB_Progress, hwndBar "\:" seg "\:(?P<hWnd>([^,]+|.+))",p)) ? phWnd :

         } else {

            If (RegExMatch(Ops,"i)-smooth"))
               dwStyle ^= 0x1

            hWndProg := DllCall("CreateWindowEx","uint",0,"str","msctls_progress32"
               ,"uint",0,"uint", dwStyle
               ,"int",0,"int",0,"int",0,"int",0 ; segment-progress :: X/Y/W/H
               ,"uint",DllCall("GetAncestor","uInt",hwndBar,"uInt",1) ; gui hwnd
               ,"uint",0,"uint",0,"uint",0)

            SB_Progress .= (StrLen(SB_Progress) ? "," : "") hwndBar ":" Seg ":" hWndProg

         } ; If InStr Prog <-> Seg

         ; HTML Colors
         Black:=0x000000,Green:=0x008000,Silver:=0xC0C0C0,Lime:=0x00FF00,Gray:=0x808080
         Olive:=0x808000,White:=0xFFFFFF,Yellow:=0xFFFF00,Maroon:=0x800000,Navy:=0x000080
         Red:=0xFF0000,Blue:=0x0000FF,Fuchsia:=0xFF00FF,Aqua:=0x00FFFF,Orange:=0xFFBB00

         If (RegExMatch(ops,"i)\bBackground(?P<C>[a-z0-9]+)\b",bg)) {
              if ((strlen(bgC)=6)&&(RegExMatch(bgC,"i)([0-9a-f]{6})")))
                  bgC := "0x" bgC
              else if !(RegExMatch(bgC,"i)^0x([0-9a-f]{1,6})"))
                  bgC := %bgC%
              if (bgC+0!="")
                  SendMessage, PBM_SETBKCOLOR, 0
                      , ((bgC&255)<<16)+(((bgC>>8)&255)<<8)+(bgC>>16) ; BGR
                      ,, ahk_id %hwndProg%
         } ; If RegEx BGC
         If (RegExMatch(ops,"i)\bc(?P<C>[a-z0-9]+)\b",fg)) {
              if ((strlen(fgC)=6)&&(RegExMatch(fgC,"i)([0-9a-f]{6})")))
                  fgC := "0x" fgC
              else if !(RegExMatch(fgC,"i)^0x([0-9a-f]{1,6})"))
                  fgC := %fgC%
              if (fgC+0!="")
                  SendMessage, PBM_SETBARCOLOR, 0
                      , ((fgC&255)<<16)+(((fgC>>8)&255)<<8)+(fgC>>16) ; BGR
                      ,, ahk_id %hwndProg%
         } ; If RegEx FGC

         If ((RegExMatch(ops,"i)(?P<In>[^ ])?range((?P<Lo>\-?\d+)\-(?P<Hi>\-?\d+))?",r))
              && (rIn!="-") && (rHi>rLo)) {    ; Set new LowRange and HighRange
              SendMessage,0x406,rLo,rHi,,ahk_id %hWndProg%
         } else if ((rIn="-") || (rLo>rHi)) {  ; restore defaults on remove or invalid values
              SendMessage,0x406,0,100,,ahk_id %hWndProg%
         } ; If RegEx Range

         If (RegExMatch(ops,"i)\bEnable\b"))
            Control, Enable,,, ahk_id %hWndProg%
         If (RegExMatch(ops,"i)\bDisable\b"))
            Control, Disable,,, ahk_id %hWndProg%
         If (RegExMatch(ops,"i)\bHide\b"))
            Control, Hide,,, ahk_id %hWndProg%
         If (RegExMatch(ops,"i)\bShow\b"))
            Control, Show,,, ahk_id %hWndProg%

         ControlGetPos,xb,yb,,,,ahk_id %hwndBar%
         ControlMove,,xb+n1,yb+n2,n3-n1,n4-n2,ahk_id %hwndProg%
         SendMessage,PBM_SETPOS,value,0,,ahk_id %hWndProg%

      } ; if Seg greater than count
   } ; if Seg greater zero

   If (regExMatch(rErrorLevel,"^FAIL")) {
      ErrorLevel := rErrorLevel
      Return -1
   } else
      Return hWndProg

}

LB_SetSelAll(hwin, LbName:="ListBox1")                                                 	{

	hwin := Trim(hwin)
	If (hwin ~= "i)^\d+$")
		hwin := "ahk_id " hwin
	else If (hwin ~= "i)^(0x)*[\dA-F]+$")
		hwin := "ahk_id 0x" RegExReplace(hwin, "^0x")

	If IsObject(notselected := LB_notSelected(hwin, LbName))
		return notselected

	If (notSelected > 0) {
		Loop 2 {
			res := SendMessage(0x0185, 1, -1, LbName, hwin)  		; alle Zeilen auswählen  LB_SetSet wParam=1 ausgewählt, lPAram=-1 alle
			If ((notSelected := LB_notSelected(hwin, LbName)) = 0)
				return true
			Sleep 300
		}
	}

return notselected=0 ? true : {"error": "select failure", "notSelected": notselected}
}

LB_notSelected(hwin, LbName:="ListBox1")                                              	{

		hwin := Trim(hwin)
	If (hwin ~= "i)^\d+$")
		hwin := "ahk_id " hwin
	else If (hwin ~= "i)^(0x)*[\dA-F]+$")
		hwin := "ahk_id 0x" RegExReplace(hwin, "i)^0x")

	If !(LBCount := SendMessage(0x018B,,, LbName, hwin))
		return {"error": "no items"}

	LBSelCount := SendMessage(0x0190,,, LbName, hwin)  ; Wieviele Zeilen sind selektiert?
	notselected := LBCount-LBSelCount

return notselected
}

EM_SetMargins(Hwnd, Left := "", Right := "")                                           	{
   ; EM_SETMARGINS = 0x00D3 -> http://msdn.microsoft.com/en-us/library/bb761649(v=vs.85).aspx
   Set := 0 + (Left <> "") + ((Right <> "") * 2)
   Margins := (Left <> "" ? Left & 0xFFFF : 0) + (Right <> "" ? (Right & 0xFFFF) << 16 : 0)
   Return DllCall("User32.dll\SendMessage", "Ptr", HWND, "UInt", 0x00D3, "Ptr", Set, "Ptr", Margins, "Ptr")
}

LV_GetSortedColumn(hLV)                                                               	{

	static 	ptr 						:= A_PtrSize   		? ("ptr"		, ptrSize := A_PtrSize)
				         				     								: ("uint"		, ptrSize := 4)
	static 	LVM_GETCOLUMN 	:= A_IsUnicode  	? (4191	, LVM_SETCOLUMN := 4192)
						               									: (4121 	, LVM_SETCOLUMN := 4122)

	VarSetCapacity(lvColumn, ptrSize + 4)
	NumPut(1, lvColumn, "uint")

	VarSetCapacity(LVCOL, 56, 0)
	NumPut(1, LVCOL, "UInt") ; LVCF_FMT

	hdrH := DllCall("SendMessage", "uint", hLV, "uint", 4127) 						; LVM_GETHEADER
	hdrC := DllCall("SendMessage", "uint", hdrH, "uint", 4608) 			     		; HDM_GETITEMCOUNT

	t .= "hLV: " hLV "`thdrH: " GetHex(hdrH) "`thdrC: " hdrC "`n"
	t .= "Col`tfmt  `tfmtHex  `tfmt&1024`n"
	Loop % hdrC {

		VarSetCapacity(lvColumn, ptrSize + 4)
		NumPut(A_Index-1, lvColumn, "uint")
		DllCall("SendMessage", ptr, hdrH, "uint", LVM_GETCOLUMN, "uint", A_Index-1, ptr, &lvColumn)
		If (fmt := NumGet(lvColumn, 4, "UInt"))
			fmtHex := GetHex(fmt)
		else {
			nix := 0
		}

		t .= A_Index "`t" fmt "`t" GetHex(fmt) "  `t" fmt & 1024 "`n"

	}

	dbg ? SciTEOutput(A_ThisFunc "() |`n" t) : ""

}

LV_GetRowIndex(MapID, hLV) 																											     		{
return DllCall("SendMessage", "ptr", hLV, "uint", 0x10BB, "uint", MapID)
}

LV_GetRowMapID(rowIndex, hLV) 																									     		{
return DllCall("SendMessage", "ptr", hLV, "uint", 0x10B4, "uint", rowIndex-1)
}

class PerformanceCounter                                                               	{	;This class implements an object oriented interface to the system's performance counter.

	;Get the frequency of the system's performance counter and store it in a class variable to
	;eliminate .dll call overhead.
	static __frequency := PerformanceCounter.__initialiaze_frequency()
	__initialiaze_frequency(){
		DllCall("QueryPerformanceFrequency", "Int64*", frequency)
		return frequency
	}

	query(){
		;Returns the current count of the performance counter.

		DllCall("QueryPerformanceCounter", "Int64*", count)
		return count
	}

	frequency[]{
		get{
			;The frequency, in hertz, of the performance counter.

			return this.__frequency
		}
	}
}

class PerformanceTimer                                                                	{	;This class implements a timer based on the system's performance counter.

	; Depending on the
	;system it has a precision on the order of microseconds.

	__started := false
	__paused := false
	__elapsed_counts := 0

	start(){		;Starts the timer.  Also unpauses the timer if it has been paused.

		;Make sure that the timer has not already been started, or is paused.
		if (this.__started and not this.__paused)
			return
		this.__started := true
		this.__paused := false
		this.__start_count := PerformanceCounter.query()
	}

	pause(){		;Pauses the timer without resetting it.  Call start() to restart the timer.

		elapsed_counts := PerformanceCounter.query() - this.__start_count

		;Make sure that the timer is not already paused.
		if (this.__paused){
			return
		}
		this.__paused := true

		this.__elapsed_counts += elapsed_counts
	}

	reset(){		;Stops the timer and resets the elapsed time to 0.
		this.__started := false
		this.__paused := false
		this.__elapsed_counts := 0
	}

	elapsed[]{
		get{
			elapsed_counts := PerformanceCounter.query() - this.__start_count

			;If the timer has started, and is not paused, add the number of counts elapsed since
			;start() was last called.
			if (this.__started and not this.__paused)
				this.__elapsed_counts += elapsed_counts

			;Convert the elapsed counts to microseconds.  Parenthesis enforce the order of
			;operations that will result in the least loss of precision.  See:
			;https://msdn.microsoft.com/en-us/library/windows/desktop/dn553408%28v=vs.85%29.aspx#examples_for_acquiring_time_stamps
			return ((this.__elapsed_counts * 1000000) // PerformanceCounter.frequency)
		}
	}
}

;}

; Icon 																															;{
; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_QuickExporter_ico(NewHandle := False) {
Static hBitmap := Create_QuickExporter_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGcAAABnAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFAxQVDV4jFpwrG7wvHc8wHtMvHdMvHdMvHdMuHNMuHNMuHNMuHNMuHNItG9ItG9ItG9ItG9ItG9IsGtIsGtIsGtIsGtIsGtIrGdIrGdIrGdIrGdIqGNIqGNIqGNIpF9IpF9IoF84kFLsgEpoTC10EAhQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAQQaEWsyIcsyIdMzIdQzIdMzIdMzIdMzIdMyINMyINMyINMxINMxH9MxH9MxH9MxH9MwH9MwHtMwHtMwHtMwHtMvHtMvHdMzIdMyIdMyINMyINMyINMuHNMxH9MxH9MwH9MwHtMvHtMvHdMuHNMvHdMuHNMrGskVDGkBAAMAAAAAAAAAAAAAAAAAAAADAg4nGp41I9Q2JdQ4JtQ4JtQ4J9Q5KNQ6KNQ5KNQ5J9Q5J9Q5J9Q4J9Q4J9Q4JtQ4JtQ3JtQ3JtQ3JdQ3JdQ3JdQ2JdQ2JdQ2JNQ2JNQ2JNQ1JNQ1JNQ4J9Q1I9Q4JtQ3JtQzItQ2JNQ1JNQ0ItQ0ItMzIdMwH9MuHNMtG9IgE5wDAg4AAAAAAAAAAAABAQQnGp03JdQ4JtQ5J9U7KdU9LNU9LNU/LdVALtVAL9VAL9VALtU/LtU/LtU/LdU+LdU+LdU+LdU+LdU+LNU+LNU9LNU9LNU9LNU9K9U9K9VALtU/LtU8KtU/LtU/LtU+LdU+LNU9LNU7KtU6KdU6KdQ4J9Q1JNQzItMyINMvHdMtG9IhE6ABAQUAAAAAAAAaEWs1JNQ4JtQ6KNU8KtU+LNVALtZBMNZDMtZFNNZGNdZDMtZDMtZDMdZCMdZCMdZCMdZCMNZCMNZBMNZBMNZBMNZBMNZAL9ZAL9ZEMtZDMtZDMtZDMtZCMdZCMdZFNNZFNNZEM9ZDMdZBMNZBMNU/LtU7KtU4J9Q2JNQzItMxH9MvHdMsGtIVDGkAAAAFAxQyIcs3JdQ5KNU8KtU+LNVAL9ZDMtZFNNZHNtdJONdLOtdKOddKOddJONdJONdJONdJONdIN9dIN9dIN9dIN9dIN9dHNtdHNtdKOtdHNtdHNtZKOddGNdZJONdJONdJONdLO9dJOddHNtZHNtdFNNZCMdY+LdU7KtU4J9Q0I9MzIdMwH9MtG9IoF8gEAhQVDV41I9Q3JtQ7KdU9LNVBL9ZDMtZGNdZJONdLO9dNPNhQQNhQP9hRQNhQP9hQP9hQP9hPP9hPPthPPthPPthOPthOPthOPdhOPdhOPdhNPdhRQNhNPNdNPNdQP9hPP9hSQthRQNhPPthNPddMO9dIN9dEM9ZAL9U+LNU7KtQ4JtQ1I9QxH9MuHNIrGdISClwkFp01I9Q4JtQ7KtU/LdVCMdZFNNZIN9dLOtdPPthSQdhUQ9lXRtlXRtlXR9lXRtlWRtlWRtlWRdlWRdlVRdlVRdlVRdlVRNlVRNlURNlYR9lURNhUQ9hTQ9hWRtlZSdlYSNlWRdlVRNlSQthPP9hLO9dHNtZEM9ZAL9U9LNU5KNQ2JNQyINMvHdMrGdIdEJsrHLw0ItQ5KNU9K9VAL9ZEM9ZHNtdKOddOPdhRQdhWRdlYSNlaStldTdpeTtpeTdpdTdpdTdpdTNpcTNpcTNpcTNpcS9pbS9pbS9pbS9pbS9lbStlaStldTdpgUNpgUNpdTdpcTNpZSdlWRtlTQthOPddKOddGNdZCMdY/LtU7KdQ4JtQ0ItMwH9MsGtIjE7kwHs81I9Q5KNU9LNVCMNZGNdZJONdNPNdQP9hURNlYSNlcTNpfT9piUtplVdtkVNtkVNtkVNtjU9tjU9tjU9tjU9tiUttiUttiUttlVdtlVdtlVdtoWNtnWNtnV9tiUttkVNthUdpeTtpZSdlVRdlRQNhNPNdJONdEM9ZBMNU9LNU6KNQ1JNQxINMtG9InFs4xH9M1JNQ7KdU+LdVDMtZHNtdLOtdPPthSQdhWRtlbStpeTtpiUttmVttpWdttXdxrW9xqW9xqWtxqWtxqWtxpWtxpWdxpWdxpWdxpWdxvX9xrXNxrW9xuXtxtXtxsXNxpWtxmVttiUtpcTNpYR9lURNhQP9hMO9dHNtZDMtY/LtU7KtQ2JdQyINMtG9IqGNIvHtM1I9Q6KdVALtVEM9ZJONdNPNdRQNhXR9mJfOS1re7V0PXW0vXY0/bY1PbZ1fba1vbZ1fbZ1fbZ1fbZ1fbZ1fbZ1fbZ1fba1fba1vbZ1fba1vba1vba1vba1vbZ1fbY1PbX0/bW0vXU0PXTz/XEvvKlnetvYd5JONdFNNZCMdU9LNU4J9QzIdMtG9IoFtEvHdM0ItM6KNRBMNZFNNZLOtdPPtiQhubs6vv////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////RzPVmWN1DMtY/LtU7KtQyINMtG9InFdEuHNIzIdM5J9RAL9VGNdZQP9i+t/D////////v7/j1+vD4/+z4/+v3/ur1/Onz+ebx+OXv9ePt9OHs8uDr8d/p8N3p793p7t3p7t3p7t3p793p8N3r8d/s8uDt9OHv9ePx+OXz+eb1/On3/ur4/+v4/+31+PL29vv////29f16beJAL9U5J9QyINMtG9InFdEtG9IzINM4JtQ/LtVFNNa3se/////49/31+u72/+f2/+b2/+b1/uXy++Lv+ODs9d3q8trn79jl7dbj6tTh6NLg59Hf5tDe5c/e5c/e5c/f5tDg59Hh6NLj6tTl7dbn79jq8trs9d3v+ODy++L1/uX2/+b2/+b2/+j2+fX////29f1fUdw4J9QxINMsGtInFdEtG9IyINM3JtQ+LdV6beL////6+v73/uz2/+f2/+b2/+b2/+bz++Pv+ODs9d3p8drm7tfk69Xh6dPf5tDd5M7b4s3a4czZ4MvZ4MvZ4Mva4czb4s3d5M7f5tDg59Hk69Xm7tfp8drs9N3v+ODy++P2/ub2/+b2/+b2/+f4+vb////OyvQ4J9QxINMsGtInFdEsGtIxH9M3JdQ+LNXU0Pb////09/H2/+f2/+f2/+b2/+b0/eTw+eHt9d7q8tvn7tjk69Xh6NLe5c/S28bAz7y3yLa4yrfI1L/U2sbU28bV3MfX3cnZ38vb4s2TvqzP4Mrj69Tm7tfq8trt9d3w+eDz/OT2/+b2/+b2/+b3/+n7+/3///9ZS9sxH9MsGtInFdEsGdIxH9M2JNRQQNn+/v/8/P72/+j2/+f2/+b2/+b2/+by++Lv99/r89zo8Nnj6tTL1sGouaiEnpBrj4JciXxOg3hAgHUzfnM6h3xkopSmw7HT2cXV28fY3sm90r5Hmo7S4czP483N5M7r89vu99/y+uL1/uX2/+b2/+b2/+f7/vb///+Ng+YxH9MsGtImFNErGdIwHtM2JNRuYeD////y8vn2/+f2/+f2/+b2/+b1/eXx+uHt9t7o8NnQ2sWuu6qYq5uRp5eVq5uasaGUrp6Ip5lsmYs9g3gkfXIYfXMLgHUyloqnyLbV28fY38qxzrs8mo7W5M5asaLk79nt9d3w+eH0/eT2/+b2/+b2/+f7//P///+imusxH9MrGtImFNErGNIwHtM1I9R4bOL////y8/j2/+f2/+f2/+b2/+b0/eTw+eHm7tfN1sHAy7fH0r7T28fb4czX3snU2sbQ1sLN0r/Kz7zEyriWtKQ9joILf3UAgngCi4Bgs6TO2sba4cykyrdFo5V8wbBvv67s9Nzv+ODz/OP2/+b2/+b2/+f7//X///+ZkOkwH9MrGdImFNEqGNIvHdM0ItNmWd3////29vv2/+f2/+f2/+b2/+b0/eTw+eDh6tPi6dPl7dbi6dPe5c/a4czX3cnT2cXP1cLM0b7IzbvGy7nKzrzBybiBr6AKhnsAioAAkocepJew1cHd5M5/va1htaQyp5mr28bW79jt+eL2/+b2/+b2/+b9/vr///91aeEwH9MrGdImFNEqF9IvHdM0ItNAL9X49/3////3/+v2/+f2/+b2/+b0/eTw+eHt9d3p8drm7dfi6dPe5dDb4szX3snU2sbQ1sLN0r/Jz7zHzLrFyrh8rJ2ivaysxLMll4wAkocAm48KppmY0sDh6NJQrZ57xLIjppni89yp4s32/+b2/+b5/+3////y8fxALtYwHtMrGdIlE9EpF9IuHNIzIdM6KdS0re/////6/Pf2/+f2/+b2/+b1/uXx+uLu9t7q8tvn7tjj69Tg59Hc487Z38vV3MfS2MTP1cLN0r/L0L2qxLPDzrpDmYy/zrvG078vopUAm48Ao5YCrJ+B0sDa6dIqo5ZtxLJpx7SJ1sTi+OD4/+r9/v3///+Vi+g3JdQwHtMrGdIlE9EoFtIuG9IzIdM6KNRURdr19P3////7//P2/+b2/+b2/+by++Pv99/s9Nzo8Nnl7Nbh6dPe5dDb4s3Y38rV3MfT2cXR18PP1cLA0Lw6oZOQvaxEnpHT28XL2cUwqp0Ao5YAq54Bs6WM2seo2MNevq0dq52r4s9x0sX+/v7////Ev/I/LtU2JdQwHtMqGdIlE9EoFtEtG9IyINM5J9Q+LdV8cOL8/P7////8/vf3/+r2/+b0/eTx+eHt9t7q8tvn79jk69Xh6NLe5dDb4s3Z4MvX3cnV3MfU2sbU2saaxrU2pZdvtKVmsKHa4czK3cgZqp0Aq54As6UCvK6f5NCb2Mea2tE8vLJrz8jj9va1ru9GNdc9LNU2JdQvHtMqGNIlE9EoFdEtG9IyINM5J9Q+LNVFNNZ8cOLw7/z////////5//D2/+bz/OPw+ODt9d3q8trn79jk7NXh6dPf5tHd5M/b4s3a4czZ4MvY38rZ38pct6druqpEqJp+vq7g59Kx28YErZ8As6UAu60MxrbR9en8/v4esaZ+utQ1lMBJONdEM9Y9LNU2JdQvHdMqGNIlE9EnFdEsGtIxH9M4JtQ9LNVEM9ZLOtdeTtvEvvH////5/Pb2/+b2/+bz++Pw+ODt9d3q8tvo79jl7dbj69Th6dLg59He5tDe5c/d5M/G3snQ4cswrJ6Oy7kro5aFx7Tm7tdxzbsAs6UAu60Aw7Qp1MX4/v1SocMkhbcVlbRGPtVDM9Y9K9U2JNQvHdMqGNIlE9EnFNEsGtIxH9M4JtQ9K9VEMtZKOddRQdh1aOD////7+/33/+j2/+b2/+bz/OPw+eHu9t7r89zp8drn79jm7dfk7NXj69Tj6tTi6tPi6tOO0b603Mgfqp2Z08Axq5xyxLHk8dsiva0Au60Aw7QAzLx+6eF2jdcueb4ne78zX8pDMtY8K9U2JNQvHdMqGNIkEtEmFNErGdIwHtM3JdQ8K9VDMtZKOddRQNhXR9nTz/X////5+/j3/+f2/+b2/+b0/eTx+uP0+unv9+Hs9Nzq8tvp8dro8Nnn79jn79jn79jl7tdiybeg2scVrJ624MzI59Lo9t2Y4MsAu60Aw7QAzLwG1cRdb9tWSNg8W8xBRNJDMtY8K9U1JNQuHdIpGNIkEtEmE9ErGdIwHtM3JdQ8KtRDMdZJONdQP9hXRtlwYd7t7Pv////7+/37//T6//D7//P9//z////7/Prw+eDv99/u9t7t9d3s9N3s9N3s9N3s9d3n891nzryS2scTsKKg3sr0/eTy/uUgxLUAw7QAzLwA1MMlpdBXR9lNRdg7WdhDMtY8K9U1JNQuHdIpF9IkEtElE9EqGNIvHdM2JNQ7KdRCMNVJONZPPtdWRtldTNp0Zt/SzfT////////////////////8/P7////6/vHz/OPy++Py+uLx+uHx+eHx+eHx+uHy+uLu+uJi0b+H3Mc1v6+x6NL2/+aT49IAw7QAzLwA1MMD18wip9QF4doVvd5CMtY8K9U1I9MuHNIpF9IkEtElEtEqF9IvHdM2JNQ7KdRBMNVIN9ZPPtdVRNhcS9lkVNtqWtuEd+Gupeu8tO6xqOuekuatour+/v/+/v/5/+/2/+b2/+b2/+b2/+b2/+b2/+b2/+b2/+b1/+aD3spx2MX0/uX3/+vv+vkKxbcAzLwA1MMA3MsA5NIA7NoRx95CMdY7KtU1I9MuHNIpF9IjEdEkEdEpF9IuHNI1I9M6KNRBL9VHNtZOPddUQ9hXR9lgT9pnV9trW9xwYNx1Zt16a96Bct+FduC/t+7////////6/vP2/+f2/+b2/+b2/+b2/+b2/+b2/+b2/+b2/+f2/+fm+eqU3tk4yr8Aw7QAzLwA1MMA3MsA5NIA7NoN0t9CMdY7KtQ0I9MtHNIoF9IjEdEkEdEpFtIuG9I0ItM5J9RALtVGNdZLOtdSQdhXRtheTdpjU9pnV9trW9xwYNx1Zt17a95/cN+FduC2ruz7+v7////7+/z6/vP3/+v2/+f2/+f2/+f2/+f3/+n5//D6+/p81M0As6UAu60Aw7QAzLwA1MMA3MsA5NIA7NoJ3OBCMdU7KtQ0I9MtHNIoFtIjEdEjENEoFdEtG9I0IdM4J9Q/LdVAL9VKOdZPPtdUQ9hZSdldTdpjUtpmVttrW9xvX9x1Zd15at5+bt6Bct+XiuTd2Pb////////////8/P35+fv4+fv7+/3////////////o5flUosgAu60Aw7QAzLwA1MMA3MsA5NIA7NoF5+BBMNU7KdQ0ItMtG9IoFtIjEdEiENEnFdEsGtIzINM3JdQ9K9VDMdVJN9ZMO9dQP9hUQ9hZSdleTdpiUtpmVttqWttwYNx0ZN13aN57bN58bd58bd6ViOTFvvDm4/n9/P7////////9/P7u7PvJw/KdkedxYtxrW9w4hsYAw7QAzLwA1MMA3MsA5NIA7NoB8eBAMNU6KdQzItMtG9IoFtEjEdEiD9AmFNErGdIuHNI3JdQ9K9VAL9VFNNZJONZNPNdQP9dUQ9hYR9lcTNlgUNplVdtpWdtuXtxwYdxxYd11Zd11Zt15at55at56at56at6DdeCFd+F7bN53aN1zY91vX9xqW9xmVttjU9o4fsoBy70A1MMA3MsA5NIA7NoA9OE7NtY4J9QzItMsG9InFdEiENEhDtAlE9ErGdIwHtM2JNM6KdQ/LdVCMdVFNNZINtZLOtdPPtdTQthYR9lcTNlfT9piUtpnV9tqWttqWttuXtxrW9xrW9xvX9xvX9xvX9xvYNxwYNxwYNxxYd1uXtxqWttlVdthUdpfTtpZSdk0es0B08MA3MsA5NIA7NoA9OE3Qdc2JdQyINMtHNInFdEiENEbCcwlEtErGNIwHdI1I9M3JdQ6KdQ9K9RALtVDMtZHNtZLOtdPPtdSQdhXRtlaSdlcS9lfTtpfT9pgT9pgT9pgUNpnV9toV9toWNtoWNtoWNtpWNtpWdtpWdtnV9tkU9tgUNpcTNlaSdlWRdhSQdgzcdAB2ssA5NIA7NoA9OExSdg2JdQxINMsG9ImFNEgD80dDLskEdEqF9EtG9IxH9M0ItM2JNQ5J9Q8KtQ/LtVEMtZHNtZKOddNPNdRQNhTQthWRthYR9lYSNlcTNlZSNldTNldTNldTNldTdpdTdpeTdpeTdpeTtpeTtpfTtpdTdpaSdlXRtlUQ9hRQNhNPNdJONcwZtIC39IA7NoA9OEsU9g0ItMwHtMrGdImFNEcDLgXCZwkEdEoFtEsGdItG9IwHdIyINM1I9M4JtQ8KtQ/LtVCMNVGNdZHNtZLOtdOPddRQNhRQNhVRNhVRNhVRNhVRdhWRdhWRdhWRdhWRtlXRtlXRtlXRtlXRtlYR9lYR9lVRNhSQdhPPtdMO9dJONdGNdZDMdYtY9QC5toA9OEnXdkxH9MuHNIoFtIlE9EYCpoOBV8iENAlE9EoFdEqF9ItGtIvHdIyINM1ItM3JdQ7KdQ/LdVALtVDMtZGNNZJN9ZKONdKOddKOddOPddOPddPPddPPtdPPtdPPtdPPtdPPtdQP9dQP9hQP9hRP9hRQNhPPtdNO9dKOddHNtZEM9ZCMNU+LdU7KdQpV9YC6+AiaNovHdMpF9IoFtEjEdEOBl0DARQgDcgjENElEtEoFdEpF9IsGdIvHdIxH9M0ItM2JNQ5J9Q8KtQ+LdU+LNU/LdVDMdU/LtVDMtZEMtZEMtZHNtZINtZINtZIN9ZIN9ZIN9ZJN9ZJONZJONZJONdKONdKOddHNtZFNNZDMdZALtU+LNU6KdQ3JtQ1I9QmUdYgZtosGtIpFtIlEtEfDccDARQAAAAQB2shDtAjEdElEtEnFdEpF9IsGtItG9IwHtMyINM1I9M0ItM2JNM3JtQ4JtQ4JtQ8KtQ8KtQ5J9Q9K9U5J9Q9K9U9K9U9LNU+LNU+LNU+LNU+LNU+LdU/LdU/LdU/LdVDMdY/LtU+LNU8KtQ4JtQ2JNQ0ItMwHtMuHNMsGtIpF9IkEdEhD9EQB2kAAAAAAAABAAQXCZshDtAjENElEtEmE9EoFdEqGNIsGtIuHNItG9IvHNIwHtMxHtMxH9M1I9M1I9MyH9M1I9MyINM2JNM2JNQ2JNQ2JNQ3JNQ3JdQ3JdQ3JdQ3JdQ4JdQ4JtQ4JtQ4JtQ7KdQ4JtQ3JdQ0ItMzINMwHdMtGtIqF9IoFtIlE9EhDtEYCp4BAAUAAAAAAAAAAAACAQ8XCZ0gDdAiD9AjENElEtEmFNEmE9EnFNEoFdEoFtEpF9IqF9ItG9IuG9IuHNIuHNIrGNIvHNIvHNIvHdIvHdIvHdMvHdMwHdMwHtMwHtMwHtMwHtMxHtMxH9MxH9MyH9MzIdMxH9MvHdMuHNIsGdIqF9ImFNEkEtEiD9EYCpsCAQ4AAAAAAAAAAAAAAAAAAAABAAQPBmseDMggDdAfDNAgDdAgDtAhDtAhD9AiD9AiENAjENEjENEnFNEjEdEnFNEnFdEnFdEoFdEoFdEoFdEoFtEoFtEpFtEpFtIpFtIpF9IpF9IqF9IqF9IqF9IqGNIrGNIsGtIrGNIpFtIoFdEmE9ElEtEhDsgQB2kBAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADARQMBF0UBpoYB7kbCMwbCM8bCM8bCM8bCM8cCc8cCdAgDdAgDdAgDdAgDdAgDdAhDtAhDtAhDtAhDtAhDtAiD9AiD9AiD9AiD9EiD9EiENEjENEjENEjENEkEdEkEdEmE9EkEs0eDbsZC5wPBl4DARQAAAAAAAAAAAAAAAAAAAD4AAAAAB8AAOAAAAAABwAAwAAAAAADAACAAAAAAAEAAIAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAQAAgAAAAAABAADAAAAAAAMAAOAAAAAABwAA+AAAAAAfAAA="
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
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
DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}
;}

;}

;{ ⚞──────────────────        INCLUDES           	──────────────────⚟

	#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Ini.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
	#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
	#include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk

	#Include %A_ScriptDir%\..\..\lib\ACC.ahk
	#Include %A_ScriptDir%\..\..\lib\class_Loaderbar.ahk
	#Include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
	#Include %A_ScriptDir%\..\..\lib\class_Crypt.ahk
	#Include %A_ScriptDir%\..\..\lib\class_SQLiteDB.ahk
	#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
	#Include %A_ScriptDir%\..\..\lib\Class_LV_InCellEdit.ahk
	#Include %A_ScriptDir%\..\..\lib\class_UIA_Browser.ahk
	#Include %A_ScriptDir%\..\..\lib\class_UIA_Constants.ahk
	#Include %A_ScriptDir%\..\..\lib\class_UIA_Interface.ahk
	#Include %A_ScriptDir%\..\..\lib\class_7zip.ahk
	#Include %A_ScriptDir%\..\..\lib\Gdip_All.ahk
	#Include %A_ScriptDir%\..\..\lib\IFileDialog.ahk
	#Include %A_ScriptDir%\..\..\lib\objTree.ahk
	#Include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
	#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
	#Include %A_ScriptDir%\..\..\lib\SciteOutput.ahk
	#Include %A_ScriptDir%\..\..\lib\Sift.ahk

;}


/* Müllplatz

cdbxpcmd --burn-data -folder:"Pfad\zum\Ordner" -device:0 -speed:16 -close

QEGetPatSet()


	; Prüft ob der Datensatz bereits vorhanden ist
		;~ If IsObject(result := SQLiteRecordExist("EXPORTS", "PatID", Pat.ID)) {
			;~ PatSet.Record 	    	:= result.Record                 	; die passende Zeile mit allen Spalten
			;~ For colname, column in result.Columns {
				;~ PatSet.Columns[colname].db     	:= column.db		  		; geordnet nach Spaltenbezeichnung
				;~ PatSet.Columns[colname].dbCol 	:= column.dbCol
			;~ }
		;~ else if result {

		----------
		Gui, QE: Default
		Gui, QE: ListView, LvExp
		Loop % LV_GetCount() {
			cRow := A_Index
			LV_GetText(PatID, cRow, 1)
			If (PatID = PatSet.ID) {
				PatSet.Columns["PatID"].lv    := PatSet.ID
				PatSet.Columns["PatID"].LVCol := 1
				Loop % LV_GetCount("Columns")-1 {
					LV_GetText(cellTxt	, cRow	, colNr := A_Index+1)
					LV_GetText(lvCN	, 0       	, colNr)								; liest die Spaltenbezeichnungen
					dbColName := ExportsTbl[colNr]
					PatSet.Columns[dbColName].lv         	:= cellTxt
					PatSet.Columns[dbColName].lvColName 	:= lvCN
					PatSet.Columns[dbColName].lvCol      	:= colNr
				}
				break		; <<< ist richtig hier oder nicht?
			}


*/








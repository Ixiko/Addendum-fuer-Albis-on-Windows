;{ Debugging Einstellungen
global ScriptStart            	:= A_TickCount
global DebugFunctions		:= 0							; Debug Funktionen wie die Erstellung der ScanPool...Log.md Dateien ein- oder ausschalten
global ListScriptLines         	:= 0							; Skriptablauf protokollieren
global OnMessageByPass	:= 1							; OnMessage Funktionen sind immer angeschaltet
;}
;###############################################################
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------------------- Addendum für AlbisOnWindows -----------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;----------------------------------------------------- Modul: ScanPool ---------------------------------------------------------
																   Version := "0.9.90"
;------------------------------------------------------------------------------------------------------------------------------------
;{      Ein Skript zur schnellen Organisation / Signierung neu eingegangener Patientenbefunde (PDF-Format)
	; und  zur vollautomatischen Übernahme in die jeweilige Patientenakte.
	; Sehen Sie sofort, wenn Ihre Angestellten ein neues Dokument bereit gestellt haben und zwar auf jedem Ihrer
	; verbundenen Netzwerk-PC.
	; Verschwenden Sie keine Zeit mit der Suche in Papierdokumenten. ScanPool erkennt automatisch die geöffnete
	; Patientenakte und mit einem Click oder einer Tastenkombination läßt sich ScanPool starten und zeigt Ihnen
	; zum Patienten gehörende Befunde an. Sie können sich auch eine Liste mit allen Dokumenten anzeigen lassen.
	; Öffnen Sie mehrere Dokumente im Foxitreader und drücken Sie auf den 'Signieren'-Button.
	; Das Skript beginnt mit dem Vorgang des passwortgeschützten Signierens. Revisionssicher? PDF Dateien sind
	; nicht sicher zu schützen!
	; Schreibfehler bei der Benennung der Dateinamen mittels Nachname, Vorname, Art des Befundes und Datum,
	; welche immer entstehen können, werden durch unscharfe Stringvergleichsfunktionen weitestgehend
	; automatisch korrigiert werden. Sicherlich ließe es sich komfortabler mit Zugriff auf die Albis-Datenbank
	; gestalten. Doch möchte ich einen eventuellen Datenverlust bei mir oder Ihnen aufgrund meines unbeholfenen
	; Programmierstiles vermeiden. Mir nimmt das Skript einiges an Arbeit ab, gibt mir schneller eine Übersicht
	; und ich spare mir die monatlichen Kosten meiner Archivsoftware, zu welcher ich mehrfach im Jahr keinen
	; Zugriff bekam, da entweder der Zugriff durch einen anderen Nutzer im Moment blockiert sei oder aber
	; ich plötzlich in Besitz einer nicht lizensierten Version war.Eine weitere Zeitersparnis, da ich den Software-
	; betreuer deshalb nicht mehr kontaktieren muss.
;}
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------------- written by Ixiko -this version is from 31.08.2019 ------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ------------------------
;------------------------------ use subject: "Addendum" so that you don't end up in spam folder -----------------------
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ----------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;###############################################################

;TODO's :
;
;{ erledigte TODO's
;| **01.04.2019** | **F~** | Fehlerbehebung in den Funktionen: Übertragen der Befunde, Öffnen einer Patientenakte, Anzeigen aller Befunde (V0.9.85) |
;| **31.07.2019** | **F~** | UTF8 String encoding für die richtige Darstellung von Umlauten hinzugefügt |
;}

;{ unerledigte TODO's
;02. Einstellungen	- Hilfe Gui oder Animation  +0.5
;05. Einstellungen	- Hotkey Anzeige
;08. Einstellungen	- Gui für die notwendigen Einstellungen der ScanPool Gui

;06. Kontrolle		- Meldung an das Praxomat Skript das gerade Befunde einsortiert werden, um zu verhindern das das Praxomatskript eventuell mit Albis interagiert, ebenso sollten auch alle anderen Module blockiert werden (AutoSuspend)
;25. Kontrolle		- !!!! Skript pausieren falls Albis aus irgendeinem Grund abstürzen sollte !!!!

;40. Vorschau 		- automatisiertes DREHEN am besten auch in der Original-PDF Datei, leeres Blatt automatisch raus oder die nächste Seite anzeigen

;38. Scannen		- per OCR die Seitenzahlen erkennen und automatisches Anordnen programmieren,  Seitenzahlerkennung der echten Seite (OCR geht ja nur) vielleicht geht das mit der Routine von feiyue
;39. Scannen		- OCR per Tesseract integrieren oder anders
;47. Scannen     	- von Namen beim Scannen - für die Schwestern
;48. Scannen     	- des Dokumenttitel beim Scannen - z.B. Epikrise Gastroenterologie 4.9.-18.9.2018 per OCR

;11. Albis Pdf's		- !!!!in Albis HINWEIS - das neue BEFUNDE für PATIENTEN DA SIND - über Addendum.ahk!!!
;50. Albis Pdf's		- Index der von Albis umbenannten PDF Dateien erzeugen und diese dem jeweiligen Patienten zuordnen (also die kryptischen Dateinamen sind gemeint)

;54. Signieren		- PDF-Signierung am liebsten mit einem anderen Programm also ohne den FoxitReader - dann kann ein anderer freier PDF Viewer benutzt werden

;42. sonstiges		- Notiz zu einer PDF Datei hinterlassen, z.B. Pat. muss einbestellt werden, oder C 18 PP K 20 - oder noch anderes
;43. sonstiges		- Schnellaufruf zum Eintragen eines Patienten z.B. eine Schwester - um den Patienten einzubestellen , Aufruf über PDF Patientennamen
;53. sonstiges		- Auto-Installation des FoxitReaders auf Clients die diesen noch nicht installiert haben, gibt es eine portable Version?

;49. Listview			- checkbox Häkchen ersetzen gegen andere Hintergrundfarbe
;56. Listview			- schon signierte Dateien im Listview hervorheben

;59. Import			- abgebrochener Importvorgang sollte alle Fenster schließen


;67.
;68.

;}

;{ 1. Scripteinstellungen / Tray Menu

		#NoEnv
		#SingleInstance force
		#Persistent
		#MaxMem 4095	; INCREASE MAXIMUM MEMORY ALLOWED FOR EACH VARIABLE - NECESSARY FOR THE INDEXING VARIABLE
		#KeyHistory 0
		;#WinActivateForce

		#ErrorStdOut
		#MaxThreads, 250
		#MaxThreadsBuffer, On
		#MaxThreadsPerHotkey, 2
		;#Warn All, StdOut

		SendMode Input
		SetWorkingDir %A_ScriptDir%
		SetTitleMatchMode, 2
		DetectHiddenWindows, Off
		DetectHiddenText, Off
		SetControlDelay -1
		SetWinDelay, -1
		SetBatchLines -1
		CoordMode, Mouse, Screen
		CoordMode, Pixel, Screen
		CoordMode, Menu, Screen
		CoordMode, Caret, Screen
		CoordMode, Tooltip, Screen

		FileEncoding, UTF-8

		OnError("FehlerProtokoll")

		If !ListScriptLines
			ListLines, Off

	; Überprüfen, ob ein Albisfenster geöffnet ist und Skript beenden falls nicht
		IfWinExist, ahk_class OptoAppClass
		{
				WinActivate, ahk_class OptoAppClass
				AlbisWinID := AlbisWinID()
				PatID		 := AlbisAktuellePatID()
				If (PatID =1) or (PatID = 2)
								PatID:=""
		} else {
					MsgBox, 4144, Addendum für AlbisOnWindows - ScanPool - Info, Dieses Modul setzt ein laufendes Albis-Programm voraus.`nDas Programm muss nun beendet werden!, 10
					ExitApp
		}

	; Ladefortschritt anzeigen
		Progress, P0 B2 cW202842 cBFFFFFF cTFFFFFF zH25 w400 WM400 WS500,  ......starte ScanPool Modul %Version%, Addendum für AlbisOnWindows - ScanPool - Info, Addendum für AlbisOnWindows - ScanPool Startvorgang, Futura Bk Bt
		WinMove, Addendum für AlbisOnWindows - ScanPool Startvorgang,,, % A_ScreenHeight-(A_ScreenHeight//3)

	; Tray Menu erstellen
		Menu, Tray, NoStandard
		Menu, Tray, Tip, % "      ScanPool" ScanPoolVersion "`n-----------------------`n         Addendum   `nfür Albis on Windows"
		Menu, Tray, Add, über ScanPool, NoNo
		Menu, Tray, Add,
		Menu, Tray, Add, Zeige HotKey's, Hotkeys
		Menu, Tray, Add, Zeige Skript-Objekte, ScriptObjects
		Menu, Tray, Add, Zeige Skript-Variablen, ScriptVars
		Menu, Tray, Add, Skript Neu Starten, SkriptReload
		Menu, Tray, Add, Fenster anzeigen, ShowGui
		Menu, Tray, Add, Beenden, BOGuiClose
	; Tray Eintrag Fenster maximieren auf minimieren ändern
		Menu, Tray, Disable, Fenster anzeigen

	; Taskbar/TrayIcon wird erstellt
		Menu Tray, Icon, % "hIcon: " Create_ScanPool_ico(true)

		OnExit("ScanPoolEndsHere")

	  ; GDI starten
		If (!pToken := Gdip_Startup()) 	{
				MsgBox, 0x40048, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
				ExitApp
		}

		Progress, 5

;}

;{ 2. Variblen Setup / Registry auslesen

	;{ a) ---------------------------------------------------------------- globale Variablen -----------------------------------------------------------------------

		global AddendumDir	:= FileOpen("C:\albiswin\AddendumDir-DO_NOT_DELETE","r").Read()
		global AlbisWinID                           				;ACHTUNG: diese sollten in jedem Skript global sein
		global AlbisWorkDir

		; Gui - Variablen
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		global hBO                                               	;Handle des ScanPool Gui (main gui handle)
		global hBOPV 	 				                 			;Handle des Pdf Preview Gui
		global hBOPT                                            	;Handle des Process Infofenster
		global PTText3                                        		;Name des Static-Control im Process Infofenster
		global hLV1                                                	;Handle der Listview
		global Listview1                                            	;Name des Listview-Control
		global MOPdf                           	               	;beinhaltet den Namen der Pdf-Datei über der gerade der Mauszeiger steht
		global RamDisk                                        		;Laufwerksbezeichnung für die RamDisk
		global RamDiskPath                    					;Laufwerkspfad für die RamDisk
		global dpiFactor, DpiShow

		; Variablen zur Kontrolle des Skriptablaufes
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		global P, PAdd                                        		;Variablen für das Fortschritt-Gui (automatisierter Signierprozeß) um den Fortschritt anzeigen zu können
		global Running, WhatIsRunning                    	;neue Idee der Scriptüberwachung, enthält Info welche andere Timer, Prozeß, Funktion gerade läuft
		global ExecPID                                        		;Consolen PID
		global PreviewerWaiting := 1				    		;für PdfPreviewer - verhindert das der Previewer vor Beendigung unterbrochen wird

		; alle Variablen für Daten zu den PDF-Dateien
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		global BefundOrdner                    					;BefundOrdner - der Ordner in dem sich die ungelesenen PdfDateien befinden
		global PdfIndexFile                                    	;Pfad und Name der PatIndex Datei
		global PageSum 	:= 0                    				;die Gesamtzahl aller Seiten der Dateien im Befundordner
		global FileCount 	:= 0                    				;Anzahl der Pdf Dokumente im Ordner
		global ScanPool	:= []                    				;Array enthält die Daten zu den Pdf Dateien
		global pdfError		:= 0                    				;zählt Lesefehler von PDF Dateien
		global oPat          	:= Object()                      	;Patientendatenbank Objekt für Addendum
		global PdfData     	:= Object()                      	;enthält Daten der aktuell in Arbeit befindlichen PdfDatei
		global SPData		:= Object()                    	;Objekt für ini Daten

		; Hook - Variablen
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		global SpEHookHwnd1, SpEvent1					;Handle des WinEventHook auslösenden Controls oder Fenster und Event-Nr - für neue Fenster
		global SpEHookHwnd2, SpEvent2					;Handle des WinEventHook auslösenden Controls oder Fenster und Event-Nr - für schließende Fenster
		global SpClass, SpTitle, SpProc1, SpProc2		;eventhook: Class, Title und ProcessName
		global ShHookHWnd, ShHookEvent				;shellhook	: hwnd (lParam) und event (wParam)
		global handlerrunning1,handlerrunning2		;für die WinEvent Funktionen, wenn diese flag gesetzt ist, läßt sich die jeweilige Funktion nicht erneut aufrufen
		global foxitHook                                     		;handle Kopie wird erst auf 0 gesetzt wenn der hook Funktion vollständig bearbeitet ist
 		global hDokSig                                        		;Handle des gerade bearbeiteten Signierfensters
		global PatientNichtVorhanden := 0				;flag für alle Prozesse welche eine Patientenakte in Albis öffnen müssen

		; sonstige - Variablen
		; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		global CopyOfData                    					;Interskript Kommunikationsvariable
		global HighDpi                                        		;flag für eine skalierbare Darstellung der Programmfenster
		global Hinweis1, Hinweis2, Hinweis3
		global Hinweis4, Hinweis5
		global docImported
		global PDFReaderWinClass                    		;ahk_class des PDFViewer/Reader

	;}

	;{ b)	-------------------------------------------------------------- für Fensteranimation ----------------------------------------------------------------------

			FADE     		:= 524288
			SHOW   		:= 131072
			HIDE      		:= 65536

		;Converting the above to Hexdecimal value
			FADE_SHOW:= GetHex(FADE+SHOW) 	; Converts to 0xa0000
			FADE_HIDE	:= GetHex(FADE+HIDE)     	; Converts to 0x90000
			Duration 		:= 300                             	; Duration of Animation in milliseconds
	;}

	;{ c)	---------------------------------------------------- diverse Variablen setzen oder definieren ---------------------------------------------------------

			Pat                	:= []                                                              	;
			AlbisPatient      	:= []                                                              	;Vor- und Nachname des aktuellen Patienten in Albis
			arr                	:= []                                                              	;
			chfiles           	:= []                                                              	;Dateinamen der ausgewählten Pdf-Dateien (derzeit werden geöffnete FoxitReaderfenster dafür ermittelt)
			Name           	:= []                                                              	;Array der die aus den Dateinamen ausgelesenen Patientennamen enthält
			PdfViewer			:= Object()                                        				;PID und WinTitle des PDF Anzeigeprogrammes
			NPRWin_old		:= 0
			NPRWin			:= 0                                                             	;Counter für dieses Objekt
			RQueue        	:= Object()
			UserDpi        	:= Object()                                                      	;
			Praxomat        	:= Object()                    	 	            				;enthält später die Koordinaten des Praxomat Gui
			BO                	:= Object()                    	 			            		;enthält später die Koordinaten des ScanPool Hauptfensters
			BOPV				:= Object()                                        				;Koordinaten: des Vorschaufensters
			SWI               	:= Object()
			CWI					:= Object()                                        				;Koordinaten: steht für ChildWindow
			BOPT				:= Object()                                        				;Koordinaten: BOProcessTip
			BOPic            	:= Object()                    	 		            			;enthält pBitmap und hBitmap für die Preview
			pdfshow        	:= 0                                                              	;
			DocIndex        	:= 0                                                              	;
			Aktualisieren    	:= 0                                                              	;
			pdfError        	:= 0                                                              	;
			LVBGColor    	:= "6d8fff"                                                      	;Listview Background Farben, old LVBGColor2:= "194bbb"
			LVBGColor1    	:= "6d8fff"
			LVBGColor2    	:= "8fA2ff"
			Switch           	:= 1                                                              	;
			PraxoWas    		:= 0                                                               	;PraxoWas - wenn das PraxomatGui da ist wird die Flag auf 1 gesetzt
			NoRestart			:= 1                                                               	;flag ermöglicht das neu einlesen der PatIndex-Datei
			LVDelFlag 		:= 0                                                            		;flag für die inkrementelle Suche
			Signate				:= 0                                                            		;flag für BOUebertragen

		;Versuch eine RamDisk (ImDisk.exe) zu starten, funktioniert leider uunzuverlässig (es läuft auf einem Client, auf einem anderen nicht)
			;RamDiskPath:= AddendumDir . "\include\RamDisk"
			;RamDisk:= CreateRamDrive(16, "Addendum", RamDiskPath)
	;}

	;{ d) ------------------------------------------------------------------ Hinweisdialoge -------------------------------------------------------------------------

			Hinweis1 =
			(LTrim
			Es konnten keine neuen Dokumenten für den Patienten %Vorname% %Name%
			gefunden werden. Soll ich eine Volltextsuche in den PDF Dokumenten durchführen?

			Beachten Sie das eine Volltextsuche länger dauern kann.
			)

			Hinweis2 =
			(LTrim
																						Lieber Nutzer,

			die folgenden Programmteile versuchen anhand des Dateinamens eine Zuordnung des Befundes zum entsprechenden Patienten vorzunehmen.
			Im ersten Versuch wird dieses Skript einen Patientenaufruf mit den im Dateinamen vorhanden Patientennamen versuchen. Z.B. Schulz, Michaerl.
			Wie unschwer zu erkennen, enthält der Name einen Schreibfehler. Albis wird somit den Patienten nicht finden können.
			Als nächstes wird das Skript den vermeintlichen Nachnamen - hier 'Schulz' eingeben. In dem sich öffnenden Listview Fenster wird das Skript dann
			die naheliegendste Variante von Michaerl suchen und nach Bestätigung durch Sie den PDF Befund in die Akte importieren.
			Ich plane in Zukunft "smarter" Vorzugehen. Soweit mein Verständnis cognitiver digitaler Verarbeitung reicht. Die Userinteraktion soll in Zukunft
			exakt und so minimal wie möglich werden. Zum Starten eines Vorganges sollen maximal ein Tasten- oder Mausklick und zum Abschließen auch
			maximal ein Klick reichen. Ziel ist das der Computer für uns arbeitet und nicht wir für ihn wie es bisher der Fall ist!

			)

			Hinweis3 =
			(LTrim
			Bitte überprüfen Sie ob Albis noch läuft!
			Wenn Albis keine Probleme zu machen scheint,
			drücken Sie bitte auf 'Ja'!
			Bei Auswahl von 'Nein' erfolgt jetzt ein Abbruch
			der Signatur Funktion.
			)

			Hinweis4 =
			(LTrim
			Die richtige Patientenakte scheint noch immer nicht geöffnet zu sein.
			Sie können jetzt mit 'JA' bestätigen das die gewünschte Patientenakte
			dennoch geöffnet ist oder drücken Sie 'NEIN' um die manuelle Eingabe
			des Namens noch einmal durchzuführen.
			)

        	Hinweis5 =
			(LTrim
			Die eingestellte Signatur: "%Darstellungstyp%" ist auf diesem Computer
			nicht oder nicht mehr vorhanden. Wählen Sie bitte eine Signatur aus und
			drücken Sie dann in diesem Dialog auf 'Ok'.
			)
;}

	;{ e) ------------------------------------------- ScanPool BefundOrdner/TempPath/PDFReader/SignaturRechte ----------------------------------------

		 ; auslesen ob das ScanPool den PdfIndex nicht zu Ende speichern konnte
			IniRead, ScriptAblauf		    	, %AddendumDir%\Addendum.ini, ScanPool, ScriptAblauf
			If InStr(ScriptAblauf, "canceled")
					NoRestart := 0                    	;hier müsste eine Funktion hin welche die PdfIndex auf Integrität testen sollte

		 ; Albis Arbeitsverzeichnis einlesen
			IniRead, AlbisWorkDir			, %AddendumDir%\Addendum.ini, Albis, AlbisWorkDir

		 ; BefundOrdner = Scan-Ordner für neue Befundzugänge
			IniRead, BefundOrdner			, %AddendumDir%\Addendum.ini, ScanPool, BefundOrdner

		  ; der Name des verwendeten PDF Anzeigeprogramms benötigt für eine Abfrage
			IniRead, PDFReaderName		, %AddendumDir%\Addendum.ini, ScanPool, PDFReaderName

		  ; das ClassNN des Hauptfenster des verwendeten PDF Anzeigeprogramms - wird benötigt für alle Automationsprozesse
		  ; (ich hoffe man kann so auch andere Reader einbinden)
			IniRead, PDFReaderWinClass	, %AddendumDir%\Addendum.ini, ScanPool, PDFReaderWinClass

		  ; der PDFReader der verwendet werden soll, Achtung: mein Automation funktioniert nur mit dem FoxitReader
			IniRead, PDFReaderFullPath	, %AddendumDir%\Addendum.ini, ScanPool, PDFReaderFullPath
			PDFReaderFullPath:=  StrReplace(PDFReaderFullPath, "%AddendumDir%", AddendumDir) 			;falls User den Reader aus einem anderen Verzeichnis benutzen möchte

		  ; welche Clients dürfen überhaupt signieren und/oder welche Nutzer dürfen
			IniRead, SignatureClients		, %AddendumDir%\Addendum.ini, ScanPool, SignatureClients

		  ; der Übersichtlichkeit wegen werden nach und nach die einzelnen ini-Variablen in diesem Objekt untergebracht werden
			SPData:= BuildScanPoolDataObject()

		; Consolenprogramme um ohne jeglichen PDFReader PDF Dateien digital signieren zu können
			;IniRead, PdfMachinePath		, %AddendumDir%\Addendum.ini, ScanPool, PdfMachinePath

		  ; Temporären- und Backuppfad im Befundordner anlegen
			PdfIndexFile	:= BefundOrdner . "\PdfIndex.txt"
			TempPath		:= BefundOrdner . "\Temp"
			BackupPath	:= BefundOrdner . "\Backup"

			If !InStr(FileExist(TempPath), "D")
			{
					FileCreateDir, % TempPath
					If ErrorLevel
							MsgBox, 1, Addendum für Albis on Windows, % "Der Ordner für temporäre Vorschaubilder`n(" TempPath ")`nkonnte nicht angelegt werden"
					ExitApp
			}

			If !InStr(FileExist(BackupPath), "D")
			{
					FileCreateDir, % BackupPath
					If ErrorLevel
							MsgBox, 1, Addendum für Albis on Windows, % "Der Ordner für temporäre Vorschaubilder`n(" BackupPath ")`nkonnte nicht angelegt werden"
					ExitApp
			}


	;}

	;{ f) -------------------------------------------------------------- Client PC Name finden --------------------------------------------------------------------

			CompName:= StrReplace(A_ComputerName, "-")
	;}

	;{ g) ---------------------------------------------- ScanPool Gui Einstellungen aus der INI einlesen ------------------------------------------------------

			IniRead, BOFontOptions, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOFontOptions"			, S8 CDefault q5
			IniRead, BOFont			, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOFont"                    	, Futura Bk Bt
			;IniRead, BOPosXY     	, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOXYPos-" CompName
			IniRead, BOw				, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOwUSer"					, 450
			IniRead, BOh				, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOhUSer"					, 400
			IniRead, LVBGColor		, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOListViewBgColor	" 	, 6d8fff
			IniRead, LVBGColor1		, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOListViewBgColor1"	, 6d8fff
			IniRead, LVBGColor2		, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOListViewBgColor2"	, 8fA2ff
			IniRead, UserChoice		, % AddendumDir "\Addendum.ini", % "ScanPool", % "DpiPdfVorschau"			, 1

			UserDpi			:= {1: "Aus", 2: "winzig", 3: "klein", 4: "mittel", 5: "groß", 6: "riesig", "Aus": 0, "winzig": 28, "klein": 36, "mittel": 48, "groß": 72, "riesig": 144, "Font1": 0, "Font2": 1, "Font3": 2, "Font4": 4, "Font5": 6, "Font6": 8, "Wahl": ""}
			UserDpi[Wahl]	:= UserChoice
			dpiShow			:= UserDpi[UserChoice]
			dpi					:= UserDpi[(dpiShow)]
			FontChoose		:= "Font" . UserChoice
			FontSize			:= UserDpi[(FontChoose)]
	;}

	;{ h) ------------------------------------------------------------------- PDFReader Pfad -----------------------------------------------------------------------

			SplitPath, PDFReaderFullPath, PDFReaderExe
	;}

	;{ i) ---------------------------------------------- Position und Größe des ScanPool Fenster -----------------------------------------------

			BOxMrg					:= 5                                        	; der GUI y Rand
			BOyMrg					:= 0                                        	; der GUI y Rand
			BOw                     	:= BOw - 14			            		; 14 ist die Rahmenbreite des Fenster insgesamt

		; Anpassungen an verschiedene Monitorauflösungen (2k oder 4k))
			If InStr(BOPosXY, "ERROR") || !BOPosXY
				BOMon             	:= GetMonitorIndexFromWindow(AlbisWinID())
			else
				BOMon                	:= StrSplit(BOPosXy).3            	; Nr des Monitor auf welchem das Fenster zuvor war

			SysGet, mon, MonitorWorkArea, % BOMon
			monHeight            	:= monBottom - monTop

			If monRight > 1920
					HighDpi := 1, PVZoom:= 144/dpi, DPIFactor := 1.5
			else
					HighDpi := 0, PVZoom:= 72/dpi	, DPIFactor := 1

		; Position der Gui auf die zuletzt gespeicherte Position setzen
			If InStr(BOPosXY, "ERROR") || !BOPosXY
			{
					BOx := monRight - BOw - 12	                		; Abstand zum Monitorrand - per Defaul 30px
					BOy := BOyNormal	:= 43		                 		; Abstand zum oberen Bildschirmrand (y Position AfxControlBar901 in Albis)
			}
			else
			{
					BOx	:= StrSplit(BOPosXy).1
					BOy	:= BOyNormal := StrSplit(BOPosXy).2

				; damit das Fenster nicht ausserhalb des sichtbaren Monitorbereiches liegt
					If BOx > monRight || BOy > monHeight
					{
							BOx := monRight - BOw - 12
					        BOy := BOyNormal	:= 43
					}
			}

			ToolTip, % BOx ", " BOy

		; restlichen Größen und Positionen
			BOhShrink				:= 100                    					; innerhalb der Gui - Platz lassen zum unteren Rand für das Suchfeld
			BOText1W				:= 180                    					; x,y werden dynamisch erzeugt - Text wird zentriert in der GUI
			BOEdit1W		     		:= 180                    					; x,y werden dynamisch erzeugt - Text wird zentriert in der GUI
			BOh	:= BOhMax		:= monHeight - BOy	+ 5   		; maximale Höhe der GUI anhand der MonitorWorkArea
			BOLv1H					:= (BOh - BOhShrink - 5) < 300 ? 300 : (BOh - BOhShrink - 5)
	;}

	;{ j) ------------------------------- bestimmen ob eine Patientenakte geöffnet ist, durch Ermitteln eines Namens ------------------------------------

		;wenn eine Patientenakte geöffnet ist, wird nur ein kleines Fenster angezeigt für die zugehörigen Patientenbefunde
			If InStr(AlbisGetActiveWindowType(), "Patientenakte")
			{
					PatID        	:= AlbisAktuellePatID()
					AlbisPatient  	:= VorUndNachname(AlbisCurrentPatient())
					Bezeichner	:= "PatID: " PatID "; " AlbisPatient[1] ", " AlbisPatient[2]
			}
			  else
			{
					PatID        	:= ""
					Bezeichner	:= "Liste aller ScanPool Dokumente"
					BOh            	:= monHeight - BOyNormal
			}


	;}

	;{	k)	---------------------------------------------------------------- Listview Konstanten ----------------------------------------------------------------------

				global tipDuration                     		:= 4000
				global LVIR_LABEL                     		:= 0x0002		;LVM_GETSUBITEMRECT constant - get label info
				global LVM_GETITEMCOUNT 			:= 4100			;gets total number of rows
				global LVM_SCROLL 				    		:= 4116			;scrolls the listview
				global LVM_GETTOPINDEX 				:= 4135			;gets the first displayed row
				global LVM_GETCOUNTPERPAGE   	:= 4136			;gets number of displayed rows
				global LVM_GETSUBITEMRECT    		:= 4152			;gets cell width,height,x,y

	;}

	;{ l) ---------------------------------------------------------- Patientendatenbank einlesen ----------------------------------------------------------------

			IniRead, AddendumDBPath, %AddendumDir%\Addendum.ini, Addendum, AddendumDBPath
			AddendumDBPath:= StrReplace(AddendumDBPath, "%AddendumDir%", AddendumDir)
			If !InStr(AddendumDBPath, "Error")
						oPat:= ReadPatientDatabase(AddendumDBPath)

	;}

	;{ m)	--------------------------------------------- eventuell noch vorhandene Vorschaubilder entfernen -------------------------------------------------
		Loop
		{
				picname:= % TempPath . "\sppreview-" . SubStr("00000" . A_Index, -6) . ".png"
				If !FileExist(picname)
					break
				FileDelete, % picname
		}

	;}

	Progress, 30

;}

;{ 3.	ScanPool Gui + Einlesen des Index - die Variable FileCount wird zur Erstellung des Listview gebraucht

	;-: Einlesen des Index
		NoRestart	:= 0
		;FileCount	:=MaxFiles:= ReadIndex()
		FileCount	:=MaxFiles:= ReadPdfIndex(BefundOrdner . "\PdfIndex.txt")

	;-: -------------------------------------------------------------------------
	;-: ---------------	 allgemeine Einstellungen 	----------------
	;-: ------------------------------------------------------------------------- ;{
		Gui, BO: NEW,  +Resize -Caption +ToolWindow +MinSize386x400 +MaxSize%BOw%x%BOhMax% +LastFound +HwndhBO +OwnDialogs +AlwaysOnTop 0x94030000
		Gui, BO: Color  	, % LVBGColor
		Gui, BO: Font    	, % BOFontOptions	, % BOFont
		Gui, BO: Margin	, % BOxMrg          	, % BOyMrg
		Gui, Add, Progress, % "x-1 y-1 w" BOw " h" BOhMax " vBOProg1 Background202842 Disabled hwndHPROG1"
		Control, ExStyle, -0x20000, , ahk_id %HPROG1%
	;}

	;-: -------------------------------------------------------------------------
	;-: ----------------	 oberer Gui Bereich 	----------------
	;-: ------------------------------------------------------------------------- ;{

	;-: ----------------	           Buttons      	----------------	;{
	;-: Öffne alle Befunde des ausgewählten Patienten
		Gui, BO: Add, Button, % "x5 y" (BOyMrg + 2) " BackgroundTrans 	        			  vBOButton1 gBOtheButtons"		    					, % "Öffne gleiche"
		GuiControlGet, BOButton1, BO: Pos	;für dynamische Positionierung
	;}
	;-: ----------------	         Signieren     	----------------	;{
		Gui, BO: Add, Button, % "x" (5 + BOButton1W + 5)  " y" (BOyMrg + 2) 				" vBOButton2 gBOtheButtons Disabled"				, % "alphabet."
		GuiControlGet, BOButton2, BO: Pos
	;}
	;-: ----------------	       alle Befunde   	----------------	;{
		Gui, BO: Add, Button, % "x"(5 + BOButton2X + BOButton2W) " y" (BOyMrg + 2) " vBOButton3 gBOtheButtons"                    			, % "alle Befunde"
	;}
	;-: ----------------	 Vorschau - UpDown ----------------	;{
		Gui, BO: Font, S8 q5 , %BOFont%
		Gui, BO: Add, Text, % "x+2 y" (BOyMrg+2) " w65 cWhite Center Backgroundtrans"                                                            			, % "Pdf Vorschau"
		Gui, BO: Font, S12 q5 , %BOFont%
		Gui, BO: Add, Text, % 		 "y+" (BOyMrg)   " w65 cWhite Center Backgroundtrans vBOPreviewDpi"                                        		, % dpiShow
		GuiControlGet	, BOPreviewDpi	, BO: Pos
		GuiControl		, BO: Move	, BOPreviewDpi, % "y" BOPreviewDpiY-3
		If HighDpi
			Gui, BO: Add, UpDown, % "x+3 y" (BOyMrg+2) " h" BOButton1H " vBOUpDown1 gBOtheButtons Range1-6 -16"                    	,  % UserChoice  ;, Aus|winzig|klein|mittel|groß|riesig
		else
			Gui, BO: Add, UpDown, % "x+3 y" (BOyMrg+2) " h" BOButton1H " vBOUpDown1 gBOtheButtons Range1-5 -16"                    	,  % UserChoice  ;, Aus|winzig|klein|mittel|groß|riesig
		BOUpDown1:= UserChoice
	;}
	;-: ----------------	  Fenster schließen 	----------------	;{
		Gui, BO: Font, S8 q5, %BOFont%
		Gui, BO: Add, Button, % "x" (BOw - 70) " y" (BOyMrg + 2) " Center vBOCancel gBOtheButtons"                                        				, % "X"
		GuiControlGet, BOCancel, BO: Pos                                                                                                                                                                			;für dependente Positionierung
		GuiControl, BO: Move, BOCancel, % "x" (BOw - BOCancelW - BOxMrg) " y" (BOyMrg+2) " h" BOButton1H
		;}
	;-: ----------------	 Fenster minimieren	----------------	;{
		Gui, BO: Add, Button, % "x" (BOw - 70) " y" (BOyMrg + 2)  " Center vBOMin gBOtheButtons"                                            				, % "_"
		GuiControlGet, BOMin, BO: Pos
		GuiControl, BO: Move, BOMin, % "x" (BOw - BOCancelW - BOMinW - BoxMrg - 5) " y" (BOyMrg+2) " h" BOButton1H
	;}
	;-: ----------------	            Hilfe         	----------------	;{
		Gui, BO: Font, %BOFontOptions% , %BOFont%
		Gui, BO: Add, Button, % "x" BOCancelX - 10 " y" (BOyMrg+2) " vBOButton4 gBOtheButtons Hide"                                        		, % "?"
		GuiControlGet, BOButton4, BO: Pos                                                                                                                                                                			;für dependente Positionierung
		GuiControlGet, BOMin, BO: Pos
		GuiControl, BO: Move, BOButton4	, % "x" (BOMinX - BOButton4W - 5)
	;}
	;-: ----------------	          ListView       	----------------	;{
	;-: Listview1 Element wird dynamisch unterhalb der oberen Buttons positioniert :-
		BOLv1Y:= ( BOButton1Y + BOButton1H + BOyMrg + 5 ) ;- 15
		BOLv1H:= ( BOh - BOhShrink - 5 ) < 300 ? 300 : ( BOh - BOhShrink - 5 )
		GuiControl, BO: Move, BOProg1, % "h" BOLv1Y + 2                                                                                                                                            	                                	;Progress1 Neupositionierung
		/* benutzte Listview-Styles

		 	LVS_EX_DOUBLEBUFFER    	= LV0x10000                    			LVS_EX_LABELTIP            		= LV0x04000
			LVS_EX_FLATSB	                	= LV0x00100
			LVS_EX_FULLROWSELECT		= LV0x00020
			----------------------------------------------------
														= LV0x10120

			LVS_SHOWSELALWAYS			=     0x00008	;Standard
			LVS_SORTASCENDING			=     0x00010
			-----------------------------------------------------
														=     0x00018
		*/
		Gui, BO: Add, Listview, % " x0 y" (BOLv1Y) " w" (BOw) " h" (BOLv1H) " vListView1 HWNDhLV1 r30 gBOListView Background" LVBGColor " AltSubmit 0x18 LV0x10120 Checked NoSort NoSortHdr Count" FileCount, % "    Befund(e)|S." ;LV0x15100
		LV_ModifyCol(1, BOw - 55)
		LV_ModifyCol(2, "30 Integer Right")
	;}

	;}

	;-: -------------------------------------------------------------------------
	;-: ----------------	 unterer Gui Bereich 	----------------
	;-: ------------------------------------------------------------------------- ;{

	;-: ----------------	'Suche nach Befund'	----------------	;{
		BOText1X := BOw - BOText1W - 5
		BOText1Y := BOLv1H + 6
		Gui, Add, Progress, % "x" -1 " y" (BOLv1H + BOLv1Y) " w" (BOw + 2) " h10 vBOProg2 HWNDhProg2 Background202842 Disabled"                                    		;Progress2 wird hinzugefügt
		Gui, BO: Font, cFFFFFF, %BOFont%
		Gui, BO:Add, Text, % " x" BOText1X " y" BOText1Y " w" BOText1W " BackgroundTrans Center vText1"				                        			, Suche nach Befund(en):
	;}
	;-: ----------------	    Suchfeld Edit1   	----------------	;{
		GUIControlGet, Text1, BO: Pos
		BOText1H := Text1H
		BOEdit1X	 := BOw - BOEdit1W- 5,
		BOEdit1Y	 := BOText1Y + BOText1H + 6                                        	 ;plus ein kleiner Abstand - jetzt mit 10px Abstand zum rechten Rand
		Gui, BO: Font, Normal c000000, % BOFont
		Gui, BO:Add, Edit, % " x" BOEdit1X " y" BOEdit1Y " w" BOEdit1W " vBOEdit1 gBOIncFuzzySearch HWNDhBOEdit1"
	;}
	;-: ----------------	  Text: Addendum..	----------------	;{
		; Fenstertextbezeichnung in dieser GUI nicht als Titelleiste sondern unten links in der Gui - mehr Info's auf kleinem Raum möglich :-
		Gui, BO: Font, S8 cFFFFFF q5, Futura Bk Bt
		Gui, BO: Add, Text, % " xm y" BOText1Y " vWTitle1 BackgroundTrans"		                                                                                     		, Addendum für AlbisOnWindows
		GuiControlGet, BoCtrl_, BO: Pos, WTitle1
		BOProg2H := BoCtrl_H
	;}
	;-: ----------------	   Text: ScanPool..   	----------------	;{
		Gui, BO: Font, S24 cFFFFFF q5 Bold, Futura Bk MD
		Gui, BO: Add, Text, % " xm y" BOEdit1Y " vWTitle2 BackgroundTrans"                                                                                            		, ScanPool
		GUIControlGet, BoCtrl_, BO: Pos, WTitle2
		BOProg2H += BoCtrl_H
	;}
	;-: ----------------	    Text: Version..   	----------------	;{
		Gui, BO: Font, s7 cFFFFFF q5, Futura Bk Bt
		Gui, BO: Add, Text, % " x" (BoCtrl_X + BoCtrl_W - 10) " y" (BoCtrl_Y + BoCtrl_H - 2) " w" 50 " vWTitle3 BackgroundTrans Center"			, % Version
		GuiControlGet, BoCtrl_, BO: Pos, WTitle3
	;}

	;}

	;-: -------------------------------------------------------------------------
	;-: ---------------    Ausrichten und Anzeigen  ----------------
	;-: ------------------------------------------------------------------------- ;{
		GuiControl, BO: Move, BOProg2, % "h" ( BOProg2H + 1 )

	;-: zwei Texte die nur angezeigt werden wenn es notwendig ist
		Gui			, BO: Font		, s12 c000000 q5 Normal, % BOFont
		Gui			, BO: Add		, Text		, % "x5 y80   w" (BOw-19) " vNDoc1 Center BackgroundTrans", % "Keine neuen Dokumente vorhanden für:"
		GuiControl, BO: Hide, NDoc1
		Gui			, BO: Font		, s22 c000000 q5 Bold, % BOFont
		Gui			, BO: Add		, Text		, % "x5 y100 w" (BOw-19) " vNDoc2 Center BackgroundTrans", % " "
		GuiControl, BO: Hide, NDoc2

	;-: StatusBar wird erstellt :-
		Gui, BO: Font, S8 c000000 q5 Normal, Futura Bk Bt
		Gui, BO: Add, StatusBar
		SB_SetParts(30, 250, 0, 150)

	;-: Gui anzeigen - Gui Default machen
		Gui, BO: Show, % "x" BOx " y" BOy " w" BOw " h" BOh " Hide", ScanPool - Befunde                                        				;w%BOw% h%BOh%

	;-: ScanPool Gui zeigen
		If !WinExist("ahk_class TV_ControlWin") and !IsRemoteSession()
				AnimateWindow(hBO, Duration, FADE_SHOW)
		Gui, BO: Show
		Gui, BO: Default

	;--kleiner Trick damit das Fenster in seiner maximalen Breite angezeigt wird
		WinMove, ahk_id %hBO%,,,% A_GuiHeight-1, % A_GuiWidth
		WinGetPos,,, GuiWidth,, ahk_id %hBO%
		GuiControl, BO: MoveDraw, Listview1, % "w" GuiWidth
		WinSet, Redraw,, ahk_id %hBO%
	;}

	;-: -------------------------------------------------------------------------
	;-: ---------------	Tooltip Handler + Coloring ----------------
	;-: ------------------------------------------------------------------------- ;{
		Running :="bereite Datei-`nanzeige vor", addingFiles := 1
		SetTimer, BOProcessTipHandler, 40

	;-: Listview coloring - Vorbereitung (Library - LVA.ahk)
		LVA_ListViewAdd("ListView1", "0x" . LVBGColor1)

	;-: ToolTips für Buttons
		;leer
	;}

	;-: -------------------------------------------------------------------------
	;-: ---------------	   ListView Context Menu 	----------------
	;-: ------------------------------------------------------------------------- ;{
		Menu, BOContextMenu, Add, Dateinamen ändern                    	    	, BOEditItem
		Menu, BOContextMenu, Add, zugehörige Akte öffnen                     	, BOAkteOeffnen
		Menu, BOContextMenu, Add, alle Befunde mit Anzeigen                 	, BOAlleOeffnen
		Menu, BOContextMenu, Add, Befund(e) übertragen                         	, BOUebertragen
		;Menu, BOContextMenu, Add, Befund(e) übertragen und Signieren     	, BOUeberSignieren
		Menu, BOContextMenu, Add, Metadaten anzeigen                         	, BOMetaDatenAnzeigen
		Menu, BOContextMenu, Add, Pdf ContentCheck                             	, BOContentCheck
		Menu, BOContextMenu, Add, Ordner neu indizieren                       	, BONewIndex
		Menu, BOContextMenu, Add,
		Menu, BOContextMenu, Add, Datei(en) löschen                    		       	, BODeleteItem
		Menu, BOContextMenu, Add,
		Menu, BOContextMenu, Add, Skript beenden                    			       	, BOGuiCheckClose
	;}

		Progress, 50
;}

;{ 4. Labelaufruf Einlesen des ScanPool-Ordners, Timer Labels und OnMessage Funktion (Autoexecute)

		OnMessage(0x4E, "LVA_OnNotify")

	;erstes Einlesen aller PDF Dateien
		gosub BOAddFilesToScanPool

	;Progressfenster Fortschritt erhöhen
		Progress, 70

	;Progressfenster wird beendet
		Progress, 100
		sleep, 300
		Progress, Off
		WinKill, % "ahk_id " hProgress1

	; markiert schon signierte Befunde und sammelt diese in einem Array
		If Instr(SignatureClients, CompName)		{
				MSPdfIdx:= 0, Signed:= [], SignedIdx:= 0
				SetTimer, BOMarkSigned, -10
		}

	; Gui auf Default setzen, Listview1 auf Default setzen
		Gui, BO: Default
		GuiControl, MoveDraw, BO: Listview1, % "w" A_GuiWidth

		PraxTT( Round((A_TickCount-ScriptStart)/1000, 2) "s für das Einlesen und Aktualisieren`nvon " FileCount " Befunden.", "5 4")

	; freigeben von Variablen die nicht mehr gebraucht werden
		BOButton1X:=BOButton1Y:=BOButton1W:=BOButton1H:=BOButton2X:=BOButton2Y:=BOButton2W:=BOButton2H:=BOButton4X:=BOButton4Y:=BOButton4W:=BOButton4H:=""
		BOCancelX:=BOCancelY:=BOCancelW:=BOCancelH:=Text1X:=Text1Y:=Text1W:=Text1H:=BOPreviewDpiX:=BOPreviewDpiY:=BOPreviewDpiW:=BOPreviewDpiH:=""

	; wenn das Praxomatfenster beim Start des Skriptes geöffnet ist dann
		If (SpEHookHwnd1:= WinExist("PraxomatGui ahk_class AutoHotkeyGUI"))
               	gosub, SpEvHook_WinHandler1

	; Farbanzeige jetzt auffrischen
		LVA_Refresh("ListView1")

	; Gui verschieben per OnMessagefunktion
		;OnMessageStatus(1)
		OnMessage(0x200, "BOWM_MOUSEMOVE")
		OnMessage(0x201, "BOWM_LBUTTONDOWN")

	; Skriptkommunikation
		OnMessage(0x4a, "Receive_WM_COPYDATA")

	; WinEventHooks initialisieren
		gosub InitializeWinEventHook
		;gosub InitializeShellHook


;}

;{ 5. Hotkeys

	;-: Hotkeys für die Interaktion mit dem Vorschaufenster
		Hotkey, IfWinActive   	, ScanPool - Pdf Vorschau
		HotKey, WheelUp	   	, BOPVPageUP
		HotKey, WheelDown 	, BOPVPageDown
		HotKey, Left       	   	, BOPVPageTurnLeft
		HotKey, Right       	   	, BOPVPageTurnRight
		Hotkey, IfWinActive

	;-: Hauptfenster Hotkeys
		Hotkey, IfWinActive   	, ScanPool - Befunde
		Hotkey, ^!s            	, ScriptVars
		Hotkey, ^!l            	, ScriptLines
		Hotkey, ^!n            	, SkriptReload
		Hotkey, IfWinActive

		Hotkey, IfWinExist   	, ScanPool - Befunde
		Hotkey, Esc              	, BOGuiClose
		Hotkey, IfWinExist

		Hotkey, IfWinActive	, ahk_class %PDFReaderWinClass%
		Hotkey, Insert			, PdfSignatur
		Hotkey, IfWinActive

return
;}

;{ 6. ScanPool BO GUI LABELS - alle Routinen damit in der Gui etwas angezeigt wird
;     ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

BOtheButtons:                               	;{ 	ScanPool Gui g-Label und GuiEvent Abhandlungen

	;keine Reaktion bei ButtonClick wenn ein Hinweisfenster geöffnet ist
		If StrLen(Running) > 0
				return

	;{ ----------------------Abbruch Button--------------------------

		if Instr(A_GuiControl, "BOCancel") {
            BOExit ++
            ExitApp
		}

	;}

	;{ --------------------------gleiche--------------------------------

		if Instr(A_GuiControl, "BOButton1") {
            gosub BOAlleOeffnen
            return
		}

	;}

	;{ ---------------------sortieren Button--------------------------

		if Instr(A_GuiControl, "BOButton2")
		{
           		;gosub BOSaveSignedPdf
            	return
		}

	;}

	;{ -----------------------alle Befunde-----------------------------

		if Instr(A_GuiControl, "BOButton3")
		{

				Vorname:=Name:=Pat[1]:=Pat[2]:=AlbisPatient[1]:=AlbisPatient[2]:= ""
				NoDoc	 := 0
				Running := "alle Befunde eines`nPatienten im`nPdf-Viewer anzeigen"

            ;NoDoc Text ausblenden
            	GuiControl, BO: Hide, NDoc1
            	GuiControl, BO: Hide, NDoc2

            ;Listview aktivieren, Suchfeld deaktivieren
            	GuiControl, BO: Enable, ListView1

            ;Löschen eines eventuell vorhandenen Eintrags im Suchfeld der Gui, ohne addingFiles=1 wird das gLabel ausgeführt
            	addingFiles := 1
            	GuiControl, BO: , BOEdit1, % ""

			;Listview leeren
				Gui, BO: Default                                                            		;steht das nicht vor jedem modifizierendem LV_ Befehl, funktioniert dieser nicht
				LV_Delete()

            ;von Patientenanzeige auf Anzeigen der Befund aller Patienten umschalten oder wenn alle angezeigt (ab '}else{') dann Aktualisieren der Daten des Befundordners
            	If Aktualisieren = 0
				{
                    	;ScanPool Gui auf maximale Größe setzen
                    		BO:= GetWindowInfo(hBO)
                    		WinMove, ahk_id %hBO%,,,,,% BOhMax - BO.WindowY
                        ;Ändern des Button mit einem neuen Text
                        	Aktualisieren:= 1
                        	Gui, BO: Default
                        	GuiControl, BO: , BOButton3, % "akt. Patient"
                        ;Reihen neu erstellen
                        	PatID:=""                                                                         		;keine Patienten ID - zeigt er alle Befunde
                        	NoRestart:= 0,  AlbisPatient[1]:= AlbisPatient[2]:= ""
            	}
				else
				{
                         ;ist in Albis eine Patientenakte(PatID) geöffnet, dann werden nur die Befunde zu diesem Patienten angezeigt
                        	If ( PatID:= Trim(AlbisAktuellePatID()) ) {
                        		Aktualisieren:= 0
                        		Gui, BO: Default
                        		GuiControl, BO: , BOButton3, % "alle Befunde"
                        		AlbisPatient 		:= StrSplit(AlbisCurrentPatient(), ",")
                        		AlbisPatient[1]	:= Trim(	AlbisPatient[1]	)
                    			AlbisPatient[2]	:= Trim(	AlbisPatient[2]	)
                        	}
                        ;Befundordner neu einlesen
                        	NoRestart:= 0
            	}

			;Dateien dem Listview hinzufügen
				gosub BOAddFilesToScanPool

            ;Anzahl der Reihen des Listview zählen
            	Gui, BO: Default
            	LVA_Refresh("ListView1")

            return
		}
	;}

	;{ -----------------------Minimieren------------------------------

		if Instr(A_GuiControl, "BOMin") {
            ;Timer und OnMessage Prozesse anhalten spart CPU Zeit für andere Programme
            	SetTimer, BOProcessTipHandler, Off
            	;OnMessageStatus(0)
            ;jetzt erst minimieren(verstecken))
            	minimized:=1
            	WinMinimize	, ahk_id %hBO%
            	WinHide		, ahk_id %hBO%
            	If WinExist("ahk_id " . hProgress)
                        Progress, Hide
            	If WinExist("ScanPool - Pdf Vorschau")
                        gosub BOPVGuiClose
            ;Tray Menu Fenster anzeigen aktivieren
				Menu, Tray, Enable, Fenster anzeigen
		return
		}

	;}

	;{ -------------Wahl der Pdf-Vorschaugröße-------------------
		if Instr(A_GuiControl, "BOUpDown1")
		{
				;UserDpi:= {1:"Aus", 2:"winzig", 3:"klein", 4:"mittel", 5:"groß", 6:"riesig", "Aus":0, "winzig":28, "klein":36, "mittel":48, "groß":72, "riesig":144, "Font1": 0, "Font2": 1, "Font3": 1, "Font4": 3, "Font4": 5}
				Gui, BO: Submit, NoHide

				If InStr(BOUpDown1, UD_old)
				{
                    	If InStr(UD_old, "Aus")
                    			If HighDpi
                    				GuiControl, BO:, BOUpDown1, "riesig"
                    			else
                    				GuiControl, BO:, BOUpDown1, "groß"

                    	If InStr(UD_old, "groß") and !HighDpi
                    			GuiControl, BO:, BOUpDown1, "Aus"
                    	else If InStr(UD_old, "riesig")
                    			GuiControl, BO:, BOUpDown1, "Aus"
				}

				UserDpi[Wahl]	:= BOUpDown1
				dpiShow			:= UserDpi[(BOUpDown1)]
				FontChoose		:= "Font" . BOUpDown1
				FontSize			:= UserDpi[(FontChoose)]
				dpi                    := UserDpi[(dpiShow)]
				UD_old				:= BOUpDown1

				GuiControl, BO: Text, BOPreviewDpi,  % dpiShow
				If WinExist("ScanPool - Pdf Vorschau") && Instr(dpiShow, "Aus")
                    			gosub BOPVGuiClose

				}

	;}

return
;}

BOListView:                                    	;{

	;{ --------- Doppelklick öffnet die Pdf Datei mit dem eingestellten PDF-Reader ------------------------------

		if Instr(A_GuiEvent, "DoubleClick") and !WinExist("ScanPool - arbeitet")
		{
             ;Zeile ermitteln und des Textes aus dem Feld der ersten Spalte
            	Gui, BO: Default
            	LV_GetText(BOPdfFile, nRow:= A_EventInfo)
				If InStr(last_PdfFile, BOPdfFile)
                    	return

            ;schauen ob Pdf Dokument schon angezeigt wird, sonst nicht nochmal anzeigen
            	If !WinExist(BOPdfFile " ahk_class " PDFReaderWinClass)
            	{
                        If FileExist(BefundOrdner . "`\" . BOPdfFile . ".pdf")
                    	{
                    			Running := "zeige PDF-Datei`nmit  " PDFReaderName
                        		Run, %PDFReaderFullPath% `"%BefundOrdner%\%BOPdfFile%.pdf`"
                    			WinWait, % BOPdfFile " ahk_class " PDFReaderWinClass
                    			Running:= "", last_PdfFile:= BOPdfFile
                        } else
                        		PraxTT("Die ausgewählte PDF Datei ist`nscheinbar nicht mehr vorhanden.", "4 2")

                        LV_Modify(nROw, "Check")
            	} else
                        PraxTT("Diese PDF Datei ist schon geöffnet!", "4 2")

				return
		}

	;}

	;{ --------- einfacher Klick zeigt eine Pdf Vorschau ---------------------------------------------------------------
		;return
		if !Instr(dpiShow, "Aus") && InStr(A_GuiEvent,"Normal") && !WinExist("ScanPool - arbeitet")
		{
				Gui, BO: Default
				LV_GetText(MOPdf	, A_EventInfo, 1)
				If !FileExist(BefundOrdner "\" MOPdf ".pdf")
				{
						PageSum -= ScanPoolArray("Delete", MOPdf)
						FileCount := ScanPoolArray("ValidKeys")
						SetTimer, BOSBarText, -10
						return
				}
				LV_GetText(MOPages, A_EventInfo, 2)
				SB_SetText(MOPDf, 2)
				SetTimer, BOPdfPreviewer, -20
		}

	;}

return
;}

BOGuiContextMenu:                     	;{ 	enthält sämtliche Label für das Listview-ContextMenu

	;Display the menu only for clicks inside the ListView.
		If ( Running <> "" ) or (A_GuiControl <> "Listview1")
				return

	;Info's auslesen
		Gui, BO: Default
		LV_GetText(CRowText, ctRow:= A_EventInfo, 1)
		Name := ExtractNamesFromFileName(CRowText)
		MouseGetPos, mx, my, owin, octrl, 2
		Menu, BOContextMenu, Show , % mx - 20 , % my + 10

return

BOEditItem:                                   	;{		Datei umbenennen

		OnMessage(0x200, "")
		OnMessage(0x201, "")
		SetTimer, BOProcessTipHandler, Off
		;Running := "Dokument`numbenennen`n"
		BORenamePdf(CRowText, ctRow, SPData["PDFReaderWinClass"])
		WinWaitClose, Addendum für AlbisOnWindows - Scan Pool, Ändern Sie den Dateinamen
		OnMessage(0x200, "BOWM_MOUSEMOVE")
		OnMessage(0x201, "BOWM_LBUTTONDOWN")
		SetTimer, BOProcessTipHandler, 40
		Running:= PreFile:= PostFile:= CRowText:= ctRow:= ""

return

BORenamePdf(PdfFileName, ctRow, PDFReaderWinClass) {

		static P := Object()

		WinSet, AlwaysOnTop, Off, % "ahk_id " hBOPV
		OnMessage(0x200, "")

	;Überprüfen ob die Datei per PDF-Reader geöffnet ist, wenn ja muss dieser zunächst geschlossen werden
		while WinExist(PdfFileName " ahk_class " PDFReaderWinClass)
		{
                    PdfReader_Close(PdfFileName, PDFReaderWinClass)
                    If WinExist(PdfFileName " ahk_class " PDFReaderWinClass)
                    {
                    				PraxTT(	"Das PDF-Readerfenster mit dem Befund:`n" PdfFileName "`nkonnte nicht geschlossen werden.`nBitte schließen Sie das Fenster manuell`nund wiederholen das Kommando!", "15 2")
                    				return
                    }
		}

	;InputBox Position
		If WinExist("ahk_id " hBOPV)
		{
				P := GetWindowInfo(hBOPV)
				If P.WindowW > 500
						Iw:= P.WindowW + 20
				else
						Iw:= 500 + 20
				InputBoxX:= P.WindowX + (P.WindowW//2) - Iw//2
				InputBoxY:=  (P.WindowY + P.WindowH) -118
		}
		else
		{
				P := GetWindowInfo(hBO)
				InputBoxX:= P.WindowX - 500 - 20
				InputBoxY:= (P.WindowY + P.WindowH)//2 - 125//2
		}

	;Inputbox für Namensänderung durch Nutzereingabe einblenden
		SetTimer, InputBoxNameChange, -50
		InputBox, PostFile, Addendum für AlbisOnWindows - Scan Pool, Ändern Sie den Dateinamen und drücken Sie 'Ok' oder die Enter Taste`.,, % Iw, 125, % InputBoxX , % InputBoxY,,, % PreFile:= PdfFileName
		WinWaitClose, Addendum für AlbisOnWindows - Scan Pool, Ändern Sie den Dateinamen
		If (PostFile <> "") && (PostFile <> PreFile)
		{
				;eine Abfrage, ob die Originaldatei noch besteht ist sehr wichtig, bevor man eine Datei umbenennt
					If FileExist(BefundOrdner . "\" . PreFile . ".pdf")
					{
							FileMove, % BefundOrdner . "\" . PreFile . ".pdf", % BefundOrdner . "\" . PostFile . ".pdf" , 1
							;war das Umbenennen nicht erfolgreich, geht eine Meldung raus, nach -else- muß noch die .cnt Datei umbenannt und die ListView aufgefrischt werden
							If ErrorLevel
										MsgBox,, Addendum für AlbisOnWindows - Scan Pool, % "Das Umbenennen der Datei:`n" PreFile ".pdf" "`nist fehlgeschlagen."
							else
							{
									;wie frische ich die Listview jetzt auf oder sortiere ich alles neu?
										Gui, BO: Default
										LV_Modify(ctRow, "", PostFile)
										LV_ModifyCol(1, "Sort")
										res:= ScanPoolArray("Rename", PreFile . ".pdf", PostFile . ".pdf")
										PraxTT("Die Datei wurde erfolgreich umbenannt.","2 1")
							}
					}
					 else
					{
							MsgBox,, Addendum für AlbisOnWindows - Scan Pool, % "`t`tUpps...`n`nDie Datei: " PreFile ".pdf" " konnte nicht umbenannt werden`n`, da sie anscheinend nicht mehr existiert.", 10
							return ""
					}
		}

		WinSet, AlwaysOnTop, On, % "ahk_id " hBOPV
		WinSet, AlwaysOnTop, On, % "ahk_id " hBO

return

InputBoxNameChange:
hBONC:= WinExist("Addendum für AlbisOnWindows - Scan Pool", "Ändern Sie den Dateinamen")
WinSet, AlwaysOnTop, On, % "ahk_id " hBONC
WinActivate, % "ahk_id " hBONC
ControlFocus, Edit1, % "ahk_id " hBONC
ControlSend , Edit1, {Right}, % "ahk_id " hBONC
return
}

	;}
BODeleteItem:                               	;{		Datei löschen

		Running:="Dokument löschen"
	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Nutzerabfrage
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		MsgBox,4, Addendum für AlbisOnWindows - Scan Pool, % "Wollen Sie die Datei:`n" PdfFile "`nwirklich löschen?"
		IfMsgBox, Yes
		{
					If !PdfReader_Close(CRowText, PDFReaderWinClass)
					{
								PraxTT(	"Das PDF-Readerfenster mit dem Befund:`n" CRowText "`nkonnte nicht geschlossen werden.`nBitte schließen Sie das Fenster manuell`nund wiederholen das Kommando!", "15 2")
								Running:="", CRowText:=""
								return
					}
		}
		IfMsgBox, No
		{
				Running:="", CRowText:=""
				PraxTT(	"Der Löschvorgang wurde abgebrochen!", "5 2")
				return
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; vor dem Löschen prüfen ob die Datei überhaupt noch existiert
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If FileExist(BefundOrdner . "\" . CRowText . ".pdf")
		{
			;eventuell per PDFReader angezeigten Befund vor dem Löschen der Datei schließen
				If PdfReader_Close(CRowText . ".pdf", PDFReaderWinClass)
				{
						;zu löschende Datei in das Backup-Verzeichnis verschieben
							FileMove, % BefundOrdner . "\" . CRowText . ".pdf", % BackupPath . "\" . CRowText . ".pdf"

						;hier mal anders als bei anderen Befehlen wird ErrorLevel bei Erfolg auf 0 gesetzt, deshalb heißt es jetzt !ErrorLevel
							If !ErrorLevel
							{
									Gui, BO: Default
									LV_Delete(ctRow)
									If !ErrorLevel
										PraxTT("Listenzeile " . ctRow . " ließ sich nicht löschen!", "6 2")
							}
				}
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; aus dem ScanPoolArray löschen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		PageSum -= ScanPoolArray("Delete", CRowText . ".pdf")
		FileCount := ScanPoolArray("ValidKeys")
		pdffile:=""
		SetTimer, BOSBarText, -10
	;}

		Running:="", CRowText:="", ctRow:=""

return
	;}
BOAkteOeffnen:                            	;{ 	zugehörige Patientenakte löschen

		Running:="Akte öffnen"

		Name:= ExtractNamesFromFileName(CRowText)
		PraxTT("Öffne die Akte des Patienten: `n" . Name[1], "3 3")
			sleep, 500
		If AlbisOeffneAkte(Name) = 0
		{
				PraxTT("Die Akte des Patienten`nließ sich nicht öffnen!", "3 2")
					sleep, 2000
		}

		Running:="", CRowText:="", ctRow:=""

return
;}
BOAlleOeffnen:                             	;{ 	öffnet alle Befunde des Patienten

		If !ctRow
			CRowText := "", return

		Running := "'Öffnen aller Befunde`ndes ausgewählten`nPatienten.'"

	;Feststellen des Patientennamen und Nachfrage ob es der vom Nutzer gewünschte Namen ist
		MsgBox, 0x2124, Addendum für AlbisOnWindows - ScanPool - Info, % "Sie haben folgenden Patienten ausgewählt:`n`n" Name[1] "`n`nSoll ich alle Befunde des Patienten mit dem " PDFReaderName " anzeigen?"
		IfMsgBox, No
			Running:= "", CRowText:="", return

	;Pdf Previewer ausschalten
		PrevBackup:= dpiShow, dpiShow:= "Aus"

		Gui, BO: Default
		SB_SetText("ausgewählter Patient: " Name[1], 2)

	;jetzt durch alle Zeilen gehen und nach diesem Patienten suchen
		RowName:=[], docNr:=0
		Loop, % ScanPoolArray.MaxIndex()
		{
					RowName:= ExtractNamesFromFileName(ScanPool[A_Index])
					BOPdfFile:= StrSplitEx(ScanPool[A_Index], 1, "|")
				;normale (exakte!) Suchfunktion - die nächste soll Sift3 nutzen um noch mehr Vorschläge zu machen
					If ( Instr(RowName[1], Name[3]) and Instr(RowName[1], Name[4]) ) 													;Name[3]und[4] sind einzelne Namen
					{
							If FileExist(BefundOrdner . "`\" . BOPdfFile)
							{
									If !WinExist(BOPdfFile " ahk_class " PDFReaderWinClass)
									{
											docNr ++
											Run, %PDFReaderFullPath% `"%BefundOrdner%\%BOPdfFile%`"
											If docNr = 1
											{
													PraxTT("Warte bis zu 20s auf das Öffnen`nder ersten PDF-Datei.", "20 2")
													WinWaitActive, % BOPdfFile " ahk_class " PDFReaderWinClass,, 20
													If ErrorLevel
													{
															MsgBox, 1, Addendum für Albis on Windows, % "Die Patientenakte ließ sich nicht öffnen!"
															DpiShow:= PrevBackup
															Running := "", CRowText:="", ctRow:=""
															return
													}
											}
											GuiControl, BO: Enable, BOButton2
									}
							}
					}
		}

		DpiShow:= PrevBackup
		Running := "", CRowText:="", ctRow:=""

return
;}
BOUeberSignieren:                        	;{ 	Übertragen eines oder mehrere Befund mit automatischem Signieren
		Signate:=1
;}
BOUebertragen:                             	;{ 	übertragen eines Befundes in die Akte

		Uebertragungsfehler	:= 0
		ReaderID_old			:= 0
		NPRWin_old				:= NPRWin
		SelectedPdfs				:= Object()
		SelectedPdfs				:= LV_GetSelectedText()
		PdfData					:= Object()

		If !SelectedPdfs.MaxIndex() {
				PraxTT("Es wurde kein Befund zum Übertragen ausgewählt.`nDer Vorgang wurde abgebrochen.", "8 2")
				return
		} else {
				EinOderMehrere := SelectedPdfs.MaxIndex() > 1 ? "Dokument(e)" : "Dokument"
		}

	; es darf kein PDFReader-Fenster jetzt mehr aktiv sein
		AlbisActivate(5)
		AlbisIsBlocked(AlbisWinID(), 2)

	; verschiedene Hinweistexte
		If (Signate = 0)
			Running := EinOderMehrere " nach Albis übertragen"
		else if (Signate = 1)
			Running := EinOderMehrere " nach Albis übertragen`nund automatisch signieren"

	; BOProcessTip Zeit zum Abfangen geben
		Sleep, 300

	; selektierte Dateien nach und nach übertragen
		For key, Selected in SelectedPdfs
		{
				; leeren das PdfData-Objektes
					PraxTT("übertrage die Pdf-Datei:`n" Selected.Pdf, "20 2")
					PdfData.KarteikartenText	:= ""
					PdfData.PdfFullPath			:= ""
					PdfData.ReaderID				:= ""
					PdfData.AlbisPdfPath			:= ""

				; eventuell noch geöffnetes Dokumentfenster schließen da der FoxitReader ein kopieren sonst verhindert
					PdfReader_Close(Selected.Pdf, PDFReaderWinClass)

				; Pdf-Dokument wird übertragen
					retcode:= AlbisPdfUebertragen(Selected.Pdf ".pdf", Selected.Row)
					If (retcode > 1)
							continue

				; löschen des Eintrages aus dem ScanPool[] Array
					PageSum -= ScanPoolArray("Delete", Selected.Pdf ".pdf")
					PraxTT("Die interne Pdf-Liste wurde aufgefrischt.", "3 2")

				; Daten zum Speicherort der PDF-Datei sichern
					PdfBezeichnung:= StrReplace(PdfData.KarteikartenText, "(" Name[4] ", ", "")
					PdfBezeichnung:= StrReplace(PdfBezeichnung, Name[3] ") - ", "")

				; löschen des Pdf-Eintrages aus dem Listview
					Gui, BO: Default
					LV_Delete(Selected.Row)
					WinSet, Redraw,, % "ahk_id " hBO

				; Statusbar auffrischen
					SetGuiStatusBar("BO", LV_GetCount(), "Pdf Dateien mit " . PageSum . " Seiten")

				; Signieren der Datei falls gewünscht
				/*
					PraxTT("Titel: " WinGetTitle(ReaderID) "`nID:    " ReaderID, "4 2")
					If (Signate = 1)
					{
							FoxitReader_SignaturSetzen(ReaderID, PDFReaderWinClass)
							;	SetTimer, SignatureHelper, 300
							WinWait		 , % "Speichern unter ahk_class #32770",, 60
							WinWaitClose, % "Speichern unter ahk_class #32770",, 60
					}
					else
					{
							If InStr(PDFReaderName, "FoxitReader")
								FoxitInvoke("SaveAs", PdfData["ReaderID"])
							WinWait		 , % "Speichern unter ahk_class #32770",, 60
							WinWaitClose, % "Speichern unter ahk_class #32770",, 60
					}

					FileAppend, % "`n" ( Name[5] = "" ? (-------) : Name[5] ) ";" Name[4] ";" Name[3] ";" PdfBezeichnung ";" PdfData.AlbisPdfPath, % AddendumDBPath "\PdfPfade.csv"
				*/
		}

	; alles zurücksetzen - fertig
		AlbisSetzeProgrammDatum(A_DD "." A_MM "." A_YYYY)
		WinActivate   	, % "ahk_class " PDFReaderWinClass
		WinWaitActive	, % "ahk_class " PDFReaderWinClass,, 3
		Running:= CRowText:= ctRow:="", Signate := 0

return
;}
BOMetaDatenAnzeigen:                   	;{		zeigt alle Metadaten der ausgewählten Pdf Datei an
	MetaDatenAnzeigen(BefundOrdner "\" CRowText ".pdf")
return
;}
BOContentCheck:                         	;{		experimentell

	;PdfFuzzyContents(Name, BefundOrdner . "`\" . CRowText)
	PdfUntersuchen(CRowText . ".pdf", 1)

return

PdfUntersuchen(PdfFileName, MaxPages) {

		PdfText:= PdfToText(BefundOrdner "\" PdfFileName, MaxPages)
		clipboard:= PdfText
		MsgBox, % PdfText
}

;}
BOGuiCheckClose:                       	;{ 	andere Möglichkeit das Skript zu beenden
		MsgBox,4, Addendum für AlbisOnWindows - Scan Pool, % "Wollen Sie ScanPool wirklich beenden?"
		IfMsgBox, Yes
				gosub BOGuiClose
		CRowText:="", ctRow:=""
return
;}
BONewIndex:                                	;{ 	löscht den Index und erstellt einen neuen

	;ScanPool macht gerade was anderes dann Abbruch
		If Running <> "" or addingfiles
			return

		Running := "Erstelle neuen Index"

	;PdfPreview-Fenster schliessen
		If WinExist("ScanPool - Pdf Vorschau")
			gosub BOPVGuiClose

	;Variablen setzen
		NoRestart:= 0, AlbisPatient[1]:= "", AlbisPatient[2]:="", PatID:=""

	;das Listview leeren
		Gui, BO:default
        LV_Delete()

		FileDelete, % BefundOrdner "\pdfIndex.txt"

	;Dateien neu indizieren
		NoRestart 	:= 1									;mit NoRestart = 1 wird der komplette Index neu erstellt
		Running	:= ""
		gosub BOAddFilesToScanPool

		CRowText:= ctRow:=""
		Gui, BO:default

return
	;}

;}

BOProcessTipHandler:                   	;{  	verschiebt das ScanPool Gui falls die PraxomatGui geschlossen oder geöffnet wird

	;zeigt einen Hinweis das das Skript im Moment arbeitet, damit der Nutzer nicht versehentlich störend eingreift
		If Running And !WinExist("ScanPool - arbeitet")
		{
				ButtonsOnOff("BOButton3|BoButton2|BOButton1|BOUpDown1|BOEdit1", "BO", "Disable")
				;OnMessageStatus(0)

				If WinExist("ScanPool - Pdf Vorschau")
						gosub BOPVGuiClose

				BOGuiProcessTip()
				Gui, BO: Default

				Procedure:= StrReplace(Running, "`n", " ")
				Procedure:= StrReplace(Procedure, "- ", "")
				RQIndex:= RQueue.Push({"Procedure" : (Procedure), "starttime": (A_TickCount), "endtime": 0, "Duration": 0})
				If DebugFunctions
				{
						RQList.= "'" RQueue[RQIndex].Procedure "' wurde gestartet`n"
						If StrSplit(RQList, "`n").MaxIndex() > 5
								RQList:= LineDelete(RQList, 1)
						ToolTip, % RQList, 600, 5, 17
				}
		}
		  else if WinExist("ScanPool - arbeitet") and !InStr(WhatIsRunning, Running)
		{
				RQueue[RQIndex].EndTime 	:= A_TickCount
				RQueue[RQIndex].Duration	:= Round((RQueue[RQIndex].endtime - RQueue[RQIndex].starttime)/1000, 3)
				Procedure:= StrReplace(Running, "`n", " ")
				Procedure:= StrReplace(Procedure, "- ", "")
				RQIndex:= RQueue.Push({"Procedure" : (Procedure), "starttime": (A_TickCount), "endtime": 0, "Duration": 0})
				WhatIsRunning := Running

				ButtonsOnOff("BOButton3|BoButton2|BOButton1|BOUpDown1|BOEdit1", "BO", "Disable")
				;OnMessageStatus(0)

				GuiControl, BOPT: , PTText3, % Running
				Gui, BO: Default
		}
		  else if WinExist("ScanPool - arbeitet") and ( Running = "" )
		{
				RQueue[RQIndex].EndTime 	:= A_TickCount
				RQueue[RQIndex].Duration	:= Round((RQueue[RQIndex].endtime - RQueue[RQIndex].starttime)/1000, 3)
				If DebugFunctions
				{
						RQList.= "'" RQueue[RQIndex].Procedure "' lief für " RQueue[RQIndex].Duration " Sekunden." "`n"
						If StrSplit(RQList, "`n").MaxIndex() > 5
								RQList:= LineDelete(RQList, 1)
						ToolTip, % RQList, 600, 5, 17
				}

				ButtonsOnOff("BOButton3|BoButton2|BOButton1|BOUpDown1|BOEdit1", "BO", "Enable")
				;OnMessageStatus(1)

				If !WinExist("ahk_class TV_ControlWin") and !IsRemoteSession()
								AnimateWindow(hBOPT, 150, FADE_HIDE)

				Gui, BOPT: Destroy
				Gui, BO: Default
				hBOPT := ""

		}


return
;}

BOGuiSize:                                   	;{  	Routinen für angepasste Größenänderung der ScanPool Gui Elemente bei Usereingriffen

		Critical, Off			;Trick damit GuiSize sofort aufgerufen wird
		if A_EventInfo = 1  ; Das Fenster wurde minimiert.  Keine Aktion notwendig.
		return
		Critical
		;OnMessageStatus(0)
		BOLv1H	:= A_GuiHeight - BOProg2H - BOLv1Y - 20
		BOText1X	:= A_GuiWidth  - BOText1W - 10	, BOText1Y	:= BOLv1H + 35
		BOEdit1X	:= A_GuiWidth  - BOEdit1W - 10	, BOEdit1Y	:= BOText1Y + BOText1H + 2
		GuiControl, Move		   , ListView1	, % "w" A_GuiWidth " h" BOLv1H
		GuiControl, MoveDraw, Text1		, % "x" BOText1X " y" BOText1Y
		GuiControl, MoveDraw, BOProg2	, % "y" BOLv1Y + BOLv1H - 1
		GuiControl, MoveDraw, Edit1		, % "x" BOEdit1X " y" BOEdit1Y
		GuiControl, MoveDraw, WTitle1		, % "y" BOText1Y
		GuiControl, MoveDraw, WTitle2		, % "y" BOEdit1Y - 5
		GuiControlGet, WTitle2, BO: Pos
		GuiControl, MoveDraw, WTitle3, % " x" (WTitle2X + WTitle2W - 8) " y" (WTitle2Y + WTitle2H - 17)
		Critical, Off
		;SetTimer, BORedraw, % "-" A_GuiHeight//32
return

BORedraw:
		WinMove	, ahk_id %hBO%					;,,,, 			;%BOx%, %BOy%
		WinSet		, Redraw,, ahk_id %hBO%
		;OnMessageStatus(1)
return
;}

BOSBarText:                                  	;{ 	Timerlabel für die Anzeige von Statustexten (im Moment nur einer)

	Gui, BO: Default
	SB_SetText(FileCount, 1)
	SB_SetText(" Pdf Dateien mit " . PageSum . " Seiten", 2)

return
;}

BOAddFilesToScanPool:                 	;{ 	Auslesen des ScanPool Ordners mit Filterung und Übergabe der Daten an ein ListView

/*	Beschreibung

		Problem:							Befunde können nicht gefunden werden
		Problembeschreibung:		die Dateinamen der PDF Dateien enthalten Name, Vorname des Patienten Untersuchungsart und Untersuchungsdatum
												die Patientennamen werden leider oft nicht korrekt eingegeben:
												1. Schreibfehler
												2. verkürzte Schreibweise bei Doppelname Hans-Joachim nur Hans
		kein Problem machte: 		vertauschen der Position von Name, Vorname , die If-Abfrage trennt die Namen und sucht unabhängig der Stellung

		Lösungsmöglichkeiten:		1.Autokorrektur der Patientenname bei Erstellung der Dateien - in Arbeit - aber ohne sicheren Erfolg
												2.flexiblere Handhabung des Patientennamen im Nachhinein
												3.Patientennamen über Datenbankabfrage ermitteln, habe bisher keinen Datenbankzugriff auf die aktuellen Datenbanken hinbekommen, Datenbankdatei musste immer erst kopiert werden

	*/

	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen zurücksetzen
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		docIndex    	:= 0
		Switch          	:= 1			; flag für Hintergrundfarbwechsel bei anderem Namen
		NoDoc         	:= 0			; wenn kein passenden Dokumente gefunden wurden = 0
		last_row		:= []			; speichert hier Vor- und Nachname der letzten Listview-Zeile
	;}

	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; für schnelleres Anzeigen - Controls vorübergehend abschalten
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		Running :="bereite Datei-`nanzeige vor"
		Sleep 300
		GuiControl, BO: -Redraw	, ListView1
		Gui, BO: Default
		GuiControl, BO: Hide	, NDoc1
		GuiControl, BO: Hide	, NDoc2
	;}

	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; der Ordner wird neu eingelesen wenn NORestart=0 ist
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		if NORestart
		{
				NoRestart	:= 0
				FileCount	:= MaxFiles:= ReadPdfIndex(BefundOrdner . "\PdfIndex.txt")
		}

	;}

	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; fügt dem Listview die Daten zu, 1. zeigt alle Dateien wenn keine Patientenakte geöffnet ist - oder - 2. zeigt nur die Befunde des geöffneten Patienten an
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		If !InStr(AlbisGetActiveWindowType(), "Patientenakte")   or   (Aktualisieren = 1)
		{
					For key, val in ScanPool					;wenn keine PatID vorhanden ist, dann ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)
					{
									if (val = "")
										continue

								;Dokumenten Index erhöhen
									docIndex ++
									col:= StrSplit(val, "|")
									LV_Add("", StrReplace(col[1], ".pdf"), col[3])				;1=pdfname, pdfgröße, pdfseiten

								;stimmen die Namen der vorherigen Zeile und der aktuellen Zeile überein - erhalten diese Listviewreihen diesselbe Hintergrundfarbe
									Name 		:= ExtractNamesFromFileName(col[1])
									If ( StrDiff(last_row[1] . last_row[2], Name[3] . Name[4]) > 0.15 )   &&   ( StrDiff(last_row[1] . last_row[2], Name[4] . Name[3]) > 0.15 )
											Switch:= -Switch

								;jetzt je nach Switch-Status die Hintergrundfarbe wählen
									If Switch = 1
										LVA_SetCell("ListView1", (docIndex), 0, "0x" . LVBGColor1, "0x000000")
									else
										LVA_SetCell("ListView1", (docIndex), 0, "0x" . LVBGColor2, "0x000000")

								;vorhergehende Zeile speichern
									last_row[1]:= Name[3], last_row[2]:= Name[4]
					}
		}
 		  else if (Aktualisieren = 0)
		{
					AlbisPatient 	:= VorUndNachname(AlbisCurrentPatient())
					For key, val in ScanPool
					{
									if (val = "")
										continue
								;Sift Suche - bei Übereinstimmung von größer 85% könnte die Datei dazu gehören, Schreibfehler im Namen
									col			:= StrSplit(val, "|")
									Name 		:= ExtractNamesFromFileName(col[1])
									If ( StrDiff(AlbisPatient[1] . AlbisPatient[2], Name[3] . Name[4]) <= 0.15 )   ||   ( StrDiff(AlbisPatient[1] . AlbisPatient[2], Name[4] . Name[3]) <= 0.15 )
									{
											docIndex++
											LV_Add("", StrReplace(col[1], ".pdf"), col[3])				;1=pdfname, pdfgröße, pdfseiten
									}
					}
		}

	;}

	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; keine oder Dateien zum Patienten gefunden
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		If (docIndex = 0)
		{
				;Darstellung eines Splashtext-Fenster in der Mitte der ScanPool Gui - wenn kein Befund vorhanden ist, als extra Thread eingerichtet - Progress und Splashimage behindern sich, Progress ging nicht zu schliessen
					NoDoc := 1
				;Anzeige keine Befunde für diesen Patienten vorhanden
					GuiControl,					, NDoc2, % AlbisPatient[1] ", " AlbisPatient[2]
					GuiControl, BO: Hide	, NDoc1
					GuiControl, BO: Show	, NDoc1
					GuiControl, BO: Hide	, NDoc2
					GuiControl, BO: Show	, NDoc2
		}
		 else
		{
					GuiControl, BO: Enable		, ListView1
					GuiControl, BO: +Redraw	, ListView1
		}
	;}

	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Listview auffrischen, Status, Index anzeigen und anderes
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
		Running := ""

		SB_SetText(docIndex, 1)
		If PageSum = 0
				SB_SetText("Pdf Dateien", 2)
		else
				SB_SetText("Pdf Dateien mit " . PageSum . " Seiten", 2)

		NoRestart  	:= 0																										;flag um Dateien nicht neu einzulesen nach Anzeigemoduswechsel
	;}

return

;}

BOCheckPDFReader:                     	;{ 	schaut nach ob ein PDFReader Prozeß existiert, setzt ein Häkchen vor angezeigte Dokumente und managed überhaupt das ScanPool Gui

	SetTimer, BOCheckPdfReader, Off
	IfWinExist, ScanPool - Info
						return

 ;Label wird per Hook aktiviert - also muss das hier nicht mehr sein
	Process, Exist, %PDFReaderExe%
	If !ErrorLevel
	{
		pdfshow :=0
		Gui, BO: Default
		LV_Modify(0, "-Check")
		GuiControl, BO: Disable, BOButton2
		return
	}

	iPdf:=0, pdfList:=""

	WinGet, WinList, List, ahk_class classFoxitReader
	Loop % WinList
	{
				WinGetTitle, WT, % "ahk_id " WinList%A_Index%
				pdfList.= WT . "|"
				WinList%A_Index%:=""
	}

	Gui, BO: Default

	Loop % LV_GetCount()
	{
				LV_GetText(RetrievedText, A_Index)
				If Instr(pdfList, RetrievedText)
				{
					LV_Modify(A_Index, "Check")
					iPdf++
					continue
				}
				LV_Modify(A_Index, "-Check")
	}

	pdfshow:= iPDF=0 ? 0 : 1

	if !pdfshow
		GuiControl, BO: Disable, BOButton2
	else
		GuiControl, BO: Enable	, BOButton2

	iPDF:= pdfList:= RetrievedText:=""

return
;}

BOMarkSigned:                             	;{  	zeigt farbig an welche Dateien signiert sind

	; Dauer der einzelnen Befehle im Labels sind nicht berechenbar, deshalb wird der Timer ausgestellt, am Ende wieder angestellt
		SetTimer, BOMarkSigned, Off
		;return
		If !IsObject(MarkSigned)
		{
				MarkSigned 	:= Object()
				MarkSigned.Index					:= 0
				MarkSigned.SignedPositive    	:= Array()
				MarkSigned.SignedPositive		:= ScanPoolArray("Signed")
				MarkSigned.SignedNegative	:= Array()
				MarkSigned.SignedNegative  := ScanPoolArray("NotSigned", MarkSigned["SignedPositive"])
		}

	; Listview Zähler
		MarkSigned.Index ++
		;MsgBox, % "MaxIndex: " MarkSigned["SignedNegative"].MaxIndex() "`nNotSigned[1]: " MarkSigned["SignedNegative"][1] "`nIndex: " MarkSigned["SignedPositiv"].MaxIndex()
		;ObjTree(ScanPool)

	; Ende ist erreicht, dann hier raus
		If MarkSigned.Index > docIndex
				return





		LV_GetText(MSrow, MSPdfIdx, 1)


return
;}

BOGuiClose:                                 	;{ 	schließen noch evtl. geöffneter Befunde im PDFReader, Beenden der OnMessage-Funktionen, Beenden der Hooks, Freigeben von Speicher
BOGuiEscape:

		OnExit

		PraxTT("Das Skript wird beendet ......`nSpeichere noch Daten und Einstellungen....", "8 2")

	; Fensterposition speichern
		MonitorPos 	:= GetMonitorIndexFromWindow(hBO)
		ScPwin	    	:= GetWindowSpot(hBO)
		IniWrite, % ScPwin.X "|" ScPwin.Y "|" MonitorPos, % AddendumDir "\Addendum.ini", % "ScanPool", % "BOXYPos-" CompName

	; die Daten des ScanPool-Ordners sichern
		ScanPoolArray("Save")

	; Einstellungen sicheren
		Gui, BO: Submit, NoHide
		IniWrite, % BOUpDown1     		, % AddendumDir "\Addendum.ini", % "ScanPool", % "DpiPdfVorschau"
		IniWrite, % SPData.Signatures 	, % AddendumDir "\Addendum.ini", % "ScanPool", % "Signatures"

	;Timer beenden, OnMessage beenden
		SetTimer, BOCheckPDFReader	, Off
		SetTimer, BOProcessTipHandler	, Off
		SetTimer, BOMarkSigned	    	, Off
		OnMessage(0x200, "")								;WM_MouseOver wird deregistriert
		OnMessage(0x201, "")
		OnMessage(0x4E	, "")
		OnMessage(0x4A	, "")

	;RemoveRamDrive(RamDisk, RamDiskPath)

	;MsgBox, 4, Addendum für AlbisOnWindows, Sollen noch alle Fenster im %PDFReader% mit`nnoch geöffneten Patientenbefunden geschlossen werden?
	;IfMsgBox, Yes
	;	%PDFReader%_CloseAllPatientPDF()						;dynamische Implementation eines Funktionsaufrufes um hier keinen neuen Programmcode für ein anderes PDF Anzeige Tool erstellen zu müssen

	;temp pictures löschen
		Loop
		{
				picname:= % TempPath . "\sppreview-" . SubStr("00000" . A_Index, -6) . ".png"
				If !FileExist(picname)
					break
				FileDelete, % picname
		}

	;keine Animation bei eingeschalten TeamViewer
    	If !WinExist("ahk_class TV_ControlWin") and !IsRemoteSession()
			AnimateWindow(hBO, Duration, FADE_HIDE)
		Gui, BO: Destroy
		Gui, BOPV: Destroy

	;unhook
		If hWinEventHook1
				UnhookWinEvent(hWinEventHook1, HookProcAdr1)
		If hWinEventHook2
				UnhookWinEvent(hWinEventHook2, HookProcAdr2)

	;eventuell noch vorhandene Console freigeben
		If WinExist("ahk_pid " ExecPID) {
				DllCall("FreeConsole")
				Process, Close, %ExecPID%
		}

ExitApp

ScanPoolEndsHere(ExitReason, ExitCode) {
	;ListLines
	;Pause
	;FileAppend, % A_Now " ,ER: " ExitReason " , EC: " ExitCode "`n", %A_ScriptDir%\ExitGruende.txt
	gosub BOGuiClose
	return 1
}
;}
;}

;{ 7. PDF Suche | Pdf Preview

BOFTextSearch:                             	;{ 	--> funktionslos- PDF Volltext Namenssuche - noch nicht anfangen

	return
;}

BOIncSearch:                                	;{ 	inkrementelles Suchen von Patientenbefunden, nützlich wenn alle Dateien im Befundordner angezeigt werden

	If addingfiles = 1 or Running != ""
		return

	;Controls auslesen
	Gui, BO: Default
	GuiControlGet, Edit1
	Edit1:= StrReplace(Edit1, ",")
	LV_Delete()
	GuiControl, BO: -Redraw, ListView1

	;Suchfunktion
	For each, row in ScanPool
	{
			colums:= StrSplit(row, "|")
			colums[1]:= StrReplace(colums[1], ".pdf")
			If !( Edit1 = "" ) && InStr(colums[1], Edit1)
						LV_Add("", colums[1], colums[3])
			else
						LV_Add("", colums[1], colums[3])

	}

	;Statusanzeige
	Gui, BO: Default
	GuiControl, BO: +Redraw, Listview1
	SB_SetText(LV_GetCount(), 1)
	SB_SetText(" von " . FileCount . " Dateien", 2)
	GuiControl, BO: +Redraw, ListView1

	return
;}

BOIncFuzzySearch:                           	;{ 	inkrementelles Suchen von Patientenbefunden, nützlich wenn alle Dateien im Befundordner angezeigt werden

		If (addingfiles = 1) or (Running <> "")
				return

		GuiControlGet, query,, BO: BOEdit1
		If ( query = "" )
		{
				Gui, BO: Default
				LV_Delete()
				GuiControl, BO: -Redraw, ListView1
				gosub BOAddFilesToScanPool
				return
		}
		else
		{
				Gui, BO: Default
				LV_Delete()
				GuiControl, BO: -Redraw, ListView1
		}

	;Suchfunktion
		For each, row in ScanPool
		{
					colums:= StrSplit(row, "|")
					col:= PrepareString(colums[1])

				;der Loop vergleicht jedes einzelne Wort im query gegen jedes einzelne Wort im PDF Namen,
					query			:= StrReplace(query, ",", " ")
					query			:= StrSplit(query, " ")
					wordspassed := 0

					Loop, % col.MaxIndex()
					{
							BOidx:= A_Index

							Loop, % query.MaxIndex()
								If StrDiff(col[BOidx], query[A_Index]) <= 0.2
										wordspassed ++, continue

							If ( wordspassed = query.MaxIndex() )			;Beispiel: 2 Suchwörter sollten zwei Treffer erzielen
							{
									LV_Add("", StrReplace(colums[1], ".pdf", ""), colums[3])
									continue
							}

								wordspassed:=0
					}
			}

			For key in query
					  query[key]	:=""
			For key in Columns
				Columns[key]	:=""

	;Statusanzeige
		Gui, BO: Default
		GuiControl, BO: +Redraw, Listview1
		SB_SetText(LV_GetCount(), 1)
		SB_SetText(" von " . FileCount . " Dateien", 2)

return

PrepareString(str) {
	str:= StrReplace(str, ".pdf")
	str:= StrReplace(str, ",", "#")
	str:= StrReplace(str, ";", "#")
	str:= StrReplace(str, ":", "#")
	str:= StrReplace(str, " ", "#")
	str:= StrReplace(str, "###", "#")
	str:= StrReplace(str, "##", "#")
return StrSplit(str, "#")
}

;}

BOPdfPreviewer:                            	;{ 	zeigt das von pdftopng gerenderte Bild in einem Guifenster an, nichts besonderes

       	;{ Initialisierung der Gui

		If InStr(last_MOPdf, MOPdf) || !MOPdf || StrLen(MOPdf)<3
		{
				SB_SetText("nichts ausgewählt!", 2)
				gosub BOPVGuiClose
				return
		}

		PreviewerWaiting := 0
		Running     	:= "erstelle neue PDF-Vorschau"
		prevPdf			:= BefundOrdner . "\" . MOPdf . ".pdf"
		last_MOPdf	:= MOPdf

	;-: Startparameter
		page := 1, arc := 90, BOPV_isTurned :=0, BOPVFirst :=1
		PreviewFilename:= % TempPath . "\PdfPreview-" . SubStr("00000" . page, -6) . ".png"
		If HighDpi
				PVZoom:= 144/dpi
		else
				PVZoom:= 96/dpi + 0.1 				;10% kleiner noch machen

	;-: Anzeigen der Vorschauoberfläche
		gosub BOPVPaintPreviewWindow
		SetTimer, BOPVCheckMouseOver, 250

	;-: Gui, Default zurück an BO
		Gui, BO: Default

		PreviewerWaiting := 1
		Running := ""

return
;}

BOPVPaintPreviewWindow:               	;{

		Gui, BOPV: New					, -Resize +ToolWindow -Caption -DPIScale +Parent%hBO% HWndhBOPV	;+Border
		Gui, BOPV: Margin				, % BOxMrg, % BOyMrg
		Gui, BOPV: Color					, 202842
		Gui, BOPV: Font					, % BOFontOptions , % BOFont

		;theBitmap := LoadPicture(PreviewFilename, prevPdf, dpi, 1, "hBitmap", 0)

    ;-: Fenstertitel
		Gui, BOPV: Font					, % "s" (14 // PVZoom) " cWhite"
		Gui, BOPV: Add, Text			, % " xm y2 vPVText1 BackgroundTrans", Addendum für AlbisOnWindows																		    				;vPVText1 - Addendum für...
		GuiControlGet	, PVCtrl_    	, BOPV: Pos, PVText1
		Gui, BOPV: Font					, % "s" (34 // PVZoom) " cFFFFFF q5 Bold", % BOFont
		Gui, BOPV: Add, Text			, % " xm y" PVCtrl_H - 7 " BackgroundTrans Center vPVText2", % "ScanPool"							    				              				    			;vPVText2 - ScanPool // ?? PVCtrl_H + 5//PVZoom
		GuiControlGet	, PVCtrl_ 		, BOPV: Pos, PVText2

		;BOPV_Turn ? Gdip_RotateBitmap( LoadPicture(picname, "pBitmap", 1) , 90, 1) : LoadPicture(picname, "hBitmap", 1)

	;-: Preview Picture
		theBitmap := BOPV_Turn ? Gdip_RotateBitmap( LoadPicture(PreviewFilename, prevPdf, dpi, page, "pBitmap", 0) , arc, 1) : LoadPicture(PreviewFilename, prevPdf, dpi, 1, "hBitmap", 0)
		Gui, BOPV: Add, Picture		, % "xm y" (PVCtrl_Y + PVCtrl_H + 5) " vPVPic1 Section 0xE BackGroundTrans", % "HBITMAP: " theBitmap							    				;vPVPic1
		GuiControlGet	, PVPic1		, BOPV: Pos

	;-: Bild Unterschrift
		Gui, BOPV: Font					, % "s" (16 // PVZoom) " cWhite"
		Gui, BOPV: Add, Text			, % "xs y+" 10  " Center 	vBOUse", % "Mausrad blättern`, rechte Maustaste drehen.   akt.Ausrichtung: " arc "°"					;vBOUse
		GuiControlGet	, BOUse		, BOPV: Pos
		Gui, BOPV: Add, Text			, % "x" PVPic1W - 200 " y" BOUseY " vPVPages"	, % page "`/" MOPages																			    				;vPVPages
		GuiControlGet	, PVPages  	, BOPV: Pos

		GuiControlGet	, PVPages  	, BOPV: Pos
		If !PVPic1X
			PVPic1X:=10, PVPic1Y:=30, PVPic1W:=400, PVPic1X:=200

	;-: Gui unsichtbar anzeigen
		Gui, BOPV: Show, AutoSize NA Hide, ScanPool - Pdf Vorschau

	;-: momentane Fenstergröße ermitteln, Position des Fensters in Bezug zur BO Gui berechnen
		BO			:= GetWindowInfo(hBO)
		BOPV		:= GetWindowInfo(hBOPV)
		BOPVx		:= Floor( BO.WindowX 	- BOPV.WindowW - 10 )
		BOPVy		:= Floor( BO.WindowY + (BO.WindowH//2 - PV.WindowH//2) )
		BOPVy		:= BOPVy < 1 ? BOy : BOPVy

	;-: Fenster bewegen, ein Control versetzen
		SetWindowPos(hBOPV, BOPVx, BOPVy, BOPV.WindowW, BOPV.WindowH)

	;-: PdfVorschau Schrift ←↑→ ▶◀◁▷►▻◄◅
		Gui					, BOPV: Font							, % "s" (22 // PVZoom) " cFFFFFF q5 Bold", % BOFont
		Gui					, BOPV: Add	 , Text				, % "x50 y0 vPVText3 BackgroundTrans Center HWNDhPVText3", PDF Vorschau																	;vText3 - PdfVorschau
		GuiControlGet	, PVCtrl_		 , BOPV: Pos		, PVText3
		GuiControl		, BOPV: Move, PVText3			, % "x" (PVPic1W - PVCtrl_W + 5)

		If PVPic1H > 20
		{
			GuiControl		, BOPV: Move, PVPages			, % "x" (PVPic1W - PVPagesW - 10//PVZoom) " y" (PVPic1Y + PVPic1H + 10)
			GuiControl		, BOPV: Move, BOUse			, % "y" (PVPic1Y + PVPic1H + 10)
		}
		else
		{
			BOPV := GetWindowInfo(BOPV)
			GuiControl		, BOPV: Move, PVPages			, % "x" (BOPV.WindowW - PVPagesW - 10//PVZoom) " y" (BOPV.WindowY - PVPagesH - 10)
			GuiControl		, BOPV: Move, BOUse			, % "y" (BOPV.WindowH - BOUseH - 10)
		}

	;-: Büroklammer
		Gui					, BOPV: Add	 , Picture			, % "x" ( PVPic1W - 100//PVZoom ) " y" (PVPic1Y - 22//PVZoom) " w" ( 50//PVZoom ) " h" ( 100//PVZoom ) " vLNice1 BackgroundTrans", % A_ScriptDir . "\assets\paperclip50x.png"

	;-: verschieben vom ScanPool Text
		GuiControlGet	, PVCtrl_ 		 , BOPV: Pos, PVText2			;Position genau oberhalb des Bildes jetzt
		GuiControl		, BOPV: Move, PVText2			, % "y" (PVPic1Y - PVCtrl_H - 1)

	;-: Dateibezeichnung einblenden
		Gui					, BOPV: Font							, % "s" (16 // PVZoom) " cFFFF00 q5 Normal", % BOFont
		Gui, BOPV: Add, Text	, % "x" (PVCtrl_X + PVCtrl_W + 10) " y2 vPVText4" , % "-- " StrReplace(StrReplace(prevPdf, BefundOrdner . "\", ""), ".pdf") " --"
		GuiControlGet	, PVCtrl_ 		, BOPV: Pos, PVText4
		GuiControl		, BOPV: Move, PVText4			, % "y" (PVPic1Y - PVCtrl_H - 3)

	;-: Seite vor und zurück
		BOPV:= GetWindowInfo(hBOPV)
		Gui, BOPV: Font, % "s" 12 " cWhite"

	;-: Fenster neu zeichnen
		WinSet, Redraw,, ahk_id %hBOPV%

	;-: Fenster zeigen
		Gui, BOPV: Show, NA, ScanPool - Pdf Vorschau
		WinActivate, % "ahk_id " hBO

		PVText1X:=PVText1Y:=PVText1W:=PVText1H:=PVText2X:=PVText2Y:=PVText2W:=PVText2H:=PVText3X:=PVText3Y:=PVText3W:=PVText3H:=BOPVx:=BOPVy:=PVPic1X:=PVPic1Y:=PVPic1W:=PVPic1H:= ""
		PVSeiteX:=PVSeiteY:=PVSeiteW:=PVSeiteH:=BOPVPagesX:=BOPVPagesY:=BOPVPagesW:=BOPVPagesH:= ""

return
;}

BOPVPageTurnLeft:                                         	;{
BOPV_Turn:= -1
;}
BOPVPageTurnRight:                                      	;{
If !BOPV_turn
	BOPV_turn:= 1
;}
BOPVPageTurn:                                             	;{

	If !MouseIsOver(hBOPV)  or !MouseIsOver(hBO)
				return

    ;arc := (arc = 270) ? 0 : arc + (BOPV_Turn * 90)
    arc += BOPV_Turn * 90
	arc  := (arc = 360) ? 0 : arc
	arc  := (arc < 0) ? 270 : arc

	gosub BOPVPaintPreviewWindow
	BOPV_Turn	:= 0

return
;}

BOPVPageDown:                                           	;{
		PageDown:= 1
BOPVPageUp:
		If !MouseIsOver(hBOPV) or !MouseIsOver(hBO)
				PageDown:=0,	return

		If !PageDown
				ttSymbol:= "▷", page:= (page > MOPages ) ? 1        	  : page	+	1
		else
				ttSymbol:= "◁", page:= (page < 1    		  ) ? MOPages : page -	1

		page:= page > MOPages	? 1           	: page
		page:= page < 1         	? MOPages	: page

		ToolTip, % ttSymbol , % MOx +10 , % MOy + 20, 9
		SetTimer, BOPVSetNewPagePreview	, -5
		SetTimer, TT9Off							, -1000

		PageDown:= 0

return
;}

BOPVSetNewPagePreview:                             	;{

	theBitmap := LoadPicture(TempPath . "\PdfPreview-" . SubStr("00000" . page, -6) . ".png", prevPdf, dpi, page, "hBitmap", 0)

	GuiControl		, BOPV: 		 , PVPic1			, % "HBITMAP: " theBitmap
	GuiControl		, BOPV: 		 , PVPages			, % page "/" MOPages

	GuiControlGet	, PVPic1		 , BOPV: Pos
	GuiControlGet	, PVPages	 	 , BOPV: Pos

	If PVPic1H > 20
	{
			GuiControl		, BOPV: Move, PVPages			, % "x" (PVPic1W - PVPagesW - 10//PVZoom) " y" (PVPic1Y + PVPic1H + 10)
			GuiControl		, BOPV: Move, BOUse			, % "y" (PVPic1Y + PVPic1H + 10)
	}
	else
	{
			BOPV := GetWindowInfo(BOPV)
			GuiControl		, BOPV: Move, PVPages			, % "x" (BOPV.WindowW - PVPagesW - 10//PVZoom) " y" (BOPV.WindowY - PVPagesH - 10)
			GuiControl		, BOPV: Move, BOUse			, % "y" (BOPV.WindowH - BOUseH - 10)
	}
	WinSet, Redraw,, ahk_id %hBOPV%

return
;}

BOPVCheckMouseOver:                                 	;{	weil OnMessageMouseMove nicht absolut zuverlässig ist

	MouseGetPos,,, hWinOver
	if hWinOver not in %hBO%,%hBOPV%
			gosub BOPVGuiClose

return
;}

BOPVGuiClose:	                                            	;{
BOPVGuiEscape:

	SetTimer, BOPVCheckMouseOver, Off

	If WinExist("ScanPool - Pdf Vorschau")
	{
			PreviewerWaiting := 1
			If !WinExist("ahk_class TV_ControlWin") and !IsRemoteSession()
								AnimateWindow(hBOPV, 150, FADE_HIDE)
			Gui, BOPV: Destroy
		; temporäre Dateien löschen, die xpdf Tools überschreiben vorhandene Dateien nicht!
			tempfiles:= ReadDir(TempPath, "png")
			Loop, Parse, tempfiles, `n, `r
					If FileExist(A_LoopField)
							FileDelete, % A_LoopField

			tempfiles:=hBOPV:=MOPdf:=MOPdfPages:=prevPdf:=""
	}

return
;}

TT9Off:                                                         	;{
	ToolTip,,,,9
return
;}

MouseIsOver(thisHwnd) {
	MouseGetPos,,, hWinNow
	If hWinNow = thishwnd
			return true
return false
}

;}

;}

;{ 9. spezielle in PDF Suchfunktionen

PdfFuzzyContents(Pat, PdfPath) {                                                                            	;-- startet ein externes Skript für die Suche nach einem Patientennamen in den Pdf-Dateien

	q:= Chr(0x22)
	;ToolTip, % "PdfFuzzy.ahk " q . Pat[3] . q . " " . q . Pat[4] . q . " " . q . PdfPath . q, 200, 200, 14
	;RunWait, PdfFuzzy.ahk % " " . q . Pat[3] . "|" . Pat[4] . "|" . PdfPath . q

	Run, % "PdfFuzzy.ahk " q . Pat[3] . q . " " . q . Pat[4] . q . " " . q . PdfPath . q

return
}

PdfFuzzyContents2(Pat, PdfPath) {

		PdfName:= StrReplace(PdfPath, "`,", " ")
		PdfName:= StrReplace(PdfPath, "`;", " ")
		PdfName:= StrReplace(PdfPath, "   ", " ")
		PdfName:= StrReplace(PdfPath, "  ", " ")
		PdfName:= StrReplace(PdfPath, " ", "_")
		PdfName:= StrReplace(PdfName, "`.pdf ", "")

		LastPage:= pdfinfo(BOPdfFile, "Pages")
		pdfText:= pdftotext(PdfPath, LastPage)
		If PdfText:="" {
			MsgBox, Bei diesem Dokument wurde bisher keine Texterkennung durchgeführt
			return
		}

		Condensed:="", wC1:=0, wC2:=0
		Loop, Parse, PdfText, `n, `r
		{
				line:= StrReplace(line, "`.", " ")
				line:= StrReplace(line, "`,", " ")
				line:= StrReplace(line, "`;", " ")
				line:= StrReplace(line, "`:", " ")
				line:= StrReplace(line, "`#", " ")
				line:= StrReplace(line, "`+", " ")
				line:= StrReplace(line, "   ", " ")
				line:= StrReplace(line, "  ", " ")

				condensed.= Trim(A_LoopField) . " "
		}

		FileAppend % Condensed, %A_ScriptDir%\PdfData\%PdfName%-condensed.txt
		Loop, Parse, Condensed, %A_Space%
		{

			If (StrReplace(PdfName, " ", "") = "")
								continue

			FileAppend, % A_LoopField "`n", %A_ScriptDir%\pdfer.txt
			a:= StrDiff(A_LoopField, Pat[3])
			b:= StrDiff(A_LoopField, Pat[4])

			if (a<0.2) {
					lst1.= A_LoopField . ", "
					FileAppend, % Pat[3] ": " a "  `[ " A_LoopField " `]`n", %A_ScriptDir%\PdfData\%PdfName%-sift.txt
					wC1++
			}

			if (b<0.2) {
					lst2.= A_LoopField . ", "
					FileAppend, % Pat[4] ": " b "  `[ " A_LoopField " `]`n", %A_ScriptDir%\%PdfName%-sift.txt
					wC2++
			}

		}

		If (wC1 AND wC2) {

					MsgBox, % "Folgende relavanten Übereinstimmungen wurden gefunden zu: `n"         Pat[3] "`n`n" lst1 "`n`n`nund zu: "        Pat[4] "`n" lst2

		}else{
					PraxTT("Der ähnlichkeitsbasierende Stringvergleich hat keine wesentlichen`nÜbereinstimmungen zum Patientennamen in der Pdf-Datei gefunden!", "7 6")
		}

}

Encode(input, password, DoubleBasicEncoding=0) {													;{

    If (StrLen(input) < 1 or StrLen(password) < 1){
        Return
    }
    input := Flip(input)
	numberchain := Abs(Convert(SHA1(password), "base16", "base10"))
    KBase:=Mul:=Add:=Manipulation1:=Manipulation2:=Manipulation3:=num:= ""
    Manipulation4 := Array()
    state:=switch:=count:= 0
    Loop  {
        Loop, parse, numberchain
        {
            token := A_LoopField
            If (state = 0 and token > 0 and token < 4) {
                KBase .= token
                state++
                Continue
            }
            If (state = 1 and (KBase < 3 or token < 7)) {
                KBase .= token
                state++
                Continue
            }
            If (state = 2 and token > 0) {
                Mul .= token
                state++
                Continue
            }
            If (state = 3) {
                If (StrLen(Mul) = 4)
                {
                    state++
                    Continue
                }
                Mul .= token
            }
            If (state = 4 and token > 0) {
                Add .= token
                AddLen := Round(sum(sum(numberchain)) / 10^4)
                state++
                Continue
            }
            If (state = 5) {
                If (StrLen(Add) >= AddLen)
                {
                    state++
                    Continue
                }
                Add .= token
            }
            If (state = 6 and token > 0) {
                Manipulation1 .= token
                If (StrLen(Manipulation1) >= 10)
                {
                    state++
                    Continue
                }
            }
            If (state = 7 and token > 0) {
                Manipulation2 .= token
                If (StrLen(Manipulation2) >= 10)
                {
                    state++
                    Continue
                }
            }
            If (state = 8) {
                If (switch = 0) {
                    tmp := token
                    switch := 1
                    Continue
                }
                If (switch = 1) {
                    tmp .= token
                    switch := 2
                    Continue
                }
                If (switch = 2)
                {
                    If (tmp != "00")
                        Manipulation3 .= Convert(tmp, "base10", "base36")
                    Else
                        Manipulation3 .= 0
                    If (StrLen(Manipulation3) >= 5) {
                        Manipulation3 := SubStr(Manipulation3, 1, 5)
                        state++
                        switch := 0
                        Continue
                    }
                    switch := 0
                    Continue
                }
            }
            If (state = 9) {
                If (switch = 0) {
                    num .= token
                    If (StrLen(num) >= 2){
                        num := SubStr(num, 1, 2)
                        switch := 1
                        Continue
                    }
                }
                If (switch = 1) {
                    count++
                    Manipulation4[count] := num
                    switch := 0
                    num := ""
                    If (count >= 10)
                        state++
                    Continue
                }
            }
            If (state = 10) {
                Break
            }
        }
        If (state = 10) {
            Break
        }
        If (state = 0) {
            KBase := 3
            state++
        }
        If (state = 1) {
            KBase .= 6
            state++
        }
    }
    tmp := ""
    GNI_Manipulation1_counter := GNI_Manipulation2_counter := GNI_Manipulation3_counter := 1
    counter1 := ready := 0
    len := StrLen(input)
    If (DoubleBasicEncoding=1) {
        Loop, % len
        {
            tmp .= Convert((Asc(SubStr(input, A_Index, 1)) * Mul + Add), "base10", "base" . KBase) . (A_Index < len ? "|" : "")
        }
        input := tmp
        len := StrLen(input)
        tmp := ""
    }
    string1 := ""
    Loop {
        i := GetNextItem("Manipulation1")
        Loop, %i%
        {
            string1 .= Convert(((Asc(SubStr(input, (counter1 + A_Index), 1)) + GetNextItem("Manipulation2")) * Mul + Add), "base10", "base" . KBase)
            If ((counter1 + A_Index) = len) {
                ready := 1
                Break
            }
            string1 .= "|"
        }
        If ready
            Break
        counter1 += i
        string1 .= Convert(((Asc(GetNextItem("Manipulation2")) + GetNextItem("Manipulation2")) * Mul + Add), "base10", "base" . KBase) . "|"
    }
    string2 := ""
    Loop, parse, string1, |
    {
        token := GetNextItem("Manipulation3")
        StringReplace, field, A_LoopField, %token%, |, 1
        string2 .= field . token
    }
    string3 := ""
    m4counter := 1
    len := StrLen(string2)
    Loop, % len
    {
        string3 .= Chr(Asc(SubStr(string2, A_Index, 1)) + Manipulation4[m4counter])
        m4counter ++
        If (m4counter > 10)
            m4counter := 1
    }
    Return string3
}

Decode(in, pwd, DoubleBasicEncoding=0) {
    global
    If (StrLen(in) < 1 or StrLen(pwd) < 1)
    {
        Return
    }
    numberchain := Abs(Convert(SHA1(pwd), "base16", "base10"))
    KBase := Mul := Add := Manipulation1 := Manipulation2 := Manipulation3 := num := ""
    Manipulation4 := Array()
    state := switch := count := 0
    Loop {
        Loop, parse, numberchain
        {
            token := A_LoopField
            If (state = 0 and token > 0 and token < 4) {
                KBase .= token
                state++
                Continue
            }
            If (state = 1 and (KBase < 3 or token < 7)) {
                KBase .= token
                state++
                Continue
            }
            If (state = 2 and token > 0) {
                Mul .= token
                state++
                Continue
            }
            If (state = 3) {
                If (StrLen(Mul) = 4)
                {
                    state++
                    Continue
                }
                Mul .= token
            }
            If (state = 4 and token > 0) {
                Add .= token
                AddLen := Round(sum(sum(numberchain)) / 10^4)
                state++
                Continue
            }
            If (state = 5) {
                If (StrLen(Add) >= AddLen)
                {
                    state++
                    Continue
                }
                Add .= token
            }
            If (state = 6 and token > 0) {
                Manipulation1 .= token
                If (StrLen(Manipulation1) >= 10)
                {
                    state++
                    Continue
                }
            }
            If (state = 7 and token > 0) {
                Manipulation2 .= token
                If (StrLen(Manipulation2) >= 10)
                {
                    state++
                    Continue
                }
            }
            If (state = 8) {
                If (switch = 0) {
                    tmp := token
                    switch := 1
                    Continue
                }
                If (switch = 1) {
                    tmp .= token
                    switch := 2
                    Continue
                }
                If (switch = 2)
                {
                    If (tmp != "00")
                        Manipulation3 .= Convert(tmp, "base10", "base36")
                    Else
                        Manipulation3 .= 0
                    If (StrLen(Manipulation3) >= 5) {
                        Manipulation3 := SubStr(Manipulation3, 1, 5)
                        state++
                        switch := 0
                        Continue
                    }
                    switch := 0
                    Continue
                }
            }
            If (state = 9) {
                If (switch = 0) {
                    num .= token
                    If (StrLen(num) >= 2){
                        num := SubStr(num, 1, 2)
                        switch := 1
                        Continue
                    }
                }
                If (switch = 1) {
                    count++
                    Manipulation4[count] := num
                    switch := 0
                    num := ""
                    If (count >= 10)
                        state++
                    Continue
                }
            }
            If (state = 10) {
                Break
            }
        }
        If (state = 10) {
            Break
        }
        If (state = 0) {
            KBase := 3
            state++
        }
        If (state = 1) {
            KBase .= 6
            state++
        }
    }
    GNI_Manipulation1_counter := 1
    GNI_Manipulation2_counter := 1
    GNI_Manipulation3_counter := 1

    string3 := in
    string2 := ""
    string1 := ""
    string0 := ""
    m4counter := 1
    len := StrLen(string3)
    Loop, % len
    {
        string2 .= Chr(Asc(SubStr(string3, A_Index, 1)) - Manipulation4[m4counter])
        m4counter ++
        If (m4counter > 10)
            m4counter := 1
    }
    Loop {
        token := GetNextItem("Manipulation3")
        endpos := InStr(string2, token)
        field := SubStr(string2, 1, (endpos - 1))
        StringTrimLeft, string2, string2, %endpos%
        StringReplace, field, field, |, %token%, 1
        string1 .= field . "|"
        If (StrLen(string2) = 0)
            Break
        If (StrLen(string2) = lastlen)
            Break
        lastlen := StrLen(string2)
        startpos := endpos + 1
    }
    StringTrimRight, string1, string1, 1
    i := GetNextItem("Manipulation1") + 1
    Loop, parse, string1, |
    {
        If (A_Index = i) {
            i += GetNextItem("Manipulation1") + 1
            GetNextItem("Manipulation2")
            GetNextItem("Manipulation2")
        } Else {
            string0 .= Chr((((Convert(A_LoopField, "base" . KBase, "base10") - Add) / Mul) - GetNextItem("Manipulation2")))
        }
    }
    If (DoubleBasicEncoding = 1)
    {
        tmp := ""
        Loop, parse, string0, |
        {
            tmp .= Chr(((Convert(A_LoopField, "base" . KBase, "base10") - Add) / Mul))
        }
        string0 := tmp
        tmp := ""
    }
    string0 := Flip(string0)
    Return string0
}
GetNextItem(varname) {
    global
    item := SubStr(%varname%, GNI_%varname%_counter, 1)
    GNI_%varname%_counter ++
    If (GNI_%varname%_counter > StrLen(%varname%))
        GNI_%varname%_counter := 1
    Return item
}
sum(input) {
    output := 0
    Loop, parse, input
    {
        output += A_LoopField
    }
    Return output
}
Flip(in) {
    VarSetCapacity(out, n:=StrLen(in))
    Loop %n%
        out .= SubStr(in, n--, 1)
    return out
}
Convert(value, from, to) {
    if !(value and from and to)   {
    Return "?"
    }
    base2 = Base2|Binary|Bin|Digital|Binär|Dual|Di|B
    base3 = Base3|Ternary|Triple|Trial|Ternär
    base4 = Base4|Quaternary|Quater|Tetral|Quaternär
    base5 = Base5|Quinary|Pental|Quinär
    base6 = Base6|Senary|Hexal|Senär
    base7 = Base7|Septenary|Peptal|Heptal
    base8 = Base8|Octal|Oktal|Oct|Okt|O
    base9 = Base9|Nonary|Nonal|Enneal
    base10 = Base10|Decimal|Dezimal|Denär|Dekal|Dec|Dez|D
    base11 = Base11|Undenary|Monodecimal|Monadezimal|Hendekal
    base12 = Base12|Duodecimal|Dedezimal|Dodekal
    base13 = Base13|Tridecimal|Tridezimal|Triskaidekal
    base14 = Base14|Tetradecimal|Tetradezimal|Tetrakaidekal
    base15 = Base15|Pentadecimal|Pentadezimal|Pentakaidekal
    base16 = Base16|Hexadecimal|Hexadezimal|Hektakaidekal|Hex|H
    base17 = Base17|Peptaldecimal|Peptaldezimal|Heptakaidekal
    base18 = Base18|Octaldecimal|Oktaldezimal|Octakaidekal|Oktakaidekal
    base19 = Base19|Nonarydecimal|Nonaldezimal|Enneakaidekal
    base20 = Base20|Vigesimal|Eikosal
    base30 = Base30|Triakontal
    base40 = Base40|Tettarakontal
    base50 = Base50|Pentekontal
    base60 = Base60|Sexagesimal|Hektakontal
    StringReplace, value_form, value,(,,all
    StringReplace, value_form, value_form ,),, all
    if value_form is not Alnum
    {
    Return "?"
    }
    if (InStr(from, "base"))
    {
        StringTrimLeft, base_check, from, 4
        if base_check is not Integer
        {
        Return "?"
        }
        else
            from := base_check
    }
    if (InStr(to, "base"))
    {
        StringTrimLeft, base_check, to, 4
        if base_check is not Integer
        {
            Return "?"
        }
        else
            to := base_check
    }
    base_loop := 1
    loop, 60
    {
        if from is Integer
            if to is Integer
                Break
        if (base_loop < 20)
            base_loop ++
        else
            base_loop += 10
        if (base_loop > 60)
            Break
        base := base%base_loop%
        loop parse, base, |
        {
            if (from = A_LoopField)
                from := base_loop
            if (to = A_LoopField)
                to := base_loop
        }
    }
    if (from < 11)
        if value is not Integer
        {
            Return "?"
        }
    con_letter := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    result_dec=
    length := StrLen(value)
    counter := 0
    parenthesis := False
    loop, %length%
    {
        StringMid, char, value, (length + 1 - A_Index), 1
        if (char = ")")
        {
            if parenthesis
            {
                Return "?"
            }
            parenthesis := True
            Continue
        }
        else if (char = "(")
        {
            if !parenthesis
            {
                Return "?"
            }
            parenthesis := False
            if !par_char
            {
                Return "?"
            }
            char := par_char
        }
        else if parenthesis
        {
            if char is not Integer
            {
                Return "?"
            }
            par_char := char . par_char
            Continue
        }
        else if char is Alpha
        {
            StringGetPos, char_pos, con_letter, %char%
            StringReplace, char, char, %char%, %char_pos%
            char += 10
            if (char >= from)
            {
                Return "?"
            }
        }
        if (char >= from)
        {
            Return "?"
        }
        result_dec += char * (from**counter)
        counter ++
    }
    if (to = 10)
        Return %result_dec%
    result=
    while (result_dec)
    {
        char := Mod(result_dec , to)
        if (char > 35)
            char := "(" . char . ")"
        else if (char > 9)
            StringMid, char, con_letter, (char - 9), 1
        result :=  char . result
        result_dec := Floor(result_dec / to)
    }
    Return %result%
}
varum() {
	String0:=String1:=String2:=String3:=Manipulation1:=Manipulation2:=Manipulation3:=Manipulation4:=field:=Mul:=""
	GNI_Manipulation1_counter:=GNI_Manipulation2_counter:=GNI_Manipulation3_counter:=count:=counter:=counter1:=""
	Add:=AddLen:=endpos:=Kbase:=lastlen:=len:=m4counter:=item:=numberchain:=ready:=startpos:=switch:=tmp:=token:=""
return
}
SHA1(string) {
        return HashFromString(string, 0x8004)
}
HashFromString(string, algid, key=0) {
        len := strlen(string)
        if (A_IsUnicode) {
            VarSetCapacity(data, len)
            StrPut := "StrPut"
            %StrPut%(string, &data, len, "cp0")
            return HashFromAddr(&data, len, algid, key)
        }
        data := string
        return HashFromAddr(&data, len, algid, key)
}
HashFromAddr(pData, len, algid, key=0) {
        hProv := size := hHash := hash := ""
        ptr := (A_PtrSize) ? "ptr" : "uint"
        aw := (A_IsUnicode) ? "W" : "A"
        if (DllCall("advapi32\CryptAcquireContext" aw, ptr "*", hProv, ptr, 0, ptr, 0, "uint", 1, "uint", 0xF0000000)) {
            if (DllCall("advapi32\CryptCreateHash", ptr, hProv, "uint", algid, "uint", key, "uint", 0, ptr "*", hHash)) {
                if (DllCall("advapi32\CryptHashData", ptr, hHash, ptr, pData, "uint", len, "uint", 0)) {
                    if (DllCall("advapi32\CryptGetHashParam", ptr, hHash, "uint", 2, ptr, 0, "uint*", size, "uint", 0)) {
                        VarSetCapacity(bhash, size, 0)
                        DllCall("advapi32\CryptGetHashParam", ptr, hHash, "uint", 2, ptr, &bhash, "uint*", size, "uint", 0)
                    }
                }
                DllCall("advapi32\CryptDestroyHash", ptr, hHash)
            }
            DllCall("advapi32\CryptReleaseContext", ptr, hProv, "uint", 0)
        }
        int := A_FormatInteger
        SetFormat, Integer, h
        Loop, % size
        {
            v := substr(NumGet(bhash, A_Index-1, "uchar") "", 3)
            while (strlen(v)<2)
            v := "0" v
            hash .= v
        }
        SetFormat, Integer, % int
        return hash
}
;}


;}

;{ 10. andere Funktionen - | OnMessage | Albishandling | Gui | Kommunikation | Bitmapbearbeitung | Listview | RamDrive | Fensterhooks | integriertes Icon

;---------------------------------------ONMESSAGE                                                      	;{

BOWM_MOUSEMOVE(wparam, lparam, msg, hwnd) {         				;-- Gui verschieben, Pdf Preview Fenster ausblenden

	;hwnd - ändert sich nur wenn die Maus über Skripteigenen Gui's steht, fremde Fensterhandle werden nicht erkannt, deshalb wird hier MouseGetPos verwendet
	MouseGetPos,,, hWinMouseOver
	If !InStr(Running, "umbenennen") && ( hWinMouseOver not in %hBOPV%,%hBO%,%hBOPT%,%hBONC%)
				SetTimer, BOPVGuiClose, -1

;return
}

BOWM_LBUTTONDOWN(wparam, lparam, msg, hWMLbd) {				;-- Klick mit der li. Maustaste verschiebt das Hauptfenster und schließt das PDF-Vorschau Gui

	 ;ToolTip % "Meldung " GetHex(msg) " angekommen:`nWPARAM: " GetHex(wParam) "`nLPARAM: " GetHex(lParam)

	if (hWMLbd = hBO) { ; LButton
			If WinExist("ScanPool - Pdf Vorschau")
					SetTimer, BOPVGuiClose, -0
			PostMessage, 0xA1, 2,,, A 					; WM_NCLBUTTONDOWN
	}

return
}

BOWM_RBUTTONDOWN(wparam, lparam, msg, hWMRbd) {			;-- Gui Contextmenu Funktion (ausgelöst durch re.Maustaste)

	global

	if (hWMRbd = hLV1)  and  (Running = "")  and  (addingfiles = 0)					;wparam=2 rechte Mouse      ;(wparam = 2)  And
				a:= 1 ;SetTimer, BOGuiContextMenu, -20

return
}

OnMessageStatus(status:=1) {											             	;-- alle genutzen OnMessage-Aufrufe gesamt an- oder abschalten

	If OnMessageByPass
			status := 1

	If status	{
			OnMessage(0x200	, "BOWM_MOUSEMOVE")
			OnMessage(0x201	, "BOWM_LBUTTONDOWN")
			;OnMessage(0x202	, "BOWM_RBUTTONDOWN")
			OnMessage(0x2A3	, "BOWM_MOUSEMOVE")
			OnMessage(0x4E		, "LVA_OnNotify")
	}
	else	{
			OnMessage(0x200	, "")
			OnMessage(0x201	, "")
			;OnMessage(0x202	, "")
			OnMessage(0x2A3	, "")
			;OnMessage(0x4E		, "")
	}

return
}

;}

;---------------------------------------ALBISHANDLING                                                 	;{

BOAkteOeffnen(Name1, Name2) {

	LastPopUpWin:= DLLCall("GetLastActivePopup", "uint", AlbisWinID)

	if (Name2=!"")
		Name2:= "`," . Name2

	AlbisAkteOeffnen(Name1 . Name2)
	PopUpWin:= WaitForNewPopUpWindow(AlbisWinID, LastPopUpWin, 5)    ;neue Art WinWait - wartet auf neue Fenster, unabhängig irgendeines Namens

return PopUpWin
}

;}

;---------------------------------------GUI                                                                     	;{
ButtonsOnOff(Buttons, GuiID, State:="Enable") {			    				                 		;-- eine mit '|' Zeichen getrennte Liste an vButton Namen übergeben um mehrere Buttons gleichzeitig anzusprechen

	;State kann jeden Befehl enthalten - Disable, Enable, Show, Hide ....
		Loop, Parse, Buttons, `|
				GuiControl, %GuiID%: %State%, %A_Loopfield%

return
}

LastNameGui(Pat) {                                                                                    	    	;-- manuelle Nachnamen Auswahl

	Gui, LName: New
	Gui, LName: Font, S10, Futura Bk Bt
	Gui, LName: Add, Text, xCenter ym w500, Bitte wählen Sie den Nachnamen aus!
	Gui, LName: Add, Button, xm y+m vBOLNB1 gBOLN1, % Pat[1]
	GuiControlGet, BOLNB1, LName: Pos
	Gui, LName: Add, Button, % "x" (BOLNB1X + BOLNB1W + 20) " yp vBOLNB2 gBOLN2", % Pat[2]
	Gui, LName: Show, AutoSize, Addendum für AlbisOnWindows - Wahl des Nachnamens
	WinWaitClose, Wahl des Nachnamens

	return B

	BOLN1:
	B=1
	Gui, LName: Destroy
	return

	BOLN2:
	B=2
	Gui, LName: Destroy
	return

}

BOGuiProcessTipDispatcher(procStr) {                                                                 	;-- entscheidet über Anfragen zum Prozeßablauf, gibt Information zum Status zurück

	/*	Funktionsbeschreibung
	   *
		 * Funktionen/Labels können durch Aufruf der Funktion ohne Argument erfahren ob sie im Moment ausgeführt werden dürfen,
		 * dies ist wegen der implementierten WinEventHooks notwendig
		 * Parameter procstr sollte folgende Angaben enthalten: "automatische Signatur On" oder "automatische Signatur Off"
		 * Abfrageparameter leer ("")	: gibt die laufende Funktion oder einen leeren String zurück
		 * Abfrageparameter "last"		: gibt die vorherig laufende Funktion zurück egal ob derzeit nichts läuft und auch wenn etwas läuft,
		 *                                      	  gibt immer einen benannten Prozeß zurück, niemals einen leeren String
	   *
	*/

	If !Running = ""
			SetTimer, BOWaitProcess, 100

return
}

BOGuiProcessTipHandler(procStr) {

	;zeigt einen Hinweis das das Skript im Moment arbeitet, damit der Nutzer nicht versehentlich störend eingreift
		If Running And !WinExist("ScanPool - arbeitet")		{

				ButtonsOnOff("BOButton3|BoButton2|BOButton1|BOUpDown1|BOEdit1", "BO", "Disable")
				;OnMessageStatus(0)

				If WinExist("ScanPool - Pdf Vorschau")
						gosub BOPVGuiClose

				BOGuiProcessTip()
				Gui, BO: Default

				Procedure:= StrReplace(Running, "`n", " ")
				Procedure:= StrReplace(Procedure, "- ", "")
				RQIndex:= RQueue.Push({"Procedure" : (Procedure), "starttime": (A_TickCount), "endtime": 0, "Duration": 0})
				If DebugFunctions
				{
						RQList.= "'" RQueue[RQIndex].Procedure "' wurde gestartet`n"
						If StrSplit(RQList, "`n").MaxIndex() > 5
								RQList:= LineDelete(RQList, 1)
						ToolTip, % RQList, 600, 5, 17
				}
		}
		  else if WinExist("ScanPool - arbeitet") and !InStr(WhatIsRunning, Running)		{

				RQueue[RQIndex].EndTime 	:= A_TickCount
				RQueue[RQIndex].Duration	:= Round((RQueue[RQIndex].endtime - RQueue[RQIndex].starttime)/1000, 3)
				Procedure:= StrReplace(Running, "`n", " ")
				Procedure:= StrReplace(Procedure, "- ", "")
				RQIndex:= RQueue.Push({"Procedure" : (Procedure), "starttime": (A_TickCount), "endtime": 0, "Duration": 0})
				WhatIsRunning := Running

				ButtonsOnOff("BOButton3|BoButton2|BOButton1|BOUpDown1|BOEdit1", "BO", "Disable")
				;OnMessageStatus(0)

				GuiControl, BOPT: , PTText3, % Running
				Gui, BO: Default
		}
		  else if WinExist("ScanPool - arbeitet") and ( Running = "" )		{

				RQueue[RQIndex].EndTime 	:= A_TickCount
				RQueue[RQIndex].Duration	:= Round((RQueue[RQIndex].endtime - RQueue[RQIndex].starttime)/1000, 3)
				If DebugFunctions				{
						RQList.= "'" RQueue[RQIndex].Procedure "' lief für " RQueue[RQIndex].Duration " Sekunden." "`n"
						If StrSplit(RQList, "`n").MaxIndex() > 5
								RQList:= LineDelete(RQList, 1)
						ToolTip, % RQList, 600, 5, 17
				}

				ButtonsOnOff("BOButton3|BoButton2|BOButton1|BOUpDown1|BOEdit1", "BO", "Enable")
				;OnMessageStatus(1)

				If !WinExist("ahk_class TV_ControlWin") and !IsRemoteSession()
								AnimateWindow(hBOPT, 150, FADE_HIDE)

				Gui, BOPT	: 	Destroy
				Gui, BO: 		Default
				hBOPT := ""

		}


return
}

BOGuiProcessTip() {                                                                                           	;-- zeigt einen Hinweis das ein Vorgang in Bearbeitung ist

			If Running = ""
					return

			static PTText1, PTText2, PTText3

			WhatIsRunning:= Running

			Gui								, BOPT: New	 			,  -Resize +ToolWindow -Caption +AlWaysOnTop +0x90040000 +Disabled +E0x8000200 HWndhBOPT				; +Border
			Gui								, BOPT: Margin 		, 5				 			, 5
			Gui								, BOPT: Color	 		, 202842
			Gui								, BOPT: Font	 			, % BOFontOptions, % BOFont
			Gui								, BOPT: Font	 			, % "s" 10*DPIFactor " q5 cWhite"
			Gui								, BOPT: Add	 	, Text, % "xm ym Section 	  vPTText1	BackgroundTrans Center", % "Addendum für AlbisOnWindows"
			GuiControlGet, PTText1	, BOPT: Pos
			Gui								, BOPT: Font	 			, % "s" 22*DPIFactor " cFFFF00 q5 Bold Underline"
			Gui								, BOPT: Add	 	, Text, % "xs w" (PTText1W)  " vPTText2	BackgroundTrans Center", % "   ScanPool   "
			Gui								, BOPT: Font	 			, % "s" 14*DPIFactor " cFFFFFF q5 norm"
			Gui								, BOPT: Add	 	, Text, % "xs w" (PTText1W) "           	BackgroundTrans Center", % "Vorgang in`nBearbeitung:"
			Gui								, BOPT: Font	 			, % "s" 14*DPIFactor " cFF6666 q5 bold"
			Gui								, BOPT: Add	 	, Text, % "xs w" (PTText1W) " vPTText3	BackgroundTrans Center", % Running
			Gui								, BOPT: Font	 			, % "s" 14*DPIFactor " cFFFFFF q5 norm"
			Gui								, BOPT: Add	 	, Text, % "xs w" (PTText1W) " 				BackgroundTrans Center", % "...bitte warten..."
			Gui								, BOPT: Show	 		, % "xCenter yCenter AutoSize Hide NA"																				    , Addendum für AlbisOnWindows - ScanPool - arbeitet

		;ProcessTip Fenster im Hauptfenster zentriert positionieren
			BO					:= GetWindowInfo(hBO)
			BOPT				:= GetWindowInfo(hBOPT)
			BOPT.WindowX	:= ( 	BO.WindowX   	+ 	BO.WindowW//2 	- 	BOPT.WindowW//2 	)
			BOPT.WindowY	:= ( 	BO.WindowY	+ 	BO.WindowH//2 	- 	BOPT.WindowH//2 	)

			SetWindowPos(hBOPT, BOPT.WindowX, BOPT.WindowY, BOPT.WindowW, BOPT.WindowH)

		;einblenden mit Animation wenn keine Teamviewerverbindung besteht
			If !WinExist("ahk_class TV_Control") and !IsRemoteSession()
					AnimateWindow(hBOPT, 150, FADE_SHOW)

			Gui, BOPT: Show, NA
			Gui, BO:	  Default
			;Gui, BO: Listview, ListView1

return
}

BOWaitProcess:

return

SetGuiStatusBar(Gui, row1, row2)	{
	Gui, %Gui%: Default
	SB_SetText(row1, 1)
	SB_SetText(row2, 2)
return
}
;}

;---------------------------------------INTER-SKRIPT-KOMMUNIKATION                         	;{
Receive_WM_COPYDATA(wParam, lParam) {                                                                	;-- empfängt Nachrichten von anderen Skripten die auf demselben Client laufen
	global InComing
    StringAddress := NumGet(lParam + 2*A_PtrSize)
	InComing := StrGet(StringAddress)
	SetTimer, MessageWorker, -10
    return true
}

MessageWorker:                                                                                                      	;{

	MWmsg:= StrSplit(InComing, "|")

	If InStr(MWmsg[1], "PatDB")
	{
			If MWmsg.2 = "PatID"
					recPatID:= MWmsg.3
	}

	If InStr(MWmsg[1], "TPCount")
			result := Send_WM_COPYDATA("TPCount|" TPCount 		, MWmsg[2])

return
;}

received:                                                                                                                	;{
	Loop 100
		Sleep 20
	TrayTip
return
;}
;}

;---------------------------------------GDI                                                                     	;{

Gdip_RotateBitmap(pBitmap, Angle, Dispose=1) { 														;-- returns rotated bitmap. By Learning one.

		static rot:=0

		Gdip_GetImageDimensions(pBitmap, Width, Height)
		Gdip_GetRotatedDimensions(Width, Height, Angle, RWidth, RHeight)
		Gdip_GetRotatedTranslation(Width, Height, Angle, xTranslation, yTranslation)

		pBitmap2 	:= Gdip_CreateBitmap(RWidth, RHeight)
		G2 			:= Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
		Gdip_TranslateWorldTransform(G2, xTranslation, yTranslation)
		Gdip_RotateWorldTransform(G2, Angle)
		Gdip_DrawImage(G2, pBitmap, 0, 0, Width, Height)

		hBitmap 	:= Gdip_CreateHBITMAPFromBitmap(pBitmap2)

		Gdip_ResetWorldTransform(G2)
		Gdip_DeleteGraphics(G2)
		if Dispose
				Gdip_DisposeImage(pBitmap)

return {"hBitmap": (hBitmap), "pBitmap": (pBitmap2), "Width": (Width), "Height": (Height), "Rotation": rot}
} ; http://www.autohotkey.com/community/viewtopic.php?p=477333#p477333

LoadPicture(PreviewFilename, PdfFileName, dpi, page, cmdReturn:="hBitmap", uselast:= 0) {      				  	;-- lädt ein Bild und gibt die Handles als Key:Value Objekt zurück

	; ein leerer cmdReturn Parameter gibt ein Objekt zurück
	; uselast = 1 - wenn später z.B. das pBitmap für Gdip_RotateBitmap gebraucht wird
	; RamDiskPath - ist am Anfang des Skriptes global gemacht worden

	static pBitmap, hBitmap

	If !uselast
	{
			SplitPath, PreviewFilename,, PreviewPath
			png		 	:= PdfToPng(PdfFileName, dpi, page, PreviewPath)			;this must be the full path of pdf file with extension
			pBitmap 	:= Gdip_CreateBitmapFromFile(PreviewFilename)
			hBitmap 	:= Gdip_CreateHBITMAPFromBitmap(pBitmap)
			Gdip_DisposeImage(pBitmap)
	}

	If InStr(cmdReturn, "hBitmap")
		return hBitmap
	else if InStr(cmdReturn, "pBitmap")
		return pBitmap

return {"hBitmap": (hBitmap), "pBitmap": (pBitmap)}
}

CheckPlace(hwnd, x1, y1, y2, y3, tolerance=0.3) {                                 				   	;-- leer , gedacht für Auto-OCR



}

;}

;---------------------------------------LISTVIEW                                                             	;{

LV_ItemIs(lvhwnd, row) {

	SendMessage, 4140, row - 1, 0xF000, ahk_id %lvhwnd%  ; 4140 ist LVM_GETITEMSTATE.  0xF000 ist LVIS_STATEIMAGEMASK.
	IstMarkiert := (ErrorLevel >> 12) - 1  ; Setzt IstMarkiert auf True, wenn Reihennummer markiert ist, ansonsten auf False.

return IstMarkiert
}

LV_GetSelectedText() {																			;-- speziell für ScanPool entworfene Funktion

		; modified for only use in this script by Ixiko 08-2018

		Retrieved:=Object(), Column:=[], RIndex:=0, RowNumber := 0
		Gui, BO: Default

		Loop, % LV_GetCount()
		{
				RowNumber := LV_GetNext(RowNumber)
				if !RowNumber
						break

				LV_GetText(RetrievedText, RowNumber, 1)
				RowOne:= Trim(RetrievedText, " `r`n`t")
				LV_GetText(RetrievedText, RowNumber, 2)
				RowTwo:= Trim(RetrievedText, " `r`n`t")

				Retrieved.Push({"Pdf": RowOne, "Pages": RowTwo, "Row": RowNumber})
		}

return Retrieved
}

LV_GetCheckedItems(lvhwnd, StartRow:=1, EndRow:=99) {

	chRowIdx:=0, checked:=[]

	Gui, BO: Default
	Gui, ListView, Listview1

	Loop % LV_GetCount()
	{
				;chfiles = checked files, sammelt alle Dateien ein die im PdfReader angezeigt werden, diese Dateien werden dann dem Prozeß für die Signierung übergeben
				RowState:= LV_GetItemState(hLV1, A_Index)
				 If RowState.checked
				 {
						SB_SetText("row: " . A_Index . " is checked", 4)
						chRowIdx ++
						LV_GetText(RnText, A_Index)
						checked[chRowIdx]:= RnText
				}
	}

	;Button wurde gedrückt, aber es war nichts ausgewählt, nichts machen - return
	If !chRowIndex
			return

return checked
}
;}

;---------------------------------------RAMDRIVE                                                           	;{

CreateRamDrive(RamSize, RamDriveName, fpath) {													;-- CreateRamDrive()

	; Creates a RamDrive (X:/Y:/Z: whichever is availabe) as a RemovableDisk and assigns "RamDisk" as label
	;RamSize in MB

	q := Chr(0x22)

	SetWorkingDir, %fpath%\ImDisk_Portable

	RunWait, sc create AWEAlloc 	binPath= "%fpath%\ImDisk_Portable\awealloc.sys" 	DisplayName= "AWE Memory Allocation Driver" 	Type= kernel 	Start= demand,, Hide
		Progress, 6
	RunWait, sc create ImDisk 		binPath= "%fpath%\ImDisk_Portable\imdisk.sys" 	DisplayName= "ImDisk Virtual Disk Driver" 			Type= kernel 	Start= demand,, Hide
		Progress, 7
	RunWait, sc create ImDskSvc 	binPath= "%fpath%\ImDisk_Portable\imdsksvc.exe" DisplayName= "ImDisk Virtual Disk Driver Helper" Type= own 		Start= demand,, Hide
		Progress, 8
	RunWait, sc description AWEAlloc 			"Driver for physical memory allocation through AWE"	,, Hide
		Progress, 9
	RunWait, sc description ImDisk 				"Disk emulation driver"												,, Hide
		Progress, 10
	RunWait, sc description imdsksvc 			"Helper service for ImDisk Virtual Disk Driver."				,, Hide
		Progress, 11
	RunWait, net start ImDskSvc																									,, Hide
		Progress, 12
	RunWait, net start AWEAlloc																									,, Hide
		Progress, 13
	RunWait, net start ImDisk																										,, Hide
	NewDriveArray := ["E:","F:","G:", "H:", "I:", "J:", "K:", "L:", "M:", "N:", "O:", "P:", "Q:", "R:", "S:", "T:", "W:", "X:", "Y:", "Z:"]
	Loop, 20
	{
		tdrive := NewDriveArray[A_Index]
		;MsgBox, % fpath "`n" tdrive "`n" q . fpath . "\ImDisk_Portable\imdisk.exe" . q . " -a -s " RamSize "M -m " tdrive " -o rem -p " . q . "/fs:ntfs /v:RamDisk /A:512 /q /c /y" . q,
		DriveGet, status, Status, %tdrive%
		If (status = "Invalid")
		{
			Run, % q . fpath . "\ImDisk_Portable\imdisk.exe" . q . " -a -s " RamSize "M -m " tdrive " -o rem -p " . q . "/fs:ntfs /v:RamDisk /A:512 /q /c /y" . q,, Hide
			Loop
				DriveGet, status2, Status, %tdrive%
			Until Instr(status2,"Ready")
			Break
		}
	}
	temp:= "RamDrive"
	expression:=(%temp%:=tdrive) ;create global variable
	Progress, 15
	SetWorkingDir, %A_ScriptDir%
	Return, tdrive
}

RemoveRamDrive(RamDrv, fpath) {																			;-- Removes the RamDrive with "RamDisk" as label
	;fpath :=RamDrivePath()
	If (RamDrv="")
		RamDrv := GetRamDriveLetter()
	else
		RunWait, "%fpath%\ImDisk_Portable\imdisk.exe" -D -m %RamDrv%,, Hide
	Sleep, 100
	RunWait, sc delete ImDskSvc,,  Hide
	RunWait, sc delete AWEAlloc,,  Hide
	RunWait, sc delete ImDisk,,  Hide
	Run, taskkill /f /IM "imdisk.exe",,  Hide
	Run, taskkill /f /IM "imdsksvc.exe",,  Hide
	RunWait, sc delete ImDskSvc,,  Hide
	RegDelete, HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\ImDskSvc /f
	RunWait, sc delete AWEAlloc,,  Hide
	RegDelete, HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\AWEAlloc /f
	RunWait, sc delete ImDisk,,  Hide
	RegDelete, HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\ImDisk /f
}

GetRamDriveLetter() {
																								;-- Removes the RamDrive with "RamDisk" as label and return the drive letter
	DriveGet TempText, List, REMOVABLE
	Loop, Parse, TempText
	{
		DriveGet TempText, Label, %A_LoopField%:
		If (TempText = "RamDisk")
				Drv := A_LoopField ":"
	}

Return Drv
}

;}

;---------------------------------------WINEVENTHOOK                                                	;{

;fängt Fenster des Foxitreaders ab. | Anpassung der ScanPool Gui Größe, PopUp Fenster-Handler für den FoxitReader

InitializeWinEventHook:                     	;{

	;https://docs.microsoft.com/en-us/windows/desktop/winauto/event-constants ;{
	;	EVENT_OBJECT_CREATE 				:= 0x8000
	;	EVENT_OBJECT_DESTROY	       	:= 0x8001
	;	EVENT_OBJECT_SHOW		        	:= 0x8002
	;	EVENT_OBJECT_HIDE					:= 0x8003
	;	EVENT_OBJECT_REORDER		   	:= 0x8004			;A container object has added, removed, or reordered its children. (header control, list-view control, toolbar control, and window object - !z-order!)
	;	EVENT_OBJECT_FOCUS				:= 0x8005
	;	EVENT_OBJECT_INVOKED			:= 0x8013			;An object has been invoked; for example, the user has clicked a button.
	;	EVENT_OBJECT_NAMECHANGE 	:= 0x800C
	;	EVENT_OBJECT_VALUECHANGE	:= 0x800E			;An object's Value property has changed.  edit, header, hot key, progress bar, slider, and up-down, scrollbar

	;	EVENT_SYSTEM_CAPTUREEND		:= 0x0009			;A window has lost mouse capture.
	;	EVENT_SYSTEM_CAPTURESTART	:= 0x0008			;A window has received mouse capture. This event is sent by the system, never by servers.
	;	EVENT_SYSTEM_DIALOGEND		:= 0x0011			;A dialog box has been closed.
	;	EVENT_SYSTEM_DIALOGSTART		:= 0x0010			;A dialog box has been displayed.
	;	EVENT_SYSTEM_FOREGROUND	:= 0x0003			;The foreground window has changed. The system sends this event even if the foreground window has changed to another window in the same

		EVENT_SKIPOWNTHREAD				:= 0x0001
		EVENT_SKIPOWNPROCESS			:= 0x0002
		EVENT_OUTOFCONTEXT				:= 0x0000
	;}

	; creating the hooks
		HookProcAdr1 						:= RegisterCallback("SpWinProc1", "F")
		hWinEventHook1 					:= SetWinEventHook( 0x8000, 0x8000, 0, HookProcAdr1, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook1 					:= SetWinEventHook( 0x8002, 0x8002, 0, HookProcAdr1, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		hWinEventHook1 					:= SetWinEventHook( 0x0003, 0x0003, 0, HookProcAdr1, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )

		;HookProcAdr2 						:= RegisterCallback("SpWinProc2"			, "F")
		;hWinEventHook2 					:= SetWinEventHook( 0x800C, 0x800C, 0, HookProcAdr2, 0, 0, EVENT_SKIPOWNTHREAD | EVENT_SKIPOWNPROCESS )
		;gosub InitializeShellHook

return
;}

SpWinProc1(hWinEventHook, Event, SpHwnd1, idObject, idChild, eventThread, eventTime) {				;Hookverteiler
		Critical
		static filter := {1:"#32770", 2:"AutoHotkeyGUI", 3:"classFoxitReader", 4:"OptoAppClass"}
		Sleep, 10
		WinGetClass, SPClass, % "ahk_id " SpEHookHwnd1:= GetHex(SpHwnd1)
		TrayTip("Hook", "SpClass: " SPClass, 2)
		;If SPClass in "#32770,AutoHotkeyGUI,classFoxitReader,OptoAppClass
		For Index, filterClass in filter
			If InStr(SPClass, filterClass) {
				SetTimer, SpEvHook_WinHandler1,  -1
				break
			}
return 0
}

SpWinProc2(hWinEventHook, Event, Hwnd, idObject, idChild, eventThread, eventTime) {						;Hookverteiler

		Critical
		Sleep, 50
		If WinGetClass(SpEHookHwnd2:= GetHex(hWnd))
				WinGetClass, wclass, % "ahk_id " hwnd
		If wclass = ""
				return
		SpProc2 	:= WinGet(SpEHookHwnd2, "ProcessName")
		If InStr(SpProc2, "FoxitReader")
				SetTimer, SpEvHook_WinHandler2,  -0
return 0
}

SpEvHook_WinHandler1:                     	;{

		handlerrunning1	:= 1
		SpProc1 	:= WinGet(SpEHookHwnd1, "ProcessName")
		ToolTip, % "proc: " SpProc1
		TrayTip("Hook", "proc: " SpProc1, 2)
		If InStr(SpProc1, "FoxitReader") {
				FoxitTitle1	:= WinGetTitle(SpEHookHwnd1)
				TrayTip("Foxit Hook", "Title: " FoxitTitle1, 2)
				If InStr(FoxitTitle1, "Speichern unter bestätigen")
						FoxitReader_ConfirmSaveAs(SpEHookHwnd1)
				else if InStr(FoxitTitle1, "Speichern unter")
						FoxitReader_CloseSaveAs(SpEHookHwnd1)
				else if InStr(FoxitTitle1, "Sign Document") || InStr(FoxitTitle1, "Dokument signieren") 						; für die englische und deutsche Variante
						FoxitReader_SignDoc(SpEHookHwnd1, FoxitTitle1)
		}
		else if InStr(SpProc1, "Autohotkey") {

				hPraxomat:= WinExist("PraxomatGui ahk_class AutoHotkeyGUI")                                             	; Anpassung der ScanPool Gui Größe, PopUp Fenster-Handler für den FoxitReader
				If (hPraxomat and !PraxoWas)
				{
							BO:= GetWindowInfo(hBO), Praxomat:= GetWindowInfo(PraxoWas:= hPraxomat)   ; <--- Änderung hier am 08.04.2019 gemacht
							If RectOverlapsRect(BO.WindowX, BO.WindowY, BO.WindowW, BO.WindowH, Praxomat.WindowX, Praxomat.WindowY, Praxomat.WindowW, Praxomat.WindowH, "")
							{
									BOy:= Praxomat.WindowY + Praxomat.WindowH + 5
									BOh:= (BO.WindowH + BOy) > (monHeight - BOyNormal + 5) ? (monHeight - BOyNormal + 5) : BO.WindowH
									SetWindowPos(hBO, BO.WindowX, BOy, BO.WindowW, BOh)
							}
				}
				else if (PraxoWas = SpEHookHwnd1)
				{
							BOy:= BOyNormal, BOh:= BOhMax, PraxoWas:= 0
							SetWindowPos(hBO, BO.WindowX, BOy, BO.WindowW, BOh)
				}

		}
		else if InStr(SpProc1, "albis") {

				albisPopUpText:= WinGetText(SpEHookHwnd1)
				albisPopUpTitle:= WinGetTitle(SpEHookHwnd1)
				;FileAppend, % ", (3) " albisPopUpTitle, %A_ScriptDir%\logs\FoxitTitleLog.txt
				If RegExMatch(albisPopUpText, "i)Patient.*nicht vorhanden" )
							PatientNichtVorhanden := 1
				else if InStr(albisPopUpText, "Patient hat in diesem Quartal")
							VerifiedClick("Button1", "ALBIS", "Patient hat in diesem Quartal")
				else if InStr(albisPopUpTitle, "Herzlichen Glückwunsch zum Geburtstag")
							VerifiedClick("Button1", "Herzlichen Glückwunsch zum Geburtstag", "")

		}

		handlerrunning1	:= 0
		SpProc1            	:= FoxitTitle1:= ""
		SpEHookHwnd1	:= 0

return
;}

SpEvHook_WinHandler2:                     	;{

		handlerrunning2	:= 1
		FoxitTitle2	:= WinGetTitle(SpEHookHwnd2)

		handlerrunning 	:= 0

return
;}

SignatureHelper: ;{

	activeTitle:= WinGetTitle( activeID:= WinExist("A") )
	if (InStr(activeTitle, "Sign Document") or InStr(activeTitle, "Dokument signieren"))	{
				SetTimer, SignatureHelper, Off
				If !handlerrunning1				{
						SpEHookHwnd1:= activeID
						SetTimer, SpEvHook_WinHandler1, -0
				}
	}

	activeTitle:= activeID:= ""

return
;}

;}	Speichern unter bestätigen

;---------------------------------------SHELLHOOK                                                        	;{
InitializeShellHook:						                                                		;{

		DllCall( "RegisterShellHookWindow", UInt, GetDec(hBO))
		MsgNum := DllCall( "RegisterWindowMessage", Str,"SHELLHOOK" )
		OnMessage( MsgNum, "ShellMessage" )

return
;}

ShellMessage( wParam, lParam ) {					    	;{

		global ShTitle, SHProcName

		If (wParam not in 1,2) && !InStr(WinGetClass(lParam), PDFReaderWinClass)
				return

		WinGetTitle	, ShTitle									, % "ahk_id  " ShHookHwnd:= GetHex(lParam)
		WinGet			, SHProcName, ProcessName	, % "ahk_id  " ShHookHwnd
		SetTimer, ShMessage_WinHandler, -0

}

ShMessage_WinHandler:																						;{

		If !PdfViewer.HasKey(ShHookHwnd)
		{
				NPRWin := ShHookHwnd
				PdfViewer[ShHookHwnd] := ShTitle
				FileAppend, "PatID: " AlbisAktuellePatID() ", geöffnet: " ShHookHwnd ", " ShTitle, %A_ScriptDir%\hook.txt
		}
		else
		{
				FileAppend, "PatID: " AlbisAktuellePatID() ", geschlossen: " ShHookHwnd ", " ShTitle, %A_ScriptDir%\hook.txt
		}

return
;}
;}
;}

;---------------------------------------INI READ ALS FUNKTION                                     	;{

BuildScanPoolDataObject() {

		obj:= Object(), value:= []

	  ; pdftk.exe - für Schreiben von neuen Metadaten in die PDF-Datei (Patientendaten)
		inistrings:= AddendumDir "\Addendum.ini, ScanPool  	, PDFtkPath				," 		"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool  	, xpdfPath					,"		"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool  	, Signatures				, 0"	"`n"
		inistrings.= AddendumDir "\Addendum.ini, Addendum	, AddendumDBPath	,"    	"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool    	, BefundOrdner			,"    	"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool    	, PDFReaderName		,"    	"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool    	, PDFReaderWinClass,"    	"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool    	, PDFReaderFullPath	,"    	"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool    	, SignatureClients		,"    	"`n"
		inistrings.= AddendumDir "\Addendum.ini, ScanPool    	, AutoCloseReader		, 1"

	  ;----------------------------------------------------------------------------------------------------------------------------------------------
	  ; Erstellung des Objektes
	  ;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Loop, Parse, inistrings, `n, `r
		{
				args	:= StrSplit(A_LoopField, ",", "`r`t`n" A_Space)
				temp	:= IniRead(args[1], args[2], args[3], args[4])
				If args[5] = ""
					obj[(args[3])]:= StrReplace(temp, "%AddendumDir%", AddendumDir)
				else
					obj[(args[5])]:= StrReplace(temp, "%AddendumDir%", AddendumDir)
		}

		obj.ContentPath:= obj.AddendumDBPath "\Befunde\Daten"
		obj.PdfIndexFile	:= obj.BefundOrdner "\PdfIndex.txt"
		obj.TempPath   	:= obj.BefundOrdner "\Temp"
		obj.BackupPath	:= obj.BefundOrdner "\Backup"

	  ;----------------------------------------------------------------------------------------------------------------------------------------------
	  ; ein paar Sonderbehandlungen zuerst danach lasse ich das Objekt automatisch erstellen
	  ;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If InStr(obj.PDFtkPath, "Error")		{
				MsgBox, 1, Addendum für Albis on Windows,
				(LTrim
				Programm zum erneuern von PDF-Metadaten!

				.....
				)
				IfMsgBox, No
					obj.PDFtkPath:= ""
		}

		If !InStr(FileExist(obj.ContentPath), "D")		{
				FileCreateDir, % obj.ContentPath
				If ErrorLevel
						MsgBox,, % "Addendum für AlbisOnWindows", % "Das Verzeichnis für die Datenbank: `n" obj.ContentPath "\Befunde\Daten" "`nkonnte nicht angelegt werden."
		}
		;}

return obj
}

IniRead(iniPath, iniSection, iniKey, iniDefault:="") {
	IniRead, val, % inipath, % iniSection, % inikey, % iniDefault
return val
}

;}

;}

;---------------------------------------ICON                                                                  	;{

Create_ScanPool_ico(NewHandle := False) {
Static hBitmap := Create_ScanPool_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGAAAABgAAAAAAAAAAAAAAD/////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo//oo//oo//oo//oo//oo/8oI3/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRhyRzX/oo//oo9sQzFCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo9jPSxCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdFKhmFU0HbjHj/oo//oo//oo//oo//oo+BUT9EKhlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdJLBySXErnkn//oo//oo//oo//oo//oo//oo//oo//oo//oo+OWUhILBpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdQMSCfZFPvmIX/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+cY1BPMB9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdXNSSsbVr2nIn/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+qa1lWNSRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdiPCu5dWP7oIz/oo//oo//oo//oo//oo//oo//oo//oo//oo+laFalaFb/oo//oo//oo//oo//oo//oo//oo//oo//oo/7n422c2BfOylCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdtRDPFfGr+oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo93SzlDKRhCKBdCKBdDKRh9Tj3/oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY7Ce2hqQjBCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBd3SjnRhXL/oo//oo//oo//oo//oo//oo//oo//oo/+oY7/oo9qQjBCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBduRDP/oo/+oY7/oo//oo//oo//oo//oo//oo//oo//oo/Qg3F3SjlCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdbOSfWiHX/oo//oo//oo//oo//oo//oo//oo//oo/7n43/oo9fOylCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPiz/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9WNSRCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdQMSC8d2X9oI3/oo//oo//oo//oo//oo/1nIn/oo9VNSRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdZNyX/oo/4nor/oo//oo//oo//oo//oo/7n43/oo9MLh1CKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdaOCevblz/oo//oo//oo+dY1FPMB9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdXNiT/oo//oo9YNiVCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdRMiCiZlP/oo//oo/0moj/oo9WNSRCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9DKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdpQS//oo/9oI3/oo//oo/9oI3/oo9mPy5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdFKhn/oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd1STj/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo90SDdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdEKhmCUUD/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+AUD9DKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdJLRuQW0n/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+NWEdHKxpCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdOLx//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/Beme9d2X/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+ZYE5MLx5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdVNST/oo/1nIn/oo//oo//oo//oo//oo//oo//oo//oo//oo+UXktMLh1CKBdCKBdMLh2XX03/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+laFZTMyJCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdeOin/oo/7n43/oo//oo//oo//oo//oo//oo//oo//oo//oo+FVEJFKhlCKBdCKBdCKBdCKBdCKBdCKBdGKxmKV0T/oo//oo//oo//oo//oo//oo//oo//oo//oo/5noz/oo9bOSdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdSMiH/oo/+oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo96TDpDKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRh+Tj3/oo//oo//oo//oo//oo//oo//oo//oo//oo/9oI3/oo9KLRxCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdaOCf/oo//oo//oo//oo//oo//oo//oo/+oY7/oo9sQzJCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdxRjT/oo/+oY7/oo//oo//oo//oo//oo//oo//oo9TMyJCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdyRzX/oo/+oY7/oo/5noz/oo9hPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdKLRz/oo//oo9EKhlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdkPi3/oo/7oIz/oo/+oY7/oo9uRDNCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo9MLh1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdUMyL/oo//oo//oo//oo//oo//oo9SMyFCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdPMB//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBddOSj/oo//oo//oo//oo//oo//oo//oo//oo/4nYv/oo9aOCdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdoQC//oo/9oI3/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/8oI3/oo9kPi1CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd0SDf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY7/oo9xRjRCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdDKRiAUD7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo99TjxDKRhCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdHKxqNWEf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+KV0RGKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdMLx6ZYE7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+UXktJLRtCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdKLRz/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9ILBpCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdgPCr/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9cOShCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdJLRuOWUf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo+JVkRGKxlCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdDKRiAUD//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo99TjxDKRhCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBd0SDf/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/+oY7/oo9xRzVCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdoQC//oo/+oY7/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo/9oI3/oo9mPi5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBddOSj/oo/6n4z/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9bOSdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdVNCP/oo/0moj/oo//oo//oo//oo//oo//oo//oo//oo//oo//oo9SMyFCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdMLx6aYU//oo//oo//oo//oo//oo//oo//oo9MLx5CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdILBqNWEf/oo//oo//oo9GKxlCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo//oo//oo//oo//oo//oo//oo/8oI3/oo9CKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBf/oo/8oI3/oo//oo//oo//oo//oo//oo//oo//oo//oo9iPCtCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdCKBdjPSz/oo//oo//oo//oo//oo//oo//////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//////////////oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo//oo////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
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

LineDelete(V, L, R := "", O := "", ByRef M := "") {	                                       	;-- Deletes a specific line or a range of lines from a variable containing one or more lines of text.

	; DESCRIPTION of function LineDelete() - see AHK.Rare on GitHub (Ixiko)

	T := StrSplit(V, "`n").MaxIndex()
	If(L > 0 && L <= T && (O = "" || O = "B")){
		V := StrReplace(V, "`r`n", "`n"), S := "`n" V "`n"
		P := (O = "B") ? InStr(S, "`n",,, L + 1)
		   : InStr(S, "`n",,, L)
		M := (R <> "" && R > 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 2 + (R - L)) - P - 1)
		   : (R <> "" && R < 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 3 + (R - L + T)) - P - 1)
		   : (R <> "" && R > 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, R - L) - P - 1)
		   : (R <> "" && R < 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, 1 + (R - L + T)) - P - 1)
		   : SubStr(S, P + 1, InStr(S, "`n",, P, 2) - P - 1)
		X := SubStr(S, 1, P - 1) . SubStr(S, P + StrLen(M) + 1), X := SubStr(X, 2, -1)
	}
	Else If(L < 0 && L >= -T && (O = "" || O = "B")){
		V := StrReplace(V, "`r`n", "`n"), S := "`n" V "`n"
		P := (R <> "" && R < 0 && O = "" ) ? InStr(S, "`n",,, R + T + 1)
		   : (R <> "" && R > 0 && O = "" ) ? InStr(S, "`n",,, R)
		   : (R <> "" && R < 0 && O = "B") ? InStr(S, "`n",,, R + T + 2)
		   : (R <> "" && R > 0 && O = "B") ? InStr(S, "`n",,, R + 1)
		   : InStr(S, "`n",,, L + T + 1)
		M := (R <> "" && R < 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 2 + (L - R)) - P - 1)
		   : (R <> "" && R > 0 && O = "" ) ? SubStr(S, P + 1, InStr(S, "`n",, P, 3 + (T - R + L)) - P - 1)
		   : (R <> "" && R < 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, (L - R)) - P - 1)
		   : (R <> "" && R > 0 && O = "B") ? SubStr(S, P + 1, InStr(S, "`n",, P, 1 + (T - R + L)) - P - 1)
		   : SubStr(S, P + 1, InStr(S, "`n",, P, 2) - P - 1)
		X := SubStr(S, 1, P - 1) . SubStr(S, P + StrLen(M) + 1), X := SubStr(X, 2, -1)
	}
	Return X
}

PrintArr(Arr, Option := "w800 h500, Object name", GuiNum:= 90		                                    	;-- show values of an array in a listview gui for debugging
, Colums:= "Nr|PatID|", statustext:="Gesamtzahl") {
	static initGui:= []
	Option:= StrSplit(Option, ",")

    for index, obj in Arr {
        if (A_Index = 1) {
            for k, v in obj {
                Columns .= k "|"
                cnt++
            }
			If !(init[GuiNum]) {
					Gui, %GuiNum%: Margin, 0, 0
					Gui, %GuiNum%: Add, ListView, % Option[1], % Columns
					Gui, %GuiNum%: Add, Statusbar
					init[GuiNum]:= 1
			}
        }
        RowNum := A_Index
        Gui, %GuiNum%: default
        LV_Add("")
        for k, v in obj {
            LV_GetText(Header, 0, A_Index)
            if (k <> Header) {
                FoundHeader := False
                loop % LV_GetCount("Column") {
                    LV_GetText(Header, 0, A_Index)
                    if (k <> Header)
                        continue
                    else {
                        FoundHeader := A_Index
                        break
                    }
                }
                if !(FoundHeader) {
                    LV_InsertCol(cnt + 1, "", k)
                    cnt++
                    ColNum := "Col" cnt
                } else
                    ColNum := "Col" FoundHeader
            } else
                ColNum := "Col" A_Index
            LV_Modify(RowNum, ColNum, (IsObject(v) ? "Object()" : v))
        }
    }
    loop % LV_GetCount("Column")
        LV_ModifyCol(A_Index, "AutoHdr")
	SB_SetText("   " LV_GetCount() " " statustext)

    Gui, %GuiNum%: Show,, % Option[2]
	return RowNum
}

LV_EX_FindString(HLV, Str, Start := 0, Partial := False) {                       				;-- gibt die Zeilennummer zurück in welchem sich der gesuchte Text befindet

   ; LVM_FINDITEM -> http://msdn.microsoft.com/en-us/library/bb774903(v=vs.85).aspx
   Static LVM_FINDITEM := A_IsUnicode ? 0x1053 : 0x100D ; LVM_FINDITEMW : LVM_FINDITEMA
   Static LVFISize := 40

   VarSetCapacity(LVFI, LVFISize, 0) ; LVFINDINFO
   Flags := 0x0002 ; LVFI_STRING
   If (Partial)
      Flags |= 0x0008 ; LVFI_PARTIAL
   NumPut(Flags	, LVFI, 0        	, "UInt")
   NumPut(&Str	, LVFI, A_PtrSize	, "Ptr")
   SendMessage, % LVM_FINDITEM, % (Start - 1), &LVFI,, % "ahk_id " HLV

Return (ErrorLevel > 0x7FFFFFFF ? 0 : ErrorLevel + 1)
}

LV_GetItemState(HLV, Row) {                                                                         	;-- den Status einer Listviewzeile ermitteln

   Static LVM_GETITEMSTATE := 0x102C
   Static LVIS := {Cut: 0x04, DropHilited: 0x08, Focused: 0x01, Selected: 0x02, Checked: 0x2000}
   Static ALLSTATES := 0xFFFF ; not defined in MSDN
   SendMessage, % LVM_GETITEMSTATE, % (Row - 1), % ALLSTATES, , % "ahk_id " . HLV
   If (ErrorLevel + 0) {
      States := ErrorLevel
      Result := {}
      For Key, Value In LVIS
         Result[Key] := !!(States & Value)
      Return Result
   }

 Return False
}

LV_GetItemState2(HLV, Row) {                                                                        	;-- wie darüber, da LV_GetItemState nicht immer funktionierte

	; letzte Änderung 01.12.2020

	Static LVM_GETITEMSTATE := 0x102C
	Static LVIS1 := {Cut: 0x04, DropHilited: 0x08, Focused: 0x01, Selected: 0x02, Checked: 0x2000}
	Static LVIS2 := {"0x04":"Cut", "0x08":"DropHilited", "0x01":"Focused", "0x02":"Selected", "0x2000":"Checked"}
	Static ALLSTATES := 0xFFFF ; not defined in MSDN

	SendMessage, % LVM_GETITEMSTATE, % (Row - 1), % ALLSTATES, , % "ahk_id " HLV

	If (ErrorLevel + 0) {

		States := Format("0x{:x}", ErrorLevel)
		For Key, Value in LVIS2
			If (States = Value)
				(result := Value), break           ; #### funktioniert das?

		Return result.= " (" Key ")"
   }

 Return False
}

LV_ItemText(hListView, iItem, iSubItem, ByRef lpString, nMaxCount) {               	;--

        ;const
        LVNULL                            	:= 0
        PROCESS_ALL_ACCESS 	:= 0x001F0FFF
        INVALID_HANDLE_VALUE	:= 0xFFFFFFFF
        PAGE_READWRITE         	:= 4
        FILE_MAP_WRITE             	:= 2
        MEM_COMMIT             	:= 0x1000
        MEM_RELEASE               	:= 0x8000
        LV_ITEM_mask              	:= 0
        LV_ITEM_iItem               	:= 4
        LV_ITEM_iSubItem         	:= 8
        LV_ITEM_state                 	:= 12
        LV_ITEM_stateMask       	:= 16
        LV_ITEM_pszText              	:= 20
        LV_ITEM_cchTextMax      	:= 24
        LVIF_TEXT                     	:= 1
        LVM_GETITEM                	:= 0x1005
        SIZEOF_LV_ITEM             	:= 0x28
        SIZEOF_TEXT_BUF         	:= 0x104
        SIZEOF_BUF                     := 0x120
        SIZEOF_INT                     	:= 4
        SIZEOF_POINTER             	:= 4

        ;var
        result        	:= 0
        hProcess    	:= LVNULL
        dwProcessId	:= 0

        if (lpString <> LVNULL) && (nMaxCount > 0)        {

            DllCall("lstrcpy", "Str",lpString, "Str","")
            DllCall("GetWindowThreadProcessId", "UInt", hListView, "UIntP", dwProcessId)
            hProcess := DllCall("OpenProcess", "UInt", PROCESS_ALL_ACCESS, "Int", false, "UInt", dwProcessId)
            if (hProcess <> LVNULL)  {

                ;var
                lpProcessBuf  	:= LVNULL
                hMap            	:= LVNULL
                hKernel         	:= DllCall("GetModuleHandle", Str,"kernel32.dll", UInt)
                pVirtualAllocEx	:= DllCall("GetProcAddress", UInt,hKernel, Str,"VirtualAllocEx", UInt)

                if (pVirtualAllocEx == LVNULL) {

                    hMap := DllCall("CreateFileMapping", "UInt",INVALID_HANDLE_VALUE, "Int",LVNULL, "UInt",PAGE_READWRITE, "UInt",0, "UInt",SIZEOF_BUF, UInt)
                    if (hMap <> LVNULL)
                        lpProcessBuf := DllCall("MapViewOfFile", "UInt",hMap, "UInt",FILE_MAP_WRITE, "UInt",0, "UInt",0, "UInt",0, "UInt")

                }
                else {

                    lpProcessBuf := DllCall("VirtualAllocEx", "UInt",hProcess, "UInt",LVNULL, "UInt",SIZEOF_BUF, "UInt",MEM_COMMIT, "UInt",PAGE_READWRITE)

                }

                if (lpProcessBuf <> LVNULL)   {

                    ;var
                    VarSetCapacity(buf, SIZEOF_BUF, 0)

                    InsertInteger(LVIF_TEXT, buf, LV_ITEM_mask, SIZEOF_INT)
                    InsertInteger(iItem, buf, LV_ITEM_iItem, SIZEOF_INT)
                    InsertInteger(iSubItem, buf, LV_ITEM_iSubItem, SIZEOF_INT)
                    InsertInteger(lpProcessBuf + SIZEOF_LV_ITEM, buf, LV_ITEM_pszText, SIZEOF_POINTER)
                    InsertInteger(SIZEOF_TEXT_BUF, buf, LV_ITEM_cchTextMax, SIZEOF_INT)

                    if (DllCall("WriteProcessMemory", "UInt",hProcess, "UInt",lpProcessBuf, "UInt",&buf, "UInt",SIZEOF_BUF, "UInt",LVNULL) <> 0)
                        if (DllCall("SendMessage", "UInt",hListView, "UInt",LVM_GETITEM, "Int",0, "Int",lpProcessBuf) <> 0)
                            if (DllCall("ReadProcessMemory", "UInt",hProcess, "UInt",lpProcessBuf, "UInt",&buf, "UInt",SIZEOF_BUF, "UInt",LVNULL) <> 0)  {
                                DllCall("lstrcpyn", "Str",lpString, "UInt",&buf + SIZEOF_LV_ITEM, "Int",nMaxCount)
                                result := DllCall("lstrlen", "Str",lpString)
                            }
                }

                if (lpProcessBuf <> LVNULL)
                    if (pVirtualAllocEx <> LVNULL)
                        DllCall("VirtualFreeEx", "UInt",hProcess, "UInt",lpProcessBuf, "UInt",0, "UInt",MEM_RELEASE)
                    else
                        DllCall("UnmapViewOfFile", "UInt",lpProcessBuf)

                if (hMap <> LVNULL)
                    DllCall("CloseHandle", "UInt",hMap)

                DllCall("CloseHandle", "UInt",hProcess)
            }

        }

return result
}
;{Sub	for LV_GetItemText and LV_GetText

ExtractInteger(ByRef pSource, pOffset = 0, pIsSigned = false, pSize = 4) {

   SourceAddress := &pSource + pOffset  ; Get address and apply the caller's offset.
   result := 0  ; Init prior to accumulation in the loop.
   Loop % pSize { ; For each byte in the integer:
      result := result | (*SourceAddress << 8 * (A_Index - 1))  ; Build the integer from its bytes.
      SourceAddress += 1  ; Move on to the next byte.
   }
   if (!pIsSigned OR pSize > 4 OR result < 0x80000000)
      return result  ; Signed vs. unsigned doesn't matter in these cases.
   ; Otherwise, convert the value (now known to be 32-bit) to its signed counterpart:
   return -(0xFFFFFFFF - result + 1)
}

InsertInteger(pInteger, ByRef pDest, pOffset = 0, pSize = 4) {
; To preserve any existing contents in pDest, only pSize number of bytes starting at pOffset
; are altered in it. The caller must ensure that pDest has sufficient capacity.

   mask := 0xFF  ; This serves to isolate each byte, one by one.
   Loop % pSize {  ; Copy each byte in the integer into the structure as raw binary data.
      DllCall("RtlFillMemory"	, "UInt", &pDest + pOffset + A_Index - 1, "UInt", 1              	; Write one byte.
			                        	, "UChar", (pInteger & mask) >> 8 * (A_Index - 1))              	; This line is auto-merged with above at load-time.
      mask := mask << 8  ; Set it up for isolation of the next byte.
   }

}

;}

;----------------------------------------------------------------------------------------------------------------------------------------------
; Namen extrahieren
;----------------------------------------------------------------------------------------------------------------------------------------------
ExtractNamesFromFileName(Pat) {					                      	;-- die Namen müssen die ersten zwei Worte im FileNamen der PDF-Datei sein

	Name:=[]
	;diese Funktion basiert auf der manuellen Erstellung des Dateinamens während des Scanvorganges, es ist sehr kompliziert dem Computer aus OCR Text nach Patientennamen zu durchsuchen

	; Ermitteln des Patientennamen aus dem PDF-Dateinamen
		;Pat:= RegExReplace(Pat, "\.|;|,", A_Space) 			;entferne alle Punkt, Komma und Semikolon
		Pat:= StrReplace(Pat, "Dr.", " ")
		Pat:= StrReplace(Pat, "Prof.", " ")
		Pat:= StrReplace(Pat, ",", " ")				;(^,+|,+(?= )|,+$)	;
		Pat:= StrReplace(Pat, ".", " ")				;(^,+|,+(?= )|,+$)	;
		Pat:= StrReplace(Pat, ";", " ")				;\.|;|,
		Pat:= RegExReplace(Pat, "^Frau\s", " ")
		Pat:= RegExReplace(Pat, "\s*Frau\s", " ")
		Pat:= RegExReplace(Pat, "^Herr\s", " ")
		Pat:= RegExReplace(Pat, "\s*Herr\s", " ")
	  ; entfernt alle aufeinander folgenden Spacezeichen! https://autohotkey.com/board/topic/26575-removing-redundant-spaces-code-required-for-regexreplace/
		Pat:= RegExReplace(RegExReplace(Pat, "\s{2,}", ""),"(^\s*|\s*$)")
		Pat:= StrSplit(Pat, A_Space)
		Name[1]:= Pat[1] ", " Pat[2]
		Name[2]:= Pat[2] ", " Pat[1]
		Name[3]:= Pat[1]
		Name[4]:= Pat[2]

return Name
}

;}

;{11. TrayIcon - Prozesse und Hotkey - Labels

Hotkeys: 								;{			Hotkey Gui

		;--Anzeige aller benutzten Hotkeys mit Beschreibung durch Auslesen des Skriptes
		script:= A_ScriptDir . "\" . A_ScriptName


return
;}

ScriptObjects:						;{ 		zeigt Variablen die nicht über die normale Anzeige ausgegeben werden können

	PrintArr(ReStructure(ScanPool), "w900 h500, Pdf in der Datenbank", 90, "Dateiname|Größe|Seiten", "Gesamtseitenzahl")

return

ReStructure(ScanPool) {			; baut den ScanPool Array so um das er mit PrintArr() angezeigt werden kann

	reStructured:= Object()

	For key, val in ScanPool
		 Restructured[key]:= {"filename": StrSplitEx(val,1), "filesize": StrSplitEx(val,2), "pages": StrSplitEx(val,3)}

return Restructured
}

;}

ShowGui: 								;{ 		Wiederanzeigen des Fenster wenn es minimiert wurde

	If minimized {
		;OnMessage Farbanzeige wieder einschalten
			;OnMessageStatus(1)
		;Farbanzeige jetzt auffrischen
			LVA_Refresh("ListView1")
		;Timer wieder herstellen
			SetTimer, BOProcessTipHandler, 40
		;Gui auf Default setzen, Listview1 auf Default setzen
			Gui, BO: Default
			;Gui, BO: Listview, ListView1
		;erst jetzt das Gui wieder anzeigen
			WinShow, ahk_id %hBO%
			WinMaximize, ahk_id %hBO%
			If WinExist("ahk_id " . hProgress)
					Progress, Show
			If WinExist("ahk_id " . hBOPV)
					WinShow, ahk_id %hBOPV%
		;Tray Eintrag Fenster maximieren auf minimieren ändern
			Menu, Tray, Disable, Fenster anzeigen
	}

return
;}

SkriptReload: 						;{			Skript neu starten (ist kein reload!)

		Script:= A_ScriptDir "\" A_ScriptName
		scriptPID := DllCall("GetCurrentProcessId")
		run,  Autohotkey.exe /f "%AddendumDir%\include\RestartScript.ahk" "%Script%" "2" "%A_ScriptHwnd%" "%scriptPID%"
		ExitApp

return
;}

ScriptVars:								;{
	ListVars
return
;}

ScriptLines:							;{
	If !A_LineNumber = ""
			ListLines
return
;}

NoNo:                              	;{     	Anzeige von Programminfo's

		FileRead, rawlines, %A_ScriptDir%\%A_ScriptName%
		rIndex:=0
		Loop, Parse, rawlines, `n
		{
			If Instr(A_LoopField, ";###") {
				rIndex++
				continue
			}
			If (rIndex=2)
				break

			line:= StrReplace(A_LoopField, ";", "")
			line:= Trim(StrReplace(line, "-", ""))
			B.= line "`n"

		}

		Gui UberPt: New, -Caption Border HWNDhUberPt
		Gui UberPt: Font, S11 CDefault, Futura Bk Bt
		Gui UberPt: Add, Edit, r20 Center ReadOnly -WantCtrlA -WantReturn -Wrap HWNDhUberTextPT, %B%
		Gui UberPt: Add, Button, ys vUberOk gUberEnde, OK
		Gui UberPt: Show, AutoSize, ScanPoolInfo
		ControlFocus, UberOk, ahk_id %hUberPt%
		SendInput, {Left}
		DllCall("HideCaret","Int",hUberTextPT)

return

UberEnde:

		If (A_GuiControl=="UberOk") {
			B:=""									;--Variable B wieder leeren um RAM zu sparen
			Gui UberPt: Destroy
		}

return


	;}

PdfSignatur:							;{

	FoxitID:= WinExist("A")
	;FoxitReader_SignaturSetzen(ReaderID, PDFReaderWinClass)
	FoxitInvoke("Place_Signature", FoxitID)

return
;}

;}

;{12. Includes

	#Include %A_ScriptDir%\..\..\..\include\Addendum_Functions.ahk
	#Include %A_ScriptDir%\..\..\..\include\Addendum_DB.ahk
	#Include %A_ScriptDir%\..\..\..\include\Addendum_FoxitReader.ahk
	#Include %A_ScriptDir%\..\..\..\include\Addendum_Protocol.ahk
	#Include %A_ScriptDir%\..\..\..\include\Addendum_PdfHelper.ahk
	#include %A_ScriptDir%\..\..\..\Include\Gui\PraxTT.ahk

	#Include %A_ScriptDir%\..\..\..\lib\ACC.ahk
	#Include %A_ScriptDir%\..\..\..\lib\class_JSONahk
	#Include %A_ScriptDir%\..\..\..\lib\ini.ahk
	#Include %A_ScriptDir%\..\..\..\lib\FindText.ahk
	#include %A_ScriptDir%\..\..\..\lib\GDIP_All.ahk
	#include %A_ScriptDir%\..\..\..\lib\LVA.ahk
	#include %A_ScriptDir%\..\..\..\lib\Math.ahk
	#Include %A_ScriptDir%\..\..\..\lib\ObjDump.ahk
	#Include %A_ScriptDir%\..\..\..\lib\RemoteBuf.ahk
	#Include %A_ScriptDir%\..\..\..\lib\_Struct.ahk
	#Include %A_ScriptDir%\..\..\..\lib\LV.ahk
	#Include %A_ScriptDir%\..\..\..\lib\sizeof.ahk
	#Include %A_ScriptDir%\..\..\..\lib\ObjTree.ahk
	#Include %A_ScriptDir%\..\..\..\lib\attach.ahk
	#Include %A_ScriptDir%\..\..\..\lib\sift.ahk
	;#include %A_ScriptDir%\..\..\..\include\RamDisk\RamDrive.ahk

;}

;----------------------------------------------------------------------------------------------------------------------------------------------------------------------
; PDF DATENBANK / SCANPOOL
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------;{
ScanPoolArray(cmd, param="", opt="") {                    		;-- verarbeitet den Datei-Array der die Datei-Informationen des Befundordners bereit hält

/*							BESCHREIBUNG

		ein Array mit dem Namen 'ScanPool' muss superglobal gemacht werden am Anfang aufrufenden Skriptes
			WICHTIG!!:	egal was man mit dem Array dann macht >>> niemals innerhalb dieser Funktion löschen oder neu initialisieren!
								wie es gemacht werden sollte und zwar nur ausserhalb dieser Funktion:
									• ScanPool := ""        	- leert den Array
									• ScanPool := Array()	- erstellt das Array neu!
								führt man die oben erwähnten Zuordnungen innerhalb dieser Funktion aus, ist ScanPool zwar weiterhin superglobal angelegt
								, aber ausserhalb dieser Funktion wird der Array leer sein, selbst wenn man ihn wieder befüllt. Scheinbar existieren dann zwei Variablen
								anstatt nur einer Variable!!

		Beschreibung: param
		was param beinhalten kann, ist vom übergebenen Befehl (cmd) abhängig, z.B. ScanPoolArray("GetCell", "5|3") ; ermittle den Inhalt der 3.Spalte der 5.Zeile, weiteres siehe folgende Zeilen:
		cmd:    "Delete"			- löschen einer Datei innerhalb des ScanPool-Array, param - Name der zu entfernenden Datei
					"Sort"   			- sortiert die Dateien im Array und somit auch für die Anzeige im Listviewfenster
					"Load"			- lädt aus einer Datei den zuvor indizierten Ordner samt entsprechender Daten
					"Save"			- speichert die Daten auf Festplatte in eine Textdatei (mit '|' getrennte Speicherung der einzelnen Felder Dateiname,Größe,Seiten,Signiert ja=1;nein=Feld bleibt leer)
					"Rename" 		- Umbenennen einer Datei innerhalb des ScanPool-Array
					"ValidKeys"	- zählt die vorhandenen Datensätze, da auch manchmal leere Datensätze gespeichert wurden, werden nur nicht leere gezählt
					"Find"	    	- sucht nach einem Dateinamen und gibt den Index zurück
					"CountPages"- ermittelt die Gesamtzahl aller Seiten der Pdf-Dateien im BefundOrdner
					"Reports"     	- fuzzy Personennamen Suche, param - Nachname Vorname (des Patienten)
					"Signed"     	- erstellt einen Array in dem alle im Befundordner signierten Dateien aufgelistet sind (kontrolliert nicht ob neue Dateien hinzugefügt wurden)
					"NotSigned"	- wie "Signed" nur alle bekannt unsignierten um diese auf eine vorhandene Signatur zu prüfen
*/

	static Loaded := false
	static PdfIndexFile

	newPDF := 0, col :=[], res := 0, allfiles := ""
	FileEncoding, UTF-8

	;diese Zeilen sichern ab das zuerst der ScanPool-Array erstellt wird bevor die anderen Befehle aufgerufen werden können
	If !Loaded && (StrLen(param) = 0) 	{
		MsgBox, 0, % "Funktion ScanPoolArray", % "Dieser Funktion muss als erstes per 'Load' command`nder Pfad zur PdfIndex.txt Datei übergeben werden."
		ExitApp
		;Loaded:= ScanPoolArray("Load")
	}
  ;--------------------------------------------------------------------- Befehle -----------------------------------------------------------------------------------

	If Instr(cmd, "Delete") || Instr(cmd, "Remove")	{  	;param: gesuchter Dateiname			           			, Rückgabe: Wert       	- Seitenanzahl der entfernten Pdf-Datei

			param := RegExReplace(param, "\.pdf$", "") ".pdf" ;muss man nicht immer dran denken die Dateiendung zu übergeben
			For key, val in ScanPool
				If Instr(val, param)	{
					delpages := StrSplitEx(val, 3)
					ScanPool.Delete(key)
					break
				}

			FileCount := CountValidKeys(ScanPool)						;ist sozusagen dann der "NEUE" MaxIndex()
			return delpages
	}
	else If Instr(cmd, "Load")         	{                         	;param: Pfad zur PDFIndex.txt Datei                      	, Rückgabe: Wert       	- Gesamtzahl der Befunde in der pdfIndex.txt Datei

		If FileExist(param)	{
			Loaded      	:= true											        	;zur Überprüfung - Load muss vor allen anderen Befehlen als erstes stattgefunden haben
			PdfIndexFile	:= param
			allfiles       	:= FileOpen(PdfIndexFile, "r").Read()
			Sort, allfiles
			ScanPool   	:= StrSplit(allfiles, "`n", "`r")
			return ScanPool.MaxIndex()
		}
		else
			return 0

	}
	else If Instr(cmd, "Rename")    	{                         	;param: Original-Dateiname, opt: neuer Name     	, Rückgabe: Wert       	- ist der Index im ScanPool-Array

			For key, val in ScanPool
				If Instr(val, param)		{
					col := StrSplit(ScanPool[key], "|")
					ScanPool[key] := opt "|" col[2] "|" col[3]
					return key
				}
			return 0
	}
	else If Instr(cmd, "Save")         	{                         	;param: und opt - unbenutzt					            	, Rückgabe: ErrorLevel	- erfolgreich = 1, Speicherung nicht möglich = 0

			File := FileOpen(PdfIndexFile, "w", "UTF-8")
			For key, val in ScanPool
				allfiles .= (StrLen(Trim(val)) >0) ? val "`n" : ""

				;~ If (StrLen(Trim(val)) >0)
					;~ allfiles .= val "`n"

			allfiles := RTrim(allfiles, "`n")
			File.Write(allfiles)
			File.Close()
			If !ErrorLevel
				return 1
			else
				return 0

	}
	else if Instr(cmd, "Sort")           	{                         	;param: und opt - unbenutzt			            			, Rückgabe: ohne      	- der ScanPool-Array wird sortiert

			For key, val in ScanPool
				If (Trim(val) != "")
					allfiles .= val "`n"

			Sort, allfiles
			ScanPool := StrSplit(RTrim(allfiles, "`n"), "`n")

	}
	else if Instr(cmd, "ValidKeys")   	{                         	;param: und opt - unbenutzt				           			, Rückgabe: Wert       	- Anzahl der im Array gespeicherten Pdf-Dateien
		return CountValidKeys(ScanPool)
	}
	else If InStr(cmd, "Find")           	{                         	;param: gesuchter Dateiname		            			, Rückgabe: Wert       	- ist der Indexwert oder KeyIndex im ScanPool-Array

			For key, val in ScanPool
				If Instr(val, param)
			   		return key

			return ""
	}
	else If Instr(cmd, "CountPages")  {                        	;param: und opt - unbenutzt			           				, Rückgabe: Wert       	- Gesamtzahl aller Seiten in den Pdf-Dateien des Befundordners

			tpgs := 0
			For key, val in ScanPool
				tpgs += !StrSplitEx(val, 3) ? 0 : StrSplitEx(val, 3)

			return tpgs
	}
	else If InStr(cmd, "Reports")       	{                      	;param: Patientenname (Nachname, Vorname)     	, Rückgabe: Array      	- Pdf Befunde mit passendem Patientennamen

			Reports := []

			RegExMatch(StrReplace(param, "-"), "(?P<Nachname>[\w\p{L}]+)[\,\s]+(?P<Vorname>[\w\p{L}]+)", Such)
			SuchName := SuchNachname SuchVorname

			For key, val in ScanPool	{				;wenn keine PatID vorhanden ist, dann ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)

				if (StrLen(val) = 0)
					continue

				filename := StrReplace(StrSplit(val, "|").1, ".pdf")
				RegExMatch(StrReplace(filename, "-"), "(?P<Nachname>[\w\p{L}]+)[\,\s]+(?P<Vorname>[\w\p{L}]+)", pdf)

				a := StrDiff(SuchName, pdfNachname pdfVorname)
				b := StrDiff(SuchName, pdfVorname pdfNachname)

				If ((a < 0.12) || (b< 0.12))
					Reports.Push(filename)

			}

			return Reports
	}
	else If InStr(cmd, "Refresh")       	{                       	;param: unbenutzt                                                	, Rückgabe: Wert        	- Anzahl neuer Funde

		If InStr(A_ScriptName, "ScanPool")
			newPDF := RefreshPdfIndex(BefundOrdner)
		else
			newPDF := RefreshPdfIndex(Addendum.BefundOrdner)

		If (newPDF > 0) {
			anzeigetext := newpdf = 1 ? "1 Dokument hinzugefügt." : newpdf " Dokumente hinzugefügt."
			PraxTT(anzeigetext, "4 0")
		}
		return newPDF

	}
	else If InStr(cmd, "Signed")      	{                      	;param: und opt - unbenutzt						           	, Rückgabe: Array       	- signierte Befunde

			Signed := []
			for key, val in ScanPool
				If (StrSplitEx(val, 4) = 1)
					Signed.Push(StrSplitEx(val, 1))

			return Signed
	}
	else If InStr(cmd, "NotSigned") 	{                       	;param: Array (signierter Befunde "Signed")           	, Rückgabe: Array       	- unsignierte Befunde

			if !IsObject(param)
					return ScanPool

			NotSigned := [], BOidx := 0
			For key, val in ScanPool
				allfiles .= val "`n"
			For key, val in param
				allfiles .= val "`n"
			allfiles := RTrim(allfiles, "`n")

			for key, val in ScanPool
				If InStr(val, param)
					If (StrSplitEx(val, 4) = 1)
						SignedArr.Push(StrSplitEx(val, 1))

			return SignedArr
	}

return "EndOfFunction"
}

CountValidKeys(arr) {                                                     	;-- zählt die gültigen Einträge im Array

	counter:=0, notValid:=""
	For key, val in arr
		If !(val = "")
			counter++

return counter
}

ReadDir(dir, ext) {                                                           	;-- liest ein Verzeichnis ein, ext=Dateiendung
	tlist := Array()
	Loop, Files, % dir "\*." ext
		tlist.Push(A_LoopFileName)
return tlist
}

ReadPdfIndex(PdfIndexFile) {	                                        	;-- erstellt das ScanPool Object### geändert jetzt fehlerhaft - ist noch für scanpool.ahk

	; ein Teil der Variablen sind globale Variablen
		PageSum	:= FileCount	:= 0
		allfiles   	:= ""
		tmpPool	:= Array()

		If (FileCount := ScanPoolArray("Load", PdfIndexFile))					; erstellt den files Array aus der pdfIndex.txt Datei
			PageSum := ScanPoolArray("CountPages")

		PDFfiles := ReadDir(BefundOrdner, "pdf")

	;nicht mehr vorhandene Dateien aus dem Index nehmen
		For key, val in ScanPool {

			SPfile := StrSplit(val, "|").1
			filefound := false
			For idx, pdfname in PDFfiles
				If InStr(pdfname, SpFile) {
					tmpPool.Push(val)
					PDFFiles[idx] := "---"
					filefound := true
					break
				}

		}
		ScanPool := Array()
		For idx, val in tmpPool
			ScanPool.Push(val)

	;nach noch nicht aufgenommenen Dateien suchen
		For idx, pdfname in PDFfiles {
			If (pdfname != "---")	{
				pdfPath := Addendum.BefundOrdner "\" pdfname
				FileGetSize	, FSize	, % pdfPath, K
				FileGetTime	, FTime	, % pdfPath, M
				pages := PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
				ScanPool.Push(pdfname "|" FSize "|" FTime "|" pages)
				continue
			}

		}

	;Sortieren der eingelesenen und aktualisierten Dateien
		ScanPoolArray("Sort")
		ScanPoolArray("Save")

return CountValidKeys(ScanPool)
}

RefreshPdfIndex(BefundOrdner) {	                                   	;-- frischt das ScanPool Object auf

	; WICHTIG!: braucht eine globale Variable im aufrufenden Skript: ScanPool := Object()

		newPDFs 	:= 0
		tmpObj 	:= ScanPool                      	; Kopie des ScanPool-Objektes anlegen
		ScanPool	:= Object()                    		; ScanPool-Objekt leeren

	; alle pdf Dokumente einlesen
		PDFfiles	:= ReadDir(BefundOrdner, "pdf")

	; nicht mehr vorhandene Dateien aus dem Index nehmen
		For key, val in tmpObj {

			SPfile := StrSplit(val, "|").1
			filefound := false
			For idx, pdfname in PDFfiles
				If InStr(pdfname, SpFile) {
					ScanPool.Push(val)
					PDFFiles[idx] := "---"
					filefound := true
					break
				}

		}

	; neue PDF Dateien hinzufügen
		For idx, pdfname in PDFfiles {
			If (pdfname != "---")	{
				newPDFs ++
				pdfPath := BefundOrdner "\" pdfname
				FileGetSize	, FSize	, % pdfPath, K
				FileGetTime	, FTime	, % pdfPath, M
				pages           	:= PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
				isSearchable 	:= PDFisSearchable(pdfPath) ? 1 : 0
				ScanPool.Push(pdfname "|" FSize "|" FTime "|" pages "|" isSearchable)
			}

		}

	;Sortieren der Daten
		For key, val in ScanPool
			If (StrLen(Trim(val)) > 0)
				allfiles .= val "`n"
		Sort, allfiles
		ScanPool := StrSplit(RTrim(allfiles, "`n"), "`n")

		ScanPoolArray("Save")

return newPDFs
}

;}

ReadPatientDatabase(PatDBPath) {										               				;-- liest die .csv Datei Patienten.txt als Object() ein

	PatDB := Object()
	If !RegExMatch(PatDBPath, "\.json$")
		PatDBPath := Addendum.DBPath "\Patienten.json"

	If !FileExist(PatDBPath)
		return PatDB

return JSONData.Load(PatDBPath,, "UTF-8")
}

PatInDB(PatDb, Name) {						     													;-- für ScanPool - Fuzzy-Patientensuche

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Name[3] - ist Nachname, Name[4] sollte der Vorname sein
	;----------------------------------------------------------------------------------------------------------------------------------------------
		If !IsObject(PatDB)
			Exceptionhelper(A_ScriptName, "PatInDb(PatDB,cmd:="") {", "Die Funktion hat einen Fehler ausgelöst,`nweil kein Objekt übergeben wurde.", A_LineNumber)

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; versucht auch durch vertauschen von Vor- und Nachnamen den Patienten in der Datenbank zu finden
	;----------------------------------------------------------------------------------------------------------------------------------------------
		N1	:= Name[3]
		N2	:= Name[4]
		If (PatID := FindPatData(PatDb, "PatID", "Nn", N1, "Vn", N2))
			return PatID
		else If (PatID := FindPatData(PatDb, "PatID", "Nn", N2, "Vn", N1))
			return PatID

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Fuzzy Teil
	;----------------------------------------------------------------------------------------------------------------------------------------------
		For key, Pat in PatDb		{
			If (StrSplit(Pat["Vn"], " ").MaxIndex() > 1)	{
				Loop % StrSplit(Pat["Vn"], " ").MaxIndex() {		;geht jeden Vornamen durch und vergleicht ihn mit dem gesuchten Namen

					nr := A_Index
					Loop % StrSplit(N2, " ").MaxIndex() {

						NVdb	:= (Pat.Nn) . (StrSplit(Pat.Vn, " ")[nr])
						NVcmp	:= N1 . StrSplit(N2, " ")[A_Index]
						VNcmp	:= N2 . StrSplit(N1, " ")[A_Index]

						If ( StrDiff(NVdb, NVcmp) <= 0.2 ) || ( StrDiff(NVdb, Name[4] StrSplit(Name[3], " ")[A_Index]) <= 0.2 )
							suggestion.= key . ": " . Pat["Nn"] . ", " . Pat["Vn"] . "|"
					}
				}
			}
				else
				{
						If ( StrDiff(Pat["Nn"] . Pat["Vn"], Name[3] . Name[4]) <= 0.2 ) || ( StrDiff(Pat["Nn"] . Pat["Vn"], Name[4] . Name[3]) <= 0.2 )
								suggestion.= key . ": " . Pat["Nn"] . ", " . Pat["Vn"] . "|"
				}
		}

return RTrim(suggestion, "|")
}

PatInDB_SiftNgram(Name) {

		static Haystack, Haystack_Init:=0
		found:=Object(), Needle:= [], idx:= 0

		If !IsObject(oPat)
			MsgBox, Kein Objekt!

		If !Haystack_Init		{
			Haystack_Init:= 1
			For PatNr, Pat in oPat
				Haystack .= Pat["Nn"] " " Pat["Vn"] "`n"
		}


		Needle[1]:= Name[3] " " Name[4]
		Needle[2]:= Name[4] " " Name[3]

	; from https://www.autohotkey.com/boards/viewtopic.php?f=76&t=28796 - Getting closest string
		Loop, 2		{
			for key, element in Sift_Ngram(Haystack, Needle[A_Index], 0,, 2, "S")
			If (element.delta > 0.90) 	{
					idx ++
					element := StrSplit(element.data, " ")
					found[idx] := GetFromPatDb("PatID|Nn|Vn|Gd", 1, element[1], element[2])
			}
		}

return (idx=0) ? 0 : found
}

GetFromPatDb(getString:="PatID|Gd", OnlyFirstMatch:=0, Nn:="", Vn:="", Gt:="", Gd:="", Kk:="") {

	returnObj	:= Object()
	getter    	:= StrSplit(getString, "|")

	For PatID, Pat in oPat
		If InStr(Pat["Nn"], Nn) && InStr(Pat["Vn"], Vn) && InStr(Pat["Gt"], Gt) && InStr(Pat["Gd"], Gd) && InStr(Pat["Kk"], Kk)			{
			Loop, % getter.MaxIndex()
				If InStr(getter[A_Index], "PatID")
					returnObj[getter[A_Index]] := PatID
				else
					returnObj[getter[A_Index]] := Pat[getter[A_Index]]

				If OnlyFirstMatch
					break
			}

return returnObj
}

FindPatData(tPat, returnValue, pA1, pA2, pB1:="", pB2:="") {

	;zB. FindPatData(oPat, "PatID", "Nn", "Mustermann", "Vn", "Max" ) gibt die PatientenID zurück wenn ein Datensatz vorhanden ist
	;zB. FindPatData(oPat, "Birth"	, "Nn", "Mustermann", "Vn", "Max" ) gibt das Geburtsdatum zurück wenn ein Datensatz vorhanden ist
	;;Nn - Nachname, Vn- Vorname, Gt - Geschlecht, Gd - Geburtsdatum, Kk - Krankenkasse

	For FPD_PatID, PatData in tPat
	{
			If (PatData[(pA1)] = pA2)
			{
						If (pB1 = "Vn") && (StrSplit(PatData[pB2], " ").MaxIndex() > 1)
						{
								For k, v in StrSplit(PatData[pB2], " ")
									If v = pB2
										return ( InStr(returnValue, "PatID") ? k : PatData[(returnValue)] )
						}
						else if (PatData[(pB1)] = pB2)
						{
								return ( InStr(returnValue, "PatID") ? k : PatData[(returnValue)] )
								;~ If InStr(returnValue, "PatID")
										;~ return FPD_PatID
								;~ else
										;~ return PatData[(returnValue)]
						}
			}
		}

return 0
}

AlbisImportierePdf(PdfName, LVRow:=0) {                                                 	;-- ### Originalfunktion für Scanpool !! Importieren einer PDF-Datei in eine Patientenakte

	; benötigt ein globales Objekt mit dem Namen SPData -> Aufbau siehe ScanPool.ahk
	; oPat : globals Object, enthält die Addendum interne Patientendatenbank
	; Addendum: globales Object - enthält Daten zu Albis-Ordnern

		suggestions:= Object()
		BlockInput, On

	; für das ScanPool-Skript, ansonsten wird davon ausgegangen das die richtige Patientenakte geöffnet ist
		If InStr(A_ScriptName, "ScanPool") {

			; Pdf Previewer ausschalten
				PrevBackup:= dpiShow, dpiShow:= "Aus"

			; Patientenname ermitteln
				Name := ExtractNamesFromFileName(PdfName)

			; Prüfen ob der Patient in der Datenbank vorhanden ist und ermitteln einer Patienten ID wenn möglich
				If oPat.Count()	{

					suggestions:= PatInDB_SiftNgram(Name)
					If IsObject(suggestions)	{
						If suggestions.MaxIndex() = 1
							Name[5]:= suggestions.1.PatID
					}
					else	                        	{
						InputBox, newName, ScanPool, % "Patient: " Name[1] "konnte in der Addendum-Datenbank nicht gefunden werden.`nBitte geben Sie hier den korrekten Namen ein!",, 400,,,,,, % Name[1]
						Name:= ExtractNamesFromFileName(newName)
					}

				}

			; Patientenakte öffnen
				PraxTT("Öffne die Akte von: `n" . Name[1] , "3 2")
				If !AlbisOeffneAkte(Name) {
						PraxTT("Die Akte des Patienten`n" . Name[1] . "`nließ sich nicht öffnen! Der angeforderte Vorgang wird jetzt abgebrochen", "6 2")
						return 0
				}

		}

	; Eingangsdatum des Befundes auslesen
		If InStr(A_ScriptName, "ScanPool")
			creationtime := FormatedFileCreationTime(SPData.BefundOrdner "\" PdfName)
		else
			creationtime := FormatedFileCreationTime(Addendum.BefundOrdner "\" PdfName)

	; öffnet den Dialog 'Grafischer Befund' durch Eingabe eines Kürzel in die Akte - falls Sie ein anderes Kürzel verwenden - müssen Sie dies hier ändern
		PraxTT("Importiere den Befund:`n" PdfName, "15 2")
		AlbisSetzeProgrammDatum(creationtime)

	; Eingabefokus in die Karteikarte setzen
		If !AlbisKarteikartenFocusSetzen("Edit3") {
			PraxTT("Beim Übertragen des Befundes ist ein Problem aufgetreten.`nEs konnte kein Eingabefocus in die Karteikarte gesetzt werden.", "6 2")
			BlockInput, Off
			return 0
		}

		SendRaw, Scan
		sleep, 100
		SendInput, {Tab}

	; Erstellen des Karteikartentextes
		KarteikartenText := StrReplace(PdfName, ".pdf")            	;entfernen von '.pdf'

		If InStr(A_ScriptName, "ScanPool") {
			KarteikartenText := StrReplace(KarteikartenText, Name[3], "")
			KarteikartenText := StrReplace(KarteikartenText, Name[4], "")
		} else {
			KarteikartenText := RegExReplace(KarteikartenText, "^[\w\p{L}-]+\s*\,*\s*[\w\p{L}]+\s*\,*\s*", "")
		}

		KarteikartenText := RegExReplace(KarteikartenText, "^\s*\,*\s*")
		KarteikartenText := StrReplace(KarteikartenText, "   ", "°")
		KarteikartenText := StrReplace(KarteikartenText, "  ", "°")
		KarteikartenText := StrReplace(KarteikartenText, "°", " ")

	; Übertragen des Befundes in die Akte
		If !AlbisUebertrageGrafischenBefund(PdfName, KarteikartenText) 	{
			PraxTT("Beim Übertragen des Befundes`nist ein Problem aufgetreten.", "6 2")
			BlockInput, Off
			return 0
		}

	; ScanPool nutzt derzeit noch andere Variablen
		If InStr(A_ScriptName, "ScanPool") {

			; Daten sichern (PdfData ist global - ein return nicht notwendig)
				PdfData.KarteikartenText 	:= KarteikartenText
				PdfData.PdfFullPath       	:= PdfName
				PdfData.ReaderID 	        	:= ReaderID

			; Zurücksetzen von gemachten Einstellungen
				DpiShow	:= PrevBackup
		}
		else	{

				SPData.KarteikartenText 	:= KarteikartenText
				SPData.PdfFullPath        	:= PdfName
				SPData.ReaderID 	        	:= ReaderID

		}

	; auf aktuelles Tagesdatum zurücksetzen
		AlbisSetzeProgrammDatum()

	BlockInput, Off

		PraxTT("Der Befund wurde importiert.", "3 2")

return creationtime
}

AlbisOeffneAkte(Pat, PdfPath:="") {                                                            	;-- ursprüngl. ScanPool.ahk - Akte lässt sich auch öffnen, selbst wenn der Name nicht ganz korrekt geschrieben ist

		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

		AllreadyOpen       	:= 0
		CurrPat    				:= Object()
		CurrPat    				:= VorUndNachname(AlbisCurrentPatient())

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Überprüfen ob die Patientenakte schon geöffnet ist
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		If (Pat[5] = "")		{
			PraxTT("Öffne die Akte des Patienten`n#2" Pat[1], "3 3")
			AllreadyOpen := FuzzyNameMatch(CurrPat, Pat)
		}
		else		{
			If InStr(AlbisAktuellePatID(), Pat[5])
				AllreadyOpen := 1
		}

		If AllreadyOpen		{
			AlbisZeigeKarteikarte() 		; sicherstellen das die Karteikarte angezeigt wird
			PraxTT("Die Patientenakte ist schon geöffnet!", "3 3")
			return 1
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Patient öffnen - Dialog: Aufruf des Fensters und Texteingabe
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		AlbisActivate(2)

	; 32768 Menu Kommando für den 'Patient öffnen' Auswahldialog
		while !WinExist(Win_PatientOeffnen)		{

			PostMessage, 0x111, 32768,,, % "ahk_id " AlbisWinID()
			Sleep, 200
			If (A_Index > 10) {
				PraxTT("Der Dialog zum Öffnen einer Patientenakte`nließ sich nicht aufrufen.`nDer Vorgang wurde abgebrochen!", "5 0")
				return 0
			}

		}

		WinActivate, % Win_PatientOeffnen
		WinWaitActive, % Win_PatientOeffnen,, 3

		If VerifiedSetText("Edit1", (Pat[5]="" ? Pat[1] : Pat[5]), Win_PatientOeffnen, 200)	{		;Pat enthält nur die ID wenn es kein Array ist

			MsgBox, 1, % "Addendum für Albis on Windows - " A_ScriptName,
                           (LTrim
                       	   Der Patientenname konnte nicht eingetragen werden.

							1. Bitte manuell per Strg+c Taste einfügen (Kopie aus dem Clipboard).
							2. Drücken Sie dann auf 'Ok" im 'Patient öffnen' - Dialog
							3. Warten Sie das Öffnen der Akte ab,
							4. Drücken Sie erst jetzt auf den 'Ok"-Button in diesem Dialog!
							)

		}
		else		{
			  ; Patient öffnen - Dialog: Ok oder Enter drücken
				LastPopUpWin := GetLastActivePopup(AlbisWinID())
				while WinExist(Win_PatientOeffnen)		{

					VerifiedClick("Button2", Win_PatientOeffnen)			; Button OK drücken
					WinWaitClose, % Win_PatientOeffnen,, 1

					if WinExist(Win_PatientOeffnen)		{                 	; Fenster ist immer noch da? Dann sende ein ENTER.
							WinActivate, % Win_PatientOeffnen
								WinWaitActive, % Win_PatientOeffnen,, 1
							ControlFocus, Edit1, % Win_PatientOeffnen
							SendInput, {Enter}
					}

					Sleep, 300

				}
		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; 2 Folgedialoge sind möglich - 1. Patient nicht vorhanden , 2. Listview-Auswahl -> den will ich durch möglichst vermeiden
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		PopWin:= WaitForNewPopUpWindow(AlbisWinID(), LastPopUpWin, 10)

	  ;Dialog Patient <.....,......> nicht vorhanden schließen.
		If ( Instr(PopWin.Title, "ALBIS") AND Instr(PopWin.Class, "#32770") AND  Instr(PopWin.Text, "nicht vorhanden") ) {
			LastPopUpWin := GetLastActivePopup(AlbisWinID())
			VerifiedClick("Button1", "", "", PopWin.Hwnd)			;soll das Fenster schließen
			PopWin:= WaitForNewPopUpWindow(AlbisWinID(), LastPopUpWin, 30)
		}
	  ;Dialog Patient auswählen mit eingebettetem ListView-Fenster
		If ( Instr(PopWin.Title, "Patient") AND Instr(PopWin.Class, "#32770") AND  Instr(PopWin.Text, "List1") ) {
				;Scite_OutPut("Listenfenster ist aufgetaucht", 0, 1, 0)
			return 0
		}
	  ;Dialog Patient öffnen schließen der sich nach dem Schließen des "<Patient ... nicht vorhanden>" Dialoges erneut öffnet
		If ( Instr(PopWin.Title, "Patient öffnen") AND Instr(PopWin.Class, "#32770") ) {
			LastPopUpWin := GetLastActivePopup(AlbisWinID())
			VerifiedClick("Button3", "", "", PopWin.Hwnd)			;soll das Fenster schließen
			PopWin := WaitForNewPopUpWindow(AlbisWinID(), LastPopUpWin, 30)
			return 0
		}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; wartet auf das Öffnen der angeforderten Akte
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		PraxTT( "Warte bis zu 10s auf die Patientenakte (" Name[5] ", " Name[1] ")", "10 3" )
		start:= A_TickCount
		Loop
		{
				If FuzzyNameMatch(CurrPat, Pat) or InStr(AlbisAktuellePatID(), Pat[5])
				{
						PraxTT("Die Patientenakte ist jetzt bereit!", "3 3")
						return 1
				}

				Sleep, 50
				timesince:= (A_TickCount - start)/1000
				GuiControl, PraxTT:, PraxTTCounter2, % Round( (10-timesince), 1 ) "s"
				GuiControl, PraxTT:, PraxTTCounter1, % Round( (10-timesince), 1 ) "s"

		} until timesince > 9.9

		PraxTT("Die richtige Patientenakte konnte nicht geöffnet werden!!", "3 3")

	;}

return 0
}

FuzzyNameMatch(Name1, Name2, diffmax := 0.12) {                                  	;-- Fuzzy Suchfunktion für Vor- und Nachnamensuche

	surname1	:= Trim(Name1[1])
	prename1	:= Trim(Name1[2])
	surname2	:= Trim(Name2[3])
	prename2	:= Trim(Name2[4])

	;FileAppend, % "Name1: " prename1 ", " surname1 "    matching with Name2: " prename2 ", " surname2 "`n", %A_ScriptDir%\logs\MatchingMethodLog.txt
	;FileAppend, % StrDiff(surname1 . prename1, prename2 . surname2) "`n", %A_ScriptDir%\logs\MatchingMethodLog.txt
	;FileAppend, % "Name1: " prename1 ", " surname1 "    matching with Name2: " surname2 ", " prename2 "`n", %A_ScriptDir%\logs\MatchingMethodLog.txt
	;FileAppend, % StrDiff(surname1 . prename1, surname2 . prename2) "`n`n", %A_ScriptDir%\logs\MatchingMethodLog.txt

	if ( StrDiff(surname1 . prename1, surname2 . prename2) <= diffmax )  ||  ( StrDiff(surname1 . prename1, prename2 . surname2) <= diffmax )
			return 1

return 0
}

FuzzySearch(string1, string2) {

	lenl := StrLen(string1)
	lens := StrLen(string2)
	if(lenl > lens)
	{
		shorter := string2
		longer := string1
	}
	else if(lens > lenl)
	{
		shorter := string1
		longer := string2
		lens := lenl
		lenl := StrLen(string2)
	}
	else
		return StrDiff(string1, string2)
	min := 1
	Loop % lenl - lens + 1
	{
		distance := StrDiff(shorter, SubStr(longer, A_Index, lens))
		if(distance < min)
			min := distance
	}
	return min
}

VorUndNachname(Name) {				                              		        				;-- teilt einen Komma-getrennten String und entfernt Leerzeichen am Anfang und Ende eines Namens
	Arr    	:=[]
	Arr[1]	:= StrSplitEx(Name, 1, ",")		;Trimmed den String
	Arr[2]	:= StrSplitEx(Name, 2, ",")
return Arr
}

StrSplitEx(str, nr=1, splitchar:="|") {                                                       		;-- Trim-Split mit Rückgabe eines Wertes (keine Array-Rückgabe)
	splitArr:= StrSplit(str, splitchar)
return Trim(splitArr[nr])
}

ExceptionHelper(libPath, SearchCode, ErrorMessage, codeline) { 					;-- searches for the given SearchCode in an uncompiled script as a help to throw exceptions

	If !A_IsCompiled {

	;Fehlerfunktion bei Eingabe eines falschen Parameter
		FileRead, Pfunc, % AddendumDir "\" libPath
		FileOpen(AddendumDir "\" libPath, "r", "UTF-8").Read()
		For idx, line in StrSplit(Pfunc, "`n", "`r") {
			If Instr(line, SearchCode) {
				scriptline	:= A_Index
				ScriptText	:= line
				break
			}
		}

		Exception(ErrorMessage)

	} else {

		msg=
		(Ltrim
		This message is shown, because the script wanted
		to call a function that works only in uncompiled scripts.

		A function was called to show a runtime error.
		This function was called from %A_ScriptName%
		at line: %codeline%. The code to show ist:
		%SearchCode%
		with the following error-message:
		%ErrorMessage%
		)

		MsgBox, % "Addendum für AlbisOnWindows - " A_ScriptName,  % msg

	}

}

; -----------------------------------------------------------------------------------------------------------------------------------------------
; -----------------------------------------------------------------------------------------------------------------------------------------------

;{ Hades :-( Exeatis hoc mundo! Nihil praeparari! Codice continet obsoleta. :-)

;				else if InStr(FoxitTitle1, "FoxitReader.exe")
;							FoxitReader_ExceptionDialog(SpEHookHwnd1)

/*
Loop, % ScanPool.MaxIndex()-1
		{
			row:= A_Index
			ToolTip, % "row: " row "`n""`nPdf: `t" BOPdfFile "`nBOPCount: `t" BOPCount
			;Defaults festlegen
				Gui, BO: Default
				Gui, BO: Listview, Listview1
			LV_GetText(Pages, row, 3)
			If (Pages="") {

				;Auslesen des Dateinamen , Spalte 1
				Gui, BO: Default
				Gui, BO: Listview, Listview1
				LV_GetText(BOPdfFile, row)
				BOPdfFile:= BefundOrdner . "`\" . BOPdfFile
				BOPcFile:= StrReplace(BOPdfFile, "`.pdf", "`.cnt")

				;ursprüngliche if Abfrage, zuvor wurden die Seitenzahlen erst nach dem kompletten Füllen der Listview mit Daten hinzugefügt
					If FileExist(BOPcFile) {
						FileReadLine, BOPCount, %BOPcFile%, 1
					} else {
						;BOPCount:= PdfPageCounter(BOPdfFile)
						PdfInfo:= PdfInfo(BOPdfFile)
						result:= RegExMatch(PdfInfo, "(Pages:\s+\d)", Match)
						BOPCount:= Trim(StrReplace(Match1, "Pages:", ""))
						ToolTip, % "row: " row "`nPdf: `t" BOPdfFile "`nBOPCount: `t" BOPCount
						;sleep, 3300
						;FileAppend, %BOPCount%, %BOPcFile%
					}
			}
			;zusammenzählen und anzeigen
				PageSum+= BOPCount
				If (BOPCount = "")
							BOPCount:= "?"
				Gui, BO: Default
				LV_Modify(BOPageCount, "Col3", BOPCount)
		}

*/

/*		ehemals BOPdfCheckReader - ist nur noch die letzte Codezeile übrig
	;nach sehen ob eine Reihe ausgewählt ist, die nicht ausgewählt sein sollte, z.B. weil der User diese ausgewählt hat
		LVRowState:= LV_GetItemState(hLV1, A_Index)
		If LVRowState.checked			;wenn Reihe ist checked = 1 dann
		{
			If !Instr(pdfList, RetrievedText) {
				Gui, BO: Default
				LV_Modify(A_Index, "-Check")
			}
		}
	*/

/*
 while !StrLen(fDir)
 {
					Round:= A_Index
					If A_Index = 1
						ControlClick,, ahk_id %hDir%,,,, NA
					If A_Index = 2
					{
							ControlGetPos	, cx	, cy	,,,, ahk_id %hDir%
							MouseGetPos	, mx	, my
							MouseMove		, % cx+10, % cy + 10, 0
							MouseClick		, Left
							MouseMove		, % mx, % my, 0
					}
					If A_Index = 2
							break

					;hEdit2	:= GetHex(GetChildHWND(hHook, "Edit2"))
					;ControlGetText, fDir,, ahk_id %hEdit2%
					fDir := Controls("Edit2", "GetText", hHook, true, true, "slow")

					If InStr(fDir, "Speichern")
							goto SpeichernUnterHandlerExit
					else if StrLen(fDir) > 0
							break

					Sleep, 3000
			}

*/

/*

else if RegExMatch(FoxitTitle1, "i)\d+\w\w\.pdf") and ((A_TickCount - docImported)/1000 < 10)
						{
									docImported:=0
									;FoxitInvoke("Close", SpEHookHwnd1)	; für den Fall das ich einen Weg finde Tabs im FoxitReader anzusteuern
									;Win32_SendMessage(SpEHookHwnd1)							;, "", )
						}

*/

/*			Datei Speichern unter Dialog
			;----------------------------------------------------------------------------------------------------------------------------------------------
			; f)	Logfile schreiben
			;----------------------------------------------------------------------------------------------------------------------------------------------
				SpeichernUnterHandlerExit:
				write:= "| " A_DD A_MM SubStr(A_YYYY, -2) "," A_Hour A_Min A_Sec ":" A_MSec " | " (JPos=0 ? (hClass " `/ " GetHex(hHook)) : "Abbruch: " JPos) " | " hEdit1 " | " hEdit2a " | " hEdit2 " | " efname "/" fname " | " cFocus " | " (fDir="" ? "read error " : fDir) ( f="" ? "" : (" " f " " fname ")") ) " |`n"
				FileAppend, % write, %A_ScriptDir%\ScanPoolLog.MD
				If RegExMatch(fname, "\d\dA\w\.pdf")
						FileAppend, % fDir "\" fname "," AlbisAktuellePatID() "`n", % AddendumDir "\logs'n'data\_DB\PdfReferenzen\PdfLinks.txt"
*/


;}



; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;               					       										⚫      INFOFENSTER		⚫
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;      Funktionen:   		▫ Verwaltung/Bearbeitung (Texterkennung/Kategorisierung) gescannter Post vor dem Import in die Patientenkartei
;                           	▫ Überwachung des Dokumentverzeichnisses (Hotfolder) für eine automatisierte
;                           	▫ Texterkennung mittels Tesseract in eigenem Thread ("Hintergrundanwendung")
;								▫ Autonaming (automatische Erkennung von Patientennamen und Klassifizierung)
;                           	▫ Anzeigen von Praxisinformationen
;                           	▫ ## Netzwerkkommunikation - funktioniert noch? nicht
;                           	▫ RDP Sessions starten
;                           	▫ Tagesprotokolle gruppiert nach PCs im Netzwerk
;								▫ Programme und Skripte starten
;								▫ Abrechnungshelfer
;								▫ COVID19-Schutzimpfungsstatistik
;								▫ Zusatzfunktionen: Laborblattdruck
;
;      Basisskript: 	Addendum.ahk
;
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - letzte Änderung 31.07.2023 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

; Infofenster
AddendumGui(admShow="", admGuiDebug="false")                         	{

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global

		static admLVPatOpt, admLVJournalOpt, admLVTProtokoll ;, DDLPrinter
		static ImportChecks := 0
		static fn_admGui := Func("AddendumGui")

		local MObj, MNr, cpx, cpy, cpw, cph, Item
		local APos, SPos, mox, moy, moWin, moCtrl, regExt, Aktiv, Pat, nfo, res
		local ClientName, LanClients, ClientIP, onlineclient, OffLineIndex, InfoText, found, YPlus, Y, ProtClients

		If admGui_Exist()
			return

		func_admGui       	:= Func("AddendumGui")

		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		; Initialisierung von Variablen
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If !Addendum.iWin.Init {

			; Befunde (Klassen-Objekt) - Dateihandler für alle Dateizugriffe auf den Befundordner
				Befunde.options(false)

				Addendum.iWin.Init            	:= true
				Addendum.iWin.FilesStack 	:= []
				Addendum.iWin.OCRStack 	:= []
				Addendum.iWin.RowColors	:= false
				Addendum.iWin.Imports     	:= 0
				Addendum.iWin.ImportsLast 	:= -1
				Addendum.iWin.Check      	:= 0
				Addendum.iWin.ReCheck   	:= 0
				Addendum.iWin.paint        	:= 0
				Addendum.iWin.TPClient      	:= !Addendum.iWin.TPClient ? compname : Addendum.iWin.TPClient
				Addendum.tessOCRRunning	:= false        	; OCR Vorgang läuft
				Addendum.iWin.Impfstatistik	:= FileExist(Addendum.Dir "\include\Gui\Impfstatistik.html") ? true : false
				Addendum.Flags.ImportFromPat   	:= false           	; Import aus Patient Tab läuft
				Addendum.Flags.ImportFromJrnl  	:= false           	; Import aus dem Journal Tab läuft

				factor                	:= A_ScreenDPI / 96

				admTags            	:= {	"positive"	: "Karteikarte|Laborblatt|Biometriedaten|Rechnungsliste"
												 ,	"negative"	: "Abrechnung|Privatabrechnung|Überweisung"}

				admTabNames  	:= "Patient|Journal|Protokoll|Extras|Netzwerk|Info" (Addendum.iWin.Impfstatistik ? "|Impfen" : "")
				admGuiTabs     	:= {"Patient":0, "Journal":1, "Protokoll":2, "Extras":3, "Netzwerk":4, "Info":5, "Impfen":6}
				admLVPatOpt    	:= " r" (Addendum.iWin.LVScanPool.r)
				admBGColor     	:= "BackgroundF0F0F0 "
				admLVPatOpt    	.= " gadm_Reports vadmReports   	 HWNDadmHReports     AltSubmit -Hdr ReadOnly LV0x10020"
											. 	 " 0x5001804B " admBGColor
				admLVJournalOpt	:= " gadm_Journal vadmJournal 	 HWNDadmHJournal 	 AltSubmit                        LV0x10020"
											.	 " 0x50210049 -E0x200 " admBGColor
				admLVTProtokoll 	:= " gadm_TProtLV vadmTProtokoll HWNDadmHTProtokoll AltSubmit        ReadOnly  LV0x10020"
											.	 " 0x50210049 -E0x200 " admBGColor

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Druckerliste anlegen
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				DDLPrinter := Addendum.Drucker.StandardA4 "|" Addendum.Drucker.PDF "|"
				DDLPrinter := RegExReplace(DDLPrinter, "\|\|", "|")
				DDLPrinter := RegExReplace(DDLPrinter, "^\|$", "")

				for Item in ComObjGet( "winmgmts:" ).ExecQuery("Select * from Win32_Printer")
					DDLPrinter .= !InStr(DDLPrinter, Item.Name) ? Item.Name "|" : ""

				DDLPrinter := RTrim(DDLPrinter, "|")
			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; ICON Liste laden
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				admHBM := Array()
				admHBM.Push(LoadPicture(Addendum.Dir "\assets\ModulIcons\PDFImage.ico"))	; PDF ohne OCR
				admHBM.Push(LoadPicture(Addendum.Dir "\assets\ModulIcons\image.ico"))    	; Bilddatei
				admHBM.Push(LoadPicture(Addendum.Dir "\assets\ModulIcons\PDFOCR.ico"))	; PDF mit OCR

				;-: lädt die Icons
				admImageListID  	:= IL_Create(3)
				IL_Add(admImageListID, "HBITMAP:" admHBM[1], 0x00000) 		;1    - 0xFFFFFF
				IL_Add(admImageListID, "HBITMAP:" admHBM[2], 0xFFFFFF)		;2
				IL_Add(admImageListID, "HBITMAP:" admHBM[3], 0x00000) 		;3

			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Kontextmenu Funktion für das Journal
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				func_admOpen 	:= Func("admGui_CM").Bind("JOpen")
				func_admImport	:= Func("admGui_CM").Bind("JImport")
				func_admRename	:= Func("admGui_CM").Bind("JRename")
				func_admSplit   	:= Func("admGui_CM").Bind("JSplit")
				func_admDelete	:= Func("admGui_CM").Bind("JDelete")
				func_admView1 	:= Func("admGui_CM").Bind("JView1")
				func_admView2 	:= Func("admGui_CM").Bind("JView2")
				func_admExport 	:= Func("admGui_CM").Bind("JExport")

				func_admMail    	:= Func("admGui_CM").Bind("JMail")
				func_admRefresh	:= Func("admGui_CM").Bind("JRefresh")
				func_admRecog1	:= Func("admGui_CM").Bind("JRecog1")
				func_admRecog2	:= Func("admGui_CM").Bind("JRecog2")
				func_admRecog3	:= Func("admGui_CM").Bind("JRecog3")
				func_admOCR  	:= Func("admGui_CM").Bind("JOCR")
				func_admOCRAll	:= Func("admGui_CM").Bind("JOCRAll")

			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Kontextmenu (Rechtsklickmenu) für das Journal
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				Menu, admJCM, Add, % "Karteikarte öffnen"                     	, % func_admOpen

				; Dokument anzeigen
				If (StrLen(Addendum.PDF.Reader) > 0 && StrLen(Addendum.PDF.ReaderAlternative) > 0) {
					Menu, admJCMView, Add, % Addendum.PDF.ReaderName            	, % func_admView2
					Menu, admJCMView, Add, % Addendum.PDF.ReaderAlternativeName	, % func_admView1
					Menu, admJCM, Add, % "Anzeigen mit"                         	, :admJCMView
				} else if (StrLen(Addendum.PDF.Reader) = 0 && StrLen(Addendum.PDF.ReaderAlternative) > 0) {
					Menu, admJCM, Add, % "Anzeigen"                              	, % func_admView1
				} else if (StrLen(Addendum.PDF.Reader) > 0 && StrLen(Addendum.PDF.ReaderAlternative) = 0) {
					Menu, admJCM, Add, % "Anzeigen"                              	, % func_admView2
				}

				; Import & Export
				Menu, admJCM    	, Add, % "Importieren"                             	, % func_admImport
				Menu, admJCM    	, Add, % "Exportieren"                              	, % func_admExport

				; Versenden per Telegram
				If Addendum.Telegram.Bots.Count() && Addendum.Telegram.Chats{

					If Addendum.Telegram.Bots[1].active {

						For name, chatID in Addendum.Telegram.Chats {
							func_admTgram 	:= Func("admGui_CM").Bind("JTgram_" name)
							Menu, admJCMSend	, Add, % "an " name " (" (chatID~="^\-" ? "Gruppe" : "Nutzer") ")" 	, % func_admTgram
						}

						bot := Addendum.Telegram.Bots[1].botname
						Menu, admJCM    	, Add, % "Versenden mit Telegram (" bot ")", :admJCMSend
					}
				}

				; Versenden per EMail
				Menu, admJCM    	, Add, % "per EMail versenden"                	, % func_admMail

				; Dateioperationen
				Menu, admJCM     	, Add, % "Umbennen"                              	, % func_admRename
				Menu, admJCM    	, Add, % "Löschen"                                    	, % func_admDelete

				; Kategorisierung
				Menu, admJCM    	, Add, % "Texterkennung ausführen"          	, % func_admOCR
				Menu, admJCM    	, Add, % "Inhaltserkennung"      	            	, % func_admRecog1
				If Addendum.IamTheBoss
					Menu, admJCM	, Add, % "Inhaltserkennung (Debug)"      	, % func_admRecog3

				Menu, admJCM    	, Add
				Menu, admJCM    	, Add, % "Texterkennung - alle Dateien"    	, % func_admOCRAll
				Menu, admJCM    	, Add, % "automatische Benennung"          	, % func_admRecog2
				Menu, admJCM    	, Add, % "Befundordner neu indizieren"    	, % func_admRefresh

				; Menu bei Mehrfachauswahl von Dateien
				Menu, admJCMX    	, Add, % "Importieren"                             	, % func_admImport
				Menu, admJCMX    	, Add, % "Exportieren"                              	, % func_admExport
				Menu, admJCMX    	, Add, % "Umbennen"                              	, % func_admRename
				Menu, admJCMX    	, Add, % "Löschen"                                  	, % func_admDelete
				Menu, admJCMX    	, Add, % "Inhaltserkennung"                   	, % func_admRecog1
				If Addendum.IamTheBoss
					Menu, admJCMX, Add, % "Inhaltserkennung (Debug)"         	, % func_admRecog3

				;Menu, admJCM, Add, Aufteilen                                   	, % func_admSplit

			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Hotkey's für alle Tab's
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				func_admCMJournal := Func("GuiControlActive").Bind("admJournal"	, "adm")
				func_admSuche 		:= Func("GuiControlActive").Bind("admTPSuche"	, "adm")
				func_admTPSuche		:= Func("admGui_TPSuche")

			; Hotkeys - Journal
				Hotkey, If, % func_admCMJournal
				Hotkey, Enter 	, % func_admView1
				Hotkey, F2     	, % func_admRename
				Hotkey, F3     	, % func_admView1
				Hotkey, F5     	, % func_admRefresh
				Hotkey, F6     	, % func_admOCR
				If Addendum.IamTheBoss
					Hotkey, F7   	, % func_admRecog3
				else
					Hotkey, F7  	, % func_admRecog1
				Hotkey, Delete	, % func_admDelete
				Hotkey, If

			; Hotkeys - Protokoll
				Hotkey, If, % func_admSuche
				Hotkey, Enter 	, % func_admTPSuche
				Hotkey, If
			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; OnMessage
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
					LISTENERS := [ 	WM_ACTIVATE                      	:= 0x06
									,	WM_PARENTNOTIFY             	:= 0x210
									, 	WM_MDIACTIVATE                	:= 0x222]

			;}

		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Überprüfungen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; Abbruch des Selbstaufrufes der Funktion
	;  ( Albis wurde beendet oder es wurden mehr als 9 Versuche gebraucht um die Gui zu integrieren )
		If !WinExist("ahk_class OptoAppClass") || (Addendum.iWin.ReCheck > 9) {
			admGui_Destroy()
			If (Addendum.iWin.ReCheck > 9) {
				SetTimer, % func_admGui, Off
				IF (compname = "SERVER") {
					res                := Controls("", "Reset", "")
					hMDIFrame	:= Controls("AfxMDIFrame1401"	, "ID", AlbisMDIChildGetActive())     	; AlbisMDIChildGetActive(): Steuerelement-ID im aktiven Fenster
					hStamm		:= Controls("#327701"             	, "ID", hMDIFrame)                        	; Stammdaten Steuerelement-ID
					TrayTip, AddendumGui, % "Die Integration des Infofenster ist fehlgeschlagen.`n"
														.	"AfxMDIFrame1401: " hMDIFrame "`n"
														.	"#327701" hStamm "`n", 2
				}
			}
			Addendum.iWin.ReCheck := Addendum.iWin.paint := Addendum.iWin.Check := Addendum.iWin.lastPatID := 0
			return
		}

	; die Z-Position der Albisfenster für Integration des Infofensters finden
		If !(aPatID := AlbisAktuellePatID()) {
			Addendum.iWin.lastPatID := 0
			return
		}

		Addendum.iWin.Check ++

	; Albis ist beschäftigt (z.B. Abrechnung oder anderes), dann das Fenster nicht zeichnen da sich Addendum wahrscheinlich
	; dabei aufhängt oder ausgebremst wird
		If !CheckWindowStatus(AlbisWinID()) {
			SetTimer, % func_admGui, -1000
			return
		}

		Aktiv          	:= Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		res                := Controls("", "Reset", "")
		hMDIFrame	:= Controls("AfxMDIFrame1401"	, "ID", AlbisMDIChildGetActive())     	; AlbisMDIChildGetActive(): Steuerelement-ID im aktiven Fenster
		hStamm		:= Controls("#327701"             	, "ID", hMDIFrame)                        	; Stammdaten Steuerelement-ID

		If (!RegExMatch(Aktiv, "i)(?<Type>" admTags.positive ")", win) || !hMDIFrame || !hStamm) {   	; || RegExMatch(Aktiv, "i)(" Tags.negative ")")
			Addendum.iWin.lastPatID := AlbisAktuellePatID() && winType ? AlbisAktuellePatID() : 0
			Addendum.iWin.ReCheck ++
			If Addendum.iWin.debug {
				SciTEOutput( "admGui ReCheck: " Addendum.iWin.ReCheck "|"
							.	   (wintype 		? "" : "Kontext inkorrekt [" wintype "] |")
							.	   (hMDIFrame 	? "" : "hMDIFrame ist leer|")
							. 	   (hStamm 		? "" : "hStamm ist leer"))
			}
			SetTimer, % func_admGui, -1000
			return
		}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeichnen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	Addendum.iWin.PatID := aktivePatID := AlbisAktuellePatID()
	If (Addendum.iWin.lastPatID = aktivePatID) {
		return
	}
	If (Addendum.iWin.lastPatID <> aktivePatID) {

			Addendum.iWin.lastPatID := aktivePatID

		; Position von Albis und der Stammdaten ;{
			APos          	:= GetWindowSpot(AlbisWinID())
			SPos          	:= GetWindowSpot(hStamm)
		;}

		; legt die Höhe der Gui anhand der Größe des MDIFrame-Fenster fest ;{
			Addendum.iWin.StammY	:= Round(SPos.Y / factor)
			Addendum.iWin.X	         	:= APos.CW - Addendum.iWin.W ; ClientWidth passt besser!
			Addendum.iWin.H          	:= admHeight := SPos.CH

		;}

		; der Albisinfobereich (Stammdaten, Dauerdiagnosen, Dauermedikamente) wird für das Einfügen der Gui verkleinert ;{
			If (SPos.W >= APos.CW-(admWidth:= Addendum.iWin.W)) ; verhindert ein wiederholtes Verkleinern
				SetWindowPos(hStamm, 2, 2, APos.CW-admWidth-4, SPos.CH)
		;}

		; umgeht die seltene Problematik eines nicht mehr vorhandenen Handle für das MDIFrame (z.B. hat es geschlossen und blieb unbemerkt)
		; das Infofenster hier als Kind-Fenster anzubinden würde zu einer Fehlermeldung führen
			Sleep 200
			try {
				Gui, adm: New	, % "-Caption +ToolWindow 0x50000000 E0x0802009C +Parent" hMDIFrame " +HWNDhadm " (Addendum.noDPIScale ? "-DPIScale":"")
			} catch {
				SetTimer, % func_admGui, -1000
				Addendum.iWin.ReCheck ++
				return
			}

			SetTimer, % func_admGui, Off
			Gui, adm2: 	Show, Hide
			Gui, RN:   	Show, Hide
			Gui, RNP:  	Show, Hide
			Addendum.Flag.RNRunning := false

			; sicher gehen das keine DPI Skalierung verwendet wird (erst +DPIScale dann -DPIScale - warum Versuch macht klug!)
			Gui, adm: -DPIScale

		;-: Gui 	 :- Tab Start                      	;{
			Gui, adm: Margin	, 0, 0
			Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
			Gui, adm: Add   	, Tab     	, % "x0 y0 w" admWidth " h" admHeight+2 " HWNDadmHTab vadmTabs gadm_Tabs"       	, % admTabNames
		;}

		;-: Tab1 :- Patient                          	;{
			Gui, adm: Tab  	, 1

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Text      	, % "x10 y25 w" admWidth-15 " BackgroundTrans vadmPatientTitle Section"      	, % "(es wird gesucht ....)"

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Button    	, % "x" admWidth-55 " y23 h16 vadmButton1 gadm_BefundImport"                  	, % "Importieren"
			GuiControlGet, cpos, adm: Pos, admButton1
			GuiControl, adm: Move, admButton1, % "x" admWidth-cposW-5

			;-: Patientenbefunde
			admCol1W :=50, admCol2W := admWidth - admCol1W - 10
			Gui, adm: Font  	, s7
			Gui, adm: Add  	, ListView	, % "x2 y+1 	w" admWidth-6	" " admLVPatOpt                                                	, % "S|Befundname"

			Gui, adm: ListView, admReports
			LV_SetImageList(admImageListID)
			LV_ModifyCol(1, admCol1W " Integer Right NoSort")
			LV_ModifyCol(2, admCol2W " Text Left NoSort")

			;-: Abrechnungshelfer
			Gui, adm: Font  	, s7
			Gui, adm: Add  	, Text    	, % "x2 y+0 	w" (cpw := Floor((admWidth-6)/1.9)) " Center vadmGP1"            	, % "Abrechnungshelfer und andere Hinweise"

			GuiControlGet, cpos, adm: Pos, AdmGP1
			GCOption	:= "x2 y" cposY+12 " w" cpw " h" admHeight-cposY-cposH-4 " t6 t12 t22 vadmNotes"
			Gui, adm: Font  	, s8
			Gui, adm: Add, Edit, % GCOption , % (Addendum.iWin.AbrHelfer  ? admGui_Abrechnungshelfer(Addendum.iWin.PatID, cPat.GEBURT(Addendum.iWin.PatID)) : "")

			;-: zusätzliche Karteikartenfunktionen (Laborblatt drucken/Laborblatt per Mail versenden)
			GuiControlGet, cpos, adm: Pos, admNotes
			cpx := cposX + cposW + 3, cpy := cposY - 1, cpw	:= admWidth - cpx - 120 - 6
			Gui, adm: Add, Button       	, % "x" cpx " y" cpy " w78          	vadmLBD          	gadm_LBDruck"                 	, % "Labor drucken"
			Gui, adm: Add, Edit           	, % "x+0  y" cpy+1 " w" 52 "     	vadmLBSpalten"
			Gui, adm: Add, UpDown   	, % "x+0                               	vadmLBDUD     	gadm_LBDruck"                   	, % "1"

			Gui, adm: Font 	, s7
			Gui, adm: Add, Checkbox		, % "x+2  y" cpy " w" 84 "         	vadmLBAnm      	gadm_LBDruck"                   	, % "Anmerkung/`nProbeddaten"
			GuiControl, adm: , admLBAnm, % Addendum.iWin.LBAnm

			Gui, adm: Font 	, s7
			GuiControlGet, cpos, adm: Pos, admLBD
			cpw	:= admWidth - cpx - 5, cpy := cposY+cposH+1
			Gui, adm: Add	, DDL          	, % "x" cpx " y" cpy " w" cpw " r8 	vadmLBDrucker 	gadm_LBDruck"                 	, % DDLPrinter
			GuiControl, adm: ChooseString, admLBDrucker, % Addendum.iWin.LBDrucker
			Gui, adm: Font	, s8

			;-: Laborblattversand per EMail
			GuiControlGet, cpos, adm: Pos, admLBDrucker
			If cPat.Get(AlbisAktuellePatID(), "EMAIL") {
				Gui, adm: Add	, Button       	, % "x" cpx-1 " y+5 w" cpW+2 " vadmLBMail  gadm_LBDruck"                       	, % "Laborblatt per Mail versenden an:"
				Gui, adm: Font	, s8 underline italic cBlue
				Gui, adm: Add	, Text           	, % "x" cpx-1 " y+0 w" cpW+2 " vadmLBMAdr  Center Backgroundtrans"       	, % cPat.Get(AlbisAktuellePatID(), "EMAIL")
				Gui, adm: Font	, s8 Normal cBlack
				GuiControlGet, cpos, adm: Pos, admLBMAdr
			}

			;-: Exploreranzeige des Versand/Dokumentordner bei Bedarf anbieten
			If InStr(FileExist(Addendum.iWin.PatDocs := Addendum.ExportOrdner "\_Befundberichte nicht fertig\(" Addendum.iWin.PatID ") " cPat.Name(Addendum.iWin.PatID, false)), "D")
				Gui, adm: Add	, Button    	, % "x" cpx " y+5 w" cpW+2 " vadmDocFolder  gadm_PatDocsFolder"                 	, % "Dokumentverzeichnis anzeigen"

			;-: Kontaktdaten ins Clipboard
			Gui, adm: Add	, Button       	, % "x" cpx " y+5 w" (cpW+2)/2-5 " vadmAddress  gadm_Address"                      	, % "Kontaktdaten `nins Clipboard"

			;-: Karteikarte drucken
			Gui, adm: Add	, Button       	, % "x+5 w" (cpW+2)/2-5 " vadmKKDruck  gadm_KKDruck"               	, % "Behandlungsdaten`n2010-" SubStr(A_YYYY, 3, 2) " ausgeben"

			;-: sucht in den gesicherten LDT Dateien nach dem Namen des Patienten
			If Addendum.iWin.LDTSearch
				Gui, adm: Add	, Button    	, % "x" cpx-1 " y+10 w" cpW+2 " vadmLDTSearch gadm_LBDruck"                    	, % "vermisste Laborwerte suchen"


			;}

		;-: Tab2 :- Journal                         	;{
			Gui, adm: Tab  	, 2

			bgt := " BackgroundTrans "
			Gui, adm: Font  	, s7 q5 Normal cBlack
			Gui, adm: Add  	, Text      	, % "x10 y25 w" Floor(admWidth/4) 	bgt " vadmJournalTitle   	Section"                     	, % "(es wird gesucht...)"
			Gui, adm: Add  	, Text      	, % "x+2"                                      	bgt " vadmJournalTitle2 	Section"                     	, % "Filter"

			tmps := ["Center AltSubmit", "Alle|Neueste|ohne Klassifikation|mit Klassifikation"]
			Gui, adm: Font  	, s7
			Gui, adm: Add  	, DDL      	, % "x+5"  " y23 h16 w" (12*7) " r4 Choose1 vadmJrnFilter gadm_Journal " tmps.1       	, % tmps.2
			GuiControl, adm: Choose, admJrnFiler, 1
			GuiControlGet  	, cpos, adm: Pos	, admJrnFilter

			Gui, adm: Font  	, s7
			Gui, adm: Add  	, Button    	, % "x" admWidth-145 " y23 h" cposH " w" (11*7) "	vadmJImport  gadm_Journal Center" 	, % "Importieren"
			Gui, adm: Add  	, Button    	, % "x20 y23 h" cposH "                                      	vadmJUpdate gadm_Journal"           	, % "Aktualisieren"

			GuiControlGet  	, cpos, adm: Pos	, admJImport
			GuiControlGet  	, dpos, adm: Pos	, admJUpdate
			GuiControl        	, adm: Move, admJImport, % "x" admWidth - cposW - 5
			GuiControl        	, adm: Move, admJUpdate, % "x" admWidth - cposW - dposW - 5
			Gui, adm: Add  	, Button    	, % "x20 y23 h" cposH "                                       	vadmJOCR gadm_Journal"          	, % "OCR ausführen  "

			GuiControlGet  	, cpos, adm: Pos, admJUpdate
			GuiControlGet  	, dpos, adm: Pos, admJOCR
			GuiControl        	, adm: Move, admJOCR, % "x" cposX - dposW

			Gui, adm: Font  	, s7
			Gui, adm: Add  	, ListView	, % "x2 y+5 		w" admWidth - 6 " h" admHeight - 47 " " admLVJournalOpt                 	, % "Befund|S|Eingang|TimeStamp"

			Journal.Default("admJournal")
			Journal.Arrange(admWidth)
			Gui, adm: ListView, admJournal
			LV_SetImageList(admImageListID)

		;}

		;-: Tab3 :- Protokoll                       	;{
			Gui, adm: Tab  	, 3

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, DDL     	, % "x5 y22  	w100 r5       	           	vadmTPClient gadm_TP "          ; 	, % ""
			Gui, adm: Add  	, Text      	, % "x+2 y26	w40                          	vadmTProtokollTitel " bgt

			Gui, adm: Font  	, s8 q5 Normal cBlack
			Gui, adm: Add  	, Text      	, % "x+10 "                                                                                       	, % "Pat:"
			Gui, adm: Add  	, Edit     	, % "x+0 y22 w170 r1                      	vadmTPSuche"                             	, % ""

			admTPRightX := admWidth-80-5
			Gui, adm: Add  	, Text      	, % "x" admTPRightX " y26 w80          	vadmTPTag       	gadm_TP " bgt 	, % Addendum.iWin.TProtDate
			Gui, adm: Add  	, Button    	, % "x+10 y25 h16                          	vadmTPZurueck	 	gadm_TP"        	, % "<"
			Gui, adm: Add  	, Button    	, % "x+0 	y25 h16						   	vadmTPVor   		gadm_TP"       	, % ">"

			GuiControlGet  	, dpos, adm: Pos, admTPZurueck
			GuiControlGet  	, cpos, adm: Pos, admTPVor
			GuiControl        	, adm: Move, admTPVor    	, % "x" admTPRightX - cposW - 2
			GuiControl        	, adm: Move, admTPZurueck	, % "x" admTPRightX - cposW - dposW - 2

			TPRotCols := "RF|Nachname, Vorname|Geburtstag|Alter|PatID|Uhrzeit"
			Gui, adm: Font  	, s8 q5 Normal cBlack
			Gui, adm: Add  	, ListView	, % "x2 y+5 w" admWidth-6 " h" admHeight-47 " " admLVTProtokoll       	, % TProtCols

			Gui, adm: ListView, admTProtokoll
			LV_ModifyCol(1, 30            	)
			LV_ModifyCol(2, 200         	)
			LV_ModifyCol(3, 70            	)
			LV_ModifyCol(4, 40 " Integer")
			LV_ModifyCol(5, 50 " Integer")
			LV_ModifyCol(6, 60            	)

			GuiControl, adm: ChooseString, admTPClient, % compname
			admGui_TProtokoll(Addendum.iWin.TProtDate, 0, Addendum.iWin.TPClient)

		;}

		;-: Tab4 :- Extras                           ;{
			Gui, adm: Tab  	, 4

			; MODUL BUTTON
			Gui, adm: Font , s10 q5 bold cWhite, % Addendum.Standard.Font
			Gui, adm: Add, Button	, % "x10 y35 w" admWidth-20 " h20 Center hwndadmHPB", % "< - - - - - - - - - - - - - - M O D U L E - - - - - - - - - - - - - - >"  ; vadmMMM
			Opt1 := { 1:0, 2:0xFF35BB34, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1}  ; 0xFF009140
			Opt2 := { 1:0, 2:0xFF35BB34, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1}
			Opt3 := { 1:0, 2:0xFF35BB34, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1}
			ImageButton.Create(admHPB, Opt1, Opt2, Opt3)

			; MODUL BUTTONS (feste Größe)
			Gui, adm: Font  	, s8 q5 Normal cWhite, Calibri ; % Addendum.Standard.Font
			ImageButton.SetGuiColor("Green")
			GuiControlGet, EX, adm: Pos, % admHPB
			Mw := Mh := 15, MEw := MEh :=20
			mStep 	:= 0 	, mxWidth	:= 100, stepX := mxWidth + 5
			mdX  	:= 15	, mdY     	:= mdY := EXY +EXH + 7
			mCols 	:= Floor((admWidth - (mdX*2))/(mxWidth+5))
			stepX 	:= Round(((admWidth - (mdX*2)) - (mCols*mxWidth))/(mCols-1))
			stepX 	:= stepX < 5 ? 5 : stepX

			For iWMIndex, iWModul in Addendum.Module {

				IWx := mStep < mCols ? (mdX + mStep*(mxWidth+stepX)) : (admWidth-mdX-mxWidth)
				Gui, adm: Add, Button, % "x" IWx " y" mdY " w" mxWidth " vadmMX" iWMIndex " hwndadmHPB  gadm_extras", % "  " iWModul.name " " ;Center
				Opt1 := { 1:0, 2:0xFF35BB34, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1, icon:{file:iWModul.ico, x: 9        	, w: Mw		, h: Mh	}}
				Opt2 := { 1:0, 2:0xFF35BB34, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1, icon:{file:iWModul.ico, x: 7, y: 3	, w: MEw   	, h: MEh}}
				Opt3 := { 1:0, 2:0xFF35BB34, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1, icon:{file:iWModul.ico, x: 9        	, w: Mw		, h: Mh	}}
				ImageButton.Create(admHPB, Opt1, Opt2, Opt3)
				iWModul.hwnd := Format("0x{:X}", admHPB)
				mStep ++
				If (mStep = mCols) {
					mStep := 0
					GuiControlGet, EX, adm: Pos, % "admMX" iWMIndex
					mdY := EXY + EXH + 5
				}

			}

			; Tools können ausgeblendet werden
			If Addendum.iWin.ShowTools {

				; TOOL BUTTON  0xFFAfC2ff -> 0xFFAFFFF4
				Gui, adm: Font, s10 q5 bold cBlack, Calibri
				Gui, adm: Add, Button, % "x10 y+15 w" admWidth-20 " h20 Center hwndadmHPB", % "< - - - - - - - - - - - - - - - T O O L S - - - - - - - - - - - - - - - >" ; vadmTTT
				Opt1 := { 1:0, 2:0xFFAFFFF4	, 4:"Black", 5:"H", 6:"WHITE" , 7:"Black", 8:1 	 }
				Opt2 := { 1:0, 2:0xFFAFFFF4	, 4:"Black", 5:"H", 6:"WHITE" , 7:"Black", 8:1 	 }
				Opt3 := { 1:0, 2:0xFFAFFFF4	, 4:"Black", 5:"H", 6:"WHITE" , 7:"Black", 8:1	 }
				ImageButton.Create(admHPB, Opt1, Opt2, Opt3)

				; TOOL BUTTONS
				Gui, adm: Font, s9 q5 Normal cBlack, Calibri
				GuiControlGet, EX, adm: Pos, % admHPB
				IWx := mdX, mdY := EXY + EXH + 7, mStep := 0

				For IWtIdx, IWTool in Addendum.Tools {

					; neuer Button soll nicht aus dem Fenster herausragen
					thisWidth := CalcIdealWidthEx("", IWTool.name,, "s9 q5 Normal", Calibri, 0) + MW
					If (IWtIdx > 1) && (EXW+EXX+thisWidth > admWidth - (2*mdX))
						mStep := 0, IWx := mdX, mdY := EXY + EXH + 5
					else if (IWtIdx > 1)
						IWx := EXX + EXW + 5

					Gui, adm: Add, Button	, % "x" IWx " y" mdY " w" thisWidth " Center vadmTX" IWtIdx " hwndadmHTB gadm_extras", % IWTool.name
					Opt1 := { 1:0, 2:0xFFAFFFF4, 4:"Black", 5:"H", 6:"WHITE" , 7:"Black", 8:1, icon:{file:IWTool.ico, x: 8     	, w: Mw		, h: Mh	}}
					Opt2 := { 1:0, 2:0xFFAFFFF4, 4:"Black", 5:"H", 6:"WHITE" , 7:"Black", 8:1, icon:{file:IWTool.ico, x: 6, y: 4, w: MEw	, h: MEh}}
					Opt3 := { 1:0, 2:0xFFAFFFF4, 4:"Black", 5:"H", 6:"WHITE" , 7:"Black", 8:1, icon:{file:IWTool.ico, x: 8     	, w: Mw		, h: Mh	}}
					ImageButton.Create(admHTB, Opt1, Opt2, Opt3)
					mStep ++
					GuiControlGet, EX, adm: Pos, % "admTX" IWtIdx

				}

			}

		;}

		;-: Tab5 :- Netzwerk                       	;{
			Gui, adm: Tab  	, 5

			dposX := 2, dposY := 60, cMaxW := 0
			Gui, adm: Font  	, s8 q5 Normal underline cBlack, Arial
			Gui, adm: Add   	, Button	, % "x2 y25 vadmLanCheck gadm_Lan", % "Netzwerkgeräte aktualisieren"
			Gui, adm: Font  	, s8 q5 Normal cBlack, Arial

			;-: zeichnet Buttons mit den Clients im Netzwerk und stellt einen weiteren Button für direkten RDP Zugriff bereit ;{
			GuiControlGet, cpos, adm: Pos, % "admLanCheck"
			dposY := cposY+cposH+2
			defaultcmds := "|restart Addendum|send A_TimeIdle"
			For clientName, client in Addendum.LAN.Clients {
				If (compname = clientName)
					continue
				Gui, adm: Add   	, Button	     	, % "x" dposX " y" dposY " Center	vadmClient_" 	clientName " gadm_Lan"	, % clientName
				Gui, adm: Add   	, Button   		, % "x150		  y" dposY " Center	vadmRDP_"   	clientName " gadm_Lan"	, % "RDP"
				Gui, adm: Add   	, Button	    	, % "x+2		  y" dposY " Center	vadmCMDB_" 	clientName " gadm_Lan"	, % "Befehl senden"
				Gui, adm: Add   	, ComboBox 	, % "x+2		  y" dposY "        	vadmCMDE_" 	clientName " gadm_Lan"	, % defaultcmds
				GuiControlGet, cpos, adm: Pos, % "admClient_" clientName
				dposY	:= cposY + cposH
				cMaxW	:= cposW > cMaxW ? cposW : cMaxW
			}

			For clientName, client in Addendum.LAN.Clients {
				If (compname = clientName)
					continue
				GuiControl    	, adm: Move     	, % "admClient_" 	clientName,	% "w" cMaxW
				GuiControl    	, adm: Move     	, % "admRDP_" 	 	clientName,	% "x"	 dposX + cMaxW + 2
				GuiControlGet	, epos, adm: Pos	, % "admRDP_"  	clientName
				GuiControl    	, adm: Move     	, % "admCMDB_" 	clientName,	% "x"	 eposX + eposW + 2
				GuiControlGet	, epos, adm: Pos	, % "admCMDB_"	clientName
				GuiControl    	, adm: Move     	, % "admCMDE_"	clientName,	% "x"	 eposX + eposW +2 " w" admWidth-eposX-eposW-2-5
				GuiControlGet	, cpos, adm: Pos	, % "admClient_" 	clientName
				Gui, adm: Add, Picture, % "x10 y" cposY+2 " w" cposW-4 " h" cposH-4 " Backgroundtrans vadmConn_" clientName

			}

			;-: Ausgabe für Netzwerkoperationen
			Gui, adm: Add, Edit, % "xm y+3 w" admWidth-6 " h" 100 " Backgroundtrans vadmNetMsg hwndadmHNetMsg"

			GuiControlGet, cpos, adm: Pos, % "admNetMsg"
			GuiControl, adm: Move, % "admNetMsg", % "h" admHeight-eposY-15
			DllCall("HideCaret","Int", admHNetMsg)

		 ;}

			;}

		;-: Tab6 :- Info                           	;{
			Gui, adm: Tab  	, 6

			tW 	:= A_ScreenWidth > 1920 ? 150 :130
			fS 	:= A_ScreenWidth > 1920 ? 11 : 10
			fWS	:= A_ScreenWidth > 1920 ? "         " : "                        "
			YPlus := 5
			Sprechstunde := Addendum.Praxis.Sprechstunde[A_DDDD]
			Gui, adm: Font  	, % "s13 q4 bold underline cBlack", % Addendum.StandardFont
			Gui, adm: Add  	, Text, % "x10 y30 w" admWidth-20 " ", % (" " A_DDDD ", " A_DD "." A_MM "." A_YYYY . fWS)

			Gui, adm: Font  	, % "s" fs " normal bold cBlack"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus+8 " w" tw, % "⌚ Sprechstunde"
			Gui, adm: Font  	, Normal
			Gui, adm: Add  	, Text, x+5, % ((StrLen(Sprechstunde) > 0) ? ("von " Sprechstunde " Uhr") : ( "heute keine Sprechstunde"))

			Gui, adm: Font  	, % "s" fs " bold cDarkBlue"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw, % "✍ Tagesprotokoll"
			Gui, adm: Font  	, Normal
			TProto := TProtGetDay(A_DD "." A_MM "." A_YYYY, compname)
			Gui, adm: Add  	, Text, x+5, % (TProto.MaxIndex()="" ? 0 : TProto.MaxIndex()) " Patienten"

			Gui, adm: Font  	, % "s" fs " bold cGreen"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw, % "⚕ Patienten ges."
			Gui, adm: Font  	, Normal
			Gui, adm: Add  	, Text, x+5, % cPat.ItemsCount()

			Gui, adm: Font  	, % "s" fs " bold cRed"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw, % "✉ Signaturen"
			Gui, adm: Font  	, Normal
			Gui, adm: Add  	, Text, x+5, % Addendum.PDF.SignatureCount ", Seitenzahl: " Addendum.PDF.SignaturePages

			Gui, adm: Font  	, % "s" fs " bold cAA2277"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw " hwndadmHMT", % "☀ Monitor" (Addendum.Monitor.Count()>1 ? "e" : "")

			GuiControlGet, cp, adm: Pos, % admHMT
			cpy -= 1

			Loop % Addendum.Monitor.Count() {
				Gui, adm: Font	, % "s" fs+12 " Normal"
				Gui, adm: Add 	, Text, % "x" tw+15 " y" cpy-3 " BackgroundTrans", % SubStr("⚀⚁⚂⚃⚄⚅", A_Index, 1)
				Gui, adm: Font	, % "s" fs
				Gui, adm: Add 	, Text, % "x+2  y" cpy+8 " hwndadmHPB", % Addendum.Monitor[A_Index]
				GuiControlGet, dp, adm: Pos, % admHPB
				cpy += dph+4
			}

			Gui, adm: Font  	, % "s" fs " bold c172842"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw " hwndadmHMT", % "⚗ Laborabruf"
			Gui, adm: Font  	, % "s" fs " Normal"
			Gui, adm: Add   	, Text, % "x+5 w" admWidth-tw-10 " vadmILAB1 " , % "letzter Abruf:`n "
			Gui, adm: Add   	, Text, % "y+2 w" admWidth-tw-10 " vadmILAB2 " , % "nächster Abruf:`n"

			GuiControlGet, cp, adm: Pos, % "admILAB2"

			Gui, adm: Font  	, % "s" fs " bold c172842"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw      	, % "🔄 Infofenster"
			Gui, adm: Font  	, % "s" fs-1 " Normal"
			Gui, adm: Add  	, Text, % "x+5 w" admWidth-tw-10 " h" fs*3 " vadmIWInit"	, % "[hparent" ":" Addendum.iWin.Init ":0]"

		;}

		;-: Tab7 :- Impfstatistik                  	;{
			If Addendum.iWin.Impfstatistik {
				Gui, adm: Tab  	, 7
				Gui, adm: Add   	, Text 	, % "x-5 y25 w" admWidth " h" admHeight-25 " cWhite Backgroundtrans hwndadmHEmbeddedWeb"
				If IsObject(neutron)
					neutron := ""
			}
		;}

		Gui, adm: Tab
		Addendum.iWin.paint ++

	}
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeigen
	;---------------------------------------------------------------------------------------------------------------------------------------------;{
		; MDIFrame handle gehört dem aktuellen AfxMDIFrame Steuerelement ?
			hMDIFrameC := Controls("AfxMDIFrame1401", "ID", AlbisMDIChildGetActive())
			If (hMDIFrame <> hMDIFrameC)
			GuiControl, adm:, admIWInit, % "[hparent: " hMDIFrame ", hadm:" hadm  "]"

		; Gui dem aktuellen AfxMidiFrame als Child zuordnen
			SetParentByID(hMDIFrame, hadm)
			WinSet, Style   	, 0x54000000	, % "ahk_id " hadm
			WinSet, ExStyle 	, 0x0802009C	, % "ahk_id " hadm  ; evtl. +MDIChildWindow 0x40 und/oder AppWindow 0x80000

		; den letzten Tab nach vorne holen
			admGui_ShowTab(Addendum.iWin.firstTab)
			If admGui_ActiveTab("Extras")
				OnMessage(0x200, "admGui_OnHover")
			else if admGui_ActiveTab("Impfen")
				admGui_Impfen()

		; Fenster anzeigen
			Gui, adm: Show, % "x" Addendum.iWin.X " y" Addendum.iWin.Y " w" Addendum.iWin.W  " h" Addendum.iWin.H
									. 	 " NA " admShow, AddendumGui

		; Neuzeichnen der Gui abwarten
			Sleep 300

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tabs / Listviews mit Daten befüllen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		; Statistik
			Addendum.iWin.Init ++
			GuiControl, adm:, admIWInit, % "[Painted: " Addendum.iWin.paint "x, Init: " Addendum.iWin.Init "x, reChecks: "Addendum.iWin.ReCheck "x]"
			Addendum.iWin.Check := Addendum.iWin.ReCheck := 0

		; Inhalte auffrischen und anzeigen
			admGui_Journal(false)	; ohne ReIndizierung
			If !Addendum.Flags.ImportFromJrnl {
				PatDocs := admGui_Reports()
				; diese nur Zeichnen wenn das Tab aktiv ist
				If admGui_ActiveTab("Protokoll")
					admGui_TProtokoll(Addendum.iWin.TProtDate, 0, Addendum.iWin.TPClient)
				else if admGui_ActiveTab("Info")
					admGui_CountDown()
			}
	;}

	; Journalimport bei Bedarf fortsetzen
		If Addendum.Flags.ImportFromJrnl
			Journal.ShowImportStatus(Addendum.Flags.ImportFromJrnl)

	; PDF Thumbnailpreview
		;OnMessage(0x200, "admGui_OnHover")

return
}

; -------- Gui Labels / gFunktionen                                                        	;{
admLAN:                                                                                        	;{	zeichnet den LAN Tab

	; Array mit verfügbaren Netzwerkgeräten
		LanClients := ClientsOnline()

		GuiControlGet, cpos, adm: Pos, AdmLanT1
		Gui, adm: Font  	, s8 q5 Normal cBlue, Arial

	; Anzeigen der verfügbaren Geräte und deren Optionen
		For ClientIndex, onlineclient in LanClients {

			onlineclient := StrReplace(onlineclient, "-")
			YPlus := (ClientIndex = 1 ? 5 : 2), Y := cposY + cposH + YPlus

			Gui, adm: Add, Progress	, % "x10 y" Y     	 " w114 cCCCC00 vAdmLanP" ClientIndex               	, 100
			Gui, adm: Add, Text      	, % "x17 y" Y + 2 " w110 BackgroundTrans vAdmLanC" ClientIndex  	, % onlineclient

			GuiControlGet	, cpos, adm: Pos, % "AdmLanP" ClientIndex
			GuiControl    	, adm: Move, % "AdmLanP"  ClientIndex, % "h" cposH + 4

			Gui                	, adm: Add, Checkbox	, % "x+3 y" Y " h" cposH + 4 " w10 vAdmLanAdm" ClientIndex
			GuiControlGet	, cpos, adm: Pos, % "AdmLanAdm" ClientIndex

			If Addendum.LAN.Clients[onlineclient].remoteShutdown
				Gui, adm: Add, Button	, % "x+10 y" Y " h" cposH + 4 " Checked0 vAdmLanSd" ClientIndex 	, % "PC herunterfahren"

		}

		Gui, adm: Font  	, s9 q5 Normal underline cBlack, Arial
		Y := cposY + cposH + 10
		Gui, adm: Add    	, Text, % "x10 y" Y " w110 BackgroundTrans Center vAdmLanT2", % "OFFLINE"

		GuiControlGet, cpos, adm: Pos, AdmLanT2
		Gui, adm: Font  	, s8 q5 bold cSilver, Arial

	; zeigt offline Netzwerkgeräte an
		OffLineIndex := 0
		For clientName, prop in Addendum.LAN.Clients {

			found := false
			For idx, onlineclient in LanClients
				If (StrReplace(onlineclient, "-") = clientName) {
					found := true
					break
				}

			If !found {
				OfflineIndex	++
				ClientIndex	++
				YPlus := (OfflineIndex = 1 ? 5 : 2), Y := cposY + cposH + YPlus
				Gui, adm: Add, Progress	, % "x10 y" Y      	" w114 cGray vAdmLanP" ClientIndex                 	, 100
				Gui, adm: Add, Text      	, % "x17 y" Y + 2	" w110 BackgroundTrans vAdmLanC" ClientIndex  	, % clientName
				GuiControlGet, cpos, adm: Pos, % "AdmLanP" ClientIndex
				GuiControl, adm: Move, % "AdmLanP"  ClientIndex, % "h" cposH + 4
			}

		}

return ;}

adm_Lan:                                                                                       	;{	Netzwerk Tab 	- gLabel

		If (A_GuiControl = "admLanCheck")                                              	{ 	; zeigt an welche Netzwerkgeräte online sind
		admGui_CheckNetworkDevices()
	}
	else if RegExMatch(A_GuiControl, "admRDP_(?<lient>.*)", c)       	{
		If FileExist(RDPFile := Addendum.LAN.Clients[compname].rdpPath "\" client ".rdp")
			Run % q RDPFile q
		else
			PraxTT("Eine .rdp Datei mit Namen <" client "> ist nicht vorhanden.", "3 1")
	}
	else if RegExMatch(A_GuiControl, "i)admCMDB_(?<lient>.*)", c)	{	; Befehle über das Netzwerk an andere PC's senden

		Gui, adm: Submit, NoHide
		;~ RegExMatch("admCMDE_" client, "(?<cmd> )")
		If !IsObject(admTcp)
			admTCP := Object()
		If !IsObject(admTcp[client]) {
			admTcp[client] := new SocketTCP()
			admTcp[client].connect(Addendum.LAN.Clients[client].ip, Addendum.LAN.Clients[client].port)
			admTcp[client].onrecv := Func("admGui_Receive").Bind(compname)
		}
		vctrl := "admCMDE_" client
		SciTEOutput(A_ThisFunc ": " vctrl " = " %vctrl%)
		If (ip_cmd := %vctrl%) {
			Edit_Append(admHNetMsg, "send cmd: " ip_cmd " to " client " [" Addendum.LAN.Clients[client].ip "]`n`r")
			admTcp[client].sendText("[" compname "] " ip_cmd)
		}

	}

return ;}

adm_Reports:                                                                                  	;{ Patienten Tab 	- gLabel

		If (A_GuiControl = "admReports" && A_EventInfo = 0)
			return

	; zurück wenn Datei nicht existiert
		pdfPath := ""
		LV_GetText(admFile, EventInfo := A_EventInfo, 1)
		For RepIdx, file in PatDocs
			If InStr(file.name, admFile)
				break

		If !FileExist(Addendum.BefundOrdner "\" file.name) 	{
			PraxTT(	"Die Dateioperation ist nicht möglich,`nda die Datei nicht mehr vorhanden ist.", "3 2")
			admGui_Reports()
			RedrawWindow(hadm)
			return
		}

	; PDF Vorschau
		If Instr(A_GuiEvent, "DoubleClick")
			admGui_View(file.name)

return ;}

adm_Journal:                                                                                   	;{ Journal Tab   	- gLabel

		Critical


		; das Triggern durch mehrfach Events vermeiden
		If (RegExMatch(A_GuiControl, "admJournal") && !A_EventInfo)
			return

		; ausgewählte(s) Dokument(e) ermitteln
		If RegExMatch(A_GuiControl, "(admJournal|admJImport)") {
				admFile := Journal.GetSelected()
				If !IsObject(admFile) && !FileExist(Addendum.BefundOrdner "\" admFile) {
					PraxTT("Das ausgewählte Dokument ist nicht vorhanden!", "3 1")
					Journal.RemoveFromAll(admFile)  ; entfernt Datei aus den Objekten, Listviews
					LastJournal_EventInfo := ""
					return
				}
		}

		;~ SciTEOutput("admFile: " admFile)
		Switch A_GuiControl {

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Listview Events
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			case "admJournal":

					; Text des Importieren Buttons ändern
					LastJournal_EventInfo := A_EventInfo
					GuiControl, adm: , admJImport, % "Importieren" (IsObject(admFile) ? " [" admFile.Count() "]" : "")

					; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					; Kontextmenu
					; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					If InStr(A_GuiEvent	, "RightClick") && (StrLen(admFile)>0 || IsObject(admFile)) {
						If !Addendum.Flags.ImportFromJrnl    ; es läuft kein Import zur Zeit
							Journal.ShowMenu(admFile)
						else
							PraxTT("Es läuft ein Importvorgang!", "2 1")
						return
					}

					; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					; Klick auf den Listviewheader - Spaltensortierung
					; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					else 	If 	InStr(A_GuiEvent	, "ColClick")    	{
						If LastJournal_EventInfo
							journal.Sort(LastJournal_EventInfo)
						LastJournal_EventInfo := ""
						return
					}

					; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					; PDF/Bild-Programm aufrufen
					; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					else If Instr(A_GuiEvent	, "DoubleClick")
						admGui_View(admFile)

				;}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Listview Filter
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			case "admJrnFilter":

				JrnFilter := Journal.GetFilter()
				If (JrnFilter <> Journal.lastFilter) {
					Journal.lastFilter := JrnFilter
					Journal.UseFilter(JrnFilter)
				}

			;}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Dokumente importieren
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			case "admJImport":

					LastJournal_EventInfo := ""
					If Addendum.Flags.ImportFromJrnl {
						PraxTT("Es läuft bereits ein Importvorgang!", "2 1")
						return
					}

					fno := 1
					MsgBox, 0x1004, 	% "Ausschluß vom Dokumentimport"
											, 	% "Sollen nur die vollständig mit Name, Vorname, Dokumentbezeichnung und Datum benannten Dokumente importiert werden?"
					IfMsgBox, No
						fno := 0
					admGui_ImportJournalAll(admFile, "FullnamedOnly=" fno " Debug=true" )

			;}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Befundordner indizieren
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			case "admJUpdate":

				If !Addendum.Flags.ImportFromJrnl
					admGui_Reload()
				LastJournal_EventInfo := ""

			;}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; OCR starten/abbbrechen
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
			case "admJOCR":

				LastJournal_EventInfo := ""
				If Addendum.Thread["tessOCR"].ahkReady() {
					MsgBox, 4	, Addendum für Albis on Windows, % "Soll die laufende Texterkennung`nabgebrochen werden?"
					IfMsgBox, No
						return
					Addendum.Thread["tessOCR"].ahkTerminate[]
					PraxTT("Texterkennung wurde abgebrochen", "3 1")
					Addendum.tessOCRRunning := false
					journal.OCR("+OCR ausführen")
				}
				else
					admGui_OCRAllFiles()

			;}

		}


return ;}

adm_BefundImport:                                                                             	;{ läuft wenn Befunde oder Bilder importiert werden sollen

		If (PatDocs.MaxIndex() = 0) || Addendum.Flags.ImportFromPat 	; versehentlichen Aufruf bei leerem Array verhindern
			return

		PreReaderListe := admGui_GetAllPdfWin()              	; derzeit geöffnete Readerfenster einlesen
		admGui_ImportGui(true, "...importiere alle Befunde")	; Hinweisfenster anzeigen
		Imports := admGui_ImportFromPatient()                   	; Importvorgang starten
		Journal.InfoText()                                                      	; Kurzinfo Journal und Patient aktualisieren
		admGui_ImportGui(false)                                        	; Hinweisfenster wieder schliessen
		admGui_Journal(false)                                             	; Journalinhalt auffrischen
		admGui_ShowPdfWin(PreReaderListe)                     	; holt den/die PdfReaderfenster in den Vordergrund

return ;}

adm_Tabs()                                                                	{               	;	Tab            	- gLabel

	global neutron, admWidth, admHeight, admHEmbeddedWeb, PatDocs
	global adm, admTabs, admTPClient, admTPTag, covax, fn_AutoSwitchTab, fn_AutoTabSwitch
	static vacstatsHTML

	Critical

	If (A_GuiEvent = "Normal") {

		; OnMessage für alle Tabs abschalten
		OnMessage(0x200, "")

		; aktuelles Tab ermitteln und speichern
		GuiControlGet, aTab, adm:, admTabs
		Addendum.iWin.lastTab 	:= RegExMatch(Addendum.iWin.firstTab, "i)(Patient|Journal|Protokoll|Impfen)")  	; ? wozu das?
													 ? Addendum.iWin.firstTab : Addendum.iWin.lastTab                                    	; ? wozu das?
		Addendum.iWin.firstTab 	:= RegExMatch(aTab, "i)(Patient|Journal|Protokoll|Impfen)")								; ? wozu das?
													 ? aTab : Addendum.iWin.firstTab																	; ? wozu das?

		; Tabauswahl
		Switch aTab               	{

			case "Info":
				admGui_CountDown()
				AutoSwitchTab := 60

			case "Extras":
				OnMessage(0x200, "admGui_OnHover")
				AutoSwitchTab := 60

			case "Protokoll":
				GuiControl, adm: , admTPClient	, % Addendum.iWin.TPClient
				GuiControl, adm: , admTPTag    	, % (Addendum.iWin.TProtDate := !Addendum.iWin.TProtDate ? "Heute" : Addendum.iWin.TProtDate)
				admGui_TProtokoll(Addendum.iWin.TProtDate, 0, Addendum.iWin.TPClient)

			case "Netzwerk":
				AutoSwitchTab := 180

			case "Impfen":
				admGui_Impfen()

		}

	}

	; Info, Extras, Netzwerk Tab nur eine zeitlang anzeigen
	If AutoSwitchTab {
		fn_AutoSwitchTab := Func("admGui_ShowTab").Bind(Addendum.iWin.lastTab)
		SetTimer, % fn_AutoSwitchTab, % -1*AutoSwitchTab*1000
	}

	; AutoTabSwitch -> nach Patient, wenn Dokumente vorhanden sind
	If (PatDocs.Count()>0 && !IsObject(fn_AutoTabSwitch))
			admGui_AutoTabSwitch()

return
}

adm_Statistics()                                                          	{

	;~ global admVacDate, adm, admVacOut

	;~ Critical

	;~ Gui, adm: Submit, NoHide

	;~ SciTEOutput("gc: " A_GuiControl ", " admVacDate)

	;~ If (A_GuiControl = "admVacStats") {

		;~ vaccinationsDay := neutron.doc.getElementById("vacdate").innerText
		;~ If !RegExMatch(vaccinationsDay, "\d{1,2}\.\d{1,2}\.\d{2}|\d{4}") {
			;~ MsgBox, Dies ist kein gültiges Datumsformat
			;~ return
		;~ }

		;~ If RegExMatch(vaccinationsDay "#", "(?<ddMM>\d{1,2}\.\d{1,2}\.)(?<YH>\d{2})#", dt_)
			;~ vaccinationsDay := dt_ddMM SubStr(A_YYYY, 1, 2) dt_YH

		;~ admGui_Impfstatistik(vaccinationsDay)

	;~ }


}

adm_TP()                                                                   	{               	;	Protokoll Tab 	- gLabel

	global admTPClient

	Critical

	TProtDate := Addendum.iWin.TProtDate
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If         	InStr(A_GuiControl, "admTPZurueck")	             	; einen Tag zurück
			admGui_TProtokoll(TProtDate, -1, Addendum.iWin.TPClient)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else If	InStr(A_GuiControl, "admTPVor")    	             	; einen Tag weiter
		If (TProtDate = "Heute")
			return
		else
			admGui_TProtokoll(TProtDate, 1, Addendum.iWin.TPClient)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else if InStr(A_GuiControl, "admTPTag")                      	; Datum geklickt
		If (TProtDate = "Heute")
			return
		else
			admGui_TProtokoll("Heute", 0, Addendum.iWin.TPClient)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else if InStr(A_GuiControl, "admTPClient")   {                	; Clientwechsel
		Gui, adm: Submit, NoHide
		Addendum.iWin.TPClient := admTPClient
		admGui_TProtokoll(TProtDate, 0, admTPClient)
	}

return
}

adm_TProtLV()                                                              	{               	; Tagesprotokoll 	- gLabel

	admGui_Default("admTProtokoll")

	If Instr(A_GuiEvent, "DoubleClick") {
		LV_GetText(PatID, A_EventInfo, 5)
		AlbisAkteOeffnen("", PatID)
	}

return
}

adm_LBDruck()                                                              	{               	; Laborblattdruck - gLabel

	 ; admLBDrucker 	= DDL Druckertreiber
	 ; admLBSpalten 	= zu druckende Spalten des Laborblatts
	 ; admLBAnm    	= Anmerkungen/Probedaten (Checkbox)

		global adm, admLBSpalten, admLBDrucker, admLBAnm, admLBDUD, DDLPrinter

	; sichert die Einstellung bei Änderung des Druckertreibers
		Critical
		Gui, adm: Submit, NoHide

		If (A_GuiControl = "admLBDrucker") {
			If (Addendum.iWin.LBDrucker <> admLBDrucker) {
				IniWrite, % admLBDrucker, % Addendum.Ini, % compname, % "Infofenster_Laborblatt_Drucker"
				Addendum.iWin.LBDrucker := admLBDrucker
			}
		}

	; wenn Zähler 0 zeigt "Alles" anzeigen
		else If (A_GuiControl = "admLBDUD") {
			If (admLBDUD < 1)
				GuiControl, adm:, admLBSpalten, % "Alles"
			GuiControl, adm: Focus, admLBDUD
		}

	; Drucken ausführen
		else If (A_GuiControl = "admLBD")
			admGui_LaborblattDruck(admLBSpalten, admLBDrucker, admLBAnm)

	; Mailversand
		else If (A_GuiControl = "admLBMail") {

			;~ savePath 	:= admGui_LaborblattDruck(admLBSpalten, "Microsoft Print to PDF", admLBAnm)

			PatID     	:= AlbisAktuellePatID()
			savePath 	:= AlbisLaborblattPDFDruck(admLBSpalten, admLBAnm)
			;~ SciTEOutput(savePath)
			If (!savePath || !FileExist(savePath) || !RegExMatch(savePath, "\\")) {
				PraxTT("Der E-Mail-Versand der Laborwerte ist nicht möglich!`nDas Laborblatt konnte nicht exportiert werden.", "4 1")
				return
			}
			RegExMatch(PDFText:=IFilter(savePath), "geb\.\s+[\d\.]+\s*[\n\r\s]+(?<Datum>[\d\.]+)", Laborblatt_)
			PraxTT("erstelle Outlook-EMail ...", "0 1")
			MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
			MailItem.Recipients.Add(cPat.Get(PatID, "EMAIL"))
			MailItem.attachments.Add(savePath)
			MailItem.Subject := "Ihre Laboruntersuchung vom " Laborblatt_Datum
			body := "Sehr geehrte" (cPat.Get(PatID, "GESCHLECHT")=2 ? " Frau" : "r Herr") " "
			body .= cPat.Get(PatID, "VORNAME") " " cPat.Get(PatID, "NACHNAME") ",`n`n"
			body .= "wie besprochen sende ich Ihnen die Ergebnisse Ihrer Laboruntersuchung vom " Laborblatt_Datum " zu."
			body .= "`n`n`nMit freundlichen Grüßen`n" StrReplace(Addendum.Praxis.MailStempel, "##", "`n")
			MailItem.body := body
			MailItem.display
			PraxTT("", "off")

		}

	; Checkbox Status sichern
		else if (A_GuiControl = "admLBAnm")
			Addendum.iWin.LBAnm := admLBAnm

return
}

adm_Extras()                                                               	{     	      		; Extras Tab     	- gLabel (Startet Module/Tools)

	; letzte Änderung: 19.07.2022

	global fn_AutoSwitchTab

	If !RegExMatch(A_GuiControl, "adm(?<Obj>[A-Z]+)(?<Nr>\d+)", M)
		return

	Switch MObj 	{

	 ; Module
		case "MX":

			EModulExe    	:= Addendum.Module[MNr].command
			EModulName	:= Addendum.Module[MNr].name

			If FileExist(EModulExe) {

				; Skriptmodul ausführen
				If !WinExist(EModulName " ahk_class AutoHotkeyGUI") {
					PraxTT("starte Modul : " EModulName, "2 1")
					SplitPath, EModulExe, WorkPath
					Run, % EModulExe, % WorkPath,, EModulPID
				}

				; Skript DocPrinz die aktuelle PatientenID übergeben
				If Instr(EModulExe, "DocPrinz") {

					while !(EModulHwnd := WinExist(DokExClass := "DocPrinz ahk_class AutoHotkeyGUI")) && (A_Index < 40)
						Sleep 100

					If EModulHwnd {
						WinActivate   	, % DokExClass
						WinWaitActive	, % DokExClass,, 7
						Send_WM_CopyData("search|" AlbisAktuellePatID() "|" GetAddendumID() "|admGui_Callback", EModulHwnd)
					}

				}

			}

		; Tools
		case "TX":

			EToolExe       	:= Addendum.Tools[MNr].command
			EToolName   	:= Addendum.Tools[MNr].name

			If FileExist(EToolExe) {
				PraxTT("starte Tool: " EToolName, "1 1")
				SplitPath, EToolExe, WorkPath
				Run, % EToolExe, % WorkPath
			}
			else {
				SplitPath, EToolExe,, EToolPath
				PraxTT(EToolName " ist nicht vorhanden`n[" EToolPath "\]", "1 1")
			}

	}

	; schaltet nachdem ein Programm
		fn_AutoSwitchTab := Func("admGui_ShowTab").Bind(Addendum.iWin.lastTab)
		SetTimer, % fn_AutoSwitchTab, % -1*AutoSwitchTab*1000

return
}

adm_Address(PatID:="", fields:="")                                         	{               	; kopiert die Kontaktdaten des aktuellen Pat. ins Clipboard

	If !isObject(Addendum.AdressGui)
		Addendum.AdressGui := Object()

	; übergebene PatID nutzen oder aus dem Albisfenstertitel holen, anschließend die PatID prüfen
	PatID := !PatID || !PatID ~= "^\d+$" ? AlbisAktuellePatID() : PatID
	If (!PatID || !cPat.Exist(PatID)) {
		PraxTT((!PatID 	?	"Die Daten von " AlbisCurrentPatient() " konnten nicht ausgelesen werden."
										:	"Die Patienten ID (" PatID ") gehört zu keinem Patienten.") , "2 1")
		return
	}

	; einen Datensatz aus Patient.dbf holen und Rückgabeparameter prüfen
	sPat := cPat.GetAdditionalData(PatID, "")
	If !IsObject(sPat) {
		PraxTT("Die Daten von " AlbisCurrentPatient() " konnten nicht aus der Datenbank gelesen werden.", "2 1")
		return
	}

	For sIndex, sData in sPat {
		If (PatID = sData.NR) {
			nPat := sData
			break
		}
	}


	PatData := Object()
				; Adressdaten

	PatData.Adresse :=                                     "`nPatID:         `t"          	PatID
									.                                   	 "`nName:         `t"            (nPat.TITEL  ?   nPat.TITLE " "  : "") nPat.NAME ", " nPat.VORNAME
									.                                   	 "`nGeburtstag:   `t"           	ConvertDBASEDate(nPat.GEBURT)
									.                                    	 "`nAdresse:      `t"           	nPat.STRASSE .  (nPat.HAUSNUMMER ? " " nPat.HAUSNUMMER : "") " in " nPat.PLZ " " nPat.ORT
									. 	(nPat.ENTFERNUNG  	?              "`nEntfernung:   `t"           	nPat.ENTFERNUNG : "")

	If nPat.STRASSE
		PatData.Postanschrift := (nPat.ANREDE ? RegExReplace(nPat.ANREDE, "^Herrn", "Herr") " " : "") (nPat.TITEL ? nPat.TITEL " " : "")  nPat.VORNAME ", " nPat.NAME "`n" nPat.STRASSE (nPat.HAUSNUMMER ? " " nPat.HAUSNUMMER : "") "`n`n" nPat.PLZ " " nPat.ORT
	If nPat.POSTFACH
		PatData.Postanschrift_Postfach := (nPat.ANREDE ? RegExReplace(nPat.ANREDE, "^Herrn", "Herr") " " : "") (nPat.TITEL ? nPat.TITEL " " : "")  nPat.VORNAME ", " nPat.NAME "`n" nPat.POSTFACH "`n`n" nPat.PLZ " " nPat.ORT


				; Kommunikationsdaten
	PatData.Kommunikation := 	nPat.Telefon1 || nPat.Telefon2
													|| nPat.TELEFAX  ?           	"`nTelefone:           `t" RTrim((nPat.Telefon1       	? "T1: "  	nPat.Telefon1 ", " : "")
																																						. 	      	 (nPat.Telefon2       	? "T2: "  	nPat.Telefon2 ", " : "")
																																						. 	      	 (nPat.TELEFAX        	? "Fax: " 	nPat.TELEFAX       : ""))
									.  	(cPat.EMail        ?             	"`nEMail:              `t"      	nPat.EMAIL : "")

				; Stammdaten
	PatData.Stammdaten  :=                              	"`nStammdaten:         `t"     	"VNR: "	nPat.VNR ", VKNR: "  nPat.VKNR
									.	                                  	"`nKVK eingelesen:     `t"	   	ConvertDBASEDate(nPat.LESTAG)
									. 						                				"`nletzte Behandlung:  `t"     	nPat.LASTBEH
									.	                                  	"`ngültig bis:         `t"    	nPat.GUELTIG
									.	 						                      	"`nin Behandlung seit:`t"     	ConvertDBASEDate(nPat.SEIT)
									.	                                   	"`nPat. seit:          `t"    	nPat.SEIT(PatID)
									. 	(nPat.LASTBEH    ?               	"`nletzte Behandlung:  `t"     	nPat.LASTBEH : "")

				; sonstiges
	PatData.Arbeit  := 	(nPat.ARBEIT     ?               	"`nArbeit:             `t"    	nPat.ARBEIT : "")

	For theme, data in PatData
		outdata .= (outdata ? "`n" : "")  "`n-------------------------------------------------------`n" theme ":`n-------------------------------------------------------`n" data "`n"

	ClipBoard := outdata
	ClipWait, 1
	PraxTT("Die Daten von " nPat.Anrede " " nPat.NAME ", " nPat.VORNAME " *" ConvertDBASEDate(nPat.Geburt) " [" PatID "] `nsind jetzt aus dem Clipboard abrufbar.", "2 1")

	adressgui := new addressGui(PatData)
	Addendum.AddressGui[adressgui.guihwnd] := PatID

}

class addressGui                                                            {

	__New(PatData) {

		gcount := 0
		this.PatData := PatData
		this.ref := Object()

		Gui New, HWNDguihwnd -DPIScale
		this.guihwnd := guihwnd
		Gui, %guihwnd%: Default


		;~ hwndVarName := "Adr_" this.guihwnd "_hBtn" gcount
		For theme, pData in PatData {

			gcount ++

			Gui, Font, s11 q5 Bold

			Gui Add, Text, % "xm " (gcount = 1 ? "ym" : "y+15") " BackgroundTrans hwndhStatic" , % theme ":"
			cp := GuiControlGet(guihwnd, "Pos", hStatic)
			Gui, Font, s10 q5 Normal
			Gui Add, Button, % "x800 y" cp.Y-10 " hwndhCtrl "   	, % "...diese Daten kopieren"

			glabelfunc := ObjBindMethod(this, "addressClipboardHandler", theme, hCtrl)
			GuiControl, %guihwnd%: +g, % hCtrl, % glabelfunc

			this.ref[hCtrl] := {"theme":theme, "glabelfunc":glabelfunc}  ; Referenz speichern

			Gui, Font, s10 q5 Normal
			Gui Add, Edit, % "xm y+3 w1000 r" StrSplit(pData, "`n").Count() " ReadOnly "	, % pData

		}

		Gui Add, Button, % "xm y+15 hwndhCtrl", % "Fenster schließen"
		glabelfunc := ObjBindMethod(this, "addressClipboardHandler", "Close", hCtrl)
		GuiControl, %guihwnd%: +g, % hCtrl, % glabelfunc
		this.ref[hCtrl] := {"theme":"Close", "glabelfunc":glabelfunc}

		Gui, %guihwnd%: Show, AutoSize, % "fixe Patientendaten für die langsame Bürokratie (WinID: " this.guihwnd ")"
		SendInput, {Tab} ; die automatische Textauswahl des letzten Editsteuerelementes verhindern
		ControlFocus, % "ahk_id " hCtrl

	}

	addressClipboardhandler(theme, hCtrl) {

			Critical

			SciTEOutput("Ja bin dabei")

			if (theme = "Close") {
				guihwnd := this.guihwnd
				Gui, %guihwnd%: Destroy
				Addendum.AddressGui.Delete(guihwnd)
				for hCtrl, item in this.ref {
					GuiControl, %guihwnd%: -g, % hCtrl, % item.glabelfunc
					ObjRelease(item.glabelfunc)
				}

				this := ""
				return
			}

			clipboard := this.PatData[theme]
			ClipWait, 2
			If (clipboard = this.PatData[theme])
				PraxTT("Die angeforderten Daten zum THema '" theme "' sind ins Clipboard kopiert worden")
			else
				PraxTT("Das Clipboard hat sich nicht mit den Patientendaten befüllen lassen.")

	}


}

adm_KKDruck(PatID:="", from:="", to:="")                                  	{               	; Behandlungsdaten als PDF Dokument ausgeben

}

adm_PatDocsFolder()                                                       	{               	; Dokumentverzeichnis im Explorer öffnen

	If  InStr(A_GuiControl, "admDocFolder")
		Run, % "explorer.exe " q Addendum.iWin.PatDocs q

}

admGuiDropFiles(hwnd,files,elHwnd,dropX,dropY)                            	{               	; Dokumente per Drag-Drop hinzufügen

	; letzte Änderung 27.09.2021:
	; in Funktion umgeschrieben, class Journal erweitert
	; wenn WatchFolder eingeschaltet ist, wird die Datei über die Funktion admGui_FolderWatch() verwaltet

		MsgBox, 0x1004, % StrReplace(A_ScriptName, ".ahk"), % "[Drag'n'Drop Befundordner]`n`n"
								. "Möchten Sie die Datei" (files.Count()>1 ? "en" : "") " im Ursprungsverzeichnis löschen?"
		IfMsgBox, Yes
			RemoveFileOriginals := true
		For idx, fullfilepath in files 	{

			; Datei kommt direkt aus dem Befundordner, nicht aufnehmen
				SplitPath, fullfilepath, filename, filepath, fileext
				If FileExist(Addendum.BefundOrdner "\" filename) || !RegExMatch(filename, "i)\.(pdf|doc|docx|jpg|png|tiff|bmp|wav|mov|avi)$")
					continue

			; benennt die Datei gleich mit dem Namen des aktuellen Patienten, wenn man die Datei auf den Patienten Tab gezogen hatte
				If InStr(A_GuiControl, "admReports")
					filename := PatName ", " RegExReplace(StrReplace(filename, PatName), "^\s*,\s*")

			; Datei sicherheitshalber in den Befund und Backup Ordner kopieren
				FileCopy, % fullfilepath, % Addendum.BefundOrdner "\" filename
				If RemoveFileOriginals
					FileDelete, % fullfilepath

			; wenn das Hotfolder Feature (FolderWatch) ausgeschaltet ist, wird hier das Dokument dem Pool hinzugefügt
				If (!Addendum.OCR.WatchFolderStatus = "running") {
					path := Befunde.Backup(filename)                                	; Backup der Datei anlegen falls bisher nicht erfolgt
					PDFpool.Add(Addendum.BefundOrdner, filename)        	; Dokument dem Pool hinzufügen
					Journal.Add(filename)                                                 	; in den Listview's des Journal und Patient TAB anzeigen
					SciTEOutput(A_ThisFunc ": " path) ; Debug
				}

		}

return
}
;}

; -------- Gui-Funktionen                                                                  	;{
admGui_Default(LVName)                                             	{               	; Gui Listview als Default setzen
	Gui, adm: Default
	Gui, adm: ListView, % LVName
return
}

admGui_ActiveTab(TabName="")                                     	{                 	; welches TAB wird angezeigt
	global adm, admTabs
	Gui, adm: Submit, NoHide
return !TabName ? admTabs : TabName = admTabs ? true : false
}

admGui_ShowTab(TabName, hTab="", callfrom:="")       	{               	; ein bestimmten Tab nach vorne holen

		global adm, admHTab, hadm, fn_AutoSwitchTab, fn_AutoTabSwitch, PatDocs
		global admGuiTabs

		TabToShow	:= (RegExMatch(TabName, "^\d+$") ? TabName : admGuiTabs[TabName])
		hTab         	:= !hTab ? admHTab : hTab

	; werden Dateien umbenannt, wird die Tabumschaltung für 30 Sekunden verzögert
		If WinExist("Addendum (Datei umbenennen)", "ahk_class AutoHotkeyGUI") {
			If !isObject(fn_AutoTabSwitch)
				fn_AutoTabSwitch := Func("admGui_ShowTab").Bind(TabName,, "AutoTabSwitch")
			SetTimer, % fn_AutoTabSwitch, % -1*30*1000
			return
		}

	; AutoSwitchTab Timer beenden, firstTab und lastTab ändern
		If IsObject(fn_AutoSwitchTab) || IsObject(fn_AutoTabSwitch) {
			GuiControlGet, aTab, adm:, admTabs
			Addendum.iWin.lastTab 	:= RegExMatch(Addendum.iWin.firstTab, "i)(Patient|Journal|Protokoll|Impfen)")
													 ? Addendum.iWin.firstTab : Addendum.iWin.lastTab
			Addendum.iWin.firstTab 	:= RegExMatch(aTab, "i)(Patient|Journal|Protokoll|Impfen)")
													 ? aTab : Addendum.iWin.firstTab
			If IsObject(fn_AutoSwitchTab) {
				SetTimer, % fn_AutoSwitchTab, Delete
				fn_AutoSwitchTab := ""
			}
		}

	; TCM_GETCURSEL (0x130B)     	: ausgewählten TAB ermitteln
		SendMessage, 0x130B,,,, % "ahk_id " hTab
		CurrentTab := ErrorLevel

	; zurück wenn Tab schon angezeigt wird
		If (TabToShow = CurrentTab)
			return 1

	; TCM_SETCURFOCUS (0x1330)	: TAB wählen
		SendMessage, 0x1330, % TabToShow,,, % "ahk_id " hTab

	; TCM_GETCURSEL (0x130B)     	: gewählten TAB ermitteln
		SendMessage, 0x130B,,,, % "ahk_id " hTab
		CurrentTab := ErrorLevel

	; aktuellen Tabnamen ermitteln
		For CurrentTabName, tabnr in admGuiTabs
			If (CurrentTab = tabnr)
				break

	; OnMessage nur bei bestimmten Tabs nutzen
		OnMessage(0x200, (CurrentTabName="Extras" ? "admGui_OnHover" : "")) ; schaltet OnMessage ein oder aus

	; automatisch auf den "Patient" Tab schalten, während der Sprechstunde (Datenschutz)
		If vacation.isConsultationTime()
			If (!CurrentTabName~="i)(Journal|Protokoll)" && !IsObject(fn_AutoTabSwitch))
				admGui_AutoTabSwitch()

return CurrentTabName
}

admGui_AutoTabSwitch(state="On",tab="Patient",delay=30){

	global adm, admHTab, hadm, fn_AutoTabSwitch

	If (state~="i)(Off|Delete)") && IsObject(fn_AutoTabSwitch) {
		SetTimer, % fn_AutoTabSwitch, % state
		fn_AutoTabSwitch := ""
		return
	}

	fn_AutoTabSwitch := Func("admGui_ShowTab").Bind(tab,, "AutoTabSwitch")
	SetTimer, % fn_AutoTabSwitch, % -1*delay*1000

}

admGui_ColorRow(hLV, rowNr, paint=true)                      	{                	; eine Zeile einfärben

	global 	hadm
	static 	bcolor := 0x995555, tcolor := 0xFFFFFF

	If !Addendum.iWin.RowColors
		return

	res := paint ? LV_Colors.Row(hLV, rowNr, bcolor, tcolor) : LV_Colors.Row(hLV, rowNr)
	WinSet, Redraw, , % "ahk_id " hadm

}

admGui_Exist()                                                           		{              	; gibt hwnd des Infofenster zurück
	Controls("", "reset", "")
return Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", WinExist("ahk_class OptoAppClass"))
}

admGui_OnHover(lparam, wparam, msg, hwnd)              	{                	; Eingabedialoge anzeigen beim Überfahren eines Symbols anzeigen

	global admHTab, hadm, adm, admTabs
	static lasthwnd, fn_onhover

	hwnd := Format("0x{:X}", hwnd)
	If (lasthwnd = hwnd || !admGui_ActiveTab("Extras"))
		return 0

	lasthwnd 	:= hwnd
	Sleep 50                                   	; kurze Pause für Imagebutton Funktion
	;~ Critical
	If IsObject(fn_onhover) {
		SetTimer, % fn_onhover, Off
		fn_onhover := ""
	}
	MouseGetPos, mx , my,, cHwnd, 2
	For index, modul in Addendum.Module
		If (modul.hwnd = hwnd && modul.name = "in Labordaten suchen") {
			fn_onhover := Func("admGui_OnHoverStart").Bind(modul)
			SetTimer, % fn_onhover, -250
			break
		}

return 0
}

admGui_OnHoverStart(modul)                                         	{               	; Mousehover

	static minis := Object(), miniNr := 0
	static miniText1, miniText2, miniEdit1, miniEdit2, miniBtn1, miniBtn2, hmini1, hmini2
	static modulname, guiname, sstring

	If !RegExMatch(modul.name, "(in Labordaten suchen)")
		return

	m := GetWindowSpot(modul.hwnd)
	modulname := modul.name
	If !IsObject(minis[modul.name]) {

		minis[modul.name] := Object()
		miniNr ++
		guiname := "mini" miniNr
		Gui, % guiname ": New"    	, % "-Caption -DPIScale +ToolWindow +hwndhmini" miniNr
		Gui, % guiname ": Margin"	, 15, 5
		Gui, % guiname ": Color"      	, % Addendum.Default.BGColor3, % Addendum.Default.BGColor2
		Gui, % guiname ": Font"       	, % "s9 q5 cBlack", % Addendum.Default.Font
		Gui, % guiname ": Add"        	, Text 	, % "xm ym vminiText" miniNr, % " Geburtsdatum o. Nachname, Vorname o. ,Vorname"
		Gui, % guiname ": Font"       	, % "s8 q5 c" Addendum.Default.FntColor
		Gui, % guiname ": Add"        	, Edit  	, % "y+2 w200 r1 vminiEdit" miniNr " hwndhedit", % sstring
		Gui, % guiname ": Add"        	, Button	, % "x+5 vminiBtn" miniNr, % "Go"
		Gui, % guiname ": Show"    	, % "x" m.X  " y" m.Y+m.H " Hide", % modul.name

		WinSet, ExStyle, 0x0, % "ahk_id " hedit
		EM_SetMargins(hEdit, 2, 2)
		GuiControlGet, cp, % guiname ": Pos", % "miniText" miniNr
		GuiControlGet, dp, % guiname ": Pos", % "miniBtn" miniNr
		GuiControlGet, ep, % guiname ": Pos", % "miniEdit" miniNr
		GuiControl, % guiname ": Move", % "miniBtn" miniNr	, % "x" (BtnX := cpX+cpW-dpW) " h" epH
		GuiControl, % guiname ": Move", % "miniEdit" miniNr	, % "w" (BtnX-10-cpX)

		hmini := "hmini" miniNr
		minis[modul.name].hwnd  	:= %hmini%
		minis[modul.name].miniNr 	:= miniNr

		g := GetWindowSpot(minis[modul.name].hwnd)
		WinSet, Region, % "0-0 w" g.W " h" g.H " R30-30", % "ahk_id " minis[modul.name].hwnd
		CS_DROPSHADOW := 0x20000
		ClassStyle := GetGuiClassStyle()
		SetGuiClassStyle(minis[modul.name].hwnd, ClassStyle | CS_DROPSHADOW)

		Gui, % guiName ": Show"    	, % "x" m.X+Floor(m.W//2-g.W//2) " y" m.Y+m.H " NA", % modul.name


	}
	else {
		guiname := "mini" minis[modul.name].miniNr
		g := GetWindowSpot(minis[modul.name].hwnd)
		Gui, % guiname ": Show"    	, % "x" m.X+Floor(m.W//2-g.W//2) " y" m.Y+m.H " NA"
	}

	SetTimer, OnHoverGuiOff, -2000

return
OnHoverGuiOff:
	MouseGetPos,,, hWin,, 1
	If (minis[modulname].hwnd = GetHex(hWin)) {
		SetTimer, OnHoverGuiOff, -2000
		return
	}
	Gui, % guiName ": Show"    	, % "Hide"
return
}

admGui_Destroy()                                                             {                	; schließt das Infofenster

	global hadm, hadm2, adm, adm2, RN

	Gui, RN:  Show, Hide
	Gui, RNP:  Show, Hide
	Gui, adm2:  Show, Hide
	Gui, adm:	Destroy
	Addendum.iWin.lastPatID	:= hadm := hadm2 := 0

return
}

;}

; -------- Netzwerk-Funktionen                                                             	;{
admGui_OnAccept(this)                                                 	{

	global admClients, admAccepts

	If !IsObject(admClients)
		admClients := Object()
	If !IsObject(admAccepts)
		admAccepts := Array()

	admAccepts.Push(this.accept())
	CI := admAccepts.MaxIndex()
	admAccepts[CI].onrecv := Func("admGui_Receive").Bind(CI)
	admAccepts[CI].sendText("[" compname "] tell name CI" CI)

}

admGui_OnDisconnect()                                                	{

	global admClients, admAccepts

}

admGui_Receive(from, answer)                                         	{                 	; empfängt und sendet Netzwerknachrichten

	global admClients, admAccepts, admHNetMsg

	msg := answer.recvText()
	SciTEOutput("lan received: " msg " from " from)
	RegExMatch(msg, "i)\[(?<from>.*?)\]\s*(?<cmd>.+?)$", m)

	If RegExMatch(mcmd, "i)tell\s+name\s+CI(?<I>\d+)", C) {
		Edit_Append(admHNetMsg, "send my client name to: " mfrom " " Addendum.LAN.Clients[mfrom].ip "`n`r")
		If !admClients.haskey(mfrom) {
			admAccepts[CI].clientname := mfrom
			admAccepts[CI].onrecv := Func("admGui_Receive").Bind(mfrom)
			admClients[mfrom] := CI
		}
		admAccepts[CI].sendText("[" compname "] CI" CI " name: " compname)
	}
	else if RegExMatch(mcmd, "i)\brestart\s+all") && (compname <> Addendum.LAN.Server.IamServer){
		Edit_Append(admHNetMsg, "cmd received: " mcmd " [" mfrom " " Addendum.LAN.Clients[mfrom].ip "]`n`r")
		PraxTT("Addendum remote restart in 60s", "12 0")
		fn := Func("SkriptReload").Bind("AutoRestart")
		SetTimer, % fn, -60000
	}
	else if RegExMatch(mcmd, "i)\bsend\s+(?<var>.*?)$", m) {
		Edit_Append(admHNetMsg, "cmd received: " mcmd " [" mfrom " " (ip := Addendum.LAN.Clients[mfrom].ip) "]`n`r")
		If ip {
			CI 	:= admClients[mfrom]
			val 	:= %mvar%
			admAccepts[CI].sendText("[" compname "] answer: " mvar " = " val)
		}

	}
	else
		Edit_Append(admHNetMsg, "received: " mcmd " [" mfrom " " Addendum.LAN.Clients[mfrom].ip "]`n`r")


}

admGui_CheckNetworkDevices()                                    	{              	; zeigt an welche Netzwerkgeräte online sind

	global adm, hadm, admHNetMsg

	Gui, adm: Default
	GuiControl, % "adm: Enable0", % "admLanCheck"

	For clientname, client in Addendum.LAN.Clients {
		If (clientName = compname)
			continue
		If (ping := IPHelper.Ping(client.ip))
			Edit_Append(admHNetMsg, "ping: " ping "ms  [" clientname " " client.ip "]`n`r")
		GuiControl, % "adm: Enable" (ping ? 1 : 0), % "admClient_" 	clientName
		GuiControl, % "adm: Enable" (ping ? 1 : 0), % "admRDP_"   	clientName
		GuiControl, % "adm: Enable" (ping ? 1 : 0), % "admCMDB_" 	clientName
		GuiControl, % "adm: Enable" (ping ? 1 : 0), % "admCMDE_" 	clientName
	}

	GuiControl, % "adm: Enable1", % "admLanCheck"

}

;}

; -------- Inhalte erstellen                                                               	;{
admGui_Reload(ReIndex=true)                                  	{                   	; aktualisiert die Inhalte und zeichnet das Fenster neu

	global hadm, PatDocs

	admGui_Journal(ReIndex)
	PatDocs := admGui_Reports()
	RedrawWindow(hadm)

}

admGui_Abrechnungshelfer(PatID, Geburtsdatum)          	{                	; berechnet und zeigt die letzten Untersuchungstermine an

		global adm, admNotes

		;~ If (StrLen(Addendum.Praxis.ArztUser) = 0)
			;~ NR := Addendum.Praxis.StandardArzt

		;~ If InStr(Addendum.Praxis.Arzt[NR].Fach, "Hausarzt")
			;~ return

		Kassenabrechnung := false

		SArten   	:= "Scheinarten: "
		infT       	:= ""
		infTr      	:= "⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈ ⎈`n"
		infT2     	:= ""

		GVUMin 	:= Addendum.GVUminAlter
		AbrSchein := AlbisAbrechnungsscheinAktuell()
		QDaten 	:= QuartalTage({"TDatum": A_YYYY A_MM A_DD, "TFormat": "DBASE"})
		PatGeburt	:= StrReplace(Geburtsdatum, ".")
		PatGeburt	:= SubStr(PatGeburt, 5, 4) SubStr(PatGeburt, 3, 2) SubStr(PatGeburt, 1, 2)
		PatAge    	:= HowLong(PatGeburt, A_YYYY A_MM A_DD)

	; angelegte Abrechnungsscheine werden angezeigt
		If (AbrSchein.Count() > 0)
			For abrIndex, Abr in AbrSchein {
				SArten .= (abIndex > 1 ? "," : "")  (Abr.Scheinart = "Privat" ? "P " Abr.Datum : SubStr(Abr.Scheinart,1,1) " (" Abr.Quartal ")")
				If (Abr.Scheinart = "Abrechnung")
					Kassenabrechnung := true
			}
		else
			SArten .= "keine"

	; bei ausschließlicher Abrechnung nach EBM (Privatpatienten) bricht die Funktion hier ab
		If !Kassenabrechnung
			return SArten

	; Abrechnungsregeln für verschiedene Ziffern
		For key, val in Addendum.PatExtra[PatID] {

			If    	  (key = "Z01732")	{

					gkey := StrReplace(key, "Z")
					gvQ 	:= QuartalTage({"TDatum":val, "TFormat":"DBASE"})
						gvT 	:= ConvertDBASEDate(val) "`tQ " SubStr(gvQ["Aktuell"], 1, 2) "/" SubStr(gvQ["Aktuell"], 3, 2)
					gvuspan	:= HowLong(val, QDaten.DBaseEnd)
					timeString	:= gvuspan.Years "J " gvuspan.Months "M " gvuspan.Days "T`n"
					If (gvuspan.Years >= 3)
						gvZ :=  ((gvuspan.Years >= 3) ? "☘" : "⛔") "   `t" timeString

			}
			else If (key = "Z01746" || key = "Z01745")	{

					hkey	:= StrReplace(key, "Z")
					hkQ	:= QuartalTage({"TDatum":val, "TFormat":"DBASE"})
					hkT 	:= ConvertDBASEDate(val) "`tQ " SubStr(hkQ["Aktuell"], 1, 2) "/" SubStr(hkQ["Aktuell"], 3, 2)
					hkspan	:= HowLong(val, QDaten.DBaseEnd)
					timeString	:= "vor " hkspan.Years "J " hkspan.Months "M " hkspan.Days "T`n"
					If (hkspan.Years >= 3)
						hkZ := "HKS(" key ")" ((hkspan.Years >= 3) ? "☘" : "⛔") "   `t" timeString

			}
			else If (key = "Z01740")	{

					coZ := "01740: "
					cokey := StrReplace(key, "Z")
					coQ := QuartalTage({"TDatum":val, "TFormat":"DBASE"})
					coZ := cokey ":#" ConvertDBASEDate(val) " `tQ " SubStr(coQ["Aktuell"], 1, 2) "/" SubStr(coQ["Aktuell"], 3, 2) "`n"

			}
			else If (key = "Z03220")	{

					crzQ1 := QuartalTage({"TDatum":val, "TFormat":"DBASE"})
					cz1 := "03220: #" ConvertDBASEDate(val) " `tQ " SubStr(crzQ1["Aktuell"], 1, 2) "/" SubStr(crzQ1["Aktuell"], 3, 2) "`n"

			}
			else If (key = "Z03221")	{

					crzQ2 := QuartalTage({"TDatum":val, "TFormat":"DBASE"})
					cz2 := "03221: #" ConvertDBASEDate(val) " `tQ " SubStr(crzQ2["Aktuell"], 1, 2) "/" SubStr(crzQ2["Aktuell"], 3, 2) "`n"

			}

		}

	; Regeln der Vorsorgeuntersuchungen
		If      	(PatAge.Years >= GVUMin) && (StrLen(gkey) > 0) && (StrLen(hkey) > 0) && (gvQ.Aktuell = hkQ.Aktuell ) {

			tabs := "`t`t`t"
			infT1 .= gkey "/" hkey . " `t" . gvT "`n"
			If gvz
				infT2 .= "GVU(" gkey ") / HKS(" hkey ") " gvz "`n"

		}
		else if 	(PatAge.Years >= GVUMin) && (StrLen(gkey) > 0) && (StrLen(hkey) > 0) {

			tabs := "`t"
			infT1 .= gkey . tabs . gvl "`n" hkey . tabs . hkl "`n"
			If hkZ
				infT2 .= "GVU(" gkey ") " gvz "`n" hkZ "`n"

		}
		else if 	(PatAge.Years >= GVUMin) && (StrLen(gkey) =0) && (StrLen(hkey) = 0) {
			infT2 .= "GVU / HKS - bisher noch ohne GVU`n"
		}

	; Text zusammensetzen
		infT .= StrReplace(coz, "#", tabs)
		infT .= infT1
		infT .= StrReplace(cz1, "#", tabs)
		infT .= StrReplace(cz2, "#", tabs)

		If IstChronischKrank(PatID) {
			If (crzQ1.Aktuell <> QDaten.Aktuell)
				infT2 .= "Chronikerpauschale 1: ☘`t`tQ " . SubStr(crzQ1["Aktuell"], 1, 2) "/" SubStr(crzQ1["Aktuell"], 3, 2) "`n"
			If (crzQ2.Aktuell <> QDaten.Aktuell)
				infT2 .= "Chronikerpauschale 2: ☘`t`tQ " . SubStr(crzQ2["Aktuell"], 1, 2) "/" SubStr(crzQ2["Aktuell"], 3, 2) "`n"
		}

		If !InStr(infT, "01740") && (PatAge.Years >= 55)
				infT2 .= "01740: ☘" tabs " ! ! ! ! "

return SArten "`n" . infT . infTr . (Kassenabrechnung ? infT2 : "")
}

admGui_CountDown()                                                   	{              	; berechnet die Zeit bis zum nächsten Laborabruf

	global hadm, adm, admILAB1

	IniRead, LCall1	, % Addendum.Ini, % "LaborAbruf", % "Letzter_Abruf"
	IniRead, LCall2	, % Addendum.Ini, % "LaborAbruf", % "Letzter_Abruf_ohne_Daten"
	IniRead, LCall3	, % Addendum.Ini, % "LaborAbruf", % "Letzter_Abruf_mit_Daten"

	pdate 	:= A_YYYY A_MM A_DD "000000"
	heute 	:= A_MM A_DD
	pdate 	+= -1, days
	gestern	:= SubStr(pdate, 5, 2) SubStr(pdate, 7, 2)
	pdate 	+= +2, days
	morgen	:= SubStr(pdate, 5, 2) SubStr(pdate, 7, 2)

	If RegExMatch(LCall1, "(?<Y>\d{4})-(?<M>\d{2})-(?<D>\d{2})\s+(?<H>\d{2}):(?<Min>\d{2}):(?<S>\d{2})[\s\|]*", T) {
		Abruftag := TM TD
		lastLBCall := "letzter Abruf:`n"
						. 	(	Abruftag = heute  	? "heute"
							: 	Abruftag = gestern 	? "gestern" : "am " TD "." TM "." (TY = A_YYYY ? "" : SubStr(TY, 3, 2)))
						. 	" um " TH ":" TMin " Uhr, "
						. 	(LCall1 = LCall2 ? "ohne neue" : LCall1 = LCall3 ? "mit neuen" : "??") " Daten"
	}


	If !Addendum.Labor.AbrufTimer
		nextLBCall := "der zeitgesteuerte Abruf ist aus"
	else {
		IniRead, nextCall, % Addendum.Ini, % "LaborAbruf", % "naechster_Abruf"
		If RegExMatch(nextCall, "(?<D>\d{2})\.(?<M>\d{2})\.(?<Y>\d{4})\s*,\s*(?<hm>\d+:\d+)", T) {
			Abruftag 	:= TM TD
			nextLBCall	:= "nächster Abruf:`n"
							. 	( Abruftag = heute  	? "heute, "
								: Abruftag = morgen	? "morgen, " : "am " TD "." TM "." (TY = A_YYYY ? "" : TY))
							.	  " um " Thm " Uhr"

		}
	}

	GuiControl, adm:, admILAB1, % lastLBCall
	GuiControl, adm:, admILAB2, % nextLBCall

return
}

ExecuteFunc(neutron, event, params*)                            	{               	; ruft eine Funktion auf wenn diese noch nicht ausgeführt wird

	; verhindert einen Funktionsaufruf solang die Funktion noch ausgeführt wird

	function := "admGui_" event.target.innerText
	If !Addendum.Flags[function]
		%function%(event, params*)

}

admGui_Impfstatistik(event:="", call:="")                           	{               	; Impfstatistik mit KVB-Webformularbefüllung

		global 	adm, neutron, fn_AutoTabSwitch, fn_AutoSwitchTab
		global 	covax
			static 	COVIDStatsPath
			static	calls := 0

		If !COVIDStatsPath
			COVIDStatsPath := Addendum.DBPath "\sonstiges\COVID19-Schutzimpfungen.json"

		If !isObject(covax) {
			If FileExist(COVIDStatsPath)
				covax := JSONData.Load(COVIDStatsPath, "", "UTF-8")
			else
				covax := Object()
		}

		If Addendum.Flags[A_ThisFunc]
			return
		Addendum.Flags[A_ThisFunc] := true

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Progresselement ändern
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		If (request := IsObject(event) || call = "newVacDate" ? true : false) {

			; Autoumschalten auf einen anderen Tab verhindern wenn die Impfdaten frisch erstellt werden
			If IsObject(fn_AutoTabSwitch) {
				SetTimer, % fn_AutoTabSwitch, Delete
				fn_AutoTabSwitch := ""
			}
			If IsObject(fn_AutoSwitchTab) {
				SetTimer, % fn_AutoSwitchTab, Delete
				fn_AutoSwitchTab := ""
			}
			admGui_ImpfProgress(1)

		}
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Datum ermitteln
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If IsObject(neutron) && (request || call = "load") {
			event.preventDefault()    ; verhindert das versehentliche Auslösen einer Default-Aktion
			element	:= neutron.doc.getElementById("vacdate")
			vacDate  	:= element.value
		}
		else if RegExMatch(event, "\d{2}\.\d{2}\.(\d{2}|\d{4})") {
			vacDate   	:= event
		}
		else {
			PraxTT(A_ThisFunc ": falscher Aufruf der Funktion!", "2 2")
			Addendum.Flags[A_ThisFunc] := false
			return
		}

		If !RegExMatch(vacDate, "\d{2}\.\d{2}\.\d{2}|\d{4}") {
			PraxTT("COVID19-Impfstatístik`nDatumsformat ungültig: <<" vacDate ">>", "2 3")
			Addendum.Flags[A_ThisFunc] := false
			return
		}

		If (call = "load" && !IsObject(covax[ConvertToDBASEDate(vacDate)])) {
			PraxTT("COVID19-Impfstatístik`nDatum: " vacDate " ohne erstellte Statistik bisher. ", "2 3")
			Addendum.Flags[A_ThisFunc] := false
			return
		}

		covaxDate := ConvertToDBASEDate(vacDate)
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Datum sichern
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If (request && Addendum.iWin.Impfstatistikdatum <> vacDate)
			IniWrite, % (Addendum.iWin.Impfstatistikdatum := vacDate), % Addendum.Ini, Infofenster, Impfstatistikdatum
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Progresselement ändern
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If request
			admGui_ImpfProgress(2)
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Patientendaten laden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		;~ If !IsObject(PatDB)                         ; # wird nicht benötigt?
			;~ PatDB := ReadPatientDBF(Addendum.AlbisDBPath,,, vacDate-1)
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Progresselement ändern
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If request
			admGui_ImpfProgress(20)
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Statisitik erstellen                              (Datenbank auslesen)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If request {
			praxis := new AlbisDB()
			covax[covaxDate] := praxis.Impfstatistik(vacDate)
			praxis := ""
		}
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Progresselement ändern
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If request
			admGui_ImpfProgress(100)
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Ausgabe vorbereiten                               (Summen bilden)
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		vax := covax[covaxDate]
		; - - - - - - - - - - - - -
		; Tabellte 1
		; - - - - - - - - - - - - -
		B  	:= vax.B, M := vax.M, A := vax.A, J := vax.J
		tBs	:= B.NR.1	+ B.NR.2	+ B.NR.3
		tMs	:= M.NR.1+ M.NR.2	+ M.NR.3
		tAs	:= A.NR.1	+ A.NR.2	+ A.NR.3
		tJs 	:= J.NR.2	+	J.NR.3

		; - - - - - - - - - - - - -
		; Tabellte 2
		; - - - - - - - - - - - - -
		VOs	:= vax.VO.1	+ vax.VO.2	+ vax.VO.3
		VMs	:= vax.VM.1	+ vax.VM.2	+ vax.VM.3
		VJs	:= vax.VJ.1	+ vax.VJ.2 	+ vax.VJ.3
		VKs	:= vax.VK.1	+ vax.VK.2 	+ vax.VK.3
		VS1 	:= vax.VO.1	+ vax.VJ.1	+ vax.VK.1	+ vax.VM.1
		VS2 	:= vax.VO.2	+ vax.VJ.2	+ vax.VK.2	+ vax.VM.2
		VS3 	:= vax.VO.3	+ vax.VJ.3	+ vax.VK.3	+ vax.VM.3
		Vs 	:= VS1     	+ VS2        	+ VS3
		noData := !Vs ? true : false
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Ausgabe als HTML-Tabelle im Infofenster
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If (IsObject(neutron) && (request || call = "load")) {

			; - - - - - - - - - - - - -
			; Tabellte 1
			; - - - - - - - - - - - - -
			neutron.doc.getElementById("tB1").innerText 	:= B.NR.1     	? B.NR.1 : ""
			neutron.doc.getElementById("tB2").innerText 	:= B.NR.2     	? B.NR.2 : ""
			neutron.doc.getElementById("tB3").innerText 	:= B.NR.3     	? B.NR.3 : ""
			neutron.doc.getElementById("tBs").innerText	:= tBs           	? tBs : ""
			neutron.doc.getElementById("tM1").innerText 	:= M.NR.1     	? M.NR.1 : ""
			neutron.doc.getElementById("tM2").innerText 	:= M.NR.2     	? M.NR.2 : ""
			neutron.doc.getElementById("tM3").innerText 	:= M.NR.3     	? M.NR.3 : ""
			neutron.doc.getElementById("tMs").innerText	:= tMs          	? tMs : ""
			neutron.doc.getElementById("tA1").innerText 	:= A.NR.1     	? A.NR.1 : ""
			neutron.doc.getElementById("tA2").innerText 	:= A.NR.2     	? A.NR.2 : ""
			neutron.doc.getElementById("tAs").innerText	:= tAs            	? tAs : ""
			neutron.doc.getElementById("tJ2").innerText 	:= J.NR.2      	? J.NR.2 : ""
			neutron.doc.getElementById("tJs").innerText 	:= tJs             	? tJs : ""

			; - - - - - - - - - - - - -
			; Tabelle2
			; - - - - - - - - - - - - -
			; >=60
			neutron.doc.getElementById("tAO1").innerText	:= vax.VO.1 	? vax.VO.1 : ""
			neutron.doc.getElementById("tAO2").innerText	:= vax.VO.2 	? vax.VO.2 : ""
			neutron.doc.getElementById("tAO3").innerText	:= vax.VO.3 	? vax.VO.3 : ""
			neutron.doc.getElementById("tAOs").innerText	:= VOs           	? VOs : ""
			; 18-59
			neutron.doc.getElementById("tAM1").innerText	:= vax.VM.1    	? vax.VM.1 : ""
			neutron.doc.getElementById("tAM2").innerText	:= vax.VM.2    	? vax.VM.2 : ""
			neutron.doc.getElementById("tAM3").innerText	:= vax.VM.3    	? vax.VM.3 : ""
			neutron.doc.getElementById("tAMs").innerText	:= VMs          	? VMs : ""
			; 12-17
			neutron.doc.getElementById("tAJ1").innerText	:= vax.VJ.1    	? vax.VJ.1 : ""
			neutron.doc.getElementById("tAJ2").innerText	:= vax.VJ.2    	? vax.VJ.2 : ""
			neutron.doc.getElementById("tAJ3").innerText	:= vax.VJ.3    	? vax.VJ.3 : ""
			neutron.doc.getElementById("tAJs").innerText	:= VJs              	? VJs : ""
			; 5-11
			neutron.doc.getElementById("tAK1").innerText	:= vax.VK.1    	? vax.VK.1 : ""
			neutron.doc.getElementById("tAK2").innerText	:= vax.VK.2    	? vax.VK.2 : ""
			neutron.doc.getElementById("tAK3").innerText	:= vax.VK.3    	? vax.VK.3 : ""
			neutron.doc.getElementById("tAKs").innerText	:= VKs             	? VKs : ""
			; Summen
			neutron.doc.getElementById("tS1").innerText	:= VS1         	? VS1 : ""
			neutron.doc.getElementById("tS2").innerText	:= VS2         	? VS2 : ""
			neutron.doc.getElementById("tS3").innerText	:= VS3         	? VS3 : ""
			neutron.doc.getElementById("tSs").innerText 	:= Vs            	? Vs : ""

			; Progress ändern
			If request {
				progressEL.value := 100
				ControlClick, Internet Explorer_Server1, % "ahk_id " neutron.hwnd,, Left, NA
				Sleep 120
				ControlClick,, % "ahk_id " neutron.hwnd,, Left
				stopwatch.innerText := Round((A_TickCount - tck)/1000, 2) "s"
				vacstateEL	:= neutron.doc.getElementById("vacstate")
				vacstateEL.focus
				vacstateEL.click()
				vacstateEL.innerText := "COVID19-IMPFSTATISTIK vom " vacDate (noData ? " [keine Daten erfasst]" : "")
			}

		}
		else if (request && !nodata) {
			x := "   `t|  `t"
			y := "        `t      |  `t    "
			z := " `t     |  "
			t .= " Datum: " vacDate "`n`n"
			t .= " Hersteller [Impfstoff]`t| Erstimpfung`t| Abschlußimpfung | Auffrischimpfung | gesamt`n"
			t .= " --------------------------------------------------------------------------------------------`n"
			t .= " Biontech [Comirnaty] `t| "		(B.NR.1 	? B.NR.1 	: "") 	x  (B.NR.2 ? B.NR.2	: "") 	y (B.NR.3	? B.NR.3	: "")   	z  (tBs 	? tBs	: "")	"`n"
			t .= " Moderna [Spikevax]   `t| "		(M.NR.1	? M.NR.1	: "") 	x  (M.NR.2? M.NR.2 	: "") 	y (M.NR.3	? M.NR.3	: "")   	z  (tMs 	? tMs: "")	"`n"
			t .= " Astrazeneca [Spikevax]`t| " 	(A.NR.1	? A.NR.1	: "")   	x  (A.NR.2	? A.NR.2 	: "") 	y (A.NR.3	? "--"     	: "--") 	z  (tAs 	? tAs	: "")  	"`n"
			t .= " Johnson [Johnson]     `t| "                  "--"   	               	x  (J.NR.2	? A.NR.2 	: "") 	y (J.NR.3	? J.NR.3	: "")      	z  (tJs	? tJs 	: "") 	"`n"
			t .= " --------------------------------------------------------------------------------------------`n"
			t .= " Altersgruppe    `t|          `t|          `t  |          `t     | gesamt`n"
			t .= " --------------------------------------------------------------------------------------------`n"
			t .= "     >60         `t|       "         	vax.VO.1 		x 	vax.VO.2     	y 	vax.VO.3      	z 	VOs      	"`n"
			t .= "     <18         `t|       "         	vax.VJ.1   		x 	vax.VJ.2      	y 	vax.VJ.3        	z 	VJs          	"`n"
			t .= "     18-59       `t|       "         	vax.VM.1 		x 	vax.VM.2     	y 	vax.VM.3      	z 	VMs        	"`n"
			t .= "     alle          `t|       "          	vax.VS.1 		x 	vax.VS.2    	y 	vax.VS.3       	z 	vax.sum  	"`n"
			FileOpen(A_Temp "\COVID19Impfstatistik.txt", "w", "UTF-8").Write(t)
			Run, notepad.exe %A_Temp%\COVID19Impfstatistik.txt
		}
	;}

		; sichert berechnete Statistiken
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If request && !noData
			JSONData.Save(COVIDStatsPath, covax, true,, 1, "UTF-8")

		Addendum.Flags[A_ThisFunc] := false

		If (noData || call="load")
			return
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Ausfüllen des KVB-WEB Formulars
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		admGui_RKIWebassist:
		MsgBox, 0x1044, % "RKI/KBV Webassist", % 	"Setze zuerst den Cursor ins Feld: 'Biontech Erstimpfungen'`n"
																					.     	"Drücke dann Taste 'S' um die Daten automatisch eintragen zu lassen.`n"
																					.     	"(Escape bricht ab.)`n`n"
																					.     	"Jetzt loslegen?"
		IfMsgBox, No
			return

		msg  	:=  "Cursor ins erste Feld [Biontech Erstimpfungen] setzen.`nDann 's' für Start drücken.`n"
		msgB 	:=  "Jetzt starten mit Taste 's'.`n"

		Loop {
			If Mod(A_Index, 3) = 0 {
				MouseGetPos, mx, my
				ToolTip, msg . "(mit Escape hier abbrechen)", % mx, % my, 12
			}
			If GetKeyState("Esc")
				return
			else if GetKeyState("S")
				break
			else if GetKeyState("LButton")
				msg := msgB
			Sleep 40
		}

		ToolTip,,,, 12
		sltime := 80
		SendInput	, {LControl Down}a{LControl Up}
		Sleep % sltime

		; BIONTECH
		; ------------------------;{
		If vax.B.NR.1 {
			SendRaw	, % vax.B.NR.1
			Sleep % sltime
		}
		SendInput	, {TAB}
		Sleep % sltime
		If vax.B.NR.2 {
			SendRaw	, % vax.B.NR.2
			Sleep % sltime
		}
		SendInput	, {TAB}
		Sleep % sltime
		If vax.B.NR.3 {
			SendRaw	, % vax.B.NR.3
			Sleep % sltime
		}
		SendInput	, {TAB}
		Sleep % sltime
		;}

		; AstraZeneca
		; ------------------------;{
		If vax.A.NR.1 {
			SendRaw	, % vax.A.NR.1
			Sleep % sltime
		}
		SendInput	, {TAB}
		Sleep % sltime
		If vax.A.NR.2 {
			SendRaw	, % vax.A.NR.2
			Sleep % sltime
		}
		SendInput	, {TAB}
		Sleep % sltime
		;}

		; Moderna
		; ------------------------;{
		If vax.M.NR.1 {
			SendRaw	, % vax.M.NR.1
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		If vax.M.NR.2 {
			SendRaw	, % vax.M.NR.2
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		If vax.M.NR.3 {
		SendRaw	, % vax.M.NR.3
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		;}

		; Johnson & Johnson
		; ------------------------;{
		If vax.J.NR.2 {
		SendRaw	, % vax.J.NR.2
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		;}

		; BioNTech/Pfizer für 5 bis 11-Jährige
		; ------------------------;{
		SendInput, {TAB}
		Sleep % sltime
		SendInput, {TAB}
		Sleep % sltime
		;}

		; Altersverteilung
		; ------------------------;{
		If vax.VO.1 {
			SendRaw	, % vax.VO.1
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		If vax.VO.2 {
			SendRaw	, % vax.VO.2
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		If vax.VO.3 {
			SendRaw	, % vax.VO.3
			Sleep % sltime
		}

		SendInput, {TAB}
		Sleep % sltime

		If vax.VJ.1 {
			SendRaw	, % vax.VJ.1
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		If vax.VJ.2 {
			SendRaw	, % vax.VJ.2
			Sleep % sltime
		}
		SendInput, {TAB}
		Sleep % sltime
		If vax.VJ.3 {
			SendRaw	, % vax.VJ.3
			Sleep % sltime
		}
		;}

		MsgBox, 0x1004, % A_ThisFunc, % "Nicht korrekt ausgefüllt?`n"
														. 	"Noch ein Versuch?"
		IfMsgBox, Yes
			goto admGui_RKIWebassist

	;}


return covax
}

admGui_ImpfProgress(percent)                                      	{               	; ändert den Fortschrittsbalken

	; soll auch als Callbackfunktion eingesetzt werden

	global 	adm, neutron
	global 	covax
	static 	progressEL, stopwatch, tck, lastProgressposition

	If (percent=1) {

		tck := A_TickCount
		progressEL	:= neutron.doc.getElementById("vacstatsprogress")
		stopwatch 	:= neutron.doc.getElementById("stopwatch")
		progressEL.style.Width 	:= "0%"
		stopwatch.innerText 	:= "0.00s"
		lastProgressPosition 	:= 0
		return

	}
	else if !percent
		return lastProgressposition

	progressEL.style.Width  	:=  (lastProgressPosition := percent) "%"
	If (percent > 19 && percent < 90)
		progressEL.style.background :=  "linear-gradient(rgb(47, 175, 175), rgb(96, 226, 226), rgb(47, 175, 175))"
	If (percent > 95)
		progressEL.style.background :=  "linear-gradient(rgb(68, 175, 47), rgb(119, 235, 104), rgb(41, 156, 66))"
	stopwatch.innerText   	:= "[" Round((A_TickCount - tck)/1000, 2) "s]"

}
;}

; -------- Gui Tabs                                                                       	;{
admGui_Reports()                                                         	{	            	; zeigt Befunde des Patienten aus dem Befundordner

		global 	PatDocs
		static 	PdfImport 		:= false
		static 	ImageImport	:= false
		static 	rxPerson1    	:= "[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL-]+)*"
		static 	rxPerson2    	:= "[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"

		journal.Default("admReports")
		LV_Delete()

	; Pdf Befunde des Patienten ermitteln und entfernen bestimmter Zeichen aus dem Patientennamen für die fuzzy Suchfunktion
		PatDocs	:= Array()
		PatID    	:= AlbisAktuellePatID()
		PatNV   	:= RegExReplace(cPat.Get(PatID, "NAME"), "[\s\-]") . RegExReplace(cPat.Get(PatID, "VORNAME"), "[\s\-]")

	; PatDocs erstellen - enthält nur die Befunde/Bilder des aktuellen Patienten
		For key, pdf in ScanPool	{				; ohne PatID ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)
			If !RegExMatch(pdf.name, "i)\.pdf$") || !FileExist(Addendum.BefundOrdner "\" pdf.name) {
				pdfpool.remove(pdf.name)
				continue
			}
			RegExMatch(pdf.name, "^\s*(?<Nachname>" rxPerson2 ")[\,\s]+(?<Vorname>" rxPerson2 ")", doc)  ; ## ersetzen
			a := StrDiff(PatNV, RegExReplace(docNachname docVorname, "[\s\-]"))
			b := StrDiff(PatNV, RegExReplace(docVorname docNachname, "[\s\-]"))
			If (a < 0.11) || (b < 0.11)
				PatDocs.Push(pdf)
		}

	; PDF Befunde anzeigen
		For key, pdf in PatDocs {
			If !RegExMatch(pdf.name, "i)\.pdf$") || !FileExist(Addendum.BefundOrdner "\" pdf.name)
				continue
			displayname := xstring.Replace.Names(xstring.Replace.FileExt(pdf.name))
			If pdf.isSearchable
				LV_Add("ICON3", pdf.pages " S:" , displayname)
			else
				LV_Add("ICON1", pdf.pages " S:" , displayname)
		}

	; Bilddateien aus dem Befundordner einlesen
		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(A_LoopFileName, "\.jpg|png|tiff|bmp|wav|mov|avi$") {
				RegExMatch(A_LoopFileName, "^\s*(?<Nachname>" rxPerson2 ")[\,\s]+(?<Vorname>" rxPerson2 ")", Such) ; ## ersetzen
				SuchNV := RegExReplace(SuchNachname SuchVorname, "[\s\-]")
				SuchVN := RegExReplace(SuchVorname SuchNachname, "[\s\-]")
				a	:= StrDiff(PatNV, SuchNV), b	:= StrDiff(PatNV, SuchVN)
				Diff := (a <= b) ? a : b
				If (Diff < 0.11) {
					LV_Add("ICON2",,  xstring.Replace.Names(A_LoopFileName)) ; entfernt den Namen
					PatDocs.Push({"name": A_LoopFileName})
				}
			}

	; InfoText auffrischen
		Journal.InfoText()

return PatDocs
}

admGui_Journal(ReIndex=false)                                    	{	            	; befüllt das Journal mit Pdf und Bildbefunden aus dem Befundordner

		global admHJournal

		Journal.Default()			     		; LISTVIEW: DEFAULT MACHEN
		Journal.Empty()							; LISTVIEW: LEEREN
		pdfpool.Refresh(ReIndex)	     	; SCANPOOL OBJECT AUFFRISCHEN		---
		journal.OCR("-OCR ausführen") ; PDF BEFUNDE DES PATIENTEN HINZUFÜGEN
		journal.AddDocumentsAll()		; DOKUMENTE IM BEFUNDORDNER AUFLISTEN
		journal.AddPictures() 	        		; BILD-DOKUMENTE HINZUFÜGEN
		Journal.InfoText()			    		; INFOTEXT AUFFRISCHEN
		journal.Sort(0, true)		    		; ANZEIGE WIRD MIT DER GESPEICHERTEN SPALTENSORTIERUNG ANGEZEIGT (Addendum.ini)

return
}


admGui_TProtokoll(datestring, addDays, TPClient)             	{                	; listet alle Karteikarten des übergebenen Datums auf

	; letzte Änderung: 06.11.2021

	global admTProtokoll, admHTProtokoll, admTPClient, admTProtokollTitel, admTPTag
	static last_prep

	admGui_Default("admTProtokoll")

	datestring	:= Addendum.iWin.TProtDate := !datestring ? "Heute" : datestring
	datestring	:= (datestring="Heute" || StrLen(datestring)=0) ? A_DD "." A_MM "." A_YYYY : datestring
	lDate     	:= RegExMatch(datestring, "(?<dd>\d+)\.(?<MM>\d+)\.(?<YYYY>\d+)", d) ? dYYYY dMM ddd : A_YYYY A_MM A_DD
	If addDays
		lDate += addDays, days
	lDate := SubStr(lDate, 1, 8)

	RegExMatch(ldate, "^(?<YYYY>\d{4})(?<MM>\d{2})(?<dd>\d{2})", d)
	FormatTime, weekdayshort, % lDate, ddd
	FormatTime, weekdaylong	, % lDate, dddd
	FormatTime, month       	, % lDate, MMMM

	; doppelten Aufruf des TAB verhindern
	IF (last_prep =  lDate "|" TPClient)
		return

	; Tagesprotokoll auslesen
	TProto := TProtGetDay(lDate, TPClient)

	; stellt nur Clients zur Auswahl bereit bei denen Protokolleinträge vorhanden sind
	GuiControl, adm:, admTPClient, % "|"
	PClients := compname
	For clientName, clientdata in Addendum.LAN.Clients {
		If (compname=clientName || InStr(PClients, clientName))
			continue
		oTmp  	:= TProtGetDay(lDate, clientName)
		PClients.= (oTmp.Count()>0 ? "|" clientName : "")
	}
	GuiControl, adm:, admTPClient	, % "|" PClients
	GuiControl, adm: ChooseString	, admTPClient, % TPClient

	; Tagesprotokoll anzeigen
	Gui, adm: ListView, admTProtokoll
	LV_Delete()
	SubStrFactor := StrLen(TProto.Count())-1 < 1 ? 1 : StrLen(TProto.Count())-1
	For idx, PatIDAndTime in TProto {
		RegExMatch(PatIDAndTime, "^(?<atID>\d+)\(*(?<Time>\d+:\d+)*", P)
		If !PatID
			continue
		LV_Insert(1,, SubStr("0000000" idx, -1*SubStrFactor)            	; Index absteigend
						 , cPat.NAME(PatID)                                           	; Name
						 , cPat.GEBURT(PatID, true)	                        		; Geburtsdatum
						 , Age(cPat.GEBURT(PatID), datestring)	               	; Alter an diesem Tag
						 , PatID 						                    					; Patientennummer
						 , PTime)                                                              	; Uhrzeit des ersten Karteikartenabrufes
	}

	LV_ModifyCol(2, "Auto")
	If (LV_GetColWidth(admHTProtokoll, 2) < 140)
		LV_ModifyCol(2, 140)

	TPTag := SubStr(lDate, 1, 8) = SubStr(A_Now, 1, 8) ? "Heute" : weekdayshort ", " ddd "." dMM "." dYYYY
	GuiControl, adm: , admTProtokollTitel	, % "[" (TProto.Count()-1 < 0 ? 0 : TProto.Count()-1) " Pat.]"
	GuiControl, adm: , admTPTag            	, % TPTag

	last_prep                        	:= lDate "|" TPClient
	Addendum.iWin.TPClient 	:= TPClient
	Addendum.iWin.TProtDate	:= TPTag

return
}

admGui_Impfen()                                                          	{               	; zeigte eine HTML-Seite im Tab an

	global neutron, admHEmbeddedWeb, covax, admWidth, admHeight
	static vacstatsHTML

	 ; HTML laden
		If !vacstatsHTML
			vacstatsHTML := FileOpen(Addendum.Dir "\include\Gui\Impfstatistik.html", "r", "UTF-8").Read()

		If !IsObject(neutron) {
			neutron := new NeutronWindow(vacstatsHTML,,, "Coronaschutzimpfungsstatistik", "-Resize +Parent" admHEmbeddedWeb)
			WinSet, Top  	,                   	, % "ahk_id " neutron.hwnd
			WinSet, Style 	, 0x54000000	, % "ahk_id " neutron.hwnd
			WinSet, ExStyle	, 0x08000000	, % "ahk_id " neutron.hwnd
			neutron.Show("w" admWidth " h" admHeight)
			SetWindowPos(neutron.hwnd, 0, 0, admWidth, admHeight-25)
			Addendum.iWin.Impfstatistikdatum := RegExMatch(Addendum.iWin.Impfstatistikdatum, "\d{1,2}\.\d{1,2}\.(\d{2}|\d{4})")
																											? Addendum.iWin.Impfstatistikdatum : A_DD "." A_MM "." A_YYYY
			neutron.doc.getElementById("vacdate").innerText := Addendum.iWin.Impfstatistikdatum
			Redraw(neutron.hwnd)

		}

		Sleep 200
		Redraw(neutron.hwnd)

		admGui_Impfstatistik(Addendum.iWin.Impfstatistikdatum, "load")

}
;}

; -------- Kontextmenu                                                                     	;{
admGui_CM(MenuName)                                               	{                 	; Kontextmenu des Journal

		; Tastaturkürzel in allen Tabs (01.07.2020)
		; Mehrfachauswahl ermöglicht (26.09.2021)

		global 	admHJournal, admHReports, admFile, hadm, PatDocs
		static 	newadmFile, rowNr

		Addendum.iWin.firstTab := "Journal"
		If RegExMatch(MenuName, "i)^J")
			Journal.Default("admJournal")

		blockthis 	:= false
		;~ If !admFile {
			;~ journal.Default("admJournal")
			;~ admFile := Journal.GetSelected()
		;~ }
		admFile 	:= Journal.GetSelected()
		If (!IsObject(admFile) && !admFile) {
			PraxTT(	"Es wurde keine Datei ausgewählt", "1 0")
			return
		}

		jfile := !IsObject(admFile) ? [admFile] : admFile.Clone()
		rowSel  	:= Journal.getNext(1, admFile)
		For key, inprogress in Addendum.iWin.FilesStack
			If (inprogress = admFile) {
				blockthis := true
				break
			}

		; Callback nicht mehr ausführen wenn ein Menupunkt ausgeführt werden soll
			Addendum.PopUpMenuCallback := ""

	; Menupunkte ohne PDF Schreibzugriffe auf PDF Datei
		If      	InStr(MenuName, "JRefresh")        	{	; Listview aktualisieren
			admGui_Reload()
			return
		}
		else if	InStr(MenuName, "JOpen")         	{	; Karteikarte anzeigen
			If blockthis
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
			else if FuzzyKarteikarte(admFile)
				Addendum.iWin.firstTab := "Patient"
			return
		}
		else if	InStr(MenuName, "JView")          	{	; PDF mit dem Standard PDF Anzeigeprogramm öffnen
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_View(admFile, MenuName)
			return
		}
		else if	InStr(MenuName, "JTgram")          	{	; PDF mit dem Standard PDF Anzeigeprogramm öffnen
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			RegExMatch(MenuName, "Oi)JTgram_(?<name>.*)$", chat)
			RegExMatch(admFile, "O)[\d+\.]+\s+(?<Nr>\d+)\s+(?<Sender>.*?)\.pdf", Fax)
			If Fax.Sender
				Notification := "📠 Fax von <u><b>" Fax.Sender "</b></u>\n📇 Nummer: " Fax.NR
			else
				Notification := "anbei ein Dokument zur Bearbeitung..."
			Telegram_SendDocument(Addendum.BefundOrdner "\" admFile, chat.name, Notification)
		}

	; zurück wenn kein Schreibzugriff möglich ist
		If !IsObject(admFile) && (FileIsLocked(Addendum.BefundOrdner "\" admFile) || blockthis) {
			PraxTT(	"Die Dateioperation ist nicht möglich,`nda ein anderes Programm den Dateizugriff sperrt.", "5 2")
			return
		}

	; Menupunkte mit Schreibzugriff auf PDF Datei
		If      	InStr(MenuName, "JDelete")        	{	; Datei(en) löschen

		; Nutzer fragen
			msg := IsObject(admFile) ? "Wollen Sie " admFile.Count() " Dateien entfernen?" : "Wollen Sie die Datei:`n" admFile "`nwirklich entfernen?"
			MsgBox, 0x1024, Addendum für AlbisOnWindows, % msg
			IfMsgBox, No
				return

		; erstellt Backups und löscht die Original und zugehörigen Dateien
			admFiles := IsObject(admFile) ? admFile : [admFile]
			For fileNr, admFile in admFiles
				res := Befunde.FileDelete(admFile)                         	;~ res := Journal.RemoveFromAll(admFile)

		}
		else if	InStr(MenuName, "JOCR")          	{	; OCR einer Datei

			; Abbruch wenn gerade ein OCR Vorgang läuft
				If Addendum.Thread["tessOCR"].ahkReady() || blockthis {
					PraxTT((blockthis ? "Diese Datei ist gerade in Bearbeitung." : "Ein Texterkennungsvorgang läuft derzeit noch") "`n...bitte warten...", "2 0")
					return
				}

			; Nutzerhinweise anzeigen
				If !PDFisSearchable(Addendum.BefundOrdner "\" admFile) {
					PraxTT("PDF [" admFile "] bisher ohne Texterkennung. Starte Tesseract..", "2 0")
					Sleep 2000
				}
				else                                                                                 	{
					MsgBox, 4	, Addendum für Albis on Windows
									, %	">> OCR Text vorhanden <<`n" admFile "`n"
										. 	"Bei erneuter Ausführung werden die bisherigen Textdaten gelöscht.`n"
										.	"Dennoch ausführen?"
					IfMsgBox, No
						return
					PraxTT("Die Texterkennung der Datei < " admFile " > wird jetzt ausgeführt.", "2 0")
					Sleep 2000
				}

			; OCR Vorgang anzeigen
				Addendum.iWin.FilesStack.InsertAt(1, admFile)
				journal.OCR("+OCR abbrechen")

			; OCR Starten als echter Thread, nach Abschluß eines Vorgangs schickt der Thread eine Nachricht ans Addendum-Skript
				TesseractOCR(admfile)

			return
		}
		else if	InStr(MenuName, "JOCRAll")      	{	; OCR aller offenen Dateien
			If Addendum.Thread["tessOCR"].ahkReady() || (Addendum.iWin.FilesStack.Count() > 0) {
				PraxTT("Ein Texterkennungsvorgang läuft im Moment noch.`n...bitte warten...", "3 0")
				return
			}
			admGui_OCRAllFiles()
			return
		}
		else if	(MenuName ~= "JRecog(1|3)") 	{	; Patienten zuordnen, Dokumentdatum finden
			If blockthis {
				PraxTT("Diese Datei ist gerade in Bearbeitung.`n...bitte warten...", "1 0")
				return
			}
			admCM_JRecog(admFile, MenuName ~= "1" ? 0 : 1, "CM")
			return
		}
		else if	InStr(MenuName, "JRecog2")        	{	; alle Dateien kategorisieren
			If Addendum.Thread["tessOCR"].ahkReady() || (Addendum.iWin.FilesStack.Count() > 0) {
				PraxTT("Ein Texterkennungsvorgang läuft im Moment noch.`n...bitte abwarten...", "1 0")
				return
			}
			admCM_JRecogAll()
			return
		}
		else if	InStr(MenuName, "JRename")     	{	; manuelles Umbenennen
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile)
			return
		}
		else if	InStr(MenuName, "JSplit")           	{	; manuelles Aufteilen einer PDF Datei
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile	,	"Aufteilen der PDF-Seiten auf mehrere Datein.`n"
													.	"Schreiben Sie: 1-2 D1, 3-4 D2`n"
													.	"Datei 1 erhält die Seiten 1-2 und Datei 2 die Seiten 3-4`n" )
			return
		}
		else if	InStr(MenuName, "JImport")        	{	; Datei in geöffnete Karteikarte importieren
			If blockthis
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
			else
				admGui_ImportFromJournal(admFile)
			return
		}
		else if	InStr(MenuName, "JExport")        	{	; Export der/von Datei/en
			If blockthis {
				PraxTT("Die wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Export(admFile)
			return
		}

return
}
;}

; -------- Autonaming                                                                      	;{
admCM_JRecog(admFile, dbg:=false, callstate:="")          	{              	; Kontextmenu: JRecog - Dateibezeichnung aus PDF Text erstellen

		; ** ermittelt im Dokumenttext enthaltene Namen von Patienten.  **
		; ** macht dies bei einer Datei                                                	**

		; dbg     	=	true/false - Debugging Option
		; calstate 	=	CM - Funktion wurde über das  Kontextmenu aufgerufen.
		;							jeder andere String oder ein leerer Parameter führt die Funktion im stillen Modus aus
		;							(zukünftig werden per Multithreading mehrere Dateien gleichzeitig verarbeitet)

		; letzte Änderung:  17.12.2022


		; Pfade
		pdfPath	:= Addendum.Befundordner "\" admFile
		txtPath	:= Addendum.Befundordner "\Text\"  StrReplace(admFile, ".pdf", ".txt")
		F      	:= xstring.get.AllSubStrings(admFile)
		saveTxt	:= false

		; zurück wenn PDF keinen Textlayer hat
		If !PDFisSearchable(pdfPath) {
			PraxTT("PDF Datei enthält keinen Textlayer.`nFühren Sie zuerst eine Texterkennung durch!", "2 1")
			return
		}

		; PDFText auslesen
		PdfText := IFilter(pdfPath)                                     ; PDF to Text per IFilter Funktion ist am schnellsten
		If (StrLen(Trim(PdfText)) = 0) {
			PraxTT("Der extrahierte Text enthält keine Zeichen.`n[" admFile "]", "2 1")
			return
		}

		; extrahierten Text speichern
		FileGetSize, fSize, % txtPath
		If (!FileExist(txtPath) || fSize=0) {
			FileOpen(txtPath, "w", "UTF-8").Write(PdfText)
			FileGetSize, fSize, % txtPath
			If (!FileExist(txtPath) || fSize=0)
				PraxTT("Der PDF-Text konnte nicht gespeichert werden!`n[" admFile "]", "2 1")
		}

		; ImportGui: Anzeigen von Wort und Zeichenzahl der Datei
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (callstate ~= "i)(CM1|CMAll|autocall)") {    ; && !WinExist("admImportLayer ahk_class AutoHotkeyGUI")
			newmsg := "[Textanalyse:`t"     	admFile "]"                                                	"`r`n"
						. 	"Wörter:        `t"       	StrSplit(PdfText, " ").MaxIndex() 						"`r`n"
						. 	"Zeichenzahl:`t"       	StrLen(RegExReplace(PdfText, "[\v\n\r\f]"))   	"`r`n`r`n"
						. 	"....bitte warten`n#"  ; Header endet erst hier
						. 	"Schritt 1/4:`tSuche nach Personennamen`r`n▏"
			admGui_ImportGui("Show", newmsg, "admJournal")
		}

	/*

		▏▏
		▍
		▌
		▋
		▋▏
		▊▏
		▊▏▏
		▊▏▏▏
		▉▎▏
		▉▍▏
		▉▌▏
		▉▋▏
		▉▊▏
		▉▉▏

	*/

		; RegEx Auswertungen
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Namen von Patienten finden
			;~ If ((callstate="autocall" && !xstring.ContainsName(admFile)) || callstate != "autocall")
			If !xstring.ContainsName(admFile)
				fDoc  	:= FindDocNames(PdfText, dbg)  ; "admGui_ImportGui"

		; Dokument- oder Behandlungsdatum ermitteln
			If (callstate ~= "i)(CM1||CMAll|autocall)")
				admGui_ImportGui("update message", "Schritt 2/4: Ermittlung des Dokumentdatum", "admJournal")
			;~ fDays 	:= FindDocDate(fDoc.Text, fDoc.names, dbg)
			fDays 	:= FindDocDate(Text, fDoc.names, dbg)

		; neuen Dateinamen erstellen
		; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (IsObject(fDoc.names) && fDoc.names.Count() > 0) {

			; der Name des Patienten als erstes
			if (fDoc.names.Count() = 1) {

				For PatID, Pat in fDoc.names {
					newfilename := fDoc.names[PatID].Nn ", " fDoc.names[PatID].Vn ", " (F.K ? F.K : "") (F.I ? " - " F.I : "") " "
					break
				}

				If dbg
					SciTEOutput("   " A_ThisFunc " 1: " newfilename)

				; als letztes das gefundene Dokumentdatum
				; es wurde ein von bis Datum extrahiert, dieses Datum wird jetzt noch gekürzt um möglichst wenig Zeichen zu verbrauchen
				If !IsObject(fDays) && InStr(fDays, "-") {
					docDate := ShrinkTwoDatesStr(fDays)
					newfilename .= docDate
					If dbg
						SciTEOutput("   " A_ThisFunc " 2a: " fDays ", " docDate)
				}
				else If (IsObject(fDays)) {    ; && !F.D1 && !F.D2
					docDate     	:= fDays.Behandlung[1] ? fDays.Behandlung[1] : fDays.Dokument[1]
					newfilename .= RegExReplace(docDate	, "^[\s,v\.om]+")
					If dbg
						SciTEOutput("   " A_ThisFunc " 2b: " fDays ", " docDate)
				}
				else {
					newfilename .= (F.D ? F.D : "")
				}

				If dbg
					SciTEOutput("   " A_ThisFunc " 3: " newfilename)

				; wird umbenannt
				admGui_Rename(admFile, "", newfilename := (Trim(newfilename) (!(newfilename ~= "\.pdf$") ? ".pdf" : "")), dbg)
			}

			; keine eindeutige Zuordnung möglich
			else if (fDoc.names.Count() > 1)  {

				For PatID, Pat in fDoc.names
					t .= "   [" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"
				PraxTT("Das Dokument konnte nicht eindeutig zugeordnet werden. Folgende Patienten stehen zu Auswahl:`n"  t, "15 1")
			}

		}

		; es wurde kein Patient gefunden
		else {

			PraxTT("Dokument: " admFile "`nkonnte keinem Patienten zugeordnet werden", "4 1")

			 If (IsObject(fDays)) {    ; && !F.D1 && !F.D2
					docDate     	:= fDays.Behandlung[1] ? fDays.Behandlung[1] : fDays.Dokument[1]
					newfilename := RegExReplace(docDate	, "^[\s,v\.om]+")
			} else
				newfilename := GetFileTime(pdfPath)

			newfilename := (Trim(newfilename) (!(newfilename ~= "\.pdf$") ? ".pdf" : ""))

			; wird umbenannt
			If (callstate ~= "i)(CM1||CMAll|autocall)")
				admGui_ImportGui("update message", "Dokument: " admFile " ließ sich keinem Patienten zugeordnen,"
																		. "die Dokumentbezeichnung ist automatisch vergeben worden`n " newfilename  , "admJournal")

			admGui_Rename(admFile, "", newfilename, dbg)

		}


	 ; nur eine Datei soll erkannt werden dann schliessen
		If (callstate = "CM1")
			admGui_ImportGui(false)

return (newfilename ? newfilename : "")
}

admCM_JRecogAll(files:="")                                            	{                	; Kontextmenu: automatisiert die Benennung von PDF-Dateien

	; letzte Änderung 12.07.2022


	; Variablen
		static newadmFile, lastnewadmFile, RecognitionDocs
		global Recog, RecogCancel, RecogCounter
		OCRDone := []

	; ein Parameteraufruf der Funktion hat stattgefunden
		If IsObject(files) && !IsObject(RecognitionDocs) {

			RecognitionDocs := []
			for idx, filename in files
				If Befunde.FileExist(filename)
					RecognitionDocs.Push(filename)

		}

	; Dateinamen unbenannter PDF-Dateien einsammeln
		If !IsObject(RecognitionDocs)	{

			MsgBox, 0x2003, Addendum für AlbisOnWindows, Sollen nur bisher nicht kategorisierte Dokumente verarbeitet werden?
			IfMsgBox, Cancel
			{
				PraxTT("Abbruch durch Nutzer", "1 1")
				return
			}
			IfMsgBox, Yes
				RecognitionDocs := pdfpool.GetUnamedDocuments()
			IfMsgBox, No
				RecognitionDocs := pdfpool.GetAllDocuments()

		}

	; Abbruch wenn keine Dokumente zum Umbenennen gefunden wurden
		If (RecognitionDocs.MaxIndex() = 0 || !IsObject(RecognitionDocs)) {
			PraxTT("Keine Dateien für eine automatische Umbennung vorhanden!", "1 2")
			return
		}

	; Vorbereitungen
		cPrefix := ""
		Loop % (cLen	:= StrLen(RecognitionDocs.Count())-1)
			cPrefix .= "0"

	; kleine Extra-Gui einblenden, welche den laufenden Prozess anzeigt ;{
		If IsObject(RecognitionDocs_X) {     ; ## ausgestellt
			RecogEnd := false
			Gui, Recog: New   	, -Caption +ToolWindow -DPIScale AlwaysOnTop HWNDhRecog
			Gui, Recog: Margin	, 0, 0
			Gui, Recog: Color  	, cEFF0F1
			Gui, Recog: Font   	, % "s10 cWhite"                                                                  	, % Addendum.Default.Font
			Gui, Recog: Add    	, Text	, x2 y2                                                                     	, % "Autonaming Datei: "
			Gui, Recog: Add    	, Text	, x+2 vRecogCounter                                            	, % prefix "1/" RecognitionDocs.Count()
			Gui, Recog: Font   	, % "s8 cWhite"                                                                  	, % Addendum.Default.Font
			Gui, Recog: Add    	, Button, x2 y+5 vRecogCancel gRecogHandler                    	, % "Abbruch"
			Gui, Recog: Show    	, % "AutoSize Hide"                                                            	, % "Addendum Autonaming"

			aw:= GetWindowSpot(AlbisWinID())
			rw	:= GetWindowSpot(admHRecog)
			Gui, Recog: Show     	, % "x" aw.X " y" aw.CH - rw.H " NA"
		}
	;}

	; Patienten und Dokumentdatum finden
		For idx, filename in RecognitionDocs {

			; Abbruch nach Nutzerinteraktion mit ExtraGui
				If RecogEnd
					break

			; Pfade festlegen
				path  	:= Befunde.GetPaths(filename)
				counter	:= SubStr(cPrefix . idx, -1*cLen) "/" RecognitionDocs.Count()
				Befunde.Backup(path.fF)
				PdfText := ""

			; Debug
				;~ If !IsObject(files)
					GuiControl, Recog:, RecogCounter, % counter
				If (idx > 1) && (Weitermachen("Datei: " counter "`nBez.: " filename "`nWeitermachen?",, 2) = 0)
				 break

			; Text extrahieren, falls noch nicht geschehen
				If PDFisSearchable(Path.fF) && !FileExist(path.fT) {

					; Fortschritt anzeigen
						If !IsObject(files)
							admCM_JRnProgress([counter, filename, "Text wird extrahiert....", "---", "---", lastnewadmFile])

					; pdftotext extrahiert den Text aus der PDF Datei
						PdfText := IFilter(Path.fF)
						FileOpen(Path.fT, "w", "UTF-8").Write(PdfText)
						If !FileExist(Path.fT) {
							PraxTT("Datei [" counter "] " filename "`nTextextraktion fehlgeschlagen", "2 1" )
							continue
						}

				}
				else if !PDFisSearchable(Path.fF) {

					; Protokollierung
						PraxTT("Datei [" counter "] " filename " ist nicht durchsuchbar", "2 1" )
						Sleep 500
						continue
				}

			; extrahierten Text lesen
				If !PdfText
					PdfText := IFilter(Path.fF)
				If (StrLen(Trim(PdfText)) = 0) {
					PraxTT("Der extrahierte Text enthält keine Zeichen.`n[" filename "]", "2 1")
					FileDelete, % Path.fT
					continue
				}

			; Fortschritt anzeigen
				If IsObject(RecognitionDocs_X)
					admCM_JRnProgress([ counter
													, filename
													, "Textinhalt wird analysiert...."
													, (words:= StrSplit(PdfText, " ").MaxIndex())
													, (chars	:= StrLen(RegExReplace(PdfText, "[\s\n\r\f\v]")))
													, lastnewadmFile])

			; Patientennamen und Behandlungs- oder Dokumentdatum finden
				fDoc  	:= FindDocNames(PdfText, false)
				fDays 	:= FindDocDate(fDoc.Text, fDoc.names, false)
				fRecom	:= FindDocEmpfehlung(PdfText)

			; Datei umbenennen
				If      	(IsObject(fDoc.names) && fDoc.names.Count() = 1) {

					; Name des Patienten
						newadmFile := ""
						For PatID, Pat in fDoc.names {
							If (StrLen(Trim(Pat.Nn)) > 0 && StrLen(Trim(Pat.Vn)) > 0) { 				; Abbruch wenn keine Name erkannt wurde
								newadmFile := Pat.Nn ", " Pat.Vn ", "
								break
							}
						}
						If (StrLen(newadmFile) = 0)
							continue

					; Behandlungzeitraum oder Datum des Dokuments
						If IsObject(fDays)
							newadmFile .= " " (fDays.Behandlung[1] ? fDays.Behandlung[1]	: fDays.Dokument[1])

					; Anzeige von Wort und Zeichenzahl der Datei
						lastnewadmFile := newadmFile
						If IsObject(files_X)				;If !IsObject(files)
							admCM_JRnProgress([counter, filename, "neuer Dateiname erstellt", words, chars, lastnewadmFile])
						else
							SciTEOutput("  - Autonaming: " newadmFile)

					; Umbenennen von Original, Backup und Textdatei
						If (StrLen(newadmFile) > 0) {
							admGui_FWPause(true, [filename  	, newadmFile], 12)   	; FolderWatch pausieren (6s bis zur Wiederherstellung)
							row	:= journal.Replace(filename		, newadmFile)        	; Journal auffrischen
							pdf	:= pdfpool.Rename(filename		, newadmFile)        	; Objektdaten ändern
							res	:= Befunde.Rename(filename		, newadmFile)        	; Dateien umbenennen (Original, Backup, Text)
							admGui_FWPause(false, [filename  	, newadmFile], 12)   	; FolderWatch fortsetzen
							OCRDone.Push(newadmfile)
							newadmFile := ""
						}

				}
				else if	(IsObject(fDoc.names) && fDoc.names.Count() < 1) {
					If !IsObject(files)
						PraxTT("  - Dokument: " filename "`nkonnte keinem Patienten zugeordnet werden", "1 1")
					else
						SciTEOutput("  - Autonaming: keine eindeutige Zuordnung möglich gewesen")
				}
				else If 	(IsObject(fDoc.names) && fDoc.names.Count() > 1) {

					t := ""
					For PatID, Pat in fDoc.names
						t .= " `t[" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"

					SciTEOutput("  - [" counter "]")
					SciTEOutput("  - 1: " filename)
					SciTEOutput("  - 2: #keine eindeutige Zuordnung möglich")
					SciTEOutput("  - Namen:       `t" 	fDoc.names.Count())
					SciTEOutput("  - Beh. Datum:`t"  	fDays.Behandlung.Count())
					SciTEOutput("  - Doc. Datum:`t" 	fDays.Dokument.Count())
					SciTEOutput(t)
					SciTEOutput("  -------------------------------------------------------------------" )

				}
				else {
					If !IsObject(RecognitionDocs)
						PraxTT("Dokument: " filename "`nkonnte keinem Patienten zugeordnet werden", "1 1")
					else
						SciTEOutput("  - Autonaming: keine eindeutige Zuordnung möglich gewesen")
				}

			; Abbruch nach Nutzerinteraktion mit ExtraGui
				If RecogEnd
					break

		}

	; Anzeige schliessen
		RecognitionDocs := ""
		;~ admGui_ImportGui(false)
		;~ Gui, Recog: Destroy
		;~ If IsObject(RecognitionDocs) {
			;~ admGui_ImportGui(false, "A U T O N A M I N G`n" filename "`nläuft", "admJournal")
		;~ }

return OCRDone

RecogHandler: ;{

	If (A_GuiControl = "RecogCancel")
		RecogEnd := true

return ;}
}

admCM_JRnProgress(ProgressText)                                 	{              	; zeigt den Fortschritt des Umbennens an

		static outmsgTitle := "Dokumentklassifizierung ..#"
		static outmsg_last
		static call:=1

	; nur eine Zahl übergeben dann wird die Progressbar verändert
		If IsObject(ProgressText) {

			outmsg := StrReplace(outmsgTitle, "#", SubStr("................", 1, call)) "`n"
						.	 "- - - - - - - - - - - - - - - - - - - - - - - - - "	"`n"
						.	 "Zähler:   `t" 	ProgressText.1             	"`n"
						.	 "Datei:    `t" 	ProgressText.2             	"`n"
						. 	 "Status:   `t"  	ProgressText.3            	"`n"
						. 	 "Wörter:  `t" 	ProgressText.4	          	"`n"
						. 	 "Zeichen:`t"  	ProgressText.5           	"`n"
						. 	 "letzter:   `t"   	ProgressText.6

			admGui_ImportGui(outmsg_last ? true : "update header & message", outmsg, "admJournal")
			call := call = 5 ? 1 : call+1
			outmsg_last := outmsg

		}

}
;}

; -------- Zusatzfunktionen                                                               	;{
admGui_TPSuche()                                                                	{        	; wann war ein Patient da - nicht beendet

	TPFullPath	:= Addendum.DBPath "\Tagesprotokolle"
	Gui, adm: Submit, NoHide

	If (StrLen(admTPSuche) = 0)
		return

	;~ Loop, Files, *.txt, R
	;~ {

	;~ }

}

admGui_LaborblattDruck(Columns, Printer, PrintAnnotation)	  	{        	; automatisiert den Laborblattdruck

		; automatisiert den Laborblattdruck und stellt die vorherige Ansicht wieder her
		; admGui_LaborblattDruck(admLBSpalten, admLBDrucker, admLBAnm)

		static AlbisView

		AlbisView 	:= AlbisGetActiveWindowType(true)
		savePath 	:= InStr(Printer, "Microsoft Print to PDF") ? Addendum.ExportOrdner "\Laborwerte von " RegExReplace(AlbisCurrentPatient(), "[\s]") ".pdf" : ""
		If (Columns = "Alles") {
			MsgBox, 0x1004, % "Sie sind gerade im Begriff alle Laborwerte ausdrucken zu wollen.`nSind Sie sich sicher?"
			IfMsgBox, No
				return
		}

		resPrint := AlbisLaborblattExport(Columns, savePath, Printer, PrintAnnotation)
		If RegExMatch(resPrint, "LBEx\d+")
			PraxTT("Der Ausdruck von Laborwerten ist fehlgeschlagen.`nFehlercode: " resPrint, "4 1")

		If (AlbisView <> AlbisGetActiveWindowType(true))
			AlbisKarteikartenAnsicht(AlbisView)

return savePath
}
;}

; -------- *Dokument- und Listviewklassen                                                  	;{

class Journal                                                                	{	; Listview-Handler des Journal und Patient Tab

	; Funktion:         	verwaltet das Listview und andere Steuerelemente des Journal und teilweise des Patient TAB
	;
	; Anmerkung:		Die Listview des Journal ist als Standard festgelegt. Bevor mit der Patient-TAB Listview gearbeitet werden kann,
	; 	                      	muss diese als Standard in den Listvieweinstellungen über this.Default("admReports") gesetzt werden.
	;
	; Abhängigkeiten: Addendum_DB.ahk, Addendum_Albis.ahk, Addendum_Internal.ahk
	;
	; letzte Änderung: 06.10.2022

	static lastFilter := "", initFilter := true

	Add(fname)                                	{        	; Dokument hinzufügen oder Anzeigedaten ändern

		If !RegExMatch(fname, "i)\.(pdf|doc|docx|jpg|png|tiff|bmp|wav|mov|avi|mp\d)$")
			return

		; Listview auswählen, Neuzeichnen verhindern
		this.Default("admJournal")
		this.RedrawLV(false)

		; Listviewzeile der Datei/Dokument
		row := this.GetRow(fname)

		; nur PDF Dateien
		If RegExMatch(fname, "i)\.pdf$") {

			If !(docID := pdfpool.inPool(fname))
				return  ; ### im Moment noch return

			pdf := ScanPool[docID]
			If !row                 ; Dokument wird  noch nicht angezeigt
				LV_Add("ICON" (pdf.isSearchable ? 3:1), pdf.name, pdf.pages, pdf.filetime, pdf.timestamp)
			else
				LV_Modify(row, "ICON" (pdf.isSearchable ? 3:1), pdf.name, pdf.pages, pdf.filetime, pdf.timestamp)

		}
		; alle anderen Dateitypen
		else {
			file := pdfpool.FileTime(Addendum.BefundOrdner "/" fname)
			If !row
				LV_Add("ICON2", fname,, file.filetime, file.TimeStamp)
			else
				LV_Modify(row, "ICON2", fname,, file.filetime, file.TimeStamp)
		}

		; diese Listview neu zeichnen lassen
		this.RedrawLV(true)

		; Dokument auch in der Listview des Patient-TAB anzeigen
		this.ReportAdd(fname)

		; Infotexte aktualisieren
		this.InfoText()

	}

	AddDocumentsAll()                         	{

		this.Default("admJournal")
		journal.Empty()                                    	; Listview leeren
		this.RedrawLV(false)                            	; Neuzeichnen aus

		For docID, pdf in ScanPool	{          	;
			LV_Add("ICON" (pdf.isSearchable ? 3 : 1), pdf.name, pdf.pages, pdf.filetime, pdf.timeStamp)	; anderes Symbol für OCR-PDF Dateien
			this.InfoText()
			If !Addendum.iWin.OCREnabled && !pdf.isSearchable
				journal.OCR("+OCR " (Addendum.Thread["tessOCR"].ahkReady() ? "abbrechen" : "ausführen"))
		}

		this.RedrawLV(true)
		this.InfoText()                                      	; Infotexte aktualisieren

	}

	AddPictures()                              	{       	; Bild-Dokumente hinzufügen

		; Listview auswählen, Neuzeichnen verhindern
		this.Default("admJournal")
		this.RedrawLV(false)								; Neuzeichnen an

		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(Trim(A_LoopFileLongPath), "\.(jpg|png|tiff|bmp|wav|mov|avi)$") {
				FileGetTime, timeStamp, % A_LoopFileLongPath, C
				FormatTime, filetime  	, % timeStamp, dd.MM.yyyy
				FormatTime, timeStamp	, % timeStamp, yyyyMMdd
				SplitPath, A_LoopFileLongPath, BildBefund
				LV_Add("ICON2", BildBefund,, filetime, timeStamp)
			}

		this.RedrawLV(true)
	}

	Arrange(LVWidth)                          	{        	; richtet die Spaltenbreiten und Datenformate ein

		Col2W := 25
		Col3W := 58
		Col4W := 0
		Col1W := LVWidth - Col2W - Col3W - 25
		this.Default("admJournal")
		LV_ModifyCol(1, Col1W " Left NoSort")
		LV_ModifyCol(2, Col2W " Left Integer")
		LV_ModifyCol(3, Col3W " Right NoSort")
		LV_ModifyCol(4, Col4W " Left Integer")      	; versteckte Spalte enthält Integerzeitwerte für die Sortierung nach dem Datum

	}

	Default(LVName="")                        	{        	; Gui und Listview zum Standard machen

		global admJournal, adm, admHTab, admHJournal, admReports
		static tabs	:= {"admJournal":"Journal"        	, "admReports":"Patient"}
		static hLV	:= {"admJournal":"admHJournal"	, "admReports":"admHReports"}

		LVName := !this.DefaultLV && !LVName ? "admJournal" : LVName
		If (LVName && LVName <> this.DefaultLV) {
			hwnd            	:= hLV[LVName]
			this.DefaultLV 	:= LVName
			this.DefaultTab 	:= tabs[LVName]
			this.DefaultHLV	:= %hwnd%
		}

		Gui, adm: Default
		Gui, adm: ListView, % this.DefaultLV

	}

	Focus()                                    	{        	; Eingabefokus auf das Journallistview setzen
		global admJournal, adm, admHTab, admHJournal
		this.Default()
		GuiControl, adm: Focus, % this.DefaultLV
	return ErrorLevel
	}

	DeleteRow(row="")                         	{       	; Zeile(n) im Journal entfernen

		; row kann die Nummer der zu löschenden Zeile oder die Bezeichnung des Dokumentes sein
		; Rückgabewert ist die Nummer der entfernten Zeile

		this.Default()

		; leerer Parameter oder row größer als Anzahl der Listviewzeilen oder null
		If !row || (RegExMatch(row, "^\d+$") && (row > this.GetCount() || row = 0))
			return 0
		; row enthält nicht nur Zahlen, Zeile läßt sich nicht finden
		else If !RegExMatch(row, "^\d+$")
			If !(row := this.GetRow(Trim(row)))
				return 0

		; Zeile entfernen
		LV_Delete(row)
		EL := ErrorLevel

		; Infotexte aktualisieren
		this.InfoText()

	return EL = 1 ? 0 : row
	}

	Empty()                                    	{       	; Listview komplett leeren
		this.Default()
		LV_Delete()
	}

	FindFileRow(filename)                     	{        	; anhand Dateinamen finden

		this.Default("admJournal")
		this.RedrawLV(false)

		row := 0
		Loop % LV_GetCount() {
			LV_GetText(rowFilename, A_Index, 1)
			If (rowFilename = filename) {
				this.ListView()
				LV_Modify(row := A_Index, "Vis")
				break
			}
		}

		this.RedrawLV(true)

	return row		; 0 = nichts ersetzt
	}

	GetCount(LV="admJournal")                 	{        	; Anzahl der Tabellenzeilen
		this.Default(LV)
	return LV_GetCount()
	}

	GetFilter()                                	{        	; ausgewählten Filter erhalten

		; wird bei eingestellter AltSubmit Eigenschaft die ausgewählte Zeile sein. Ansonsten wird der Zeilentext zurückgegeben.
		GuiControlGet, JrnFilter, adm:, admJrnFilter

	return JrnFilter
	}

	GetLastFilter()                            	{        	; ausgewählten Filter erhalten
	return this.lastFilter
	}

	GetNext(rLV:=0, LVNI:=2)                  	{        	; findet die nächste ausgewählte/fokussierte Zeile
		global admJournal, adm, admHTab, admHJournal, admHReports
		; rLV = 1 based index of the starting row for the flag search. Omit or 0 to find first occurance specified flags.
		; LVNI_ALL := 0x0, LVNI_FOCUSED := 0x1, LVNI_SELECTED := 0x2
		hLV := this.DefaultHLV
	return DllCall("SendMessage", "uint", hLV, "uint", 4108, "uint", rLV-1<0?0:rLV-1 , "uint", LVNI) + 1
	}

	GetRow(fname, col:=1, DLV:="")            	{        	; Zeilennummer mit diesem Dateinamen

		row := 0
		If !DLV
			this.Default()
		else
			Gui, adm: ListView, % DLV
		Loop % LV_GetCount() {
			LV_GetText(rText, A_Index, col)
			If (rText = fname) {
				row := A_Index
				break
			}
		}

	return row
	}

	GetSelected(col:=1)                       	{        	; alle ausgewählten Zeilen auslesen

		row 		:= 0
		JFiles	:= Array()

		this.Default("admJournal")
		while (row := LV_GetNext(row)) {
			LV_GetText(fname, row, col)
			JFiles.Push(fname)
		}

		If (JFiles.Count()=1)
			return JFiles[1]
		else if (JFiles.Count()>1)
			return JFiles

	}

	GetText(row, col:=1)                      	{        	; Text einer Zeile, Spalte auslesen
		this.Default()
		LV_GetText(rText, row, col)
	return rText
	}

	GetTopIndex()                            		{     		; Index des obersten sichtbaren Elementes
			; LVM_GETTTOPINDEX
		SendMessage, 0x1027,,,, % "ahk_id " this.DefaultHLV
	return ErrorLevel
	}

	gLabel(register)                           	{       	; g-Label bereitstellen oder abschalten

		global adm, admJournal, hadm
		GuiControl, % "adm: " (register ? "+gadm_Journal" : "-g"), % this.DefaultLV

	}

	InfoText()                                 	{        	; Infotext (Journal/Patient) wird aufgefrischt

		global admButton1, adm, hadm, admJournalTitle, admPatientTitle, fn_AutoPatientTab

		Gui, adm: Default

		; Zahl benannter Dokumente
		fullnamed := pdfpool.GetFullNamedDocuments()

		; Journal
		InfoText := (ScanPool.Count() = 0) ? " keine Dokumente" : (ScanPool.Count() = 1) ? "1 Dokument" : ScanPool.Count() " Dok. (" fullnamed.Count() " vollst.)"
		GuiControl, adm:, admJournalTitle, % InfoText

		; Patient
		InfoText := "Posteingang: (" (!PatDocs.Count() ? "keine neuen Befunde)" : PatDocs.Count() = 1 ? "1 Befund)" : PatDocs.Count() " Befunde)") ", insgesamt: " ScanPool.Count()
		GuiControl, adm:, admPatientTitle, % InfoText

		; "Auto(SwitchTo)PatientTab" falls für den aktuellen Patienten neue Dokumente vorhanden sind
		If (PatDocs.Count() > 0) && !Addendum.Flags.RNRunning {
			fn_AutoPatientTab := Func("admGui_ShowTab").Bind("Patient")
			SetTimer, % fn_AutoPatientTab, -30000
		}
		; "Auto(SwitchTo)PatientTab" Timer beenden wenn PatDocs leer ist
		else if (PatDocs.Count() = 0 && IsObject(fn_AutoPatientTab)) {
			SetTimer, % fn_AutoPatientTab, Delete
			fn_AutoPatientTab := ""
		}


	}

	ListView(LVName:="")                      	{        	; nur dieses Listview zum Standard machen
		global admJournal, adm, admHTab, admHJournal, admReports
		Gui, adm: ListView, % (LVName ? LVName : "admJournal")
	}

	SelectRow(row, unselect:=true)            	{       	; eine Zeile auswählen

		; ♀ 📧
		; row - Index der Zeile oder ein String für den Vergleich mit dem Inhalt (Tabelle wird nach dem ersten Vorkommen durchsucht)

			global adm, admJournal, admReports, hadm
			static 	LVIS_FOCUSED:=1, LVIS_SELECTED:=2

		; Deregistriere die gosub-Label Verknüpfung vor dem Auswählen einer Zeile (kein versehentliches Auslösen von Ereignissen)
			this.gLabel(false)
			this.Default("admJournal")

		; es wurde ein Textstring anstatt einer Reihennummer übergeben
			row :=  !RegExMatch(row, "^\d+$") ? this.GetRow(row) : row

		; Listviewzeilenauswahl wird für alle Zeilen entfernt
			If unselect
				LV_Modify(0, "-Select -Focus")

		; Zeile auswählen, alle anderen werden abgewählt
			this.Focus()
			LV_Modify(row, "Select1 Focus Vis")

		; sicherstellen das die ausgewählte Zeile im sichtbaren Bereich ist
			this.EnsureVisible(row)

		; Registriere das g-Label wieder
			this.gLabel(true)

	}

	RedrawLV(rdw)                             	{        	; Listview neuzeichnen an oder aus
		global admJournal, adm, admHTab, admHJournal, admHReports,hadm
		Gui, adm: Default
		Gui, adm: ListView, % (this.DefaultLV ? this.DefaultLV : "admJournal")
		GuiControl, % "adm: " (rdw ? "+" : "-")  "Redraw", % (this.DefaultLV ? this.DefaultLV : "admJournal")
		EL := ErrorLevel
		If rdw
			RedrawWindow(admHJournal), RedrawWindow(admHReports)
	return EL
	}

	Remove(fname:="", rLV:="")                  {        	; entfernt eine Datei aus der Listview

		; ein Dateiname oder eine Zeilennummer können übergeben werden
		global admJournal, adm, admHTab, admHJournal

		this.Default("admJournal")
		this.RedrawLV(false)
		row := fname ? this.GetRow(fname) : rLV ? rLV : 0
		res := row > 0 ? LV_Delete(row) : -1
		this.RedrawLV(true)
		this.InfoText()

	return res
	}

	RemoveFromAll(fname)                      	{	      	; entfernt die Datei aus allen Objekten und Listviews

		; entfernt die Datei aus dem Befund- und PDFPool-Objekt und aus beiden Listviews (Journal und Patient)

		res := Befunde.Remove(fname)	    	; verschiebt die Datei in das Backupverzeichnis (löscht nicht!)
		res := PDFpool.Remove(fname)       	; entfernt sie aus dem Scanpool Objekt
		res := this.Remove(fname)              	; entfernt ihre sichtbare Anzeige im Journal Listview
		res := this.ReportRemove(fname)     	; entfernt ihre sichtbare Anzeige im Patient Listview

		this.InfoText()

	}

	Replace(sourcefile, targetfile)           	{        	; Dateinamen ändern

		this.Default()
		this.RedrawLV(false)

		row := 0
		Loop % LV_GetCount() {
			row := A_Index
			LV_GetText(rowFilename, row, 1)
			If (rowFilename = sourcefile) {
				this.ListView()
				LV_Modify(row, "Vis", targetfile)
				break
			}
		}

		this.RedrawLV(true)

	return row		; 0 = nichts ersetzt
	}

	SetFilter(JrnFilter)                      	{        	; angezeigten Filter des Steuerelementes ändern

		If !JrnFilter || !(JrnFilter~="^\s*(\d+|[\pL\s\-])\s*$")
			return
		GuiControl, % "adm: " (JrnFilter ~= "^\d+$" ? "Choose" : "ChooseString"), admJrnFilter, % JrnFilter
		this.UseFilter(JrnFilter)

	}

	Show()                                     	{        	; zeigt den Journal Tab an
		global admHTab
	return admGui_ShowTab(this.DefaultTab)
	}

	ShowImportStatus(ImportRuns)               	{        	; Interaktionskontrolle der 'Importieren' Steuerelemente

		global admJImport, admButton1, adm, hadm, PatDocs

		; Patient Tab, Import Button
			this.Default("admReports")
			GuiControl, adm:, admButton1, % (ImportRuns ? "...importiere" : "Importieren")
			If ImportRuns
				GuiControl, % "adm: Disable", admButton1
			else
				GuiControl, % "adm: " (PatDocs.MaxIndex() > 0 ? "Enable" : "Disable"), admButton1

		; Journal Tab, Import Button
			this.Default("admJournal")
			GuiControl, adm:, admJImport, % (ImportRuns ? "...importiere" : "Importieren")
			If ImportRuns
				GuiControl, % "adm: Disable", admJImport
			else
				GuiControl, % "adm: " (this.GetCount()>0 ? "Enable" : "Disable"), admJImport

	}

	ShowMenu(admFileObj)                      	{        	; Kontextmenu anzeigen

		global admJCM, admJCMX

		MouseGetPos, mx, my, mWin, mCtrl, 2
		Menu, % (IsObject(admFileObj) ? "admJCMX" : "admJCM")	, Show, % mx - 20, % my + 10

	return
	}

	SortByDocNames(desc=false)                	{        	; sortiert nach Dokumentnamen
		this.Default()
		LV_ModifyCol(1, (!desc ? "Sort" : "SortDesc"))
	}

	SortByDocDates(desc=false)                	{        	; sortiert nach Dokumentdatum
		this.Default()
		LV_ModifyCol(3, (!desc ? "Sort" : "SortDesc"))
	}

	Sort(EventNr, LV_Init=false)               	{        	; sortiert die Journalspalten, zeigt Symbol für Sortierreihenfolge

		; Funktion wird gebraucht für die Wiederherstellung der letzten Sortierung und für das Sichern der Einstellungen bei Nutzerinteraktion
		; LV_Init - nutzen um gespeicherte Sortierungseinstellung wiederherzustellen

		global 	admHJournal, hadm
		static 	admJCols, LVSortStr, JSort, JColDir, JRow, JSortDir := []

		journal.Default("admJournal")

		If LV_Init || (EventNr = 0) {

				; LVM_GETHEADER = LVM_FIRST (0x1000) + 31 = 0x101F.
			admhHdr	:= DllCall("SendMessage", "Uint", admHJournal	, "Uint", 0x101F, "Uint", 0, "Uint", 0)
			; HDM_GETITEMCOUNT = HDM_FIRST (0x1200) + 0 = 0x1200.
			admJCols	:= DllCall("SendMessage", "Uint", admhHdr     	, "Uint", 0x1200, "Uint", 0, "Uint", 0)

			Loop % admJCols
				JSortDir[A_Index] := "0"

			; Spalte "3" , Sortierrichtung 1 = Aufsteigend : 0 = Absteigend, Listview scrollen zur Reihe Nr.
			Addendum.iWin.JournalSort := IniReadExt(compname, "Infofenster_JournalSortierung", "3 1 1")
			RegExMatch(Addendum.iWin.JournalSort, "\s*(\d+)\s+(\d)\s+(\d+)", S)
			EventNr            	:= S1= 0 	? 1 : S1
			JSortDir[EventNr] 	:= S2= ""	? 1 : S2
			TopIndex             	:= S3

		}

		; Sortierung je nach gewählter Spalte vornehmen
		; Idee von: https://www.autohotkey.com/boards/viewtopic.php?t=68777
		; Spalte 3 (Eingangsdatum) - sortiert wird nach der unsichtbaren Spalte 4 (Zeitstempel)
			EventNr := EventNr = 3 ? 4 : EventNr

			this.ColSort("admJournal", EventNr, JSortDir[EventNr] := !JSortDir[EventNr])
			TopIndex := journal.GetTopIndex()
			Addendum.iWin.JournalSort := EventNr " " JSortDir[EventNr] " " TopIndex

	}

	ColSort(LVName, col, SortDest)            	{       	; Helfer von .Sort()

		global 	admHJournal

		If (LVName ~= "i)(admJournal|admReport)")
			journal.Default(LVName)
		else
			admGui_Default(LVName)

		If (LVName = "admJournal") {
			Type	:= col=2 ? "Integer" : col=3 ? "Integer" : "Text"
			col 	:= col=3 ? 4 : col
		}

		LV_ModifyCol(col, (SortDest ? "Sort" : "SortDesc") " " Type)
		LV_SortArrow(admHJournal, col, (SortDest ? "up":"down"))

	}

	UseFilter(JrnFilter)                       	{	      	; Anwenden des ausgewählten Dokumentfilters

		If !JrnFilter
			return

		Switch JrnFilter {

			case "1" :
				Docs := pdfpool.GetAllDocuments()
			case "2" :
				Docs := pdfpool.GetAllDocuments()
			case "3" :
				Docs := pdfpool.GetUnamedDocuments()
			case "4" :
				Docs := pdfpool.GetNamedDocuments()

		}

		this.Default("admJournal")


	}

	EnsureVisible(row)                        	{        	; sicherstellen das die Listview Reihe sichtbar ist

		; versucht die übergebene Zeile in der Mitte der Listview zu halten

		; Default machen wenn nicht Default
			If (this.DefaultLV <> "admJournal")
				this.Default("admJournal")

		; Anzahl der sichtbaren Zeilen der Listview bekannt?
			If !IsObject(this.visrows[this.DefaultLV])
				this.VisibleRows()

		; max. Zeilenabstand berechnen
			If !this.visrows[this.DefaultLV].visplus && this.visrows[this.DefaultLV].total
				this.visrows[this.DefaultLV].visplus := Floor(this.visrows[this.DefaultLV].total/2)

		; EnsureVisible ausführen
			ensureVisRow := row+this.visrows[this.DefaultLV].visplus > this.GetCount() ? this.GetCount() : row+this.visrows[this.DefaultLV].visplus
			SendMessage, 0x1013, % ensureVisRow - 1,,, % "ahk_id " this.DefaultHLV

	}

	VisibleRows()                              	{        	; gibt die Anzahl der sichtbaren Reihen zurück

		VisibleRange := {}

		this.Default("admJournal")
		If !IsObject(this.visrows[this.DefaultLV])
			this.visrows[this.DefaultLV] := Object()

		Loop % this.GetCount() {

			SendMessage, 0x10B6, % A_Index-1,,, % "ahk_id " this.DefaultHLV 	; LVM_ISITEMVISIBLE -> findet das erste sichtbares Item
			If (!VisibleRange.Count() && ErrorLevel) {
				this.visrows[this.DefaultLV].first 	:= VisibleRange.firstrow := A_Index
			}
			else if (VisibleRange.Count() && !ErrorLevel) {
				this.visrows[this.DefaultLV].last 	:= VisibleRange.lastrow := A_Index
				this.visrows[this.DefaultLV].total 	:= VisibleRange.visrows := VisibleRange.lastrow-VisibleRange.firstrow
				visrows := this.visrows[this.DefaultLV]
				;~ SciTEOutput("firstvisible item:" visrows.first "`n lastvisible item: " visrows.last "`nvisible rows: " visrows.total)
				return VisibleRange
			}

		}

	}

; ------------------------- Journal andere Steuerelemente --------------------------------------------------------
	OCR(status)                               	{	      	; OCR Button schalten und flag setzen

		global admJOCR, adm

		RegExMatch(status " ", "i)^\s*(?<Enabled>[\+\-])\s*(?<Text>[\pL\-\s]+)", OCR_)
		OCRState := "Enable" (OCR_Enabled="+" ? "1" : OCR_Enabled="-" ? "0" : "1")

		Addendum.iWin.OCREnabled := OCR_Enabled="+" ? true : false

		this.Default("admJournal")
		GuiControl, % "adm: " 	OCRState 	, admJOCR
		GuiControl, % "adm: ",	admJOCR	, % OCR_Text

	}


; ------------------------- PATIENT TAB -----------------------------------------------------------------------------
;   PatDocs ist hier globales Objekt

	ReportAdd(fname)                     	{        	; Dokument in der Listview der Patientdokumente anzeigen

		global PatDocs

		If !IsObject(PatDocs)
			PatDocs := Array()

		; zurück wenn kein Name enthalten ist
		doc  	:= xstring.get.Names(fname)
		If (!doc.Nn || !doc.Vn)
			return 0

		; vergleicht Patientennamen des Dokumentes mit dem des aktuell angezeigten Patienten
		matchID	:= 0
		PatID       	:= AlbisAktuellePatID()
		SSIDs     	:= admDB.StringSimilarityID(doc.Nn, doc.Vn)
		For idx, SSID in SSIDs
			If (SSID = PatID) {
				matchID := SSID
				break
			}

		; Dokument gehört nicht dem aktuellen Patienten
		If !matchID
			return 0

		; Dokument der Listview hinzufügen
		this.ListView("admReports")
		If RegExMatch(fname, "i)\.pdf$") {
			If (docID 	:= PDFpool.inPool(fname)) {
				pdf   	:= ScanPool[docID]
				PatDocs.Push(pdf)
				LV_Add((pdf.isSearchable ? "ICON3" : "ICON1")                           	; Icon
							, pdf.pages " S:"                                                               	; Seitenzahl
							, xstring.Replace.Names(xstring.Replace.FileExt(fname)))     	; Dokumentbezeichnung ohne Dateiendung
			}
		}
		else {
			LV_Add("ICON2",, xstring.replace.Names(fname))
			PatDocs.Push({"name": fname})
		}

		this.ListView("admJournal")

	return 1
	}

	ReportRemove(fname:="", rLV:="")	{         	; Dokument aus der Report (Patientenlistview) entfernen

		; bereinigt PatDocs
		global PatDocs

	; zurück wenn es zum Patienten keine Befunde  gibt
		If (!IsObject(PatDocs) || PatDocs.Count() = 0)
			return

		fname 	:= xstring.Replace.Names(fname)
		fname 	:= xstring.Replace.FileExt(fname)
		row   	:= fname ? this.GetRow(fname, 2, "admReports") : rLV ? rLV : 0
		res    	:= row > 0 ? LV_Delete(row) : -1

	 ; sucht in PatDocs (Objekt welches nur die Befunde eines Patienten enthält)
		PatDocFound := false
		For fNR, file in PatDocs
			If (file.name = fname) {
				PatDocs.RemoveAt(fNR)
				PatDocFound := true
				break
			}

		; Dokument in PatDocs, dann das Dokument auch im Patient Listview entfernen
		this.ListView("admReports")
		Loop % LV_GetCount() {
			LV_GetText(rText, row, 2)
			If RegExMatch(fname, "i)" rText "\.*\w*$")
				LV_Delete(row)
		}
		this.ListView("admJournal")

	return res
	}

}

class Befunde                                                                	{	; Datei-Handler

	; Funktion: verwaltet die physisch vorhandenen Dateien auf der Festplatte
	; die Dateioperationen finden in einem Hauptordner mit 2 Unterordnern und jeweils einem Backup-Ordner statt.

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Objektdaten erhalten / setzen
		dbgStatus[]                                                            	{ 	; getter und setter
			get {
				return this.debug
			}
		}

		FilePath[]                                                                 	{	; Aufrufen um den Hauptpfad abzurufen
			get {
				return this.fPath
			}
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Objekt Operationen
		Backup(fname)                                                       	{	; Backupdatei erstellen

			; path.fF	- original file path
			; path.fB - backup path of original files
			; path.iB - backup path of imported documents
			; path.tB - text backup/import path
			; Path.fT - OCR text path
			; Path.iT - OCR import(backup) path
			path 	:= this.GetPaths(fname)

			If !FileExist(path.fB) {
				FileCopy, % path.fF, % path.fB    ; fileFile, fileBackup
				ErrorLevel ? (path.eL := A_LastError " Backup für <" fname "> konnte nicht erstellt werden") : ""
				If this.debug && ErrorLevel
					SciTEOutput(path.eL)
			}

		return path
		}

		Rename(oldfile, newfile)                                          	{	; umbenennen

			; oF - original file path, oB - original backup path
				eL 	:= []
				path	:= this.GetPaths(newfile, oldfile)

				If this.debug
					SciTEOutput("Rename " oldfile " to " newfile)

			; Dokument umbenennen
				If !FileExist(path.oF) {
					PraxTT("Datei: <" oldfile ">`nist nicht vorhanden!", "2 0")
					eL.1 := "no file"
					return eL
				}
				FileMove, % path.oF, % path.fF, 1
				If ErrorLevel {
					ErrorLevel ? (eL.2 := A_LastError) : ""
					PraxTT("Das Umbenennen der Datei`n#1<" oldfile ">`nwurde mit dem Fehlercode (" A_LastError ") abgelehnt.", "4 0")
					return eL
				}

			; Backup-Datei umbennen, ggf. erstellen
				If !FileExist(path.oB) {

					FileCopy, % path.fF, % path.fB                                    ; wird mit dem neuen Dateinamen erstellt
					ErrorLevel ? (eL.3 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("Backup konnte nicht erstellt werden")
					else If this.debug && !ErrorLevel
						SciTEOutput("Backup Datei wurde angelegt.")

				} else {

					FileMove, % path.oB, % path.fB, 1
					ErrorLevel ? (eL.4 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("Backup konnte nicht umbenannt werden")

				}

			; Textdatei umbenennen
				If (path.oX = "pdf") {
					If FileExist(path.oT) {
						FileMove, % path.oT, % path.fT, 1
						ErrorLevel ? (eL.5 := A_LastError) : ""
						If this.debug && ErrorLevel
							SciTEOutput("Text-Datei konnte nicht umbenannt werden")
					} else {
						If this.debug
							SciTEOutput("Text-Datei nicht vorhanden")
					}
				}

		return eL.Count() > 0 ? eL : 1
		}

		RegExRename(fname, rnPattern)                              	{	; ## RegEx umbenennen für Batch-Renaming
			If RegExMatch(rnPattern, "^(?<Options>.*)?\)(?<pattern>.*)$", rx)
				RegExMatch(fname, "o" )     ; ???
			this.Rename(fname, newfile)
		}

		MoveToImports(fname)                                           	{	; Dateien in den Importiert Ordner verschieben

			; es gibt keine Funktion zum Löschen von Dokumenten! Dokumente werden nur in einen Backup-Ordner verschoben.
			; So gehen bei Programmierfehlern keine Dokumente verloren. Der PDF Backup-Ordner ist nur manuell zu leeren!

			; .fB - document backup path
			; .iB - document import path 	(second backup path)
				eL 	:= []
				path	:= this.GetPaths(fname)

			; Backup-Dokument - verschieben nach Imports
				If FileExist(path.fB) {
					FileMove, % path.fB, % path.iB, 1
					ErrorLevel ? (eL.1 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("Backup konnte nicht verschoben werden")
				}

			; Textdatei wird verschoben, insofern vorhanden
				If FileExist(path.fT) {
					FileMove, % path.fT, % path.iT, 1
					ErrorLevel ? (eL.2 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("Text-Datei konnte nicht verschoben werden")
				}

		return eL.Count() > 0 ? eL : 1
		}

		CopyToExports(fname, PatID, LastName, FirstName)  	{	; Datei in den Export-Ordner kopieren

			; kopiert die Datei aus dem Scanpool Ordner in einen patientenindividuellen Exportpfad (hinterlegt in Addendum.ini)
			; der Exportpfad wird bei Bedarf angelegt

			If !(PatPath := this.GetExportPath(PatID, LastName, FirstName))
				return 0

			If !FileExist(PatPath "\" fname) {
				FileCopy, % Addendum.BefundOrdner "\" fname, % PatPath "\" fname
				return {"success" : (ErrorLevel ? false : true), "LastError" : A_LastError, "PatPath" : PatPath}
			}

		return {"success" : true, "LastError" : "", "PatPath" : PatPath}
		}

		Remove(fname)                                                      	{	; verschiebt eine Datei aus dem Befundordner in den Backup-Ordner

			; Backup der Originaldatei erstellen wenn noch nicht erfolgt
				path := this.Backup(fname)

			; Abbruch: bei Fehlermeldung beim Kopiervorgang
				If path.eL
					return path.eL

			; Datei aus dem Befundordner nur löschen wenn ein Backup existiert
				eL := []
				If FileExist(path.fB) {
					FileDelete, % path.fF
					ErrorLevel ? (eL.Push(A_LastError " Die Orignaldatei <" path.fFN "> konnte nicht entfernt werden.")) : ""
					If this.debug && ErrorLevel
						SciTEOutput(eL[eL.Count()])
				}

			; Backup der extrahierten OCR-Textdatei erstellen
				If FileExist(path.fT) && !FileExist(path.iT) {
					FileCopy, % path.fT, % path.iT
					ErrorLevel ? (eL.Push(A_LastError " Das Backup für <" path.fFN ".txt> konnte nicht erstellt werden.")) : ""
					If this.debug && ErrorLevel
						SciTEOutput(eL[eL.Count()])
				}

			; OCR-Textdatei wird nur gelöscht wenn es ein Backup gibt
				If FileExist(path.iT) {
					FileDelete, % path.fT
					ErrorLevel ? (eL.Push(A_LastError " Die OCR Textversion von <" path.fFN "> konnte nicht entfernt werden.")) : ""
					If this.debug && ErrorLevel
						SciTEOutput(eL[eL.Count()])
				}

		return eL.Count()>0 ? eL : ""
		}

		CleanImport(fromDate)                                            	{	; ## Dateien die älter sind werden vollständig gelöscht



		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; kombinierte Datei- und Objektoperationen
		FileDelete(files, ask:=false)                                     	{ 	; löscht eine oder mehrere Dateien physisch und entfernt die Einträge aus den Objekten und Listviews

			/*  Beschreibung

				im Skript ist jetzt nur noch eine Zeile für einen Löschvorgang notwendig
					res := Befunde.FileDelete(files, true)

			*/

			; Parameterinhalt überprüfen
			If (!IsObject(files) && StrLen(files)=0) {
				PraxTT("Keine Dateinnamen zum löschen übergeben", "3 1")
				return
			}

		; Dateiarray erstellen
			filesToDel := IsObject(files) ? files : [files]

		; Nutzer fragen
			If ask {
				msg := "Wollen Sie die " (filesToDel.Count()=1 ? "eine":filesToDel.Count()) " Datei" (filesToDel.Count()>1?"en":"") " wirklich löschen?"
				MsgBox, 0x1024, Addendum für AlbisOnWindows, % msg
				IfMsgBox, No
					return
			}

		; erstellt Backups und löscht die Original und zugehörigen Dateien
			errmsg := []
			For fileNr, file in filesToDel {

				; löscht die Datei von der Festplatte [wird ein Objekt zurückgegeben, gab es eine Fehlermeldung von Windows]
				If IsObject(eL := Befunde.Remove(file))
					for errIndex, thisError in eL {
						errmsg.Push([file, thisError])
						t := t . (file ": " thisError "`n")
					}

				; Einträge aus den Objekten und Listviews entfernen
				res := PDFpool.Remove(file)       		; entfernt den Namen und die Daten aus dem PDFPool
				res := Journal.Remove(file)             	; entfernt die Datei im Listview des Journal
				res := Journal.ReportRemove(file)   	; entfernt die Datei im Listview des Patienten

			}

			If (StrLen(t)>0)
				SciTEOutput(A_ThisFunc ": Fehlermeldungen beim Löschen von Dateien:`n" t)

		return errmsg
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Hilfsfunktionen
		GetFileExt(fname)                                                   	{	; Dateiendung
			RegExMatch(fname, "i)\.(?<Ext>[a-z]+)$", File)
		return FileExt
		}

		GetPaths(fname:="", oldfile:="", path:="")                 	{	; erstellt alle notwendigen Dateipfade

			/* 	Beschreibung

				- 	erstellt Strings für alle 3 Dateipfade für die Aufteilung in Originaldatei, Backupdatei und Textdatei.
				- 	vereinfacht Lösch/Änderungsvorgänge. Jede Änderung des Dateinamen oder des des Dateipfades (z.B. nach Albis-Import) wird automatisch
					auf die anderen beiden Dateien übertragen
				-	Die Namen der Variablen sind sehr kurz gewählt um den Programmcode schmal zu halten.
					Die verkürzten Variablennamen ergeben sich aus folgenden abgekürzten Schreibweisen:

						this.Path.[BasisName][BasisOrdner oder andere Bezeichnung]

						[BasisName]
							.f		= Dateiname wenn kein Quell-Dateiname übergeben wurde, oder neuer Dateiname
							.o		= old oder Quell-Dateiname
							.i		= Dateibasis für importierte Dateien

						[BasisOrdner oder andere Bezeichnung]
							F		= Pfad + Dateiname zum Stammordner
							B		= Pfad + Dateiname zum Backupordner
							T		=	Pfad + Dateiname zum Textdateiordner
							X		= Dateiendung
							N		= nur der Dateiname

						.fFN ist nur der Dateiname (ohne Pfad) der Datei im Stammordner

				-	## austehend: legt bei Bedarf alle Unterverzeichnisse an

				-	letzte Änderung: 23.03.2021 - erstellt nur die Basis-Pfade wenn fname und oldfile leer ist

			*/

			; Hauptordner eintragen
				fpathChanged := false
				If (StrLen(path) = 0) && (StrLen(this.fPath) = 0) {
					fpathChanged := true
					this.fPath := Addendum.BefundOrdner
				}
				else if path && (this.fPath<>path) {
					fpathChanged := true
					this.fPath := path
				}

			; Unterpfade eintragen
				If this.initPaths {

					this.initPaths	:= false
					this.fBPath		:= this.fPath "\Backup"
					this.iBPath		:= this.fPath "\Backup\Importiert"
					this.fTPath		:= this.fPath "\Text"
					this.iTPath		:= this.fPath "\Text\Backup"

					If !this.pathExist(this.fBPath)
						throw A_ThisFunc ": Dokument Backup Pfad ist nicht vorhanden!"
					If !this.pathExist(this.iBPath)
						throw A_ThisFunc ": Dokument Import Pfad ist nicht vorhanden!"
					If !this.pathExist(this.fTPath)
						throw A_ThisFunc ": Textdatei Pfad ist nicht vorhanden!"
					If !this.pathExist(this.iBPath)
						throw A_ThisFunc ": Textdatei Backup Pfad ist nicht vorhanden!"
				}

			; zurück wenn beide Dateinamen leer sind
				If !fname && ! oldfile
					return this.fPath

			; Pfadstrings leeren
				this.Path := Object()

			; prüft die Dateien auf passende Dateierweiterungen
				this.Path.fX := this.GetFileExt(fname)
				If (StrLen(oldfile) > 0) {

					; Dateierweiterung erhalten
						this.Path.oX := this.GetFileExt(oldfile)

					; vergleicht die Dateierweiterungen und ergänzt bei Bedarf eine fehlende
						If (StrLen(this.Path.fX) = 0) && (StrLen(this.Path.oX) > 0)
							this.Path.fFN := (fname .= "." this.Path.oX), this.Path.fX := this.Path.oX
						else if (StrLen(this.Path.fX) > 0) && (StrLen(this.Path.oX) > 0) && (this.Path.oX <> this.Path.fX) {
							PraxTT("Dateinamen haben unterschiedliche Dateierweiterungen.`nDatei wird nicht umbenannt!", "1 0")
							return 0
						}
						else If (StrLen(this.Path.fX) > 0) && (StrLen(this.Path.oX) = 0)
							this.Path.oFN := (oldfile .= "." this.Path.fX), this.Path.oX := this.Path.fX

					; Quelldatei
						this.Path.oF := this.fPath	"\"	oldfile        	; Haupt-Pfad
						this.Path.oB := this.fBPath	"\" oldfile        	; Backup-Pfad
						If (this.Path.oX = "pdf")
							this.Path.oT := this.fTPath "\" RegExReplace(oldfile, "\.\w+$", ".txt")

				}

			; Dateipfad, Backup sowie Pfade für importierte Dateien
				this.Path.fF	:= this.fPath  	"\" fname        	; Hauptpfad
				this.Path.fB	:= this.fBPath	"\" fname        	; Backup
				this.Path.iB	:= this.iBPath 	"\" fname        	; Backup\Importpfad
				If (this.Path.fX = "pdf") {
					this.Path.fT := this.fTPath	"\" RegExReplace(fname, "\.\w+$", ".txt")
					this.Path.iT := this.iTPath	"\" RegExReplace(fname, "\.\w+$", ".txt")
				}

			; Ausgabe
				If this.debug
					For key, path in this.Path
						SciTEOutput(key ": " path)

		return this.Path
		}

		GetExportPath(PatID, LastName, FirstName)             	{	; Exportpfad für Dokumente eines Patienten erstellen

			; Dokumentenexportpfad wird bei Bedarf anlegt
				If (PatID && LastName && Firstname)
					PatSubPath := "(" PatID ") " LastName ", " Firstname
				If this.debug && PatSubPath
					SciTEOutput("PatSubPath: " PatSubPath)
				If !InStr(FileExist(PatPath := Addendum.ExportOrdner "\" PatSubPath . (PatSubPath ? "\" : "")), "D") {
					FileCreateDir, % PatPath "\"
					If !ErrorLevel {
						CreatePathFailed := true

						return 0
					}
				}

		return PatPath
		}

		FileExist(fname)                                                     	{	; gibt true zurück wenn der Befund/das Dokument physisch vorhanden ist

			path	:= this.GetPaths(fname)
			If FileExist(path.fF)
				return true

		return false
		}

		pathExist(path)                                                        	{	; gibt true zurück wenn ein Pfad vorhanden ist
			If InStr(FileExist(path), "D")
				return true
		return false
		}

		options(dbgstatus, path="")                                     	{	; Debug-Ausgaben ein oder ausschalten

			; nicht nur Debug-Ausgabe aktivieren , sondern auch für Änderungen des Scanpool Ordners ohne Skriptneustart (z.B. für Testzwecke)

			this.debug := InStr(dbgStatus, "toggle") ?  !this.debug : dbgStatus

			If path && !RegExMatch(path, "^([A-Z]\:\\|\\\\w+)")
				throw A_ThisFunc ": No valid path!`n" path
			else If path && !InStr(FileExist(path), "D")
				throw A_ThisFunc ": Path not exists!`n" path
			else if path
				this.fPath := path, this.initPaths := true
			else
				this.fPath := Addendum.BefundOrdner, this.initPaths := true

		}

}

class pdfpool                                                                	{	; verwaltet das ScanPool Objekt

	; Funktion:         	- verwaltet das Objekt (ScanPool) als virtueller Dateimanager.
	;                        	- hält deshalb zusätzliche Informationen zu den PDF-Dokumenten bereit
	; Variablen:       	- ScanPool - ist noch ein externes Objekt und muss ist deshalb Super-Global sein
	; Abhängigkeiten:	- Addendum_PdfHelper.ahk
	;
	; letzte Änderung: 30.09.2022

	inPool(fname)                                                         	{	; sucht nach Dateinamen im Pool

	; Rückgabewert ist der Indexwert im ScanPool Array oder 0 wenn die Datei nicht gefunden werden konnte

		For docID, pdf in ScanPool
			If (pdf.name = fname)
				return docID

	return 0
	}

	Add(path, fname)                                                    	{ 	; Dokument hinzufügen + Metadaten erstellen

		If !FileExist(path "\" fname)
			return

		docID	:= this.inPool(fname)
		oPDF	:= this.FileMeta(path, fname)

		If !docID
			ScanPool.Push(oPDF)
		else
			ScanPool[docID] := oPDF

	return oPDF
	}

	Update(ReIndex=false, save=true, path="")            	{ 	; neue Dateien hinzufügen, nicht mehr vorhandene entfernen

		; WICHTIG!: braucht eine globale Variable im aufrufenden Skript: ScanPool := Object()
			global admPatientTitle, admJournalTitle

		; ReIndex = true - ScanPool Inhalte werden neu erstellt oder falls leer wird das gespeicherte Objekt geladen
			ReIndex ? this.Empty() : fcount := this.Load(path)

		; alle pdf Dokumente des Befundordners dem ScanPool Objekt hinzufügen
			files  	:= this.GetFiles(Addendum.BefundOrdner, "*.pdf")
			iLen  	:= StrLen(files.Count())-1
			fcount	:= SubStr("00000" files.Count(), -1*iLen)

		; gefundene Dateien dem ScanPool Objekt hinzufügen, den Fortschritt anzeigen
			For idx, filename in files {
				If !this.inPool(filename)
					this.Add(Addendum.BefundOrdner, filename)
				GuiControl, adm: , admPatientTitle	, % (InfoText := "indiziere Dokument: " SubStr("00000" A_Index, -1*iLen) " von " fcount)
				GuiControl, adm: , admJournalTitle	, % InfoText
			}

		; nicht mehr vorhandene Dateien entfernen
			this.RemoveNoFile()

		; Speichern
			res := save ? this.Save() : ""

	return ScanPool.Count()
	}

	Remove(fname)                                                     	{	; Datei aus Scanpool Objekt entfernen
		If (docID := this.inPool(fname))
			return ScanPool.RemoveAt(docID)
	return docID
	}

	RemoveNoFile()                                                     	{	; entfernt Dateien die nicht mehr existieren
		removed := 0
		If ScanPool.MaxIndex() {
			Loop % (max := ScanPool.MaxIndex()-1)
				If !FileExist(Addendum.BefundOrdner "\" ScanPool[(docID := max-A_Index+1)].name) {
					removed ++
					ScanPool.RemoveAt(docID)
				}
		}
	return removed
	}

	Rename(oldfile, newfile)                                          	{	; Datei umbenennen
		If (docID := this.inPool(oldfile)) {
			ScanPool[docID].name := newfile
			return ScanPool[docID]
		}
	return
	}

	GetDoc(docIDName)                                               	{ 	; PDF Daten aus ScanPool erhalten
		docID := !RegExMatch(docIDName, "^\d+$") ? this.inPool(docIDName) : docIDName
	return IsObject(ScanPool[docID]) ? ScanPool[docID] : ""
	}

	GetAllDocuments()                                                 	{	; alle Dateinamen zurückgeben
		docNames := Array()
		For docID, PDF in ScanPool
			docNames.Push(PDF.name)
		If (docNames.MaxIndex() = 0)
			return
	return docNames
	}

	GetNoOCRFiles()                                                     	{ 	; PDF Dateien ohne Texterkennung
		noOCR := Array()
		For idx, pdf in ScanPool
			If !pdf.isSearchable {
				InStack := false
				For stidx, stfile in Addendum.iWin.FilesStack  	; Datei ist in Bearbeitung (z.B. OCR Vorgang), dann ignorieren
					If (stfile = pdf.name) {
						InStack := true
						break
					}
				If !Instack                              					 	; Datei hinzufügen
					noOCR.Push(pdf.name)
			}
	return noOCR
	}

	GetNewDocuments()                                              	{	; Filter für Dokumente der letzten 2 Werktage
		newDocuments := Array()
		today := A_YYYY A_MM A_DD
		weekDay := GetWeekday
		lastWorkingDay := DateAddEx(today, "3y")
		;~ For idx, pdf in ScanPool
	}

	GetNextUnnamed(StartDocID:=0)                                 	{	; das nächste nicht vollständig benannte Dokument finden
		For docID, PDF in ScanPool
			If (docID >= StartDocID) && !xstring.isFullNamed(PDF.name)
				return PDF
	return
	}

	GetNamedDocuments(refreshDocs=false)               	{	; erstellt einen Array mit den Dateinamen benannter Dokumente
		If refreshDocs
			admGui_Reload(true)
		named := Array()
		For docID, PDF in ScanPool
			If xstring.isNamed(PDF.name)
				named.Push(PDF.name)
		If (named.Count() = 0)
			return
	return named
	}

	GetFullNamedDocuments(refreshDocs=false)            	{	; erstellt einen Array mit den Dateinamen vollständig benannter Dokumente

		If refreshDocs
			admGui_Reload(true)
		fullnamed := Array()
		For docID, PDF in ScanPool
			If xstring.isFullNamed(PDF.name)
				fullnamed.Push(PDF.name)
		If (fullnamed.Count() = 0)
			return
	return fullnamed
	}

	GetUnamedDocuments(refreshDocs=false)               	{	; Array mit Dateinamen nicht vollständig benannter Dokumente
		If refreshDocs
			admGui_Reload(true)
		unnamed := Array()
		For docID, PDF in ScanPool
			If !xstring.ContainsName(PDF.name)
				unnamed.Push(PDF.name)
		If (unnamed.Count() = 0)
			return
	return unnamed
	}

	Empty()                                                                 	{	; das ScanPool Objekt komplett leeren
		If !IsObject(ScanPool) {
			PraxTT("ScanPool ist kein Objekt!", "2 1")
			return
		}
		Loop % (mxIdx := ScanPool.MaxIndex())
			ScanPool.Pop()
	return mxIdx
	}

	Load(path="")                                                         	{	; ScanPool laden
		If FileExist(ScanPoolPath := this.GetValidPath(path) "\PdfDaten.json") {
			pdfdata := FileOpen(ScanPoolPath, "r", "UTF-8").Read()
			If (RegExMatch(pdfdata, "^\{\}") || StrLen(pdfdata) = 0)
				return 0
			For docID, pdf in JSONData.Load(ScanPoolPath, "", "UTF-8")
				If FileExist(Addendum.BefundOrdner "\" pdf.name) && !this.inPool(pdf.name)
					ScanPool.Push(pdf)
		}
	return ScanPool.Count()
	}

	Refresh(ReIndex:=false)	                                          	{	; ScanPool auffrischen

		static FirstRun := true

		; ScanPool wird bei ReIndex=true komplett neu aufgebaut
		If (ReIndex || FirstRun) {

			; (ReIndex = false, save = true, Speicherpfad - Objektdaten)
			this.Update(ReIndex, true, Addendum.DBPath "\sonstiges")

			; Skriptneustart: Texterkennung auf Nachfrage starten wenn PDF Dateien ohne Textlayer vorhanden sind
			If (FirstRun && !Addendum.ExecStatus = "reload")
				If (!Addendum.Thread["tessOCR"].ahkReady() && Addendum.OCR.Client = compname)
					If IsObject(noOCRJ := this.GetNoOCRFiles()) {
						fn := Func("FirstRunOCR").Bind(noOCRJ.Count())
						SetTimer, % fn, -10000
					}

			FirstRun := false
		}
		else
			this.RemoveNoFile()                            	; nicht vorhandene Dateien aus dem Objekt entfernen

	}

	Save(path="")                                                          	{	; ScanPool speichern
	return JSONData.Save(this.GetValidPath(path) "\PdfDaten.json", ScanPool, true,, 1, "UTF-8")
	}

	; + + + + + + + INTERN + + + + + + +
	FileSize(fullfilepath)                                                 	{	; PDF Dateigröße
		FileGetSize	, FSize, % fullfilepath, K
	return FSize
	}

	FileTime(fullfilepath)                                               	{	; PDF Zeitstempel
		FileGetTime	, TimeStamp	, % fullfilepath	, C
		FormatTime	, FileTime     	, % TimeStamp	, dd.MM.yyyy
	return {"FTime" : FileTime, "TimeStamp": TimeStamp}
	}

	FileMeta(path, fname)                                            	{	; PDF Metadaten
		file := this.FileTime(fullfilepath := path "\" fname)
		return {	"name"          	: fname
				, 	"timestamp"    	: file.TimeStamp                                         				; Änderungsdatum
				, 	"filetime"        	: file.FTime                                                        		; Lesbares Tagesdatum von TimeStamp
				, 	"filesize"        	: this.FileSize(fullfilepath)
				, 	"pages"         	: PDFGetPages(fullfilepath, Addendum.PDF.qpdfPath)
				, 	"isSearchable"	: (PDFisSearchable(fullfilepath) ? true : false)}
	}

	GetFiles(path, filepattern:="*.*", options:="")           	{	; Array mit allen Dateien im Pfad erstellen
		files := Array()
		Loop, Files, % path "\" filepattern, % (options ? options : "")
			files.Push(A_LoopFileName)
	return files
	}

	GetValidPath(path)                                                  	{	; ScanPool-Pfad String prüfen
	return (this.isPath(path) ? path : Addendum.DBPath "\sonstiges")
	}

	isPath(path)                                                           	{ 	; prüft den Pfad auf Existenz
	return (StrLen(path)>0 && RegExMatch(path, "i)[A-Z]\:\\") && InStr(FileExist(path), "D") ? true : false)
	}

	FirstRunOCR(noOCRCount)                                     	{	; Auto-OCR Start bei Scriptaufruf

		MsgBox, 0x1004, % StrReplace(A_ScriptName, ".ahk")
								 , % 	"Bei " noOCRCount " PDF-Dateien wurde noch keine Texterkennung durchgeführt.`n"
								 .   	"Soll die Texterkennung jetzt gestartet werden?"
								 , 10
		IfMsgBox, No
			return
		IfMsgBox, TimeOut
			return

		admGui_OCRAllFiles()

	return
	}

}

;}

; -------- Umbenennen                                                                      	;{
admGui_Rename(filename, prompt="", newfilename="", debug=false) 	{               	; Dialog für Dokument umbenennen

	; letzte Änderung: 03.03.2023

	global  ; global zu halten sind die Variablen:  classifier, RN, RNE3-6, oPV (Sumatra Objekt)

; Initialisierung                                      	;{

	; locals & statics                                                                   	;{
		local edW, row, rows, rowFilename, startrow, lvFilename, rowL
		local cp, dp, wp, gs, title, vctrl, ENr  ;admPos
		local monIndex, admScreen
		local FileExists, BEZChanged, each, item, Items, ftime, fSize, fileout
		local RNHinweise, dcpLen, dcpMid, gsLeft, gsMid, charsNow
		local idx, found, res, pdf, mx, my

		static RNGui_init:=false                                                                        		; Initialisierung
		static loaderstate                                                                                  		; Initialisierung
		static RNminWidth := 570                                                                        	; minimalste Gui Breite
		static RNFieldNames := [ "Edit: Name, Vorname", "Edit: Dokumentdatum" 	; Debug Beschreibungen des Eingabefokus
											, "Edit: Kategorie", "Edit: Inhalt", "Combobox: Inhalt"
											, "Edit: Zusatzwörter Kategorie"
											, "Edit: Zusatzwörter Inhalt", "Edit: OCR Text"]
		static BEZPath, BEZLastSize := BEZLastTime :=0                                        	; vollständiger Pfad zur, Größe, letzte Änderungszeit Befundbezeichner Datei
		static RNAutoComplete1, RNAutoComplete2												; handle der eAutocomplete Felder
		static RNx, RNy, RNWidth                                                                         	; Koordinaten der Gui
		static wT, PatName, category, content, Datum
		static categoryO, categoryX, contentO, contentX, RNItems := 0   				; letzte Kategorie und letzter Inhalt
				; oPV                                                                                              	; Sumatra PDFReader Objekt (global)
		static oldadmFile, newadmFile, FileExt, FileOutExt, oldfilename                 	; Arbeitsdateivariablen
		static mainClassifactions, classified, subClassifications, doctxt                     	; Autonaming
		static autonaming := Object()                                                                 	; Autonaming
		static MsgBoxTimer, dbg, names                                                    	    	; Patientennamen


		; Gui-Variablen
		wT          	:= 0
		edW     	:= 100
		dbg      	:= debug
		fontSize    	:= (A_ScreenWidth > 1920 ? 10 : 9)
		admPos	:= GetWindowSpot(hadm)
		RNWidth	:= admPos.W - (2*admPos.BW)
		RNWidth	:= RnWidth < RNminWidth ? RNminWidth : RnWidth
	;}
	dbg := true
	; Ladefenster anzeigen beim ersten Aufruf                             	;{
		admGui_Loader("Hide")
		ShowLoader := !WinExist("Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI") || !RNGui_init ? true : false
		If ShowLoader {
			admGui_Loader("Show")
			admGui_Loader(loaderstate := 1)
		}
	;}

	; Patientennamen aus der Datenbank erhalten                     	;{
		If !IsObject(names) {
			names := cPat.GetNamesArray()
			If ShowLoader
				admGui_Loader(loaderstate += 1)
		}
	;}



	; Vorbereitung: Dokument umbenennen                             	;{
		If (StrLen(filename)=0)
			return

		Addendum.Flags.RNRunning := true

		SplitPath, filename,,, FileExt, FileOutExt
		oldadmfile 	:= oldfilename :=  filename
		newfilename	:= RegExReplace(newfilename, "\s*\.\w+$", "")
		examine    	:= StrLen(newfilename)=0 ? filename : newfilename
		examine   		:= RegExReplace(examine, "\.\w+$", " ")
		examine   		:= RegExReplace(examine, "i)v(\.|om)", " ")
		examine   		:= RegExReplace(examine, ",\s+", ", ")
		examine   		:= RegExReplace(examine, "\s{2,}", " ") " "

		dbg ? SciTEOutput( " [" A_ThisFunc "] " (newfilename ? "newfilename: " newfilename: "filename: " filename)) : ""

		; untersucht den übergebenen Dateinamen
		If RegExMatch(examine, 	"O)^\s*"
											. 	"(?<N>[\pL\-\s\.]+\s*,\s*[\pL\-\s\.]+)*"                                                                               	; (Nachame, Vorname)*
											.	"\s*,\s"                                                                                                                                	; ein Komma muss sein
											. 	"((?<K1>[\pL\d\-\s\.']+)\s\-\s*(?<I>[\pL\d\-\s\.']+)\s"                                                         	; (Kategorie - Inhalt
											. 	"|(?<K2>[\pL\d\-\s\.']+)\s)*"                                                                                               	; nur Kategorie)*
											. 	"((?<D1>\d{1,2}\.\d{0,2}\.*(\d{0,4}|\d{0,2})\s*\-\s*\d{1,2}\.\d{1,2}\.(\d{4}|\d{2}))"   	; Zeitraum oder
											.	"|(?<D2>\d+\.\d+\.(\d{4}|\d{2})))*"                                                                               	; Tagesdatum
											.	"(\s|$)", F) {

			; überflüssige Zeichen entfernen
			Patname	:= RegExReplace(F.N                     	, "[\s,]+$")
			category	:= RegExReplace(F.K1 ? F.K1 : F.K2	, "[\s,\-]+$")
			category	:= RegExReplace(category              	, "i)v[om\.\s]*")
			content 	:= RegExReplace(F.I                       	, "[\s,]+$")

			; ohne Kategorisierung und Inhalt landet ein Datum bei F.Kx (K  = category)
			If (!F.D1 && !F.D2 && RegExMatch(category, "(\d+\.\d*\.*\d*)*(\s*\-\s*)*(\d+\.\d+\.\d+)" ))
				Datum	:= category, category := ""
			else if (F.D1 || F.D2)
				Datum	:= F.D1 ? F.D1 : F.D2

			; debug
			dbg ? SciTEOutput(" [" A_ThisFunc "] (B)" examine "`nDatum: " Datum ", category: " category, ", Inhalt: "  Inhalt) : ""

			; wenn kein Datum ermittelt werden konnte, wird der Zeitpunkt der Erstellung verwendet
			Datum   	:= !Datum ? GetFileTime(Addendum.BefundOrdner "\" filename, "C") : Datum

		}
		else {
			Patname	:= ""
			category 	:= examine ~= "^\s*[\d_]+(\.\w+)*$" ? "" : examine
			Datum  	:= GetFileTime(Addendum.BefundOrdner "\" filename, "C")
		}


		;}

	; Klassifizierungen laden oder auffrischen                             	;{

		; Addendum DB Pfad bei Bedarf anlegen
		dbg ? SciTEOutput(A_ThisFunc "(): " Addendum.DBPath "\Dictionary") : ""
			If !InStr(FileExist(Addendum.DBPath "\Dictionary"), "D") {
				FilePathCreate(Addendum.DBPath "\Dictionary")
				If !InStr(FileExist(Addendum.DBPath "\Dictionary"), "D")
					throw A_ThisFunc ": der Pfad zum Speichern der Dateibezeichnungen konnte nicht angelegt werden`n"  Addendum.DBPath "\Dictionary"
			}
			If ShowLoader
				admGui_Loader(loaderstate += 1)

		; eine automatische Dokumentklassifizierung ist mit dieser Klasse möglich
			If ShowLoader {
				classifier := new autonamer(Addendum.DBPath "\Dictionary", "", "debug=" (dbg=true ? "true":"false") " ")
				admGui_Loader(loaderstate := 90)
			} else {
				classificationsState := classifier.update()
			}
	;}

	; Dokumentinhalt klassifizieren, Patienten zuweisen                	;{
		dbg ? SciTEOutput(A_ThisFunc "() " IsObject(classifier) ", fpath: " Addendum.BefundOrdner "\" filename ", FileExt: " FileExt) : ""
		If (IsObject(classifier) && FileExt = "pdf") {

			; Dokumenttext extrahieren, Untersuchungsmethoden festlegen
			doctxt := classifier.GetDocumentText(Addendum.BefundOrdner "\" filename)
			Messages    := (catMsg := doctxt && (!category && !content) ? 1 : 0)
			Messages  += (PatMsg := doctxt && !PatName ? 1 : 0)
			HeaderMsg := "Dokumentbezeichnung ändern`r`n" (catMsg ? "Klassifizierung":"") (catMsg & PatMsg ? " & " : "")  (PatMsg ? "Patient zuweisen": "") "`r`n"
			dbg ? SciTEOutput("Pat: " PatName ", Kategorie: " category ", Bez.: " content "`n" HeaderMsg) : ""

			; bei fehlender Klassifikation ausführen
			If catMsg {

				admGui_ImportGui("Show", HeaderMsg "# Schritt 1/" Messages ": Dokumentklassifizierung gestartet", "admJournal")

				classified   	:= classifier.matchScore(doctxt, filename)
															 classifier.rmwords := false                                                                                    	; Stopwörter nicht entfernen
				doctxt        	:= classifier.cleanup(doctxt)
				category     	:= classified.maxMainTitle	? classified.maxMainTitle 	: category
				content        	:= classified.maxSubTitle	? classified.maxSubTitle 	: content

				; sichert die ursprüngliche Liste für statistische Berechnungen (Vergleiche) ### begonnen
				autonaming.mainwords 	:= classified.maxMainTitle ? classifier.getMainWords(classified.maxMainTitle) 	: ""
				autonaming.subwords    	:= classified.maxMainTitle && classified.maxSubTitle ? classifier.getSubWords(classified.maxMainTitle, classified.maxSubTitle) : ""
				autonaming.filename    	:= filename
				autonaming.newfilename	:= ""

			}

			; bei fehlendem Patientennamen ausführen   ## wurde bereits nach der Texterkennung ausgeführt
			If PatMsg {

				;~ admGui_ImportGui((catMsg ? "update message":"Show"), (catMsg ? "" : HeaderMsg "#") "Schritt "
				;~ . (catMsg ? "2" : "1") "/" Messages ": Patientensuche gestartet", "admJournal")
				;~ fDoc  	:= FindDocNames(doctxt, dbg)
				;~ if (fDoc.names.Count() = 1)
					;~ For PatID, Pat in fDoc.names {
						;~ PatName := Pat.Nn ", " Pat.Vn
						;~ break
					;~ }

			}

			; Abschluß
			If catMsg || PatMsg
				admGui_ImportGui("update message", "Dokumentklassifizierung beendet", "admJournal")

	}
	;}

	; Gui nicht erneut erstellen, falls diese angezeigt wird          	;{
		dbg ? SciTEOutput(A_ThisFunc "() ShowLoader: " ShowLoader) : ""
		If !ShowLoader {
			gosub RNEFill
			return
		}
	;}

		; Ladefenster Fortschritt
		If ShowLoader
			admGui_Loader(loaderstate += 1)

;}

; Gui                                                  	;{

	If !RNGui_init {

		dbg ? SciTEOutput(A_ThisFunc "(): Gui zeichnen" ) : ""
		RNGui_init:=true

		;-: Gui Start                                                    	;{
		Gui, RN: new	 , % "-DpiScale  +Border +AlwaysOnTop HWNDhadmRN" ;+ToolWindow  -SysMenu +OwnDialogs
		Gui, RN: Margin, 5, 5
		Gui, RN: Color	 , cCCD8FF ;}

		;-: Hinweise                                                    	;{
		Gui, RN: Font, % "s7 cDarkBlue", Calibri
		Gui, RN: Add, Text, % "x2 y2              	BackgroundTrans ", % "Bezeichner:`nNamen:"
		Gui, RN: Add, Text, % "x+2  vRNDBG 	BackgroundTrans ", % "00000`n00000"

		cp := GuiControlGet("RN", "Pos", "RNDBG")
		RNHinweise 	:= "[ Pfeil ⏶& ⏷ von Feld zu Feld, Eingabe für Umbenennen ]"
		RNHWx     	:= cp.X
		RNHWw     	:= RNWidth-RNHWx-10

		Gui, RN: Font, % "s" fontSize - 1 " cBlack", Arial
		Gui, RN: Add, Text, % "x+5 y" cp.Y " w" RNHWw " Center vRNC1 BackgroundTrans"            	, % RNHinweise

		; debug für Eingabefokus
		Gui, RN: Font, % "s7 cDarkBlue", Arial
		Gui, RN: Add, Text, % "x2 y" cp.Y+cp.H " BackgroundTrans ", % "Eingabe:"
		Gui, RN: Add, Text, % "x" RNHWx " y" cp.Y+cp.H " w" RNHWw " vRNFocus BackgroundTrans "   		, % ""
		;}

		;-: Eingabefelder                                               	;{
		PNWidth 	:= Floor((RNWidth/3)*2)	- 5
		DTWidth 	:= Floor(RNWidth/3)     	- 10
		CTWidth 	:= RNWidth                    	- 10

		;: RNE1 - Nachname, Vorname
		Gui, RN: Font         	, % "s" fontSize-1 " cBlack", Segoe script
		Gui, RN: Add, Text  	, % "xm y+4 Left vRNT1 "                                                                            	, % "Nachname, Vorname:"
		Gui, RN: Font        	, % "s" fontSize, % Addendum.Default.Font
		Gui, RN: Add, Edit    	, % "y+0 w" PNWidth	    	" r1 vRNE1 HWNDRNhE1 gRNEHandler"            	, % Patname
		cp := GuiControlGet("RN", "Pos", "RNT1")

		;: RNE2 - Dokumentdatum/Behandlungszeitraum
		Gui, RN: Font        	, % "s" fontSize-1, Segoe script
		Gui, RN: Add, Text  	, % "x+5"  " y" cp.Y " Left vRNT2 "                                                                	, % "Dokumentdatum (*von-bis):"
		Gui, RN: Font        	, % "s" fontSize, % Addendum.Default.Font
		Gui, RN: Add, Edit    	, % "y+0 w" DTWidth      	" r1 vRNE2 HWNDRNhE2 gRNEHandler"            	, % Datum

		;: RNE3 - Kategorie
		CBOpt := " gRNEHandler Simple " ; AlbSubmit
		Gui, RN: Font        	, % "s" fontSize-1, Segoe script
		Gui, RN: Add, Text  	, % "xm y+3 Left vRNT3 " 	                                                                        	, % "Kategorie:"
		Gui, RN: Font        	, % "s" fontSize, % Addendum.Default.Font
		Gui, RN: Add, Combobox, % "xm y+0 w" CTWidth	" r8 vRNE3 HWNDRNhE3" CBOpt                     	, % category
		DllCall("UxTheme.dll\SetWindowTheme", "Ptr", RNhE3, "WStr", "Explorer", "Ptr", 0)

		;: RNE4	- Inhalt
		Gui, RN: Font        	, % "s" fontSize-1, Segoe script
		Gui, RN: Add, Text  	, % "xm y+3 Left vRNT4 " 	                                                                        	, % "Inhalt:"
		Gui, RN: Font        	, % "s" fontSize, % Addendum.Default.Font
		Gui, RN: Add, Combobox, % "xm y+0 w" CTWidth	" r8 vRNE4 HWNDRNhE4" CBOpt                  	, % content
		DllCall("UxTheme.dll\SetWindowTheme", "Ptr", RNhE4, "WStr", "Explorer", "Ptr", 0)

		;}

		;-: maximale Zeichenzahl                                        	;{
		cp := GuiControlGet("RN", "Pos", "RNE4")
		Gui, RN: Font             	, % "s" fontSize - 1 " cWhite", Segoe script
		Gui, RN: Add, Progress	, % "x0      	y" cp.Y+cp.H+2 	" w10 h20 c7B89BB vRNPG1 HWNDRNhPG1" , 100
		Gui, RN: Add, Text    	, % "x" cp.X "	y" cp.Y+cp.H+3  	"  Left Backgroundtrans vRNHW1"                 	, % "verbrauchte Zeichen (Kategorie+Inhalt+Datum):"
		dp := GuiControlGet("RN", "Pos", "RNHW1")
		GuiControl, RN: Move, % "RNPG1", % "h" dp.H+4

		Gui, RN: Font            	, % "s" fontSize - 1 " c55EC68 Bold", Segoe script	; c80C188
		Gui, RN: Add, Text    	, % "x+10                                    	Left Backgroundtrans vRNLen "       	, % SubStr("00" StrLen(category . content . Datum ), -1) . " / 70"
		;}

		;-: Dateinamenvorschau                                           	;{
		cp := GuiControlGet("RN", "Pos", "RNPG1")
		dp := GuiControlGet("RN", "Pos", "RNHW1")
		Gui, RN: Font            	, % "s" fontSize - 1 " cBlack Normal", Arial
		Gui, RN: Add, Progress	, % "x0 y" dp.Y+dp.H-1     " w" CTWidth " h" cp.H+2 " cAfC2ff vRNPG2 HWNDRNhPG2"   	, 100
		Gui, RN: Add, Text    	, % "x5 y" dp.Y+dp.H+2    " w" CTWidth " h" dp.H       " Center	Backgroundtrans vRNPV " 	, % admGui_FileName(	Patname	, category
																																																			, 	content    	, Datum)
		;}

		;-: aktuelle Datei                                               	;{
		cp := GuiControlGet("RN", "Pos", "RNPV")
		Gui, RN: Font             	, % "s" fontSize - 1 " cWhite", Segoe script
		Gui, RN: Add, Progress	, % "x0 y" cp.Y+cp.H+3           " w" CTWidth " h20 c7B89BB vRNPG4 HWNDRNhPG4"          	, 100
		Gui, RN: Add, Text    	, % "x5 y" cp.Y+cp.H+4  " w" CTWidth " h" dp.H "  Center Backgroundtrans vRNHW2"       	, % "aktuelle Datei:"

		cp := GuiControlGet("RN", "Pos", "RNHW2")
		Gui, RN: Font             	, % "s" fontSize - 1 " cBlack", % Addendum.Default.Font
		Gui, RN: Add, Progress	, % "x0 y" cp.Y+cp.H " w" CTWidth " h" cp.H   " cAfC2ff vRNPG5 HWNDRNhPG5"          	, 100
		Gui, RN: Add, Text    	, % "x5	y" cp.Y+cp.H+1 " w" CTWidth " h" cp.H " Center Backgroundtrans vRNAFN"        	, % filename

		GuiControl, RN: Move, % "RNCancel", % "h" cp.H + dp.H + 10		;}

		;-: OK, Importieren,Löschen,Abbruch                              	;{
		Gui, RN: Font, % "s" fontSize-1
		Gui, RN: Add, Button, % "xm                         	Center vRNOK  	gRNProceed"         	, % "Umbennen"
		Gui, RN: Add, Button, % "x+5                       	Center vRNImport	gRNProceed"         	, % "Umbenennen`n/Importieren"
		Gui, RN: Add, Button, % "x+20                     	Center vRNPrev 	gRNProceed"         	, % "⏮"
		Gui, RN: Add, Button, % "x+20                     	Center vRNNext 	gRNProceed"         	, % "⏭"
		Gui, RN: Add, Button, % "x+5                     	Center vRNDelete 	gRNProceed"         	, % "Dokument`nlöschen"
		Gui, RN: Add, Button, % "x+10                  	Center vRNCancel	gRNProceed"         	, % "Abbruch"
		cp     	:= GuiControlGet("RN", "Pos", "RNPrev")
		dp     	:= GuiControlGet("RN", "Pos", "RNNext")
		Gui, RN: Add, Text 	, % "x" cp.X " y" cp.Y+cp.H " w" cp.W+dp.W+20 " Center"	    	, % "Dokumente"
		cp     	:= GuiControlGet("RN", "Pos", "RNDelete") 	;}

		;-: Stichwörter                                                 	;{
		halfW 	:= Floor(CTWidth/2)-2
		Gui, RN: Add, Progress	, % "x0 y" cp.Y+cp.H+7 " w" CTWidth " h20  c7B89BB vRNPG3 HWNDRNhPG3"         	, 100

		cp := GuiControlGet("RN", "Pos", "RNPG3")
		Gui, RN: Font            	, % "s" fontSize-1 " cWhite Bold", Segoe script
		Gui, RN: Add, Text     	, % "xm y" cp.Y+2 " w" CTWidth "  Center Backgroundtrans vRNT5 " 	, % "Kategorie        < - - Klassifikationswörter - - >             Inhalt"

		cp := GuiControlGet("RN", "Pos", "RNT5")
		Gui, RN: Font        	, % "s" fontSize-1 " cBlack Normal", % Addendum.Default.Font
		Gui, RN: Add, Edit 	, % "xm y+2	w" halfW    	" r10 vRNE5 HWNDRNhE5 gRNCHandler"              	, % ""
		Gui, RN: Add, Edit 	, % "x+5      	w" halfW    	" r10 vRNE6 HWNDRNhE6 gRNCHandler"                	, % ""

		If FileExist(Addendum.DBPath "\Dictionary\Befundbezeichner.json") && Addendum.IamTheBoss {
			Gui, RN: Font, % "s8 cGrey",  % Addendum.Default.Font
			Gui, RN: Add, Button, % "xm                     	Center vRNJson         	gRNProceed"         	, % "Bezeichner editieren"
			Gui, RN: Add, Button, % "x+10                 	Center vRNANResults 	gRNProceed"         	, % "Ergebnisse ins Clipboard"
			xtra := true
		} 	;}

		Gui, RN: Add, Button, % (xtra?"x+10":"xm") "  Center vRNANSave   	gRNProceed"         	, % "Änderungen speichern"

		;-: OCR Text                                                     	;{
		dp := GuiControlGet("RN", "Pos", "RNANResults")
		Gui, RN: Font        	, % "s" fontSize-2, Segoe script
		Gui, RN: Add, Text  	, % "x" RNWidth-50 " y" dp.Y " Left vRNT6"                                        	, % "OCR Ergebnis:"
		cp := GuiControlGet("RN", "Pos", "RNT6")
		GuiControl, RN: Move,  RNT6, % "x" RNWidth-cp.W-5 " y" (dp.Y ? dp.Y+dp.H-cp.H : "+3")

		cp := GuiControlGet("RN", "Pos", "RNE4")
		dp := GuiControlGet("RN", "Pos", "RNT6")
		Gui, RN: Font        	, % "s" fontSize-2, Arial
		Gui, RN: Add, Edit, % "x" cp.X " y" dp.Y+dp.H+1 " w" CTWidth " r20 vRNE7 HWNDRNhE7 gRNCHandler" ;}

		;-: Show Hide                                                   	;{
		monIndex  	:= GetMonitorIndexFromWindow(hadm)
		admScreen	:= ScreenDims(monIndex)
		RNx := admPos.X - RnWidth, RNy := 0
		Gui, RN: Show, % "x" RNx " y" RNy " w" RNWidth " Hide", % "Addendum (Datei umbenennen)"		;}

		;-: Titel und Preview anpassen                                  	;{
		gs := GetWindowSpot(hadmRN)
		GuiControl, RN: Move, % "RNC1"	, % "w"   	gs.CW - 5
		GuiControl, RN: Move, % "RNPV"	, % "w"  	gs.CW - 5
		;}

		;-: Zeichenzähler zentrieren                                    	;{
		dp := GuiControlGet("RN", "Pos", "RNHW1"), cp := GuiControlGet("RN", "Pos", "RNLen")
		dcpLen	:= dp.W+5+cp.W
		GuiControl, RN: Move, % "RNHW1"	, % "x" (gsLeft := Floor(gs.w/2) - Floor(dcpLen/2))
		GuiControl, RN: Move, % "RNLen"  	, % "x" gsLeft + dp.W + 5
		;}

		;-: Abbruch Button verschieben                                  	;{
		dp := GuiControlGet("RN", "Pos", "RNCancel")
		GuiControl, RN: Move, % "RNCancel", % "x" gs.W - dp.W - 10		;}

		;-: Progress anpassen                                           	;{
		cp := GuiControlGet("RN", "Pos", "RNPV")
		dp := GuiControlGet("RN", "Pos", "RNT5")
		GuiControl, RN: Move, % "RNPG1"	, % "x0 w" gs.W
		GuiControl, RN: Move, % "RNPG2"	, % "x0 w" gs.W
		GuiControl, RN: Move, % "RNPG3"	, % "x0 w" gs.W " h" dp.H+4
		GuiControl, RN: Move, % "RNPG4"	, % "x0 w" gs.W ;" h" dp.H+4
		GuiControl, RN: Move, % "RNPG5"	, % "x0 w" gs.W ;" h" dp.H+4
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG1
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG2
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG3
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG4
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG5
		;}

		;-: Hotkeys
		admGui_RenameHotkeys()
		RNGui_init:=true

	}

	; nach vorne holen durch anzeigen
	Gui, RN: Show
;}

; Kassifizierung laden, Felder befüllen                	;{
	If ShowLoader
		admGui_Loader(loaderstate += 3)

	;-: Journal Show
	Journal.VisibleRows()

	; RNE1:  Patientennamen (muss nur einmal gesetzt werden, deshalb nicht bei RNEFill)
	If IsObject(names)
		RNAutoComplete1 := IAutoComplete_Create(RNhE1, Names, ["AUTOSUGGEST", "USETAB"], true)

	; RNE3:	Kategorie
	RNAutoComplete2 := IAutoComplete_Create(RNhE3, classifier.getMains()	, ["AUTOSUGGEST", "WORD_FILTER", "USETAB"], true)  ; "AUTOAPPEND"

	If WinExist("Rename LoaderGui ahk_class AutoHotkeyGUI")
		admGui_Loader(loaderstate := 100)

	;}

RNEFill:                 	                							;{	Inhalte einfügen

		LoadPDF := true

	 ; PDF Vorschau in neuem Sumatrafenster
		If !IsObject(oPV) {
			If !IsObject(oPV := admGui_PDFPreview(fileName)) {
				PraxTT("SumatraPDF konnte nicht gestartet werden.", "10 1")
				gosub RNGuiClose
				return
			}
		}

	; PDF Vorschau in vorhandenem Sumatrafenster
		If IsObject(oPV) {
			If (oPV.Previewer = "Sumatra") {

				; Dokument laden/nach vorne holen
				If !InStr(WinGetTitle(oPV.ID), filename) {
					SumatraDDE(oPV.ID, "OpenFile"	, Addendum.BefundOrdner "\" filename, 0, 0, 0)
					while (!InStr(WinGetTitle(oPV.ID), filename) && A_Index < 300)
						Sleep 20
				}

			; nur eine Seite anzeigen
				SumatraInvoke("SinglePage"	, oPV.ID)
				Sleep 200
				SumatraInvoke("FitContent"	, oPV.ID)
				Sleep 200
				Gui, RN: Default
			}
		}

	; OnMEssage wird helfen SumatraPDF zusammen mit dem Umbenennen Dialog zu minimieren oder zu maximieren
		If !Addendum.Flag.RenameGUi {
			Addendum.Flag.RenameGUi := true
			OnMessage(0x5, "admGui_OnWM_SIZE")
		}

		admGui_Loader("Hide")
		admGui_ImportGui("Hide", "", "admJournal")

		Gui, RN: Show
		Gui, RN: Default

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RNE1/2: Editfelder befüllen und Vorschau anzeigen - RNE1 u. RNE2
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
		GuiControl, RN:, RNPV	, % admGui_FileName(Patname, category, content, Datum)
		GuiControl, RN:, RNE1	, % PatName
		GuiControl, RN:, RNE2	, % Datum
	;}


	; ▪ KATEGORIE ▪ KATEGORIE ▪ KATEGORIE ▪ KATEGORIE ▪ KATEGORIE ▪ KATEGORIE ▪ KATEGORIE ▪
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RNE3: Kategorien
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (maintitles := classifier.getMains("", "|")) {
			GuiControl, RN: -Redraw, RNE3
			GuiControl, RN:, RNE3, % "|" maintitles               ; wird überschrieben
			GuiControl, RN: +Redraw, RNE3
		}

		; RNE3: Kategorien - Autocomplete auffrischen
		If (ClassificationsState = "update") {
			ClassificationsState := ""
			If IsObject(mainClassifactions := classifier.getMains())   ; brauche ich nicht mehr
				RNAutoComplete2.UpdStrings(mainClassifactions)
		}
		GuiControl, RN:, RNDBG, % classifier.countSubclassifications() "`n" names.Count()

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RNE5: Kategorie - Wortliste
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If category {

			;~ GuiControl, % "RN: " (classifier.hasMain(category) ? "ChooseString" : "Text"), RNE3, % category
			If classifier.hasMain(category)
				GuiControl, RN: ChooseString, RNE3, % category

			; RNE5: Wortliste einfügen
			GuiControl, RN:, RNE5	, % (mainwords := classifier.getMainwords(category, "`n"))  "`n"

		}


	; ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪ INHALT ▪
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RNE4: Inhalt - Liste der zur Kategorie gehörenden Inhaltsbeschreibungen anzeigen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If category {

		 ; RNE6: Wortliste einfügen
			GuiControl, RN: , RNE4, % "|" (subtitles := classifier.getSubs(category, "", "|"))		; Befüllen
			RegExReplace(subtitles, "\|",, RNItems)  ; Anzahl der Einträge ermitteln und in RNItems ablegen

		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; RNE6: Inhalt - Inhaltsbeschreibung auswählen und Wortliste anzeigen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If  category && content {

			; Inhalt auswählen
			If classifier.hasSub(category, content)
				GuiControl, RN: ChooseString, RNE4, % content

			; RNE6: Wortliste einfügen
			GuiControl, RN: , RNE6, % (subwords := classifier.getSubwords(category, content, "`n")) "`n"
			subWCount := StrSplit(subwords, "`n").Count()

		}


	; ▪ DATEI ▪ OCR ▪ DATEI ▪ OCR ▪ DATEI ▪ OCR ▪ DATEI ▪ OCR ▪ DATEI ▪ OCR ▪ DATEI ▪ OCR ▪ DATEI ▪
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; aktueller Dateinamen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		GuiControl, RN:, RNAFN, % oldadmfile     ; ja oldadmfile ist aktuell

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; OCR Text
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		GuiControl, RN:, RNE7	, % doctxt 		;

		; debug
		dbg ? SciTEOutput(" [" A_ThisFunc "] (Fill) " examine "`nDatum: " Datum ", Kategorie: " category, ", content: "  content) : ""

	; Focus je nach Inhalt der Eingabefelder setzen
		GuiControl, RN: Focus, % (!PatName ? "RNE1" : !Datum ? "RNE2" : !category ? "RNE3" : !content ? "RNE4" : "RNE1")

	; Nutzerzufriedenheit mit dem Ergebnis der Klassifizierung ;{
	/*  UserSatisfaction

		Der Nutzer wird nicht befragt. Ändert der Nutzer den Dateinamenvorschlag wird sich das sofort in der als negative Bewertung
		nieder schlagen.
		es wird am Ende eine Zahl zwischen 0-100 welche die Abweichung von der erwarteten Treffergenauigkeit (100%) zur tatsächlichen
		Zufriedenheit darstellen. Wobei jeder Aspekt (Kategorie, Inhalt) jeweils 50 Punkte beinhaltet. Wählt der Nutzer eine
		andere bereits vorhandene Kategorie aus sinkt die Zufriedenheit auf 0%. Die Wahl einer anderen Inhaltsbeschreibung reduziert
		das Ergebnis um 50%.
		Zusätzlich wird die Anzeil der Wörter in einer Liste ins Verhältnis zu 50 gesetzt. Nicht erkannte Worte führen zu einem weiteren Punktabzug.
		Dementsprechend gäbe es bei gerade mal 2 Worten in einer Liste für jedes nicht erkannte Worte 25 Punkte Abzug.
		Dieses Beispiel rechnet sich dann so: 	Kategorie 	   50
																		Inhalt	+(50 - 25) ein nicht erkanntes Wort von zwei Worten = 25
																-----------------------------------------------------------------------------------
														UserSatisfaction	  75 Punkte

		Struktur des autonaming Objektes
			- UserSatisfaction ist beschrieben
			- mainwords
			- subwords
			- filename			: der ursprüngliche Dateiname
			- newfilename	: der Dateiname nach Umbenennung

		das autonaming Objekt wird zur Auswertung einem eigens dafür entworfenen Skript übergeben

	*/
		autonaming.UserSatisfaction := category & content ? 100 : -1

	;}

;}

return

RNCHandler:                                           	;{	 Stichwörter (Klassifizierlisten)



return ;}

RNEHandler:                                            	;{   Dateinamenvorschau und Vorschläge

		Critical
		Gui, RN: Default
		Gui, RN: Submit, NoHide

	; aktuelles Eingabenfeld
		RegExMatch(A_GuiControl, "\d", gcNr)
		GuiControl, RN:, RNFocus, % StrReplace(RNFieldNames[gcNr], ":", gcNr ":") " [ " oldadmFile " ]"

	; Dateibezeichnungskonform einkürzen
		category := RegExReplace(RNE3    	, "[\*\<\>\:\\\/\|\?]")
		category := RegExReplace(category	, "\s{2,}")
		content := RegExReplace(RNE4      	, "[\*\<\>\:\\\/\|\?]")
		content := RegExReplace(content   	, "\s{2,}")

	; zu lange Zeicheneingabe verhindern (Albis begrenzt Dateinamen auf 70 Zeichen)
		charsNow := StrLen(category . content . RNE2)
		GuiControl, RN:, RNLen, % SubStr("00" charsNow, -1) " / 70 "
		If (charsNow > 70) {
			BlockInput, Send
			Send, % "{BackSpace " (charsNow - 70) "}"
			BlockInput, Off
		}

	; Kategorie - aktualisieren auch der Wortlistenanzeige RNE5 und RNE6
		If (gcNr = 3) && (categoryO != category)   {
			categoryO := category
			If (categoryX := classifier.hasMain(category, "pattern")) {
				GuiControl, RN: ChooseString, RNE3, % categoryX
				subs := classifier.getSubs(categoryX, "", "|")
				GuiControl, RN: -Redraw, RNE4
				GuiControl, RN: , RNE4, % "|" (subs < 1 ? "" : subs)
				GuiControl, RN: +Redraw, RNE4
				mainwords := classifier.getMainWords(categoryX, "`n")
				GuiControl, RN:, RNE5, % (mainwords < 1 ? "" : mainwords)
				GuiControl, RN:, RNE6, % ""
			}
			else {
				GuiControl, RN:, RNE4, % "|"
				GuiControl, RN:, RNE5, % ""
				GuiControl, RN:, RNE6, % ""
			}

		}
	 ;Inhalt - und zugehörige Wortliste RNE4 und RNE6
		else if (gcNr = 4) && (contentO != content) {
			contentO := content
			If (contentX := classifier.hasSub(category, content, "pattern")) {
				GuiControl, RN: ChooseString, RNE4, % contentX
				subwords := classifier.getSubWords(category, contentX, "`n")
				GuiControl, RN:, RNE6, % (subwords < 1 ? "" : subwords)
			}
			else {
				GuiControl, RN:, RNE6, % ""
			}
		}

	; Dateinamenvorschau auffrischen
		Gui, RN: Submit, NoHide
		GuiControl, RN:, RNPV, % admGui_FileName(RNE1, categoryX ? categoryX : category, contentY ? contentY : content, RNE2)
	; letzten Fokus

return ;}

RNProceed:                                            	;{   Datei wird umbenannt, PDF-Viewer-Vorschau wird geschlossen

		Critical
		Gui, RN: Default
		Gui, RN: Submit, NoHide
		RNBtnState := A_GuiControl
		RegExMatch(A_GuiControl, "\d", gcNr)
		GuiControlGet	, gcfocused, RN: FocusV
		GuiControl    	, RN:, RNFocus, % StrReplace(RNFieldNames[gcNr], ":", gcNr ":") ", " oldadmFile   	; Debug - zeigt Eingabefokus
		SciTEOutput("Buttonstate: " RNBtnState ", " gcfocused)

		; Datei kann nach dem Umbenennen gleich importiert werden
		If (A_GuiControl = "RNImport")                                                                   	{
			Critical, Off
			RNDocImport := true
			return
		}
		; Datei(en) löschen
		else If (A_GuiControl = "RNDelete")                                                             	{
			Critical, Off
			If !oldadmfile || !FileExist(Addendum.BefundOrdner "\" oldadmfile) {
				SciTEOutput((!oldadmfile ? "oldadmfile ist leer " : oldadmfile " - Datei ist nicht mehr vorhanden"))
				return
			}
			row 	:= Journal.FindFileRow(oldadmfile)             ; damit er eine Einsprungstelle für die Suche hat
			res 	:= Befunde.FileDelete(oldadmfile, true)
						 Journal.SelectRow(row)
			goto RNGuiBeenden
			return
		}
		; Befundbezeichner.json im Editor aufrufen
		else If (A_GuiControl = "RNJson")                                                               	{
			Critical, Off
			PraxTT("Öffne Befundbezeichner.json zum editieren.`n" Addendum.DBPath "\Dictionary\Befundbezeichner.json", "3 1")
			Run, % Addendum.DBPath "\Dictionary\Befundbezeichner.json"
			return
		}
		; mit einer Datei davor oder nach der aktuellen Datei fortfahren
		else if (A_GuiControl ~= "i)(RNNext|RNPrev)")                                             	{
			Critical, Off
			If (row := Journal.FindFileRow(oldadmfile)) {
				journal.SelectRow(row += (A_GuiControl = "RNNext" ? 1 : -1))
				If (lvFilename := journal.GetText(row, 1)) {
					MsgBox	, 0x1024, % StrReplace(A_ScriptName, ".ahk"), % "Mit dieser Datei ≪" lvFilename "≫ fortfahren?"
					IfMsgBox, Yes
						GuiControl, RN:, RNAFN, % (oldadmfile := lvFilename)
				}
				return
			}
			return
		}
		; Autonaming Ergebnisse ins Clipboard legen
		else if (A_GuiControl = "RNANResults")                                                       	{
			Critical, Off
			If IsObject(classified) {
				classifiedClone := classified.Clone()
				cleanedtxt := classifiedClone.Delete("txt")
				ClipBoard := cleanedtxt "`n`n" cJSON.Dump(classifiedClone, 1)
				ClipWait, 2
				classifiedClone := cleanedtxt := ""
				PraxTT(	"Die Autonaming Ergebnisse sind ins Clipboard kopiert worden.`n"
						. 	"Die besten " classified.titles.Count() " Treffer sind jetzt als Text verfügbar (Strg+v).", "5 1")
			}
			return
		}
		; Datenfelder speichern
		else if (A_GuiControl = "RNANSave")                                                          	{
			Critical Off
			docName := Classify_Save(dbg)
			return
		}
		; Abbruch ohne Dateinamenänderung
		else If (A_GuiControl = "RNCancel")                                                           	{    ;A_ThisHotkey != "Enter"
			Critical, Off
			;~ If (A_ThisHotkey != "Enter")
			gosub RNGuiClose
			return
		}
		; Umbenennen
		else if (A_GuiControl = "RNOK")                                                                  	{
			Critical, Off
		}
		; Enter - unterschiedliches Verhalten je nach Steuerelement
		else if (A_ThisHotkey = "Enter")                                                                    	{
			Critical, Off

			GuiControl, RN:, RNFocus, % StrReplace(RNFieldNames[gcNr], ":", gcNr ":") ", " oldadmFile
			If (gcNr > 4 || !gcNr) {
				SendInput, {Enter}
				return
			}
		}
		else
			return

		Critical, Off

	; Enter oder Button OK wurde gedrückt
	; - - - - - - - - - - - - - - - - - - - - - -
	; Dateinamen zusammensetzen
		Gui, RN: Default
		Gui, RN: Submit, NoHide
		newadmFile := admGui_FileName(RNE1, RNE3, RNE4, RNE2)
		filename := oldfilename

	; Daten zum autonaming Objekt hinzufügen vor Auswertung
		; ------

	; Check: Nutzer hat OK oder Enter gedrückt, aber nichts geändert oder der Nutzer hat alles gelöscht
		If ((newadmfile . FileExt) = filename || !newadmfile ) {
			PraxTT((!newadmfile ? "Der neue Dateiname enthält keine Zeichen!" : "Sie haben den Dateinamen nicht geändert!"), "2 1")
			gosub RNGuiBeenden
			return
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Autonaming - Inhaltsbeschreibung speichern               	;{

		; Beschreibung vorhanden?
		found := false, ncategory := ncontent := ""

		; Nutzer kann Teile einer Dateibezeichnung von der Übernahme in die Liste mit Vorschlägen aussschließen
		; entfernt alles zwischen zwei Sternchen (Elisabeth Stift - Kardiologie *Weihnachtsfeier Einladung* -> Elisabeth Stift - Kardiologie)
		; oder ab einem Sternchen oder zwei Großbuchstaben gefolgt von einem kleinen Buchstaben
		; (Elisabeth Stift - Kardiologie WEihnachtsfeier Einladung -> Elisabeth Stift - Kardiologie)
		docName := Classify_Save(false)
		msg := ""
		admGui_ImportGui("update message", msg .= docName "`r`nneue Klassifizierungsdaten wurden gespeichert`r`n", "admJournal")

		; Klassifizierer mit Eingabe auffrischen
		GuiControl, RN: -Redraw, RNE3
		GuiControl, RN:, RNE3, % "|" classifier.getMains("", "|")
		GuiControl, RN: +Redraw, RNE3

	; Felder leeren
		GuiControl, RN:, RNE1, % ""
		GuiControl, RN:, RNE2, % ""
		GuiControl, RN: Text, RNE3, % ""
		GuiControl, RN:, RNE4, % "|"
		GuiControl, RN: Text, RNE4, % ""
		GuiControl, RN:, RNE5, % ""
		GuiControl, RN:, RNE6, % ""
		GuiControl, RN:, RNE7, % ""
		GuiControl, RN:, RNAFN, % ""
	;}

	; FolderWatch pausieren, Dateie umbenennen, FolderWatch fortsetzen
		newadmFile .= "." FileExt
		admGui_FWPause(true, [newadmFile 	, oldadmfile], 3)    	; FolderWatch für 3s pausieren
		res	:= Befunde.Rename(oldadmfile  	, newadmFile)        	; Dateien umbenennen (Original, Backup, Text)
		pdf	:= PDFpool.Rename(oldadmfile 	, newadmFile)        	; ScanPool - Objektdaten ändern
		If !(row	:= Journal.Replace(oldadmfile 	, newadmFile)) {    	; Journal auffrischen
			SciTEOutput("Journal.Replace: konnte nicht die Zeile mit dem geänderten Dateinamen zurückgeben.`nSetze row/startrow auf Zeile 1 deshalb.")
			row := 1
		}

;}

RNGuiBeenden:                                         	;{   weitere Dateien umbenennen

		journal.Default("admJournal")

	; nachschauen ob weitere unvollständig benannte Dokumente vorhanden sind
		;oldadmfile 	:= newadmFile ? newadmFile : "- - - -"
		startrow    	:= row
		maxrows      	:= Journal.GetCount()

		Loop {

			; vorherige Zeile abwählen
			journal.SelectRow(row)

			; nächste Zeile auswählen. Macht weiter mit der ersten Zeile, wenn die letzte erreicht war. Jede Zeile wird nur einmal gelesen.
			row 	:= row+1 > maxrows ? 1 : row+1
			If (row = startrow || A_Index > maxrows)
				break

			; Zeile (row) auswählen, in sichtbaren Bereich scrollen, weitere ausgewählte Zeilen abwählen vorher
			journal.SelectRow(row)

			; Dokumentbezeichnung der ausgewählten Zeile auslesen
			If !xstring.isFullNamed(lvFilename := journal.getText(row, 1)) {
				res := Journal.EnsureVisible(row)                                 	; Zeile in sichbaren Bereich schieben
				break
			}

		}

	; weiteres Dokument umbenennen?
		If lvFilename {

			MsgBox	, 0x1024, % StrReplace(A_ScriptName, ".ahk"), % "Datei ≪" lvFilename "≫ umbenennen?"
			IfMsgBox, Yes
			{
				;~ admGui_ImportGui("update message", (msg .= "berechne die nächste Dateibezeichnung`r`n[" lvFilename  "] aufgerufen`r`n")	, "admJournal")
				msg := "berechne die nächste Dateibezeichnung`r`n[" lvFilename  "] aufgerufen`r`n"
				admGui_ImportGui("update message", msg, "admJournal")

				newfilename := admCM_JRecog(lvFilename, dbg, "autocall")

				msg := "neue Bezeichnung: " newfilename  "`r`nwird " A_ThisFunc " übergeben. `r`n"
				admGui_ImportGui("update message", msg, "admJournal")

				RNBtnState := "JRecog"

			 ; der Selbstaufruf löscht vermutlich den Dateinamen
				Addendum.Flag.RNRunning := false
				Patname := Datum := category := ncategory := categoryO := categoryX := content := ncontent := contentO := contentX := ""

			; verzögerte Aufruf um die Daten der ImportGui lesen zu können
				fn := Func("admGui_Rename").Bind(lvFilename, "", "", dbg)  ; ?? newfilename ??
				SetTimer, % fn, -2000
				return
			}

		}

;}

RNGuiClose:                                           	;{   Fenster definitiv schliessen
RNGuiEscape:

	SciTEOutput("Yes we close this!")

	; globale/statische Variablen zurücksetzen
		wT := Patname := category := ncategory := categoryO := categoryX := content := ncontent := contentO := contentX := Datum := ""
		RNAutoComplete1 := RNAutoComplete2 := ""
		newadmFile := FileExt := FileOutExt := oldadmfile := oldfilename := ""
		classifier := mainClassifactions := subClassifications := classified := doctxt := ""
		RNx := RNy := RNWidth := MsgBoxTimer := names := ""
		RNBtnState := RNItems := RNGui_init := 0
		;~ autonaming := Befunde := ""

	; PDF Previewer beenden
		If (oPV.Previewer = "Sumatra") {
			WinSet, AlwaysOnTop, Off, % "ahk_id " oPV.hwnd
			oPV := Sumatra_Close(oPV.hwnd)
			oPV := ""
		}

	; Gui beenden
		OnMessage(0x5, "")

		Gui, RN:   	-AlwaysOnTop
		Gui, RN: 		Destroy
		Gui, RNP: 	Destroy
		Gui, adm2: 	Destroy

		Addendum.Flag.RenameGUi := false
		Addendum.Flags.RNRunning := false

	; Journal fokussieren
		Journal.Focus()

return ;}

RNGuiUpDown:                                          	;{   mit den Pfeiltasten zwischen den Eingabefeldern wechseln

	; wenn Gui Inaktiv ist, wird das überhaupt ausgelöst?
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (A_ThisHotkey != "LButton" && !WinActive("Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI")) {
			SendInput, % "{" A_ThisHotkey "}"
			dbg ? SciTEOutput(A_LineNumber " ausgelöst durch: <" A_ThisHotkey ">") : ""
			return
		}

	; fokussiertes Steuerelement feststellen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		GuiControlGet, ctrlfocused, RN: FocusV
		RegExMatch(ctrlfocused, "\d+", ENr)
		GuiControl, RN:, RNFocus, % StrReplace(RNFieldNames[ENr], ":", ENr ":")  ", Hotkey: <" (lastHotkey := A_ThisHotkey) ">"
		RNENr 	:= % "RNE" 	ENr
		RNhENr:= % "RNhE"	ENr

	; springen zwischen den Steuerelementen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If (A_ThisHotkey = "Up")        	{
			If (ENr = 1)
				GuiControl, RN:Focus, RNE3
			else if (ENr ~= "(3|4)") {    ; 3 - Hauptkategorie, 4 - Unterkategorie
				RNEItems := CB_GetCount(%RNhENr%)
				GuiControlGet, itemNR, RN:, % RNENr
				itemNR := itemNR-1 < 1 ? RNEItems : itemNR-1
				GuiControl, RN:, RNFocus, % StrReplace(RNFieldNames[ENr], ":", ENr ":") ", Hotkey: <" lastHotkey ">, itemNr: " itemNR
				GuiControl, RN: Choose, % RNENr, % itemNR
				GuiControl, RN: Focus	, % RNENr
			}
			else if (ENr ~= "(5|6)")   	; 5 - Hauptkategorie - identifizierende Worte/Sätze, 6 - Unterkategorie -  identifizierende Worte/Sätze
				SendInput, % "{" A_ThisHotkey "}"
			else
				SendInput, {LShift Down}{Tab}{LShift Up}
		}
		else if (A_ThisHotkey = "Down")	{
			If (ENr ~= "(3|4)") {
				RNEItems := CB_GetCount(%RNhENr%)
				GuiControlGet, itemNR, RN:, % RNENr
				itemNR := itemNR+1 > RNEItems  || !itemNR ? 1 : itemNR+1
				GuiControl, RN:, RNFocus, % StrReplace(RNFieldNames[ENr], ":", ENr ":") ", Hotkey: <" lastHotkey ">, itemNr: "  itemNR
				GuiControl, RN: Choose, % RNENr, % itemNR
				GuiControl, RN: Focus	, % RNENr
			}
			else if (ENr ~= "(5|6)")
				SendInput, % "{" A_ThisHotkey "}"
			else
				SendInput, {Tab}
		}

	; Clipboardinhalt einfügen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		else if (A_ThisHotkey = "$^v") 	{
			toSend := ClipBoard
			SendInput, {LControl Down}v{LControl Up}
			if (ENr ~= "(5|6)")
				SendInput, {Enter}
		}

	;~ SendMessage, 0x146, 0, 0,, % "ahk_id " %RNhENr% ; CB_GETCOUNT:= 0x0146

return ;}

RNGuiChoose:                                          	;{   mit den Pfeiltasten zwischen den Eingabefeldern wechseln

	; je nach AutoComplete Listbox
		;~ ahwnd := WinExist("A")
		;~ WinGetActiveTitle, wTtitle
		;~ GuiControlGet, gcFocus, RN: FocusV
		;SciTEOutput("active: " wTitle)
		;~ SendInput, % (!ACGShow ? "{Tab}" : "{Space}" )

return
;}

MoveNextMBox:                                         	;{   kürzerer Mausweg, MsgBox zum Mauspfeil verschieben

	MsgBoxTimer ++
	If (MsgBoxTimer > 20) {
		SetTimer, MoveNextMBox, Off
		return
	}
	If !(hwnd := WinExist("Addendum für ahk_class #32770", "Möchten Sie mit der nächsten"))
		return
	SetTimer, MoveNextMBox, Off

	dbg ? SciTEOutput("hMsgBox: " GetHex(hwnd)) : ""
	MouseGetPos, mx, my
	wp := GetWindowSpot(hwnd)
	my := my-wp.H < 0 ? 1 : my
	If !IsInsideVisibleArea(mx, my, wp.W, wp.H, CoordInjury) {
		If (CoordInjury ~= "(x|w)")
			mx := wp.X
		If (CoordInjury ~= "(y|h)")
			my := wp.Y
	}
	SetWindowPos(hwnd, mx, my, wp.W, wp.H)

return ;}

}

admGui_RenameHotkeys(state:="On")                                	{                	; schaltet Hotkey's an oder aus

	Hotkey, IfWinActive, % "Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI"
	Hotkey, Escape  	, RNGuiClose	    , % state
	Hotkey, Up        	, RNGuiUpDown	, % state
	Hotkey, Down   	, RNGuiUpDown	, % state
	Hotkey, ^v         	, RNGuiUpDown	, % state
	Hotkey, $Enter     	, RNProceed			, % state
	;~ Hotkey, ~LButton	, RNGuiUpDown	, % state
	Hotkey, IfWinActive

}

Classify_Save(dbg:=false)                                        	{	              	; eingegebene Daten dem Klassifizierer übergeben

	global classifier, RN, RNE3, RNE4, RNE5, RNE6
	global ncategory, nocontent

	If !IsObject(classifier)
		return

	Gui, RN: Submit, NoHide

	classifier.Debug := dbg
	res := classifier.setSub(admGui_RemoveTemps(RNE3), RNE5, admGui_RemoveTemps(RNE4) , RNE6)
	res := classifier.Save()
	classifier.Debug := false

	dbg ? SciTEOutput("RNE5: " RNE5 "`n- - - - - - - - - - - `nRNE6: " RNE6 "`n.Save(): " res) : ""

return ncategory (ncontent ? " - " ncontent : "")
}

Classify_Files(showProgress:=true)                               	{                	; alle PDF Dokumente im (Befund-)Ordner klassifizieren lassen

	; für die weitestgehend vollständige Automatisierung der Erstellung von Dokumentbezeichnungen gedacht.
	; Entwicklungsschritt 1: alle Dateien untersuchen und die gewonnenen Ergebnisse in PDFDaten.json speichern
	; Entwicklungsschritt 2: Multithreading implementieren
	; Entwicklungsschritt 3: das Dokumentdatum muss besser erkannt werden (Dokumenttext)
	; Entwicklungsschritt 4: falls das Autonaming circa 99% Genauigkeit erreichen sollte, dann soll so eine Datei auf gleich einen neue Bezeichnung erhalten

	classifier := new autonamer(Addendum.DBPath "\Dictionary"
											, {"pdfpath":Addendum.BefundOrdner, "pdfdatapath":Addendum.DBPath "\sonstiges\PdfDaten.json"}
											,  "Debug=false ShowCleanupedText=false RemoveStopWords=true Save_immediately=false")

	If IsObject(documents := classifier.readDirectory("only_unnamed=true")) {

		; einfache Progress Gui
		If showprogress {
			Progress	, % "B2 zH25 w600 WM400 WS500"
						.  	 " 	cW" 	Addendum.Default.BgColor1
						. 	 " 	cB" 	Addendum.Default.PRGColor
						. 	 " 	cT" 	Addendum.Default.FntColor
						, % "Klassifiziere Dokument: "
						, % "Addendum für AlbisOnWindows"
						, % "Dokumentklassifizierung"
						, % Addendum.Default.Font
			hprg := WinExist("Dokumentklassifizierung ahk_class AutoHotkey2")
		}

		; Dateiliste klassifizieren
		factor := 100/documents.Count()
		For fIndex, file in documents {

			If showprogress {
				rc := Round(fIndex * factor)
				ControlSetText, Static2, % "Klassifiziere Dokument [" SubStr("000" fIndex, -1*(StrLen(documents.Count())-1)) "/" documents.Count() "]: " file.name
									, % "ahk_id " hprg
				Progress % rc
			}

			If (!file.isSearchable || file.category || !FileExist(file.path "\" file.name))
				continue

			doctxt	:= classifier.getDocumentText(file.path "\" file.name)
			result 	:= classifier.matchscore(doctxt, file.name)
			val    	:= result.Delete("txt")

			; die besten 3 Ergebnisse werden herausgesucht
			bestof 	:= [Array(), Array(), Array()]
			For score, data in result.titles {

				if (score ~= "(" (result.max1 ? result.max1 : "#") "|" (result.max2 ? result.max2 : "#") "|" (result.max3 ? result.max3 : "#")  ")") {
						obj := {"category":data.maintitle, "content":(data.subTitle="--" ? "" : data.subTitle)}
							If (score = result.max1)
							bestof.1.Push(obj)
						else If (score = result.max2)
							bestof.2.Push(obj)
						else If (score = result.max3)
							bestof.3.Push(obj)
				}

			}

			; Entfernen leerer Einträge
			Loop 3
			 If (bestof[4-A_Index].Count() = 0)
				bestof.RemoveAt(4-A_Index)

			; bestes Ergebnis sichern
			file.category 	:= bestof.1.1.category
			file.content 	:= bestof.1.1.content
			file.bestof    	:= bestof.titles

		}

		; Backup der PDF-Daten.json Datei
		If !InStr(FileExist(Addendum.DBPath "\sonstiges\_Backup"), "D")
			FileCreateDir, % Addendum.DBPath "\sonstiges\_Backup"
		Loop 6	{	; maximal 6 Backups

			If !FileExist(fpath := Addendum.DBPath "\sonstiges\_Backup\_backup-PdfDaten" A_Index ".json") {
				fpath := Addendum.DBPath "\sonstiges\_Backup\_PdfDaten"  (A_Index > 1 ? A_Index-1 : 1) ".json"
				break
			}

			FileGetTime, ftime, % fpath , M
			If (!ftimeLast || ftimeLast > ftime) {
				ftimeLast := ftime
				fpath := Addendum.DBPath "\sonstiges\_Backup\_backup-PdfDaten" A_Index ".json"
			}

		}

		FileCopy, % Addendum.DBPath "\sonstiges\PdfDaten.json", % fpath, 1
		FileOpen(Addendum.DBPath "\sonstiges\PdfDaten.json", "w", "UTF-8").Write(cJSON.Dump(documents, 1))
	}

}

Classify_TextIdentifier(doctxt)                                  	{
					; Telefon- und Faxnummern																			Betriebsstättennummer
	ident := ["i)(?<TelFax>(Tel|Fax):*)*\s*(?<Nr>0\d{2}[\s/]*\d[\s\-]*[\d\s\-]{8,15})", "i)(?<BSNR>BSNR:*)\s*(?<Nr>[\d\s/]+)"
					; Postleitzahl und Ort
			,	  "i)(?<PLZ>\d{5})\s*(?<Ort>[A-ZÄÖÜ][\pL-]+)", "(?<street>[A-ZÄÖÜ][\pL\-]+\s+(Straße|Str\.*|Platz|Weg)\s+\d+\s*[A-Za-z]{0,1}\s)"]


}

admGui_FileName(PatName, category, content, Datum)              	{                	; Hilfsfunktion: admGui_Rename

	SPos     	:= 1
	category 	:= RegExReplace(Trim(category), "\s{2,}")
	category	:= RegExReplace(category, "[\*\<\>\:\\\/\|\?]")
	while (SPosN := RegExMatch(category, "[A-ZÄÖÜ]\K[A-ZÄÖÜ](?=[a-zäöüß])", upperchar, SPos)) {
		category := RegExReplace(category, "[A-ZÄÖÜ]\K[A-ZÄÖÜ](?=[a-zäöüß])", Format("{:L}", upperchar),, 1, SPos)
		SPos := SPosN
	}

	SPos 	:= 1
	content 	:= RegExReplace(Trim(content), "\s{2,}", " ")
	content	:= RegExReplace(content, "[\*\<\>\:\\\/\|\?]") ; entfernt Zeichen die in Dateinamen nicht vorkommen dürfen
	while (SPosN := RegExMatch(content, "[A-ZÄÖÜ]\K[A-ZÄÖÜ](?=[a-zäöüß])", upperchar, SPos)) {
		content := RegExReplace(content, "[A-ZÄÖÜ]\K[A-ZÄÖÜ](?=[a-zäöüß])", Format("{:L}", upperchar),, 1, SPos)
		SPos := SPosN
	}
	content 	:= RegExReplace(content, "\s{2,}", " ")

	retStr := RegExReplace(PatName, "[\s,]+$") ", "
	retStr .= Trim(RegExReplace(category, "[\s,]+$")) " - "
	retStr .= Trim(RegExReplace(content, "[\s,]+$"))
	retStr .= Datum ? " " Trim(RegExReplace(Datum, "^[\s,vom\.]+")) : ""

return retStr
}

admGui_RemoveTemps(string)                                       	{               	; Strings bearbeiten
string := RegExReplace(string, "\*.*?\*")
;~ string := RegExReplace(string, "(\*|[A-ZÄÖÜ][A-ZÄÖÜ][a-zäöüß]).*$")
return string
}

admGui_Loader(state)                                             	{               	; hält den Nutzer bei Laune indem es ihn informiert

	; Fenster wird versteckt

	global admPos, hRNP, RNPprg1
	static RNPText1, RNPText2

	If (state~="i)Show") {
		If !WinExist("Rename LoaderGui ahk_class AutoHotkeyGUI") {

			Gui, RNP: New, % "AlwaysOnTop -DpiScale -Caption +Toolwindow hwndhRNP"
			Gui, RNP: Color, % "c" Addendum.Default.BGColor
			Gui, RNP: -DpiScale

			Gui, RNP: Font, s12, % Addendum.Default.BoldFont
			Gui, RNP: Add, Text    	, % "xm ym w" admPos.CW-40 " 	vRNPText1 cLightGreen Center BackgroundTrans"                	, % "bitte warten.... "

			fgcolor := DefaultProgressColor, bgcolor := Addendum.Default.BGColor1
			Gui, RNP: Add, Progress, % "x0	y+5 w" admPos.CW " h12 vRNPprg1 -Smooth c" fgcolor " Background" bgcolor           	,  0

			Gui, RNP: Font, s21 Normal, % Addendum.Default.BoldFont
			Gui, RNP: Add, Text    	, % "xm	y+5 w" admPos.CW-40 " 	vRNPText2 cWhite Center BackgroundTrans"                        	, % "Daten werden vorbereitet"

		}
		Gui, RNP: Show, % "x" admPos.X " y" admPos.Y+admPos.H " w" admPos.CW-40 " AutoSize NA", % "Rename LoaderGui"
	}
	else if (state~="i)(Hide|Destroy)") {
		Gui, RNP: Show, Hide
	}
	else if (state ~= "^\d+$") {
		GuiControl, RNP:, % "RNPprg1", % state
	}

}

admGui_OnWM_SIZE(wParam, lParam, msg, hWnd)                     	{	              	; minimiert o. maximiert SumatraPDF synchron zum Umbenennungsdialog

		global oPV, hadm, hadmRN
		static smtra

		If (hwnd = hadmRN) {

			;~ SciTEOutput(GetHex(hwnd) "=" hadmRN ": " oPV.hwnd ", " msg ", " wParam "," GetHex(lParam) )
			If (oPV.Previewer = "Sumatra") {
				If (wParam=1) {      	; is Minimized
					smtra := GetWindowInfo(oPV.hwnd)
					WinMinimize, % "ahk_id " oPV.hwnd   ; Sumatra PDF auch minimieren
				}
				else if (wParam=0) {	; anzeigen
					WinMaximize, % "ahk_id " oPV.hwnd
					SetWindowPos(oPV.hwnd, smtra.X, smtra.Y, smtra.W, smtra.H)   ; Sumatra PDF auf ursprüngliche Größe zurück
				}
			}
		}

}


class documents                                                  	{              	; ### begonnen


	__New(fname) {

		; hält nachher die Daten
			this.Designation := Array()

		; Addendum DB Pfad bei Bedarf anlegen
			If !InStr(FileExist(Addendum.DBPath "\Dictionary"), "D") {
				FilePathCreate(Addendum.DBPath "\Dictionary")
				If !InStr(FileExist(Addendum.DBPath "\Dictionary"), "D")
					throw A_ThisFunc ": der Pfad zum Speichern der Dateibezeichnungen konnte nicht angelegt werden`n"  Addendum.DBPath "\Dictionary"
			}

		; Dateipfad
			this.filepath := Addendum.DBPath "\Dictionary\" fname

	}

	Load() {

		; Datei laden wenn vorhanden
			If FileExist(this.filepath) {
				fileItems := FileOpen(BEZPath, "r", "UTF-8").Read()
				this.sortitems(fileitems)
			}

	return this.Designations
	}

	Save() {

		tmpitems := []
		For each, item in this.Designations
			If (StrLen(Trim(item)) > 0)
				tmpitems.Push(item)

		this.sortitems(tmpitems)

	}

	SortItems(ItemsToSort) {



	}


}

;}

; -------- PDF/Bild Viewer                                                                 	;{
admGui_PDFPreview(fileName)                                      	{               	; PDF Vorschau

		global 	hadm, PIDSumatra, hadmRN

		static 	hadmPV, PVPic1, PVbw, PVfw, imagick
		static 	SmtraExist := false, SmtraInit := true
		static 	SumatraCMD
		static 	SmtraClass := "ahk_class SUMATRA_PDF_FRAME"
		static		A4AspectRatio	:= 1.25

	; Fensterpositionen ermitteln
		admPos	:= GetWindowSpot(hadm)
		admRN  	:= GetWindowSpot(hadmRN)
		Mon      	:= GetMonitorInfoFromWindow(hadmRN)

	; Dateipfad anpassen
		pdfPath	:= RegExMatch(filename, "[A-Z]\:\\") ? fileName : Addendum.BefundOrdner "\" filename
		pages 	:= PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)

	; nutzt den Sumatra PDF READER zur Vorschau (sehr schnelle PDF Darstellung!)
	; prüft ob Sumatra installiert ist, falls nicht wird die PDF Datei mit imagemagick konvertiert
		If SmtraInit {
			SmtraInit := false
			SumatraCMD := GetAppImagePath("SumatraPDF")
			If (StrLen(SumatraCMD) > 0) && FileExist(SumatraCMD)
				SmtraExist := true
			else
				SumatraCMD := ""
		}

	; ohne SumatraPDF Bilder mit ImageMagick convert auspacken und in eigener Gui zeigen
		If !SumatraCMD {

				If (StrLen(imagick) = 0) && FileExist("T:\cmd\convert.exe") {
					tempPath 	:= "T:"
					imagick  	:= tempPath "\cmd\convert.exe"
				} else {
					tempPath	:= A_Temp
					imagick 	:= Addendum.PDF.imagickPath "\convert.exe"
				}

				pdfconvert	:= imagick " -resize x" Round(Mon.H * 0.8) " -border 3 " q pdfPath q " " q tempPath "\prevPic.jpg" q
				prevPic  	:= tempPath "\prevPic" (pages = 1 ? "" : "-1") ".jpg"
				Stdout   	:= StdoutToVar(pdfconvert) ;,, Hide

				SciTEOutput(A_LineNumber ": Ich bin mal hier und wieder weg.")

				If !FileExist(prevPic) {
					PraxTT("PDF Vorschau konnte nicht erstellt werden.", "2 0")
					return
				}

				Gui, admPV: New 	, % "+ToolWindow -DPIScale +HWndhadmPV +Border"
				Gui, admPV: Margin	, 5, 5
				Gui, admPV: Add   	, Button	, % "xm 	ym 	vPVbw 	0xE BackGroundTrans gPVHandler"	, % " << "
				Gui, admPV: Add   	, Button	, % "x+10 ym 	vPVfw 	0xE BackGroundTrans gPVHandler" 	, % " >> "
				Gui, admPV: Add   	, Picture, % "xm  	y+5 	vPVPic1	0xE BackGroundTrans"                   	, % prevPic
				Gui, admPV: Show	, % "AutoSize Hide", % "Pdf Vorschau - Seite 1/" pages " [" filename "]"
				PV := GetWindowSpot(hadmPV)
				Gui, admPV: Show, % "x" (admPos.X - PV.W - 5) " y" Round(Mon.B/2) - Round(PV.H/2) " NA"
				activeID := hadmPV, Previewer := "admPV"

		}
	; automatische Größenanpassung des Sumatra PDF Reader
		else {

			; DetectHidden... merken
				lDetectHiddenWin := A_DetectHiddenWindows
				DetectHiddenWindows, On

				Run, % q SumatraCMD q " -view " q "single page" q " -zoom " q "fit page" q " " q pdfPath q,, UseErrorLevel, PIDSumatra   ; -new-window
				WinWait        	, % SmtraClass,, 12
				WinActivate		, % SmtraClass
				WinWaitActive	, % SmtraClass,, 4

				hSumatra  	:= WinExist(SmtraClass)
				WinGet      	,	SumatraPID, PID, % "ahk_id " hSumatra
				ControlGet	,	hSumatraCnvs, HWND,, SUMATRA_PDF_CANVAS1, % "ahk_id " hSumatra
				WinGetPos	,,	stY,, stH, ahk_class Shell_TrayWnd

				rnGui    		:= WinExist("Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI") ? GetWindowSpot(hadmRN) : 0
				smtraWin   	:= GetWindowSpot(hSumatra)
				smtraCnvs 	:= GetWindowSpot(hSumatraCnvs)
				smtraTBarH	:= smtraWin.H - smtraCnvs.H
				smtraH     	:= Mon.H - stH
				smtraW     	:= Floor(smtraH/A4AspectRatio)              			; A4 Seite 29,7:21,0 (cm) ~ 1.41 --- geht nur 1.2?
				smtraW     	:= admRN.X - smtraW - 5 < Mon.X ? admRN.X-Mon.X-5 : smtraW
				smtraX      	:= admRN.X - smtraW - 5

				;~ SciTEOutput("hSumatra: " hSumatra ", " smtraW ", " smtraH  "," Mon.X ", " monIndex)
				DllCall("SetWindowPos", "Ptr", hSumatra, "Ptr", 0, "Int", smtraX, "Int", Mon.Y, "Int", smtraW, "Int", smtraH, "UInt", 0x40) ;SHOW_WINDOW
				WinSet, AlwaysOnTop, On, % "ahk_id " hSumatra

				DetectHiddenWindows, % lDetectHiddenWin
				activeID := hSumatra, activePID := SumatraPID, Previewer := "Sumatra"
			}

return {"Previewer":previewer, "hwnd":activeID, "ID":activeID, "PID":activePID, "Path":SumatraCMD}

PVHandler:
return
}

admGui_View(filename, MenuName="")                              	{               	; Befund-/Bildanzeigeprogramm aufrufen

	; letzte Änderung 19.06.2022

	;~ static pdfReaderOnePath, pdfReaderTwoPath

	filepath := Addendum.BefundOrdner "\" filename
	If (StrLen(filename) = 0 || !FileExist(filepath))
		return 0

	If RegExMatch(filename, "\.jpg|png|tiff|bmp|wav|mov|avi$") {
		PraxTT("Anzeigeprogramm wird geöffnet", "3 0")
		Run % q filepath q
		return 1
	}
	else {

		If !FileExist(Addendum.PDF.ReaderAlternative) && !FileExist(Addendum.PDF.Reader) {
			PraxTT("Es konnte kein PDF-Anzeigeprogramm gefunden werden", "2 0")
			return 0
		}

		pdfReaderPath := !MenuName || MenuName="JView1"
								? 	(Addendum.PDF.ReaderAlternative 	? Addendum.PDF.ReaderAlternative	: Addendum.PDF.Reader)
								: 	(Addendum.PDF.Reader                 	? Addendum.PDF.Reader               	: Addendum.PDF.ReaderAlternative)

		If !PDFisCorrupt(filepath) {
			PraxTT("PDF Datei [" filename "] wird angezeigt", "2 0")
			Run % q pdfReaderPath q " " q filepath q
			return 1
		} else {
			PraxTT("PDF Datei [" fileName "] ist defekt", "2 0")
			return 0
		}
	}


}

admGui_GetAllPdfWin()                                            	{               	; Namen aller angezeigten PdfReader ermitteln

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDF.ReaderWinClass
	Loop % tmpListe
		PreReaderListe .= tmpListe%A_Index% "`n"

return PreReaderListe
}

admGui_ShowPdfWin(PreReaderListe)                                	{               	; PDF-Readerfenster nach vorne holen

	; funktioniert nicht wenn man nur ein Fenster (eine Instanz) des PdfReader eingestellt hat
	; die Tabs im FoxitReader bekomme ich nicht ermittelt und somit auch nicht angesteuert

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDF.ReaderWinClass
	Loop % tmpListe
		ReaderListe .= tmpListe%A_Index% "`n"
	ReaderListe := PreReaderListe "`n" ReaderListe
	Sort, ReaderListe, U                              	; entfernt die vor dem Importvorgang geöffneten Readerfenster aus der Aktivierungsliste
	Loop, Parse, ReaderListe, `n
	{
		WinActivate    	, % "ahk_id " A_LoopField
		WinWaitActive	, % "ahk_id " A_LoopField,, 1
	}

}

admGui_ThumbnailView(lparam, wParam, xParam)                    	{                	; PDF-Dateien sollen hiermit einblendbar werden ohne einen PDF-Reader aufzurufen

	CoordMode, Mouse, Screen
	global admHJournal, admHReports, hadm, adm

	MouseGetPos, mx, my, hWin, hControl, 2
	oAcc := AccObj_FromPoint(mx, my)
	role := "hwin: " xParam ", hControl: " wParam "`nChildCount: " oAcc.accChildCount "`nName: " oAcc.accName(1) "`nValue: " oAcc.accValue(1)
	;ToolTip, % role, 800, 1, 15

}

;}

; -------- Importieren / Exportieren                                                      	;{
admGui_ImportGui(ShowGui=true, imprtmsg="", gCtrl="")            	{                 	; Fenster für laufende Prozesse

	; letzte Änderung: 17.12.2022

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Variablen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		global admImporter, admImporterPrg, admJournal, admImportCancel, admImportHeader1, admImportHeader2
		global admHTab, breakImportAll, adm, adm2, hadm, hadm2, admImportPrg1

		static  last_msg, last_htext, last_ctrl, thisMsg, admHImporter
		static adm2Init := false, dbg := false
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Aufruf ohne Parameter verhindern
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		;~ If !imprtmsg {
			;~ if (ShowGui~="i)^\s*(0|false|hide|destroy|off|)\s*$")
				;~ gosub adm2_Hide
			;~ return
		;~ }
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; das richtige Tab-Control anzeigen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If (!gCtrl || gCtrl = "admReports")
			res := admGui_ShowTab("Patient"), gCtrl := "admReports"
		else if (gCtrl = "admJournal")
			res := journal.Show()
		last_ctrl	:= gctrl
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; VORBEREITUNGEN
	; Kopf und flexiblen Nachrichtenteil trennen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		RegExMatch(imprtmsg, "O)^((?<Header>.+?)\n(?<HeaderText>.+)#)*(?<BodyText>.+)*$", thisMsg)
		If !StrLen(Trim(thisMsg.Header . thisMsg.HeaderText . thisMsg.BodyText)) {
			;~ SciTEOutput(A_ThisFunc " " SDbg() "imprtmsg is empty: " imprtmsg)
			Header := ["Aufruf des Fortschrittsfensters mit fehlender Nachricht", "msg = " (imprtmsg ? imprtmsg : " leer") ", Show = " ShowGui ", gCtrl = " gCtrl ]
			;~ SciTEOutput(A_ThisFunc " " SDbg() "Header: " thisMsg.Header "`n                                 HeaderText: " thisMsg.HeaderText "`n                                 " thisMsg.BodyText)
			gosub adm2_Hide
			return
		}

		last_htext := thisMsg.HeaderText
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; GUI
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	; die Gui wird nur einmal erstellt
		If !adm2Init {

			adm2Init 	:= true
			rp := GetWindowSpot(hadm)
			W	:= rp.CW-(2*rp.BW)

			Gui, adm2: New	, +ToolWindow AlwaysOnTop +HWNDhadm2 -DPIScale 	; -Caption +Parent%hadm%
																																; spätestens nach einem Karteikartenwechsel stirbt die GUI dann automatisch
			Gui, adm2: Color	, c909090, cF0F0F0
			Gui, adm2: Margin, 0, 0

			;-: Titel
			Gui, adm2: Font	, s14 q5 cBlue Bold underline, % Addendum.Default.Font
			Gui, adm2: Add	, Text, % "xm y10 w" W " Center Backgroundtrans vadmImportHeader1"                  	, % thisMsg.Header
			GuiControlGet, cp, adm2: Pos, admImportHeader1

			;-: Titeltext
			Gui, adm2: Font	, s11 cGrey Normal
			Gui, adm2: Add, Text, % "xm y+3 w" W " Center Backgroundtrans  vadmImportHeader2"                   	, % thisMsg.HeaderText
			GuiControlGet, cp, adm2: Pos, admImportHeader2

			;-: Nachricht
			Gui, adm2: Font	, s9 cBlack
			adm2opt := (gCtrl = "admReport" ? "Center" :"") " T28 T30 T32 T40 T48 T64 hwndadmHImporter vadmImporter"
			Gui, adm2: Add	, Edit	, % "x10 y+5 w" W-10 " h" (rp.CH-cpY-cpH-5) " " adm2opt                           	, % thisMsg.BodyText
			Gui, adm2: Add	, Progress	, % "x10 y+0 w" W-10 " h5 vadmImportPrg1 "                                         	, 0

			;-: Button
			Gui, adm2: Add	, Button	, % "x20 y+10  vadmImportCancel gadm2_Handler"                                       	, % "Vorgang abbrechen"

			;-: Anpassen der Elemente und der Gui
			Gui, adm2: Show	, % "x" rp.X " y" rp.Y " w" rp.CW " h" rp.H+rp.BW*2 " Hide"

			GuiControlGet	, cp, adm2: Pos	, admImportCancel
			GuiControlGet	, dp, adm2: Pos, admImporter

			GuiControl       	, adm2: Move	, admImportCancel	, % "x" rp.CW-cpW-30 " y" rp.CH-cpH-10
			GuiControl       	, adm2: Move	, admImporter       	, % "h" rp.CH-dpY-cpH-20

			EM_SETMARGINS(admHImporter, 3, 3)                                               ; Ränder Edit Steuerelement

			WinSet, Style         	, 0x50000004	, % "ahk_id " admHImporter
			WinSet, ExStyle         	, 0x00000000	, % "ahk_id " admHImporter
			WinSet, Style         	, 0x50000000	, % "ahk_id " hadm2
			WinSet, ExStyle	    	, 0x0808008C	, % "ahk_id " hadm2
			WinSet, Transparent	, 235             	, % "ahk_id " hadm2

			Gui, adm2: Show	, % (hasParent ? "x0 y0" : "x" rp.X " y" rp.Y)  " w" rp.CW " h" rp.H+rp.BW*2 " NA"   	, % "admImportLayer"

			If dbg
				SciTEOutput("ImportGui wird angezeigt!`n " thisMsg.HeaderText "`n"
							.    "Message [ mit " StrSplit(thisMsg.BodyText, "`n", "`r").Count() " Zeilen]: " thisMsg.BodyText "`n"
							. 	 "msg Parameter: " imprtmsg)

			return
		}
	;}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Importfortschritt-Gui zeigen oder beenden
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
		If (ShowGui~="i)^\s*(Show|true|1$|update)") {

			If (last_msg <> imprtmsg)
				last_msg := imprtmsg

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Position anpassen, warten falls das Infofenster noch nicht angezeigt wird
			If (WinGetTitle(hadm) != "AddendumGui")
				hafx := Controls("AfxMDIFrame1401", "hwnd", "ALBIS ahk_class OptoAppClass")

			rp := GetWindowSpot(hadm ? hadm : hafx)
			X 	:= hadm ? rp.X : rp.W - Addendum.iWin.W - 2
			W	:= hadm ? rp.CW : Addendum.iWin.W - 2
			SetWindowPos(hadm2, X, rp.Y, W, rp.H+rp.BW*2)

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				; HEADER: der Header kann unabhängig vom Nachrichtentext aufgefrischt werden
			If (ShowGui ~= "i)update.*header\W" && thisMsg.Header)
				GuiControl, adm2:, admImportHeader1       	, % thisMsg.Header

		 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			 ; HEADERTEXT: der Header kann unabhängig vom Nachrichtentext aufgefrischt werden
			If (ShowGui ~= "i)update.*header(text)*" && thisMsg.HeaderText)
				GuiControl, adm2:, admImportHeader2    	, % thisMsg.HeaderText

		 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			 ; NACHRICHT: Nachricht erneuern   ; "update message" oder "update header"
			If (ShowGui ~= "i)update.*message") {
				Edit_Append(admHImporter, thisMsg.BodyText "`n")
			}

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; Gui nach vorne holen
			WinSet, AlwaysOnTop, % "ahk_id " hadm2
			Gui, adm2: Show, NA

	}

	; Gui schliessen
		else if (ShowGui~="i)^\s*(0|false|hide|destroy|off|)\s*$")
			gosub adm2_Hide
	;}

return
adm2_Handler:     	;{

	Critical, 20
	GuiControlGet, cText, adm2:, admImportCancel
	If InStr(cText, "Vorgang abbrechen") {
		breakImportAll := true
		admGui_ImportGui("update message", last_msg "`n`r...Vorgang wird gleich beendet", last_Ctrl)
		return
	}

;}
adm2GuiClose:
adm2GuiEscape:
adm2_Hide:           	;{

	Gui, adm2: Show, Hide
	GuiControl, adm2:, admImportCancel, % "Vorgang abbrechen"
	GuiControl, adm2:, admImporter, % ""
	last_msg := last_ctrl := ""
	Addendum.Flags.ImportFromPat 	:= false
	Addendum.Flags.ImportFromJrnl	:= false

return ;}
}

admGui_ImportJournalAll(files="", Options="")                   	{               	; Journal: 	Import aller vollständig benannten Dateien

	; letzte Änderung: 28.07.2022

	; Variablen
		global admHTab, breakImportAll, admImportCancel, PatDocs, admImportStatus
		static ImptDbg      	:= false
		static waituser      	:= 6
		static NoImports   	:= 0
		static basemsg

		admImportStatus 	:= 0x1
		breakImportAll     	:= false
		Importstart           	:= A_Hour ":" A_Min ":" A_Sec
		ImportstartTC    	:= A_TickCount

	; Optionen parsen
		RegExMatch(Options, "i)(?<=FullNamedOnly\=)\d+"               	, fullnamedonly)
		RegExMatch(Options, "i)(?<=Debug\=)(?<bg>\d+|true|false)"	, d)
		basemsg   	:= "D O K U M E N T 1 M P O R T`n"
								. (fullnamedonly ? "[ nur vollständig benannte Dokumente werden importiert ]" : "[ alle zuweisbaren Dokumente werden importiert ]")
								. "`n`n#"

		If ImptDbg
			SciTEOutput("[" A_ThisFunc "] files: " files.Count())

	; Import Button inaktiv schalten
		Journal.ShowImportStatus(Addendum.Flags.ImportFromJrnl := true)
		Journal.Default()

	; das Journal wird aufsteigend nach Dokumentbezeichnung sortiert
		Journal.SortByDocNames()

	; Import Gui anzeigen
		admGui_ImportGui(true, basemsg, "admJournal")

	; Dokumente ohne Personennamen werden ignoriert
	; Dokumente ohne genaue Bezeichnung (Name und Dokumentdatum nicht)!
		rowNr := 0, JRowsStart := JRows := Journal.GetCount()

	; Abbruchbedingungen nach while ...
		while (JRows>0 && NoImports<=JRows && rowNr<=JRowsStart && !breakImportAll) {

			; Zeilentext auslesen
				rowNr ++
				Journal.SelectRow(rowNr)
				If !(rowFile := Journal.GetText(rowNr))
					continue
				fname :=RegExReplace(rowFile, "i)\.[a-z]+$")

			 ; enthält keinen Personennamen dann weiter
				I_Dont_Like_This_Import := false
				isFullNamed	:= xstring.isFullNamed(fname)
				isNamed   	:= xstring.isNamed(fname)

				If ImptDbg
					SciTEOutput("isfullnamed: " isFullNamed " , " fname)

				If (fullnamedonly && !isFullNamed) || (!fullnamedonly && !isNamed)  {
					JRows := Journal.GetCount()
					NoImports ++
					I_Dont_Like_This_Import := true
				}

			; Fortschritt anzeigen
				admGui_ImportGui("update message"	, "["   	rowNr "/" JRowsStart                  	"] "
																									. "" 	rowFile                                     	" "
																									. "" 	TimeFormatEx(Floor((A_TickCount - ImportstartTC)/1000))
																									, "admJournal")

			; weiter wenn Bedingung nicht erfüllt war
				If I_Dont_Like_This_Import
					continue

			; Abbruch durch Nutzer ermöglichen
				If (Addendum.iWin.Imports > 0) {
					MsgBox, 0x1024, % "Dokumentimport", % "Weiteres Dokument importieren?`n[" rowFile "]", % (rowNr <= 5 ? waituser : waituser/3 )
					IfMsgBox, No
					{
						Addendum.Flags.ImportFromJrnl	:= false
						breakImportAll                          	:= true
						break
					}
				}

			; Nutzer hat den Vorgang manuell abgebrochen (breakImportAll = true), die Schleife wird hier verlassen
				If breakImportAll
					break

			; Importieren (breakImportAll = false),
				filePath := Addendum.BefundOrdner "\" rowFile
				If RegExMatch(filePath, "\.pdf$") && FileExist(filePath) && !FileIsLocked(filePath) && PDFisSearchable(filepath) {

						If FuzzyKarteikarte(rowFile) {                                                        	; Karteikarte öffnen 	- erfolgreich

							Pat := xstring.Get.Names(rowFile)

								; hat der Pat. mehrere Dokumente wird admGui_ImportFromPatient() aufgerufen
							; wenn nur vollständig benannte Dokumente importertiert werden sollen dann muss PatDocs vorher bereinigt werden
							PatDocs := admGui_Reports()
							If (PatDocs.Count() > 1 && fullnamedonly) {
								tmpDocs := Array()
								For Each, PDoc in PatDocs {
									PDocName := RegExReplace(PDoc.name, "\.\w+$")
									PDocName := Trim(RegExReplace(PDocName, "i)\sv*\.*\s*[\d\-\s.]+$"))
									If (StrLen(PDocName) > 0 && !FileIsLocked(Addendum.BefundOrdner "\" PDocName))
										tmpDocs.Push(PDoc)
								}
								PatDocs := Array()
								For Each, PDoc in tmpDocs
									PatDocs.Push(PDoc)
							}

							; Pat. hat (auch nach der Bereinigung) noch mehrere Dokumente
							; es werden dann alle Dokumente des Pat. importieret. Achtung: Keine Nutzernachfrage!
							PatDocImports := PatDocs.Count()>1	? admGui_ImportFromPatient(Pat) : admGui_ImportFromJournal(rowFile)
							If PatDocImports                                                                      	; Importieren         	- erfolgreich
								Addendum.iWin.ImportsLast  := Addendum.iWin.Imports, Addendum.iWin.Imports += PatDocImports
							else                                                                                      	; Importieren - fehlgeschlagen
								NoImports += PatDocs.Count()>1	? PatDocs.Count() : 1

						}
						else {                                                                                        	; Karteikarte öffnen 	-  fehlgeschlagen
							PraxTT("Es konnte kein passender Patient gefunden werden.", "3 0")
							NoImports ++
						}

				}

			; Zeilenzahl ermitteln
				JRows := Journal.GetCount()
				If (NoImports >= JRows )
					break

		}

	; das Journal wird wieder nach Dokumenteingangsdatum sortiert
		Journal.SortByDocDates(true)

	; Statistik zeigen
		timeDiff := Floor((A_TickCount - ImportstartTCt)/1000)
		outmsg := StrReplace(basemsg, "$", "Statistik")                                                                                           	"`r`n"
					. "insgesamt:        `t" 	JRowsStart   "/" JRows                                                                               	"`r`n"
					. "importiert:         `t" 	Addendum.iWin.Imports                                                                        	"`r`n"
					. "fehlgeschlagen:`t"  	NoImports                                                                                             	"`r`n"
					. "gestartet:         `t"   	ImportStart                                                                                            	"`r`n"
					. "beendet:          `t"   	(ImportEnd := A_Hour ":" A_Min ":" A_Sec)                                                	"`r`n"
					. "benötigte Zeit:  `t"  	(needed := TimeFormatEx(timediff))                                                         	"`r`n"
					. "Durchschnitt:    `t"  	(avgTime := Round(timeDiff/Addendum.iWin.Imports, 1)) " s je Dokument"	"`r`n"
		admGui_ImportGui("update header+message", outmsg, "admJournal")

	; Statistik sichern
		stats := !FileExist(Addendum.DBPath "\ImportZeit.log") ? "Datum | Startzeit | Endzeit | benötigt | Durchschnitt | Importe | Fehlschläge" : ""
		stats .= A_DD "." A_MM "." A_YYYY " | " ImportStart " | " ImportEnd " | " needed " | " avgTime " | " Addendum.iWin.Imports " | "  NoImports  "`n"
		FileAppend, % stats, % Addendum.DBPath "\ImportZeit.log", UTF-8

	; ImportGui Schliessen Button anzeigen
		GuiControl, adm2:, admImportCancel, % "Statistik schliessen"
		fn_ImportGuiOff := Func("admGui_ImportGui").Bind(false)
		SetTimer, % fn_ImportGuiOff, -200000                                   ; ~20s

	; Zurücksetzen der Zähler
		Addendum.Flags.ImportFromPat       	:= false
		Addendum.iWin.ImportsLast            	:= -1
		Addendum.iWin.Imports                	:= 0
		admImportStatus 	                         	:= 0x0
		breakImportAll                                 	:= false
		basemsg                                        	:= ""
		NoImports := rowNr := ImportStart := ImportstartTC := 0

	; Steuerelement aktivieren, Infotexte ändern
		Journal.SelectRow(0)
		Journal.ShowImportStatus(Addendum.Flags.ImportFromJrnl := false)
		Journal.InfoText()

return
}

admGui_ImportFromJournal(filename)                              	{               	; Journal: 	Einzelimport-Funktion

	; letzte Änderung: 14.06.2022
	; cPat - muss global sein

		global 	admHPDFfilenames, hadm, admHTab, admJournal

	; Nutzer befragen ;{
		currPat 		:= AlbisCurrentPatient()
		currPatID 	:= AlbisAktuellePatID()
		Pat	    	:= xstring.Get.Names(filename)
		If Addendum.iWin.ConfirmImport || (!InStr(currPat, Pat.Nn) && !InStr(currPat, Pat.Vn)) {
			MsgBox, 262148	, % "Befund importieren?", % "Wollen Sie die Datei:`n"
																				.  (StrLen(filename) > 30 ? SubStr(filename, 1, 30) "..." : filename) "  dem Pat.`n"
																				.  currPat " zuordnen ?"
			IfMsgBox, No
			{
				Addendum.Flags.ImportFromJrnl := false
				return 0
			}
		}

		Addendum.Flags.ImportFromJrnl := true

		;}

	; Dokumentdatum aus dem Dateierstellungsdatum nehmen, ansonsten aus dem Dateinamen entnehmen (falls enthalten, sonst leer)
	; DocDate ist manchmal das Geburtsdatum des Patienten - DocDate enthält dann das Dateidatum
	; ein Dokumentdatum das vor dem ersten Behandlungstag liegt, wird ebenso verworfen
		DocDate 	:= xstring.get.DocDate(filename)
		DBDate		:= ConvertToDBASEDate(DocDate)
		If (DBDate=cPat.GEBURT(currPatID) || DBDate<cPat.SEIT(currPatID) || !RegExMatch(DocDate, "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})"))
			DocDate := GetFileTime(Addendum.BefundOrdner "\" filename, "C")

	; Befund importieren
		If      	RegExMatch(filename, "\.pdf")                                             	{

			Befunde.Backup(filename)                          	; Backup der Datei anlegen falls bisher nicht erfolgt
			If !AlbisImportierePdf(filename,, DocDate)  {  	; Importfunktion aufrufen
				Addendum.Flags.ImportFromJrnl := false
				return 0
			}

			pdf	:= pdfpool.Remove(filename)            	; aus ScanPool entfernen
			res	:= Befunde.MoveToImports(filename) 	; Backup- und Textdatei in ihre Unterverzeichnisse verschieben

			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			; automatisch aufgrund eines Dokumentitels in ein Wartezimmer setzen # funktioniert nie
			; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			If Addendum.iWin.AutoWZ
				If RegExMatch(filename, "i)(?<ger1>" Addendum.iWin.AutoWZTitles ")", trig) && RegExMatch(filename, "i)(?<ger2>" Addendum.iWin.AutoWZStates ")", trig) {
					PraxTT("{AutoWartezimmer}`nTrigger gefunden`n<" trigger1 "> und <" trigger2 ">", "4 1")
					For titles, WZ in Addendum.iWin.AutoWZAssigns
						If RegExMatch(match, "i)(?<entar>" titles ")", Komm) {
							currPat := AlbisCurrentPatient()
							geschl := AlbisPatientGeschlecht()
							MsgBox, 0x1004, Auto Wartezimmer, % "Für " currPat " ist eine Anfrage von/vom: " Kommentar " eingetroffen.`n"
																					.	"Wollen Sie " (geschl="m" ? "den Patienten" : geschl="w" ? "die Patientin" : "den Fall")
																					. " ins Wartezimmer ('" WZ "') setzen?"
							IfMsgBox, Yes
							{
								admGui_PatInWZ(currPat, WZ, "Kommentar", WZKommentar)
								If !InStr(WZKommentar, Kommentar) {
									AlbisWZKommentar(WZ, (WZKommentar ? ", " : "") Kommentar, "Abwesend")
									; ein Pat. der schon im Wartezimmer ist, wird bei erneutem Setzen eines Wartezimmerkommentars ins Wartezimmer "Abgemeldet" verlegt
									; den Abgemeldet Pat. wieder ins ursprüngliche Wartezimmer zu legen wird hier später erfolgen, noch habe ich keine Idee wie ich das realisiere
									;~ If !admGui_PatInWZ(currPat, WZ, "Kommentar", WZKommentar)
										;~ If admGui_PatInWZ(currPat, "Abgemeldet", "Kommentar", WZKommentar)
								}
								MsgBox, 0x1004, Auto Wartezimmer, % "Hat es funktioniert?"
								IfMsgBox, No
									PraxTT("Ach verdammt!", "2 4")
							}

							break
						}

				}

		}
		else If	RegExMatch(filename, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{
			If !AlbisImportiereBild(filename, xstring.Replace.Names(filename), DocDate) {
				Addendum.Flags.ImportFromJrnl := false
				return 0
			}
			res := Befunde.MoveToImports(filename)
		}
		else {
			Addendum.Flags.ImportFromJrnl := false
			return 0
		}


	; Re-Indizieren, Gui auffrischen, Importprotokoll
		If !FilePathCreate(Addendum.DBPath "\logs")
			PraxTT("Protokolldateipfad: `n" Addendum.DBPath "\logs`nkonnte nicht angelegt werden.", "1 1" )
		else
			FileAppend, % datestamp() "| " filename "   >>>   [" currPatID "] " currPat "`n", % Addendum.LogPath "\PdfImportLog.txt"

	; Anzeigen und interne Objekte bearbeiten
		Addendum.Flags.ImportFromJrnl := false
		result	:= 	Journal.Remove(filename)             	; die Listview wird von diesem Eintrag befreit
		PatDocs:= 	admGui_Reports()
		result	:= 	Journal.Show()
											Journal.InfoText()

return 1
}

admGui_ImportFromPatient(Pat:="")                               	{               	; Patient:	Befundimport alle Befunde

	; letzte Änderung: 12.05.2022

		global 	PatDocs

		Addendum.Flags.ImportFromPat := true

		Imports        	:= 0
		Docs         	:= Object()
		currPatID  	:= AlbisAktuellePatID()
		currPatBirth	:= AlbisPatientGeburtsdatum()

	; Importschleife
		For key, file in PatDocs		{

			Addendum.FuncCallback := ""
			Befunde.Backup(file.name)                                          	; Backup der Datei anlegen falls bisher nicht erfolgt

		; Abbruch wenn die falsche Karteikarte geöffnet ist (z.b. hat der User während des Imports eine andere Karteikarte ausgewählt)
			currPat 	:= AlbisCurrentPatient()
			If !InStr(currPat, Pat.Nn) && !InStr(currPat, Pat.Vn)
				return Imports

			KKText    	:= xstring.Karteikartentext(file.name)
			DocDate 	:= xstring.Get.DocDate(file.name)

			If (DocDate = currPatBirth || !RegExMatch(DocDate, "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})"))
				DocDate := GetFileTime(Addendum.BefundOrdner "\" file.name, "C")

		; Bild importieren
			If RegExMatch(file.name, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{

				If !AlbisImportiereBild(file.name, KKText, DocDate) {
					Docs.Push(file)
					continue
				}
				else
					Imports ++

				res := Befunde.MoveToImports(file.name)     	; auch Bilddateien erhalten ein Backup
			}
		; Pdf importieren
			else {

				; sind ohne .pdf ?
					Addendum.FuncCallback := "GetPDFViewerArg"
					If !AlbisImportierePdf(file.name, KKText, DocDate) {
						Docs.Push(file)
						continue
					}
					else
						Imports ++

				; Logfile schreiben. Backup-PDF-Datei und PDF-Textdatei verschieben
					res	:= Befunde.MoveToImports(file.name)

			}

			Journal.RemoveFromAll(file.name)   ; aus Journal und Patienten Listview entfernen
			If !FilePathCreate(Addendum.DBPath "\logs")
				PraxTT("Protokolldateipfad: `n" Addendum.DBPath "\logs`nkonnte nicht angelegt werden.", "1 1" )
			else
				FileAppend, % datestamp() "| " filename "   >>>   [" currPatID "] " currPat "`n", % Addendum.LogPath "\PdfImportLog.txt"

		}

	; PatDocs  leeren und mit nicht verarbeiteten Dateien füllen
		Loop % PatDocs.MaxIndex()
			PatDocs.RemoveAt(1)
		For index, file in Docs
			PatDocs.Push(file)

	; Flag aus
		Addendum.Flags.ImportFromPat := false

return Imports
}

admGui_Export(files)                                             	{                	; Kontextmenu: Exportieren von Dateien in den Exportordner

	; auch für die Verarbeitung von Dateistapeln
		If !IsObject(files) {
			tmpfile 	:= files
			files  	:= []
			files.Push(tmpfile)
		}

	; Dateistapel abarbeiten
		For fidx, file in files {

			; aus dem Dokumentnamen den Patienten ermitteln
			Pat		:= xstring.Get.Names(file)
			PatIDs 	:= admDB.StringSimilarityID(Pat.Nn, Pat.Vn)
			If (PatIDs.Count() <> 1)
				Pat := "", PatIDs := ""

			; Datei kopieren und Ergebnis anzeigen
			res 		:= Befunde.CopyToExports(file, PatIDs[1], Pat.Nn, Pat.Vn)
			If !IsObject(res) {
				PraxTT("Anlegen des Exportpfades `n⋘ " res.PatPath " ⋙`n" (PatID ? " für (" PatID ") " Name ", " Vorname : "Exportpfad")  " ist fehlgeschlagen.`nDie Datei konnte nicht kopiert werden.", "3 1")
			} else {
				If res.success
					PraxTT("Datei: " file "`nnach " res.PatPath " kopiert", "3 1")
				else
					PraxTT("Datei: " file "`nFehler beim kopieren [" res.LastError "]", "3 1")
			}
			Sleep 2000

		}

}

admGui_PatInWZ(PatName, WZ, retParam, ByRef retVal)             	{               	; sucht nach einem Namen im angegebenen Wartezimmer

	; admGui_PatInWZ("Mustermann, Martin", "Arzt", "Kommentar", WZKommentar)

	WZTable := AlbisWZListe(WZ)
	For rowNR, col in WZTable
		If InStr(col.Name, currPat) {
			retVal := Trim(col[retParam])
			return rowNR
		}

return 0
}
;}

; -------- Karteikarte                                                                    	;{
FuzzyKarteikarte(NameStr)                                        	{               	; fuzzy name matching function, öffnet eine Karteikarte

	; letzte Änderung 12.05.2022
	; cPat - ist ein globales Klassenobjekt, wird in Addendum.ahk erstellt

	; prüft ob Namen übergeben wurden
		If !IsObject(Pat := xstring.Get.Names(NameStr)) {
			PraxTT(	"Die Dateibezeichnung enthält keinen Namen eines Patienten.`nDer Karteikartenaufruf wird abgebrochen!", "3 0")
			return 0
		}

	; passende Patienten suchen
		mPats := cPat.StringSimilarityEx(Pat.Nn, Pat.Vn)
		;~ m := mPats.diff
		m := mPats.Delete("diff")

	; Karteikartenfunktion aufrufen
		If (mPats.Count() = 1) {

			For PatID, Patient in mPats
				return admGui_Karteikarte(PatID)

		}
		else If (mPats.Count() > 1) {

			index := 1
			For PatID, Patient in mPats {

				NNVN  	:= RegExReplace(cPat.Get(PatID, "NAME") cPat.Get(PatID, "VORNAME"), "[\s-]")
				VNNN  	:= RegExReplace(cPat.Get(PatID, "VORNAME") cPat.Get(PatID, "NAME"), "[\s-]")
				PatNNVN	:= Pat.Nn Pat.Vn
				If (NNVN=PatNNVN || VNNN=PatNNVN)
					return admGui_Karteikarte(PatID)
				else
					t .= (index++) ": [" PatID "] " cPat.Name(PatID, true) (index=1 ? "":"`n")
				If (index = 9) {
					If (mPats.Count() > index)
						t .= "......"
					break
				}

			}

			PraxTT("Karteikarte öffnen: Es wurden mehrere Namen gefunden:`n" t, "8 0")
			return 0

		}
		else {

			PraxTT("Karteikarte öffnen: Es wurde kein passender Name gefunde.", "3 0")
			return 0

		}


return 1
}

admGui_Karteikarte(PatID)                                        	{               	; fragt ob eine Karteikarte geöffnet werden soll

	; cPat - ist ein globales Klassenobjekt, wird in Addendum.ahk erstellt

	If !Addendum.PDF.PatAkteSofortOeffnen {
		MsgBox, 0x1004, % "Patientenakte öffnen ?"	, % "Möchte Sie die Akte des Patienten:  `n"
																			. 	  "(" PatID ") " cPat.Name(PatID, true) "`n"
																			.	  "öffnen?"
		IfMsgBox, No
			return 0
	}

return AlbisAkteOeffnen(cPat.Name(PatID, false), PatID)
}
;}

; -------- Gui Einstellungen speichern                                                     	;{
admGui_PatLog(PatID, cmd, LogText)                             	{                 	; individuelles Patienten Logbuch

	; zurück falls Nutzer keine Protokollerstellung wünscht
		If !Addendum.PatLog
			return

		PatientDir 	:= PatDir(PatID)
		If InStr(cmd, "AddToLog") {
			timeStamp := A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min ":" A_Sec "`t"
			FileAppend, % timeStamp . LogText "`n", % PatientDir "\log.txt"
		}

}

admGui_SaveStatus()                                                      	{               	; Anzeigestatus des Infofenster sichern

	; letzte Änderung: 06.11.2021

		global admHJournal, adm, admTPTag, admTabs, admLBAnm, admLBDrucker, admTPClient

	; Gui auslesen
		Gui, adm: Submit, NoHide
		JournalPos := LV_GetScrollViewPos(admHJournal)

	; firstTab
		firstTab       	:= IniReadExt(compname, "Infofenster_aktuelles_Tab")
		thisTab      	:= Addendum.iWin.firstTab
		If (firstTab <> thisTab)
			IniWrite, % thisTab, % Addendum.Ini, % compname, % "Infofenster_aktuelles_Tab"

	; angezeigtes Protokoll
		iniTProtDate	:= IniReadExt(compname, "Infofenster_Tagesprotokoll_Datum")
		TProtDate  	:= Addendum.iWin.TProtDate
		If (iniTProtDate <> TProtDate)
			IniWrite, % TProtDate, % Addendum.Ini, % compname, % "Infofenster_Tagesprotokoll_Datum"

	; für welchen Client
		iniTPClient 	:= IniReadExt(compname, "Infofenster_Tagesprotokoll_Client")
		TPClient     	:= Addendum.iWin.TPClient
		If (iniTPClient <> TPClient)
			IniWrite, % TPClient, % Addendum.Ini, % compname, % "Infofenster_Tagesprotokoll_Client"

	; Drucker
		iniLBDrucker	:= IniReadExt(compname, "Infofenster_Laborblatt_Drucker")
		LBDrucker 	:= Addendum.iWin.LBDrucker
		If (iniLBDrucker <> LBDrucker)
			IniWrite, % LBDrucker, % Addendum.Ini, % compname, % "Infofenster_Laborblatt_Drucker"

	; Ausdruck mit Anmerkungen/Probedaten
		iniLBAnm  	:= IniReadExt(compname, "Infofenster_Laborblatt_Anmerkung_drucken")
		LBAnm      	:= Addendum.iWin.LBAnm
		If (iniLBAnm <> LBAnm)
			IniWrite, % (LBAnm ? "Ja" : "Nein")	, % Addendum.Ini, % compname, % "Infofenster_Laborblatt_Anmerkung_drucken"

		; JournalPosition
		If JournalPos 	{
			If (Addendum.iWin.JournalPos  	<> JournalPos) 	{
				IniWrite, % JournalPos, % Addendum.Ini, % compname, % "Infofenster_JournalPosition"
				Addendum.iWin.JournalPos	:= JournalPos
			}
		}


}
;}

; -------- Callback Funktionen                                                             	;{
admGui_CheckJournal(pdffile, ThreadID)                          	{                	; Hilfsfunktion für tesseract OCR und Gui Anzeige

		; Ausführung jedesmal nach erfolgreicher Erstellung einer PDF durch TesseractOCR()
		; Aufgaben: 	1. Start der Autonaming Funktionen
		;               	2. ScanPool-Objekt auffrischen
		;					nach Abschluß den OCR Vorgang per IPC-Nachricht fortsetzen

		global 	admHJournal
		static	 	AutonamingRunning := false

	; Autonaming starten
		;~ SciTEOutput("  - AutoNaming: gestartet")
		AutonamingRunning	:= true
		newfilenames         	:= admCM_JRecogAll([pdfFile])
		AutonamingRunning	:= false
		;~ SciTEOutput("  - AutoNaming: beendet")

	; Datei aus dem Stapel entfernen
		OCRStack := Addendum.iWin.FilesStack
		Addendum.iWin.FilesStack := []
		For idx, inprogress in OCRStack {

			FileIsProcessed := false
			For newNr, newfilename in newfilenames {
				If (inprogress = pdfFile) {
					FileIsProcessed := true
					break
				}
			}

			If !FileIsProcessed
				Addendum.iWin.FilesStack.InsertAt(1, inprogress)

		}

	; ScanPool-Objekt updaten
		If !IsObject(newfilenames) || (newfilenames.MaxIndex() = 0)
			oPDF := PDFpool.Add(Addendum.BefundOrdner, pdfFile)

	; Nachricht an den OCR-Thread: OCR Vorgang fortsetzen
		Send_WM_CopyData("continue", ThreadID)

}

admGui_FolderWatch(path, changes)                              	{                	; überwacht den Befundordner

	/* Beschreibung

		- 	überwacht den Befundordner auf Dateiveränderungen
		- 	bei neuen PDF-Dateien ohne Textlayer wird automatisch eine Texterkennung durchgeführt
		- 	jede Dateiveränderung wird sofort im Infofenster angezeigt

		- 	Infofenster sowie WatchFolder können für sich allein oder gemeinsam ausgeführt werden,
			dementsprechend muss sich diese Funktion an die Nutzereinstellungen anpassen und insbesondere
			darf es keine zeitgleichen Zugriffe durch Funktionen des Infofensters auf dasselbe Listview-Element geben

		letzte Änderung: 	24.01.2022 (bei FWPause und Treffer in der Dateiliste wird die For-Schleife fortgesetzt und nicht abgebrochen)
									## - es fehlt die Verarbeitung von Bilddateien

	 */

		global hadm
		static func_AutoOCR	:= func("admGui_OCRAllFiles"), OSDTip := false
		static WFActions := ["ADDED", "REMOVED", "MODIFIED", "RENAMED", "NEWNAME"]
		static last_actiontext

	; untersucht alle veränderten oder hinzugfügten Dateien
		For Each, Change In Changes 	{

			action	:= change.action
			name	:= change.name

		; Abbruch wenn eine Infofenster Funktion Dateien bearbeitet
			If (Addendum.OCR.FWPause || Addendum.OCR.FWIgnore.Count() > 0)
				For iFIdx, ignoreFile in Addendum.OCR.FWIgnore
					If (name = ignoreFile)
						continue

		; Protokoll führen (Ausschluß der Erfassung von Dateien mit bestimmten Endungen)
			If (Addendum.OCR.WFLog && !RegExMatch(name, "i)\.(pd~|dat|json)$")) {
				actiontext := WFActions[action] " | " name
				If (last_actiontext <> actiontext)
					FileAppend,	% A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min " | " (last_actiontext := actiontext) "`n"
									, 	% Addendum.DBPath "\logs\WatchFolder-Log.txt", UTF-8
			}

		; PDF Dateien
			If RegExMatch(name, "i)\.pdf$")
				If RegExMatch(action, "1|3|4|5")	{

					SplitPath, name, filename, filepath
					If IsObject(oPDF := PDFpool.Add(filepath, filename)) {

						; neue Datei zu Listview hinzufügen und Info's ändern
							If admGui_Exist()
								Journal.Add(oPDF.name)

						; Dateizähler erhöhen
							Addendum.OCR.staticFileCount 	++
							Addendum.OCR.filecount           	++

						; Texterkennung: ausstehend? Dann wird ein Thread* zeitlich verzögert gestartet.
						;                          	(*wird nur auf dem Client ausgeführt der in der Addendum.ini hinterlegt ist)
							If (Addendum.OCR.AutoOCR && Addendum.OCR.Client=compname)
								If !oPDF.isSearchable
									SetTimer, % func_AutoOCR, % "-" (Addendum.OCR.AutoOCRDelay*1000)

					}

				}
			; Datei wurde gelöscht
				else If RegExMatch(action, "2")	{

					SplitPath, name, filename, filepath
					oPDF := PDFpool.Remove(filename)

					; wenn Infofenster sichtbar
					If admGui_Exist()
						If !(delRow := Journal.DeleteRow(filename)) && Addendum.iWin.Debug
							SciTEOutput("> " A_ThisFunc ": delRow = " delRow ", pdf: " filename)

				}

		}

		If Addendum.OCR.filecount   	{
			msg := Addendum.OCR.filecount " Befund" (Addendum.OCR.filecount = 1 ? " wurde" : "e wurden" ) " hinzufügt!"
			Controls("", "reset", "")
			If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {
				admPos := GetWindowSpot(hadm)
				Splash(msg, 6, admPos.X+30, admPos.Y + admPos.H - 30, admPos.W - 60, hadm)
			}
			else
				TrayTip("Befundordner", msg, 6)

		}

		Addendum.OCR.filecount := 0

return
}

admGui_FWPause(FWPause, files, PTime=8)                  	{               	; FolderWatch für bestimmte Dateien anhalten/fortsetzen

	; letzte Änderung: 24.01.2022

	; WatchFolder ist nicht eingeschaltet. Ausführung ist nicht notwendig.
		If !Addendum.OCR.WatchFolder
			return

	; Resettimer verlängern
		If Addendum.OCR.FWPause && FWPause {
			SetTimer, FWPauseOff, % (-1*PTime*1000)
			return
		}

	; FW pausieren:
		If FWPause {

		; es wird ein Array angelegt (egal ob 2 oder mehr oder nur eine Datei übergeben wurde)
			Addendum.OCR.FWIgnore := IsObject(files) ? files :  StrLen(files) > 0 ? [files] : ""

		; Timer nur starten wenn Dateinamen übergeben wurden
			If (Addendum.OCR.FWIgnore.Count() > 0) {
				Addendum.OCR.FWPause := true
				SetTimer, FWPauseOff, % -1*PTime*1000
				return
			}

		}

FWPauseOff:
	Addendum.OCR.FWIgnore	:= ""
	Addendum.OCR.FWPause 	:= false
return
}
;}

; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
; -------- Texterkennung                                                                  	;{
TesseractOCR(files)                                                         	{               	; erstellt einen zusätzlichen echten Thread

	 /*  Funktionsweise


	#	Erstellung des Threads:

		Ein Teil des Threads wird in Addendum.ini in die Variable >Addendum.Threads.OCR< geladen (Addendum_OCR.ahk)
		Die Funktion TesseractOCR() verbindet Skript und die Tesseract/OCR-Einstellungen zu einem lauffähigen Skript.
		Anschließend wird das zusammengefügte Skript mit dem AutohotkeyH Befehl ''AHKThread(script)'' ausgeführt.

	#	Kommunikation zwischen Thread und Skript:

		Es werden Nachrichten zwischen dem Thread und dem aufrufenden Skript ausgetauscht.
		Der Thread meldet die Fertigstellung einer Datei und wartet auf eine Nachricht des Skripts, bevor er fortfährt.
		Dadurch kann das Skript seine eigenen Operationen ausführen und beenden ohne durch die Fertigstellung
		einer weiteren Texterkennung unterbrochen zu werden.
		- die Kommunikationsfunktionen des Skriptes finden sich in Addendum.ahk und die des Thread in Addendum_OCR.ahk

	# genutzte Kommandozeilenprogramme

		convert.exe 		- imagemagick CLI Tool in \include\OCR\imagick
		gswin64c.exe	- Ghostscript 9x+ wird von convert.exe benötigt, muss dafür leider installiert sein

	*/

	global Settings, autoexecScript

	If (Addendum.OCR.Client <> compname) || (StrLen(Addendum.OCR.Client) = 0)
		return

	; Konfiguration und Skript;{
		tessconfig1 =
					(LTrim Join|
							tessedit_create_boxfile                    	1
							tessedit_create_hocr                      	1
							tessedit_create_pdf                        	1
							tessedit_create_tsv                           	1
							tessedit_create_txt                          	1

							load_bigram_dawg	                    	True
							tessedit_enable_bigram_correction		True
							tessedit_bigram_debug	                  	3
							save_raw_choices	                         	True
							save_alt_choices	                         	True
					)

		tessconfig2 =
					(LTrim Join|
							tessedit_create_pdf                        	1
							tessedit_create_tsv                           	1
							tessedit_create_txt                           	1
					)

		autoexecScript =
		(
		#NoTrayIcon
		#Persistent
		SetBatchLines, -1
		DetectHiddenWindows On
		SetTitleMatchMode 2

		global hOCR
		global threadcontrol               	; global flag - Thread externe Funktion (Addendum.ahk)
		global q := Chr(0x22)              	; q := "

		###files
		###settings

		Gui, OCR: New, +hwndhOCR +ToolWindow
		Gui, OCR: Show, Hide, Addendum tessOCR processor
		OnMessage(0x4a, "Receive_WM_COPYDATA")

		If isObject(files) {
			filesmax := files.MaxIndex()
			For idx, pdffile in files {
				threadcontrol := true
				SciTEOutput("``n> Bearbeite " idx "/" filesmax ": [" pdffile "]")
				pdfText := tessOCRPdf(pdffile, settings)
				Send_WM_CopyData("OCR_processed|" pdffile "|" hOCR, settings.ScriptID)
				While threadcontrol
					sleep, 50
			}
		}
		else {
			filesmax := 1
			threadcontrol := true
			SciTEOutput("``n> Bearbeite " 1 "/" 1 ": [" pdffile "]")
			pdfText := tessOCRPdf(files, settings)
			Send_WM_CopyData("OCR_processed|" files "|" hOCR, settings.ScriptID)
			While threadcontrol
				sleep, 50
		}

		Send_WM_CopyData("OCR_ready|" filesmax "|" hOCR, settings.ScriptID)
		ExitApp

		)

		Settings := {	"SetName"                             	: "threadOCR"
								,	"UseRamDisk"                        	: "T:"
								,	"imagickPath"                           	: Addendum.PDF.imagickPath
								,	"qpdfPath"                             	: Addendum.PDF.qpdfPath
								,	"tessPath"                               	: Addendum.PDF.tessPath
								,	"xpdfPath"            	                 	: Addendum.PDF.xpdfPath
								,	"documentPath"                        	: Addendum.BefundOrdner
								,	"backupPath"                	         	: Addendum.BefundOrdner "\backup"
								,	"txtOCRPath"        	                 	: Addendum.BefundOrdner "\Text"
								,	"OCRLogPath"      	                  	: Addendum.LogPath "\OCRTime_Log.txt"
								,	"tessconfig"                            	: tessconfig2
								,	"tessCmdPlus"                           	: "--psm 1 --oem 2"
								,	"useData"                               	: "best"
								,	"uselang"                               	: "deu"
								,	"preprocessing"                       	: "imagemagick"
								,	"convertwith"                           	: "pdfimages"
								,	"debug"                                  	: "1"
								,	"writelog"             	                 	: "true"
								,	"backupfiles"                            	: "true"
								,	"callbackFunc"                          	: "OCRReady"
								,	"ProcessID"                             	: DllCall("GetCurrentProcessId")                      	; nicht ändern!
								,	"ScriptID"                                	: (Addendum.MsgGui.hMsgGui)}                    	; nicht ändern!

	;}

	; schreibe das/die Objekt/e als Skriptcode
		If IsObject(files) {
			repfiles := "files := ["
			For key, val in files
				repfiles .= q val q ","
			repfiles := RTrim(Trim(repfiles), ",") "]`n"
		}
		else
			repfiles := "files := " q files q

	; stellt die Einstellungen zusammen
		repset := "settings := {"
		For key, val in Settings
			repset .= q key q ":" . ((val = "true") || (val = "false") || RegExMatch(val, "^(\d+|0x\w+)$") ? val : q val q) ", "
		repset := RTrim(Trim(repset), ",") "}`n"

	; Skript zusammensetzen
		autoexec	:= StrReplace(autoexecScript	, "###files"   	, repfiles)
		autoexec	:= StrReplace(autoexec       	, "###settings"	, repset)
		script := autoexec "`n" Addendum.Threads.OCR

	; OCR-Thread starten
		If !Addendum.Thread["tessOCR"].ahkReady() {
			journal.OCR("+OCR abbrechen")
			Addendum.Thread["tessOCR"] := AHKThread(script)
			Addendum.tessOCRRunning := true
		}

}

admGui_OCRAllFiles()                                                      	{               	; Texterkennung aller unbearbeiteten PDF-Dateien

		static noOCR

	; Gui Button ändern
		If Addendum.OCR.AutoOCR && (Addendum.OCR.Client = compname) && Addendum.Thread["tessOCR"].ahkReady() {
			Addendum.OCR.RestartAutoOCR := true
			journal.OCR("+OCR abbrechen")
			return
		}

	; Dateinamen ohne Texterkennung ermitteln. Dateinamen in FileStack ablegen
		noOCR := pdfpool.GetNoOCRFiles()
		If noOCR.Count() {

			For key, filename in noOCR {
				Instack := false
				For fstack, fstackname in Addendum.iWin.FilesStack
					If (filename = fstackname) {
						Instack := true
						break
					}
				If !Instack
					Addendum.iWin.FilesStack.InsertAt(1, filename)
			}

			journal.OCR("+OCR abbrechen")
			Addendum.OCR.RestartAutoOCR := false
			TesseractOCR(noOCR)

		}
		else
			journal.OCR("-OCR ausführen")

return
}
;}

; -------- Hilfsfunktionen                                                                 	;{
ClientsOnline2()                                                            	{               	;-- gibt einen Array mit den Namen der Netzwerkgeräte zurück

	For clientName, prop in Addendum.LAN.Clients {



	}


return
}

ClientsOnline()                                                              	{               	;-- gibt einen Array mit den Namen der Netzwerkgeräte zurück

	; der Shell Befehl 'net view' gibt eine Liste aller vorhandenen (eingeschalteten) Netzwerkgeräte im LAN aus
	; die Liste wird nach Gerätenamen durchsucht und mit den in der Addendum.ini vorhandenen Netzwerkgeräten verglichen
	; nur die in der ini gelisteten Geräte werden von der Funktion zurückgegeben

		clients := Array()
		c := StdoutToVar("net view") ; shell Befehl ausführen

		Loop, Parse, c, `n ,`r
			If RegExMatch(A_LoopField, "^\\\\(?<Client>.*?\s|$)", Lan) {

				LanClient := Trim(StrReplace(LanClient, "-"))

				For clientName, prop in Addendum.LAN.Clients
					If RegExMatch(clientName, "i)^" LanClient "$") {
						clients.Push(clientName)
						break
					}

			}

return clients
}

GetLastGVU(PatID)                                                        	{

	static GVFile
	static LGVULenDataset := 6+8

	If (StrLen(GVFile) = 0)
		GVFile := Addendum.DBPath "\lastGVU.db"

	If !FileExist(GVFile) {

		MsgBox, % "File lastGVU.db not found`n" GVFile
		return
	}
	VarSetCapacity(recordbuf, LGVULenDataset, 0x20)

	lgvu := FileOpen(GVFile, "r", "UTF-8")
	lgvu.Seek(((PatID-1)*LGVULenDataset) + 1, 0)
	bytes	:= lgvu.RawRead(recordbuf, LGVULenDataset)
	set 	:= StrGet(&recordbuf, LGVULenDataset, "UTF-8")
	LGVUPatID	:= LTrim(SubStr(set, 1, 6), "0")
	LGVUDate	:= SubStr(Set, 7, 8)

return set " [" LGVUPatID "] . [" LGVUDate "]"
}

MsgBox_XY(Message, X:=0, Y:=0)                                 	{

	global MsgBox_X, MsgBox_Y	                                                                            	; Make global for other function

	MouseGetPos, mx, my
	MsgBox_X := !X ? mx : X
	MsgBox_Y := !Y ? my : Y

	OnMessage(0x44, "Move_MsgBox")	                                                                	; When a dialog appears run Move_MsgBox

}

Move_MsgBox(wParam)                                                 	{

	global MsgBox_X, MsgBox_Y
	global hadmRN

	if (wParam = 1027) {					                                                                    	; Make sure it is a AHK dialog

		Process, Exist	                                                                                                	; Get the Process PID into ErrorLevel
		mbxPID := ErrorLevel
		DetectHiddenWindows, % (ADHWs := A_DetectHiddenWindows) ? "On" : "On"	; Round about way of changing DetectHiddenWindows settting
																																; and saving the setting all in one line
		if (hwnd := WinExist("ahk_class #32770 ahk_pid " mbxPID)) {                          	; Make sure dialog still exist
			mbx      	:= GetWindowSpot(hwnd)
			rn          	:= GetWindowSpot(hadmRN)
			MsgBox_Y	:= rn.Y + (rn.H//2-mbx.H//2)
			MsgBox_X	:= rn.X + (rn.W//2-mbx.W//2)
			SetWindowPos(hwnd, MsgBox_X, MsgBox_Y, mbx.W, mbx.H)                      	; Move the default window from above WinExist
		}

		DetectHiddenWindows, % ADHWs
		OnMessage(0x44, "")	                                                                             			; switch off OnMessage for MsgBoxes

	}

}

SDbg()                                                                          	{
return " " A_Hour ":" A_Min ":" A_Sec " [" A_MSec  "] | "
}

;}

; -------- Debugfunktionen                                                                 	;{
ParseStdOut(stdOut, indent="    `t")                               	{

	For i, line in StrSplit(stdOut, "`n", "`r")
		t .= StrLen(line) > 0 ? indent . line "`n" : ""

return RTrim(t, "`n")
}
;}

; -------- Gui / Fenster Funktionen                                                       	;{
CalcIdealWidthEx(hLB, Content="", Delim="|"
	, FontOptions="", FontName="", rows=0)                   	{              	;-- berechnet die Anzeigebreite eines Textes in Pixel

		DestroyGui := MaxW := 0
		static SM_CVXSCROLL

		If !SM_CVXSCROLL
			SysGet, SM_CVXSCROLL, 2                                        	; width of a vertical scrollbar

		If !hLB	{
			If (StrLen(Content) = 0)
				Return -1

				Gui, LB_EX_CalcContentWidthGui: New	, % "+Delimiter" Delim
				Gui, LB_EX_CalcContentWidthGui: Font	, % FontOptions, % FontName
				Gui, LB_EX_CalcContentWidthGui: Add	, ListBox, % "+HWNDhLB " (rows = 0 ? "" : "r" rows), % Content
				DestroyGui := True

		}

		ControlGet, Content, List,,, % "ahk_id " hLB                        ; content ermitteln
		Items := StrSplit(Content, "`n")

		SendMessage, 0x31, 0, 0,, % "ahk_id " hLB                     	; WM_GETFONT
		hFont	:= ErrorLevel

		hDC  	:= DllCall("User32.dll\GetDC", "Ptr", hLB, "UPtr")
		DllCall("Gdi32.dll\SelectObject", "Ptr", hDC, "Ptr", hFont)

		VarSetCapacity(SIZE, 8, 0)
		For Each, Item In Items	{
			DllCall("Gdi32.dll\GetTextExtentPoint32", "Ptr", HDC, "Ptr", &Item, "Int", StrLen(Item), "UIntP", Width)
			MaxW := Width > MaxW ? Width : MaxW
		}

		DllCall("User32.dll\ReleaseDC", "Ptr", hLB, "Ptr", hDC)

	; einfachste Umsetzung um einfach nur herauszufinden ob eine vertikale Scrollbar existiert
		SBVERT_Exist := (Items.MaxIndex() > rows) ? true : false

		If (DestroyGui)
			Gui, LB_EX_CalcContentWidthGui: Destroy

Return MaxW + (SBVERT_Exist = true ? SM_CVXSCROLL : 0) + 8 ; + 8 for the margins
}

GuiControlActive(CtrlName, GuiName)                               	{	;-- bestimmt ob ein GuiControl den Eingabefocus hat

	Gui, %GuiName%: Default
	GuiControlGet, fCtrl, FocusV ;% "ahk_id " hWin
	If (fCtrl = CtrlName)
		return true

return false
}

CB_GetCount(hwnd)                                                  	{	;-- Comboboxeinträge zählen
	SendMessage, 0x146, 0, 0,, % "ahk_id " hwnd ; CB_GETCOUNT:= 0x0146
return ErrorLevel
}

EM_SetMargins(Hwnd, Left := "", Right := "")                      	{
	 ; EM_SETMARGINS = 0x00D3 -> http://msdn.microsoft.com/en-us/library/bb761649(v=vs.85).aspx
	 Set := 0 + (Left <> "") + ((Right <> "") * 2)
	 Margins := (Left <> "" ? Left & 0xFFFF : 0) + (Right <> "" ? (Right & 0xFFFF) << 16 : 0)
	 Return DllCall("User32.dll\SendMessage", "Ptr", HWND, "UInt", 0x00D3, "Ptr", Set, "Ptr", Margins, "Ptr")
}

RedrawWindow(hwnd=0)                                              	{	;-- zeichnet eine Autohotkey Gui komplett neu

	global hadm

	static RDW_INVALIDATE 	:= 0x0001
	static	RDW_ERASE           	:= 0x0004
	static RDW_FRAME          	:= 0x0400
	static RDW_ALLCHILDREN	:= 0x0080

return DllCall("RedrawWindow", "Ptr", (hwnd = 0 ? hadm : hwnd), "Ptr", 0, "Ptr", 0, "UInt", RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN)
}

DropShadow(HGUI:="", Style:="", GetGuiClassStyle:=""
, SetGuiClassStyle:="")                                                   	{

	;================================================================================
	; https://www.autohotkey.com/boards/viewtopic.php?f=6&p=264108#p264108
	;--------------------------------------------------------------------------------

	if (GetGuiClassStyle) {
		ClassStyle:=GetGuiClassStyle()
		return ClassStyle
	}

	if (SetGuiClassStyle) {
		SetGuiClassStyle(HGUI, Style)
	}

}

GetGuiClassStyle()                                                 	{
	Gui, GetGuiClassStyleGUI:Add, Text
	Module := DllCall("GetModuleHandle", "Ptr", 0, "UPtr")
	VarSetCapacity(WNDCLASS, A_PtrSize * 10, 0)
	ClassStyle := DllCall("GetClassInfo", "Ptr", Module, "Str", "AutoHotkeyGUI", "Ptr", &WNDCLASS, "UInt")
								 ? NumGet(WNDCLASS, "Int")
								 : ""
	Gui, GetGuiClassStyleGUI:Destroy
	Return ClassStyle
}

SetGuiClassStyle(HGUI, Style)                                      	{
	Return DllCall("SetClassLong" . (A_PtrSize = 8 ? "Ptr" : ""), "Ptr", HGUI, "Int", -26, "Ptr", Style, "UInt")
}

AccObj_FromPoint(X := "", Y := "")                                 	{

	; AccessibleObjectFromPoint()
	; Retrieves the address of the IAccessible interface pointer for the object displayed at a specified point on the screen.
	; --------------------------------------------------------------------------------------------------------------------------------

	If (X = "") || (Y = "")
		DllCall("GetCursorPos", "Int64P", PT)
	Else
		PT := (X & 0xFFFFFFFF) | (Y << 32)
	VarSetCapacity(CID, 24, 0)
	HR := DllCall("Oleacc.dll\AccessibleObjectFromPoint", "Int64", PT, "PtrP", IAP, "Ptr", &CID, "UInt")

Return (HR = 0 ? New AccObj_Object(IAP, NumGet(CID, 8, "Int")) : False)
}
;================================================================================
;}

; -------- IAutoComplete                                                                   	;{
IAutoComplete_Create(HEDT, Strings, Options := "", WantReturn := False, Enable := True) {	;-- aus dem Autohotkey Forum
	 Return New IAutoComplete(HEDT, Strings, Options, WantReturn, Enable)
}

Class IAutoComplete {
	 Static Attached := []
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 ; Constructor - see AutoComplete_Create()
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 __New(HEDT, Strings, Options := "", WantReturn := False, Enable := True) {

			Static IAC2_Init := A_PtrSize * 3
			If AutoComplete.Attached[HEDT]
			 Return ""
			This.HWND := HEDT
			This.SubclassProc := 0
			If !(IAC2 := ComObjCreate(	"{00BB2763-6A77-11D0-A535-00C04FD7D062}"
													, 	"{EAC04BC0-3791-11d2-BB95-0060977B464C}"))
			 Return ""
			This.IAC2 := IAC2
			If !(IES := IEnumString_Create())
			 Return ""
			If !IEnumString_SetStrings(IES, Strings) {
			 DllCall("GlobalFree", "Ptr", IES)
			 Return ""
			}
			This.IES := IES
			This.VTBL := NumGet(IAC2 + 0, "UPtr")
			If DllCall(NumGet(This.VTBL + IAC2_Init, "UPtr"), "Ptr", IAC2 + 0, "Ptr", HEDT, "Ptr", IES, "Ptr", 0, "Ptr", 0, "UInt")
			 Return ""
			This.SetOptions(Options = "" ? ["AUTOSUGGEST"] : Options)
			This.Enabled := True
			If !(Enable)
			 This.Disable()
			If (WantReturn) {
			 ControlGet, Styles, Style,,, ahk_id %HEDT%
			 If !(Styles & 0x0004) && (CB := RegisterCallback("IAutoComplete_SubclassProc")) { ; !ES_MULTILINE
				If DllCall("SetWindowSubclass", "Ptr", HEDT, "Ptr", CB, "Ptr", &This, "Ptr", CB, "UInt")
					 This.SubclassProc := CB
				Else
					 DllCall("GlobalFree", "Ptr", CB, "Ptr")
			 }
			}
			AutoComplete.Attached[HEDT] := True

	 }
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 ; Destructor
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 __Delete() {
			; The edit control keeps own references to IAC2. Hence autocompletion has to be disabled before it can be released.
			; The only way to reenable autocompletion is to assign a new autocompletion object to the edit.
			If (This.IAC2) {
				 This.Disable()
				 ObjRelease(This.IAC2)
			}
			If (This.SubclassProc) {
				 DllCall("RemoveWindowSubclass", "Ptr", This.HWND, "Ptr", This.SubclassProc, "Ptr", &This)
				 DllCall("GlobalFree", "Ptr", This.SubclassProc)
			}
			AutoComplete.Attached.Delete(This.HWND)
	 }
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 ; Enables / disables autocompletion.
	 ;     Enable   -  True or False
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 Enable(Enable := True) {
			Static IAC2_Enable := A_PtrSize * 4
			If !(This.VTBL)
				 Return False
			This.Enabled := !!Enable
			Return !DllCall(NumGet(This.VTBL + IAC2_Enable, "UPtr"), "Ptr", This.IAC2, "Int", !!Enable, "UInt")
	 }
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 ; Disables autocompletion.
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 Disable() {
			Return This.Enable(False)
	 }
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 ;  Sets the autocompletion options.
	 ;     Options  -  Simple array of option strings corresponding to the keys defined in ACO
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 SetOptions(Options) {
			Static IAC2_SetOptions := A_PtrSize * 5
			Static ACO := {NONE: 0, AUTOSUGGEST: 1, AUTOAPPEND: 2, SEARCH: 4, FILTERPREFIXES: 8, USETAB: 16
									 , UPDOWNKEYDROPSLIST: 32, RTLREADING: 64, WORD_FILTER: 128, NOPREFIXFILTERING: 256}
			If !(This.VTBL)
				 Return False
			Opts := 0
			For Each, Opt In Options
				 Opts |= (Opt := ACO[Opt]) <> "" ? Opt : 0
			Return !DllCall(NumGet(This.VTBL + IAC2_SetOptions, "UPtr"), "Ptr", This.IAC2, "UInt", Opts, "UInt")
	 }
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 ; Updates the autocompletion strings.
	 ;     Strings  -  Simple array of strings. If you pass a non-object value the string table will be emptied.
	 ; -------------------------------------------------------------------------------------------------------------------------------
	 UpdStrings(Strings) {
			Static IID_IACDD := "{3CD141F4-3C6A-11d2-BCAA-00C04FD929DB}" ; IAutoCompleteDropDown
					 , IACDD_ResetEnumerator := A_PtrSize * 4
			If !(This.IES)
				 Return False
			If !(IEnumString_SetStrings(This.IES, Strings))
				 Return False
			If (IACDD := ComObjQuery(This.IAC2, IID_IACDD)) {
				 DllCall(NumGet(NumGet(IACDD + 0, "UPtr") + IACDD_ResetEnumerator, "UPtr"), "Ptr", This.IAC2, "UInt")
				 ObjRelease(IACDD)
			}
			Return True
	 }
}

RegisterSyncCallback(FunctionName, Options:="", ParamCount:=""){

		if !(fn := Func(FunctionName)) || fn.IsBuiltIn
				throw Exception("Bad function", -1, FunctionName)
		if (ParamCount == "")
				ParamCount := fn.MinParams
		if (ParamCount > fn.MaxParams && !fn.IsVariadic || ParamCount+0 < fn.MinParams)
				throw Exception("Bad param count", -1, ParamCount)

		static sHwnd := 0, sMsg, sSendMessageW
		if !sHwnd    {
				Gui RegisterSyncCallback: +Parent%A_ScriptHwnd% +hwndsHwnd
				OnMessage(sMsg 	:= 0x8000, Func("RegisterSyncCallback_Msg"))
				sSendMessageW 	:= DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "user32.dll", "ptr"), "astr", "SendMessageW", "ptr")
		}

		if !(pcb := DllCall("GlobalAlloc", "uint", 0, "ptr", 96, "ptr"))
				throw
		DllCall("VirtualProtect", "ptr", pcb, "ptr", 96, "uint", 0x40, "uint*", 0)

		p := pcb
		if (A_PtrSize = 8)    {

				p := NumPut(0x54894808244c8948, p+0)
				p := NumPut(0x4c182444894c1024, p+0)
				p := NumPut(0x28ec834820244c89, p+0)
				p := NumPut(   0xb9493024448d4c, p+0) - 1
				lParamPtr := p, p += 8

				p := NumPut(0xba, p+0       	, "char") ; mov edx, nmsg
				p := NumPut(sMsg, p+0       	, "int")
				p := NumPut(0xb9, p+0       	, "char") ; mov ecx, hwnd
				p := NumPut(sHwnd, p+0     	, "int")
				p := NumPut(0xb848, p+0    	, "short") ; mov rax, SendMessageW
				p := NumPut(sSendMessageW	, p+0)

				p := NumPut(0x00c328c48348d0ff, p+0)
		}
		else     { ;(A_PtrSize = 4)
				p := NumPut(0x68, p+0, "char")      ; push ... (lParam data)
				lParamPtr := p, p += 4
				p := NumPut(0x0824448d, p+0, "int") ; lea eax, [esp+8]
				p := NumPut(0x50, p+0, "char")      ; push eax
				p := NumPut(0x68, p+0, "char")      ; push nmsg
				p := NumPut(sMsg, p+0, "int")
				p := NumPut(0x68, p+0, "char")      ; push hwnd
				p := NumPut(sHwnd, p+0, "int")
				p := NumPut(0xb8, p+0, "char")      ; mov eax, &SendMessageW
				p := NumPut(sSendMessageW, p+0, "int")
				p := NumPut(0xd0ff, p+0, "short")   ; call eax
				p := NumPut(0xc2, p+0, "char")      ; ret argsize
				p := NumPut((InStr(Options, "C") ? 0 : ParamCount*4), p+0, "short")
		}
		NumPut(p, lParamPtr+0) ; To be passed as lParam.
		p := NumPut(&fn, p+0)
		p := NumPut(ParamCount, p+0, "int")
return pcb
}

RegisterSyncCallback_Msg(wParam, lParam){
		if (A_Gui != "RegisterSyncCallback")
				return
		fn := Object(NumGet(lParam + 0))
		paramCount := NumGet(lParam + A_PtrSize, "int")
		params := []
		Loop % paramCount
				params.Push(NumGet(wParam + A_PtrSize * (A_Index-1)))
		return %fn%(params*)
}

; Used internally to pass the return key to single-line edit controls.
IAutoComplete_SubclassProc(HWND, Msg, wParam, lParam, ID, Data) {
	 If (Msg = 0x0087) && (wParam = 13) ; WM_GETDLGCODE, VK_RETURN
			Return 0x0004 ; DLGC_WANTALLKEYS
	 If (Msg = 0x0002) { ; WM_DESTROY
			DllCall("RemoveWindowSubclass", "Ptr", HWND, "Ptr", Data, "Ptr", ID)
			DllCall("GlobalFree", "Ptr", Data)
			If (IAutoCompleteAC := Object(ID))
				 IAutoCompleteAC.SubclassProc := 0
	 }
	 Return DllCall("Comctl32.dll\DefSubclassProc", "Ptr", HWND, "UInt", Msg, "Ptr", wParam, "Ptr", lParam)
}

IEnumString_Create() {
	 Static IESInit := True, IESSize := A_PtrSize * 10, IESVTBL
	 If (IESInit) {
			VarSetCapacity(IESVTBL, IESSize, 0)
			Addr := &IESVTBL + A_PtrSize
			Addr := NumPut(RegisterSyncCallback("IEnumString_QueryInterface"), Addr + 0, "UPtr")
			Addr := NumPut(RegisterSyncCallback("IEnumString_AddRef")          	, Addr + 0, "UPtr")
			Addr := NumPut(RegisterSyncCallback("IEnumString_Release")          	, Addr + 0, "UPtr")
			Addr := NumPut(RegisterSyncCallback("IEnumString_Next")           	, Addr + 0, "UPtr")
			Addr := NumPut(RegisterSyncCallback("IEnumString_Skip")            	, Addr + 0, "UPtr")
			Addr := NumPut(RegisterSyncCallback("IEnumString_Reset")          	, Addr + 0, "UPtr")
			Addr := NumPut(RegisterSyncCallback("IEnumString_Clone")          	, Addr + 0, "UPtr")
			IESInit := False
	 }
	 If !(IES := DllCall("GlobalAlloc", "UInt", 0x40, "Ptr", IESSize, "UPtr"))
			Return False
	 DllCall("RtlMoveMemory", "Ptr", IES, "Ptr", &IESVTBL, "Ptr", IESSize)
	 NumPut(IES + A_PtrSize, IES + 0, "UPtr")
	 Return IES
}

IEnumString_SetStrings(IES, ByRef Strings) {
	 PrevTbl := NumGet(IES + (A_PtrSize * 9), "UPtr")
	 StrSize := 0
	 StrArray := []
	 Loop, % Strings.Length()
			If ((S := Strings[A_Index]) <> "")
				 L := StrPut(S, "UTF-16") * 2
				 , StrSize += L
				 , StrArray.Push({S: S, L: L})
			Else
				 Break
	 StrCount := StrArray.Length()
	 StrTblSize := (A_PtrSize * 2) + (StrCount * A_PtrSize * 2) + StrSize
	 If !(StrTbl := DllCall("GlobalAlloc", "UInt", 0x40, "Ptr", StrTblSize, "UPtr"))
			Return False
	 Addr := StrTbl + A_PtrSize
	 Addr := NumPut(StrCount, Addr + 0, "UPtr")
	 StrPtr := Addr + (StrCount * A_PtrSize * 2)
	 For Each, Str In StrArray {
			Addr := NumPut(StrPtr, Addr + 0, "UPtr")
			Addr := NumPut(Str.L, Addr + 0, "UPtr")
			StrPut(Str.S, StrPtr, "UTF-16")
			StrPtr += Str.L
	 }
	 If (PrevTbl)
			DllCall("GlobalFree", "Ptr", PrevTbl)
	 NumPut(StrTbl, IES + (A_PtrSize * 9), "UPtr")
	 Return True
}

IEnumString_QueryInterface(IES, RIID, ObjPtr) {
	 Static IID := "{00000101-0000-0000-C000-000000000046}", IID_IEnumString := 0
				, Init := VarSetCapacity(IID_IEnumString, 16, 0) + DllCall("Ole32.dll\IIDFromString", "WStr", IID, "Ptr", &IID_IEnumString)
	 Critical
	 If DllCall("Ole32.dll\IsEqualGUID", "Ptr", RIID, "Ptr", &IID_IEnumString, "UInt") {
			IEnumString_AddRef(IES)
			Return !(NumPut(IES, ObjPtr + 0, "UPtr"))
	 }
	 Else
			Return 0x80004002
}

IEnumString_AddRef(IES) {
	 NumPut(RefCount := NumGet(IES + (A_PtrSize * 8), "UPtr") + 1,  IES + (A_PtrSize * 8), "UPtr")
	 Return RefCount
}

IEnumString_Release(IES) {
	 RefCount := NumGet(IES + (A_PtrSize * 8), "UPtr")
	 If (RefCount > 0) {
			NumPut(--RefCount, IES + (A_PtrSize * 8), "UPtr")
			If (RefCount = 0) {
				 DllCall("GlobalFree", "Ptr", NumGet(IES + (A_PtrSize * 9), "UPtr")) ; string table
				 DllCall("GlobalFree", "Ptr", IES)
			}
	 }
	 Return RefCount
}

IEnumString_Next(IES, Fetch, Strings, Fetched) {
	 Critical
	 I := 0
	 StrTbl     	:= NumGet(IES + (A_PtrSize * 9), "UPtr")
	 Current  	:= NumGet(StrTbl + 0, "UPtr")
	 Maximum 	:= NumGet(StrTbl + A_PtrSize, "UPtr")
	 StrAddr  	:= StrTbl + (A_PtrSize * 2) + (A_PtrSize * Current * 2)

	 While (Current < Maximum) && (I < Fetch)
			Ptr := NumGet(StrAddr + 0, "UPtr")
			, Len := NumGet(StrAddr + A_PtrSize, "UPtr")
			, Mem := DllCall("Ole32.dll\CoTaskMemAlloc", "Ptr", Len, "UPtr")
			, DllCall("RtlMoveMemory", "Ptr", Mem, "Ptr", Ptr, "Ptr", Len)
			, NumPut(Mem, Strings + (I * A_PtrSize), "Ptr")
			, NumPut(++I, Fetched + 0, "UInt")
			, NumPut(++Current, StrTbl + 0, "UPtr")
			, StrAddr += A_PtrSize * 2

Return (I = Fetch) ? 0 : 1
}

IEnumString_Skip(IES, Skip) {
	 Critical
	 StrTbl := NumGet(IES + (A_PtrSize * 9), "UPtr")
	 , Current := NumGet(StrTbl + 0, "UPtr")
	 , Maximum := NumGet(StrTbl + A_PtrSize, "UPtr")
	 If ((Current + Skip) <= Maximum)
			Return (NumPut(Current + Skip, StrTbl, "UPtr") & 0)
	 Return 1
}

IEnumString_Reset(IES) {
	 Return (NumPut(0, NumGet(IES + (A_PtrSize * 9), "UPtr"), "UPtr") & 0)
}
; ----------------------------------------------------------------------------------------------------------------------------------
IEnumString_Clone(IES, ObjPtr) { ; Not sure about the reference counter (IES + (A_PtrSize * 8))!
	 IESSize := DllCall("GlobalSize", "Ptr", IES, "Ptr")
	 StrTbl := NumGet(IES + (A_PtrSize * 9), "UPtr")
	 StrTblSize := DllCall("GlobalSize", "Ptr", StrTbl, "Ptr")
	 If !(IESClone := DllCall("GlobalAlloc", "UInt", 0x40, "Ptr", IESSize, "UPtr"))
			Return False
	 If !(StrTblClone := DllCall("GlobalAlloc", "UInt", 0x40, "Ptr", StrTblSize, "UPtr")) {
			DllCall("GlobalFree", "Ptr", IESClone)
			Return False
	 }
	 DllCall("RtlMoveMemory", "Ptr", IESClone, "Ptr", IES, "Ptr", IESSize)
	 DllCall("RtlMoveMemory", "Ptr", StrTblClone, "Ptr", StrTbl, "Ptr", StrTblSize)
	 NumPut(0, IESClone + (A_PtrSize * 8), "UPtr") ; Set the reference counter to zero or one in this case???
	 NumPut(StrTblClone, IESCLone + (A_PtrSIze * 9), "UPtr")
	 Return (NumPut(IESClone, ObjPtr + 0, "UPtr") & 0)
}


;}


/*



		;~ mainwords := !IsObject(mainwords) ? StrSplit(mainwords, "`n") : mainwords
		;~ For each, mainword in mainwords {
			;~ !mainword ? continue : ""
			;~ mainexists := false
			;~ For idx, mainStored in this.autonames[mainclass].main
				;~ If (mainStored = mainword) {
					;~ mainexists := true
					;~ break
				;~ }
			;~ If !mainexists
				;~ this.autonames[mainclass].main.Push(mainword)
		;~ }
		;~ subwords := !IsObject(subwords) ? StrSplit(subwords, "`n") : subwords
		;~ For each, subword in subwords {
			;~ !subword ? continue : ""
			;~ subexists := false
			;~ For idx, subStored in this.autonames[mainclass].sub[subclass]
				;~ If (subStored = subword) {
					;~ subexists := true
					;~ break
				;~ }
			;~ If !subexists
				;~ this.autonames[mainclass].sub[subclass].Push(subword)
		;~ }


*/


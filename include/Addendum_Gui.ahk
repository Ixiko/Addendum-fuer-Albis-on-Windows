; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;               															⚫      INFOFENSTER		⚫
;
;      Funktionen:   	▫ Verwaltung/Bearbeitung (Texterkennung/Kategorisierung) gescannter Post vor dem Import in die Patientenkartei
;                       	▫ Anzeigen von Praxisinformationen
;                       	▫ Netzwerkkommunikation
;                       	▫ RDP Sessions starten
;                       	▫ erweitertes Tagesprotokoll
;							▫ Programme und Skripte starten
;							▫ Abrechnungshelfer
;							▫ Zusatzfunktionen: Laborblattdruck
;
;      Basisskript: 	Addendum.ahk
;
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - letzte Änderung 12.07.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

; Infofenster
AddendumGui(admShow="", admGuiDebug="false")                         	{

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global

		static admWidth, admHeight, admTabNames, admLVPatOpt, admLVJournalOpt, admLVTProtokoll ;, DDLPrinter
		static ImportChecks := 0
		static admTags

		local MObj, MNr, cpx, cpy, cpw, cph
		local APos, SPos, mox, moy, moWin, moCtrl, regExt, Aktiv, Pat, nfo, res
		local ClientName, LanClients, ClientIP, onlineclient, OffLineIndex, InfoText, found, YPlus, Y

		If WinExist("AddendumGui ahk_class AutohotkeyGui") {
			SciTEOutput("Infofenster vorhanden.")
			RedrawWindow(hadm)
			return
		}

		If !Addendum.iWin.Init {

			; Befunde (Klassen-Objekt) - Dateihandler für alle Dateizugriffe auf den Befundordner
				Befunde.options(false)

				Addendum.iWin.Init            	:= 1
				Addendum.iWin.FilesStack 	:= []
				Addendum.iWin.OCRStack 	:= []
				Addendum.iWin.RowColors	:= false
				Addendum.iWin.Imports     	:= 0
				Addendum.iWin.ImportsLast 	:= -1
				Addendum.iWin.Check      	:= 0
				Addendum.iWin.ReCheck   	:= 0
				Addendum.iWin.paint        	:= 0
				Addendum.ImportFromPat   	:= false           	; Import aus Patient Tab läuft
				Addendum.ImportFromJrnl   	:= false           	; Import aus dem Journal Tab läuft
				Addendum.tessOCRRunning	:= false        	; OCR Vorgang läuft

				func_admGui       	:= Func("AddendumGui")
				factor                	:= A_ScreenDPI / 96

				admTags            	:= {	"positive"	: "Karteikarte|Laborblatt|Biometriedaten|Rechnungsliste"
											   ,	"negative"	: "Abrechnung|Privatabrechnung|Überweisung"}

				admTabNames  	:= "Patient|Journal|Protokoll|Extras|Netzwerk|Info"
				admLVPatOpt    	:= " r" (Addendum.iWin.LVScanPool.r)
				admBGColor    	:= "BackgroundF0F0F0 "
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

				;~ admNetW := []
				;~ admNetW.Push(LoadPicture(Addendum.Dir "\assets\ModulIcons\connected.png"))
				;~ admNetW.Push(LoadPicture(Addendum.Dir "\assets\ModulIcons\disconnected.png"))

				;hbmPdf_Ico   	:= Create_PDF_ico()
				;hbmImage_Ico	:= Create_Image_ico()
				;~ IL_Add(ImageListID, "HBITMAP: " hbmPdf_Ico 		, 0xFFFFFF, 0) 		;1
				;~ IL_Add(ImageListID, "HBITMAP: " hbmImage_Ico	, 0xFFFFFF, 0)		;2

			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Kontextmenu (Rechtsklickmenu) für das Journal
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				func_admOpen 	:= Func("admGui_CM").Bind("JOpen")
				func_admImport	:= Func("admGui_CM").Bind("JImport")
				func_admRename	:= Func("admGui_CM").Bind("JRename")
				func_admSplit   	:= Func("admGui_CM").Bind("JSplit")
				func_admDelete	:= Func("admGui_CM").Bind("JDelete")
				func_admView1 	:= Func("admGui_CM").Bind("JView1")
				func_admView2 	:= Func("admGui_CM").Bind("JView2")
				func_admRefresh	:= Func("admGui_CM").Bind("JRefresh")
				func_admRecog1	:= Func("admGui_CM").Bind("JRecog1")
				func_admRecog2	:= Func("admGui_CM").Bind("JRecog2")
				func_admOCR  	:= Func("admGui_CM").Bind("JOCR")
				func_admOCRAll	:= Func("admGui_CM").Bind("JOCRAll")

				Menu, admJCM, Add, Karteikarte öffnen                      	, % func_admOpen

				If (StrLen(Addendum.PDF.Reader) > 0) && (StrLen(Addendum.PDF.ReaderAlternative) > 0) {
					Menu, admJCMView, Add, % Addendum.PDF.ReaderName            	, % func_admView2
					Menu, admJCMView, Add, % Addendum.PDF.ReaderAlternativeName	, % func_admView1
					Menu, admJCM, Add, Anzeigen mit                          	, :admJCMView
				} else if (StrLen(Addendum.PDF.Reader) = 0) && (StrLen(Addendum.PDF.ReaderAlternative) > 0)
					Menu, admJCM, Add, Anzeigen                                	, % func_admView1
				else if (StrLen(Addendum.PDF.Reader) > 0) && (StrLen(Addendum.PDF.ReaderAlternative) = 0)
					Menu, admJCM, Add, Anzeigen                                	, % func_admView2

				Menu, admJCM, Add, Importieren                             	, % func_admImport
				Menu, admJCM, Add, Umbennen                              	, % func_admRename
				;Menu, admJCM, Add, Aufteilem                                   	, % func_admSplit
				Menu, admJCM, Add, Löschen                                     	, % func_admDelete
				Menu, admJCM, Add, Texterkennung ausführen           	, % func_admOCR
				Menu, admJCM, Add, Inhaltserkennung                      	, % func_admRecog1
				Menu, admJCM, Add
				Menu, admJCM, Add, Texterkennung - alle Dateien       	, % func_admOCRAll
				Menu, admJCM, Add, automatische Benennung           	, % func_admRecog2
				Menu, admJCM, Add, Befundordner neu indizieren    	, % func_admRefresh

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
				Hotkey, F4     	, % func_admRecog1
				Hotkey, F5     	, % func_admRefresh
				Hotkey, F6     	, % func_admOCR
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
				TrayTip, AddendumGui, % "Die Integration des Infofenster ist fehlgeschlagen.", 2
			}
			Addendum.iWin.ReCheck := Addendum.iWin.paint := Addendum.iWin.Check := Addendum.iWin.lastPatID := 0
			return

		}

	; Position innerhalb des Albisfenster für die Integration finden
		Aktiv          	:= Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		res                := Controls("", "Reset", "")
		hMDIFrame	:= Controls("AfxMDIFrame1401"	, "ID", AlbisMDIChildGetActive())     	; AlbisMDIChildGetActive(): Steuerelement-ID
																																		; im aktiven Fenster
		hStamm		:= Controls("#327701"             	, "ID", hMDIFrame)                        	; Stammdaten Steuerelement-ID
		Addendum.iWin.Check ++
		If (!RegExMatch(Aktiv, "i)(" Tags.positive ")", winType) || !hMDIFrame || !hStamm) {   	; || RegExMatch(Aktiv, "i)(" Tags.negative ")")
			Addendum.iWin.lastPatID := 0
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
	If (Addendum.iWin.lastPatID <> (Addendum.iWin.PatID := AlbisAktuellePatID())) {

			Addendum.iWin.lastPatID := Addendum.iWin.PatID

		; Position von Albis und der Stammdaten ;{
			APos          	:= GetWindowSpot(AlbisWinID())
			SPos          	:= GetWindowSpot(hStamm)
		;}

		; legt die Höhe der Gui anhand der Größe des MDIFrame-Fenster fest ;{
			Addendum.iWin.StammY	:= Round(SPos.Y / factor)
			Addendum.iWin.X	         	:= APos.CW - Addendum.iWin.W ; ClientWidth passt besser!
			Addendum.iWin.H          	:= admHeight := SPos.CH
			admWidth                     	:= Addendum.iWin.W
			;~ Addendum.iWin.lastPatID	:= AlbisAktuellePatID()    ; warum nochmal?
		;}

		; der Albisinfobereich (Stammdaten, Dauerdiagnosen, Dauermedikamente) wird für das Einfügen der Gui verkleinert ;{
			If (SPos.W > APos.CW - admWidth) ; verhindert ein wiederholtes Verkleinern
				SetWindowPos(hStamm, 2, 2, APos.CW - admWidth - 4, SPos.CH)
		;}

		; umgeht die seltene Problematik der eines nicht mehr vorhandenen Handle eines MDIFrame, wenn der Nutzer eine
		; Karteikarte schon geschlossen hat und das Skript noch dieses Handle erfasst hatte
			Sleep 100
			try {
			Gui, adm: New	, -Caption -DPIScale +ToolWindow 0x50000000 E0x0802009C +Parent%hMDIFrame% +HWNDhadm
		} catch {
			SetTimer, % func_admGui, -1000
			Addendum.iWin.ReCheck ++
			return
		}

		;-: Gui 	 :- Tab Start                       	;{
			Gui, adm: Margin	, 0, 0
			Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Tab     	, % "x0 y0  	w" admWidth " h" admHeight+2 " HWNDadmHTab vadmTabs gadm_Tabs"       	, % admTabNames
		;}

		;-: Tab1 :- Patient                          	;{
			Gui, adm: Tab  	, 1

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Text      	, % "x10 y25 w" admWidth-15 " BackgroundTrans vadmPatientTitle Section"                     	, % "(es wird gesucht ....)"

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Button    	, % "x" admWidth-55 " y23 h16 vadmButton1 gadm_BefundImport"                                   	, % "Importieren"
			GuiControlGet, cpos, adm: Pos, admButton1
			GuiControl, adm: Move, admButton1, % "x" admWidth-cposW-5

		  ;-: Patientenbefunde
			admCol1W :=50, admCol2W := admWidth - admCol1W - 10
			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, ListView	, % "x2 y+1 	w" admWidth-6	" " admLVPatOpt                                                                   	, % "S|Befundname"
			LV_SetImageList(admImageListID)
			LV_ModifyCol(1, admCol1W " Integer Right NoSort")
			LV_ModifyCol(2, admCol2W " Text Left NoSort")

		  ;-: Abrechnungshelfer
			cpw := Floor((admWidth-6)/1.9)
			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Text    	, % "x2 y+0 	w" cpw " Center vadmGP1"   	                                                                    	, % "Abrechnungshelfer und andere Hinweise"

			Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
			GuiControlGet, cpos, adm: Pos, AdmGP1

			cph       	:= admHeight - cposY - cposH - 4 - 70
			abrInfo    	:= Addendum.iWin.AbrHelfer ? admGui_Abrechnungshelfer(AlbisAktuellePatID(), AlbisPatientGeburtsdatum()) : ""
			GCOption	:= "x2 y" cposY+12 " w" cpw " h" cph " t6 t12 t22 vadmNotes"
			Gui, adm: Add, Edit, % GCOption                                                                                                                              	, %  abrInfo

		  ;-: kontextsensitive Hotkey's anzeigen
			;~ admHKIOpt := "ReadOnly -TabStop -VScroll"
			 ;~ Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
			 ;~ Gui, adm: Add  	, Edit    	, % "x2 y+5 	w" cpw " h68   vadmHKI1 " admHKIOpt                                                           	, % "Shift+F6 - Vorlagen GOÄ Ziffer Begründungen"
			 ;~ Gui, adm: Add  	, Edit    	, % "x+2    	w" admWidth-cpw-7 " h68  vadmHKI2 " admHKIOpt                                       	, % "Shift+F5 - Heute"

		  ;-: zusätzliche Karteikartenfunktionen (Laborblatt drucken/Laborblatt per Mail versenden)
			GuiControlGet, cpos, adm: Pos, admNotes
			cpx := cposX + cposW + 3, cpy := cposY - 1, cpw	:= admWidth - cpx - 120 - 6
			Gui, adm: Add, Button       	, % "x" cpx " y" cpy " w78 vadmLBD  gadm_LBDruck"                                                        	, % "Labor drucken"
			Gui, adm: Add, Edit           	, % "x+0  y" cpy+1 " w" 52 " vadmLBDCB"                                                                       ;	, % "1|alle"
			Gui, adm: Add, UpDown   	, % "x+0  vadmLBDUD gadm_LBDruck"                                                                            	,  1

			Gui, adm: Font 	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add, Checkbox		, % "x+2  y" cpy " w" 84 " vadmLBCk  gadm_LBDruck"                                                        	, % "Anmerkung/`nProbeddaten"
			GuiControl, adm: , admLBCk, % Addendum.iWin.LBAnm

			Gui, adm: Font 	, s7 q5 Normal cBlack, Arial
			GuiControlGet, cpos, adm: Pos, admLBD
			cpw	:= admWidth - cpx - 5, cpy := cposY+cposH+1
			Gui, adm: Add	, DDL          	, % "x" cpx " y" cpy " w" cpw " r8 vadmLBDPD gadm_LBDruck"                                           	, % DDLPrinter
			GuiControl, adm: ChooseString, admLBDPD, % Addendum.iWin.LBDrucker

		;-: Laborblattversand per EMail
			If oPat[AlbisAktuellePatID()].EMAIL {
				GuiControlGet, cpos, adm: Pos, admLBDPD
				Gui, adm: Font	, s8 q5 Normal cBlack, Arial
				Gui, adm: Add	, Button       	, % "x" cpx-1 " y+5 w" cpW+2 " vadmLBMail  gadm_LBDruck"                                                	, % "Laborblatt per Mail versenden an:"
				Gui, adm: Font	, s8 q5 underline italic cBlue, Arial
				Gui, adm: Add	, Text           	, % "x" cpx-1 " y+0 w" cpW+2 " vadmLBMAdr  Center Backgroundtrans"                                	, % oPat[AlbisAktuellePatID()].EMAIL
				Gui, adm: Font	, s8 q5 Normal cBlack, Arial
			}
			;}

		;-: Tab2 :- Journal                         	;{
			Gui, adm: Tab  	, 2

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Text      	, % "x10 y25 w" Floor(admWidth/3) " BackgroundTrans vadmJournalTitle 	Section"          	, % "(es wird gesucht ....)"
			Gui, adm: Add  	, Text      	, % "x+5                                          BackgroundTrans vadmJournalTitle2 	Section"          	, % "                          "

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, Button    	, % "x" admWidth-145 " y23 h16 w" (11*7) " 	vadmButton2 gadm_Journal Center"             	, % "Importieren"
			Gui, adm: Add  	, Button    	, % "x20 y23 h16                                       	vadmButton3 gadm_Journal"                        	, % "Aktualisieren"

			GuiControlGet  	, cpos, adm: Pos	, admButton2
			GuiControlGet  	, dpos, adm: Pos	, admButton3
			GuiControl        	, adm: Move, admButton2, % "x" admWidth - cposW - 5
			GuiControl        	, adm: Move, admButton3, % "x" admWidth - cposW - dposW - 5
			Gui, adm: Add  	, Button    	, % "x20 y23 h16                                        	vadmButton4 gadm_Journal"                       	, % "OCR ausführen  "

			GuiControlGet  	, cpos, adm: Pos, admButton3
			GuiControlGet  	, dpos, adm: Pos, admButton4
			GuiControl        	, adm: Move, admButton4, % "x" cposX - dposW

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			Gui, adm: Add  	, ListView	, % "x2 y+5 		w" admWidth - 6 " h" admHeight - 47 " " admLVJournalOpt                        	, % "Befund|S|Eingang|TimeStamp"

			admCol4W := 0, admCol3W := 58, admCol2W := 25
			admCol1W := admWidth - admCol2W - admCol3W - 25

			Gui, adm: ListView, % "admJournal"
			LV_ModifyCol(1, admCol1W " Left NoSort")
			LV_ModifyCol(2, admCol2W " Left Integer")
			LV_ModifyCol(3, admCol3W " Right NoSort")    	; versteckte Spalte enthält Integerzeitwerte für die Sortierung nach dem Datum
			LV_ModifyCol(4, admCol4W " Left Integer")       	; versteckte Spalte enthält Integerzeitwerte für die Sortierung nach dem Datum
			Gui, adm: ListView, % "admJournal"
			LV_SetImageList(admImageListID)

		;}

		;-: Tab3 :- Protokoll (letzte Patienten)	;{
			Gui, adm: Tab  	, 3

			Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
			ProtText := "[" compname "] [" (TProtokoll.Count()-1 < 0 ? 0 : TProtokoll.Count()-1) " Pat.]"
			Gui, adm: Add  	, Text      	, % "x5 y25  w130 BackgroundTrans 	vadmTProtokollTitel"                                              	, % ProtText

			Gui, adm: Font  	, s8 q5 Normal cBlack
			Gui, adm: Add  	, Text      	, % "x+10 "                                                                                                                   	, % "Pat:"
			Gui, adm: Add  	, Edit     	, % "x+0 y22 w100 r1  vadmTPSuche"                                                                               	, % ""

			admTPRightX := admWidth - 80 - 5
			bgt := "BackgroundTrans"
			Gui, adm: Add  	, Text      	, % "x" admTPRightX " y25 w80          	vadmTPTag       	gadm_TP " bgt                            	, % Addendum.iWin.TProtDate
			Gui, adm: Add  	, Button    	, % "x+10 y23 h16                          	vadmTPZurueck	 	gadm_TP"                                  	, % "<"
			Gui, adm: Add  	, Button    	, % "x+0 	y23 h16						   	vadmTPVor   		gadm_TP"                                  	, % ">"
			GuiControlGet  	, dpos, adm: Pos, admTPZurueck
			GuiControlGet  	, cpos, adm: Pos, admTPVor
			GuiControl        	, adm: Move, admTPVor    	, % "x" admTPRightX - cposW - 2
			GuiControl        	, adm: Move, admTPZurueck	, % "x" admTPRightX - cposW - dposW - 2

			Gui, adm: Font  	, s8 q5 Normal cBlack
			Gui, adm: Add  	, ListView	, % "x2 y+5 w" admWidth-6 " h" admHeight-47 " " admLVTProtokoll                                    	, % "RF|Nachname, Vorname|Geburtstag|Alter|PatID"

			admcol1W := 30, admcol2W := 160, admcol3W := 70, admcol4W := 40, admCol5W := 70
			Gui, adm: ListView, admTProtokoll
			LV_ModifyCol(1, admcol1W )
			LV_ModifyCol(2, admcol2W )
			LV_ModifyCol(3, admcol3W )
			LV_ModifyCol(4, admcol4W )
			LV_ModifyCol(5, admcol5W )
		;}

		;-: Tab4 :- Extras                               ;{
			Gui, adm: Tab  	, 4

			; MODUL BUTTON
			Gui, adm: Font , s10 q5 bold cBlack, Calibri
			Gui, adm: Add, Button	, % "x10 y35 w" admWidth-20 " h20 Center hwndadmHPB", % "< - - - - - - - - - - - - - - M O D U L E - - - - - - - - - - - - - - >"  ; vadmMMM
			Opt1 := { 1:0, 2:0xFFAfC2ff , 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1}
			Opt2 := { 1:0, 2:0xFFAfC2ff	, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1}
			Opt3 := { 1:0, 2:0xFFAfC2ff	, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1}
			ImageButton.Create(admHPB, Opt1, Opt2, Opt3)

			; MODUL BUTTONS (feste Größe)
			Gui, adm: Font  	, s9 q5 Normal cBlack, Calibri
			ImageButton.SetGuiColor("Green")
			GuiControlGet, EX, adm: Pos, % admHPB
			Mw :=Mh := 15, MEw := MEh :=20
			mStep 	:= 0 	, mxWidth	:= 95, stepX := mxWidth + 5
			mdX  	:= 15	, mdY     	:= mdY := EXY +EXH + 7
			mCols 	:= Floor((admWidth - (mdX*2))/(mxWidth+5))
			stepX 	:= Round(((admWidth - (mdX*2)) - (mCols*mxWidth))/(mCols-1))
			stepX 	:= stepX < 5 ? 5 : stepX
			For IWmIdx, IWModul in Addendum.Module {

				IWx := mStep < mCols ? (mdX + mStep*(mxWidth+stepX)) : (admWidth-mdX-mxWidth)
				Gui, adm: Add, Button, % "x" IWx " y" mdY " w" mxWidth " vadmMX" IWmIdx " hwndadmHPB  gadm_extras", % "  " IWModul.name " " ;Center
				Opt1 := { 1:0, 2:0xFFAfC2ff	, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1, icon:{file:IWModul.ico, x: 9      	, w: Mw		, h: Mh	}}
				Opt2 := { 1:0, 2:0xFF8fA2ff	, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1, icon:{file:IWModul.ico, x: 7, y: 3	, w: MEw	, h: MEh}}
				Opt3 := { 1:0, 2:0xFFAfC2f	, 4:"Black", 5:"H", 6:"WHITE", 7:"Black", 8:1, icon:{file:IWModul.ico, x: 9      	, w: Mw		, h: Mh	}}
				ImageButton.Create(admHPB, Opt1, Opt2, Opt3)
				mStep ++
				If (mStep = mCols) {
					mStep := 0
					GuiControlGet, EX, adm: Pos, % "admMX" IWmIdx
					mdY := EXY + EXH + 5
				}

			}

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

		;}

		;-: Tab5 :- Netzwerk                       	;{
			Gui, adm: Tab  	, 5

			dposX := 40, dposY := 60, cMaxW := 0
			Gui, adm: Font  	, s8 q5 Normal underline cBlack, Arial
			Gui, adm: Add   	, Button	, % "x10 y30 vAdmLanCheck gadm_Lan", % "Netzwerkgeräte aktualisieren"
			Gui, adm: Font  	, s8 q5 Normal cBlack, Arial

		  ;-: zeichnet Buttons mit den Clients im Netzwerk und stellt einen weiteren Button für direkten RDP Zugriff bereit ;{
			For clientName, client in Addendum.LAN.Clients {
				If (compname = clientName)
					continue
				Gui, adm: Add   	, Button	, % "x" dposX " y" dposY " vadmClient_"	clientName " Center gadm_Lan", % clientName
				Gui, adm: Add   	, Button	, % "x150		  y" dposY " vadmRDP_"  	clientName " Center gadm_Lan", % "RDP Sitzung starten"
				GuiControlGet	, cpos, adm: Pos, % "AdmClient_" clientName
				dposY	:= cposY + cposH
				cMaxW	:= cposW > cMaxW ? cposW : cMaxW
			}
			For clientName, client in Addendum.LAN.Clients {
				If (compname = clientName)
					continue
				GuiControl, adm: Move, % "AdmClient_" clientName,	% "w" cMaxW
				GuiControl, adm: Move, % "AdmRDP_" 	 clientName,	% "x"	 dposX + 20 + cMaxW
				GuiControlGet	, cpos, adm: Pos, % "AdmClient_" clientName
				Gui, adm: Add, Picture, % "x10 y" cposY+2 " w" cposH-4 " h" cposH-4 " Backgroundtrans vAdmConn_" clientName ;, % "HBITMAP:" Addendum.Dir "\assets\ModulIcons\connected.png"
			}
			;Gui, adm: Add    	, Text, % "x10 y30 w110 BackgroundTrans Center vAdmLanT1", % "ONLINE"
			;SetTimer, admLan, -300
		 ;}

			;}

		;-: Tab6 :- Info                              	;{
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
			Gui, adm: Add  	, Text, x+5, % (TProtokoll.MaxIndex() = "" ? 0 : TProtokoll.MaxIndex()) " Patienten"

			Gui, adm: Font  	, % "s" fs " bold cGreen"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw, % "⚕ Patienten ges."
			Gui, adm: Font  	, Normal
			Gui, adm: Add  	, Text, x+5, % oPat.Count()

			Gui, adm: Font  	, % "s" fs " bold cRed"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw, % "✉ Signaturen"
			Gui, adm: Font  	, Normal
			Gui, adm: Add  	, Text, x+5, % Addendum.PDF.SignatureCount ", Seitenzahl: " Addendum.PDF.SignaturePages

			Gui, adm: Font  	, % "s" fs " bold cAA2277"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw " hwndadmHMT", % "☀ Monitor"

			GuiControlGet, cp, adm: Pos, % admHMT
			cpy -= 1

			Loop, % Addendum.Monitor.Count() {
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
			Gui, adm: Add   	, Text, % "x+5 w" admWidth-tw-10 " vadmILAB1 " , % " `n "
			Gui, adm: Add   	, Text, % "y+2 w" admWidth-tw-10 " vadmILAB2 " , % ""
			;Gui, adm: Add  	, Text, % "x10 y" admHeight-50 " w" admWidth-20 " h40 vadmILAB1"  	, % ""

			Addendum.iWin.paint ++
			GuiControlGet, cp, adm: Pos, % admILAB2

			Gui, adm: Font  	, % "s" fs " bold c172842"
			Gui, adm: Add  	, Text, % "x10 y+" YPlus " w" tw      	, % "🔄 Infofenster"
			Gui, adm: Font  	, % "s" fs " Normal"
			Gui, adm: Add  	, Text, % "x+5 w100 vadmIWInit"	, % "[" Addendum.iWin.paint ":" Addendum.iWin.Init ":0]"

		;}

		;-: Tab7 :- Einstellungen                  	;{

		;}

		Gui, adm: Tab

		admGui_ShowTab(Addendum.iWin.firstTab, admHTab)
		Addendum.iWin.paint ++
	}
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeigen
	;---------------------------------------------------------------------------------------------------------------------------------------------;{
		; MDIFrame handle gehört dem aktuellen AfxMDIFrame Steuerelement ?
			Controls("Reset")
			If !InStr(Controls("", "ControlFind, " hMDIFrame ",, return class", AlbisMDIChildGetActive()), "AfxMDIFrame1401")
				hMDIFrame := Controls("AfxMDIFrame1401", "ID", AlbisMDIChildGetActive())
				; evtl. das MDIFrame auf WS_Clipsiblings 0x400000 setzen

		; Gui dem aktuellen AfxMidiFrame als Child zuordnen
			SetParentByID(hMDIFrame, hadm)
			WinSet, Style   	, 0x54000000	, % "ahk_id " hadm
			WinSet, ExStyle 	, 0x0802009C	, % "ahk_id " hadm  ; evtl. +MDIChildWindow 0x40 und/oder AppWindow 0x80000

		; Fenster anzeigen
			Gui, adm: Show, % "x" Addendum.iWin.X " y" Addendum.iWin.Y " w" Addendum.iWin.W  " h" Addendum.iWin.H
									. 	 " NoActivate " admShow, AddendumGui

		; Neuzeichnen der Gui abwarten
			Sleep 300

		; Fenster ist kein Child (WS_POPUP = 0x80000000) - schliessen und neuzeichnen
			WinGet, Style, Style, % "ahk_id " hadm
			If (Style & 0x80000000) {
				admGui_Destroy()
				AddendumGui()
			}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tabs / Listviews mit Daten befüllen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		; Inhalte auffrischen und anzeigen
			Addendum.iWin.Init ++
			GuiControl, adm:, admIWInit, % "[P" Addendum.iWin.paint ":I" Addendum.iWin.Init ":R"Addendum.iWin.ReCheck "]"
			Addendum.iWin.Check := Addendum.iWin.ReCheck := 0
			admGui_Journal( (!Addendum.iWin.Init ? true : false) )
			If !Addendum.ImportFromJrnl {
				admGui_ShowTab(Addendum.iWin.firstTab, admHTab)
				admGui_CountDown()
				admGui_TProtokoll(Addendum.iWin.TProtDate, 0, compname)
				PatDocs := admGui_Reports()
			}
			else
				admGui_ShowTab("Journal", admHTab)

	;}

	; Journalimport bei Bedarf fortsetzen
		Journal.ShowImportStatus(Addendum.ImportFromJrnl)
		If Addendum.ImportFromJrnl {
			;SetTimer, adm_ImportRunCheck, -1000
			nil := 0
		}

	; PDF Thumbnailpreview
		;OnMessage(0x200,"admGui_ThumbnailView")

return
}

; -------- Gui Labels                                                                            	;{
adm_ImportRunCheck:                                                                       	;{ SetTimer Label

		ImportChecks ++
		ToolTip, % "Warte auf Ende des aktuellen Importvorgangs..`n"
					. 	"ImportCheck nr:  " Importchecks "`n"
					. 	"Imports: "     	Addendum.iWin.Imports "`n"
					. 	"ImportsLast: " 	Addendum.iWin.ImportsLast

		If (Addendum.iWin.ImportsLast < Addendum.iWin.Imports) {
			If (Addendum.iWin.ImportsLast >= 0)
				SetTimer, adm_ImportRunCheck, -100
			else
				ToolTip
			return
		}

		SetTimer, adm_ImportRunCheck, Off
		SetTimer, adm_ImportNext, -1000

return

adm_ImportNext:

		ToolTip
		ImportChecks := 0
		res := admGui_ShowTab("Journal", admHTab)
		admGui_ImportJournalAll()

return
;}

admLAN:                                                                                          	;{ zeichnet den LAN Tab

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

adm_Tabs:                                                                                        	;{ Tab Handler

	If (A_GuiEvent = "Normal") {

		; aktueller Tab ist Info, Countdown auffrischen
			SendMessage, 0x130B,,,, % "ahk_id " admHTab 	; TCM_GETCURSEL
			If (ErrorLevel = 5)  ; Info Tab angewählt
				admGui_CountDown()

		; sichern der aktuellen Tab-Ansicht
			GuiControlGet, admTabName, adm:, admTabs
			Addendum.iWin.firstTab := admTabName

	}

return ;}

adm_Lan:                                                                                         	;{ Netzwerk/LAN Tab Handler

	If (A_GuiControl = "AdmLanCheck") {


		;~ For clientname, Client in Addendum.LAN.Clients {

			;~ ;SciTEOutput("IP: " clientname " , Ping: " IPHelper.Ping(client.ip))
			;~ GuiControl, adm:, % "AdmConn_" clientName, % "HBITMAP:" Addendum.Dir "\assets\ModulIcons\connected.png"

		;~ }
			;~ LANMsg := ""
			;~ admSendText("192.168.100.45", "answer|" A_ComputerName "|" A_IPAddress1 "|Status Ok" )
			;~ ;admSendText("192.168.100.25", "answer")
			;~ myTcp := new SocketTCP()
			;~ myTcp.connect("localhost", 1337)
			;~ MsgBox, % myTcp.recvText()

	} else if RegExMatch(A_GuiControl, "admRDP_") {
			Run, % q Addendum.LAN.Clients[compname].rdpPath "\" StrReplace(A_GuiControl, "admRDP_") ".rdp" q
	} else if RegExMatch(A_GuiControl, "admClient_(?<name>.*)", client) {

			SciTEOutput("clientname: " clientName)

			myTcp := new SocketTCP()
			myTcp.connect(Addendum.LAN.admServer.IP, Addendum.LAN.admServer.Port)
			MsgBox, % myTcp.recvText()
			myTcp.sendText("Hello Server!")

	}

return ;}

adm_Extras:		                                                                         			;{ Startet Module / Tools

	If !RegExMatch(A_GuiControl, "adm(?<Obj>[A-Z]+)(?<Nr>\d+)", M)
		return

	Switch MObj 	{

		case "MX":

			EModulExe    	:= Addendum.Module[MNr].command
			EModulName	:= Addendum.Module[MNr].name
			If FileExist(EModulExe) {
				PraxTT("starte Modul : " EModulName, "1 1")
				SplitPath, EModulExe, WorkPath
				Run, % EModulExe, % WorkPath,, EModulPID
				If Instr(EModulExe, "Addendum_Exporter") {
					WinWait         	, % (DokExClass := "Dokumentexport ahk_class AutoHotkeyGUI"),, 2
					Sleep 1000
					WinActivate   	, % DokExClass
					WinWaitActive	, % DokExClass,, 7
					while !(EModulHwnd := WinExist(DokExClass)) && (A_Index < 20)
						Sleep 100
					If EModulHwnd
						Send_WM_CopyData("search|" AlbisAktuellePatID() "|" GetAddendumID() "|admGui_Callback", EModulHwnd)
				}
			}

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

return ;}

adm_Reports:                                                                                    	;{ Patienten Tab 	- Listview Handler

		If InStr(A_GuiControl, "admReports") && (A_EventInfo = 0)
			return

	; zurück wenn Datei nicht existiert
		pdfPath := ""
		LV_GetText(admFile, EventInfo := A_EventInfo, 1)
		For idx, file in PatDocs
			If InStr(file.name, admFile)
				break

		If !FileExist(Addendum.BefundOrdner "\" file.name) 	{
			PraxTT(	"Die Dateioperation ist nicht möglich,`nda die Datei nicht mehr vorhanden ist.", "3 2")
			admGui_Reports()
			RedrawWindow(hadm)
			return
		}

	;PDF Vorschau
		If Instr(A_GuiEvent, "DoubleClick")
			admGui_View(file.name)

return ;}

adm_Journal:                                                                                     	;{ Journal Tab

	; Default Listview, gewählter Dateiname
		If (A_EventInfo = EventInfo)
			return
		EventInfo	:= A_EventInfo
		admFile 	:= Journal.GetSelected()

	; nur ausgewählte Importieren
		If !Addendum.ImportFromJrnl {
			If IsObject(admFile)
				GuiControl, adm: , admButton2, % "Importieren [" admFile.Count() "]"
			else
				GuiControl, adm: , admButton2, % "Importieren"
		}

	; Kontextmenu
		If InStr(A_GuiEvent	, "RightClick") && (StrLen(admFile) > 0) && !IsObject(admFile)	{
			If !Addendum.ImportFromJrnl
				Journal.ShowMenu(admFile)
			return
		}
	; Spaltensortierung
		else 	If 	InStr(A_GuiEvent	, "ColClick")    	{
			If (EventInfo > 0) && !Addendum.ImportFromJrnl
				admGui_Sort(EventInfo)
			return
		}
	; Dokumente importieren
		else 	If 	InStr(A_GuiControl, "admButton2")	{
			EventInfo := ""
			If !Addendum.ImportFromJrnl
				admGui_ImportJournalAll(admFile)
			return
		}
	; Befundordner indizieren
		else 	If	InStr(A_GuiControl, "admButton3") {
			EventInfo := ""
			If !Addendum.ImportFromJrnl
				admGui_Reload()
			return
		}
	; OCR starten/abbbrechen
		else 	If	InStr(A_GuiControl, "admButton4") {
			EventInfo := ""
			If Addendum.Thread["tessOCR"].ahkReady() {
				MsgBox, 4	, Addendum für Albis on Windows, % "Soll die laufende Texterkennung`nabgebrochen werden?"
				IfMsgBox, No
					return
				Addendum.Thread["tessOCR"].ahkTerminate[]
				admGui_OCRButton("+OCR ausführen")
				return
			}
			else {
				admGui_OCRAllFiles()
			}
		}

	; zurück wenn Datei nicht existiert
		EventInfo := ""
		If !FileExist(Addendum.BefundOrdner "\" admFile) || IsObject(admFile)
			return

	; PDF/Bild-Programm aufrufen
		If Instr(A_GuiEvent	, "DoubleClick")
			admGui_View(admFile)

return ;}

adm_TP:                                                                                             	;{ Protokoll Tab 	- Button Handler

	GuiControlGet, day, adm:, AdmTPTag

	If         	InStr(A_GuiControl, "admTPZurueck")	             	; einen Tag zurück
			admGui_TProtokoll(day, -1, compname)

	else If	InStr(A_GuiControl, "admTPVor")    	             	; einen Tag weiter

		If (day = "Heute")
			return
		else
			admGui_TProtokoll(day, 1, compname)

	else if InStr(A_GuiControl, "admTPTag")

		If (day = "Heute")
			return
		else
			admGui_TProtokoll("Heute", 0, compname)

return ;}

adm_TProtLV:                                                                                    	;{ Tagesprotokoll 	- Listview Handler

	admGui_Default("admTProtokoll")

	If Instr(A_GuiEvent, "DoubleClick") {
		LV_GetText(PatID, A_EventInfo, 5)
		AlbisAkteOeffnen("", PatID)
	}

return ;}

adm_GuiDropFiles:                                                                             	;{ Dokumente per Drag-Drop hinzufügen

		PatName 	:= AlbisCurrentPatient()

	; übergebene Dateien liegen durch ein mit `n getrennte Liste in A_GuiEvent vor (genial einfach))
		Loop, Parse, A_GuiEvent, `n
		{

			; Datei kommt direkt aus dem Befundordner, nicht aufnehmen (, "^" EscapeStrRegEx(Addendum.BefundOrdner) "\\[^\\]+$")
				SplitPath, A_LoopField, filename, filepath
				If (A_LoopField = (Addendum.BefundOrdner "\" filename))
					continue

			; benennt die Datei gleich mit dem Namen des aktuellen Patienten, wenn man die Datei auf den Patienten Tab gezogen hatte
				If InStr(A_GuiControl, "admReports") {
					filename := StrReplace(filename, PatName)
					filename := PatName ", " RegExReplace(filename, "^\s*,\s*")
				}

			; Datei sicherheitshalber in den Befund und Backup Ordner kopieren
				If !FileExist(Addendum.BefundOrdner "\" filename){
					FileCopy, % A_LoopField, % Addendum.BefundOrdner "\" filename
					FileCopy, % A_LoopField, % Addendum.BefundOrdner "\Backup\" filename, 1
				}

			; in den Listviews anzeigen
				displayname := string.Replace.Names(filename)
				If RegExMatch(filename, "\.pdf$")		{
					PatDocs.Push(StrReplace(filename, ".pdf"))
					LV_Add("ICON" 1,, displayname)
				}
				else If RegExMatch(filename, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{
					PatDocs.Push(filename)
					LV_Add("ICON" 2,, displayname)
				}

		}

	; Importbuttons aktivieren und später nach Befundart ausschalten
		admGui_InfoText("Patient")
		admGui_InfoText("Journal")

return ;}

adm_BefundImport:                                                                              	;{ läuft wenn Befunde oder Bilder importiert werden sollen

		If (PatDocs.MaxIndex() = 0) || Addendum.ImportFromPat 	; versehentlichen Aufruf bei leerem Array verhindern
			return

		PreReaderListe := admGui_GetAllPdfWin()              	; derzeit geöffnete Readerfenster einlesen
		admGui_ImportGui(true, "...importiere alle Befunde")	; Hinweisfenster anzeigen
		Imports := admGui_ImportFromPatient()                   	; Importvorgang starten
		admGui_InfoText("Patient")                                      	; Kurzinfo aktualisieren
		admGui_InfoText("Journal")                                      	; Kurzinfo aktualisieren
		admGui_ImportGui(false)                                        	; Hinweisfenster wieder schliessen
		admGui_Journal(false)                                             	; Journalinhalt auffrischen
		admGui_ShowPdfWin(PreReaderListe)                     	; holt den/die PdfReaderfenster in den Vordergrund

return ;}

adm_LBDruck:                                                                                     	;{ Laborblattdruck mit einem Tastendruck

	Gui, adm: Submit, NoHide

	; sichert die Einstellung
		If (A_GuiControl = "admLBDPD") {
			If (Addendum.iWin.LBDrucker <> admLBDPD)
				IniWrite, % admLBDPD, % Addendum.Ini, % compname, % "Infofenster_Laborblatt_Drucker"
			Addendum.iWin.LBDrucker := admLBDPD
		}
	; wenn Zähler 0 zeigt "Alles" anzeigen
		else If (A_GuiControl = "admLBDUD") {
			If (admLBDUD < 1) {
				GuiControl, adm:, admLBDCB, % "Alles"
				Sleep 50
				GuiControl, adm:, admLBDCB, % "Alles"
			}
			GuiControl, adm: Focus, admLBDUD
		}
	; Drucken ausführen
		else If (A_GuiControl = "admLBD")
			admGui_LaborblattDruck(admLBDPD, admLBDCB, admLBCk)
	; Mailversand
		else If (A_GuiControl = "admLBMail") {
			If InStr(DDLPrinter, "Microsoft Print to PDF") {
				If !(savePath 	:= admGui_LaborblattDruck("Microsoft Print to PDF", admLBDCB, admLBCk)) {
					PraxTT("Das Laborblatt konnte nicht als PDF exportiert werden.`nDie Funktion bricht hier ab.")
					return
				}
				RegExMatch(IFilter(savePath), "geb\.\s+[\d\.]+\s*[\n\r\s]+(?<Datum>[\d\.]+)", Laborblatt)
				PraxTT("erstelle Outlook-EMail ...", "40 1")
				MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
				MailItem.Recipients.Add(oPat[AlbisAktuellePatID()].EMAIL)
				MailItem.attachments.Add(savePath)
				MailItem.Subject := "Ihre Laboruntersuchung vom " LaborblattDatum
				body := "Sehr geehrte " (oPat[AlbisAktuellePatID()].Gt = "w" ? " Frau " : "r Herr ") oPat[AlbisAktuellePatID()].Vn " " oPat[AlbisAktuellePatID()].Nn ",`n`n"
				body .= "wie besprochen sende ich Ihnen die Ergebnisse Ihrer Laboruntersuchung vom " LaborblattDatum " zu."
				body .= "`n`n`nMit freundlichen Grüßen`n" StrReplace(Addendum.Praxis.MailStempel, "##", "`n")
				MailItem.body := body
				MailItem.display
				PraxTT("", "off")
			}
			else
				PraxTT("Funktion benötigt den PDF-Druckertreiber`nMicrosoft Print to PDF", "4 1")
		}
	; Checkbox Status sichern
		else if (A_GuiControl = "admLBCk") {
			Addendum.iWin.LBAnm := admLBCk
		}

return ;}
;}

; -------- Gui-Funktionen                                                                     	;{
admGui_Default(LVName)                                             	{               	; Gui Listview als Default setzen

	Gui, adm: Default
	Gui, adm: ListView, % LVName

return
}

admGui_ShowTab(TabName, hTab="")                          	{               	; ein bestimmten Tab nach vorne holen

		global admHTab, hadm
		static admGuiTabs := {"Patient":0, "Journal":1, "Protokoll":2, "Extras":3, "Netzwerk":4, "Info":5}

		TabToShow := (RegExMatch(TabName, "^\d+$") ? TabName : admGuiTabs[TabName])
		hTab := !hTab ? admHTab : hTab

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
		For CurrentTabName, tabnr in admGuiTabs
			If (CurrentTab = tabnr)
				return CurrentTabName

return ErrorLevel
}

admGui_Sort(EventNr, LV_Init=false)                              	{                 	; sortiert die Journalspalten und zeigt ein Symbol für die Sortierreihenfolge an

		; Funktion wird gebraucht für die Wiederherstellung der letzten Sortierung und für das Sichern der Einstellungen bei Nutzerinteraktion
		; LV_Init - nutzen um gespeicherte Sortierungseinstellung wiederherzustellen

		global 	admHJournal, hadm
		static 	admJCols, LVSortStr, JSort, JColDir, JRow, JSortDir := []

		admGui_Default("admJournal")

		If LV_Init {

		  ; Spalte "3" , Sortierrichtung 1 = Aufsteigend : 0 = Absteigend, Listview scrollen zur Reihe Nr.
			Addendum.iWin.JournalSort:= IniReadExt(compname, "Infofenster_JournalSortierung", "3 1 1")

   		  ; LVM_GETHEADER = LVM_FIRST (0x1000) + 31 = 0x101F.
			admHhdr	:= DllCall("SendMessage", "uint", admHJournal	, "uint", 0x101F, "uint", 0, "uint", 0)
		  ; HDM_GETITEMCOUNT = HDM_FIRST (0x1200) + 0 = 0x1200.
			amdCols	:= DllCall("SendMessage", "uint", admHhdr     	, "uint", 0x1200, "uint", 0, "uint", 0)

			Loop % admJCols
				JSortDir[A_Index] := "0"

		}

		If LV_Init || (EventNr = 0) {
			RegExMatch(Addendum.iWin.JournalSort, "\s*(\d+)\s*(\d)\s*(\d+)", S)
			EventNr            	:= S1= 0 	? 1 : S1
			JSortDir[EventNr] 	:= S2= ""	? 1 : S2
			JRow                 	:= S3
		}

	; Sortierung je nach gewählter Spalte vornehmen
	; Idee von: https://www.autohotkey.com/boards/viewtopic.php?t=68777
		If (EventNr = 3)     ; Spalte 3 (Eingangsdatuim) - sortiert wird nach der unsichtbaren Spalte 4 (Timestamp)
			EventNr := 4

		admGui_ColSort("admJournal", EventNr, (JSortDir[EventNr] := !JSortDir[EventNr]))
		Addendum.iWin.JournalSort := EventNr " " (JSortDir[EventNr])

}

admGui_ColSort(LVName, col, SortDest)                         	{                	; Helfer für admGui_Sort

	global 	admHJournal

	admGui_Default(LVName)

	If (LVName = "admJournal")
		If (col = 2)
			Type := "Integer"
		else If (col = 3)
			col := 4, Type := "Integer"
		else
			Type := "Text"

	LV_ModifyCol(col, (SortDest ? "Sort" : "SortDesc") " " Type)
	LV_SortArrow(admHJournal, col, (SortDest ? "up":"down"))

}

admGui_ColorRow(hLV, rowNr, paint=true)                      	{                	; eine Zeile einfärben

	global 	hadm
	static 	bcolor := 0x995555, tcolor := 0xFFFFFF

	If !Addendum.iWin.RowColors
		return

	If paint
		res := LV_Colors.Row(hLV, rowNr, bcolor, tcolor)
	else
		res := LV_Colors.Row(hLV, rowNr)

	;SciTEOutput("ColorRow: " res)
	;GuiControl, adm: +Redraw, admJournal
	WinSet, Redraw, , % "ahk_id " hadm

}

admGui_OCRButton(status)                                           	{                	; ändert den Text und Modus des OCR Buttons im Journal Tab

	global admButton4, adm

	RegExMatch(status " ", "i)^\s*(?<Enabled>[\+\-])\s*(?<Text>[\pL\-\s]+)", OCR_)

	If (OCR_Enabled = "+")
		OCRStatus := "Enable1"
	else if (OCR_Enabled = "-")
		OCRStatus := "Enable0"
	else
		OCRStatus := "Enable1"

	admGui_Default("admJournal")

	GuiControl, % "adm: " 	OCRStatus 	, admButton4
	GuiControl, % "adm: ",	admButton4	, % OCR_Text

}

admGui_Receive(answer)                                               	{                 	; empfängt Netzwerknachrichten
	SciTEOutput(answer)
}

admGui_Destroy()                                                             {                	; schliesst das Infofenster

	global admHJournal, admHTab

	Gui, adm: Submit, Hide
	GuiControlGet, ProtokollTag, adm:, admTPTag

	Addendum.TProtDate       	:= ProtokollTag
	Addendum.iWin.firstTab	:= admTabs
	Addendum.iWin.lastPatID	:= 0

	Gui, adm: 	Destroy
	Gui, adm2: 	Destroy

	hadm := hadm2 := 0

return
}
;}

; -------- Inhalte erstellen                                                                      	;{
admGui_Reload(refreshDocs=true)                                  	{                	; aktualisiert die Inhalte und zeichnet das Fenster neu

	global hadm, PatDocs

	admGui_Journal(refreshDocs)
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

	If RegExMatch(LCall1, "(?<Y>\d{4})-(?<M>\d{2})-(?<D>\d{2})\s+(?<H>\d{2}):(?<Min>\d{2}):(?<S>\d{2})[\s\|]*", T)
		lastLBCall := "letzter:  " TD "." TM (TY = A_YYYY ? "." : "." TY) " " TH ":" TMin "Uhr`nDaten: " (LCall1 = LCall2 ? "[-]" : LCall1 = LCall3 ? "[+]" : "[?]")

	IniRead, nextCall, % Addendum.Ini, % "LaborAbruf", % "naechster_Abruf"
	If RegExMatch(nextCall, "(?<D>\d{2})\.(?<M>\d{2})\.(?<Y>\d{4})", T)
		nextLBCall := "nächster: " (nextCall ? (TY = A_YYYY ? RegExReplace(nextCall, "\." TY "\,\s*", ". ") : nextCall) "Uhr" : "")

	GuiControl, adm:, admILAB1, % lastLBCall
	GuiControl, adm:, admILAB2, % nextLBCall

return
}
;}

; -------- Gui Tabs                                                                              	;{
admGui_Reports()                                                         	{	            	; zeigt Befunde des Patienten aus dem Befundordner

		global 	PatDocs
		static 	PdfImport 		:= false
		static 	ImageImport	:= false
		static 	rxPerson1    	:= "[A-ZÄÖÜ][\pL]+(\-[A-ZÄÖÜ][\pL-]+)*"
		static 	rxPerson2    	:= "[A-ZÄÖÜ][\pL]+([\-\s][A-ZÄÖÜ][\pL]+)*"

		admGui_Default("admReports")
		LV_Delete()

	; Pdf Befunde des Patienten ermitteln und entfernen bestimmter Zeichen aus dem Patientennamen für die fuzzy Suchfunktion
		PatDocs	:= Array()
		PatID    	:= AlbisAktuellePatID()
		PatNV   	:= RegExReplace(oPat[PatID].Nn, "[\s\-]") . RegExReplace(oPat[PatID].Vn, "[\s\-]")

	; PatDocs erstellen - enthält nur die Befunde/Bilder des aktuellen Patienten
		For key, pdf in ScanPool	{				;wenn keine PatID vorhanden ist, dann ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)
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
			displayname := string.Replace.Names(string.Replace.FileExt(pdf.name))
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
					LV_Add("ICON2",,  string.Replace.Names(A_LoopFileName)) ; entfernt den Namen
					PatDocs.Push({"name": A_LoopFileName})
				}
			}

	; InfoText auffrischen
		admGui_InfoText("Journal")

return PatDocs
}

admGui_Journal(refreshDocs=false)                              	{	            	; befüllt das Journal mit Pdf und Bildbefunden aus dem Befundordner

		global admHJournal
		static FirstRun := true

		admGui_Default("admJournal")
		LV_Delete()

	; SCANPOOL OBJECT AUFFRISCHEN
		If (refreshDocs || FirstRun) {
			FirstRun := false
			pdfpool.Update((refreshDocs?true:false), true, Addendum.DBPath "\sonstiges")  ; (ReIndex = false, save = true)
		} else
			pdfpool.RemoveNoFile()

	; PDF BEFUNDE DES PATIENTEN HINZUFÜGEN
		admGui_OCRButton("-OCR ausführen")
		Addendum.iWin.OCREnabled := false
		For docID, pdf in ScanPool	{
			LV_Add((pdf.isSearchable ? "ICON3" : "ICON1"), pdf.name, pdf.pages, pdf.filetime, pdf.timeStamp) 	; anderes Symbol für durchsuchbare PDF Dateien
			If !Addendum.iWin.OCREnabled && !pdf.isSearchable
				Addendum.iWin.OCREnabled := true
		}

	; OCR BUTTON
		If Addendum.iWin.OCREnabled {
			If Addendum.Thread["tessOCR"].ahkReady()
				admGui_OCRButton("+OCR abbrechen")
			else
				admGui_OCRButton("+OCR ausführen")
		}

	; BILD-DOKUMENTE HINZUFÜGEN
		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(Trim(A_LoopFileLongPath), "\.(jpg|png|tiff|bmp|wav|mov|avi)$") {
				FileGetTime, timeStamp, % A_LoopFileLongPath, C
				FormatTime, filetime  	, % timeStamp, dd.MM.yyyy
				FormatTime, timeStamp	, % timeStamp, yyyyMMdd
				SplitPath, A_LoopFileLongPath, BefundName
				LV_Add("ICON2", BefundName,, filetime, timeStamp)
			}

	; INFOTEXT AUFFRISCHEN
		admGui_InfoText("Journal")

	; SCROLLT DAS ZULETZT GESPEICHERTE ELEMENT IN SICHT
		;~ IWin :=  Addendum.iWin
		;~ If (IWin.firstTab = "Protokoll") && (IWin.firstTabPos > 0)
			;~ SendMessage, 0x1013, % IWin.firstTabPos, 0,, % "ahk_id " admHJournal

	; ANZEIGE WIRD MIT DER GESPEICHERTEN SPALTENSORTIERUNG ANGEZEIGT (Addendum.ini)
		admGui_Sort(0, true)

return
}

admGui_TProtokoll(datestring, addDays, client)              	{                	; listet alle Karteikarten übergebenen Datums auf

	global admHTProtokoll

	admGui_Default("admTProtokoll")

	Addendum.iWin.TProtDate := datestring
	datestring := datestring = "Heute" || StrLen(datestring) = 0 ? (A_DD "." A_MM "." A_YYYY) : (datestring)
	RegExMatch( datestring, "(?<dd>\d+)\.(?<MM>\d+)\.(?<YYYY>\d+)", d)
	lDate := dYYYY dMM ddd
	If (addDays != 0)
		lDate += addDays, Days

	RegExMatch(ldate, "^(?<YYYY>\d{4})(?<MM>\d{2})(?<dd>\d{2})", d)
	FormatTime, weekdayshort, % lDate, ddd
	FormatTime, weekdaylong	, % lDate, dddd
	FormatTime, month       	, % lDate, MMMM

	If !FileExist(TPFullPath	:= Addendum.DBPath "\Tagesprotokolle\" dYYYY "\" dMM "-" month "_TP.txt")
		return

	IniRead, TProto, % TPFullPath, % (SectionDate := weekdaylong "|" ddd "." dMM), % client
	If InStr(TProto, "Error")
		TProto := ""
	else
		TProto := InStr(TProto, ",") ? StrSplit(RTrim(TProto, ","), ",") : StrSplit(TProto, ">", "<")

	LV_Delete()
	For idx, PatIDTime in TProto {
		RegExMatch(PatIDTime, "^\d+", PatID)
		If (StrLen(PatID) = 0)
			continue
		LV_Insert(1,, SubStr("00" idx, -1)                                          	; Index absteigend
						 , oPat[PatID].Nn ", " oPat[PatID].Vn                    	; Name
						 , oPat[PatID].Gd						                    		; Geburtsdatum
						 , Age(oPat[PatID].Gd, ddd "." dMM "." dYYYY)		; Alter an diesem Tag
						 , PatID)						                    					; Patientennummer
	}

	LV_ModifyCol(2, "Auto")
	If (LV_GetColWidth(admHTProtokoll, 2) < 115)
		LV_ModifyCol(2, 115)

	GuiControl, adm: , AdmTProtokollTitel	, % "[" compname "] [" (TProto.Count()-1 < 0 ? 0 : TProto.Count()-1) " Pat.]"
	If (SubStr(lDate, 1, 8) = SubStr(A_Now, 1, 8))
		GuiControl, adm: , AdmTPTag          	, % "Heute"
	else
		GuiControl, adm: , AdmTPTag          	, % weekdayshort ", " ddd "." dMM "." dYYYY

}

admGui_InfoText(TabTitel)                                             	{               	; Zusammenfassungen aktualisieren

	; letzte Änderung: 17.05.2021

	global adm, admButton1, admButton2
	global admJournalTitle, admPatientTitle, admTProtokollTitel

	Gui, adm: Default

	If (TabTitel = "Journal") || (TabTitel = "Patient")	{
		Journal.InfoText()
		Journal.ShowImportStatus(false)
	}
	else If InStr(TabTitel, "TProtokoll") 	{
		InfoText := "[" compname "] [" TProtokoll.MaxIndex() " Patienten]"
		GuiControl, adm: , admTProtokollTitel, % InfoText
	}

}
;}

; -------- Kontextmenu                                                                        	;{
admGui_CM(MenuName)                                               	{                 	; Kontextmenu des Journal

		; fehlt: nachschauen ob im Befundordner ein Backup-Verzeichnis angelegt ist
		; auch für Tastaturkürzelbefehle in allen Tabs (01.07.2020)

		global 	admHJournal, admHReports, admFile, rcEvInfo, hadm, PatDocs
		static 	newadmFile, rowNr

		Addendum.iWin.firstTab	:= "Journal"
		If RegExMatch(MenuName, "^J")
			admGui_Default("admJournal")

		blockthis 	:= false
		rowSel  	:= Journal.getNext(1, admFile)
		For key, inprogress in Addendum.iWin.FilesStack
			If (inprogress = admFile) {
				blockthis := true
				break
			}

		; Callback nicht ausführen wenn ein Menu ausgeführt wird
			Addendum.PopUpMenuCallback := ""

	; Menupunkte ohne PDF Schreibzugriffe auf PDF Datei
		If      	InStr(MenuName, "JRefresh")	{	; Listview aktualisieren
			admGui_Reload()
			return
		}
		else if	InStr(MenuName, "JOpen") 	{	; Listview auffrischen
			If blockthis {
				PraxTT("Die Datei wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			FuzzyKarteikarte(admFile)
			return
		}
		else if	InStr(MenuName, "JView")  	{	; PDF mit dem Standard PDF Anzeigeprogramm öffnen
			If blockthis {
				PraxTT("Die wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_View(admFile, MenuName)
			return
		}

	; zurück wenn der Schreibzugriff gesperrt ist
		If FileIsLocked(Addendum.BefundOrdner "\" admFile) 	{
			PraxTT(	"Die Dateioperation ist nicht möglich,`nda ein anderes Programm die Datei sperrt.", "5 2")
			return
		}

	; Menupunkte mit Schreibzugriff auf PDF Datei
		If      	InStr(MenuName, "JDelete")	{	; Datei löschen

			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}

		; Nutzer fragen
			MsgBox, 0x1024, Addendum für AlbisOnWindows, % "Wollen Sie die Datei:`n" admFile "`nwirklich löschen?"
			IfMsgBox, No
				return

		; erstellt Backups und löscht die Original und zugehörigen Dateien
			res := Befunde.Remove(admFile)
			res := PDFpool.Remove(admFile)
			res := Journal.Remove(admFile)

		}
		else if	InStr(MenuName, "JOCR")  	{	; OCR einer Datei

			; Abbruch wenn gerade ein OCR Vorgang läuft
				If Addendum.Thread["tessOCR"].ahkReady() {
					PraxTT("Es läuft gerade noch ein Texterkennungsvorgang!`n...bitte warten...", "2 0")
					return
				}
				else If blockthis {
					PraxTT("Diese Datei ist gerade in Bearbeitung.`n...bitte warten...", "1 0")
					return
				}

			; Nutzerhinweise anzeigen
				If !PDFisSearchable(Addendum.BefundOrdner "\" admFile) {
					PraxTT("PDF [" admFile "] bisher ohne Texterkennung. Starte Tesseract..", "1 0")
					Sleep 1000
				}
				else                                                                                 	{
					MsgBox, 4	, Addendum für Albis on Windows
									, %	">> OCR Text vorhanden <<`n" admFile "`n"
										. 	"Bei erneuter Ausführung werden die bisherigen Textdaten gelöscht.`n"
										.	"Dennoch ausführen?"
					IfMsgBox, No
						return
					PraxTT("Die Texterkennung der Datei < " admFile " > wird jetzt ausgeführt.", "1 0")
					Sleep 500
				}

			; OCR Vorgang anzeigen
				Addendum.iWin.FilesStack.InsertAt(1, admFile)
				admGui_OCRButton("+OCR abbrechen")

			; OCR Starten als echter Thread, nach Abschluß eines Vorgangs schickt der Thread eine Nachricht ans Addendum-Skript
				TesseractOCR(admfile)

			return
		}
		else if	InStr(MenuName, "JOCRAll")	{	; OCR aller offenen Dateien
			If Addendum.Thread["tessOCR"].ahkReady() || (Addendum.iWin.FilesStack.Count() > 0) {
				PraxTT("Ein Texterkennungsvorgang läuft im Moment noch.`n...bitte abwarten...", "3 0")
				return
			}
			admGui_OCRAllFiles()
			return
		}
		else if	InStr(MenuName, "JRecog1")	{	; Datei automatisch kategorisieren
			If blockthis {
				PraxTT("Diese Datei ist gerade in Bearbeitung.`n...bitte warten...", "1 0")
				return
			}
			admCM_JRecog(admFile)
			return
		}
		else if	InStr(MenuName, "JRecog2")	{	; alle Dateien kategorisieren
			If Addendum.Thread["tessOCR"].ahkReady() || (Addendum.iWin.FilesStack.Count() > 0) {
				PraxTT("Ein Texterkennungsvorgang läuft im Moment noch.`n...bitte abwarten...", "1 0")
				return
			}
			admCM_JRecogAll()
			return
		}
		else if	InStr(MenuName, "JRename")	{	; manuelles Umbenennen
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile)
			return
		}
		else if	InStr(MenuName, "JSplit")   	{	; manuelles Aufteilen einer PDF Datei
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile	,	"Aufteilen der PDF-Seiten auf mehrere Datein.`n"
													.	"Schreiben Sie: 1-2 D1, 3-4 D2`n"
													.	"Datei 1 erhält die Seiten 1-2 und Datei 2 die Seiten 3-4`n" )
			return
		}
		else if	InStr(MenuName, "JImport")	{	; Datei in geöffnete Karteikarte importieren
			If blockthis {
				PraxTT("Die wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_ImportFromJournal(admFile)
			return
		}

return
}
;}

; -------- Autonaming                                                                         	;{
admCM_JRecog(admFile)                                              	{              	; Kontextmenu: JRecog - Dateibezeichnung aus PDF Text erstellen

	; Pfade
		pdfPath	:= Addendum.Befundordner "\" admFile
		txtPath	:= Addendum.Befundordner "\Text\"  StrReplace(admFile, ".pdf", ".txt")

	; zurück wenn PDF keinen Textlayer hat
		If !PDFisSearchable(pdfPath) {
			PraxTT("PDF Datei enthält keinen Textlayer.`nFühren Sie zuerst eine Texterkennung durch!", "2 1")
			return
		}

	; Textlayer der PDF extrahieren und laden
		If !FileExist(txtPath) {

			; PDF to Text per commandline Befehl falls kein Text extrahiert wurde
				stdout := StdOutToVar(	Addendum.PDF.xpdfPath "\pdftotext.exe"
												.	" -f 1 -l " PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
												. 	" -bom -enc"
												. 	" " q "UTF-8" 	q
												.	" " q pdfPath 	q
												. 	" " q txtPath 	q)

				If !FileExist(txtPath) {
					PraxTT("Es konnte kein Text extrahiert werden!`n[" admFile "]", "2 1")
					return
				}

		}

	; Textinhalt lesen
		PdfText := FileOpen(txtPath, "r").Read()
		If (StrLen(Trim(PdfText)) = 0) {
			PraxTT("Der extrahierte Text enthält keine Zeichen.`n[" admFile "]", "2 1")
			FileDelete, % txtPath
			return
		}

	; ImportGui: Anzeigen von Wort und Zeichenzahl der Datei
		outmsg := "[Textanalyse:`t" 	admFile "]"                                                	"`n"
					. 	"Wörter:        `t"   	StrSplit(PdfText, " ").MaxIndex() 						"`n"
					. 	"Zeichenzahl:`t"   	StrLen(RegExReplace(PdfText, "[\s\n\r\f]"))
		admGui_ImportGui(true, outmsg, "admJournal")

	; RegEx Auswertungen
		fNames	:= FindDocNames(PdfText)
		fDates 	:= FindDocDate(PdfText, fNames, true)
		admGui_ImportGui(true, "Autonaming Dokument`n" admFile "`nläuft", "admJournal")

	; neuen Dateinamen erstellen
		If IsObject(fNames)  {

			; Patientenname als erstes
				if (fNames.Count() = 1) {
					For PatID, Pat in fNames {
						newfilename := fNames[PatID].Nn ", " fNames[PatID].Vn ", "
						break
					}

			; als letztes das gefundene Dokumentdatum
				If IsObject(fDates)
					newfilename .= "v. " (fDates.Behandlung[1] ? fDates.Behandlung[1] : fDates.Dokument[1])

			; wird umbenannt
				admGui_Rename(admFile,, newfilename ".pdf")
				return newfilename

			}
			else {
				For PatID, Pat in fNames
					t .= "   [" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"
				SciTEOutput("  Das Dokument konnte nicht eindeutig zugeordnet werden. Folgende Namen stehen zu Auswahl:`n" t)
			}
		}
		else {

			PraxTT("Dokument: " admFile "`nkonnte keinem Patienten zugeordnet werden", "4 1")

		}

return
}

admCM_JRecogAll(files:="")                                            	{                	; Kontextmenu: automatisiert die Benennung von PDF-Dateien

	; letzte Änderung 17.06.2021:	kann mit Parameter gestartet werden, files sollte eine Array mit Dateinamen aus ihrem Befundordner sein
	;                                             	geplant automatisches Umbenennen nach dem Tesseract die Texterkennung beendet hat

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

			MsgBox, 0x2003, Addendum für AlbisOnWindows, Sollen nur bisher unbenannte Dokumente verarbeitet werden?
			IfMsgBox, Cancel
			{
				PraxTT("Abbruch durch Nutzer", "1 1")
				return
			}
			IfMsgBox, Yes
				RecognitionDocs := pdfpool.GetUnamedDocuments()
			IfMsgBox, No
				RecognitionDocs := pdfpool.GetAllDocNames()

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
		If !IsObject(files) {
			RecogEnd := false
			Gui, Recog: New   	, -Caption +ToolWindow -DPIScale AlwaysOnTop HWNDhRecog
			Gui, Recog: Margin	, 0, 0
			Gui, Recog: Color  	, cEFF0F1
			Gui, Recog: Font   	, % "s10 cWhite"                                                                  	, % Addendum.Default.Font
			Gui, Recog: Add    	, Text	, x2 y2                                                                     	, % "Autonaming Datei: "
			Gui, Recog: Add    	, Text	, x+2 vRecogCounter                                            	, % prefix "1/" RecognitionDocs.Count()
			Gui, Recog: Font   	, % "s8 cWhite"                                                                  	, % Addendum.Default.Font
			Gui, Recog: Add    	, Button, x2 y+5 vRecogCancel gRecogHandler                    	, % "Abbruch"
			Gui, Recog: Show    	, % "AutoSize Hide NoActivate"                                             	, % "Addendum Autonaming"

			aw:= GetWindowSpot(AlbisWinID())
			rw	:= GetWindowSpot(admHRecog)
			Gui, Recog: Show     	, % "x" aw.X " y" aw.CH - rw.H " NoActivate"
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

			; Debug
				If !IsObject(files)
					GuiControl, Recog:, RecogCounter, % counter
				If (idx > 1) && (Weitermachen("Datei: " counter "`nBez.: " filename "`nWeitermachen?",, 2) = 0)
				 break

			; Text extrahieren, falls noch nicht geschehen
				If PDFisSearchable(Path.fF) && !FileExist(path.fT) {

					; Fortschritt anzeigen
						;SciTEOutput("Textpath: " path.fT)
						If !IsObject(files)
							admCM_JRnProgress([counter, filename, "Text wird extrahiert....", "---", "---", lastnewadmFile])

					; pdftotext extrahiert den Text aus der PDF Datei
						stdout := StdOutToVar(Addendum.PDF.xpdfPath "\pdftotext.exe"
													                     	. 	" -f 1 -l " PDFGetPages(Path.fF, Addendum.PDF.qpdfPath)
													                     	. 	" -bom -enc UTF-8"
													                     	.	" " q Path.fF q
													                     	. 	" " q Path.fT q)
						If !FileExist(Path.fT) {
							PraxTT("Datei [" counter "] " filename "`nTextextraktion fehlgeschlagen", "2 1" )
							continue
						}

				} else if !PDFisSearchable(Path.fF) {

					; Protokollierung
						PraxTT("Datei [" counter "] " filename " ist nicht durchsuchbar", "2 1" )
						Sleep 500
						continue
				}

			; extrahierten Text lesen
				doctext := FileOpen(Path.fT, "r").Read()
				If (StrLen(Trim(doctext)) = 0) {
					PraxTT("Der extrahierte Text enthält keine Zeichen.`n[" filename "]", "2 1")
					FileDelete, % Path.fT
					continue
				}

			; Fortschritt anzeigen
				If !IsObject(files)
					admCM_JRnProgress([ counter
													, filename
													, "Textinhalt wird analysiert...."
													, (words	:= StrSplit(doctext, " ").MaxIndex())
													, (chars	:= StrLen(RegExReplace(doctext, "[\s\n\r\f\v]")))
													, lastnewadmFile])

			; Patientennamen und Behandlungs- oder Dokumentdatum finden
				fNames	:= FindDocNames(doctext, false)
				fDates 	:= FindDocDate(doctext, fNames, false)
				fRecom	:= FindDocEmpfehlung(doctext)

			; Datei umbenennen
				If      	(IsObject(fNames) && fNames.Count() = 1) {

					; Name des Patienten
						newadmFile := ""
						For PatID, Pat in fNames {
							If (StrLen(Trim(Pat.Nn)) > 0 && StrLen(Trim(Pat.Vn)) > 0) { 				; Abbruch wenn keine Name erkannt wurde
								newadmFile := fNames[PatID].Nn ", " fNames[PatID].Vn ", "
								break
							}
						}
						If (StrLen(newadmFile) = 0)
							continue

					; Behandlungzeitraum oder Datum des Dokuments
						If IsObject(fDates)
							newadmFile .= "v. " (fDates.Behandlung[1] ? fDates.Behandlung[1]	: fDates.Dokument[1])

					; Anzeige von Wort und Zeichenzahl der Datei
						lastnewadmFile := newadmFile
						If !IsObject(files)
							admCM_JRnProgress([counter, filename, "neuer Dateiname erstellt", words, chars, lastnewadmFile])
						else
							SciTEOutput("  - Autonaming: " filename)

					; Umbenennen von Original, Backup und Textdatei
						If (StrLen(newadmFile) > 0) {
							admGui_FWPause(true, [filename  	, newadmfile], 12)    	; FolderWatch pausieren (6s bis zur Wiederherstellung)
							row	:= journal.Replace(filename		, newadmfile)        	; Journal auffrischen
							pdf	:= pdfpool.Rename(filename		, newadmfile)        	; Objektdaten ändern
							res	:= Befunde.Rename(filename		, newadmfile)        	; Dateien umbenennen (Original, Backup, Text)
							admGui_FWPause(false, [filename  	, newadmfile], 12)    	; FolderWatch pausieren (6s bis zur Wiederherstellung)
							OCRDone.Push(newadmfile)
							newadmFile := ""
						}

				}
				else if	(IsObject(fNames) && fNames.Count() < 1) {
					If !IsObject(files)
						PraxTT("Dokument: " filename "`nkonnte keinem Patienten zugeordnet werden", "1 1")
					else
						SciTEOutput("  - Autonaming: keine eindeutige Zuordnung möglich gewesen")
				}
				else If 	(IsObject(fNames) && fNames.Count() > 1) {

					t := ""
					For PatID, Pat in fNames
						t .= " `t[" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"

					SciTEOutput("  - [" counter "]")
					SciTEOutput("  - 1: " filename)
					SciTEOutput("  - 2: #keine eindeutige Zuordnung möglich")
					SciTEOutput("  - Namen:       `t" 	fNames.Count())
					SciTEOutput("  - Beh. Datum:`t" 	fDates.Behandlung.Count())
					SciTEOutput("  - Doc. Datum:`t" 	fDates.Dokument.Count())
					SciTEOutput(t)
					SciTEOutput("  -------------------------------------------------------------------" )

				}
				else {
					If !IsObject(files)
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
		If !IsObject(files) {
			admGui_ImportGui(false, "A U T O N A M I N G`n" filename "`nläuft", "admJournal")
			Gui, Recog: Destroy
		}

return OCRDone

RecogHandler: ;{

	If (A_GuiControl = "RecogCancel")
		RecogEnd := true

return ;}
}

admCM_JRnProgress(ProgressText)                                 	{              	; zeigt den Fortschritt beim Umbennen an

		static outmsgTitle := "A U T O N A M I N G ..#"
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

			admGui_ImportGui(true, outmsg, "admJournal")
			call := call = 5 ? 1 : call+1

		}

}
;}

; -------- Zusatzfunktionen                                                                     	;{
admGui_TPSuche()                                                                	{        	; wann war ein Patient da - nicht beendet

	TPFullPath	:= Addendum.DBPath "\Tagesprotokolle"
	Gui, adm: Submit, NoHide

	If (StrLen(admTPSuche) = 0)
		return

	;~ Loop, Files, *.txt, R
	;~ {


	;~ }

}

admGui_LaborblattDruck(Printer, Druckspalten, PrintAnnotation)	{        	; automatisiert den Laborblattdruck

		; automatisiert den Laborblattdruck und stellt die vorherige Ansicht wieder her

		static AlbisViewType

		AlbisViewType := AlbisGetActiveWindowType(true)

		If InStr(Printer, "Microsoft Print to PDF")
			savePath := Addendum.ExportOrdner "\Laborwerte-" RegExReplace(AlbisCurrentPatient(), "[\s]") ".pdf"
		else
			savePath := ""

		If !(resPrint := AlbisLaborblattExport(Druckspalten, savePath, Printer, PrintAnnotation))
			PraxTT("Der Ausdruck von Laborwerten ist fehlgeschlagen.`nFehlercode: " res, "4 1")

		If (AlbisViewType <> AlbisGetActiveWindowType(true))
			AlbisKarteikartenAnsicht(AlbisViewType)

return resPrint ? savePath : ""
}
;}

; -------- Dokument- und Listviewklassen                                                	;{
class Journal                                                                 	{                	; Gui-Handler des Journal Tab

		; Funktion:	verwaltet das Listview-Steuerelement des Journal und teilweise das Patient TAB
		;
		;					die Listview des Journal ist als Standard festgelegt, vor dem Bearbeiten der Listview des Patient TAB muss
		; 					die Standard Listvieweinstellungen der Gui über Default("admReports")

		GetCount()                              	{        	; Anzahl der Tabellenzeilen
			this.Default()
		return LV_GetCount()
		}

		GetNext(rLV:=0, LVNI:=2)       	{        	; findet die nächste ausgewählte Zeile
			global admJournal, adm, admHTab, admHJournal, admHReports
			; rLV = 1 based index of the starting row for the flag search. Omit or 0 to find first occurance specified flags.
			; LVNI_ALL := 0x0, LVNI_FOCUSED := 0x1, LVNI_SELECTED := 0x2
			hLV := this.DefaultHLV
			hLV := %hLV%
		return DllCall("SendMessage", "uint", hLV, "uint", 4108, "uint", rLV-1, "uint", LVNI) + 1
		}

		GetRow(fname, col:=1)            	{         	; Zeilennummer mit diesem Dateinamen

			row := 0
			this.Default()
			Loop % LV_GetCount() {
				LV_GetText(rText, A_Index, col)
				If (rText = fname) {
					row := A_Index
					break
				}
			}

		return row
		}

		GetSelected(col:=1)                 	{        	; alle ausgewählten Zeilen auslesen

			row 		:= 0
			JFiles	:= Array()

			this.Default()
			while (row := LV_GetNext(row)) {
				LV_GetText(fname, row, col)
				JFiles.Push(fname)
			}

			if (JFiles.MaxIndex() = 1)
				return JFiles[1]
			else if (JFiles.MaxIndex() > 1)
				return JFiles
			else
				return ""

		}

		GetText(row, col:=1)                	{        	; Text einer Zeile, Spalte auslesen
			this.Default()
			LV_GetText(rText, row, col)
		return rText
		}

		SelectRow(row)                       	{      	; eine Zeile auswählen

				global adm, admJournal, admReports, hadm

			; Deregistriere die gosub-Label Verknüpfung vor dem Auswählen einer Zeile
				this.gLabel(false)

			; es wurde ein Textstring anstatt einer Reihennummer übergeben
				If !RegExMatch(row, "^\d+$")
					row := this.GetRow(row)

			; Zeile auswählen, alle anderen werden abgewählt
				this.Default()
				this.Focus()
				Loop % this.GetCount()
					LV_Modify(A_Index, (A_Index = row ? "Select1 Focus Vis" : "-Select"))

			; Registriere das g-Label wieder
				this.gLabel(true)

		}

		Remove(fname:="", rLV:="")      {        	; entfernt eine Datei aus der Listview

			; ein Dateiname oder eine Zeilennummer können übergeben werden
			global admJournal, adm, admHTab, admHJournal

			this.Default()
			this.RedrawLV(false)
			row := fname ? this.getRow(fname) : rLV ? rLV : 0
			If (row > 0)
				res := LV_Delete(row)
			else
				res := -1
			this.RedrawLV(true)

		return res
		}

		Replace(sourcefile, targetfile)    	{         	; Dateinamen ändern

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

		ShowMenu(admFileObj)        	{        	; Kontextmenu anzeigen

			MouseGetPos, mx, my, mWin, mCtrl, 2
			If IsObject(admFileObj)
				Menu, admJCM1	, Show, % mx - 20, % my + 10
			else
				Menu, admJCM	, Show, % mx - 20, % my + 10

		return
		}

		ShowImportStatus(ImportRuns)  	{         	; Interaktionskontrolle der 'Importieren' Steuerelemente

			global admButton2, admButton1, adm, hadm, PatDocs

			; Patient Tab, Import Button
				this.Default("admReports")
				GuiControl, adm:, admButton1, % (ImportRuns ? "...importiere" : "Importieren")
				If ImportRuns
					GuiControl, % "adm: Disable", admButton1
				else
					GuiControl, % "adm: " (PatDocs.MaxIndex() > 0 ? "Enable" : "Disable"), admButton1

			; Journal Tab, Import Button
				this.Default("admJournal")
				GuiControl, adm:, admButton2, % (ImportRuns ? "...importiere" : "Importieren")
				If ImportRuns
					GuiControl, % "adm: Disable", admButton2
				else
					GuiControl, % "adm: " (this.GetCount()>0 ? "Enable" : "Disable"), admButton2

		}

		InfoText()                                  	{         	; Infotext (Journal/Patient) wird aufgefrischt

			global admButton1, adm, hadm, admJournalTitle, admPatientTitle

			Gui, adm: Default

			; Journal
			InfoText := (ScanPool.MaxIndex() = 0) ? " keine Dokumente" : (ScanPool.MaxIndex() = 1) ? "1 Dokument" : ScanPool.MaxIndex() " Dokumente"
			GuiControl, adm:, admJournalTitle, % InfoText

			; Patient
			InfoText := "Posteingang: (" (!PatDocs.MaxIndex() ? "keine neuen Befunde)" : PatDocs.MaxIndex() = 1 ? "1 Befund)" : PatDocs.MaxIndex() " Befunde)") ", insgesamt: " ScanPool.MaxIndex()
			GuiControl, adm:, admPatientTitle, % InfoText

		}

		Default(LVName="")                	{        	; Gui und Listview Default machen
			global admJournal, adm, admHTab, admHJournal, admReports
			static tabs := {"admJournal":"Journal", "admReports":"Patient"}, hLV := {"admJournal":"admHJournal", "admReports":"admHReports"}
			this.DefaultLV 	:= LVName ? LVName : this.DefaultLV ? this.DefaultLV : "admJournal"
			this.DefaultTab 	:= tabs[this.DefaultLV]
			this.DefaultHLV	:= hLV[this.DefaultLV]
			Gui, adm: Default
			Gui, adm: ListView, % this.DefaultLV
		}

		Focus()                                    	{        	; das Listview erhält den Eingabefokus
			global admJournal, adm, admHTab, admHJournal
			this.Default()
			GuiControl, adm: Focus, % this.DefaultLV
		return ErrorLevel
		}

		gLabel(register)                     	{       	; g-Label bereitstellen oder abschalten

			global adm, admJournal, hadm

			If register
				GuiControl, adm: +gadm_Journal	, % this.DefaultLV
			else
				GuiControl, adm: -g                  	, % this.DefaultLV

		}

		ListView(LVName)                    	{        	; nur das Listview Default machen
			global admJournal, adm, admHTab, admHJournal, admReports
			Gui, adm: ListView, % (LVName ? LVName : "admJournal")
		}

		RedrawLV(rdw)                         	{      	; Listview neuzeichnen an oder aus
			global admJournal, adm, admHTab, admHJournal, hadm
			GuiControl, % "adm: " (rdw?"+":"-") "Redraw", % this.DefaultLV
			EL := ErrorLevel
			If rdw
				RedrawWindow(hadm)
		return EL
		}

		Show()                                   	{        	; zeigt den Journal Tab an
			global admHTab
		return admGui_ShowTab(this.DefaultTab, admHTab)
		}

}

class Befunde                                                                	{               	; Datei-Handler

	; Funktion: verwaltet die physisch vorhandene Dateien auf der Festplatte
	; die Dateioperationen finden in einem Hauptordner mit 2 Unterordnern und jeweils einem Backup-Ordner statt.

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Objektdaten erhalten / setzen
		dbgStatus[]                                                            	{ 	; getter und setter
			get {
				return this.debug
			}
		}

		fPath[]                                                                    	{	; aufrufen um den Hauptpfad abzurufen
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
			eL 	:= []
			path	:= this.GetPaths(fname)

			If !FileExist(path.fB) {
				FileCopy, % path.fF, % path.fB
				ErrorLevel ? (eL.1 := A_LastError) : ""
				If this.debug && ErrorLevel
					SciTEOutput("Backup für <" fname "> konnte nicht erstellt werden")
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

		RegExRename(fname, rnPattern)                              	{	; RegEx umbenennen für Batch-Renaming
			If RegExMatch(rnPattern, "^(?<Options>.*)?\)(?<pattern>.*)$", rx)
				RegExMatch(fname, "o" )
			this.Rename(fname, newfile)
		}

		MoveToImports(fname)                                           	{	; Dateien in den Import-Ordner verschieben

			; es gibt keine Funktion zum Löschen von Dokumenten! Dokumente werden nur in einen Backup-Ordner verschoben.
			; Grund: bei Fehlern im Programmcode sind die Dokumente noch vorhanden. Der Backup-Ordner ist im Moment noch manuell zu leeren.

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

		Remove(fname)                                                      	{	; verschiebt eine Datei aus dem Befundordner in den Backup-Ordner

			; Backup der Originaldatei erstellen wenn noch nicht erfolgt
				path := this.Backup(fname)

			; Datei aus dem Befundordner nur löschen wenn ein Backup existiert
				If FileExist(path.fB) {
					FileDelete, % path.fF
					ErrorLevel ? (eL.1 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("Orignaldatei <" path.fFN "> konnte nicht entfernt werden")
				}

			; Backup der extrahierten Textdatei erstellen
				If FileExist(path.fT) && !FileExist(path.iT) {
					FileCopy, % path.fT, % path.iT
					ErrorLevel ? (eL.1 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("Backup für <" path.fFN ".txt> konnte nicht erstellt werden")
				}

			; Textdatei wird nur gelöscht wenn es ein Backup gibt
				If FileExist(path.iT) {
					FileDelete, % path.fT
					ErrorLevel ? (eL.1 := A_LastError) : ""
					If this.debug && ErrorLevel
						SciTEOutput("OCR Textversion von <" path.fFN "> konnte nicht entfernt werden")
				}


		}

		CleanImport(fromDate)                                            	{	; ## Dateien die älter sind werden vollständig gelöscht



		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Hilfsfunktionen
		GetFileExt(fname)                                                   	{	; Dateiendung
			RegExMatch(fname, "i)\.(?<Ext>[a-z]+)$", File)
		return FileExt
		}

		GetPaths(fname:="", oldfile:="", path:="")                 	{	; erstellt alle notwendigen Dateipfade

			/* 	Beschreibung

				- 	erstellt in einem Durchgang alle Dateipfade für die in 3 Ordner verteilten PDF und Textdateien
				- 	vereinfacht dadurch Löschvorgänge, das Umbenennen oder das Importieren, da mit jeder Änderung der eigentlichen
					PDF-Datei die anderen beiden Dateien entsprechend verschoben, entfernt oder umbenannt werden müssen
				-	erstellte Strings werden im Object Path gespeichert. Die Variablennamen sind dabei sehr kurz gewählt.
					verwendete Kurzzeichen:

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
				If this.debug = true
					For key, path in this.Path
						SciTEOutput(key ": " path)

		return this.Path
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

			If InStr(dbgStatus, "toggle")
				this.debug := !this.debug
			else
				this.debug := dbgStatus

			If path && !RegExMatch(path, "^([A-Z]\:\\|\\\\w+)")
				throw A_ThisFunc ": No valid path!`n" path
			else If path && !InStr(FileExist(path), "D")
				throw A_ThisFunc ": Path not exists!`n" path
			else if path
				this.FPath := path, this.initPaths := true
			else
				this.FPath := Addendum.BefundOrdner, this.initPaths := true

		}

}

class PDFpool                                                                 	{               	; verwaltet das ScanPool Objekt

	; Funktion: verwaltet ein Objekt mit virtuellen Dateien. Als Abbild der physischen Dateien auf der Festplatte hält das Objekt zusätzliche Informationen zu den Dateien bereit
	; Abhängigkeiten - Addendum_PdfHelper.ahk
	; ScanPool - muss ein Super-Globales Objekt sein

		inPool(fname)                                                         	{	; sucht nach Dateinamen im Pool

			For docID, PDF in ScanPool
				If (PDF.name = fname)
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

			; ReIndex = true - ScanPool wird komplett neu erstellt oder falls leer wird das gespeicherte Objekt geladen
				If ReIndex
					this.Empty()
				else {
					max := this.Load()
				}

			; alle pdf Dokumente des Befundordners dem ScanPool Objekt hinzufügen
				files  	:= this.GetFiles(Addendum.BefundOrdner, "*.pdf")
				iLen  	:= StrLen(files.MaxIndex())-1
				fcount	:= SubStr("00000" files.MaxIndex(), -1*iLen)

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

		return ScanPool.MaxIndex()
		}

		Remove(fname)                                                     	{	; Datei entfernen
			If (docID := this.inPool(fname))
				return ScanPool.RemoveAt(docID)
		return
		}

		RemoveNoFile()                                                     	{	; entfernt Dateien die nicht mehr existieren
			removed := 0
			If (ScanPool.MaxIndex() > 0) {
				Loop, % (max := ScanPool.MaxIndex()-1)
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

		GetAllDocNames()                                                 	{	; alle Dateinamen zurückgeben
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

		GetNamedDocuments(refreshDocs=false)               	{	; erstellt einen Array mit den Dateinamen umbenannter Dokumente
			If refreshDocs
				admGui_Reload(true)
			named := Array()
			For docID, PDF in ScanPool
				If string.isNamed(PDF.name)
					named.Push(PDF.name)
			If (named.MaxIndex() = 0)
				return
		return named
		}

		GetUnnamed(StartDocID:=0)                                 	{	; Unbenanntes Dokument finden
			For docID, PDF in ScanPool
				If (docID >= StartDocID) && !string.isFullNamed(PDF.name)
					return PDF
		return
		}

		GetUnamedDocuments(refreshDocs=false)               	{	; erstellt einen Array mit den Dateinamen unbezeichneter Dokumente
			If refreshDocs
				admGui_Reload(true)
			unnamed := Array()
			For docID, PDF in ScanPool
				If !string.ContainsName(PDF.name)
					unnamed.Push(PDF.name)
			If (unnamed.MaxIndex() = 0)
				return
		return unnamed
		}

		Empty()                                                                 	{	; ScanPool komplett leeren
			If !IsObject(ScanPool) {
				PraxTT("ScanPool ist kein Objekt!", "2 1")
				return
			}
			Loop % (mxIdx := ScanPool.MaxIndex())
				ScanPool.Pop()
		return mxIdx
		}

		Load(path="")                                                         	{	; ScanPool laden
			If FileExist(scanpoolPath := this.GetValidPath(path) "\PdfDaten.json") {
				For docID, pdf in JSONData.Load(scanpoolPath, "", "UTF-8")
					If FileExist(Addendum.BefundOrdner "\" pdf.name) && !this.inPool(pdf.name)
						ScanPool.Push(pdf)
			}
		return ScanPool.MaxIndex()
		}

		Save(path="")                                                          	{	; ScanPool speichern
		return JSONData.Save(this.GetValidPath(path) "\PdfDaten.json", ScanPool, true,, 1, "UTF-8")
		}

	; + + + + + + + INTERN + + + + + + +
		FileSize(fullfilepath)                                                 	{
			FileGetSize	, FSize, % fullfilepath, K
		return FSize
		}

		FileTime(fullfilepath)                                               	{
			FileGetTime	, TimeStamp 	, % fullfilepath, C
			FormatTime	, FileTime     	, % timeStamp, dd.MM.yyyy
		return {"FileTime" : FileTime, "timeStamp": TimeStamp}
		}

		FileMeta(path, fname )                                           	{
			file := this.FileTime(path "\" fname)
			return {	"name"          	: fname
					, 	"filesize"        	: this.FSize(path "\" fname)
					, 	"timestamp"    	: file.timeStamp
					, 	"filetime"        	: file.FileTime
					, 	"pages"         	: PDFGetPages(path "\" fname, Addendum.PDF.qpdfPath)
					, 	"isSearchable"	: (PDFisSearchable(path "\" fname) ? 1 : 0 )}
		}

		GetFiles(path, filepattern:="*.*", options:="")           	{
			files := Array()
			Loop, Files, % path "\" filepattern, % (options ? options : "")
				files.Push(A_LoopFileName)
		return files
		}

		GetValidPath(path)                                                  	{
		return (this.isPath(path) ? path : Addendum.DBPath "\sonstiges")
		}

		isPath(path)                                                           	{
		return (StrLen(path) > 0 && RegExMatch(path, "i)[A-Z]\:\\") && InStr(FileExist(path),"D") ? true : false)
		}

}
;}

; -------- PDF Dateien umbenennen / teilen                                            	;{

admGui_Rename(filename, prompt="", newfilename="") 	{               	; Dialog für Dokument umbenennen oder teilen

		global

; Initialisierung                                   	;{

	; locals & statics                                                                    	;{
		local fSize, edW	, txtfile, row, rows, rowFilename
		local cp, dp, gs, admPos, title, fileout, vctrl, ENr
		local key, pdf, TLines, found, monIndex, admScreen, RNx, RNWidth
		local names

		static RNminWidth := 350
		static BEZPath, Beschreibung, wT, Patname, Inhalt, Datum, oPV
		static oldadmFile, newadmFile, FileExt, FileOutExt, oldfilename
		static Dateinamen     ; bezeichner
	;}

	; PDF Vorschau in neuem oder vorhandenem Sumatrafenster	;{
		If IsObject(oPV) {
			If (oPV.Previewer = "Sumatra") {
				If (StrLen(WinGetTitle(oPV.ID)) > 0) {
					SumatraDDE(oPV.ID, "OpenFile"	, Addendum.BefundOrdner "\" filename, 0, 0, 0)
					Sleep 100
					SumatraDDE(oPV.ID, "SetView"	, Addendum.BefundOrdner "\" filename, "single page", "-1") 	; -1 = fit page
					Sleep 100
					SumatraDDE(oPV.ID, "SetView"	, Addendum.BefundOrdner "\" filename, "single page", "-1") 	; 2x senden, einmal reicht manchmal nicht
					Gui, RN: Default
				}
				else
					oPV := ""
			}
		} else if !IsObject(oPV) {
			oPV := admGui_PDFPreview(fileName)
			If !IsObject(oPV)
				return
		}
	;}

	; Datei mit Inhaltsbezeichnungen laden                                	;{
		If !IsObject(BefundBezeichner) {
			If !FileExist(BEZPath := Addendum.DBPath "\Dictionary\Befundbezeichner.txt") {
				FilePathCreate(Addendum.DBPath "\Dictionary")
			} else {
				BefundBezeichner 	:= []
				Dateinamen 	:= FileOpen(BEZPath, "r", "UTF-8").Read()
				;~ Dateinamen 	:= RegExReplace(Dateinamen, "[\n\r]{2,}", "")
				;~ Dateinamen 	:= RegExReplace(Dateinamen, "[\n\r]+$", "")
				Sort, Dateinamen, U
				;~ Dateinamen 	:= RegExReplace(Dateinamen, "[\n\r]+", "|")
				For idx, bez in StrSplit(Dateinamen, "`n", "`r")
					If (StrLen(bez) > 0)
						BefundBezeichner.Push(bez)
			}
		}
	;}

	; Namen aus der Datenbank erhalten                                 	;{
		names := admDB.GetNamesArray()
	;}

	; Vorbereitung: Dokument zerlegen
		If Instr(prompt, "Aufteilen")
			title := "Dokument zerlegen", fileout := ""
	; Vorbereitung: Dokument umbenennen
		else {

			oldadmFile := filename
			If (StrLen(newfilename) = 0)
				newfilename := filename

			SplitPath, newfilename,,, FileExt, FileOutExt

			If (FileExt = "pdf") {
				If RegExMatch(FileOutExt, "^\s*(?<N>.*)?,\s*(?<I>.*)?v\.\s(?<D>.*)\s*$", FOE) {
					Patname	:= RegExReplace(FOEN	, "[\s,]+$")
					Inhalt 		:= RegExReplace(FOEI 	, "[\s,]+$")
					Datum 		:= RegExReplace(FOED	, "^[\s,v\.]+")
				} else {
					Patname	:= ""
					Inhalt    	:= RegExMatch(FileOutExt, "^\s*[\d_]+$") ? "" : FileOutExt
					Datum  	:= GetFileTime(Addendum.BefundOrdner "\" filename, "C")
				}
			} else {
					Patname 	:= ""
					Inhalt    	:= RegExMatch(FileOutExt, "^\s*[\d_]+$") ? "" : FileOutExt
					Datum  	:= GetFileTime(Addendum.BefundOrdner "\" filename, "C")
			}

		}

	; Gui wird bei erneutem Aufruf nicht erneut erstellt
		If WinExist("Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI") {
			gosub RNEFill
			return
		}

;}

; Gui                                                  	;{

	;-: Variablen                                                  	;{
	wT          	:= 0
	edW     	:= 100
	fSize     	:= (A_ScreenWidth > 1920 ? 10 : 9)
	admPos	:= GetWindowSpot(hadm)
	RNWidth	:= admPos.W - (2*admPos.BW)
	RNWidth	:= RnWidth < RNminWidth ? RNminWidth : RnWidth		;}

	;-: Gui Start                                                   	;{
	Gui, RN: new, % "AlwaysOnTop -DpiScale +ToolWindow -Caption +Border -SysMenu +OwnDialogs HWNDhadmRN"
	Gui, RN: Margin, 5, 5		;}

	;-: Hinweise                                                    	;{
	Gui, RN: Font, % "s7 cDarkBlue", Calibri
	Gui, RN: Add, Text, % "x2 y2              	BackgroundTrans ", % "Bezeichner:`nNamen:"
	Gui, RN: Add, Text, % "x+2  vRNDBG 	BackgroundTrans ", % "00000`n00000"
	Gui, RN: Font, % "s" fSize - 1 " cBlack", Arial
	RNHinweise =
	(LTrim
	[ Pfeil ⏶& ⏷ von Feld zu Feld, Eingabe für Umbenennen ]
	)
	Gui, RN: Add, Text, % "xm ym Center vRNC1 BackgroundTrans"                         	, % RNHinweise		;}

	;-: Eingabefelder                                              	;{
	Gui, RN: Font, % "s" fSize
	Gui, RN: Add, Text  	, % "xm y+8 Right vRNT1 "                                             	, % "Name"
	Gui, RN: Add, Edit    	, % "x+5 w" edW " r1 vRNE1 HWNDRNhE1 gRNEHandler"	, % Patname
	Gui, RN: Add, Text  	, % "xm y+5 Right vRNT2 " 	                                         	, % "Inhalt"
	Gui, RN: Add, Edit    	, % "x+5 w" edW " r1 vRNE2 HWNDRNhE2 gRNEHandler"	, % Inhalt
	Gui, RN: Add, Text  	, % "xm y+5 Right vRNT3 "                                             	, % "Datum"
	Gui, RN: Add, Edit    	, % "x+5 w" edW " r1 vRNE3 HWNDRNhE3 gRNEHandler"	, % Datum
	;}

	;-: Name, Inhalt, Datum - Breite angleichen    	;{
	; Edit-Felder maximieren
	Loop, 3 {
		cp 	:= GuiControlGet("RN", "Pos", "RNT" A_Index)
		wT	:= cp.W > wT ? cp.W : wT
	}
	Loop, 3 {
		GuiControl, RN:Move, % "RNT" A_Index, % "w" wT
		cp := GuiControlGet("RN", "Pos", "RNT" A_Index)
		GuiControl, RN:Move, % "RNE" A_Index, % "x" cp.X+wT+5 " y" cp.Y-3 " w" RNWidth-wT-15
	}
	;}

	;-: maximale Zeichenzahl                                 	;{
	cp := GuiControlGet("RN", "Pos", "RNE3")
	Gui, RN: Font             	, % "s" fSize - 1 " cWhite", Calibri
	Gui, RN: Add, Progress	, % "x0      	y" cp.Y+cp.H+5 	" w10 h20 c7B89BB vRNPG1 HWNDRNhPG1"          	, 100
	Gui, RN: Add, Text    	, % "x" cp.X "	y" cp.Y+cp.H+7   	"  Left Backgroundtrans vRNHW1"                          	, % "verbrauchte Zeichen (Inhalt+Datum):"
	dp := GuiControlGet("RN", "Pos", "RNHW1")
	GuiControl, RN: Move, % "RNPG1", % "h" dp.H + 5

	Gui, RN: Font            	, % "s" fSize - 1 " cWhite Bold", Consolas
	Gui, RN: Add, Text    	, % "x+5                                    	Left Backgroundtrans vRNLen "                           	, % SubStr("00" (StrLen(Inhalt) + StrLen(Datum)), -1) . " / 70"
	;}

	;-: Dateinamenvorschau                                 	;{
	cp := GuiControlGet("RN", "Pos", "RNPG1")
	dp := GuiControlGet("RN", "Pos", "RNHW1")
	Gui, RN: Font            	, % "s" fSize - 1 " cBlack Normal", Arial
	Gui, RN: Add, Progress	, % "x0 y" dp.Y+dp.H+5 " w10 h" cp.H*2 " cAfC2ff vRNPG2 HWNDRNhPG2"             	, 100
	Gui, RN: Add, Text    	, % "x5 y" dp.Y+dp.H+7 " w" edW " h" dp.H*2 " Center	Backgroundtrans vRNPV "   	, % admGui_FileName(Patname, Inhalt, Datum)
	cp := GuiControlGet("RN", "Pos", "RNPV")
	GuiControl, RN: Move, % "RNCancel", % "h" cp.H + dp.H + 10		;}

	;-: OK, Abbruch                                            	;{
	Gui, RN: Font, % "s" fSize
	Gui, RN: Add, Button, % "xm                         	Center vRNOK  	gRNProceed"         	, % "Umbennen"
	Gui, RN: Add, Button, % "x+20                  	Center vRNCancel	gRNProceed"         	, % "Abbruch"
	cp := GuiControlGet("RN", "Pos", "RNE3"), dp := GuiControlGet("RN", "Pos", "RNCancel")
	GuiControl, RN: Move, % "RNCancel", % "x" cp.X + cp.W - dp.W		;}

	;-: Dateibezeichner                                           	;{
	;~ Gui, RN: Font, % "s" fSize
	;~ Gui, RN: Add, ListView	, % "xm y+10 w" dp.W-10 " r10 -Hdr -E0x200 vRNLV", % "DocNames"  ; gRNListview
	;}

	;-: Show Hide                                                	;{
	monIndex  	:= GetMonitorIndexFromWindow(hadm)
	admScreen	:= ScreenDims(monIndex)
	RNx := ((admPos.X - 2*admPos.BW) + RnWidth) > admScreen.W ? admScreen.W - RnWidth : admPos.X - 2*admPos.BW
	Gui, RN: Show, % "x" RNx " y" admPos.Y + admPos.H " w" RNWidth " Hide", Addendum (Datei umbenennen)		;}

	;-: Titel und Preview anpassen                          	;{
	gs := GetWindowSpot(hadmRN)
	GuiControl, RN: Move, % "RNC1"	, % "w"   	gs.CW - 5
	GuiControl, RN: Move, % "RNPV"	, % "w"  	gs.CW - 5
	GuiControl, RN: Move, % "RNE1"	, % "w"  	gs.CW - wT - 15
	GuiControl, RN: Move, % "RNE2"	, % "w"  	gs.CW - wT - 15
	GuiControl, RN: Move, % "RNE3"	, % "w"  	gs.CW - wT - 15		;}

	;-: Zeichenzähler zentrieren                             	;{
	dp := GuiControlGet("RN", "Pos", "RNHW1")
	cp := GuiControlGet("RN", "Pos", "RNLen")
	dcpLen := dp.W + 5 + cp.W
	dcpMid := Floor(dcpLen/2)
	gsMid	:= Floor(gs.w/2)
	gsLeft	:= (gsMid - dcpMid)
	GuiControl, RN: Move, % "RNHW1"	, % "x" gsLeft
	GuiControl, RN: Move, % "RNLen"  	, % "x" gsLeft + dp.W + 5		;}

	;-: Abbruch Button verschieben                        	;{
	dp := GuiControlGet("RN", "Pos", "RNCancel")
	GuiControl, RN: Move, % "RNCancel", % "x" gs.W - dp.W - 5		;}

	;-: Progress anpassen                                     	;{
	cp := GuiControlGet("RN", "Pos", "RNPV")
	GuiControl, RN: Move, % "RNPG1"	, % "x0 w" gs.W
	GuiControl, RN: Move, % "RNPG2"	, % "x0 w" gs.W
	WinSet, ExStyle, 0x0, % "ahk_id " RNhPG1
	WinSet, ExStyle, 0x0, % "ahk_id " RNhPG2		;}

	;-: Show
	Gui, RN: Show

	If IsObject(names)
		RNAutoComplete1 := IAutoComplete_Create(RNhE1, names              	, ["AUTOSUGGEST", "UseTab"], true)
	If IsObject(BefundBezeichner) {
		;iWinBezeichner(BefundBezeichner)
		RNAutoComplete2 := IAutoComplete_Create(RNhE2, BefundBezeichner	, ["AUTOSUGGEST", "UseTab"], true)
	}

	GuiControl, RN:, RNDBG, % BefundBezeichner.MaxIndex() "`n" names.MaxIndex()

	admGui_RenameHotkeys()

	;}

RNEFill:                                               	;{   Inhalte einfügen

	; Rename-Gui Default machen
		Gui, RN: Default

	; Editfelder befüllen und Vorschau anzeigen
		GuiControl, RN:, RNPV	, % admGui_FileName(Patname, Inhalt, Datum)
		GuiControl, RN:, RNE1	, % PatName
		GuiControl, RN:, RNE2	, % Inhalt
		GuiControl, RN:, RNE3	, % Datum

	; Focus je nach Inhalt der Eingabefelder setzen
		If (StrLen(PatName) = 0)
			GuiControl, RN: Focus, RNE1
		else
			GuiControl, RN: Focus, RNE2

return ;}

RNEHandler:                                      	;{   Dateinamenvorschau und Vorschläge

		Gui, RN: Default
		Gui, RN: Submit, NoHide
		RegExMatch(A_GuiControl, "\d", gcNr)

	; zu lange Zeicheneingabe verhindern
		GuiControl, RN:, RNLen, % SubStr("00" (charsNow := StrLen(RNE2) + StrLen(RNE3)), -1) " / 70 "
		If (charsNow > 70) {
			BlockInput, Send
			Send, % "{BackSpace " (charsNow - 70) "}"
			BlockInput, Off
		}
		Gui, RN: Submit, NoHide

	; Dateinamenvorschau auffrischen
		GuiControl, RN:, RNPV, % admGui_FileName(RNE1, RNE2, RNE3)

	; Dateibezeichnungsauswahl
		;~ If (gcNr=2) {

			;~ Gui, RN: ListView, RNLV
			;~ Gui, RN: ListView, -Redraw
			;~ RNWords := StrSplit(RNE2, A_Space)
			;~ RNWmax := RNWords.Count()

			;~ LV_Delete()


			;~ Gui, RN: ListView, +Redraw
		;~ }

return ;}

RNListview:                                         	;{	  Dateibeschreibung auswählen

return
;}

RNProceed:                                       	;{   Datei wird umbenannt, PDF-Viewer-Vorschau wird geschlossen

	; Abbruch ohne Dateinamenänderung
		If (A_GuiControl = "RNCancel") {
			gosub RNGuiClose
			return
		}	else if (A_GuiControl <> "RNOK") && (A_ThisHotkey <> "Enter") {
			gosub RNGuiClose
			return
		}

		Gui, RN: Default
		Gui, RN: Submit, NoHide
		newadmFile := admGui_FileName(RNE1, RNE2, RNE3)

	; Nutzer hat OK oder Enter gedrückt, aber nichts geändert
		If ((newadmfile . FileExt) = filename) {
			PraxTT("Sie haben den Dateinamen nicht geändert!", "2 1")
			gosub RNGuiBeenden
			return
		} else If (StrLen(newadmfile) = 0) { ; Nutzer hat alles gelöscht
			PraxTT("Ihr neuer Dateiname enthält keine Zeichen.", "2 1")
			gosub RNGuiBeenden
			return
		}

	; neue Inhaltsbeschreibung speichern                  	;{
		found := false
		For idx, beschreibung in BefundBezeichner
			If (beschreibung = RNE2) {
				found := true
				break
			}
		If !found {
			BefundBezeichner.Push(RNE2)
			Dateinamen := RNE2 "`n"
			For idx, beschreibung in BefundBezeichner
				Dateinamen .= beschreibung "`n"
			Sort, Dateinamen, U
			GuiControl, RN:, RNDBG, % BefundBezeichner.MaxIndex() . "`n" . names.MaxIndex()
			Dateinamen := RegExReplace(Dateinamen, "[\n\r]{2,}", "`n")
			FileOpen(BEZPath, "w", "UTF-8").Write(Dateinamen)
		}
	;}

	; FolderWatch pausieren/alle Dateien umbenennen/FolderWatch fortsetzen
		newadmfile .= "." FileExt
		admGui_FWPause(true, [newadmfile	, oldadmfile], 6)    	; FolderWatch pausieren (20s bis zur Wiederherstellung)
		res	:= Befunde.Rename(oldadmfile	, newadmfile)        	; Dateien umbenennen (Original, Backup, Text)
		pdf	:= PDFpool.Rename(oldadmfile	, newadmfile)        	; Objektdaten ändern
		row	:= Journal.Replace(oldadmfile	, newadmfile)        	; Journal auffrischen

;}

RNGuiBeenden:                                   	;{   weitere Dateien umbenennen

	; nachschauen ob eine weitere Datei zum Umbenennen zur Verfügung steht
		oldfilename := newadmFile, startrow := rows := LV_GetCount()
		Loop {

			rowL := row
			row ++
			row := row > rows ? 1 : row
			If (row = startrow) || (A_Index > rows)
				break

			LV_Modify(rowL	, "-Select")
			LV_Modify(row	, "Focus")
			LV_Modify(row	, "Select")

			LV_GetText(lvFilename, row, 1)
			If !string.isFullNamed(lvFilename)
				break

		}

	; weiteres Dokument umbenennen?
		If (StrLen(lvFilename) > 0) && !string.isFullNamed(lvFilename) && (lvFilename <> oldfilename) {
			MsgBox, 0x1024, Addendum für Albis on Windows, % "Möchten Sie mit der nächsten`n[" lvFilename "]`nDatei fortfahren ?", 30
			IfMsgBox, Yes
			{
				admGui_Rename(lvFilename)
				return
			}
		}
;}

RNGuiClose:                                      	;{   Fenster definitiv schliessen
RNGuiEscape:

	; PDF Previewer beenden
		If (oPV.Previewer = "Sumatra") {
			oPV := SumatraClose(oPV.PID, oPV.ID)
			oPV := ""
		}

	; Gui beenden
		Gui, RN: Destroy
		admGui_RenameHotkeys("Off")

	; Journal fokussieren
		Journal.Focus()

return ;}

RNGuiUpDown:                                   	;{   mit den Pfeiltasten zwischen den Eingabefeldern wechseln

	; AutoComplete Listbox wird angezeigt
	If ACGShow
		return

	Gui, RN: Submit, NoHide
	GuiControlGet, vctrl, RN:Focus
	RegExMatch(vctrl, "Edit(?<nr>\d)", E)

	If (A_ThisHotkey = "Up")
		If (ENr = 1)
			GuiControl, RN:Focus, RNE3
		else
			Send, {LShift Down}{Tab}{LShift Up}
	else if (A_ThisHotkey = "Down")
		If (ENr = 3)
			GuiControl, RN:Focus, RNE1
		else
			Send, {Tab}

return ;}
}

iWinBezeichner(obj) {                                                          					; speichert und zeigt das Addendum Objekt an

	SciTEOutput("Bezeichner: " obj.MaxIndex())
	For idx, bez in obj
		t .= bez "`n"
	SciTEOutput(t)

}

admGui_RenameHotkeys(status:="On")                          	{                	; schaltet Hotkey's an oder aus

	Hotkey, IfWinActive, % "Addendum (Datei umbenennen)"
	Hotkey, Escape 	, RNGuiClose	    , % status
	Hotkey, Up    	, RNGuiUpDown	, % status
	Hotkey, Down	, RNGuiUpDown	, % status
	Hotkey, Enter		, RNProceed			, % status
	Hotkey, IfWinActive


}

admGui_FileName(PatName, inhalt, Datum)                  	{                	; Hilfsfunktion: admGui_Rename

	retStr := RegExReplace(PatName, "[\s,]+$") ", "
	retStr .= Trim(RegExReplace(inhalt, "[\s,]+$"))
	retStr .= StrLen(Datum) > 0 ? " v. " Trim(RegExReplace(Datum, "^[\s,v\.]+")) : ""

return retStr
}

;}

; -------- PDF/Bild Viewer                                                                    	;{
admGui_PDFPreview(fileName)                                      	{               	; PDF Vorschau

		global 	hadm, PIDSumatra

		static 	hadmPV, PVPic1, PVbw, PVfw, imagick
		static 	SmtraExist := false, SmtraInit := true
		static 	SumatraCMD
		static 	SmtraClass := "ahk_class SUMATRA_PDF_FRAME"

	; Fensterposition ermitteln
		admPos	:= GetWindowSpot(hadm)
		monIndex	:= GetMonitorIndexFromWindow(hadm)
		Mon     	:= ScreenDims(monIndex)

	; Dateipfad anpassen
		If RegExMatch(filename, "[A-Z]\:\\")
			pdfPath	:= fileName
		else
			pdfPath	:= Addendum.BefundOrdner "\" filename

		pages := PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)

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
		If       	(StrLen(SumatraCMD) = 0) {

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

				If !FileExist(prevPic) {
					PraxTT("PDF Vorschau konnte nicht erstellt werden.", "2 0")
					return
				}

				Gui, admPV: New 	, % "+ToolWindow -DPIScale +HWndhadmPV +Border"
				Gui, admPV: Margin	, 5, 5
				Gui, admPV: Add   	, Button	, % "xm 	ym 	vPVbw 	0xE BackGroundTrans gPVHandler"	, % " << "
				Gui, admPV: Add   	, Button	, % "x+10 ym 	vPVfw 	0xE BackGroundTrans gPVHandler" 	, % " >> "
				Gui, admPV: Add   	, Picture, % "xm  	y+5 	vPVPic1	0xE BackGroundTrans"                   	, % prevPic
				Gui, admPV: Show	, % "AutoSize NA Hide", % "Pdf Vorschau - Seite 1/" pages " [" filename "]"
				PV := GetWindowSpot(hadmPV)
				Gui, admPV: Show, % "x" (admPos.X - PV.W - 5) " y" Round(Mon.B/2) - Round(PV.H/2) " NA"
				activeID := hadmPV, Previewer := "admPV"

		}
	; automatische Größenanpassung des Sumatra PDF Reader
		else if 	(StrLen(SumatraCMD) > 0) {

			; DetectHidden... merken
				lDetectHiddenWin := A_DetectHiddenWindows
				DetectHiddenWindows, On

				Run, % q SumatraCMD q " -view " q "single page" q " -zoom " q "fit page" q " " q pdfPath q,, UseErrorLevel, PIDSumatra   ; -new-window
				WinWait        	, % SmtraClass,, 12
				WinActivate		, % SmtraClass
				WinWaitActive	, % SmtraClass,, 4

				hSumatra  	:= WinExist(SmtraClass)
				;classnn        	:= GetClassName(hSumatra)
				WinGet      	,	SumatraPID, PID, % "ahk_id " hSumatra
				ControlGet	,	hSumatraCnvs, HWND,, SUMATRA_PDF_CANVAS1, % "ahk_id " hSumatra
				WinGetPos	,,	stY,, stH, ahk_class Shell_TrayWnd

				AspectRatio	:= 1.5
				smtraWin   	:= GetWindowSpot(hSumatra)
				smtraCnvs 	:= GetWindowSpot(hSumatraCnvs)
				smtraTBarH	:= smtraWin.H - smtraCnvs.H
				smtraH     	:= Mon.H - stH
				smtraW     	:= Floor(smtraH/1.2)                 			; A4 Seite 29,7:21,0 (cm) ~ 1.41 --- geht nur 1.2?
				smtraX        	:= admPos.X - smtraW

				DllCall("SetWindowPos", "Ptr", hSumatra	, "Ptr", 0, "Int", smtraX, "Int", 5, "Int", smtraW, "Int", smtraH, "UInt", 0x40) ;SHOW_WINDOW

				DetectHiddenWindows, % lDetectHiddenWin
				activeID := hSumatra, activePID := SumatraPID, Previewer := "Sumatra"
			}


return {"Previewer": previewer, "ID": activeID, "PID": activePID, "Path": SumatraCMD}


PVHandler:

return
}

admGui_View(filename, MenuName="")                           	{               	; Befund-/Bildanzeigeprogramm aufrufen

	; letzte Änderung 07.03.2021

	filepath := Addendum.BefundOrdner "\" filename
	If (StrLen(filename) = 0) || !FileExist(filepath)
		return 0

	If RegExMatch(filename, "\.jpg|png|tiff|bmp|wav|mov|avi$") {
		PraxTT("Anzeigeprogramm wird geöffnet", "3 0")
		Run % q filepath q
	}
	else {

		If !RegExMatch(Addendum.PDF.Reader, "i)[A-Z]\:\\")
			If !RegExMatch(Addendum.PDF.Reader, "[\\\/,;%\(\)]")
				pdfReaderPath := GetAppImagePath(RegExReplace(Addendum.PDF.Reader, "\.exe$") ".exe")

		If !RegExMatch(Addendum.PDF.ReaderAlternative, "i)[A-Z]\:\\")
			If !RegExMatch(Addendum.PDF.ReaderAlternative, "[\\\/,;%\(\)]")
				pdfReaderAlternativePath := GetAppImagePath(RegExReplace(Addendum.PDF.ReaderAlternative, "\.exe$") ".exe")

		If !FileExist(pdfReaderAlternativePath)
			If !FileExist(pdfReaderPath) {
				PraxTT("Es konnte kein PDF-Anzeigeprogramm gefunden werden", "2 0")
				return
			}

		If !MenuName
			pdfReader := pdfReaderAlternativePath
		else if (MenuName = "JView1")
			pdfReader := pdfReaderAlternativePath
		else if (MenuName = "JView2")
			pdfReader := pdfReaderPath

		PraxTT("PDF-Datei wird angezeigt", "3 0")
		If !PDFisCorrupt(filepath)
			Run % q pdfReader q " " q filepath q
		else
			PraxTT("PDF Datei:`n>" fileName "<`nist defekt", "2 0")
	}

	Sleep 3000
	PraxTT("", "off")

return 1
}

admGui_GetAllPdfWin()                                                 	{               	; Namen aller angezeigten PdfReader ermitteln

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDF.ReaderWinClass
	Loop % tmpListe
		PreReaderListe .= tmpListe%A_Index% "`n"

return PreReaderListe
}

admGui_ShowPdfWin(PreReaderListe)                             	{               	; PDF-Readerfenster nach vorne holen

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

admGui_ThumbnailView(lparam, wParam, xParam)        	{                	; PDF-Dateien sollen hiermit einblendbar werden ohne einen PDF-Reader aufzurufen

	CoordMode, Mouse, Screen
	global admHJournal, admHReports, hadm, adm

	MouseGetPos, mx, my, hWin, hControl, 2
	oAcc := AccObj_FromPoint(mx, my)
	role := "hwin: " xParam ", hControl: " wParam "`nChildCount: " oAcc.accChildCount "`nName: " oAcc.accName(1) "`nValue: " oAcc.accValue(1)
	;ToolTip, % role, 800, 1, 15

}

AccObj_FromPoint(X := "", Y := "") {

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
;}

; -------- Importieren                                                                          	;{
admGui_ImportGui(Show=true, msg="", gctrl="")              	{                 	; Fenster für laufende Prozesse

	global adm, adm2, hadm, hadm2, admImporter, admImporterPrg, admJournal, last_msg, last_ctrl
	global admHTab

	; das richtige Tab-Control anzeigen
		If (StrLen(gCtrl) = 0) || (gCtrl = "admReports")
			res := admGui_ShowTab("Patient", admHTab), gCtrl := "admReports"
		else if (gCtrl = "admJournal")
			res :=journal.Show()

	; Gui zeigen oder beenden
		If Show {

			; zeigt den Hinweis auch an wenn das Infofenster neu gezeichnet wurde
				If Addendum.ImportFromPat && (StrLen(msg) > 0) {
					last_ctrl	:= gctrl
					GuiControl, adm2:, admImporter, % msg
					If (last_msg <> msg) {
						last_msg := msg
					}
					return
				}
				else if !Addendum.ImportFromPat {
					last_msg := msg, last_ctrl 	:= gctrl
					Addendum.ImportFromPat := true
					If (StrLen(msg) = 0)
						msg := "Importvorgang läuft...`nbitte nicht unterbrechen!"
				}

				;MsgOpt := RegExMatch(msg, "i)(Dateinamen|Datei)") ? "" : ""
				GuiControlGet, rp, adm: Pos, % gCtrl

				Gui, adm2: New	, +HWNDhadm2 +ToolWindow -Caption +Parent%hadm%
				Gui, adm2: Color	, c3170A0
				Gui, adm2: Font	, s10 q5 cFFFFFF Bold, % Addendum.Default.Font
				Gui, adm2: Add	, Text    	, % "x20 y1 " MsgOpt " vadmImporter"        	, % msg
				Gui, adm2: Show	, % "x" rpX " y" rpY " w" rpW " h" rpH " Hide NoActivate" 	, admImportLayer

				; Text zentrieren
				GuiControlGet	, cp, adm2: Pos	, admImporter
				GuiControl    	, adm2: Move	, admImporter, % "x" Floor(rpW/2 - cpW/2)  " y" Floor(rpH/2 - cpH/2)
				Gui, adm2: Show	, % "NoActivate"

				WinSet, Style         	, 0x50020000	, % "ahk_id " hadm2
				WinSet, ExStyle	    	, 0x0802009C	, % "ahk_id " hadm2
				WinSet, Transparent	, 220             	, % "ahk_id " hadm2

		} else {

				Gui, adm2: Destroy
				last_msg := last_ctrl := ""
				Addendum.ImportFromPat := false

		}

}

admGui_ImportJournalAll(files="")                                    	{               	; Journal: 	Import aller vollständig benannten Dateien

	; Variablen
		global admHTab
		static breakImportAll	:= false
		static ImptDbg        	:= false
		static waituser        	:= 6
		static NoImports    	:= 0

	; Import Button inaktiv schalten
		Journal.ShowImportStatus(Addendum.ImportFromJrnl := true)
		Journal.Default()

	; Dokumente ohne Personennamen werden ignoriert
	; Dokumente ohne genaue Bezeichnung (Name und Dokumentdatum nicht)!
		JRowsStart := JRows := Journal.GetCount(), rowNr := 0
		If ImptDbg
			SciTEOutput("JRowsStart: " JRowsStart)

	; Abbruchbedingungen nach while ...
		while (JRows > 0) && (NoImports < JRows) && (rowNr <= JRowsStart) && !breakImportAll {

			; Zeilentext auslesen
				Journal.SelectRow(rowNr++)
				rowFile :=	Journal.GetText(rowNr)

			 ; enthält keinen Personennamen dann weiter
				If !string.isNamed(rowFile) {
					If ImptDbg
						SciTEOutput(" [" rowNr "] not named - " rowFile)
					JRows := Journal.GetCount()
					NoImports ++
					continue
				}

			; Abbruch durch Nutzer ermöglichen
				If (Addendum.iWin.Imports > 0) {

					If ImptDbg
						SciTEOutput(" [" rowNr "] weiteres ? - " rowFile )
					MsgBox, 0x1024, % "Addendum", % "Weiteres Dokument importieren?`n[" rowFile "]", % (rowNr <= 5 ? waituser : waituser/3 )
					IfMsgBox, No
					{
						Addendum.ImportFromJrnl:= false
						breakImportAll              	:= true
						If ImptDbg
							SciTEOutput(" [" rowNr "] Abbruch -" rowFile)
						break
					}
				}

			; Importieren
				If !breakImportAll {

					filePath := Addendum.BefundOrdner "\" rowFile
					If RegExMatch(filePath, "\.pdf$") && FileExist(filePath) && !FileIsLocked(filePath) && PDFisSearchable(filepath) {

						; -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
						; Karteikarte wurde geöffnet
							If FuzzyKarteikarte(rowFile) {

								; Importieren - erfolgreich
									If admGui_ImportFromJournal(rowFile) {
										Addendum.iWin.ImportsLast  := Addendum.iWin.Imports
										Addendum.iWin.Imports ++
										If ImptDbg
											SciTEOutput(" [" rowNr "] importiert - " rowFile)
								; Importieren - fehlgeschlagen
									} else {
										NoImports ++
										If ImptDbg
											SciTEOutput(" [" rowNr "] Importfehler - " rowFile)
									}

						; Öffnen der Karteikarte ist fehlgeschlagen
							} else {

								PraxTT("Es konnte kein passender Patient gefunden werden.", "3 0")
								If ImptDbg
									SciTEOutput(" [" rowNr "] filenotfound - " rowFile)
								NoImports ++
								JRows := Journal.GetCount()
								If (JRows = NoImports)
									break

							}

					}

				}
			; Nutzer hat den Vorgang manuell abgebrochen (breakImportAll = true), die Schleife wird hier verlassen
				else {
					break
				}

				JRows := Journal.GetCount()

		}

	; Zurücksetzen der Zähler
		Addendum.ImportFromPat   	:= false
		Addendum.iWin.ImportsLast  := -1
		Addendum.iWin.Imports    	:= 0
		breakImportAll                     	:= false
		NoImports := rowNr            	:= 0

	; Importieren Steuerelemente aktivieren, Infotexte ändern
		Journal.SelectRow(0)
		Journal.ShowImportStatus(Addendum.ImportFromJrnl := false)
		Journal.InfoText()

return
}

admGui_ImportFromJournal(filename)                              	{               	; Journal: 	Einzelimport-Funktion ### LV Ausgabe anpassen

	; letzte Änderung: 12.03.2021

		global 	admHPDFfilenames, hadm, admHTab, admJournal
		static 	WZEintrag	:= {"LASV":"Anfragen", "Anfrage":"Anfragen", "Antrag":"Anfragen", "Lageso":"Anfragen", "Lebensversicherung":"Anfragen"}

	; Nutzer befragen
		currPat 		:= AlbisCurrentPatient()
		currPatID 	:= AlbisAktuellePatID()
		Pat	    	:= string.Get.Names(filename)
		If Addendum.iWin.ConfirmImport || (!InStr(currPat, Pat.Nn) && !InStr(currPat, Pat.Vn)) {
			MsgBox, 262148	, % "Befund importieren?"
										, % "Wollen Sie die Datei:`n"
										. 	 (StrLen(filename) > 30 ? SubStr(filename, 1, 30) "..." : filename) "  dem Pat.`n"
										. 	currPat " zuordnen ?"
			IfMsgBox, No
				return 0
		}

	; Dokumentdatum aus dem Dateinamen entnehmen (falls enthalten, sonst leer)
		DocDate := string.get.DocDate(filename)

	; Befund importieren
		If      	RegExMatch(filename, "\.pdf")                                             	{

			Befunde.Backup(filename)                          	; Backup der Datei anlegen falls bisher nicht erfolgt
			If !AlbisImportierePdf(filename, DocDate)     	; Importfunktion aufrufen
				return 0

			pdf 	:= pdfpool.Remove(filename)            	; aus ScanPool entfernen
			res	:= Befunde.MoveToImports(filename) 	; Backup- und Textdatei in ihre Unterverzeichnisse verschieben

		}
		else If	RegExMatch(filename, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{
			If !AlbisImportiereBild(filename, string.Replace.Names(filename), DocDate)
				return 0
			res	:= Befunde.MoveToImports(filename)
		}
		else
			return 0

	; Re-Indizieren, Gui auffrischen, Importprotokoll
		FileAppend, % datestamp() "| " filename "   >>>   [" currPatID "] " currPat "`n", % Addendum.DBPath "\sonstiges\PdfImportLog.txt"

	; Anzeigen und interne Objekte bearbeiten
		result	:= 	Journal.Remove(filename)             	; die Listview wird von diesem Eintrag befreit
		PatDocs:= 	admGui_Reports()
		result	:= 	Journal.Show()
		                	Journal.InfoText()

return 1
}

admGui_ImportFromPatient()                                            	{               	; Patient:	Befundimport alle Befunde

	; letzte Änderung: 17.02.2021

		global 	PatDocs

		Docs 	:= Object()

	; Importschleife
		For key, file in PatDocs		{

			Addendum.FuncCallback := ""
			Befunde.Backup(file.name)                                          	; Backup der Datei anlegen falls bisher nicht erfolgt

			DocDate 	:= string.get.DocDate(file.name)
			KKText    	:= string.Karteikartentext(file.name)

		; Bild importieren
			If RegExMatch(file.name, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{

				If !AlbisImportiereBild(file.name, KKText, DocDate) {
					Docs.Push(file)
					continue
				}
				res := Befunde.MoveToImports(file.name)     	; Backup- und Textdatei in ihre Unterverzeichnisse verschieben

			}
		; Pdf importieren
			else {

				; sind ohne .pdf ?
					Addendum.FuncCallback := "GetPDFViewerArg"
					If !AlbisImportierePdf(file.name, KKText, DocDate) {
						Docs.Push(file)
						continue
					}

				; Logfile schreiben. Backup-PDF-Datei und PDF-Textdatei verschieben
					res	:= Befunde.MoveToImports(file.name)
					pdf	:= pdfpool.Remove(file.name)

			}

			journal.Remove(file.name)                                      	; aus Listview entfernen
			FileAppend, % datestamp() "| " file.name "`n", % Addendum.DBPath "\sonstiges\PdfImportLog.txt"

		}

	; PatDocs  leeren und mit nicht verarbeiteten Dateien füllen
		Loop % PatDocs.MaxIndex()
			PatDocs.RemoveAt(1)
		For index, file in Docs
			PatDocs.Push(file)

return Imports
}
;}

; -------- Karteikarte                                                                           	;{
FuzzyKarteikarte(NameStr)                                             	{               	; fuzzy name matching function, öffnet eine Karteikarte

	; prüft ob Namen übergeben wurden
		If !IsObject(Pat := string.Get.Names(NameStr)) {
			PraxTT(	"Die Dateibezeichnung enthält keinen Namen eines Patienten.`nDer Karteikartenaufruf wird abgebrochen!", "3 2")
			return 0
		}

	; passende Patienten suchen
		Patients := admDB.StringSimilarityEx(Pat.Nn, Pat.Vn)
		m := Patients.diff
		Patients.Delete("diff")

	; Karteikartenfunktion aufrufen
		If (Patients.Count() = 1)
			For PatID, Patient in Patients
				return admGui_Karteikarte(PatID)
		else
			return 0

return 1
}

admGui_Karteikarte(PatID)                                             	{               	; fragt ob eine Karteikarte geöffnet werden soll

	If !Addendum.PDF.PatAkteSofortOeffnen {
		MsgBox, 4	, % "Patientenakte öffnen ?"
						, % "Möchte Sie die Akte des Patienten:`n"
						. 	  "(" PatID ") " oPat[PatID].Nn ", " oPat[PatID].Vn " geb.am " oPat[PatID].Gd "`n"
						.	  "öffnen?"
		IfMsgBox, No
			return 0
	}

return AlbisAkteOeffnen(oPat[PatID].Nn ", " oPat[PatID].Vn, PatID)
}
;}

; -------- Gui Einstellungen speichern                                                  	;{
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

	; letzte Änderung: 05.06.2021

		global admHJournal, adm, hadm, admTPTag, admTabs, admLBCk, admLBDPD

	; Gui auslesen
		Gui, adm: Submit, Hide
		GuiControlGet, ProtokollTag, adm:, admTPTag
		JournalSort := LV_GetScrollViewPos(admHJournal)

	; bei Bedarf in der ini speichern
		If (Addendum.iWin.firstTab     	<> admTabs)
			IniWrite, % admTabs      	, % Addendum.Ini, % compname, % "Infofenster_aktuelles_Tab"
		If (Addendum.iWin.JournalSort 	<> JournalSort)
			IniWrite, % JournalSort   	, % Addendum.Ini, % compname, % "Infofenster_JournalSortierung"
		If (Addendum.iWin.TProtDate 	<> ProtokollTag)
			IniWrite, % ProtokollTag 	, % Addendum.Ini, % compname, % "Infofenster_Tagesprotokoll_Datum"
		If (Addendum.iWin.LBDrucker <> admLBDPD)
				IniWrite, % admLBDPD	, % Addendum.Ini, % compname, % "Infofenster_Laborblatt_Drucker"
		If (Addendum.iWin.LBAnm       	<> admLBCk)
			IniWrite, % (admLBCk ? "Ja" : "Nein")	, % Addendum.Ini, % compname, % "Infofenster_Laborblatt_Anmerkung_drucken"

	; sichern im Addendum Objekt
		Addendum.iWin.firstTab    	:= admTabs
		Addendum.iWin.TProtDate 	:= ProtokollTag
		Addendum.iWin.JournalSort	:= JournalSort
		Addendum.iWin.LBDrucker 	:= admLBDPD
		Addendum.iWin.LBAnm     	:= admLBCk

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
		SciTEOutput("  - AutoNaming: gestartet")
		AutonamingRunning	:= true
		newfilenames         	:= admCM_JRecogAll([pdfFile])
		AutonamingRunning	:= false
		SciTEOutput("  - AutoNaming: beendet")

	; Datei aus dem Stapel entfernen
		OCRStack := Addendum.iWin.FilesStack
		Addendum.iWin.FilesStack := []
		For idx, inprogress in OCRStack {

			For newNr, newfilename in newfilenames {
				FileIsProcessed := false
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

	 */
	; letzte Änderung: 08.03.2021

		global hadm
		static func_AutoOCR	:= func("admGui_OCRAllFiles")
		static OSDTip := false

	; untersucht alle veränderten oder hinzugfügten Dateien
		For Each, Change In Changes 	{

			action	:= change.action
			name	:= change.name

		; abbrechen wenn eine Infofenster Funktion Dateien ändert
			If Addendum.OCR.FWPause || (Addendum.OCR.FWIgnore.Count() > 0) {
				For IFnr, ignoreFile in Addendum.OCR.FWIgnore
					If (name = ignoreFile)
						return
			}

			FileAppend, % A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min " | " action " | " name "`n", % Addendum.DBPath "\WatchFolder-Log.txt", UTF-8

		; PDF Dateien
			If RegExMatch(name, "i)\.pdf$")
				If RegExMatch(action, "1|3|4")	{

					SplitPath, name, filename, filepath
					If IsObject(oPDF := pdfpool.Add(filepath, filename)) {

						; neue Datei zu Listview hinzufügen und Info's ändern
							If hadm {
								admGui_Default("admJournal")
								LV_Add((oPDF.isSearchable ? "ICON3" : "ICON1"), oPDF.name, oPDF.pages, oPDF.filetime, oPDF.timeStamp)
								admGui_InfoText("Journal")
							}

						; Dateizähler erhöhen
							Addendum.OCR.staticFileCount 	++
							Addendum.OCR.filecount           	++

						; Datei hat noch keine Texterkennung* erhalten, Timer mit Verzögerung wird gestartet
						;                                  	*wird nur auf dem Client ausgeführt der in der Addendum.ini hinterlegt ist
							If Addendum.OCR.AutoOCR && (Addendum.OCR.Client = compname) && !oPDF.isSearchable
								SetTimer, % func_AutoOCR, % "-" (Addendum.OCR.AutoOCRDelay*1000)

					}

				}
			; Datei wurde gelöscht
				else If RegExMatch(action, "2")	{

					SplitPath, name, filename, filepath
					oPDF := pdfpool.Remove(filename)
					Controls("", "reset", "")
					If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {  ; Infofenster wird angezeigt?
						admGui_Default("admJournal")
						row := LV_FindRow("admJournal", 1, oPDF.name)
						LV_Delete(row)
						admGui_InfoText("Journal")
					}

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

	; WatchFolder ist nicht eingeschaltet. Ausführung ist nicht notwendig.
		If !Addendum.OCR.WatchFolder
			return

	; Resettimer verlängern
		If Addendum.OCR.FWPause && FWPause {
			SetTimer, FWPauseOff, % (-1*PTime*1000)
			return
		}

	; zu ignorierende Dateiliste anlegen
		If !FWPause {
			Addendum.OCR.FWIgnore:= ""
			Addendum.OCR.FWPause	:= false
			return
		}
		else {
			Addendum.OCR.FWIgnore := Array()
			Addendum.OCR.FWPause	:= true
			If IsObject(files) {
				For fidx, filename in files
					Addendum.OCR.FWIgnore.Push(filename)
			} else if (StrLen(files) > 0) {
				Addendum.OCR.FWIgnore.Push(files)
			}
			SetTimer, FWPauseOff, % -1*PTime*1000
		}

		;SciTEOutput("FWPause at  " A_Min ":" A_Sec ": " (Addendum.OCR.FWPause ? "true":"false"))

return
FWPauseOff:
	Addendum.OCR.FWIgnore	:= Array()
	Addendum.OCR.FWPause 	:= false
	;SciTEOutput("FWPause at  " A_Min ":" A_Sec ": " (Addendum.OCR.FWPause ? "true":"false"))
return
}
;}

; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------
; -------- Texterkennung                                                                       	;{
TesseractOCR(files)                                                         	{               	; erstellt einen zusätzlichen echten Thread

	 /*  Funktionsweise


	#	Erstellung des Threads:

		- 	ein Teil des Threads wird in Addendum.ini in die Variable >Addendum.Threads.OCR< geladen (Addendum_OCR.ahk)
		- 	die Funktion TesseractOCR() bringt die Skriptzeilen die als Autoexec-Bereich eines Autohotkey-Skriptes bezeichnet werden
			und die Parameter für die aufgerufene Funkion tessOCRPdf() in Form von 2 Objekten
		-	die beiden Teile werden zu einem Skript zusammengefügt und per AutohotkeyH Befehl AHKThread(script) ausgeführt


	#	Kommunikation zwischen Thread und Skript:

		-	zwischen dem Thread und Skript werden Nachrichten getauscht
		-	der Thread meldet die Fertigstellung einer Datei und wartet auf Nachricht vom Skript das es	weitermachen soll.
			Dies ermöglicht dem Skript eigene Vorgänge auszuführen oder abzuschliessen bevor die nächste
			Datei fertiggestellt wird.
		- die Kommunikationsfunktionen des Skriptes finden sich in Addendum.ahk und vom Thread in Addendum_OCR.ahk

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

		Settings := {	"SetName"                             	: "threadOCR"
						,	"UseRamDisk"                        	: "T:"
						,	"imagickPath"                           	: Addendum.PDF.imagickPath
						,	"qpdfPath"                             	: Addendum.PDF.qpdfPath
						,	"tessPath"                               	: Addendum.PDF.tessPath
						,	"xpdfPath"            	                 	: Addendum.PDF.xpdfPath
						,	"documentPath"                        	: Addendum.BefundOrdner
						,	"backupPath"                	         	: Addendum.BefundOrdner "\backup"
						,	"txtOCRPath"        	                 	: Addendum.BefundOrdner "\Text"
						,	"OCRLogPath"      	                  	: Addendum.DBPath "\OCRTime_Log.txt"
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
		Gui, OCR: Show, Hide NA, Addendum tessOCR processor
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
			admGui_OCRButton("+OCR abbrechen")
			Addendum.Thread["tessOCR"] := AHKThread(script)
			Addendum.tessOCRRunning := true
		}

}

admGui_OCRAllFiles()                                                      	{               	; Texterkennung aller unbearbeiteten PDF-Dateien

		static noOCR

	; Gui Button ändern
		If Addendum.OCR.AutoOCR && (Addendum.OCR.Client = compname) && Addendum.Thread["tessOCR"].ahkReady() {
			Addendum.OCR.RestartAutoOCR := true
			admGui_OCRButton("+OCR abbrechen")
			return
		}

	; Dateinamen ohne Texterkennung ermitteln. Dateinamen in FileStack ablegen
		noOCR := pdfpool.GetNoOCRFiles()
		If (noOCR.MaxIndex() > 0) {

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

			admGui_OCRButton("+OCR abbrechen")
			Addendum.OCR.RestartAutoOCR := false
			TesseractOCR(noOCR)

		}
		else
			admGui_OCRButton("-OCR ausführen")

return
}
;}

; -------- Hilfsfunktionen                                                                     	;{
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
;}

; -------- Debugfunktionen                                                                     	;{
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

		ControlGet, Content, List,,, % "ahk_id " hLB                        ; Inhalt ermitteln
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

GuiControlActive(CtrlName, GuiName)                          	{                	;-- bestimmt ob ein GuiControl den Eingabefocus hat

	Gui, %GuiName%: Default
	GuiControlGet, fCtrl, FocusV ;% "ahk_id " hWin
	If (fCtrl = CtrlName)
		return true

return false
}

LV_GetSelected(LV_Name)                                             	{              	;-- ermittelt alle ausgewählten Einträge

	admGui_Default(LV_Name)
	cRow := 0, JFiles := Array()

	Loop
		If (cRow := LV_GetNext(cRow)) {
			LV_GetText(fname, cRow, 1)
			JFiles.Push(fname)
		}
		else
			break

	if JFiles.MaxIndex() = 1
		return JFiles[1]
	else if (JFiles.MaxIndex() > 1)
		return JFiles
	else
		return ""
}

RedrawWindow(hwnd=0)                                              	{                 	;-- zeichnet eine Autohotkey Gui komplett neu

	global hadm

	static RDW_INVALIDATE 	:= 0x0001
	static	RDW_ERASE           	:= 0x0004
	static RDW_FRAME          	:= 0x0400
	static RDW_ALLCHILDREN	:= 0x0080

return DllCall("RedrawWindow", "Ptr", (hwnd = 0 ? hadm : hwnd), "Ptr", 0, "Ptr", 0, "UInt", RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN)
}
;}

; -------- Extern                                                                                  	;{
IAutoComplete_Create(HEDT, Strings, Options := "", WantReturn := False, Enable := True) {
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
      If !(IAC2 := ComObjCreate("{00BB2763-6A77-11D0-A535-00C04FD7D062}", "{EAC04BC0-3791-11d2-BB95-0060977B464C}"))
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
         ControlGet, Styles, Style, , , ahk_id %HEDT%
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
      Addr := NumPut(RegisterSyncCallback("IEnumString_QueryInterface") , Addr + 0, "UPtr")
      Addr := NumPut(RegisterSyncCallback("IEnumString_AddRef")         , Addr + 0, "UPtr")
      Addr := NumPut(RegisterSyncCallback("IEnumString_Release")        , Addr + 0, "UPtr")
      Addr := NumPut(RegisterSyncCallback("IEnumString_Next")           , Addr + 0, "UPtr")
      Addr := NumPut(RegisterSyncCallback("IEnumString_Skip")           , Addr + 0, "UPtr")
      Addr := NumPut(RegisterSyncCallback("IEnumString_Reset")          , Addr + 0, "UPtr")
      Addr := NumPut(RegisterSyncCallback("IEnumString_Clone")          , Addr + 0, "UPtr")
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


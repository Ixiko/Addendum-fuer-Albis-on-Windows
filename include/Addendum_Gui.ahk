; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
;                        	by Ixiko started in September 2017 - letzte Änderung 16.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

; Infofenster
AddendumGui(admShow="") {

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global

		static admWidth, admHeight, admTabNames, admLVPatOpt, admLVJournalOpt, admLVTProtokoll, DDLPrinter

		local MObj, MNr, cpx, cpy, cpw, cph
		local APos, SPos, mox, moy, moWin, moCtrl, regExt, Aktiv, Pat, nfo, res
		local ClientName, LanClients, ClientIP, onlineclient, OffLineIndex, InfoText, found, YPlus, Y

		If !Addendum.iWin.Init {

				Addendum.iWin.FileStack   	:= Array()
				Addendum.iWin.OCRStack   	:= Array()     	; sammelt zunächst alle PDF Dateien die noch keine Texterkennung erhalten haben
				Addendum.iWin.RowColors	:= false
				Addendum.iWin.Imports     	:= 0
				Addendum.iWin.ReCheck   	:= 0
				Addendum.Importing                     	:= 0

				func_admGui       	:= Func("AddendumGui")
				factor                	:= A_ScreenDPI / 96

				admTabNames  	:= "Patient|Journal|Protokoll|Extras|Netzwerk|Info"
				admLVPatOpt    	:= " r" (Addendum.iWin.LVScanPool.r)
				admLVPatOpt    	.= " gadm_Reports         	vadmReports     	HWNDadmHReports     	AltSubmit -Hdr	ReadOnly	BackgroundF0F0F0 LV0x10020 0x5001804B "
				admLVJournalOpt	:= " gadm_Journal        	vadmJournal 		HWNDadmHJournal 	 	AltSubmit                      	BackgroundF0F0F0 LV0x10020 0x50210049 -E0x200 "
				admLVTProtokoll 	:= " gadm_TProtokollLV 	vadmTProtokoll 	HWNDadmHTProtokoll  	AltSubmit      	ReadOnly	BackgroundF0F0F0 LV0x10020 0x50210049 -E0x200"

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
				func_admView1	:= Func("admGui_CM").Bind("JView1")
				func_admView2	:= Func("admGui_CM").Bind("JView2")
				func_admRefresh	:= Func("admGui_CM").Bind("JRefresh")
				func_admRecog1	:= Func("admGui_CM").Bind("JRecog1")
				func_admRecog2	:= Func("admGui_CM").Bind("JRecog2")
				func_admOCR  	:= Func("admGui_CM").Bind("JOCR")
				func_admOCRAll	:= Func("admGui_CM").Bind("JOCRAll")

				Menu, admJCM, Add, Karteikarte öffnen                      	, % func_admOpen
				Menu, admJCM, Add, Anzeigen                                 	, % func_admView1
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
	; Vorbereitungen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; Albis wurde beendet -return-
		If !WinExist("ahk_class OptoAppClass") {
			admGui_Destroy()
			return
		}

	; mehr als 9 Versuche das MDIFrame zu finden, führen zum Abbruch des Selbstaufrufes der Funktion
		If (Addendum.iWin.ReCheck > 9) {
			Addendum.iWin.ReCheck := 0
			TrayTip, AddendumGui, % "Die Integration des Infofenster ist fehlgeschlagen.", 2
			return
		}

	; Position innerhalb des Albisfenster für die Integration finden
		Aktiv         	:= Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		res                := Controls("", "Reset", "")
		hMDIFrame	:= Controls("AfxMDIFrame1401"	, "ID", AlbisGetActiveMDIChild())
		hStamm		:= Controls("#327701"             	, "ID", hMDIFrame)
		If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten|Rechnungsliste") || (GetDec(hMDIFrame) = 0) || (GetDec(hStamm) = 0) {
			SetTimer, % func_admGui, -1000
			Addendum.iWin.lastPatID := 0
			Addendum.iWin.ReCheck ++
			return
		}

	; legt die Höhe der Gui anhand der Größe des MDIFrame-Fenster fest
		If (Addendum.iWin.lastPatID <> AlbisAktuellePatID()) {

			APos  	:= GetWindowSpot(AlbisWinID())
			SPos 	:= GetWindowSpot(hStamm)

			Addendum.iWin.StammY	:= Round(SPos.Y / factor)
			Addendum.iWin.X	         	:= APos.CW - Addendum.iWin.W ; ClientWidth passt besser!
			Addendum.iWin.H          	:= admHeight := SPos.CH
			admWidth                     	:= Addendum.iWin.W

		; der Albisinfobereich (Stammdaten, Dauerdiagnosen, Dauermedikamente) wird für das Einfügen der Gui verkleinert
			If (SPos.W > APos.CW - admWidth) ; verhindert ein wiederholtes Verkleinern
				SetWindowPos(hStamm, 2, 2, APos.CW - admWidth - 4, SPos.CH)
		}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeichnen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	If (Addendum.iWin.lastPatID <> AlbisAktuellePatID()) {
		Addendum.iWin.lastPatID := AlbisAktuellePatID()

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
		Gui, adm: Add  	, Tab     	, % "x0 y0  	w" admWidth " h" admHeight " HWNDadmHTab vadmTabs gadm_Tabs"       	, % admTabNames
	;}

	;-: Tab1 :- Patient                          	;{
		Gui, adm: Tab  	, 1

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y27 w" admWidth-15 " BackgroundTrans vadmPatientTitle Section"                     	, % "(es wird gesucht ....)"

		Gui, adm: Font  	, s6 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-55 " y23 w50 vadmButton1 gadm_BefundImport"                                    	, % "Importieren"

	  ;-: Patientenbefunde
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+2 	w" admWidth-6	" " admLVPatOpt                                                                   	, % "S|Befundname"
		LV_SetImageList(admImageListID)
		admCol1W :=50, admCol2W := admWidth - admCol1W - 10
		LV_ModifyCol(1, admCol1W " Integer Right NoSort")
		LV_ModifyCol(2, admCol2W " Text Left NoSort")

	  ;-: Abrechnungshelfer
		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		cpw := Floor((admWidth-6)/1.8)
		Gui, adm: Add  	, Text    	, % "x2 y+0 	w" cpw " Center vadmGP1"   	                                                                    	, % "Abrechnungshelfer und andere Hinweise"

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		GuiControlGet, cpos, adm: Pos, AdmGP1

		cph := admHeight - cposY - cposH - 4
		GCOption	:= "x2 y" cposY+12 " w" cpw " h" cph " t6 t12 t22 vadmNotes"

		If Addendum.iWin.AbrHelfer
			abrInfo    	:= admGui_Abrechnungshelfer(AlbisAktuellePatID(), AlbisPatientGeburtsdatum())
		else
			abrInfo 	:= ""
		Gui, adm: Add, Edit, % GCOption                                                                                                                              	, %  abrInfo

	  ;-: zusätzliche Karteikartenfunktionen
		GuiControlGet, cpos, adm: Pos, admNotes
		cpx := cposX + cposW + 3, cpy := cposY - 1, cpw	:= admWidth - cpx - 120 - 6
		Gui, adm: Add, Button       	, % "x" cpx " y" cpy " w100 vadmLBD  gadm_LBDruck"                                                        	, % "Laborblatt Druck"
		Gui, adm: Add, Edit          	, % "x+2 y" cpy+1 " w" cpw " vadmLBDCB"                                                                      ;	, % "1|alle"
		Gui, adm: Add, UpDown   	, % "x+0 y" cpy+1 " vadmLBDUD gadm_LBDruck"                                                              	,  1

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		cpw	:= admWidth - cpx - 5
		Gui, adm: Add, DDL          	, % "x" cpx+1 " y+2 w" cpw " r8 vadmLBDPD gadm_LBDruck"                                           	, % DDLPrinter
		GuiControl, adm: ChooseString, admLBDPD, % Addendum.iWin.LBDrucker
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
		Gui, adm: Add  	, Button    	, % "x20 y23 h16                          	vadmButton4 gadm_Journal"                                        	, % "OCR ausführen  "

		GuiControlGet  	, cpos, adm: Pos, admButton3
		GuiControlGet  	, dpos, adm: Pos, admButton4
		GuiControl        	, adm: Move, admButton4, % "x" cposX - dposW

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5 		w" admWidth - 6 " h" admHeight - 47 " " admLVJournalOpt                        	, % "Befund|S|Eingang|TimeStamp"

		admCol4W := 0
		admCol3W := 58
		admCol2W := 25
		admCol1W := admWidth - admCol2W - admCol3W - 25

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
		ProtText := "[" compname "] [" TProtokoll.MaxIndex() " Patienten]"
		Gui, adm: Add  	, Text      	, % "x10 y25  w120 BackgroundTrans 	vadmTProtokollTitel"                                          	, % ProtText

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial

		Gui, adm: Add  	, Text      	, % "x+10 "                                                                                                                   	, % "Pat:"
		Gui, adm: Add  	, Edit     	, % "x+0 y22 w100 r1  vadmTPSuche"                                                                               	, % ""

		Gui, adm: Add  	, Text      	, % "x" admWidth - 100 - 5 " y25 w100 BackgroundTrans 	vadmTPTag"                          	, % Addendum.iWin.TProtDate
		Gui, adm: Add  	, Button    	, % "x+10 y23 h16                             	vadmTPZurueck	 	gadm_TP"                              	, % "<"
		Gui, adm: Add  	, Button    	, % "x+2 	y23 h16						    	vadmTPVor   		gadm_TP"                              	, % ">"
		GuiControlGet	, cpos, adm: Pos, admTPVor
		GuiControlGet	, dpos, adm: Pos, admTPZurueck
		GuiControl    	, adm: Move, admTPVor    	, % "x" admWidth - 100 - 5 - cposW - 5
		GuiControl    	, adm: Move, admTPZurueck	, % "x" admWidth - 100 - 5 - cposW - dposW - 10

		;Gui, adm: Add  	, Text      	, % "x+5 y25 w30 BackgroundTrans	 Right vAdmTPTag          	"                                     	, % "Heute"
		;Gui, adm: Add  	, Text      	, % "x+5 y25 w70 BackgroundTrans	vAdmTPDatum                 	"                                      	, % A_DD "." A_MM "." A_YYYY

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5 w" admWidth-6 " h" admHeight-47 " " admLVTProtokoll                                    	, % "RF|Nachname, Vorname|Geburtstag|PatID"

		admcol1W := 30, admcol2W := 160, admcol3W := 70, amdcol4W := 70
		Gui, adm: ListView, AdmTProtokoll
		LV_ModifyCol(1, admcol1W )
		LV_ModifyCol(2, admcol2W )
		LV_ModifyCol(3, admcol3W )
		LV_ModifyCol(4, admcol4W )
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
		mStep 	:= 0 	, mxWidth	:= 115, stepX := mxWidth + 5
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

	;-: Tab5 :-Netzwerk                       	;{
		Gui, adm: Tab  	, 5
		Gui, adm: Font  	, s8 q5 Normal underline cBlack, Arial

		;~ nprop	:= {	"title"         	: "admNet"
						;~ , 	"x"             	: 1
						;~ , 	"y"             	: 20
						;~ , 	"w"             	: Addendum.iWin.W -	2
						;~ , 	"h"            	: Addendum.iWin.H -	22
						;~ , 	"hparent"    	: hadm
						;~ ,	"parentgui"	: "adm"}
		;~ htmlPath := CreateHTML(nprop)
		;~ MoNet := new NeutronEmbedded(htmlpath,,, nprop)

		Gui, adm: Add   	, Button	, % "x10 y30 vAdmLanCheck gadm_Lan", % "Netzwerkgeräte aktualisieren"

		dposX := 40, dposY := 60, cMaxW := 0
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		For clientName, client in Addendum.LAN.Clients {

			If (compname = clientName)
				continue

			Gui, adm: Add   	, Button	, % "x" dposX " y" dposY " vadmClient_"	clientName " Center gadm_Lan", % clientName
			Gui, adm: Add   	, Button	, % "x150		  y" dposY " vadmRDP_"  	clientName " Center gadm_Lan", % "RDP Sitzung starten"
			GuiControlGet	, cpos, adm: Pos, % "AdmClient_" clientName
			dposY := cposY + cposH
			cMaxW := cposW > cMaxW ? cposW : cMaxW

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
		cpy -= 10

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

	;}

		Gui, adm: Tab
	}

	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeigen
	;---------------------------------------------------------------------------------------------------------------------------------------------;{
		; Fensterstyles anpassen
			If !hMDIFrame {
				res                := Controls("", "Reset", "")
				hMDIFrame	:= Controls("AfxMDIFrame1401"	, "ID", AlbisGetActiveMDIChild())
			}
			SetParentByID(hMDIFrame, hadm)

			WinSet, Style              	, 0x50000000	, % "ahk_id " hadm
			WinSet, ExStyle          	, 0x0802009C	, % "ahk_id " hadm

		; Fenster anzeigen
			Gui, adm: Show, % "x" Addendum.iWin.X " y" Addendum.iWin.Y " w" Addendum.iWin.W  " h" Addendum.iWin.H " HIDE NA " admShow, AddendumGui
			WinGet, Style, Style, % "ahk_id " hadm
			If (Style & 0x80000000) {
				admGui_Destroy()
				AddendumGui()
			}


	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tabs - Listviews mit Daten befüllen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Addendum.iWin.PatID := AlbisAktuellePatID()
		If (Addendum.iWin.PatID <> Addendum.iWin.lastPatID)
			Addendum.iWin.ReIndex := true, Addendum.iWin.lastPatID := Addendum.iWin.PatID
		else
			Addendum.iWin.ReIndex := false

		; Inhalte auffrischen und anzeigen
			res := admGui_ShowTab(Addendum.iWin.firstTab, admHTab)
			admGui_CountDown()
			admGui_Journal( (!Addendum.iWin.Init ? true : false) )
			Addendum.iWin.Init ++
			admGui_TProtokoll(Addendum.TProtDate, 0, compname)
			PatDocs := admGui_Reports()
			Addendum.iWin.ReCheck := 0
	;}

	; PDF Thumbnailpreview
		;OnMessage(0x200,"admGui_ThumbnailView")

	; Journalimport läuft gerade
		If Addendum.ImportRunning {
			res := admGui_ShowTab("Journal", admHTab)
			admGui_ImportJournalAll()
		}

return

admLAN:                                                                                          	;{ zeichnet den LAN Tab

	; Array mit verfügbaren Netzwerkgeräten
		LanClients := ClientsOnline()

		GuiControlGet, cpos, adm: Pos, AdmLanT1
		Gui, adm: Font  	, s8 q5 Normal cBlue, Arial

	; Anzeigen der verfügbaren Geräte und deren Optionen
		For ClientIndex, onlineclient in LanClients {

			onlineclient := StrReplace(onlineclient, "-")
			YPlus := (ClientIndex = 1 ? 5 : 2)
			Y := cposY + cposH + YPlus

			Gui, adm: Add, Progress	, % "x10 y" Y     	 " w114 cCCCC00 vAdmLanP" ClientIndex               	, 100
			Gui, adm: Add, Text      	, % "x17 y" Y + 2 " w110 BackgroundTrans vAdmLanC" ClientIndex  	, % onlineclient

			GuiControlGet, cpos, adm: Pos, % "AdmLanP" ClientIndex
			GuiControl, adm: Move, % "AdmLanP"  ClientIndex, % "h" cposH + 4

			Gui, adm: Add, Checkbox	, % "x+3 y" Y " h" cposH + 4 " w10 vAdmLanAdm" ClientIndex
			GuiControlGet, cpos, adm: Pos, % "AdmLanAdm" ClientIndex

			If Addendum.LAN.Clients[onlineclient].remoteShotdown
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
				YPlus := (OfflineIndex = 1 ? 5 : 2)
				Y := cposY + cposH + YPlus
				Gui, adm: Add, Progress	, % "x10 y" Y      	" w114 cGray vAdmLanP" ClientIndex                 	, 100
				Gui, adm: Add, Text      	, % "x17 y" Y + 2	" w110 BackgroundTrans vAdmLanC" ClientIndex  	, % clientName
				GuiControlGet, cpos, adm: Pos, % "AdmLanP" ClientIndex
				GuiControl, adm: Move, % "AdmLanP"  ClientIndex, % "h" cposH + 4
			}

		}

return ;}

}

adm_Tabs:                                                                                        	;{ Tab Handler

	If (A_GuiEvent = "Normal") {
		; aktueller Tab ist Info, Countdown auffrischen
			SendMessage, 0x130B,,,, % "ahk_id " admHTab
			If (ErrorLevel = 5)  ; Info Tab angewählt
				admGui_CountDown()
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

	} else if RegExMatch(A_GuiControl, "admRDP_")
			Run, % q Addendum.LAN.Clients[compname].rdpPath "\" StrReplace(A_GuiControl, "admRDP_") ".rdp" q

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
					Sleep 1000
					WinWaitActive, % "Dokumentexport ahk_class AutoHotkeyGUI",, 7
					If (EModulHwnd := WinExist("Dokumentexport ahk_class AutoHotkeyGUI"))
						Send_WM_CopyData("search|" AlbisAktuellePatID() "|" GetAddendumID() "|admGui_Callback", EModulHwnd)
				}

			}

		case "TX":
			EToolExe    	:= Addendum.Tools[MNr].command
			EToolName	:= Addendum.Tools[MNr].name
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
		admGui_Default("admJournal")
		admFile	:= LV_GetSelected("admJournal")
		EventInfo	:= A_EventInfo

		If IsObject(admFile)
			GuiControl, adm: , admButton2, % "Importieren [" admFile.Count() "]"
		else
			GuiControl, adm: , admButton2, % "Importieren"

	; Kontextmenu
		    	If InStr(A_GuiEvent 	, "RightClick")  	{
			If (StrLen(admFile) = 0) && !IsObject(admFile)
				return
			MouseGetPos, mx, my, mWin, mCtrl, 2
			If IsObject(admFile)
				Menu, admJCM1, Show, % mx - 20, % my + 10
			else
				Menu, admJCM, Show, % mx - 20, % my + 10
			return
		}
	; Spaltensortierung
		else 	If 	InStr(A_GuiEvent	, "ColClick")    	{
			If (EventInfo > 0)
				admGui_Sort(EventInfo)
			return
		}
	; Dokumente importieren
		else 	If 	InStr(A_GuiControl, "admButton2")	{
			admGui_ImportJournalAll(admFile)
			return
		}
	; Befundordner indizieren
		else 	If	InStr(A_GuiControl, "admButton3") {
			admGui_Reload()
			return
		}
	; OCR starten/abbbrechen
		else 	If	InStr(A_GuiControl, "admButton4") {
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
		If !FileExist(Addendum.BefundOrdner "\" admFile) || IsObject(admFile)
			return

	; PDF/Bild-Programm aufrufen
		If Instr(A_GuiEvent	, "DoubleClick")
			admGui_View(admFile)

return ;}

adm_TP:                                                                                             	;{ Protokoll Tab 	- Button Handler

	GuiControlGet, day, adm:, AdmTPTag

	If          	InStr(A_GuiControl, "admTPZurueck")	{             	; einen Tag zurück
		admGui_TProtokoll(day, -1, compname)
	} else If 	InStr(A_GuiControl, "admTPVor")    	{             	; einen Tag weiter
		If (day = "Heute")
			return
		admGui_TProtokoll(day, 1, compname)
	}


return ;}

adm_TProtokollLV:                                                                               	;{ Tagesprotokoll 	- Listview Handler

	admGui_Default("admTProtokoll")

	If Instr(A_GuiEvent, "DoubleClick") {
		LV_GetText(PatID, EventInfo:= A_EventInfo, 4)
		AlbisAkteOeffnen("", PatID)
	}

return ;}

adm_GuiDropFiles:                                                                             	;{ nicht erkannte Dokumente einfach in das Posteingangfenster ziehen!

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

		If (PatDocs.MaxIndex() = 0) || Addendum.Importing 	; versehentlichen Aufruf bei leerem Array verhindern
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
			admGui_LaborblattDruck(admLBDPD, admLBDCB)

return ;}


; -------- Gui-Funktionen
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

	; TCM_GETCURSEL (0x130B)     	: ausgewählten TAB ermitteln
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

admGui_Sort_old(EventNr, LV_Init=false)                          	{               	; sortiert die Journalspalten und zeigt ein Symbol für die Sortierreihenfolge an

		; Funktion wird gebraucht für die Wiederherstellung der letzten Sortierung und für das Sichern der Einstellungen bei Nutzerinteraktion
		; LV_Init - nutzen um gespeicherte Sortierungseinstellung wiederherzustellen

		global 	admHJournal
		static 	JCol1Dir, JCol2Dir, LVSortStr, JColDir := []

		admGui_Default("admJournal")

		If LV_Init {

			RegExMatch(Addendum.iWin.JournalSort, "\s*(\d)\s*(\d)", JSort)
			EventNr := JSort1

			If (JSort1 = 1) {
				If (JSort2 = 1)
					JCol1Dir := true
			}
			else {
				If (JSort2 = 1)
					JCol3Dir := true
			}

			If JCol1Dir
				JCol3Dir := false

			If JCol3Dir
				JCol1Dir := false

		}

	; Sortierung je nach gewählter Spalte vornehmen
	; Idee von: https://www.autohotkey.com/boards/viewtopic.php?t=68777
		If      	(EventNr = 3) {    ; Spalte 3 (Eingangsdatuim) - sortiert wird nach der unsichtbaren Spalte 4 (Timestamp)

			admGui_ColSort("admJournal", EventNr, (JCol3Dir := !JCol3Dir))
			Addendum.iWin.JournalSort := EventNr " " (JCol2Dir ? "1":"0")

		}
		else If	(EventNr = 1) {

			admGui_ColSort("admJournal", EventNr, (JCol1Dir := !JCol1Dir))
			Addendum.iWin.JournalSort := EventNr " " (JCol1Dir ? "1":"0")

		}


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

admGui_Watcher()                                                         	{                	; Timerfunktion. Änderungen im Albisfenster werden leider nicht zuverlässig erkannt

	global hadm

	If !WinExist("ahk_class OptoAppClass") || !Addendum.AddendumGui
		return

	Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
	If RegExMatch(Addendum.AktiveAnzeige, "i)Karteikarte|Laborblatt|Biometriedaten") {
		WinGet, winList, ControlList, % "ahk_id" AlbisGetActiveMDIChild()
		If InStr(winlist, "AutohotkeyGui1")
			return
		else
			AddendumGui()
	}

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

	;~ for i, message in LISTENERS
		;~ OnMessage(message, "")

return
}


; -------- Inhalte erstellen
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


; -------- TAB Anzeigen auffrischen
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

	; PatDocs erstellen - enthält nur die Dateien zum aktuellen Patienten
		For key, pdf in ScanPool	{				;wenn keine PatID vorhanden ist, dann ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)
			If !RegExMatch(pdf.name, "\.pdf$") || !FileExist(Addendum.BefundOrdner "\" pdf.name) {
				pdfpool.remove(pdf.name)
				continue
			}
			RegExMatch(pdf.name, "^\s*(?<Nachname>" rxPerson2 ")[\,\s]+(?<Vorname>" rxPerson2 ")", doc)  ; ## ersetzen
			a := StrDiff(PatNV, RegExReplace(docNachname docVorname, "[\s\-]"))
			b := StrDiff(PatNV, RegExReplace(docVorname docNachname, "[\s\-]"))
			If (a < 0.11) || (b < 0.11)
				PatDocs.Push(pdf)
		}

	; Pdf Befunde anzeigen
		For key, pdf in PatDocs {
			If !RegExMatch(pdf.name, "\.pdf$") || !FileExist(Addendum.BefundOrdner "\" pdf.name)
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

		ImageImport := false
		PdfImport   	:= false

		admGui_Default("admJournal")
		LV_Delete()

	; SCANPOOL OBJECT UND "PdfIndex.json" DATEI AUFFRISCHEN
		If (refreshDocs || FirstRun) {
			FirstRun := false
			newPDF := BefundIndex()
		}

	; PDF BEFUNDE DES PATIENTEN HINZUFÜGEN
		Addendum.iWin.OCREnabled := false
		For key, pdf in ScanPool	{
			If FileExist(filePath := Addendum.BefundOrdner "\" pdf.name) {
				LV_Add((pdf.isSearchable ? "ICON3" : "ICON1"), pdf.name, pdf.pages, pdf.filetime, pdf.timeStamp) 	; anderes Symbol für durchsuchbare PDF Dateien
				If !pdf.isSearchable && !Addendum.iWin.OCREnabled
					Addendum.iWin.OCREnabled := true
			}
		}

	; OCR BUTTON
		If Addendum.iWin.OCREnabled {
			If Addendum.Thread["tessOCR"].ahkReady()
				admGui_OCRButton("+OCR abbrechen")
			else
				admGui_OCRButton("+OCR ausführen")
		}
		else
			admGui_OCRButton("-OCR ausführen")

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

	datestring := datestring = "Heute" || StrLen(datestring) = 0 ? (A_DD "." A_MM "." A_YYYY) : (datestring)
	RegExMatch( datestring, "(?<dd>\d+)\.(?<MM>\d+)\.(?<YYYY>\d+)", d)
	lDate := dYYYY dMM ddd
	If (addDays != 0)
		lDate += addDays, Days

	RegExMatch(ldate, "^(?<YYYY>\d{4})(?<MM>\d{2})(?<dd>\d{2})", d)
	FormatTime, weekdayshort, % lDate, ddd
	FormatTime, weekdaylong	, % lDate, dddd
	FormatTime, month       	, % lDate, MMMM

	TPFullPath	:= Addendum.DBPath "\Tagesprotokolle\" dYYYY "\" dMM "-" month "_TP.txt"
	If !FileExist(TPFullPath)
		return

	SectionDate := weekdaylong "|" ddd "." dMM
	IniRead, TProto, % TPFullPath, % SectionDate, % client
	If InStr(TProto, "Error")
		TProto := ""

	If InStr(TProto, ",")
		TProto := StrSplit(RTrim(TProto, ","), ",")
	else
		TProto := StrSplit(TProto, ">", "<")

	LV_Delete()
	For idx, PatIDTime in TProto {
		RegExMatch(PatIDTime, "^\d+", PatID)
		If (StrLen(PatID) = 0)
			continue
		LV_Insert(1,, SubStr("00" idx, -1) , oPat[PatID].Nn ", " oPat[PatID].Vn, oPat[PatID].Gd, PatID)
	}

	LV_ModifyCol(2, "Auto")
	If (LV_GetColWidth(admHTProtokoll, 2) < 115)
		LV_ModifyCol(2, 115)

	GuiControl, adm: , AdmTProtokollTitel	, % "[" compname "] [" TProto.MaxIndex() " Patienten]"
	If (SubStr(lDate, 1, 8) = SubStr(A_Now, 1, 8))
		GuiControl, adm: , AdmTPTag           	, % "Heute"
	else
		GuiControl, adm: , AdmTPTag           	, % weekdayshort ", " ddd "." dMM "." dYYYY

}

admGui_InfoText(TabTitel)                                             	{               	; Zusammenfassungen aktualisieren

	; letzte Änderung: 15.02.2021

	global adm, admButton1, admButton2
	global admJournalTitle, admPatientTitle, admTProtokollTitel
	global PatDocs

	Gui, adm: Default

	If (TabTitel = "Journal") || (TabTitel = "Patient")	{

		; Journal
			InfoText := (ScanPool.MaxIndex() = 0) ? " keine Dokumente" : (ScanPool.MaxIndex() = 1) ? "1 Dokument" : ScanPool.MaxIndex() " Dokumente"
			admGui_Default("admJournal")
			GuiControl, adm: , admJournalTitle, % InfoText
			If (LV_GetCount() > 0)
				GuiControl, adm: Enable, admButton2
			else
				GuiControl, adm: Disable, admButton2

		; Patient
			InfoText := "Posteingang: (" (!PatDocs.MaxIndex() ? "keine neuen Befunde)" : PatDocs.MaxIndex() = 1 ? "1 Befund)" : PatDocs.MaxIndex() " Befunde)") ", insgesamt: " ScanPool.MaxIndex()
			admGui_Default("admReports")
			GuiControl, adm: , admPatientTitle, % InfoText
			If (LV_GetCount() > 0)
				GuiControl, adm: Enable, admButton1
			else
				GuiControl, adm: Disable, admButton1

	}

	If InStr(TabTitel, "TProtokoll") 	{
		InfoText := "[" compname "] [" TProtokoll.MaxIndex() " Patienten]"
		GuiControl, adm: , admTProtokollTitel, % InfoText
	}

}


; -------- Kontextmenu und zugehörige Funktionen
admGui_CM(MenuName)                                               	{                 	; Kontextmenu des Journal

		; fehlt: nachschauen ob im Befundordner ein Backup-Verzeichnis angelegt ist
		; auch für Tastaturkürzelbefehle in allen Tabs (01.07.2020)

		global 	admHJournal, admHReports, admFile, rcEvInfo, hadm, PatDocs
		static 	newadmFile, rowNr

		Addendum.iWin.firstTab	:= "Journal"
		If RegExMatch(MenuName, "^J")
			admGui_Default("admJournal")

		rowSel  	:= LV_FindRow("admJournal", 1, admFile)
		blockthis 	:= false
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
			admGui_View(admFile)
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

			MsgBox,0x1024, Addendum für AlbisOnWindows, % "Wollen Sie die Datei:`n" admFile "`nwirklich löschen?"
			IfMsgBox, No
				return

		; kopiert ins Backupverzeichnis bevor gelöscht wird, Datei kann so wiederhergestellt werden
			If !FileExist(Addendum.BefundOrdner "\backup\" admFile)
				FileCopy, % Addendum.BefundOrdner "\" admFile, % Addendum.BefundOrdner "\backup\" admFile
			FileDelete, % Addendum.BefundOrdner "\" admFile
			FileDelete, % Addendum.BefundOrdner "\Text\" RegExReplace(admFile, "\.pdf$", ".txt")

		; Datei aus der Listview entfernen
			admGui_Default("admJournal")
			LV_Delete(rowSel)

		; Ordner neu einlesen
			admGui_Journal()
			admGui_Reports()

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
				else                                                                                                                                                  	{
					MsgBox, 4	, Addendum für Albis on Windows, %	">> OCR Text vorhanden <<`n" admFile "`n"
																						. 	"Bei erneuter Ausführung werden die bisherigen Textdaten gelöscht.`n"
																						.	"Dennoch ausführen?"
					IfMsgBox, No
						return
					PraxTT("Die Texterkennung der Datei < " admFile " > wird jetzt ausgeführt.", "1 0")
					Sleep 1000
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
			admCM_JRenAll()
		}
		else if	InStr(MenuName, "JRename")	{	; manuelles Umbenennen
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile)
		}
		else if	InStr(MenuName, "JSplit")   	{	; manuelles Aufteilen einer PDF Datei
			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile	,	"Aufteilen der PDF-Seiten auf mehrere Datein.`n"
													.	"Schreiben Sie: 1-2 D1, 3-4 D2`n"
													.	"Datei 1 erhält die Seiten 1-2 und Datei 2 die Seiten 3-4`n" )
		}
		else if	InStr(MenuName, "JImport")	{	; Datei in geöffnete Karteikarte importieren
			If blockthis {
				PraxTT("Die wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_ImportFromJournal(admFile)
		}

return
}

admCM_JRecog(admFile)                                              	{              	; Kontextmenu: JRecog - einzelner Datei automatisch einen Namen geben

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
		fDates 	:= FindDocDate(PdfText)
		admGui_ImportGui(false, "Autonaming Dokument`n" admFile "`nläuft", "admJournal")
		If IsObject(fNames)  {

			if (fNames.Count() = 1) {
				For PatID, Pat in fNames {
					newfilename := fNames[PatID].Nn ", " fNames[PatID].Vn ", " ;(StrLen(withoutname) = 0 ? "Befund" : withoutname)	; " [" PatID "]
					break
				}

				If IsObject(fDates)
					If fDates.Behandlung[1]
						newfilename .= "v. " fDates.Behandlung[1]
					else
						newfilename .= "v. " fDates.Dokument[1]

			; wird umbenannt
				admGui_Rename(admFile,, newfilename ".pdf")
				return

			}
			else {
				For PatID, Pat in found
					t .= "   [" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"
				SciTEOutput("  Das Dokument konnte nicht eindeutig zugeordnet werden. Folgende Namen stehen zu Auswahl:`n" t)
			}
		}
		else {

			PraxTT("Dokument: " admFile "`nkonnte keinem Patienten zugeordnet werden", "4 1")

		}

return
}

admCM_JRenAll()                                                           	{              	; Kontextmenu: automatisiert die Benennung aller unbenannten PDF-Dateien

	/*  mögliche Varianten Datumsangaben

		Therapiekontrolle bei COPD am 15.12.20
		CT horn and Akatomee mit LV, Kontrastmittelgabe vom 16.12.2020;
		ambulante Vorstellung zuletzt am 14.12.2020.
		MRT der BWS vom 14.12.2020
		Aufnahme am: 20.12.2020 um 21:08 Uhr
		Stationärer Aufenthalt: \n\r Vom 17.12.2020 bis zum 20.12.2020 auf unserer neurochirurgischen Station M115B

	 */

	; Variablen
		static newadmFile, lastnewadmFile

		filenamechanged := false

	; Dateinamen unbenannter PDF-Dateien einsammeln
		If !IsObject(unnamed := GetUnamedDocuments()) {
			PraxTT("Keine Dateien für eine automatische Umbennung vorhanden!", "1 2")
			return
		}

	; Patienten und Dokumentdatum finden
		docPath := Addendum.BefundOrdner
		For idx, filename in unnamed {

			; Pfade festlegen
				pdfPath  	:= docPath "\" filename
				bupPath 	:= docPath "\Backup\" filename
				txtPath   	:= docPath "\Text\" StrReplace(filename, ".pdf", ".txt")
				counter  	:= idx "/" unnamed.Count()

			; optional Text extrahieren
				If PDFisSearchable(pdfPath) && !FileExist(txtpath) {

					; Fortschritt anzeigen
						admCM_JRnProgress([counter, filename, "Text wird extrahiert....", "", "", lastnewadmFile])

					; pdftotext extrahiert den Text aus der PDF Datei
						cmdline := Addendum.PDF.xpdfPath "\pdftotext.exe"
														. 	" -f 1 -l " PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
														. 	" -bom -enc UTF-8"
														.	" " q pdfPath q
														. 	" " q txtpath q
						stdout := StdOutToVar(cmdline)

						If !FileExist(txtpath) {
							;SciTEOutput(" [" cmdline "]`n" stdout)
							continue
						}

				} else if !PDFisSearchable(pdfPath) {
					; Protokollierung
						PraxTT("Datei " idx " " filename " ist nicht durchsuchbar" )
						continue
				}

			; extrahierten Text lesen
				doctext := FileOpen(txtpath, "r").Read()
				If (StrLen(Trim(doctext)) = 0) {
					PraxTT("Der extrahierte Text enthält keine Zeichen.`n[" admFile "]", "2 2")
					FileDelete, % txtPath
					continue
				}

			; Fortschritt anzeigen
				words := StrSplit(doctext, " ").MaxIndex(), chars := StrLen(RegExReplace(doctext, "[\s\n\r\f]"))
				admCM_JRnProgress([counter, filename, "Textinhalt wird analysiert....", words, chars, lastnewadmFile])

			; Patientennamen und Behandlungs- oder Dokumentdatum finden
				fNames	:= FindDocNames(doctext)
				fDates 	:= FindDocDate(doctext)

			; Datei umbenennen
				If (IsObject(fNames) && fNames.Count() = 1) {

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
							If fDates.Behandlung[1]
								newadmFile .= "v. " fDates.Behandlung[1] " "
							else
								newadmFile .= "v. " fDates.Dokument[1] " "

					; Anzeige von Wort und Zeichenzahl der Datei
						admCM_JRnProgress([counter, filename, "neuen Dateinamen erstellt", words, chars, lastnewadmFile])
						lastnewadmFile := newadmFile

					; Debug
						 If ((res := Weitermachen("Datei: " counter "`nBez.: " newadmFile "`nWeitermachen?") <> 1))
							break

					; Umbenennen von Original, Backup und Textdatei
						If (StrLen(newadmFile) > 0) {
							filenamechanged := true
							FileMove, % pdfPath	, % docPath "\" newadmFile ".pdf"           	, 1               	; Originaldatei umbenennen
							FileMove, % bupPath	, % docPath "\Backup\" newadmFile ".pdf"	, 1               	; Backupdatei umbenennen
							FileMove, % txtPath	, % docPath "\Text\" newadmFile ".txt"     	, 1           		; zugehörige Text Datei umbenennen
							newadmFile := ""
						}

				}
				else if (IsObject(fNames) && fNames.Count() < 1) {
					PraxTT("Dokument: " filename "`nkonnte keinem Patienten zugeordnet werden", "4 1")
				}
				else If (IsObject(fNames) && fNames.Count() > 1) {
					t := ""
					For PatID, Pat in fNames
						t .= " `t[" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"

					SciTEOutput("  [" idx "/" unnamed.Count() "]`n  1: " filename "`n  2: #keine eindeutige Zuordnung möglich")
					SciTEOutput("  Namen:       `t" 	fNames.Count())
					SciTEOutput("  Beh. Datum:`t" 	fDates.Behandlung.Count())
					SciTEOutput("  Doc. Datum:`t" 	fDates.Dokument.Count())
					SciTEOutput(t)
					SciTEOutput("  -------------------------------------------------------------------" )
				}
				else {
					PraxTT("Dokument: " filename "`nkonnte keinem Patienten zugeordnet werden", "4 1")
				}

		}

	; Anzeige schliessen
		admGui_ImportGui(false, "Autonaming Dokument`n" admFile "`nläuft", "admJournal")

	; Anzeige auffrischen
		If filenamechanged {
			admGui_Journal(true)
			PatDocs := admGui_Reports()
			RedrawWindow(hadm)
		}

return
}

admCM_JRnProgress(ProgressText)                                 	{              	; zeigt den Fortschritt beim Umbennen an

		static outmsgTitle := "Automatische Dateinamengenerierung`n`n"

	; nur eine Zahl übergeben dann wird die Progressbar verändert
		If !RegExMatch(ProgrressText, "^\d+$") {

			outmsg :="Zähler:   `t"		ProgressText.1      	"`n"
						.	"Name:   `t"  	ProgressText.2      	"`n"
						. 	"Status:   `t"     	ProgressText.3     	"`n"
						. 	"Wörter:  `t"   	ProgressText.4	   	"`n"
						. 	"Zeichen:`t"   	ProgressText.5   	"`n"
						. 	"`n- - - letzter Fund - - -`n" newadmFile

			admGui_ImportGui(true, outmsgTitle . outmsg, "admJournal")

		}

}


; -------- Zusatzfunktionen
admGui_TPSuche()                                                        	{                	; wann war ein Patient da

	TPFullPath	:= Addendum.DBPath "\Tagesprotokolle"
	Gui, adm: Submit, NoHide

	If (StrLen(admTPSuche) = 0)
		return

	;~ Loop, Files, *.txt, R
	;~ {


	;~ }

}

admGui_LaborblattDruck(Printer, Druckspalten)                 	{                	; automatisiert den Laborblattdruck, stellt vorherige Ansicht wieder her

		static AlbisViewType

		AlbisViewType := AlbisGetActiveWindowType(true)

		If InStr(Printer, "Microsoft Print to PDF")
			savePath := Addendum.ExportOrdner "\Laborwerte-" RegExReplace(AlbisCurrentPatient(), "[\s]") ".pdf"
		else
			savePath := ""

		If !(res := AlbisLaborblattExport(Druckspalten, savePath, Printer))
			PraxTT("Der Ausdruck von Laborwerten ist fehlgeschlagen.`nFehlercode: " res, "4 1")

		If (AlbisViewType <> AlbisGetActiveWindowType(true))
			AlbisKarteikartenAnsicht(AlbisViewType)

}


; -------- PDF Dateien umbenennen / teilen
admGui_Rename(filename, prompt="", newfilename="") 	{               	; Dialog für Dokument umbenennen oder teilen

		global

		local fSize, edW	, txtfile, row, rowFilename
		local cp, dp, gs, admPos, title, fileout, vctrl, ENr
		local key, pdf, TLines, found, monIndex, admScreen, RNx, RNWidth

		static RNminWidth := 350
		static BEZPath, Beschreibung
		static wT, Patname, Inhalt, Datum, Bezeichner, oPV
		static oldadmFile, newadmFile, FileExt, FileOutExt, oldfilename
		static rxMonths := "\(Jan*u*a*r*|Febr*u*a*r*|Mä*a*rz|Apr*i*l*|Mai|Jun*i*|Jul*i*|Aug*u*s*t*|Sept*e*m*b*e*r*|Okt*o*b*e*r*|Nov*e*m*b*e*r*|Dez*e*m*b*e*r*)"


	; PDF Vorschau anzeigen in neuem Sumatrafenster oder wenn geöffnet in diesem Fenster
		If !IsObject(oPV) {
			oPV := admGui_PDFPreview(fileName)
			If !IsObject(oPV)
				return
		} else {
			If (oPV.Previewer = "Sumatra") {
				SumatraDDE(oPV.ID, "OpenFile"	, Addendum.BefundOrdner "\" filename, 0, 1, 0)
				Sleep 400
				SumatraDDE(oPV.ID, "SetView"	, Addendum.BefundOrdner "\" filename, "single page", "fit page")
			}
			else
				return
		}

	; Vorbereitung: Dokument zerlegen
		If Instr(prompt, "Aufteilen") {
			title    	:= "Dokument zerlegen"
			fileout	:= ""
		}
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

	; Datei mit Inhaltsbezeichnungen laden
		If !FileExist(BEZPath := Addendum.DBPath "\Dictionary\Befundbezeichner.txt") {
			FilePathCreate(Addendum.DBPath "\Dictionary")
		} else {
			Bezeichner := FileOpen(BEZPath, "r", "UTF-8").Read()
			Bezeichner := StrSplit(Bezeichner, "`n", "`r")
		}

	; Gui wird bei erneutem Aufruf nicht erneut erstellt
		If WinExist("Addendum (Datei umbenennen) ahk_class AutoHotkeyGUI") {
			gosub RNEFill
			return
		}

; Rename Gui                                     	;{

		wT          	:= 0
		edW     	:= 100
		fSize     	:= (A_ScreenWidth > 1920 ? 10 : 9)
		admPos	:= GetWindowSpot(hadm)
		RNWidth	:= admPos.W - (2*admPos.BW)
		RNWidth	:= RnWidth < RNminWidth ? RNminWidth : RnWidth

		Gui, RN: new, % "AlwaysOnTop -DpiScale -ToolWindow -Caption +Border HWNDhadmRN"
		Gui, RN: Margin, 5, 5

		;-: Hinweise
		Gui, RN: Font, % "s" fSize - 1
		Gui, RN: Add, Text, % "xm ym Center vRNC1"                     	, % "[ Pfeil ⏶& ⏷ von Feld zu Feld, Eingabe für Umbenennen ]"

		;-: Eingabefelder
		Gui, RN: Font, % "s" fSize
		Gui, RN: Add, Text, % "xm y+8 Right vRNT1 "                                             	, % "Name"
		Gui, RN: Add, Edit, % "x+5 w" edW " r1 vRNE1 HWNDRNhE1 gRNEHandler"	, % Patname
		Gui, RN: Add, Text, % "xm y+5 Right vRNT2 " 	                                         	, % "Inhalt"
		Gui, RN: Add, Edit, % "x+5 w" edW " r1 vRNE2 HWNDRNhE2 gRNEHandler"	, % Inhalt
		Gui, RN: Add, Text, % "xm y+5 Right vRNT3 "                                             	, % "Datum"
		Gui, RN: Add, Edit, % "x+5 w" edW " r1 vRNE3 HWNDRNhE3 gRNEHandler"	, % Datum

		;-: Name, Inhalt, Datum - Steuerelemente auf gleiche Breite bringen, Edit-Felder maximieren
		Loop, 3 {
			cp 	:= GuiControlGet("RN", "Pos", "RNT" A_Index)
			wT	:= cp.W > wT ? cp.W : wT
		}
		Loop, 3 {
			GuiControl, RN:Move, % "RNT" A_Index, % "w" wT
			cp := GuiControlGet("RN", "Pos", "RNT" A_Index)
			GuiControl, RN:Move, % "RNE" A_Index, % "x" cp.X + wT + 5 " y" cp.Y - 3 " w" RNWidth - wT - 15
		}

		;-: maximale Zeichenzahl
		cp := GuiControlGet("RN", "Pos", "RNE3")
		Gui, RN: Font, % "s" fSize - 1 " cWhite", Calibri
		Gui, RN: Add, Progress	, % "x0      	y" cp.Y+cp.H+5 " w10 h20 c7B89BB vRNPG1 HWNDRNhPG1"     	, 100
		Gui, RN: Add, Text    	, % "x" cp.X "	y" cp.Y+cp.H+7 "  Left Backgroundtrans vRNHW1"                     	, % "verbrauchte Zeichen (Inhalt+Datum):"
		dp := GuiControlGet("RN", "Pos", "RNHW1")
		Gui, RN: Font, % "s" fSize - 1 " cWhite Bold", Consolas
		chars := (StrLen(Inhalt) + StrLen(Datum))
		chars := SubStr("00" chars, -1) . " / 70"
		Gui, RN: Add, Text, % "x+5                                          	Left Backgroundtrans vRNLen "                        	, % chars
		GuiControl, RN: Move, % "RNPG1", % "h" dp.H + 5

		;-: Dateinamenvorschau
		Gui, RN: Font, % "s" fSize - 1 " cBlack Normal", Arial
		Gui, RN: Add, Progress	, % "x0 y" dp.Y+dp.H+5 " w10 h20 cAfC2ff vRNPG2 HWNDRNhPG2"             	, 100
		Gui, RN: Add, Text, % "x5 y" dp.Y+dp.H+7 " w" edW " h" dp.H*2 " Center	Backgroundtrans vRNPV "     	, % admGui_FileName(Patname, Inhalt, Datum)
		cp := GuiControlGet("RN", "Pos", "RNPV")
		GuiControl, RN: Move, % "RNCancel", % "h" cp.H + dp.H + 10

		;-: OK, Abbruch
		Gui, RN: Font, % "s" fSize
		Gui, RN: Add, Button, % "xm                         	Center vRNOK  	gRNProceed"         	, % "Umbennen"
		Gui, RN: Add, Button, % "x+20                  	Center vRNCancel	gRNProceed"         	, % "Abbruch"
		cp := GuiControlGet("RN", "Pos", "RNE3"), dp := GuiControlGet("RN", "Pos", "RNCancel")
		GuiControl, RN: Move, % "RNCancel", % "x" cp.X + cp.W - dp.W

		;-: Show Hide
		monIndex  	:= GetMonitorIndexFromWindow(hadm)
		admScreen	:= ScreenDims(monIndex)
		RNx := ((admPos.X - 2*admPos.BW) + RnWidth) > admScreen.W ? admScreen.W - RnWidth : admPos.X - 2*admPos.BW

		Gui, RN: Show, % "x" RNx " y" admPos.Y + admPos.H " w" RNWidth " Hide", Addendum (Datei umbenennen)

		;-: Titel und Preview anpassen
		gs := GetWindowSpot(hadmRN)
		GuiControl, RN: Move, % "RNC1"	, % "w"   	gs.CW - 5
		GuiControl, RN: Move, % "RNPV"	, % "w"  	gs.CW - 5
		GuiControl, RN: Move, % "RNE1"	, % "w"  	gs.CW - wT - 15
		GuiControl, RN: Move, % "RNE2"	, % "w"  	gs.CW - wT - 15
		GuiControl, RN: Move, % "RNE3"	, % "w"  	gs.CW - wT - 15

		;-: Zeichenzähler zentrieren
		dp := GuiControlGet("RN", "Pos", "RNHW1")
		cp := GuiControlGet("RN", "Pos", "RNLen")
		dcpLen := dp.W + 5 + cp.W
		dcpMid := Floor(dcpLen/2)
		gsMid	:= Floor(gs.w/2)
		gsLeft	:= (gsMid - dcpMid)
		GuiControl, RN: Move, % "RNHW1"	, % "x" gsLeft
		GuiControl, RN: Move, % "RNLen"  	, % "x" gsLeft + dp.W + 5

		;-: Abbruch Button verschieben
		dp := GuiControlGet("RN", "Pos", "RNCancel")
		GuiControl, RN: Move, % "RNCancel", % "x" gs.W - dp.W - 5

		;-: Progress anpassen
		cp := GuiControlGet("RN", "Pos", "RNPV")
		GuiControl, RN: Move, % "RNPG1"	, % "x0 w" gs.W
		GuiControl, RN: Move, % "RNPG2"	, % "x0 w" gs.W
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG1
		WinSet, ExStyle, 0x0, % "ahk_id " RNhPG2

		;-: Show
		Gui, RN: Show

		Hotkey, IfWinActive, % "Addendum (Datei umbenennen)"
		Hotkey, Escape 	, RNGuiClose
		Hotkey, Up    	, RNGuiUpDown
		Hotkey, Down	, RNGuiUpDown
		Hotkey, Enter	, RNProceed
		Hotkey, IfWinActive

		;}

RNEFill:                                               	;{   Inhalte einfügen

	; Editfelder befüllen und Vorschau anzeigen
		GuiControl, RN:, RNPV	, % admGui_FileName(Patname, Inhalt, Datum)
		GuiControl, RN:, RNE1	, % PatName
		GuiControl, RN:, RNE2	, % Inhalt
		GuiControl, RN:, RNE3	, % Datum

	; Focus je nach Inhalt der Eingabefelder setzen
		If !PatName
			GuiControl, RN: Focus, RNE1
		else
			GuiControl, RN: Focus, RNE2

return ;}

RNEHandler:                                      	;{   Dateinamenvorschau

	Gui, RN: Submit, NoHide

	RegExMatch(A_GuiControl, "\d", nr)

	GuiControl, RN:, % "RNLen", % SubStr("00" (charsNow := StrLen(RNE2) + StrLen(RNE3)), -1) " / 70 "
	If (charsNow > 70) {
		BlockInput, Send
		xc := charsNow - 70
		Send, % "{BackSpace " xc "}"
		BlockInput, Off
	}

	Gui, RN: Submit, NoHide
	GuiControl, RN:, RNPV, % admGui_FileName(RNE1, RNE2, RNE3)

return ;}

RNProceed:                                       	;{   Datei wird umbenannt, PDF-Viewer-Vorschau wird geschlossen

	; Abbruch ohne Dateinamenänderung
		If (A_GuiControl = "RNCancel") {
			gosub RNGuiClose
			return
		}	else if (A_GuiControl <> "RNOK") && (A_ThisHotkey <> "Enter")
			return

		Gui, RN: Submit, NoHide
		newadmFile := admGui_FileName(RNE1, RNE2, RNE3)

	; Nutzer hat OK oder Enter gedrückt, aber nichts geändert
		If ((newadmFile . FileExt) = filename) {
			PraxTT("Sie haben den Dateinamen nicht geändert!", "2 1")
			gosub RNGuiBeenden
			return
		}

	; Nutzer hat alles gelöscht
		If (StrLen(newadmFile) = 0) {
			PraxTT("Ihr neuer Dateiname enthält keine Zeichen.", "2 1")
			gosub RNGuiBeenden
			return
		}

	; neue Inhaltsbeschreibung speichern
		found := false, TLines := ""
		For idx, Beschreibung in Bezeichner {
			TLines .= Beschreibung "`n"
			If (Beschreibung = RNE2) {
				found := true
				break
			}
		}
		If !found {
			TLines .= RNE2
			Sort, TLines
			FileOpen(BEZPath, "w", "UTF-8").Write(TLines)
		}

	; Originaldatei umbennen
		newadmFile := newadmFile "." FileExt
		FileMove, % Addendum.BefundOrdner "\" oldadmFile, % Addendum.BefundOrdner "\" newadmFile, 1
		If ErrorLevel {
			MsgBox, 0x1024, Addendum für Albis on Windows
									, % 	"Das Umbenennen der Datei`n"
									.	"  <" oldadmFile ">  `n"
									. 	"wurde von Windows aufgrund des`n"
									.	"Fehlercode (" A_LastError ") abgelehnt."
			gosub RNGuiBeenden
			return
		}

	; Backup Datei umbennen
		If (fileext = "pdf") && FileExist(Addendum.BefundOrdner "\Backup\" oldadmFile)
			FileMove, % Addendum.BefundOrdner "\Backup\" oldadmFile, % Addendum.BefundOrdner "\Backup\" newadmFile, 1

	; Textdokument umbennen
		txtfile := Addendum.BefundOrdner "\Text\" RegExReplace(oldadmFile, "\.\w+$", ".txt")
		If (FileExt = "pdf") && FileExist(txtfile)
			FileMove, % txtfile, % Addendum.BefundOrdner "\Text\" RegExReplace(newadmFile, "\.pdf$", ".txt"), 1

	; ScanPool auffrischen
		For key, pdf in ScanPool
			If (pdf.name = oldadmFile) {
				ScanPool[key].name := newadmFile
				break
			}

	; Journal auffrischen
		admGui_Default("admJournal")
		Loop % LV_GetCount() {
			LV_GetText(rowFilename, row := A_Index, 1)
			If Instr(rowFilename, oldadmFile) {
				LV_Modify(row,, newadmFile)
				break
			}
		}

;}

RNGuiBeenden:                                   	;{   weitere Dateien umbenennen

	; nachschauen ob eine weitere Datei zum Umbenennen zur Verfügung steht
		admGui_Default("admJournal")

	; nächste Datei finden
		rows := LV_GetCount()
		LV_GetNext(row)
		startrow := row := row = 0 ? 1 : row
		 ToolTip, Suche nächstes Dokument
		Loop {

			row := row = rows ? 1 : row + 1
			LV_Modify(row = 1 ? rows : row-1, "-Select")
			LV_Modify(row, "Select")
			Sleep 200
			If (row = startrow)
				break
			If (A_Index > rows) {
				startrow := row
				break
			}

			LV_GetText(filename, row, 1)
			If !string.isFullNamed(FileName)
				break

		}
		ToolTip

	; weiteres Dokument umbenennen?
		If (row <> startrow) && (filename <> oldfilename) {
			oldfilename := filename
			MsgBox, 0x4, Addendum für Albis on Windows, % "Möchten Sie mit der nächsten`n[" filename "]`nDatei fortfahren ?", 30
			IfMsgBox, Yes
			{
				admGui_Rename(filename)
				return
			}
		}
;}

RNGuiClose:
RNGuiEscape:                                    	;{   Fenster definitiv schliessen

	; Gui beenden
		Gui, RN: Destroy

	; PDF Previewer beenden
		If (oPV.Previewer = "Sumatra")
			oPV := admGui_SumatraClose(oPV.PID, oPV.ID)

	; Hotkeys freigeben
		;~ Hotkey, IfWinActive, % "Addendum (Datei umbenennen)"
		;~ Hotkey, Up, Off
		;~ Hotkey, Down, Off
		;~ Hotkey, Enter, Off
		;~ Hotkey, IfWinActive

	; Journal fokussieren
		admGui_Default("admJournal")
		GuiControl, adm: Focus, AdmJournal

		oPV := ""

return ;}

RNGuiUpDown:                                   	;{   mit den Pfeiltasten zwischen den Eingabefeldern wechseln

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

admGui_SumatraClose(PID, ID="")                                 	{                	; Hilfsfunktion: admGui_Rename

	Process, Close, % PID
	If !ErrorLevel
		SumatraInvoke("Exit", ID)

}

admGui_FileName(PatName, inhalt, Datum)                  	{                	; Hilfsfunktion: admGui_Rename

	retStr := RegExReplace(PatName, "[\s,]+$") ", "
	retStr .= Trim(RegExReplace(inhalt, "[\s,]+$"))
	retStr .= StrLen(Datum) > 0 ? " v. " Trim(RegExReplace(Datum, "^[\s,v\.]+")) : ""

return retStr
}


; --------- PDF/Bild Viewer
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
			pdfPath	:= Addendum.BefundOrdner "\" fileName

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
				WinWait        	, % SmtraClass,, 6
				WinActivate		, % SmtraClass
				WinWaitActive	, % SmtraClass,, 1

				hSumatra := WinExist(SmtraClass)
				WinGet      	,	SumatraPID, PID, % "ahk_id " hSumatra
				ControlGet	,	hSumatraCnvs, HWND,, SUMATRA_PDF_CANVAS1, % "ahk_id " hSumatra
				WinGetPos	,,	stY,, stH, ahk_class Shell_TrayWnd
				;WinSet      	,	Style, 0x16C40000, % "ahk_id " hSumatra  ; Minimieren, Maximieren und Beenden werden entfernt

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

admGui_View(filename)                                                 	{               	; Befund-/Bildanzeigeprogramm aufrufen

	filepath := Addendum.BefundOrdner "\" filename
	If (StrLen(filename) = 0) || !FileExist(filepath)
		return 0

	If RegExMatch(filename, "\.jpg|png|tiff|bmp|wav|mov|avi$") {
		PraxTT("Anzeigeprogramm wird geöffnet", "3 0")
		Run % q filepath q
	}
	else {
		If !FileExist(pdfReaderPath := Addendum.PDF.ReaderAlternative)
			If !FileExist(pdfReaderPath := Addendum.PDF.Reader) {
				PraxTT("Es konnte kein PDF-Anzeigeprogramm gefunden werden", "2 1")
				return
			}
		PraxTT("PDF-Anzeigeprogramm wird geöffnet", "3 0")
		If !PDFisCorrupt(filepath)
			Run % q pdfReaderPath q " " q filepath q
		else
			PraxTT("PDF Datei:`n>" fileName "<`nist defekt", "2 0")
	}


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


; -------- Importieren
admGui_ImportGui(Show=true, msg="", gctrl="")              	{                 	; Fenster für laufende Prozesse

	global adm, adm2, hadm, hadm2, admImporter, admImporterPrg, admJournal, last_msg, last_ctrl
	global admHTab

	If (StrLen(gCtrl) = 0) {
		gCtrl := "admReports"
		res := admGui_ShowTab("Patient", admHTab)
	}
	else if (gCtrl = "admJournal") {
		res := admGui_ShowTab("Journal", admHTab)
	}

	If Show {

		; zeigt den Hinweis auch an wenn das Infofenster neu gezeichnet wurde
			If Addendum.Importing && (StrLen(msg) > 0) {
				gctrl := last_ctrl
				If (last_msg <> msg) {
					GuiControl, adm2:, AdmImporter, % msg
					last_msg := msg
				}
				return
			}
			else if !Addendum.Importing {
				last_msg := msg, last_ctrl 	:= gctrl
				Addendum.Importing := true
				If StrLen(msg) = 0
					msg := "Importvorgang läuft...`nbitte nicht unterbrechen!"
			}

			GuiControlGet, rpos, adm: Pos, % gCtrl

			Gui, adm2: New	, +HWNDhadm2 +ToolWindow -Caption +Parent%hadm%
			Gui, adm2: Color	, cAA0000
			Gui, adm2: Font	, s10 q5 cFFFFFF Bold, % Addendum.Default.Font

			MsgOpt := InStr(msg, "Dateinamen") ? "" : "Center", PrgOpt := ""
			Gui, adm2: Add	, Text    	, % "x30 y1 w" (rposW-30) " " MsgOpt " vadmImporter"  	, % msg

			PrgOpt := ""
			;Gui, adm2: Add	, Progress	, % "x30 y" rposH-30 " w" rposW-30 " vAdmImporterPrg"

			Gui, adm2: Show	, % "x" rposX " y" rposY " w" rposW " h" rposH " Hide NA"              	, admImportLayer

			GuiControlGet, cpos, adm2: Pos, admImporter
			GuiControl, adm2: Move, admImporter, % "y" Floor(rposH // 2 - cposH // 2)
			Gui, adm2: Show	, % "NA"

			WinSet, Style         	, 0x50020000	, % "ahk_id " hadm2
			WinSet, ExStyle	    	, 0x0802009C	, % "ahk_id " hadm2
			WinSet, Transparent	, 180             	, % "ahk_id " hadm2

	} else {

			Gui, adm2: Destroy
			last_msg := last_ctrl := ""
			Addendum.Importing := false

	}

}

admGui_ImportJournalAll(files="")                                    	{               	; Journal: 	Import aller vollständig benannten Dateien

		global admHTab

		admGui_Default("admJournal")
		GuiControl, % "adm: Enable0"	, AdmButton2
		GuiControl, % "adm:"            	, AdmButton2, % "...importiere"

		Addendum.ImportRunning := true
		waituser := 10
		Loop % LV_GetCount() {                                                ; Dokumente ohne Personennamen werden ignoriert

			 ; enthält keinen Personennamen dann weiter
				LV_GetText(rowFile, A_Index, 1)
				If !string.isFullNamed(rowFile)
					continue

			; weiter mit Nachfrage
				If (Addendum.iWin.Imports > 0) {
					MsgBox, 0x1021, % "Addendum", % "Weiteres Dokument importieren?`n[" rowFile "]", % (A_Index < 3 ? waituser : waituser/2 )
					IfMsgBox, Cancel
					{
						Addendum.ImportRunning	:= false
						break
					}
				}

			; Importieren jetzt
				filePath := Addendum.BefundOrdner "\" rowFile
				If RegExMatch(filePath, "\.pdf$") && FileExist(filePath) && !FileIsLocked(filePath) && PDFisSearchable(filepath) {

					If FuzzyKarteikarte(rowFile) {
						If admGui_ImportFromJournal(rowFile) {
							Addendum.iWin.Imports ++
							return
						}
					}
					else {
						PraxTT("Es konnte kein passender Patient gefunden werden.", "3 0")
						break
					}
				}

		}

		Addendum.iWin.Imports	:= 0
		Addendum.ImportRunning	:= false

		GuiControl, % "adm:"            	, AdmButton2, % "Importieren"
		If (LV_GetCount() > 0)
			GuiControl, % "adm: Enable1"	, AdmButton2

		admGui_InfoText("Journal")
		admGui_InfoText("Patient")

return
}

admGui_ImportFromJournal(filename)                              	{               	; Journal: 	Einzelimport-Funktion

		global 	admHPDFfilenames, hadm, admHTab
		static 	WZEintrag	:= {"LASV":"Anfragen", "Anfrage":"Anfragen", "Antrag":"Anfragen", "Lageso":"Anfragen", "Lebensversicherung":"Anfragen"}

	; Nutzer befragen
		currPat 	:= AlbisCurrentPatient()
		Pat		:= string.Get.Names(filename)
		If Addendum.iWin.ConfirmImport || (!InStr(currPat, Pat.Nn) && !InStr(currPat, Pat.Vn)) {
			MsgBox, 262148	, % "Befund importieren?"
										, % "Wollen Sie die Datei:`n"
										. 	 (StrLen(filename) > 30 ? SubStr(filename, 1, 30) "..." : filename) "  dem Pat.`n"
										. 	currPat " zuordnen ?"
			IfMsgBox, No
				return 0
		}

	; Dokumentdatum aus dem Dateinamen entnehmen, falls enthalten
		DocDate := string.Get.DocDate(filename)

	; Befund importieren
		If      	RegExMatch(filename, "\.pdf")                                             	{
				Addendum.FuncCallback := "GetPDFViewerArg"
				If !AlbisImportierePdf(filename, DocDate)
					return 0
				admGui_MoveFile(filename)             	; PDF in einen anderen Ordner verschieben
				docID := pdfpool.Remove(filename)	; aus ScanPool entfernen
		}
		else If	RegExMatch(filename, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{
				If !AlbisImportiereBild(filename, string.Replace.Names(filename), DocDate)
					return 0
		}
		else
			return 0

	; Re-Indizieren, Gui auffrischen, Importprotokoll
		FileAppend, % datestamp() "| " filename "`n", % Addendum.Befundordner "\PdfImportLog.txt"
		while, % (StrLen(Addendum.FuncCallback) > 0) {
			Sleep, 50
			If (A_Index > 60)  ; 3 Sekunden
				break
		}

		admGui_Journal(false)
		PatDocs := admGui_Reports()
		admGui_ShowTab("Journal", admHTab)
		admGui_InfoText("Journal")

return 1
}

admGui_ImportFromPatient()                                            	{               	; Patient:	Befundimport alle Befunde

	; letzte Änderung: 17.02.2021

		global 	PatDocs

		Docs 	:= Object()

	; Importschleife
		For key, file in PatDocs		{

			Addendum.FuncCallback := ""
			If !FileExist(Addendum.BefundOrdner "\Backup\" file.name)
				continue

			DocDate 	:= string.Get.DocDate(file.name)
			KKText    	:= string.Karteikartentext(file.name)

			; Bild importieren
			If RegExMatch(file.name, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{

				FileCopy, % Addendum.BefundOrdner "\" file.name, % Addendum.BefundOrdner "\Backup\Importiert\" file.name, 1
				If !AlbisImportiereBild(file.name, KKText, DocDate) {
					Docs.Push(file)
					continue
				}

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
					admGui_MoveFile(file.name)
					docID := pdfpool.Remove(file.name)

				; wartet auf den FoxitReader
					while, % (StrLen(Addendum.FuncCallback) > 0) {
						Sleep, 50
						If (A_Index > 40)
							break
					}

				SciTEOutput("FuncCallback: " Addendum.FuncCallback)

			}

			FileAppend, % datestamp() "| " file.name "`n", % Addendum.Befundordner "\PdfImportLog.txt"
			admGui_Default("admReports")
			Loop % LV_GetCount() {
				LV_GetText(celltext, A_Index, 1)
				If (imported = celltext) {
					LV_Delete(A_Index)
					break
				}
			}
		}

	; PatDocs  leeren und mit nicht verarbeiteten Dateien füllen
		Loop % PatDocs.MaxIndex()
			PatDocs.RemoveAt(1)
		For index, file in Docs
			PatDocs.Push(file)

return Imports
}

admGui_MoveFile(filename)                                             	{               	; behandelt die PDF Backup Dateien

	; Backup PDF in anderen Ordner verschieben
		If FileExist(Addendum.BefundOrdner "\Backup\" filename)
			FileMove, % Addendum.BefundOrdner "\Backup\" filename, % Addendum.BefundOrdner "\Backup\Importiert\" filename, 1

	; PDF-Text in anderen Ordner verschieben
		txtfilename := RegExReplace(filename, "\.pdf$", ".txt")
		If FileExist(Addendum.BefundOrdner "\Text\" txtfilename)        ; PDF in einen anderen Ordner verschieben
			FileMove, % Addendum.BefundOrdner "\Text\" txtfilename, % Addendum.BefundOrdner "\Text\Backup\" txtfilename, 1

}


; -------- Karteikarte
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
		If (Patients.Count() = 1) {

			For PatID, Patient in Patients
				return admGui_Karteikarte(PatID)

		} else {

			SciTEOutput(NameStr ": " Pat.Nn ", " Pat.Vn)
			SciTEOutput(" (matches: " Patients.Count() ") mindiff: " m.minDiff ", [" m.bestID "] " m.bestNn ", " m.bestVn)
			For PatID, Pat in Patients
				SciTEOutput(" [" PatID "] " Pat.Nn ", " Pat.Vn)
			return 0
		}

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


; -------- Gui Einstellungen speichern
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

	; letzte Änderung: 13.02.2021

		global admHJournal, adm, hadm, admTPTag, admTabs

	; Gui auslesen
		Gui, adm: Submit, Hide
		GuiControlGet, ProtokollTag, adm:, admTPTag
		JournalSort := LV_GetScrollViewPos(admHJournal)

	; bei Bedarf in der ini speichern
		If (Addendum.iWin.firstTab <> admTabs)
			IniWrite, % admTabs     	, % Addendum.Ini, % compname, % "Infofenster_aktuelles_Tab"
		If (Addendum.iWin.JournalSort <> JournalSort)
			IniWrite, % JournalSort  	, % Addendum.Ini, % compname, % "Infofenster_JournalSortierung"
		If (Addendum.iWin.TProtDate <> ProtokollTag)
			IniWrite, % ProtokollTag	, % Addendum.Ini, % compname, % "Infofenster_Tagesprotokoll_Datum"

	; sichern im Addendum Objekt
		Addendum.iWin.firstTab    	:= admTabs
		Addendum.iWin.TProtDate 	:= ProtokollTag
		Addendum.iWin.JournalSort	:= JournalSort

}


; -------- Callback Funktionen
admGui_CheckJournal(pdffile, ThreadID)                          	{                	; Ausführung jedesmal nach erfolgreicher Erstellung einer PDF durch TesseractOCR()

		global admHJournal

		admGui_Default("admJournal")

	; Farb-Markierung entfernen
		If (rowNr := LV_FindRow("admJournal",1, pdffile)) {
			LV_Modify(rowNr, "ICON3")                                        	; ICON für durchsuchbare PDF hinzufügen
			;admGui_ColorRow(admHJournal, rowNr, false)          	; Hintergrundfarbe der Zeile wird entfernt
		}

	; Datei aus dem Stapel entfernen
		For idx, inprogress in Addendum.iWin.FilesStack
			If (inprogress = pdfFile) {
				Addendum.iWin.FilesStack.RemoveAt(idx)
				break
			}

	; ScanPool Array ändern und BefundIndex sichern
		oPDF := pdfpool.Add(Addendum.BefundOrdner, pdfFile)

	; Nachricht an den OCR-Thread: OCR Vorgang fortsetzen
		Send_WM_CopyData("continue", ThreadID)

}

admGui_FolderWatch(path, changes)                             	{                	; wird aufgerufen wenn neue PDF-Dateien im Befundordner vorhanden sind

	; letzte Änderung: 13.02.2021

		global hadm

		static staticFileCount	:= 0
		static filecount        	:= 0
		static func_AutoOCR	:= func("admGui_OCRAllFiles")

	; untersucht alle veränderten oder hinzugfügten Dateien
		For Each, Change In Changes {

			action   	:= change.action
			name    	:= change.name

			If RegExMatch(name, "\.pdf$")
				If RegExMatch(action, "1|3|4")	{

					SplitPath, name, filename, filepath
					If IsObject(oPDF := pdfpool.Add(filepath, filename)) {

						; neue Datei zu Listview hinzufügen und Info's ändern
							;Controls("", "reset", "")
							;If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {
							If hadm {
								admGui_Default("admJournal")
								LV_Add((oPDF.isSearchable ? "ICON3" : "ICON1"), oPDF.name, oPDF.pages, oPDF.filetime, oPDF.timeStamp)
								admGui_InfoText("Journal")
							}

						; Dateizähler erhöhen
							staticFileCount 	++
							filecount        	++

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
					If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {
						admGui_Default("admJournal")
						row := LV_FindRow("admJournal", 1, oPDF.name)
						LV_Delete(row)
						admGui_InfoText("Journal")
					}

				}
		}

		If filecount {
			msg := filecount " Befund" (filecount = 1 ? " wurde" : "e wurden" ) " hinzufügt!"
			Controls("", "reset", "")
			If Controls("", "ControlFind, AddendumGui, AutoHotkeyGUI, return hwnd", "ahk_class OptoAppClass") {
				admPos := GetWindowSpot(hadm)
				Splash(msg, 6, admPos.X+30, admPos.Y + admPos.H - 30, admPos.W - 60, hadm)
			}
			else
				TrayTip("Befundordner", msg, 6)

		}
		filecount := 0

}

class pdfpool                                                                 	{               	; verwaltet das ScanPool Objekt

	; Abhängigkeiten - Addendum_PdfHelper.ahk

	; sucht nach Dateinamen im Pool
		inPool(filename) {

			For docID, PDF in ScanPool
				If (PDF.name = filename)
					return docID

		return 0
		}

	; fügt Datei mit allen notwendigen Informationen dem Scanpool hinzu
		Add(path, filename) {

			If !this.inPool(filename) {

				FileGetSize	, FSize       	, % (pdfPath := path "\" filename), K
				FileGetTime	, timeStamp 	, % pdfPath, C
				FormatTime	, FTime     	, % timeStamp, dd.MM.yyyy

				oPDF := {"name"          	: filename
							, 	"filesize"        	: FSize
							, 	"timestamp"    	: timeStamp
							, 	"filetime"        	: FTime
							, 	"pages"         	: PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
							, 	"isSearchable"	: (PDFisSearchable(pdfPath)?1:0)}

				ScanPool.Push(oPDF)

			return oPDF
			}

		return
		}

	; Datei entfernen
		Remove(filename) {
			If (docID := this.inPool(filename))
				return ScanPool.RemoveAt(docID)
		return
		}

	; ScanPool komplett leeren
		Empty() {

			Loop % (mxIdx := ScanPool.MaxIndex())
				ScanPool.Pop()

		return mxIdx
		}

	; ScanPool speichern/laden
		Load(path) {
			return JSONData.Load(path "\PdfDaten.json", "", "UTF-8")
		}

		Save(path) {
			return JSONData.Save(path "\PdfDaten.json", ScanPool, true,, 1, "UTF-8")
		}

}

GetPDFViewerArg(class, hwnd)                                       	{               	; commandline des PDF Anzeigeprogramms auslesen

	WinGet PID, PID, % "ahk_id " hwnd

    StrQuery := "SELECT * FROM Win32_Process WHERE ProcessId=" . PID
    Enum := ComObjGet("winmgmts:").ExecQuery(StrQuery)._NewEnum
    If (Enum[Process]) {
		Viewercmdline := Process.CommandLine
		;SciTEOutput("Foxit: " Viewercmdline)
	}
	Addendum.FuncCallback := ""

	;SciTEOutput("Foxit: " Viewercmdline)
	; FileAppend, % datestamp() " - " report "`n", % Addendum.Befundordner "\PdfImportLog.txt"
}


; ------------------------ Texterkennung
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
							tessedit_create_boxfile        	1
							tessedit_create_hocr           	1
							tessedit_create_pdf            	1
							tessedit_create_tsv               	1
							tessedit_create_txt               	1

							load_bigram_dawg	                    	True
							tessedit_enable_bigram_correction		True
							tessedit_bigram_debug	                  	3
							save_raw_choices	                         	True
							save_alt_choices	                         	True
					)

		tessconfig2 =
					(LTrim Join|
							tessedit_create_pdf            	1
							tessedit_create_tsv               	1
							tessedit_create_txt               	1
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
		global threadcontrol
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
					sleep, 100
			}
		}
		else {
			filesmax := 1
			threadcontrol := true
			SciteOutPut(" ")
			pdfText := tessOCRPdf(files, settings)
			Send_WM_CopyData("OCR_processed|" files "|" hOCR, settings.ScriptID)
			While threadcontrol
				sleep, 100
		}

		Send_WM_CopyData("OCR_ready|" filesmax "|" hOCR, settings.ScriptID)
		;SciteOutPut("send message to: " settings.ScriptID " - processing is done. ")

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
		For key, val in Settings {
			repset .= q key q ":" . ((val = "true") || (val = "false") || RegExMatch(val, "^(\d+|0x\w+)$") ? val : q val q) ", "
		}
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

	If Addendum.OCR.AutoOCR && (Addendum.OCR.Client = compname) && Addendum.Thread["tessOCR"].ahkReady() {
		Addendum.OCR.RestartAutoOCR := true
		admGui_OCRButton("+OCR abbrechen")
		return
	}

	noOCR := GetNoOCRFiles()
	If (noOCR.MaxIndex() > 0) {

		For key, filename in noOCR
			Addendum.iWin.FilesStack.Push(admFile)

		admGui_OCRButton("+OCR abbrechen")
		Addendum.OCR.RestartAutoOCR := false
		TesseractOCR(noOCR)

	}
	else
		admGui_OCRButton("-OCR ausführen")

return
}

GetNoOCRFiles()                                                           	{               	; stellt PDF Dateien ohne Text zusammen für TesseractOCR

	global admHJournal

	noOCR := Array()

	For idx, pdf in ScanPool
		If !pdf.isSearchable {

			; Datei ist in Bearbeitung (z.B. OCR Vorgang), dann ignorieren
				InStack := false
				For stackindex, stackfile in Addendum.iWin.FilesStack
					If (stackfile = pdf.name) {
						InStack := true
						break
					}

			; Datei hinzufügen
				If !Instack {
					noOCR.Push(pdf.name)
					;admGui_ColorRow(admHJournal, LV_FindRow("admJournal", 1, pdf.name), true) 			; kann per Flag ausgeschaltet werden
				}

		}

return noOCR
}

GetUnamedDocuments(refreshDocs=false)                      	{                	; erstellt einen Array mit den Dateinamen unbezeichneter Dokumente

	; ScanPool ist global
	unnamed := Array()

	If refreshDocs
		admGui_Reload(true)

	For idx, pdf in ScanPool
		If !string.ContainsName(pdf.name)
			unnamed.Push(pdf.name)

	If (unnamed.MaxIndex() = 0)
		return

return unnamed
}


; ------------------------ Dateien indizieren
BefundIndex(save=true)                                                  	{	               	;-- erststellt das ScanPool Objekt neu

	; WICHTIG!: braucht eine globale Variable im aufrufenden Skript: ScanPool := Object()
		global admPatientTitle, admJournalTitle

	; ScanPool nicht neu indizieren, nur die Einträge einzeln löschen, damit es kein neues Objekt wird
		pdfpool.Empty()

	; alle pdf Dokumente des Befundordners dem ScanPool Objekt hinzufügen
		files  	:= GetFilesInDir(Addendum.BefundOrdner, "*.pdf")
		iLen  	:= StrLen(files.MaxIndex())-1
		fcount	:= SubStr("00000" files.MaxIndex(), -1*iLen)

	; gefundene Dateien dem ScanPool Objekt hinzufügen, den Fortschritt anzeigen
		For idx, filename in files {
			pdfpool.Add(Addendum.BefundOrdner, filename)
			InfoText := "indiziere Dokument: " SubStr("00000" A_Index, -1*iLen) " von " fcount
			GuiControl, adm: , admPatientTitle, % InfoText
			GuiControl, adm: , admJournalTitle	, % InfoText
		}

return ScanPool.MaxIndex()
}

GetPDFData(path, pdfname)                                           	{               	;-- erstellt Daten über PDF-Dateien für das ScanPool-Objekt

	pdfPath := path "\" pdfname
	FileGetSize	, FSize       	, % pdfPath, K
	FileGetTime	, timeStamp 	, % pdfPath, C
	FormatTime	, FTime     	, % timeStamp, dd.MM.yyyy

return {	"name"          	: pdfname
		, 	"filesize"          	: FSize
		, 	"timestamp"   	: timeStamp
		, 	"filetime"       	: FTime
		, 	"pages"          	: PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
		, 	"isSearchable"	: (PDFisSearchable(pdfPath)?1:0)}
}


; ------------------------ Hilfsfunktionen
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

GetLastGVU(PatID) {

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


; ------------------------ Debugfunktionen
ParseStdOut(stdOut, indent="    `t")                               	{

	For i, line in StrSplit(stdOut, "`n", "`r")
		t .= StrLen(line) > 0 ? indent . line "`n" : ""

return RTrim(t, "`n")
}


; ------------------------ Gui / Fenster Funktionen
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
	cRow:=0, JFiles:=Array()

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

return dllcall("RedrawWindow", "Ptr", (hwnd = 0 ? hadm : hwnd), "Ptr", 0, "Ptr", 0, "UInt", RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN)
}


GetFilesInDir(path, filepattern:="*.*", options:="")               	{

	files := Array()
	Loop, Files, % path "\" filepattern, % (options ? options : "")
		files.Push(A_LoopFileName)

return files
}

SumatraDDE(hSumatra, cmd, params*) 								{         			;-- Befehle an Sumatra per DDE schicken

	/*  DESCRIPTION

		https://github.com/sumatrapdfreader/sumatrapdf/issues/1398
		https://gist.github.com/nod5/4d172a31a3740b147d3621e7ed9934aa
		functions Send_WM_COPYDATA() and RECEIVE_WM_COPYDATA() are required
		Required data to tell SumatraPDF to interpret lpData as DDE command text, always 0x44646557

		SumatraPDF DDE command unicode text, https://www.sumatrapdfreader.org/docs/DDE-Commands.html

		DDE Commands
		Sumatra can be controlled in a limited way from other software by sending DDE commands. They are mostly
		used to use SumatraPDF as a preview tool from e.g. LaTeX editors that generate PDF files.

		Format of DDE comands
			Single DDE command:   	[Command(parameter1, parameter2, ..., )]
			Multiple DDE commands: 	[Command1(parameter1, parameter2, ..., )][Command2(...)][...]

		List of DDE commands:
        	[Open file]
			format:     	[Open("<pdffilepath>"[,<newwindow>,<focus>,<forcerefresh>])]
			arguments:	if newwindow is 1 then a new window is created even if the file is already open
								if focus is 1 then the focus is set to the window
								if forcerefresh is 1 the command forces the refresh of the file window if already open
								(useful for files opened over network that don't get file-change notifications)".
			example:   	[Open("c:\file.pdf", 1, 1, 0)]

			[Forward-Search]
			format: [ForwardSearch(["<pdffilepath>",]"<sourcefilepath>",<line>,<column>[,<newwindow>,<setfocus>])]
			arguments:
			pdffilepath:     	path to the PDF document (if this path is omitted and the document isn't already open,
	                    			SumatraPDF won't open it for you)
			column:         	this parameter is for future use (just always pass 0)
			newwindow:  	1 to open the document in a new window (even if the file is already opened)
			focus:             	1 to set focus to SumatraPDF's window.
			examples:     	[ForwardSearch("c:\file.pdf","c:\folder\source.tex",298,0)]
                                   	[ForwardSearch("c:\folder\source.tex",298,0,0,1)]

           	[GotoNamedDest]
           	format:         	[GotoNamedDest("<pdffilepath>","<destination name>")]
           	example:       	[GotoNamedDest("c:\file.pdf", "chapter.1")]
           	note:             	the pdf file must be already opened

           	[Go to page]
           	format:         	[GotoPage("<pdffilepath>",<page number>)]
           	example:       	[GotoPage("c:\file.pdf", 37)]
           	note:             	the pdf file must be already opened.

        	[SetView]
   			format: 			[SetView("<pdffilepath>","<view mode>",<zoom level>[,<scrollX>,<scrollY>])]
   			arguments:
   			view mode: 		"single page"
   									"facing"
    								"book view"
    								"continuous"
    								"continuous facing"
    								"continuous book view"
   			zoom level : 		either a zoom factor between 8 and 6400 (in percent) or one
	                            	of -1 (Fit Page), -2 (Fit Width) or -3 (Fit Content)
   			scrollX, scrollY: 	PDF document (user) coordinates of the point to be visible in the top-left of the window
   			example: 			[SetView("c:\file.pdf","continuous",-3)]
   			note: 				the pdf file must already be opened
    		Example:			[SetView("c:\file.pdf","continuous",-3)]

	 */

		static dwData := 0x44646557

		lpData := { 	"OpenFile"         	: ("[Open(""p1"",p2,p3,p4)]")
														; p1=filepath, p2=sourcefilepath, p3=1 for focus, p4=1 for force refresh

														; [p1=filepath,] p2=sourcefilepath, p3=line, p4=column[, p5=1 for new window, p6=1 to set focus]
						,	"ForwardSearch" 	: ("[ForwardSearch(""p1"",""p2"",p3,p4,p5,p6)]")

														; p1=filepath, p2=destination name
						,	"GotoNamedDest"	: ("[GotoNamedDest(""p1"",""p2"")]")

														; p1=filepath, p2=PageNr
						,	"GotoPage"        	: ("[GotoPage(""p1"",p2)]")

														; p1=filepath, p2=view mode, p3=zoom level [, p4=scrollX, p5=scrollY>]
						,	"SetView"           	: ("[SetView(""p1"",""p2"",p3,p4,p5)]")}

		For index, param in params
			lpData[cmd] := StrReplace(lpData[cmd], "p" index, param)

		lpData[cmd] := RegExReplace(lpData[cmd], ",*\""*\s*p\d\s*\""*\,*")

		;SciTEOutput(" " lpData[cmd])

return SendEx_WM_COPYDATA(hSumatra, dwData, lpData[cmd])
}

SendEx_WM_COPYDATA(hWin, dwData, lpData) 				{              	;-- für die Kommunikation mit Sumatra und SumatraDDE()

	VarSetCapacity(COPYDATASTRUCT, 3*A_PtrSize, 0)
    cbData := (StrLen(lpData) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(dwData	, COPYDATASTRUCT, 0*A_PtrSize)
    NumPut(cbData 	, COPYDATASTRUCT, 1*A_PtrSize)
    NumPut(&lpData	, COPYDATASTRUCT, 2*A_PtrSize)
	SendMessage, 0x4a, 0, &COPYDATASTRUCT,, % "ahk_id " hWin ; 0x4a WM_COPYDATA

return ErrorLevel == "FAIL" ? false : true
}

/*

CheckJournal() { ...
;~ For idx, pdf in ScanPool
	;~ If (pdf.name = pdfFile) {
		;~ ScanPool[idx] := GetPDFData(Addendum.BefundOrdner, pdfFile)
		;~ break
	;~ }

BefundIndex() { ....
ScanPool.Push(GetPDFData(Addendum.BefundOrdner, filename))

*/


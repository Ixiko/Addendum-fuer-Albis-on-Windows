; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                     INFOFENSTER -
;
;      Funktion:   	Verwaltung/Bearbeitung (Texterkennung/Kategorisierung) neuer gescannter Post vor dem Import in die Karteikarte
;                       	Export und Druck von Befunden direkt aus der Karteikarte
;                       	Anzeige von Praxisinformationen
;                       	Netzwerkkommunikation
;                       	erweitertes Tagesprotokoll
;      Basisskript: 	Addendum.ahk
;
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - last change 22.11.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; Infofenster
AddendumGui(admShow="") {

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Variablen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		global

		admLVPatOpt    	:= " r" (Addendum.InfoWindow.LVScanPool.r)
		admLVPatOpt    	.= " gadm_PdfReports   	vadmPdfReports 	HWNDadmHPDFReports 	AltSubmit -Hdr	ReadOnly	BackgroundF0F0F0 LV0x10020 0x5001804B "
		admLVJournalOpt	:= " gadm_Journal        	vadmJournal 		HWNDadmHJournal 	 	AltSubmit                      	BackgroundF0F0F0 LV0x10020 0x50210049 -E0x200 "
		admLVTProtokoll 	:= " gadm_TProtokollLV 	vadmTProtokoll 	HWNDadmHTProtokoll  	AltSubmit      	ReadOnly	BackgroundF0F0F0 LV0x10020 0x50210049 -E0x200"

		local LVContent
		local hbmPdf_Ico, hbmImage_Ico
		local adm, adm2, adm3, admWidth, admHeight, TabSize, TabSizeX, TabSizeY, TabSizeW, TabSizeH
		local APos, SPos, rpos, mox, moy, moWin, moCtrl, regExt, Aktiv, Pat
		local ClientName, LanClients, ClientIP, onlineclient, OffLineIndex, InfoText, found, YPlus, Y
		local admReCheck := 0

		If (Addendum.InfoWindow.Init = false) {

				Addendum.InfoWindow.Init ++
				Addendum.InfoWindow.FileStack   	:= Array()
				Addendum.InfoWindow.RowColors	:= false

				adm_importing	:= false
				admTabNames	:= "Patient|Journal|Protokoll|Info|Netzwerk"
				factor            	:= A_ScreenDPI / 96
				fc_admGui    	:= Func("AddendumGui")

				admHBM := Array()
				admHBM.Push(LoadPicture(Addendum.AddendumDir "\assets\ModulIcons\PDFImage.ico"))	; PDF ohne OCR
				admHBM.Push(LoadPicture(Addendum.AddendumDir "\assets\ModulIcons\image.ico"))    	; Bilddatei
				admHBM.Push(LoadPicture(Addendum.AddendumDir "\assets\ModulIcons\PDFOCR.ico"))	; PDF mit OCR

				;-: lädt die Icons
				admImageListID  	:= IL_Create(3)
				IL_Add(admImageListID, "HBITMAP:" admHBM[1], 0x00000) 		;1    - 0xFFFFFF
				IL_Add(admImageListID, "HBITMAP:" admHBM[2], 0xFFFFFF)		;2
				IL_Add(admImageListID, "HBITMAP:" admHBM[3], 0x00000) 		;3

				admHConnected   	:= LoadPicture(Addendum.AddendumDir "\assets\ModulIcons\connected.png")
				admHDisconnected	:= LoadPicture(Addendum.AddendumDir "\assets\ModulIcons\disconnected.png")

				;hbmPdf_Ico   	:= Create_PDF_ico()
				;hbmImage_Ico	:= Create_Image_ico()
				;~ IL_Add(ImageListID, "HBITMAP: " hbmPdf_Ico 		, 0xFFFFFF, 0) 		;1
				;~ IL_Add(ImageListID, "HBITMAP: " hbmImage_Ico	, 0xFFFFFF, 0)		;2

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Kontextmenu (Rechtsklickmenu) für das Journal
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				fc_admOpen 	:= Func("admGui_CM").Bind("JOpen")
				fc_admImport	:= Func("admGui_CM").Bind("JImport")
				fc_admRename	:= Func("admGui_CM").Bind("JRename")
				fc_admDelete	:= Func("admGui_CM").Bind("JDelete")
				fc_admView		:= Func("admGui_CM").Bind("JView")
				fc_admRefresh	:= Func("admGui_CM").Bind("JRefresh")
				fc_admRecog	:= Func("admGui_CM").Bind("JRecognize")
				fc_admRenAll	:= Func("admGui_CM").Bind("JRenAll")
				fc_admOCR  	:= Func("admGui_CM").Bind("JOCR")
				fc_admOCRAll	:= Func("admGui_CM").Bind("JOCRAll")

				Menu, admJCM, Add, Karteikarte öffnen                      	, % fc_admOpen
				Menu, admJCM, Add, Datei - anzeigen                        	, % fc_admView
				Menu, admJCM, Add, Datei - importieren                     	, % fc_admImport
				Menu, admJCM, Add, Datei - umbennen                      	, % fc_admRename
				Menu, admJCM, Add, Datei - löschen                           	, % fc_admDelete
				Menu, admJCM, Add, Datei - Texterkennung ausführen 	, % fc_admOCR
				Menu, admJCM, Add, Datei - Inhalt erkennen            	, % fc_admRecog
				Menu, admJCM, Add
				Menu, admJCM, Add, Texterkennung - alle Dateien       	, % fc_admOCRAll
				Menu, admJCM, Add, automatische Benennung           	, % fc_admRenAll
				Menu, admJCM, Add, Befundordner neu indizieren    	, % fc_admRefresh
			;}

			;----------------------------------------------------------------------------------------------------------------------------------------------
			; Hotkey's für alle Tab's
			;----------------------------------------------------------------------------------------------------------------------------------------------;{
				fc_admCMJournal := Func("GuiControlActive").Bind("AdmJournal", "adm")
				Hotkey, If, % fc_admCMJournal ; ahk_class AutoHotkeyGUI"
				Hotkey, Enter       	, % fc_admView
				Hotkey, F2        	, % fc_admRename
				Hotkey, F3         	, % fc_admView
				Hotkey, F4         	, % fc_admRecog
				Hotkey, F5        	, % fc_admRefresh
				Hotkey, F6        	, % fc_admOCR
				Hotkey, BackSpace, % fc_admDelete
				Hotkey, If
			;}

		}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Vorbereitungen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{

	; Albis wurde beendet -return-
		If !WinExist("ahk_class OptoAppClass")
			return

		Result := DllCall("User32\SetProcessDpiAwarenessContext", "UInt" , -1)

		If (Addendum.InfoWindow.lastPatID = AlbisAktuellePatID())
			return

	; mehr als 9 Versuche das MDIFrame zu finden, führen zum Abbruch des Selbstaufrufes der Funktion
		If (admReCheck > 9) {
			admReCheck := 0
			TrayTip, AddendumGui, % "Die Integration des Infofenster ist fehlgeschlagen.", 2
			return
		}

		Aktiv         	:= Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		res                := Controls("", "Reset", "")
		hMDIFrame	:= Controls("AfxMDIFrame1401"	, "ID", AlbisGetActiveMDIChild())
		hStamm		:= Controls("#327701"             	, "ID", hMDIFrame)
		If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten|Abrechnung|Rechnungsliste") || (GetDec(hMDIFrame) = 0) || (GetDec(hStamm) = 0) {
			SetTimer, % fc_admGui, -1000
			admReCheck ++
			return
		}

	; legt die Höhe der Gui anhand der Größe des MDIFrame-Fenster fest
		APos  	:= GetWindowSpot(AlbisWinID())
		SPos 	:= GetWindowSpot(hStamm)

		Addendum.InfoWindow.StammY:= Round(SPos.Y / factor)
		Addendum.InfoWindow.X	       	:= APos.CW - Addendum.InfoWindow.W ; ClientWidth passt besser!
		Addendum.InfoWindow.H       	:= admHeight := SPos.CH
		admWidth 	:= Addendum.InfoWindow.W

	; der Albisinfobereich (Stammdaten, Dauerdiagnosen, Dauermedikamente) wird für das Einfügen der Gui verkleinert
		If (SPos.W > APos.CW - admWidth) ; verhindert erneutes verkleinern
			SetWindowPos(hStamm, 2, 2, APos.CW - admWidth - 4, SPos.CH)
	;}

	;---------------------------------------------------------------------------------------------------------------------------------------------
	; Gui zeichnen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	; umgeht die seltene Problematik der eines nicht mehr vorhandenen Handle eines MDIFrame, wenn der Nutzer eine Karteikarte
	; schon geschlossen hat und das Skript noch dieses Handle erfasst hatte
		Sleep 100
		try {
			Gui, adm: New	, -Caption -DPIScale +ToolWindow 0x50020000 E0x0802009C +Parent%hMDIFrame% +HWNDhadm
		} catch {
			SetTimer, % fc_admGui, -1000
			admReCheck ++
			return
		}

		Gui, adm: Margin	, 0, 0
		Gui, adm: Add  	, Tab     	, % "x0 y0  	w" admWidth " h" admHeight " hwndadmHTab vadmTabs"                           	, % admTabNames

	;-: Tab1 :- Patient                          	;{
		Gui, adm: Tab  	, 1

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y27 w" admWidth-15 " BackgroundTrans vadmPdfReportTitel Section"                  	, % "(es wird gesucht ....)"

		Gui, adm: Font  	, s6 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-55 " y23 w50 vadmButton1 gadm_BefundImport"                                    	, % "Importieren"

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+2 	w" admWidth-6	" " admLVPatOpt                                                                   	, % "S|Befundname"
		LV_SetImageList(admImageListID)
		admCol1W :=35, admCol2W := admWidth - admCol1W - 10
		LV_ModifyCol(1, admCol1W " Integer Right NoSort")
		LV_ModifyCol(2, admCol2W " Text Left NoSort")

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text, % "x2 y+0 	w" admWidth-6 " Center vadmGP1"   	                                                             	, % "Notizen/Erinnerungen"

		GuiControlGet, cpos, adm: Pos, AdmGP1
		Gui, adm: Add  	, Edit     	, % "x2 y" cposY+12 " w" admWidth-6 " h" admHeight - cposY - cposH - 4 " gadm_Notes vadmNotes"

		;}

	;-: Tab2 :- Journal                         	;{
		Gui, adm: Tab  	, 2

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Text      	, % "x10 y25  	w" admWidth-22 " BackgroundTrans vadmJournalTitel Section"               	, % "(es wird gesucht ....)"

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth-145 " y23 h16 	vadmButton2 gadm_Journal"                                        	, % "Einzelimport"
		Gui, adm: Add  	, Button    	, % "x20 y23 h16                          	vadmButton3 gadm_Journal"                                        	, % "Aktualisieren"
		GuiControlGet  	, cpos, adm: Pos, AdmButton2
		GuiControlGet  	, dpos, adm: Pos, AdmButton3
		GuiControl        	, adm: Move, AdmButton2, % "x" admWidth - cposW - 5
		GuiControl        	, adm: Move, AdmButton3, % "x" admWidth - cposW - dposW - 10
		Gui, adm: Add  	, Button    	, % "x20 y23 h16                          	vadmButton4 gadm_Journal"                                        	, % "OCR ausführen  "
		GuiControlGet  	, cpos, adm: Pos, AdmButton3
		GuiControlGet  	, dpos, adm: Pos, AdmButton4
		GuiControl        	, adm: Move, AdmButton4, % "x" cposX - dposW - 20

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5 		w" admWidth - 6 " h" admHeight - 47 " " admLVJournalOpt                        	, % "Befund|Eingang|TimeStamp"
		admCol2W := 58, admCol1W := admWidth - admCol2W - 25, admCol3W := 0
		LV_ModifyCol(1, admCol1W " Left NoSort")
		LV_ModifyCol(2, admCol2W " Left NoSort")
		LV_ModifyCol(3, admCol3W " Left Integer")       	; versteckte Spalte enthält Integerzeitwerte für die Sortierung nach dem Datum
		Gui, adm: ListView, % "admJournal"
		LV_SetImageList(admImageListID)

	;}

	;-: Tab3 :- Protokoll (letzte Patienten)	;{
		Gui, adm: Tab  	, 3

		Gui, adm: Font  	, s7 q5 Normal cBlack, Arial
		ProtText := "[" compname "] [" TProtokoll.MaxIndex() " Patienten]"
		Gui, adm: Add  	, Text      	, % "x10 y25  	w" admWidth - 22 " BackgroundTrans vAdmTProtokollTitel"                     	, % ProtText

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, Button    	, % "x" admWidth - 225 " y23 h16      	vadmTPZurueck	 	gadm_TP"                              	, % "<"
		Gui, adm: Add  	, Button    	, % "x+2 y23 h16							    	vadmTPVor   		gadm_TP"                              	, % ">"
		Gui, adm: Add  	, Text      	, % "x+5 y25 w100 BackgroundTrans 	vAdmTPTag                    	"                                	, % Addendum.InfoWindow.TProtDate

		;Gui, adm: Add  	, Text      	, % "x+5 y25 w30 BackgroundTrans	 Right vAdmTPTag          	"                                     	, % "Heute"
		;Gui, adm: Add  	, Text      	, % "x+5 y25 w70 BackgroundTrans	vAdmTPDatum                 	"                                      	, % A_DD "." A_MM "." A_YYYY

		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial
		Gui, adm: Add  	, ListView	, % "x2 y+5 w" admWidth-6 " h" admHeight-47 " " admLVTProtokoll                                    	, % "RF|Nachname, Vorname|Geburtstag|PatID"

		col1W := 30, col2W := 160, col3W := 70, col4W := 70
		Gui, adm: ListView, AdmTProtokoll
		LV_ModifyCol(1, col1W )
		LV_ModifyCol(2, col2W )
		LV_ModifyCol(3, col3W )
		LV_ModifyCol(4, col4W )
	;}

	;-: Tab4 :- Info                              	;{
		Gui, adm: Tab  	, 4
		Gui, adm: Font  	, s8 q5 Normal cBlack, Arial

		InfoText	:= "⌚ Sprechstunde  : " ;{
		Sprechstunde := Addendum.Praxis.Sprechstunde[A_DDDD]
		InfoText .= ((StrLen(Sprechstunde) > 0) ? (A_DDDD ",der " A_DD "." A_MM "." A_YYYY " von " Sprechstunde " Uhr") : ( "heute keine Sprechstunde")) "`n"
		InfoText	.= "⚕ Tagesprotokoll: "	(TProtokoll.MaxIndex() = "" ? 0 : TProtokoll.MaxIndex()) " Patienten`n"
		InfoText	.= "⚕ Patienten ges. : "	oPat.Count() "`n"
		InfoText	.= "⚕ Signaturen      : "	Addendum.SignatureCount ", ges. Seitenzahl: " Addendum.SignaturePages "`n"
		InfoText	.= "☀ Monitoranzahl : "	Addendum.Monitor.Count() (IsRemoteSession() ? " (Remotedesktopsitzung!)" : "") "`n"
		Loop, % Addendum.Monitor.Count()
			InfoText .= "`t" SubStr("⚀⚁⚂⚃⚄⚅", A_Index, 1) " " Addendum.Monitor[A_Index] "`n"

		InfoText .= "`n"
		InfoText .= "⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯⚯`n"
		InfoText	.= "① Albis    `t: x"     	APos.X " y" APos.Y " w" APos.W " h" APos.H ", CW" APos.CW " CH" APos.CH "`n"
		InfoText	.= "② Stamm  `t: "     	hStamm ", x" SPos.X " y" SPos.Y " w" SPos.W " h" SPos.H " cw" SPos.CW " ch" SPos.CH "`n"
		InfoText	.= "③ nStammW: "     	(SPos.W - admWidth - 2) "`n"
		InfoText	.= "③ hTab   `t:"          	admHTab "`n"
		InfoText	.= "④ GuiSize`t: x"     	Addendum.InfoWindow.X " y" Addendum.InfoWindow.Y " w" Addendum.InfoWindow.W  " h" Addendum.InfoWindow.H "`n"
		InfoText	.= "⑤ Iconlist ID`t: 1=" GetHex(admHBM[1]) ", 2=" GetHex(admHBM[2]) ", 3=" GetHex(admHBM[3])  "`n"
		InfoText	.= "⑤ rcolors.`t: "     	(Addendum.InfoWindow.RowColors ? "true":"false") "`n"

		;}

		Gui, adm: Add  	, Edit     	, % "x2 y30 w" admWidth - 6 " h" admHeight - 32 " t6 t12 t18 ReadOnly -E0x200", % InfoText

	;}

	;-: Tab5 :-Netzwerk                       	;{
		Gui, adm: Tab  	, 5
		Gui, adm: Font  	, s8 q5 Normal underline cBlack, Arial

		;~ nprop	:= {	"title"         	: "admNet"
						;~ , 	"x"             	: 1
						;~ , 	"y"             	: 20
						;~ , 	"w"             	: Addendum.InfoWindow.W -	2
						;~ , 	"h"            	: Addendum.InfoWindow.H -	22
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

			Gui, adm: Add   	, Button	, % "x" dposX " y" dposY " vAdmClient_"	clientName " Center gadm_Lan", % clientName
			Gui, adm: Add   	, Button	, % "x150		  y" dposY " vAdmRDP_" 	clientName " Center gadm_Lan", % "RDP Sitzung starten"
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
			Gui, adm: Add, Pic, % "x10 y" cposY+2 " w" cposH-4 " h" cposH-4 " Backgroundtrans vAdmConn_" clientName, % "HBITMAP: " hBmDisconnected
		}

		;Gui, adm: Add    	, Text, % "x10 y30 w110 BackgroundTrans Center vAdmLanT1", % "ONLINE"
		;SetTimer, admLan, -300

	;-:          :-
		Gui, adm: Tab
		;}

	;-: Gui zeigen                                 	;{

		; Listview coloring -
			Gui, adm: ListView, AdmJournal
			If Addendum.InfoWindow.RowColors {
				LV_Colors.OnMessage()
				If !LV_Colors.Attach(admHJournal, true) {
					GuiControl, adm: +Redraw, AdmJournal
					SciTEOutput("LV_Color attach impossible")
				}
			}

		; Fenster anzeigen
			Gui, adm: Show	, % "x" Addendum.InfoWindow.X " y" Addendum.InfoWindow.Y " w" Addendum.InfoWindow.W  " h" Addendum.InfoWindow.H " " admShow " NA ", AddendumGui
			Gui, adm: Default

		; den zuletzt angezeigten TAB wiederherstellen (TCM_SETCURFOCUS (0x1330):)
			For tabNr, TabName in StrSplit(admTabNames, "|")
				If InStr(TabName, Addendum.InfoWindow.firstTab) {
					SendMessage, 0x1330, % tabNr - 1,,, % "ahk_id " admHTab
					break
				}

		; Fensterstyles anpassen
			WinSet, Style  	, 0x50020000, % "ahk_id " hadm
			WinSet, ExStyle , 0x0802009C, % "ahk_id " hadm
			WinSet, AlwaysOnTop, On 	 , % "ahk_id " hadm

	;}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Tabs - Listviews mit Daten befüllen
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
		Addendum.InfoWindow.PatID := AlbisAktuellePatID()
		If (Addendum.InfoWindow.PatID <> Addendum.InfoWindow.lastPatID)
			Addendum.InfoWindow.ReIndex := true, Addendum.InfoWindow.lastPatID := Addendum.InfoWindow.PatID
		else
			Addendum.InfoWindow.ReIndex := false

	; Inhalte auffrischen und anzeigen
		admGui_Journal((Addendum.InfoWindow.Init ? false : true))
		admGui_TProtokoll(Addendum.TProtDate, 0, compname)
		PdfReports := admGui_Reports()

	; Gui aktivieren damit es angezeigt wird
		If admShow
			RedrawWindow(hadm)

		;AlbisActivate(1)

	;}

		If Addendum.import
			admGui_ImportGui(true)

		admReCheck := false

		SetTimer, toptop, % Addendum.InfoWindow.RefreshTime

return

admLAN:                                                                           	;{ zeichnet den LAN Tab

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

adm_Lan:                                                                                         	;{ Netzwerk/LAN Tab Handler

	If (A_GuiControl = "AdmLanCheck") {

			;~ LANMsg := ""
			;~ admSendText("192.168.100.45", "answer|" A_ComputerName "|" A_IPAddress1 "|Status Ok" )
			;~ ;admSendText("192.168.100.25", "answer")
			;~ myTcp := new SocketTCP()
			;~ myTcp.connect("localhost", 1337)
			;~ MsgBox, % myTcp.recvText()

	} else if RegExMatch(A_GuiControl, "AdmRDP_") {

			rdpPath 	:= Addendum.LAN.Clients[compname].rdpPath
			rdpClient	:= StrReplace(A_GuiControl, "AdmRDP_")

			SciTEOutput("rdpfile: " rdpPath "\" rdpclient ".rdp")

			Run, % q rdpPath "\" rdpclient ".rdp" q

	}

return ;}

adm_Notes:			                                                                    			;{ Notizen           	- Edit Handler

return ;}

adm_PdfReports:                                                                               	;{ Patienten Tab 	- Listview Handler

		If InStr(A_GuiControl, "AdmPdfReports") && (A_EventInfo = 0)
			return

	; zurück wenn Datei nicht existiert
		pdfPath := ""
		LV_GetText(admFile, EventInfo := A_EventInfo, 1)
		For idx, pdf in PdfReports
			If InStr(pdf.name, admFile) {
				pdfPath	:= Addendum.BefundOrdner "\" pdf.name
				break
			}

		If (StrLen(pdfPath) = 0)
			return

		If !FileExist(pdfPath) 	{
			PraxTT(	"Die Dateioperation ist nicht möglich,`nda die Datei nicht mehr vorhanden ist.", "3 2")
			admGui_Reports()
			RedrawWindow(hadm)
			return
		}

	;PDF Vorschau
		If Instr(A_GuiEvent, "DoubleClick")
			admGui_View(pdf.name)

return ;}

adm_Journal:                                                                                     	;{ Journal Tab   	- Listview Handler

	; Default Listview, gewählter Dateiname
		admGui_Default("admJournal")
		LV_GetText(admFile, A_EventInfo, 1)
		EventInfo := A_EventInfo
		;ToolTip, % "file: " admFile "`nEvInfo: " EventInfo "`nGuiEvent: " A_GuiEvent "`nControl: " A_GuiControl, 1000, 400, 15

		If       	InStr(A_GuiEvent	, "RightClick")  	{                    	; Kontextmenu
			If (StrLen(admFile) = 0)
				return
			MouseGetPos, mx, my, mWin, mCtrl, 2
			Menu, admJCM, Show, % mx - 20, % my + 10
			return
		}
		else if 	InStr(A_GuiEvent	, "ColClick")    	{                    	; Spaltensortierung
			If (EventInfo > 0)
				admGui_Sort(EventInfo)
			return
		}
		else If 	InStr(A_GuiControl, "admButton2")	{                    	; einzelnes Dokument importieren

			Loop % LV_GetCount() {                                                ; Dokumente ohne Personennamen werden ignoriert
				LV_GetText(rowFile, A_Index, 1)
				If !RxNames(rowFile, "ContainsName")   ; enthält keinen Personennamen dann weiter
					continue
				filePath := Addendum.BefundOrdner "\" rowFile
				If RegExMatch(filePath, "\.pdf$") && FileExist(filePath) && !FileIsLocked(filePath) && isSearchablePDF(filepath) {
					If FuzzyKarteikarte(rowFile)
						admGui_ImportFromJournal(rowFile)
					else
						PraxTT("Es konnte kein passender Patient gefunden werden.", "3 0")
					return
				}
			}

			return
		}
		else If	InStr(A_GuiControl, "admButton3") {                    	; Befundordner neu einlesen
			admGui_Reload()
			return
		}
		else If	InStr(A_GuiControl, "admButton4") {                    	; OCR Vorgang starten oder abbbrechen
			If Addendum.Thread["tessOCR"].ahkReady() {
				MsgBox, 4	, Addendum für Albis on Windows, % "Soll die laufende Texterkennung`nabgebrochen werden?"
				IfMsgBox, No
					return
				Addendum.Thread["tessOCR"].ahkTerminate[]
				admGui_OCRButton("+OCR ausführen")
				return
			}
			else {
				OCRRenameAllFiles()
			}
		}

		If !FileExist(Addendum.BefundOrdner "\" admFile) 	{        	; zurück wenn Datei nicht existiert
			;PraxTT(A_GuiEvent ", " A_GuiControl	" - Eine Dateioperation ist nicht möglich,`nda die Datei`n>" admFile "<`nnicht mehr vorhanden ist.", "3 1")
			;admGui_Reload()
			return
		}
		If Instr(A_GuiEvent	, "DoubleClick")                                 	; PDF/Bild-Programm aufrufen
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

	If Instr(A_GuiEvent, "DoubleClick") {
		LV_GetText(PatID, EventInfo:= A_EventInfo, 4)
		AlbisAkteOeffnen("", PatID)
	}

return ;}

adm_GuiDropFiles:                                                                             	;{ nicht erkannte Dokumente einfach in das Posteingangfenster ziehen!

	PatName 	:= StrReplace(AlbisCurrentPatient(), " ")
	PatName 	:= StrReplace(PatName, "-")
	PatName 	:= StrSplit(PatName, ",")

	; übergebene Dateien liegen als `n getrennte Liste in A_GuiEvent vor (genial einfach))
		Loop, Parse, A_GuiEvent, `n
		{
				SplitPath, A_LoopField, filename

			; Datei muss zunächst sicherheitshalber in den Befundordner kopiert werden, wenn sie dort nicht vorliegt
				If !InStr(A_LoopField, Addendum.BefundOrdner )
					FileCopy, % A_LoopField, % Addendum.BefundOrdner "\" filename

				If RegExMatch(filename, "\.pdf")
				{
						PdfReports.Push(StrReplace(filename, ".pdf"))
						;LV_Add("ICON" 1,, RegExReplace(filename, "^[\w\p{L}-]+[\s,]+[\w\p{L}]+\s*\,*\s*", ""))
						LV_Add("ICON" 1,, filename)
				}
				else If RegExMatch(filename, "\.jpg")
				{
						PdfReports.Push(filename)
						LV_Add("ICON" 2,, RegExReplace(filename, "^[\w\p{L}-]+\s*[,]*\s*[\w\p{L}]+\,*\s*", ""))
				}
		}

	; Importbuttons aktivieren und später nach Befundart ausschalten
		admGui_InfoText("admReports")

return ;}

adm_BefundImport:                                                                              	;{ läuft wenn Befunde oder Bilder importiert werden sollen

		If (PdfReports.MaxIndex() = 0) || Addendum.Importing 	; versehentlichen Aufruf bei leerem Array verhindern
			return

		;Addendum.Importing := true                             		; globale flag für Importvorgang setzen
		PreReaderListe := admGui_GetAllPdfWin()              	; derzeit geöffnete Readerfenster einlesen
		admGui_ImportGui(true, "...importiere alle Befunde")	; Hinweisfenster anzeigen
		Imports := admGui_FileImport()                               	; Importvorgang starten
		admGui_RemoveImports(Imports)                            	; PdfReports - Importe entfernen
		admGui_InfoText("Reports")                                    	; Kurzinfo aktualisieren
		admGui_Journal(true)                                               	; Journalinhalt auffrischen
		admGui_ShowPdfWin(PreReaderListe)                     	; holt den/die PdfReaderfenster in den Vordergrund
		admGui_ImportGui(false)                                        	; Hinweisfenster wieder schliessen
		;Addendum.Importing := false                                 	; globalen flag für Importvorgang ausschalten

		;SetTimer, toptop, % Addendum.InfoWindow.RefreshTime

return ;}

toptop:                                                                                              	;{ zeichnet das Fenster in Abständen neu

	; Gui schliessen unter bestimmten Bedingungen
		Aktiv := Addendum.AktiveAnzeige := AlbisGetActiveWindowType()
		If !RegExMatch(Aktiv, "i)Karteikarte|Laborblatt|Biometriedaten") || !WinExist("ahk_class OptoAppClass") {
			admGui_Destroy()
			return
		}

	; Fenster ab und zu - neu zeichnen lassen
		RedrawWindow(hadm)
		try {
			Gui, adm: Show, NA
		}

	; ist Albis inaktiv, dann wird das wiederholte neuzeichnen der Gui zeitlich reduziert
		SetTimer, toptop, % (Addendum.InfoWindow.aRT := !WinActive("ahk_class OptoAppClass") ? Addendum.InfoWindow.RefreshTime*2 : Addendum.InfoWindow.RefreshTime)
		;ToolTip, % "Toptop: on, timer interval: " Round(Addendum.InfoWindow.aRT/1000, 1) "s", 1000, 1, 11

return ;}

; ------------------------ Gui-Funktionen
admGui_Default(LVName)                                             	{               	; Gui Listview als Default setzen

	Gui, adm: Default
	Gui, adm: ListView, % LVName

return
}

admGui_Destroy()                                                             {                	; schliesst das Infofenster

	global admHJournal

	Gui, adm: Submit, Hide
	GuiControlGet, ProtokollTag, adm:, admTPTag

	Addendum.TProtDate             	:= ProtokollTag
	Addendum.InfoWindow.firstTab	:= admTabs

	If Addendum.InfoWindow.RowColors
		LV_Colors.Detach(admHJournal)

	SetTimer, toptop, % (Addendum.InfoWindow.aRT := Addendum.InfoWindow.RefreshTime*3)
	;SetTimer, toptop, Off
	Gui, adm: 	Destroy
	Gui, adm2: 	Destroy

return
}

admGui_Receive(answer)                                               	{                 	; empfängt Netzwerknachrichten

	SciTEOutput(answer)

}

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

; -------- TAB Inhalte
admGui_Reload()                                                           	{                	; aktualisiert die Inhalte und zeichnet das Fenster neu

	admGui_Journal(true)
	PdfReports := admGui_Reports()
	RedrawWindow(hadm)

}

admGui_Reports()                                                         	{	            	; zeigt Befunde des Patienten aus dem Befundordner

		global PdfReports
		PdfReports := []

		static PatName
		static PdfImport 	:= false
		static ImageImport	:= false

		admGui_Default("admPdfReports")
		LV_Delete()


	; Pdf Befunde des Patienten ermitteln und entfernen bestimmter Zeichen aus dem Patientennamen für die fuzzy Suchfunktion
		RegExMatch(AlbisCurrentPatient(), "(?<Nachname>[\pL-]+)[\,\s]+(?<Vorname>[\pL-]+)", Pat)
		PatNVame	:= RegExReplace(PatNachname PatVorname, "\s-")
		PatVName	:= RegExReplace(PatVorname PatNachname, "\s-")

	; PdfReports erstellen - enthält nur die Dateien zum aktuellen Patienten
		For key, pdf in ScanPool	{				;wenn keine PatID vorhanden ist, dann ist die if-Abfrage immer gültig (alle Dateien werden angezeigt)
			RegExMatch(pdf.name, "^\s*(?<Nachname>[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*)[\,\s]+(?<Vorname>[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*)", doc)
			a := StrDiff(PatNVame, RegExReplace(docNachname docVorname, "\s-"))
			b := StrDiff(PatNVame, RegExReplace(docVorname docNachname, "\s-"))
			If (a < 0.11) || (b< 0.11)
				PdfReports.Push(pdf)
		}

	; Pdf Befunde anzeigen
		For key, pdf in PdfReports {
			If !RegExMatch(pdf.name, "\.pdf$") || !FileExist(Addendum.BefundOrdner "\" pdf.name)
				continue
			If pdf.isSearchable
				LV_Add("ICON3", pdf.pages, RxNames(StrReplace(pdf.name, ".pdf"), "ReplaceNames"))
			else
				LV_Add("ICON1", pdf.pages, RxNames(StrReplace(pdf.name, ".pdf"), "ReplaceNames"))
		}

	; Bilddateien aus dem Befundordner einlesen
		Loop, Files, % Addendum.BefundOrdner "\*.*"
			If RegExMatch(A_LoopFileName, "\.jpg|png|tiff|bmp|wav|mov|avi$") {
				RegExMatch(A_LoopFileName, "^\s*(?<Nachname>[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*)[\,\s]+(?<Vorname>[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*)", Such)
				SuchName := RegExReplace(SuchNachname SuchVorname, "\s-")
				a	:= StrDiff(SuchName, PatNVame)
				b	:= StrDiff(SuchName, PatVName)
				Diff := (a <= b) ? a : b
				If (Diff < 0.11) {
					LV_Add("ICON2",, RxNames(A_LoopFileName, "ReplaceNames")) ; entfernt den Namen
					PdfReports.Push({"name": A_LoopFileName})
				}
			}

	; InfoText auffrischen
		admGui_InfoText("Reports")

return PdfReports
}

admGui_CheckJournal(pdffile, ThreadID)                          	{                	; Ausführung jedesmal nach erfolgreicher Erstellung einer PDF durch TesseractOCR()

		global admHJournal

		admGui_Default("admJournal")

	; Farb-Markierung entfernen
		If (rowNr := LV_FindRow("admJournal",1, pdffile)) {
			LV_Modify(rowNr, "ICON3")                                        	; ICON für durchsuchbare PDF hinzufügen
			admGui_ColorRow(admHJournal, rowNr, false)          	; Hintergrundfarbe der Zeile wird entfernt
		}

	; Datei aus dem Stapel entfernen
		For idx, inprogress in Addendum.InfoWindow.FilesStack
			If (inprogress = pdffile) {
				Addendum.InfoWindow.FilesStack.RemoveAt(idx)
				break
			}

	; ScanPool Array ändern und BefundIndex sichern
		For idx, pdf in ScanPool
			If InStr(pdf.name, pdfFile) {
				pdfPath := Addendum.BefundOrdner "\" pdffile
				FileGetSize	, FSize	, % pdfPath, K
				FileGetTime	, FTime	, % pdfPath, C
				ScanPool[idx].filesize      	:= FSize
				ScanPool[idx].filetime      	:= FTime
				ScanPool[idx].pages       	:= GetPDFPages(pdfPath)
				ScanPool[idx].isSearchable	:= isSearchablePDF(pdfPath)
				break
			}

	; Nachricht an den OCR-Thread: OCR Vorgang fortsetzen
		Send_WM_CopyData("continue", ThreadID)

}

admGui_Journal(refreshDocs=false)                              	{	            	; befüllt das Journal mit Pdf und Bildbefunden aus dem Befundordner

		global admHJournal, tessOCRRunning
		static FirstRun := true

		ImageImport := false
		PdfImport   	:= false

		admGui_Default("admJournal")
		LV_Delete()

	; SCANPOOL OBJECT UND "PdfIndex.json" DATEI AUFFRISCHEN
		If (refreshDocs || FirstRun) {
			newPDF := BefundIndex()
			FirstRun := false
		}

	; PDF BEFUNDE DES PATIENTEN HINZUFÜGEN
		GuiControl, adm: -Redraw, admJournal
		Addendum.InfoWindow.OCREnabled := false
		For key, pdf in ScanPool	{

			If FileExist(filePath := Addendum.BefundOrdner "\" pdf.name) {

				FormatTime, filetime      	, % pdf.filetime, dd.MM.yyyy
				FormatTime, timeStamp 	, % pdf.filetime, yyyyMMdd

				; anderes Symbol für durchsuchbare PDF Dateien
				If pdf.isSearchable
					LV_Add("ICON3", pdf.name, filetime, timeStamp)
				else
					LV_Add("ICON1", pdf.name, filetime, timeStamp)

				If !pdf.isSearchable && !Addendum.InfoWindow.OCREnabled
					Addendum.InfoWindow.OCREnabled := true

			}
		}
		GuiControl, adm: +Redraw, admJournal

	; OCR BUTTON
		If Addendum.InfoWindow.OCREnabled {
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
				LV_Add("ICON2", BefundName, filetime, timeStamp)
			}

	; INFOTEXT AUFFRISCHEN
		admGui_InfoText("Journal")

	; SCROLLT DAS ZULETZT GESPEICHERTE ELEMENT IN SICHT
		;~ IWin :=  Addendum.InfoWindow
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

	global PdfReports

	Gui, adm: Default

	If InStr(TabTitel, "Journal")  	{
		InfoText := (ScanPool.MaxIndex() = 0) ? " keine Dokumente" : (ScanPool.MaxIndex() = 1) ? "1 Dokument" : ScanPool.MaxIndex() " Dokumente"
		admGui_Default("admJournal")
		GuiControl, adm: , AdmJournalTitel, % InfoText
		If (LV_GetCount() > 0)
			GuiControl, adm: Enable, AdmButton2
		else
			GuiControl, adm: Disable, AdmButton2
	}

	If InStr(TabTitel, "Reports")  	{
		InfoText := "Posteingang: (" (!PdfReports.MaxIndex() ? "keine neuen Befunde)" : PdfReports.MaxIndex() = 1 ? "1 Befund)" : PdfReports.MaxIndex() " Befunde)") ", insgesamt: " ScanPool.MaxIndex()
		admGui_Default("admPatLV")
		GuiControl, adm: , AdmPdfReportTitel, % InfoText
		If (LV_GetCount() > 0)
			GuiControl, adm: Enable, AdmButton1
		else
			GuiControl, adm: Disable, AdmButton1
	}

	If InStr(TabTitel, "TProtokoll") 	{
		InfoText := "[" compname "] [" TProtokoll.MaxIndex() " Patienten]"
		GuiControl, adm: , AdmTProtokollTitel, % InfoText
	}

}

; -------- Hilfsfunktionen
admGui_Sort(EventNr, LV_Init=false)                              	{                 	; sortiert die Journalspalten und zeigt ein Symbol für die Sortierreihenfolge an

		; Funktion wird gebraucht für die Wiederherstellung der letzten Sortierungseinstellung und für das Sichern der Einstellungen bei Nutzerinteraktion
		; LV_Init - nutzen um gespeicherte Sortierungseinstellung wiederherzustellen

		global 	admHJournal
		static 	JCol1Dir, JCol2Dir, LVSortStr, JColDir := []

		admGui_Default("admJournal")

		If LV_Init {

			RegExMatch(Addendum.InfoWindow.JournalSort, "\s*(\d)\s*(\d)", JSort)
			EventNr := JSort1

			If (JSort1 = 1) {
				If (JSort2 = 1)
					JCol1Dir := true
			}
			else {
				If (JSort2 = 1)
					JCol2Dir := true
			}

			If JCol1Dir
				JCol2Dir := false

			If JCol2Dir
				JCol1Dir := false

		}

	; Sortierung je nach gewählter Spalte vornehmen
	; Idee von: https://www.autohotkey.com/boards/viewtopic.php?t=68777
		If      	(EventNr = 2) {    ; Spalte 2 - sortiert wird nach der unsichtbaren Spalte 3

			admGui_ColSort("admJournal", EventNr, (JCol2Dir := !JCol2Dir))
			Addendum.InfoWindow.JournalSort := EventNr " " (JCol2Dir ? "1":"0")

		}
		else If	(EventNr = 1) {

			admGui_ColSort("admJournal", EventNr, (JCol1Dir := !JCol1Dir))
			Addendum.InfoWindow.JournalSort := EventNr " " (JCol1Dir ? "1":"0")

		}


}

admGui_ColorRow(hLV, rowNr, paint=true)                      	{                	; eine Zeile einfärben

	global 	hadm
	static 	bcolor := 0x995555, tcolor := 0xFFFFFF

	If !Addendum.InfoWindow.RowColors
		return

	If paint
		res := LV_Colors.Row(hLV, rowNr, bcolor, tcolor)
	else
		res := LV_Colors.Row(hLV, rowNr)

	SciTEOutput("ColorRow: " res)
	;GuiControl, adm: +Redraw, admJournal
	WinSet, Redraw, , % "ahk_id " hadm

}

admGui_ColSort(LVName, col, SortDest)                         	{                	; Helfer für admGui_Sort

	global 	admHJournal

	admGui_Default(LVName)

	If (LVName = "admJournal")
		If (col = 2)
			col := 3, Type := "Integer"
		else
			Type := "Text"

	LV_ModifyCol(col, (SortDest ? "Sort" : "SortDesc") " " Type)
	LV_SortArrow(admHJournal, col, (SortDest ? "up":"down"))

}

admGui_OCRButton(status)                                           	{                	; ändert den Text und Modus des OCR Buttons im Journal Tab

	RegExMatch(status, "i)^\s*(?<Enabled>[\+\-])\s*(?<Text>[\pL-\s]+)\s*$", Button)

	If (ButtonEnabled ="+")
		ButtonEnabled :="Enable"
	else if (ButtonEnabled ="-")
		ButtonEnabled :="Disable"
	else
		ButtonEnabled :="Enable"

	Gui, adm: Default
	GuiControl, % "adm: " ButtonEnabled, AdmButton4
	GuiControl, % "adm: ", AdmButton4, % ButtonText

}

; -------- Kontextmenu
admGui_CM(MenuName)                                               	{                 	; Kontextmenu des Journal

		; fehlt: nachschauen ob im Befundordner ein Backup-Verzeichnis angelegt ist
		; auch für Tastaturkürzelbefehle in allen Tabs (01.07.2020)

		global 	admHJournal, admHPDFReports, admFile, rcEvInfo, hadm, PdfReports
		static 	newadmFile, rowNr

		If RegExMatch(MenuName, "^J")
			admGui_Default("admJournal")

		rowSel  	:= LV_FindRow("admJournal", 1, admFile)
		blockthis 	:= false
		For key, inprogress in Addendum.InfoWindow.FilesStack
			If (inprogress = admFile) {
				blockthis := true
				break
			}

	; Menupunkte ohne PDF Schreibzugriffe auf PDF Datei
		If      	InStr(MenuName, "JRefresh")	{	; Listview aktualisieren
			admGui_Reload()
			return
		}
		else if 	InStr(MenuName, "JOpen") 	{  ; Listview auffrischen
			If blockthis {
				PraxTT("Die wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			FuzzyKarteikarte(admFile)
			return
		}
		else if 	InStr(MenuName, "JView")  	{ 	; PDF mit dem Standard PDF Anzeigeprogramm öffnen
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
		If         	InStr(MenuName, "JDelete") 	{	; Datei löschen

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

		; Datei aus der Listview entfernen
			admGui_Default("admJournal")
			LV_Delete(rowSel)

		; Ordner neu einlesen
			admGui_Journal()
			admGui_Reports()

		}
		else if 	InStr(MenuName, "JOCR")   	{	; OCR einer Datei

				If blockthis {
					PraxTT("Diese Datei ist gerade in Bearbeitung.`n...bitte warten...", "1 0")
					return
				}

			; Nutzerhinweise anzeigen
				InfoText := ""
				If !isSearchablePDF(Addendum.BefundOrdner "\" admFile) {
					SciTEOutput("`nPDF [" admFile "] bisher ohne Texterkennung. Starte Tesseract..")
				}
				else                                                                                                                                                  	{
					InfoText := ">> OCR Text vorhanden <<`n" admFile "`nBei erneuter Ausführung werden die bisherigen Textdaten gelöscht.`nDennoch ausführen?"
					MsgBox, 4	, Addendum für Albis on Windows, % InfoText
					IfMsgBox, No
						return
				}

				PraxTT("Die Texterkennung der Datei < " admFile " > wird jetzt ausgeführt.", "1 0")
				Sleep 1000
				Addendum.InfoWindow.FilesStack.InsertAt(1, admFile)
				admGui_ColorRow(admHJournal, rowSel)                                				; Listview coloring - in der Funktion wird geprüft, ob die coloring aktiviert ist
				admGui_OCRButton("+OCR abbrechen")                                           	; OCR Vorgang anzeigen

			; OCR Starten als echter Thread, nach Abschluß eines Vorgangs schickt der Thread eine Nachricht ans Addendum-Skript
				TesseractOCR(admfile)

			return
		}
		else if 	InStr(MenuName, "JOCRAll") 	{	; OCR aller offenen Dateien

			If Addendum.Thread["tessOCR"].ahkReady() || (Addendum.InfoWindow.FilesStack.Count() > 0) {
				PraxTT("Ein Texterkennungsvorgang läuft im Moment noch.`n...bitte abwarten...", "1 0")
				return
			}
			OCRRenameAllFiles()

		}
		else if 	InStr(MenuName, "JRecog")   	{	; automatische Dateinamenbenennung

				If blockthis {
					PraxTT("Diese Datei ist gerade in Bearbeitung.`n...bitte warten...", "1 0")
					return
				}

				;Addendum.InfoWindow.FilesStack.Push(admFile)

			; nicht ändern wenn Datei einen Patientennamen trägt
				;~ If RxNames(admFile, "ContainsName") {
					;~ InfoText := "Die Datei wurde schon einem Patienten zugeordnet.`n" admFile "`nMöchten Sie automatische Benennung dennoch durchführen?"
					;~ MsgBox, 4	, Addendum für Albis on Windows, % InfoText
					;~ IfMsgBox, No
						;~ return
					;~ withoutName := RxNames(admFile, false)                                           	; alles nach dem Namen speichern
					;~ withoutName := RegExReplace(withoutname, "\s*\,\s*Befund\s*\.pdf")
				;~ }

			; Pfadevariablen
				pdfPath	:= Addendum.Befundordner "\" admFile
				txtPath	:= Addendum.Befundordner "\Text\"  StrReplace(admFile, ".pdf", ".txt")

			; Textlayer der PDF extrahieren und laden
				If FileExist(txtPath)
					PdfText := FileOpen(txtPath, "r").Read()
				else 	{
					if InStr(FileExist("T:\"), "D")
						tmpPath := "T:"
					else
						tmpPath := A_Temp

					PdfToText(pdfPath, 0, "UTF-8", tmpPath "\tmp.txt")
					PdfText := FileOpen(tmpPath "\tmp.txt", "r").Read()
					FileDelete, % tmpPath "\tmp.txt"
				}

				If (StrLen(PdfText) = 0) {
					PraxTT("Es konnte kein Text extrahiert werden.`nDer Vorgang wurde abgebrochen.", "2 0")
					return
				}

			; Ausgabe von Wort und Zeichenzahl der Datei
				wordcount	:= StrSplit(PdfText, " ").MaxIndex()
				charcount	:= StrLen(RegExReplace(PdfText, "\s"))
				SciTEOutput("`n[Textanalyse startet]`n  Dokument: " pdfPath " enthält: " wordcount " Worte mit insgesamt " charcount  " Zeichen")

			; RegEx Auswertungen
				found := FindName(PdfText)
				If IsObject(found)  {

					if (found.Count() = 1) {

						For PatID, Pat in found {
							newadmFile := found[PatID].Nn ", " found[PatID].Vn ", " (StrLen(withoutname) = 0 ? "Befund" : withoutname)	; " [" PatID "]
							break
						}

						SciTEOutput("  neuer Name:`t" newadmFile ".pdf")
						;admGui_Rename(admFile,, newadmFile)
						return

						admGui_Reload()

					}
					else {

						For PatID, Pat in found
							t .= "   [" PatID "] " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`n"
						SciTEOutput("  Das Dokument konnte nicht eindeutig zugeordnet werden. Folgende Namen stehen zu Auswahl:`n" t)

					}

				}
				else {
					SciTEOutput("  Im Dokument konnten keine Patientennamen erkannt werden.")
				}

			return
		}
		else if 	InStr(MenuName, "JRename") {	; manuelles Umbenennen

			If blockthis {
				PraxTT("Die Datei wird gerade bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_Rename(admFile)

		}
		else if 	InStr(MenuName, "JRenAll")   	{	; automatische Umbenennung aller unbenannten Dateien

			If Addendum.Thread["tessOCR"].ahkReady() || (Addendum.InfoWindow.FilesStack.Count() > 0) {
					PraxTT("Ein Texterkennungsvorgang läuft im Moment noch.`n...bitte abwarten...", "1 0")
					return
			}

			notNamed := Array()
			For idx, pdf in ScanPool
				If !RxNames(pdf.name, "ContainsName")
					notNamed.Push(pdf.name)

			If (notNamed.Count() > 0)
				Autonaming(notNamed)

		}
		else if 	InStr(MenuName, "JImport")	{	; Datei in geöffnete Karteikarte importieren
			If blockthis {
				PraxTT("Die wird bearbeitet.`n...bitte warten...", "1 0")
				Return
			}
			admGui_ImportFromJournal(admFile)
		}

return
}

admGui_OCR(pdfDoc, Mode="Text")                             	{                	; kompletten Text einer PDF-Datei per tesseract cmdline tool extrahieren

	If      	(Mode = "Text1") 	{

		Loop % (PageMax := PdfInfo(pdfDoc, "Pages")) {

			PraxTT(	"OCR Vorgang für Dokument:"                           	"`n"
					. 	pdfDoc                                                             	"`n"
					.	"Seite " A_Index " von " PageMax " wird bearbeitet."	"`n"
					.	"konvertiere in png Datei.", "0 2")

			pngPath := PdfToPng(pdfDoc, A_Index)

			PraxTT(	"OCR Texterkennungsvorgang für Dokument:"     	"`n"
					.	pdfDoc                                                                	"`n"
					.	"Seite " A_Index " von " PageMax " wird bearbeitet."	"`n"
					.	"Tesseract OCR-Vorgang läuft.", "0 2")

			;PdfText .= OCR(pngPath, "deu") "`n"
			;SciTEOutput(A_Index ": Textlänge: " StrLen(PdfText))

		}
		return PdfText

	}
	else If 	(Mode = "OCR") 	{

	}


}

admGui_Rename(filename, prompt="", newfilename="")   	{                	; Dialog für Dokument umbenennen

	; Variablen
		global hadm
		static admRen_Width 	:= 400
		static admRen_Height	:= 145

		if (StrLen(prompt) = 0)
			prompt := "Ändern Sie hier die Dateibezeichnung.`nDrücken Sie anschließend 'OK' oder 'ENTER'"

		admPos	:= GetWindowSpot(hadm)
		RnWidth	:= admPos.W + (2 * admPos.BW) + 10
		RnWidth	:= RnWidth < 400 ? admRen_Width : RnWidth

	; PDF Vorschau anzeigen
		hPV := admGui_PDFPreview(fileName)

	; Inputbox anzeigen
		SplitPath, filename,,, fileext, fileout
		newfilename := Trim(newfilename)
		If (StrLen(newfilename) > 0)
			SplitPath, newfilename,,,, fileout

		InputBox, newadmFile	, % "Dokument umbenennen - Addendum für AlbisOnWindows"
									    	, % prompt
											,   ; Hide
											, % RnWidth
											, % admRen_Height
											, % admPos.X + admPos.W -	 RnWidth + 5
											, % admPos.Y + admPos.H
											, % "Locale"
											,  ; Timeout
											, % fileout

	; PDF Vorschau schliessen
		Gui, admPV: Destroy

	; Nutzereingabe auswerten
		newadmFile .= "." fileext
		If ErrorLevel || (newadmFile = filename) || (StrLen(newadmFile) = 0) {

			If (newadmFile = filename) && !ErrorLevel
				PraxTT("Sie haben den Dateinamen nicht geändert!", "2 1")
			else if (StrLen(newadmFile) = 0)
				PraxTT("Ihr neuer Dateiname enthält keine Zeichen.", "2 1")

			RedrawWindow(hadm)
			return

		}

	; Datei umbenennen
		FileMove, % Addendum.BefundOrdner "\" filename, % Addendum.BefundOrdner "\" newadmFile, 1
		If ErrorLevel {
			MsgBox, 0x1024, Addendum für Albis on Windows, % 	"Das Umbenennen der Datei`n"
																							.	"  <" filename ">  `n"
																							. 	"wurde von Windows aufgrund des`n"
																							.	"Fehlercode (" A_LastError ") abgelehnt."
			RedrawWindow(hadm)
			return
		}

	; zugehörige Text Datei umbenennen
		If (fileext = "pdf") && FileExist(Addendum.BefundOrdner "\Text\" fileout ".txt")
			FileMove, % Addendum.BefundOrdner "\Text\" fileout ".txt", % Addendum.BefundOrdner "\Text\" StrReplace(newadmFile, ".pdf", ".txt"), 1

	; Backup Datei umbenennen
		If (fileext = "pdf") && FileExist(Addendum.BefundOrdner "\Backup\" fileName)
			FileMove, % Addendum.BefundOrdner "\Backup\" fileName, % Addendum.BefundOrdner "\Backup\" newadmFile, 1

	; Anzeige auffrischen
		admGui_Reload()

}

admGui_PDFPreview(fileName)                                         	{                	; zeigt eine Vorschau der PDF Datei an

		global 	hadm
		static 	hadmPV, PVPic1

		PdftoPngexe := Addendum.xpdfPath "\PdftoPng.exe"
		If !FileExist(PdfToPngexe) {
			SciTEOutput(PdftoPngexe " ist nicht vorhanden.")
			return
		}

		If (filetest := FileOpen("T:\test.txt", "w")) {
			filetest.Close()
			tempPath := "T:"
		} else
			tempPath := A_Temp

		pngExtract_cmdline := PdftoPngexe " -f 1 -l 1 -r 100 -aa yes -freetype yes -aaVector yes " q Addendum.BefundOrdner "\" fileName q " " q tempPath "\PdfPreview" q
		RunWait, % pngExtract_cmdline,, Hide

		If !FileExist(tempPath "\PdfPreview-000001.png") {
			PraxTT("PDF Vorschau konnte nicht erstellt werden.", "2 0")
			return
		}

		scaleX 	:= 700/900
		picH 	:= A_ScreenHeight - 50
		picW 	:= Round(picH * scaleX)

		Gui, admPV: New 	, +ToolWindow -DPIScale +HWndhadmPV	+Border
		Gui, admPV: Margin	, 5, 5
		;Gui, admPV: Add   	, Picture, % "xm ym vPVPic1 0xE BackGroundTrans", % tempPath "\PdfPreview-000001.png"
		pic := tempPath "\PdfPreview-000001.png"
		Gui, admPV: Add  	, ActiveX, % "w" picW " h" picH, % "mshtml:<img src='" pic "' height='100%' />"
		;Gui, admPV: Add  	, ActiveX, % "w" picW " h" picH, % "mshtml:<iframe src=" q Addendum.BefundOrdner "\" fileName q " height=" q "100%" q "></iframe>"
		Gui, admPV: Show	, % "AutoSize NA Hide", Pdf Vorschau
		admPos	:= GetWindowSpot(hadm)
		PV    		:= GetWindowSpot(hadmPV)
		Gui, admPV: Show, % "x" (admPos.X - PV.W - 5) " y0 NA", Pdf Vorschau
		;Gui, admPV: Show, % "x50 y50 NA", ScanPool - Pdf Vorschau


return hadmPV
}

admGui_View(filename)                                                 	{              	; Befund-/Bildanzeigeprogramm aufrufen

	;SciTEOutput(Addendum.BefundOrdner "\" filename)
	If (StrLen(filename) = 0) || !FileExist(Addendum.BefundOrdner "\" filename)
		return 0

	If RegExMatch(filename, "\.jpg|png|tiff|bmp|wav|mov|avi$") {
		PraxTT("Anzeigeprogramm wird geöffnet", "3 0")
		Run % q Addendum.BefundOrdner "\" filename q
	}
	else {
		PraxTT("PDF-Anzeigeprogramm wird geöffnet", "3 0")
		If !isCorruptPDF(Addendum.BefundOrdner "\" filename)
			Run % Addendum.PDFReaderFullPath " " q Addendum.BefundOrdner "\" filename q
		else
			PraxTT("PDF Datei:`n>" fileName "<`nist defekt", "2 0")
	}


	PraxTT("", "off")

return 1
}


admGui_ImportGui(Show=true, IGmsg=""                                           	; Hinweisfenster für laufenden Importvorgang
	, gctrl="AdmPdfReports")                                            	{

	global adm2, hadm, hadm2, AdmImporter, AdmJournal

	If Show {

	; zeigt den Hinweis auch an wenn das Infofenster neu gezeichnet wurde
		If Addendum.Importing && (StrLen(last_IGmsg) > 0)
			IGmsg := last_IGmsg, gctrl := last_ctrl
		else if !Addendum.Importing
			last_IGmsg := IGmsg, last_ctrl := gctrl, Addendum.Importing := true

		If StrLen(IGmsg) = 0
			IGmsg := "Importvorgang läuft...`nbitte nicht unterbrechen!"

		GuiControlGet, rpos, adm: Pos, % gCtrl

		Gui, adm2: New	, +HWNDhadm2 +ToolWindow -Caption +Parent%hadm%
		Gui, adm2: Color	, cAA0000
		Gui, adm2: Font	, s10 q5 cFFFFFF Bold, % Addendum.StandardFont
		Gui, adm2: Add	, Text, % "x" 0 " y" Floor(rposH//4) " w" rposW " Center vAdmImporter"	, % IGmsg
		Gui, adm2: Show	, % "x" rposX " y" rposY " w" rposW " h" rposH " Hide NA"                     	, AdmImportLayer

		GuiControlGet, cpos, adm2: Pos, AdmImporter
		GuiControl, adm2: Move, AdmImporter, % "y" Floor(rposH // 2 - cposH // 2)
		Gui, adm2: Show	, % "NA"

		WinSet, Style         	, 0x50020000	, % "ahk_id " hadm2
		WinSet, ExStyle	    	, 0x0802009C	, % "ahk_id " hadm2
		WinSet, Transparent	, 220             	, % "ahk_id " hadm2

	} else {

		Gui, adm2: Destroy
		last_IGmsg := last_ctrl := ""
		Addendum.Importing := false

	}

}

admGui_ImportFromJournal(report)                               	{                 	; Einzelimport-Funktion vom Journal

		global 	admHPDFReports, hadm
		static 	WZEintrag	:= {"LASV":"Anfragen", "Anfrage":"Anfragen", "Antrag":"Anfragen", "Lageso":"Anfragen", "Lebensversicherung":"Anfragen"}
		static 	rxPerson 	:= "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"

	; Nutzer abfragen
		If Addendum.InfoWindow.ConfirmImport {
			message :=	"Wollen Sie die Datei:`n"
							. 	(StrLen(Report) >30 ? SubStr(Report, 1, 30) "..." : Report) "  dem Pat.`n"
							. 	AlbisCurrentPatient() " zuordnen ?"
			MsgBox, 262148, Importieren ? - Addendum für AlbisOnWindows, % message
			IfMsgBox, No
				return
		}

	; Befund importieren
		If      	RegExMatch(report, "\.pdf")                                             	{

			Addendum.FuncCallback := "GetPDFViewerArg"
			If AlbisImportierePdf(report) {

				If FileExist(Addendum.BefundOrdner "\Backup\" report)        ; PDF in einen anderen Ordner verschieben
					FileMove, % Addendum.BefundOrdner "\Backup\" report, % Addendum.BefundOrdner "\Backup\Importiert\" report, 1
				For idx, pdf in ScanPool
					If InStr(pdf.name, report)
						ScanPool.RemoveAt(idx)

			}

		}
		else If	RegExMatch(report, "\.jpg|png|tiff|bmp|wav|mov|avi$")	{
			If !AlbisImportiereBild(report, RegExReplace(report, "^\s*[\p{L}-]+[\,\s]+[\p{L}-]+"))
				return
		}
		else
			return

	; Re-Indizieren, Gui auffrischen, Importprotokoll
		FileAppend, % A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min " - " report "`n", % Addendum.Befundordner "\PdfImportLog.txt"
		while, % (StrLen(Addendum.FuncCallback) > 0) {
			Sleep, 50
			If (A_Index > 20)  ; 1 Sekunde
				break
		}

		admGui_Journal(true)
		admGui_Reports()
		admGui_ShowTab("Journal")

return
}

admGui_ShowTab(TabName, hTab="")                          	{                	; ein bestimmten Tab nach vorne holen

		global admHTab
		static admGuiTabs := {"Patient":0, "Journal":1, "Protokoll":2, "Info":3, "Netzwerk":4}

	; (TCM_SETCURFOCUS (0x1330)
		SendMessage, 0x1330, % (RegExMatch(TabName, "^\d+$") ? TabName : admGuiTabs[TabName]),,, % "ahk_id " (StrLen(hTab) = 0 ? admHTab : hTab)

		;~ SciTEOutput("EL: " ErrorLevel ", " StrLen(hTab))
		;~ SciTEOutput("SendMessage, 0x1330, " (RegExMatch(TabName, "^\d+$") ? TabName : admGuiTabs[TabName]) ",,, % " q "ahk_id " (StrLen(hTab) = 0 ? admHTab : hTab) q )

return ErrorLevel
}

admGui_FileImport()                                                     	{                	; Befundimport für alle Befunde eines Patienten

	global PdfReports

	Imports := ""
	admGui_Default("admPdfReports")                         	; Journallistview als Default setzen

	For key, pdf in PdfReports		{

		imported	:= ""
		report 	 	:= pdf.name

		If RegExMatch(report, "\.(jpg|png|tiff|bmp|wav|mov|avi)$")	{	; Image importieren

			If !AlbisImportiereBild(report, RegExReplace(report, "^\s*[\p{L}-]+[\,\s]+[\p{L}-]+")) {
				;SciTEOutput("Datei (" Report ") wurde nicht importiert.", 0, 1)
				continue
			} else {
				If FileExist(Addendum.BefundOrdner "\Backup\" report)        ; PDF in einen anderen Ordner verschieben
					FileMove, % Addendum.BefundOrdner "\Backup\" report, % Addendum.BefundOrdner "\Backup\Importiert\" report, 1
				FileAppend, % A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min " - " report "`n", % Addendum.Befundordner "\PdfImportLog.txt"
				imported := Report
			}

		}
		else {                                                                                    ; Pdf importieren

			Addendum.FuncCallback := "GetPDFViewerArg"
			If !AlbisImportierePdf(report) {
				;SciTEOutput("Datei (" Report ") wurde nicht importiert.", 0, 1)
				continue
			}	else {
				FileAppend, % A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min " - " report ".pdf`n", % Addendum.Befundordner "\PdfImportLog.txt"
				while, % (StrLen(Addendum.FuncCallback) > 0) {
					Sleep, 50
					If (A_Index > 40)
						break
				}
				imported := Report
				For key, pdf in ScanPool
					If InStr(pdf.name, report)
						ScanPool.RemoveAt(key)
			}

		}

		If (StrLen(imported) > 0)                   			{                        	; importierten Eintrag aus der Listview entfernen

				Imports .= imported "`n"
				admGui_Default("admPdfReports")
				Loop % LV_GetCount() {
					LV_GetText(celltext, A_Index, 2)
					If InStr(imported, celltext) {
						LV_Delete(A_Index)
						break
					}
				}

		}

	}

	PdfReports := Array()

return Imports
}

admGui_RemoveImports(ImportList)                               	{                 	; PdfReport-Array aufräumen

	global PdfReports

	Loop, Parse, % RTrim(ImportList, "`n"), `n
		Loop % PdfReports.MaxIndex()
			If InStr(PdfReports[A_Index], A_LoopField) {
				PdfReports.RemoveAt(A_Index)
				continue
			}

}

admGui_GetAllPdfWin()                                                 	{                	; Namen aller angezeigten PdfReader ermitteln

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDFReaderWinClass
	Loop % tmpListe
		PreReaderListe .= tmpListe%A_Index% "`n"

return PreReaderListe
}

admGui_ShowPdfWin(PreReaderListe)                             	{               	; PdfReaderfenster nach vorne holen

	; funktioniert nicht wenn man nur ein Fenster (eine Instanz) des PdfReader eingestellt hat
	; die Tabs im FoxitReader bekomme ich nicht ermittelt und somit auch nicht angesteuert

	WinGet, tmpListe, List, % "ahk_class " Addendum.PDFReaderWinClass
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

admGui_SaveStatus()                                                        	{                	; Anzeigestatus des Infofenster sichern

	global
	Gui, adm: Submit, Hide
	GuiControlGet, ProtokollTag, adm:, admTPTag
	;SciTEOutput("Tab: " admTabs ", Tag: " ProtokollTag)
	IniWrite, % ProtokollTag                                 	, % Addendum.AddendumIni, % compname, % "Infofenster_Tagesprotokoll_Datum"

	If (InStr(admTabs, "Protokoll")) {        ; speichert die Scrollviewposition
		admGui_Default("admProtokoll")
		admTabs .= " " LV_GetScrollViewPos(admHJournal)
	}
	IniWrite, % admTabs                                      	, % Addendum.AddendumIni, % compname, % "Infofenster_aktuelles_Tab"
	IniWrite, % Addendum.InfoWindow.JournalSort	, % Addendum.AddendumIni, % compname, % "Infofenster_JournalSortierung"

}

admGui_Karteikarte(PatID)                                             	{              	; fragt ob eine Karteikarte geöffnet werden soll

	If !Addendum.PatAkteSofortOeffnen {
		MsgBox, 4, Patientenakte öffnen?, % "Möchte Sie die Akte des Patienten:`n(" PatID ") " oPat[PatID].Nn ", " oPat[PatID].Vn " geb.am " oPat[PatID].Gd "`nöffnen?"
		IfMsgBox, No
			return
	}

	AlbisAkteOeffnen(oPat[PatID].Nn ", " oPat[PatID].Vn, PatID)

return
}

; ------------------------ Texterkennung
TesseractOCR(files)                                                         	{                	; erstellt einen zusätzlichen echten Thread

	global Settings, autoexecScript, tessOCRRunning
	;SciTEOutput("OCRThread: " Addendum.Threads.OCR)

	Settings := {	"SetName"                             	: "threadOCR"
					,	"UseRamDisk"                        	: "T:"
					,	"tesspath"                               	: Addendum.AddendumDir "\include\OCR"
					,	"xpdfpath"            	                 	: Addendum.AddendumDir "\include\xpdf"
					,	"documentPath"                        	: Addendum.BefundOrdner
					,	"backupPath"                	         	: Addendum.BefundOrdner "\backup"
					,	"txtOCRPath"        	                 	: Addendum.BefundOrdner "\Text"
					,	"OCRLogPath"      	                  	: Addendum.DBPath "\OCRTime_Log.txt"
					,	"tessconfig"                            	: ""
					,	"cmdlinePlus"                            	: "--psm 6 --oem 3"
					,	"useData"                               	: "best"
					,	"uselang"                               	: "deu"
					,	"useleptonica"                           	: "false"
					,	"convertwith"                           	: "convert"
					,	"ProcessID"                             	: DllCall("GetCurrentProcessId")
					,	"ScriptID"                                	: (Addendum.MsgGui.hMsgGui)
					,	"showprogress"                        	: "true"
					,	"writelog"             	                 	: "true"
					,	"backupfiles"                            	: "true"
					,	"corrupt_files_try_again"           	: 3
					,	"stopOCR_num_fails"              	: 2
					,	"callbackFunc"                          	: "OCRReady"}

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
		pdfText := tessOCRPdf(files, settings)
		Send_WM_CopyData("OCR_processed|" files "|" hOCR, settings.ScriptID)
		While threadcontrol
			sleep, 100
	}

	Send_WM_CopyData("OCR_ready|" filesmax "|" hOCR, settings.ScriptID)
	;SciteOutPut("send message to: " settings.ScriptID " - processing is done. ")

	ExitApp
	)

	; schreibe das/die Objekt/e als Skriptcode
		If IsObject(files) {
			repfiles := "files := ["
			For key, val in files
				repfiles .= q val q ","
			repfiles := RTrim(Trim(repfiles), ",") "]`n"
		}
		else
			repfiles := "files := " q files q

		repset := "settings := {"
		For key, val in Settings {
			repset .= q key q ":" . ((val = "true") || (val = "false") || RegExMatch(val, "^(\d+|0x\w+)$") ? val : q val q) ", "
		}
		repset := RTrim(Trim(repset), ",") "}`n"

	; Skript zusammensetzen
		autoexec	:= StrReplace(autoexecScript	, "###files"   	, repfiles)
		autoexec	:= StrReplace(autoexec       	, "###settings"	, repset)
		ocrfuncs	:= Addendum.Threads.OCR
		script := autoexec "`n" ocrfuncs
		FileOpen("T:\OCRScript.ahk", "w", "utf-8").write(script)

	; OCR-Thread starten
		If !Addendum.Thread["tessOCR"].ahkReady()
			Addendum.Thread["tessOCR"] := AHKThread(script)


}

GetNoOCRFiles()                                                           	{              	; stellt PDF Dateien ohne Text zusammen für TesseractOCR

	global admHJournal

	noOCR := Array()

	For idx, pdf in ScanPool
		If !pdf.isSearchable {

		; Datei ist in Bearbeitung (z.B. OCR Vorgang), dann ignorieren
			InStack := false
			For stackindex, stackfile in Addendum.InfoWindow.FilesStack
				If (stackfile = pdf.name) {
					InStack := true
					break
				}
			If Instack
				continue

		; Datei hinzufügen
			noOCR.Push(pdf.name)
			admGui_ColorRow(admHJournal, LV_FindRow("admJournal", 1, pdf.name), true) 			; kann per Flag ausgeschaltet werden

		}


return noOCR
}

OCRRenameAllFiles()                                                     	{                	; Texterkennung und automatische Bezeichnung aller unbearbeiteten PDF-Dateien

	static noOCR

	noOCR := GetNoOCRFiles()
	If (noOCR.MaxIndex() > 0) {

		For key, filename in noOCR
			Addendum.InfoWindow.FilesStack.Push(admFile)

		admGui_OCRButton("+OCR abbrechen")
		TesseractOCR(noOCR)
		;Autonaming(noOCR)
		admGui_OCRButton("-OCR ausführen")

	}
	else
		admGui_OCRButton("-OCR ausführen")

	SciTEOutput("noOCR: " noOCR.MaxIndex())

return
}

Autonaming(notNamed)                                                	{                	; automatisiert die Benennung aller unbenannten PDF-Dateien

	; pdftotext -f 1 -l 1 -layout -enc UTF-8 -bom test4.pdf test4.txt

	; Variablen
		docPath := Addendum.BefundOrdner

	; verarbeitet Dateilisten in Form eines linearen Array
	; es kann auch einfach nur ein Dateiname als String übergeben werden
		If !IsObject(notNamed) {
			noName := Array()
			noName.Push(notNamed)
		} else
			noName := notNamed

	; for-loop through files
		For idx, filename in noName
			If !RxNames(filename, "ContainsName")  {            	; PDF sollen in der Form Nachname[-Nachname2], Vorname[-Vorname2], Beschreibungstext.......pdf vorliegen

				newfilename	:= ""
				pdfPath     	:= docPath "\" filename
				bupPath     	:= docPath "\Backup\" filename
				txtPath      	:= docPath "\Text\" StrReplace(filename, ".pdf", ".txt")

				If !isSearchablePDF(pdfPath) ;|| !FileExist(txtPath) || !FileExist (pdfPath)
					continue

				PdfText := FileOpen(txtpath, "r").Read()
				SciTEOutput(fileName ": " StrLen(PdfText))
				found := FindName(PdfText)
				If IsObject(found) {

					If (found.Count() = 1)
						For PatID, Pat in found {
							newfilename := found[PatID].Nn ", " found[PatID].Vn ", Befund.pdf"
							SciTEOutput("  neuer Name:   `t" newfilename)
							break
						}
					else If (found.Count() > 1) {
						newfilename := "--enthält " found.Count() " Patientennamen-- " filename
						SciTEOutput("  Dokument:   `t" newfilename)
					}

					If (StrLen(newfilename) > 0) {
							FileMove, % pdfPath	, % docPath "\" newfilename, 1                                                	; Originaldatei umbenennen
							FileMove, % bupPath	, % docPath "\Backup\" newfilename, 1 		                               	; Backupdatei umbenennen
							FileMove, % txtPath	, % docPath "\Text\" StrReplace(newfilename, ".pdf", ".txt"), 1 		; zugehörige Text Datei umbenennen
					}
				}
			}


	; Anzeige auffrischen
		admGui_Journal(true)
		admGui_Reports()
		RedrawWindow(hadm)

return
}

; ------------------------ PDF Funktionen, Umbennen
isSearchablePDF(pdfFilePath)                                             {               	;-- durchsuchbare PDF Datei?

	If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252")) {
		SciTEOutput("fileopen error: " pdfFilePath)
		return 0
	}

	while !fileobject.AtEof {
		line := fileobject.ReadLine()
		If RegExMatch(line, "i)\/PDF\s*\/Text") {
			filepos := fileObject.pos
			fileObject.Close()
			return filepos
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file) {    	; binären Inhalt überspringen
			fileobject.seek(fileseek, 1)                                           	; •1 (SEEK_CUR): Current position of the file pointer.
		}
	}

	fileObject.Close()

return 0
}

GetPDFPages(pdfFilePath)                                              	{               	;-- gibt die Anzahl der Seiten einer PDF zurück

	If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fileobject.AtEof {
		line := fileobject.ReadLine()
		If RegExMatch(line, "i)\/Count\s+(\d+)", pages) {
			fileObject.Close()
			return pages1
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file) {    	; binären Inhalt überspringen
			fileobject.seek(fileseek, 1)                                           	; •1 (SEEK_CUR): Current position of the file pointer.
		}
	}

	fileObject.Close()

return 0
}

isCorruptPDF(pdfFilePath)                                                 	{                	;-- prüft ob die PDF Datei defekt ist

	If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	VarSetCapacity(EndOfFile, 5)
	fileobject.seek(fileobject.Length - 6, 1)
	fileobject.RawRead(EndOfFile, 5)

return InStr(StrGet(&EndOfFile, 5, "CP0"), "EOF") ? false : true
}

FileIsLocked(FullFilePath)                                                 	{              	;-- ist die Datei gesperrt?

	f := FileOpen(FullFilePath, "rw")
	LE		:= A_LastError
	iO 	:= IsObject(f)
	If iO
		f.Close()

return LE = 32 ? true : false
}

GetPDFViewerArg(class, hwnd)                                        	{               	;-- Prozess commandline des PDF Anzeigeprogramms auslesen

	WinGet PID, PID, % "ahk_id " hwnd

    StrQuery := "SELECT * FROM Win32_Process WHERE ProcessId=" . PID
    Enum := ComObjGet("winmgmts:").ExecQuery(StrQuery)._NewEnum
    If (Enum[Process]) {
		Viewercmdline := Process.CommandLine
		SciTEOutput("Foxit: " Viewercmdline)
	}
	Addendum.FuncCallback := ""

	SciTEOutput("Foxit: " Viewercmdline)
	; FileAppend, % A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min " - " report ".pdf`n", % Addendum.Befundordner "\PdfImportLog.txt"
}

zlib_Decompress(Byref Decompressed
, Byref CompressedData, DataLen
, OriginalSize = -1)                                                          	{              	;-- zum Dekodieren von stream-Objekten in PDF Dokumenten

	; http://www.autohotkey.com/forum/viewtopic.php?t=68170

	OriginalSize := (OriginalSize > 0) ? OriginalSize : DataLen*10 ;should be large enough for most cases
	VarSetCapacity(Decompressed,OriginalSize)
	ErrorLevel := DllCall("zlib1\uncompress", "Ptr", &Decompressed, "UIntP", OriginalSize, "Ptr", &CompressedData, "UInt", DataLen)

return ErrorLevel ? 0 : OriginalSize
}

; ------------------------ Dateien kategorisieren/benennen
FindName(text)                                                             	{              	;-- PDF renaming - Hilfsfunktion

		;				        			,	"(?<name2>[A-Z][\pL-]+)[.,;\s]+(?<name1>[A-Z][\pL-]+)[.,;\s]+[(gebo).(ren)(am)\*\s]*[,:.;\s]*\d{1,2}[.,;\s]+\w+[.,;:\s]+\d{2,4}"
		static ThousandYear	:= SubStr(A_YYYY, 1, 2)
		static HundredYear	:= SubStr(A_YYYY, 3, 2)

		static MonthNrs
		static excludeIDs    	:= "2"
		static rxMonths      	:= "Janu*a*r*|Febr*u*a*r*|Mä*a*rz|Apri*l*|Mai|Juni*|Juli*|Augu*s*t*|Septe*m*b*e*r*|Okto*b*e*r*|Nove*m*b*e*r*|Deze*m*b*e*r*"
		static	rxDates         	:= [	"(?<Tag>\d{1,2})[.,;](?<Monat>\d{1,2})[.,;](?<Jahr>(\d{4}|\d{2}))[\s:,;.]"
						                	,	"(?<Tag>\d{1,2})[.,;\s]+(?<Monat>\d{1,2}|" rxMonths ")[.,;\s]+(?<Jahr>(\d{4}|\d{2}))"]
		static rxPerson       	:= "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"
		static RxNames       	:= [	"[(Frau)(Herr)\s]*(?<name2>" rxPerson ")[.,;]+\s*(?<name1>" rxPerson ")[.,;\s]+[(gebo).(ren)(am)\*\s]*[,:.;\s]*\d{1,2}[.,;\s]+\w+[.,\s]+\d{2,4}"
								        	,	"[Nn]ame[\s\w]*[;:\s]+(?<name2>" rxPerson "[.,\s]+(?<name1>" rxPerson ")"
								        	,	"[(Pat.)(Patient)(Patienti|en)]+[;,:.\s]+(?<name2>" rxPerson ")[.,\s]+(?<name1>" rxPerson ")"
								        	,	"i)(Nach)*name[:,;.\s]+(?<name2>[A-Z][\pL-]+).*Vo(rn|m)ar*me[:,;.\s]+(?<name1>[A-Z][\pL-]+)"
						        			,	"i)Vo(rn|m)ar*me[:,;.\s]+(?<name1>[A-Z][\pL-]+).*(Nach)*name[:,;.\s]+(?<name2>[A-Z][\pL-]+)"]

		If !IsObject(MontNrs)
			MonthNrs := StrSplit(rxMonths, "|")

		PatBuf	:= Object()
		sText := StrSplit(text, "`n", "`r")

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 1:	RegEx Strings nutzen
	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; - benutzt RegEx-Strings aus RxNames um Patientennamen zu finden
	; -
	;
		For RxIdx, rxString in RxNames {

			spos	:= spose := 1
			Next	:= false

			;while (spos := RegExMatch(text, rxString, Pat, spos)) {
			Loop, % sText.MaxIndex() {

				If !RegExMatch(sText[A_Index], rxString, Pat)
					continue

			  ; trennt den String in Nachname (PatName2) und Vorname (PatName1)
				If (StrLen(PatName1) > 0) && (StrLen(PatName2) > 0) {
					spose  += StrLen(PatName2 PatName1)
				} else if (StrLen(PatName) > 0) {
					spose  += StrLen(PatName)
					PatName1 := StrSplit(PatName, "\s+").1
					PatName2 := StrSplit(PatName, "\s+").2
				}

				spos  += StrLen(Pat)
				;SciTEOutput("  <" PatName2 ", " PatName1 ">")
			  ; schaut ob der gefundene Name in der Datenbank ist
				PatIDArr := admDB.StringSimilarityID(PatName1, PatName2, 0.12)
				If IsObject(PatIDArr)
					For idx, PatID in PatIDArr {

						If PatID in %excludeIDs%
							continue

						;Next := true
						If !PatBuf.haskey(PatID) {
							PatBuf[PatID] := {"Nn":oPat[PatID].Nn, "Vn":oPat[PatID].Vn, "Gd":oPat[PatID].Gd, "hits":1, "method":1}
							SciTEOutput("  Methode 1[" RxIdx "]:`t" PatID " , hits: " PatBuf[PatID].hits ", Name: " PatBuf[PatID].Nn ", "PatBuf[PatID].Vn ", " PatBuf[PatID].Gd)
						}

						;SciTEOutput("Methode 1: " PatID " , hits: " PatBuf[PatID] ", Name: " PatName1 ", " PatName2)

					}

			}

			If Next
				break
		}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 2:	Suche nach Patientennamen über gefundene Geburtstage
	; ---------------------------------------------------------------------------------------------------------------------------------------------------
		spos := 1
		while (spos := RegExMatch(text, rxDates[2], D, spos)) {

			spos += StrLen(D)

		  ; Monat wurde als Wort geschrieben, dann die Nummer des Monats ermitteln
			If !RegExMatch(DMonat, "\d{1,2}")
				For nr, rxMonth in MonthNrs
					If RegExMatch(DMonat, rxMonth) {
						DMonat := nr
						break
					}

		  ; für Datumsformate 01.08.24, erstellt bei größeren Jahren als das aktuelle Jahr (zB. 20), 01.08.1924
			If (StrLen(DJahr) = 2)
				DJahr := (DJahr > HundredYear) ? (SubStr("00" (ThousendYear - 1), -1) . DJahr) : (SubStr("00" (ThousendYear), -1) . DJahr)

		  ; sucht nach Patienten mit passenden Geburtstagen
			Datum   	:= SubStr("00" DTag, -1) "." SubStr("00" DMonat, -1) "." DJahr
			PatIDArr	:= admDB.MatchID("gd", Datum)
			If IsObject(PatIDArr)
				For idx, PatID in PatIDArr {

					If PatID in %excludeIDs%
						continue

				 ; bekommt zwei Treffer-Punkte wenn zuvor
					If PatBuf.haskey(PatID) {

						If (PatBuf[PatID].method = 1)
							PatBuf[PatID].hits += 2
						else
							PatBuf[PatID].hits += 1

					} else {

						PatBuf[PatID] := {"Nn":oPat[PatID].Nn, "Vn":oPat[PatID].Vn, "Gd":oPat[PatID].Gd, "hits":1, "method":2}
						If InStr(text, oPat[PatID].Nn) && InStr(text, oPat[PatID].Vn)
							PatBuf[PatID].hits += 2

					}

					SciTEOutput("  Methode 2:    `t" PatID " , hits: " PatBuf[PatID].hits ", Datum: " RegExReplace(D, "[\n\r]"))

				}

		}

	; ---------------------------------------------------------------------------------------------------------------------------------------------------
	; Methode 3:	Findet zwei aufeinander folgende Worte mit Großbuchstaben am Anfang (Personennamen, Eigennamen ...)
	;               	und vergleicht diese per StringSimiliarity-Algorithmus mit den Namen in der Addendum-Patientendatenbank
	; ---------------------------------------------------------------------------------------------------------------------------------------------------
		goto FindNameLabel
		spos := 1
		while (spos := RegExMatch(text, "([A-Z][\pL-]+)[\s,.]+([A-Z][\pL-]+)", name, spos)) {

				;SciTEOutput("spos: " spos)
			; überflüssige Zeichen entfernen
				name1 := RegExReplace(name1, "[\s\n\r]"),
				name2 := RegExReplace(name2, "[\s\n\r]")

			; Bindestrichworte ignorieren
				If RegExMatch(name1, "\-$") {
					spos += StrLen(name1)                    ; ein Wort weiter
					continue
				} else if RegExMatch(name2, "\-$") {
					spos += StrLen(name)                     ; beide Wörter weiter
					continue
				} else if (StrLen(name1) = 0) || (StrLen(name2) = 0) {
					spos += StrLen(name)
					continue
				}

				PatIDArr := admDB.StringSimilarityID(name1, name2)
				If IsObject(PatIDArr) {
					spos += StrLen(name)
					For idx, PatID in PatIDArr {
						If PatBuf.haskey(PatID)
							PatBuf[PatID] += 1
						else
							PatBuf[PatID] := 1
						SciTEOutput("  Methode 3:`t" PatID " , hits: " PatBuf[PatID])
					}

				}
				else
					spos += StrLen(name1)

		}
FindNameLabel:
	; höchste Trefferzahl ermitteln
		maxHits := 0, equal := false, BestHits := Object()
		For PatID, Pat in PatBuf {
			SciTEOutput("  hitpoints:   `t`t" Pat.hits " - (" PatID ") " Pat.Nn ", " Pat.Vn)
			If (Pat.hits > maxHits)
				maxHits := Pat.hits
		}

	; alle Namen mit gleicher Trefferanzahl werden behalten
		For PatID, Pat in PatBuf
			If (Pat.hits = maxHits)
				BestHits[PatID] := {"Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}

return BestHits
}

RxNames(Str, RxMatch="GetNames")                                	{                	;-- Stringfunktion für die Behandlung von PDF-Dateinamen

	static rxPerson	:= "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"
	static RxExtract	:= "^\s*(?<NN>" rxPerson ")[\s,]+(?<VN>" rxPerson ")[\s,]+(?<DokTitel>.*)*" ;?\.[a-z]*
	static RxName	:= "^\s*(?<NN>" rxPerson ")[\s,]+(?<VN>" rxPerson ")[\s,]+"

	if (RxMatch = "ContainsName")
		return RegExMatch(Str, RxName)
	else If (RxMatch = "GetNames") {
		RegExMatch(Str, RxExtract, Pat)
		return {"Nn":PatNN, "Vn":PatVN, "DokTitel":PatDokTitel}
	}
	else if (RxMatch ="ReplaceNames")
		return RegExReplace(Str, RxName)

}

FuzzyKarteikarte(NameStr, OnlyPatID=false)                   	{                 	;-- fuzzy name matching function, öffnet eine Karteikarte

	; mit OnlyPatID = true kann man nur einen String mit enthaltenem Patientennamen testen und bekommt bei einem Treffer die Patienten-ID zurück

		static rxPerson := "[A-ZÄÖÜ][\pL]+[\s-]*([A-ZÄÖÜ][\pL-]+)*"

		DiffSafe 	:= Array()
		minDiff		:= 100
		PatFound 	:= true
		umlauts 	:= {"Ä":"Ae", "Ü":"Ue", "Ö":"Oe", "ä":"ae", "ü":"ue", "ö":"oe", "ß":"sz"}


		If !IsObject(Pat := RxNames(NameStr, "GetNames")) {
			PraxTT(	"Die Dateibezeichnung enthält keinen Namen eines Patienten.`nDer Karteikartenaufruf wird abgebrochen!", "3 2")
			Sleep 3000
			return
		}

		NVName 	:= RegExReplace(Pat.Nn Pat.Vn	, "[\s]")         	; NachnameVorname
		VNName 	:= RegExReplace(Pat.Vn Pat.Nn	, "[\s]")         	; VornameNachname
		;SciTEOutput(NVName "`n" VNName)

	; Stringdifference Methode mit Suche des Patienten in der Addendum Patientendatenbank     	;{
		For PatID, Pat in oPat 		{

				If InStr(PatID, "MaxPat")
					break

			; Leerzeichen und Minus aus den Namen entfernen
				DbName	:= RegExReplace(Pat.Nn Pat.Vn, "[\s]")         	; NachnameVorname

			; Stringdifferenz Methode anwenden
				DiffA     	:= StrDiff(DBName, NVname)
				DiffB     	:= StrDiff(DbName, VNname)
				Diff        	:= DiffA <= DiffB ? DiffA : DiffB

			 ; Treffergenauigkeit (Stringsimilarity) kleiner gleich 0.11
			 ; -> passender Patient wurde gefunden. Karteikarte des Patienten wird aufgerufen
				If (Diff <= 0.11) {
					If !OnlyPatID
						admGui_Karteikarte(PatID)
					return {"ID":PatID, "Nn":Pat.Nn, "Vn":Pat.Vn, "Gd":Pat.Gd}
				}

			; sammelt den kleinsten Differenzwert
				If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Pat.Nn, bestVn := Pat.Vn, bestGd := Pat.Gd

			; enthält eine der beiden Zeichenketten 2 Zeichen mehr als die andere wird die Patienten ID gepuffert
				If ( Diff <= 0.13 && Diff > 0.11 )
					If (Abs(StrLen(DbName) - StrLen(NVname)) > 2)
						DiffSafe.Push(PatID)

		}

		If (minDiff < 0.2) {
			If !OnlyPatID
				admGui_Karteikarte(bestDiff)
			return {"ID":bestDiff, "Nn":bestNn, "Vn":bestVn, "Gd":bestGd}
		}

	; kein Treffer dann leeren String zurückgeben
		If OnlyPatID
			return "#" minDiff ", " bestDiff
		; kann entfernt werden?
		;~ If RegExMatch(DbName, "^" VNname) || RegExMatch(DbName, "^" NVname) 	{

				;~ MsgBox, 4, Patientenakte öffnen?, % "Möchten Sie die Akte des Patienten:`n(" PatID ") " Pat.Nn ", " Pat.Vn " geb.am " Pat.Gd "`nöffnen?"
				;~ IfMsgBox, Yes
					;~ AlbisAkteOeffnen(Pat.Nn ", " Pat.Vn, PatID)
				;~ IfMsgBox, No
					;~ return false
				;~ return true

		;~ }

		;}

	; Suchen über die Albis Patientensuche                                                                                	;{

		; Code nimmt die ersten beiden Buchstaben vom Vor- und Zuname und übergibt sie dem Patient öffnen Dialog
		; stehen mehrere Personen zur Auswahl im Dialog "Patient auswählen" wird per String Differenz der dichteste Treffer gesucht
		; wenn Albis für diese Kombination nur einen Treffer hat, wird sofort die Akte geöffnet (unabhängig vom Skript)
		; --- die umlautveränderten Namen werden hier für die Suche noch nicht genutzt!!
		PatMatches 	:= Object()

		PraxTT("Der Patient ist nicht in der Datenbank!`nVersuche es mit einer Patientensuche über den Albisdialog.", "6 0")
		Sleep, 200

	; öffnet den Dialog "Öffne Patient" und übergibt nur die ersten beiden Buchstaben vom Vor- und Zunamen
		AlbisDialogOeffnePatient("open", SubStr(name[1], 1, 2) ", " SubStr(name[2], 1, 2))

		while !(WinExist("Patient auswählen ahk_class #32770") || WinExist("ALBIS ahk_class #32770", "Patient <") || (A_Index > 40))
			Sleep 50

		; Patient ist nicht vorhanden ;{
			while (hwnd := WinExist("ALBIS ahk_class #32770", "Patient <")) && (A_Index < 10) 	{
				PatFound := false
				VerifiedClick("Button1", hwnd,,, true) ; OK
				WinWait, % "Patient öffnen ahk_class #32770",, 2
				If WinExist("Patient öffnen ahk_class #32770") {
					AlbisDialogOeffnePatient("close")
					PraxTT("Albis hat keine Patienten zur Auswahl gestellt.`nEs kann keine Suche durchgeführt werden.", "6 2")
					return false
				}
				sleep, 200
			}
		;}

		If (hwnd := WinExist("Patient auswählen ahk_class #32770")) {

			PatFound	:= true
			LVPats   	:= AlbisLVContent(hwnd, "SysListView321", "Name|Vorname|Geb.-Datum")
			For row, Patient in LVPats	{
				Patient.2 := StrReplace(Patient.2, "-")
				Patient.3 := StrReplace(Patient.3, "-")
				If (StrDiff(Patient.2 Patient.3, Name.1 Name.2) <= 0.12 ) || (StrDiff(Patient.2 Patient.3, Name.2 Name.1) <= 0.12 )
					PatMatches.Push(Patient)
			}

			If (PatMatches.MaxIndex() = 1) {

				hinweis :=  "Ich denke dieser Patient könnte es sein:`n`n" PatMatches[1][2] "," PatMatches[1][3] ", geb. am " PatMatches[1][4]
				hinweis .= "`nletzter Behandlungstag: " PatMatches[1][6] "`n`nSoll ich diese Akte öffnen?"

				Msgbox, 0x1024, Addendum für Albis on Windows, % hinweis
				IfMsgBox, No
					return

				VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true)
				WinWait, % "Patient öffnen ahk_class #32770",, 3
				If ErrorLevel	{
					MsgBox, Der Patient öffnen Dialog fehlt mir jetzt`num weiter zu machen!
					return false
				}
				VerifiedSetText("Edit1", PatMatches[1][5], "Patient öffnen ahk_class #32770", 200)

			; Akte wird jetzt geöffnet durch drücken von OK
				while WinExist("Patient öffnen ahk_class #32770") 	{

					VerifiedClick("Button2", "Patient öffnen ahk_class #32770")	    	; Button OK drücken
					WinWaitClose, Patient öffnen ahk_class #32770,, 1
					If WinExist("Patient öffnen ahk_class #32770") {	                     	; Fenster ist immer noch da? Dann sende ein ENTER.
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

			}
			else
				PatFound := false

			VerifiedClick("Button2", "Patient auswählen ahk_class #32770",,, true) ; Abbruch
			AlbisDialogOeffnePatient("close")

		}

	;}

return PatFound
}

class admDB {                                                                                    	;-- AddendumDB - Klasse für DB (Objekt) Handling (im Moment nur Suche)

	; das globale oPat-Objekt wird benötigt (wird per )

	; Patienten ID Suche per Stringvergleich
	; key:	kann sein Nn, Vn, Gd, Kasse
	; val: 	zu suchender Eintrag
		MatchID(key, val) {

			PatIDArr	:= Array()

			For PatID, PatData in oPat
				If (PatData[key] = val)
					PatIDArr.Push(PatID)

		return PatIDArr
		}

	; Patienten ID Suche per String-Similarity Funktion
		StringSimilarityID(name1, name2, diffmin=0.09) {

			PatIDArr	:= Array()
			minDiff 	:= 100
			NVname 	:= RegExReplace(name1 . name2, "[\s]")
			VNname 	:= RegExReplace(name2 . name1, "[\s]")

			For PatID, Pat in oPat 		{

				If InStr(PatID, "MaxPat")
					break

				DbName	:= RegExReplace(Pat.Nn Pat.Vn, "[\s]")
				DiffA     	:= StrDiff(DBName, NVname)
				DiffB     	:= StrDiff(DbName, VNname)
				Diff        	:= DiffA <= DiffB ? DiffA : DiffB

				If (Diff <= diffmin)
					PatIDArr.Push(PatID)
				else If (Diff < minDiff)
					minDiff	:= Diff, bestDiff := PatID, bestNn := Pat.Nn, bestVn := Pat.Vn, bestGd := Pat.Gd

			}

		return PatIDArr
		}

}

BefundIndex(save=true)                                                  	{	               	;-- liest den Befundordner ein

	; WICHTIG!: braucht eine globale Variable im aufrufenden Skript: ScanPool := Object()

		PDFfiles 	:= Array()

		Loop, % (MaxIndex := ScanPool.MaxIndex())
			ScanPool.RemoveAt(MaxIndex + 1 - A_Index)

	; alle pdf Dokumente in ein weiteres temp. Objekt einlesen
		Loop, Files, % Addendum.BefundOrdner "\*.pdf"
			PDFFiles.Push(A_LoopFileName)

		For idx, pdfname in PDFfiles {

			pdfPath := Addendum.BefundOrdner "\" pdfname
			FileGetSize	, FSize	, % pdfPath, K
			FileGetTime	, FTime	, % pdfPath, C
			ScanPool.Push({	"name"          	: pdfname
									, 	"filesize"        	: FSize
									,	"filetime"        	: FTime
									,	"pages"         	: GetPDFPages(pdfPath)
									,	"isSearchable"	: (isSearchablePDF(pdfPath) ? 1 : 0)})
		}

return ScanPool.MaxIndex()
}

SaveBefundIndex()                                                         	{                	;-- speichert Daten des Befundordners in einer Datei

	; speichern des Index im json Format
		JSONStr := JSON.Dump(ScanPool,, 4)
		indexfile := FileOpen(Addendum.BefundOrdner "\PDFIndex.json", "w", "UTF-8")
		indexfile.Write(JSONStr)
		indexfile.Close()

}

; ------------------------ Hilfsfunktionen
PatDir(PatID)                                                                 	{               	;-- berechnet den Addendum-Datenpfad für einen Patienten

	; Addendum speichert zusätzliche Daten zum Patienten in einem extra Ordnerverzeichnis nach folgendem Schema:
	; 	- die Sortierung erfolgt anhand der Patienten ID von Albis
	;	- maximal 1000 Unterverzeichnisse pro Ordner, ein Ordner mit der Bezeichnung 8000 enthält alle Patientenunterordner zwischen 7001-8000
	;
	; die Funktion legt noch nicht vorhandene Ordner/Unterordner automatisch an

	BaseNum	:= Round((PatID / 1000) ) * 1000
	IDBase  	:= (PatID - BaseNum <= 0) ? BaseNum : BaseNum + 1000
	PatientDir 	:= Addendum.DBPath "\PatData\" IDBase "\" PatID
	If !InStr(FileExist(PatientDir), "D")
			FileCreateDir, % PatientDir

return PatientDir
}

ClientsOnline2()                                                              	{              	;-- gibt einen Array mit den Namen der Netzwerkgeräte zurück

	For clientName, prop in Addendum.LAN.Clients {



	}


return
}

ClientsOnline()                                                              	{                	;-- gibt einen Array mit den Namen der Netzwerkgeräte zurück

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

; ------------------------ Gui / Fenster Funktionen
LV_GetColWidth(hLV, ColN) {                                               	;-- gets the width of a column

	; from AutoGui
    SendMessage 0x101F, 0, 0,, % "ahk_id " hLV ; LVM_GETHEADER
    hHeader := ErrorLevel
    cbHDITEM := (4 * 6) + (A_PtrSize * 6)
    VarSetCapacity(HDITEM, cbHDITEM, 0)
    NumPut(0x1, HDITEM, 0, "UInt") ; mask (HDI_WIDTH)
    SendMessage, % A_IsUnicode ? 0x120B : 0x1203, ColN - 1, &HDITEM,, % "ahk_id " hHeader ; HDM_GETITEMW

Return (ErrorLevel != "FAIL") ? NumGet(HDITEM, 4, "UInt") : 0
}

LV_GetScrollViewPos(hwnd) {

	Loop, % LV_GetCount() {
		SendMessage, 0x10B6, % A_Index,,, % "ahk_id " hwnd 	; LVM_ISITEMVISIBLE -> findet das erste sichtbares Item
		If ErrorLevel {
			SciTEOutput("firstvisible item:" A_Index)
			return A_Index
		}
	}

}

LV_FindRow(LV, col, searchStr) {                                           	;-- search for a string in listview col, returns row

	Gui, adm: ListView, % LV

	Loop % LV_GetCount() {
		LV_GetText(cmpStr, rowFind := A_Index, col)
		If InStr(cmpStr, searchStr)
			return rowFind
		}

return 0
}

GuiControlActive(CtrlName, GuiName) {                             	;-- bestimmt ob ein GuiControl den Eingabefocus hat

	Gui, %GuiName%: Default
	GuiControlGet, fCtrl, FocusV ;% "ahk_id " hWin
	If (fCtrl = CtrlName)
		return true

return false
}

RedrawWindow(hwnd) {                                                       	;-- zeichnet eine Autohotkey Gui komplett neu

	static RDW_INVALIDATE 	:= 0x0001
	static	RDW_ERASE           	:= 0x0004
	static RDW_FRAME          	:= 0x0400
	static RDW_ALLCHILDREN	:= 0x0080

	dllcall("RedrawWindow", "Ptr", hwnd, "Ptr", 0, "Ptr", 0, "UInt", RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN)

}

UpdateWindow(hwnd) {                                                      	;-- sends WM_Paint to Update a window
return dllcall("UpdateWindow", "Ptr", hwnd)
}

; ------------------------- WEB Gui
CreateHTML(nprop) {

	static networkHTML := "
( ; html
	<!DOCTYPE html><html>
	<head>
	<meta http-equiv='X-UA-Compatible' content='IE=edge'>
	<style>

	html, body {
			width: 100%; height: 100%;
			margin: 0; padding: 0;
			font-family: sans-serif;
		}

	body {
			display: flex;
			flex-direction: column;
		}

	.main {
			 font-size: 9pt;
			padding: 1em;
			overflow: auto;
		}

	.row {
		  display: flex;
		  flex-wrap: wrap;
		  justify-content: space-between;
		}

    .w-50 {
      width: 48%;
    }

    .w-100 {
      width: 100%;
    }

	</style>
	</head>

	<body>

	<header>
	</header>

	<div class='main'>

		<div class='row'>
			###ROWBUTTONS
		</div>

	</div>

	</body>
	</html>

)"

/*




*/

	networkHTML := StrReplace(networkHTML, "###ROWBUTTONS", "<button class='w-50' onclick='ahk.Example1_Button(event)'>Anmeldung</button>")
	htmlpath := A_Temp "\admGui_LanInterface.html"
	FileOpen(htmlpath, "w", "UTF-8").write(networkHTML)

return htmlpath

}

class NeutronEmbedded {

	/* NeutronEmbedded.ahk v1.0.0

		this is a modification of the Neutron WebGui class in order to be able to integrate it into an existing Gui

		Copyright (c) 2020 Philip Taylor (known also as GeekDude, G33kDude)
		https://github.com/G33kDude/Neutron.ahk

		MIT License

		Permission is hereby granted, free of charge, to any person obtaining a copy
		of this software and associated documentation files (the "Software"), to deal
		in the Software without restriction, including without limitation the rights
		to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
		copies of the Software, and to permit persons to whom the Software is
		furnished to do so, subject to the following conditions:

		The above copyright notice and this permission notice shall be included in all
		copies or substantial portions of the Software.

		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
		IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
		FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
		AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
		LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
		OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
		SOFTWARE.

	 */


	; --- Constants ---          	;{
	static VERSION := "1.0.0"

	; Windows Messages
	, WM_DESTROY := 0x02
	, WM_SIZE := 0x05
	, WM_NCCALCSIZE := 0x83
	, WM_NCHITTEST := 0x84
	, WM_NCLBUTTONDOWN := 0xA1
	, WM_KEYDOWN := 0x100
	, WM_KEYUP := 0x101
	, WM_SYSKEYDOWN := 0x104
	, WM_SYSKEYUP := 0x105
	, WM_MOUSEMOVE := 0x200
	, WM_LBUTTONDOWN := 0x201

	; Virtual-Key Codes
	, VK_TAB           	:= 0x09
	, VK_SHIFT         	:= 0x10
	, VK_CONTROL 	:= 0x11
	, VK_MENU       	:= 0x12
	, VK_F5             	:= 0x74

	; Non-client hit test values (WM_NCHITTEST)
	, HT_VALUES := [[13, 12, 14], [10, 1, 11], [16, 15, 17]]

	; Registry keys
	, KEY_FBE := "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\MAIN"
	. "\FeatureControl\FEATURE_BROWSER_EMULATION"

	; Undoucmented Accent API constants
	; https://withinrafael.com/2018/02/02/adding-acrylic-blur-to-your-windows-10-apps-redstone-4-desktop-apps/
	, ACCENT_ENABLE_BLURBEHIND	:= 3
	, WCA_ACCENT_POLICY          	:= 19

	; Other constants
	, EXE_NAME := A_IsCompiled ? A_ScriptName : StrSplit(A_AhkPath, "\").Pop()


	; --- Instance Variables ---

	LISTENERS := [this.WM_DESTROY, this.WM_SIZE, this.WM_NCCALCSIZE
	, this.WM_KEYDOWN, this.WM_KEYUP, this.WM_SYSKEYDOWN, this.WM_SYSKEYUP
	, this.WM_LBUTTONDOWN]

	; Maximum pixel inset for sizing handles to appear
	border_size := 6

	; The window size
	;w := 800
	;h := 1024

	; Modifier keys as seen by neutron
	MODIFIER_BITMAP := {this.VK_SHIFT: 1<<0, this.VK_CONTROL: 1<<1
	, this.VK_MENU: 1<<2}
	modifiers := 0

	; Shortcuts to not pass on to the web control
	disabled_shortcuts :=
	( Join ; ahk
	{
		0: {
			this.VK_F5: true
		},
		this.MODIFIER_BITMAP[this.VK_CONTROL]: {
			GetKeyVK("F"): true,
			GetKeyVK("L"): true,
			GetKeyVK("N"): true,
			GetKeyVK("O"): true,
			GetKeyVK("P"): true
		}
	}
	)
	;}

	; --- Properties ---          	;{
	; Get the JS DOM object
	doc[]	{

		get
		{
			return this.wb.Document
		}

	}

	; Get the JS Window object
	wnd[]	{

		get
		{
			return this.wb.Document.parentWindow
		}

	}

	;}

	; --- Construction, Destruction, Meta-Functions --- ;{
	__New(html:="", css:="", js:="", nprop:="") 	{

		/* nprop description

			nprop := {"title"         	: "Neutron"
						, 	"x"             	: 1
						, 	"y"             	: 20
						, 	"w"             	: 600
						, 	"h"            	: 400
						, 	"hparent"    	: hadm
						,	"parentgui"	: "adm"}

		 */

		static wb

		; Create necessary circular references
		this.bound := {}
		this.bound._OnMessage := this._OnMessage.Bind(this)

		; Bind message handlers
		for i, message in this.LISTENERS
			OnMessage(message, this.bound._OnMessage)

		; Create and save the GUI
		If IsObject(nprop) {
			this.hWnd      	:= nprop.hparent
			this.x             	:= nprop.x
			this.y             	:= nprop.y
			this.w            	:= nprop.w
			this.h            	:= nprop.h
			this.parentgui	:= nprop.parentgui
			this.title          	:= nprop.title
		}

		; Creating an ActiveX control with a valid URL instantiates a
		; WebBrowser, saving its object to the associated variable. The "about"
		; URL scheme allows us to start the control on either a blank page, or a
		; page with some HTML content pre-loaded by passing HTML after the
		; colon: "about:<!DOCTYPE html><body>...</body>"

		; Read more about the WebBrowser control here:
		; http://msdn.microsoft.com/en-us/library/aa752085

		; For backwards compatibility reasons, the WebBrowser control defaults
		; to IE7 emulation mode. The standard method of mitigating this is to
		; include a compatibility meta tag in the HTML, but this requires
		; tampering to the HTML and does not solve all compatibility issues.
		; By tweaking the registry before and after creation of the control we
		; can opt-out of the browser emulation feature altogether with minimal
		; impact on the rest of the system.

		; Read more about browser compatibility modes here:
		; https://docs.microsoft.com/en-us/archive/blogs/patricka/controlling-webbrowser-control-compatibility

		RegRead, fbe, % this.KEY_FBE, % this.EXE_NAME
		RegWrite, REG_DWORD, % this.KEY_FBE, % this.EXE_NAME, 0

		Gui, % this.hWnd ":Add", ActiveX, % "vwb hWndhWB x" this.x " y" this.y " w" this.w " h" this.h, about:blank
		if (fbe = "")
			RegDelete, % this.KEY_FBE, % this.EXE_NAME
		else
			RegWrite, REG_DWORD, % this.KEY_FBE, % this.EXE_NAME, % fbe

		; Save the WebBrowser control to reference later
		this.wb  	:= wb
		this.hWB	:= hWB

		; Connect the web browser's event stream to a new event handler object
		ComObjConnect(this.wb, new this.WBEvents(this))

		; Compute the HTML template if necessary
		if !(html ~= "i)^<!DOCTYPE")
			html := Format(this.TEMPLATE, css, title, html, js)

		; Write the given content to the page
		this.doc.write(html)
		this.doc.close()

		; Inject the AHK objects into the JS scope
		this.wnd.neutron 	:= this
		this.wnd.ahk     	:= new this.Dispatch(this)

		; Wait for the page to finish loading
		while wb.readyState < 4
			Sleep, 50

		; Subclass the rendered Internet Explorer_Server control to intercept
		; its events, including WM_NCHITTEST and WM_NCLBUTTONDOWN.
		; Read more here: https://forum.juce.com/t/_/27937
		; And in the AutoHotkey documentation for RegisterCallback (Example 2)

		dhw := A_DetectHiddenWindows
		DetectHiddenWindows, On
		ControlGet, hWnd, hWnd,, Internet Explorer_Server1	, % "ahk_id" this.hWnd
		this.hIES    	:= hWnd
		ControlGet, hWnd, hWnd,, Shell DocObject View1  	, % "ahk_id" this.hWnd
		this.hSDOV 	:= hWnd
		DetectHiddenWindows, %dhw%

		this.pWndProc   	:= RegisterCallback(this._WindowProc, "", 4, &this)
		this.pWndProcOld	:= DllCall("SetWindowLong" (A_PtrSize == 8 ? "Ptr" : "")
													, "Ptr"	, this.hIES          	; HWND     	hWnd
													, "Int" 	, -4                    	; int          	nIndex (GWLP_WNDPROC)
													, "Ptr"	, this.pWndProc 	; LONG_PTR	dwNewLong
													, "Ptr")                             	; LONG_PTR

		; Stop the WebBrowser control from consuming file drag and drop events
		this.wb.RegisterAsDropTarget := False
		DllCall("ole32\RevokeDragDrop", "UPtr", this.hIES)
	}
	;}

	; --- Event Handlers ---  	;{
	_OnMessage(wParam, lParam, Msg, hWnd)	{

		if (hWnd == this.hWnd)		{

			; Handle messages for the main window
			 if 	(Msg == this.WM_DESTROY)			{

				; Clean up all our circular references so that the object may be
				; garbage collected.
				for i, message in this.LISTENERS
					OnMessage(message, this.bound._OnMessage, 0)
				ComObjConnect(this.wb)
				this.bound := []

			}

		}
		else if (hWnd == this.hIES || hWnd == this.hSDOV)		{

			; Handle messages for the rendered Internet Explorer_Server
			pressed 	:= (Msg == this.WM_KEYDOWN	|| Msg == this.WM_SYSKEYDOWN)
			released	:= (Msg == this.WM_KEYUP     	|| Msg == this.WM_SYSKEYUP)

			if (pressed || released)			{

				; Track modifier states
				if (bit := this.MODIFIER_BITMAP[wParam])
					this.modifiers := (this.modifiers & ~bit) | (pressed * bit)

				; Block disabled key combinations
				if (this.disabled_shortcuts[this.modifiers, wParam])
					return 0


				; When you press tab with the last tabbable item in the
				; document already selected, focus will be taken from the IES
				; control and moved to the SDOV control. The accelerator code
				; from the AutoHotkey installer uses a conditional loop in an
				; attempt to work around this behavior, but as implemented it
				; did not work correctly on my system. Instead, listen for the
				; tab up event on the SDOV and swap it for a tab down before
				; translating it. This should prevent the user from tabbing to
				; the SDOV in most cases, though there may still be some way to
				; tab to it that I am not aware of. A more elegant solution may
				; be to subclass the SDOV like was done for the IES, then
				; forward the WM_SETFOCUS message back to the IES control.
				; However, given the relative complexity of subclassing and the
				; fact that this message substution approach appears to work
				; just as well, we will use the message substitution. Consider
				; implementing the other approach if it turns out that the
				; undesirable behavior continues to manifest under some
				; circumstances.
				Msg := hWnd == this.hSDOV ? this.WM_KEYDOWN : Msg

				; Modified accelerator handling code from AutoHotkey Installer
				Gui +OwnDialogs ; For threadless callbacks which interrupt this.
				pipa := ComObjQuery(this.wb, "{00000117-0000-0000-C000-000000000046}")
				VarSetCapacity(kMsg, 48), NumPut(A_GuiY, NumPut(A_GuiX
				, NumPut(A_EventInfo, NumPut(lParam, NumPut(wParam
				, NumPut(Msg, NumPut(hWnd, kMsg)))), "uint"), "int"), "int")
				r := DllCall(NumGet(NumGet(1*pipa)+5*A_PtrSize), "ptr", pipa, "ptr", &kMsg)
				ObjRelease(pipa)

				if (r == 0) ; S_OK: the message was translated to an accelerator.
					return 0
				return
			}
		}
	}

	_WindowProc(Msg, wParam, lParam)	{
		Critical
		hWnd := this
		this := Object(A_EventInfo)

		if (Msg == this.WM_NCHITTEST)		{
			; Check to see if the cursor is near the window border, which
			; should be treated as the "non-client" drag-to-resize area.
			; https://autohotkey.com/board/topic/23969-/#entry155480

			; Extract coordinates from LOWORD and HIWORD (preserving sign)
			x := lParam<<48>>48, y := lParam<<32>>48

			; Get the window position for comparison
			WinGetPos, wX, wY, wW, wH, % "ahk_id" this.hWnd

			; Calculate positions in the lookup tables
			row	:= (x < wX + this.BORDER_SIZE) ? 1 : (x >= wX + wW - this.BORDER_SIZE) ? 3 : 2
			col	:= (y < wY + this.BORDER_SIZE) ? 1 : (y >= wY + wH - this.BORDER_SIZE) ? 3 : 2

			return this.HT_VALUES[col, row]

		}
		else if (Msg == this.WM_NCLBUTTONDOWN)		{
			; Hoist nonclient clicks to main window
			return DllCall("SendMessage", "Ptr", this.hWnd, "UInt", Msg, "UPtr", wParam, "Ptr", lParam, "Ptr")
		}

		; Otherwise (since above didn't return), pass all unhandled events to the original WindowProc.
		Critical, Off
		return DllCall("CallWindowProc"
							, 	"Ptr"   	, this.pWndProcOld 	; WNDPROC lpPrevWndFunc
							, 	"Ptr"   	, hWnd                   	; HWND    hWnd
							, 	"UInt"	, Msg                    	; UINT    Msg
							, 	"UPtr"	, wParam              	; WPARAM  wParam
							, 	"Ptr"  	, lParam                	; LPARAM  lParam
							, 	"Ptr")                                 	; LRESULT
	}
	;}

	; --- Instance Methods ---	;{
	; Loads an HTML file by name (not path). When running the script uncompiled,
	; looks for the file in the local directory. When running the script
	; compiled, looks for the file in the EXE's RCDATA. Files included in your
	; compiled EXE by FileInstall are stored in RCDATA whether they get
	; extracted or not. An easy way to get your Neutron resources into a
	; compiled script, then, is to put FileInstall commands for them right below
	; the return at the bottom of your AutoExecute section.
	;
	; Parameters:
	;   fileName - The name of the HTML file to load into the Neutron window.
	;              Make sure to give just the file name, not the full path.
	;
	; Returns: nothing
	;
	; Example:
	;
	; ; AutoExecute Section
	; neutron := new NeutronWindow()
	; neutron.Load("index.html")
	; neutron.Show()
	; return
	; FileInstall, index.html, index.html
	; FileInstall, index.css, index.css
	;
	Load(fileName)	{
		; Complete the path based on compiled state
		if A_IsCompiled
			url := "res://" this.wnd.encodeURIComponent(A_ScriptFullPath) "/10/" fileName
		else
			url := A_WorkingDir "/" fileName

		; Navigate to the calculated file URL
		this.wb.Navigate(url)

		; Wait for the page to finish loading
		while this.wb.readyState < 3
			Sleep, 50

		; Inject the AHK objects into the JS scope
		this.wnd.neutron 	:= this
		this.wnd.ahk     	:= new this.Dispatch(this)

		; Wait for the page to finish loading
		while this.wb.readyState < 4
			Sleep, 50
	}

	; Shorthand method for document.querySelector
	qs(selector)	{
		return this.doc.querySelector(selector)
	}

	; Shorthand method for document.querySelectorAll
	qsa(selector)	{
		return this.doc.querySelectorAll(selector)
	}

	; Passthrough method for the Gui command, targeted at the Neutron Window
	; instance
	Gui(subCommand, value1:="", value2:="", value3:="")	{
		Gui, % this.hWnd ":" subCommand, %value1%, %value2%, %value3%
	}
	;}

	; --- Static Methods ---    	;{
	; Given an HTML Collection (or other JavaScript array), return an enumerator
	; that will iterate over its items.
	;
	; Parameters:
	;     htmlCollection - The JavaScript array to be iterated over
	;
	; Returns: An Enumerable object
	;
	; Example:
	;
	; neutron := new NeutronWindow("<body><p>A</p><p>B</p><p>C</p></body>")
	; neutron.Show()
	; for i, element in neutron.Each(neutron.body.children)
	;     MsgBox, % i ": " element.innerText
	;
	Each(htmlCollection)	{
		return new this.Enumerable(htmlCollection)
	}

	; Given an HTML Form Element, construct a FormData object
	;
	; Parameters:
	;   formElement - The HTML Form Element
	;   useIdAsName - When a field's name is blank, use it's ID instead
	;
	; Returns: A FormData object
	;
	; Example:
	;
	; neutron := new NeutronWindow("<form>"
	; . "<input type='text' name='field1' value='One'>"
	; . "<input type='text' name='field2' value='Two'>"
	; . "<input type='text' name='field3' value='Three'>"
	; . "</form>")
	; neutron.Show()
	; formElement := neutron.doc.querySelector("form") ; Grab 1st form on page
	; formData := neutron.GetFormData(formElement) ; Get form data
	; MsgBox, % formData.field2 ; Pull a single field
	; for name, element in formData ; Iterate all fields
	;     MsgBox, %name%: %element%
	;
	GetFormData(formElement, useIdAsName:=True)	{

		formData := new this.FormData()

		for i, field in this.Each(formElement.elements)		{
			; Discover the field's name
			name := ""
			try ; fieldset elements error when reading the name field
				name := field.name
			if (name == "" && useIdAsName)
				name := field.id

			; Filter against fields which should be omitted
			if (name == "" || field.disabled || field.type ~= "^file|reset|submit|button$")
				continue

			; Handle select-multiple variants
			if (field.type == "select-multiple")			{
				for j, option in this.Each(field.options)
					if (option.selected)
						formData.add(name, option.value)
				continue
			}

			; Filter against unchecked checkboxes and radios
			if (field.type ~= "^checkbox|radio$" && !field.checked)
				continue

			; Return the field values
			formData.add(name, field.value)
		}

		return formData
	}

	; Given a potentially HTML-unsafe string, return an HTML safe string
	; https://stackoverflow.com/a/6234804
	EscapeHTML(unsafe)	{
		unsafe := StrReplace(unsafe, "&", "&amp;")
		unsafe := StrReplace(unsafe, "<", "&lt;")
		unsafe := StrReplace(unsafe, ">", "&gt;")
		unsafe := StrReplace(unsafe, """", "&quot;")
		unsafe := StrReplace(unsafe, "''", "&#039;")
		return unsafe
	}

	; Wrapper for Format that applies EscapeHTML to each value before passing
	; them on. Useful for dynamic HTML generation.
	FormatHTML(formatStr, values*)	{
		for i, value in values
			values[i] := this.EscapeHTML(value)
		return Format(formatStr, values*)
	}
	;}

	; --- Nested Classes ---    	;{
	; Proxies method calls to AHK function calls, binding a given value to the
	; first parameter of the target function.
	;
	; For internal use only.
	;
	; Parameters:
	;   parent - The value to bind
	;
	class Dispatch	{

		__New(parent)		{
			this.parent := parent
		}

		__Call(params*)		{
			; Make sure the given name is a function
			if !(fn := Func(params[1]))
				throw Exception("Unknown function: " params[1])

			; Make sure enough parameters were given
			if (params.length() < fn.MinParams)
				throw Exception("Too few parameters given to " fn.Name ": " params.length())

			; Make sure too many parameters weren't given
			if (params.length() > fn.MaxParams && !fn.IsVariadic)
				throw Exception("Too many parameters given to " fn.Name ": " params.length())

			; Change first parameter from the function name to the neutron instance
			params[1] := this.parent

			; Call the function
			return fn.Call(params*)
		}
	}

	; Handles Web Browser events
	; https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa768283%28v%3dvs.85%29
	;
	; For internal use only
	;
	; Parameters:
	;   parent - An instance of the Neutron class
	;
	class WBEvents	{

		__New(parent)		{
			this.parent := parent
		}

		DocumentComplete(wb)		{
			; Inject the AHK objects into the JS scope
			wb.document.parentWindow.neutron := this.parent
			wb.document.parentWindow.ahk := new this.parent.Dispatch(this.parent)
		}
	}

	; Enumerator class that enumerates the items of an HTMLCollection (or other
	; JavaScript array).
	;
	; Best accessed through the .Each() helper method.
	;
	; Parameters:
	;   htmlCollection - The HTMLCollection to be enumerated.
	;
	class Enumerable	{
		i := 0

		__New(htmlCollection)		{
			this.collection := htmlCollection
		}

		_NewEnum()		{
			return this
		}

		Next(ByRef i, ByRef elem)		{
			if (this.i >= this.collection.length)
				return False
			i := this.i
			elem := this.collection.item(this.i++)
			return True
		}
	}

	; A collection similar to an OrderedDict designed for holding form data.
	; This collection allows duplicate keys and enumerates key value pairs in
	; the order they were added.
	class FormData	{

		names := []
		values := []

		; Add a field to the FormData structure.
		;
		; Parameters:
		;   name - The form field name associated with the value
		;   value - The value of the form field
		;
		; Returns: Nothing
		;
		Add(name, value)		{
			this.names.Push(name)
			this.values.Push(value)
		}

		; Get an array of all values associated with a name.
		;
		; Parameters:
		;   name - The form field name associated with the values
		;
		; Returns: An array of values
		;
		; Example:
		;
		; fd := new NeutronWindow.FormData()
		; fd.Add("foods", "hamburgers")
		; fd.Add("foods", "hotdogs")
		; fd.Add("foods", "pizza")
		; fd.Add("colors", "red")
		; fd.Add("colors", "green")
		; fd.Add("colors", "blue")
		; for i, food in fd.All("foods")
		;     out .= i ": " food "`n"
		; MsgBox, %out%
		;
		All(name)		{
			values := []
			for i, v in this.names
				if (v == name)
					values.Push(this.values[i])
			return values
		}

		; Meta-function to allow direct access of field values using either dot
		; or bracket notation. Can retrieve the nth item associated with a given
		; name by passing more than one value in when bracket notation.
		;
		; Example:
		;
		; fd := new NeutronWindow.FormData()
		; fd.Add("foods", "hamburgers")
		; fd.Add("foods", "hotdogs")
		; MsgBox, % fd.foods ; hamburgers
		; MsgBox, % fd["foods", 2] ; hotdogs
		;
		__Get(name, n := 1)		{
			for i, v in this.names
				if (v == name && !--n)
					return this.values[i]
		}

		; Allow iteration in the order fields were added, instead of a normal
		; object's alphanumeric order of iteration.
		;
		; Example:
		;
		; fd := new NeutronWindow.FormData()
		; fd.Add("z", "3")
		; fd.Add("y", "2")
		; fd.Add("x", "1")
		; for name, field in fd
		;     out .= name ": " field ","
		; MsgBox, %out% ; z: 3, y: 2, x: 1
		;
		_NewEnum()		{
			return {"i": 0, "base": this}
		}
		Next(ByRef name, ByRef value)		{
			if (++this.i > this.names.length())
				return False
			name := this.names[this.i]
			value := this.values[this.i]
			return True
		}
	}

	;}

}




; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                     DOKUMENT EXPORTER
;
;      Funktion:  		 	    -	Dokumente (PDF, Bilder, Briefe) finden, wählen und mit Karteikartentext als Dateinamen exportieren
;                       		    -	Laborblatt und Karteikarte können mit einem Klick zusammen mit den Dokumenten exportiert werden
;									-	bei Aufruf des Skriptes über das Infofenster wird der aktuelle Patient vorausgewählt
;									-	gut nutzbar auch für die Patientensuche z.B. alle aus einer Straße oder Ort oder Arbeit oder Jahrgang
;
;      	Basisskript:    		-	keines
;		Abhängigkeiten: 	-	siehe #includes
;
;	                    				Addendum für Albis on Windows
;                        				by Ixiko started in September 2017 - last change 09.06.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;	Version 0.88

	; Skripteinstellungen                                        	;{

		#NoEnv
		#KeyHistory                   	, 0
		#MaxThreads                  	, 100
		#MaxThreadsBuffer       	, On
		#MaxMem                    	, 256

		SetBatchLines                	, -1
		SetControlDelay               	, -1
		SetWinDelay                  	, -1
		ListLines                        	, Off

	;}

	; Variablen                                                       	;{
		global q:=Chr(0x22)
		global DBFPatient, AlbisPath, AlbisDBPath, exportordner, AddendumDir, suchfeldeintragung, qpdfPath
		global DXP, hDXP, DXP_Pat, DXPhPat, DXP_PLV, DXP_DLV, DXP_PE, DXP_SB, DXP_FF, DXP_AB, DXP_EB, DXP_FB, DXP_NR, DXP_NV
		global FSize, edWidth
		global docfilter_full, docfilter_checked, pdfViewer, imgViewer, docViewer
		global PatDok := Object(), PREVID := Object()
		global Addendum := Object()
		global PatChoosed := false


		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		Addendum.qpdfPath    	:= AddendumDir "\include\OCR\qpdf"
		Addendum.AlbisDBPath	:= GetAlbisPath() "\db"

		Menu, Tray,  Icon, % AddendumDir "\assets\ModulIcons\DokExport.ico"
	;}

	; Albis Basispfad                                            	;{
		SetRegView, % (A_PtrSize = 8 ? 64 : 32)
		RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
		AlbisDBPath := AlbisPath "\DB"
	;}

	; INI Read                                                     	;{
		IniRead, docfilter_full          	, % A_ScriptDir "\admExporter.ini", DokExporter, docfilter_full            	, % "pdf,doc,docx,rtf,txt,jpg,bmp,gif,png,wav,avi,mov"

		IniRead, docfilter_checked  	, % A_ScriptDir "\admExporter.ini", DokExporter, docfilter_checked    	, % "pdf,doc,docx,rtf,txt,jpg,bmp"
		If InStr(docfilter_checkd, "ERROR")
			docfilter_checkd := ""

		IniRead, docfilter_bezeichner	, % A_ScriptDir "\admExporter.ini", DokExporter, docfilter_bezeichner	, % "Verordnung,Einnahme,Rezept,Fax,Medikamenten,chron. krank, Notfall-Vertretungsschein"

		IniRead, suchfeldeintragung  	, % A_ScriptDir "\admExporter.ini", DokExporter, suchfeldeintragung
		If InStr(suchfeldeintragung, "ERROR")
			suchfeldeintragung := ""

		IniRead, ini_yearsback       	, % A_ScriptDir "\admExporter.ini", DokExporter, exportJahre           	, % "2"

		IniRead, pdfViewer             	, % A_ScriptDir "\admExporter.ini", PREVIEWER, pdf                        		, % "SumatraPDF.exe "
		If InStr(pdfViewer, "ERROR")
			pdfViewer := ""

		IniRead, imgViewer             	, % A_ScriptDir "\admExporter.ini", PREVIEWER, img                       		, % "i_view64.exe "
		If InStr(imgViewer, "ERROR")
			imgViewer := ""

		IniRead, docViewer             	, % A_ScriptDir "\admExporter.ini", PREVIEWER, doc                         	, % "notepad.exe"
		If InStr(docViewer, "ERROR")
			docViewer := ""

		IniRead, exportordner        	, % A_ScriptDir "\admExporter.ini", DokExporter, exportordner
		If InStr(exportordner, "ERROR") || (StrLen(exportordner) = 0) {
			IniRead, exportordner, % AddendumDir "\Addendum.ini", % "Scanpool", % "Exportordner"
			If InStr(exportordner, "ERROR") || (StrLen(exportordner) = 0) {
				IniRead, exportordner, % "C:\albiswin.loc\Local.ini", % "bvl", "VP6001"
				If InStr(exportordner, "ERROR") || (StrLen(exportordner) = 0) {
					MsgBox, 1, % "Addendum für Albis on Windows", % "Das Skript kann kein Verzeichnis für den Export der Patientendokumente finden.`n"
										. "Bitte hinterlegen Sie einen Pfad im entsprechenden Feld in der Gui"
				}
				exportordner := A_ScriptDir "\Dokumentexport"
			}
		}


	;}

	; Karteikartenkürzel                                         	;{
		DBaccess 	:= new AlbisDB(Addendum.AlbisDBPath)
		KKFilter 	:= DBaccess.Karteikartenfilter()
		For filter, data in KKFilter
			DbFilter .= filter "|"
		DBFilter := RTrim(DBFilter, "|")
	;}

	; DIE GUI 	                                                       	;{
	AuswahlGui:

	;-: VARIABLEN/GUIGRÖSSE                           	;{
		debug   	:= 0
		FSize     	:= (A_ScreenHeight > 1080 ? 10 : 9)
		tcolor    	:= "DDDDBB"
		gdxp     	:= " gDXP_Handler "
		gopts1  	:= " Center BackgroundTrans "
		extRows 	:= 15

		sbWidth	:= (6 	 * FSize) + 10
		edWidth	:= (150 * FSize) - sbWidth - 15
		edWidth	:= edWidth > A_ScreenWidth - 20	 ? A_ScreenWidth - 20 : edWidth
		sfWidth 	:= (70 * FSize) - sbWidth - 15
		otWidth 	:= edWidth + sbWidth + 5
	;}

	;-: GUI  ANFANG 0x00180000                        	;{
		Gui, DXP: new, HWNDhDXP +E0x100000 ;-DPIScale
		Gui, DXP: Color, cCCCCAA, cDDDDBB
	;}

	;-: PATIENTENSUCHE                                       	;{
		Gui, DXP: Font, % "s" FSize " q5 Normal", Futura Bk Bt
		Gui, DXP: Add, Text          	, % "xm ym BackgroundTrans"      	, % "Pat.Nr. | Nachname | Nachname, Vorname | , Vorname | Geburtsdatum | A:Arbeit | O:Ort | Straße Nr"
		Gui, DXP: Font, % "s" FSize-3 " q5 Normal"
		Gui, DXP: Add, Text          	, % "x+1 ym-1 BackgroundTrans"                                                                          	, % "1"
		Gui, DXP: Add, Text          	, % "x+1 ym-1 c2222AA BackgroundTrans   vDXP_One"                                        	, % "1"
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Text          	, % "x+10 ym+1 c2222AA BackgroundTrans vDXP_OneText"                                 	, % "   Zahl oder/und Buchstabe oder * z.B. Feldweg 1, Ahornallee 5A oder Hauptstraße *"
		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Combobox	, % "xm y+5 w" sfWidth " r20 vDXP_Pat gDXP_Handler +hwndDXPhPat"                  	, % suchfeldeintragung
		GuiControl, DXP: Choose, DXP_Pat, 1
		Gui, DXP: Add, Button      	, % "x+5 w" sbWidth  " vDXP_SB   gDXP_Handler"                                                	, % "Suchen"
		GuiControlGet, cp, DXP: Pos, DXP_Pat
		GuiControl, DXP: Move, DXP_SB, % "y" cpY-1 " h" cpH+2

		Gui, DXP: Add, Text          	, % "x+50	y" cpY+cpH-15 "   "                                                                             	, % "gefundene Patienten: "
		Gui, DXP: Add, Text          	, % "x+5 	y" cpY+cpH-15 "    vDXP_PZ"                                                                  	, % "                            "
	;}

	;-: LISTVIEW PATIENTEN                                  	;{
		Gui, DXP: Font, % "s" FSize " q5 Normal"
		LVOPtions := "AltSubmit -Multi Grid Background" tcolor
		Gui, DXP: Add, ListView    	, % "xm y+5 w" otWidth " r12 vDXP_PLV gDXP_Handler " LVOptions, % "NR|Name|Vorname|Geburtstag|PLZ|Ort|Strasse|Telefon|Arbeit|letzte Beh|timestamp"

		LV_ModifyCol(1 	, Round(edWidth/16) " Integer")	    	; NR
		LV_ModifyCol(2 	, Round(edWidth/9))                        	; Name
		LV_ModifyCol(3 	, Round(edWidth/9))                        	; Vorname
		LV_ModifyCol(4 	, Round(edWidth/12) " Integer Left")   	; Geburtstag
		LV_ModifyCol(5 	, Round(edWidth/15) " Integer Left") 	; PLZ
		LV_ModifyCol(6 	, Round(edWidth/8))                       	; Ort
		LV_ModifyCol(7 	, Round(edWidth/8))                        	; Straße
		LV_ModifyCol(8 	, Round(edWidth/11))                       	; Telefon
		LV_ModifyCol(9 	, Round(edWidth/6))                        	; Arbeit
		LV_ModifyCol(10	, Round(edWidth/12))                       	; letzte Behandlung
		LV_ModifyCol(11	, "0 Integer")
	;}

	;-: LISTVIEW FÜR DOKUMENTE                      	;{
		DLVWidth := Round(otWidth*0.60)
		Gui, DXP: Add, Text          	, % "xm y+5 BackgroundTrans" 	                                                                            	, % "Dokumente von ["
		Gui, DXP: Add, Text          	, % "x+2 BackgroundTrans vDXP_NR"    	                                                                	, % "888888"
		Gui, DXP: Add, Text          	, % "x+2 BackgroundTrans "                                                                                   	, % "] "
		Gui, DXP: Add, Text          	, % "x+0 w350 BackgroundTrans vDXP_NV"                                                            	, % ""
		GuiControl,  DXP: , DXP_NR, % "             "

		LVOptions := "Grid Count Checked AltSubmit Background" tcolor " hwndDXP_hDLV"
		DocCols := "exp.|OCR|Seiten"
		Gui, DXP: Add, ListView   	, % "xm y+3 w" DLVWidth " r20 vDXP_DLV	gDXP_Handler " LVOptions                      	, % "Datum|Art|Bezeichnung|exp.|OCR|Seiten|LFDNR"
		LV_ModifyCol(1, Round(DLVWidth/8) " Integer")
		LV_ModifyCol(2, Round(DLVWidth/10) )
		LV_ModifyCol(3, Round(DLVWidth/2))
		LV_ModifyCol(4, "40 Center")
		LV_ModifyCol(5, "40 Center")
		LV_ModifyCol(6, "50 Center")
		LV_ModifyCol(7, "70 Center") ; Spalte für interne Nutzung
	;}

	;-: EINSTELLUNGEN/FILTER                           	;{
		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Text, % "x+10  BAckgroundTrans                                	vDXP_ET"                                              	, % "Basispfad für den Dokumentenexport"

		PBWidth := 6 * (FSize-1) - 5
		PEWidth := otWidth - DLVWidth - PBWidth - 15
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Edit           	, % "y+3  w" PEWidth "        						vDXP_PE " gdxp                                       	, % exportordner
		Gui, DXP: Add, Button        	, % "x+5  w" PBWidth "       						vDXP_PB " gdxp                                      	, % "ändern"
		GuiControlGet, cp, DXP: pos, DXP_PE

		;-: Dateifilter
		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Text, % "x" cpX " y" (cpY + cpH +15) " BackgroundTrans	vDXP_FT "                                             	, % "Dateien die erfasst werden sollen"
		GuiControlGet, cp, DXP: pos, DXP_FT
		Gui, DXP: Font, % "s" FSize " q5 Normal"

		For chkidx, ext in StrSplit(docfilter_full, ",")
			Gui, DXP: Add, Checkbox	, % "x" (A_Index=1 ? cpX    	: Mod(A_Index,extRows)=0 ? cpX : "+5")
													. 	      (A_Index=1 ? " y+5"	: Mod(A_Index,extRows)=0 ? " y+5 " : "") " vDXP_C" (lastExtCB_Index := A_Index)
													. 	  	  ( inFilter(ext, docfilter_checked) ? " Checked" : "") " gDXP_Handler"	, % ext

		;-: Bezeichnungsfilter
		;{
		GuiControlGet, cp	, DXP: pos, % "DXP_C" 	lastExtCB_Index
		GuiControlGet, dp, DXP: pos, % "DXP_C1"
		Y := (cpY + cpH +15), docfX := dpX, docfW := PBWidth + PEWidth + 5
		;}

		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Text, % "x" dpX " y" Y "                  BackgroundTrans    	vDXP_FFT "                        							, % "Dokumentbezeichnungen die gefiltert werden sollen(Komma getrennt)"
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Edit, % "x" dpX " y+3 w" docfW " r4                     		vDXP_FF"									gdxp            	, % docfilter_bezeichner
   ;}

	;-: AUSWAHL                                                	;{
		;{
		GuiControlGet, cp, DXP: pos, % "DXP_FF"
		Y1 := (cpY + cpH + 20)
		;}

		Gui, DXP: Font, % "s" FSize-1 " q5 Italic "
		Gui, DXP: Add, Text, % "x" cpX+2 " y" Y1-17 "          BackgroundTrans    	               "                        						, % "Dokumente für den Export auswählen"

		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Button        	, % "x" cpX " y" Y1 "	                            	vDXP_AB"									gdxp          	, % "Alle"
		Gui, DXP: Add, Button        	, % "x+5           	                            	vDXP_KB"									gdxp          	, % "Keines"

		;{
		GuiControlGet, cp, DXP: pos, % "DXP_AB"
		xtraOpt := " y" Y1+2 " w" Floor((FSize)*2.2) " h" cpH-4 " Number	"
		;}
		Gui, DXP: Add, Button        	, % "x+5           	                            	vDXP_LB"									gdxp          	, % "zurückliegende Jahre"
		Gui, DXP: Add, Edit            	, % "x+2" xtraOpt "                            	vDXP_JB"      	                     		gdxp          	, % ini_yearsback

		Gui, DXP: Add, Button        	, % "x" cpX " y+10 Center                     	vDXP_EA"									gdxp           	, % "000000 Dokumente exportieren"
		Gui, DXP: Add, Button        	, % "y+10         	                            	vDXP_EP"									gdxp           	, % "Auswahl, Laborblatt und Karteikarte exportieren"

		;{
		GuiControl, DXP:          	, DXP_EA, % "Dokumente exportieren"
		GuiControl, DXP: Disable	, DXP_EA
		GuiControlGet, cp, DXP: pos, % "DXP_EP"
		bpX := cpX+cpW+10
		;}

		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Text            	, % "x" bpX " y" cpY-(FSize)-12 " 	    		vDXP_VBT" 								gopts1         	, % "von - bis"
		Gui, DXP: Font, % "s" FSize+1 " q5 Normal"
		Gui, DXP: Add, Edit           	, % "y+5 w" (FSize-1)*20 " h" cpH "     	vDXP_VB   	hwndDXP_hVB"                         	, % "01.01.2010-" A_DD "." A_MM "." A_YYYY

		;{
		WinSet, ExStyle, 0x0 , % "ahk_id " DXP_hVB
		GuiControlGet, dp, DXP: pos, % "DXP_VBT"
		rY1 := cpY + Floor((cpH/2-dpH/2)/2) - 2
		GuiControl, DXP: Move, DXP_VB	, % "y" rY1
		GuiControlGet, ep, DXP: pos, % "DXP_VB"
		rY2 := epY - dpH - 2
		GuiControl, DXP: Move, DXP_VBT	, % "x" epX + Floor((cpW/2-dpW/2)/2) " y" rY2
		X := epX+epW+10, W := docfX+docfW - X, H := cpH
		;}

		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Text            	, % "x" X " y" Y1-10 " w" (FSize-1)*10 " 	vDXP_KKFT " 							gopts1         	, % "Karteikartenfilter"
		Gui, DXP: Font, % "s" FSize+2 " q5 Normal"
		Gui, DXP: Add, ListBox          	, % "x" X " y" Y1 " w" W " h" H " r1 			vDXP_KKF 	hwndDXP_hKKF "	gdxp           	, % DbFilter

		;{
		WinSet, ExStyle, 0x0 , % "ahk_id " DXP_hKKF
		GuiControl, DXP: Move, DXP_KKF	, % "h" cpH - 1
		GuiControlGet, dp, DXP: pos, % "DXP_KKFT"
		GuiControlGet, ep, DXP: pos, % "DXP_KKF"
		Y := epY - dpH - 2
		GuiControl, DXP: Move, DXP_KKFT	, % "x" epX + Floor((epW/2-dpW/2)) " y" Y
		W := dpW+epW+10
		;}

		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Text            	, % "x" epX " y" epY+epH+3 " w" epW " 	vDXP_KKFIT " 			    			gopts1        	, % "[Filtername]"
		Gui, DXP: Font, % "s" FSize " q5 Normal"
		Gui, DXP: Add, Edit            	, % "x" epX " y+1"  " w" epW " h150 		vDXP_KKFI 	hwndDXP_hKKFI"                    	, %  KKFilter["*ohne*"].inhalt

		;{
		WinSet, Style 	, 0x50000904 , % "ahk_id " DXP_hKKFI
		WinSet, ExStyle	, 0x00000000 , % "ahk_id " DXP_hKKFI
		;}

		Gui, DXP: Font, % "s" FSize " q5 Normal"
		Gui, DXP: Add, Button        	, % "x" cpX " y" cpY+cpH+26 " 				vDXP_FB" 						    		gdxp            	, % "Exportpfad im Explorer öffnen"
	;}

	;-: STATUS BAR                                               	;{
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, StatusBar  	,                                                                                                                           	, % ""
		Gui, DXP: Show, % "AutoSize", % "Dokumentexport"

		WinGetPos, wx, wy, ww, wh, % "ahk_id " hDXP
		w := GetWindowSpot(hDXP)

		GuiControlGet, cp, DXP: Pos, DXP_OneText
		GuiControlGet, dp, DXP: Pos, DXP_One
		GuiControl, DXP: Move, DXP_OneText	, % "x" (w.CW-cpW-10)
		GuiControl, DXP: Move, DXP_One     	, % "x" (w.CW-cpW-dpW-3)

		Gui, DXP: Show

		sb1Width	:= Round(ww/4)
		sb2Width	:= 5*FSize
		sb3Width	:= Round(ww/3.5)
		sb4Width	:= ww - sb1Width - sb2Width - sb3Width
		SB_SetParts(sb1Width, sb2Width, sb3Width, sb4Width )
		SB_SetText("...warte auf Deine Eingabe", 1)
		SB_SetText(AlbisPath " | " exportordner, 3)
	;}

	;-: HOTKEY                                                    	;{
		fn_EnterSearch := Func("DXPSearch").Bind("Hotkey")
		Hotkey, IfWinActive, % "Dokumentexport ahk_class AutoHotkeyGUI"
		Hotkey, Enter, % fn_EnterSearch
		Hotkey, IfWinActive
	;}

	;-: INTERSKRIPTKOMMUNIKATION                	;{
		OnMessage(0x4A	, "Receive_WM_COPYDATA")

		;~ Gui, DXP: Submit, NoHide
		GuiControlGet, KUERZEL, DXP:, DXP_KKF

	;}

return ;}

	; GUI LABELS                                                   	;{

DXP_Handler:                                                    	;{

	Gui, DXP: Submit, NoHide

	Switch A_GuiControl	{

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Listviews füllen --
		Case "DXP_SB":                                                     ; Patienten suchen
			DXPSearch("Button")

		Case "DXP_PLV":                                                	; Patienten Listview
			If (A_GuiControlEvent = "Normal")
				ZeigeDokumente(A_EventInfo)

		Case "DXP_DLV":                                                	; Dokumenten Listview
			If (A_GuiEvent = "DoubleClick")
				DokumentVorschau(A_EventInfo)
			else If (A_GuiEvent = "Normal") || (A_GuiEvent = "C")
				LV_CountCheckedDocs()

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Auswählen --
		Case "DXP_AB":                                                 	; Alle Befunde Button
			LV_ChooseAll()

		Case "DXP_KB":                                                 	; keine Befunde Button
			LV_ChooseNone()

		Case "DXP_LB":                                                    	; die letzten X Jahre Button
			LV_ChooseLastYears(DXP_JB)

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Export --
		Case "DXP_EA":                                                 	; Auswahl exportieren Button
			If PatChoosed
				LVExport()

		Case "DXP_EP":                                                 	; Karteikarte, Laborblatt u. Auswahl exportieren Button
			If PatChoosed {
				LVExport()
				KarteikartenExport()
			}

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Exportordner --
		Case "DXP_FB":                                                  	; Exportpfad im Explorer öffnen
			If PatChoosed
				LVShowExplorer()

		Case "DXP_PB":                                                  	; Exportordner ändern
			ChangeFolder()

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		Case "DXP_KKF":
			GuiControl, DXP:, DXP_KKFIT, % "[" DXP_KKF "]"
			GuiControl, DXP:, DXP_KKFI	, % KKFilter[DXP_KKF].inhalt

	}

return ;}

DXPGuiClose:
DXPGuiEscape:                                                 	;{

	; Klickfilter sichern
	For index, ext in StrSplit(docfilter_full, ",") {

		GuiControlGet, isChecked, DXP:, % "DXP_C" A_Index
		If (isChecked = 1)
			checkedext .= (A_Index=1?"":",") ext
	}
	If (docfilter_checked <> checkedext)
		IniWrite, % checkedext, % A_ScriptDir "\admExporter.ini", DokExporter, docfilter_checked

	; Sucheintragungen sichern
	suchfeld := CB_GetEntries(DXPhPat)
	If (suchfeldeintragung <> suchfeld)
		IniWrite, % suchfeld, % A_ScriptDir "\admExporter.ini", DokExporter, suchfeldeintragung

	; exportordner
	GuiControlGet, export, DXP:, % "DXP_PE"
	IF (exportordner <> export)
		IniWrite, % export, % A_ScriptDir "\admExporter.ini", DokExporter, exportordner

	; Dokumentbezeichner
	GuiControlGet, docbez, DXP:, % "DXP_FF"
	IF (docfilter_bezeichner <> docbez)
		IniWrite, % docbez, % A_ScriptDir "\admExporter.ini", DokExporter, docfilter_bezeichner

	Gui, DXP: Submit, Hide
	IF (ini_yearsback <> DXP_JB)
		IniWrite, % DXP_JB, % A_ScriptDir "\admExporter.ini", DokExporter, exportJahre_zurueckliegend


ExitApp ;}

;}

DimmerGui(show, DBASEFileName:="")             	{

	global

	DPGS_init := true

	If (show = "off") {
		Gui, DMR: Destroy
		hDimmer 	:= 0
		DPGS_init	:= false
		return
	} else if (show = "on") && hDimmer {
		Gui, DMR: Default
		GuiControl, DMR: , DMR_DBFN	, % DBASEFileName
		return
	}

	If !DBASEFileName
		DBASEFileName := "                            "

	DXPPos := GetWindowSpot(hDXP)
	ControlGetPos,,,,SBh, msctls_statusbar321, % "ahk_id " hDXP

	Gui, DMR: New    	, -Caption +LastFound + ToolWindow +hwndhDimmer +E0x4 +Parent%hDXP%
	Gui, DMR: Color   	, Gray

	Gui, DMR: Font	   	, s100 q5 bold italic cBlue, Calibri
	Gui, DMR: Add    	, Text    	, % "x0   	y" Floor(DXPPos.CH/4) " w" DXPPos.CW " vDMR_SL1 Backgroundtrans Center"	, % "...Suche läuft"

	GuiControlGet, cp	, DMR: Pos	, DMR_SL1
	Gui, DMR: Font	   	, s24 q5 bold cBlue
	Gui, DMR: Add    	, Text    	, % "x0 y" cpY+cpH-35 " w" DXPPos.CW " vDMR_DBFN Center Backgroundtrans"            	, % DBASEFileName

	Gui, DMR: Font	   	, s28 q5 bold cDarkBlue
	Gui, DMR: Add    	, Text    	, % "x0 y+15 Right vDMR_SL2 Backgroundtrans"                                                           	, % "0000000"
	Gui, DMR: Add    	, Text    	, % "x+0 Left vDMR_SL3 Backgroundtrans"                                                                  	, % "/0000000"

	GuiControlGet, cp	, DMR: Pos	, DMR_SL2
	Gui, DMR: Add    	, Progress	, % "x10 	y" cpY+cpH " w" DXPPos.CW-20 " h40 vDMR_PGS1 HWNDDMR_hPGS1 ", 0

	Gui, DMR: Show   	, % "Hide NA x" 0 " y" 0 " w" DXPPos.CW " h" DXPPos.CH-SBh, Dimmer 100
	WinSet, Redraw,, % "ahk_id " hDXP

	DMRPos := GetWindowSpot(hDimmer)
	GuiControlGet, dp, DMR: Pos	, DMR_SL2
	GuiControlGet, ep, DMR: Pos	, DMR_SL3
	bW := dpW+epW
	GuiControl, DMR: Move, DMR_SL2, % "x" Floor(DMRPos.CW/2 - bw/2) " y" dpY-25
	GuiControlGet, dp, DMR: Pos	, DMR_SL2
	GuiControl, DMR: Move, DMR_SL3, % "x" dpX + dpW + 5 " y" dpY

	GuiControl, DMR: , DMR_SL2, % ""
	GuiControl, DMR: , DMR_SL3, % "/"

	WinSet, Redraw,, % "ahk_id " hDXP
	WinSet, Transparent, 200, % "ahk_id " hDimmer
	WinSet, Top,,  % "ahk_id " hDimmer
	WinSet, AlwaysOnTop, On,  % "ahk_id " hDimmer

	Gui, DMR: Show
	WinSet, Redraw,, % "ahk_id " hDXP

return
}

DimmerPGS(recordNr, RecordsMax, len)             	{

	global DPGS_init, DBASEFileName, DBASEFileName_old

	Gui, DMR: Default

	If DPGS_init {
		DPGS_init := 0
		GuiControl, DMR: , DMR_SL3	, % "/" RecordsMax
		GuiControl, DMR: , DMR_DBFN	, % DBASEFileName
		DBASEFileName_old := DBASEFileName
	}

	If (DBASEFileName_old <> DBASEFileName) {
		GuiControl, DMR: , DMR_DBFN	, % DBASEFileName
		DBASEFileName_old := DBASEFileName
	}

	GuiControl, DMR: , DMR_SL2, % recordNr
	GuiControl, DMR:, DMR_PGS1, % Floor(pgs := recordNr*100/RecordsMax)

	Gui, DXP: Default
	SB_SetText("`t`t" Round(pgs), 2)

}

inFilter(ext, filter)                                                 	{
	For i, sfilt in StrSplit(filter, ",")
		If (ext = sfilt)
			return true
return false
}

GetCheckedFilter()                                             	{
	Gui, DXP: Default
}

LVShowExplorer()                                                 	{                       	;-- öffnet den Exportpfad in einem Explorer-Fenster

	Pat := LV_GetSelectedRow()
	PatPath := exportordner "\(" Pat.NR ") " Pat.NAME ", " Pat.VORNAME "\"
	Gui, DXP: Default

	If InStr(FileExist(PatPath), "D") {
		SB_SetText("exportordner: " PatPath " geöffnet.", 1)
		Run, % "explorer.exe " q PatPath q
	} else
		SB_SetText("exportordner: " PatPath " ist nicht vorhanden", 1)

}
;}


;--- Dokumente auflisten        	;{
ZeigeDokumente(row)                                                                                	{

	;global DXP_FxF, DXP_PLV
		global DPGS_init, PatChoosed, DXP, DXP_DLV
		static outBezeichner

		Gui, DXP: Default

	; PATNR, NAME, VORNAME ermitteln
		Pat := LV_GetSelectedRow()
		If !RegExMatch(Pat.NR, "\d+")
			return

		PatChoosed := true, DPGS_init := true
		DimmerGui("on")

	; Patientendaten anzeigen
		GuiControl, DXP:, DXP_NR, % Pat.NR
		GuiControl, DXP:, DXP_NV, % Pat.NAME ", " Pat.VORNAME

	; aktuelle Dateifilter ermitteln
		For index, ext in StrSplit(docfilter_full, ",") {
			GuiControlGet, isChecked, DXP:, % "DXP_C" A_Index
			If (isChecked = 1)
				checkedext .= (A_Index=1?"":",") ext
		}

	; Dokument Listview aktivieren
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV
		LV_Delete()

	; out Filter lesen ;{
		Gui, DXP: Submit, NoHide
		outBezeichner := RTrim(RegExReplace(DXP_FF, "\s*,\s*", ","), ",")
		outBezeichner := RegExReplace(outBezeichner, "\*", ".*")
		outBezeichner := RegExReplace(outBezeichner, "?", ".")
		filters              := StrSplit(outBezeichner, ",")
		If (outBezeichner <> DXP_FF)
			GuiControl, DXP:, DXP_FF, % outBezeichner
	;}

	; Statusbar anzeige ändern
		sbtext := "Suche Dokumente von " Pat.NAME ", " Pat.VORNAME
		SB_SetText(sbtext, 1)
		SB_SetText("`t`t0", 2)

	; Dokumente aus den Datenbanken ermitteln
		If !IsObject(PatDok[Pat.NR])
			PatDok[Pat.NR] := DBDokumente(Pat.NR, checkedext)

	; anhand der Filter anzeigen
		Gui, DMR: Default
		GuiControl, DMR:,  DMR_SL1	, % "Dokumente auflisten..."
		GuiControl, DMR: , DMR_SL3	, % "/" (DokMax := PatDok[Pat.NR].Count())
		GuiControl, DMR: , DMR_DBFN	, % "der gefunden Dokumente sind gelistet"

		PatPath := exportordner "\(" Pat.NR ") " Pat.NAME ", " Pat.VORNAME
		lvdocNr := AllPages := exportedPages := 0
		For lfdnr, file in PatDok[Pat.NR] {

				DimmerPGS(lvdocNr ++, DokMax, StrLen(DokMax))
				Gui, DXP: Default
				Gui, DXP: ListView, DXP_DLV

				outmatched := false
				For filterIndex, out in filters
					If RegExMatch(file.Bezeichnung, "i)" out) {
						outmatched := true
						break
					}
				If outmatched
					continue

		; Daten zu den Dokumenten zusammenstellen
			RegExMatch(file.filename, "i)\.(?<xt>[a-z]+)$", e)
			exported    	:= searchable := ""
			pages       	:= 0
			datum 	    	:= SubStr(file.datum, 7, 2) "." SubStr(file.datum, 5, 2) "."SubStr(file.datum, 1, 4)
			docPath    	:= PatPath "\" file.datum " - " file.Bezeichnung "." file.ext
			originalPath 	:= AlbisPath "\Briefe\" file.filename
			docExist     	:= false

		  ; Exportstatus
			If FileExist(docPath) {
				exported	:= "Ja"
				docExist  	:= true
				viewPath 	:= docPath
			} else if FileExist(originalPath) {
				exported	:= "-"
				docExist 	:= true
				viewPath 	:= originalPath
			}

		  ; Dokument ist durchsuchbar und Seitenzahl
				If (docExist && RegExMatch(file.ext,"i)pdf")) {
					searchable 	:= PDFisSearchable(viewPath) ? "Ja" : ""
					pages       	:= PDFGetPages(viewPath, Addendum.qpdfPath)
				}
				else If docExist && RegExMatch(file.ext,"i)(txt|doc|docx|odt|xml)") {
					searchable	:= "Ja"
				}

		; Eintragen in die Listview
				Gui, DXP: Default
				Gui, DXP: ListView, DXP_DLV
				LV_Add((!exported ? "Check" : ""), datum, ext, file.Bezeichnung, exported, searchable, (pages = 0 ? "" : pages), LFDNR)

				AllPages += pages
				If (exported = "Ja")
					exportedPages += pages

		}

	; Spaltenbreite an den Inhalt anpassen
		If (LV_GetCount() > 0) || (PatDok[Pat.NR].Count() > 0) {
			LV_ModifyCol()
			LV_ModifyCol(4, "40")
			LV_ModifyCol(5, "40")
			LV_ModifyCol(6, "50")
			LV_ModifyCol(7, "70")
		}

	; neuen Statustest setzen
		DokumentInfo := (PatDok[Pat.NR].Count() > 0 ? PatDok[Pat.NR].Count() : "keine") " Dokumente"
		DokumentInfo .= (PatDok[Pat.NR].Count() > 0 ? " , Gesamtseitenzahl: " AllPages ", exportierte Dokumente:" exportedPages : "")
		SB_SetText(DokumentInfo)
		SB_SetText("", 2)
		SB_SetText("PatNR: " Pat.NR, 3)
		SB_SetText("Anzahl Patienten in PatDOK: " PatDok.Count(), 4)

	; Overlay-Gui ausschalten
		DimmerGui("off")

}

DBDokumente(PATNR, dokext="pdf,jpg,gif,png,wav,avi,mov,doc,docx") 	{

		global DPGS_init, DBASEFileName

		ftoex 	:= Array()
		dokext 	:= "(" . RegExReplace(dokext, "[,]", "|") . ")"

	;-------------------------------------------------------------------------------------------------------------------------------------------
	; neues Datenbankobjekt anlegen. new DBASE(Datenbankpfad [, debug])
	; Datenbankpfad	- vollständiger Pfad zur Datenbank in der Form D:\albiswin\db\BEFUND.dbf
	; debug                	- Ausgabe von Werten zur Kontrolle des Ablaufes, Voreinstellung: keine Ausgabe
	;-------------------------------------------------------------------------------------------------------------------------------------------
		SB_SetText("% der Datensätze der BEFTEXT.DBF durchsucht", 3)
		DPGS_init	:= 1
		DimmerGui("on", DBASEFileName)

	  ; Dateinamen
		beftexte   	:= new dBase(AlbisDBPath "\BEFTEXTE.dbf", false, {"name":"DXP", "control":"Statusbar", "sbNR":2})
		pattern  	:= {"PATNR": PATNR, "TEXT": "rx:\w+\\\w+\.[a-z]{1,4}"}
		res2        	:= beftexte.OpenDBF()
		matches2 	:= beftexte.Search(pattern, 0, "DimmerPGS")
		res2         	:= beftexte.CloseDBF()

		For i, m in matches2 {
			lfds .= (i > 1 ? "|" : "") m.LFDNR
			m.Text := Trim(RTrim(m.TEXT, "`r`n"))
			RegExMatch(m.TEXT, "i)\.(?<xt>[a-z]+)$", e)
			If RegExMatch(ext, "i)" dokext)
				ftoex[m.LFDNR] := {"filename":m.TEXT, "datum":m.DATUM, "ext":ext}
		}

		SB_SetText("% der Datensätze der BEFUND.DBF durchsucht", 3)
		DPGS_init	:= 1
		DimmerGui("on", DBASEFileName)

	  ; Bezeichnungstext in der Karteikarte
		befund   	:= new dBase(AlbisDBPath "\BEFUND.dbf", 0, {"name":"DXP", "control":"Statusbar", "sbNR":2})
		pattern		:= {"PATNR": PATNR, "TEXTDB": ("rx:(" RTrim(lfds, "|") ")")}
		res1        	:= befund.OpenDBF()
		matches1 	:= befund.Search(pattern, 0, "DimmerPGS")
		res1         	:= befund.CloseDBF()

		Loop, 127
			ws .= " "

	  ; Sonderzeichen entfernen für Dateinamenanpassung
		For i, m in matches1 {

			INHALT   	:=  Trim(m.INHALT)
			If (StrLen(INHALT) > 0) {
				INHALT 	:=  RegExReplace(INHALT " ", "i)[\s\-][A-Z]{2}\s", " ")
				INHALT 	:=  RegExReplace(INHALT " ", "i)\s[A-Z]\s", " ")
				INHALT 	:=  RegExReplace(INHALT, "[,._;:§%&\!\-\?\$\{\}\[\]\(\))\\\/\s]+$")
				INHALT 	:=  RegExReplace(INHALT, "[~" Chr(0x22) "#%&\*<>\?\\\{\|\}]", " ")
				INHALT 	:=  RegExReplace(INHALT, "[&\:\/]", "-")
				INHALT 	:=  RegExReplace(INHALT, "\s{2,}", " ")
				INHALT 	:=  RegExReplace(INHALT, "([a-z])(\d)", "$1 $2")
			}
			; unbenannte Dokumente als unbenannt kennzeichnen
			else
				INHALT := "unbenannt"


			; doppelte Dateinamen finden und aktuellen Dateinamen mittels Zähler eindeutig machen
				filecount := 0
				For tmpLFDNR, file in ftoex {
					RegExMatch(file.Bezeichnung, "(?<Bezeichnung>.+?)(\-\s*(?<counter>\d*)\s*\.|$)", file)
					If (INHALT = fileBezeichnung) && (m.datum = file.datum)
						filecount := filecounter ? filecounter+1 : 1
				}
				INHALT .= (filecount ? " - " filecount : INHALT = "unbenannt" ? "- 1" : "")
				;SciTEOutput(SubStr("00" i, -1) ": " INHALT)

			ftoex[m.TEXTDB].Bezeichnung :=  INHALT	; TEXTDB = LFDNR

		}

return ftoex
}


;}

;--- Dokumente exportieren    	;{
LVExport() {                                                                                                           	;-- exportiert ausgewählte Dokumente

		Pat        	:= LV_GetSelectedRow()
		dokument := Array()

	; Listview default machen
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV

	; zu exportierende Dateien zusammenstellen
		row := 0
		while (row := LV_GetNext(row, "C")) {
			LV_GetText(LFDNR, row, 7)
			dokument[A_Index] := PatDok[Pat.NR][LFDNR]
		}

	; Exportvorgang ausführen
		If (dokument.Count() = 0) {
			SB_SetText("KEINE EXPORTAUSWAHL GETROFFEN!", 1)
			return 0
		}

		filesLeft :=Dokumentexport(Pat.NR, Pat.NAME ", " Pat.VORNAME, dokument, exportordner)
		PraxTT("Dateien exportiert. Rest " filesLeft, "3 0")


return filesLeft = 0 ? 1 : -1*filesLeft
}

Dokumentexport(PATNR, PATName, dokument, exportPath) {                                     	;-- kopiert die ausgewählten Dokumente aus albiswin/briefe in den Exportordner

		global AlbisPath, exportordner

	; Dokumentenexportpfad anlegen
		PatPath := CreatePatPath(PATNR, PATName, exportPath)
		If (StrLen(PatPath) = 0) {
			PraxTT(A_ScriptName "`n" A_ThisFunc "`nDer Patientenexportpfad konnte nicht angelegt werden", "3 1")
			return -99
		}

	; Gui Default
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV

	; Dateien kopieren
		For row, file in dokument {
			exportFile := PatPath "\" file.datum " - " file.Bezeichnung "." file.ext
			If FileExist(exportFile)
				continue
			FileCopy, % AlbisPath "\Briefe\" file.filename, % exportFile, 1
			SB_SetText(AlbisPath "\Briefe\" file.filename, 1)
			SB_SetText(PatPath "\" file.datum " - " file.Bezeichnung "." file.ext, 4)
			SB_SetText(ErrorLevel, 2)
			If !ErrorLevel {
				copynr ++
				SB_SetText(copynr " von " dokument.Count() " Dateien exportiert.")
				LV_Modify(row, "-Check")
			}
		}

return dokument.Count() - copynr
}

KarteikartenExport() {                                                                                               	;-- Tagesprotokoll und Laborblatt werden exportiert

		global DXP_VB, DXP_KKF

		Gui, DXP: Default
		Gui, DXP: ListView, DXP_PLV
		Gui, DXP: Submit, NoHide

	; Patienten Nummer bekommen
		Pat := LV_GetSelectedRow()
		If !RegExMatch(Pat.NR, "^\d+$")
			return

	; Dokumentenexportpfad anlegen bei Bedarf
		If !(PatPath := CreatePatPath(Pat.NR, Pat.NAME ", " Pat.VORNAME, exportordner)) {
			PraxTT(A_ScriptName "`n" A_ThisFunc "`nDer Patientenexportpfad konnte nicht angelegt werden", "3 1")
			return 0
		}
		KKFilePath	:= PatPath "\Karteikartenauszug von " Pat.NAME ", " Pat.VORNAME

	; Karteikarte des Patienten öffnen
		AlbisAkteOeffnen(Pat.NAME ", " Pat.VORNAME, Pat.NR)

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Laborblatt drucken
	;----------------------------------------------------------------------------------------------------------------------------------------------
		; Werte exportieren
			If !FileExist(PatPath "\Laborkarte von " Pat.NAME ", " Pat.VORNAME) {
				PraxTT("Laborwerte werden exportiert.", "20 1")
				LBlattRes := LaborblattExport(Pat.NAME ", " Pat.VORNAME, PatPath)
			}
			else {
				PraxTT("Laborwerte wurden bereits exportiert.", "3 1")
				Sleep 3000
			}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; TAGESPROTOKOLL EXPORTIEREN (Karteikarte!)
	;----------------------------------------------------------------------------------------------------------------------------------------------
		; Tagesprotokoll für Patient erstellen
			PraxTT("Karteikarteneinträge werden exportiert.", "20 1")
			options := {	"Periode"	: DXP_VB
							, 	"Patienten"	: "aktiver_Patient"
							, 	"Kuerzel" 	: DXP_KKF
							, 	"Cave"   	: false}
							SciTEOutput("Kürzel: " DXP_KKF)
			TProtRes := AlbisErstelleTagesprotokoll(options,, false)
			If (TProtRes := !RegExMatch(TProtRes, "i)[A-Z]\:\\") ? 1 : 0) {

				; Tagesprotokoll aktivieren
					AlbisMDIChildActivate("Tagesprotokoll")

				; Protokoll über Drucken speichen
					If FileExist(KKFilepath ".pdf")
						FileDelete, % KKFilepath ".pdf"
					TProtRes := AlbisSaveAsPDF(KKFilePath)
					If (TProtRes <> 1)
						return 0

				; erstelltes Tagesprotokoll schliessen
					AlbisMDIChildWindowClose("Tagesprotokoll")
					PraxTT("", "Off")

			}

return  {"Patient":Pat.NAME ", " Pat.VORNAME " (" Pat.NR ")", "TP": TProtRes, "LB": LBlattRes}
}

LaborblattExport(name, PatPath, Printer="", Spalten="Alles")                 	{                	;-- automatisiert den Laborblattexport

		static AlbisViewType

		AlbisViewType	:= AlbisGetActiveWindowType(true)
		savePath       	:= PatPath "\Laborkarte von " name
		Printer            	:= Printer ? Printer : "Microsoft Print to PDF"
		Spalten         	:= Spalten ? Spalten : "Alles"

		If FileExist(savePath ".pdf")
			FileDelete, % savePath ".pdf"

	; Aufruf der Automatisierungsfunktion
		res := AlbisLaborblattExport(Spalten, savePath, Printer)
		If (res = 0) || (res > 1) {
			PraxTT("Der Export der Laborwerte ist fehlgeschlagen.`nFehlercode: " res, "4 1")
			return 0
		}

	; ursprüngliche Ansicht wiederherstellen
		If (AlbisViewType <> AlbisGetActiveWindowType(true))
			AlbisKarteikartenAnsicht(AlbisViewType)

}

CreatePatPath(PATNR, PATName, exportPath:="") {                                                   	;-- legt den Patientenexportpfad an

	If !exportPath
		exportPath := ExportOrdner

	If !InStr(FileExist(PatPath := exportPath "\(" PATNR ") " PATName), "D") {
		FileCreateDir, % PatPath "\"
		If ErrorLevel
			return PatPath
		return
	}

return PatPath
}


;}

;--- Patientensuche                 	;{
DXPSearch(callfrom) {                                                 	;-- für Interskript-Kommunikation (Fernsteuerung,Abfrage durch andere Skripte)

		global DXP, DXP_Pat, DXP_DLV, DXP_PLV

	; Aufruf durch HOTKEY
		If (callfrom = "Hotkey") {
			GuiControlGet, fc, DXP: FocusV
			If (fc <> "DXP_Pat")
				return
		}

	; DXP_Pat Eingabe ermitteln
		Gui, DXP: Submit, NoHide
		If (StrLen(DXP_Pat) = 0)
			return

	; Fortschrittsanzeige anschalten
		DimmerGui("on", "PATIENT.dbf")

	; die Combobox um einen Eintrag erweitern (bei Übergabe eine PatID allerdings hier noch nicht)
		If !(callfrom = "MessageWorker") && !RegExMatch(DXP_Pat, "^\d+$")
			AddNewComboxItem(DXP_Pat, DXPhPat)

	; neue Suche - Objekte leeren
		PreVID	:= Object()

	; beide Listen leeren
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV
		LV_Delete()
		Gui, DXP: ListView, DXP_PLV
		LV_Delete()

	; passenden Patienten suchen
		SB_SetText("...Suche nach Patienten")
		PatMatches := DBASEGetPatID(DXP_Pat, 0)
		Gui, DXP: Default
		SB_SetText( (PatMatches.Count() > 0 ? PatMatches.Count() " Patienten gefunden" : "nichts gefunden") )

	; Patienten auflisten
		LV_ModifyCol(11, "0 Integer")
		For i, m in PatMatches {
			GEBURT	:= ConvertDBASEDate(m.GEBURT)
			LASTBEH	:= ConvertDBASEDate(m.LAST_BEH)
			Mortal   	:= m.Mortal ? "♱" : ""
			LV_Add("", m.NR, m.NAME ,m.VORNAME, GEBURT, m.PLZ, m.ORT ,m.STRASSE " " m.HAUSNUMMER, m.TELEFON, m.ARBEIT, LASTBEH, m.GEBURT)
		}

	; Spaltenbreite anpassen
		If (PatMatches.Count() > 0) {
			LV_ModifyCol()
			LV_ModifyCol(4, "Integer Left")   	; Geburtstag
			LV_ModifyCol(5, "Integer Right") 	; PLZ
			LV_ModifyCol(6, "108")              	; Ort
			LV_ModifyCol(11, "0 Integer")
		}

	; Combobox anstatt eine ID den Patientennamen anfügen
		If RegExMatch(DXP_Pat, "^\d+$") && PatMatches.Count() = 1
			AddNewComboxItem(DXP_Pat, DXPhPat)

	; Zahl der gefunden Patienten anzeigen
		GuiControl, DXP:, DXP_PZ, % PatMatches.Count()

	DimmerGui("off")

return
}

DBASEGetPatID(searchstr, debug=false) {                     	;-- sucht in der PATIENT.DBF

	static DBFPatient

	searchstr := Trim(searchstr)

	; Öffnen der Datenbankdatei
		If !IsObject(DBFPatient) {
			DBFPatient := new dBase(AlbisDBPath "\PATIENT.dbf", debug)
			If !IsObject(DBFPatient) {
				MsgBox, 1, % "Addendum für Albis on Windows", % "Es gab ein Problem beim Öffnen der PATIENT.dbf Datenbanl.`nDas Skript wird beendet."
				ExitApp
			}
		}

	;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	; Parsen des Suchstrings
	;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~;{

	;-----------------------------------------------------------------------------------------------
	; Geburtstag
	;-----------------------------------------------------------------------------------------------
		If RegExMatch(searchstr . " ", "(?<Tag>[*?\d]{1,2})\.(?<Monat>[*?\d]{1,2})\.(?<Jahr>\*|[*?\d]{4}|[*?\d]{2})", s) {

			sJahr	:= RegExReplace(sJahr	 , "[?*]", ".")
			sMonat	:= RegExReplace(sMonat, "[?*]", ".")
			sTag  	:= RegExReplace(sTag	 , "[?*]", ".")
			lenJahr	:= StrLen(sJahr)
			TNow 	:= SubStr(A_YYYY, 1, 1)
			HNow 	:= SubStr(A_YYYY, 2, 1)
			JHNow := SubStr(A_YYYY, 1, 2)
			JNow 	:= SubStr(A_YYYY, 3, 2)

			If (sJahr = ".")        	{
				Jahr := "...."
			}
			else If (lenJahr = 2)  	{

				if RegExMatch(sJahr, "^\d{2}$")
					Jahr := (sJahr > JNow) ? (HNow - 1) sJahr : Jahr := ".." sJahr
				else
					Jahr := ".." sJahr

			}
			else if (lenJahr = 4) 	{

				 If RegExMatch(sJahr, "^\d\d", Jahrhundert) {
					If (Jahrhundert > JHNow)
						sJahr := JHNow SubStr(sJahr, 3, 2)
				}
				else if RegExMatch(sJahr, "^\d", Tausender) {
					If (Tausender = 0) || (Tausender > TNow)
						sJahr := "." SubStr(sJahr, 2, 3)
				}

				Jahr := sJahr

			}

			If RegExMatch(sMonat, "^\.$")
				Monat := ".."
			else If RegExMatch(sMonat, "^\d{1}$")
				Monat := (sMonat = 0) ? "10" : "0" sMonat
			else if RegExMatch(sMonat, "^\d{2}$")
				Monat := (sMonat > 12 ? "12" : sMonat = 0 ? "01" : sMonat)
			else If RegExMatch(sMonat, "^(\d)\.$", n)
				Monat := n1 > 1 ? "1." : sMonat
			else If RegExMatch(sMonat, "^\.(\d)$", n)
				Monat := sMonat
			else
				Monat := sMonat

			If RegExMatch(sTag, "^\.$")
				Tag := ".."
			else If RegExMatch(sTag, "^\d{1}$")
				Tag := (sTag = 0) ? "01" : "0" sTag
			else if RegExMatch(sTag, "^\d{2}$")
				Tag := sTag > 31 ? "31" : sTag = 0 ? "01" : sTag
			else If RegExMatch(sTag, "^(\d)\.$", n)
				Tag := n1 > 3 ? "3." : sTag
			else If RegExMatch(sTag, "^\.(\d)$", n)
				Tag := sTag
			else
				Tag := sTag

		  ; DBASE Datumsformat YYYYMMDD
			searchstr := Jahr Monat Tag
			spattern  := {"GEBURT": searchstr}
			;SB_SetText(searchstr "|" sJahr " - " sMonat " - " sTag )

		}

	;-----------------------------------------------------------------------------------------------
	; Nachname (Wort mit Bindestrich, auch zwei Worte mit Leerzeichen)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^[a-zßäöü\p{L}\-]+\s*,*$")
			spattern := {"NAME" : ("rx:" searchstr ".*?[^\d]")}         ; findet sonst auch Straßennamen

	;-----------------------------------------------------------------------------------------------
	; Nachname, Vorname (wie oben beschrieben)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^(?<nn>[a-zßäöü\p{L}\-]+)\s*,\s*(?<vn>[a-zßäöü\p{L}\-]+)$", s)
			spattern := {"NAME" : ("rx:.*?" snn), "VORNAME": ("rx:.*?" svn)}

	;-----------------------------------------------------------------------------------------------
	; Vorname (wird erkannt durch ein Komma vor dem Wort)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^,\s*(?<vn>[a-zßäöü\p{L}\s\-]+)$", s)
			spattern := {"VORNAME" : ("rx:.*?" svn)}

	;-----------------------------------------------------------------------------------------------
	; Straße Hausnummer (anstatt einer Nummer ein Sternchen verwenden)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^(?<streetn>[a-zßäöü-]+)\s(?<streetnr>\*|[\da-z]+)$", s) {
			If !InStr(sstreetnr, "*")
				spattern := {"STRASSE" : ("rx:.*?" sstreetn), "HAUSNUMMER" : ("rx:.*?" sstreetnr)}
			else
				spattern := {"STRASSE" : ("rx:.*?" sstreetn)}
		}

	;-----------------------------------------------------------------------------------------------
	; Patienten Nr (nur Zahleneingabe)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^\d+$")
			spattern := {"NR" : searchstr}

	;-----------------------------------------------------------------------------------------------
	; Arbeit (A:)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^A\s*\:\s*(?<Arbeit>.*)$", s)
			spattern := {"ARBEIT" : ("rx:.*?" sArbeit)}

	;-----------------------------------------------------------------------------------------------
	; Ort (O:)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^O\s*\:\s*(?<Ort>.*)$", s)
			spattern := {"Ort" : ("rx:.*?" sOrt)}
		else {
			SB_SetText("kein passender Suchstring!", 4)
			return
		}
	;}

	; Suche starten
		filepos1  	:= DBFPatient.OpenDBF()
		matches 	:= DBFPatient.Search(spattern,, "DimmerPGS")
		filepos2  	:= DBFPatient.CloseDBF()

		DBFPatient := ""

return matches
}
;}

;--- Extra Funktion                 	;{
Antragzaehler() {

	antragfilter:= ["i)lageso", "^Au\s|$", "Anfrage", "Antrag", ",ausgefüllt", "MDK Anfrage", "Bericht medizinisch.*Dienst", "Auszahlschein", "Bericht.*lebens.*versich", "Bescheinigung.*Versicherung"
						, "Befundbericht.*REHA", "Antrag.*Rente", "Erinnerung.*DRV", "RLV.*Anfrage", "REHA.*Antrag", "Landesamt.*Versorgung", "Landesamt.*Sozial", "Ärztliches Attest", "zahlschein"
						, "ausgefüllt.*Rentenver*", "Antwortschreiben", "Auszahlschein", "Med.*Dienst", "Dt\s", "ausgefüllt.*Schreib[en]+", "Rücksendung", "Antwortschreiben", "Rückantwort",
						, "Deutsch.*Versich", "Bundesagentur", "LASV", "Folge\s", "Zahlung.*Krankengeld", "Befundbericht"]


}

;}

;--- Gui Funktionen                	;{
AddNewComboxItem(newItem, cbhwnd, maxItems=50) {

		global DXP, DXP_Pat

		Gui, DXP: Default
		ControlGet, cbList, List,,, % "ahk_id " cbhwnd
		cbitems := StrSplit(cbList, "`n")

	; removes same items
		itemfound := false
		For itempos, item in cbitems
			If RegExMatch(item, "i)^" newitem) {
				cbitems.RemoveAt(itempos)
				;~ itemfound := true
				break
			}

	; add new item to list,
		cbitems.InsertAt(1, newitem)									; add item at beginning
		If (cbitems.Length() > maxItems)							; remove items from list
			cbitems.Pop()

	; builds a new combobox list from array
		cbList := ""
		For itempos, item in cbitems
			If (StrLen(item) > 0)
				cbList .= (itempos=1 ? "":"|") . item

	; set new content
		Gui, DXP: Default
		SendMessage, 0x14B,,,, % "ahk_id " cbhwnd     	; CB_RESETCONTENT
		GuiControl, DXP:, DXP_Pat, % cbList
		GuiControl, DXP: ChooseString , DXP_Pat, % newitem

}

ChangeFolder() {

	newFolder := SelectFolder(0, exportordner)
	If (StrLen(newFolder) = 0) && (newFolder = PropSaveDir)
		return

	GuiControl, DXP:, DXP_PE, % (exportordner := newfolder)
	IniWrite, % exportordner, % A_ScriptDir "\admExporter.ini", DokExporter, exportordner

}

CB_GetEntries(hwnd) {
	ControlGet, cbList, List,,, % "ahk_id " hwnd
return StrReplace(cbList, "`n", "|")
}

LV_ChooseAll() {

	Gui, DXP: ListView, DXP_DLV
	Loop, % LV_GetCount()
		LV_Modify(A_Index, "Check")

	GuiControl, DXP: Enable, DXP_EA
	GuiControl, DXP:, DXP_EA, % LV_GetCount() " Dokumente exportieren"

}

LV_ChooseNone() {

	Gui, DXP: ListView, DXP_DLV
	Loop, % LV_GetCount()
		LV_Modify(A_Index, "-Check")

	GuiControl, DXP: Disable, DXP_EA
	GuiControl, DXP:, DXP_EA, % "Dokumente exportieren"

}

LV_ChooseLastYears(yearsback:=2) {

	Gui, DXP: ListView, DXP_DLV
	fromYear    	:= A_YYYY - yearsback
	iamchecked 	:= 0
	Loop, % LV_GetCount() {
		LV_GetText(LVDate, A_Index, 1)
		RegExMatch(LVDate, "\d{4}", LVyear)
		If (LVyear >= fromYear) {
			LV_Modify(A_Index, "Check")
			iamchecked ++
		}
	}

	GuiControl, DXP: Enable, DXP_EA
	GuiControl, DXP:, DXP_EA, % iamchecked " Dokument" (iamchecked>1 ? "e": "") " exportieren"

}

LV_GetSelectedRow(LVName:="DXP_PLV") {

	global DXP

	Gui, DXP: Default
	Gui, DXP: ListView, % LVName
	frow := 0
	Loop {
		frow := LV_GetNext(frow, "F")
		LV_GetText(PATNR     	, frow, 1)
		LV_GetText(NAME      	, frow, 2)
		LV_GetText(VORNAME	, frow, 3)
		break
	}

return {"NR":PATNR, "NAME":NAME , "VORNAME":VORNAME}
}

LV_CountCheckedDocs() {

		global DXP, DXP_LV, DXP_EA

		Gui, DXP: ListView, DXP_DLV

	; abgehakte Einträge zählen
		row := checkedRows := 0
		Loop % LV_GetCount() {
			If !(row := LV_GetNext(row, "Checked"))
				break
			checkedRows ++
		}

		If (checkedRows > 0)
			GuiControl, DXP: Enable, DXP_EA
		else
			GuiControl, DXP: Disable, DXP_EA

		GuiControl, DXP:, DXP_EA, % (checkedRows = 0 ? "" : checkedRows) " Dokument" (checkedRows=1 ? "": "e") " exportieren"

}

;}

;--- Dateien anzeigen              	;{
EndPreviewers(PATNR) {  ; maybe sometimes ???

	If IsObject(PREVID) {

	For PNR, LFDNR in PREVID {

		If (PNR = PATNR)
			continue

		Process, Exist, % LFDNR.PID
		If !ErrorLevel
			PreVID[PNR].Delete(LFDNR)
		else {

			ControlGet,Customcaption,  hwnd,, % "ahk_Id " LFDNR.ID


		}



	}


	}


}

DokumentVorschau(row) {

	Gui, DXP: Default

	; PATNR ermitteln
		GuiControlGet, PATNR, DXP:, DXP_NR
		If !RegExMatch(PATNR, "\d+")
			return

	; LFDNR aus der Dokumentlistview holen
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV
		LV_GetText(ext   	, row	, 2)
		LV_GetText(LFDNR, row	, 5)
		If !RegExMatch(LFDNR, "\d+") || (StrLen(ext) = 0)
			return

	; Prüfen ob das Dokument vorhanden ist
		If FileExist(AlbisPath "\Briefe\" PATDok[PATNR][LFDNR].filename)
			SB_SetText("zeige Dokument: " LFDNR " - " PATDok[PATNR][LFDNR].filename, 1)
		else {
			SB_SetText("Dokument nicht vorhanden: " LFDNR " - " PATDok[PATNR][LFDNR].filename, 1)
			return
		}

	; Fensterpositionen / Monitogrößen ermitteln
		sDXP   	:= GetWindowSpot(hDXP)
		hMon 	:= MonitorFromWindow(hDXP)
		Mon   	:= GetMonitorInfo(hMon)

	; Dokumenten mit Standard-Anzeigeprogramm aufrufen oder wenn pdf dann mit SumatraPDF falls vorhanden
		file := PATDok[PATNR][LFDNR]
		;RegExMatch(file.filename, "i)\.(?<xt>[a-z]{3})$", e)
		If (ext = "pdf") && InStr(pdfViewer, "Sumatra") {

			; eigene Gui wird versetzt
				DllCall("SetWindowPos"	, "Ptr"	, hDXP
													, "Ptr"	, 0
													, "Int" 	, (Mon.WR - sDXP.W - 5)   	; x
													, "Int" 	, sDXP.Y                            ; y
													, "Int" 	, sDXP.W                        	; w
													, "Int" 	, sDXP.h							; h
													, "UInt"	, 0x40)								; SWP_SHOWWINDOW:= 0x40

			; SumataPDF wird aufgerufen
				smtra := Sumatra_Show(AlbisPath "\Briefe\" file.filename)
				If !IsObject(PreVID[PATNR])
						PreVID[PATNR] := Object()
				else if !IsObject(PreVID[PATNR][LFDNR])
						PreVID[PATNR][LFDNR] := {"exe":"Sumatra", "id": sumatra.id, "PID": SumatraPID, "file":file.filename}

			; Größenberechnungen
				mwidth 	:= Mon.WR - sDXP.W - 10
				mheigth	:= Mon.WB
				sw := Round(mheight * smtra.AR) > mwidht ? mwidth : Round(mheight * smtra.AR)
				sh := sw = mwidth ? Round(sw / smtra.AR)

			; Positionsberechung
				sx := sw < mwidth ? (Mon.WR - sDXP.W - 10 - sw) ? 0
				sy := sh = mheight ? 0 : Round(mheight/2) - Round(sh/2)

			; SumataPDF verschieben und Inhalt anpassen
				DllCall("SetWindowPos"	, "Ptr", smtra.ID, "Ptr", 0, "Int", 0, "Int", 0, "Int", mwidth, "Int", mheigth, "UInt", 0x40)	; SWP_SHOWWINDOW:= 0x40
				Sleep 300
				SumatraInvoke("SinglePage", smtra.ID)
				Sleep 300
				SumatraInvoke("FitPage", smtra.ID)

		}
		else {

			; starten mit Standard-Anzeigeprogramm
				Run, % AlbisPath "\Briefe\" file.filename

		}



return
}


;}

;--- Interskriptkommunikation 	;{
MessageWorker(InComing) {                                                                                    	;-- verarbeitet die eingegangen Nachrichten

		global DXP, DXPhPat, DXP_Pat

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}


	; Kommunikation vom aufrufenden Skript - sendet z.B. die PatNR
		if InStr(recv.txtmsg, "Search") {
			Gui, DXP: Default
			GuiControl, DXP: Focus, DXP_Pat
			AddNewComboxItem(recv.opt, DXPhPat, 50)
			DXP_Pat := recv.opt
			DXPSearch(A_ThisFunc)
		}

return
}
;}

;--- Dateidialog Funktionen    	;{
IFileDialogEvents_new(){
	vtbl := IFileDialogEvents_Vtbl()
	fde := DllCall("GlobalAlloc", "UInt", 0x0000, "Ptr", A_PtrSize + 4, "Ptr")
	if (!fde)
		return 0

	NumPut(vtbl, fde+0,, "Ptr")
	NumPut(1, fde+0, A_PtrSize, "UInt")

	return fde
}

IFileDialogEvents_Vtbl(ByRef vtblSize := 0){
	static vtable
	if (!VarSetCapacity(vtable)) {

		extfuncs := ["QueryInterface", "AddRef", "Release", "OnFileOk", "OnFolderChanging", "OnFolderChange", "OnSelectionChange", "OnShareViolation", "OnTypeChange", "OnOverwrite"]

		VarSetCapacity(vtable, extfuncs.Length() * A_PtrSize)

		for i, name in extfuncs
			NumPut(RegisterCallback("IFileDialogEvents_" . name), vtable, (i-1) * A_PtrSize)
	}
	if (IsByRef(vtblSize))
		vtblSize := VarSetCapacity(vtable)
	return &vtable
}

IFileDialogEvents_QueryInterface(this_, riid, ppvObject){                                         	; Called on a "ComObjQuery"
	static IID_IUnknown, IID_IFileDialogEvents
	if (!VarSetCapacity(IID_IUnknown))
		VarSetCapacity(IID_IUnknown, 16), VarSetCapacity(IID_IFileDialogEvents, 16)
		,DllCall("ole32\CLSIDFromString", "WStr", "{00000000-0000-0000-C000-000000000046}", "Ptr", &IID_IUnknown)
		,DllCall("ole32\CLSIDFromString", "WStr", "{973510db-7d7f-452b-8975-74a85828d354}", "Ptr", &IID_IFileDialogEvents)

	if (DllCall("ole32\IsEqualGUID", "Ptr", riid, "Ptr", &IID_IFileDialogEvents) || DllCall("ole32\IsEqualGUID", "Ptr", riid, "Ptr", &IID_IUnknown)) {
		NumPut(this_, ppvObject+0, "Ptr")
		IFileDialogEvents_AddRef(this_)
		return 0
	}

	; Else
	NumPut(0, ppvObject+0, "Ptr")
	return 0x80004002
}

IFileDialogEvents_AddRef(this_){                                                                             	; Called on an "ObjAddRef"
	NumPut((_refCount := NumGet(this_+0, A_PtrSize, "UInt") + 1), this_+0, A_PtrSize, "UInt")
	return _refCount
}

IFileDialogEvents_Release(this_) {                                                                              	; Called on an "ObjRelease"
	_refCount := NumGet(this_+0, A_PtrSize, "UInt")
	if (_refCount > 0) {
		_refCount -= 1
		NumPut(_refCount, this_+0, A_PtrSize, "UInt")
		if (_refCount == 0)
			DllCall("GlobalFree", "Ptr", this_, "Ptr")
	}
	return _refCount
}

IFileDialogEvents_OnFileOk(this_, pfd){
	return 0x80004001
}

IFileDialogEvents_OnFolderChanging(this_, pfd, psiFolder){
	return 0x80004001
}

IFileDialogEvents_OnFolderChange(this_, pfd){
	return 0x80004001
}

IFileDialogEvents_OnSelectionChange(this_, pfd){
	if (DllCall(NumGet(NumGet(pfd+0)+14*A_PtrSize), "Ptr", pfd, "Ptr*", psi) >= 0) {
         GetDisplayName := NumGet(NumGet(psi + 0, "UPtr"), A_PtrSize * 5, "UPtr")
         If !DllCall(GetDisplayName, "Ptr", psi, "UInt", 0x80028000, "PtrP", StrPtr) {
            SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
			ToolTip % SelectedFolder
		 }
		ObjRelease(psi)
	}
	return 0
}

IFileDialogEvents_OnShareViolation(this_, pfd, psi, pResponse){
	return 0x80004001
}

IFileDialogEvents_OnTypeChange(this_, pfd){
	return 0x80004001
}

IFileDialogEvents_OnOverwrite(this_, pfd, psi, pResponse){
	return 0x80004001
}

SelectFolder(fde:=0, initFolder:="") {

   Static OsVersion 	:= DllCall("GetVersion", "UChar")
   Static Show          	:= A_PtrSize * 3
   Static SetOptions	:= A_PtrSize * 9
   Static GetResult    	:= A_PtrSize * 20
   SelectedFolder    	:= initFolder

   If (OsVersion < 6) {
      FileSelectFolder, SelectedFolder
      Return SelectedFolder
   }

   If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
      Return ""
   VTBL := NumGet(FileDialog + 0, "UPtr")
   DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", 0x00000028, "UInt")

	if (fde) {
		DllCall(NumGet(NumGet(FileDialog+0)+7*A_PtrSize), "Ptr", FileDialog, "Ptr", fde, "UInt*", dwCookie := 0)
	}

	showSucceeded := DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", 0) >= 0

	if (dwCookie)
		DllCall(NumGet(NumGet(FileDialog+0)+8*A_PtrSize), "Ptr", FileDialog, "UInt", dwCookie)

   If (showSucceeded) {
	   If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt") {
         GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
         If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr)
            SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
         ObjRelease(ShellItem)
      }
   }

   ObjRelease(FileDialog)

Return SelectedFolder
}

SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {                               	;-- modified version for Addendum für Albis on Windows

	; last change 17.08.2020

	; some variables
		static	LinesOut           	:= 0
			, 	SCI_GETLENGTH	:= 2006
			,	SCI_GOTOPOS		:= 2025

	; gets Scite COM object
		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

	; move Caret to end of output pane to prevent inserting text at random positions
		SendMessage, 2006,,, Scintilla2, % "ahk_id " SciObj.SciteHandle
		endPos := ErrorLevel
		SendMessage, 2025, % endPos,, Scintilla2, % "ahk_id " SciObj.SciteHandle

	; shows count of printed lines in case output pane was erased
		If InStr(Text, "ShowLinesOut") {
			;SciObj.Output("SciteOutput function has printed " LinesOut " lines.`n")
			return
		}

	; Clear output window
		If (Clear=1) || ((StrLen(Text) = 0) && (LinesOut = 0))
			SendMessage, SciObj.Message(0x111, 420)
		;~ else if (Clear = "CL") {
			;~ ControlSend, Scintilla2, {Left}, % "ahk_id " SciObj.SciteHandle
			;~ ControlSend, Scintilla2, {LShift Down}{Home}{LShift Up}, % "ahk_id " SciObj.SciteHandle
		;~ }


	; send text to SciTE output pane
		If (StrLen(Text) != 0) {
			Text .= (LineBreak ? "`r`n": "")
			SciObj.Output(Text)
			LinesOut += StrSplit(Text, "`n", "`r").MaxIndex()
		}

		If Exit {
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

}
;}

;---- includes                           	;{
#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk

#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DBASE.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DATUM.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PDFReader.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PraxTT.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
;}

;---- Hades                            	;{

	;~ items := StrSplit(cbList, "`n")
	;~ For i, item in items
		;~ If (StrLen(item) > 0)
			;~ cbitems .= item (items.Count() = i ? "" : "|")

;return cbitems

		; add new item to list,
		;~ If !itemfound {
		;~ cbitems.InsertAt(1, newitem)									; add item at beginning
		;~ If (cbitems.Length() > maxItems)							; remove items from list
			;~ cbitems.Pop()
		;~ }
			;cbitems.Delete(maxItems+1, cbitems.Length())

;}

^!#::Reload
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 		DocPrinz
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
;                        				by Ixiko started in September 2017 - last change 15.05.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;	Version 1.01

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

		global q := Chr(0x22)
		global DBFPatient, AlbisPath, AlbisDBPath, exportordner, AddendumDir, suchfeldeintragung, qpdfPath
		global DXP, hDXP, DXP_Pat, DXPhPat, DXP_PLV, DXP_DLV, DXP_PE, DXP_SB, DXP_FF, DXP_AB, DXP_EB, DXP_FB, DXP_NR, DXP_NV
		global FSize, edWidth
		global docfilter_full, docfilter_checked, pdfViewer, imgViewer, docViewer
		global PatDok         	:= Object()
		global PREVID         	:= Object()
		global Addendum 	:= Object()
		global PatChoosed 	:= false
		global SumatraCMD, SumatraExist

		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)
		Addendum.qpdfPath    	:= AddendumDir "\include\OCR\qpdf"
		Addendum.AlbisDBPath	:= GetAlbisPath() "\db"

	; Sumatra für Vorschau und den Druck von PDF Dokumenten
		SumatraCMD := GetAppImagePath("SumatraPDF")
		If (StrLen(SumatraCMD) > 0) && FileExist(SumatraCMD)
			SumatraExist := true

	  ; Tray Icon erstellen
		If (hIcon := DocPrinz_ico())
			Menu, Tray, Icon, % "hIcon: " hIcon
		;~ else
			;~ Menu, Tray,  Icon, % AddendumDir "\assets\ModulIcons\DocPrinz.ico"

	;}

	; Albis Basispfad                                            	;{
		SetRegView, % (A_PtrSize = 8 ? 64 : 32)
		RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
		AlbisDBPath := AlbisPath "\DB"
	;}

	; INI Read                                                     	;{
		global scriptini := A_ScriptDir "\" StrReplace(A_ScriptName, ".ahk") ".ini"

		IniRead, docfilter_full          	, % scriptini, Exporter, docfilter_full            	, % "pdf,doc,docx,rtf,txt,jpg,bmp,gif,png,wav,avi,mov"

		IniRead, docfilter_checked  	, % scriptini, Exporter, docfilter_checked    	, % "pdf,doc,docx,rtf,txt,jpg,bmp"
		If InStr(docfilter_checkd, "ERROR")
			docfilter_checkd := ""

		IniRead, docfilter_bezeichner	, % scriptini, Exporter, docfilter_bezeichner	, % "Verordnung,Einnahme,Rezept,Fax,Medikamenten"
																														. ",chron. krank, Notfall-Vertretungsschein"

		IniRead, suchfeldeintragung  	, % scriptini, Exporter, suchfeldeintragung
		If InStr(suchfeldeintragung, "ERROR")
			suchfeldeintragung := ""

		IniRead, ini_yearsback       	, % scriptini, Exporter, exportJahre           	, % "2"

		IniRead, Startdatum           	, % scriptini, Exporter, Startdatum             	, % "01.01.2010"

		IniRead, kopiepreis            	, % scriptini, Exporter, kopiepreis             	, % "0.50€ ab 50 Seiten 0.15€"
		If (StrLen(kopiepreis) = 0 || !RegExMatch(Kopiepreis, "i)^(?<Preis1>[\d\.\,]+).*?a*b*\s*(?<AbSeite>[\d]*)\s*(Seiten*)\s*(?<Preis2>[\d\.\,]*)", k))
			kopiepreis := "0.50€ ab 50 Seiten 0.15€"

		IniRead, LastPrinter            	, % scriptini, Exporter, DokumentDrucker   	, % ""
		If InStr(LastPrinter, "ERROR")
			LastPrinter := ""

		IniRead, Autopreview           	, % scriptini, PREVIEWER, Autopreview    		, % "1"
		If InStr(Autopreview, "ERROR")
			Autopreview := "1"

		IniRead, pdfViewer             	, % scriptini, PREVIEWER, pdf                 		, % "SumatraPDF.exe "
		If InStr(pdfViewer, "ERROR")
			pdfViewer := ""

		IniRead, imgViewer             	, % scriptini, PREVIEWER, img                		, % "i_view64.exe "
		If InStr(imgViewer, "ERROR")
			imgViewer := ""

		IniRead, docViewer             	, % scriptini, PREVIEWER, doc                   	, % "notepad.exe"
		docViewer := InStr(docViewer, "ERROR") ? "" : docViewer

		IniRead, exportordner        	, % scriptini, Exporter, exportordner
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

		; Exportfilter
		IniRead, lastFilter, % scriptini, Exporter, Exportfilter
		lastFilter := InStr(lastFilter, "ERROR") ? "" : lastFilter

	;}

	; Karteikartenkürzel                                         	;{
		DBaccess 	:= new AlbisDB(Addendum.AlbisDBPath)
		KKFilter 	:= DBaccess.Karteikartenfilter()
		DbFilter 	:= ""

		For filtername, data in KKFilter
			DbFilter .= filtername "|"
		DBFilter := RTrim(DBFilter, "|")

		cJSON.EscapeUnicode := "UTF-8"
	;}

	; DIE GUI 	                                                       	;{
	AuswahlGui:

	;-: VARIABLEN/GUIGRÖSSE                           	;{
		debug   	:= 0
		FSize     	:= (A_ScreenHeight > 1080 ? 10 : 9)
		tcolor    	:= "DDDDBB"
		gdxp     	:= " gDXP_Handler "
		gopts1  	:= " Center BackgroundTrans "
		extRows 	:= 7

		sbWidth	:= (FSize	* 6) + 10
		edWidth	:= (FSize	* 140) 	- sbWidth - 15
		sfWidth 	:= (FSize	* 70) 	- sbWidth - 15
		edWidth	:= edWidth > A_ScreenWidth - 20 ? A_ScreenWidth - 20 : edWidth
		otWidth 	:= edWidth + sbWidth + 5
	;}

	;-: GUI  ANFANG 0x00180000                        	;{
		Gui, DXP: new, HWNDhDXP +E0x100000 ;-DPIScale
		Gui, DXP: Color, cCCCCAA, cDDDDBB
	;}

	;-: PATIENTENSUCHE                                       	;{
		Gui, DXP: Font, % "s" FSize " q5 Normal", Futura Bk Bt
		title := "Pat.Nr. | Nachname | Nachname, Vorname | , Vorname | Geburtsdatum | A:Arbeit | O:Ort | Straße Nr"
		Gui, DXP: Add, Text          	, % "xm ym BackgroundTrans"                                                                               	, % title

		Gui, DXP: Font, % "s" FSize-3 " q5 Normal"
		Gui, DXP: Add, Text          	, % "x+1 ym-1 BackgroundTrans"                                                                          	, % "1"
		Gui, DXP: Add, Text          	, % "x+1 ym-1 c2222AA BackgroundTrans   vDXP_One"                                        	, % "1"

		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		title := "   Zahl oder/und Buchstabe oder * z.B. Feldweg 1, Ahornallee 5A oder Hauptstraße *"
		Gui, DXP: Add, Text          	, % "x+10 ym+1 c2222AA BackgroundTrans vDXP_OneText"                                 	, % title

		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Combobox	, % "xm y+5 w" sfWidth " r20 vDXP_Pat gDXP_Handler +hwndDXPhPat"                  	, % suchfeldeintragung

		GuiControl, DXP: Choose, DXP_Pat, 1
		Gui, DXP: Add, Button      	, % "x+5 w" sbWidth  " vDXP_SB   gDXP_Handler"                                                	, % "Suchen"

		GuiControlGet, cp, DXP: Pos, DXP_Pat
		GuiControl, DXP: Move, DXP_SB, % "y" cpY-1 " h" cpH+2

		Gui, DXP: Add, Text          	, % "x+50	y" cpY+cpH-15 "   "                                                                             	, % "gefundene Patienten: "
		Gui, DXP: Add, Text          	, % "x+5 	y" cpY+cpH-15 "    vDXP_PZ"                                                                	, % "                            "
	;}

	;-: LISTVIEW PATIENTEN                                  	;{
		Gui, DXP: Font, % "s" FSize " q5 Normal"
		LVOptions 	:= " AltSubmit -Multi Grid Background" tcolor
		LVColumns 	:= "NR|Name|Vorname|Geburtstag|PLZ|Ort|Strasse|Telefon|Arbeit|letzte Beh.|timestamp|Seit"
		Gui, DXP: Add, ListView    	, % "xm y+5 w" otWidth " r12 vDXP_PLV gDXP_Handler " LVOptions, % LVColumns

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
		LV_ModifyCol(12	, Round(edWidth/12))                       	; letzte Behandlung
	;}

	;-: LISTVIEW FÜR DOKUMENTE                      	;{
		DLVWidth := Round(otWidth*0.55)
		Gui, DXP: Add, Text          	, % "xm y+5 BackgroundTrans" 	                                                           	, % "Dokumente von ["
		Gui, DXP: Add, Text          	, % "x+2 BackgroundTrans vDXP_NR"    	                                               	, % "888888"
		Gui, DXP: Add, Text          	, % "x+2 BackgroundTrans "                                                                 	, % "] "
		Gui, DXP: Add, Text          	, % "x+0 w350 BackgroundTrans vDXP_NV"                                         	, % ""
		GuiControl,  DXP: , DXP_NR, % "             "

		Gui, DXP: Add, Checkbox	, % "x+30 vDXP_AV gDXP_Handler"                                                     	, % "Autovorschau"
		GuiControl, % "DXP: " , DXP_AV, % Autopreview ; (Autopreview ? "Check1" : "Check0")

		LVOptions 	:= "Grid Count Checked AltSubmit Background" tcolor " hwndDXP_hDLV"
		LVColumns 	:= "Datum|Art|Bezeichnung|exp.|OCR|Seiten/Größe|LFDNR"
		DocCols := "exp.|OCR|Seiten"
		Gui, DXP: Add, ListView   	, % "xm y+3 w" DLVWidth " r20 vDXP_DLV	gDXP_Handler " LVOptions   	, % LVColumns
		LV_ModifyCol(1, Floor(DLVWidth/8) " Integer")
		LV_ModifyCol(2, Floor(DLVWidth/20) )
		LV_ModifyCol(3, Floor(DLVWidth/2.3))
		LV_ModifyCol(4, "40 Center")
		LV_ModifyCol(5, "40 Center")
		LV_ModifyCol(6, "90 Center")
		LV_ModifyCol(7, "0 Center Integer") ; Spalte für interne Nutzung
	;}

	;-: EINSTELLUNGEN/FILTER                           	;{

		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Text, % "x+10  BAckgroundTrans                                	vDXP_ET"                         	, % "Basispfad für den Dokumentenexport "

		PBWidth := 6 * (FSize-1) - 5
		PEWidth := otWidth - DLVWidth - PBWidth - 15
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Edit           	, % "y+3  w" PEWidth "        						vDXP_PE " gdxp              	, % exportordner
		Gui, DXP: Add, Button        	, % "x+5  w" PBWidth "       						vDXP_PB " gdxp                	, % "ändern"
		GuiControlGet, cp, DXP: pos, DXP_PE

		;-: Dateifilter
		Gui, DXP: Font, % "s" FSize " q5 Italic"
		Gui, DXP: Add, Text, % "x" cpX " y" (cpY + cpH +10) " BackgroundTrans	vDXP_FT "                        	, % "Dateien die erfasst werden sollen"
		GuiControlGet, cp, DXP: pos, DXP_FT
		Gui, DXP: Font, % "s" FSize " q5 Normal"

		cpY := cpY+cpH-15
		For chkidx, ext in StrSplit(docfilter_full, ",")
			Gui, DXP: Add, Checkbox	, % "x" 	(chkidx=1 || Mod(chkidx,extRows)=0 ? cbX:=cpX 	:  cbX+=50)
													. 	 " y" 	(chkidx=1 || Mod(chkidx,extRows)=0 ? cpY+=15  	:  cpY)
													. 		  	" vDXP_C" (lastExtCB_Index := chkidx)
													. 	  	  (inFilter(ext, docfilter_checked) ? " Checked" : "") " gDXP_Handler"	, % ext

		;-: Bezeichnungsfilter
		;{
		GuiControlGet, cp	, DXP: pos, % "DXP_C" 	lastExtCB_Index
		GuiControlGet, dp, DXP: pos, % "DXP_C1"
		Y := (cpY + cpH +15), docfX := dpX, docfW := PBWidth + PEWidth + 5
		;}

		Gui, DXP: Font, % "s" FSize " q5 Italic"
		title := "Dokumentbezeichnungen die gefiltert werden sollen(Komma getrennt)"
		Gui, DXP: Add, Text, % "x" dpX " y" Y "                  BackgroundTrans    	vDXP_FFT "            				, % title
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Edit, % "x" dpX " y+3 w" docfW " r4                     		vDXP_FF"			gdxp            	, % docfilter_bezeichner
   ;}

	;-: AUSWAHL                                                	;{
		;{
		GuiControlGet, cp, DXP: pos, % "DXP_FF"
		Y1 := (cpY + cpH + 20)
		xtraOpt := " Number HWNDDXP_hJB "
		;}

		Gui, DXP: Font, % "s" FSize-1 " Italic "
		Gui, DXP: Add, Text, % "x" cpX+2 " y" Y1-17 "          BackgroundTrans    	               "   						, % "Dokumente für den Export auswählen"

		Gui, DXP: Font, % "s" FSize-1 " Normal"
		Gui, DXP: Add, Button        	, % "x" cpX " y" Y1 "	                            	vDXP_AB"		gdxp              	, % "Alle"
		Gui, DXP: Add, Button        	, % "x+5           	                            	vDXP_KB"		gdxp              	, % "Keines"
		Gui, DXP: Add, Button        	, % "x+5           	                            	vDXP_LB"		gdxp              	, % "zurückliegende Jahre"

		Gui, DXP: Font, % "s" FSize 	" Normal"
		Gui, DXP: Add, Edit            	, % "x+2 yp+1 " xtraOpt " Center          	vDXP_JB"      	gdxp              	, % ini_yearsback

		;{
		WinSet, ExStyle, 0x0, % "ahk_id " DXP_hJB

		GuiControl, DXP: Disable        	, DXP_AB
		GuiControl, DXP: Disable        	, DXP_KB
		GuiControl, DXP: Disable        	, DXP_LB

		GuiControlGet, cp, DXP: pos, % "DXP_AB"
		awX := cpX
		xtraOpt1 	 := " y" Y1 " w" Floor((FSize)*2.2) " h" cpH-4 " Number   	HWNDDXP_hABRP "
		xtraOpt2 	 := " y" Y1+2 " w" Floor((FSize)*2.2) " h" cpH-4 " ReadOnly	hwndDXP_hGO "
		DDLPrinter := GetPrinters()

		;}

		Gui, DXP: Add, Button        	, % "x" awX " y+10 Center                     	vDXP_EA"			gdxp           	, % "000 Dokumente exportieren"
		Gui, DXP: Add, Button        	, % "x+10                                				vDXP_FB" 			gdxp         	, % "Exportpfad im Explorer öffnen"
		Gui, DXP: Add, Button        	, % "x" awX " y+15                               	vDXP_EP"			gdxp          	, % "Auswahl, Laborblatt und Karteikarte exportieren"
		Gui, DXP: Add, Button        	, % "x" awX " y+10 Center                     	vDXP_PRN"		gdxp           	, % "000 Dokumente drucken"
		Gui, DXP: Add, DDL           	, % "x+10 r5                       	         		vDXP_PRD "      	gdxp          	, % DDLPrinter               ; PRD = PrinterDevice
		Gui, DXP: Add, Button        	, % "x" awX " y+15 Center                     	vDXP_ABR"		gdxp           	, % "000 Kopien abrechnen"
		Gui, DXP: Add, Edit            	, % "x+2 w200 r1 " xtraOpt1 "              	vDXP_ABRP" 		gdxp           	, % kopiepreis
		Gui, DXP: Add, Text           	, % "x" awX " y+3 Right                        	vDXP_GOT"		gdxp           	, % "Gebührentext:"
		Gui, DXP: Font, % "s" FSize-1 " Normal"
		Gui, DXP: Add, Edit            	, % "x+2 w200 r1 " xtraOpt2 "               	vDXP_GO "                       	, % ""

		;{
		WinSet, ExStyle, 0x0, % "ahk_id " DXP_hABRP

		GuiControl, DXP:                    	, DXP_EA	, % "Dokumente exportieren"
		GuiControl, DXP:                     	, DXP_PRN, % "Dokumente drucken"
		GuiControl, DXP:                     	, DXP_ABR, % "Kopien abrechnen"
		GuiControl, DXP: Disable        	, DXP_EA
		GuiControl, DXP: Disable        	, DXP_EP
		GuiControl, DXP: Disable        	, DXP_PRN
		GuiControl, DXP: Disable        	, DXP_ABR
		GuiControl, % "DXP: " (LastPrinter ? "ChooseString" : "Choose"), DXP_PRD, % (LastPrinter ? LastPrinter : 1)

		GuiControlGet, cp, DXP: pos, % "DXP_EA"
		GuiControl, DXP: Move  	, DXP_PRN, % "w"	cpW
		GuiControl, DXP: Move  	, DXP_PRD, % "x"	cpX + cpW + 5 " w" cpW
		GuiControl, DXP: Move  	, DXP_ABR, % "w"	cpW

		GuiControlGet, dp, DXP: pos, % "DXP_ABR"
		GuiControlGet, cp, DXP: pos, % "DXP_ABRP"
		GuiControl, DXP: Move  	, DXP_ABRP, % "x" dpX + dpW + 5 " y" dpY+(dpH-cpH) " w" (ABRPw := cpW)

		GuiControlGet, dp, DXP: pos, % "DXP_ABRP"
		GuiControlGet, cp, DXP: pos, % "DXP_EP"
		bpX := cpX+cpW+10
		ABRPy := dpY
		;}

		alts := "AltSubmit "
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Text            	, % "x" bpX " y" cpY-(FSize-1) "        	vDXP_VBT" 	gopts1 	, % "von - bis"
		Gui, DXP: Font, % "s" FSize+1 " q5 Normal"
		Gui, DXP: Add, Edit           	, % "y+5 w" (FSize-1)*20 " h" cpH " vDXP_VB hwndDXP_hVB " alts gopts1, % Startdatum "-" A_DD "." A_MM "." A_YYYY

		Gui, DXP: Font, % "s" FSize-2 " q5 Normal"
		Gui, DXP: Add, Text            	, % "x" dpX " y" dpY-FSize-4 " w400   	vDXP_ABRPT"	gopts1 	, % "Regel für den Kopiepreis (Preise in Euro angeben)"

		;{
		WinSet, ExStyle, 0x0 , % "ahk_id " DXP_hVB

		; Datum verschieben
		GuiControlGet, dp, DXP: pos, % "DXP_VBT"
		GuiControl, DXP: Move, DXP_VB, % "y" cpY + Floor((cpH/2-dpH/2)/2) - 2

		; von - bis verschieben
		GuiControlGet, ep, DXP: pos, % "DXP_VB"
		GuiControl, DXP: Move, DXP_VBT	, % "x" epX + Floor((cpW/2-dpW/2)/2) " y" epY - dpH
		X := epX+epW+10

		; Druckerauswahl verschieben
		GuiControlGet, cp, DXP: pos, % "DXP_PRN"
		GuiControl, DXP: Move, DXP_PRD	, % "w" epX + epW  - cpW - cpX - 5 " y" cpY
		GuiControl, DXP: Move, DXP_ABRP, % "w" epX + epW  - cpW - cpX - 5

		; Dokument drucken Höhe anpassen
		GuiControlGet, cp, DXP: pos, % "DXP_PRD"
		GuiControl, DXP: Move, DXP_PRN	, % " h" cpH

		; Kopiepreisbeschriftung verschieben
		GuiControlGet, dp, DXP: pos, % "DXP_ABRPT"
		GuiControl, DXP: Move, DXP_ABRPT, % "y" ABRPy - dpH - 2 " w" cpW

		; Gebührentext verschieben
		GuiControlGet, ep, DXP: pos, % "DXP_FF"
		GuiControlGet, dp, DXP: pos, % "DXP_ABR"
		GuiControlGet, cp, DXP: pos, % "DXP_ABRT"
		GuiControl, DXP: Move, DXP_GOT, % "x" dpX " y" dpy+dpH+2 " w" dpW
		WinSet, Style 	, 0x50000880, % "ahk_id" DXP_hGO
		WinSet, ExStyle	, 0x00000000, % "ahk_id" DXP_hGO
		GuiControl, DXP: Move, DXP_GO	, % "x" dpX+dpW+5 " y" dpy+dpH+1 " w" epX + epW - dpX - dpW - 5

		W := docfX+docfW - X, H := cpH

		; ersten Filter nehmen in diesem Fall
		If !lastFilter
			For lastfilter, INHALT in KKFilter
				break
		;}

		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, Text            	, % "x" X " y" Y1-10 " w" (FSize-1)*10 " vDXP_KKFT "					        			 	gopts1 	, % "Exportfilter"
		Gui, DXP: Font, % "s" FSize+1 " q5 Normal"
		Gui, DXP: Add, ListBox          	, % "x" X " y" Y1  " w" W  " r1 	         	vDXP_KKF hwndDXP_hKKF "   			 	gdxp    	, % DbFilter

		;{
		GuiControl, DXP: ChooseString, DXP_KKF, % lastFilter
		WinSet, ExStyle, 0x0 , % "ahk_id " DXP_hKKF
		GuiControl        	, DXP: Move	, DXP_KKF	, % "h" cpH+4

		GuiControlGet, dp, DXP: pos  	, % "DXP_KKFT"
		GuiControlGet, ep, DXP: pos 	, % "DXP_KKF"
		Y := epY - dpH - 2
		GuiControl, DXP: Move, DXP_KKFT, % "x" epX + Floor((epW/2-dpW/2)) " y" Y
		W := dpW+epW+10

		;}

		Gui, DXP: Font, % "s" FSize " q5 Normal"
		Gui, DXP: Add, Edit            	, % "x" epX " y+1"  " w" epW " h150 " opt " vDXP_KKFI 	hwndDXP_hKKFI"                    	, %  KKFilter[lastFilter].inhalt

		;~ opt := "0x50400804 E0x00000200"

		;{
		WinSet, Style 	, 0x50400904 	, % "ahk_id " DXP_hKKFI
		WinSet, ExStyle	, 0x00000200 	, % "ahk_id " DXP_hKKFI
		GuiControlGet	, dp, DXP: Pos	, % "DXP_DLV"
		GuiControlGet	, cp, DXP: Pos	, % "DXP_EP"
		;}

		;~ Gui, DXP: Font, % "s" FSize " q5 Normal"
		;~ Gui, DXP: Add, Button        	, % "x" cpX " y" dpY+dpH-cpH-2 "		vDXP_FB" 	  		gdxp                                 	, % "Exportpfad im Explorer öffnen"
	;}

	;-: STATUS BAR                                               	;{
		Gui, DXP: Font, % "s" FSize-1 " q5 Normal"
		Gui, DXP: Add, StatusBar  	,                                                                                                                           	, % ""
		Gui, DXP: Show, % "AutoSize", % "DocPrinz"

		WinGetPos, wx, wy, ww, wh, % "ahk_id " hDXP
		w := GetWindowSpot(hDXP)

		GuiControlGet, cp, DXP: Pos, DXP_OneText
		GuiControlGet, dp, DXP: Pos, DXP_One
		GuiControl, DXP: Move, DXP_OneText	, % "x" (w.CW-cpW-10)
		GuiControl, DXP: Move, DXP_One     	, % "x" (w.CW-cpW-dpW-3)

		Gui, DXP: Show

		sb1Width	:= Round(ww/3)
		sb2Width	:= 5 * FSize
		sb3Width	:= Round(ww/3.5)
		sb4Width	:= ww - sb1Width - sb2Width - sb3Width

		SB_SetParts(sb1Width, sb2Width, sb3Width, sb4Width )
		SB_SetText("...warte auf Deine Eingabe", 1)
		SB_SetText(AlbisPath " | " exportordner, 3)
	;}

	;-: HOTKEY                                                    	;{
		fn_EnterSearch := Func("DXPSearch").Bind("Hotkey")
		Hotkey, IfWinActive, % "DocPrinz ahk_class AutoHotkeyGUI"
		Hotkey, Enter, % fn_EnterSearch
		Hotkey, IfWinActive
	;}

	;-: INTERSKRIPTKOMMUNIKATION                	;{
		OnMessage(0x4A		, "Receive_WM_COPYDATA")
		OnMessage(0x020A	, "DXPGui_Extras")              	; WM_MouseWheel
		OnMessage(0x0201	, "DXPGui_Extras")		    		; WM_LButtonDown

		;~ Gui, DXP: Submit, NoHide
		GuiControlGet, KUERZEL, DXP:, DXP_KKF

	;}

return

	; GUI LABELS                                                   	;{

DXP_Handler:                                                    	;{

	Critical
	Gui, DXP: Submit, NoHide

	;SciTEOutput("GC: " A_GuiControl ", GE: " A_GuiEvent ", EvI: " A_EventInfo)
	Switch A_GuiControl	{

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Listviews füllen --
		Case "DXP_SB":                                                     ; Patienten suchen
			DXPSearch("Button")

		Case "DXP_PLV":                                                	; Patienten Listview
			If (A_GuiControlEvent = "Normal")
				Documents := Dokumente_Show(A_EventInfo)

		Case "DXP_DLV":                                                	; Dokumenten Listview
			If (A_GuiEvent = "DoubleClick")
				Dokumente_Vorschau(A_EventInfo)
			else If RegExMatch(A_GuiEvent, "(Normal|C|I)") {
				LV_CountCheckedDocs()
				If (DXP_AV && A_EventInfo)
					Dokumente_Vorschau(A_EventInfo)        	; Dokumentvorschau
			}

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Auswählen --
		Case "DXP_AB":                                                 	; Alle Befunde
			LV_ChooseAll()
			LV_CountCheckedDocs()

		Case "DXP_KB":                                                 	; keine Befunde
			LV_ChooseNone()
			LV_CountCheckedDocs()

		Case "DXP_LB":                                                    	; die letzten X Jahre
			LV_ChooseLastYears(DXP_JB)
			LV_CountCheckedDocs()

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Export --
		Case "DXP_EA":                                                 	; 'Dokumente exportieren'
			If PatChoosed
				Dokumente_Export(PatChoosed)

		Case "DXP_EP":                                                 	; 'Karteikarte, Laborblatt u. Auswahl exportieren'
			If PatChoosed {
				Dokumente_Export(PatChoosed)
				Karteikarte_Export()
			}

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Drucken --
		Case "DXP_PRN":                                                 	; 'Dokumente drucken'
			If PatChoosed
				Dokumente_Print()

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Exportordner --
		Case "DXP_FB":                                                  	; Exportpfad im Explorer öffnen
			If PatChoosed
				Dokumente_ShowWithExplorer()

		Case "DXP_PB":                                                  	; Exportordner ändern
			ChangeFolder()

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Filter --
		Case "DXP_KKF":                                                	; Karteikartenfilter
			GuiControl, DXP:, DXP_KKFIT, % "[" DXP_KKF "]"
			Gui, DXP: Submit, NoHide
			GuiControl, DXP:, DXP_KKFI	, % KKFilter[DXP_KKF].inhalt

		; --------------------------------------------------------------------------------------------------------------------------------------------------
		; -- Abrechnung --
		Case "DXP_ABR":                                                	; gedruckte Kopien abrechnen
			Dokumente_Abrechnung()

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
		IniWrite, % checkedext, % scriptini, Exporter, docfilter_checked

	; Sucheintragungen sichern
	suchfeld := CB_GetEntries(DXPhPat)
	If (suchfeldeintragung <> suchfeld)
		IniWrite, % suchfeld	, % scriptini, Exporter, suchfeldeintragung

	; Exportordner
	GuiControlGet, export, DXP:, % "DXP_PE"
	If (exportordner <> export)
		IniWrite, % export  	, % scriptini, Exporter, exportordner

	; Dokumentbezeichner
	GuiControlGet, docbez, DXP:, % "DXP_FF"
	If (docfilter_bezeichner <> docbez)
		IniWrite, % docbez 	, % scriptini, Exporter, docfilter_bezeichner

	; zurückliegende Jahre
	Gui, DXP: Submit, Hide
	If (ini_yearsback <> DXP_JB)
		IniWrite, % DXP_JB	, % scriptini, Exporter, exportJahre_zurueckliegend

	; nur das erste Datum
	RegExMatch(DXP_VB, "(?<datum2>\d\d\.\d\d\.\d\d\d\d)", Start)
	If (Startdatum2 <> Startdatum)
		IniWrite, % Startdatum2	, % scriptini, Exporter, Startdatum

	; eingestellter Filter
	If (lastFilter <> DXP_KKF)
		IniWrite, % DXP_KKF	, % scriptini, Exporter, Exportfilter

	; ausgewählter Drucker
	If (LastPrinter <> DXP_PRD)
		IniWrite, % DXP_PRD	, % scriptini, Exporter, DokumentDrucker

	; Kopiepreis
	If (kopiepreis <> DXP_ABRP)
		IniWrite, % DXP_ABRP, % scriptini, Exporter, kopiepreis

	; Autopreview
	If (Autopreview <> DXP_AV)
		IniWrite, % DXP_AV   	, % scriptini, PREVIEWER, Autopreview

	If ReloadGui
		Reload

ExitApp ;}

;}

DXPGui_Extras()                                                   	{

	global DXP_hKKF, KKFilter, DXP, DXP_KKFI, hDXP

	MouseGetPos, mx, my,, hControl, 2
	If (hControl = DXP_hKKF) {
		Sleep 100
		SendMessage, 0x0201,,,,  % "ahk_id " hControl
		Sleep 50
		SendMessage, 0x0202,,,,  % "ahk_id " hControl
	}

}

DimmerGui(show, DBASEFileName:="")             	{       ;-- Fortschrittsanzeige

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

	Gui, DMR: Font	   	, s90 q5 bold italic cBlue, Calibri
	Gui, DMR: Add    	, Text    	, % "x0 	y" Floor(DXPPos.CH/4) " w" DXPPos.CW " vDMR_SL1 Backgroundtrans Center" 	, % "...Suche läuft"

	GuiControlGet, cp	, DMR: Pos	, DMR_SL1
	Gui, DMR: Font	   	, s24 q5 bold cBlue
	Gui, DMR: Add    	, Text    	, % "x0 y" cpY+cpH-25 " w" DXPPos.CW " vDMR_DBFN Center Backgroundtrans"            	, % DBASEFileName

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

DimmerPGS(recordNr, RecordsMax, len)             	{		;-- Callbackfunktion für die Fortschrittsanzeige

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

;}


;--- Dokumente auflisten        	;{
Dokumente_Show(row)                                                                                       	{

	; Variablen [PatDok]
		global DPGS_init, PatChoosed, DXP, DXP_DLV
		global PatMatches

		static outBezeichner

	; DIMMER-Gui einblenden
		PatChoosed := true, DPGS_init := true
		DimmerGui("on")

	; PATNR, NAME, VORNAME ermitteln
		Gui, DXP: Default
		PatChoosed := Pat := LV_GetPatient()
		If !RegExMatch(Pat.NR, "\d+") {
			PatChoosed := false
			MsgBox, % (msg := "Patient auswählen um Dokumente anzeigen zulassen.")
			SB_SetText(msg, 4)
			DimmerGui("off")
			return
		}

	; Patientendaten anzeigen
		GuiControl, DXP:, DXP_NR, % Pat.NR
		GuiControl, DXP:, DXP_NV, % Pat.NAME ", " Pat.VORNAME

	; von Datum mit Patient.Seit-Datum aus der Datenbank ersetzen
		For each, dbPat in PatMatches
			If (dbPat.NR = Pat.NR) {
				GuiControl, DXP:, DXP_VB, % ConvertDBASEDate(dbPat.Seit) "-" A_DD "." A_MM "." A_YYYY
			}

	; aktuelle Dateifilter ermitteln
		For each, extension in StrSplit(docfilter_full, ",") {
			GuiControlGet, isChecked, DXP:, % "DXP_C" A_Index
			If (isChecked = 1)
				checkedext .= (A_Index=1 ? "" : ",") extension
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
			PatDok[Pat.NR] := Dokumente_GetFromDB(Pat.NR, checkedext)

	; anhand der Filter anzeigen
		Gui, DMR: Default
		GuiControl, DMR:,  DMR_SL1	, % "Dokumente auflisten..."
		GuiControl, DMR: , DMR_SL3	, % "/" (DokMax := PatDok[Pat.NR].Count())
		GuiControl, DMR: , DMR_DBFN	, % "der gefunden Dokumente sind gelistet"

		PatPath := exportordner "\(" Pat.NR ") " Pat.NAME ", " Pat.VORNAME
		lvdocNr := AllPages := ExportedDocs := exportedPages := 0
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
			datum 	        	:= SubStr(file.datum, 7, 2) "." SubStr(file.datum, 5, 2) "." SubStr(file.datum, 1, 4)
			docPath        	:= PatPath "\" file.datum " - " file.Bezeichnung "." file.ext
			originalPath     	:= AlbisPath "\Briefe\" file.filename
			exported        	:= searchable := ""
			docExist         	:= false
			pages           	:= 0
			fileSize           	:= 0

		  ; Exportstatus
			If FileExist(docPath) {
				exported	:= "Ja"
				docExist  	:= true
				viewPath 	:= docPath
				ExportedDocs += 1
			} else if FileExist(originalPath) {
				exported	:= ""
				docExist 	:= true
				viewPath 	:= originalPath
			}

		  ; Dokumenteigenschaften: durchsuchbar (OCR oder Text) und Seitenzahl, wenn Seitenzahl nicht verfügbar dann Dateigröße
				If (docExist && RegExMatch(file.ext,"i)pdf")) {
					searchable 	:= PDFisSearchable(viewPath) ? "Ja" : ""
					pages       	:= PDFGetPages(viewPath, Addendum.qpdfPath)
					pages         	:= pages > 0 ? pages : pages
				}
				else If docExist && RegExMatch(file.ext,"i)(txt|doc|docx|odt|xml)") {
					searchable	:= "Ja"
					If RegExMatch(file.ext,"i)(doc|docx|odt)") {
						SplitPath, originalPath, sourcename, sourcepath
						fex    	:= Filexpro(originalPath,, "System.Document.PageCount")
						pages	:= fex["System.Document.PageCount"]
					}
				}
				else If docExist && RegExMatch(file.ext,"i)(txt|jpg|png|avi|wmi|wmv|mp3)") {
					FileGetSize, FSize, % originalPath, K
					fileSize     	:= FSize >= 1024 ? Round(FSize/1024, 1) " MB" : FSize " KB"
				}

		; Eintragen in die Listview
				Gui, DXP: Default
				Gui, DXP: ListView, DXP_DLV
				LV_Add(""	, datum
								, ext
								, file.Bezeichnung (!docExist ? " (Datei fehlt!)" : "")
								, exported
								, searchable
								, (pages ? pages : filesize)
								, LFDNR)

				AllPages += pages
				If (exported = "Ja")
					ExportedPages 	+= pages

		}

	; Spaltenbreite an den Inhalt anpassen
		If (LV_GetCount() > 0) || (PatDok[Pat.NR].Count() > 0) {
			LV_ModifyCol()
			LV_ModifyCol(4, "40")
			LV_ModifyCol(5, "40")
			LV_ModifyCol(6, "90")
			LV_ModifyCol(7, "0")
			GuiControl, DXP: Enable, DXP_AB
			GuiControl, DXP: Enable, DXP_KB
			GuiControl, DXP: Enable, DXP_LB
			GuiControl, DXP: Enable, DXP_EP
		}

	; neuen Statustest setzen
		DokumentInfo := (PatDok[Pat.NR].Count() > 0 ? PatDok[Pat.NR].Count() " Dokumente, exportiert: " ExportedDocs : "keine Dokumente")
		DokumentInfo .= (PatDok[Pat.NR].Count() > 0 ? "| Seiten gesamt: " AllPages ", exportiert: " ExportedPages : "")
		SB_SetText(DokumentInfo)
		SB_SetText("", 2)
		SB_SetText("PatNR: " Pat.NR, 3)
		SB_SetText("Anzahl Patienten in PatDOK: " PatDok.Count(), 4)

	; Overlay-Gui ausschalten
		DimmerGui("off")

return PATDok
}

Dokumente_GetFromDB(PATNR, dokext="pdf,jpg,gif,bmp,tif,png,wav,avi,mov,doc,docx") 	{

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
		pattern  	:= {"PATNR": "rx:\s*" PATNR, "TEXT": "rx:\w+\\\w+\.[a-z]{1,4}"}
		res2        	:= beftexte.OpenDBF()
		matches2 	:= beftexte.Search(pattern, 0, "DimmerPGS")
		res2         	:= beftexte.CloseDBF()
		;~ SciTEOutput("BefTexte rx: " beftexte.SearchRegExStr )
		beftexte 	:= ""

		For i, m in matches2 {
			If m.removed
				continue
			lfds .= (i > 1 ? "|" : "") m.LFDNR
			m.Text := Trim(RTrim(m.TEXT, "`r`n"))
			RegExMatch(m.TEXT, "i)\.(?<xt>[a-z]+)$", e)
			If RegExMatch(ext, "i)" dokext)                                    ; Dateiendungen werden erkannt
				ftoex[m.LFDNR] := {"filename":m.TEXT, "datum":m.DATUM, "ext":ext}
		}

		SB_SetText("% der Datensätze der BEFUND.DBF durchsucht", 3)
		DPGS_init	:= 1
		DimmerGui("on", DBASEFileName)

	  ; Bezeichnungstext in der Karteikarte
		befund   	:= new dBase(AlbisDBPath "\BEFUND.dbf", 0, {"name":"DXP", "control":"Statusbar", "sbNR":2})
		pattern		:= {"PATNR": "rx:\s*" PATNR, "TEXTDB": ("rx:(" RTrim(lfds, "|") ")")}
		res1        	:= befund.OpenDBF()
		matches1 	:= befund.Search(pattern, 0, "DimmerPGS")
		res1         	:= befund.CloseDBF()
		;~ SciTEOutput("Befund rx: " befund.SearchRegExStr )
		befund 		:= ""

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
				ftoex[m.TEXTDB].Bezeichnung :=  INHALT	; TEXTDB = LFDNR

		}

return ftoex
}
;}

;--- Dokumente exportieren    	;{
Dokumente_Export(Pat:="")                                   	{	;-- exportiert ausgewählte Dokumente

		global PATDok, exportordner

	; Patient
		If !IsObject(Pat)
			Pat := LV_GetPatient()
		dokumente 	:= Array()

	; Listview default machen
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV                                                    	; DLV = DokumentListView

	; zu exportierende Dateien zusammenstellen
		row := 0
		while (row := LV_GetNext(row, "C")) {
			LV_GetText(LFDNR, row, 7)                                                	; LFDNR = LaufendeNummer - Bezeichnung aus Albis-Datenbank BEFUND.DBF
			dokumente.Push(PatDok[Pat.NR][LFDNR])
		}

	; Exportvorgang ausführen
		If !dokumente.Count() {
			SB_SetText(msg := "KEINE EXPORTAUSWAHL GETROFFEN!", 4)
			MsgBox, 0x1000, DocPrinz, % msg, 3
			return 0
		}

		filesLeft :=Dokumente_Filecopy(Pat.NR, Pat.NAME ", " Pat.VORNAME, dokumente, exportordner)
		If (filesLeft = -99)  {
			PraxTT(StrReplace(A_ScriptName, ".ahk") "`n" A_ThisFunc "`nDer Patientenexportpfad konnte nicht angelegt werden.`n"
					. "Dokumente zum Exportieren: " dokumente.Count(), "3 1")
		}
		else
			PraxTT(dokumente.Count() " Dokumente sollten exportiert werden. Rest " filesLeft, "3 0")

return filesLeft = 0 ? 1 : -1*filesLeft
}

Dokumente_Filecopy(PatID, PatName, FilesToCopy, exportPath)         	{	;-- kopiert die ausgewählten Dokumente aus albiswin/briefe in den Exportordner

		global AlbisPath, exportordner

		copynr := 0

	; Dokumentenexportpfad anlegen
		If !(PatPath := CreatePatPath(PatID, PatName, exportPath)) {
			PraxTT(A_ScriptName "`n" A_ThisFunc "`nDer Patientenexportpfad konnte nicht angelegt werden", "3 1")
			return -99
		}

	; Gui Default
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV

	; Dateien kopieren
		For row, file in FilesToCopy {
			exportFile := PatPath "\" file.datum " - " file.Bezeichnung "." file.ext
			If FileExist(exportFile) {
				Dokumente_Uncheck(file.Bezeichnung)
				continue
			}
			FileCopy, % AlbisPath "\Briefe\" file.filename, % exportFile, 1
			If exportFile {
				copynr ++
				Dokumente_Uncheck(file.Bezeichnung)
				SB_SetText(copynr " von " FilesToCopy.Count() " Dateien exportiert.")
			}
			SB_SetText(AlbisPath "\Briefe\" file.filename, 1)
			SB_SetText(PatPath "\" file.datum " - " file.Bezeichnung "." file.ext, 4)
			SB_SetText(ErrorLevel, 2)
		}

return FilesToCopy.Count() - copynr
}

Dokumente_Print()                                                                            	{	;-- gewählte Dokumente drucken

	; PatDok ist globales Objekt
	; Sumatra PDFReader wird für das Drucken von PDF Dateien übr die console benötigt

	; Variablen
		global SecureGOText
		row := checkedRows := pages := 0

	; Drucker
		Gui, DXP: Submit, NoHide
		GuiControlGet, Printer, DXP: , DXP_PRD
		RegExMatch(DXP_ABR, "\d+" docCount)
		GuiControl, DXP: Disable, DXP_DLV
		SecureGOText := true

	; Patient
		Pat := LV_GetPatient()

	; Statusbar Informationen erneuern
		SB_SetText("PatNR: " Pat.NR, 3)
		SB_SetText("Anzahl Patienten in PatDOK: " PatDok.Count(), 4)

	; abgehakte Einträge einsammeln
		Gui, DXP: ListView, DXP_DLV

		while (row := LV_GetNext(row, "C")) {

			checkedRows ++
			LV_GetText(LFDNR, row, 7)
			LV_GetText(pagecount, row, 6)
			pages += (RegExMatch(pagecount, "^\d+$") ? pagecount : 1)  ; Bilddateien werden als einseitige Dokumente gerechnet

			If PATDok[Pat.NR][LFDNR].print {
				MsgBox, 0x4, DocPrinz, % PATDok[Pat.NR][LFDNR].Bezeichnung "`nwurde schon gedruckt. Soll das Dokument erneut gedruckt werden?"
				IfMsgBox, No
					continue
			}

			filepath := AlbisPath "\Briefe\" PATDok[Pat.NR][LFDNR].filename
			RegExMatch(filepath, "\.(?<ext>[a-z]+)$", file)

			If (fileext="pdf") {
				printresult := Sumatra_PrintPDF(filepath, Printer)
			}

			PATDok[Pat.NR][LFDNR].print := print := printresult = 1 ? true : false
			SB_SetText("Druckauftrag: " (print ? "":"nicht ") "gesendet, " PATDok[Pat.NR][LFDNR].filename
						  . (print ? " [" pagecount " Seite" (pagecount >1 ? "n" : "") "]"))

		}


	; Gebührentext erstellen
		GOText := pages > 0 ? AlbisKopiekosten(DXP_ABRP, pages, true) : ""                                ; gibt nur den Gebührentext zurück
		GuiControl, DXP:, DXP_GO, % (IsObject(GOText) ? GOText.1 (GOText.2 ? "-" GOText.2 : "") : "")

	; Listview aktivieren
		GuiControl, DXP: Enable, DXP_DLV
		SecureGOText := false

}

Dokumente_Abrechnung()                                                               	{	;-- Gebührentext erstellen und eintragen

	global DXP_GO

	Gui, DXP: Submit, NoHide
	If (StrLen(Trim(DXP_GO)) = 0)
		return
	Clipboard := DXP_GO
	ClipWait, 1



}

Dokumente_Uncheck(filedescription)                                                 	{	;-- Haken entfernen

		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV
		Loop % LV_GetCount() {

			LV_GetText(txt, A_Index, 3)
			If (txt = filedescription) {
				row := A_Index
				LV_Modify(row, "-Check")
				LV_Modify(row, "Col4", "Ja")
				return row
			}

		}

return 0
}

Karteikarte_Export(Pat:="")                                                              	{	;-- Tagesprotokoll und Laborblatt werden exportiert

		global DXP_VB, DXP_KKF, DXP_PLV,
		global exportordner
		global PatMatches

		Gui, DXP: Default
		Gui, DXP: Submit, NoHide

	; Patienten Nummer bekommen
		Gui, DXP: ListView, DXP_PLV
		If !IsObject(Pat) {
			Pat := LV_GetPatient()
			If !RegExMatch(Pat.NR, "^\d+$")
				return
		}

	; Dokumentenexportpfad anlegen bei Bedarf
		If !(PatPath := CreatePatPath(Pat.NR, Pat.NAME ", " Pat.VORNAME, ExportOrdner)) {
			PraxTT(A_ScriptName "`n" A_ThisFunc "`nDer Patientenexportpfad konnte nicht angelegt werden", "3 1")
			return 0
		}

	; Karteikarte des Patienten öffnen
		AlbisAkteOeffnen(Pat.NAME ", " Pat.VORNAME, Pat.NR)

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Laborblatt drucken
	;----------------------------------------------------------------------------------------------------------------------------------------------
		; Werte exportieren
			If !FileExist(PatPath "\Laborkarte von " Pat.NAME ", " Pat.VORNAME) {
				PraxTT("Laborwerte werden exportiert.", "20 1")
				LBlattRes := Laborblatt_Export(Pat.NAME ", " Pat.VORNAME, PatPath)
			}
			else {
				PraxTT("Laborwerte wurden bereits exportiert.", "3 1")
				Sleep 3000
			}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; TAGESPROTOKOLL EXPORTIEREN (Karteikarte!)
	;----------------------------------------------------------------------------------------------------------------------------------------------
		; Tagesprotokoll für Patient erstellen
			PraxTT("Karteikarteneinträge werden exportiert.", "0 1")
			TProtRes := AlbisErstelleTagesprotokoll({	"Periode"	                                    	: DXP_VB            	; Tagesprotokoll Dialogeinstellungen
																	, 		"Patienten"                                      	: "aktiver_Patient"
																	,		"Filter"                                            	: DXP_KKF
																	,		"Cave"                                           	: false
																	,		"Hinweis_bei_fehlender_Diagnose"   	: false
																	,		"Sortierung_nach_Namen"             	: false}
																	,    	""                                                                               	; SaveFolder
																	, 	false)                                                                            	; CloseProtokoll
			;SciTEOutput("Kürzel: " DXP_KKF)
			;~ If (!RegExMatch(TProtRes, "i)[A-Z]\:\\") ? 1 : 0) {

		; Tagesprotokoll aktivieren
			AlbisMDIChildActivate("Tagesprotokoll")

		; Protokoll über Drucken speichen
			KKFilePath	:= PatPath "\Karteikartenauszug von " Pat.NAME ", " Pat.VORNAME
			If FileExist(KKFilepath ".pdf")
				FileDelete, % KKFilepath ".pdf"
			MsgBox, Bitte manuell als PDF ausdrucken!

			;~ TProtRes := AlbisSaveAsPDF(KKFilePath,, true)
			;~ If (TProtRes <> 1)
				;~ return 0

		; erstelltes Tagesprotokoll schliessen
			AlbisMDIChildWindowClose("Tagesprotokoll")
			PraxTT("", "Off")

return {"Patient":Pat.NAME ", " Pat.VORNAME " (" Pat.NR ")", "TP": TProtRes, "LB": LBlattRes}
}

Laborblatt_Export(name, PatPath, Printer="", Spalten="Alles")             	{ 	;-- automatisiert den Laborblattexport

		static AlbisViewType

		AlbisViewType	:= AlbisGetActiveWindowType(true)
		savePath       	:= PatPath "\Laborbefunde von " name
		Printer            	:= Printer  	? Printer 	: "Microsoft Print to PDF"
		Spalten         	:= Spalten	? Spalten 	: "Alles"

		If FileExist(savePath ".pdf")
			FileDelete, % savePath ".pdf"

	; Aufruf der Automatisierungsfunktion
		res := AlbisLaborblattExport(Spalten, savePath, Printer)

	; ursprüngliche Ansicht wiederherstellen
		If (AlbisViewType <> AlbisGetActiveWindowType(true))
			AlbisKarteikartenAnsicht(AlbisViewType)

	; Fehlercodeanzeige
		If (res = 0) || (res > 1) {
			PraxTT("Der Export der Laborwerte ist fehlgeschlagen.`nFehlercode: " res, "4 1")
			return 0
		}

return 1
}

CreatePatPath(PATNR, PATName, exportPath:="")                             	{ 	;-- legt den Patientenexportpfad an

	global ExportOrdner

	If !exportPath
		exportPath := ExportOrdner

	If !InStr(FileExist(PatPath := exportPath "\(" PATNR ") " PATName), "D") {
		FileCreateDir, % PatPath "\"
		If !ErrorLevel                             ; ErrorLevel ist null bei Erfolg
			return PatPath
		return
	}

return PatPath
}

;}

;--- Patientensuche                 	;{
DXPSearch(callfrom)                                                                       	{	;-- für Interskript-Kommunikation (Fernsteuerung,Abfrage durch andere Skripte)

		global DXP, DXP_Pat, DXP_DLV, DXP_PLV
		global PatMatches

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
			LV_Add("", m.NR, m.NAME ,m.VORNAME, GEBURT, m.PLZ, m.ORT ,m.STRASSE " " m.HAUSNUMMER, m.TELEFON, m.ARBEIT, LASTBEH, m.GEBURT, m.Seit)
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

return PatMatches.Count()
}

DBASEGetPatID(searchstr, debug=false)                                         	{	;-- sucht in der PATIENT.DBF

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
		else If RegExMatch(searchstr, "i)^[a-zßäöü\p{L}\-\s]+\s*,*$")
			spattern := {"NAME" : ("rx:" searchstr ".*?[^\d]")}         ; findet sonst auch Straßennamen

	;-----------------------------------------------------------------------------------------------
	; Nachname, Vorname (wie oben beschrieben)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^(?<nn>[a-zßäöü\p{L}\-\s]+)\s*,\s*(?<vn>[a-zßäöü\p{L}\-\s]+)$", s)
			spattern := {"NAME" : ("rx:.*?" snn), "VORNAME": ("rx:.*?" svn)}

	;-----------------------------------------------------------------------------------------------
	; Vorname (wird erkannt durch ein Komma vor dem Wort)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^,\s*(?<vn>[a-zßäöü\p{L}\s\-]+)$", s)
			spattern := {"VORNAME" : ("rx:.*?" svn)}

	;-----------------------------------------------------------------------------------------------
	; Straße Hausnummer (anstatt einer Nummer ein Sternchen verwenden)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "i)^(?<streetn>[a-zßäöü\p{L}\-\s]+)\s+(?<streetnr>\*|[\da-z]+)$", s) {
			If !InStr(sstreetnr, "*")
				spattern := {"STRASSE" : ("rx:.*?" sstreetn), "HAUSNUMMER" : ("rx:.*?" sstreetnr)}
			else
				spattern := {"STRASSE" : ("rx:.*?" sstreetn)}
		}

	;-----------------------------------------------------------------------------------------------
	; Patienten Nr (nur Zahleneingabe)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "^\d+$")
			spattern := {"NR" : searchstr}

	;-----------------------------------------------------------------------------------------------
	; Arbeit (A:)
	;-----------------------------------------------------------------------------------------------
		else If RegExMatch(searchstr, "^A[\s\:]+(?<Arbeit>.*)$", s)
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
Antragzaehler()                                                                             	{

	antragfilter:= ["i)lageso", "^Au\s|$", "Anfrage", "Antrag", ",ausgefüllt", "MDK Anfrage", "Bericht medizinisch.*Dienst", "Auszahlschein", "Bericht.*lebens.*versich", "Bescheinigung.*Versicherung"
						, "Befundbericht.*REHA", "Antrag.*Rente", "Erinnerung.*DRV", "RLV.*Anfrage", "REHA.*Antrag", "Landesamt.*Versorgung", "Landesamt.*Sozial", "Ärztliches Attest", "zahlschein"
						, "ausgefüllt.*Rentenver*", "Antwortschreiben", "Auszahlschein", "Med.*Dienst", "Dt\s", "ausgefüllt.*Schreib[en]+", "Rücksendung", "Antwortschreiben", "Rückantwort",
						, "Deutsch.*Versich", "Bundesagentur", "LASV", "Folge\s", "Zahlung.*Krankengeld", "Befundbericht"]


}

Filexpro( sFile := "", Kind := "", P* )                                                 	{	;-- v.90 By SKAN on D1CC @ goo.gl/jyXFo9
Local
Static xDetails

	If ( sFile = "" )   {                                                           ;   Deinit static variable
        xDetails := ""
        Return
    }

  fex := {}, _FileExt := ""

  Loop, Files, % RTrim(sfile,"\*/."), DF
    {
        If not FileExist( sFile:=A_LoopFileLongPath )
          {
              Return
          }

        SplitPath, sFile, _FileExt, _Dir, _Ext, _File, _Drv

        If ( p[p.length()] = "xInfo" )                          ;  Last parameter is xInfo
          {
              p.Pop()                                           ;         Delete parameter
              fex.SetCapacity(11)                               ; Make room for Extra info
              fex["_Attrib"]    := A_LoopFileAttrib
              fex["_Dir"]       := _Dir
              fex["_Drv"]       := _Drv
              fex["_Ext"]       := _Ext
              fex["_File"]      := _File
              fex["_File.Ext"]  := _FileExt
              fex["_FilePath"]  := sFile
              fex["_FileSize"]  := A_LoopFileSize
              fex["_FileTimeA"] := A_LoopFileTimeAccessed
              fex["_FileTimeC"] := A_LoopFileTimeCreated
              fex["_FileTimeM"] := A_LoopFileTimeModified
          }
        Break
    }

  If Not ( _FileExt )                                   ;    Filepath not resolved
    {
        Return
    }


  objShl := ComObjCreate("Shell.Application")
  objDir := objShl.NameSpace(_Dir)
  objItm := objDir.ParseName(_FileExt)

  If ( VarSetCapacity(xDetails) = 0 )                           ;     Init static variable
    {
        i:=-1,  xDetails:={},  xDetails.SetCapacity(309)

        While ( i++ < 309 )
          {
            xDetails[ objDir.GetDetailsOf(0,i) ] := i
          }

        xDetails.Delete("")
    }

  If ( Kind and Kind <> objDir.GetDetailsOf(objItm,11) )        ;  File isn't desired kind
    {
        Return
    }

  i:=0,  nParams:=p.Count(),  fex.SetCapacity(nParams + 11)

  While ( i++ < nParams )
    {
        Prop := p[i]

        If ( (Dot:=InStr(Prop,".")) and (Prop:=(Dot=1 ? "System":"") . Prop) )
          {
              fex[Prop] := objItm.ExtendedProperty(Prop)
              Continue
          }

        If ( PropNum := xDetails[Prop] ) > -1
          {
              fex[Prop] := ObjDir.GetDetailsOf(objItm,PropNum)
              Continue
          }
    }

  fex.SetCapacity(-1)
Return fex

} ;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

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
	IniWrite, % exportordner, % scriptini, Exporter, exportordner

}

CB_GetEntries(hwnd) {
	ControlGet, cbList, List,,, % "ahk_id " hwnd
return StrReplace(cbList, "`n", "|")
}

LV_ChooseAll() {

	Gui, DXP: ListView, DXP_DLV
	Loop, % LV_GetCount()
		LV_Modify(A_Index, "Check")

}

LV_ChooseNone() {

	Gui, DXP: ListView, DXP_DLV
	Loop, % LV_GetCount()
		LV_Modify(A_Index, "-Check")

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

}

LV_GetPatient() {

	global DXP, DXP_PLV

	Gui, DXP: Default
	Gui, DXP: ListView, % "DXP_PLV"        ; PLV = PatientListView
	row := LV_GetNext(0, "F")
	LV_GetText(PATNR     	, row, 1)
	LV_GetText(NAME      	, row, 2)
	LV_GetText(VORNAME	, row, 3)

	;~ frow := 0
	;~ Loop {
		;~ break
	;~ }

return {"NR":PATNR, "NAME":NAME , "VORNAME":VORNAME}
}

LV_CountCheckedDocs() {

		global DXP, DXP_DLV, DXP_EA, DXP_ABRP, DXP_PRN, DXP_ABR
		global SecureGOText

		Gui, DXP: Submit, NoHide
		Gui, DXP: ListView, DXP_DLV

	; abgehakte Einträge zählen
		row := checkedRows := pages := 0
		Loop % LV_GetCount() {

			If !(row := LV_GetNext(row, "Checked"))
				break
			checkedRows ++
			LV_GetText(pagecount, row, 6)
			pages += (RegExMatch(pagecount, "^\d+$") ? pagecount : 1)  ; Pauschal werden reine Bilddateien als einseitig angenommen

		}

		Dokumentzahl 	:= (checkedRows = 0 ? "" : checkedRows) " Dokument" (checkedRows=1 ? " ": "e ")

		GuiControl, % "DXP: " (checkedRows 	? "Enable" : "Disable"), DXP_EA
		GuiControl, % "DXP: " (checkedRows 	? "Enable" : "Disable"), DXP_PRN
		GuiControl, % "DXP: " (pages         	? "Enable" : "Disable"), DXP_ABR

		GuiControl, DXP:, DXP_EA 	, % Dokumentzahl "exportieren"
		GuiControl, DXP:, DXP_PRN	, % Dokumentzahl "drucken"
		GuiControl, DXP:, DXP_ABR  	, % (pages = 0 ? "" : pages) " Kopie" (pages=1 ? " ": "n ") "abrechnen"

		If !SecureGOText {
			GOText := pages > 0 ? AlbisKopiekosten(DXP_ABRP, pages, true) : ""                                ; gibt nur den Gebührentext zurück
			GuiControl, DXP:, DXP_GO  	, % (IsObject(GOText) ? GOText.1 (GOText.2 ? "-" GOText.2 : "") : "")
		}

}

GetPrinters() {

	for Item in ComObjGet( "winmgmts:" ).ExecQuery("Select * from Win32_Printer")
			DDLPrinter .= !InStr(DDLPrinter, Item.Name) ? Item.Name "|" : ""

return  RTrim(DDLPrinter, "|")
}

;}

;--- Dokumente anzeigen          	;{
EndPreviewers(PATNR)                                         	{  ; maybe sometimes ???

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

Dokumente_Vorschau(row)                                	{

		global AlbisPath
		static SumatraWasMovedBefore := false

		PreVID := Object()

		Gui, DXP: Default

	; PATNR ermitteln
		Pat := LV_GetPatient()
		If !IsObject(Pat) {
			SB_SetText("Kein Patient ermittelbar", 4)
			return
		}

	; LFDNR aus der Dokumentlistview holen
		Gui, DXP: Default
		Gui, DXP: ListView, DXP_DLV
		LV_GetText(ext   	, row	, 2)
		LV_GetText(LFDNR, row	, 7)
		If !RegExMatch(LFDNR, "\d+") || (StrLen(ext) = 0)
			return

		file      	:= PATDok[PAT.NR][LFDNR]
		filepath	:= AlbisPath "\Briefe\" file.filename

	; Prüfen ob das Dokument vorhanden ist
		If FileExist(filepath)
			SB_SetText("zeige Dokument: " LFDNR " - " file.filename, 1)
		else {
			SB_SetText("Dokument nicht vorhanden: " LFDNR " - " file.filename, 1)
			return
		}

	; Fensterpositionen / Monitogrößen ermitteln
		sDXP   	:= GetWindowSpot(hDXP)
		hMon 	:= MonitorFromWindow(hDXP)
		Mon   	:= GetMonitorInfo(hMon)

	; Dokumenten mit Standard-Anzeigeprogramm aufrufen oder wenn pdf dann mit SumatraPDF falls vorhanden
		RegExMatch(file.filename, "i)\.(?<xt>[a-z\d]+)$", e)
		If (ext = "pdf") && InStr(pdfViewer, "Sumatra") {

			; DocPrinz Gui wird versetzt, aber nur beim ersten Start
				If !WinExist("ahk_class SUMATRA_PDF_FRAME")
					SumatraIsNotHere := true
				If (!SumatraWasMovedBefore || SumatraIsNotHere) {
					DllCall("SetWindowPos"	, "Ptr"	, hDXP
														, "Ptr"	, 0
														, "Int" 	, (Mon.WR - sDXP.W - 5)   	; x
														, "Int" 	, sDXP.Y                            ; y
														, "Int" 	, sDXP.W                        	; w
														, "Int" 	, sDXP.h							; h
														, "UInt"	, 0x40)								; SWP_SHOWWINDOW:= 0x40
				}

			; SumataPDF wird aufgerufen
				smtra := Sumatra_Show(filepath)
				If !IsObject(PreVID[PAT.NR])
					PreVID[PAT.NR] := Object()
				else if !IsObject(PreVID[PAT.NR][LFDNR])
					PreVID[PAT.NR][LFDNR] := {"exe":"Sumatra", "id": sumatra.id, "PID": SumatraPID, "file":file.filename}

			; Größenberechnungen
				mwidth 	:= Mon.WR - sDXP.W - 10
				mheigth	:= Mon.WB
				sw := Round(mheight * smtra.AR) > mwidht ? mwidth : Round(mheight * smtra.AR)
				sh := sw = mwidth ? Round(sw / smtra.AR)

			; Positionsberechnung
				sx := sw < mwidth ? (Mon.WR - sDXP.W - 10 - sw) ? 0
				sy := sh = mheight ? 0 : Round(mheight/2) - Round(sh/2)

			; SumataPDF verschieben und Inhalt anpassen
				If (!SumatraWasMovedBefore || SumatraIsNotHere) {
					SumatraWasMovedBefore := true
																														; SWP_SHOWWINDOW:= 0x40
					DllCall("SetWindowPos"	, "Ptr", smtra.ID, "Ptr", 0, "Int", 0, "Int", 0, "Int", mwidth, "Int", mheigth, "UInt", 0x40)
				}

		}
		else {

			; starten mit Standard-Anzeigeprogramm
				Run, % filepath

		}



return
}

Dokumente_ShowWithExplorer()                            	{ 	;-- öffnet den Exportpfad in einem Explorer-Fenster

	Gui, DXP: Default
	Pat := LV_GetPatient()
	If IsObject(Pat)
		PatPath := exportordner "\(" Pat.NR ") " Pat.NAME ", " Pat.VORNAME "\"
	else
		PatPath := exportordner

	If InStr(FileExist(PatPath), "D") {
		SB_SetText("exportordner: " PatPath " geöffnet.", 1)
		Run, % "explorer.exe " q PatPath q
	} else
		SB_SetText("exportordner: " PatPath " ist nicht vorhanden", 1)

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

		; Per IPC erhaltene Suchanfrage in das Suchfeld eintragen
			AddNewComboxItem(recv.opt, DXPhPat, 50)
			DXP_Pat := recv.opt

		 ; Patient suchen
			PatientsFound := DXPSearch(A_ThisFunc)

		  ; sofort die Dokumente auflisten lassen
			If (RegExMatch(DXP_Pat, "^\d+$") && PatientsFound = 1) {
				Gui, DXP: ListView, DXP_DLV
				LV_Modify(1, "Focus Select")
				;~ Dokumente_Show(1)
			}
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

	if (fde)
		DllCall(NumGet(NumGet(FileDialog+0)+7*A_PtrSize), "Ptr", FileDialog, "Ptr", fde, "UInt*", dwCookie := 0)

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
DocPrinz_ico(NewHandle := False) {
Static hBitmap := DocPrinz_ico()
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 9812 << !!A_IsUnicode)
B64 := "AAABAAEAMDAAAAEAGACoHAAAFgAAACgAAAAwAAAAYAAAAAEAGAAAAAAAAAAAAGQAAABkAAAAAAAAAAAAAACAgECAgECAgECAgEBvhC9WixZHjwdBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFkAVRjBFphip/gECAgECAgECAgECAgECAgEB9gTxUixVAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBPjQ96gjqAgECAgECAgEB8gTxLjgxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJjwl7gTuAgECAgEBTjBNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBSjBKAgEBthS1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBuhS5TjBNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkwVjpS+Dt1mVwnGly4emzIikyoWUwXCCt1hkpjBElAZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBWjBZFjwZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBQmhWSwG7L4bn5/Pf////////////////////////////////////7/frR5MGcxntWnh5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBIjwlAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBboSS616P6/Pj////////////////////////////////////////////////////////////9/vzC3K1ipC1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBDkANAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBElAaly4f7/Pn////////////////////////////////////////////////////////////////////////////9/vyz05lJlgxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkANAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBWnh3X6Mn////////////////////////////////////////////////////////////////////////////////////////////d7NJZnyFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkANAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBOmRPl8Nz////////////////////////////////////////////////////////////////////////////3+vR2sEhSmxi72KT////////q8+NTnBpAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQC72KT///////////////////////////////////////////////////////////////////////////////+bxXlAkQBAkQBDkwTb6s/////////I37VAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBQmhb9/v3////////////////////////////////////////////////////////////////////////////4+/VPmhRBkgKpzYxEkwWEuFv///////////9YnyBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBzrkT////////////////////////////p8uK616OcxnuDt1lqqThXnh9NmBFNmBFNmBFQmhZfoyl/tVSXw3RjpS9AkQB/tVT///+GuV1HlQny+O7///////9/tVRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB0r0b////////////////v9umkyoVipC1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkgLe7NP////P479AkQC31p////////+LvGRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB0r0X////////y+O6SwG5HlQlAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCHul/////////8/ftLlw93sUr///////+LvGRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB0r0X////E3bBQmhZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBrqjn3+vT///////////99tFJAkQC516L///+Ju2JAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkAJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB0r0XV58dElAZAkQBAkQBAkQBAkQBAkQBAkQBWnh1rqjp9tFGMvWaPvmmVwnGMvWaMvWaMvWaOvmiiyoPC3K77/fr///////////////+z05lAkQBElAbL4bmJu2JAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBRmxdOmRNAkQBAkQBAkQBNmRKPvmnA26vt9ef////////////////////////////////////////////////////////////////////o8uBAkQBAkQBIlgtcoSZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBPmhSt0JL5/Pf///////////////////////////////////////////////////////////////////////////////////9ZoCJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB2sEjv9un///////////////////////////////////////////////////////////////////////////////////////////+EuFtAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCCt1j9/vz///////////////////////////////////////////////////////////////////////////////////////////////+v0ZRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFlAfy+O7////////////////////////////////////////////////////////////////////////////////////////////////////a6c1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBqqTj////////////////////////////////////////////////////////////////////////////////////////////////////////t9edAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB4sUv////////////////9/vzR5MGw0pahyYKTwW+SwG6ZxHeZxHeqzo2616PN4rzf7NTv9ur///////////////////////////////////+MvWZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB6s07////////x9+ygyIBSmxhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBLlw92sEjC3K3////////////////////Z6cxBkgJAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB9tFH////C3K5PmhRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCgyID////////////+//5rqjpAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB+tVO82KVBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBLlw////////////+41qBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBHlQpElAZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJlgxipS5urD57s0+Hul+MvWWAtlWAtlVoqDZAkQBAkQBAkQD4+/b////r8+RNmBFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBfoyqSwG3C3K3m8d7+//7////////////////////////////+//6Gul5AkQBMmBD////9/v10r0VAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGlAiKvGPZ6cz////////////////////////////////////////////////////m8N1AkQBsqjv///+cxntAkQCEuFuTwW9JlgxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBrqjni7tj////////////////////////////////////////////////////////////9/v1BkgJ8tFC82KVBkQFdoif3+vT////s9OaCt1hAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQCMvWb9/vz////////////////////////////////////////////////////////////////s9OZAkQBFlAdAkQBMmBDg7db////////////+//6SwG5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBBkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBNmBH7/Pn////////////////////////////////////////////////////////////////////5/PdXnh9AkQBHlQrP47/////////////////////7/fpYnyBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQByrkP////////////////////////////////////////////////////////////////////////////z+O/P47/t9ef///////////////////////////+FuVxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB4sUv///////////////////////////////////////////////////////////////////////////////////////////////////////////////////+Ju2JAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQB7s0////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+GuV1AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQByrkP///////////////////////////////////////////////////////////////////////////////////////////////////////////////////92sEhAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBFlAfp8uL////////////////////////////////////////////////////////////////////////////////////////////////////////////v9ulNmBFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBnpzTq8+P////////////////////////////////////////////////////////////////////////////////////////////////////r8+RoqDZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBJlgyhyYHz+O/////////////////////////////////////////////////////////////////////////////////////2+vOt0JJNmBFAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGjwZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBIlguJu2HA26vy+O7////////////////////////////////////////////////////////////4+/XL4bmLvGRMmBBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBGjwZUjBRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBCkgNpqTeKvGOkyoW+2ajQ5MDZ6czi7tjm8N3f7NTU5sbK4Li/2qqsz5CNvWdrqjpElAZAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBUjBRthS5AkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBshSyAgEBTjBNAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBPjRB/gECAgEB8gTxLjgxAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBHkAd5gTqAgECAgECAgEB8gTxTjBRAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBNjQ55gjmAgECAgECAgECAgECAgECAgEBvhS5VixVGjwdAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBAkQBEkAVQjRBohih/gD+AgECAgECAgEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
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

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
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


;}

^!#::
ReloadGui := true
gosub DXPGuiClose
return

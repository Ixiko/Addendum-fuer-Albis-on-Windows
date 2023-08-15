; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                 		Addendum Dokument Finder
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;      Funktion:           	- Volltextsuche in PDF, MS Word und Textdateien
;									-
;									-
;
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 19.02.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Skripteinstellungen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	#NoEnv
	#Persistent
	#SingleInstance               	, Force
	#KeyHistory                  	, 0

	SetBatchLines                	, -1
	;~ ListLines                         	, Off

	SetKeyDelay                    	, -1, -1
	SetMouseDelay	               	, -1
	SetDefaultMouseSpeed  	, 0
	SetWinDelay                 	, -1
	SetControlDelay            	, -1
	SendMode                    	, Input

	OnExit, fdrGuiClose

	;~ Process, Priority, , A

  ; startet Windows Gdip
   	If !(pToken:=Gdip_Startup()) {
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Variablen / Einstellungen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
  ; Albis Datenbankpfad / Addendum Verzeichnis
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)

	;~ global pdffiles 	:= Object()
	global adm   	:= Object()
	global load_Bar, percent
	global PatDB
	global smtra
	global q := Chr(0x22)

	adm.Dir               	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"

	adm.compname	:= StrReplace(A_ComputerName, "-")                                                    	; der Name des Computer auf dem das Skript läuft
	workini              	:= IniReadExt(adm.Ini)

	Albis                               	:= GetAlbisPaths()
	adm.AlbisExe                 	:= Albis.Exe
	adm.AlbisDocumentsPath 	:= Albis.Briefe
	adm.AlbisDBPath            	:= Instr(FileExist(Albis.db), "D") ? Albis.db : IniReadExt("Albis", "AlbisWorkDir") "\db"

	adm.PDFPaths                	:= Array()
	adm.PDFPaths.Push(adm.AlbisDocumentsPath)
	adm.PDFPaths.Push(IniReadExt("ScanPool", "Befundordner"))
	adm.PDFPaths.Push(IniReadExt("ScanPool", "ExportOrdner"))


  ; Tray Icon erstellen
	;~ hIconLabJournal	:= Laborjournal_ico()
   	;~ Menu, Tray, Icon, % (hIconLabJournal ? "hIcon: " hIconLabJournal : adm.Dir "\assets\ModulIcons\LaborJournalS.ico")

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ;  Einstellungsbereich in der Addendum.ini anlegen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ;{
	;~ If !IniReadExt("PDFFinder", "") {
		;~ backupfailure := ""
		;~ If !FilePathCreate(adm.Dir "\_backup") {
			;~ backupfailure := "Der Backup-Pfad konnte nicht angelegt werden.`n<<" adm.Dir "\_backup>> "
			;~ PraxTT("Der Backup-Pfad konnte nicht angelegt werden!", "4 1")
		;~ }
		;~ else {
			;~ FileCopy, % adm.ini, % adm.Dir "\_backup\" A_YYYY A_MM A_DD "Addendum.ini", 1
			;~ If ErrorLevel
				;~ backupfailure := "Die Addendum.ini konnte nicht nach \_backup\" A_YYYY A_MM A_DD "Addendum.ini kopiert werden."
		;~ }
		;~ If !backupfailure {
			;~ thisIniPart =
			;~ (LTrim
			;~ [PDFFinder]

			;~ [B----
			;~ )
			;~ inifull := FileOpen(adm.ini, "r", "UTF-8").Read()
			;~ inifull := RegExReplace(inifull, "\[B\-{4}", thisIniPart)
			;~ FileOpen(adm.ini, "w", "UTF-8").Write(inifull)
		;~ }
		;~ else {
			;~ MsgBox, 0x1000, Laborjournal, % "Die Erstellung eines Backup der Addendum.ini ist aus folgenden Grund fehlgeschlagen:`n" backupfailure
			;~ ExitApp
		;~ }

	;~ }

  ;}

	tmp:=key:=index:=param:=iniVal:=inifull:=labjournalIni:=backupfailure:=filter:=""

	;}

	FinderGui()



return

~ESC:: ;{
	If WinActive("PDF Finder") {
		;~ SaveGuiPos(labJ.hwnd)
		ExitApp
	}
return ;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Basis
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
FinderGui() {

	global hfdr, fdrSearchStr, fdrBtnSearch, fdrRegEx, fdrResults, fdrPView, fdrLV, fdrHLV, fdrSB, fdrHBar, fdrHFDate, fdrFDate1,fdrFDate2
	global fdrspdf, fdrsdoc, fdrsdocx, fdrstxt, fdrHPView
	global fgColor
	global PauseSearch := false, SearchRunState := false
	static fdrHSearch, canvas

	MarginX	:= 5
	MarginY	:= 5
	A4Ratio 	:= 1.415
	guiH     	:= 960
	preW    	:= guiH // A4Ratio

	Gui, fdr: new 	, Hwndhfdr ;+AlwaysOnTop
	Gui, fdr: Color 	, % "c355C7D" , % "c99B898"
	Gui, fdr: Margin, % MarginX, % MarginY

	Gui, fdr: Font	, s9 q5 Normal cWhite, Segoe Script
	Gui, fdr: Add 	, Text, xm ym+2 BackgroundTrans, Suchtext hier eingeben

	Gui, fdr: Font	, s10
	Gui, fdr: Add 	, Checkbox, x+20 yp-2	BackgroundTrans -E0x200 vfdrRegEx  	, RegEx Suche
	Gui, fdr: Add 	, Checkbox, x+20      	BackgroundTrans -E0x200 vfdrspdf    	, pdf
	Gui, fdr: Add 	, Checkbox, x+5    		BackgroundTrans -E0x200 vfdrsdoc    	, doc
	Gui, fdr: Add 	, Checkbox, x+5    		BackgroundTrans -E0x200 vfdrsdocx    	, docx
	Gui, fdr: Add 	, Checkbox, x+5    		BackgroundTrans -E0x200 vfdrstxt      	, txt
	GuiControl, fdr:, fdrspdf, 1

	Gui, fdr: Font	, s10 cBlack, Arial
	Gui, fdr: Add 	, Edit, xm y+2 w500 h20 -E0x200 hwndfdrHSearch vfdrSearchStr, % adm.finder.lastsearch
	Edit_SetMargin(fdrHSearch, 0, 2,1,0)

	Gui, fdr: Add 	, Button, x+5 h20 -E0x200 vfdrBtnSearch gfdrHandler, Suchen

	Gui, fdr: Font	, s10 cWhite, Segoe Script
	Gui, fdr: Add 	, Text, x+10 ym+2	BackgroundTrans, Datum von
	Gui, fdr: Font	, s10 cBlack, Arial
	Gui, fdr: Add 	, Edit, y+0    	w80 h20 -E0x200 hwndfdrHFDate1 vfdrFDate1

	Gui, fdr: Font	, s10 cWhite, Segoe Script
	Gui, fdr: Add 	, Text, x+5 ym+2   	BackgroundTrans, Datum bis
	Gui, fdr: Font	, s10 cBlack, Arial
	Gui, fdr: Add 	, Edit, y+0    	w80 h20 -E0x200 hwndfdrHFDate2 vfdrFDate2

	Gui, fdr: Font	, s9 cBlack, Arial
	Gui, fdr: Add 	, ListView, x+10 ym w424 h58 -Hdr Checked -E0x200 hwndfdrHLV vfdrLV, PDF Pfade
	For rowNr, pdfPath in adm.PDFPaths
		LV_Add("Check", pdfPath)

	Gui, fdr: Font	, s9 cWhite, Segoe Script
	Gui, fdr: Add 	, Text, xm y+-14 	w600 	BackgroundTrans Center	, Ergebnisse
	Gui, fdr: Add 	, Text, x+5                 		BackgroundTrans 				, Vorschau

	LVOpt := "xm y+-2 	w800 h" guiH "  gfdrHandler AltSubmit vfdrResults hwndfdrHResults"  ; -E0x200
	Gui, fdr: Font	, s9 cBlack, Arial
	Gui, fdr: Add 	, ListView	, % LVOpt,  Pfad/Subpfad|Erstelldatum|Textausschnitt|DateTime
	LV_ModifyCol(1,300)
	LV_ModifyCol(2,80)
	LV_ModifyCol(3,350)
	LV_ModifyCol(4,0)

	fgColor := new LV_Colors(fdrHResults)
	fgColor.Critical := 200
	fgColor.SelectionColors("0x0078D6")

  ;-: previewer Größe
	cp := GuiControlGet("fdr", "Pos", "fdrResults")
	If !IsObject(smtra := Sumatra_Embed(hfdr, cp.X+cp.W+MarginX, cp.Y, preW, guiH))
		MsgBox, % "Das einbetten eines Sumatra PDF Fenster ist fehlgeschlagen!"

  ;-: Statusbar
	Gui, fdr: Add, StatusBar, hwndfdrHBar  vfdrSB

  ;-: Gui anzeigen
	Gui, fdr: Show	, % "w" cp.X+cp.w+2*MarginX+preW,  Dokument Finder

	If IsObject(smtra) {
		canvas := Sumatra_EmbedFinalize(hfdr, smtra.hwnd)
		Sumatra_Open(smtra.hwnd, A_ScriptDir "\resources\DefaultBack.pdf",, canvas)
	}

	SB_SetParts(100, 220, 630, 130, 120)

	DocSearch("collectfiles")
	;~ SB_SetText("demotext",2)
	;~ hwnd := SB_SetProgress(50,2,"BackgroundGreen cLime show")

return canvas

fdrHandler: ;{ (Wese.*?Michael|Michael.*?Wese)

	Critical

	;~ SciTEOutput("X: " A_GuiControl ", RunState: " SearchRunState ", Pause: " PauseSearch)
	Gui, fdr: Submit, NoHide

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Suchen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If (A_GuiControl = "fdrBtnSearch") {

		If !PauseSearch && !SearchRunState {

			If !fdrSearchStr
				return
			GuiControl, fdr: Text, fdrBtnSearch, % "Stop"
			fn := Func("DocSearch")
			SetTimer, % fn, -10
			SearchRunState := true

		}
		else if !PauseSearch && SearchRunState
			PauseSearch := true

	}
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Vorschau einer ausgewälten Datei
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else If (A_GuiControl = "fdrResults") {

		thisrow := A_EventInfo

		If (A_GuiEvent = "Normal") {

			Gui, fdr: ListView, fdrResult
			LV_GetText(fname, thisrow, 1)

		 ; Basispfad auslesen
			Loop % LV_GetCount() {
				LV_GetText(item, A_Index, 1)
				If RegExMatch(item, "[A-Z]:") {
					thisbasePath := RTrim(item, "\")
				}
			}

			If FileExist(fpath := thisbasePath "\" Trim(fname))
				Sumatra_Open(smtra.hwnd, thisbasePath "\" Trim(fname),, canvas)
			else if !FileExist(thisbasePath "\" Trim(fname)) {
				MsgBox, 0x1024, Dokument Finder, % "Datei: " Trim(fname) "`nbefindet sich nicht im Ordner: "  thisbasePath, 3
				GuiControl, fdr: Focus, fdResults
			}

		}

	}

	Critical, Off

return ;}

fdrGuiClose:
fdrGuiEscape: ;{

	OnExit

	MsgBox, 0x1004, % A_ScriptName, % "Beenden, bei 'Nein' wird das Skript neu gestartet?"
	IfMsgBox, Yes
	{
		If !Sumatra_Close(smtra.hwnd)
			MsgBox % "Bitte schließe den eingebetten Sumatra PDF Reader noch"
		ExitApp
	}

	TrayTip, % A_ScriptName, % "Dokument Finder wird neu gestartet", 1
	If !Sumatra_Close(smtra.hwnd)
		MsgBox % "Bitte schließe den eingebetten Sumatra PDF Reader noch"

	Reload

return  ;}
}

DocSearch(cmd:="") {

	global fdr, hfdr, fdrSearchStr, fdrRegEx, fdrResults, fdrLV, fdrHLV, fgColor, fdrHResults, PauseSearch, SearchRunState
	global collecting:=true

	static filepaths := Array()
	static matches := Array()
	static filesinDir := 0

	filepos := txtlen := 0

	Gui, fdr: Submit, NoHide

  ; Dateien zusammenstellen
	If (filepaths.Count()=0) {

		collecting := true
		GuiControl, fdr: Enable0, fdrBtnSearch

		For rowNr, pdfPath in adm.PDFPaths {

			SB_SetText("sammle Dateinamen im Ordner: " pdfPath, 2)
			filepaths[rowNr]	:= Array()
			LenMainPath     	:= StrLen(pdfPath)

			Loop, Files, % pdfPath "\*.*", R
				If A_LoopFileExt in pdf,doc,docx,txt
				{
					subPath := SubStr(A_LoopFileFullPath, LenMainPath+2, StrLen(A_LoopFileFullPath)-LenMainPath-1)
					filepaths[rowNr].Push(subPath)
					filesinDir ++
					If (Mod(filesinDir, 1000) = 0)
						SB_SetText("`t`t" filesinDir, 1)
				}

		}

		SB_SetText("Dateien in " adm.PDFPaths.Count() " Verzeichnissen gefunden", 2)
		SB_SetText("`t`t" filesinDir, 1)

		collecting := false
		GuiControl, fdr: Enable1, fdrBtnSearch
		If (cmd="collectfiles")
			return

	}

	; date filter

  ; Dokumenttypfilter                	;{
    finddocExt := "i)\.("
	GuiControlGet, IsChecked, fdr:, fdrspdf
	finddocExt .= (IsChecked ? "pdf|":"")

	GuiControlGet, IsChecked, fdr:, fdrsdoc
	finddocExt .= (IsChecked ? "doc|":"")

	GuiControlGet, IsChecked, fdr:, fdrsdocx
	finddocExt .= (IsChecked ? "docx|":"")

	GuiControlGet, IsChecked, fdr:, fdrtxt
	finddocExt .= (IsChecked ? "txt":"")
	finddocExt := RTrim(finddocExt, "|") ")$"
  ;}

  ; Dateipfadfilter                     	;{
	filesinDir   	:= 0
	searchPath 	:= []

	Gui, fdr: Default
	Gui, fdr: ListView, fdrLV

	Loop % LV_GetCount() {

		LV_GetText(path, A_Index)
		SendMessage, 0x102C, A_Index-1, 0xF000,, % "ahk_id " fdrHLV
		IsChecked := (ErrorLevel >> 12) - 1  ; Setzt IsChecked auf 1 (true), wenn rowNr abgehakt ist, ansonsten auf 0 (false).
		searchPath[A_Index] := IsChecked ? path : ""
		filesinDir += IsChecked ? filepaths[A_Index].Count() : 0

	}
	ShowAt := StrLen(filesinDir) < 5 ? 1 : 10
	;}

  ; Dateien suchen
	SB_SetText("untersuchte Dateien. aktuelle Datei:", 2)
	SB_SetText("`t`t untersuchte Zeichen:", 4)
	Gui, fdr: ListView, fdrResults

	For pathNr, path in filepaths {

		If PauseSearch
			break

	  ; abgewählte Pfade überspringen
		If !(basePath := searchPath[pathNr])
			continue

	  ; Leerzeile
		If (LV_GetCount()>0)
			LV_Add("", " ")

	  ; Ausgabe des aktuellen Verzeichnisses
		LV_Add("", basePath "\")
		fgColor.Row(LV_GetCount(), 0x527250, 0xFFFFFF)
		SendMessage, 0x102A, % LV_GetCount()-1,,, % "ahk_id " fdrHResults

	  ; Dateinamenschleife
		For each, filename in path {

			If PauseSearch
				break

			filepos ++
			thisfile := basePath "\" filename
			If (Mod(filepos, ShowAt) = 0) {
				SB_SetText("`t`t" filepos " / " filesinDir, 1)
				SB_SetText(thisfile, 3)
				SB_SetText(txtlen, 5)
			}

		  ; nicht vorhandene Dateien überspringen
			If !FileExist(thisfile) || !RegExMatch(thisfile, finddocExt)
				continue

		  ; leere Dateien ebenso
			FileGetSize, FSize, % thisfile
			If (FSize = 0)
				continue

          ; Dateierstellungsfilter anwenden (# nicht begonnen)
			FileGetTime, FTime, % thisfile, C

		  ; Volltextsuche
			If !(fpos := RegExMatch(filename, fdrSearchStr)) {

			  ; Text extrahieren oder bei Textdateien direkt lesen
				If (SubStr(filename, -2) = "txt")
					txt := FileOpen(thisfile, "r", "UTF-8").Read()
				else {
					try
						txt := IFilter(thisfile)
					catch
						continue
				}

				If (StrLen(txt)=0)
					continue
				txtlen += StrLen(txt)

				If (fpos := RegExMatch(txt, fdrSearchStr)) {

                  ; Textpassage für die Vorschau extrahieren
					preview     	:= ""
					subtxt        	:=  SubStr(txt, 1, fpos-1)
					subwords  	:= CleanAndSplit(subtxt)
					allWords    	:= CleanAndSplit(txt)
					MatchStart   	:= subwords.Count() - 3 < 0 ? 1 : subwords.Count() - 3
					MaxWords   	:= MatchStart + 8 > allWords.Count() ? allWords.Count() - MatchStart : 8

					Loop % MaxWords
						preview .= (A_Index > 1 ? " " : "") allWords[MatchStart+A_Index]

					FDay := SubStr(FTime, 7, 2) "." SubStr(FTime, 5, 2) "."  SubStr(FTime, 1, 4)
					LV_Add("", " " filename, FDay, preview, FTime)
				}

			}
			else
				LV_Add("", " " filename, FDay, "im Dateinamen gefunden", FTime)

		}

	}

	GuiControl, fdr:Text, fdrBtnSearch, % "Suchen"
	PauseSearch := SearchRunState := false

	MsgBox, Fertig

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Hilfsfunktionen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
SaveGuiPos(hwnd) {

	MinMaxState := WinGetMinMaxState(hwnd)
	If (MinMaxState <> "i") {
		win 			:= GetWindowSpot(hwnd)
		winPos	 	:= "x" win.X " y" win.Y " w" win.CW " h" win.CH   (MinMaxState = "z" ? " Maximize " : " ")
		IniWrite, % winpos	, % adm.Ini, % adm.compname, LaborJournal_Position
	}

}

PosStringToObject(string) {

	p := Object()
	For wIdx, coord in StrSplit("XYWH") {
		RegExMatch(string, "i)" coord "(?<Pos>\d+)", w)
		p[coord] := wPos
	}

	p.Maximize := InStr(string, "Maximize") ? true : false

return p
}

MessageWorker(InComing) {                                                                                	;-- verarbeitet die eingegangen Nachrichten

		recv := {	"txtmsg"		: (StrSplit(InComing, "|").1)
					, 	"opt"     	: (StrSplit(InComing, "|").2)
					, 	"fromID"	: (StrSplit(InComing, "|").3)}

	; Laborabruf_*.ahk
		; per WM_COPYDATA die Daten im Laborjournal auffrischen (Skript wird neu gestartet)
		if RegExMatch(recv.txtmsg, "i)\s*reload") {
			Send_WM_COPYDATA("reload Laborjournal||" labJ.hwnd, recv.fromID)
			Reload
		}

return
}

hWnd_to_hBmp( hWnd:=-1, Client:=0, A:="", C:="" ) {                                       	;-- Capture fullscreen, Window, Control or user defined area of these

	; By SKAN C/M on D295|D299 @ bit.ly/2lyG0sN

	A      := IsObject(A) ? A : StrLen(A) ? StrSplit( A, ",", A_Space ) : {},     A.tBM := 0
	Client := ( ( A.FS := hWnd=-1 ) ? False : !!Client ), A.DrawCursor := "DrawCursor"
	hWnd   := ( A.FS ? DllCall( "GetDesktopWindow", "UPtr" ) : WinExist( "ahk_id" . hWnd ) )

	A.SetCapacity( "WINDOWINFO", 62 ),  A.Ptr := A.GetAddress( "WINDOWINFO" )
	A.RECT := NumPut( 62, A.Ptr, "UInt" ) + ( Client*16 )

	If (DllCall("GetWindowInfo",   "Ptr",hWnd, "Ptr",A.Ptr ) && DllCall("IsWindowVisible", "Ptr",hWnd ) && DllCall("IsIconic", "Ptr",hWnd )=0) {
        A.L	:= NumGet( A.RECT+ 0, "Int" )	, A.X 	:= ( A.1 <> "" ? A.1 : (A.FS ? A.L : 0) )
        A.T	:= NumGet( A.RECT+ 4, "Int" )	, A.Y 	:= ( A.2 <> "" ? A.2 : (A.FS ? A.T : 0 ))
        A.R	:= NumGet( A.RECT+ 8, "Int" )	, A.W	:= ( A.3  >  0 ? A.3 : (A.R - A.L - Round(A.1)) )
        A.B	:= NumGet( A.RECT+12, "Int" )	, A.H 	:= ( A.4  >  0 ? A.4 : (A.B - A.T - Round(A.2)) )

        A.sDC	:= DllCall( Client ? "GetDC" : "GetWindowDC", "Ptr",hWnd, "UPtr" )
        A.mDC	:= DllCall( "CreateCompatibleDC", "Ptr",A.sDC, "UPtr")
        A.tBM 	:= DllCall( "CreateCompatibleBitmap", "Ptr",A.sDC, "Int",A.W, "Int",A.H, "UPtr" )

        DllCall( "SaveDC", "Ptr",A.mDC )
        DllCall( "SelectObject", "Ptr",A.mDC, "Ptr",A.tBM )
        DllCall( "BitBlt",       "Ptr",A.mDC, "Int",0,   "Int",0, "Int",A.W, "Int",A.H, "Ptr",A.sDC, "Int",A.X, "Int",A.Y, "UInt",0x40CC0020 )

        A.R := ( IsObject(C) || StrLen(C) ) && IsFunc( A.DrawCursor ) ? A.DrawCursor( A.mDC, C ) : 0
        DllCall( "RestoreDC", "Ptr",A.mDC, "Int",-1 )
        DllCall( "DeleteDC",  "Ptr",A.mDC )
        DllCall( "ReleaseDC", "Ptr",hWnd, "Ptr",A.sDC )
    }

Return A.tBM
}

SavePicture(hBM, sFile) {                                                                                      	;-- By SKAN on D293 @ bit.ly/2krOIc9
Local V,  pBM := VarSetCapacity(V,16,0)>>8,  Ext := LTrim(SubStr(sFile,-3),"."),  E := [0,0,0,0]
Local Enc := 0x557CF400 | Round({"bmp":0, "jpg":1,"jpeg":1,"gif":2,"tif":5,"tiff":5,"png":6}[Ext])
  E[1] := DllCall("gdi32\GetObjectType", "Ptr",hBM ) <> 7
  E[2] := E[1] ? 0 : DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "Ptr",hBM, "UInt",0, "PtrP",pBM)
  NumPut(0x2EF31EF8,NumPut(0x0000739A,NumPut(0x11D31A04,NumPut(Enc+0,V,"UInt"),"UInt"),"UInt"),"UInt")
  E[3] := pBM ? DllCall("gdiplus\GdipSaveImageToFile", "Ptr",pBM, "WStr",sFile, "Ptr",&V, "UInt",0) : 1
  E[4] := pBM ? DllCall("gdiplus\GdipDisposeImage", "Ptr",pBM) : 1
Return E[1] ? 0 : E[2] ? -1 : E[3] ? -2 : E[4] ? -3 : 1
}

Edit_SetMargin(hEdit, mLeft:=0, mTop:=0, mRight:=0, mBottom:=0) {

	static dpi := A_ScreenDPI / 96

	VarSetCapacity(RECT, 16, 0 )

	; SendMessage, 0xB2,, &RECT,, ahk_id %hEdit% ; EM_GETMARGIN
	DllCall("GetClientRect", "ptr", hEdit, "ptr", &RECT)
	right  := NumGet(RECT, 8, "Int")
	; bottom := NumGet(RECT, 12, "Int")

	NumPut(0     + Ceil(mLeft*dpi) , RECT, 0, "Int")
	NumPut(0     + Ceil(mTop*dpi)  , RECT, 4, "Int")
	NumPut(right - Ceil(mRight*dpi), RECT, 8, "Int")
	; NumPut(bottom - mBottom, RECT, 12, "Int")
	SendMessage, 0xB3, 0x0, &RECT,, % "ahk_id " hEdit ; EM_SETMARGIN
}

LoadBar_Gui(show:=1, opt:="") {

	global load_BarGUI, load_Bar

	If !IsObject(opt)
		opt:={"col": ["0x4D4D4D","0xFFFFFF","0xEFEFEF"], "w":320,	"h":36}

	Gui, load_BarGUI: -Border -Caption +ToolWindow HWNDhLoad_BarWin
	Gui, load_BarGUI: Color, % opt.col.1, % opt.col.2

	load_Bar := new LoaderBar("load_BarGUI", 3, 3, opt.W, opt.H, "LABORJOURNAL", 1, opt.col.3)

	wW :=load_Bar.Width + 2*load_Bar.X
	wH :=load_Bar.Height + 2*load_Bar.Y

	Gui, load_BarGUI: Show, % "w" wW " h" wH, % "Laborjournal lädt..."
	load_Bar.hWin := hLoad_BarWin

}

LoadBar_Callback(param*) {


}

class LoaderBar {

	__New(GUI_ID:="Default",x:=0,y:=0,w:=280,h:=28,Title:="",ShowDesc:=0,FontColorDesc:="2B2B2B",FontColor:="EFEFEF",BG:="2B2B2B|2F2F2F|323232",FG:="66A3E2|4B79AF|385D87") {
		SetWinDelay,0
		SetBatchLines,-1
		if (StrLen(A_Gui))
			_GUI_ID:=A_Gui
		else
			_GUI_ID:=1
		if ( (GUI_ID="Default") || !StrLen(GUI_ID) || GUI_ID==0 )
			GUI_ID:=_GUI_ID

		this.GUI_ID := GUI_ID
		Gui, %GUI_ID%:Default

		this.BG     	:= StrSplit(BG,"|")
		this.BG.W  	:= w
		this.BG.H  	:= h
		this.Width 	:=w
		this.Height	:=h
		this.FG       	:= StrSplit(FG,"|")
		this.FG.W  	:= this.BG.W - 2
		this.FG.H   	:= (fg_h:=(this.BG.H - 2))
		this.Percent 	:= 0
		this.X        	:= x
		this.Y        	:= y
		fg_x            	:= this.X + 1
		fg_y            	:= this.Y + 1
		this.FontColor := FontColor
		this.ShowDesc := ShowDesc

		;DescBGColor:="4D4D4D"
		DescBGColor:="Black"
		this.DescBGColor := DescBGColor

		this.FontColorDesc := FontColorDesc

		Gui,Font,s10
		Gui, Add, Text, % "x" x " y" y " w" w " h" h " BackgroundTrans 0xE hwndhLoaderBarTitle", % Title
		this.hLoaderBarTitle := hLoaderBarTitle

		Gui,Font,s8
		Gui, Add, Text, % "x" x " y+1 w" w " h" h " 0xE hwndhLoaderBarBG"
		this.ApplyGradient(this.hLoaderBarBG	:= hLoaderBarBG,this.BG.1, this.BG.2, this.BG.3,1)

		Gui, Add, Text, x%fg_x% y%fg_y% w0 h%fg_h% 0xE hwndhLoaderBarFG
		this.ApplyGradient(this.hLoaderBarFG   	:= hLoaderBarFG,this.FG.1, this.FG.2, this.FG.3,1)

		Gui, Add, Text, x%x% y%y% w%w% h%h% 0x200 border center BackgroundTrans hwndhLoaderNumber c%FontColor%, % "[ 0 % ]"
			this.hLoaderNumber := hLoaderNumber

		if (this.ShowDesc) {
			Gui, Add, Text, xp y+2 w%w% h16 0x200 Center border BackgroundTrans hwndhLoaderDesc c%FontColorDesc%, Loading...
			this.hLoaderDesc := hLoaderDesc
			this.Height:=h+18
		}

		Gui,Font

		Gui, %_GUI_ID%:Default
	}

	Set(p,w:="Loading...") {
		if (StrLen(A_Gui))
			_GUI_ID:=A_Gui
		else
			_GUI_ID:=1
		GUI_ID := this.GUI_ID

		Gui, %GUI_ID%:Default
		GuiControlGet, LoaderBarBG, Pos, % this.hLoaderBarBG

		this.BG.W := LoaderBarBGW
		this.FG.W := LoaderBarBGW - 2
		this.Percent:=(p>=100) ? p:=100 : p

		PercentNum	:= Round(this.Percent,0)
		PercentBar	:= Floor((this.Percent/100)*(this.FG.W))

		hLoaderBarTitle		:= this.hLoaderBarTitle
		hLoaderBarFG  	:= this.hLoaderBarFG
		hLoaderNumber 	:= this.hLoaderNumber

		GuiControl,Move	,% hLoaderBarFG  	, % "w" PercentBar
		GuiControl,       	,% hLoaderNumber 	, % "[" PercentNum "% ]"

		if (this.ShowDesc) {
			hLoaderDesc := this.hLoaderDesc
			GuiControl,,%hLoaderDesc%, %w%
		}
		Gui, %_GUI_ID%:Default
	}

	ApplyGradient( Hwnd, LT := "101010", MB := "0000AA", RB := "00FF00", Vertical := 1 ) {
		Static STM_SETIMAGE := 0x172
		ControlGetPos,,, W, H,, ahk_id %Hwnd%
		PixelData := Vertical ? LT "|" LT "|" LT "|" MB "|" MB "|" MB "|" RB "|" RB "|" RB : LT "|" MB "|" RB "|" LT "|" MB "|" RB "|" LT "|" MB "|" RB
		hBitmap := this.CreateDIB( PixelData, 3, 3, W, H, True )
		oBitmap := DllCall( "SendMessage", "Ptr",Hwnd, "UInt",STM_SETIMAGE, "Ptr",0, "Ptr",hBitmap )
		Return hBitmap, DllCall( "DeleteObject", "Ptr",oBitmap )
	}

	CreateDIB( PixelData, W, H, ResizeW := 0, ResizeH := 0, Gradient := 1  ) {
		; http://ahkscript.org/boards/viewtopic.php?t=3203                  SKAN, CD: 01-Apr-2014 MD: 05-May-2014
		Static LR_Flag1 := 0x2008 ; LR_CREATEDIBSECTION := 0x2000 | LR_COPYDELETEORG := 8
			,  LR_Flag2 := 0x200C ; LR_CREATEDIBSECTION := 0x2000 | LR_COPYDELETEORG := 8 | LR_COPYRETURNORG := 4
			,  LR_Flag3 := 0x0008 ; LR_COPYDELETEORG := 8
		WB := Ceil( ( W * 3 ) / 2 ) * 2,  VarSetCapacity( BMBITS, WB * H + 1, 0 ),  P := &BMBITS
		Loop, Parse, PixelData, |
		P := Numput( "0x" A_LoopField, P+0, 0, "UInt" ) - ( W & 1 and Mod( A_Index * 3, W * 3 ) = 0 ? 0 : 1 )
		hBM := DllCall( "CreateBitmap", "Int",W, "Int",H, "UInt",1, "UInt",24, "Ptr",0, "Ptr" )
		hBM := DllCall( "CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_Flag1, "Ptr" )
		DllCall( "SetBitmapBits", "Ptr",hBM, "UInt",WB * H, "Ptr",&BMBITS )
		If not ( Gradient + 0 )
			hBM := DllCall( "CopyImage", "Ptr",hBM, "UInt",0, "Int",0, "Int",0, "UInt",LR_Flag3, "Ptr" )
		Return DllCall( "CopyImage", "Ptr",hBM, "Int",0, "Int",ResizeW, "Int",ResizeH, "Int",LR_Flag2, "UPtr" )
	}
}

DownloadToString(url, encoding = "utf-8") {
    static a := "AutoHotkey/" A_AhkVersion
    if (!DllCall("LoadLibrary", "str", "wininet") || !(h := DllCall("wininet\InternetOpen", "str", a, "uint", 1, "ptr", 0, "ptr", 0, "uint", 0, "ptr")))
        return 0
    c := s := 0, o := ""
    if (f := DllCall("wininet\InternetOpenUrl", "ptr", h, "str", url, "ptr", 0, "uint", 0, "uint", 0x80003000, "ptr", 0, "ptr"))    {
        while (DllCall("wininet\InternetQueryDataAvailable", "ptr", f, "uint*", s, "uint", 0, "ptr", 0) && s > 0)        {
            VarSetCapacity(b, s, 0)
            DllCall("wininet\InternetReadFile", "ptr", f, "ptr", &b, "uint", s, "uint*", r)
            o .= StrGet(&b, r >> (encoding = "utf-16" || encoding = "cp1200"), encoding)
        }
        DllCall("wininet\InternetCloseHandle", "ptr", f)
    }
    DllCall("wininet\InternetCloseHandle", "ptr", h)
    return o
}

GetGuiClassStyle()                                                             	{
	Gui, GetGuiClassStyleGUI:Add, Text
	Module := DllCall("GetModuleHandle", "Ptr", 0, "UPtr")
	VarSetCapacity(WNDCLASS, A_PtrSize * 10, 0)
	ClassStyle := DllCall("GetClassInfo", "Ptr", Module, "Str", "AutoHotkeyGUI", "Ptr", &WNDCLASS, "UInt")
                 ? NumGet(WNDCLASS, "Int")
                 : ""
	Gui, GetGuiClassStyleGUI:Destroy
	Return ClassStyle
}

SetGuiClassStyle(HGUI, Style)                                             	{
	Return DllCall("SetClassLong" . (A_PtrSize = 8 ? "Ptr" : ""), "Ptr", HGUI, "Int", -26, "Ptr", Style, "UInt")
}

Sumatra_Open(hSumatra, fpath, altPath:="", canvas:="")      	{

	If !(SumatraPID := WinGet(hSumatra, "PID")) || !FileExist(fpath)
		If FileExist(altPath)
			fpath := altpath

	SplitPath, fpath, fname

  ; schließt das vorherige Dokument
	If (fname <> "DefaultBack.pdf") {
		tabs := Sumatra_GetActiveTab(hSumatra, "tabs")
		If (tabs.Count() > 0) {
			res := SumatraInvoke("Close", hSumatra)
			while (tabs.Count()>0 && A_Index < 100) {
				If (tabs.Count()=1 && tabs[1]="DefaultBack.pdf")
					break
				tabs := Sumatra_GetActiveTab(hSumatra, "tabs")
				Sleep 30
			}
		}
	}

  ; öffnet die PDF Datei und überprüft den Erfolg
	res1 := SumatraDDE(hSumatra, "OpenFile", fpath, 0, 0, 1)
	tabname := Sumatra_GetActiveTab(hSumatra, "name")
	while (!InStr(tabname := Sumatra_GetActiveTab(hSumatra, "name"), fname) && A_Index < 100) ;~10s
		Sleep 20

  ; Toolbar ausblenden
	Sumatra_ShowToolbar(hSumatra, false)

  ; -1 = fit page
	res2 := SumatraDDE(hSumatra, "SetView", fpath, "single page", "-1") * 2
	Sleep 100

  ; PDF Renderbereich nach jedem Öffnen neu anpassen bei eingebundenen Fenstern
	If IsObject(canvas)
		WinMove, % "ahk_id " canvas.hwnd,, 0, 0, % canvas.W, % canvas.H

return res+res2
}

CleanAndSplit(txt) {

	txt := RegExReplace(txt, "[\n\r]+", " ")
	txt := RegExReplace(txt, "\s{2,}"	, " ")

return StrSplit(txt, A_Space)
}


SB_SetProgress(Value=0,Seg=1,Ops="", hGui=0) {

	; SB_SetProgress
	; (w) by DerRaphael / Released under the Terms of EUPL 1.0
	; see http://ec.europa.eu/idabc/en/document/7330 for details
	; https://www.autohotkey.com/board/topic/34593-stdlib-sb-setprogress/

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
   ControlGet, hwndBar, hWnd,,msctls_statusbar321, % (hGui ? "ahk_id " hGui : "")
	SciTEOutput("hwndbar " hwndBar)

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
         Red:=0xFF0000,Blue:=0x0000FF,Fuchsia:=0xFF00FF,Aqua:=0x00FFFF

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
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Includes
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
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

#Include %A_ScriptDir%\..\..\lib\acc.ahk
#Include %A_ScriptDir%\..\..\lib\class_CJSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_Neutron.ahk
#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
#include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;}





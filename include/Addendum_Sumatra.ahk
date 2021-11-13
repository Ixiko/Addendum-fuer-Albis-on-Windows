;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;       	RPA (robotic prozess automation) FUNKTIONEN für
;   	    SUMATRA PDF READER
;
;			ABHÄNGIGKEITEN:
;                                                                                  	------------------------
;          	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 04.02.2021 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
return

;Sumatra PDF Viewer
SumatraInvoke(command, SumatraID="") {                                                                     	;-- wm_command wrapper for SumatraPDF V3.1 & 3.2

	/* DESCRIPTION of FUNCTION:  SumatraInvoke()

		                                                                               by Ixiko (version 30.11.2020)
		---------------------------------------------------------------------------------------------------
		a wm_command wrapper for SumatraPdf V3.1 & V3.2
		...........................................................
		Remark:
		- SumatraPDF has changed all wm_command codes from V3.1 to V3.2
		- the script tries to automatically recognize the version of the addressed
		  SumatraPDF process in order to send the correct commands
		- maybe not all commands are listed !
		---------------------------------------------------------------------------------------------------
		Parameters:
		- command:	the names are borrowed from menu or toolbar names. However,
							no whitespaces or hyphens are used, only letters
		- SumatraID:	by use of a valid handle, this function will post your command to
							Sumatra. Otherwise by use of a version string ("3.1" or "3.2") this
							function returns the wm_command code.
		...........................................................
		Rersult:
		- You have to control the success of postmessage command yourself!
		---------------------------------------------------------------------------------------------------

		---------------------------------------------------------------------------------------------------
		      EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES - EXAMPLES

		SumatraInvoke("ShowToolbar", "3.2")      SumatraInvoke("DoublePage", SumatraID)
		.................................................       ...................................................................
		this one only returns the Sumatra              sends the command "DoublePage" to
        command-code                                             your specified Sumatra process using
																 parameter 2 (SumatraID) as window handle.
															 		          command-code will be returned too
		---------------------------------------------------------------------------------------------------

	*/

		static	SumatraCmds
		local	SumatraPID

		If !IsObject(SumatraCmds) {

			SumatraCmds := Object()
			SumatraCmds["3.1"] := { 	"NewWindow":                      	0      	; not available in this version -dummy command
												, 	"Open":                                 	400  	; File
												,	"Close":                                 	401  	; File
												,	"ShowInFolder":                        	0      	; not available in this version -dummy command
												,	"SaveAs":                               	402  	; File
												,	"Rename":                              	580  	; File
												,	"Print":                                   	403  	; File
												,	"SendMail":                           	408  	; File
												,	"Properties":                           	409  	; File
												,	"OpenLast1":                         	510  	; File
												,	"OpenLast2":                         	511  	; File
												,	"OpenLast3":                         	512  	; File
												,	"Exit":                                    	405  	; File
												,	"SinglePage":                         	410  	; View
												,	"DoublePage":                       	411  	; View
												,	"BookView":                           	412  	; View
												,	"ShowPagesContinuously":     	413  	; View
												,	"MangaMode":                       	0      	; not available in this version -dummy command
												,	"TurnCounterclockwise":         	415  	; View
												,	"RotateLeft":                           	415  	; View
												,	"TurnClockwise":                    	416  	; View
												,	"RotateRight":                          	416  	; View
												,	"Presentation":                        	418  	; View
												,	"Fullscreen":                           	421  	; View
												,	"Bookmark":                          	000  	; View - do not use! empty call!
												,	"ShowToolbar":                      	419  	; View
												,	"SelectAll":                             	422  	; View
												,	"CopyAll":                              	420  	; View
												,	"NextPage":                           	430  	; GoTo
												,	"PreviousPage":                      	431  	; GoTo
												,	"FirstPage":                            	432  	; GoTo
												,	"LastPage":                            	433  	; GoTo
												,	"GotoPage":                          	434  	; GoTo
												,	"Back":                                  	558  	; GoTo
												,	"Forward":                             	559  	; GoTo
												,	"Find":                                   	435  	; GoTo
												,	"FitASinglePage":                    	[410, 440] ; to hat V3.2 function (not tested)
												,	"FitPage":                              	440  	; Zoom
												,	"ActualSize":                          	441  	; Zoom
												,	"FitWidth":                             	442  	; Zoom
												,	"FitContent":                          	456  	; Zoom
												,	"CustomZoom":                     	457  	; Zoom
												,	"Zoom6400":                        	443  	; Zoom
												,	"Zoom3200":                        	444  	; Zoom
												,	"Zoom1600":                        	445  	; Zoom
												,	"Zoom800":                          	446  	; Zoom
												,	"Zoom400":                          	447  	; Zoom
												,	"Zoom200":                          	448  	; Zoom
												,	"Zoom150":                          	449  	; Zoom
												,	"Zoom125":                          	450  	; Zoom
												,	"Zoom100":                          	451  	; Zoom
												,	"Zoom50":                            	452  	; Zoom
												,	"Zoom25":                            	453  	; Zoom
												,	"Zoom12.5":                          	454  	; Zoom
												,	"Zoom8.33":                          	455  	; Zoom
												,	"AddPageToFavorites":           	560  	; Favorites
												,	"RemovePageFromFavorites": 	561  	; Favorites
												,	"ShowFavorites":                    	562  	; Favorites
												,	"CloseFavorites":                    	1106  	; Favorites
												,	"CurrentFileFavorite1":           	600  	; Favorites
												,	"CurrentFileFavorite2":           	601  	; Favorites -> I think this will be increased with every page added to favorites
												,	"ChangeLanguage":               	553  	; Settings
												,	"Options":                             	552  	; Settings
												,	"AdvancedOptions":               	597  	; Settings
												,	"VisitWebsite":                        	550  	; Help
												,	"Manual":                              	555  	; Help
												,	"CheckForUpdates":               	554  	; Help
												,	"About":                                	551}  	; Help
			SumatraCmds["3.2"] := { 	"NewWindow":                      	450  	; File
												, 	"Open":                                 	400  	; File
												,	"Close":                                 	404  	; File
												,	"ShowInFolder":                        	410  	; File
												,	"SaveAs":                               	406  	; File
												,	"Rename":                              	610  	; File
												,	"Print":                                   	408  	; File
												,	"SendMail":                           	418  	; File
												,	"Properties":                           	420  	; File
												,	"OpenLast1":                         	570  	; File
												,	"OpenLast2":                         	571  	; File
												,	"OpenLast3":                         	572  	; File
												,	"Exit":                                    	412  	; File
												,	"SinglePage":                         	422  	; View
												,	"DoublePage":                       	423  	; View
												,	"BookView":                           	424  	; View
												,	"ShowPagesContinuously":     	425  	; View
												,	"MangaMode":                       	426  	; View
												,	"RotateLeft":                              	432  	; View
												,	"RotateRight":                            	434  	; View
												,	"Presentation":                        	438  	; View
												,	"Fullscreen":                          	444  	; View
												,	"Bookmark":                          	000  	; View - do not use! empty call!
												,	"ShowToolbar":                      	440  	; View
												,	"SelectAll":                             	446  	; View
												,	"CopyAll":                              	442  	; View
												,	"NextPage":                           	460  	; GoTo
												,	"PreviousPage":                      	462  	; GoTo
												,	"FirstPage":                            	464  	; GoTo
												,	"LastPage":                            	466  	; GoTo
												,	"GotoPage":                          	468  	; GoTo
												,	"Back":                                  	596  	; GoTo
												,	"Forward":                             	598  	; GoTo
												,	"Find":                                   	470  	; GoTo
												,	"FindNext":                             	472  	; Toolbar
												,	"FindPrevious":                          	474  	; Toolbar
												,	"MatchCase":                          	476  	; Toolbar
												,	"FitWithContinuously":               	3026	; Toolbar
												,	"FitASinglePage":                    	3027	; Toolbar
												,	"ZoomIn":                               	3012	; Toolbar
												,	"ZoomOut":                            	3013	; Toolbar
												,	"FitPage":                              	480  	; Zoom
												,	"ActualSize":                          	481  	; Zoom
												,	"FitWidth":                             	482  	; Zoom
												,	"FitContent":                          	496  	; Zoom
												,	"CustomZoom":                     	497  	; Zoom
												,	"Zoom6400":                        	483  	; Zoom
												,	"Zoom3200":                        	484  	; Zoom
												,	"Zoom1600":                        	485  	; Zoom
												,	"Zoom800":                          	486  	; Zoom
												,	"Zoom400":                          	487  	; Zoom
												,	"Zoom200":                          	488  	; Zoom
												,	"Zoom150":                          	489  	; Zoom
												,	"Zoom125":                          	490  	; Zoom
												,	"Zoom100":                          	491  	; Zoom
												,	"Zoom50":                            	492  	; Zoom
												,	"Zoom25":                            	493  	; Zoom
												,	"Zoom12.5":                          	494  	; Zoom
												,	"Zoom8.33":                          	495  	; Zoom
												,	"AddPageToFavorites":           	600  	; Favorites
												,	"RemovePageFromFavorites": 	602  	; Favorites
												,	"ShowCloseFavorites":               	604  	; Favorites
												,	"CurrentFileFavorite1":           	700  	; Favorites
												,	"CurrentFileFavorite2":           	701  	; Favorites -> I think this will be increased with every page added to favorites
												,	"ChangeLanguage":               	588  	; Settings
												,	"Options":                             	586  	; Settings
												,	"AdvancedOptions":               	632  	; Settings
												,	"VisitWebsite":                        	582  	; Help
												,	"Manual":                              	592  	; Help
												,	"CheckForUpdates":               	590  	; Help
												,	"About":                                	584  	; Help
												,	"HighlightLinks":                       	616  	; Debug
												,	"ToggleEBookUI":                     	624  	; Debug
												,	"MuiDebugPaint":                     	626  	; Debug
												,	"MuiDebugPaint":                     	626  	; Debug
												,	"AnnotationFromSelection":      	628  	; Debug
												,	"DownloadSymbols":                 	630}  	; Debug

		}

	; ---------------------------------------------------------------------------------------------------------------------
	; try to determine the version of the running SumatraPDF  process from the passed window handle
	; ---------------------------------------------------------------------------------------------------------------------
	; parts of following code was taken from WinSpy

		WinGetClass, class, % "ahk_id " SumatraID
		If InStr(class, "SUMATRA_PDF_FRAME") {

			WinGet SumatraPID, PID, % "ahk_id " SumatraID
			Enum := ComObjGet("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" SumatraPID)._NewEnum
			If (Enum[Process])
				FileGetVersion ProgVer, % Process.ExecutablePath

			RegExMatch(ProgVer, "\d\.\d", VSumatra)

			If (SumatraCmds[VSumatra][command] = 0)
				return "" 																	;return on dummy command
			else If !SumatraCmds[VSumatra].haskey(command)
				throw "Parameter #1 [" command "] unknown in SumatraPDF version " VSumatra

			If IsObject(SumatraCmds[VSumatra][command]) {

				wmcmds := SumatraCmds[VSumatra][command]
				For i, cmd in wmcmds {
					PostMessage, 0x111, % cmd,,, % "ahk_id " SumatraID
					If (i < wmcmds.Count())
						Sleep, 300                 ; I think a little delay is necessary here
				}

			}
			else
				PostMessage, 0x111, % SumatraCmds[VSumatra][command],,, % "ahk_id " SumatraID

		}
		else {

			If RegExMatch(SumatraID, "\d\.\d", VSumatra) {

				If (SumatraCmds[VSumatra][command] = 0)
					return "" 																	;return on dummy command
				else If !SumatraCmds[VSumatra].haskey(command)
					throw "Parameter #1 [" command "] unknown in SumatraPDF version " VSumatra
				else
					return SumatraCmds[VSumatra][command]
			}
			else
				throw "Parameter #2 invalid! The passed SumatraID was neither a correct window handle nor a valid string for a program version."

		}

}

Sumatra_GetPages(SumatraID="") {                                                                                	;-- aktuelle und maximale Seiten des aktuellen Dokumentes ermitteln

	If !SumatraID
		SumatraID := WinExist("ahk_class SUMATRA_PDF_FRAME")

	ControlGetText, PageDisp, Edit3 	, % "ahk_id " SumatraID
	ControlGetText, PageMax, Static3	, % "ahk_id " SumatraID
	RegExMatch(PageMax, "\s*(?<Max>\d+)", Page)

return {"disp":PageDisp, "max":PageMax}
}

Sumatra_ToPrint(SumatraID="", Printer="") {                                                                     	;-- Druck Dialoghandler - Ausdruck auf übergebenen Drucker

		; druckt das aktuell angezeigte Dokument
		; abhängige Biblitheken: LV_ExtListView.ahk

		static sumatraprint	:= "i)[(Print)|(Drucken)]+ ahk_class i)#32770 ahk_exe i)SumatraPDF.exe"

		rxPrinter:= StrReplace(Trim(Printer), " ", "\s")
		rxPrinter:= StrReplace(rxPrinter, "(", "\(")
		rxPrinter:= StrReplace(rxPrinter, ")", "\)")

		OldMatchMode := A_TitleMatchMode
		SetTitleMatchMode, RegEx                                                              	; RegEx Fenstervergleichsmodus einstellen

		SumatraInvoke("Print", SumatraID)                                                  	; Druckdialog wird aufgerufen
		WinWait, % sumatraprint,, 6                                                             	; wartet 6 Sekunden auf das Dialogfenster
		hSumatraPrint := GetHex(WinExist(sumatraprint))                               	; 'Drucken' - Dialog handle
		ControlGet, hLV, Hwnd,, SysListview321, % "ahk_id " hSumatraPrint    	; Handle der Druckerliste (Listview) ermittlen
		sleep 200                                                                                       	; Pause um Fensteraufbau abzuwarten
		ControlGet, Items	, List  , Col1 	,, % "ahk_id " hLV                             	; Auslesen der vorhandenen Drucker
		ItemNr := 0                                                                                    	; ItemNr auf 0 setzen
		Loop, Parse, Items, `n                                                                    	; Listview Position des Standarddrucker suchen
			If RegExMatch(A_LoopField, "i)^" rxPrinter) {                                	; Standarddrucker gefunden
				ItemNr := A_Index                                                                	; nullbasierende Zählung in Listview Steuerelementen
				break
			}
		If ItemNr {                                                                                    	; Drucker in der externen Listview auswählen
			objLV := ExtListView_Initialize(sumatraprint)                                	; Initialisieren eines externen Speicherzugriff auf den Sumatra-Prozeß
			ControlFocus,, % "ahk_id " objLV.hlv                                            	; Druckerauswahl fokussieren
			err	 := ExtListView_ToggleSelection(objLV, 1, ItemNr - 1)            	; gefundenes Listview-Element (Drucker) fokussieren und selektieren
			ExtListView_DeInitialize(objLV)                                                     	; externer Speicherzugriff muss freigegeben werden
			Sleep 200
			err	:= VerifiedClick("Button13", hSumatraPrint)                           	; 'Drucken' - Button wählen
			WinWaitClose, % "ahk_id " hSumatraPrint,, 3                              	; wartet max. 3 Sek. bis der Dialog geschlossen wurde
		}

		SetTitleMatchMode, % OldMatchMode                                            	; TitleMatchMode zurückstellen

return {"DialogID":hSumatraPrint, "ItemNr":ItemNr}                                 	; für Erfolgskontrolle und eventuelle weitere Abarbeitungen
}

; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; this functions are stolen from github.com/nod5/HighlightJump
; functions not tested yet!!!
; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SumatraCreateSmxWithHeader(vFile) {											     								;-- function: create .smx with standard header data
  ; -------------------------
  ; .smx header format
  ; -------------------------
  ; # SumatraPDF: modifications to "Example File.pdf"
  ; [@meta]
  ; version = 3.2
  ; filesize = 1065035
  ; timestamp = 2020-01-30T19:23:12Z
  ;
  ; -------------------------

  SplitPath, vFile, vFilename
  FileGetSize, vBytes, % vFile
  FormatTime, vTimestamp,, yyyy-MM-ddTHH:mm:ssZ

  vSmxHeader =
  (LTrim
  # SumatraPDF: modifications to "%vFilename%"
  [@meta]
  version = 3.2
  filesize = %vBytes%
  timestamp = %vTimestamp%
  )


  FileAppend, % vSmxHeader, % vFile ".smx", UTF-8-RAW
}

SumatraSaveSmx(vFile, vSmx) {																							;-- function: save new highlight to .smx file
  If !FileExist(vFile ".smx")
    Return

  FileDelete, % vFile ".smx"
  FileAppend, % vSmx, % vFile ".smx", UTF-8-RAW

  RefreshSumatraDocument()
}

SumatraRefreshSumatraDocument() {																					;-- function: reload active SumatraPDF document, keeps current page view

	; https://github.com/sumatrapdfreader/sumatrapdf/blob/master/src/resource.h
	; IDM_REFRESH  := 406
	; WM_COMMAND   := 0x111

	vWinId := WinExist("A")
	SendMessage, 0x111, 406,0,, % "ahk_class SUMATRA_PDF_FRAME ahk_id " vWinId

}

SumatraSaveAnnotationToSmx() {																						;-- function: save annotation in active SumatraPDF document

	; https://github.com/sumatrapdfreader/sumatrapdf/blob/master/src/resource.h
	; IDM_SAVE_ANNOTATIONS_SMX  := 439
	; WM_COMMAND   := 0x111
	; Note: In SumatraPDF prerelease older than 2020-02-02: crash if no new annotation
	;       See https://github.com/sumatrapdfreader/sumatrapdf/issues/1442
	vWinId := WinExist("A")
	SendMessage, 0x111, 439,0,, % "ahk_class SUMATRA_PDF_FRAME ahk_id " vWinId
}

SumatraGotoPageDDE(vFile, vPage) {																				;-- function: go to page in active SumatraPDF window via SendMessage

	; references
	; https://github.com/sumatrapdfreader/sumatrapdf/issues/1398
	; https://gist.github.com/nod5/4d172a31a3740b147d3621e7ed9934aa
	; Control a SumatraPDF window from AutoHotkey unicode 32bit/64bit scripts
	; through DDE command text packed in SendMessage WM_COPYDATA

  WinTitle := "ahk_id " WinExist("A")

  ; Required data to tell SumatraPDF to interpret lpData as DDE command text, always 0x44646557
  dwData := 0x44646557

  ; SumatraPDF DDE command unicode text, https://www.sumatrapdfreader.org/docs/DDE-Commands.html
  ; Example: [GotoPage("C:\file.pdf", 4)]
  ; Remember to escape " in AutoHotkey expressions with ""
  lpData := "[GotoPage(""" vFile """, " vPage ")]"

  Send_WM_COPYDATA(WinTitle, dwData, lpData)
}

SumatraGetFile(ByRef vFile) {																								;-- function: get filepath to current document in active SumatraPDF window
  if vLegacyMethods ; super-global variable
    vFile := SumatraGetDocumentFilepathFromTitle()
  else
    vFile := SumatraGetDocumentFilepath()
  if vFile
    Return 1
}

SumatraGetSmx(ByRef vSmx, vFile := "") {																			;-- function: get .smx file data
If vFile
  FileRead, vSmx, %vFile%.smx
If vSmx
  Return 1
}

SumatraGetPageLen(ByRef vPage, ByRef vLen) {																	;-- function: get page number and document length from SumatraPDF go to control

  ; note: reading the control works even if toolbar is hidden
  ControlGetText, vPage, Edit1, A
  ; Get adjacent text
  ControlGetText, vRightText, Static3, A

  ; ------------------------------------------------
  ; Notes on how pagenumbers are shown in the SumatraPDF UI
  ; ------------------------------------------------
  ; - The SumatraPDF goto editbox can show virtual or real pagenumbers
  ; - Examples of virtual pagenumbers: i ii iii D1 toc 24 ...
  ; - Detect which by examining the adjacent UI text:
  ;     case1: "(6 / 102)" -> virtual number in editbox, real number is 6
  ;     case2: "/ 102"     -> real    number in editbox
  ; - Use real pagenumber to control SumatraPDF via DDE or command line
  ; - The number after the frontslash is document length in real pagenumbers
  ; ------------------------------------------------

  ; if case1: get real (not virtual) pagenumber from adjacent text
  RegExMatch(vRightText, "\((\d+)", vRealPageCheck)
  If vRealPageCheck
    vPage := vRealPageCheck1

  ; get document length: "/ 102" -> 102
  ; use pattern that handles both case1 and case2 text
  RegExMatch(vRightText, "^.*/ (\d+)(?:\D|)", vLenMatch)
  vLen := vLenMatch1

  If (vPage and vLen)
    Return 1
}

SumatraGetCanvasPosAtCursor(ByRef x, ByRef y) {																;-- function: get canvas X Y pos in pt units at mouse cursor in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_REPLY_POS = 1503
	; note: SumatraPDF returns up to 10 digits (, 32-bit signed integer, range up to 2147483648)
	; with postion data packed in format XXXXXYYYYY
	; X < 21474. Y <= 99999. Y is zero padded.

	  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
		Return
	  SendMessage, 0x111, 1503, 0,, A
	  vReturn := ErrorLevel
	  if (vReturn and vReturn != "FAIL")	  {
		;get X Y (add 0 to remove zero-padding)
		y := 0 + SubStr(vReturn, -4)    ;last 5 digits
		x := 0 + SubStr(vReturn, 1, -5) ;all except last 5 digits
	  }

}

SumatraGetPageAtCursor(ByRef vPage) {																			;-- function: get page number under mouse cursor in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_REPLY_PAGE = 1504
	; note: SumatraPDF returns 0 if cursor not over any page

  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
    Return
  SendMessage, 0x111, 1504, 0,, A
  vPage := ErrorLevel = "FAIL" ? 0 : ErrorLevel
}

SumatraClickPageGetPageNumber(vToolTipText) {																;-- function: show tooltip asking user to click page and return its page number

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp (in called functions)
	; show tooltip and wait for left mouse click to get page under mouse
	; cancel on Esc or timeout or active window change

  ToolTip, % vToolTipText
  vTick := A_TickCount
  Loop  {
    ;note: KeyWait ErrorLevel: 1 if timeout, 0 if key detected
    KeyWait, Lbutton, D T0.02
    If !ErrorLevel    {
      ToolTip
      SumatraGetPageAtCursor(vPage) ; ByRef
      Return vPage
    }
    KeyWait, Esc, D T0.02
    If !ErrorLevel or (A_TickCount > vTick + 3000) or !WinActive("ahk_class SUMATRA_PDF_FRAME")    {
      ToolTip
      return
    }
  }
}

SmartPageSelect(vToolTipText) {																							;-- function: get foreground page or else let user click to select a page

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp (in called functions)
	; vToolTipText: string to fill blank in tooltip "Click a page to _____ (Esc = cancel)"
	; "foreground" page = only one page >0% visible or only one page >50% visible
	; if no unique foreground page then let user click to select a page

  SumatraGetForegroundPage(vPage) ; ByRef
  if !vPage  {
    ; show tooltip and wait for Lbutton, cancel via Esc or timeout
    vToolTipText := "Click a page to " vToolTipText "`n(Esc = cancel)"
    vPage := ClickPageGetPageNumber(vToolTipText)
  }
  Return vPage
}

SumatraGetForegroundPage(ByRef vForegroundPage) {														;-- function: get foreground page pagenumber in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_REPLY_FOREGROUND_PAGE = 1505
	; "foreground" page = only one page >0% visible or only one page >50% visible
	; returns zero if no unique foreground page found

  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
    Return
  SendMessage, 0x111, 1505, 0,, A
  vReturn := ErrorLevel
  vForegroundPage := 0
  if (vReturn and vReturn != "FAIL")
    vForegroundPage := vReturn
}

SumatraCopySelection(ByRef vClip) {																					;-- function: copy selected text in active SumatraPDF window

	; IDM_COPY_SELECTION              420

  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
    Return
  vClipBackup := clipboardall
  Clipboard := ""
  ; copy selection
  SendMessage, 0x111, 420, 0,, A
  ; hide notification "Select content with Ctrl+left mouse button" shown if no selection
  Control, Hide, , SUMATRA_PDF_NOTIFICATION_WINDOW1, A
  vClip := Clipboard
  Clipboard := vClipBackup
  vClipBackup := ""
}

SumatraGetDocumentFilepath() {																						;-- function: get filepath for active document in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; - make SumatraPDF return data via WM_COPYDATA
	; - IDC_REPLY_FILE_PATH = 1500
	; - Works in SumatraPDF 32bit/64bit and AutoHotkey 32bit/64bit unicode, all combinations

	  DetectHiddenWindows, On
	  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
		Return
	  vWinId := WinExist("A")

  ; start listener for WM_COPYDATA that SumatraPDF will send after we first call
	OnMessage(0x4a, "Receive_WM_COPYDATA")

  ; clear super-global variable
	vFilepathReturn := ""

  ; make first call to SumatraPDF
  ; - 0x111 is WM_COMMAND
  ; - IDM_COPY_FILE_PATH = 1500
  ; - A_SCriptHwnd is Hwnd to this script's hidden window
  ;   Note: A_SCriptHwnd is hex, WinTitle ahk_id accepts both hex and dec values
	SendMessage, 0x111, 1500, A_ScriptHwnd, , % "ahk_id " vWinId

  ; todo test more if this delays operation and/or improves stability
  ; wait for message up to 250 ms
  ;if !vFilePathReturn
  ;  While (!vFilePathReturn and A_Index < 50)
  ;    sleep 5

  ; after Receive_WM_COPYDATA has reacted, stop listener
	OnMessage(0x4a, "Receive_WM_COPYDATA", 0)

  ; check for \ to ensure filepath an not only filename
	If InStr(vFilepathReturn, "\")
		return vFilepathReturn
}

SumatraRemoveAnnotationAtCursor() {																				;-- function: remove annotation in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_REMOVE_ANNOTATION = 1506
	; note: if lparam is a pagenumber then removes all annotations on that page
	;       if lparam is zero         then removes annotations under mouse cursor

	if WinActive("ahk_class SUMATRA_PDF_FRAME")
		SendMessage, 0x111, 1506, 0,, A
}

SumatraRemoveAnnotationsOnPage(vPage) {																		;-- function: remove all annotations on page vPageNum in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_REMOVE_ANNOTATION = 1506
	; note: if lparam is a pagenumber then removes all annotations on that page
	;       if lparam is zero         then removes annotations under mouse cursor

	if WinActive("ahk_class SUMATRA_PDF_FRAME")
		SendMessage, 0x111, 1506, vPage,, A
}

SumatraAnnotateSelection(vColor, vRGB := 1) {																	;-- function: color annotate selection in active SumatraPDF document

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_ANNOTATE_SEL_COLOR = 1501
	; note: SumatraPDF expects vColor to be type COLORREF
	; https://docs.microsoft.com/en-us/windows/win32/gdi/colorref
	; "COLORREF value has the following hexadecimal form: 0x00bbggrr"

  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
    Return
  if vRGB
    ; shift RGB to BGR value
    vColor := RGBColorToBGR(vColor)
  vColor := 0x00 vColor
  SendMessage, 0x111, 1501, vColor,, A
}

SumatraAnnotateDotAtCursor(vColor, vRGB := 1) {															;-- function: color annotate dot at cursor on page in active SumatraPDF window

	; dependency: SumatraPDF source code edits in Resource.h , SumatraPDF.cpp
	; IDC_ANNOTATE_DOT = 1502
	; note: the dot is an 8 x 8 pt filled rectangle
	; note: SumatraPDF expects vColor to be type COLORREF
	; https://docs.microsoft.com/en-us/windows/win32/gdi/colorref
	; "COLORREF value has the following hexadecimal form: 0x00bbggrr"

  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
    Return
  if vRGB
    ; shift RGB to BGR value
    vColor := RGBColorToBGR(vColor)
  vColor := 0x00 vColor
  SendMessage, 0x111, 1502, vColor,, A
}

/*  notes on SumatraPDF position helper notification

	 ------------------------------------------
	 Shortcut "m" shows position helper notification in top left document corner.
	 Reports X Y document canvas position that the mouse is over.
	   Example: top left document corner reports 0 0 in fixed page mode.
	 Press "m" again to cycle notification units: pt -> mm -> in
	 .smx files use unit pt for highlight rect corner data.
	 Once notification exists we can hide it and still read with ControlGetText.
	 "Esc" closes the notification even when hidden.

	 Notification text format:
	 "Cursor position: 51,0 x 273,5 pt"
	 "Cursor position: 18,0 x 96,4 mm"
	 "Cursor position: 0,71 x 3,8 in"

	 - Do not use the ":" in regex because not present in all translations
	 - Fraction separator can be "," or "." depending on translation/locale
	     Example: Dutch language setting: "Cursor positie 0,0 x 4,77 in"
	 - Thousands separator can be (comma dot space) depending on translation/locale
	 - The pt/mm/in unit string is not translated so can be used as pattern.
	 - Check for "pt" via SubStr on last two characters because
	     InStr could give false positive on "pt" in preceding text in some language.
	 - SumatraPDF source: SumatraPDF.cpp function UpdateCursorPositionHelper
	 - https://github.com/nod5/HighlightJump/issues/6
	 ------------------------------------------

*/

SumatraPrepareForCanvasPosCheck() {																				;-- function: prepare SumatraPDF for document pt position check

	; Returns 1 if prepared

	; check if popup exists (hidden or visible)
		ControlGetText, vPos, SUMATRA_PDF_NOTIFICATION_WINDOW1, A

  ; if not exist then send "m" to create position helper notification
  ; note: SumatraPDF changes the mouse cursor to a cross while the notification exists
  ; Check for "pt" via SubStr on last two characters
  If !(SubStr(vPos, -1) = "pt")  {
    Loop, 3    {
      ; show popup
      Send m
      ; hide popup (we can still read it and toggle its unit)
      Control, Hide, , SUMATRA_PDF_NOTIFICATION_WINDOW1, A
      ; read popup
      ControlGetText, vPos, SUMATRA_PDF_NOTIFICATION_WINDOW1, A
      If (SubStr(vPos, -1) = "pt")
        Break
      Sleep 30
    }
  }

  If (SubStr(vPos, -1) = "pt")
    Return 1
}

SumatraGetCanvasPosFromNotification(ByRef x, Byref y) {													;-- function: Read SumatraPDF notification to get mouse position in canvas pt units

  ; get mouse position in SumatraPDF canvas in pt units (one decimal)
	ControlGetText, vPos, SUMATRA_PDF_NOTIFICATION_WINDOW1, A

  ; extract X Y pos with decimals and round later

  ; - Do not use the ":" in regex because not present in all translations
  ; - Fraction separator can be "," or "." depending on translation/locale
  ;     Example: Dutch language setting: "Cursor positie 0,0 x 4,77 in"
  ; - Thousands separator can be (comma dot space) depending on translation/locale
  ; - The pt/mm/in unit string is not translated so can be used as pattern.
  ; - Check for "pt" via SubStr on last two characters because
  ;     InStr could give false positive on "pt" in preceding text in some language.
  ; - SumatraPDF source: SumatraPDF.cpp function UpdateCursorPositionHelper
  ; - https://github.com/nod5/HighlightJump/issues/6

  ; Hybrid test string that covers all character variants
  ; ": 4.401 000,3 x 341,000.3 pt"
  ; "a 4.401 000,3 x 341,000.3 pt"
  ; regex pattern match
	  vPattern := "\D ([\d \.,]+)[\.,](\d+) x ([\d \.,]+)[\.,](\d+) pt$"
	  RegExMatch(vPos, vPattern, vPos)

  ; remove separators to get integers
	vPos1 := RegExReplace(vPos1, "[ \.,]", "")
	vPos3 := RegExReplace(vPos3, "[ \.,]", "")

  ; concatenate (integer)(dot-separator)(fraction)
	x := vPos1 "." vPos2
	y := vPos3 "." vPos4

}

SumatraGetDocumentFilepathFromTitle() {																			;-- function: Get document filepath by parsing window title

	; Dependency: SumatraPDF > Advanced Options > FullPathInTitle = true

	WinGetTitle, vTitle, A

  ; ---------------------------------------------
  ; SumatraPDF window title format with advanced setting "FullPathInTitle = true"
  ; ---------------------------------------------
  ; format1 "<filepath> - [<metadata document title>] - SumatraPDF"
  ; format2 "<filepath> - SumatraPDF"
  ; --------------------------------------------
  ; note: in format1 both filepath and metadata can use characters "-"  "["  and  " "
  ; That enables edge cases like "C:\a.pdf - [.pdf - [x] - SumatraPDF"
  ; which has two possible solutions
  ; 1 "C:\a.pdf"          with metadata ".pdf - [x"
  ; 2 "C:\a.pdf - [.pdf"  with metadata "x"
  ; and more complex cases with even more solutions.
  ; In format2 there is always one solution.

  ; Detect format1 or format2
  If (SubStr(vTitle, -13) = "] - SumatraPDF")  {
    ; format1: try each instance of " - [" until a file exist
    ; probably good enough in most circumstances
    Loop, 20    {
      vFilepathLen := InStr(vTitle, " - [", , , A_Index) - 1
      vFile := SubStr(vTitle, 1, vFilepathLen)
      if FileExist(vFile)
        break
    }
  }
  Else  {
    ; format2: find last instance of " - SumatraPDF"
    vFilepathLen := InStr(vTitle, " - SumatraPDF") - 1
    vFile := SubStr(vTitle, 1, vFilepathLen)
  }

  ; Check for :\ and existance to ensure string is filepath and not only filename
  If InStr(vFile, ":\") and FileExist(vFile)
    Return vFile
}






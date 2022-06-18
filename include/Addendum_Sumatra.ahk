;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;       	RPA (robotic prozess automation) FUNKTIONEN für
;   	    SUMATRA PDF READER
;
;			ABHÄNGIGKEITEN:
;                                                                                  	------------------------
;          	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 04.02.2021 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MsgBox, % SumatraGetDocumentFilepath()
return

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

  SumatraRefreshDocument()
}

SumatraRefreshDocument() {							               														;-- function: reload active SumatraPDF document, keeps current page view

	; https://github.com/sumatrapdfreader/sumatrapdf/blob/master/src/resource.h
	; IDM_REFRESH  := 406
	; WM_COMMAND   := 0x111

	vWinId := WinExist("A")
	SendMessage, 0x111, 406,0,, % "ahk_class SUMATRA_PDF_FRAME1 ahk_id " vWinId

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

  if !WinActive("ahk_class SUMATRA_PDF_FRAME")
    Return
  vWinId := WinExist("A")
  DetectHiddenWindows, On
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

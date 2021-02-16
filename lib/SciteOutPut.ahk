SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {     	; modified version for Addendum für Albis on Windows

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
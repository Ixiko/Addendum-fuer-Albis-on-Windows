SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {     	; modified version for Addendum

	; last change 30.07.2020

		static LinesOut           	:= 0,
				SCI_GETLENGTH	:= 2006,
				SCI_GOTOPOS		:= 2025

		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

		; move Caret to end of output window
		ControlSend, Scintilla2, {LControl Down}{End}{LControl Up} , % "ahk_id " SciObj.SciTEHandle

		If InStr(Text, "ShowLinesOut") {
			SciObj.Output("SciteOutput function has printed " LinesOut " lines.`n")
			return
		}

		If Clear || ((StrLen(Text) = 0) && (LinesOut = 0))
			SendMessage, SciObj.Message(0x111, 420)                   	;If clear=1 Clear output window

		If (StrLen(Text) != 0) {
			SciObj.Output(Text (LineBreak ? "`r`n": ""))                            	;send text to SciTE output pane
			LinesOut += StrSplit(Text, "`n").MaxIndex()
		}

		If Exit {
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

}
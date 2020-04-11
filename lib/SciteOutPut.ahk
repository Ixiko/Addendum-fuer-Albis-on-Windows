SciTEOutput(Text, Clear:=1, LineBreak:=1, Exit:=0) {
		SciObj := ComObjActive("SciTE4AHK.Application")             	;get pointer to active SciTE window
		If (Clear=1)
				SendMessage, SciObj.Message(0x111, 420)               	;If clear=1 Clear output window
		If (LineBreak=1)
				Text := "`r`n" Text                                                      	;If LineBreak=1 prepend text with `r`n
		SciObj.Output(Text)                                                           	;send text to SciTE output pane
		If (Exit=1)
		{
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}
}
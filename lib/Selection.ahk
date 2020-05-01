GetSelection(PlainText = 1, RestoreCB = 0) {

	/*! Function: GetSelection([PlainText])

		Link: https://github.com/number1nub/MouseGestures/blob/8ae7afd93af29bf0693e002ad19253269d92de36/lib/MGR_UDF.ahk
		Get the user's selection and converts it to text format

		Parameters:
		PlainText - (Optional) This flag determines if the return value will be
		treated as plain text or returned in the standard format. Defaults to true &
		returns plain text unless either 0, false or an empty string ("") is specified
		RestoreCB - (Optional) Flag that sets whether or not to restore the clipboard to its
		original content. Default is 0 ([b]don't[/b] restore clipboard)

		Returns:
		Returns the selection (as plaintext or default formatting, depending on value of the [b]PlainText[/b] flag
		passed (default is plain text).

		Throws:
		Throws an exception if an error is encountered (i.e. unable to copy selection to

	*/

	cbBU := ClipboardAll	                                    	; Save the previous clipboard
	Clipboard := ""				                                    	; Start off empty to allow ClipWait to detect when the text has arrived
	SendInput, {Blind}^c	                                    	; Simulate the Copy hotkey (Ctrl+C)
	ClipWait, 1, 1				                                    	; Wait 2 seconds for the clipboard to contain text
	If ErrorLevel
		throw { what: "GetSelection", message: "Unable to copy selection to clipboard" }
	If !(PlainText)
		return, ClipboardAll
	Selection := Clipboard	                                		; Put the clipboard in the variable Selection
	Clipboard := RestoreCB ? cbBU : Clipboard			; Restore the previous clipboard

return Selection		                                            		; return the selection
}

SaveSelection(fPath:="") {

	/*! Function: SaveSelection()

		Saves the user's selection in a txt file

		Parameters:
		fPath - (Optional) The path of the file in which to save the selection. If blank
					or omitted, a file save dialog will be displayed to the user.
	*/

	try {
		sel := GetSelection()
		FileSelectFile, fPath, S 24,, Where should the file be saved?, Text Files (*.txt)
		FileDelete, % fPath
		FileAppend, % sel, % fPath
	}
	catch e {
		errAction := InStr(e.what, "GetSelection") ? "Get the current selection."
				 : ((InStr(e.what, "SelectFile") ? "Invalid file selection:"
				   : InStr(e.what, "Delete") ? "Delete the specified file:"
				   : InStr(e.what, "Append") ? "Write the selection to the file:") "`n`t""" fPath """")
		m("Oops..`n", "An error occurred while trying to " errAction, "!")
	}

}

mgr_MonitorGesture() { ; Monitor the mouse directions to get the gesture

	While GetKeyState(mainHotkey, "P") {
		MouseGetPos, x1, y1
		While GetKeyState(mainHotkey, "P") {
			Sleep, 10
			MouseGetPos, x2, y2
			if (Sqrt((x2-x1)**2+(y2-y1)**2)>=15)				                        	    	      	; if the module is greater or equal than 35,
			{
				Direction := mgr_GetDirection(x2-x1, y2-y1)	                        				;	Get hotkey modifiers & the mouse movement direction
				x1 := x2 , y1 := y2										                                    	;	Update the origin point
				if (Direction && LastDirection && Direction<>LastDirection)                	;	if the direction has changed,
					Break												                                        		;		get the next direction
				Gesture	:= GetMods() mgr_RemoveDups(Directions "-" Direction, "-")	; Set the gesture with the different directions
				Command	:= mgr_GetCommand(Gesture, 1)						            	;	if there is a description, get the description instead of the command
				if (Gesture && Gesture<>LastGesture) 					                        		;	if the gesture has changed,
					if (showTT) {
						ToolTip, % Command ? Command : Gesture				            		;		display the command else display the gesture
						LastGesture := Gesture						                        					;	Usefull to know if the gesture has changed
						SetTimer, RemoveTips_tmr, 1000						                			;	Remove tooltips (+ traytips) after 1.5 seconds
					}
			}
			LastDirection := Direction	; Usefull to know if the direction has changed
		}
		Directions .= "-" LastDirection
		LastDirection := Direction	:= ""
	}

return Gesture
}

mgr_GetDirection(X_Offset, Y_Offset) {

	; https://github.com/number1nub/MouseGestures/blob/8ae7afd93af29bf0693e002ad19253269d92de36/lib/mgr%20GetDirection.ahk

	static dirList := ["R", "UR", "U", "UL", "L", "DL", "D", "DR", "R"]

	Module    	:= Sqrt((X_Offset**2) + (Y_Offset**2))	    	; Distance between the center and the mouse cursor
	Argument	:= ACos(X_Offset/Module) * (45/ATan(1))		; Angle between the mouse and the X-axis from the center
	Argument	:= Y_Offset<0 ? Argument : 360-Argument	; (Screen Y-axis is inverted)
	Direction 	:= Ceil((Argument-22.5)/45)			            	; Converts the argument into a slice number
	;~ Direction := Direction=0 ? "R" : Direction=1 ? "UR" : Direction=2 ? "U" : Direction=3 ? "UL" : Direction=4 ? "L" : Direction=5 ? "DL" : Direction=6 ? "D" : Direction=7 ? "DR" : Direction=8 ? "R" : ""

return dirList[Direction+1]
}
; Workerthread für QuickExporter - sucht nach falsch abgelegten Dateien

#NoEnv
SetBatchLines, -1
CoordMode, ToolTip, Screen

;~ for each, value in A_Args
	;~ t .= each " = " value "`n"
;~ MsgBox, % "t: " t "`ncritObj: " critObj.1

critObj:= CriticalObject(A_Args.1)
RemoveWrongFiles(critObj.1)
Sleep 500
ExitApp

RemoveWrongFiles(stdPath) {

	rmfFunc := "running"
	wrongfiles := 0
	xDataPaths := {}

	WinGetPos, wx, wy, ww, wh, % "Addendum - Quickexporter ahk_class AutoHotkeyGUI"

	If !FileExist(A_ScriptDir "\WrongFilesDeleted.txt")
		FileAppend, % "[" stdPath "]"  , % A_ScriptDir "\WrongFilesDeleted.txt", UTF-8

	Loop, Files, % stdPath "\*.*", R
	{
		fileIndex := A_Index
		If (Mod(A_Index, 100) = 0)
			ToolTip, % "[" SubStr("000000" A_Index, -5) "] - wrongfiles found: " wrongfiles "`n" StrReplace(A_LoopFileFullPath, stdPath "\") , % wx+ww-300, % wy+20, 3

		If RegExMatch(A_LoopFileName, "i)^\s*Laborbefunde_\d+\s+\-.*\.pdf") {
			wrongfiles ++
			FileAppend, % "`n[" SubStr("00000" wrongfiles, -4) "] " StrReplace(A_LoopFileFullPath, stdPath "\")  , % A_ScriptDir "\WrongFilesDeleted.txt", UTF-8
			FileDelete, % A_LoopFileFullPath
		}

		if !InStr(A_LoopFilePath, stdPath "\xData")
			if InStr(A_LoopFileFullPath, "\xData") {
				xDataPath := RegExReplace(A_LoopFileFullPath, "i)\\xData.+", "\xData")
				if !xDataPaths.haskey(xDataPath) {
					xDataPaths[xDataPath] := 1
					wrongfiles ++
					FileAppend, % "`n[" SubStr("00000" wrongfiles, -4) "] " StrReplace(A_LoopFileFullPath, stdPath "\")  , % A_ScriptDir "\WrongFilesDeleted.txt", UTF-8
				}
		}

	}

	For xDataPath, count in xDataPaths {
		FileRemoveDir, % xDataPath, 1
		SciTEOutput("entfernt: " StrReplace(xDataPath, stdPath "\"))
	}


	ToolTip,,,, 3
	rmfFunc := "ready"
	;~ SciTEOutput("state: " rmfFunc "`nexamined: " fileIndex " files")
} ;RWF-End


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
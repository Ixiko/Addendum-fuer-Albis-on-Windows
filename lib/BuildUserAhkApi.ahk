BuildUserAhkApi(AhkScriptPaths, OverwriteAhkApi:="1", RecurseIncludes:="1", Labels:="1", WrapWidth:="265", AhkApiPath:="C:\Users\User\Documents\AutoHotkey\SciTE\user2.ahk.api", RecursionCall:="0"){
	; Written by XeroByte
	; Generates the User.ahk.api file to add custom function & label intellisense!
	; to initate: use BuildUserAhkApi(A_ScriptFullPath,1) from the main script
	; Requires: grep() by Titan/Polyethelene https://github.com/camerb/AHKs/blob/master/thirdParty/titan/grep.ahk
	; Requires: RegExMatchGlobal() by Just Me - https://autohotkey.com/board/topic/14817-grep-global-regular-expression-match/page-2
	; Requires: tooltip(),
	; Requires: isDir(),
	; Optionally include TF_Wrap() from the TF Library https://github.com/hi5/TF/blob/master/tf.ahk

	TF_Wrap 				:= "TF_Wrap" ; This is for a workaround to make the script still work even without the TF library - need to call the function dynamically.
	FileExclusionList 	:= "Socket.ahk|DirMenu2-include.ahk"
	FuncExclusionList 	:= "if|While|VarZ_Save|__New|__Delete|__Call|__Get|__Set"
	LabelExclusionList 	:= "FindReplaceErr|MoveCurser"
	if(!IsObject(AhkScriptPaths))
		AhkScriptPaths := Array(AhkScriptPaths)
	for index, AhkScriptPath in AhkScriptPaths
	{
		if(isDir(AhkScriptPath)){
			Loop, %AhkScriptPath%\*.ahk,0,1
			{
				if(InStr("|" FileExclusionList "|", "|" A_LoopFileName "|",false)=0)
					AhkApiText .= BuildUserAhkApi(A_LoopFileLongPath, 0, RecurseIncludes, Labels, WrapWidth, AhkApiPath, 1)
			}
		}else{
			SplitPath,AhkScriptPath,AhkScriptFname
			if(InStr("|" FileExclusionList "|", "|" AhkScriptFname "|",false)=0)
			{
				FileRead, ThisScriptTxt, %AhkScriptPath%
				; Retrieve Functions
					MatchCollection := RegExMatchGlobal(ThisScriptTxt, "m)(?P<PreComment>(?:(?:(?:^[ \t]*;.*)+(?:\r|\n)+)|^[ \t]*\/\*(?:(?!^[ \t]*\*\/)(?:.|\r|\n))+^[ \t]*\*\/(?:\r|\n)+)*)^[ \t]*(?P<FuncDefinition>(?P<FuncName>[a-zA-Z0-9_-]*)\((?P<FuncParams>.*)\))\s*(?P<Comment1>(?:[ \t]*;.*(?:\r|\n)+)*)\{[ \t]*(?P<Comment2>(?:[ \t]*;.*)?(?:(?:(?:\r|\n)+^[ \t]*;.*)|(?:\r|\n)+^[ \t]*\/\*(?:(?!^[ \t]*\*\/)(?:.|\r|\n))+^[ \t]*\*\/)*)")
					For k, v in MatchCollection
					{
						if(InStr("|" FuncExclusionList "|", "|" v[3] "|",0)>0)
							continue
						CustomFunc := v[3] " (" v[4] ")`r`n" v[1] "`r`n" v[5] "`r`n" v[6]
						CustomFunc := StrReplace(CustomFunc,"\","\\")
						CustomFunc := RegExReplace(CustomFunc,"S)(^\s*|\s*$)")
						CustomFunc := RegExReplace(CustomFunc, "^\s*[\r\n]+","")
						if(IsFunc("TF_Wrap"))
							CustomFunc := %TF_Wrap%(CustomFunc, WrapWidth)
						CustomFunc := RegExReplace(CustomFunc,"S)\R+", "\n")
						CustomFunc := RegExReplace(CustomFunc,"S)\t", " ")
						CustomFunc .= "\n; Location: " . RegExReplace(AhkScriptPath,"\\","\\")
						CustomFunc := RegExReplace(CustomFunc,"S)(?:\\n)+ *", "\n")
						;~ CustomFunc := RegExReplace(CustomFunc,"S)(?:\\t)+ *", "\t")
						AhkApiText .= CustomFunc "`n"
					}
					AhkApiText .= "`n"
				; Retrieve Labels
				if(Labels){
			grep(ThisScriptTxt, "mS)^\s*[a-zA-Z0-9_-]+\:\s(\s*;.*?$)*", MatchCollection,1,0,"§")
			loop, parse, MatchCollection, §
					{
						if(RegExMatch(A_LoopField,"iS)^\s*(" LabelExclusionList ")\:\s")>0)
							continue
						StringReplace, CustomFunc, A_LoopField, \, \\, All
						CustomFunc := RegExReplace(CustomFunc,"mS)(^\s*|\s*$)")
						if(IsFunc("TF_Wrap"))
							CustomFunc := %TF_Wrap%(CustomFunc, WrapWidth)
						CustomFunc := RegExReplace(CustomFunc,"S)\R", "\n")
						CustomFunc := RegExReplace(CustomFunc,"S)\t", "\t")
						StringReplace, CustomFunc, CustomFunc, :, %A_Space%
						AhkApiText .= CustomFunc . "\n; Location: " . RegExReplace(AhkScriptPath,"\\","\\") . "`n"
					}
					AhkApiText .= "`n"
				}
				; Recurse into includes
				if(RecurseIncludes){
			grep(ThisScriptTxt, "imS)^\s*\#include (.*)\s*$", IncludeCollection,1,1,"§")
					AhkScriptFolder := substr(AhkScriptPath,1,InStr(AhkScriptPath,"\",0,0))
			loop, parse, IncludeCollection, §
						AhkApiText .= BuildUserAhkApi((instr(A_LoopField,":")) ? A_LoopField : AhkScriptFolder . A_LoopField, 0, 1, Labels, WrapWidth, AhkApiPath, 1)
				}
			}
		}
	}
	; If this function call is the original function call (and not an autocall by the recursion).
	if(!RecursionCall){
		if(OverwriteAhkApi)
			FileDelete, %AhkApiPath%
		AhkApiText := RegExReplace(AhkApiText, "(\R){2,}","$1$1")
		FileAppend, %AhkApiText%, %AhkApiPath%
		ToolTip("Imported Custom Funcs & Labels from`n" . AhkScriptPath)
	}
	return AhkApiText
}

grep(h, n, ByRef v, s = 1, e = 0, d = "") {
	v =
	StringReplace, h, h, %d%, , All
	Loop
		If s := RegExMatch(h, n, c, s)
			p .= d . s, s += StrLen(c), v .= d . (e ? c%e% : c)
		Else Return, SubStr(p, 2), v := SubStr(v, 2)
}

ToolTip(Message, TimeToDisplay = 500, SleepWhileDisplayed = true){
	; ToolTip, Text, X, Y, WhichToolTip
	ToolTip,%Message% ; display message
	SetTimer,ToolTipClear, Off
	SetTimer,ToolTipClear,-%TimeToDisplay% ; clear tooltip after TimeToDisplay milliseconds
	If SleepWhileDisplayed
		Sleep,%TimeToDisplay% ; sleep before returning
	Return

	ToolTipClear: ; clear tooltip
	ToolTip
	Return
}

isDir(FilePattern){
	att := FileExist(FilePattern)
	return att = "" ? "" : InStr(att, "D") = 0 ? 0 : 1
}

RegExMatchGlobal(ByRef Haystack, NeedleRegEx) {
   Static Options := "U)^[imsxACDJOPSUX`a`n`r]+\)"
   NeedleRegEx := (RegExMatch(NeedleRegEx, Options, Opt) ? (InStr(Opt, "O", 1) ? "" : "O") : "O)") . NeedleRegEx
   Match := {Len: {0: 0}}, Matches := [], FoundPos := 1
   While (FoundPos := RegExMatch(Haystack, NeedleRegEx, Match, FoundPos + Match.Len[0]))
      Matches[A_Index] := Match
   Return Matches
}
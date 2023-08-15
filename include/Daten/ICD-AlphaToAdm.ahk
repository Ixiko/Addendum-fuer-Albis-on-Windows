; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;         Skript für die Konvertierung von DIMDI ICD-10-GM Dateien in ein für Addendum für Albis on Windwos  verwendbares Format
;
;                             by Ixiko started in September 2017 - last change 01.01.2023 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
	#NoEnv
	#Persistent
	#SingleInstance             	, Off
	#KeyHistory                     	, Off

	SetBatchLines                  	, -1
	ListLines                        	, Off
	AutoTrim                       	, On
	FileEncoding                 	, UTF-8

	;~ ICDTextFile := "icd10gm2022alpha_edvtxt_20211001_20220114.txt"
	ICDTextFile := "icd10gm2023alpha_edvtxt_20220930.txt"

	sourceDate := 0
	files := ReadDirectory(A_ScriptDir "\*.txt", "R")
	For each, file in files
		If RegExMatch(file, "i)(?<fname>icd10gm" A_YYYY "alpha_edvtxt)_(?<fDate>\d+)\.txt", source) {
			If (sourceDate < sourcefDate)
				sourcefile := file, sourcename := sourcefname sourcefDate,	sourceDate := sourcefDate
		}

	targetfile := "icd10gm" A_YYYY
	If !FileExist(sourcefile) {
		MsgBox, 0x1, % ScriptName, % "Eine aktuell zu konvertierende Datei ist nicht vorhanden`n"
													.  "\icd10gm" A_YYYY "\icd10gm" A_YYYY "alpha_edvtxt_*.txt`nDas Skript wird beendet.", 20
		ExitApp
	}

	ICDText := FileOpen(sourcefile, "r", "UTF-8").Read()
	icdfilelines := StrSplit(ICDText, "`n", "`r")
	Linescount := icdfilelines.Count()
	MsgBox, 0x4, % ScriptName, % "Soll diese Datei konvertiert werden?`n"
												.	"::: \icd10gm" A_YYYY  "\" sourcename " :::`n"
												.	"[" Linescount  " Zeilen]`n"
												.	"Zielverzeichnis\Dateiname: \" StrReplace(targetfile, A_ScriptDir) ".txt"
	IfMsgBox, No
		ExitApp

	file := FileOpen(A_ScriptDir "\" targetfile ".txt", "w", "UTF-8")
	ToolTip, % SubStr("              ", -1*(StrLen(Linescount)-1)) "/" Linescount
	For lineNr, line in icdfileLines {

		If (Mod(lineNr, 500) = 0)
			ToolTip, % SubStr("             ", -1*(StrLen(lineNr)-1)) "/" Linescount

		icd := StrSplit(line, "|")

		If icd.4 {

			RegExMatch(icd.4, "[\!\+]", notation)
			ICDCode 	:= notation ? notation . StrReplace(icd.4, notation) : icd.4
			ICDZusatz	:= icd.5 ? " (" icd.5 ")" : ""
			file.WriteLine(icd.8 . ICDZusatz . " {" ICDCode "}")

		}

	}

	file.close()

	MsgBox, 0x1, % ScriptName, % "Konvertierung abgeschlossen.`n \" targetfile "\" sourcename " :::`n"

ExitApp

ReadDirectory(FilePattern, Mode:="R", ExtraFilter:="") {
	files := Array()
	Loop, Files, % FilePattern , % Mode
		files.Push(A_LoopFileFullPath)
return files
}
;ABBYYScanHelper.ahk


;{ 1. Scripteinstellungen / Includes

	#NoEnv
	#SingleInstance force
	#Persistent
	;#NoTrayIcon

	SendMode Input
	SetWorkingDir %A_ScriptDir%
	SetTitleMatchMode, 2
	DetectHiddenWindows, On
	DetectHiddenText, On
	SetControlDelay -1
	SetWinDelay, -1
	SetBatchLines -1
	CoordMode, Mouse, Screen
	CoordMode, Pixel, Screen
	CoordMode, Menu, Screen
	CoordMode, Caret, Screen
	CoordMode, Tooltip, Screen

	#Include %A_LineFile%\..\..\..\include\Praxomat_Functions.ahk
	#Include %A_LineFile%\..\..\..\include\FindText.ahk
	#include %A_LineFile%\..\..\..\include\GDIP_All.ahk
	#include %A_LineFile%\..\..\..\include\Gui_PraxTT.ahk

	OnExit, WeHaveIT

	global AddendumDir
	AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")

	CompName:= A_ComputerName
	StringReplace, Compname, CompName, -,, All

;}

		;WinWait, ABBYY FineReader ahk_class FineReaderSprint,, 6
		hFR= %1%
	If hFR {
		PraxTT("Okay! Es geht los mit dem Scannen!", "4 13 B")
			sleep, 1000
		PraxTT("3!", "4 18 S")
			sleep, 1000
		PraxTT("2!", "4 18 S")
			sleep, 1000
		PraxTT("1!", "4 18 S")
			sleep, 1000
		PraxTT("Start!", "4 18 S")
	}else{
		PraxTT("Habe kein Fensterhandle vom FineReader bekommen.`nIch muss jetzt abbrechen!", "4 3")
			sleep, 4000
		ExitApp
	}


	hScanWin:= FindWindow("", "#3277011", "", "ABBYY FineReader 12 Sprint", "FineReaderSprint12MainWindowClass", "off", "off")
	MsgBox, % hScanWin "`n" SureControlClick("Dokument scannen","ABBYY FineReader",,hScanWin)

ExitApp

PraxTTOff: ;{
	AnimateWindow(hPraxTT, PraxTTDuration, BLEND)
	Gui, PraxTT: Destroy
return
;}


WeHaveIT:


ExitApp
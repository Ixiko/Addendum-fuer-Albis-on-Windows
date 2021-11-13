;------------------------------------------------------------------------------------------------ Modul Privatabrechnung versenden ------------------------------------------------------------------------------------------
;------------------------------------------------------------------------------------------------ Addendum für ALBIS on WINDOWS -----------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------- written by Ixiko -this version is from 30.06.2018 -----------------------------------------------------------------------------------
;------------------------------------------------------------------------------ please report errors and suggestions to me: Ixiko@mailbox.org ---------------------------------------------------------------------------
;--------------------------------------------------------------------------- use subject: "Addendum" so that you don't end up in the spam folder ------------------------------------------------------------------------
;------------------------------------------------------------------------------------ GNU Lizenz - can be found in main directory  - 2017 ---------------------------------------------------------------------------------
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

;{1. Scripteinstellungen

	#NoEnv
	#SingleInstance
	SetBatchlines, -1
	SetWinDelay, -1
	SetTitleMatchMode, 2
	DetectHiddenWindows, On
	DetectHiddenText, On
	FileEncoding, UTF-8
	OnExit, JetztAberRaus

	global code:= []
	debug = true

	If (debug) {
		FileRead, scriptfile, %A_ScriptFullPath%
		Loop, Parse, scriptfile, `n
		{
			code[A_Index]:= A_LoopField
		}
	}

	;{--Minimieren des SciteWindow falls das Skript zu Testzwecken aus Scite gestartet wird
		IfWinExist, ahk_class SciTEWindow
		{
					WinActivate
					 WinMinimize
		}
	;}

;}

;{2. Includes

	#include %A_ScriptDir%\..\..\include\AddendumFunctions.ahk

;}

;{3. Variablen und Konstanten

  	Global AddendumDir, AlbisWinID
	AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")

;}


If AlbisWinID:=WinExist("ahk_class OptoAppClass") {

	AlbisWaitAndActivate("ahk_class OptoAppClass", 0, 0)
	AlbisLogout()

}


return


JetztAberRaus:


	IfWinExist, ahk_class SciTEWindow
	{
				WinActivate, ahk_class SciTEWindow
					WinMaximize,  ahk_class SciTEWindow
	}

ExitApp

AlbisLogout() {

	SendMessage, 0x111, 32922,, ahk_id %AlbisWinID%

}

AlbisAkteSchließen(PatName) {																							;schließt im Moment nur die aktive Akte

	SendMessage, 0x111, 57602,, ahk_id %AlbisWinID%

}
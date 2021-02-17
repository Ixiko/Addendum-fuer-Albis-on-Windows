;###############################################################
;------------------------------------------------------------------------------------------------------------------------------------
;----------------------------------------------- Addendum für AlbisOnWindows ---------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------------------------- Modul: ScanPool_Adept----------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;-------                                 Hilfsmodul zur Kontrolle der richtigen Dateibezeichnung                               -------
;-------                                     liest Pdf Text aus oder startet ein OCR Programm                    		           	-------
;-------		ist gedacht am besten auf dem Server zu laufen und dort den Befundordner zu überwachen 	-------
;------------------------------------------------------------------------------------------------------------------------------------
Version:= "0.1 alpha", lastChange:= "02.09.2018"
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------------------------------- written by Ixiko  -------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;------------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ------------------------
;---------------------------- use subject: "Addendum" so that you don't end up in the spam folder ---------------------
;------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ----------------------------------
;------------------------------------------------------------------------------------------------------------------------------------
;###############################################################

;{ 1. Scripteinstellungen / Includes / Tray Menu

	#NoEnv
	#SingleInstance force
	#Persistent
	#MaxMem 4095	; INCREASE MAXIMUM MEMORY ALLOWED FOR EACH VARIABLE - NECESSARY FOR THE INDEXING VARIABLE

	SendMode Input
	;SetWorkingDir %A_ScriptDir%
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
	FileEncoding, UTF-8


	;Tray Menu erstellen
		Menu, Tray, NoStandard
		Menu, Tray, Tip, % "ScanPool_Adept " Version "`n-----------------------`n         Addendum   `nfür Albis on Windows"
		Menu, Tray, Add,
		Menu, Tray, Add, Skript Neu Starten, SkriptReload
		Menu, Tray, Add, Info anzeigen, ShowGui
		Menu, Tray, Add, Beenden, AdeptExit

		Menu Tray, Icon, %A_ScriptDir%\ScanPool.ico

;	;Das Scite Fenster wird wieder maximiert, wenn der Scite-Editor im Hintergrund läuft. Hilft damit man schneller im Skript Änderungen vornehmen kann.
/*
		IfWinExist, ahk_class SciTEWindow
		{
					WinActivate, ahk_class SciTEWindow
					WinMinimize, ahk_class SciTEWindow
		}
	;}
*/


	OnExit, AdeptExit


;}

;{ 2. Variblen Setup / Registry auslesen

	;-------------------------------------------------------------------- globale Variablen -----------------------------------------------------------------------
	;------------------------------------------------------------------------------------------------------------------------------------------------------------------
	global AddendumDir, AlbisWinID 					;ACHTUNG: diese sollten in jedem Skript global sein
	global P, PAdd												;Variablen für ein Prozess Gui um den Fortschritt zu Demonstrieren
	global BefundOrdner									;BefundOrdner - der Ordner in dem sich die ungelesenen PdfDateien befinden
	global files													;files-Array - enthält sämtliche benötigte Daten und kann als Datei gesichert werden
	global ExecPID												;Consolen PID
	global pdfError												;zählt Lesefehler von PDF Dateien
	global PageSum											;die Gesamtzahl an Pdf Seiten
	global FileCount											;Anzahl der Pdf Dokumente im Ordner
	global PdfDataPath										;Verzeichnis der Pdf Daten
	global PdfIndexPath										;Pfad und Dateiname des Index
	global hProgress

	;------------------------------------------------------- Ermitteln der beiden wichtigsten Pfade ------ ------------------------------------------------------
		AddendumDir := RegReadUniCode64("HKEY_LOCAL_MACHINE", "SOFTWARE\Addendum für AlbisOnWindows", "ApplicationDir")
		PdfDataPath:= A_ScriptDir . "`\PdfData\"
		PdfIndexPath:= A_ScriptDir . "`\PdfData\_BefundOrdner.xml"
	;-------------------------------------------------------- diverse Variablen setzen oder definieren ----------------------------------------------------------
	;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Pat:=[], arr:=[], chfiles:= [], addfiles:=[], Name:=[], winHandles:=[], DocIndex:=0, pdfError:=0
		files:= Object()

	;------------------------------------------------- ScanPool BefundOrdner/PDFReader/SignaturRechte ----------------------------------------------------
		IniRead, BefundOrdner, %AddendumDir%\Praxomat.ini, ScanPool, BefundOrdner
		IniRead, PDFReaderName, %AddendumDir%\Praxomat.ini, ScanPool, PDFReaderName
		IniRead, PDFReaderFullPath, %AddendumDir%\Praxomat.ini, ScanPool, PDFReaderFullPath
		IniRead, PdfMachinePath, %AddendumDir%\Praxomat.ini, ScanPool, PdfMachinePath

	;----------------------------------------------------------------- Client PC Name finden -----------------------------------------------------------------------
		CompName:= A_ComputerName
		StringReplace, Compname, CompName, -,, All

	;---------------------------------------------------- ScanPool Gui Einstellungen aus der INI einlesen -------------------------------------------------------
		IniRead, BOFontOptions, %AddendumDir%\Praxomat.ini, ScanPool, BOFontOptions, S8 CDefault q5
		IniRead, BOFont, %AddendumDir%\Praxomat.ini, ScanPool, BOFont, Futura Bk Bt

	;-------------------------------------------------------------------------- PDFReader Pfad -----------------------------------------------------------------------
		SplitPath, PDFReaderFullPath, PDFReaderExe

	;------------------------------------------------------------ Arbeitsbereich des Monitors bestimmen ----------------------------------------------------------
		SysGet, Mon1, MonitorWorkArea, 1


;}

files:= Indexer(BefundOrdner)

XA_Save(files, PdfIndexPath)

MsgBox, % "ICh bin kaputt!"

ExitApp

PraxTTOff:						;{ 	PraxTT - ToolTip im Addendum-Design - Timerlabel zum Ausblenden des Gui
	AnimateWindow(hPraxTT, PraxTTDuration, BLEND)
	Gui, PraxTT: Destroy
return
;}

Esc::ExitApp

;{12. Funktionen

Indexer(Directory) {

		FileEncoding, UTF-8

		PageSum:=0, FileCount:=0, pdfCount:= 0, P:=0, allfiles:=""
		Index:= Object(), Screen:= Object()

		Loop, %Befundordner%\*.pdf
					pdfcount++

		if FileExist(PdfIndexFile) {

		}


		Screen:= screenDims()
		PrWidth:= Screen.W // 3
		Progress, P0 B R1-%pdfCount% FM26 FS10 WM500 cW202842 cB6d8fff cTFFFFFF zH25 w%PrWidth% WS500,  ScanPool_Adept Skript -  ...indiziere Pdf Dokumente`n`n`n , Addendum für AlbisOnWindows, ScanPool_Adept Progress, Futura Bk Bt
		hPRogress:= WinExist("ScanPool_Adept Progress")
		OnMEssage(0x200, "WM_Move")
		PAdd:= 100/pdfcount

		; INDEX FILES AND UPDATE GUI WITH STATUS
		Loop, %Directory%\*.pdf, 0, 0
		{

			If (A_LoopFileExt = "")								 ; Skip any file without a file extension
				continue

			if A_LoopFileAttrib contains H,R,S				 ; Skip any file that is either H (Hidden), R (Read-only), or S (System). Note: No spaces in "H,R,S".
				continue

			if A_LoopFileExt contains PDF
			{

					FileCount++

					PdfInfo:= PdfInfo(Directory . "\" . A_LoopFileName)
					result:= RegExMatch(PdfInfo, "(Pages:\s+\d)", Match)
					NrOfPages:= Trim(StrReplace(Match1, "Pages:", ""))
					PageSum+= NrOfPages
				;PdfText einlesen
					PdfText:= PdftoText(FileCount, Directory . "\" . A_LoopFileName, NrOfPages)
					Condensed:=""
					Loop, Parse, PdfText, `n`r
					{
							condensed.= A_LoopField . " "
					}

					Condensed:= RegExReplace(Condensed, "\.", ". ")
					Condensed:= RegExReplace(Condensed, ":", ": ")
					Condensed:= RegExReplace(Condensed, ";", "; ")
					Condensed:= RegExReplace(Condensed, "#", " # ")
					;Condensed:= RegExReplace(Condensed, "\.|:|;|,|#|\\|`|´|", A_Space)
					;Condensed:= RegExReplace(Condensed, "(^\s+| +(?= )|\s+$)", "")
					;Pat:= StrReplace(condensed, A_Space . A_Space . A_Space, A_Space)
					;Pat:= StrReplace(condensed, A_Space . A_Space, A_Space)

					;ToolTip, % Condensed
					Index[FileCount]:= {PdfName: A_LoopFileName, FileSize: A_LoopFileSizeKB, PageCount: NrOfPages, Text: Condensed}

					;P+= PAdd
					Progress, % FileCount, % "ScanPool_Adept Skript -  ...indiziere Pdf Dokumente`n" Directory . "\" . A_LoopFileName "`n" SubStr("    " . FileCount, -StrLen(pdfcount)+1) "/" pdfcount " Dokumente. " SubStr("     " . PageSum, -3) " Seiten. " SubStr("  " . Round(100/pdfCount*FileCount), -2) "% indiziert." 			;Round(P)
			}

		}


		OnMessage(0x200, "")
		Progress, Off

return Index
}

WM_Move(wparam, lparam, msg, hwnd) {            														;BO Gui verschieben, Pdf Preview Fenster ausblenden
	if wparam = 1 ; LButton
		PostMessage, 0xA1, 2,,, ahk_id %hwnd% ; WM_NCLBUTTONDOWN
return 0
}


;}

;{13. TrayIcon - Prozesse

ShowGui: 								;{ 		Wiederanzeigen des Fenster wenn es minimiert wurde

	LIstVars

return
;}

SkriptReload: 						;{			neu laden des Skriptes
		Reload
return
;}


	;}

;{14. Adept exit
AdeptExit:
	ExitApp
return
;}


;#Include %A_ScriptDir%\..\..\..\include\FindText.ahk
#include %A_ScriptDir%\..\..\..\include\GDIP_All.ahk
#include %A_ScriptDir%\..\..\..\include\XA.ahk
#include %A_ScriptDir%\..\..\..\include\LVA.ahk
#include %A_ScriptDir%\..\..\..\include\WMCopyData.ahk

#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Menu.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Misc.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
#Include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk
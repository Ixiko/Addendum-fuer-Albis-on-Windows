; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .
; . . . . . . . . . .   ADDENDUM Fujitsu ScanSnap Helper V0.1
; . . . . . . . . . .   letzte Änderung 01.10.2020
; . . . . . . . . . .
; . . . . . . . . . .	ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"
; . . . . . . . . . .	BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
; . . . . . . . . . .	RUNS WITH AUTOHOTKEY_H AND AUTOHOTKEY_L IN 32 OR 64 BIT UNICODE VERSION
; . . . . . . . . . .	THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .

;----------------------------------------------------------------------------------------------------------------------------------------------
; Skripteinstellungen
;----------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	SetBatchLines, -1
;}

SPData.XpdfPath := A_ScriptDir "\..\..\include\xpdf"

for n, param in A_Args {

	t.= n ": " param "`n"
	PdfToText(param, "1-2", "UTF-8", A_Temp "\SSnHelper.text")

}
MsgBox, % A_Temp




return



;------------ allgemeine Bibliotheken ---------------------------------
#Include %A_ScriptDir%\..\..\lib\ACC.ahk
#Include %A_ScriptDir%\..\..\lib\class_CtlColors.ahk
#Include %A_ScriptDir%\..\..\lib\class_JSON.ahk
#Include %A_ScriptDir%\..\..\lib\class_socket.ahk
#Include %A_ScriptDir%\..\..\lib\crypt.ahk
#Include %A_ScriptDir%\..\..\lib\Explorer_Get.ahk
#Include %A_ScriptDir%\..\..\lib\FindText.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\lib\ini.ahk
#Include %A_ScriptDir%\..\..\lib\LV_ExtListView.ahk
#Include %A_ScriptDir%\..\..\lib\TrayIcon.ahk
#Include %A_ScriptDir%\..\..\lib\Sci.ahk
#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\Sift.ahk
#Include %A_ScriptDir%\..\..\lib\RemoteBuf.ahk
;#Include %A_ScriptDir%\..\..\lib\WatchFolder.ahk

;------------ Addendumbibliotheken für Albis on Windows --------
#Include %A_ScriptDir%\..\..\include\Addendum_Albis.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Controls.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Gdip_Specials.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_PdfHelper.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Protocol.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Screen.ahk
#Include %A_ScriptDir%\..\..\include\Addendum_Window.ahk
#Include %A_ScriptDir%\..\..\include\Gui\PraxTT.ahk
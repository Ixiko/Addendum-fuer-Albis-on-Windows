
;ListLines Off
SetTitleMatchMode, 2		;Fast is default
SetTitleMatchMode, Fast		;Fast is default
DetectHiddenWindows, Off	;Off is default
CoordMode, Mouse, Client
CoordMode, Pixel, Client
CoordMode, ToolTip, Client
CoordMode, Menu, Client
SetKeyDelay, -1, -1
SetBatchLines, -1
SetWinDelay, -1
SetControlDelay, -1
SendMode, Input
AutoTrim, On
FileEncoding, UTF-8

OnExit, DasEnde

global AlbisWinID:= AlbisWinID()
global AlbisPID:= AlbisPID()

gosub, InitializeWinEventHooks







return

InitializeWinEventHooks:        	;{                                                                                          	;falls ich mal speziell nur einen Hook auf Albis setzten will sollte ich eine Funktion oder ein Label bereitstellen

	ExcludeScriptMessages = 1 ; 0 to include
	ExcludeGuiEvents = 1 ; 0 to include
	dwFlags := ( ExcludeScriptMessages = 1 ? 0x1 : 0x0 )
	HookProcAdr := RegisterCallback("WinEventProc", "F")
	hWinEventHook := SetWinEventHook( 0x00008001, 0x00008001, 0, HookProcAdr, AlbisPID, 0, dwFlags )

return
;}

WinEventProc(hHook, event, hwnd, idObject, idChild, eventThread, eventTime) {		    				;im Moment werden nur sich öffnende Albisfenster abgefangen
	Critical
	EHookHwnd:= GetHex(hwnd)
	SetTimer, EventHook_WinHandler,  -0
return 0
}

EventHook_WinHandler:         	;{                                                                                            	;Eventhookhandler für popup (child) Fenster in Albis

	;an diese Stelle soll irgendwann eine Art Stack sein für Handles die nicht abgearbeitet werden konnten, deshalb gibt es den flag runner

		runner:=1
		EHWT		:= WinGetTitle(EHookHwnd)
		;EHWText	:= WinGetText(EHookHwnd)
		If InStr(EHWT, "Heilmittelverordnung") {
					ExitApp
		}


WinHandlerExit:
		runner:=0
Return
;}

Create_AddendumPlus_png(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 2360 << !!A_IsUnicode)
B64 := "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"
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}

DasEnde:									;{

	OnExit

	if hWinEventHook
		UnhookWinEvent(hWinEventHook, HookProcAdr)

ExitApp
;}

#include %A_ScriptDir%\..\..\..\include\GDIP_All.ahk
#Include %A_ScriptDir%\..\..\..\include\AddendumFunctions.ahk
#Include %A_ScriptDir%\..\..\..\include\ini.ahk
;#Include %A_ScriptDir%\..\..\..\include\Gui\PraxTT.ahk



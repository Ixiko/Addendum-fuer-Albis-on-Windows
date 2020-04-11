;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 07.04.2020 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Bitmap_GetWidth(hBitmap) {                                                                                                        	;-- Returns the width of a bitmap
   Static Size := (4 * 5) + A_PtrSize + (A_PtrSize - 4)
   VarSetCapacity(BITMAP, Size, 0)
   DllCall("Gdi32.dll\GetObject", "Ptr", hBitmap, "Int", Size, "Ptr", &BITMAP, "Int")
   Return NumGet(BITMAP, 4, "Int")
}

Bitmap_GetHeight(hBitmap) {                                                                                                       	;-- Returns the height of a bitmap
   Static Size := (4 * 5) + A_PtrSize + (A_PtrSize - 4)
   VarSetCapacity(BITMAP, Size, 0)
   DllCall("Gdi32.dll\GetObject", "Ptr", hBitmap, "Int", Size, "Ptr", &BITMAP, "Int")
   Return NumGet(BITMAP, 8, "Int")
}

ImageFromBase64(NewHandle := False, iB64 := "") {                                                                 	;-- Bitmap creation from BASE64-Strings

	Static hBitmap := 0

	If (NewHandle)
	   hBitmap := 0
	If (hBitmap)
	   Return hBitmap

	VarSetCapacity(B64, StrLen(iB64) << !!A_IsUnicode)
	B64:= iB64
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

GdipCreateFromBase64(strB64, HICON := 0) {

	DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &strB64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
	VarSetCapacity(B64Dec, B64Len, 0) ; pbBinary size
	DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &strB64, "UInt", 0, "UInt", 0x01, "Ptr", &B64Dec, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
	pStream := DllCall("Shlwapi.dll\SHCreateMemStream", "Ptr", &B64Dec, "UInt", B64Len, "UPtr")
	DllCall("Gdiplus.dll\GdipCreateBitmapFromStreamICM", "Ptr", pStream, "PtrP", pBitmap)

	If (HICON) {
		DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
	}

	ObjRelease(pStream)

return (HICON ? hBitmap : pBitmap)
}

GetFontTextDimension(hFont, Text, ByRef Width := "", ByRef Height := "", c := 1) {		                	;-- calculate the height and width of the text in the specified font

	/*                              	DESCRIPTION

			Syntax: GetFontTextDimension( [hFont], [Text], [out Width], [out Height] )
			Returns:
			dependings: GetStockObject()
	*/

	hFont := hFont?hFont:GetStockObject(17)
	hDC := GetDC()
	hSelectObj := SelectObject(hDC, hFont)

	VarSetCapacity(SIZE, 8, 0)

	if !DllCall("Gdi32.dll\GetTextExtentPoint32W", "Ptr", hDC, "Ptr", &Text, "Int", StrLen(Text), "Ptr", &SIZE, "Int")
		return false, ReleaseDC(0, hDC), Width := Height := 0

	VarSetCapacity(TEXTMETRIC, 60, 0)

	if !DllCall("Gdi32.dll\GetTextMetricsW", "Ptr", hDC, "Ptr", &TEXTMETRIC, "Int") ;https://msdn.microsoft.com/en-us/library/dd144941(v=vs.85).aspx
		return false, ReleaseDC(0, hDC), Width := Height := 0

	SelectObject(hDC, hSelectObj)
	ReleaseDC(0, hDC)
	Width := NumGet(SIZE, 0, "Int")
	Height := NumGet(SIZE, 4, "Int")
	Width := Width + NumGet(TEXTMETRIC, 20, "Int") * 3
	Height := Floor((NumGet(TEXTMETRIC, 0, "Int")*c)+(NumGet(TEXTMETRIC,16, "Int")*(Floor(c+0.5)-1))+0.5)+8
	return true
} ;https://msdn.microsoft.com/en-us/library/dd144938(v=vs.85).aspx

GetStockObject(nr) {
	return DllCall( "GetStockObject", UInt, nr)
}

CreateFont(pFont:="") {                                                                                                              	;-- creates font in memory which can be used with any API function accepting font handles


	;a function by majkinetor
	;https://autohotkey.com/board/topic/21003-function-createfont/
	/*			DESCRIPTION

		--------------------------------------------------------------------------
		Function:  			Creates font in memory which can be used with any API function accepting font handles.
		Parameters: 		pFont	- AHK font description, "style, face"
		Returns:				Font handle

		Example:
						>			hFont := CreateFont("s12 italic, Courier New")
						>			SendMessage, 0x30, %hFont%, 1,, ahk_id %hGuiControl%  WM_SETFONT = 0x30

	*/

	;parse font
	italic      	:= InStr(pFont, "italic")    	?  1    	:  0
	underline  	:= InStr(pFont, "underline")	?  1    	:  0
	strikeout   	:= InStr(pFont, "strikeout") 	?  1    	:  0
	weight      	:= InStr(pFont, "bold")      	? 700	: 400

	;height
	RegExMatch(pFont, "(?<=[S|s])(\d{1,2})(?=[ ,])", height)
	if (height = "")
	  height := 10
	RegRead, LogPixels, HKEY_LOCAL_MACHINE, SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI, LogPixels
	height := -DllCall("MulDiv", "int", Height, "int", LogPixels, "int", 72)
	;face
	RegExMatch(pFont, "(?<=,).+", fontFace)
	if (fontFace != "")
	   fontFace := RegExReplace( fontFace, "(^\s*)|(\s*$)")      ;trim
	else fontFace := "MS Sans Serif"

	;create font
	hFont   := DllCall("CreateFont", "int",  height, "int",  0, "int",  0, "int", 0, "int", weight, "Uint", italic, "Uint", underline,"uint", strikeOut,"Uint", nCharSet,"Uint", 0,"Uint", 0,"Uint", 0, "Uint", 0, "str", fontFace)

	return hFont
}



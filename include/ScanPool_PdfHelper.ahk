;ScanPool_PdfHelper Library für Addendum für AlbisOnWindows
;written by Ixiko from 18.05.2019
;Abhängigkeiten: 	PDFLabs - xpdf commandline tools
;							free FoxitReader - Software
;							global Object SPData
ListLines, Off

;----------------------------------------------------------------------------------------------------------------------------------------------
; Namen extrahieren
;----------------------------------------------------------------------------------------------------------------------------------------------
ExtractNamesFromFileName(Pat) {					                   ;-- die Namen müssen die ersten zwei Worte im FileNamen der PDF-Datei sein

	Name:=[]
	;diese Funktion basiert auf der manuellen Erstellung des Dateinamens während des Scanvorganges, es ist sehr kompliziert dem Computer aus OCR Text nach Patientennamen zu durchsuchen

	;Ermitteln des Patientennamen aus dem PDF-Dateinamen
		;Pat:= RegExReplace(Pat, "\.\|;|,", A_Space) 			;entferne alle Punkt, Komma und Semikolon
		Pat:= StrReplace(Pat, "Dr.", "")
		Pat:= StrReplace(Pat, "Prof.", "")
		Pat:= StrReplace(Pat, ",", A_Space)				;(^,+|,+(?= )|,+$)	;
		Pat:= StrReplace(Pat, ".", A_Space)				;(^,+|,+(?= )|,+$)	;
		Pat:= StrReplace(Pat, ";", A_Space)				;\.|;|,
		Pat:= RegExReplace(Pat, "^Frau\s", "")
		Pat:= RegExReplace(Pat, "\sFrau\s", A_Space)
		Pat:= RegExReplace(Pat, "^Herr\s", "")
		Pat:= RegExReplace(Pat, "\s*Herr\s", A_Space)
	;zuviele Space Zeichen hintereinander machen Probleme
		Pat:= RegExReplace(Pat, "(^\s+| +(?= )|\s+$)", "")								;entfernt alle aufeinander folgenden Spacezeichen! https://autohotkey.com/board/topic/26575-removing-redundant-spaces-code-required-for-regexreplace/
		Pat:= StrSplit(Pat, A_Space)
		Name[1]:= Pat[1] . ", " . Pat[2]
		Name[2]:= Pat[2] . ", " . Pat[1]
		Name[3]:= Pat[1]
		Name[4]:= Pat[2]

return Name
}

VorUndNachname(Name) {												;-- teilt einen Komma-getrennten String und entfernt Leerzeichen am Anfang und Ende eines Namens
	Arr:=[]
	Arr[1]	:= StrSplitEx(Name, 1, ",")		;Trimmed den String
	Arr[2]	:= StrSplitEx(Name, 2, ",")
return Arr
}

GetPatNames(FullName) {													;-- teilt einen Komma-getrennten String und entfernt Leerzeichen am Anfang und Ende eines Namens
	Obj	:=Object()
	obj.surname		:= StrSplitEx(FullName, 1, ",")		;Trimmed den String
	obj.prename		:= StrSplitEx(FullName, 2, ",")
return Obj
}

RegExtractAllWords(str) {

}

;----------------------------------------------------------------------------------------------------------------------------------------------
; xpdf wrapper
;----------------------------------------------------------------------------------------------------------------------------------------------
PdfToText(PdfPath, maxNr, encoding:= "UTF-8") {			   	;-- using xpdf's pdftotext, function catches the stdout

	;Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	;original By: kon - 16.4.2016 modified from: Ixiko
	q := Chr(0x22)

	newTxtFile:= StrReplace(PdfPath, ".pdf" , ".txt")
	newTxtFile:= StrReplace(newTxtFile, BefundOrdner "\" , "")


return StdOutToVar( SPData["XpdfPath"] "\pdftotext.exe" " -f 1 -l " maxNr " -bom -nopgbrk -nodiag -layout -enc " q encoding q " " q PdfPath q " " q SPData["ContentPath"] "\" newTxtFile q)			;-enc " encoding "  " " q "-" q
}

PdfToPng(PdfPath, dpi, page=1, PreviewPath:="") {	        ;-- using xpdf's pdftoPng

	;Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	;original By: kon - 16.4.2016 modified from: Ixiko

	static pdfError, q:= Chr(0x22)

	If FileExist(PreviewPath . "\PdfPreview-" . SubStr("00000" . page, -6) . ".png")
		FileDelete, % PreviewPath . "\PdfPreview-" . SubStr("00000" . page, -6) . ".png"

	StdOut:= StdOutToVar( SPData["XpdfPath"] "\pdftopng64.exe" " -f " page " -l " page " -r " dpi " -aa yes -freetype yes -aaVector yes " q PdfPath q " " q PreviewPath "\PdfPreview" q )

	If Instr(StdOut, "Error")
			pdfError ++

return pdfError
}

PdfToPng2(PdfPath, dpi, page=1) {		                            ;-- using pdftoPng.exe (xpdf) without writing the extracted picture to harddisk (not working)

	;	theres an option in
	;	based on kon's xpdf functions
	;	Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20

	;	DESCRIPTION from xpdf docs
    ;   Pdftopng converts  Portable  Document  Format  (PDF)  files  to  color, grayscale,
	;	or monochrome image files in Portable Network Graphics (PNG) format.

    ;   Pdftopng reads the PDF file, PDF-file, and writes one PNG file for each page, PNG-root-nnnnnn.png,
	;	where  nnnnnn is the page number.
	;	If PNG-root is '-', the image is sent to stdout (this is probably only useful when converting a single page).

	;	EXIT CODES
	;	 	The Xpdf tools use the following exit codes:
	;	   0      No error.
	;	   1      Error opening a PDF file.
	;	   2      Error opening an output file.
	;	   3      Error related to PDF permissions.
	;	  99     Other error.

	q := Chr(0x22)
	static pdfError

	StdOut		 	:= StdOutToVar( SPData["XpdfPath"] "\pdftopng64.exe" " -f " page " -l " page " -r " dpi " -freetype no -aaVector no " q PdfPath q " " q "-" q )

	If Instr(StdOut, "Error") {
			pdfError += 1
	} else {
			size:= VarSetCapacity(StdOut)
			ImageBin:= StdOut
			hbit:= BinImgToBITMAP(size, hBitmap, ImageBin)
			ListVars
			return {"hBitmap": (hBitmap), "pBitmap": (pBitmap)}
	}


	return 0
}

BinImgToBITMAP(size, byref hBitmap, byref Imagebin) {      ;-- Convert Binary Image to Bitmap

	; Link: https://autohotkey.com/boards/viewtopic.php?t=44115
	if hBitmap
		DllCall("DeleteObject", "ptr", hBitmap)
	hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", size, "UPtr")
	pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
	DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &ImageBin, "UPtr", size)
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

}

PdfPreviewHandler(PdfPath, PdfCachePath="") {		            ;-- additional function to cache all pages of a pdf on harddisk

	;Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	;original By: kon - 16.4.2016 modified from: Ixiko

	q := Chr(0x22)
	static pdfError

	StdOut:= StdOutToVar( SPData["XpdfPath"] "\pdftopng64.exe" " -f " page " -l " page " -r " dpi " -freetype yes -aaVector yes " q PdfPath q " " q TempPath "\sppreview" q )
	If Instr(StdOut, "Error") {
			pdfError += 1
			StdOut:=""
	}
    return pdfError
}

PdfInfo(PdfPath, opt:="", lastpage:=1) {		                     	;-- using xpdf's pdfinfo to get metadata and Pdf Info's

		;original post:	https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
		;original by:  	kon - 16.4.2016
		;modified by:		swagfag and taken by Ixiko
		;Link:            	https://autohotkey.com/boards/viewtopic.php?f=9&t=59294&hilit=pdfinfo
		;Description: 	function now returns an object (thx to swagfag for his help on RegExMatch)
		;                   	a space char in keys will be replaced from "Page size" to "Pagesize"
		;Parameters:		use a key to retreave only one result: PdfInfo2(PdfPath, "Pages")

		static q 				:= Chr(0x22)
		static PdfResults := Object()
		static PdfPath_old

		If !(PdfPath_old = PdfPath)
		{
				PdfPath_old 	:= PdfPath
				StdOut			:= StdOutToVar( SPData["XpdfPath"] "\pdfInfo.exe -f 1 -l " lastpage " -rawdates " q PdfPath q )
				If Instr(StdOut, "Error")
						StdOut := PdfPath_old:= "", return

				while(Pos := RegExMatch(StdOut, "`aimO)^([^:]+):\s*(.*)$", M, A_Index == 1 ? 1 : Pos + StrLen(M[0])))
						PdfResults[StrReplace(M[1], " ")]:= M[2]
		}

		If (opt = "")
			return PdfResults

return PdfResults[(opt)]
}

;----------------------------------------------------------------------------------------------------------------------------------------------
; PDFtk Server - wrapper
;----------------------------------------------------------------------------------------------------------------------------------------------
PdfAddMetaDataInfo(PdfPath, InfoObject, ReplaceOriginal:=true, DeleteTempFiles:=true, TempPath:= "PdfPath" ) {

	/* wrapper for pdftk.exe and function update_info_utf8 - Add new metadata to a pdf file
	* ----------------------------------------------------------------------------------------------------------
	* 	this function uses a simple key:value object to get its data
	* 	data will be written to a temporarly dumpData File
	*	please set InfoObject.pdfk key to the path of the 'PDFtk Server command line tool - pdftk.exe'
	* ----------------------------------------------------------------------------------------------------------
	* 	example of a dumpData File (encoding UTF-8) - you can add as much key's you like
	*	--------------------------------------------------------------------------------------------------------
	* 	InfoBegin
	* 	InfoKey: description
	* 	InfoValue: operation manual
	*
	* ----------------------------------------------------------------------------------------------------------
	* 	example function call:
	* ----------------------------------------------------------------------------------------------------------
	*	oInfo 		:= Object()
	*	oInfo.pdftk:= "C:\Programs\Pdftk"
	*	oInfo[1] 	:={InfoKey:"description", InfoValue:"operation manual bucksaw"}
	*	oInfo[2] 	:={InfoKey:"content", InfoValue:"Is there a dentist for saw teeth?"}
	*	If PdfAddMetaDataInfo("C:\Pdf\test.pdf", oInfo, true, true, "C:\temp")
	*			MsgbBox, Yes! Adding metadata works!
	*
	*/

	If !IsObject(InfoObject)
			return 0

	static pdfError, q:= Chr(0x22)

	If InStr(TempPath, "PdfPath")
		SplitPath, PdfPath,, TempPath
	else
		TempPath:= RTrim(TempPath, "\")

	file:= FileOpen(TempPath "\info.txt", "w", "UTF-8")

	For key, val in InfoObject
		File.Writeln("InfoBegin`nInfokey: " InfoKey "`nInfoValue: " InfoValue)

	file.Close()

	cmdline	:= InfoObject.Pdftk "\pdftk.exe " q PdfPath q " update_info_utf8 " q PdfDir "\info.txt" q " " q TempPath "\PdfTemp.pdf" q
	StdOut	:= StdOutToVar(cmdline)

	If ReplaceOriginal
		FileCopy, % TempPath "\PdfTemp.pdf", % PdfPath, 1

	If DeleteTempFiles
	{
		FileDelete, % TempPath "\PdfTemp.pdf"
		FileDelete, % TempPath "\info.txt"
	}

	If InStr(StdOut, "Error")
		return 0

return 1
}


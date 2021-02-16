;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                  	  HILFSFUNKTIONEN für die
;      DATENEXTRAKTION, DATENHANDLING VON PDF DATEIEN, ZUSÄTZLICH FUNKTIONEN ZUM AUSWERTEN VON PDF-DATEINAMEN
;                             	    	 Abhängigkeiten: 	1. xpdf commandline tools		- https://www.xpdfreader.com/download.html
;                                                            		2. FoxitReader                       	- https://www.foxitsoftware.com/
;                                                            		3. Skriptaufruf                       	- im aufrufenden Skript muss ein globales Objekt
;                                                                                                               	   mit dem Namen SPData deklariert sein
;
;                                                                                  	------------------------
;                                                	FÜR DAS AIS-ADDON: "ADDENDUM FÜR ALBIS ON WINDOWS"
;                                                                                  	------------------------
;    		BY IXIKO STARTED IN SEPTEMBER 2017 - LAST CHANGE 12.10.2020 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return
;----------------------------------------------------------------------------------------------------------------------------------------------
; Namen extrahieren
;----------------------------------------------------------------------------------------------------------------------------------------------
ExtractNamesFromFileName(Pat) {					                      	;-- die Namen müssen die ersten zwei Worte im FileNamen der PDF-Datei sein

	Name:=[]
	;diese Funktion basiert auf der manuellen Erstellung des Dateinamens während des Scanvorganges, es ist sehr kompliziert dem Computer aus OCR Text nach Patientennamen zu durchsuchen

	; Ermitteln des Patientennamen aus dem PDF-Dateinamen
		;Pat:= RegExReplace(Pat, "\.|;|,", A_Space) 			;entferne alle Punkt, Komma und Semikolon
		Pat:= StrReplace(Pat, "Dr.", " ")
		Pat:= StrReplace(Pat, "Prof.", " ")
		Pat:= StrReplace(Pat, ",", " ")				;(^,+|,+(?= )|,+$)	;
		Pat:= StrReplace(Pat, ".", " ")				;(^,+|,+(?= )|,+$)	;
		Pat:= StrReplace(Pat, ";", " ")				;\.|;|,
		Pat:= RegExReplace(Pat, "^Frau\s", " ")
		Pat:= RegExReplace(Pat, "\s*Frau\s", " ")
		Pat:= RegExReplace(Pat, "^Herr\s", " ")
		Pat:= RegExReplace(Pat, "\s*Herr\s", " ")
	  ; entfernt alle aufeinander folgenden Spacezeichen! https://autohotkey.com/board/topic/26575-removing-redundant-spaces-code-required-for-regexreplace/
		Pat:= RegExReplace(RegExReplace(Pat, "\s{2,}", ""),"(^\s*|\s*$)")
		Pat:= StrSplit(Pat, A_Space)
		Name[1]:= Pat[1] ", " Pat[2]
		Name[2]:= Pat[2] ", " Pat[1]
		Name[3]:= Pat[1]
		Name[4]:= Pat[2]

return Name
}

;----------------------------------------------------------------------------------------------------------------------------------------------
; xpdf wrapper
;----------------------------------------------------------------------------------------------------------------------------------------------
PdfToText(PdfPath, pages, enc="UTF-8", SaveToFile="") {    	;-- using xpdf's pdftotext, function catches the stdout

	; Link	: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	; By	: kon - 16.4.2016 modified from: Ixiko
	; pages = 0 -> alle Seiten extrahieren

	; parameter: pages format 	- use one number to extract only one page, use comma or - to extract a range of pages

	q := Chr(0x22)

	If !RegExMatch(pages, "(?<Start>\d+)\s*-|,\s*(?<End>\d+)", Page)
		PageStart := PageEnd := pages
	else if (pages = 0) {
		PageStart	:= 1
		PageEnd	:= PdfInfo(PdfPath, "Pages")
	}

	If (StrLen(SaveToFile) = 0)
		SplitPath, PdfPath,,,, SaveToFile

	If !InStr(SaveToFile, "\")
		SaveToFile := Addendum.BefundOrdner "\temp\" SaveToFile ".txt"

return StdOutToVar(Addendum.XpdfPath "\pdftotext.exe" " -f " PageStart " -l " PageEnd " -bom -nopgbrk -nodiag -clip -layout -enc " q enc q " " q PdfPath q " " q SaveToFile q)			;-enc " encoding "  " " q "-" q -layout
}

PdfToPng(PdfPath, page=1, dpi=300, PreviewPath="") {	            ;-- using xpdf's pdftoPng

	;Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	;original By: kon - 16.4.2016 modified from: Ixiko

	static tmpPath

	If (StrLen(PreviewPath . tmpPath) = 0)
		tmpPath := Addendum.BefundOrdner "\temp"
	else if PreviewPath && tmpPath
		tmpPath := PreviewPath

	pngPath := tmpPath "\PdfPage-" SubStr("00000" page, -6) ".png"
	If FileExist(pngPath)
		FileDelete, % pngPath

	If InStr(StdOutToVar(Addendum.XpdfPath "\pdftopng.exe" " -f " Page " -l " Page " -r " dpi " -aa yes -freetype yes -aaVector yes " q PdfPath q " " q tmpPath "\PdfPage" q), "Error")
		return "Error"

return pngPath
}

PdfPreviewHandler(PdfPath, PdfCachePath="") {    		            ;-- cache all pages of a pdf on harddisk

	; Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	; original By: kon - 16.4.2016 modified from: Ixiko

	q := Chr(0x22)
	static pdfError

	StdOut:= StdOutToVar( SPData["XpdfPath"] "\pdftopng64.exe" " -f " page " -l " page " -r " dpi " -freetype yes -aaVector yes " q PdfPath q " " q TempPath "\sppreview" q )
	If Instr(StdOut, "Error") {
		pdfError += 1
		StdOut:=""
	}

return pdfError
}

PdfInfo(PdfPath, opt:="", lastpage:=1) {		                         	;-- using xpdf's pdfinfo to get metadata and Pdf Info's

		;original post:	https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
		;original by:  	kon - 16.4.2016
		;modified by:		swagfag and taken by Ixiko
		;Link:            	https://autohotkey.com/boards/viewtopic.php?f=9&t=59294&hilit=pdfinfo
		;Description: 	function now returns an object (thx to swagfag for his help on RegExMatch)
		;                   	a space char in keys will be replaced from "Page size" to "Pagesize"
		;Parameters:		use a key to retreave only one result: PdfInfo(PdfPath, "Pages")

		static q 				:= Chr(0x22)
		static PdfResults := Object()
		static PdfPath_old

		If !(PdfPath_old = PdfPath)	{

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

;----------------------------------------------------------------------------------------------------------------------------------------------
; Sonstiges
;----------------------------------------------------------------------------------------------------------------------------------------------
MetaDatenAnzeigen(PdfPath, ShowInfo:=false)	{		        	;-- zeigt die Metadaten einer PDF Datei an

	output:="", maxlength:=0
	metadata:= Object()

	pages		:= PdfInfo(PdfPath, "Pages")
	metadata	:= PdfInfo(PdfPath, "", pages)

	If ShowInfo {
		For key, val in metadata
			maxlength:= StrLen(key) > maxlength ? StrLen(key) : maxlength

		For key, val in metadata
			output.= key SubStr("`t`t`t`t`t", 1, 1+(maxlength - StrLen(key))//4) ": " val "`n"

		MsgBox, 1, Addendum für Albis on Windows, % RTrim(output, "`n")
	}

return metadata
}

BinImgToBITMAP(size, byref hBitmap, byref Imagebin) {          ;-- Convert Binary Image to Bitmap

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

/*  PDF Dokumentstruktur - Notizen

%PDF-1.3 - "Zeilenende/trenner" Hex 0D 0A

	<<
			/Type /Pages
			/Count	2
			/Kids[ 	4 0 R          -> ..4 0 obj..	- [....] steht für ein Array, dieses Dokumentobjekt hat 2 Seiten, objekt Nr 4 und 7
						7 0 R ]        -> ..7 0 obj..
	>>

%PDF-1.7 - "Zeilenende/trenner" Hex 0A

eingebetteter OCR Text: [/PDF/Text/ImageB/ImageC/ImageI]

 */
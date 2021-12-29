;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;      DATENEXTRAKTION, DATENHANDLING VON PDF DATEIEN, ZUSÄTZLICH FUNKTIONEN ZUM AUSWERTEN VON PDF-DATEINAMEN
;
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;  		Abhängigkeiten: 	xpdf commandline tools		- https://www.xpdfreader.com/download.html
;                                	qpdf commandline tool
;                                 	Addendum_Internal.ahk
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;	 	Addendum für Albis on Windows
;   	by Ixiko started in September 2017 - last change 25.11.2021 - this file runs under Lexiko's GNU Licence
;
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return

;----------------------------------------------------------------------------------------------------------------------------------------------
; native PDF functions
;----------------------------------------------------------------------------------------------------------------------------------------------
PDFGetPages(pdfFilePath , qpdfPath:="")                           	{               	;-- gibt die Anzahl der Seiten einer PDF zurück

	; last change 25.11.2021
		pages := ""

	; PDF's are written in ANSI
		If !(fobj := FileOpen(pdfFilePath, "r", "CP0"))
			return 0

	; search string(s) that contain/s/ing page data
		while !fobj.AtEof {
			line := fobj.ReadLine()
			If RegExMatch(line, "i)\/(Count|Pages)\s+(?<ages>\d+)\/", p)
				break
			else If RegExMatch(line, "Length\s+(?<seek>\d+)", file)
				fobj.seek(fileseek, 1)
		}
		fobj.Close()

	; sometimes it detects 3 million pages - so qpdf will be used in this case
		If RegExMatch(pages, "\d+")
			If (pages > 0 && pages < 10000)
				return pages

	; #### if there's no match, sometimes the PDF XREF table is encoded/compressed - qpdf will be used
		If qpdfPath {
			qpdfPages := StdoutToVar(qpdfPath "\qpdf.exe --show-npages " q pdfFilePath q)
			If RegExMatch(qpdfPages, "\d+", qpages)
				return qpages
		}

return 0
}

PDFisCorrupt(pdfFilePath)                                               	{                	;-- prüft ob die PDF Datei defekt ist

	If !(fobj := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	VarSetCapacity(EndOfFile, 5)
	fobj.seek(fobj.Length - 6, 1)
	fobj.RawRead(EndOfFile, 5)

return InStr(StrGet(&EndOfFile, 5, "CP0"), "EOF") ? false : true
}

PDFisSearchable(pdfFilePath)                                          	{               	;-- durchsuchbare PDF Datei?

	; letzte Änderung 25.11.2021 : neuer Matchstring

	If !(fobj := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fobj.AtEof {
		line := fobj.ReadLine()
		If RegExMatch(line, "i)(Font|ToUnicode|Italic)") {
			fobj.Close()
			return true
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file)     	; binären Inhalt überspringen
			fobj.seek(fileseek, 1)                                                   	; •1 (SEEK_CUR): Current position of the file pointer.
	}

	fobj.Close()

return 0
}

PDFGetVersion(pdfFilePath)                                            	{               	;-- die Version einer PDF Datei auslesen
return FileOpen(pdfFilePath, "r", "CP0").ReadLine()
}

PDFGetEOLBytes(pdfFilePath)                                         	{               	;-- end-of-line Zeichen einer PDF ermitteln

	; je nach Version einer PDF Datei befindet sich am Ende eines Datensatzes entweder
	; ein LF (0A) oder ein CRLF "0D0A"
	; LF = Line Feed, CF = Carriage Return

	If !(fobj := FileOpen(pdfFilePath, "r", "CP0"))
		return 0

	VarSetCapacity(EOL, 1)
	fobj.seek(8, 1)
	fobj.RawRead(EOL, 1)
	fobj.Close()

return Format("{:02X}", NumGet(EOL, 0, "Int")) = "0A" ? 1 : 2
}

PDFGetXREF(pdfFilePath, EOLBytes=0)                           	{               	;-- unverschlüsselte XREF Metadaten lesen

	; FUNKTION ist nicht beendet!

	If !EOLBytes
		EOLBytes := PDFGetEOLBytes(pdfFilePath)

	If !(fobj := FileOpen(pdfFilePath, "r", "CP0"))
		return 0

	fpos := endxref := fobj.Length - (2*EOLBytes) - 5          ; 5 bytes -> %%EOF
	fobj.seek(endxref, 0)

	; loop back to find start of startxref filepointer position
	VarSetCapacity(bytes, EOLBytes)
	Loop {

		fobj.seek(-1 * EOLBytes, 1)
		fobj.RawRead(bytes, EOLBytes)
		fobj.seek(-1 * EOLBytes, 1)
		rbytes := Format("{:02X}", NumGet(bytes, 0, "Char"))
		t := 	(Mod(A_Index, 5 ) = 0   	? SubStr("00" A_Index, -1) " - " SubStr("0000000" fobj.Tell(), -7) "`n" : "")
			. 	(EOLBytes = 2               	? Format("{:02X}", NumGet(bytes, 0, "Char")) " " Format("{:02X}", NumGet(bytes, 1, "Char")) : Format("{:02X}", NumGet(bytes, 0, "Char")))
			.	(Mod(A_Index - 1, 5 ) = 0	? "`n`n" : " ") . t
		If rbytes in 0A,0D
			break
		If (A_Index > 50) {
			MsgBox, % "endxref: " endxref ", EOLBytes: " EOLBytes "`n" t
			ExitApp
		}

	}

	; calculate length of startxref pointer
	LENfpxref :=  endxref - (fobj.Tell() + EOLBytes)
	VarSetCapacity(fpxref, LENfpxref)
	fobj.seek(EOLBytes, 1)
	fobj.RawRead(fpxref, LENfpxref)
	FPxref := StrGet(&fpxref, LENfpxref, "CP0") + 0

	; set file pointer position to start of xref
	fobj.seek(FPxref, 0)
	XREF := fobj.ReadLine()

	IF !InStr(XREF, "xref")
		return ""
	XREFTable := Array(), t:=""
	line := RegExReplace(fobj.ReadLine(), "[\n\r]")
	n := StrSplit(line, " ").1 + 0
	NrObjects := StrSplit(line, " ").2 + 0
	Loop % NrObjects {
		line := RegExReplace(fobj.ReadLine(), "[\n\r]")
		XREFTable[A_Index - 1] := StrSplit(line, " ").1 + 0
		t .= XREFTable[A_Index - 1] " - " line "`n"
	}

	fobj.Close()
	MsgBox, % "Table: n=" n " objects=" NrObjects "`n" t
}

PDFObjectSearch(pdfFilePath, mstring)                            	{                	;-- freie Suche in Objekteigenschaften

	; mstring = matchstring

	If !(fobj := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fobj.AtEof {
		line := fobj.ReadLine()
		If RegExMatch(line, mstring) {
			filepos := fobj.pos
			fobj.Close()
			return filepos
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file)     	; binären Inhalt überspringen
			fobj.seek(fileseek, 1)                                                   	; •1 (SEEK_CUR): Current position of the file pointer.
	}

	fobj.Close()

return 0

}

IFilter(file, searchstring:="")                                             	{               	;-- Textsuche/Textextraktion in PDF/Word Dokumenten

	static cchBufferSize              	:= 4 * 1024
	static resultstriplinebreaks   	:= false

	static CHUNK_TEXT 	:= 1
	static STGM_READ 	:= 0
	static IFILTER_INIT_CANON_PARAGRAPHS	    	:= 1
	static IFILTER_INIT_HARD_LINE_BREAKS        		:= 2
	static IFILTER_INIT_CANON_HYPHENS 	        	:= 4
	static IFILTER_INIT_CANON_SPACES	            	:= 8
	static IFILTER_INIT_APPLY_INDEX_ATTRIBUTES	:= 16
	static IFILTER_INIT_APPLY_OTHER_ATTRIBUTES	:= 32
	static IFILTER_INIT_INDEXING_ONLY           		:= 64
	static IFILTER_INIT_SEARCH_LINKS	                	:= 128
	static IFILTER_INIT_APPLY_CRAWL_ATTRIBUTES	:= 256
	static IFILTER_INIT_FILTER_OWNED_VALUE_OK	:= 512
	static IFILTER_INIT_FILTER_AGGRESSIVE_BREAK	:= 1024
	static IFILTER_INIT_DISABLE_EMBEDDED	       	:= 2048
	static IFILTER_INIT_EMIT_FORMATTING	        	:= 4096

	static S_OK   	:= 0
	static FILTER_S_LAST_TEXT 	     	:= 268041
	static FILTER_E_NO_MORE_TEXT 	:= -2147215615

	if (!A_IsUnicode)
		throw A_ThisFunc ": The IFilter APIs appear to be Unicode only. Please try again with a Unicode build of AHK."

	if (!file)
		return -99

	SplitPath, file,,, ext
	VarSetCapacity(FILTERED_DATA_SOURCES, 4*A_PtrSize, 0)
	NumPut(&ext, FILTERED_DATA_SOURCES,, "Ptr")
	VarSetCapacity(FilterClsid, 16, 0)

	; Adobe workaround
	if (job := DllCall("CreateJobObject", "Ptr", 0, "Str", "filterProc", "Ptr"))
		DllCall("AssignProcessToJobObject", "Ptr", job, "Ptr", DllCall("GetCurrentProcess", "Ptr"))

	FilterRegistration := ComObjCreate("{9E175B8D-F52A-11D8-B9A5-505054503030}", "{c7310722-ac80-11d1-8df3-00c04fb6ef4f}")
	if (DllCall(NumGet(NumGet(FilterRegistration+0)+3*A_PtrSize), "Ptr", FilterRegistration, "Ptr", 0, "Ptr", &FILTERED_DATA_SOURCES, "Ptr", 0, "Int", false, "Ptr", &FilterClsid, "Ptr", 0, "Ptr*", 0, "Ptr*", IFilter) != 0 ) ; ILoadFilter::LoadIFilter
		throw A_ThisFunc ": can't load IFilter"

	ObjRelease(FilterRegistration)

	if (DllCall("shlwapi\SHCreateStreamOnFile", "Str", file, "UInt", STGM_READ, "Ptr*", iStream) != 0 )
		throw A_ThisFunc ": can't open filepath"

	PersistStream := ComObjQuery(IFilter, "{00000109-0000-0000-C000-000000000046}")
	if (DllCall(NumGet(NumGet(PersistStream+0)+5*A_PtrSize), "Ptr", PersistStream, "Ptr", iStream) != 0 ) ; ::Load
		throw A_ThisFunc ": can't load filestream"

	ObjRelease(iStream)

	status := 0
	MY_IFILTER := IFILTER_INIT_CANON_PARAGRAPHS | IFILTER_INIT_HARD_LINE_BREAKS | IFILTER_INIT_CANON_HYPHENS | IFILTER_INIT_CANON_SPACES
	if (DllCall(NumGet(NumGet(IFilter+0)+3*A_PtrSize), "Ptr", IFilter, "UInt", MY_IFILTER, "Int64", 0, "Ptr", 0, "Int64*", status) != 0 ) ; IFilter::Init
		throw A_ThisFunc ": can't init IFilter"

	VarSetCapacity(STAT_CHUNK, A_PtrSize == 8 ? 64 : 52)
	VarSetCapacity(buf, (cchBufferSize * 2) + 2)

	while (DllCall(NumGet(NumGet(IFilter+0)+4*A_PtrSize), "Ptr", IFilter, "Ptr", &STAT_CHUNK) == 0) { ; ::GetChunk
		if (NumGet(STAT_CHUNK, 8, "UInt") & CHUNK_TEXT) {
			while (DllCall(NumGet(NumGet(IFilter+0)+5*A_PtrSize), "Ptr", IFilter, "Int64*", (siz := cchBufferSize), "Ptr", &buf) != FILTER_E_NO_MORE_TEXT) { ; ::GetText
				text := StrGet(&buf,, "UTF-16")
				if (resultstriplinebreaks)
					text := StrReplace(text, "`r`n")
				If searchstring && (strpos := RegExMatch(text, searchstring))
					break

			}
		}
	}

	ObjRelease(PersistStream)
	ObjRelease(iFilter)
	if (job)
		DllCall("CloseHandle", "Ptr", job)

	If (strpos > 0)
		return strpos

return searchstring ? "" : Text
}


;----------------------------------------------------------------------------------------------------------------------------------------------
; xpdf wrapper
;----------------------------------------------------------------------------------------------------------------------------------------------
PdfToText(PdfPath, pages, enc="UTF-8", SaveToFile="")   	{                 	;-- using xpdf's pdftotext, function catches the stdout

	; Link	: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	; By	: kon - 16.4.2016 modified from: Ixiko
	; pages = 0 -> alle Seiten extrahieren

	; parameter: pages format 	- use one number to extract only one page, use comma or - to extract a range of pages

	q := Chr(0x22)

	if (StrLen(pages) = 0) {
		PageStart	:= 1
		PageEnd	:= PDFGetPages(PdfPath, Addendum.qpdfPath)
	}
	else If !RegExMatch(pages, "(?<Start>\d+)\s*-|,\s*(?<End>\d+)", Page)
		PageStart := PageEnd := pages

	If (StrLen(SaveToFile) = 0)
		SplitPath, PdfPath,,,, SaveToFile

	If !InStr(SaveToFile, "\")
		SaveToFile := Addendum.BefundOrdner "\temp\" SaveToFile ".txt"

return StdOutToVar(Addendum.XpdfPath "\pdftotext.exe" " -f " PageStart " -l " PageEnd " -bom -nopgbrk -nodiag -clip -layout -enc " q enc q " " q PdfPath q " " q SaveToFile q)			;-enc " encoding "  " " q "-" q -layout
}

PdfToPng(PdfPath, page=1, dpi=300, PreviewPath="")   	{	              	;-- using xpdf's pdftoPng

	;Link: https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	;original By: kon - 16.4.2016 modified from: Ixiko

	static tmpPath

	If (StrLen(PreviewPath . tmpPath) = 0)
		tmpPath := Addendum.BefundOrdner "\temp"
	else if PreviewPath && tmpPath
		tmpPath := PreviewPath

	If FileExist(pngPath := tmpPath "\PdfPage-" SubStr("00000" page, -6) ".png")
		FileDelete, % pngPath

	If InStr(StdOutToVar(Addendum.XpdfPath "\pdftopng.exe" " -f " Page " -l " Page " -r " dpi " -aa yes -freetype yes -aaVector yes " q PdfPath q " " q tmpPath "\PdfPage" q), "Error")
		return "Error"

return pngPath
}

PdfInfo(PdfPath, opt:="", lastpage:=1)                               	{	               	;-- using xpdf's pdfinfo to get metadata and Pdf Info's

	; Description:    	function now returns an object (thx to swagfag for his help on RegExMatch)
	;                       	a space char in keys will be replaced from "Page size" to "Pagesize"
	; Parameters:		use a key to retreave only one result: PdfInfo(PdfPath, "Pages")
	;
	; lastchange:			27.01.2021

	; original post:  	https://autohotkey.com/boards/viewtopic.php?t=15880&start=20
	; original by:     	kon - 16.4.2016
	; modified by:		swagfag and overtaken by Ixiko
	; Link:               	https://autohotkey.com/boards/viewtopic.php?f=9&t=59294&hilit=pdfinfo

		static qq			:= Chr(0x22)
		static PdfResults
		static PdfPath_old

		If !(PdfPath_old = PdfPath)	{
			PdfResults    	:= Object()
			PdfPath_old 	:= PdfPath
			If Instr((StdOut	:= StdOutToVar(Addendum.XpdfPath "\pdfInfo.exe -f 1 -l " lastpage " -rawdates " qq PdfPath qq)), "Error") {
				PdfPath_old := ""
				return
			}
			while(Pos := RegExMatch(StdOut, "`aimO)^([^:]+):\s*(.*)$", M, A_Index == 1 ? 1 : Pos + StrLen(M[0])))
				PdfResults[StrReplace(M[1], " ")] := M[2]
		}

		If !opt
			return PdfResults

return PdfResults[(opt)]
}

;----------------------------------------------------------------------------------------------------------------------------------------------
; PDFtk Server - wrapper
;----------------------------------------------------------------------------------------------------------------------------------------------
PdfAddMetaDataInfo(PdfPath, InfoObject, RplOrigin:=true, DelTmpFiles:=true, TempPath:= "PdfPath" ) {

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
		static qq:= Chr(0x22)

		If !IsObject(InfoObject)
				return 0

		If InStr(TempPath, "PdfPath")
			SplitPath, PdfPath,, TempPath
		else
			TempPath := RTrim(TempPath, "\")

		t := "InfoBegin`n[Infokey]`tValue`n"
		For key, val in InfoObject
			t .= "[" InfoKey "] " InfoValue "`n"
		FileOpen(TempPath "\info.txt", "w", "UTF-8").Write(RTrim(t, "`n"))

		cmdline	:= InfoObject.Pdftk "\pdftk.exe " qq PdfPath qq " update_info_utf8 " qq PdfDir "\info.txt" qq " " qq TempPath "\PdfTemp.pdf" qq
		StdOut  	:= StdOutToVar(cmdline)

		If RplOrigin
			FileCopy, % TempPath "\PdfTemp.pdf", % PdfPath, 1

		If DelTmpFiles	{
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
MetaDatenAnzeigen(PdfPath, ShowInfo:=false)	{		                         	;-- zeigt die Metadaten einer PDF Datei an

	output:="", maxlength:=0
	metadata:= Object()

	pages		:= PDFGetPages(PdfPath, Addendum.qpdfPath)
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

GetPDFData(path, pdfname) 	{                                                           	;-- Dateigröße, Durchsuchbar ...

	pdfPath := path "\" pdfname
	FileGetSize	, FSize       	, % pdfPath, K
	FileGetTime	, timeStamp 	, % pdfPath, C
	FormatTime	, FTime     	, % timeStamp, dd.MM.yyyy

return {	"name"          	: pdfname
		, 	"filesize"          	: FSize
		, 	"timestamp"   	: timeStamp
		, 	"filetime"       	: FTime
		, 	"pages"          	: PDFGetPages(pdfPath, Addendum.PDF.qpdfPath)
		, 	"isSearchable"	: (PDFisSearchable(pdfPath)?1:0)}
}






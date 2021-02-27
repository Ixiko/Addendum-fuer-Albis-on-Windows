; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                        Tesseract - OCR (Multithreading Skript)
;
;      Funktion:   	PDF zu PDF-Text Umwandlung mit der Funktion tessOCRPdf
;
;      Basisskript: 	Addendum.ahk, Addendum_Gui.ahk
;
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - last change 08.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

tessOCRPdf(PDFName, params)             	{                          	;-- tesseract Texterkennung -> erstellt eine durchsuchbare PDF Datei

	; OCR einer bestehenden PDF Datei mit tesseract 4+
	; enthaltene Bilder werden extrahiert
	; preprocessing mit Imagemagick convert
	; tesseract erkennt den Text, Seiten und Bilder werden zu einer PDF Datei zusammengefügt

		startTime := A_TickCount

  ; -------------------------------------------------------------------------------------------------------------
  ; Einstellungen

	; Variablen                                	;{

		global tess
		static iniFile, tesspath, xpdfpath, docPath, rmdrive, backupPath, txtOCRPath, OCRLogPath ;, tessconfig
		static tempPath, lastSetName, PageMax
		static sep := " -------------------------------------------------------------------"
		static ocrPreProcessing	:= 1
		static negateArg         	:= 2
		static performScaleArg	:= 1
		static scaleFactor       	:= 3.5

	;}

	; Parameter parsen                    	;{
		params.SetName := StrLen(params.SetName) = 0 ? "auto" : params.SetName

		If !IsObject(tess) && !IsObject(params) {
			SciTEOutput(" !! Function error in tessOCRPdf()`n  - second parameter: [params] must be an object!)")
			return
		}

		If !IsObject(tess) || (params.SetName <> lastSetName) {

				tess := Object()
				lastSetName := params.SetName

			; Arbeitspfade korrigieren
				;tess.iniFile         	:= Trim(RTrim(params.inipath        	, "\"))
				tess.tessPath      	:= Trim(RTrim(params.tessPath      	, "\"))
				tess.leptPath      	:= Trim(RTrim(params.leptPath      	, "\"))
				tess.xpdfPath      	:= Trim(RTrim(params.xpdfpath      	, "\"))
				tess.imagickPath   	:= Trim(RTrim(params.imagickPath   	, "\"))
				tess.qpdfPath      	:= Trim(RTrim(params.qpdfPath     	, "\"))
				tess.docPath      	:= Trim(RTrim(params.documentpath, "\"))
				tess.backupPath	:= Trim(RTrim(params.BackupPath 	, "\"))
				tess.rmdrive      	:= Trim(RTrim(params.UseRamDisk	, "\"))

			; weiteres
				tess.txtOCRPath  	:= params.txtOCRPath   	<> "" ? params.txtOCRPath   	: ""
				tess.OCRLogPath	:= params.OCRLogPath	<> "" ? params.OCRLogPath  	: ""
				tess.uselang        	:= params.uselang        	<> "" ? params.uselang        	: "deu"
				tess.convertwith  	:= params.convertwith  	<> "" ? params.convertwith    	: "convert"
				tess.preprocessing	:= params.preprocessing	<> "" ? params.preprocessing	: "imagemagick"

			; tesseract config file  ;{
				If FileExist(params.tessconfig)
					tess.tessconfig 	:= FileOpen(params.tessconfig, "r", "UTF-8").Read()
				else if (StrLen(params.tessconfig) >0)
					tess.tessconfig 	:= StrReplace(params.tessconfig, "|", "`r`n")
				else
					tess.tessconfig  	:= "
												(LTrim
													tessedit_create_pdf            	1
													tessedit_create_tsv               	1
													tessedit_create_txt               	1
												)"
			;}

			; RAM-DRIVE vorhanden ?
				If IsObject(ramdrivetest := FileOpen(tess.rmdrive "\test.txt", "w")) {

				; kopiert tesseract/leptonica               	;{
					ramdrivetest.Close()
					FileDelete, % tess.rmdrive "\test.txt"

					tess.tempPath	:= tess.rmdrive

					FileCreateDir, % tess.tempPath "\tesseract"
					FileCreateDir, % tess.tempPath "\cmd"
					FileCreateDir, % tess.tempPath "\cmd\qpdf"

					If !FileExist(tess.tempPath "\tesseract\tesseract.exe") {
						SciTEOutput("  - Kopie der tesseract Dateien wird erstellt")
						RunWait, % COMSPEC " /S Robocopy " q tess.tessPath q " " q tess.tempPath "\tesseract" q ,, % "Hide"
						SciTEOutput("  - tesseract befindet sich im RAM")
					}

				;}

				; tesseract/leptonica Pfade festlegen     	;{
					If InStr(FileExist(tess.tempPath "\tesseract\tessdata_" params.useData), "D")
						tess.tessdata 	:= tess.tempPath "\tesseract\tessdata_" params.useData ; can be fast or best or what you have
					If FileExist(tess.tempPath "\tesseract\tesseract.exe")
						tess.tessexe	:= tess.tempPath	"\tesseract\tesseract.exe"
				;}

				; kopiert cmdline tools                         	;{

					Loop, Files, % tess.xpdfPath "\*.*"
						If !FileExist(tess.tempPath "\cmd\" A_LoopFileName)
							FileCopy, % A_LoopFileFullPath, % tess.tempPath "\cmd\" A_LoopFileName

					;------------------------------- pdftopng ------------------------------
					If FileExist(tess.tempPath "\cmd\pdftopng.exe")
						tess.pdftopng := tess.tempPath "\cmd\pdftopng.exe"
					;------------------------------ pdftoppm ------------------------------
					If FileExist(tess.tempPath "\cmd\pdftoppm.exe")
						tess.Pdftoppm := tess.tempPath "\cmd\pdftoppm.exe"
					;------------------------------ pdftotext ------------------------------
					If FileExist(tess.tempPath "\cmd\pdftotext.exe")
						tess.pdftotext := tess.tempPath "\cmd\pdftotext.exe"
					;----------------------------- pdftohtml ------------------------------
					If FileExist(tess.tempPath "\cmd\pdftohtml.exe")
						tess.pdftohtml := tess.tempPath "\cmd\pdftohtml.exe"
					;---------------------------- pdftoimages ----------------------------
					If FileExist(tess.tempPath "\cmd\pdfimages.exe")
						tess.pdfimages := tess.tempPath "\cmd\pdfimages.exe"

					;---------------------- ImageMagick convert ------------------------
					Loop, Files, % tess.imagickPath "\*.*"
						If !FileExist(tess.tempPath "\cmd\" A_LoopFileName)
							FileCopy, % A_LoopFileFullPath, % tess.tempPath "\cmd\" A_LoopFileName
					If FileExist(tess.tempPath "\cmd\convert.exe")
						tess.imagick := tess.tempPath "\cmd\convert.exe"

					;-------------------------------- qpdf -----------------------------------
					Loop, Files, % tess.qpdfPath "\*.*"
						If !FileExist(tess.tempPath "\cmd\qpdf\" A_LoopFileName)
							FileCopy, % A_LoopFileFullPath, % tess.tempPath "\cmd\qpdf\" A_LoopFileName
					If FileExist(tess.tempPath "\cmd\qpdf\qpdf.exe")
						tess.qpdf := tess.tempPath "\cmd\qpdf\qpdf.exe"
				;}

				}
				else {

					tess.tempPath    	:= A_Temp "\tessOCR"
					If !InStr(FileExist(tess.tempPath), "D")
						FileCreateDir, % tess.tempPath

					tess.tessexe        	:= tess.tesspath	    	"\tesseract.exe"
					tess.tessdata      	:= tess.tessPath     	"\tessdata_" params.useData
					tess.pdftopng     	:= tess.xpdfPath     	"\pdftopng.exe"
					tess.pdfTotext    	:= tess.xpdfPath    	"\pdftotext.exe"
					tess.pdftohtml     	:= tess.xpdfPath     	"\pdftohtml.exe"
					tess.imagick      	:= tess.imagickPath	"\convert.exe"
					tess.qpdf            	:= tess.qpdfPath     	"\qpdf.exe"

				}

				tess.infilePath   	:= tess.tempPath "\infile"
				tess.configPath 	:= tess.tempPath "\tessPdf"

			}

			PDFName  	:= RegExReplace(PDFName, "\.pdf$")
			pdfPath     	:= !RegExMatch(PDFName, "[A-Z]\:\\") ? (tess.docPath "\" PDFName ".pdf") : (PDFName ".pdf")
			txtName    	:= PDFName ".txt"

	;}


  ; -------------------------------------------------------------------------------------------------------------
  ; Texterkennung

	; noch vorhandene Bilddateien löschen ;{
		FileDelete, % tess.tempPath "\work*.*"
		FileDelete, % tess.tempPath "\conv*.*"
		FileDelete, % tess.tempPath "\in.pdf"
		FileDelete, % tess.tempPath "\out*.*"
	;}

	; zu verarbeitende PDF ins Arbeits- und Backupverzeichnis kopieren ;{
		If !FileExist(tess.backupPath "\" PDFName ".pdf")
			FileCopy, % pdfPath, % tess.backupPath "\" PDFName ".pdf"

		FileCopy, % pdfPath, % tess.tempPath "\in.pdf", 1

	;}

	; Seitenbilder extrahieren            	;{

		; Seitenzahl der PDF ermitteln
			PageMax := PDFGetPages(tess.tempPath "\in.pdf")
			If (params.debug > 0)
				SciTEOutput(sep "`n  # [" PageMax "] Seite" (PageMax>1?"n":"") ": \" PDFName ".pdf")

			If (StrLen(tess.preprocessing) > 0)
				SciTEOutput("  - Preprocessing: " PageMax (PageMax > 1 ? " Seitenbilder werden" : " Seitenbild wird") " für die Texterkennung vorbereitet")

		; Originalbilder entpacken
			If (tess.convertwith = "pdftopng") {
				tess.imgpattern := 5, tess.imgext := "png"
				pdfconvert := tess.pdftopng " -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q tess.tempPath "\in.pdf" q " " q tess.tempPath "\orig" q
			}	else If (tess.convertwith = "pdftoppm") {
				tess.imgpattern := 5, tess.imgext := "ppm"
				pdfconvert := tess.pdftoppm " -tif -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q tess.tempPath "\in.pdf" q " " q tess.tempPath "\orig" q
			}	else If (tess.convertwith = "pdfimages") {
				tess.imgpattern := 3, tess.imgext := "jpg"
				pdfconvert := tess.pdfimages " -j " q tess.tempPath "\in.pdf" q " " q tess.tempPath "\orig" q
			}	else If (tess.convertwith = "imagick") {
				tess.imgpattern := 0, tess.imgext := "jpg"
				pdfconvert := tess.imagick " " q tess.tempPath "\in.pdf" q " " q tess.tempPath "\orig." tess.imgext q
			}

		 ; entfernt den Textlayer aus einer durchsuchbaren PDF
			If PDFisSearchable(tess.tempPath "\in.pdf") {

			; Bilder entpacken
				If (params.debug > 0)
					SciTEOutput("  - Preprocessing: Textlayer entfernen (Seitenbilder extrahieren)")
				tess.imgpattern := 0, tess.imgext := "jpg"
				;pdfconvert	:= tess.imagick " " q tess.tempPath "\in.pdf" q " " q tess.tempPath "\conv." tess.imgext q
				pdfconvert	:= tess.imagick " -units PixelsPerInch -density 200 " q tess.tempPath "\in.pdf" q " " q tess.tempPath "\conv." tess.imgext q
				stdOut   	:= ParseStdOut(StdoutToVar(pdfconvert))
				If (params.debug > 1)
					SciTEOutput("  - Preprocessing: (pdfconvert) cmdline Ausgabe`n    `t[" pdfconvert "]" (StrLen(stdOut) > 0 ? "`n" stdOut : ""))

			; Bilder wieder zu einer PDF zusammenlegen
				If (params.debug > 0)
					SciTEOutput(	"  - Preprocessing: extrahierte Bilder zu einer Bild-PDF vereinen")
				pdfconvert	:= tess.imagick " " q tess.tempPath "\conv*." tess.imgext q " " q tess.tempPath "\out1.pdf" q
				stdOut   	:= ParseStdOut(StdoutToVar(pdfconvert))
				If (params.debug > 1)
					SciTEOutput("  - Preprocessing: (merge to pdf) cmdline Ausgabe`n    `t[" pdfconvert "]" (StrLen(stdOut) > 0 ? "`n" stdout : ""))

			}
			else {

			  ; nicht durchsuchbar, dann die PDF Datei nur Umbenennen
				FileCopy, % tess.tempPath "\in.pdf", % tess.tempPath "\out1.pdf", 1

			}

		;}

	; OCR preprocessing                 	;{

		If (tess.preprocessing = "imagemagick") {

				If (params.debug > 1)
					SciTEOutput("  - Preprocessing: Aufbereitung der Seitenbilder für Tesseract" )

			; OCR preprocessing mit Imagemagick
				workExt      	:= "png"
				tiffOpt       	:= " -colorspace gray -compress lzw "
				pngOpt	    	:= " -colorspace gray -type Grayscale -quality 90% "
				preprocess	:= tess.imagick 	" -level 4%,98%,0.5 -deskew 40% "
															.  	(workExt = "tiff" ? tiffOpt : workExt = "png" ? pngOpt : "")
															. 	" -units PixelsPerInch -density 300 "
															.	q tess.tempPath "\in.pdf" q " " q tess.tempPath "\work." workExt q
				stdOut      	:= ParseStdOut(StdoutToVar(preprocess))
				If (params.debug > 1)
					SciTEOutput("  - Preprocessing: cmdline Ausgabe`n    `t[" preprocess "]" (StrLen(stdOut) > 0 ? "`n" stdOut : ""))

			; bearbeitete Dateinamen erfassen
				If (PageMax > 1)
					Loop % PageMax
						workimgs .= tess.tempPath "\work-" (A_Index-1) "." workExt "`n"
					else
						workimgs := tess.tempPath "\work." workExt "`n"

		}

		If (params.debug > 0)
			SciTEOutput("  - Preprocessing: abgeschlossen")

	  ; infile und config file erstellen     	;{
		workimgs := ""
		Loop, Files, % tess.tempPath "\*." workExt
		{
			pages := A_Index
			workimgs .= A_LoopFileFullPath "`n"
		}

 	   If (StrLen(workimgs) > 0) {
			FileOpen(tess.infilePath, "w", "CP0").Write(RTrim(workimgs, "`n"))
		}
		else {
			SciTEOutput("  - # Abbruch: Fehler beim Preprocessing. Bilddateien nicht vorhanden!")
			return
		}

		FileOpen(tess.configPath, "w", "CP0").Write(tess.tessconfig)

		;}

	;}

	; OCR                                        	;{

		SciTEOutput("  - Texterkennung: Tesseract wird gestartet ...")
		tess_cmdline :=    tess.tessexe
		tess_cmdline .= " --tessdata-dir " q tess.tessdata q " "                       	; Datenverzeichnis
		tess_cmdline .= q tess.infilepath q                                                  	; Eingabedateien
		tess_cmdline .= " -l " tess.uselang " "                                                  	; OCR-Sprache
		tess_cmdline .= q tess.tempPath "\out2"  q                                      	; Ausgabedatei wird im Temp Verzeichnis erstellt
		tess_cmdline .= " -c textonly_pdf=1 "                                                	; Ausgabe einer nur Text enthaltenden PDF Datei
		;tess_cmdline .= " --psm 1 --oem 2"                                                	; page segmentation mode + OCR mode
		tess_cmdline .= q tess.configpath q                                                   	; config path

		stdOut := ParseStdOut(StdoutToVar(tess_cmdline))
		if InStr(stdOut, "Error") {
			SciTEOutput("  - Texterkennung: Fehlerhaft. Versuche dennoch die Dateien zu verbinden.")
			RegExReplace(stdout, "i)error", errorCount)
			If (errorCount > 1)
				ocrfailure := true
			else
				ocrfailure := false
		}

		If (params.debug > 1)
			SciTEOutput("  - Texterkennung: Tesseract Ausgabe`n   `t[" tess_cmdline  "]" (StrLen(stdOut) > 0 ? "`n" ParseStdOut(stdout) : ""))
		else If (params.debug > 0)
			SciTEOutput("  - Texterkennung: beendet")
	;}

	; OCR postprocessing                	;{
		If !ocrfailure {
		 ; Bilder wieder zu einer PDF zusammenlegen
			SciTEOutput(	"  - Postprocessing: image-only und text-only PDF-Dateien vereinen")
			pdfmerge	:= tess.qpdf " --underlay " q tess.tempPath "\out2.pdf" q " -- " q tess.tempPath "\out1.pdf" q " " q tess.tempPath "\merged.pdf"
			stdOut   	:= StdoutToVar(pdfmerge)
			stdOut   	:= ParseStdOut(stdout)
			If (params.debug > 1)
				SciTEOutput("  - Postprocessing: merge to pdf out:`n    `t[" pdfmerge "]" (StrLen(stdOut) > 0 ? "`n" stdout : ""))
		}
	;}

  ; -------------------------------------------------------------------------------------------------------------
  ; Dateien kopieren

	If !ocrfailure {
	; erstellte Dateien kopieren          	;{
		FileGetSize, PDFSize	, % tess.tempPath "\merged.pdf"
		FileGetSize, TXTSize	, % tess.tempPath "\out2.txt"

		If (PDFSize < 2) || PDFisCorrupt(tess.tempPath "\merged.pdf") {
			if (params.debug > 0)
				SciTEOutput(	"  - Dateierzeugung: die PDF Daten sind korrupt. Eine Datei konnte nicht erstellt werden.")
			FSize := "corrupt pdf format"
		}
		else {
			; Überschreiben der Originaldatei
			FileMove, % tess.tempPath "\merged.pdf", % tess.docPath "\" PDFName ".pdf", 1
			If (params.debug > 0)
				SciTEOutput(	"  - Dateierzeugung: PDF Datei wurde durch OCR Datei ersetzt`n"
								.	"  - Dateierzeugung: Textdatei mit OCR Inhalt " ((TXTSize > 0) ? "wurde erstellt" : "konnte nicht erstellt werden"))
		}

		; Textdatei kopieren wenn es einen Pfad für Textdateien gibt
		If (tess.txtOCRPath <> "") && (TXTSize > 0) {
			PdfText :=  FileOpen(tess.tempPath "\out2.txt", "r").Read()
			FileMove, % tess.tempPath "\out2.txt", % tess.txtOCRPath "\" PDFName ".txt", 1
		}

	}
	;}

	; Daten in Logdatei schreiben	    	;{
		OCRDuration := Round((A_TickCount - startTime) / 1000, 1)
		If !ocrfailure
			LogText := "file:" PDFName ";time:" OCRDuration ";pages:" PageMax ";filesize:" (PDFSize = "" ? 0 : PDFSize) ";textlength:" (TXTSize = "" ? 0 : TXTSize)
		else
			LogText := "file:" PDFName ";time:" OCRDuration ";pages:" PageMax "; ocr failure: `n" RegExReplace(RegExReplace(stdout, "[\n\r]", "|"), "[\t]", " ")

		FileAppend, % LogText "`n", % tess.OCRLogPath, % "UTF-8"
	;}

  ;
  ; -------------------------------------------------------------------------------------------------------------

return PdfText
}

;--------------------------------- file system
PDFisSearchable(pdfFilePath)                    	{                          	;-- durchsuchbare PDF Datei?

	; letzte Änderung 08.02.2021 : neuer Matchstring

	If !(fobj := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fobj.AtEof {
		line := fobj.ReadLine()
		If RegExMatch(line, "i)(Font|ToUnicode|Italic)") {
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

PDFGetPages(pdfFilePath)                         	{                          	;-- gibt die Anzahl der Seiten einer PDF zurück

	; last change 30.11.2020

	; PDFs are written in ANSI
		If !(fileobject := FileOpen(pdfFilePath, "r", "CP0"))
			return 0

	; search string(s) that contain/s/ing page data
		while !fileobject.AtEof {

			line := fileobject.ReadLine()
			If RegExMatch(line, "i)\/Count\s(?<ages>\d+)\/*", p) {
				fileObject.Close()
				break
			}
			else If RegExMatch(line, "i)\/Pages\s(?<ages>\d+)\/", p) {
				fileObject.Close()
				break
			}
			else if RegExMatch(line,"i)Type\s+\/Pages.*?\/Count\s(?<ages>\d+)\/*", p) {
				fileObject.Close()
				break
			}
			else If RegExMatch(line, "Length\s(?<seek>\d+)", file)
				fileobject.seek(fileseek, 1)

		}
		fileObject.Close()

	; and sometimes there's nothing to count
	; using qpdf, if XREF table is encoded
		If !RegExMatch(pages, "\d+") || (pages = 0) {
			stdOut := StdoutToVar(tess.qpdf " --show-npages " q pdfFilePath q)
			If RegExMatch(stdOut, "\d+", qpages)
				return qpages
			else
				return 0
		}
		else
			return pages

return 0
}

PDFGetVersion(pdfFilePath)                    	{                            	;-- PDF Version erhalten
return FileOpen(pdfFilePath, "r", "CP0").ReadLine()
}

PDFGetEOLBytes(pdfFilePath)                 	{                            	;-- PDF Dateiendbytes lesen

	If !(fobj := FileOpen(pdfFilePath, "r", "CP0"))
		return 0

	VarSetCapacity(EOL, 1)
	fobj.seek(8, 1)
	fobj.RawRead(EOL, 1)
	fobj.Close()

return Format("{:02X}", NumGet(EOL, 0, "Int")) = "0A" ? 1 : 2
}

PDFGetXREF(pdfFilePath, EOLBytes=0)   	{                            	;-- PDF XREF Daten lesen (Metadaten)

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
		t := (Mod(A_Index, 5 ) = 0 ? SubStr("00" A_Index, -1) " - " SubStr("0000000" fobj.Tell(), -7) "`n" : "")
				. (EOLBytes = 2 ? Format("{:02X}", NumGet(bytes, 0, "Char")) " " Format("{:02X}", NumGet(bytes, 1, "Char")) : Format("{:02X}", NumGet(bytes, 0, "Char")))
				. (Mod(A_Index - 1, 5 ) = 0 ? "`n`n" : " ") . t
		If rbytes in 0A,0D
			break
		If (A_Index > 50) {
			MsgBox, % "endxref: " endxref ", EOLBytes: " EOLBytes "`n" t
			ExitApp
		}

	}

	; calculate lenght of startxref pointer
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

PDFisCorrupt(pdfFilePath)                     		{                            	;-- prüft ob die PDF Datei defekt ist

	If !(fobj := FileOpen(pdfFilePath, "r", "CP0"))
		return 0

	VarSetCapacity(EndOfFile, 5)
	fobj.seek(fobj.Length - 6, 1)
	fobj.RawRead(EndOfFile, 5)
	fobj.Close()

return InStr(StrGet(&EndOfFile, 5, "CP0"), "EOF") ? false : true
}

PDFObjectSearch(pdfFilePath,mstring)       	{                            	;-- Suchen in den Objekteigenschaften

	; mstring - matchstring

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

FileIsLocked(FullFilePath)                     		{                             	;-- ist die Datei gesperrt?

	f := FileOpen(FullFilePath, "rw")
	LE		:= A_LastError
	iO 	:= IsObject(f)
	If iO
		f.Close()

return LE = 32 ? true : false
}

GetPDFViewerArg(class, hwnd)            		{

	WinGet PID, PID, % "ahk_id " hwnd
	SciTEOutput("Foxit: " Viewercmdline)

    StrQuery := "SELECT * FROM Win32_Process WHERE ProcessId=" . PID
    Enum := ComObjGet("winmgmts:").ExecQuery(StrQuery)._NewEnum
    If (Enum[Process]) {
		Viewercmdline := Process.CommandLine
		SciTEOutput("Foxit: " Viewercmdline)
	}
	Addendum.FuncCallback := ""

	; FileAppend, % A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min " - " report ".pdf`n", % Addendum.Befundordner "\PdfImportLog.txt"
}

;--------------------------------- cmdline
StdoutToVar(sCmd, sEncoding="UTF-8", sDir="", ByRef nExitCode=0) {                               	;-- cmdline Ausgabe in einen String umleiten

    DllCall( "CreatePipe",           PtrP,hStdOutRd, PtrP,hStdOutWr, Ptr,0, UInt,0 )
    DllCall( "SetHandleInformation", Ptr,hStdOutWr, UInt,1, UInt,1                 )

            VarSetCapacity( pi, (A_PtrSize == 4) ? 16 : 24,  0 )
    siSz := VarSetCapacity( si, (A_PtrSize == 4) ? 68 : 104, 0 )
    NumPut( siSz,         si,  0,                                      		"UInt" )
    NumPut( 0x100,     si,  (A_PtrSize == 4) ? 44 : 60, 	"UInt" )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 60 : 88, 	"Ptr"  )
    NumPut( hStdOutWr, si,  (A_PtrSize == 4) ? 64 : 96, 	"Ptr"  )

    If (!DllCall( "CreateProcess", "Ptr",0, "Ptr",&sCmd, "Ptr",0, "Ptr",0, "Int",True, "UInt",0x08000000, "Ptr",0, "Ptr",sDir?&sDir:0, "Ptr",&si, "Ptr",&pi ))
        Return ""
      , DllCall( "CloseHandle", "Ptr",hStdOutWr )
      , DllCall( "CloseHandle", "Ptr",hStdOutRd )

    DllCall( "CloseHandle", "Ptr",hStdOutWr ) ; The write pipe must be closed before reading the stdout.
    While ( 1 )  { ; Before reading, we check if the pipe has been written to, so we avoid freezings.

        If (!DllCall( "PeekNamedPipe", Ptr,hStdOutRd, Ptr,0, UInt,0, Ptr,0, UIntP,nTot, Ptr,0 ))
            Break

        If ( !nTot ) { ; If the pipe buffer is empty, sleep and continue checking.
            Sleep, 100
            Continue
        } ; Pipe buffer is not empty, so we can read it.

        VarSetCapacity(sTemp, nTot+1)
        DllCall( "ReadFile", Ptr,hStdOutRd, Ptr,&sTemp, UInt,nTot, PtrP,nSize, Ptr,0 )
        sOutput .= StrGet(&sTemp, nSize, sEncoding)

    }

    ; * SKAN has managed the exit code through SetLastError.
    DllCall( "GetExitCodeProcess", Ptr,NumGet(pi,0), UIntP,nExitCode )
    DllCall( "CloseHandle",        Ptr,NumGet(pi,0)                              )
    DllCall( "CloseHandle",        Ptr,NumGet(pi,A_PtrSize)                   )
    DllCall( "CloseHandle",        Ptr,hStdOutRd                                   )

Return sOutput
}

RunCMD(CmdLine, WorkingDir:="", Codepage:="CP0", Fn:="RunCMD_Output") {  ;         RunCMD v0.94

	Local         ; RunCMD v0.94 by SKAN on D34E/D37C @ autohotkey.com/boards/viewtopic.php?t=74647
	Global A_Args ; Based on StdOutToVar.ahk by Sean @ autohotkey.com/board/topic/15455-stdouttovar

	Fn := IsFunc(Fn) ? Func(Fn) : 0
	DllCall("CreatePipe", "PtrP",hPipeR:=0, "PtrP",hPipeW:=0, "Ptr",0, "Int",0)
	DllCall("SetHandleInformation", "Ptr",hPipeW, "Int",1, "Int",1)
	DllCall("SetNamedPipeHandleState","Ptr",hPipeR, "UIntP",PIPE_NOWAIT:=1, "Ptr",0, "Ptr",0)

	P8 := (A_PtrSize=8)
	VarSetCapacity(SI, P8 ? 104 : 68, 0)                          ; STARTUPINFO structure
	NumPut(P8 ? 104 : 68, SI)                                     ; size of STARTUPINFO
	NumPut(STARTF_USESTDHANDLES:=0x100, SI, P8 ? 60 : 44,"UInt")  ; dwFlags
	NumPut(hPipeW, SI, P8 ? 88 : 60)                              ; hStdOutput
	NumPut(hPipeW, SI, P8 ? 96 : 64)                              ; hStdError
	VarSetCapacity(PI, P8 ? 24 : 16)                              ; PROCESS_INFORMATION structure

	If not DllCall("CreateProcess", "Ptr",0, "Str",CmdLine, "Ptr",0, "Int",0, "Int",True,"Int",0x08000000 | DllCall("GetPriorityClass", "Ptr",-1, "UInt"), "Int",0,"Ptr",WorkingDir ? &WorkingDir : 0, "Ptr",&SI, "Ptr",&PI)
		 Return Format("{1:}", "", ErrorLevel := -1,DllCall("CloseHandle", "Ptr",hPipeW), DllCall("CloseHandle", "Ptr",hPipeR))

	DllCall("CloseHandle", "Ptr",hPipeW)
	A_Args.RunCMD := { "PID": NumGet(PI, P8? 16 : 8, "UInt") }
	File := FileOpen(hPipeR, "h", Codepage)

	LineNum := 1,  sOutput := ""
	While (A_Args.RunCMD.PID + DllCall("Sleep", "Int",0))
		and DllCall("PeekNamedPipe", "Ptr",hPipeR, "Ptr",0, "Int",0, "Ptr",0, "Ptr",0, "Ptr",0)
			While A_Args.RunCMD.PID and (Line := File.ReadLine())
			  sOutput .= Fn ? Fn.Call(Line, LineNum++) : Line

	A_Args.RunCMD.PID := 0
	hProcess := NumGet(PI, 0)
	hThread  := NumGet(PI, A_PtrSize)

	DllCall("GetExitCodeProcess", "Ptr",hProcess, "PtrP",ExitCode:=0)
	DllCall("CloseHandle", "Ptr",hProcess)
	DllCall("CloseHandle", "Ptr",hThread)
	DllCall("CloseHandle", "Ptr",hPipeR)

	ErrorLevel := ExitCode

Return sOutput
}

ParseStdOut(stdOut, indent="    `t")      	{

	For i, line in StrSplit(stdOut, "`n", "`r")
		t .= StrLen(line) > 0 ? indent . line "`n" : ""

return RTrim(t, "`n")
}

;--------------------------------- Debugging
SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {     	; modified version for Addendum für Albis on Windows

	; last change 17.08.2020

	; some variables
		static	LinesOut           	:= 0
			, 	SCI_GETLENGTH	:= 2006
			,	SCI_GOTOPOS		:= 2025

	; gets Scite COM object
		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

	; move Caret to end of output pane to prevent inserting text at random positions
		SendMessage, 2006,,, Scintilla2, % "ahk_id " SciObj.SciteHandle
		endPos := ErrorLevel
		SendMessage, 2025, % endPos,, Scintilla2, % "ahk_id " SciObj.SciteHandle

	; shows count of printed lines in case output pane was erased
		If InStr(Text, "ShowLinesOut") {
			;SciObj.Output("SciteOutput function has printed " LinesOut " lines.`n")
			return
		}

	; Clear output window
		If Clear || ((StrLen(Text) = 0) && (LinesOut = 0))
			SendMessage, SciObj.Message(0x111, 420)

	; send text to SciTE output pane
		If (StrLen(Text) != 0) {
			Text .= (LineBreak ? "`r`n": "")
			SciObj.Output(Text)
			LinesOut += StrSplit(Text, "`n", "`r").MaxIndex()
		}

		If Exit {
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

}

;--------------------------------- Kommunikation
Send_WM_COPYDATA(ByRef StringToSend, ScriptID) {        	;-- für die Interskriptkommunikation - keine Netzwerkkommunikation!

    static TimeOutTime            	:= 4000
    Prev_DetectHiddenWindows	:= A_DetectHiddenWindows
    Prev_TitleMatchMode         	:= A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2

    VarSetCapacity(CopyDataStruct, 3*A_PtrSize, 0)
    SizeInBytes := (StrLen(StringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(SizeInBytes   	, CopyDataStruct, A_PtrSize)
    NumPut(&StringToSend	, CopyDataStruct, 2*A_PtrSize)
    SendMessage, 0x4a, 0, &CopyDataStruct,, % "ahk_id " ScriptID	;,,,, % TimeOutTime ; 0x4a is WM_COPYDATA.

    DetectHiddenWindows, % Prev_DetectHiddenWindows 	; Restore original setting for the caller.
    SetTitleMatchMode	 ,  % Prev_TitleMatchMode         	; Same.

return ErrorLevel  ; Return SendMessage's reply back to our caller.
}

Receive_WM_COPYDATA(wParam, lParam) {                     	;-- empfängt Nachrichten von anderen Skripten die auf demselben Computer ausgeführt werden

    StringAddress := NumGet(lParam + 2*A_PtrSize)
	fn_MsgWorker := Func("MessageWorker").Bind(StrGet(StringAddress))
	SetTimer, % fn_MsgWorker, -10

return
}

MessageWorker(msg) {                                                       	;{  verarbeitet die eingegangen Nachrichten

		;SciTEOutput("  - [OCR Thread] Message received: " msg)
		recv := {	"txtmsg"		: (StrSplit(msg, "|").1)
					, 	"opt"     	: (StrSplit(msg, "|").2)
					, 	"fromID"	: (StrSplit(msg, "|").3)}

	; main script receive's data from this thread and calls to continue the OCR process
		If InStr(recv.txtmsg, "continue")            	; Thread waits until ThreadControl is false
			ThreadControl := false
		else If InStr(recv.txtmsg, "exit")                ; to stop on emergency
			ExitApp

return
}





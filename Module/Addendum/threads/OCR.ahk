BatchOCR(noOCR) {

	For idx, pdfname in noOCR {
		pdfPath := Addendum.Befundordner "\" pdfname
		If FileExist(pdfPath) && !FileIsLocked(pdfPath) && isSearchablePDF(pdfPath)
			tessCreatePdf(pdfName)
	}

}

tessCreatePdf(admfile, options="")      	{                             	; tesseract Texterkennung -> erstellt eine durchsuchbare PDF Datei

	; OCR einer bestehenden PDF Datei mit tesseract 4+
	; enthaltene Seiten werden mit PDFtoPng extrahiert
	; tesseract erkennt den Text, Seiten und Bilder werden wieder zur PDF Datei zusammengefügt

	; RAM Disk erkennen, Dateien kopieren, Pfade anpassen ;{

			rmdrive       	:= "T:"
			tesspath      	:= Addendum.AddendumDir "\include\OCR"
			xpdfpath      	:= Addendum.AddendumDir "\include\xpdf"
			backupPath	:= Addendum.Befundordner "\backup"
			pdfPath     	:= Addendum.Befundordner "\" admFile
			pdfName  	:= StrReplace(admFile, ".pdf")
			txtPath      	:= StrReplace(pdfPath, ".pdf", ".txt")
			tessconfig  	:= "
									( 	LTrim
										tessedit_create_txt 1
										tessedit_create_pdf 1
									)"

			tessexe      	:= Addendum.AddendumDir "\include\OCR\tesseract\tesseract.exe"
			PdftoPngexe	:= Addendum.AddendumDir "\include\xpdf\pdftopng.exe"
			tessdatabest	:= Addendum.AddendumDir "\include\OCR\tesseract\tessdata_best"
			tempPath   	:= Addendum.Befundordner "\temp"
			infilePath   	:= Addendum.Befundordner "\temp\infile"
			configPath 	:= Addendum.Befundordner "\temp\tessPdf"

	  ; RAM Disk vorhanden, tesseract, pdftopng wird dorthin kopiert, temporäre Dateien werden dort erstellt
			ramdrivetest := FileOpen(rmdrive "\test.txt", "w")
			If IsObject(ramdrivetest) {

				ramdrivetest.Close()
				SciTEOutput("RAM Disk (" rmdrive ") ist vorhanden.")
				If !InStr(FileExist(rmdrive "\tesseract"), "D") || !InStr(FileExist(rmdrive "\leptonica_util"), "D") {

					SciTEOutput(A_Min ":" A_Sec ":" A_MSec " - Kopie des tesseract-Verzeichnisses wird erstellt")
					FileCopyDir, % tesspath, % rmdrive "\"
					SciTEOutput("tesseract Dateikopien anlegen: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
					FileCopy	 ,	% xpdfpath "\pdftopng.exe", % ramdrive "\pdftopng.exe"
					SciTEOutput("pdftopng.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
					SciTEOutput(A_Min ":" A_Sec ":" A_MSec " - tesseract befindet sich im RAM")

				}

				tessexe      	:= rmdrive "\tesseract\tesseract.exe"
				PdftoPngexe	:= rmdrive "\pdftopng.exe"
				tessdatabest	:= rmdrive "\tesseract\tessdata_best"
				tempPath   	:= rmdrive
				infilePath   	:= rmdrive "\infile"
				configPath 	:= rmdrive "\tessPdf"

			}


	;}

	; Backup der der Datei anlegen   	;{
		If !FileExist(backupPath "\" admFile)
			FileCopy, % pdfPath, %  backupPath "\" admFile
	;}

	; Seiten als Bilder extrahieren     	;{

			PdfText := "", pngPaths := ""

		; noch vorhandene Seitenbilder löschen
			;Loop % (PageMax := PdfInfo(pdfPath, "Pages")) {
			Loop % (PageMax := GetPDFPages(pdfPath)) {
				pngPaths .= tempPath "\PdfPage-" SubStr("00000" A_Index, -6) ".png`n"
				If FileExist(pngPath)
					FileDelete, % pngPath
			}

			SciTEOutput( "Das Dokument enthält " PageMax (PageMax = 1 ? " Seite." :" Seiten."))
			SciTEOutput(PdftoPngexe " -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q PdfPath q " " q tempPath "\PdfPage" q)

		; Seitenbilder extrahieren
			If InStr(output := StdOutToVar(PdftoPngexe " -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q PdfPath q " " q tempPath "\PdfPage" q), "Error") {
				SciTEOutput("OCR-Vorgang wurde abgebrochen: PdfToPng konnte keine Seitenbilder extrahieren.`n=> " output)
				return "Error"
			}

	;}

	; tess-infile und config file erstellen ;{
		If (StrLen(pngPaths) > 0) {
			txtfile := FileOpen(infilePath, "w", "CP0")
			txtFile.Write(RTrim(pngPaths, "`n"))
			txtFile.Close()
		}

		If !FileExist(configPath) {
			txtfile := FileOpen(configPath, "w", "CP0")
			txtFile.Write(tessconfig)
			txtFile.Close()
		}
		;}

	; tesseract commandline              	;{
		tess_cmdline := q tessexe q " "
		tess_cmdline .= q infilepath q " "                                                                              	; Eingabedateien
		tess_cmdline .= q StrReplace(pdfPath, ".pdf") q " "                                                     	; Ausgabedatei
		tess_cmdline .= "--tessdata-dir " q tessdatabest q " "                                                	; Datenverzeichnis
		tess_cmdline .= "-l deu "                                                                                          	; OCR-Sprache
		tess_cmdline .= q configpath q                                                                                	; config path
		;cmdline .= " --psm 3 --oem 2 -l deu "
		SciTEOutput(tess_cmdline)
	;}

	; OCR Vorgang starten                	;{
		output := StdoutToVar(tess_cmdline)
		SciTEOutput(output)
		If FileExist(txtPath) {
			SciTEOutput("OCR Textdatei wurde erstellt.")
			PdfText := FileOpen(txtPath, "r").Read()
		}
	;}


return PdfText
}

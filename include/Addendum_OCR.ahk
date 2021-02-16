; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                                                     OCR
;
;      Funktion:   	PDF zu PDF-Text Umwandlung mit der Funktion tessOCRPdf
;
;      Basisskript: 	Addendum.ahk, Addendum_Gui.ahk
;
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - last change 22.11.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

tessOCRPdf(PDFfilename, params)        	{                          	;-- tesseract Texterkennung -> erstellt eine durchsuchbare PDF Datei

	; OCR einer bestehenden PDF Datei mit tesseract 4+
	; enthaltene Seiten werden mit PDFtoPng extrahiert
	; OCR preprocessing mittels leptonica_util.exe
	; tesseract erkennt den Text, Seiten und Bilder werden zu einer PDF Datei zusammengefügt

		startTime := A_TickCount

	;{ Variablen, RAMDisk, cmdline Programme in die RAMDisk kopieren

		static tessInit := true
		static tess
		static iniFile, tesspath, xpdfpath, docPath, rmdrive, backupPath, txtOCRPath, OCRLogPath, tessconfig
		static tempPath, lastSetName

		static ocrPreProcessing	:= 1
		static negateArg         	:= 2
		static performScaleArg	:= 1
		static scaleFactor       	:= 3.5

		If tessInit {
			tessconfig 	:= "
								( 	LTrim
									tessedit_create_txt  	1
									tessedit_create_hocr	1
									tessedit_create_tsv 	1
									tessedit_create_pdf	1
								)"
		}

		If !IsObject(params) {
			SciTEOutput("tessCreatePdf(PDFfilename, [!params: must be an object!])")
			return
		}

		If (StrLen(params.SetName) = 0) {
			MsgBox, You have to give your tesseract settings a name!
			return
		}

		If (params.SetName != lastSetName) {

			lastSetName := params.SetName
			If tessInit
				tessInit := false

			tess := Object()
			tess.iniFile         	:= Trim(RTrim(params.inipath        	, "\"))
			tess.tessPath      	:= Trim(RTrim(params.tessPath      	, "\"))
			tess.xpdfpath      	:= Trim(RTrim(params.xpdfpath      	, "\"))
			tess.docPath      	:= Trim(RTrim(params.documentpath, "\"))
			tess.backupPath	:= Trim(RTrim(params.BackupPath 	, "\"))
			tess.rmdrive      	:= Trim(RTrim(params.UseRamDisk	, "\"))
			tess.txtOCRPath  	:= params.txtOCRPath
			tess.OCRLogPath	:= params.OCRLogPath
			tess.uselang        	:= params.uselang
			tess.useleptonica  	:= params.useleptonia
			tess.convertwith  	:= params.convertwith

			If FileExist(params.tessconfig)
				this.tessconfig 	:= FileOpen(tessconfig, "r").Read()
			else if StrLen(params.tessconfig)
				this.tessconfig 	:= params.tess_config
			else
				tess.tessconfig  	:= tessconfig

			If IsObject(ramdrivetest := FileOpen(tess.rmdrive "\test.txt", "w")) {

				ramdrivetest.Close()
				FileDelete, % tess.rmdrive "\test.txt"
				tess.tempPath 		:= tess.rmdrive

			; kopiert tesseract/leptonica in die RAMDisk
				If !InStr(FileExist(tess.tempPath "\tesseract"), "D") || !InStr(FileExist(tess.tempPath "\leptonica_util"), "D") {
					SciTEOutput(" -" A_Min "m " A_Sec "s " A_MSec "ms - Kopie der tesseract/leptonica Dateien wird erstellt")
					RunWait, % COMSPEC " /c Robocopy " q tess.tessPath q " " q tess.tempPath q ,, % "Hide"
					SciTEOutput(" -" A_Min "m " A_Sec "s " A_MSec "ms - tesseract/leptonica Dateien befindet sich im RAM")
				}

			; tesseract/leptonica Pfade festlegen     	;{
				If FileExist(tess.tempPath "\tesseract\tesseract.exe")
					tess.tessexe := tess.tempPath	"\tesseract\tesseract.exe"
				else
					tess.tessexe := tess.tessPath	"\tesseract\tesseract.exe"

				If FileExist(tess.tempPath "\leptonica_util\leptonica_util.exe")
					tess.leptonica := tess.tempPath	"\leptonica_util\leptonica_util.exe"
				else
					tess.leptonica := tess.tessPath	"\leptonica_util\leptonica_util.exe"
			;}

			; kopiert xpdf cmdline tools                 	;{
				If !FileExist(tess.tempPath "\pdftopng.exe")	{
					FileCopy, % tess.xpdfpath "\pdftopng.exe", % tess.tempPath "\pdftopng.exe"
					SciTEOutput("pdftopng.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
				}
				If !FileExist(tess.tempPath "\pdftoppm.exe")	{
					FileCopy, % tess.xpdfpath "\pdftoppm.exe", % tess.tempPath "\pdftoppm.exe"
					SciTEOutput("pdftoppm.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
				}
				If !FileExist(tess.tempPath "\pdftotext.exe")	{
					FileCopy, % tess.xpdfpath "\pdftotext.exe", % tess.tempPath "\pdftotext.exe"
					SciTEOutput("pdftotext.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
				}
				If !FileExist(tess.tempPath "\pdftohtml.exe")	{
					FileCopy, % tess.xpdfpath "\pdftohtml.exe", % tess.tempPath "\pdftohtml.exe"
					SciTEOutput("pdftohtml.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
				}
				If !FileExist(tess.tempPath "\pdfimages.exe")	{
					FileCopy, % tess.xpdfpath "\pdfimages.exe", % tess.tempPath "\pdfimages.exe"
					SciTEOutput("pdfimages.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
				}
				If !FileExist(tess.tempPath "\convert.exe")	{
					FileCopy, % tess.xpdfpath "\convert.exe", % tess.tempPath "\convert.exe"
					SciTEOutput("convert.exe kopieren: " (ErrorLevel ? "fehlgeschlagen" : "erfolgreich"))
				}
			;}

			; xpdf Pfade festlegen                         	;{
				If FileExist(tess.tempPath "\pdftopng.exe")
					tess.pdftopng := tess.tempPath "\pdftopng.exe",	tess.imgpattern := 5, tess.imgext := "png"
				else
					tess.pdftopng := tess.xpdfPath "\pdftopng.exe",	tess.imgpattern := 5, tess.imgext := "png"

				If FileExist(tess.tempPath "\pdftoppm.exe")
					tess.Pdftoppm := tess.tempPath "\pdftoppm.exe", tess.imgpattern := 5, tess.imgext := "ppm"
				else
					tess.pdftoppm := tess.xpdfPath "\pdftoppm.exe", tess.imgpattern := 5, tess.imgext := "ppm"

				If FileExist(tess.tempPath "\pdftotext.exe")
					tess.pdftotext := tess.tempPath "\pdftotext.exe"
				else
					tess.pdftotext := tess.xpdfPath "\pdftotext.exe"

				If FileExist(tess.tempPath "\pdftohtml.exe")
					tess.pdftohtml := tess.tempPath "\pdftohtml.exe"
				else
					tess.pdfTohtml := tess.xpdfPath "\pdftohtml.exe"

				If FileExist(tess.tempPath "\pdfimages.exe")
					tess.pdfimages := tess.tempPath "\pdfimages.exe", tess.imgpattern := 3, tess.imgext := "jpg"
				else
					tess.pdfimages := tess.xpdfPath "\pdfimages.exe", tess.imgpattern := 3, tess.imgext := "jpg"

				If FileExist(tess.tempPath "\convert.exe")
					tess.convert := tess.tempPath "\convert.exe", tess.imgpattern := 0, tess.imgext := "jpg"
				else
					tess.convert := tess.xpdfPath "\convert.exe", tess.imgpattern := 0, tess.imgext := "jpg"
				;}

			; tessdatadir Pfad festlegen                 	;{
				If InStr(FileExist(tess.tempPath "\tesseract\tessdata_" params.useData), "D")
					tess.tessdata 	:= tess.tempPath 	"\tesseract\tessdata_" params.useData ; can be fast or best or what you have
				else
					tess.tessdata 	:= tess.tessPath 	"\tesseract\tessdata_" params.useData
			;}

			}
			else {

				tess.tempPath    	:= A_Temp "\tessOCR"
				If !InStr(FileExist(tess.tempPath), "D")
					FileCreateDir, % tess.tempPath
				tess.tessexe        	:= tess.tesspath "\tesseract\tesseract.exe"
				tess.leptonica    	:= tess.tessPath "\leptonica_util\leptonica_util.exe"
				tess.tessdata      	:= tess.tessPath "\tesseract\tessdata_best"
				tess.PdftoPng     	:= tess.xpdfPath "\pdftopng.exe"
				tess.PdfToText    	:= tess.xpdfPath "\pdftotext.exe"
				tess.PdfToHtml  	:= tess.xpdfPath "\pdftotext.exe"

			}

			tess.infilePath   	:= tess.tempPath "\infile"
			tess.configPath 	:= tess.tempPath "\tessPdf"

		}

		If !RegExMatch(PDFfilename, "[A-Z]\:\\") {
			docPath := tess.docPath
			pdfPath     	:= docPath "\" PDFfilename
			pdfName  	:= StrReplace(PDFfilename, ".pdf")
			txtName    	:= StrReplace(PDFfilename, ".pdf", ".txt")
		} else {
			pdfPath     	:= PDFfilename
			SplitPath, pdfPath,, docPath,, pdfName
			txtName    	:= pdfName ".txt"
		}

	;}

	; Backup der der Datei anlegen   	;{
		If !FileExist(tess.backupPath "\" PDFfilename)
			FileCopy, % pdfPath, % tess.backupPath "\" PDFfilename
	;}

	; Seiten als Bilder extrahieren     	;{

		; noch vorhandene Bilddateien löschen
			Loop, Files, % tess.tempPath "\PdfPage*.*"
				FileDelete, % A_LoopFileFullPath

		; Seitenzahl der PDF ermitteln
			PageMax := GetPDFPages(pdfPath)
			SciTEOutput("  - " PageMax (PageMax = 1 ? " Seite wird " :" Seiten werden ") "extrahiert.")

			If (tess.convertwith = "pdftopng")
				pdfconvert_cmdline := tess.pdftopng " -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q pdfPath q " " q tess.tempPath "\PdfPage" q
			else If (tess.convertwith = "pdftoppm")
				pdfconvert_cmdline := tess.pdftoppm " -tif -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q pdfPath q " " q tess.tempPath "\PdfPage" q
			else If (tess.convertwith = "pdfimages")
				pdfconvert_cmdline := tess.pdfimages " -j " q pdfPath q " " q tess.tempPath "\PdfPage" q
			else If (tess.convertwith = "convert")
				pdfconvert_cmdline := tess.convert " -density 150 -compress jpeg -quality 30 " q pdfPath q " " q tess.tempPath "\PdfPage." tess.imgext q

			; convert -level 2%,78%,0.5 test\kontrast2.png test\level1.png
			;SciTEOutput(" [" pdfconvert_cmdline "]")
			RunWait, % pdfconvert_cmdline,, Hide

		  ; extrahierte Dateinamen erfassen
			If (PageMax > 1)
				Loop % PageMax
					imgPaths .= tess.tempPath "\PdfPage-" (tess.imgpattern = 0 ? A_Index-1 : SubStr("00000" A_Index, -1 * tess.imgpattern)) "." tess.imgext "`n"
			else
					imgPaths := tess.tempPath "\PdfPage." tess.imgext "`n"

		;}

	; OCR preprocessing                 	;{
		If tess.useleptonia {
			SciTEOutput("  - Vorverarbeitung von " (PageMax = 1 ? " einem Bild " : PageMax " Bildern ") "durch leptonica_util.exe")
		; ----------------------------------------------------------------------------------------------------------------------------------------------------------
			Loop % PageMax {

				imgPath := tess.tempPath "\PdfPage-" (tess.imgpattern = 0 ? A_Index-1 : SubStr("00000" A_Index, -1 * tess.imgpattern))
				imgPaths .= imgPath ".tif" "`n"

				leptonica_cmdline :=  	tess.leptonica " " q (imgPath "." tess.imgext) q " " q (imgPath ".tif") q " "
												.	negateArg " 0.5 " performScaleArg " " scaleFactor " " ocrPreProcessing " 5 2.5 " ocrPreProcessing " 2000 2000 0 0 0.0 "

				SciTEOutput("  - leptonica bearbeitet Bild: " A_Index "/" PageMax)

				RunWait, % leptonica_cmdline,, Hide

			}
		}
	;}

	; tess-infile und config file erstellen ;{
		If (StrLen(imgPaths) > 0) {
			FileOpen(tess.infilePath, "w", "CP0").Write(RTrim(imgPaths, "`n"))
		} else {
			SciTEOutput("  # Abbruch: Keine Bilder für Texterkennung vorhanden!")
			return
		}

		If !FileExist(tess.configPath)
			FileOpen(tess.configPath, "w", "CP0").Write(tess.tessconfig)

		;}

	; tesseract commandline              	;{
		SciTEOutput("  - starte Tesseract OCR...")
		tess_cmdline := q tess.tessexe q " "
		tess_cmdline .= q tess.infilepath q " "                                                 	; Eingabedateien
		tess_cmdline .= q tess.tempPath "\" pdfName q " "                           	; Ausgabedatei wird im Temp Verzeichnis erstellt
		tess_cmdline .= "--tessdata-dir " q tess.tessdata q " "                        	; Datenverzeichnis
		tess_cmdline .= "-l deu "                                                                 	; OCR-Sprache
		tess_cmdline .= q tess.configpath q                                                   	; config path
	;}

	; OCR Vorgang starten                	;{

		;output := StdOutToVar(tess_cmdline)
		RunWait, % tess_cmdline ,, Hide

		FileGetSize, PDFSize	, % tess.tempPath "\" pdfName ".pdf"
		FileGetSize, TXTSize	, % tess.tempPath "\" txtName

		If (PDFSize < 2) || isCorruptPDF(tess.tempPath "\" PDFfilename) {

			SciTEOutput("  - Fehler bei der Erstellung der PDF Datei. Es wurde keine neue PDF Datei erstellt.`n  - [ Größe der fehlerhaften Datei: " PDFSize " , Größe erkannter Text: " TXTSize "]")
			FSize := "corrupt pdf format"

		}
		else {

			SciTEOutput("> Texterkennung beendet, PDF Datei aktualisiert. Textdatei mit OCR Inhalt " ((TXTSize > 0) ? "wurde erstellt." : "konnte nicht erstellt werden."))
		}

		; Überschreiben der Originaldatei
		FileMove, % tess.tempPath "\" PDFfilename	, % tess.docPath "\" PDFfilename, 1
		If (TXTSize > 0) {
			PdfText :=  FileOpen(tess.tempPath "\" txtName, "r").Read()
			FileMove, % tess.tempPath "\" txtName, % tess.txtOCRPath "\" txtName	, 1
		}
		else
			TXTSize := "empty"
	;}

		OCRDuration := Round((A_TickCount - startTime) / 1000, 1)
		LogText := "file:" PDFfilename ";time:" OCRDuration ";pages:" PageMax ";filesize:" PDFSize ";textlength:" TXTSize
		FileAppend, % LogText "`n", % tess.OCRLogPath, % "UTF-8"
		SciTEOutput("  - " LogText)

return PdfText
}

;--------------------------------- file system
isSearchablePDF(pdfFilePath)                 {                            	;-- durchsuchbare PDF Datei?

	If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fileobject.AtEof {
		line := fileobject.ReadLine()
		If RegExMatch(line, "i)\/PDF\s*\/Text") {
			filepos := fileObject.pos
			fileObject.Close()
			return filepos
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file) {    	; binären Inhalt überspringen
			fileobject.seek(fileseek, 1)                                           	; •1 (SEEK_CUR): Current position of the file pointer.
		}
	}

	fileObject.Close()

return 0
}

GetPDFPages(pdfFilePath)                  	{                            	;-- gibt die Anzahl der Seiten einer PDF zurück

	If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	while !fileobject.AtEof {
		line := fileobject.ReadLine()
		If RegExMatch(line, "i)\/Count\s+(\d+)", pages) {
			fileObject.Close()
			return pages1
		}
		else If RegExMatch(line, "Length\s(?<seek>\d+)", file) {    	; binären Inhalt überspringen
			fileobject.seek(fileseek, 1)                                           	; •1 (SEEK_CUR): Current position of the file pointer.
		}
	}

	fileObject.Close()

return 0
}

isCorruptPDF(pdfFilePath)                     	{                            	;-- prüft ob die PDF Datei defekt ist

	If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
		return 0

	VarSetCapacity(EndOfFile, 5)
	fileobject.seek(fileobject.Length - 6, 1)
	fileobject.RawRead(EndOfFile, 5)

return InStr(StrGet(&EndOfFile, 5, "CP0"), "EOF") ? false : true
}

FileIsLocked(FullFilePath)                     	{                             	;-- ist die Datei gesperrt?

	f := FileOpen(FullFilePath, "rw")
	LE		:= A_LastError
	iO 	:= IsObject(f)
	If iO
		f.Close()

return LE = 32 ? true : false
}

GetPDFViewerArg(class, hwnd)            	{

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

CopyFilesAndFolders(SourcePattern, DestinationFolder, Overwrite = true) {

	; Copies all files and folders matching SourcePattern into the folder named DestinationFolder and
	; returns the number of files/folders that could not be copied.

	ErrorFilePaths := Array()

    ; First copy all the files (but not the folders):
		FileCopy, % SourcePattern , % DestinationFolder, % Overwrite
		ErrorCount := ErrorLevel

    ; Now copy all the folders:
		Loop, % SourcePattern, 2  ; 2 means "retrieve folders only".
		{
			FileCopyDir, % A_LoopFileFullPath, % DestinationFolder "\" A_LoopFileName, % Overwrite
			if ErrorLevel  ; Report each problem folder by name.
				ErrorFilePaths.Push(A_LoopFileFullPath)
		}

	; return object if not empty
		If (ErrorFilePaths.MaxIndex() > 0)
			return ErrorFilePaths

return
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
      , DllCall( "CloseHandle", Ptr,hStdOutWr )
      , DllCall( "CloseHandle", Ptr,hStdOutRd )

    DllCall( "CloseHandle", Ptr,hStdOutWr ) ; The write pipe must be closed before reading the stdout.
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



; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                         	Funktions- und Klassenbibliothek für Addendum für Albis on Windows oder als eigenständige Bibliothek für
;
;                                                              	- Texterkennung durch tesseract 4+ mit Erstellung einer durchsuchbaren PDF Datei
;                                                              	- PDF Parser als Klasse für einfaches extrahieren von Informationen aus PDF Dateien der Versionen 1.3 oder 1.5
;
;                                                                                                              	Addendum für Albis on Windows
;                                                          	by Ixiko started in September 2017 - last change 01.02.2021 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ListLines, Off
return

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; tesseract - Makro's
class tessOCR                                        {

  ; different OCR settings are possible
	__New(settings) {

		; parameter use="RamDisk\Laufwerksbuchstabe" um eine RamDisk zu benutzen.
		; Tesseract und alle notwendigen ausführbaren Dateien werden dorthin kopiert.
		; ist keine RamDisk verfügbar werden die cmdline Programme von der Festplatte in ihren Pfaden ausgeführt

		/* Parameters

			settings 	- 	AHK-Object can contain the following key's:

			Alias             	(String)	-	is the name for this setting. Tessconfig will be named like this.

			------ needed! path! ------
			tess_path       	(Path)        	-	main path to installation directory containing tesseract commandline executable
			tess_usedata 	(String)       	-	which dataset you want to use:  'fast' or 'best'
			leptonica_path 	(Path)        	-	main path to installation directory containing leptonica_util commandline executable
			xpdf_path       	(Path)        	-	main path to installation directory containing xpdf command line tools
			pdf_path           (Path)        	-	directory path to pdf files
			backup_path     (Path)        	-	backup directory path to save pdf files before manipulation
			work_path     	(Path)        	-	path where file operations will be executed, if useRamDisk = true this path containing
															all executable files
			------    optional     -------
			useRamDisk  	(boolean)  	-	default is false, if set to true - script will copy all executables to work_path
												    		RamDisk size should be about 128 MB tue to size of tesseract and leptonica files
			tess_config      	(String/Path)	-	a filepath to a tesseract config file or a string containing tesseract options
															default setting is: 	create pdf and text file
																						tessedit_create_txt 1
																						tessedit_create_pdf 1
			tess_lang       	(String)      	-	language 'eng', 'deu' ...
			create_backup	(Boolean)  	-	always backup original pdf file. By default is set to true.

		 */

		; Tesseract main path
			If !FileExist(RTrim(settings.tess_path, "\") "\tesseract.exe") {
				throw, % "wrong path to tesseract.exe `n< " RTrim(settings.tess_path, "\") " >"
				return
			}
			this.tessPath      	:= RTrim(settings.tess_path, "\")
			this.tessexe         	:= this.tessPath "\tesseract.exe"

		; Tesseract data path
			this.tessdata      	:= this.tessPath "\tessdata_" settings.tess_useData
			If !InStr(FileExist(this.tessdataPath), "D") {
				throw, % "wrong path to tessdata dir`n< ..\tessdata_" settings.tess_UseData " >"
				return
			}

		; leptonica_util main path
			If !FileExist(settings.leptonica_path "\leptonica_util.exe") {
				throw, % "wrong path to leptonica_util.exe `n< " settings.leptonica_path " >"
				return
			}
			this.leptonicaPath	:= settings.leptonica_path
			this.leptonicaexe   	:= settings.leptonica_path "\leptonica_util.exe"

		; xpdf command line tools
			this.xpdfPath      	:= settings.xpdf_path
			If !InStr(FileExist(this.xpdfPath), "D") {
				throw, % "wrong path to xpdf commandline tools`n< " this.xpdfpath " >"
				return
			}
			this.PdftoPngexe	:= this.xpdfPath "\pdftopng.exe"
			If !FileExist(this.PdftoPngexe) {
				throw, % "wrong path to pdftopng.exe `n< " this.PdftoPngexe " >"
				return
			}
			this.PdfInfoexe   	:= this.xpdfPath "\pdfinfo.exe"
			If !FileExist(this.PdfInfoexe) {
				throw, % "wrong path to pdfinfo.exe `n< " this.PdfInfoexe " >"
				return
			}

		; pdf document path
			this.docPath      	:= settings.pdf_path

		; backup path
			this.createbackup	:= StrLen(settings.create_backup) = 0 ? true : (settings.create_backup = false ? false : true) ; always true if parameter is not a boolean
			this.backupPath		:= settings.backup_path
			If this.createbackup && !InStr(FileExist(this.backupPath), "D") {            ; ## what happens if this.createbackup changes from false to true in nextcall?
				throw, % "backup path not exists`n< " this.backupPath " >"
				return
			}

		; temp or work-path
			this.workPath     	:= settings.work_path
			this.configPath   	:= this.workPath "\" Alias "config"
			this.infilePath     	:= this.workPath "\infile"

		; RAMDisk
			If settings.useRamDisk {

				RegExMatch(this.workPath, "^(?<rm>[A-Z]\:)", drive)
				If IsObject(ramdrivetest := FileOpen(rmdrive "\test.txt", "w")) {

					ramdrivetest.Close()

					; copy tesseract files
					If !InStr(FileExist(rmdrive "\tesseract"), "D") || {
						FileCopyDir, % this.tesspath, % rmdrive "\"
						If !ErrorLevel {
							RegExMatch(this.tesspath, ".*\\(.*)$", parentDir)
							this.tessexe   	:= this.workPath "\" parentDir1 "\tesseract.exe"
							this.tessdata	:= this.workPath "\" parentDir1 "\tessdata_" settings.tess_useData  ;### wrong!
						}
						SciTEOutput("copy tesseract files: " (ErrorLevel ? "dismissed" : "done"))
					}

					; copy leptonica files
					If !InStr(FileExist(rmdrive "\leptonica_util"), "D") {
						FileCopyDir, % this.leptonicaPath, % rmdrive "\"
						If !ErrorLevel {
							this.leptonicaPath	:= this.workPath "\leptonica_util"
							this.leptonicaexe   	:= this.leptonicaPath "\leptonica_util.exe"
						}
						SciTEOutput("copy leptonica files: " (ErrorLevel ? "dismissed" : "done"))
					}

					;copy xpdf files
					If !InStr(FileExist(rmdrive "\xpdf"), "D") {
						FileCopyDir, % this.xpdfPath, % rmdrive "\"
						If !ErrorLevel {
							this.PdftoPngexe	:= this.workPath "\xpdf\pdftopng.exe"
							this.PdfInfoexe   	:= this.workPath "\xpdf\pdfinfo.exe"
						}
						SciTEOutput("copy xpdf files: " (ErrorLevel ? "dismissed" : "done"))
					}

				}

			}

		; tessconfig, default Einstellung nutzen, aus Datei laden oder String übernehmen
			If FileExist(tessconfig)
				this.tessconfig 	:= FileOpen(tessconfig, "r").Read()
			else
				this.tessconfig 	:= settings.tess_config

			If (StrLen(this.tessconfig) = 0) {        ; Default wenn string leer bleibt
				this.tessconfig  	:= "
											(	LTrim
												tessedit_create_txt 1
												tessedit_create_pdf 1
											)"
			}

		; more tesseract parameter
			If !FileExist()
				this.tesslang	:= settings.tess_lang


	}

  ; tesseract OCR from existing PDF file to PDF-Text file
	OCRPDF(pdfPath)  	{

		; tesseract erkennt den Text, Seiten und Bilder werden wieder zur PDF Datei zusammengefügt

			static ocrPreProcessing	:= 1
			static negateArg         	:= 2
			static performScaleArg	:= 1
			static scaleFactor       	:= 3.5

		 ; you can temporarly overwrite settings for pdf path by passing a full file path
			If RegExMatch(pdfPath, "[A-Z]\:\\") {
				SplitPath, pdfPath, pdfName, pdfDir
				txtPath  	:= pdfDir "\" StrReplace(pdfName, ".pdf") ".txt"
			} else {
				pdfName	:= pdfPath
				pdfPath 	:= this.docPath "\" pdfPath
				txtPath  	:= this.docPath "\" StrReplace(pdfName, ".pdf") ".txt"
			}

		; create's a backup of pdf file
			If this.createbackup && !FileExist(this.backupPath "\" pdfName)
				FileCopy, % pdfPath, %  this.backupPath "\" pdfName

		; remove png images in workdir
			PdfText := "", pngPaths := ""
			Loop % (PageMax := pdf.getPages(pdfPath)) {
				pngPath	:= this.workPath "\PdfPage-" SubStr("00000" A_Index, -6) ".png"
				pngPaths	.= pngPath "`n"
				If FileExist(pngPath)
					FileDelete, % pngPath
			}

			SciTEOutput( "Das Dokument enthält " PageMax (PageMax = 1 ? " Seite." :" Seiten."))
			SciTEOutput(this.PdftoPngexe " -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q PdfPath q " " q this.workPath "\PdfPage" q)

		; now extract images
			If InStr(output := StdOutToVar(this.PdftoPngexe " -f 1 -l " PageMax " -r 300 -aa yes -freetype yes -aaVector yes " q PdfPath q " " q this.workPath "\PdfPage" q), "Error") {
				SciTEOutput("OCR-Vorgang wurde abgebrochen: PdfToPng konnte keine Seitenbilder extrahieren.`n=> " output)
				return "Error"
			}

		; tess-infile und config file erstellen
			If (StrLen(pngPaths) > 0) {
				txtfile := FileOpen(this.infilePath, "w", "CP0")
				txtFile.Write(RTrim(pngPaths, "`n"))
				txtFile.Close()
			}

			If !FileExist(configPath) {
				txtfile := FileOpen(this.configPath, "w", "CP0")
				txtFile.Write(this.tessconfig)
				txtFile.Close()
			}

		; tesseract commandline
			tess_cmdline := q this.tessexe                                   	q  	" "                                 	; tesseract.exe - full filepath
			tess_cmdline .= q this.infilepath                             	q   	" "                                  	; inputfiles
			tess_cmdline .= q StrReplace(pdfPath, ".pdf")           	q   	" "                                  	; output base file name
			tess_cmdline .= "--tessdata-dir " q this.tessdata       	q   	" "                                   	; tesseract data dir
			tess_cmdline .= "-l " this.tesslang                                      	" "                                    	; tesseract language
			tess_cmdline .= q this.tessconfig                            	q                                             	; config file full path
			;cmdline .= " --psm 3 --oem 2 -l deu "
			SciTEOutput(tess_cmdline)

		; start tesseract OCR
			output := StdoutToVar(tess_cmdline)
			SciTEOutput(output)
			If FileExist(this.txtPath) {
				SciTEOutput("OCR Textdatei wurde erstellt.")
				PdfText := FileOpen(this.txtPath, "r").Read()
			}


	return PdfText
	}


}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Autohotkey native PDF parser and manipulation class
class pdf                                       		{

	__New(pdfPath) {

		; setup a new ahk object storing pdf informations
			this.pdf          	:= Object()
			this.pdf.objects	:= Object()
			this.pdf.Path   	:= pdfPath

		; open's pdf file and begin with parsing
			If !(pdffile	:= FileOpen(pdfPath, "r")) {   ;, "CP1252"
				throw, % "can't open PDF File:`n" pdfPath
				return
			}

		; store file encoding
			this.pdf.enc   	:= pdffile.encoding

		; parse pdf and store informations
			while !pdffile.AtEof {

				line := pdffile.ReadLine()

			; first detect PDF Version
				If RegExMatch(line, "^\%PDF\-(?<version>\d+\.\d+)", pdf) {

					pdfobjects["PDFVersion"] := pdfVersion
					;SciTEOutput("PdfVersion: " pdfVersion)
					buff := ""
					continue

				}
			; new pdf object - open ahk subobject
				else If RegExMatch(line, "(?<Nr>\d+)\s\d+\sobj[\n\r]*$", obj) {

					buff := line
					thisobject := objNr
					pdfobjects[thisobject]      	:= Object()
					pdfobjects[thisobject].start	:= pdffile.pos - StrLen(line)
					continue

				}
			; stores stream data, length and absolut positions of start and end of file pointer
				else If RegExMatch(line, "^\<\<") {

					pdfobjects[thisobject].data  	:= RegExReplace(line, "[\n\r]*")

					If       	RegExMatch(line, "\/Type\s*\/(?<Type>\w+)\/Subtype\/(?<Subtype>\w+)\/(?<Subdata>.+)?\>\>", obj_) {

						pdfobjects[thisobject].type	    	:= obj_Type
						pdfobjects[thisobject].subtype  	:= obj_Subtype
						;pdfobjects[thisobject].subdata  	:= obj_Subdata

					}
					else if 	RegExMatch(line, "\/Type\s*\/(?<Type>\w+)\/(?<Subdata>.+)?\>\>", obj_) {

						pdfobjects[thisobject].type   		:= obj_Type
						;pdfobjects[thisobject].subdata	:= obj_Subdata

					}
					else if 	RegExMatch(line, "\/Subtype\/(?<Subtype>\w+)\/(?<Subdata>.+)?\>\>", obj_) {

						pdfobjects[thisobject].subtype  	:= obj_Subtype
						;pdfobjects[thisobject].subdata  	:= obj_Subdata

					}

					If RegExMatch(line, "Length\s(?<seek>\d+)", file) {

						RegExMatch(line, "\/Subtype\/(?<Subtype>\w+)", obj_)

						If (obj_Subtype = "XML") {

							pdffile.RawRead(rawbytes, fileseek)
							pdfobjects[thisobject].Metadata 	:= StrGet(&rawbytes, fileseek, file.encoding)
							;SciTEOutput(pdfobjects[thisobject].Metadata)

						}
						else {

							RegExMatch(line, "Filter\/(?<Filter>\w+)", obj_)
							RegExMatch(line, "Width\s(?<width>\w+)", obj_)
							RegExMatch(line, "Height\s(?<height>\w+)", obj_)
							RegExMatch(line, "ColorSpace\/(?<colorspace>\w+)", obj_)
							RegExMatch(line, "BitsPerComponent\s(?<bits>\d+)", obj_)

							pdfobjects[thisobject].stream          	:= Object()
							pdfobjects[thisobject].stream.start    	:= pdffile.pos
							pdfobjects[thisobject].stream.end    	:= pdffile.pos + fileseek
							pdfobjects[thisobject].stream.length	:= fileseek
							If obj_filter
								pdfobjects[thisobject].stream.filter   	:= obj_filter
							If obj_width
								pdfobjects[thisobject].stream.width   	:= obj_width
							If obj_height
								pdfobjects[thisobject].stream.height   	:= obj_height
							If obj_bits
								pdfobjects[thisobject].stream.bits    	:= obj_bits

							pdffile.RawRead(rawbytes, fileseek)
							;pdfobjects[thisobject].stream.rawbytes := rawbytes

						}
					}
				}

			buff .= line

		}

		; close pdf file
			pdffile.Close()

	}

  ; returns pages in pdf file (*works with PDF Version 1.3,1.5)
	getPages(pdfFilePath) {

			If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
				return 0

			while !fileobject.AtEof {
				line := fileobject.ReadLine()
				If RegExMatch(line, "i)\/Count\s+(\d+)", pages) {
					fileObject.Close()
					return pages1
				}
				else If RegExMatch(line, "Length\s(?<seek>\d+)", file) {    	; skip binary content (makes function faster than pdfInfo.exe)
					fileobject.seek(fileseek, 1)                                           	; •1 (SEEK_CUR): Current position of the file pointer.
				}
			}

			fileObject.Close()

		return 0
		}

  ; contains text, can be indexed (*works with PDF Version 1.3,1.5)
	isSearchable(pdfFilePath) {

		If !(fileobject := FileOpen(pdfFilePath, "r", "CP1252"))
			return 0

		while !fileobject.AtEof {
			line := fileobject.ReadLine()
			If RegExMatch(line, "i)\/PDF\s*\/Text") {
				filepos := fileObject.pos
				fileObject.Close()
				return filepos
			}
			else If RegExMatch(line, "Length\s(?<seek>\d+)", file) {    	; skip binary content (this makes function faster)
				fileobject.seek(fileseek, 1)                                           	; •1 (SEEK_CUR): Current position of the file pointer.
			}
		}

		fileObject.Close()

	return 0
	}

}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; some extra functions
IniReadL(inipath, Section, inikey, defaultvalue) {
	IniRead, value, % inipath, % Section, % inikey, % DefaultVal
	If InStr(value, "ERROR")
		throw % inikey " is not defined."
return value
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; includes

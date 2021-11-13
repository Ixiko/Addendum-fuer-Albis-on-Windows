; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	Textorizer - PDFText Kategorien Editor
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	    	-	soll helfen eigene Fachwörterbücher und Eigennamenlisten von Krankenhäusern, Abteilungen, Fachkollegen zu erstellen
;									-	gedacht für eine vollautomatische Benennung gescannter PDF-Dateien
;                                	-	Weiterverwendung der Daten zur Erstellung von Wörterbüchern auch für andere Programme
;
; 		Beschreibung:       -	Skript extrahiert Text aus Textdateien und PDF Dateien und nutzt dafür Layout-Positionen um zusammenhängende Textbereiche zu erfassen
;
;       Abhängigkeiten:   	-	\lib\SciteOutPut.ahk
;
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_AutoNaming started:       	20.11.2020
;       Addendum_AutoNaming last change:	20.11.2020
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


  ; ----------------------------------------------------------------------------------------------------------------------------------------------
  ; script settings
  ; ----------------------------------------------------------------------------------------------------------------------------------------------;{
	#NoEnv
	#KeyHistory 0
	#Persistent
	SetBatchLines, -1
	;ListLines, Off

	;
	;}

  ; ----------------------------------------------------------------------------------------------------------------------------------------------
  ; Variablen
  ; ----------------------------------------------------------------------------------------------------------------------------------------------;{
		RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)

	  ; ### CHANGE THIS PATH IF YOU WANT TO RUN THIS SCRIPT WITHOUT ADDENDUM.AHK #####
		If InStr(AddendumDir, "Addendum")
			dataPath := AddendumDir "\logs'n'data\_DB\Dictionary"
		else
			dataPath := A_ScriptDir "\Dictionary"

	  ; creates a path for collected data and property files
		If !InStr(FileExist(dataPath "\"), "D") {
			FileCreateDir, % dataPath "\"
			If ErrorLevel {
				throw Exception("Can't create data path:`n" dataPath "\", -1)
				ExitApp
			}
		}

	 ; a set properties for initializing this class object
		properties := {	"source"        	: { "dir"         	: "M:\Befunde"
										    		,	"ext"            	: "pdf"
										    		,	"updlimit"    	: 200        	; limit unprocessed files, collecting to much files takes to long
										    		,	"recurse"    	: false
										    		,	"workdir"   	: A_Temp}
							,	"collection" 	: {"path"        	: dataPath
											    	,	"name"      	: "medical1"
											    	,	"description"	: "dictionary for german medical words"}
							,	"categories"	: ["normal", "Fachwort", "Eigenname(n)"]}

	;}


  ; creates a new or edits an existing dictionary collection
	pdf := new Textorizer(properties)

  ; add some files to workspace
	pdf.AddFiles()

  ; output is running if
	pdf.GainData(10)

return



class Textorizer {

  ; Initializing, collecting filenames
	__New(properties)                                 	{

			if !IsObject(properties) {
				this.emsg := "initializing of new Textorizer object-class failed`n> PARAMETER PROPERTIES MUST BE AN OBJECT <"
				throw  Exception(this.emsg, -1, "`rclass Textorizer - line number: "  A_LineNumber - 7)
			}

			this.prop                	:= Object()
			this.dict                    	:= Object()
			this.stats                   	:= Object()
			this.files                  	:= Array()

			this.size                 	:= 0
			this.words             	:= 0
			this.chars              	:= 0
			this.processed         	:= 0
			this.filecount          	:= 0

			this.q := Chr(0x22)

		; store properties
			this.prop.collection	:= {	"path"        	: RegExReplace(properties.collection.path, "\\$")
												,	"name"      	: properties.collection.name
												,	"description" 	: properties.collection.description}
			this.prop.categories	:= properties.categories

		; store filepaths
			this.prop.source:= {	"dir"      	: RegExReplace(properties.source.dir, "\\$")
										,	"ext"       	: RegExReplace(properties.source.ext, "[.\s]")
										,	"recurse" 	: properties.source.recurse
										,	"updlimit"	: properties.source.updlimit
										,	"workdir"	: properties.source.workdir}

			this.prop.filebase           	:= this.prop.source.dir "\*." this.prop.source.ext
			this.prop.dict                 	:= this.prop.collection.path "\" this.prop.collection.name "-dictionay.json"
			this.prop.FList                	:= this.prop.collection.path "\" this.prop.collection.name "_filelist.json"
			this.prop.props                	:= this.prop.collection.path "\" this.prop.collection.name "_properties.json"

		; create dictionary object if not exist, otherwise load data
			If FileExist(this.prop.dict)
				this.LoadDictionary()

		; load file list if exists
			If FileExist(this.prop.FList)
				this.LoadFileList()

		; saves file list and properties
			this.SaveProperties()

	}

	; ------------------------------------------------
	; load and save data
    	LoadFileList()                                     	{

		  ; prevent's loading of empty data
			If !FileExist(this.prop.FList) || (this.Filesize(this.prop.FList) = 0)
				return

		  ; check this.files object
			If !IsObject(this.files)
				this.files := Array()

		  ; load Filelist object
			file := FileOpen(this.prop.FList, "r", "UTF-8")
			this.stats.LenFList := file.Length
			fList := JSON.Load(file.Read())
			file.Close()

 		  ; detect file changes, refresh stats - populate files array
			For fidx, data in fList {

				fpath := this.prop.source.dir "\" data.filename
				If !FileExist(fpath)
					continue

				FileGetSize	, size     	, % fpath
				FileGetTime	, modified	, % fpath, M

				If (data.size <> size) || (data.modified <> modified) {
					txt                    	:= FileOpen(fpath, "r", "UTF-8").Read()
					data.size	    	:= size
					data.modifed	:= modified
					data.words    	:= StrSplit(this.removeCRLF(txt), " ").MaxIndex()
					data.chars     	:= StrLen(RegExReplace(txt, "[\n\r\s]", ""))
				}

				this.files.Push(data)

			}

		  ; stores file count for later purposes
			this.LoadCount := this.files.Count()

		  ; a little statistics
			this.FilesStatistic()

		}

		LoadDictionary()                                	{

		  ; check this.files object
			If !IsObject(this.dict)
				this.dict := Object()

			this.dict := JSON.Load(FileOpen(this.prop.dict, "r", "UTF-8").Read())

		}

		AddFiles()                                          	{                	;-- adds files to file list

			Loop, Files, % this.prop.filebase, % (this.base.recurse ? "R" : "")
				If !this.FileInList(A_LoopFileName, this.prop.source.dir, this.LoadCount) {

				  ; skip hidden files
					if A_LoopFileAttrib contains H,R,S
						continue

					ext	:= A_LoopFileExt
				  ; skip not searchable PDF files
					If (ext = "pdf") && !this.isSearchablePDF(A_LoopFileFullPath)
						continue

				; this prevents long load times and saves memory space,
				; the file list is later automatically freed from the processed entries
					If !(this.files.Count() - this.processed >= this.base.updlimit)
						break

				; i will add parsers for various text file formats later
					If RegExMatch(ext, "txt|hocr|tsv|csv|ps|xml")
						txt 	:= FileOpen(A_LoopFileFullPath, "r").Read()

				; populate new data
					this.files.Push({	"basedir"	    	: this.prop.source.dir
										,	"filename"     	: A_LoopFileName
										,	"modified"     	: A_LoopFileTimeModified
										, 	"created"       	: A_LoopFileTimeCreated
										, 	"processed"    	: false
										, 	"size"              	: A_LoopFileSize
										, 	"words"     	 	: ext = "pdf" ? 0 : StrSplit(this.removeCRLF(txt), " ").MaxIndex()
										, 	"chars"      	 	: ext = "pdf" ? 0 : StrLen(RegExReplace(txt, "[\n\r\s]", " "))})

				}

		  ; return if no new files are added
			If (this.LoadCount = this.files.Count()) {
				SciTEOutput(" - .AddFiles(): no files added for processing!.`n")
				return
			}

		  ; print a message to console
			SciTEOutput(" - .AddFiles(): added " (this.files.Count() - this.LoadCount) " files`n")

		  ; save file list
			this.LoadCount := this.files.Count()
			this.FilesStatistic()
			this.SaveFileList()

		}

		SaveFileList()                                     	{

			propsave := FileOpen(this.prop.FList, "w", "UTF-8")
			propsave.Write(JSON.Dump(this.files	,,	2))
			this.stats.LenFList := propsave.Length
			propsave.Close()

		}

		SaveProperties()                                 	{

			propsave := FileOpen(this.prop.props, "w", "UTF-8")
			propsave.Write(JSON.Dump(this.prop,,	2))
			this.stats.LenPropfile := propsave.Length
			propsave.Close()

		}

		SaveDictionay(backup=10)                	{               	; ### backup!!

			If !IsObject(this.dict)
				return

			 FileOpen(this.prop.dict, "w", "UTF-8").Write(JSON.Dump(this.dict,,	2))

		}

	; ------------------------------------------------
	; file functions
		FilesStatistic()                                     	{                	;-- load and content statistics

			this.stats.size := this.stats.words := this.stats.chars := this.stats.processed:= 0

			If !IsObject(this.files) || (this.files.Count() = 0)
				return

			For fidx, data in this.files
					this.stats.size         	+= data.size
				, 	this.stats.words      	+= data.words
				, 	this.stats.chars       	+= data.chars
				, 	this.stats.processed	+= data.processed

			this.stats.filecount := this.files.Count()

		}

		FileInList(filename, basedir, StopAt)         {              	;-- checks if file ist already listed, returns position in array

		  ; StopAt - prevents searching all entries in files to safe time
			If (StopAt = 0 || this.files.Count() =0)
				return 0
			else if !IsObject(this.files) {
				this.files := Array()
				return 0
			}

			For fidx, data in this.files
				If (fidx = StopAt + 1)
					return 0
				else If (data.filename = filename) && (data.basedir = basedir)
					return fidx

		return 0
		}

		Filesize(fullfilepath)                             	{                	;-- FileGetSize wrapper
			FileGetSize, filesize, % fullfilepath
		return filesize
		}

		isSearchablePDF(fullfilepath)               	{              	;-- PDF contains text?

			If !IsObject(fObj := FileOpen(fullfilepath, "r", "CP1252"))
				return 0

			while !fObj.AtEof {

				;ToolTip, % fullfilepath "`n" fObj.pos, 200, 200, 15
				line := fObj.ReadLine()
				If RegExMatch(line, "i)\/PDF\s*\/Text")	{
					filepos := fObj.pos
					fObj.Close()
					return filepos
				}
				else If RegExMatch(line, "Length\s(?<seek>\d+)", file)    	   	; skip binary content
					fObj.seek(fileseek, 1)

			}

			fObj.Close()

		return 0
		}

		GetPDFPages(fullfilepath)                    	{               	;-- return count of pages in PDF

			If !IsObject(fObj := FileOpen(fullfilepath, "r", "CP1252"))
				return 0

			pagecount := 0
			while !fObj.AtEof {

				line := fObj.ReadLine()
				If RegExMatch(line, "i)\/Count\s+(\d+)", pages)	{
					fObj.Close()
					return pages1
				}
				else if RegExMatch(line, "i)\/PDF\s*\/Text.+?\/Type\s*\/Page\W") {		; some times we must count every page
					pagecount ++
				}
				else If RegExMatch(line, "Length\s(?<seek>\d+)", file)
					fObj.seek(fileseek, 1)

			}

			fObj.Close()

		return pagecount
		}

	; ------------------------------------------------
	; text extraction
		GainData(maxfile=0)                          	{

			wkdir := this.prop.source.workdir
				  q := this.q

			SciTEOutput(" >gaining data [" A_TickCount "] from path - " this.prop.source.dir "\ -" )

			For fidx, data in this.files {

				fpath := this.prop.source.dir "\" data.filename
				If !FileExist(fpath) || !this.isSearchablePDF(fpath)
					continue

				data.pages 	:= this.GetPDFPages(fpath)
				cmdline    	:=  "T:\pdftohtml -f 1 -l " data.pages " -r 150 " q fpath q " " q wkdir "\html" q

				SciTEOutput(	"   - processing file nr:         `t" fidx "/" this.files.Count() " with " data.pages " pages`n"
								.	"   - executing cmdline:       `t" SubStr(cmdline, 1, 50) ".....")

				stdout := stdoutToVar(cmdline, "UTF-8", wkdir "\" , nExitCode)
				stdout := RegExReplace(stdout, "[\n\r]", "                      `n`t")
				while, (stdout <> RegExreplace(stdout, "\s+\n\t$"))
					stdout := RegExreplace(stdout, "\s+\n\t$")

				SciteOutput("   - cmdline execution returns:`t" stdout (nExitCode ? " [" nExitCode "]" : ""))

				text := this.ParseXML(pages)

				SciTEOutput("   - text length:                    `t" StrLen(text))
				SciTEOutput("   - save text to:                   `t" sfile := wkdir "\" RegExReplace(data.filename, "\.\w+$") ".txt")

				FileOpen(sfile, "w", "UTF-8").Write(text)
				Run, % sfile

				;FileDelete, % wkdir "\html"

				ExitApp


			}

;\s(ff[\da-f])
		}

		ParseXML(pages)  {

			pos :=1, col := Array()

			Loop, % pages
				If FileExist(xmlFile := this.prop.source.workdir "\html\page" A_Index ".html") {
					SciTEOutput( " - ParseXML() loads html file nr: " A_Index "/" pages)
					tmptxt .= FileOpen(xmlfile, "r", "UTF-8").Read()
				}


		 ;  collect all left and top positions min,max left and top
			while 	(pos := RegExMatch(tmptxt, "<div\s+class.*?<\/div>", div, pos))
					||	(pos := RegExMatch(tmptxt, "<img\s+.*?width\=.(?<width>\d+).*?(?<height>\d+)", p, pos))
			{
				If (StrLen(pagewidth) > 0)
					blockwidth := Round(pwidth/5), Loop 5, col[A_Index] := Object()
					, width:=bLeft:=pwidth, height:=bTop:=pheight, pos +=1, continue

				pos += 1 ;StrLen(div)
				divpos := 1
				RegExMatch(div, "\<div\s+class.*?left\:(?<left>\d+).*?top\:(?<top>\d+)", c, divpos)
				divpos += 1
				while, (divpos := RegExMatch(div, "<span\s+.*?(?<px>\d+)px.*?>(?<text>.*?)<\/span>", sub, divpos)) {

					divpos += 1  ;StrLen(sub)
					column := Round(cLeft/WidthBlock) + 1
					words.Push({"left":cLeft, "top":cTop, "px":subpx, "text":subtext})
					concat .= subtext " "
					bleft := cLeft, bTop := cTop
					;SciTEOutput(SubStr("000" A_Index, -2) ": l" SubStr("0000" cLeft, -3) " t" SubStr("0000" cTop, -3) " px" subpx  " - text[" subtext "]")
				}

			}

		return concat
		}

	; ------------------------------------------------
	; gui
		AddWords()                                     		{

			Gui, TXR: new, +hwndhTXR


		}

	; debug gui



	; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; internal use	  +++++++++++++++++++++++++++++++++++++++++++++++++++++++;{
	; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		removeCRLF(str)                                 	{                	; removes all carriage returns and line feeds

			str := RegExReplace(str, "[\n\r]", "###")
			str := RegExReplace(str, "###", " ")

		return RegExReplace(str, "[\n\r]", " ")
		}
	;}

	; callback
		ErrorFunc() {

			MsgBox,1, % "fatal error", % this.emsg "`nExit script"
			ExitApp

		}

}


RunCMD(CmdLine, WorkingDir:="", Codepage:="CP0", Fn:="RunCMD_Output") {  ;         RunCMD v0.94
Local         ; RunCMD v0.94 by SKAN on D34E/D37C @ autohotkey.com/boards/viewtopic.php?t=74647
Global A_Args ; Based on StdOutToVar.ahk by Sean @ autohotkey.com/board/topic/15455-stdouttovar

  Fn := IsFunc(Fn) ? Func(Fn) : 0
, DllCall("CreatePipe", "PtrP",hPipeR:=0, "PtrP",hPipeW:=0, "Ptr",0, "Int",0)
, DllCall("SetHandleInformation", "Ptr",hPipeW, "Int",1, "Int",1)
, DllCall("SetNamedPipeHandleState","Ptr",hPipeR, "UIntP",PIPE_NOWAIT:=1, "Ptr",0, "Ptr",0)

, P8 := (A_PtrSize=8)
, VarSetCapacity(SI, P8 ? 104 : 68, 0)                          ; STARTUPINFO structure
, NumPut(P8 ? 104 : 68, SI)                                     ; size of STARTUPINFO
, NumPut(STARTF_USESTDHANDLES:=0x100, SI, P8 ? 60 : 44,"UInt")  ; dwFlags
, NumPut(hPipeW, SI, P8 ? 88 : 60)                              ; hStdOutput
, NumPut(hPipeW, SI, P8 ? 96 : 64)                              ; hStdError
, VarSetCapacity(PI, P8 ? 24 : 16)                              ; PROCESS_INFORMATION structure

  If not DllCall("CreateProcess", "Ptr",0, "Str",CmdLine, "Ptr",0, "Int",0, "Int",True
                ,"Int",0x08000000 | DllCall("GetPriorityClass", "Ptr",-1, "UInt"), "Int",0
                ,"Ptr",WorkingDir ? &WorkingDir : 0, "Ptr",&SI, "Ptr",&PI)
     Return Format("{1:}", "", ErrorLevel := -1
                   ,DllCall("CloseHandle", "Ptr",hPipeW), DllCall("CloseHandle", "Ptr",hPipeR))

  DllCall("CloseHandle", "Ptr",hPipeW)
, A_Args.RunCMD := { "PID": NumGet(PI, P8? 16 : 8, "UInt") }
, File := FileOpen(hPipeR, "h", Codepage)

, LineNum := 1,  sOutput := ""
  While (A_Args.RunCMD.PID + DllCall("Sleep", "Int",0))
    and DllCall("PeekNamedPipe", "Ptr",hPipeR, "Ptr",0, "Int",0, "Ptr",0, "Ptr",0, "Ptr",0)
        While A_Args.RunCMD.PID and (Line := File.ReadLine())
          sOutput .= Fn ? Fn.Call(Line, LineNum++) : Line

  A_Args.RunCMD.PID := 0
, hProcess := NumGet(PI, 0)
, hThread  := NumGet(PI, A_PtrSize)

, DllCall("GetExitCodeProcess", "Ptr",hProcess, "PtrP",ExitCode:=0)
, DllCall("CloseHandle", "Ptr",hProcess)
, DllCall("CloseHandle", "Ptr",hThread)
, DllCall("CloseHandle", "Ptr",hPipeR)

, ErrorLevel := ExitCode

Return sOutput
}

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


GetNumberFormat(Value, removeColon := true, Locale := 0x0400) {

	; ===============================================================================================================================
	; Function ...: GetNumberFormat
	; Return .....: Formats a number string as a number string customized for a locale specified by identifier.
	; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getnumberformata
	; ===============================================================================================================================

	if (Size := DllCall("GetNumberFormat", "uint", Locale, "uint", 0, "str", Value, "ptr", 0, "ptr", 0, "int", 0)) {
		VarSetCapacity(NumberStr, Size << !!A_IsUnicode, 0)
		if (DllCall("GetNumberFormat", "uint", Locale, "uint", 0, "str", Value, "ptr", 0, "str", NumberStr, "int", Size))
			return (removeColon = 1 ? RegExReplace(NumberStr, "\,.*") : removeColon < 1 ? RegExReplace(NumberStr, "[,0]*$", "$1") : NumberStr)
	}
	return false
}

ConvertBase(InputBase, OutputBase, nptr) {   ; Base 2 - 36

    static u := A_IsUnicode ? "_wcstoui64" : "_strtoui64"
    static v := A_IsUnicode ? "_i64tow"    : "_i64toa"
    VarSetCapacity(s, 66, 0)
    value := DllCall("msvcrt.dll\" u, "Str", nptr, "UInt", 0, "UInt", InputBase, "CDECL Int64")
    DllCall("msvcrt.dll\" v, "Int64", value, "Str", s, "UInt", OutputBase, "CDECL")
    return s
}

GetHex(hwnd) {                                                                                                                    	;-- Umwandlung Dezimal nach Hexadezimal
return Format("0x{:x}", hwnd)
}

#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\class_json.ahk

/*

;~ propLen	:= stats.LenPropfile/1024	< 10 ? stats.LenPropfile	" bytes" : Round(stats.LenPropfile/1024, 1)	" KB"
	;~ FListLen 	:= stats.LenFList/1024    	< 10 ? stats.LenFList     	" bytes" : Round(stats.LenFList/1024, 1)     	" KB"

	;~ SciTEOutput(	"filecount       `t: "  	GetNumberFormat(stats.filecount, 1)                                  	"`n"
					;~ . 	"overall filesize`t: " 	GetNumberFormat(Round(stats.size/1024, 1), 0) " KB"           	"`n"
					;~ . 	"processed     `t: "  	GetNumberFormat(stats.processed, 1)                                   	"`n"
					;~ . 	"overall words`t: "  	GetNumberFormat(stats.words, 1)                                       	"`n"
					;~ . 	"overall chars`t: "  	GetNumberFormat(stats.chars, 1)                                          	"`n"
					;~ . 	"words/file     `t: "  	GetNumberFormat(Round(stats.words / stats.filecount,1), 0)	"`n"
					;~ . 	" -------------------------------------------"                                                                   	"`n"
					;~ . 	"Len prop file `t: "  	propLen                                                	"`n"
					;~ . 	"Len File List `t: "   	FListLen                                                	"`n" )


 */


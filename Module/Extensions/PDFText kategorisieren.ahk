; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	PDFText Kategorien Editor
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	   	-	soll helfen eigene Fachwörterbücher und Eigennamenlisten von Krankenhäusern, Abteilungen, Fachkollegen zu erstellen
;									-	gedacht für eine vollautomatische Benennung eingegangener Befund
;
; 		Beschreibung:       -	Skript extrahiert Text aus PDF Dateien und nutzt dafür Layout-Positionen um zusammenhängende Textbereiche zu erfassen
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
	SetBatchLines, -1
	ListLines, Off
	;}

  ; ----------------------------------------------------------------------------------------------------------------------------------------------
  ; Variablen
  ; ----------------------------------------------------------------------------------------------------------------------------------------------;{
		OutputDebug % "Test"
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
		properties := {	"base"               	: { "dir"         	: "M:\Befunde\Text"
															,	"ext"            	: "txt"
															,	"file_limit"    	: 200                            ; limit files to read at once
															,	"recursive" 	: true}
							,	"collection"        	: {"path"         	: dataPath
															,	"name"      	: "medical1"
															,	"description"	: "dictionary for german medical words"}
							,	"categories"       	: ["normal", "Fachwort", "Eigenname(n)"]}

	;}

  ; creates a new or edits an existing dictionary collection
	pdf := new Textorizer(properties)

 ; some output for debugging purposes
	stats := pdf.stats

	propLen	:= stats.LenPropfile/1024	< 10 ? stats.LenPropfile	" bytes" : Round(stats.LenPropfile/1024, 1)	" KB"
	FListLen 	:= stats.LenFList/1024    	< 10 ? stats.LenFList     	" bytes" : Round(stats.LenFList/1024, 1)     	" KB"

	SciTEOutput(	"filecount       `t: "  	GetNumberFormat(stats.filecount, 1)                                  	"`n"
					. 	"overall filesize`t: " 	GetNumberFormat(Round(stats.size/1024, 1), 0) " KB"           	"`n"
					. 	"processed     `t: " 	GetNumberFormat(stats.processed, 1)                                   	"`n"
					. 	"overall words`t: " 	GetNumberFormat(stats.words, 1)                                       	"`n"
					. 	"overall chars`t: "  	GetNumberFormat(stats.chars, 1)                                          	"`n"
					. 	"words/file     `t: "  	GetNumberFormat(Round(stats.words / stats.filecount,1), 0)	"`n"
					. 	" -------------------------------------------"                                                                   	"`n"
					. 	"Len prop file `t: "  	propLen                                                	"`n"
					. 	"Len File List `t: "   	FListLen                                                	"`n" )


ExitApp
return


class Textorizer {

    ; Initializing, collecting filenames
	__New(properties) {

		;this.OnErrorFunc := Func(this.ErrorFunc())
		;OnError(ObjBindMethod(this, this.OnErrorFunc))

			if !IsObject(properties) {
				this.emsg := "initializing of new Textorizer object-class failed`n> PARAMETER PROPERTIES MUST BE AN OBJECT <"
				throw  Exception(this.emsg, -1, "`rclass Textorizer - line number: "  A_LineNumber - 7)
			}

			this.prop                	:= Object()
			this.prop.base          	:= Object()
			this.dict                    	:= Object()
			this.stats                   	:= Object()
			this.files                  	:= Array()

			this.size                 	:= 0
			this.words             	:= 0
			this.chars              	:= 0
			this.processed         	:= 0
			this.filecount          	:= 0

		; store properties
			this.prop.collection	:= {	"path"       	: RegExReplace(properties.collection.path, "\\$")
												,	"name"      	: properties.collection.name
												,	"description" 	: properties.collection.description}
			this.prop.categories	:= properties.categories

		; filepaths used
			this.prop.base.dir      	:= RegExReplace(properties.base.dir, "\\$")
			this.prop.base.ext      	:= RegExReplace(properties.base.ext, "\.")
			this.prop.base.recusive	:= properties.base.recursive
			this.prop.base.file_limit	:= properties.base.file_limit
			this.prop.filebase       	:= this.prop.base.dir "\*." this.prop.base.extension
			this.prop.dict             	:= this.prop.collection.path "\" this.prop.collection.name "-dictionay.json"
			this.prop.FList            	:= this.prop.collection.path "\" this.prop.collection.name "_filelist.json"
			this.prop.props            	:= this.prop.collection.path "\" this.prop.collection.name "_properties.json"

		; create dictionary object if not exist otherwise load data
			If FileExist(this.prop.dict)
				this.LoadDictionary()
			else {

			}

		; load file list if exists
			If FileExist(this.prop.FList)
				this.LoadFileList()

		; populate file list with changed or new files
			stopAt := this.files.Count()
			Loop, Files, % this.prop.filebase
				If !this.FileInList(A_LoopFileShortName, this.prop.base.dir, StopAt) {

					txt := FileOpen(A_LoopFileFullPath, "r").Read()
					this.files.Push({	"basedir"		: this.prop.base.dir
										,	"filename" 	: A_LoopFileShortName
										,	"extension" 	: A_LoopFileExt
										,	"modified" 	: A_LoopFileTimeModified
										, 	"created"   	: A_LoopFileTimeCreated
										, 	"processed"	: false
										, 	"size"         	: A_LoopFileSize
										, 	"words"      	: StrSplit(RegExReplace(txt, "[\n\r]", " "), " ").MaxIndex()
										, 	"chars"       	: StrLen(RegExReplace(txt, "[\n\r\s]", " "))})

				}

		; do some statistics
			this.FilesStatistic()

		; saves file list and properties
			this.SaveProperties()

	}

	; functions
		FilesStatistic() {

			this.stats.size := this.stats.words := this.stats.chars := this.stats.processed:= 0

			If !IsObject(this.files) || (this.files.Count() = 0)
				return

			this.stats.filecount := this.files.Count()
			For fidx, data in this.files
					this.stats.size         	+= data.size
				, 	this.stats.words      	+= data.words
				, 	this.stats.chars       	+= data.chars
				, 	this.stats.processed	+= data.processed

		}

	; gui
		AddWords() {

			Gui, TXR: new, +hwndhTXR


		}

	; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; internal use	  +++++++++++++++++++++++++++++++++++++++++++++++++++++++;{
	; ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		FileInList(filename, basedir, StopAt)     {                    ; checks if file ist already listed, returns position in array

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

		Filesize(fullfilepath)             	{
			FileGetSize, filesize, % fullfilepath
		return filesize
		}

		SaveProperties()                 	{

			propsave := FileOpen(this.prop.props, "w", "UTF-8")
			propsave.Write(JSON.Dump(this.prop,,	2))
			this.stats.LenPropfile := propsave.Length
			propsave.Close()

			propsave := FileOpen(this.prop.FList, "w", "UTF-8")
			propsave.Write(JSON.Dump(this.files	,,	2))
			this.stats.LenFList := propsave.Length
			propsave.Close()

		}

		LoadFileList()                     	{

		  ; prevent's loading of empty data
			If !FileExist(this.prop.FList) || (this.Filesize(this.prop.FList) = 0)
				return

		  ; check this.files object
			If !IsObject(this.files)
				this.files := Array()

		  ; load Filelist object
			fileO := FileOpen(this.prop.FList, "r", "UTF-8")
			this.stats.LenFList := fileO.Length
			fList := JSON.Load(fileO.Read())
			fileO.Close()

 		  ; detect file changes, refresh stats - populate files array
			For fidx, data in fList {

				fpath := this.prop.base.dir "\" data.filename "." data.extension
				If !FileExist(fpath)
					continue

				FileGetSize	, size     	, % fpath
				FileGetTime	, Modified	, % fpath, M

				If (data.size <> size) || (data.modified <> modified) {

					txt                    	:= FileOpen(fpath, "r", "UTF-8").Read()
					data.size	    	:= size
					data.modifed	:= modified
					data.words    	:= StrSplit(RegExReplace(txt, "[\n\r]", " "), " ").MaxIndex()
					data.chars     	:= StrLen(RegExReplace(txt, "[\n\r\s]", " "))

				}

				this.files.Push(data)

			}

		}

		LoadDictionary()                	{

		  ; check this.files object
			If !IsObject(this.dict)
				this.dict := Object()

			this.dict := JSON.Load(FileOpen(this.prop.dict, "r", "UTF-8").Read())

		}

		SaveDictionay(backup=10)	{               	; ### backup!!

			If !IsObject(this.dict)
				return

			 FileOpen(this.prop.dict, "w", "UTF-8").Write(JSON.Dump(this.dict,,	2))

		}
	;}

	; callback
		ErrorFunc() {

			MsgBox,1, % "fatal error", % this.emsg "`nExit script"
			ExitApp

		}

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




; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	Addendum_DBASE - V1.2 beta
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	    	-	Analyse der für Albis verwendeten DBASE-Datei Strukturen
;									-	Portierung von Daten
;									-	Suche nach Daten
;
; 		Beschreibung:       -	enthält Funktionen ausschließlich mit lesendem Zugriff auf DBASE (.dbf) Dateien im '\albiswin\db\'-Ordner
;									-	keine Verwendung von Datenbanktreibern
;									-	aufgrund der Klassenstruktur können mehrere Dateien gleichzeitig lesend geöffnet werden
;									-
;
;       Abhängigkeiten:   	-	\lib\[SciteOutPut, class_JSON]
;
;       Hinweise:          	-	ich übernehme keine Haftung bei Schäden an Ihren Datenbankdateien. Sichern Sie die Daten vor Nutzung der Bibliothek!
;									-	Klasse gibt noch keinen Fehler bei Aufruf nicht vorhandener Funktionen aus!
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_DBASE begonnen am:          	12.11.2020
;       Addendum_DBASE letzte Änderung am: 	26.03.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
class DBASE {                                     ; native DBASE Klasse nur für Albis .dbf Files

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + Construction, Destruction
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	  ; stores data from Table File Header and gets the Field Descriptor Array
		__New(filepath, debug=false, debugGui="") {

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; open dbase file for reading
			; ------------------------------------------------------------------------------------------------------------------------------------------ ;{
				SplitPath, filepath,,,, baseName

				this.headerrecordgap	:= 1                                            ; one byte after DBF Header, purpose?
				this.baseName           	:= baseName
				this.filepath                	:= filepath
				this.encoding            	:= "CP1252"
				this.debug                    	:= debug
				this.debugGui           	:= debugGui
				this.nofileaccess         	:= false

			; max. bytes for working array - you can easily change it, without touching this class
				this.maxCapacity        	:= 400 * 1024000 ; = 200 MB

			; closes file access on script exit to prevent file damage
				OnExit(ObjBindMethod(this, "_Exit"))

			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; Table File Header (Length: 32 bytes)
			; ------------------------------------------------------------------------------------------------------------------------------------------ ;{
				If !IsObject(this.dbf := FileOpen(filepath, "r", "CP1252")) {
					throw "open database file: " filepath " failed!"
					ExitApp
				}

			  ; Byte: [    0   ]	- contains the version of this dBASE file
				this.Version   	:= this._ReadNum(1, "Int")
			  ; Byte: [  1-3  ]	- Date of last update, in YYMMDD format. Each byte contains the number as a binary.
				year  	:= this._ReadNum(1, "Int")
				month	:= this._ReadNum(1, "Int")
				day   	:= this._ReadNum(1, "Int")
				this.lastupdate	:= Format("{:02X}", year) Format("{:02X}", month) Format("{:02X}", day)
				year		:= 1900 + year
				month	:= SubStr("00" month, -1)
				day   	:= SubStr("00" day, -1)
				this.lastupdateDate 	:= day "." month "." year
				this.lastupdateEng 	:= year "." month "." day

			  ; Byte: [  4-7  ]	- Number of records in the table. (Least significant byte first.)
				this.records   	:= this._ReadNum(4, "Int")
			  ; Byte: [  8-9  ]	- Number of bytes in the header. (Least significant byte first.)
				this.headerLen 	:= this._ReadNum(2, "Short")
			  ; Byte: [10-11]	- Number of bytes in the header. (Least significant byte first.)
				this.lendataset 	:= this._ReadNum(2, "Short")
			  ; Byte: [12-27]	- I think these bytes are unused for Albis
				this.dbf.Seek(16, 1)
			  ; Byte: [   28  ]	- Production MDX flag. 0x01 if a .mdx file exists for this table or 0x00 if not.
				this.mdx        	:= this._ReadNum(1, "Int")
			  ; Byte: [29-31]	- I think these bytes are unused for Albis
				this.dbf.Seek(3, 1)

			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; Field Properties Structure (each field: 32 bytes, field count: ( DBF Header Length - TFH Length ) / 32 [bytes]
			; ------------------------------------------------------------------------------------------------------------------------------------------ ;{

				this.fields      	:= Array()        ;
				this.dbfields   	:= Object()
				this.dbstruct   	:= Object()
				this.NrFields  	:= Round((this.headerLen - 32) / 32)

				fpos  	:= 0
				lenNF 	:= StrLen(this.NrFields) - 1
				dbfields	:= ""

				Loop % this.NrFields {

					flabel	:= this._ReadString(11)                 	; 11 bytes:	Field name in ASCII (zero-filled).
					ftype 	:= this._ReadString(1)                   	;   1 byte:	Field type in ASCII (B, C, D, N, L, M, @, I, +, F, 0 or G).
					this.dbf.Seek(4, 1)                                    	;           	skip 4 unused bytes
					flen   	:= this._ReadNum(1, "UChar")   	;	1 byte:	Field length in binary.
					this.dbf.Seek(14, 1)                                  	;            	skip 14 unused bytes
					fmdx 	:= this._ReadNum(1, "UChar")      	;	1 byte:	maybe this field is indexed in .mdx file

					this.fields.Push({"label": flabel, "type":ftype, "start":(fpos+1), "len":flen, "more":fmdx})
					this.dbfields[fLabel] := {"type":ftype, "start":(fpos+1), "len":flen, "more":fmdx, "pos":A_Index}
					dbfields .= flabel ","

					 If (this.debug = 4)
						x .= " > Pos: " SubStr("000" fpos, -2) "-" SubStr("000" fpos+flen-1, -2) " (" SubStr("00" flen, -1) ") [" ftype "]  - " flabel "`n"

					fpos += flen

				}

				this.dbstruct := {	"01 DBASE Version"                  	: this.Version
										,	"02 last update hex"                 	: this.lastupdate
										,	"03 last update date"               	: this.lastupdateDate
										,	"04 count of records"                	: this.records
										,	"05 length of header (bytes)"  	: this.headerLen
										,	"06 length one dataset (bytes)" 	: this.lendataset
										,	"07 uses MDXTable"                 	: this.mdx = 1 ? "true" : "false"
										,	"08 count of fields"                  	: this.NrFields
										,	"09 field names"                       	: RTrim(dbfields, ",") }

			;}

			; close this DBASE file after getting all data
				this.CloseDBF()

	    	; something to debug here
				If (this.debug = 4) {
					x .= " > average length: " fpos + 1 " - additional char *"     	"`n"
					t .= " ----------------------------------------------------------------" 	"`n"
					t .= " DBASE Version:    `t"    	Format("0x{:02X}", this.Version)	"`n"
					t .= " last Update?:    `t"      	Format("{:X}", this.lastupdate)  	"`n"
					t .= " Count of Records:`t"    	this.records                            	"`n"
					t .= " Header Length: `t"      	this.headerLen                      	"`n"
					t .= " Length Dataset:  `t"    	this.lendataSet                        	"`n"
					t .= " MDX file:            `t"    	(this.mdx ? "true":"false")           	"`n"
					t .= " Number of Fields:`t"    	this.NrFields                          	"`n"
					t .= x
					t .= " ----------------------------------------------------------------"
					SciTEOutput(t)
				}

			; calculations for viewing of progress
				this.ShowAt	:= Round(this.records/50)
				this.slen     	:= -1*(StrLen(this.records) - 1)
				this.maxRecs	:= SubStr("00000000" this.records, this.slen)

			; calculations of array size and filepointer position of first record in this database
				this.maxBytes	:= this.lendataset * this.records > this.maxCapacity ? this.maxCapacity : this.lendataset * this.records
				this.recordsStart	:= this.headerlen + this.headerrecordgap

			; recordbuf
				VarSetCapacity(this.recordbuf, this.lendataset, 0x20)

		}

	  ; file access will be closed, if this object is destroyed
		__Delete() {

			SciTEOutput("Object: " this.baseName " was deleted.")
			If IsObject(this.dbf)
				this.CloseDBF()

		}

	;}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + Database file access functions
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; starts read-access to database, automatic seek's to position of first record
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		OpenDBF() {

			If !IsObject(this.dbf := FileOpen(this.filepath, "r", "CP1252")) {
				throw "open database file: " this.filepath " failed!"
				ExitApp
			}

			this.dbf.Seek(this.recordsStart, 0)

		return this.dbf.Tell()
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; stops read-access to database, without destroying of DBASE class object
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		CloseDBF() {

			filepos := this.dbf.Tell()
			this.dbf.Close()
			this.dbf := ""

		return filepos
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; reads one record set
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		ReadRecord(outfields="") {

			bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
			set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)
			this.recordnr := Floor((this.dbf.Tell() - this.headerlen) / this.lendataset)

			If !IsObject(outfields)
				return set

			strobj := Object()
			strobj["recordNr"]	:= Floor((this.dbf.Tell() - this.headerlen) / this.lendataset)
			For index, fLabel in outfields
			If (StrLen(txt := Trim(SubStr(set, this.dbfields[fLabel].start, this.dbfields[fLabel].len))) = 0)
				continue
			else
				strobj[flabel] := txt

		return strobj
		}

	;}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + retrieve data from database records
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; REGEX BASED SEARCH, RETURNS ANY MATCHING RECORD
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Search(pattern, startrecord=0, callbackFunc="") {

		/* main description

				- parameter: pattern is designed to use only RegEx for matching, use function SearchFields() for different search algorithms!

		*/

		/* Description for using the function with less RAM usage

				The file size of a database can exceed the size limit of an Autohotkey variable. This function contains a kind of stop-resume read mode.

				- After a specified amount of memory has been reached, the number of the data record last read is transferred to the this.brkrec variable.
				- Then the entire array with the search results is returned to the caller.
				- The content of this.brkrec should be passed as the value for the startrecord parameter in this function.
				- By the way, startrecord can be used the first time you call this function, e.g. because you know the exact starting position of a data block
				- Use the phrase "use last RegEx" for the options parameter to reuse the RegEx string previously used for the search.
				- In combination with "use last RegEx" and a numerical value for the number of a data record, the search continues without detours at this point.

		*/

		; Initializing vars
			static recordbuf, rxStr
			VarSetCapacity(recordbuf	, this.lendataset, 0x20)
			VarSetCapacity(matches	, this.maxBytes)

			this.uselastRxStr := false
			;this.search.callbackFunc := Func("callbackFunc")

			matches        	:= Object()                                 	; collects the findings
			this.breakrec 	:= "#"                                       	; this variable is used to help external processing of data

		; Establish read access if not already done
			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.filepos := this.OpenDBF()
			}

		; builds a regex string, regex string can be re used by command
			If IsObject(pattern) && !this.uselastRxStr {

				this.hits	:= 0                                            	; hits counter

			; build routine
				rxStr := "^", rxPre := 0
				For index, field in this.fields
					If pattern.haskey(field.label) {

						if (rxPre > 0)
							rxStr .= ".{" rxPre "}", rxPre := 0

						If !RegExMatch(pattern[field.label], "^\s*rx.*?\:") {
							If (field.type = "N")
								rxStr .= "\s{" (field.len - StrLen(pattern[field.label])) "}" pattern[field.label]
							else
								rxStr .= pattern[field.label] ".{" (field.len - StrLen(pattern[field.label])) "}"
						}
						else {
							If (field.type = "D")
								rxStr .= RegExReplace(pattern[field.label], "^\s*rx.*?\:")
							else
								rxStr .= ".*" RegExReplace(pattern[field.label], "^\s*rx.*?\:") ".*"
						}

						lastrxLen := StrLen(rxStr)

					}
					else
						rxPre += field.len

			; needed replacements
				rxStr := SubStr(rxStr,1, lastrxLen)
				rxStr := RegExReplace(rxStr, "(\.\*){2,}", ".*")
				rxStr := RTrim(rxStr, ".*")
				;rxStr .= ".{" rxPre "}"

			; return on failure
				If (StrLen(rxStr) = 0) {
					return "can't build RegEx search string"
				}

			; neede for debugging
				If (this.debug = 1)
					SciTEOutput(rxStr)

			; publish some vars
				this.lastrxLen         	:= lastrxLen
				this.SearchRegExStr	:= rxStr

			} else If !IsObject(pattern) {
				throw "Parameter: pattern must be an object"
			}

		; seek if parameter is greater than zero
			If (startrecord > 0) {
				this.dbf.Seek(this.lendataset * startrecord, 1)
				this.filepos := this.dbf.Tell()
			}

		; search loop
			while (!this.dbf.AtEOF) {

			; reads a data set raw mode and converts it to a readable text format
				bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
				set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

			; debugging and callback function to show progress

				If (Mod(setNR := A_Index, this.ShowAt) = 0) {
					If (this.debug = 1)
						ToolTip, % "Search rx: " rxStr "`n" SubStr("00000000" A_Index, this.slen) "/" SubStr("00000000" this.records, this.slen)
					If (StrLen(callbackFunc) > 0)
						%callbackFunc%(setNR, this.records, this.slen, matches.MaxIndex())
				}

			; no match - continue
				If !RegExMatch(set, rxStr)
					continue

			; temp object collects dataset field data
				strobj := Object()
				strobj["removed"]	:= InStr(SubStr(set, -1), "*") ? true : false
				strobj["recordNr"]	:= Floor((this.dbf.Tell() - this.headerlen) / this.lendataset)
				For index, field in this.fields
					If (StrLen(txt := Trim(SubStr(set, field.start, field.len))) = 0)
						continue
					else
						strobj[field.label] := txt

			; append match object
				matches.Push(strobj)
				this.hits ++

		  ; maximum array entries, store filepointer position for later use and return matches
		    	If (this.lendataset * A_Index = this.maxbytes) {
					this.filepos	:= this.dbf.Tell()                          	; last position of filepointer
					this.breakrec	:= startrecord + A_Index            	; position of last read dataset
					ToolTip
					return matches
				}

			}

		; file access is automatically terminated if the function established this independently
			If this.nofileaccess {
				VarSetCapacity(recordbuf, 0)
				this.nofileaccess := false
				this.CloseDBF()
			}

		; indicates that the end of the file has been reached
			this.breakrec	:= "!"

			ToolTip
			If (StrLen(callbackFunc) > 0)
				%callbackFunc%(setNR, this.records, this.slen, matches.MaxIndex())

		return matches
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; Each record is first divided into fields which can be searched individually.
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		SearchExt(pattern, outfields, startrecord=0, options="", callbackFunc="") {

		/* Extended search

			- - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Date comparison / filter by date
			- - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Comparison is possible with ranges or Instr() operations. One date can be compared with another
			by only showing the dates that are greater, less or equal. This also works with pure numerical values.
			Above calculation is only triggered by field names containing words like 'DATUM' or 'DATE'.
			Using the field type instead of matching inside field name would be better in most cases, but is yet not implemented.

			- - - - - - - - - - - - - - - - - - - - - - - - - - - -
			filter return values
			- - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Outfields parameter is an array of field names you want to retrieve.

			- - - - - - - - - - - - - - - - - - - - - - - - - - - -
			callbackFunc
			- - - - - - - - - - - - - - - - - - - - - - - - - - - -
			If you want to use this, make your function variadic like 'MyCallbackFunc(p*)'
			4 value will be given to your callback function, with variadic mode you can shorten your code.

		 */

			If !IsObject(pattern)
				throw "Parameter: pattern must be an object"

		; initialize some vars
			VarSetCapacity(matches	, this.maxBytes)
			VarSetCapacity(recordbuf	, this.lendataset*100, 0x20)

			this.hits 	:= 0
			matches 	:= Array()

		; establish read access if not already done
			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.pos := this.OpenDBF()
			}

		; filepointer is set to the data record with the number from startrecord
			If (startrecord > 0) {
				this.dbf.Seek(this.lendataset * startrecord, 1)
				this.filepos := this.dbf.Tell()
			}

		; prepare pattern operations
			For pLabel, pCondition in pattern {

				If RegExMatch(pLabel, "[A-Z]+(DATUM|DATE)") {
					mObj:=Object()
					RegExMatch(pCondition, "O)(?<cd1>[!><=]+)(?<date1>\d{8})(?<AO>[&|]+)*(?<cd2>[!><=]+)*(?<date2>\d{8})*", m)
					Loop % m.Count()
						mObj[m.Name(A_Index)] := m.Value(A_Index)
					mObj.CCount := mObj.date1 && mObj.date2 ? 2 : 1
					pattern[pLabel] := mObj
				}
				else If RegExMatch(pCondition, "^rx:(?<Str>.*)", rx) {
					mObj:=Object()
					mObj.rx := rxStr
					pattern[pLabel] := mObj
				}

			}

		; prepare dataset operation. Uses every field in the comparison process or
		; only those that have been supplied with the pattern parameter
			If !IsObject(outfields) && RegExMatch(outfields, "all fields")
				outfields := ""
			this.retSubs := this.BuildSubstringData(outfields)

		; search loop
			while (!this.dbf.AtEOF) {

				; reads one dataset from database
					bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
					set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

				; debug some actions
					If (StrLen(callbackFunc) > 0) && (Mod(A_Index, this.ShowAt) = 0)
						%callbackFunc%(A_Index, this.records, this.len, matches.Count())

							;ToolTip, % "Get:`t" SubStr("00000000" A_Index + startrecord, this.slen) "/" this.maxRecs "`nmatches: "

				; compare pattern field condition with dataset
					pLabelHit := 0
					For pLabel, pCondition in pattern {

						val := Trim(SubStr(set, this.dbfields[pLabel].start, this.dbfields[pLabel].len))
						If RegExMatch(pLabel, "[A-Z]+(DATUM|DATE)") {

							cdhits := 0
							Loop % pCondition.CCount {

								cdate   	:= pCondition["date" A_Index]
								conds 	:= StrReplace(pCondition["cd" A_Index], "!")
								negate 	:= InStr(pCondition["cd" A_Index], "!") ? false : true
								cd	    	:= StrSplit(conds)

								Loop % cd.Count() {

									cda := cd[A_Index]
									cdmatched := (cda="<" && val<cDate) ? true : (cda=">" && val>cDate) ? true : (cda="=" && val=cDate) ? true : false
									If (!cdmatched && A_Index < cd.Count())
										continue
									else if cdmatched {
										cdhits ++
										break
									}

								}

							}

							If (cdhits = pCondition.CCount)
								pLabelHit ++
							else
								break

						}
						else If IsObject(pCondition) {
							If RegExMatch(val, pCondition.rx)
								pLabelHit ++
							else
								break
						}
						else if (val = pCondition)
							pLabelHit ++
						else
							break

					}

					If (pLabelHit = pattern.Count())
						matches.Push(this.GetSubData(set, this.retSubs))

			}

		return matches
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; Simple matching for fast results
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		SearchFast(pattern, outfields, startrecord=0, options="") {

		; comparison is possible with ranges or Instr() operations. One date can be compared with another
		; by only showing the dates that are greater, less or equal. This also works with pure numerical values.

			If !IsObject(pattern)
				throw "Parameter: pattern must be an object"

		; initialize some vars
			VarSetCapacity(matches 	, this.maxBytes)
			VarSetCapacity(recordbuf	, this.lendataset*100, 0x20)

			this.hits 	:= 0
			matches 	:= Array()

		; establish read access if not already done
			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.pos := this.OpenDBF()
			}

		; filepointer is set to the data record with the number from startrecord
			If (startrecord > 0) {
				this.dbf.Seek(this.lendataset * startrecord, 1)
				this.filepos := this.dbf.Tell()
			}

		; prepare dataset operation. Uses every field in the comparison process or
		; only those that have been supplied with the pattern parameter
			If !IsObject(outfields) && RegExMatch(outfields, "all fields")
				outfields := ""
			this.retSubs := this.BuildSubstringData(outfields)

		; search loop
			while (!this.dbf.AtEOF) {

				; debug some actions
					If (this.debug > 0) && (Mod(A_Index, this.ShowAt) = 0)
						If (this.debug = 1)
							ToolTip, % "Get:`t" SubStr("00000000" A_Index + startrecord, this.slen) "/" this.maxRecs "`nmatches: " matches.Count()

				; reads one dataset from database
					bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
					set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

				; compare pattern field condition with dataset
					hits := 0 ;, strobj := Object()
					For Lbl, sstr in pattern {
						val := Trim(StrGet(&recordbuf + this.dbfields[Lbl].start - 1, this.dbfields[Lbl].len, this.encoding))
						If (val = sstr)
							hits ++
					}

					If (hits = pattern.Count()) {
						matches.Push(this.GetSubData(set, this.retSubs))
					}

			}

		return matches
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; GET DATA FROM EVERY RECORD CONTAINING CERTAIN FIELD NAMES
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		GetFields(getfields) {

		; the function returns the data of all records as key:val object
		; you have to specify getfields as an array containing fieldnames

		; prepares some variables
			VarSetCapacity(recordbuf, this.lendataset, 0x20)
			data	:= Object()

		; creates an array with data for substring function retreave only fields that should be returned
			subs := this.BuildSubstringData(getfields)
			If (subs.MaxIndex() = 0)                            	; returns here if all field names passed are unknown
				return 2                                            	; "field label(s) unknown"

		; establish read access if not already done
			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.pos := this.OpenDBF()
			}

		; compile field data as an object
			while (!this.dbf.AtEOF) {

				; Tooltip progress view
					If (this.debug > 0) && (Mod(A_Index, this.ShowAt) = 0)
						If (this.debug > 0)
							ToolTip, % "GetFields:`t" SubStr("00000000" A_Index, this.slen) "/" this.maxRecs "`ndatasets: " data.Count()

				; reads a data set raw mode and converts it to a readable text format
					bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
					set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

				; push object to data array (object contains the requested field names, values and its specific recordnr)
					data.Push(this.GetSubData(set, subs))

			}

			ToolTip

		return data
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; READ A BLOCK OF DATA SETS
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		ReadBlock(NrOfSets, getfields:="", StartBlock=0) {

		; the function returns the data of all records as key:val object
		; you have to specify getfields as an array containing fieldnames

			static data, recordbuf

		; prepares some variables
			If !IsObject(this.data) || IsObject(getfields)  {

					VarSetCapacity(recordbuf, this.lendataset, 0x20)
					this.getfields	:= getfields

				; creates an array with data for substring function retreave only fields that should be returned
					this.subs       	:= this.BuildSubstringData(getfields)
					If (this.subs.MaxIndex() = 0)                               	; returns here if all field names passed are unknown
						return 2                                                        	; "field label(s) unknown"

			}
			this.data	:= Object()


		; establish read access if not already done
			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.pos := this.OpenDBF()
			}

		; filepointer is set to the data record with the number from startrecord
			If (StartBlock > 0) {
				StartRecord := ((StartBlock-1)*NrOfSets)+1
				this.dbf.Seek(this.lendataset * startrecord, 1)
				this.filepos := this.dbf.Tell()
			}

		; compile field data as an object
			blockpos := 0
			while (!this.dbf.AtEOF)  {

				; reads a data set raw mode and converts it to a readable text format
					bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
					set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)
					this.filepos := this.dbf.Tell()

				; push object to data array (object contains the requested field names, values and its specific recordnr)
					fields := this.GetSubData(set, this.subs)
					If (fields["id"] = -1) || fields["remove"]
						continue
					this.data.Push(fields)
					blockpos ++
					If (blockpos = NrOfSets)
						break
			}

		return this.data
		}


	;}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + create own index file based on parameters
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{
		CreateIndex(IndexMode, IndexFilePath, IndexFieldName:="", ReIndex:=false) {

			; creates or updates an index file with fileseek positions at the first occurrence of a certain string (date, string, quarter)
			; the function currently only indexes by quarters
			; if an index file exists, it only updates the index from the last entry
			; IndexFieldName:		The database field name which is indexed. At the moment only fields with date strings can be processed.

				IndexMode := "quarter"

			; check indexFilePath and throw error if not valid or if it doesn't exist
				If !RegExMatch(IndexFilePath, "i)[A-Z]\:\\")
					throw A_ThisFunc ": Parameter [indexFilepath] is not valid!`n" IndexFilePath
				SplitPath, IndexFilePath, idxFName, idxFDir
				If !InStr(FileExist(idxFDir "\"), "D")
					throw A_ThisFunc ": The file path does not exist!`n" idxFDir

			; create new index if doesn't exist or get's last indexed fileposition
				If !FileExist(IndexFilePath) || (ReIndex = true) {

					ReIndex     	:= true
					startrecord	:= 0
					startfilepos	:= 0
					iFile          	:= Object()

				} else {

					iFile := JSON.Load(FileOpen(indexFilePath, "r", "UTF-8").Read())
					For QZ, StartRecord in iFile
						continue
					startfilepos := this._GetFilePosFromRecord(StartRecord)
					SciTEOutput("QZ: " QZ ", seek: " seekpos, "StartRecord: " StartRecord)

				}

			; Establish read access if not already done
				If !IsObject(this.dbf) {
					this.nofileaccess := true
					this.filepos := this.OpenDBF()
				}

			; indexing is only possible for one condition
				If (IndexMode = "quarter") {
					  FStart	:= this.dbfields[IndexFieldName].start
					  FLen	:= this.dbfields[IndexFieldName].len
					  KeyN	:= "QNow"
				}

			; start indexing
				maxRecs := SubStr("00000000" this.records, this.slen)
				VarSetCapacity(recordbuf, this.lendataset, 0x20)
				while (!this.dbf.AtEOF) {

					; reads a data set raw mode and converts it to a readable text format
						bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
						set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

					; mode driven indexing
						If (IndexMode = "quarter") {
							val     	:= SubStr(set, FStart, FLen)
							IdxKey	:= SubStr(val, 1, 4) . Ceil(Substr(val, 5, 2)/3)
						}

					; Multi-debugging output: SciteWindow console or external gui
						If (this.debug > 0) && (Mod(A_Index, this.ShowAt) = 0) {

							If (this.debug = 1) {
								thispos := SubStr("00000000" A_Index + StartRecord, this.slen)
								ToolTip, % "DB"       	": "	dbname                	"`n"
											.	KeyN     	": " 	idxKEy                    	"`n"
											.	"Position"	":" 	thispos "/" maxRecs
							}
							else if (this.debugGui.Control = "Statusbar") {
								If (this.debugGui.sbNR > 0)
									SB_SetText("`t`t" Round(( (A_Index + StartRecord)/this.records)*100), this.debugGui.sbNR)
								else if StrLen(this.debugGui.sbNR = 0)
									SB_SetText(" " SubStr("00000000" A_Index + StartRecord, this.slen) "/" SubStr("00000000" this.records, this.slen) )
							}

						}

						If (idxKEy = 0 || StrLen(idxKEy) = 0 || idxKEy = idxKEy_old)
							continue
						idxKEy_old := idxKEy

						If !iFile.haskey(idxKEy)
							iFile[idxKEy] := [this._GetRecordFromFilepos(this.dbf.Tell())-1 , this.dbf.Tell()] 	   ; record nr, filepointer position

				}

			; file access is automatically terminated if the function established this independently
				If this.nofileaccess {
					VarSetCapacity(recordbuf, 0)
					this.nofileaccess := false
					this.CloseDBF()
				}

			; save index as json string
				JSONData.Save(indexFilePath, iFile, true,, 1, "UTF-8")

		return iFile
		}

	;}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + STRING FUNCTIONS
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	  ; --------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; creates an array with data for substring function to examine only certain fields. This restriction is essential for a faster search.
	  ; --------------------------------------------------------------------------------------------------------------------------------------------------------
		BuildSubstringData(getfields) {

			subs := Object()
			If IsObject(getfields) {    ; specified fields

				For FieldNr, fLabel in getfields
					For FieldsIndex, field in this.fields
						If (fLabel = field.Label) {
							subs.Push({"label":field.label, "type":field.type, "start":field.start, "len":field.len})
							break
						}

			} else {                        ; all fields

				For FieldsIndex, field in this.fields {
					subs.Push({"label":field.label, "type":field.type, "start":field.start, "len":field.len})
					If (this.debug > 2)
						SciTEOutput(" [" FieldsIndex "] " "label:" field.label ",type:" field.type ",start:" field.start ",len:" field.len)
				}

			}

		return subs
		}

	  ; --------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; builds an object containing only fields and values that are specified by BuildSubstringData() function
	  ; --------------------------------------------------------------------------------------------------------------------------------------------------------
		GetSubData(DataSet, SubstringData) {

			; this function is mostly used to reduce the size of data returned

			strobj := Object()
			strobj["removed"]	:= InStr(SubStr(set, -1), "*") ? true : false                              	; Dataset ist removed?
			strobj["recordnr"]	:= Floor((this.dbf.Tell() - this.recordsStart) / this.lendataset)    	; saves the number of record
			For index, substring in SubstringData
				strobj[substring.label] := Trim(SubStr(DataSet, substring.start, substring.len))

		return strobj
		}

	;}


	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + INTERNAL FUNCTIONS
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	; reads bytes from database and returns it as integer value
	_ReadNum(bytes, type) {
		VarSetCapacity(buffin, bytes, 0)
		this.dbf.RawRead(buffin, bytes)
	return NumGet(buffin, 0, type)
	}

	; reads bytes from database and returns the encoded string
	_ReadString(bytes) {
		VarSetCapacity(buffin, bytes, 0)
		this.dbf.RawRead(buffin, bytes)
	return StrGet(&buffin, bytes, this.encoding)
	}

	; reads one record
	_ReadRecordSet(nr, origin=2) {

		; 1 :	from first record
    	; 2 :	from current record

		static recordbuf
		If (VarSetCapacity(recordbuf) = 0)
			VarSetCapacity(recordbuf, this.lendataset, 0x20)

		If (nr > 0)
			this.pos := this._SeekToRecord(nr, origin)

		bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
		set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

	return set
	}

	; sets filepointer to new position based on record nr
	_SeekToRecord(nr, origin=2) {

		; 1 :	from first record
    	; 2 :	from current record

		If (origin = 1) {

			nr := (nr > this.records) ? this.records : (nr < 0) ? 0 : nr
			newposition := this.recordsStart + (nr * this.lendataset)
			If (this.dbf.Tell() = newposition)                                            	; file pointer is already in position
				return                                                                            	; return

			this.dbf.Seek(newposition, 0)

			return this.dbf.Tell()

		} else if (origin = 2) {

			cur_record := (this.dbf.Tell() - this.recordsStart) / this.lendataSet

			If (cur_record + nr > this.records)                                       	; new file pointer position is after EOF
				this.dbf.Seek(this.dbf.Length, 0)                                      	; seek to end of file
			else if (cur_record + nr < 0)                                              	; new file pointer position is before first record
				this.dbf.Seek(this.recordsStart, 0)                                     	; seek to first record
			else                                                                                    	; else
				this.dbf.Seek(cur_record + (nr * this.lendataset), 1)            	; seek

			return this.dbf.Tell()

		}

	}

	; converts the nr of record to file position
	_GetFilePosFromRecord(nr) {
		return Floor(this.recordsStart + (nr * this.lendataset))
	}

	; converts seekpos address to nr of record
	_GetRecordFromFilePos(filepos) {
		return Floor((filepos - this.recordsStart) / this.lendataset)
	}

	; OnExit function to ensure that file access to database is closed on exit
	_Exit(ExitReason, ExitCode) {

		If IsObject(this.dbf)
			this.dbf.Close()

	}


	;}

}



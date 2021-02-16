; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	Addendum_DBASE - V1.0 alpha
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Verwendung:	    	-	Analyse der für Albis verwendeten DBASE-Datei Strukturen
;									-	Portierung von Daten
;									-	Suche nach Daten
;
; 		Beschreibung:       -	enthält Funktionen ausschließlich mit lesendem Zugriff auf DBASE (.dbf) Dateien im '\albiswin\db\'-Ordner
;									-	keine Verwendung von Datenbanktreibern
;									-	aufgrund der Klassenstruktur können mehrere Dateien gleichzeitig lesend geöffnet werden
;									-
;
;       Abhängigkeiten:   	-	\lib\SciteOutPut.ahk
;
;       Hinweise:          	-	ich übernehme keine Haftung bei Schäden an Ihren Datenbankdateien. Sichern Sie die Daten vor Nutzung der Bibliothek!
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_DBASE started:    	12.11.2020
;       Addendum_DBASE last change:	19.11.2020
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


class DBASE {                                     ; native DBASE Klasse nur für Albis .dbf Files

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + Construction, Destruction
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	  ; stores data from Table File Header and gets the Field Descriptor Array
		__New(filepath, debug=false) {

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; open dbase file for reading
			; ------------------------------------------------------------------------------------------------------------------------------------------ ;{
				this.headerplus := {"BEFUND":1, "BEFMED":1}

				SplitPath, filepath,,,, baseName
				If !IsObject(this.dbf := FileOpen(filepath, "r", "CP1252")) {
					throw "open database file: " filepath " failed!"
					ExitApp
				}

				;this.headerrecordgap	:= this.headerplus[basename]
				this.headerrecordgap	:= 1
				this.filepath                	:= filepath
				this.encoding            	:= "CP1252"
				this.debug                    	:= debug

			  ; max. bytes for working array - you can easily change it, without touching this class
				this.maxCapacity        	:= 20 * 1024000 ; = 20 MB

				OnExit(ObjBindMethod(this, "_Exit"))

			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; Table File Header (Length: 32 bytes)
			; ------------------------------------------------------------------------------------------------------------------------------------------ ;{

			  ; Byte: [    0   ]	- contains the version of this dBASE file
				this.Version   	:= this._ReadNum(1, "Int")
			  ; Byte: [  1-3  ]	- Date of last update, in YYMMDD format. Each byte contains the number as a binary.
				this.lastupdate	:= this._ReadNum(3, "Int")
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

				this.fields    	:= Array()
				this.NrFields	:= Round((this.headerLen - 32) / 32)
				fpos := 0

				Loop % this.NrFields {

					flabel	:= this._ReadString(11)                 	; 11 bytes:	Field name in ASCII (zero-filled).
					ftype 	:= this._ReadString(1)                   	;   1 byte:	Field type in ASCII (B, C, D, N, L, M, @, I, +, F, 0 or G).
					this.dbf.Seek(4, 1)                                    	;           	skip 4 unused bytes
					flen   	:= this._ReadNum(1, "UChar")      	;	1 byte:	Field length in binary.
					this.dbf.Seek(14, 1)                                  	;            	skip 14 unused bytes
					fmdx 	:= this._ReadNum(1, "UChar")      	;	1 byte:	maybe this field is indexed in .mdx file

					this.fields.Push({"label": flabel, "type":ftype, "start":(fpos+1), "len":flen, "more":fmdx})

					 If this.debug
						x .= " > Pos: " SubStr("000" fpos + 1, -2) "-" SubStr("000" fpos+flen, -2) " (" SubStr("00" flen, -1) ") [" ftype "]  - " flabel "`n"

					fpos += flen

				}

			;}

			; close this file after getting all data
				this.dbf.Close()

			; something to debug here
				 If this.debug {
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

		}

	  ; file access will be closed, if this object is destroyed
		__Delete() {

			;SciTEOutput("Object was deleted.")
			MsgBox Object was deleted.
			If IsObject(this.dbf)
				this.dbf.Close()

		}

	;}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + Database file access functions
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	  ; starts read-access to database, moves filepointer to start position of first record
		Open() {

			If !IsObject(this.dbf := FileOpen(this.filepath, "r", "CP1252")) {
				throw "open database file: " this.filepath " failed!"
				ExitApp
			}

			this.recordsStart := this.headerlen + this.headerrecordgap
			this.dbf.Seek(this.recordsStart , 0)
			return this.dbf.Tell()

		}

	  ; stops read-access to database, without destroying of DBASE class object
		Close() {

			this.dbf.Close()
			return IsObject(this.dbf)

		}

	;}

	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	; + retrieve data from database records
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ ;{

	 ; returns every record matched
		Search(pattern, startrecord=0, export="") {

			static recordbuf

			VarSetCapacity(recordbuf, this.lendataset, 0x20)

			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.pos := this.Open()
			}

		  ; if export contains a path the matched lines will be exported, otherwise this function returns an array with all lines
			If RegExMatch(export, "[A-Z]\:\\") {
				If RegExMatch(export, "\.DBF$") || (export = this.filepath) {
					throw "File overwrite prevention!`nPath of export file is the same as path of database`n>> " export "<<"
					ExitApp
				}
				exportfile := FileOpen(export, "w", "UTF-8")
			}

		  ; builds a regex string
			If IsObject(pattern) {

				rxStr := "i)^"
				For index, field in this.fields
					If pattern.haskey(field.label) {

						If !RegExMatch(pattern[field.label], "^\s*rx.*?\:") {
							If (field.type = "N")
								rxStr .= ".{" (field.len - StrLen(pattern[field.label])) "}" pattern[field.label]
							else
								rxStr .= pattern[field.label] ".{" (field.len - StrLen(pattern[field.label])) "}"
						} else {
							If (field.type = "D")
								rxStr .= RegExReplace(pattern[field.label], "^\s*rx.*?\:")
							else
								rxStr .= ".*" RegExReplace(pattern[field.label], "^\s*rx.*?\:") ".*"
						}

						lastrxLen := StrLen(rxStr)

					}
					else
						rxStr .= ".{" field.len "}"

			}

			this.lastrxLen         	:= lastrxLen
			this.SearchRegExStr	:= rxStr := SubStr(rxStr, 1, lastrxLen)
			this.hits                 	:= 0

			this.maxBytes := this.lendataset * this.records > this.maxCapacity ? this.maxCapacity : this.lendataset * this.records
			VarSetCapacity(matches, this.maxBytes)
			matches := Array()
			while (!this.dbf.AtEOF) {

				bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
				set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

				If this.debug && (Mod(A_Index, 5000) = 0)
					ToolTip, % "p:`t" SubStr("0000000" A_Index, -6) "/" SubStr("0000000" this.records, -6) "`n" lfound

				If (StrLen(rxStr) > 0) && !RegExMatch(set, rxStr)
					continue

				this.hits ++
				;~ If this.debug
					;~ lfound := "f:`t" SubStr("0000000" A_Index, -6) "/" SubStr("0000000" this.records, -6)

				If !IsObject(exportfile) {

					strobj := Object()
					strobj.recordNr := A_Index
					For index, field in this.fields {
						txt := Trim(SubStr(set, field.start, field.len))
						If (StrLen(txt) > 0)
							strobj[(field.label)] := txt
					}
					matches.Push(strobj)

				  ; maximum array entries, store filepointer position for later use and return matches
					If (this.lendataset * A_Index = this.maxbytes) {
						this.position := this.dbf.position
						break
					}

				}
				else {
					exportfile.WriteLine(set)
				}


			}

			If this.nofileaccess {
				VarSetCapacity(recordbuf, 0)
				this.nofileaccess := false
				this.Close()
			}

			If IsObject(exportfile) {
				filelen := exportfile.Length
				exportfile.Close()
				return filelen
			}

		return matches
		}
	;}

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
	_ReadRecord() {


	}

	_SeekToRecord(nr, origin=2) {

		; 1 :	from first record
    	; 2 :	from current record

		If (origin = 1) {

			nr := (nr > this.records) ? this.records : (nr < 0) ? 0 : nr
			newposition	:= this.recordsStart + (nr * this.lendataset)
			If (this.pdf.Tell() = newposition)                                            	; file pointer is already in position
				return                                                                            	; return

			this.pdf.Seek(newposition, 0)

			return this.pdf.Tell()

		} else if (origin = 2) {

			cur_record := (this.pdf.Tell() - this.recordsStart) / this.lendataSet

			If       	(cur_record + nr > this.records)                              	; new file pointer position is after EOF
				this.pdf.Seek(this.pdf.Length, 0)                                      	; seek to end of file
			else if 	(cur_record + nr < 0)                                            	; new file pointer position is before first record
				this.pdf.Seek(this.recordsStart, 0)                                     	; seek to first record
			else                                                                                    	; else
				this.pdf.Seek(cur_record + (nr * this.lendataset), 1)            	; seek

			return this.pdf.Tell()

		}

	}

	; OnExit function to ensure that file access to database is closed on exit
	_Exit(ExitReason, ExitCode) {

		If IsObject(this.dbf)
			this.dbf.Close()

	}


	;}

}



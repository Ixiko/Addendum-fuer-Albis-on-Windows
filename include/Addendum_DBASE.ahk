; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;																				     	Addendum_DBASE - V1.6
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Dies ist ein Klassenbibliothek für den nativen Zugriff (Dateiebene) auf dBASE Datenbanken für Albis (Compugroup).
;		Geschrieben in Autohotkey_H. Möglichweise auch mit Autohotkey_L verwendbar (nicht getestet).
;		Von einer absoluten allgemeinen Verwendbarkeit mit dBASE Datenbanken anderer Hersteller sollte nicht ausgegangen werden.
; 		Dies ist anzunehmen, da Albis einige Bytes im Header der dBase-Dateien nicht nutzt und genutzte Bytes haben möglicherweise eine andere
; 		Bedeutung als üblich. Dennoch könnten die meisten Funktionen durchaus funktionieren, da diese Bibliothek nur wenige speziell auf Albis Datenbanken
; 		zugeschnittene Funktionen beinhaltet.
; 		Eine Erweiterung für diese Klasse mit spezialisierten Funktionen ist mit der Klasse 'AlbisDB' (Datei: include\Addendum_DB.ahk / class AlbisDB) vorhanden.
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;		Verwendung      		Analyse der von Albis verwendeten DBASE-Datenbank Dateien und deren Strukturen
;										Portierung von Daten
;										Suche nach Daten
;
; 		Beschreibung       	1.	ermöglicht ausschließlich lesende Zugriffe auf DBASE (.dbf) Dateien im '\albiswin\db\'-Ordner!
;									2.	rein nativer Zugriff (auf Dateiebene). Benötigt daher keine Datenbanktreiber (z.b. die Windows OLE-Treiber)
;									3.	mehrere Dateien können gleichzeitig bearbeitet werden
;									4.	enthält verschiedene Suchfunktionen.
;										-	es wird feldweise gesucht
;										-	möglich sind einfache Stringvergleiche oder RegEx-String basierende Mustersuchen
;											Achtung!:	Je mehr Suchbegriffe (insbesondere mit RegEx) verwendet werden, umso langsamer wird die Suche.
;										-	Daten werden als indizierter Array zurückgegeben.
;											Jedes Indexfeld enthält ein Objekt mit key:value Paaren. Die Schlüsselnamen (key's) entsprechen den Feldnamen der Datenbank.
;                                 	5.	Inhalte welche auf eine zweite Datenbank verteilt sind, werden und sollen in Zukunft automatisch zusammengeführt werden
;										Beispiel: BEFUND.dbf das Feld "INHALT" kann weitere Textdaten enthalten, diese finden sich dann BEFTEXT.dbf
;
;       Hinweise           		Ich übernehme keinerlei Haftung bei Schäden an Ihren Datenbankdateien durch/nach Verwendung dieser Bibliothek!
;										Sichern Sie die Daten vor Nutzung der Bibliothek!
;										Es werden noch keine Fehlermeldungen bei Aufruf nicht vorhandener Funktionen ausgegeben!
;										Auch wenn ich diese Klassenbibliothek regelmäßig verwende, gehen Sie dennoch davon aus das ich nicht
;										alle Fehler erzeugen und damit auch finden kann!
;
;       Abhängigkeiten   		\lib\[SciteOutPut, class_JSON]
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_DBASE begonnen am:          	12.11.2020
;       Addendum_DBASE letzte Änderung am: 	08.07.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

class dBASE {  ; native DBASE Klasse nur für Albis .dbf Files

	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌
	; ⚌                                                                     	CONSTRUCTION / DESTRUCTION                                                                     	⚌
	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌;{

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; stores data from Table File Header and gets the Field Descriptor Array
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		__New(filepath, debug=false, debugGui="") {

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; open dbase file for reading
			; ----------------------------------------------------------------------------------------------------------------------------------------;{
				SplitPath, filepath,,dBASEPath,, baseName

				this.headerrecordgap	:= 0                                            	; one byte after DBF Header, purpose?
				this.baseName           	:= baseName
				this.filepath                	:= filepath
				this.dBASEPath             	:= dBASEPath
				this.encoding            	:= "CP1252"                             	; main encoding is ASCII CP1252
				this.debug                    	:= debug
				this.debugGui           	:= debugGui
				this.nofileaccess         	:= false

			; ――――――――――――――――――――――――――――――――――――――――――――――――
			; max. bytes for working array - you can easily change it, without touching this class
			; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.maxCapacity        	:= 400 * 1024000 ; = 200 MB

			; ――――――――――――――――――――――――――――――――――――――――――――――――
			; database connections for some ALBIS dBASE files
			; ――――――――――――――――――――――――――――――――――――――――――――――――
			; 	DB: 			name of database
			; 	link: 	    	to retrieve the corresponding information from the other database (field name is not the same in the other one)
			;					first element is the original index nr and second element the corresponding field, if there's a third element
			;					the maximum index of this field will be stored
			;	append:	what value should be append with data retrieved from the other database
			; ――――――――――――――――――――――――――――――――――――――――――――――――
				Switch this.baseName {

					Case "BEFUND":
						this.headerrecordgap	:= 0
						this.connected           	:= {	"DBFilePath"	: (this.dBASEPath "\BEFTEXTE.dbf")
																 ,	"baseField"	: "TEXTDB"
																 ,	"linkField"  	: "LFDNR"}
								                				;~ , 	"FieldLinks"  	: ["TEXTDB", "LFDNR", "POS"]
											                	;~ , 	"append"    	: ["INHALT", "TEXT"]}
					Case "PATIENT":
						this.headerrecordgap	:= 0
				}

			; ――――――――――――――――――――――――――――――――――――――――――――――――
			; closes file access on script exit to prevent file damage
			; ――――――――――――――――――――――――――――――――――――――――――――――――
				OnExit(ObjBindMethod(this, "_Exit"))

			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; Table File Header (Length: 32 bytes)
			; ----------------------------------------------------------------------------------------------------------------------------------------;{
			  ; file object handling
				If !IsObject(this.dbf := FileOpen(filepath, "r", "CP1252")) {
					throw "open database file: " filepath " failed!"
					ExitApp
				}

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [    0   ]	 	contains the version number of this dBASE file
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.Version   	:= this._ReadNum(1, "Int")

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [  1-3  ]	 	Date of last update, in YYMMDD format. Each byte contains the number as a binary.
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; 						creates a HEX value string like: 0x14 | 0x0B | 0x1A = 20 | 11 | 26 (26.11.2020)
			  ;							Its stored as a string. Don't use that string without splitting each hex value as one number
			  ;							in this case 0x140B1A results to this integer value: 1313562. ‬
			  ;							I use this 3-HEX format solely because of its historical context.
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; 3-HEX value string creation
				year                      	:= this._ReadNum(1, "Int")
				month                   	:= this._ReadNum(1, "Int")
				day                       	:= this._ReadNum(1, "Int")
				this.lastupdate        	:= Format("{:02X}", year) Format("{:02X}", month) Format("{:02X}", day)
			  ; transforms DBASE date string in user readable format
				year	                    	:= 1900 + year
				month                   	:= SubStr("00" month, -1)
				day                       	:= SubStr("00" day, -1)
				this.lastupdateDate 	:= day "." month "." year
				this.lastupdateEng 	:= year "." month "." day
				this.lastupdatedBase 	:= year month day

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [  4-7  ]	 	Number of records in the table. (Least significant byte first.)
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.records   	:= this._ReadNum(4, "Int")

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [  8-9  ]	 	Number of bytes in the header. (Least significant byte first.)
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.headerLen 	:= this._ReadNum(2, "Short")

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [10-11]	 	Number of bytes in the header. (Least significant byte first.)
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.lendataset 	:= this._ReadNum(2, "Short")

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [12-27]	 	Skipping. I think these bytes are not used
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.dbf.Seek(16, 1)

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [   28  ]		Production MDX flag. 0x01 if a .mdx file exists for this table or 0x00 if not.
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.mdx        	:= this._ReadNum(1, "Int")

			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
			  ; Byte: [29-31]		I think these bytes are not used
			  ; ――――――――――――――――――――――――――――――――――――――――――――――――
				this.dbf.Seek(3, 1)

			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; Field Properties Structure (each field: 32 bytes, field count: ( DBF Header Length - TFH Length ) / 32 [bytes]
			; ----------------------------------------------------------------------------------------------------------------------------------------;{

				; ――――――――――――――――――――――――――――――――――――――――――――――――
				; DBASE structure will be safed in some objects
				; ――――――――――――――――――――――――――――――――――――――――――――――― ;{
					this.fields      	:= Array()                                            	; indexed array with every field name
					this.dbfields   	:= Object()                                        	; field name based object with additional data
					this.dbstruct   	:= Object()                                        	; contains basic structure informations from header
					this.NrFields  	:= Round((this.headerLen - 32) / 32)    	; count of fields
				;}

				; ――――――――――――――――――――――――――――――――――――――――――――――――
				; some necessary variables
				; ――――――――――――――――――――――――――――――――――――――――――――――― ;{
					fpos  	:= 0
					lenNF 	:= StrLen(this.NrFields) - 1
					dbfields	:= ""
				;}

				; ――――――――――――――――――――――――――――――――――――――――――――――――
				; collecting informations about every field
				; ――――――――――――――――――――――――――――――――――――――――――――――― ;{
					Loop % this.NrFields {

					  ; ――――――――――――――――――――――――――――――――――――――――――――
					  ; read raw bytes of field properties
					  ; ――――――――――――――――――――――――――――――――――――――――――――
						flabel	:= this._ReadString(11)                              	; 11 bytes	:	Field name in ASCII (zero-filled).
						ftype 	:= this._ReadString(1)                                	;   1 byte	:	Field type in ASCII (B, C, D, N, L, M, @, I, +, F, 0 or G).
										 this.dbf.Seek(4, 1)                                   	;   4 bytes	:	skip unused bytes
						flen   	:= this._ReadNum(1, "UChar")     	            	;	1 byte	:	Field length in binary.
										 this.dbf.Seek(14, 1)                                 	; 14 bytes	:	skip unused bytes
						fmdx 	:= this._ReadNum(1, "UChar")                    	;	1 byte	:	maybe this field is indexed in .mdx file

					  ; ――――――――――――――――――――――――――――――――――――――――――――
					  ; store informations in indexed array
					  ; ――――――――――――――――――――――――――――――――――――――――――――
						this.fields.Push({"label": flabel, "type":ftype, "start":(fpos+1), "len":flen, "more":fmdx})

					  ; ――――――――――――――――――――――――――――――――――――――――――――
					  ; store	informations based on field name
					  ; ――――――――――――――――――――――――――――――――――――――――――――
						this.dbfields[fLabel] := {"type":ftype, "start":(fpos+1), "len":flen, "more":fmdx, "pos":A_Index}

					  ; ――――――――――――――――――――――――――――――――――――――――――――
					  ; stores field names as comma separated string
					  ; ――――――――――――――――――――――――――――――――――――――――――――
						dbfields .= flabel ","

					  ; ――――――――――――――――――――――――――――――――――――――――――――
					  ; for debugging
					  ; ――――――――――――――――――――――――――――――――――――――――――――
						 If (this.debug = 4)
							x .= " > Pos: " SubStr("000" fpos, -2) "-" SubStr("000" fpos+flen-1, -2) " (" SubStr("00" flen, -1) ") [" ftype "]  - " flabel "`n"

					  ; ――――――――――――――――――――――――――――――――――――――――――――
					  ; start position of next field in record set
					  ; ――――――――――――――――――――――――――――――――――――――――――――
						fpos += flen

					}

				;}

				; ――――――――――――――――――――――――――――――――――――――――――――――――
				; stores basic informations about this DBASE database
				; ――――――――――――――――――――――――――――――――――――――――――――――― ;{
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

				; ――――――――――――――――――――――――――――――――――――――――――――――――
				; close this DBASE file after getting all data
				; ――――――――――――――――――――――――――――――――――――――――――――――――
					this.CloseDBF()

			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
	    	; something to debug here
			; ----------------------------------------------------------------------------------------------------------------------------------------;{
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
			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; calculations for viewing of progress
			; ----------------------------------------------------------------------------------------------------------------------------------------;{
				this.ShowAt	:= Round(this.records/50)
				this.slen     	:= -1*(StrLen(this.records) - 1)
				this.maxRecs	:= SubStr("00000000" this.records, this.slen)
			;}

			; ------------------------------------------------------------------------------------------------------------------------------------------
			; calculations of array size and filepointer position of first record in this database
			; ----------------------------------------------------------------------------------------------------------------------------------------;{
				global recordbuf

				this.maxBytes	:= this.lendataset * this.records > this.maxCapacity ? this.maxCapacity : this.lendataset * this.records
				this.recordsStart	:= this.headerlen + this.headerrecordgap

				VarSetCapacity(this.recordbuf, this.lendataset, 0x20) 	; recordbuf
				VarSetCapacity(recordbuf, this.lendataset, 0x20) 	; recordbuf
			;}

		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; file access will be closed, if this object is destroyed
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		__Delete() {

			SciTEOutput("Object: " this.baseName " was deleted.")
			If IsObject(this.dbf)
				this.CloseDBF()

		}

	;}

	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌
	; ⚌                                                                	  DATABASE FILE ACCESS FUNCTIONS                                                                    	⚌
	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌;{

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

			this.filepos := this.dbf.Tell()
			this.dbf.Close()
			this.dbf := ""

		return this.filepos
		}

	;}

	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌
	; ⚌                                                           	RETRIEVE DATA FROM DATABASE RECORDS                                                                	⚌
	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌;{

	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; REGEX BASED SEARCH. 						| RETURNS THE WHOLE RECORDSET (ALL FIELDS AND VALUES)!
	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
		Search(pattern, startrecord=0, callbackFunc="", opt="") {

		/* main description

				- parameter: pattern is designed to use only RegEx for matching, use function SearchFields() for different search algorithms!

									opt 		:  	ReturnDeleted=true, if you want to make the function return also the records with delete flag.
													Default is false. No deleted records will be returned.

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

		; Initializing vars, examining options ;{
			static recordbuf, rxStr
			VarSetCapacity(recordbuf	, this.lendataset, 0x20)
			VarSetCapacity(matches	, this.maxBytes)
			matches             	:= Object()                                 	; collects the findings
			this.flagged       	:= 0
			this.hits              	:= 0
			this.uselastRxStr 	:= false
			this.callbackFunc 	:= Func(callbackFunc)
			this.breakrec 		:= "#"                                       	; this variable is used to help external processing of data

		  ; checks opt parameters (opt is made for debugging)
			If IsObject(opt) {

				If opt.SaveMatchingSets {
					path := opt.MSetsPath
					SplitPath, path, MSetfilename, MSetDir
					If !InStr(MSetDir, "D")
						throw A_ThisFunc ": " MSetDir "\ is not a directory."
					else
						MSets := FileOpen(opt.MSetsPath, "w", "UTF-8")
				}

				If !opt.haskey("ReturnDeleted")
					opt.ReturnDeleted := false

			}

		;}

		; Establish read access if not already done
			If !IsObject(this.dbf) {
				this.nofileaccess := true
				this.filepos := this.OpenDBF()
			}

		; builds a regex string, regex string can be re used by command
			If IsObject(pattern) && !this.uselastRxStr {

			; build regexstring routine
				this.hits	:= 0                                            	; hits counter
				rxStr := "i)^\s*", rxPre := 0
				For fieldIndex, field in this.fields
					If pattern.haskey(field.label) {

						if (rxPre > 0)
							rxStr .= ".{" rxPre "}", rxPre := 0

						If !RegExMatch(pattern[field.label], "i)^\s*rx\:") {

							If (field.type = "D")
								rxStr .= "\s*{" (field.len - StrLen(pattern[field.label])) "}" pattern[field.label]
							else {
								rxExpand := (field.len - StrLen(pattern[field.label]))
								rxStr .= pattern[field.label] ;".{" (field.len - StrLen(pattern[field.label])) "}*"
							}

						}
						else {

							thispattern := RegExReplace(pattern[field.label], "^\s*rx\s*\:")
							If (field.type = "D")
								rxStr .= thispattern
							else
								rxStr .= "\s*" thispattern "\s*"

						}

						lastrxLen := StrLen(rxStr)

					}
					else
						rxPre += field.len

			; needed replacements
				rxStr := SubStr(rxStr,1, lastrxLen)
				rxStr := RegExReplace(rxStr, "(\.\*){2,}", ".*")
				rxStr := RegExReplace(rxStr, "\\s\*\\s\*", "\s*")
				rxStr := RegExReplace(rxStr, "\.\*\\s\*\.\*", ".*")
				rxStr := RegExReplace(rxStr, "\\s\*\.\*", ".*")
				rxStr := RegExReplace(rxStr, "\.\*$")

			; return on failure
				If (StrLen(rxStr) = 0)
					return "can't build RegEx search string"

			; publish some vars
				this.lastrxLen         	:= lastrxLen
				this.SearchRegExStr	:= rxStr

			} else if !IsObject(pattern) {
				this.SearchRegExStr := rxStr := pattern
			}

		; for debugging
			If RegExMatch(this.debug, "i)\bRxStr\b")
				SciTEOutput(A_ThisFunc " - RegExStr: " rxStr)

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

			; saving whole datasets for debugging or other purposes
				If opt.SaveMatchingSets
					MSets.Write(set "`n")

			; temp object collects dataset field data
				strobj    	:= Object()
				flagged  	:= SubStr(set, 1, 1)
				strobj.removed	:= SubStr(set, 1, 1) = "*" ? true : false
				set                 	:= Substr(set, 2, StrLen(set)-1)
				strobj.recordNr	:= Floor((this.dbf.Tell() - this.headerlen) / this.lendataset)

				;~ If strobj.removed
				If Trim(flagged)
					this.flagged ++

			; continue with next record, if a delete flag was found, unless this was switched off
				If strobj.removed && !opt.ReturnDeleted
					continue

			; assemble the data to a Key:Value object
				For index, field in this.fields
					If (StrLen(value := Trim(SubStr(set, field.start, field.len))) = 0)
						continue
					else
						strobj[field.label] := value

			; append match object
				matches.Push(strobj)
				this.hits ++

		  ; maximum array entries, store filepointer position for later use and return matches
		    	;~ If (this.lendataset * A_Index = this.maxbytes) {
					;~ this.filepos	:= this.dbf.Tell()                          	; last position of filepointer
					;~ this.breakrec	:= startrecord + A_Index          	; position of last read dataset
					;~ ToolTip
					;~ return matches
				;~ }

			}

			;SciTEOutput("sets: " this.hits "`nwith remove flag: " this.flagged)

		; file access is automatically terminated if the function established this independently
			If this.nofileaccess {
				VarSetCapacity(recordbuf, 0)
				this.nofileaccess := false
				this.CloseDBF()
			}

		; close MSets fileaccess
			MSets.Close()

		; indicates that the end of the file has been reached
			this.breakrec	:= "!"

		; end debug and last call the callback function
			ToolTip
			If (StrLen(callbackFunc) > 0)
				%callbackFunc%(setNR, this.records, this.slen, matches.MaxIndex())


		return matches
		}

	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; FIELD DIVIDED SEARCH. 						| EVERY FIELD CAN BE SEARCHED INDIVIDUALLY. !!MAYBE VERY SLOW, BUT PRECISE!!
	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
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

			; file access is automatically terminated if the function established this independently
				If this.nofileaccess {
					VarSetCapacity(recordbuf, 0)
					this.nofileaccess := false
					this.pos := this.CloseDBF()
				}

		return matches
		}

	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; SIMPLE STRING COMPARE. 					| FOR FAST RESULTS
	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
		SearchFast(pattern, outfields, startrecord=0, options="", callbackFunc="") {

			; comparison is possible with ranges or Instr() operations. One date can be compared with another
			; by only showing the dates that are greater, less or equal. This also works with pure numerical values.

				static matches, recordbuf, recordpos

				If !IsObject(pattern)
					throw "Parameter: pattern must be an object"

			; processed only on first start or other option than "next"
				If !RegExMatch(options, "i)\bnext\b") {

					; initialize some vars
						VarSetCapacity(matches 	, this.maxBytes)

					; establish read access if not already done
						If !IsObject(this.dbf) {
							this.nofileaccess := true
							this.filepos := this.OpenDBF()
						}

					; prepare dataset operation. Uses every field in the comparison process or
					; only those that have been supplied with the pattern parameter
						If !IsObject(outfields) && RegExMatch(outfields, "all fields")
							outfields := ""
						this.retSubs := this.BuildSubstringData(outfields)
				}

			; return data sets with delete flag, default is false
				ReturnDeleted := RegExMatch(options, "i)return_Deleted_Records") ? true : false

			; filepointer is set to the data record with the number from startrecord
				this.dbf.Seek(this.recordsStart, 0)
				If (startrecord > 0) {
					this.dbf.Seek(this.lendataset * startrecord, 1)
					this.filepos_start := this.dbf.Tell()
				}
				this.recordpos := startrecord
				VarSetCapacity(recordbuf	, this.lendataset*100, 0x20)

			; search loop
				this.foundrecord := 0, this.hits := 0, matches := Array()
				while (!this.dbf.AtEOF) {

					; debug some actions
						this.recordpos := StartRecord + A_Index
						If (Mod(setNR := A_Index, this.ShowAt) = 0) {
							If (this.debug = 1)
								ToolTip, % "Search rx: " rxStr "`n" SubStr("00000000" A_Index, this.slen) "/" SubStr("00000000" this.records, this.slen)
							If (StrLen(callbackFunc) > 0)
								%callbackFunc%(setNR, this.records, this.slen, matches.MaxIndex())
						}

					; reads one dataset from database
						bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
						set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)
						removed	:= SubStr(set, 1, 1) = "*" ? true : false
						set          	:= Substr(set, 2, StrLen(set)-1)

					; continue if returnDeleted is false
						If !ReturnDeleted && removed
							continue

					; compare pattern field condition with dataset
						hits := 0
						For fLabel, searchString in pattern {
							val := Trim(SubStr(set, this.dbfields[fLabel].start, this.dbfields[fLabel].len))
							If (val = searchString) || RegExMatch(val, searchString){
								hits ++
								this.filepos_found 	:= this.dbf.Tell()
								this.foundrecord 	:= this.recordpos
							}
						}

						If (hits = pattern.Count()) {
							matches.Push(this.GetSubData(set, this.retSubs))
							matches[matches.MaxIndex()].removed := removed
						}

				}


		return matches
		}

	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; RETURN DATA FROM EVERY RECORD. 	| RETURNS CERTAIN FIELD NAMES AND ITS VALUES
	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
		GetFields(getfields="", callbackFunc="") {

			; the function returns the data of all records as key:val object
			; you have to specify getfields as an array containing fieldnames

			; prepares some variables
				VarSetCapacity(recordbuf, this.lendataset, 0x20)
				obj	:= Object()

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

					; progress view
						If (this.debug > 0) && (Mod(A_Index, this.ShowAt) = 0) {
							If (this.debug > 0)
								ToolTip, % "GetFields:`t" SubStr("00000000" A_Index, this.slen) "/" this.maxRecs "`ndatasets: " data.Count()
							If IsFunc(callbackFunc)
								%callbackFunc%(A_Index, this.records, this.len, matches.Count())
						}

					; reads a data set raw mode and converts it to a readable text format
						bytes     	:= this.dbf.RawRead(recordbuf, this.lendataset)
						readset 	:= StrGet(&recordbuf, this.lendataset, this.encoding)
						set     		:= Substr(readset, 2, StrLen(readset)-1)

					; push object to data array (object contains the requested field names, values and its specific recordnr)
						obj.Push(this.GetSubData(set, subs))

				}

			; file access is automatically terminated if the function established this independently
				If this.nofileaccess {
					VarSetCapacity(recordbuf, 0)
					this.nofileaccess := false
					this.pos := this.CloseDBF()
				}

				If (this.debug > 0)
					ToolTip

		return obj
		}

	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; READS ONE RECORD SET                     | ONE BY ONE, STARTS FROM BEGINNING
	  ; -------------------------------------------------------------------------------------------------------------------------------------------------------------------
		ReadRecord(outfields="", callbackFunc="") {

			; you can set seekpos by

				global recordbuf

			; establish read access if not already done
				If !IsObject(this.dbf) {
					this.nofileaccess := true
					this.pos := this.OpenDBF()
				}

			; display progress
				if (this.debug && Mod(this.recordnr, this.ShowAt) = 0) {
					If RegExMatch(this.debug, "i)\bTT\b)")
						ToolTip, % SubStr("00000000" this.recordnr, this.slen) "/" SubStr("00000000" this.records, this.slen)
					If IsFunc(callbackFunc)  ;> 0)
						%callbackFunc%(this.recordnr, this.records, this.slen, "---")
				}

			; reads one dataset
				bytes		:= this.dbf.RawRead(recordbuf, this.lendataset)
				readset	:= StrGet(&recordbuf, this.lendataset, this.encoding)
				set 		:= Substr(readset, 2, StrLen(set)-1)

			; change class object filepointer position
				this.recordnr	:= Floor((this.dbf.Tell() - this.headerlen) / this.lendataset)
				this.filepos 	:= this.dbf.Tell()

			; return unparsed dataset
				If !IsObject(outfields)
					return set

			; build key:value object
				obj              	:= Object()
				obj.removed	:= SubStr(readset, 1, 1) = "*" ? true : false
				obj.recordNr	:= this.recordnr
				For index, fLabel in outfields
					obj[flabel] := Trim(SubStr(set, this.dbfields[fLabel].start, this.dbfields[fLabel].len))

		return obj
		}

	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	  ; READS A BLOCK OF RECORDSETS		| RETURNS THE WHOLE RECORDSET
	  ; ----------------------------------------------------------------------------------------------------------------------------------------------------------------
		ReadBlock(NrOfSets, getfields:="", BlockNr=0, callbackFunc="") {

			; function returns all data within a block* as key:val object
			; getfields must be an array containing the names of the fields you want to retrieve
			;	* 	A block is a consecutive number of records specified by the 'NrOfSets' parameter.
			; 		'BlockNr' parameter specifies which block within the database is to be read.

				static recordbuf

			; prepares some variables
				If !IsObject(this.data) || IsObject(getfields)  {

						VarSetCapacity(recordbuf, this.lendataset, 0x20)
						this.getfields	:= getfields

					; creates an array with data for substring function retreave only fields that should be returned
						this.subs      	:= this.BuildSubstringData(getfields)
						If (this.subs.MaxIndex() = 0)                               	; returns here if all field names passed are unknown
							return 2                                                        	; "field label(s) unknown"

				}
				this.matches	:= Object()

			; establish read access if not already done
				If !IsObject(this.dbf) {
					this.nofileaccess := true
					this.pos := this.OpenDBF()
				}

			; filepointer is set to the data record with the number from startrecord
				this.StartRecord 	:= (BlockNr*NrOfSets)
				this.Seekpos      	:= (this.lendataset * this.Startrecord)
				this.dbf.Seek(this.seekpos, 1)
				this.filepos := this.dbf.Tell()

			; compile field data as an object
				blockpos := 0
				while (!this.dbf.AtEOF)  {

					; reads a data set raw mode and converts it to a readable text format
						bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
						set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)

						this.filepos := this.dbf.Tell()
						removed	:= InStr(SubStr(set, 1, 1), "*") ? true : false
						set        	:= Substr(set, 2, StrLen(set)-1)

					; debugging and callback function to show progress
					If (Mod(setNR := A_Index, this.ShowAt) = 0) {
						If (this.debug = 1)
							ToolTip, % "Search rx: " rxStr "`n" SubStr("00000000" A_Index, this.slen) "/" SubStr("00000000" this.records, this.slen)
						If (StrLen(callbackFunc) > 0)
							%callbackFunc%(blockpos, NrOfSets, this.slen, blockpos)
					}

					; push object to data array (object contains the requested field names, values and its specific recordnr)
						fields := this.GetSubData(set, this.subs)
						fields.removed := removed
						this.matches.Push(fields)
						blockpos ++
						If (blockpos = NrOfSets)
							break
				}

			; file access is automatically terminated if the function established this independently
				If this.nofileaccess {
					VarSetCapacity(recordbuf, 0)
					this.nofileaccess := false
					this.CloseDBF()
				}

		return this.matches
		}

	;}

	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌
	; ⚌                                                       	CREATE OWN INDEX FILE BASED ON PARAMETERS                                                        	⚌
	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌;{
		CreateIndex(IndexFilePath, IndexFieldName:="", IndexMode:="", ReIndex:=false) {

			; creates or updates an index file with fileseek positions at the first occurrence of a certain string (date, string, quarter)
			; the function currently only indexes by quarters
			; if an index file exists, it only updates the index from the last entry
			; IndexFieldName:		The database field name which is indexed. At the moment only fields with date strings can be processed.

				;IndexMode := "quarter"

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
					;SciTEOutput("QZ: " QZ ", seek: " seekpos, "StartRecord: " StartRecord)

				}

			; Establish read access if not already done
				If !IsObject(this.dbf) {
					this.nofileaccess := true
					this.filepos := this.OpenDBF()
				}

			; indexing is only possible for one condition
				FStart	:= this.dbfields[IndexFieldName].start
				FLen 	:= this.dbfields[IndexFieldName].len
				KeyN	:= "QNow"
				If (FStart = 0 || FLen = 0)
					throw A_ThisFunc ": Database [" this.dbname "]`nhas no field named <" IndexFieldName ">"

			; start indexing
				maxRecs := SubStr("00000000" this.records, this.slen)
				VarSetCapacity(recordbuf, this.lendataset, 0x20)
				while (!this.dbf.AtEOF) {

					; reads a data set raw mode and converts it to a readable text format
						bytes	:= this.dbf.RawRead(recordbuf, this.lendataset)
						set 	:= StrGet(&recordbuf, this.lendataset, this.encoding)
						val   	:= SubStr(set, FStart, FLen)

					; mode driven indexing, calculates dates to quarter
						If (IndexMode = "DateToQuarter") {
							IdxKey	:= "#" SubStr(val, 1, 4) . Ceil(Substr(val, 5, 2)/3)
						} else If (IndexMode = "QYYtoYYQ" && IndexFieldName = "QUARTAL"){
							IdxKey	:= "#" SubStr(val, 2, 2) SubStr(val, 1, 1)
						} else {
							idxKEy	:= val
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

						If (IdxKey = 0 || StrLen(IdxKey) = 0 || IdxKey = idxKEy_old)
							continue
						idxKEy_old := IdxKey

						If !iFile.haskey(IdxKey)
							iFile[IdxKey] := [this._GetRecordFromFilepos(this.dbf.Tell())-1 , this.dbf.Tell()] 	   ; record nr, filepointer position

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

	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌
	; ⚌                                                                              	 STRING FUNCTIONS                                                                                	⚌
	; ⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌⚌;{

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
			strobj["removed"]	:= InStr(SubStr(set, -1), "*") || InStr(SubStr(set, 1, 1), "*" ) ? true : false         	; Dataset ist removed?
			strobj["recordnr"]	:= Floor((this.dbf.Tell() - this.recordsStart) / this.lendataset)                       	; saves the number of record
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

			If (this.dbf.Tell() <> newposition)                                         	; file pointer is not in position
				this.dbf.Seek(newposition, 0)

		} else if (origin = 2) {

			cur_record := (this.dbf.Tell() - this.recordsStart) / this.lendataSet

			If (cur_record + nr > this.records)                                       	; new file pointer position is after EOF
				this.dbf.Seek(this.dbf.Length, 0)                                      	; seek to end of file
			else if (cur_record + nr < 0)                                              	; new file pointer position is before first record
				this.dbf.Seek(this.recordsStart, 0)                                     	; seek to first record
			else                                                                                    	; else
				this.dbf.Seek(cur_record + (nr * this.lendataset), 1)            	; seek

		}

	return this.dbf.Tell()
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




AppendMatchesWithIndexing(matches, properties) {	; ##untested## appends database values, if connection to another db is known

	; there are two features inside:
	; first: 		this function will automatically create an individual index based on given parameters
	;				if there's no index created at time, the function will not waste time create an index for first. It reads and parse data at same time it's indexing.
	; second: 	it loads the data you want to have
	; dependencies: class_DBASE and class_JSON.ahk

		static dbIndex, lastfilepath

	; extract data from properties                               	;{
		filepath     	:= properties.DBFilePath
		fieldFrom  	:= properties.link[1]              	; TEXTDB 	in BEFUND means
		fieldTo      	:= properties.link[2]                	; LFDNR 	in BEFTEXT
		fieldIndex  	:= properties.link[3]	            	; POS    	in BEFTEXT means the sequence of parts of the text to append (a second index)
		appendTo  	:= properties.append[1]          	; value of INHALT (BEFUND.dbf) must append with
		appendFrom	:= properties.append[2]			; values from TEXT (BEFTEXT.dbf)
	;}

	; create outfields parameter                                 	;{
		outfields.Push(fromfield)
		If fieldIndex
			outfields.Push(fieldIndex)
	;}

	; dbIndex - prepare a new or load one                 	;{
		If !IsObject(dbIndex) || (lastfilepath <> filepath) {

			lastfilepath := filepath
			SplitPath, lastfilepath,, dir,, dbname
			dbIndexFilePath := Addendum.DBPath "\DBASE\IndexOf_" dbname ".json"

			If FileExist(dbIndexFilePath)
				dbIndex := JSONData.Load(dbIndexFilePath, "", "UTF-8")
			else
				dbIndex := Array()

		}
	;}

	; get MaxIndex from ObjectToAppend                	;{
		links := Object()
		MaxIndexToGet := 0
		For origIndex, data in matches {
			links[(dIndex := data[fieldFrom])] := origIndex
			MaxIndexToGet := MaxIndexToGet < dIndex ? dIndex : MaxIndexToGet
		}
		;SciTEOutput(MaxIndexToGet " = " links.MaxIndex() " ?" )
	;}

	; collect data to append                                     	;{
		connected 	:= new DBASE(filepath)
		res             	:= connected.dbf.OpenDBF()

		For dIndex, origIndex in links {

			; use seek position from DBIndex ;{
				StartRecord := 0
				If dbIndex.haskey(dIndex)
					StartRecord := dbIndex[dIndex].first - 1
				else {
					Loop % dIndex
						If dbIndex.haskey(dIndex-A_Index) {
							StartRecord := dbIndex[(dIndex-A_Index)].first - 1
							break
						}
				}
				;SciTEOutput(dIndex ": " StartRecord)

			; seek to this position
				connected.dbf._SeekToRecord(StartRecord, 1)

			;}

			; read recordset until data is found and allways collect data for Index
				append	:= fieldIndex ? Array() : ""
				IndexFound := false
				while (!connected.dbf.AtEOF) {

					; linked data found
						rec := connected.dbf.ReadRecord(outfields)
						if (dIndex = rec[fieldTo]) {

							IndexFound := true

						  ; collect the data to append later
							If fieldIndex
								append[(rec[fieldIndex]+1)] := rec[appendFrom]
							else
								append .= rec[appendFrom]

						  ; DBIndex-Data will be collected
							dbIndex[rec[fieldTo]] := {"first":rec["recordNr"]}
							lastIndex := rec[fieldTo]

						}
						else if (lastIndex <> rec[fieldTo]) {

							dbIndex[rec[fieldTo]] := {"last":rec["recordNr"]-1}
							lastIndex := rec[fieldTo]
							If IndexFound
								break

						}

				}

			; append found data to original matches
				If fieldIndex {
					For fidx, txt in append
						matches[origIndex][appendTo] .= txt
				}
				else
					matches[origIndex][appendTo] .= append

			}

		res            	:= connected.dbf.CloseDBF()
		connected.dbf := ""

	;}

	; save index for faster access next time
		JSONData.Save(dbIndexFilePath , dbIndex, true,, 1, "UTF-8")

return matches
}





; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;               															⚗       Labordaten		💉
;
;      Funktionen:          	▫ Suche in ldt Dateien (Labordaten)
;                               	▫ Extrahieren von einzelnen Anforderungen
;                               	▫ ignorierte Daten aufspüren
;                               	▫ (Konvertierung in menschenlesenbaren Text)
;
;      Basisskript:              noch keines
;
;		Abhängigkeiten:
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - letzte Änderung 15.03.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


/* LDT Datensätze - Beschreibung

	.ldt ⇦⇨ Deutsch

	Encodiert sind .LDT Dateien in ISO 8859-15 (https://de.wikipedia.org/wiki/ISO_8859-15).
	In Notepad++ sind die Umlaute nicht lesbar bei einer Konvertierung mit ISO 8859-15. Nimm OEM 850!
	Für das Lese-Encoding mit Autohotkey nimm: CP850 , schreiben nur mit Originalkodierung CP28605

	┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄
	                                          Zeichencodetabelle ISO 8859-15
	┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄
	Dec 	⬪ 		0 		16 	32 	48 	64 	80 	96 	112 	128 144 	160 176	192 	208 	224	240
	⬪ 		Hex 	0 		1 		2 		3 		4 		5 		6 		 7 	8 		9 		A 		B 		C 		D 		E 		F
	0 		0 						SP 	0 		@ 	P 		` 		p 								° 		À 		Ð 		à 		ð
	1 		1						! 		1		A 		Q 	a 		q 						¡ 		±		Á		Ñ 	á		ñ
	2		2						„ 		2		B 		R 		b 		r 						¢		 ² 		Â 		Ò 	â 		ò
	3		3						#		3		C 		S 		c 		s 						£		 ³ 		Ã 		Ó 	ã 		ó
	4		4						$		4		D 		T 		d		t 						€		 Ž 	Ä 		Ô 	ä 		ô
	5		5						%		5		E 		U 		e 		u 						¥		 µ 	Å 		Õ 	å 		õ
	6		6						&		6		F 		V 		f 		v 						Š		¶ 		Æ		Ö 	æ 	ö
	7		7						‚		7		G 	W 	g 		w 						§		· 		Ç 		× 	ç 		÷
	8		8						(		8		H 		X 		h 		x 						š 		ž		È 		Ø 	è 		ø
	9		9						)		9		I 		Y 		i 		y 						© 	¹ 		É 		Ù 		é 		ù
	10	A		LF 			* 		: 		J 		Z 		j 		z 						ª 		º 		Ê 		Ú 		ê 		ú
	11	B						+		; 		K		[ 		k 		{ 						« 		» 		Ë 		Û 		ë 		û
	12	C						,		< 	L 		\ 		l 		| 						¬ 	Œ 	Ì 		Ü 		ì 		ü
	13	D		CR 			-		= 	M 	] 		m 	} 						SHY	œ		Í		Ý		í		ý
	14	E						.		> 	N 	^ 	n 		~ 	    			®		Ÿ		Î		Þ		î		þ
	15	F						/		? 		O 	_ 		o 		DEL 					¯		¿		Ï		ß		ï		ÿ
	┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄ ┄

	Jeder Patienten-Datensatz beginnt mit der Feldkennung 8000, gefüllt mit dem Wert für die
	entsprechende Satzart und beinhaltet mindestens ein weiteres Objektattribut bzw. Feld.


	Feldkennungen für LDT-Version: LDT1014.01

	001 Abrechnungsgebühr
	0101 KBV-Prüfnummer
	0201 (N)BSNR
	0203 (N)BSNR-Bezeichnung
	0205 Straße der (N)BSNR
	0211 Arztname
	0212 Lebenslange Arztnummer  (LANR)
	0215 PLZ der (N)BSNR-Adresse
	0216 Ort der (N)BSNR-Adress

	3100 Namenszusatz
	3101 Nachname
	3102 Vorname
	3103 Geburtstdatum
	3104 Titel
	3105 Versichertennummer
	3107 Straße
	3108 Versichertenart
	3109 Hausnummer
	3110 Geschlecht
	3112 Postleitzahl (PLZ)
	3113 Ort
	3116 WOP
	3119 Versicherten_ID

	3622 Größe des Patienten
	3623 Gewicht des Patienten

	4132 DMP_Kennzeichnung

	4205 Auftrag
	4207 Diagnose/Verdachtsdiagnose
	4208 Befund/Medikation

	5001 Abrechnungsziffer
	5002 Art der Untersuchung
	5005 Multiplikator

	8000 Satzart (Beginn eines Datensatz?)
	8001 Sartzende
	8002 Objektident
	8003 Objektende
	8203 ??

	8100 Satzlänge
	8100 Objektattribute


	82
	8220 Satzart: L (Labor)-Datenpaket-Header
	8201


	83
	8300 Labor
	8301 Eingangsdatum
	8302 Berichtsdatum
	8303 Berichtszeit
	8310 Auftragsnummer des Einsenders
	8311 Auftragsnummer des Labors
	8320 Laborname

	84
	8401 Status (Befund/Bericht)
	8403 Gebührenordnung [1 = BMÄ  2 = EGO  3 = GOÄ 96  4 = BG-Tarif  5 = GOÄ 88]
	8406 Kosten in (€) Cent

	8410 Kurzbezeichnung Parameter (Test-Ident)
	8411 Langbezeichnung Parameter (Testbezeichung)

	8420 Ergebniswert
	8421 Einheit
	8422 Grenzwert-Indikator [+ = leicht erhöht, ++ = stark erhöht, - = mäßig erniedrigt, -- = stark erniedrigt, ! = auffällig]

	8432 Abnahmedatum
	8433 Abnahmezeit
	8434 Anforderungen

	8460 Normalwert-Text
	8461 Normalwert-Untergrenze
	8462 Normalwert-Obergrenze
	8470 Laborhinweis


	86
	8609 Abrechnungstyp
	8614 Abrechnung durch
	8615 Arzt-Nr. anfordernder Arzt


	9212 LDT Version


	([\p{L}\-\/]+)\s+([\d\.\-]+)\s+([\p{L}\.\sæ\/%]+)\s+([\+\d\.\-]+)   | $1 | $2 | $3 | $4 |
 */

/*

https://docs.microsoft.com/de-de/dotnet/api/system.text.encodinginfo.getencoding?view=net-6.0
The example produces the following output when run on .NET Core:


Info.CodePage      Info.Name                    Info.DisplayName
1200               utf-16                       Unicode
1201               utf-16BE                     Unicode (Big-Endian)
12000              utf-32                       Unicode (UTF-32)
12001              utf-32BE                     Unicode (UTF-32 Big-Endian)
20127              us-ascii                     US-ASCII
28591              iso-8859-1                   Western European (ISO)
65000              utf-7                        Unicode (UTF-7)
65001              utf-8                        Unicode (UTF-8)

The example produces the following output when run on .NET Framework:

Info.CodePage      Info.Name                    Info.DisplayName
		37                 IBM037                       IBM EBCDIC (US-Canada)
		437                IBM437                       OEM United States
		500                IBM500                       IBM EBCDIC (International)
		708                ASMO-708                     Arabic (ASMO 708)
		720                DOS-720                      Arabic (DOS)
		737                ibm737                       Greek (DOS)
		775                ibm775                       Baltic (DOS)
		850                ibm850                       Western European (DOS)
		852                ibm852                       Central European (DOS)
		855                IBM855                       OEM Cyrillic
		857                ibm857                       Turkish (DOS)
		858                IBM00858                     OEM Multilingual Latin I
		860                IBM860                       Portuguese (DOS)
		861                ibm861                       Icelandic (DOS)
		862                DOS-862                      Hebrew (DOS)
		863                IBM863                       French Canadian (DOS)
		864                IBM864                       Arabic (864)
		865                IBM865                       Nordic (DOS)
		866                cp866                        		Cyrillic (DOS)
		869                ibm869                       	Greek, Modern (DOS)
		870                IBM870                       	IBM EBCDIC (Multilingual Latin-2)
		874                windows-874                  Thai (Windows)
		875                cp875                        		IBM EBCDIC (Greek Modern)
		932                shift_jis                    		Japanese (Shift-JIS)
		936                gb2312                       	Chinese Simplified (GB2312)
		949                ks_c_5601-1987           	Korean
		950                big5                         		Chinese Traditional (Big5)
		1026               IBM1026                      	IBM EBCDIC (Turkish Latin-5)
		1047               IBM01047                     IBM Latin-1
		1140               IBM01140                     IBM EBCDIC (US-Canada-Euro)
		1141               IBM01141                     IBM EBCDIC (Germany-Euro)
		1142               IBM01142                     IBM EBCDIC (Denmark-Norway-Euro)
		1143               IBM01143                     IBM EBCDIC (Finland-Sweden-Euro)
		1144               IBM01144                     IBM EBCDIC (Italy-Euro)
		1145               IBM01145                     IBM EBCDIC (Spain-Euro)
		1146               IBM01146                     IBM EBCDIC (UK-Euro)
		1147               IBM01147                     IBM EBCDIC (France-Euro)
		1148               IBM01148                     IBM EBCDIC (International-Euro)
		1149               IBM01149                     IBM EBCDIC (Icelandic-Euro)
		1200               utf-16                       		Unicode
		1201               utf-16BE                     	Unicode (Big-Endian)
		1250               windows-1250                 Central European (Windows)
		1251               windows-1251                 Cyrillic (Windows)
		1252               windows-1252                 Western European (Windows)
		1253               windows-1253                 Greek (Windows)
		1254               windows-1254                 Turkish (Windows)
		1255               windows-1255                 Hebrew (Windows)
		1256               windows-1256                 Arabic (Windows)
		1257               windows-1257                 Baltic (Windows)
		1258               windows-1258                 Vietnamese (Windows)
		1361               Johab                        	Korean (Johab)
		10000              macintosh                    	Western European (Mac)
		10001              x-mac-japanese               Japanese (Mac)
		10002              x-mac-chinesetrad            Chinese Traditional (Mac)
		10003              x-mac-korean                 		Korean (Mac)
		10004              x-mac-arabic                 		Arabic (Mac)
		10005              x-mac-hebrew                 		Hebrew (Mac)
		10006              x-mac-greek                  		Greek (Mac)
		10007              x-mac-cyrillic               			Cyrillic (Mac)
		10008              x-mac-chinesesimp            	Chinese Simplified (Mac)
		10010              x-mac-romanian               	Romanian (Mac)
		10017              x-mac-ukrainian              		Ukrainian (Mac)
		10021              x-mac-thai                   			Thai (Mac)
		10029              x-mac-ce                     			Central European (Mac)
		10079              x-mac-icelandic              		Icelandic (Mac)
		10081              x-mac-turkish                		Turkish (Mac)
		10082              x-mac-croatian               		Croatian (Mac)
		12000              utf-32                       			Unicode (UTF-32)
		12001              utf-32BE                     			Unicode (UTF-32 Big-Endian)
		20000              x-Chinese-CNS                		Chinese Traditional (CNS)
		20001              x-cp20001                    		TCA Taiwan
		20002              x-Chinese-Eten               		Chinese Traditional (Eten)
		20003              x-cp20003                    		IBM5550 Taiwan
		20004              x-cp20004                    		TeleText Taiwan
		20005              x-cp20005                    		Wang Taiwan
		20105              x-IA5                        			Western European (IA5)
		20106              x-IA5-German                 		German (IA5)
		20107              x-IA5-Swedish                		Swedish (IA5)
		20108              x-IA5-Norwegian              		Norwegian (IA5)
		20127              us-ascii                     			US-ASCII
		20261              x-cp20261                    		T.61
		20269              x-cp20269                    		ISO-6937
		20273              IBM273                       			IBM EBCDIC (Germany)
		20277              IBM277                       			IBM EBCDIC (Denmark-Norway)
		20278              IBM278                       			IBM EBCDIC (Finland-Sweden)
		20280              IBM280                       			IBM EBCDIC (Italy)
		20284              IBM284                       			IBM EBCDIC (Spain)
		20285              IBM285                       			IBM EBCDIC (UK)
		20290              IBM290                       			IBM EBCDIC (Japanese katakana)
		20297              IBM297                       			IBM EBCDIC (France)
		20420              IBM420                       			IBM EBCDIC (Arabic)
		20423              IBM423                       			IBM EBCDIC (Greek)
		20424              IBM424                       			IBM EBCDIC (Hebrew)
		20833              x-EBCDIC-KoreanExtended 	IBM EBCDIC (Korean Extended)
		20838              IBM-Thai                     			IBM EBCDIC (Thai)
		20866              koi8-r                       			Cyrillic (KOI8-R)
		20871              IBM871                       			IBM EBCDIC (Icelandic)
		20880              IBM880                       			IBM EBCDIC (Cyrillic Russian)
		20905              IBM905                       			IBM EBCDIC (Turkish)
		20924              IBM00924                     		IBM Latin-1
		20932              EUC-JP                       			Japanese (JIS 0208-1990 and 0212-1990)
		20936              x-cp20936                    		Chinese Simplified (GB2312-80)
		20949              x-cp20949                    		Korean Wansung
		21025              cp1025                       			IBM EBCDIC (Cyrillic Serbian-Bulgarian)
		21866              koi8-u                       			Cyrillic (KOI8-U)
		28591              iso-8859-1                   		Western European (ISO)
		28592              iso-8859-2                   		Central European (ISO)
		28593              iso-8859-3                   		Latin 3 (ISO)
		28594              iso-8859-4                   		Baltic (ISO)
		28595              iso-8859-5                   		Cyrillic (ISO)
		28596              iso-8859-6                   		Arabic (ISO)
		28597              iso-8859-7                   		Greek (ISO)
		28598              iso-8859-8                   		Hebrew (ISO-Visual)
		28599              iso-8859-9                   		Turkish (ISO)
		28603              iso-8859-13                 		Estonian (ISO)
>> 	28605              iso-8859-15                 		Latin 9 (ISO)
		29001              x-Europa                     			Europa
		38598              iso-8859-8-i                 		Hebrew (ISO-Logical)
		50220              iso-2022-jp                  			Japanese (JIS)
		50221              csISO2022JP                  		Japanese (JIS-Allow 1 byte Kana)
		50222              iso-2022-jp                  			Japanese (JIS-Allow 1 byte Kana - SO/SI)
		50225              iso-2022-kr                  			Korean (ISO)
		50227              x-cp50227                    		Chinese Simplified (ISO-2022)
		51932              euc-jp                       			Japanese (EUC)
		51936              EUC-CN                       		Chinese Simplified (EUC)
		51949              euc-kr                       			Korean (EUC)
		52936              hz-gb-2312                   		Chinese Simplified (HZ)
		54936              GB18030                      		Chinese Simplified (GB18030)
		57002              x-iscii-de                   			ISCII Devanagari
		57003              x-iscii-be                   			ISCII Bengali
		57004              x-iscii-ta                   				ISCII Tamil
		57005              x-iscii-te                   				ISCII Telugu
		57006              x-iscii-as                   			ISCII Assamese
		57007              x-iscii-or                   				ISCII Oriya
		57008              x-iscii-ka                   			ISCII Kannada
		57009              x-iscii-ma                   			ISCII Malayalam
		57010              x-iscii-gu                   			ISCII Gujarati
		57011              x-iscii-pa                   			ISCII Punjabi
		65000              utf-7                        				Unicode (UTF-7)
		65001              utf-8                        				Unicode (UTF-8)

*/

	global q                     	:= Chr(0x22)                           	; ist dieses Zeichen -> "
	global Addendum

	SetRegView	, % (A_PtrSize = 8 ? 64 : 32)
	RegRead   	, AlbisMainPath 	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-MainPath
	RegRead    	, AlbisLocalPath 	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-LocalPath
	RegRead   	, AlbisExe           	, HKEY_CURRENT_USER\Software\ALBIS\Albis on Windows\Albis_Versionen, 1-Exe

	RegExMatch(A_ScriptDir, "Oi)^(?<Dir>.*?AlbisOnWindows)", adm)

	Addendum := Object()
	Addendum.Dir               	:= adm.Dir
	Addendum.DBPath 			:= adm.Dir "\logs'n'data\_DB"
	Addendum.AlbisDBPath 	:= AlbisMainPath "\db"


	ldt	:= new LDTransfer("Archiv")
	g  	:= new Output("LDT Fenster", ldt)

	;~ labdays := ["29.03.2022", "30.03.2022", "31.03.2022", "01.04.2022"]
	labdays := ["13.06.2022"] ;, "24.05.2022", "25.05.2022"]

	For each, labday in labdays {
		m		:= ldt.IgnoredTests(labday)
		x   	.= each ".Labortag: " labday " [maches: ###]`r`n" Output.Tabbed(m) "`r`n"
		t  		.= StrReplace(x, "###", m.Count() ", Länge: " StrLen(x))
				g.print(t)
	}

	m	:= ldt.PatientSearch("13.06.2022")
	If IsObject(m) {

		SciTEOutput(cJSON.Dump(m, 1))
		x 		:= g.Tabbed(m)
		t 		.= "matches: " m.Count() ", " StrLen(x) "`r`n" x
				g.print(t)


	}
		else
		SciTEOutput("kein Labortreffer")


	;~ newldt := ldt.RebuildLDT("4104920867", "20220228193007.209.X01492.LDT", "C:\LABOR")
	;~ t 		.= "`r`nnew .ldt-file path: " newldt "`r`n"
	;~ t 		.= FileOpen(newldt, "r", "CP850").Read()
	;~ g.print(t)

	;~ SciTEOutput(t)
	;~ SciTEOutput(cJSON.Dump(m, 1))


return
ExitApp


class LDTransfer {

	__New(path:="LDT", callbackFunc:="")             	{

	  ; C:\tmp oder \\SERVER\daten
		this.path 	:= path ~= "^([A-Z]:\\|\\\\\w+)" ? path : Addendum.DBPath "\Labordaten\" path
		this.cbfunc:= IsFunc(callbackFunc) ? callbackFunc : ""

		;~ SciTEOutput("2: " this.path  "`r`n3: " Addendum.Dir "`n4: " Addendum.DBPath "`n5: " Addendum.AlbisDBPath)

	}

	PatientSearch(labday, patient:="")                               	{

		; patient := "Nachname, Vorname Geburtsdatum"  (das Komma kennzeichnet einen Vornamen)

		If patient {
			RegExMatch(patient, "O)^(?<name>[\pL\-]+)*\s*(,\s*(?<prename>[\pL\-]+))*\s*(?<birth>\d\d\.\d\d.\d\d\d\d|\d{8})*", Pat)
			name 		:= Pat.name,
			prename 	:= Pat.prename
			birth      	:= InStr(Pat.birth, ".") ? this.ConvertToDBASEDate(Pat.birth) : Pat.birth

		  ; Vergleichskonditionen
			mconditions := 	name     	? 1 : 0
			mconditions += 	prename	? 1 : 0
			mconditions += 	birth      	? 1 : 0
		}

		matches 	:= Array()
		ldt         	:= this.Examinations(labday)
		For anfnr, lab in ldt {

			If patient {
				matchcount := 0
				matchcount += name   	&& lab.1_name   	= Name   	? 1 : 0
				matchcount += prename 	&& lab.2_vorname= prename 	? 1 : 0
				matchcount += birth     	&& lab.3_geburt	= birth      	? 1 : 0
			}

			If (matchcount = mconditions || !patient) {
				lab.anfnr := anfnr
				matches.Push(lab)
			}
		}

	return matches
	}

	RebuildLDT(anfnr, filename, importpath)             	    	{                  	; erstellt eine neue LDT mit alten Daten

		; gedacht für nachträgliches importieren von "vermissten Daten"

		ldt_header := true

		fobj  	:= FileOpen(importpath "\" filename, "w", "CP28605")
		ldtxt  	:= FileOpen(this.path "\" filename , "r", "CP28605").Read()
		tlines	:= StrSplit(ldtxt, "`n", "`r")


		For each, line in tlines {

			If !line
				continue

			cnt    	:= SubStr(line, 1, 3)
			key   	:= SubStr(line, 4, 4)
			val    	:= SubStr(line, 8, StrLen(line)-7)

		  ; - - - - - - - - - - - - - - - - - - - - - -
		  ; LDT-Header kopieren
		  ; - - - - - - - - - - - - - - - - - - - - - -
			If 		 ldt_header 			{

				If (key = "8000" && RegExMatch(val, "820(1|2)")) {
					ldt_header := false
					buffer := line "`r`n"
					SciTEOutput("ldt_body at " each)
				}
				else
					fobj.Write(line "`r`n")

			}

		  ; - - - - - - - - - - - - - - - - - - - - - -
		  ; LDT-Footer anfügen
		  ; - - - - - - - - - - - - - - - - - - - - - -
			else if ldt_footer 			{

				fobj.WriteLine(line "`r`n")

			}

		  ; - - - - - - - - - - - - - - - - - - - - - -
		  ; Datensätze prüfen
		  ; - - - - - - - - - - - - - - - - - - - - - -
			else If !ldt_header && !ldt_footer 	{

			  ; - - - - - - - - - - - - - - - - - - - - - -
			  ;	neuer Datensatz
			  ; - - - - - - - - - - - - - - - - - - - - - -
				If  (key = "8000" && RegExMatch(val, "(8201|8202|8221)")) {

				  ; - - - - - - - - - - - - - -
				  ;	buffer kopieren
				  ; - - - - - - - - - - - - - -
					If saveanfnr
						fobj.Write(buffer)

					saveanfnr := false, buffer := ""

				  ; - - - - - - - - - - - - - -
				  ;	Footer gefunden
				  ; - - - - - - - - - - - - - -
					If (val = "8221") {
						ldt_footer := true
						fobj.Write(line "`r`n")
						SciTEOutput("ldt_footer at " each)
					}
					else
						buffer .= line "`r`n"

				}

			  ; - - - - - - - - - - - - - - - - - - - - - -
			  ;	Anforderungsnummer
			  ; - - - - - - - - - - - - - - - - - - - - - -
				else if (key = "8310")  { ; Anforderungsnummer
					SciTEOutput("ldt_anfnr = " val " = " anfnr " " each )
					saveanfnr := val = anfnr ? true : false
					buffer .= line "`r`n"
				}

			  ; - - - - - - - - - - - - - - - - - - - - - -
			  ;	nichts davon
			  ; - - - - - - - - - - - - - - - - - - - - - -
				else
					buffer .= line "`r`n"

			}


		}

		fobj.Close()

	return importpath "\" filename
	}

	Examinations(labday:="", useWith:="Eingangsdatum") 	{

			static rxLabText := "i)^(Erreger)"
			static LDTLabel := {"#3101": "1_Name"
										, "#3102": "2_Vorname"
										, "#3103": "3_Geburt"
										, "#3110": "4_Geschlecht"
										, "#8301": "Eingangsdatum"      ; #8301|#8432
										, "#8432": "Abnahmedatum"
										, "#8433": "Abnahmezeit"
										, "#8310": "AnfNr"
										, "#8401": "6_Befundstatus"
										, "#8609": "7_Abrechnungstyp"
										, "#5001": "Abrechnung/Abrechnung"
										, "#5002": "Abrechnung/Untersuchungsart"
										, "#5005": "Abrechnung/Multiplikator"
										, "#8403": "Abrechnung/Abrechnungsart"
										, "#8406": "Abrechnung/Kosten"
										, "#8410": "Parameter/Parameter"
										, "#8434": "Parameter/Parameter"
										, "#8420": "Parameter/Wert"
										, "#8421": "Parameter/Einheit"
										, "#8422": "Parameter/Indikator"
										, "#8460": "Parameter/NormwertText"
										, "#8461": "Parameter/UGrenze"
										, "#8462": "Parameter/OGrenze"
										, "#8470": "Parameter/Bewertung"
										, "#8480": "Parameter/Wert"
										, "#8490": "Parameter/Beurteilung"}

			ldt := Object()
			files := []

		; Datum ins DBASE Format wandeln
			labday := !labday ? A_YYYY A_MM A_DD : labday
			labday := InStr(labday, ".") ? this.ConvertToDBASEDate(labday) : labday
		  ; Dateistempel, erfasse auch Dateien vor freien Tagen
			FormatTime, dayOweek, % labday "000000", % "ddd"
			daysBack 	:= 	dayOweek = "Mo" ? -3 : dayOweek = "So" ? -2 : -1
			labday  	+=	%daysBack%, days
			labday  	:= 	SubStr(labday, 1, 8)

		; alle .LDT Dateien im Backup-Pfad der Dateien ermitteln
			Loop, Files, % this.path "\*.LDT"
				files.Push({"name":A_LoopFileName, "ctime":A_LoopFileTimeCreated, "attrib":A_LoopFileAttrib})

			SciTEOutput("Labortag: ab dem " this.ConvertDBASEDate(labday) " mit " files.Count() " Untersuchungen.")
			ldtList := ""
			For fNR, file in files {

				fpath := this.path "\" file.name
				If (SubStr(file.ctime, 1, 8)+0 >= labday) {

					ldtxt := FileOpen(fpath, "r", "CP850").Read()
					Bezeichner := {}
					For Lnr, line in StrSplit(ldtxt, "`n", "`r") {

							cnt    	:= SubStr(line, 1, 3)
							key   	:= SubStr(line, 4, 4)
							val    	:= SubStr(line, 8, StrLen(line)-7)
							flabel 	:= LDTLabel["#" key]

							if (key = "8310")	    	{	; Anforderungsnummer

								anfnr := val
								If IsObject(ldt[anfnr]) {
									RegExMatch(anfnr, "\d+_(?<count>\d+)", anfnr_)
									anfnr_count := !anfnr_count ? 1 : anfnr_count +1
									anfnr := RegExReplace(anfnr, "_\d+", "_" anfnr_count)
								}

								ldt[anfnr]       	:=Object()
								ldt[anfnr].file 	:= file.name
								ldt[anfnr].path 	:= fpath

								nextanfnr 	:= false
								labparam	:= ""
								newline 	:= ""
								Bezeichner := {}

							}
							else If flabel && !nextanfnr     	{

								flabel := StrSplit(flabel, "/")

							 ; Eingangs- oder Abnahmedatum für die Filterung verwenden
							 ; die Verwendung eines Eingangsdatums filtert alle Daten die gleich alt oder jünger sind
							 ; über das Abnahmedatum werden nur die Daten eines Tages gefiltert
								If (useWith = "Abnahmedatum" && key = "8432") {
									flabel.1 := "5_Datum"
									If (val <> labday) {
										If IsObject(ldt[anfnr])
											ldt.Delete(anfnr)
										nextANFNR := true
										continue
									}
								}
								else If (useWith = "Eingangsdatum" && key = "8301") {
									flabel.1 := "5_Datum"
								}

								If !IsObject(Bezeichner[flabel.1])
									Bezeichner[flabel.1] := Object()

								If (flabel.1=flabel.2 && val <> "Auftrag" )  {

									Bezeichner[flabel.1].fsub	:= val
									If !IsObject(ldt[anfnr][flabel.1])
										ldt[anfnr][flabel.1] := Object()
									If !IsObject(ldt[anfnr][flabel.1][val])
										ldt[anfnr][flabel.1][val] := Object()
									continue

								}
								else If (val = "Auftrag" )
									Bezeichner[flabel.1].newline 	:= "`r`n"
								else if (key = "8470")                                              	; Labortext - Bewertung
									Bezeichner[flabel.1].newline 	:= " "

								fsub := Bezeichner[flabel.1].fsub
								If (flabel.1 && flabel.2)
									ldt[anfnr][flabel.1][fsub][flabel.2] .= val . Bezeichner[flabel.1].newline
								else
									ldt[anfnr][flabel.1] .= val
							}

					}
				}
			}

	return ldt
	}

	IgnoredTests(labday:="")                                             	{

		If !labday
			labday := A_YYYY A_MM A_DD

	  ; alle LDT-Eingänge ab diesem Tag ermitteln
		matches 	:= Array()
		ldt	:= this.Examinations(labday, "Abnahmedatum")
		For anfnr, lab in ldt {
			lab.anfnr := anfnr
			matches.Push(lab)
		}


	  ; alle Befundeeingänge in der LABBLATT.dbf auslesen
		aDB := new AlbisDB(Addendum.AlbisDBPath, "TT")
		Labblatt := aDB.LaborTagesDaten(labday, labday)
		;~ SciTEOutput(cJSON.Dump(Labblatt, 1))
		SciTEOutput("fertig")

	return matches
	}



	ConvertToDBASEDate(Date) {                                                                             	;-- Datumskonvertierung von DD.MM.YYYY nach YYYYMMDD
		RegExMatch(Date, "((?<Y1>\d{4})|(?<D1>\d{1,2})).(?<M>\d+).((?<Y2>\d{4})|(?<D2>\d{1,2}))", t)
	return (tY1?tY1:tY2) . SubStr("00" tM, -1) . SubStr("00" (tD1?tD1:tD2), -1)
	}

	ConvertDBASEDate(DBASEDate) {                                                                        	;-- Datumskonvertierung von YYYYMMDD nach DD.MM.YYYY
	return SubStr(DBaseDate, 7, 2) "." SubStr(DBaseDate, 5, 2) "." SubStr(DBaseDate, 1, 4)
	}

}

class Output {

	__New(guiname:="ldt", ldtobject:="") {

		this.g := this.Gui(guiname, ldtobject)

	}

	Gui(guiname, ldtobject:="") {

		global

		;~ SciTEOutput("Bin hier "  IsObject(ldtobject))
		Funcobj := ldtobject

		Gui, ldt: New, Hwndhwnd ; +AlwaysOnTop
		this.hwnd := hwnd

		Gui, ldt: Font, s10
		Gui, ldt: Add, Text, xm ym w60 , Name
		Gui, ldt: Add, Edit, x+3 yp-3 vldtName , % NAME

		Gui, ldt: Add, Text, xm y+5 w60, Vorname
		Gui, ldt: Add, Edit, x+3 yp-3  vldtPreName , % Vorname

		cp 	:= this.GuiControlGet("ldt", "Pos", "ldtName" )
		dp 	:= this.GuiControlGet("ldt", "Pos", "ldtPreName" )

		Gui, ldt: Add, Text, % "x" cp.X+cp.W+10 " y" cp.Y+4 " w130 Right", Geburtsdatum
		Gui, ldt: Add, Edit, x+3 yp-3 vldtBirth , % Geburt

		Gui, ldt: Add, Text, % "x" cp.X+cp.W+10 " y" dp.Y+2 " w130 Right " , ab Untersuchungstag
		Gui, ldt: Add, Edit, x+3 yp+1 vldtDate , % labDate

		Gui, ldt: Add, Button, x+10 yp+-4 vldtSearch gldtLabel, % "Suche starten"
		cp 	:= this.GuiControlGet("ldt", "Pos", "ldtName" )

		Gui, ldt: Font, s10, Consolas
		Gui, ldt: Add, Edit, % "xm y+10 w1100 h" 1000-cp.Y-cp.H-10 " vldtOut"

		Gui, ldt: Show, , LDT-Patientensuche

		return

		ldtGuiClose:
		ldtGuiEscape:
		ExitApp

		ldtLabel:

			Gui, ldt: Submit, NoHide

			If IsObject(Funcobj) {

				fn := Funcobj.Func("PatientSearch").Bind(ldtName (ldtPreName ? "," ldtPreName : "") " " ldtBirth)
				;~ m := fn.PatientSearch(NAME (Vorname ? "," Vorname : "") " " Geburtsdatum)
				m := %fn%()
				x   	:= this.Tabbed(m)
				t   	.= "matches: " m.Count() ", " StrLen(x) "`r`n" x

				this.print(t)

			}

		return

	}

	print(txt) {

		global ldt, ldtOut

		Gui, ldt: Default
		GuiControl, ldt:, ldtOut, % txt

	}

	Tabbed(matches) {

		t := ""
		For each, lab in matches  {

			t .= "AnfNr:  `t"    	lab.anfnr " (" (6_Befundstatus = "T" ? "Teil" : "End") "befund)"  " vom " LDTransfer.ConvertDBASEDate(lab.5_Datum) "`r`n"
			t .= "Patient:`t"  	(lab.1_Name ? lab.1_Name ", " lab.2_Vorname " *" : " -Geb.Datum ")  LDTransfer.ConvertDBASEDate(lab.3_Geburt) "`r`n"
			t .= "filepath:`t" 	lab.path "`r`n"
			t .= "filename:`t" 	lab.file "`r`n"

			x .= "[" each "] Patient:`t"  	(lab.1_Name ? lab.1_Name ", " lab.2_Vorname " *" : " -Geb.Datum ")  LDTransfer.ConvertDBASEDate(lab.3_Geburt) "`r`n"

			z := ""
			For labparam, res in lab.parameter {

				uGrenze	:= RegExReplace(res.UGrenze	, "^[0]+([0-9]+\.*[0-9]*)"              	, "$1")
				oGrenze	:= RegExReplace(res.OGrenze	, "^[0]+([0-9]+\.*[0-9]*)"              	, "$1")
				Wert        	:= (res.Beurteilung ? res.Wert : RTrim(res.Wert, "`r`n"))  res.Beurteilung
				Einheit  	:= StrReplace(res.Einheit, "æ", "µ")

				et1 	:= StrLen(labparam)<8 ? "`t`t": "`t"
				uo 	:= uGrenze && uGrenze ? 1 : 0
				ng 	:= uGrenze . (uo ? "-" : "") . ogrenze
				ng	:= RegExReplace(ng, "^[0]+([0-9]+\.*[0-9]*)", "$1")
				et2	:= StrLen(ng) 		< 8 	? "`t`t" : "`t"
				et3	:= StrLen(Einheit)	< 8 	? "`t`t" : "`t"

				t   	.= 	labparam  	. et1
							. 	ng             	. et2
							.	Einheit      	. et3
							.	res.Indikator
							. 	Wert  "`t"
							. 	res.NormalwertText "`r`n"

				If res.Bewertung
					z .= "`r`n" labparam " [Bewertung]:`r`n" res.Bewertung "`r`n"
				If res.Beurteilung
					z .= "`r`n" labparam "[Beurteilung]:`r`n" res.Beurteilung "`r`n"
			}

			t := RegExReplace(t, "[\r\n]+$") "`r`n"

			t .= z "`r`n"
			t .= "`r`n + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + `r`n"

		}

	return x . "`r`n" . t
	}

	GuiControlGet(guiname, cmd, vcontrol) {                                                        	;-- GuiControlGet wrapper
		GuiControlGet, cp, % guiname ": " cmd, % vcontrol
		If (cmd = "Pos")
			return {"X":cpX, "Y":cpY, "W":cpW, "H":cpH}
	return cp
}

}


#Include %A_ScriptDir%\..\lib\class_cJSON.ahk
#Include %A_ScriptDir%\..\lib\SciTEOutput.ahk

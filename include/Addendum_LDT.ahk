; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;               															⚗       Labordaten		💉
;
;      Funktionen:          	▫ 	Suche in ldt Dateien (Labordaten)
;                               	▫ 	eigenes Ausgabefenster mit Anzeigen aller Daten zu einzelnen Laborwerten, wie Wert, Normgrenzen und inbesondere
;										der Labortexte (Befundung, Hinweise von Fachgesellschaften ...)
;                               	▫ 	Extrahieren von einzelnen Anforderungen (nicht wirklich)
;                               	▫ 	ignorierte Daten aufspüren (noch nicht fertig)
;                               	▫ 	(Konvertierung in menschenlesenbaren Text)
;
;      Basisskript:           	keines
;
;		Hinweise:				Die Bibliothek besteht im wesentlichen aus zwei Klassenobjekte.
; 									Ruft man die Bibliothek als eigenständiges Skript auf öffnet sich ein Fenster für Suchanfragen.
;
;									LDTTransfer 		- 	stellt Daten aus gesicherten .LDT (Archiv-) Dateien zusammen
;									LDTOutput		- 	ist auf die Ausgabe der von LDTTransfert erstellten Daten angepasst
;
;									Verbindungen	-	die Klassen sind absichtlich (mit gewissen Ausnahmen) unabhängig von einander programmiert.
;																LDTransfert.Patientsearch() - verändert automatisch eine Progressbar im Ausgabefenster. Dies aber
;																nur wenn die LDTOutput zuvor initilisiert und eine globale Variable als wahr gekennzeichnet wurde.
;
;															- 	LDTOutput kann nach Übergabe des LDTTransfer-Objektes benötigte Daten lesen und hat so
;															  	Zugriff auf dessen Funktionen.
;
; 															- 	LDTTransfer nutzt globale Variablen um Daten zu erhalten. Steuerelementvariablen und Steuer-
;															  	elemente lassen sich in Autohotkey ohne Probleme von überall auslesen und auch von überall aus
;															 	in ihrem Verhalten steuern
;
;									weiteres			-	die folgende Abschnitte enthalten grundlegende Informationen zum ldt-Datenformat:
;
; 																1. "LDT Datensätze - Beschreibung" 	Es sind zunächst alle verfügbaren Textzeichen aufgeführt
;																													 	und sie finden anschließend eine sich in Ergänzung befindliche
;																													 	Liste mit Beschreibungen zu den sog. Feldkennungen
;
;																2. "Encoding Tabellen von Microsoft"
;
;		Abhängigkeiten:	▫	installiertes Albis on Windows
;									▫ 	Teile der Addendum Verzechnisstruktur werden aktuell noch benötigt.
;										Die ldt-Archivdateien werden im Verzeichnis Addendum für AlbisOnWindows\logs'n'data\_DB\Labordaten\LDT
;										oder \ein selbstgewählter Name
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - letzte Änderung 07.09.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


/* LDT Datensätze - Beschreibung

	.ldt ⇦⇨ Deutsch

	Encodiert sind .LDT Dateien in ISO 8859-15 (https://de.wikipedia.org/wiki/ISO_8859-15).
	In Notepad++ sind die Umlaute nicht lesbar bei einer Konvertierung mit ISO 8859-15. Nimm OEM 850!
	Für das Lese-Encoding mit Autohotkey nimm: CP850 , zum Erstellen von LDT-Dateien verwende am besten nur die Originalenkodierung CP28605

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
	0203 (N)BSNR-Bezeichnung (Praxisanrede))
	0205 Straße der (N)BSNR-Adresse
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
	8312 Kunden-(Arzt-)Nummer)
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


	9
	9103 Erstellungsdatum
	9106 Zeichensatzart
	9202 Datei-Ende?
	9212 LDT Version

	( Dateiende?
	01380008221
	014810000044
	017920200017305
	)

	([\p{L}\-\/]+)\s+([\d\.\-]+)\s+([\p{L}\.\sæ\/%]+)\s+([\+\d\.\-]+)   | $1 | $2 | $3 | $4 |
 */

/* Encoding Tabellen von Microsoft

https://docs.microsoft.com/de-de/dotnet/api/system.text.encodinginfo.getencoding?view=net-6.0
The example produces the following output when run on .NET Core:


Info.CodePage      Info.Name                        	Info.DisplayName
1200                	utf-16                               	Unicode
1201                	utf-16BE                           	Unicode (Big-Endian)
12000              	utf-32                               	Unicode (UTF-32)
12001              	utf-32BE                           	Unicode (UTF-32 Big-Endian)
20127              	us-ascii                             	US-ASCII
28591              	iso-8859-1                       	Western European (ISO)
65000              	utf-7                                 	Unicode (UTF-7)
65001              	utf-8                                 	Unicode (UTF-8)

The example produces the following output when run on .NET Framework:

Info.CodePage  	Info.Name                        	Info.DisplayName
		37            	IBM037                            	IBM EBCDIC (US-Canada)
		437          	IBM437                            	OEM United States
		500          	IBM500                            	IBM EBCDIC (International)
		708          	ASMO-708                       	Arabic (ASMO 708)
		720          	DOS-720                          	Arabic (DOS)
		737          	ibm737                            	Greek (DOS)
		775          	ibm775                            	Baltic (DOS)
>>	850          	ibm850                            	Western European (DOS)    	<< nehmen sie dem Codepage IBM850 in Notepad++ und in AHK (hier "CP850" verwenden) zum Lesen von .ldt Dateien.
																														   Den geeigneten Codepage für das Schreiben finden sie weiter unten.
		852          	ibm852                            	Central European (DOS)
		855          	IBM855                            	OEM Cyrillic
		857          	ibm857                            	Turkish (DOS)
		858          	IBM00858                         	OEM Multilingual Latin I
		860          	IBM860                            	Portuguese (DOS)
		861          	ibm861                            	Icelandic (DOS)
		862          	DOS-862                          	Hebrew (DOS)
		863          	IBM863                             	French Canadian (DOS)
		864          	IBM864                             	Arabic (864)
		865          	IBM865                             	Nordic (DOS)
		866          	cp866                               	Cyrillic (DOS)
		869          	ibm869                            	Greek, Modern (DOS)
		870          	IBM870                            	IBM EBCDIC (Multilingual Latin-2)
		874           	windows-874                    	Thai (Windows)
		875           	cp875                               	IBM EBCDIC (Greek Modern)
		932           	shift_jis                             	Japanese (Shift-JIS)
		936           	gb2312                            	Chinese Simplified (GB2312)
		949           	ks_c_5601-1987               	Korean
		950           	big5                                 	Chinese Traditional (Big5)
		1026         	IBM1026                          	IBM EBCDIC (Turkish Latin-5)
		1047         	IBM01047                         	IBM Latin-1
		1140         	IBM01140                         	IBM EBCDIC (US-Canada-Euro)
		1141         	IBM01141                         	IBM EBCDIC (Germany-Euro)
		1142         	IBM01142                         	IBM EBCDIC (Denmark-Norway-Euro)
		1143         	IBM01143                         	IBM EBCDIC (Finland-Sweden-Euro)
		1144         	IBM01144                         	IBM EBCDIC (Italy-Euro)
		1145         	IBM01145                         	IBM EBCDIC (Spain-Euro)
		1146         	IBM01146                         	IBM EBCDIC (UK-Euro)
		1147         	IBM01147                         	IBM EBCDIC (France-Euro)
		1148         	IBM01148                         	IBM EBCDIC (International-Euro)
		1149         	IBM01149                         	IBM EBCDIC (Icelandic-Euro)
		1200         	utf-16                               	Unicode
		1201         	utf-16BE                           	Unicode (Big-Endian)
		1250         	windows-1250                  	Central European (Windows)
		1251         	windows-1251                  	Cyrillic (Windows)
		1252         	windows-1252                  	Western European (Windows)
		1253         	windows-1253                  	Greek (Windows)
		1254         	windows-1254                  	Turkish (Windows)
		1255         	windows-1255                  	Hebrew (Windows)
		1256         	windows-1256                  	Arabic (Windows)
		1257         	windows-1257                  	Baltic (Windows)
		1258         	windows-1258                  	Vietnamese (Windows)
		1361         	Johab                               	Korean (Johab)
		10000      	macintosh                         	Western European (Mac)
		10001      	x-mac-japanese                	Japanese (Mac)
		10002      	x-mac-chinesetrad             	Chinese Traditional (Mac)
		10003      	x-mac-korean                    	Korean (Mac)
		10004      	x-mac-arabic                     	Arabic (Mac)
		10005      	x-mac-hebrew                   	Hebrew (Mac)
		10006      	x-mac-greek                      	Greek (Mac)
		10007      	x-mac-cyrillic                     	Cyrillic (Mac)
		10008      	x-mac-chinesesimp            	Chinese Simplified (Mac)
		10010      	x-mac-romanian                	Romanian (Mac)
		10017      	x-mac-ukrainian                	Ukrainian (Mac)
		10021      	x-mac-thai                        	Thai (Mac)
		10029      	x-mac-ce                          		Central European (Mac)
		10079      	x-mac-icelandic                 	Icelandic (Mac)
		10081      	x-mac-turkish                    	Turkish (Mac)
		10082      	x-mac-croatian                  	Croatian (Mac)
		12000      	utf-32                               	Unicode (UTF-32)
		12001      	utf-32BE                           	Unicode (UTF-32 Big-Endian)
		20000      	x-Chinese-CNS                 	Chinese Traditional (CNS)
		20001      	x-cp20001                        	TCA Taiwan
		20002      	x-Chinese-Eten                  	Chinese Traditional (Eten)
		20003      	x-cp20003                        	IBM5550 Taiwan
		20004      	x-cp20004                        	TeleText Taiwan
		20005      	x-cp20005                        	Wang Taiwan
		20105      	x-IA5                                	Western European (IA5)
		20106      	x-IA5-German                   	German (IA5)
		20107      	x-IA5-Swedish                   	Swedish (IA5)
		20108      	x-IA5-Norwegian              		Norwegian (IA5)
		20127      	us-ascii                             	US-ASCII
		20261      	x-cp20261                        	T.61
		20269      	x-cp20269                        	ISO-6937
		20273      	IBM273                            	IBM EBCDIC (Germany)
		20277      	IBM277                            	IBM EBCDIC (Denmark-Norway)
		20278      	IBM278                            	IBM EBCDIC (Finland-Sweden)
		20280      	IBM280                            	IBM EBCDIC (Italy)
		20284      	IBM284                            	IBM EBCDIC (Spain)
		20285      	IBM285                            	IBM EBCDIC (UK)
		20290      	IBM290                            	IBM EBCDIC (Japanese katakana)
		20297      	IBM297                            	IBM EBCDIC (France)
		20420      	IBM420                            	IBM EBCDIC (Arabic)
		20423      	IBM423                            	IBM EBCDIC (Greek)
		20424      	IBM424                            	IBM EBCDIC (Hebrew)
		20833      	x-EBCDIC-KoreanExtended	IBM EBCDIC (Korean Extended)
		20838      	IBM-Thai                           	IBM EBCDIC (Thai)
		20866      	koi8-r                               	Cyrillic (KOI8-R)
		20871      	IBM871                            	IBM EBCDIC (Icelandic)
		20880      	IBM880                            	IBM EBCDIC (Cyrillic Russian)
		20905      	IBM905                            	IBM EBCDIC (Turkish)
		20924      	IBM00924                        	IBM Latin-1
		20932      	EUC-JP                             	Japanese (JIS 0208-1990 and 0212-1990)
		20936      	x-cp20936                        	Chinese Simplified (GB2312-80)
		20949      	x-cp20949                        	Korean Wansung
		21025      	cp1025                             	IBM EBCDIC (Cyrillic Serbian-Bulgarian)
		21866      	koi8-u                               	Cyrillic (KOI8-U)
		28591      	iso-8859-1                        	Western European (ISO)
		28592      	iso-8859-2                        	Central European (ISO)
		28593      	iso-8859-3                        	Latin 3 (ISO)
		28594      	iso-8859-4                        	Baltic (ISO)
		28595      	iso-8859-5                        	Cyrillic (ISO)
		28596      	iso-8859-6                        	Arabic (ISO)
		28597      	iso-8859-7                        	Greek (ISO)
		28598      	iso-8859-8                        	Hebrew (ISO-Visual)
		28599      	iso-8859-9                        	Turkish (ISO)
		28603      	iso-8859-13                      	Estonian (ISO)
>> 	28605      	iso-8859-15                      	Latin 9 (ISO) << dieser Codepage (verwende "CP29605") funktioniert in AHK um .ldt Dateien zu erstellen
		29001      	x-Europa                           	Europa
		38598      	iso-8859-8-i                     	Hebrew (ISO-Logical)
		50220      	iso-2022-jp                      		Japanese (JIS)
		50221      	csISO2022JP                     	Japanese (JIS-Allow 1 byte Kana)
		50222      	iso-2022-jp                      		Japanese (JIS-Allow 1 byte Kana - SO/SI)
		50225      	iso-2022-kr                       	Korean (ISO)
		50227      	x-cp50227                        	Chinese Simplified (ISO-2022)
		51932      	euc-jp                               	Japanese (EUC)
		51936      	EUC-CN                           	Chinese Simplified (EUC)
		51949      	euc-kr                               	Korean (EUC)
		52936      	hz-gb-2312                      	Chinese Simplified (HZ)
		54936      	GB18030                          	Chinese Simplified (GB18030)
		57002      	x-iscii-de                           	ISCII Devanagari
		57003      	x-iscii-be                           	ISCII Bengali
		57004      	x-iscii-ta                            	ISCII Tamil
		57005      	x-iscii-te                            	ISCII Telugu
		57006      	x-iscii-as                           	ISCII Assamese
		57007      	x-iscii-or                            	ISCII Oriya
		57008      	x-iscii-ka                           	ISCII Kannada
		57009      	x-iscii-ma                          	ISCII Malayalam
		57010      	x-iscii-gu                           	ISCII Gujarati
		57011      	x-iscii-pa                           	ISCII Punjabi
		65000      	utf-7                                 	Unicode (UTF-7)
		65001      	utf-8                                 	Unicode (UTF-8)

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

 ; schon gehts los
	ldt	:= new LDTransfer("Archiv")
	g  	:= new LDTOutput("LDT Fenster", ldt)

	/*
	;~ labdays := ["23.08.2022"] ;, "30.03.2022", "31.03.2022", "01.04.2022"]
	;~ For each, labday in labdays {
		;~ m		:= ldt.IgnoredTests(labday)
		;~ x   	.= each ".Labortag: " labday " [maches: ###]`r`n" Output.Tabbed(m) "`r`n"
		;~ t  		.= StrReplace(x, "###", m.Count() ", Länge: " StrLen(x))
	;~ }
	;~ g.print(t)

	;~ m	:= ldt.PatientSearch("21.08.2022", "")
	;~ If IsObject(m) {

		;~ SciTEOutput(cJSON.Dump(m, 1))
		;~ x 		:= g.Tabbed(m)
		;~ t 		.= "matches: " m.Count() ", Anzahl übermittelter Zeichen: " StrLen(x) "`r`n" x

		;~ g.print(t)


	;~ }
		;~ else
		;~ SciTEOutput("kein Labortreffer")


	;~ newldt := ldt.RebuildLDT("4104920867", "20220228193007.209.X01492.LDT", "C:\LABOR")
	;~ t 		.= "`r`nnew .ldt-file path: " newldt "`r`n"
	;~ t 		.= FileOpen(newldt, "r", "CP850").Read()
	;~ g.print(t)

	;~ SciTEOutput(t)
	;~ SciTEOutput(cJSON.Dump(m, 1))
	 */

return
ExitApp


class LDTransfer {

	__New(path:="LDT", callbackFunc:="")                        	{

	  ; C:\tmp oder \\SERVER\daten
		this.path 	:= path ~= "^([A-Z]:\\|\\\\\w+)" ? path : Addendum.DBPath "\Labordaten\" path
		this.cbfunc:= IsFunc(callbackFunc) ? callbackFunc : ""

	}

	PatientSearch(labdays, patient:="")                               	{

		; patient := "Nachname, Vorname Geburtsdatum"  (das Komma kennzeichnet einen Vornamen)
		; labdays := Array(Tag1, Tag2, Tag3 [, TagN]) oder String "01.09.2022"
		;					Achtung: bei Übergabe eines Strings werden alle Daten beginnend mit dem Übergabedatum bis zum aktuellen Tag ausgewerten

		local rxname, rxprename, rxbirth, Name, prename, birth
		mconditions := matchcount := 0
		debug := true

		debug ? SciTEOutput("labdays: " (IsObject(labdays) ? cJson.dump(labdays) : labdays)) : ""

		If patient {
			RegExMatch(patient, "O)^(?<name>[\pL\-]+)*[\s,]*(?<prename>[\pL\-]+)*[\s,]*(?<birth>\d\d\.\d\d.\d\d\d\d|\d{8})*\s*$", Pat)
			name 		:= Pat.name
			prename 	:= Pat.prename
			birth      	:= InStr(Pat.birth, ".") ? this.ConvertToDBASEDate(Pat.birth) : Pat.birth

		  ; Vergleichskonditionen
			mconditions   := 	name     	? 1 : 0
			mconditions += 	prename	? 1 : 0
			mconditions += 	birth      	? 1 : 0
			rxname     	 :=	(name   	? ((SubStr(name, 1, 1)     	~= "^[A-ZÄÖÜ]" ? "^" : "i)") name    	) : "")
			rxprename   	 :=	(prename	? ((SubStr(prename, 1, 1)	~= "^[A-ZÄÖÜ]" ? "^" : "i)") prename	) : "")
			rxbirth       	 :=	birth ? "^" birth : ""

			debug ? SciTEOutput("Name, Vorname, Geb. Datum: " (name ? Name : "---") ", " (prename ? prename : "---") ", " (birth ? birth : "---") "`nRegExStrings: " rxname " | " rxprename " | " rxbirth) : ""
		}

		matches     	:= Array()
		opt           	:= IsObject(labdays) 	? "return_all=false" : "return_all=true"
		labdays   		:= !IsObject(labdays)	? [labdays] : labdays
		ldt             	:= this.Examinations(labdays, "", opt)
		For anfnr, lab in ldt {

			If patient {
				matchcount += (name   	&& lab.1_name   	 ~= rxname    	)	? 1 : 0
				matchcount += (prename && lab.2_vorname ~= rxprename)	? 1 : 0
				matchcount += (birth     	&& lab.3_geburt	 ~= birth         	)	? 1 : 0
			}

			If (matchcount = mconditions || !patient) {
				lab.anfnr := anfnr
				matches.Push(lab)
			}

			matchcount := 0

		}

		debug       	? (clipboard := cJson.dump(matches, 1)) 	: ""
		debug > 1 	? SciTEOutput(cJson.dump(matches, 1))  	: ""

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

	Examinations(labdays:="", useWith:="", opt:="")          	{                 		; Daten aus einer .LDT Datei in ein Autohotkey Objekt umwandeln

		global ldtProgress, ldtgui, ldtg, ldtfiletxt, ldtfilecount

		static rxLabText := "i)^(Erreger)"
		static LDTLabel := {"#3101": "1_Name"
									, "#3102": "2_Vorname"
									, "#3103": "3_Geburt"
									, "#3110": "4_Geschlecht"
									, "#8301": "Eingangsdatum"      ; #8301|#8432
									, "#8432": "Abnahmedatum"
									, "#8433": "Abnahmezeit"
									, "#8310": "AnfNr"
									, "#8311": "Berichtsdatum"
									, "#8312": "Berichtszeit"
									, "#8401": "6_Befundstatus"
									, "#8609": "7_Abrechnungstyp"
									, "#5001": "Abrechnung/Abrechnung"
									, "#5002": "Abrechnung/Untersuchungsart"
									, "#5005": "Abrechnung/Multiplikator"
									, "#8403": "Abrechnung/Abrechnungsart"
									, "#8406": "Abrechnung/Kosten"
									, "#8410": "Parameter/Parameter"
									, "#8411": "Parameter/Langtext"
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
		unknown := {}
		debug := true


	  ; Untersuchungsdatum (labdays) auf welchen Tag anwenden
		useWith := !useWith ? "Eingangsdatum" : useWith

	  ; opt parsen
		RegExMatch(opt, "i)return_all\s*=\s*(?<all>true|false|t|f|1|0|)(\s|$)", ret)
		retall := retall ~= "i)^(true|t|1|)$" ? true : false
		debug ? SciTEOutput(opt ", " retall ", Augabefenster: " ldtgui) : ""

	  ; Datum ins DBASE Format wandeln
		labday := !labday && !IsObject(labday) ? [A_YYYY A_MM A_DD] : IsObject(labday) ? labday : return

	  ; alle .LDT Dateien im Backup-Pfad der Dateien ermitteln
		Loop, Files, % this.path "\*.LDT"
			files.Push({"name":A_LoopFileName, "ctime":A_LoopFileTimeCreated, "attrib":A_LoopFileAttrib})

		If ldtgui {
			Gui, ldtg: Default
			GuiControl, ldtg: , ldtfilecount, % files.count()
		}

	  ; Progress Schritte auf max. 100 berechnen
		prgmod := Round(files.count()/100, 2)

	  ; es sollen nur Werte von genau einem Tag zurückgegeben werden, dann das Datum nicht zurückberechnet
		For labDayNr, labday in labdays {

		  ; Formatierung ändern
			slabday := labday := InStr(labday, ".") ? this.ConvertToDBASEDate(labday) : labday

		  ; Dateistempel, erfasse auch Dateien vor freien Tagen (fängt bei einem früheren Datum an) - Ausführung nur bei opt: return_all=true
			FormatTime, dayOweek, % slabday "000000", % "ddd"
			daysBack 	:= 	dayOweek = "Mo" ? -3 : dayOweek = "So" ? -2 : -1
			slabday  	+=	%daysBack%, days
			slabday  	:= 	SubStr(slabday, 1, 8)

		; debugging
			debug ? SciTEOutput("Labortag: ab dem " this.ConvertDBASEDate(labday) " mit " files.Count() " Untersuchungen.") : ""

			ldtList := ""
			For fNR, file in files {

			  ; Progress anzeigen
				If (ldtgui && Floor(Mod(fNR, prgmod)) = 0) {
					debug > 1 ? SciTEOutput(Round(fNR/prgmod) "% " prgmod ": " Floor(Mod(fNR, prgmod))  "| " fNR "/" files.count())
					GuiControl, ldtg:, ldtProgress, % Round(fNR/prgmod)
				}

				filedate := SubStr(file.ctime, 1, 8)+0
				If (filedate >= slabday) { 	;|| (!retall && filedate = labday)

				  ; debugging
					debug > 1 ? SciTEOutput(" [" fNR "] (" SubStr(file.name, 1, 8) ")" SubStr(file.name, 9, StrLen(file.name)-8)  "`t>= " slabday) : ""

					ldtxt := FileOpen( this.path "\" file.name, "r", "CP850").Read()
					Bezeichner := {}
					For Lnr, line in StrSplit(ldtxt, "`n", "`r") {

							cnt    	:= SubStr(line, 1, 3)
							key   	:= LTrim(SubStr(line, 4, 4), A_Space)
							lkey  	:= key ? "#" (StrLen(key) = 3 ? "0" : "") key : ""   ; label_key - da die führende Null sonst entfernt wird mit Raute am Anfang
							val    	:= SubStr(line, 8, StrLen(line)-7)
							flabel 	:= lkey ? LDTLabel[lkey] : ""

							If 	(lkey && !ldtLabel.haskey(lkey))
								unknown[lkey] := unknown[lkey] ? unknown[lkey]+1 : 1

							if (key = "8310")	    	{	; Anforderungsnummer

								anfnr := val
								If IsObject(ldt[anfnr]) {
									RegExMatch(anfnr, "\d+_(?<count>\d+)", anfnr_)
									anfnr_count := !anfnr_count ? 1 : anfnr_count +1
									anfnr := RegExReplace(anfnr, "_\d+", "_" anfnr_count)
								}

								ldt[anfnr]       	:= Object()
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
									flabel.1 := "5_Datum", flabel.2 := ""
									If (val <> labday) {
										If IsObject(ldt[anfnr])
											ldt.Delete(anfnr)
										nextANFNR := true
										continue
									}
								}
								else If (useWith = "Eingangsdatum" && key = "8301") {
									flabel.1 := "5_Datum", flabel.2 := ""
									If !retall
										If (val <> labday) {
											If IsObject(ldt[anfnr])
												ldt.Delete(anfnr)
										nextANFNR := true
										continue
									}
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
									ldt[anfnr][flabel.1][fsub][flabel.2] .= (ldt[anfnr][flabel.1][fsub][flabel.2] ? Bezeichner[flabel.1].newline : "") . val  ; verhindert Leerzeichen vor dem String
								else
									ldt[anfnr][flabel.1] .= !RegExMatch(ldt[anfnr][flabel.1], "(^|,)" val) ? (ldt[anfnr][flabel.1] ? ",":"" ) val : ""
							}

					}
				}
			}
		}


	; Progress auf 100% setzen
		If ldtgui {
			GuiControl, ldtg:, ldtProgress	, 100
			;~ GuiControl, ldtg:, ldtfiletxt  	, % "untersuchte LDT-Daten-Dateien:"
		}

	  ; alle unbekannten Tags ausgeben
		debug ? SciTEOutput(cJson.dump(unknown, 1)) : ""

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

	ConvertToDBASEDate(Date)                                       	{                    	;-- Datumskonvertierung von DD.MM.YYYY nach YYYYMMDD
		RegExMatch(Date, "((?<Y1>\d{4})|(?<D1>\d{1,2})).(?<M>\d+).((?<Y2>\d{4})|(?<D2>\d{1,2}))", t)
	return (tY1?tY1:tY2) . SubStr("00" tM, -1) . SubStr("00" (tD1?tD1:tD2), -1)
	}

	ConvertDBASEDate(DBASEDate)                                  	{                     	;-- Datumskonvertierung von YYYYMMDD nach DD.MM.YYYY
	return SubStr(DBaseDate, 7, 2) "." SubStr(DBaseDate, 5, 2) "." SubStr(DBaseDate, 1, 4)
	}

}

class LDTOutput {

	__New(guiname:="ldtg", ldtobject:="") {

		this.hEdit := 0
		this.hwnd := 0
		this.lastfoundpos 	:= 1
		this.debug := true
		this.funcobj := IsObject(ldtobject) ? ldtobject : ""

		this.Gui(guiname, ldtobject)

	}

	Gui(guiname, ldtobject:="") {

			global ldtName, ldtPreName, ldtBirth, ldtSearch, ldtProgress, ldtfiletxt, ldtfilecount, ldtOut, ldtResultSearch, examDay, ldtDate, ldtg
			global ldtgui := true

			Gui, ldtg: New, +Hwndhwnd 	; +AlwaysOnTop
			this.guihwnd := hwnd

			Gui, ldtg: Font, s10
			Gui, ldtg: Add, Text, xm ym w60 , Name
			Gui, ldtg: Add, Edit, x+3 yp-3 w150 vldtName , % Name

			Gui, ldtg: Add, Text, xm y+5 w60, Vorname
			Gui, ldtg: Add, Edit, x+3 yp-3 w150 vldtPreName , % Vorname

			cp 	:= this.GuiControlGet("ldtg", "Pos", "ldtName" )
			dp 	:= this.GuiControlGet("ldtg", "Pos", "ldtPreName" )

			Gui, ldtg: Add, Text, % "x" cp.X+cp.W+10 " y" cp.Y+2 " w130 Right", % "Geburtsdatum"
			Gui, ldtg: Add, Edit, % "x+3 y" cp.Y " w150 vldtBirth" , % Geburt
			Gui, ldtg: Add, Button, % "x+10 y" cp.Y-3 " vldtResultSearch", % "in gefundenen Daten suchen"
			GuiControl, ldtg: Enable0, ldtResultSearch

			Gui, ldtg: Add, Text, % "x" dp.X+dp.W+10 " y" dp.Y+2 " w130 Right vexamDay" , % "ab Untersuchungstag"
			Gui, ldtg: Add, Edit, % "x+3 y" dp.Y " w150 vldtDate" , % labDate

			Gui, ldtg: Add, Button, % "x+10 y" dp.Y-3 " vldtSearch", % "Suche starten"

			cp 	:= this.GuiControlGet("ldtg", "Pos", "ldtSearch" )
			Gui, ldtg: Add, Text    	, % "x" cp.X+cp.W+10 " yp+13 Right cBlue Backgroundtrans vldtfiletxt", % "untersuche LDT-Daten-Dateien:"
			Gui, ldtg: Add, Text    	, % "x+5 w50 cBlue Backgroundtrans vldtfilecount"

		 ; Progress
			Gui, ldtg: Add, Progress, x-2 y+6 h5 cD92626 BackgroundAAAAAA hwndhProgress vldtProgress, 0

			Gui, ldtg: Font, s10 cBlack , Consolas
			cp 	:= this.GuiControlGet("ldtg", "Pos", "ldtProgress" )
			Gui, ldtg: Add, Edit, % "xm y+6 w1100 h" 1000-cp.Y-cp.H-200 " +0x400 vldtOut hwndhEdit"
			this.hEdit := hEdit

		  ; gLabels als Klassenfunktionen einbinden
			BINDING := this.SearchInFiles.Bind(this)
			GUICONTROL, ldtg: +G, ldtSearch       	, % BINDING
			BINDING := this.GoSearch.Bind(this)
			GUICONTROL, ldtg: +G, ldtResultSearch	, % BINDING

			WinSet, ExStyle, 0x0, % "ahk_id " hProgress

			Gui, ldtg: Show, Hide , LDT-Patientensuche

			DetectHiddenWindows, On
			WinGetPos, X, Y, W, H, % "ahk_id " this.guihwnd
			DetectHiddenWindows, Off

			GuiControl, ldtg: Move, ldtProgress, % "w" W

			cp 	:= this.GuiControlGet("ldtg", "Pos", "ldtfiletxt")
			dp 	:= this.GuiControlGet("ldtg", "Pos", "ldtfilecount")
			GuiControl, ldtg: Move, ldtfiletxt      	, % "x" W-cp.W-dp.W-10
			cp 	:= this.GuiControlGet("ldtg", "Pos", "ldtfiletxt")
			GuiControl, ldtg: Move, ldtfilecount	, % "x" cp.X+cp.W+5

			Gui, ldtg: Show

			this.debug ? SciTEOutput("2 GuiHwnd, width, height: " this.guihwnd ", " W+4 ", " H) : ""

		return

		ldtgGuiClose:
		ldtgGuiEscape:
		ExitApp

		ldtLabel:

			;~ Critical
			;~ Gui, ldtg: Submit, NoHide
			;~ If IsObject(this.funcobj) {

				;~ fn := funcobj.Func("PatientSearch").Bind(ldtName "," (ldtPreName ? "," ldtPreName : "") " " ldtBirth)
				;~ m := fn.PatientSearch(NAME (Vorname ? "," Vorname : "") " " Geburtsdatum)
				;~ m := %fn%()
				;~ x   	:= this.Tabbed(m)
				;~ t   	.= "matches: " m.Count() ", " StrLen(x) "`r`n" x

				;~ this.print(t)

			;~ }

		return

	}

	SearchInFiles(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="") {

		global ldtName, ldtPreName, ldtBirth, ldtg, ldtDate

		Critical
		Gui, ldtg: Submit, NoHide
		If IsObject(this.funcobj) {

			ldtDate := StrReplace(ldtDate, A_Space)
			If !(ldtDate ~= "\d{1,2}\.\d{1,2}\.(\d{4}|\d{2})|\d{8}") {
				MsgBox, Bitte geben Sie ein Datum für die Suche ein
				return
			}

		  ; Parameterwerte für LDTTransfer.PatientSearch() aufbereiten
			patient := ldtName || ldtPreName || ldtBirth ? ldtName "," (ldtPreName ? "," ldtPreName : "") " " ldtBirth : ""
			this.debug ? SciTEOutput("Untersuchungstage: " ldtDate "`nPatient: " patient) : ""
			ldtDate := InStr(ldtDate, ",") ? StrSplit(ldtDate, ",") : ldtDate

			;~ fn := this.funcobj.PatientSearch	;.Bind(ldtDate, patient)
			;~ fn := funcobj.Func("PatientSearch").Bind(ldtDate, ldtName "," (ldtPreName ? "," ldtPreName : "") " " ldtBirth)
			;~ m := fn.PatientSearch(NAME (Vorname ? "," Vorname : "") " " Geburtsdatum)
			;~ m := %fn%(ldtDate, patient)
			m := this.funcobj.PatientSearch(ldtDate, patient)
			x   	:= this.Tabbed(m)
			t   	.= "matches: " m.Count() ", " StrLen(x) "`r`n" x

			this.print(t)

		}

	}

	GoSearch(GChwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="") {

		global ldtName, ldtPreName, ldtBirth, ldtSearch, ldtOut, examDay, ldtDate

		Gui, ldtg: Submit, NoHide


		sstring := 	(ldtName   	? ldtName     	: "")
					.  	(ldtName    	&& ldtPreName ? ",\s*"	: 	ldtName && !ldtPreName && ldtBirth	? ".*?" 	: "")
					. 	(ldtPreName	? ldtPreName 	: "")    	(ldtPreName 	&& ldtBirth                 	? "\s+" 	: "")
					. 	(ldtBirth     	? "\*" ldtBirth   	: "")

		If (this.lastsstring  != sstring) {
			this.lastsstring    	:= sstring
			this.lastfoundpos 	:= 1
		}

		ControlGetText, ctext,, % "ahk_id " this.hEdit
		fpos := RegExMatch(ctext, "i)" sstring, RegExOut, this.lastfoundpos)
		If (fpos ~= "^(\-1|)$") {
			MsgBox, % " Das Suchmuster (" RegExOut ") wurde nicht gefunden.`n" ;fpos ","  ", " this.hEdit
			return
		}

		lfc 	:= this.Edit_LineFromChar(fpos)               	; line from char
		fvl 	:= this.Edit_GetFirstVisibleLine()					; first visible line
		lvl 	:= this.Edit_GetLastVisibleLine()					; last visible line
		lpp	:= lvl - fvl                                                 ; lines per page

	  ; Edit scrollen, wenn Suchstring ausserhalb des angezeigten Bereiches liegt
		If (lfc < fvl || lfc > lvl){
			spages 	:= Floor( (lfc - fvl)/lpp)
			slines	:= lfc - (spages * lpp)
			;~ sl        	:= this.Edit_Scroll(spages, slines)
		}

		SciTEOutput( " [" sstring "] hEdit: " this.hEdit " | ctextL:" StrLen(ctext) " |"
							. "`n lfpos: " this.lastfoundpos ", fpos: " fpos ", lfchar: " lfc ", fvisl: " fvl ", lvisl: " lvl
							. "`n lperp: " lpp ", spages: " spages ", slines: " slines
							. "`n scrolllines: " sl "`n rxOut: " RegExOut)

		this.Edit_SetSel(fpos, StrLen(RegExOut))
		this.lastfoundpos 	:= fpos

	}

	print(txt) {

		global ldtg, ldtOut

		Gui, ldtg: Default
		GuiControl, ldtg:, ldtOut, % txt

	}

	Tabbed(matches) {

		t := ""
		x := "Index     AnfNr     Befund vom   Status    Patient`r`n----------------------------------------------------------------------------------------------`r`n"
		For each, lab in matches  {

			t .= "[" SubStr("00000" each, -2) "] - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - `r`n"
			t .= "AnfNr:  `t"    	lab.anfnr " (" (lab.6_Befundstatus = "T" ? "Teil" : "End") "befund)"  " vom " (thisday := LDTransfer.ConvertDBASEDate(lab.5_Datum)) "`r`n"
			t .= "Patient:`t"  	(lab.1_Name ? lab.1_Name ", " lab.2_Vorname " *" : " -Geb.Datum ")  (thisbirth := LDTransfer.ConvertDBASEDate(lab.3_Geburt)) "`r`n"
			t .= "filepath:`t" 	lab.path "`r`n"

			x .= "[" SubStr("00000" each, -2) "] " SubStr("                 " lab.anfnr, -10) "   " thisday "      " lab.6_Befundstatus "      "
					. (lab.1_Name ? lab.1_Name ", " lab.2_Vorname " *" : " -Geb.Datum *")  thisbirth " (" lab.4_Geschlecht ")" "`r`n"

			z := ""
			For labparam, res in lab.parameter {

				uGrenze	:= RegExReplace(res.UGrenze	, "^[0]+([0-9]+\.*[0-9]*)"              	, "$1")
				oGrenze	:= RegExReplace(res.OGrenze	, "^[0]+([0-9]+\.*[0-9]*)"              	, "$1")
				Wert        	:= (res.Beurteilung ? res.Wert : RTrim(res.Wert, "`r`n")) res.Beurteilung
				Einheit  	:= StrReplace(res.Einheit, "æ", "µ")

				et1 	:= StrLen(labparam)<8 ? "`t`t": "`t"
				uo 	:= StrLen(uGrenze) && StrLen(oGrenze) ? 1 : 0
				ng 	:= Trim(uGrenze) . (uo ? "-" : " ") . oGrenze
				ng	:= RegExReplace(ng, "^[0]+([0-9]+\.*[0-9]*)", "$1")
				et2	:= StrLen(ng) 		< 8 	? "`t`t" : "`t"
				et3	:= StrLen(Einheit)	< 8 	? "`t`t" : "`t"
				et4	:= Wert && Trim(res.NormalwertText)  ? "" . "`t"

				t   	.= 	labparam  	. et1
							. 	ng             	. et2
							.	Einheit      	. et3
							.	Trim(res.Indikator)
							. 	Wert           	. et4
							. 	res.NormalwertText 	"`r`n"

				If (res.Bewertung || res.Beurteilung)
					z .= "`r`n" labparam (res.Bewertung ? " [Bewertung]:`r`n" res.Bewertung : " [Beurteilung]:`r`n" res.Beurteilung) "`r`n"

			}

			t := RegExReplace(t, "[\r\n]+$") "`r`n"

			t .= z "`r`n"
			;~ t .= "`r`n 🧪  🧫  🧬 🩹  🩺  🩸  💉  🦠  ⚗  🧪  🧫  🧬 🩹  🩺  🩸  💉  🦠  ⚗  🧪  🧫  🧬 🩹  🩺  🩸  💉  🦠  ⚗  🧪  🧫  🧬 🩹  🩺  🩸  💉  🦠  ⚗  `r`n"
			;~ t .= "`r`n + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + + `r`n"

		}

	return x . "`r`n" . t
	}

	GuiControlGet(guiname, cmd, vcontrol) {                                                        	;-- GuiControlGet wrapper
		GuiControlGet, cp, % guiname ": " cmd, % vcontrol
		If (cmd = "Pos")
			return {"X":cpX, "Y":cpY, "W":cpW, "H":cpH}
	return cp
}

	Edit_SetSel(p_StartSelPos=0,p_EndSelPos=-1){

	;-----------------------------
	;
	; Function: Edit_SetSel
	;
	; Description:
	;
	;   Selects a range of characters.
	;
	; Parameters:
	;
	;   p_StartSelPos - Starting character position of the selection.  If set to -1,
	;       the current selection (if any) will be deselected.
	;
	;   p_EndSelPos - Ending character position of the selection.  Set to -1 to use
	;       the position of the last character in the control.
	;
	;-------------------------------------------------------------------------------
		Static EM_SETSEL:=0x0B1
		SendMessage EM_SETSEL,p_StartSelPos,p_EndSelPos,,% "ahk_id " this.hEdit
	}

	Edit_LineScroll(xScroll=0,yScroll=0)     {

	;-----------------------------
	;
	; Function: Edit_LineScroll
	;
	; Description:
	;
	;   Scrolls the text vertically or horizontally in a multiline Edit control.
	;
	; Parameters:
	;
	;   xScroll, yScroll - The number of characters to scroll horizontally (xScroll)
	;       or vertically (yScroll).  Use a negative number to scroll to the left
	;       (xScroll) or up (yScroll) and a positive number to scroll to the right
	;       (xScroll) or to scroll down (yScroll).  Alternatively, these parameters
	;       can contain one or more of the following values:
	;
	;       (start code)
	;       Option  Description
	;       ------   -----------
	;       Left    Scroll to the left edge of the control.
	;       Right   Scroll to the right edge of the control.
	;       Top     Scroll to the top of the control.
	;       Bottom  Scroll to the bottom of the control.
	;
	;       If more than one option is specified, the options must be delimited by a
	;       space.  Ex: "Top Left".  See the *Remarks* section for more information.
	;       (end)
	;
	; Remarks:
	;
	;   The xScroll parameter is processed first and then yScroll.  If either of
	;   these parameters contains multiple values (Ex: "Top Left"), the values are
	;   processed individually from left to right.  If there are conflicting values
	;   (Ex: "Top Bottom"), the last value specified will take precedence.
	;
	;-------------------------------------------------------------------------------

    Static Dummy3496

          ;-- Horizontal scroll values
          ,SB_LEFT :=6
          ,SB_RIGHT:=7

          ;-- Vertical scroll values
          ,SB_TOP   :=6
          ,SB_BOTTOM:=7

          ;-- Messages
          ,EM_LINESCROLL:=0xB6
          ,WM_HSCROLL   :=0x114
          ,WM_VSCROLL   :=0x115

    if xScroll is Integer
        {
        if xScroll  ;-- Any value other than 0
            SendMessage EM_LINESCROLL,xScroll,0,,% "ahk_id " this.hEdit
        }
     else
        Loop Parse,xScroll, % A_Space
            {
            if InStr(A_LoopField,"Left")
                SendMessage WM_HSCROLL,SB_LEFT,0,,% "ahk_id " this.hEdit
             else if InStr(A_LoopField,"Right")
                SendMessage WM_HSCROLL,SB_RIGHT,0,,% "ahk_id " this.hEdit
             else if InStr(A_LoopField,"Top")
                SendMessage WM_VSCROLL,SB_TOP,0,,% "ahk_id " this.hEdit
             else if InStr(A_LoopField,"Bottom")
                SendMessage WM_VSCROLL,SB_BOTTOM,0,,% "ahk_id " this.hEdit
            }

    if yScroll is Integer
        {
        if yScroll  ;-- Any value other than 0
            SendMessage EM_LINESCROLL,0,yScroll,,% "ahk_id " this.hEdit
        }
     else
        Loop Parse,yScroll, % A_Space
            {
            if InStr(A_LoopField,"Left")
                SendMessage WM_HSCROLL,SB_LEFT,0,,% "ahk_id " this.hEdit
             else if InStr(A_LoopField,"Right")
                SendMessage WM_HSCROLL,SB_RIGHT,0,,% "ahk_id " this.hEdit
             else if InStr(A_LoopField,"Top")
                SendMessage WM_VSCROLL,SB_TOP,0,,% "ahk_id " this.hEdit
             else if InStr(A_LoopField,"Bottom")
                SendMessage WM_VSCROLL,SB_BOTTOM,0,,% "ahk_id " this.hEdit
            }
    }

	Edit_Scroll(p_Pages=0,p_Lines=0)     {

	;-----------------------------
	;
	; Function: Edit_Scroll
	;
	; Description:
	;
	;   Scrolls the text vertically in a multiline Edit control.
	;
	; Parameters:
	;
	;   p_Pages - The number of pages to scroll.  Use a negative number to page up
	;       and a positive number to page down.
	;
	;   p_Lines - The number of lines to scroll.  Use a negative number to scroll up
	;       and a positive number to scroll down.
	;
	; Returns:
	;
	;   The number of lines that the command scrolls.  The value will be negative if
	;   scrolling up, positive if scrolling down, and zero (0) if no scrolling
	;   occurred.
	;
	;-------------------------------------------------------------------------------

    Static EM_SCROLL  :=0xB5
          ,SB_LINEUP  :=0x0     ;-- Scroll up one line
          ,SB_LINEDOWN:=0x1     ;-- Scroll down one line
          ,SB_PAGEUP  :=0x2     ;-- Scroll up one page
          ,SB_PAGEDOWN:=0x3     ;-- Scroll down one page

    ;-- Initialize
    l_ScrollLines:=0

    ;-- Pages
    Loop % Abs(p_Pages)        {
        SendMessage EM_SCROLL,(p_Pages>0) ? SB_PAGEDOWN:SB_PAGEUP,0,,% "ahk_id " this.hEdit
        if !ErrorLevel
            Break

        l_ScrollLines+=((ErrorLevel&0xFFFF)<<48>>48)
            ;-- LOWORD of result and converted from UShort to Short
        }

    ;-- Lines
    Loop % Abs(p_Lines)        {
        SendMessage EM_SCROLL,(p_Lines>0) ? SB_LINEDOWN:SB_LINEUP,0,,% "ahk_id " this.hEdit
        if !ErrorLevel
            Break

        l_ScrollLines+=((ErrorLevel&0xFFFF)<<48>>48)
            ;-- LOWORD of result and converted from UShort to Short
        }

    Return l_ScrollLines
    }

	Edit_ScrollCaret(){

		;-----------------------------
		;
		; Function: Edit_ScrollCaret
		;
		; Description:
		;
		;   Scrolls the caret into view in an Edit control.
		;
		;-------------------------------------------------------------------------------
		Static EM_SCROLLCARET:=0xB7
		SendMessage EM_SCROLLCARET,0,0,,% "ahk_id " this.hEdit
	}

	Edit_ScrollPage(p_HPages=0,p_VPages=0){

		;-----------------------------
		;
		; Function: Edit_ScrollPage
		;
		; Description:
		;
		;   Scrolls the Edit control by page.
		;
		; Parameters:
		;
		;   p_HPages - The number of horizontal pages to scroll.  Use a postive number
		;       to page right and a negative number to page left.
		;
		;   p_VPages - The number of vertical pages to scroll. [Optional] Use a positive
		;       number to page down and a negative number to page up.
		;
		; Remarks:
		;
		;   This function duplicates some of the functionality of <Edit_Scroll>.  If
		;   scrolling vertically and the return value is needed, use the *Edit_Scroll*
		;   function instead.
		;
		;------------------------------------------------------------------------------


		Static Dummy3535

			  ;-- Horizontal scroll values
			  ,SB_PAGELEFT :=2
			  ,SB_PAGERIGHT:=3

			  ;-- Vertical scroll values
			  ,SB_PAGEUP  :=2
			  ,SB_PAGEDOWN:=3

			  ;-- Messages
			  ,WM_HSCROLL :=0x114
			  ,WM_VSCROLL :=0x115

		;-- Horizontal
		Loop % Abs(p_HPages)
			SendMessage WM_HSCROLL,(p_HPages>0) ? SB_PAGERIGHT:SB_PAGELEFT,0,,% "ahk_id " this.hEdit

		;-- Vertical
		Loop % Abs(p_VPages)
			SendMessage WM_VSCROLL,(p_VPages>0) ? SB_PAGEDOWN:SB_PAGEUP,0,,% "ahk_id " this.hEdit

	}

	Edit_LineFromChar(p_CharPos=-1)    {
		;-----------------------------
		;
		; Function: Edit_LineFromChar
		;
		; Description:
		;
		;   Gets the index of the line that contains the specified character index.
		;
		; Parameters:
		;
		;   p_CharPos - The character index of the character contained in the line
		;       whose number is to be retrieved. [Optional] If �1 (the default), the
		;       function retrieves either the line number of the current line (the line
		;       containing the caret) or, if there is a selection, the line number of
		;       the line containing the beginning of the selection.
		;
		; Returns:
		;
		;   The zero-based line number of the line containing the character index
		;   specified by p_CharPos.
		;
		;-------------------------------------------------------------------------------
		Static EM_LINEFROMCHAR:=0xC9
		SendMessage EM_LINEFROMCHAR,p_CharPos,0,,% "ahk_id " this.hEdit
	Return ErrorLevel
	}

	Edit_GetFirstVisibleLine()    {

		;-----------------------------
		;
		; Function: Edit_GetFirstVisibleLine
		;
		; Description:
		;
		;   Returns the zero-based index of the uppermost visible line.  For single-line
		;   Edit controls, the return value is the zero-based index of the first visible
		;   character.
		;
		;-------------------------------------------------------------------------------
		Static EM_GETFIRSTVISIBLELINE:=0xCE
		SendMessage EM_GETFIRSTVISIBLELINE,0,0,, % "ahk_id " this.hEdit
	Return ErrorLevel
	}

	Edit_GetLastVisibleLine() {

	;-----------------------------
	;
	; Function: Edit_GetLastVisibleLine
	;
	; Description:
	;
	;   Returns the zero-based line index of the last visible line on the edit
	;   control.
	;
	; Calls To Other Functions:
	;
	; * <Edit_GetRect>
	; * <Edit_LineFromPos>
	;
	; Remarks:
	;
	;   To calculate the total number of visible lines, use the following...
	;
	;       (start code)
	;       Edit_GetLastVisibleLine(hEdit) - Edit_GetFirstVisibleLine(hEdit) + 1
	;       (end)
	;
	;-------------------------------------------------------------------------------
    this.Edit_GetRect(Left, Top, Right, Bottom)
	Return this.Edit_LineFromPos(0, Bottom-1)
	}

	Edit_GetRect(ByRef r_Left="",ByRef r_Top="",ByRef r_Right="",ByRef r_Bottom="")    {
		;-----------------------------
		;
		; Function: Edit_GetRect
		;
		; Description:
		;
		;   Gets the formatting rectangle of the Edit control.
		;
		; Parameters:
		;
		;   r_Left..r_Bottom - Output variables. [Optional]
		;
		; Returns:
		;
		;   The address to a RECT structure that contains the formatting rectangle.
		;
		;-------------------------------------------------------------------------------
		Static EM_GETRECT:=0xB2, RECT
		VarSetCapacity(RECT,16,0)
		SendMessage EM_GETRECT,0,&RECT,, % "ahk_id " this.hEdit
		r_Left     	:=NumGet(RECT,0,"Int")
		r_Top    	:=NumGet(RECT,4,"Int")
		r_Right  	:=NumGet(RECT,8,"Int")
		r_Bottom	:=NumGet(RECT,12,"Int")
	Return &RECT
	}

	Edit_FindText(p_SearchText, p_Min=0, p_Max=-1, p_Flags="", ByRef r_RegExOut="")    {

	/*
	;-----------------------------
	;
	; Function: Edit_FindText
	;
	; Description:
	;
	;   Find text within the Edit control.
	;
	; Parameters:
	;
	;   p_SearchText - Search text.
	;
	;   p_Min, p_Max -  Zero-based search range within the Edit control.  p_Min is
	;       the character index of the first character in the range and p_Max is the
	;       character index immediately following the last character in the range.
	;       (Ex: To search the first 5 characters of the text, set p_Min to 0 and
	;       p_Max to 5)  Set p_Max to -1 to search to the end of the text.  To
	;       search backward, the roles and descriptions of the p_Min and p_Max are
	;       reversed. (Ex: To search the first 5 characters of the control in
	;       reverse, set p_Min to 5 and p_Max to 0)
	;
	;   p_Flags - Valid flags are as follows:
	;
	;       (Start code)
	;       Flag        Description
	;       ----        -----------
	;       MatchCase   Search is case sensitive.  This flag is ignored if the
	;                   "RegEx" flag is also defined.
	;
	;       RegEx       Regular expression search.
	;
	;       Static      [Advanced feature]
	;                   Text collected from the Edit control remains in memory is
	;                   used to satisfy the search request.  The text remains in
	;                   memory until the "Reset" flag is used or until the
	;                   "Static" flag is not used.
	;
	;                   Advantages: Search time is reduced 10 to 60 percent
	;                   (or more) depending on the size of the text in the control.
	;                   There is no speed increase on the first use of the "Static"
	;                   flag.
	;
	;                   Disadvantages: Any changes in the Edit control are not
	;                   reflected in the search.
	;
	;                   Hint: Don't use this flag unless performing multiple search
	;                   requests on a control that will not be modified while
	;                   searching.
	;
	;       Reset       [Advanced feature]
	;                   Clears the saved text created by the "Static" flag so that
	;                   the next use of the "Static" flag will get the text directly
	;                   from the Edit control.  To clear the saved memory without
	;                   performing a search, use the following syntax:
	;
	;                       Edit_FindText("","",0,0,"Reset")
	;       (end)
	;
	;
	;   r_RegExOutput - Variable that contains the part of the source text that
	;       matched the RegEx pattern. [Optional]
	;
	; Returns:
	;
	;   Zero-based character index of the first character of the match or -1 if no
	;   match is found.
	;
	; Calls To Other Functions:
	;
	; * <Edit_GetText>
	; * <Edit_GetTextLength>
	; * <Edit_GetTextRange>
	;
	; Programming Notes:
	;
	;   Searching using regular expressions (RegEx) can produce results that have a
	;   dynamic number of characters.  For this reason, searching for the "next"
	;   pattern (forward or backward) may produce different results from developer
	;   to developer depending on how the values of p_Min and p_Max are determined.
	;
	;-------------------------------------------------------------------------------
	 */
    Static s_Text

    ;-- Initialize
    r_RegExOut:=""
    if InStr(p_Flags,"Reset")
        s_Text:=""

	SciTEOutput("Hier bin ich")

    ;-- Anything to search?
	SciTEOutput("StrLen(p_SearchText): " StrLen(p_SearchText))
    if !StrLen(p_SearchText)
        Return -1

    If !(l_MaxLen := this.Edit_GetTextLength())
       Return -1

	SciTEOutput("maxlen: " p_maxlen)

    ;-- Parameters
    if (p_Min<0 || p_Max>l_MaxLen)
        p_Min:=l_MaxLen

    if (p_Max<0 || p_Max>l_MaxLen)
        p_Max:=l_MaxLen

    ;-- Anything to search?
    if (p_Min=p_Max)
        Return -1

    ;-- Get text
    if InStr(p_Flags,"Static")        {
        if !StrLen(s_Text)
            s_Text:=this.Edit_GetText()
        l_Text:=SubStr(s_Text,(p_Max>p_Min) ? p_Min+1:p_Max+1,(p_Max>p_Min) ? p_Max:p_Min)

	} else {
        s_Text:=""
        l_Text:=this.Edit_GetTextRange(this.hEdit,(p_Max>p_Min) ? p_Min:p_Max,(p_Max>p_Min) ? p_Max:p_Min)
        }

    ;-- Look for it
    if !InStr(p_Flags,"RegEx")  ;-- Not RegEx
        l_FoundPos:=InStr(l_Text,p_SearchText,InStr(p_Flags,"MatchCase"),(p_Max>p_Min) ? 1:0)-1
     else { ;-- RegEx
        p_SearchText:=RegExReplace(p_SearchText,"^P\)?","",1)   ;-- Remove P or P)
        if (p_Max>p_Min)  { ;-- Search forward
            l_FoundPos:=RegExMatch(l_Text,p_SearchText,r_RegExOut,1)-1
            if ErrorLevel  {
                outputdebug,
                   (ltrim join`s
                    Function: %A_ThisFunc% - RegExMatch error.
                    ErrorLevel=%ErrorLevel%
                   )
              l_FoundPos:=-1
              }
            }
         else {  ;-- Search backward
            ;-- Programming notes:
            ;
            ;    -  The first search begins from the user-defined minimum
            ;       position.  This will establish the true minimum position to
            ;       begin search calculations.  If nothing is found, no
            ;       additional searching is necessary.
            ;
            ;    -  The RE_MinPos, RE_MaxPos, and RE_StartPos variables contain
            ;       1-based values.
            ;
            RE_MinPos	:=1
            RE_MaxPos	:=StrLen(l_Text)
            RE_StartPos 	:=RE_MinPos
            Saved_FoundPos:=-1
            Saved_RegExOut:=""
            Loop         {
                ;-- Positional search.  Last found match (if any) wins
                l_FoundPos:=RegExMatch(l_Text,p_SearchText,r_RegExOut,RE_StartPos)-1
                if ErrorLevel    {
                    outputdebug,
                       (ltrim join`s
                        Function: %A_ThisFunc% - RegExMatch error.
                        ErrorLevel=%ErrorLevel%
                       )

                    l_FoundPos:=-1
                    Break
                    }

                ;-- If found, update saved and RE_MinPos, else update RE_MaxPos
                if (l_FoundPos>-1)                    {
                    Saved_FoundPos:=l_FoundPos
                    Saved_RegExOut:=r_RegExOut
                    RE_MinPos     :=l_FoundPos+2
				}  else
                    RE_MaxPos:=RE_StartPos-1

                ;-- Are we done?
                if (RE_MinPos>RE_MaxPos ||  RE_MinPos>StrLen(l_Text))                 {
                    l_FoundPos:=Saved_FoundPos
                    r_RegExOut:=Saved_RegExOut
                    Break
                    }

                ;-- Calculate new start position
                RE_StartPos:=RE_MinPos+Floor((RE_MaxPos-RE_MinPos)/2)
                }
            }
        }

    ;-- Adjust FoundPos
    if (l_FoundPos>-1) {
        l_FoundPos+=(p_Max>p_Min) ? p_Min:p_Max
		this.RegExOut := r_RegExOut
	}

    Return l_FoundPos
    }

	Edit_GetText(p_Length=-1)    {
	;-----------------------------
	;
	; Function: Edit_GetText
	;
	; Description:
	;
	;   Returns all text from the control up to p_Length length.  If p_Length=-1
	;   (the default), all text is returned.
	;
	; Calls To Other Functions:
	;
	; * <Edit_GetTextLength>
	;
	; Remarks:
	;
	;   This function is similar to the AutoHotkey *GUIControlGet* command (for AHK
	;   GUIs) and the *ControlGetText* command except that end-of-line (EOL)
	;   characters from the retrieved text are not automatically converted
	;   (CR+LF to LF).  If needed, use <Edit_Convert2Unix> to convert the text to
	;   the AutoHotkey text format.
	;
	;-------------------------------------------------------------------------------
		Static WM_GETTEXT:=0xD
		if (p_Length<0)
			p_Length:=this.Edit_GetTextLength(this.hEdit)

		VarSetCapacity(l_Text,p_Length*(A_IsUnicode ? 2:1)+1,0)
		SendMessage WM_GETTEXT,p_Length+1,&l_Text,, % "ahk_id " this.hEdit
		Return l_Text
		}

	Edit_GetTextLength()    {

		;-----------------------------
		;
		; Function: Edit_GetTextLength
		;
		; Description:
		;
		;   Returns the length, in characters, of the text in the Edit control.
		;
		;-------------------------------------------------------------------------------
		Static WM_GETTEXTLENGTH:=0xE
		SendMessage WM_GETTEXTLENGTH,0,0,, % "ahk_id " this.hEdit
	Return ErrorLevel
	}

	Edit_GetTextRange(p_Min=0,p_Max=-1)    {

	;-----------------------------
	;
	; Function: Edit_GetTextRange
	;
	; Description:
	;
	;   Get a range of characters.
	;
	; Parameters:
	;
	;   p_Min - Character position index immediately preceding the first character
	;       in the range.
	;
	;   p_Max - Character position immediately following the last character in the
	;       range.  Set to -1 to indicate the end of the text.
	;
	; Calls To Other Functions:
	;
	; * <Edit_GetText>
	;
	; Remarks:
	;
	;   Since the Edit control does not support the EM_GETTEXTRANGE message,
	;   <Edit_GetText> (WM_GETTEXT message) is used to collect the desired range of
	;   of characters.
	;
	;-------------------------------------------------------------------------------
    Return SubStr(this.Edit_GetText(p_Max),p_Min+1)
    }

	Edit_LineFromPos(X,Y,ByRef r_CharPos="",ByRef r_LineIdx="")    {
		;-----------------------------
		;
		; Function: Edit_LineFromPos
		;
		; Description:
		;
		;   This function is the same as <Edit_CharFromPos> except the line index
		;   (r_LineIdx) is returned.
		;
		;-------------------------------------------------------------------------------
		this.Edit_CharFromPos(X,Y,r_CharPos,r_LineIdx)
		Return r_LineIdx
	}

	Edit_CharFromPos(X,Y,ByRef r_CharPos="",ByRef r_LineIdx="")    {
	/*
	;
	; Function: Edit_CharFromPos
	;
	; Description:
	;
	;   Gets information about the character and/or line closest to a specified
	;   point in the the client area of the Edit control.
	;
	; Parameters:
	;
	;   X, Y - The coordinates of a point in the Edit control's client area
	;       relative to the upper-left corner of the client area.
	;
	;   r_CharPos - [Output] The zero-based index of the character nearest the
	;       specified point. [Optional] This index is relative to the beginning of
	;       the control, not the beginning of the line.  If the specified point is
	;       beyond the last character in the Edit control, the return value
	;       indicates the last character in the control.  See the *Remarks* section
	;       for more information.
	;
	;   r_LineIdx - [Output] Zero-based index of the line that contains the
	;       character. [Optional] For single-line Edit controls, this value is zero.
	;       The index indicates the line delimiter if the specified point is beyond
	;       the last visible character in a line.  See the *Remarks* section for
	;       more information.
	;
	; Returns:
	;
	;   The value of the r_CharPos variable.
	;
	; Calls To Other Functions:
	;
	; * <Edit_GetFirstVisibleLine>
	; * <Edit_LineIndex>
	;
	; Remarks:
	;
	;   If the specified point is outside the bounds of the Edit control, the
	;   return value and all output variables (r_CharPos and r_LineIdx) are set to
	;   -1.
	;
	 */
    Static Dummy3902

          ;-- Messages
          ,EM_CHARFROMPOS        	:=0xD7
          ,EM_GETFIRSTVISIBLELINE	:=0xCE
          ,EM_LINEINDEX              	:=0xBB

		;-- Collect character position from coordinates
		SendMessage EM_CHARFROMPOS, 0, (Y<<16)|X,, % "ahk_id " this.hEdit

		;-- Out of bounds?
		if (ErrorLevel<<32>>32=-1)			{
			r_CharPos	:= -1
			r_LineIdx	:= -1
			Return -1
		}

		;-- Extract values (UShort)
		r_CharPos:=ErrorLevel&0xFFFF    ;-- LOWORD
		r_LineIdx:=ErrorLevel>>16       ;-- HIWORD

		;-- Convert from UShort to UInt using known UInt values as reference
		SendMessage EM_GETFIRSTVISIBLELINE,0,0,, % "ahk_id " this.hEdit
		FirstLine:=ErrorLevel-1
		if (FirstLine>r_LineIdx)
			r_LineIdx:=r_LineIdx+(65536*Floor((FirstLine+(65535-r_LineIdx))/65536))

		SendMessage EM_LINEINDEX,(FirstLine<0) ? 0:FirstLine,0,, % "ahk_id " this.hEdit
		FirstCharPos:=ErrorLevel
		if (FirstCharPos>r_CharPos)
			r_CharPos:=r_CharPos+(65536*Floor((FirstCharPos+(65535-r_CharPos))/65536))

    Return r_CharPos
    }

	;~ Edit_LineIndex(p_LineIdx=-1)    {
			;~ ;-----------------------------
			;~ ;
			;~ ; Function: Edit_LineIndex
			;~ ;
			;~ ; Description:
			;~ ;
			;~ ;   Gets the character index of the first character of a specified line.
			;~ ;
			;~ ; Parameters:
			;~ ;
			;~ ;   p_LineIdx - Zero-based line number. [Optional] Use -1 (the default) for the
			;~ ;       current line.
			;~ ;
			;~ ; Returns:
			;~ ;
			;~ ;   The character index of the specified line or -1 if the specified line is
			;~ ;   greater than the total number of lines in the Edit control.
			;~ ;
			;~ ;-------------------------------------------------------------------------------
		;~ Static EM_LINEINDEX:=0xBB
		;~ SendMessage EM_LINEINDEX,p_LineIdx,0,,  % "ahk_id " this.hEdit
		;~ Return ErrorLevel<<32>>32  ;-- Convert UInt to Int
	;~ }

}


#Include %A_ScriptDir%\..\lib\class_cJSON.ahk
#Include %A_ScriptDir%\..\lib\SciTEOutput.ahk

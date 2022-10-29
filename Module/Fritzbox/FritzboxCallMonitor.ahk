; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                                               	☎	Addendum Fritzbox Anrufmonitor	☎
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;      Was unterscheidet diesen Anrufmonitor von anderen wie den CallMonitor oder dem JAnrufmonitor?
;    	Dieser Monitor ist zunächst einmal speziell für Albis, da er auf Daten aus Albis-Datenbanken zugreift um die  Namen der
;     	Anrufer anzuzeigen. Die eigentlich Funktion ist die Filterung der Anrufe. Es werden nicht einzelne Anrufe angezeigt, sondern
;		Anrufe werden gebündelt nach Anrufer dargestellt, so daß jederzeit klar mit wem bereits telefoniert wurde und mit wem nicht.
;		So wird man in der Hektik einer Sprechstunde, in der sowieso nie alle Anrufe entgegen genommen werden können, zumindestens
;		vermeintlich dringende Anrufe erkennen können (ohne das man einen Sprechstunden AB benötigt).
;		weitere mögliche Vorteile: 	- geänderte Telefonnummern von Patienten erkennen
;												-
;												- Statistik der Anrufzahlen, Anrufer und gesamten Gesprächsdauer (beeindruckt sogar Patienten!)
;
;
;		Abhängigkeiten:	siehe includes
;
;	                    			Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 22.09.2022 - this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;
;
;


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Skripteinstellungen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
  	#NoEnv
	#Persistent
	#InstallKeybdHook
	#KeyHistory      	, Off
	#SingleInstance	, Force

	SetBatchLines    	, -1
	SetControlDelay	, -1
	SetWinDelay     	, -1
	ListLines            	, Off

	FileEncoding, UTF-8
	DetectHiddenWindows, On
	CoordMode, ToolTip, Screen
	CoordMode, Mouse, Screen
	CoordMode, Menu, Screen


  ; class socket Variablen
	global DC_SERV            	:= True
	global DC_CLI                	:= False

  ; Addendum Variable
	global callersDB, callmon
	global Addendum
	RegExMatch(A_ScriptDir, "i)^(?<Dir>.*)(?=\\Module)", Addendum)
	Addendum := Object()
	Addendum.Ini              	:= AddendumDir "\Addendum.ini"
	Addendum.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	Addendum.compname	:= StrReplace(A_ComputerName, "-")                                           	; der Name des Computer auf dem das Skript läuft

  ; Albisverzeichnisse
	Albis                          	:= GetAlbisPaths()
	Addendum.AlbisDBPath	:= Albis.db
	Addendum.AlbisExe    	:= Albis.Exe
	Addendum.fboxini     	:= A_ScriptDir "\FritzboxCallMonitor.ini"
	Albis := AddendumDir := ""

	if (A_OSVersion >= "10.0.17763" && SubStr(A_OSVersion, 1, 3) = "10.") {
		attr := A_OSVersion >= "10.0.18985" ? 20 : 19
		DllCall("dwmapi\DwmSetWindowAttribute", "ptr", A_ScriptHwnd, "int", attr, "int*", true, "int", 4)
	}

	Menu, Tray, Icon, % A_ScriptDir "\assets\Fritzboxanrufmonitor.ico"
	Menu, Tray, NoStandard
	fn := func("callGui_Scriptvars")
	Menu, Tray, Add, Zeige Skript Variablen       	, % fn
	fn := func("callGui_Exit").Bind(true)
	Menu, Tray, Add, Skript Neu Starten             	, % fn
	fn := func("callGui_Exit").Bind(false)
	Menu, Tray, Add, Beenden                            	, % fn


  ;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Vorbereitungen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	If !FileExist(Addendum.fboxini) {

		Firstinitialisation := true
		ini := FileOpen(Addendum.fboxini , "w", "UTF-16")
		ini.WriteLine("; ############################################################")
		ini.WriteLine("; ")
		ini.WriteLine(";                                                  Einstellungen für den AHK Fritzbox Anrufmonitor")
		ini.WriteLine("; ")
		ini.WriteLine("; ############################################################")
		ini.WriteLine("[Allgemein]")
		ini.WriteLine("Version=0.2a")
		ini.Close()

		InputBox	, TelNumbers
						, Fritzbox Anrufmonitor Initialisierung
						, % 	"Hier eine Kommagetrennte Liste zu überwachender Telefonnummern eingeben.`n"
						  . 	"Vorwahlnummern weglassen! (z.B. 123331, 123356, 1237891)"
						, , 500, 120
		IniWrite, % RegExReplace(Trim(TelNumbers), "[\s,\|\-\+\_\\\/]", "|"), % Addendum.fboxini, Telefone, ueberwachen

		InputBox	, NoTel
						, Fritzbox Anrufmonitor Initialisierung
						, % "Hier eine Komma-getrennte Liste mit Telefonnummern aller erreichbarer Geräte eintragen`n"
						  .  "z.B. Tel1=123331, Tel2=123332, Fax=123333, AB1=123334, AB2=123335"
						, , 500, 120
		NoTel := StrReplace(NoTel, " ")

	}
	;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Einstellungen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
  ; zu protokollierende eigene Telefone
	If !TelNumbers
		IniRead, TelNumbers	, % Addendum.fboxini, Telefone, ueberwachen
	Addendum.TelNumbers := InStr(TelNumbers, "ERROR") ? "(99999999999)" : "(" TelNumbers ")"

  ; zu blockierende Anrufer
	If !Telblocked
		IniRead, Telblocked	, % Addendum.fboxini, Telefone, blockierte_Anrufer
	Addendum.Telblocked := InStr(Telblocked, "ERROR") ? "(99999999999)" : "(" Telblocked ")"

  ; Durchwahlnummern aller nicht Telefongeräte mit Bezeichnung
	Addendum.devices := Object()
	If !NoTel {
		IniRead, NoTel      	, % Addendum.fboxini, andere_Geraete
		NoTel := InStr(NoTel, "ERROR") ? "" : RegExReplace(NoTel, "[\n\r]+", ",")
	}
	For index, d in StrSplit(NoTel, ",", " ") {
		RegExMatch(d, "Oi)(?<name>[\pL\d\-\_\+]+)\s*=\s*(?<telnr>\d+)", device)
		IniWrite, % device.telnr, % Addendum.fboxini, andere_Geraete, % device.name
		Addendum.devices[device.telnr] := device.name
	}

  ; Vorwahlnummern Deutschland aufbereiten
	Addendum.prefixes := LoadCallPrefixes()

  ; mit shunt ist die interne Nummer der Telefoniegeräte gemeint.
  ; um die shunt Nummer z.B. eines Anrufbeantworters zu erhalten muss die Sicherungsdatei der
  ; Fritzbox Callmonitors Strings gesichtet werden. Das Skript kann nur anhand dieser Nummern unterscheiden
  ; an welchem Gerät ein Anruf entgegen genommen wurde.
	IniRead, val, % Addendum.fboxini, Shunts
	shunts := Object()
	For shuntIndex, shunt in StrSplit(val, "`n", "`r") {
		If RegExMatch(shunt, "(?<nr>\d+)\s*=\s*(?<name>[\pL\s\-\_]+)", shunt)
			shunts[shuntnr] := shuntname
	}
	If (shunts.Count() > 0) {
		Addendum.shunts := shunts
	}

	Fax := index := d := TelNumbers := ""
	val := shunts := shuntIndex := shunt := ""

  ;}

  ; CallersDB laden
	callersDB := new phonecallers(Addendum.AlbisDBPath)

  ; Gui zeigen
	callGui()

  ; Anrufmonitor starten
	callmon	:= new Fritzbox_Callmonitor("192.168.100.254", {"savefilepath":"C:\tmp\cm.txt", "managingfunc":"CallManager"})
	callmon.Connect()

  ; bei erfolgreicher Verbindung wird die Objektvariable .connected auf wahr gesetzt
	callmon._OnReceive("savepath: " callmon.savepath)

  ; Anrufe eines Tages anzeigen. leer lassen für den aktuellen Tag oder eine Datum der Form YYYYMMDD (Jahr2stelligerMonat2stelligerTag)
	callGui_Load("")

  ; Neustart  um 0 Uhr
	func_call := Func("callGui_Exit").Bind(true)
	SetTimer, % func_call, % -1*TimerTime("00:00 Uhr")

return


!Esc::
callmon.Disconnect()
ExitApp
^+a::
	callGui_Exit(true)
return

class Fritzbox_Callmonitor                                                                         	{	;-- Verbindungsmanager zur Fritzbox

	/* 	Fritzbox Callmonitor

			- 	a class for establishing a connection to the Fritzbox call monitor and for handling the call-data output by the Fritzbox
         	- 	to handle this call-data, an internal string managing function or your own callback function can be used
			-	for the simplest connection, only the IP address of the Fritzbox is needed.
				very simple as in this example: callmon := new Fritzbox_CallMonitor("192.168.100.254")

	*/

	__New(tcpip, options)                                                                	{	;

		; check ip adress
			If !RegExMatch(tcpip, "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")
				throw A_ThisFunc ": this is no valid ip4-adress <" tcpip ">"

		; store ip adress
			this.ip   	:= tcpip
			this.port 	:= 1012

		; custom call manager function
			managingfunc := options.managingfunc
			If (StrLen(managingfunc) > 0 && IsFunc(managingfunc))
				this.managingfunc := Func(managingfunc)
			else if IsObject(managingfunc)
				this.managingfunc := managingfunc
			else
				this.managingfunc := Func("CallManager")

		; save received strings
			If options.savefilepath {
				savefilepath := options.savefilepath
				SplitPath, savefilePath, fname, fpath
				If !InStr(FileExist(fPath), "D")
					throw A_ThisFunc ": this is no valid path <<" fpath ">>"
				this.savefilepath	:= savefilepath
				this.savepath      	:= fpath
			}
			else
				this.savepath := A_Temp

		; open new TCP socket
			this.fbox := new SocketTCP()

	}

	__Call(method, args*)                                                                	{
        if (method = "")              	; For %fn%() or fn.()
            return this.Call(args*)
        if IsObject(method)         	; If this function object is being used as a method.
            return this.Call(method, args*)
    }

	__Delete()                                                                                    	{

	}

	_OnReceive(rec:="")                                                                   	{ 	;-- forwards received fritzbox string to dispatcher func

		static init, managingfunc

		If !init {
			init := true
			If !IsObject(this.managingfunc)
				throw ".managingfunc: must be a BoundFunc object!"
			managingfunc := this.managingfunc
		}

		Critical

		If IsObject(rec) {
			recv	:= rec.recvText()
			this.stringfile := FileOpen(this.savepath "\" A_YYYY A_MM A_DD "_FbxCMStrings.txt", "a", "UTF-8")
			this.stringfile.WriteLine(recv := RegExReplace(recv, "[\n\r]+"))
			this.stringfile.Close()
			call := this.CallStringParser(recv)
			%managingfunc%(call)
		}

	}

	Connect()                                                                                   	{ 	;-- connect to fritzbox callmonitor

		If (this.connected := this.fbox.Connect(this.ip, this.port)) {
			this.fbox.onRecv := ObjBindMethod(this, "_OnReceive")
			this._OnReceive("connection established to Fritzbox with ip: " this.ip ":" this.port)
		}

	return this.connected
	}

	Disconnect()                                                                              	{	;-- disconnect from callmonitor
		this.connected := this.fbox.Disconnect() ? 0 : 1
		If IsObject(this.savefile) {
			flength := this.savefile.Length
			res := this.savefile.Close()
			this.savefile := flength
		}
	return this.connected
	}

	CallStringParser(str)                                                                      	{	;-- converts the string to an ahk object

		RegExMatch(str, "Oi)(?<date>\d+\.\d+\.\d+)\s+(?<time>\d+:\d+:\d+);(?<event>\w+);(?<conID>\d+);", callh)
		call := {"date":callh.date, "time":callh.time, "event":callh.event, "conID":callh.conID}
		o := StrSplit(str, ";")
		Switch call.event {
			Case "CALL":
				call.shuntnr          	:= o.4
				call.from           	:= o.5
				call.to               	:= o.6
			Case "RING":
				call.from           	:= o.4
				call.to                	:= o.5
			Case "CONNECT":
				call.shuntnr          	:= o.4
				call.from           	:= o.5
			Case "DISCONNECT":
				call.duration        	:= o.4
		}

	return call
	}

	SaveCall(call)                                                                             	{	;-- save call data as json string

		If !this.savefilepath
			return

		If IsObject(call)
			For key, val in call
				t .= key ":" val "`t"
		else {
			t := RegExReplace(call, "[\n\r]+", "`t")
			t := RegExReplace(t	 , "[\s]+" 	, "")
		}

		this.savefile := FileOpen(this.savefilepath, "a", "UTF-8")
		this.savefile.WriteLine(t)     ; JSON.Dump(call, "", "", "UTF-8")
		this.savefile.Close()

	}

	LoadCalls(datestr:="")                                                                 	{	;-- loads messages from this day

		datestr := !datestr ? A_YYYY A_MM A_DD : datestr
		If FileExist(fullfilepath := this.savepath "\" datestr "_FbxCMStrings.txt")
			calls := StrSplit(FileOpen(fullfilepath, "r", "UTF-8").Read(), "`n", "`r")

	return calls
	}

}

class phonecallers                                                                                     	{	;-- lädt Telefonnummer-Daten aus verschiedenen Albis-Datenbank-Dateien

	__New(DBPath:="")                                     	{

	  ; Rückwärtssuche: 	"https://www.dasoertliche.de/?form_name=search_inv&ph=" telnumber
	  ; 							"http://www.google.de/search?num=0&q=%s" "$number"

	  ; Albis DBASE Datenbankenpfad
		this.AlbisDBPath := DBPath

	  ; Daten aus der PATIENT.dbf und PATTELNR.dbf zusammenführen
		PatTelNrs       	:= this.LoadPATTELNR()       	; PATTELNR.dbf PATNR TELNRNORM
		this.PatientsDB 	:= this.LoadPatients()          	;
		For PatID, PatData in this.PatientsDB
			If IsObject(PatTelNrs[PatID])
				this.PatientsDB[PatID].TELNR := PatTelNrs[PatID]

		clipboard := cJSON.Dump(this.PatientsDB, 1)

	  ; Daten aus der Addendum FritzFaxbox Telefonnummern Datei
		this.FaxSender 	:= this.LoadFaxSender()

		;this.OthersDB	:= this.LoadOthersDB()         ; UEBArzt TEL TEL2 Fax
		;this.ExtrasDB	:= this.LoadExtrasDB()

	}

	LoadPatients()                                            	{      	; lädt Patienten-Daten

		infilter		:= ["NR", "GESCHL", "NAME", "VORNAME", "GEBURT", "MORTAL"]
		outfilter		:= ["GESCHL", "NAME", "VORNAME", "GEBURT", "MORTAL"]

	  ; Daten auslesen
		tmpDB 	 	:= ReadPatientDBF(this.AlbisDBPath, infilter, "EMail=0 allData")   ; , "allData"

		PatientDB	:= Object()
		For PatID, Pat in tmpDB {
			PatientDB[PatID] := Object()
			For index, key in outfilter
				PatientDB[PatID][key] := tmpDB[PatID][key]
		}
		tmpDB := ""

	return PatientDB
	}

	LoadPATTELNR()                                        	{       	; lädt nur Telefonnummern

		matches	:= Object()
		dbf       	:= new DBASE(this.AlbisDBPath "\PATTELNR.dbf")
		res        	:= dbf.OpenDBF()

		Loop % dbf.records {
			obj    	:= dbf.ReadRecord(["PATNR", "TELNRKLAR"])
			TelNr	:= RegExReplace(obj.TelNrKLAR, "[^\d+]")
			TelNr	:= RegExReplace(TelNr, "^\+49")
			If (StrLen(TelNr) > 4)                    ; unter 5 Ziffern ist keine Telefonnummer
				If !IsObject(matches[obj.PATNR])
					matches[obj.PATNR] := [TelNr]
				else
					matches[obj.PATNR].Push(TelNr)
		}

		res         	:= dbf.CloseDBF()
		dbf        	:= ""

	return matches
	}

	LoadFaxSender()                                        	{      	; lädt Faxnummern aus anderer Datei

		IniRead, path, Addendum.Ini, ScanPool, FritzFaxbox_Telefonbuch
  		If FileExist(path) {
			this.FaxSender := Object()
			temp := FileOpen(path, "r", "UTF-8").Read()
			For each, line in StrSplit(temp, "`n", "`r")
				If (StrSplit(line, "|").1 ~= "i)\(.*Fax.*?\)") {
					sdr	:=    	Trim(StrSplit(line, "|").1)	                        	; Absender (Sender)
					nr 	:= "#" 	Trim(StrSplit(line, "|").2)	                        	; Telefonnummer
					ftr 	:=    	Trim(StrSplit(line, "|").3)	                        	; Filter
					this.FaxSender[nr] := {"sdr":sdr}
					If ftr
						this.FaxSender[nr].ftr := ftr
				}
		}

	}

	GetPrefixLacation(prefix)                            	{        	; Anzeige der zugeörigen Orte zur Vorwahl

		static prefixesObj

		If !IsObject(prefixesObj) && FileExist(fVorwahl := A_ScriptDir "\deutschland_vorwahl.json")
			prefixesObj :=  cJSON.Load(FileOpen(fVorwahl, "r", "UTF-8").Read())

		For pfxIndex, item in prefixesObj
			If (prefix = item.number)
				return item.Name

	return
	}

	GetNameFromNumber(nr, mergenames=1)	{       	; Anruferidentifikation

	; matches Struktur := Array([PatID (0 für nicht Patienten)
	;	    								, Telefon-/Faxnummer
	;										, Patienten-/Firmenname
	;                                     	, Datenquelle (Integer)])

		matches := Array()
		nr := RegExReplace(nr, "[^\d]")

	 ; Telefonnummern aus PATTELNR.dbf und PATIENT.dbf
		For PatID, Pat in this.PatientsDB {
			For TelIndex, TelNr in Pat.TELNR       	; Patiententelefonnummern heraussuchen
				If RegExMatch(nr, "^" TelNr "$") {   	 	; (TelNr && )
					If (!matches.Count() || !mergenames)
						matches.Push([PatID, TelNr, (PAT.VORNAME " " Pat.NAME), 0x1])
					else
						For pIndex, pData in matches
							If InStr(pData.3, " " Pat.NAME) {
								If !Instr(matches[pIndex].3, "&") {                                                    	; verhindert >2 Vornamen (Klaus & Bernd & Renate Müller)
									matches[pIndex].3 := PAT.VORNAME " & " matches[pIndex].3
									matches[pIndex].1 .= "|" PatID
								}
								break
							}
				}
		}

	 ; Telefonnummern aus _DB\Telefon\Telefonnummern.txt (Addendum)
		For callNr, obj in this.FaxSender {
			If RegExMatch(nr, "^" callNr "$")
				If !matches.Count()
					matches.Push([0, callNr, obj.sdr, 0x2])
				else {
					For each, item in matches
						item[4] := item[4] + 0x2
				}
		}

	; Telefonnummern aus Addendum.fboxini
		IniRead, callername, % Addendum.fboxini, Telefonbuch, % "T" nr
		If (callername := InStr(callername, "ERROR") ? "" : callername) {

			If InStr(callername, ",", A_Space )
				callername := StrSplit(callername, ",").2 " " StrSplit(callername, ",").1

			If !matches.Count()
				matches.Push([0, nr, callername, 0x4])
			else {
				For each, item in matches
					item[4] := item[4] + 0x4
			}

		}



	return matches
	}

}

callManager(call, load:=false)                                                                   	{	;-- behandelt die Verbindungsnachrichten der Fritzbox

	global 	Callers, Day, hCV, phones, uphones, ophones, ringCount, callsCount, talktime
	global	BLV, hLV1, hLV2, LVColors
	static 	callstack   	:= Array()
	static 	callerstack	:= Object()
	static 	tables   	:= Object()
	static 	cMinit    	:= 0
	static 	teltime 		:= 0
	static 	phone_prefix := {"de": {"(0151\d|0160|017[015])"             	: "Telekom"
												, "(0152\d|0162|017[234])"              	: "Vodaphone"
												, "(0157\d|0159\d|0163|017[6789])"	: "O2"
												, "015566"                                        	: "1&1 Drillisch"
												, "015888"                                           	: "TelcoVillage"}}

	If !cMinit {
		cMinit   	:= true
		phones 	:= Object()
	    uphones 	:= Object()
		ophones	:= Object()
		talktime  	:= ringCount := callsCount := 0
	}

  ; AHK arrays starts with an index of 1
	conID := call.conID + 1

  ; removes the connectionID-key/value
	call.Delete("conID")

 ; execution depending on event
	Switch call.event {

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; RING
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "RING":

			call.Delete("event")
			call.blocked := false
			call.ignore := false

		  ; in Albisdatenbank suchen
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			callers := callersDB.GetNameFromNumber(call.from)
			for idx, caller in callers {
				t 	.= caller.3 "|"
				p 	.= caller.1 "|"
			}
			t := RTrim(t, "|"), p := RTrim(p, "|")

		 ; in Callmonitor Datenbank suchen (ini Datei)
		 ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If !t && call.from {
				IniRead, callername, % Addendum.fboxini, Telefonbuch, % call.from
				t := InStr(callername, "ERROR") || !callername ? "" : callername
			}

		; Anrufe an bestimmte Nummern ignorieren (z.B. Fax)
		; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			IgnoreCall := !RegExMatch(call.to, Addendum.TelNumbers) ? true : false    	; Anruf nicht zählen, da Anruf z.B. auf Fax
			device := Addendum.devices[call.to]                                                           	; Telefongerät feststellen

	    ; blockierte Telefonnummern werden nicht angezeigt
	    ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If RegExMatch(call.from, Addendum.Telblocked)
				call.ignore := IgnoreCall := true, call.blocked := true

		 ; unbekannte Nummer aufnehmen und zählen
		 ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If !t && call.from && !IgnoreCall && !call.blocked {         ;
				If !IsObject(uphones[call.from])
					uphones[call.from] := [uphones.Count()+1, 1]
				else
					uphones[call.from].2 += 1
			}

		  ; Anrufe zu ignorierten Geräten werden ebenso gezählt
		 ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If call.ignore && !call.blocked      ; ignoriert aber nicht blockierte Nummer
				ophones[device] := !ophones.haskey(device) ? 1 : ophones[device]+1


			call.caller                 	:= t	? t	: call.from ? "unbekannter Anrufer " (uphones[call.from].1) : "Nummer unterdrückt"
			call.callerID              	:= p	? p	: call.from
			call.state                  	:= "🔔"
			cID                         	:= "#" call.callerID

			If !call.ignore {
				ringCount ++
				If call.from
					phones[call.from] := !phones.haskey(call.from) ? 1 : phones[call.from]+1
			}

			;📞🔔🔕🆎
			Gui, CV: Default
			If !IsObject(callerstack[cID]) && !call.ignore {

				Gui, CV: ListView, Callers
				LVRow := !LV_GetCount() ? 1 : LV_GetCount()+1
				callerstack[cID] := [1, LVRow, 0]                        	; Anrufe , Zeile, Anrufannahme Mensch=1 AB=2 getrennt/wartend=0

				Gui, CV: ListView, Callers
				LV_Add(""	, call.caller
								, nicenumber(call.from)
								, "((🔔))"
								, 1
								, call.time)

				BLV["Callers"].Row(LVRow, 0x99B898, 0x0)
				SendMessage, 0x102A, % LVRow-1,,, % "ahk_id " hLV1

			}
			else if IsObject(callerstack[cID]) && !call.ignore {

				callerstack[cID].1 += 1

				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "((🔔))")
				LV_Modify(callerstack[cID].2, "Col4", callerstack[cID].1)
				LV_Modify(callerstack[cID].2, "Col5", call.time)

			}
			else {                                                             	; nicht überwachte Telefonnummern werden gleich unter angenommene Anrufe einsortiert

				If !call.blocked {

					Gui, CV: ListView, Day
					LV_Add(""	, call.caller
									, nicenumber(call.from)
									, RegExMatch(call.to, Addendum.Fax) ? "((🔔)) ℻" : call.from
									,
									, 1
									, call.time)

					Gui, CV: ListView, Day
					BLV["Day"].Row((DayRow := !LV_GetCount() ? 1 : LV_GetCount()), 0x699366, 0xDDDDDD)
					SendMessage, 0x102A, % DayRow-1,,, % "ahk_id " hLV2        ; Listviewzeile updaten

				}

			}

			call.calls                     	:= callerstack[cID].1
			callstack[conID]         	:= call
			;callstack[conID].time  	:= call.time
			;callstack[conID].ignore	:= IgnoreCall
			callstack[conID].Day  	:= DayRow


	;}

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; CALL
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "CALL":

			call.Delete("event")
			callers := callersDB.GetNameFromNumber(call.to)
			for idx, caller in callers {
				t .= caller.3 "|"    	; Name
				p .= caller.1 "|"   	; PatID
			}
			t := RTrim(t, "|"), p := RTrim(p, "|")

			If !t {
				IniRead, callername, % Addendum.fboxini, Telefonbuch, % call.to
				t := InStr(callername, "ERROR") ? "" : callername
			}

			call.caller                 	:= t ? t : "unbekannter Anrufer " (unknown ++)
			call.callerID              	:= p ? p : call.to
			cID                         	:= "#" call.callerID

			callsCount ++
			phones[call.to] := !phones.haskey(call.to) ? 1 : phones[call.to]+1

			Gui, CV: Default
			If IsObject(callerstack[cID]) {
				call.state := "🆎"
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "🆎")
				LV_Modify(callerstack[cID].2, "Col5", call.time)
				callerstack[cID].3 := 2 	; zurück gerufen

			}
			else {
				call.state := "☎"
				Gui, CV: ListView, Callers
				LVRows := !LV_GetCount() ? 1 : LV_GetCount()+1
				callerstack[cID] := [1, LVRows]
				Gui, CV: ListView, Callers
				LV_Add("", call.caller, call.to, "☎", 1, call.time)
				callerstack[cID].3 := 3   ; Anruf ohne vorherigen Telefonanruf

			}

			call.calls                  	:= callerstack[cID].1
			callstack[conID]        	:= call
			callstack[conID].time := call.time

	;}

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; CONNECT
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "CONNECT":

			cID := "#" callstack[conID].callerID
			callstack[conID].shuntnr := call.shuntnr
			callstack[conID].time := call.time

			Gui, CV: Default
			If (call.shuntnr <> 40) && !callstack[conID].ignore {   ; 40 ist der AB
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "📞")
				LV_Modify(callerstack[cID].2, "Col5", call.time)
				callerstack[cID].3 := callerstack[cID].3 ? callerstack[cID].3 : 1 	; Anruf wurde angenommen
			}
			else if !callstack[conID].ignore{
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "🆎")
				LV_Modify(callerstack[cID].2, "Col5", call.time)
				callerstack[cID].4 := 1	; AB wurde abspielt  📤
			}
			else {
				callstack[conID].connect := 1
				Gui, CV: ListView, Day
				LV_Modify(callstack[conID].Day, "Col3", RegExMatch(callstack[conID].to, Addendum.Fax) ? "📥℻" : "")
				LV_Modify(callstack[conID].Day, "Col6", call.time)
			}

			callstack[conID].time := call.time
	;}

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; DISCONNECT
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "DISCONNECT":

			If !IsObject(callstack[conID])
				return

			cID := "#" callstack[conID].callerID
			callstack[conID].duration := call.duration ? call.duration : ""

		  ; Anruf wurde entgegengenommen, Anrufe mit unterdrückter Telefonnummer werden ebenso in die 2.Liste verschoben
			If callstack[conID].ignore {
				Gui, CV: ListView, Day
				LV_Modify(callstack[conID].Day, "Col3", (callstack[conID].connect ? "" : "✗") (RegExMatch(callstack[conID].to, Addendum.Fax) ? "℻" : ""))
				LV_Modify(callstack[conID].Day, "Col6", call.time)
				callstack[conID] := "-"
			}
			else {

				constr		:= 	callstack[conID]
				callNr 		:= 	callerstack[cID].3<2 ? constr.from : constr.to

				If (callerstack[cID].3>0 || InStr(constr.caller, "unbekannt")) && (constr.shuntnr<>40 && constr.duration > 0) {

					talktime += constr.duration

					Gui, CV: Default
					Gui, CV: ListView, Callers
					LV_Delete(callerstack[cID].2)

				; noch vorhandene Einträge aus vorherigen Anrufen entfernen
					Loop % LV_GetCount() {
						LV_GetText(LVTelNr	, A_Index, 2)
						LV_GetText(LVCalls	, A_Index, 4)
						LV_GetText(LVTime	, A_Index, 5)
						LVTelnr := plainnumber(LVTelNr)
						If (callNr = LVTelNr) {
							constr.calls += LVCalls
							LV_Delete(A_Index)
						}
					}

					Gui, CV: ListView, Day
					LV_Add("", constr.caller, nicenumber(callNr), constr.state, duration(constr.duration), constr.calls, constr.time)

					If !load
						callmon.SaveCall(callstack[conID])

					constr := ""
					callstack[conID] := "-"
					callerstack.Delete(cID)

				}
				else {

					Gui, CV: Default
					Gui, CV: ListView, Callers
					LV_Modify(callerstack[cID].2, "Col3", "✗" (callerstack[cID].4 ? "🆎":"🔕"))

				}
		}
	;}


	}

	callGui_SBSetText()

}


; Gui
callGui()                                                                                                      	{

	;{ Variablen
	global Callers, Day, hCV, BLV := Object(), hLV1, hLV2, hCVT1, LVColors, ontop, LVNames, CVDate
	global TThwnd1, TThwnd2
	global LVNames 	:= ["Callers", "Day"]
	global colSizes  	:= {"Callers":[230,140,50,"60 Center",125], "Day":[230,140,50,60,"55 Center",70]}

  ; Fenstergröße
	IniRead, winSize, % A_ScriptDir "\FritzboxCallMonitor.ini", % "sonstiges", % "Fensterposition"
	winSize	:=  InStr(winSize, "ERROR") ? "" : winSize
	RegExMatch(winSize, "x\s*(?<X>\-*\d+)", win)
	RegExMatch(winSize, "y\s*(?<Y>\-*\d+)", win)
	winSize	:= (winX<0 ? "" : "x" winX) (winY<0 ? "" : " y" winY " " )

  ; andere Fensterparameter
	wplus      	:= 25
	LVWidth   	:= Object()
	LVOpt		:= "NoSortHdr AltSubmit gcallGui_LVHandler hwndhLV"
	LVColors	:= ["0x99B898", "0x355C7D"]
	ontop    	:= true

  ; Listviewbreite
	For LVIndex, LVName in LVNames {
		LVWidth[LVName] := 0
		For colNr, colOpt in colSizes[LVName] {
			RegExMatch(colOpt, "^\d+", width)
			LVWidth[LVName] += width
		}
	}
	;}

  ; Gui
	Gui, CV: new		, % "hwndhCV " ( ontop ? "AlwaysOnTop" : "")
	Gui, CV: Color    	, % "c" "355C7D" , % "c" "99B898"
	Gui, CV: Margin	, 0 , 5

	Gui, CV: Font    	, s8 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "xm ym   	w100  cWhite  vCVDate"                                                                    	, % "DD, 00.00.0000"

	LVNamesW := colSizes.Callers.2 + colSizes.Callers.3
	Gui, CV: Font    	, s10 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "x" colSizes.Callers.1 " ym w" LVNamesW " Left cWhite hwndhCVT1"                         	, aktuelle Anrufe
	opt := " hwndCVhOnTop vCVOnTop gcallGui_Handler"
	Gui, CV: Add     	, Button  	, % "x" LVWidth["Callers"]+wplus-22 " y2 w20 h20 " opt, % (ontop ? "🔐" : "🔓")

	Gui, CV: Font    	, s10 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, ListView	, % "xm y+0 w" LVWidth["Callers"]+wplus 	" r15 " LVOpt "1  vCallers"                  	, Anrufer|Telefon|Status|Anrufe|Uhrzeit

	Gui, CV: Font    	, s10 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "x" colSizes.Callers.1 " y+2 w" LVNamesW " Left cWhite"                                     	, angenommene Anrufe
	Gui, CV: Font    	, s10 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, ListView	, % "xm y+0 w" LVWidth["Day"]+wplus    	" r20 " LVOpt "2 vDay"                            	, Anrufer|Telefon|Status|Dauer|Anrufe|Uhrzeit

  ; Because it looks nice
	WinSet, ExStyle, 0x0, % "ahk_id " hLV1
	WinSet, ExStyle, 0x0, % "ahk_id " hLV2
	SetExplorerTheme(hLV1)
	SetExplorerTheme(hLV2)

	Gui, CV: Show		, Hide
	Gui, CV: ListView	, Day

	Lv := GetWindowSpot(hLV2)
	rows := LV_GetCountPerPage(hLV2)
	Lv.H := 20 * rows

  ; Statusbar
	Gui, CV: Font    	, s8 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, StatusBar, % "hwndhSbar"
	DllCall("uxtheme\SetWindowTheme", "ptr", hSbar, "str", "DarkMode_Explorer", "ptr", 0)

  ; die Höhe der Listview's soll ein exaktes Vielfaches der Zeilenhöhe sein
	ControlMove,,,,, Lv.H, % "ahk_id " hLV2
	;~ ControlMove, msctls_statusbar321,, Lv.Y+Lv.H-16,,, % "ahk_id " hCVr

	Gui, CV: Show    	, % winSize:= (winX<0 ? "x1" : "x" winX) (winY<0 ? " y1" : " y" winY " " ) " NA", Fritz!Monitor

	ControlClick,, % "ahk_id " hLV1

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Breite anpassen
	For LVIndex, LVName in LVNames {
		Gui, CV: ListView, % LVName
		For colNr, colOpt in colSizes[LVName]
			LV_ModifyCol(colNr, colOpt)
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Farben
	BLV["Callers"] := new LV_Colors(hLV1)
	BLV["Callers"].Critical := 200
	BLV["Callers"].SelectionColors("0x0078D6")

	BLV["Day"] := new LV_Colors(hLV2)
	BLV["Day"].Critical := 200
	BLV["Day"].SelectionColors("0x0078D6")

  ; WM_MOUSEMOVE handler
	OnMessage(0x0200, "callGui_WM_MOUSEMOVE")

return
CVGuiClose:
CVGuiEscape:
	WinMinimize, % "ahk_id " hCV
return
}

callGui_Load(datestr:="")                                                                           	{	;-- lädt Telefondaten aus gespeicherten Fritzbox-Dateien

	DayOfWeek := DayOfWeek(dayDate := (!dateStr ? A_YYYY A_MM  A_DD : datestr), "short", "yyyyMMdd")
	GuiControl, CV:, CVDate, % " " DayOfWeek ", " SubStr(dayDate, 7, 2) "." SubStr(dayDate, 5, 2) "." SubStr(dayDate, 1, 4)

	calls := callmon.LoadCalls(datestr)

	GuiControl, CV: -Redraw, Callers
	GuiControl, CV: -Redraw, Day

	For callNr, callString in calls
		If callString {
			call := callmon.CallStringParser(callString)
			Gui, CV: Default
			callManager(call, true)
		}

	GuiControl, CV: +Redraw, Callers
	GuiControl, CV: +Redraw, Day


}

callGui_Exit(onlyReload:=true)                                                                      	{

	global hCV

	wqs := GetWindowSpot(hCV)
	If (wqs.X>-8 && wqs.Y>-8)
		IniWrite, % "x" wqs.X " y" wqs.Y, % A_ScriptDir "\FritzboxCallMonitor.ini", % "sonstiges", % "Fensterposition"

	If !FileExist(statspath := A_ScriptDir "\telefon_statistik.ini")
		IniWrite, % "Anzahl Anrufe;versch.Telefonnummern; davon Unbekannt;ausgehende Anrufe;Gesprächszeit" , % statspath, _Datenstruktur

	IniWrite, % callGui_SBSetText(), % statspath, % A_YYYY "-" A_MM, % SubStr("0" A_DD, -1)

	If onlyReload
		Reload

ExitApp
}

callGui_Handler(hwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")           	{	;-- g-Label

	global ontop, hCV, CVhOntop, CVOnTop, hLV1, hCVT1

	If (A_GuiControl="CVOnTop") {
		ontop := !ontop
		WinSet, AlwaysOnTop, % (ontop ? "on" : "off"), % "ahk_id " hCV
		GuiControl, CV:, CVOntop, % (ontop ? "🔐" : "🔓")
		ControlClick,, % "ahk_id " hLV1
		otp := GetWindowSpot(CVhOntop)
		ToolTip, % "AllwaysOnTop ist`n<" (ontop ? "eingeschaltet":"ausgeschaltet") ">", otp.X-80, otp.Y+otp.H+5, 3
		hotip := WinExist("A")
		otp := GetWindowSpot(hotip)
		SetTimer, ontopTimer, -2000
	}


return
ontopTimer:
	ToolTip,,,, 3
return
}

callGui_LVHandler(hwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")         	{	;-- g-Label Listviewhandler

	global CV, Callers, Day, ClickedLVRow, ClickedLV, BLV, hLV1, hLV2, LVColors
	global LVHandlerProc

	Critical

	If (StrLen(A_GuiControl)=0) || (StrLen(EventInfo)=0) || !EventInfo
		return

	Gui, CV: Default
	Gui, CV: ListView, % (clickedLV := A_GuiControl)
	LV_GetText(clkLVCaller	, (clickedLVRow := A_EventInfo), 1)
	LV_GetText(clkLVTelNr	, clickedLVRow, 2)
	MouseGetPos, mx, my

  ; Doppelklick
	If (A_GuiEvent = "DoubleClick") {
		clk := callGui_RenameCheck(clkLVCaller, clkLVTelNr)
		If !clk.RenameIsLocked
			callGui_Rename(EventInfo, A_GuiControl)
	}

  ; Kontextmenu
	else if (A_GuiEvent = "RightClick") {

	  ; Menu wird immer neu erstellt
		try
			Menu, cmenu, DeleteAll

	  ; Funktionsobjekte des Kontextmenu
		funcCM1 	:= Func("callGui_CM").Bind("copy"      	, clickedLV, clickedLVRow)
		funcCM2	:= Func("callGui_CM").Bind("plaincopy"	, clickedLV, clickedLVRow)
		funcCM3 	:= Func("callGui_CM").Bind("remove" 	, clickedLV, clickedLVRow)
		funcCM4	:= Func("callGui_CM").Bind("rename"	 	, clickedLV, clickedLVRow)

	  ; Namensänderungen sind nur möglich, wenn die Daten aus einer FritzboxCallMonitor oder Addendum-Datei stammen
		clk := callGui_RenameCheck(clkLVCaller, clkLVTelNr)
		;~ SciTEOutput(A_ThisFunc ": " clk.RenameIsLocked ", " clk.PatID ", " clickedLV ", " clkLVCaller ", " clkLVTelNr)

		If (clickedLV = "Callers")
			Menu, cmenu, Add, von Anruferliste nehmen       	, % funcCM3
		If !clk.RenameIsLocked
			Menu, cmenu, Add, Anrufernamen ändern          	, % funcCM4
		If clk.PatID {
			funcCM5	:= Func("callGui_CM").Bind("opencase" 	, clickedLV, clickedLVRow, clk.PatID)
			Menu, cmenu, Add, Karteikarte öffnen                	, % funcCM5
		}

		Menu, cmenu, Add
		Menu, cmenu, Add, Telefonnr. kopieren                 	, % funcCM1
		Menu, cmenu, Add, Telefonnr. (nur Zahlen) kopieren	, % funcCM2

		Menu, cmenu, Show, % mx-20, % my

	}

  ; einfacher Klick
	else if (A_GuiEvent = "Normal") {
		If !LVHandlerProc
			callGui_MarkCallers(clickedLV, clickedLVRow)
	}



}

callGui_RenameCheck(Caller, TelNr)                                                        		{ 	;-- prüft ob eine Telefonnummer geändert werden darf

	global callersDB

	 RenameIsLocked := false
   If !RegExMatch(Caller, "i)unbekannter.*Anruf") {
	   RenameIsLocked := true
		matches := callersDB.GetNameFromNumber(RegExReplace(TelNr, "[^\d]"), false)
		;~ SciTEOutput(cJSON.Dump(matches,1))
		For each, item in matches {
		clkLVPatID := item.1
			If (item.4 > 1) {
				RenameIsLocked := false
				break
			}
		}
	}

return {"RenameIsLocked": RenameIsLocked, "PatID":clkLVPatID}
}

callGui_MarkCallers(LVclicked, rowclicked:=0)                                           	{

	global CV, hCV, Callers, Day, BLV, hLV1, hLV2, LVColors
	global LVHandlerProc
	static init := false

	If !init {
		init := true
		LVHandlerProc := false
	}

	LVHandlerProc := true
	altLV := LVclicked = "Callers" ? "Day" : "Callers"

	Gui, CV: Default
	Gui, CV: ListView, % LVclicked

	LV_GetText(clkLVCaller	, rowclicked, 1)
	LV_GetText(clkLVTelNr	, rowclicked, 2)
	LV_GetText(clkLVState	, rowclicked, 3)
	LV_GetText(clkLVTime	, rowclicked, 5)
	clkLVTelNr := RegExReplace(clkLVTelNr, "[^\d]")
	clkLVTime := ClockToSeconds(clkLVTime)

	Loop % LV_GetCount() {
		BLV[LVclicked].Row(A_Index, LVColors.1, 0x000000)
		SendMessage, 0x102A, % A_Index-1,,, % "ahk_id " (LVclicked = "Callers" ? hLV1 : hLV2)
	}

	Gui, CV: Default
	Gui, CV: ListView, % altLV

	BLV[altLV].Clear()

	Loop % LV_GetCount() {

		LV_GetText(altLVCaller	, A_Index, 1)
		LV_GetText(altLVTelNr	, A_Index, 2)
		LV_GetText(altLVState	, A_Index, 3)
		LV_GetText(altLVTime	, A_index, 6)
		altLVTelNr	:=  RegExReplace(altLVTelNr, "[^\d]")
		altLVTime	:= ClockToSeconds(altLVTime)
		teltel      	:= altLVTelNr 	= clkLVTelNr 	? true : false
		callmatch	:= clkLVCaller 	= altLVCaller 	? true : false

		LV_Modify(A_Index, "-Select -Focus")
		If !InStr(altLVState, "📇") {
			BLV[altLV].Row(A_Index, (teltel ? 0xD77800 : LVColors.1), (teltel ? 0xFFFFFF : 0x000000))
			SendMessage, 0x102A, % A_Index-1,,, % "ahk_id " (altLV = "Callers" ? hLV1 : hLV2)
		}

	}

	LVHandlerProc := false

}

callGui_CM(cmd, LVclicked:="", rowclicked:=0, clkPatID:="")                        	{

	Gui, CV: Default
	Gui, CV: ListView, % LVclicked
	MouseGetPos, mx, my

	If InStr(cmd, "copy") {

		LV_GetText(TelNr, rowclicked, 2)
		Clipboard := TelNr := cmd = "plaincopy" ? RegExReplace(TelNr, "[^\d]") : TelNr
		ClipWait, 2
		ToolTip, % TelNr " kopiert", % mx-20, % my-20, 1
		SetTimer, TTipOff, -3000

	}
	else If (cmd = "rename") {

		callGui_Rename(rowclicked, LVclicked)

	}
	else if  InStr(cmd, "remove") && (LVclicked="Callers") {

		LV_GetText(Caller	, rowclicked, 1)
		LV_GetText(TelNr	, rowclicked, 2)
		LV_GetText(state	, rowclicked, 3)
		LV_GetText(calls	, rowclicked, 4)
		LV_GetText(timestr	, rowclicked, 5)
		fn := Func("MsgBoxMove").Bind(mx, my, "Frage", "Anrufer wirklich")
		SetTimer, % fn, -50
		MsgBox, 0x1004,  Frage, Anrufer wirklich von der Liste nehmen?
		IfMsgBox, No
			return
		callerstimestr := StrSplit(timestr, ":").1*3600 + StrSplit(timestr, ":").2*60 StrSplit(timestr, ":").3
		LV_Delete(rowclicked)

		Gui, CV: ListView, % "Day"
		itemmoved := itembesttime := 0
		Loop % LV_GetCount() {
			LV_GetText(DayTelNr 	, A_Index, 2)
			LV_GetText(Daycalls  	, A_Index, 5)
			LV_GetText(daytimestr	, A_Index, 6)
			If (TelNr = plainnumber(DayTelNr)) {
				daytimeStr := StrSplit(daytimeStr, ":").1*3600 + StrSplit(daytimeStr, ":").2*60 StrSplit(daytimeStr, ":").3
				itembesttime := !itembesttime && (daytimestr > callerstimestr) ? A_Index : itembesttime
				If (callerstimestr <= daytimestr) {
					LV_Modify(A_Index, "Col5", DayCalls+calls)
					itemmoved := 1
					break
				}
			}
		}

		;~ SciTEOutput("itemmoved: " itemmoved ", " itembesttime)
		itembesttime := !itembesttime ? LV_GetCount() : itembesttime
		If !itemmoved
			LV_Insert(itembesttime,, Caller, TelNr, state,, calls, timestr)

	}

return

TTipOff:
	ToolTip,,,, 1
return
}

callGui_SBSetText(txt:="", sbpos:=1)                                                             	{

	global ringCount, phones, uphones, callsCount, talktime
	txt := !txt ? (ringCount " Anrufe von " phones.Count() " Telefonnummern. "
						. uphones.Count()	" Anrufer unbekannt. "
						. callsCount          	" ausgehende Anrufe. "
						. "Gesprächszeit: " (talktime>59 ? "rund ":" ") duration(talktime, 2)) : txt
	Gui, CV: Default
	SB_SetText(txt , sbpos)

return ringCount ";" phones.Count() ";" uphones.Count() ";" callsCount ";" talktime
}

callGui_Rename(LVRow, LVName)                                                              	{

	global uphones, LVNames

	MouseGetPos, mx, my

	Gui, CV: Default
	Gui, CV: ListView, % LVName

	LV_GetText(uNAME	, LVRow, 1)
	LV_GetText(uTEL    	, LVRow, 2)

	InputBox	, callername, Fritzbox Anrufmonitor, % "Sie ändern: " uName " der Telefonnummer: " uTEL
					,, 300, 140,,,,, % (!RegExMatch(uName, "i)unbekannt.*Anruf") ? uName : "")

	If (RegExMatch(callername, "^\s*$") || RegExMatch(callername, "i)unbekannt.*Anruf") || callername = uName || ErrorLevel)
		return

  ; geänderten Namen speichern
	IniWrite, % callername, % Addendum.fboxini, Telefonbuch, % (telnr := RegExReplace(uTEL, "[^\d]"))
	uphones.Delete(TelNr)
	callGui_SBSetText()

  ; geänderten Namen in beiden Listviews anzeigen
	Gui, CV: Default
	Loop 2 {
		Gui, CV: ListView, % LVNames[A_Index]
		Loop % LV_GetCount() {
			LV_GetText(rowTel, A_Index, 2)
			rowTel := RegExReplace(rowTel, "[^\d]")
			If (rowTel = uTEL)
				LV_Modify(A_Index, 1, callername)
		}
	}

}

callGui_WM_MOUSEMOVE(wParam, lParam, Msg, Hwnd)                          	{
   ; LVM_HITTEST   -> docs.microsoft.com/en-us/windows/desktop/Controls/lvm-hittest
   ; LVHITTESTINFO -> docs.microsoft.com/en-us/windows/desktop/api/Commctrl/ns-commctrl-taglvhittestinfo
   global colSizes, hLV1, hLV2, hCV
   static 	TelefonX1, TelefonX2, LVy1, LVy2, Xinit := true

   CoordMode, ToolTip, Screen
   If Xinit {
		TelefonX1 := colSizes.Callers.1
		TelefonX2 := TelefonX1 + colSizes.Callers.2
		Xinit := false
   }

	Critical
   If (A_GuiControl ~= "i)(Callers|Day)") {
      VarSetCapacity(LVHTI, 24, 0) ; LVHITTESTINFO
      , NumPut(lParam & 0xFFFF, LVHTI, 0, "Int")
      , NumPut((lParam >> 16) & 0xFFFF, LVHTI, 4, "Int")
      , Item := DllCall("SendMessage", "Ptr", Hwnd, "UInt", 0x1012, "Ptr", 0, "Ptr", &LVHTI, "Int") ; LVM_HITTEST
      If (Item >= 0) && (NumGet(LVHTI, 8, "UInt") & 0x0E) { ; LVHT_ONITEM
		;Y := lParam >> 16
		Gui, ListView, % A_GuiControl
		LV_GetText(TT, row := Item + 1, 2)
		RegExMatch(TT, "\d+(?=\/)", prefix)
      }
   }

   If prefix {
		X := lParam & 0xFFFF
		If row && (X >= TelefonX1 && X <= TelefonX2) {
			CV 	:= GetWindowSpot(hCV), LV := GetWindowSpot(Hwnd)
			cellPos := LV_EX_GettemRect(Hwnd, 2, row, 3)
			LVy 	:= LV.Y + cellPos.Y - cellPos.H
			ToolTip, % phonecallers.GetPrefixLacation(prefix), % CV.X+TelefonX1, % LVy, 1
			SetTimer, WM_TT_Off, -500
		}
	}
return
WM_TT_Off:
	ToolTip,,,,1
return
}

callGui_ScriptVars()                                                                                    	{

	ListVars

}

; Einstellungen
LoadCallPrefixes() {

  ; Vorwahlnummern Deutschland aufbereiten
	If FileExist(fVorwahl := A_ScriptDir "\deutschland_vorwahl.json") {
		tmp := cJSON.Load(FileOpen(fVorwahl, "r", "UTF-8").Read())
		For index, prefix in tmp
			prefixes .= (index>1 ? "|" : "") prefix.number
		return prefixes
	}

return
}


;{ RPA Funktionen für Albis
AlbisAkteOeffnen(CaseTitle="", PatID="") {                                                   	;-- öffnet eine Patientenakte über Name, ID oder Geburtsdatum

	; Variablen                                                                                        	;{
		global Mdi, hMdi
		static WarteFunc, AlbisTitleFirst
		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

		If !AlbisWinID()
			return 0

		name    	:= []
		CaseTitle	:= Trim(CaseTitle)
		sStr := PatID := Trim(PatID)
	;}

	; RegExString und anderen Suchstring erstellen                                     	;{
		If RegExMatch(PatID, "^\d+$") {

			If IsObject(cPat) {
				PName := cPat.NAME(PatID)                        ; zusammengesetzter Name aus Nachname, Vorname
				GD   	:= cPat.GEBURT(PatID, true)
			}
			else If IsObject(PatDB) {
				PName	:= PatDB[PatID].Name ", " PatDB[PatID].Vorname
				GD    	:= PatDB[PatID].Geburt
			}

			rxStr      	:= "\[" PatID "\s\/"
			sStr       	:= PatID

		}
		else if RegExMatch(CaseTitle, "(?<surname>[\pL\-\s]+)*[\s,]*(?<prename>[\pL\-\s]+)*[\s,]*(?<Birth>\d{1,2}\.\d{1,2}\.(\d{4}|\d{2}))*", case) {

			caseBirth 	:= FormatDateEx(caseBirth, "DMY", "dd.MM.yyyy")
			rxStr      	:= caseName ? RegExReplace(CaseName, "\,\s+", ",\s+") : caseBirth
			sStr       	:= (caseName) (CaseName && caseBirth ? ", " : "") (caseBirth)

		}
		else if (!CaseTitle && !PatID)
			return 0

	;}

	; Karteikarte bereits geöffnet?                                                             	;{
		hMdi	 := AlbisMDIClientHandle()
		oMDI := AlbisMDIClientWindows()
		For MDITitle, thisMdi in oMDI
			If RegExMatch(MDITitle, "i)" rxStr) {
				SendMessage, 0x222, % thisMdi.ID,,, % "ahk_id " hMdi
				return 1
			}
	;}

	; Öffne Patient Dialog aufrufen                                                           	;{
		AlbisTitleFirst := WinGetTitle(AlbisWinID())               	; aktuellen Albisfenstertitel auslesen
		hPOeffnen 	:= AlbisDialogOeffnePatient()               	; Aufruf des Fenster 'Patient öffnen'
	;}

	; Übergeben des Parameter an das Albisdialogfenster                         	;{
		If !VerifiedSetText("Edit1", sStr, Win_PatientOeffnen, 200)
			If (hPOeffnen && !VerifiedSetText("Edit1", sStr, hPOeffnen, 200)) {
				VerifiedClick("Button3", Win_PatientOeffnen)
				return 0
			}

		while WinExist(Win_PatientOeffnen)		{

			If (A_Index > 1)
				sleep 100
			If (A_Index = 1)
				VerifiedClick("Button2", Win_PatientOeffnen)            	; Versuch 1: Button OK drücken
			else If (A_Index = 2)
				ControlSend, Edit1, {Enter}, % Win_PatientOeffnen	; Versuch 2: Enter simulieren
			else {
				WinClose, % Win_PatientOeffnen
				return 0
			}

		}

	;}

	; Loop der in den nächsten 10 Sekunden auf die neue Karteikarte wartet	;{
		Loop	{

			If !(hwnd:= DLLCall("GetLastActivePopup", "uint", AlbisWinID()))
				hwnd := AlbisWinID()
			newTitle	:= WinGetTitle(hwnd)
			newClass	:= WinGetClass(hwnd)
			newText	:= WinGetText(hwnd)

		; Karteikarte ist geöffnet
			If RegExMatch(newTitle, "i)" rxStr)
				return 1

		; Dialog Patient <.....,......> nicht vorhanden
			If Instr(newTitle, "ALBIS") && Instr(newClass, "#32770") && Instr(NewText, "nicht vorhanden") {
				VerifiedClick("Button1", "ALBIS ahk_class #32770", "nicht vorhanden")			;Abbrechen
				if WinExist(Win_PatientOeffnen)
					VerifiedClick("Button3", Win_PatientOeffnen)
				return 0
			}
			else if Instr(newTitle, "Patient") && Instr(newClass, "#32770") && Instr(newText, "List1") {
				return 2
			}

			If (A_Index > 50) 	{
				if WinExist(Win_PatientOeffnen)
					VerifiedClick("Button3", Win_PatientOeffnen)
				return 0
			}

			sleep, 200
		}
	;}

return 1
}

AlbisDialogOeffnePatient(command:="invoke", pattern:="" ) {                       	;-- startet Dialog zum Öffnen einer Patientenakte

	; more commands are here:
	; 	abort/close - um das Fenster zu schliessen ohne eine Suche zu durchzuführen
	;	serach/set/open [Namens-/Suchmuster]- übernimmt den eingegeben Text und kann gleichzeitig das Suchmuster eintragen
	; letzte Änderung: 30.06.2022

		static Win_PatientOeffnen := "Patient öffnen ahk_class #32770"

	; Aufrufen des Dialogfensters
		If !InStr(command, "close") {
			hwnd 	:= Albismenu(32768, "Patient öffnen", 3)
			wT 		:= WinGetTitle(hwnd)
			If !InStr(wT, "Patient öffnen") {
				WinWait, % Win_PatientOeffnen,, 1
				If !WinExist(Win_PatientOeffnen) {
					return 0
				}
				hwnd := WinExist(Win_PatientOeffnen)
			}
		}

	; command parsen
		If InStr(command, "invoke")
			return hwnd
		else If InStr(command, "close") {

			If WinExist(Win_PatientOeffnen)
				return VerifiedClick("Button3", Win_PatientOeffnen)
			else
				return

		} else If InStr(command, "open") || InStr(command, "set")  {

			; kein Suchmuster übergeben dann wird nur auf Ok gedrückt
				If (StrLen(Pattern) > 0)
					If !VerifiedSetText("Edit1", Pattern, Win_PatientOeffnen, 200)
						return 0

			; Suchmuster wurde als Parameter übergeben, aber
				If InStr(command, "set")
					return hwnd

				; Akte wird jetzt geöffnet durch drücken von OK
					while WinExist(Win_PatientOeffnen)					{
							; Button OK drücken
								VerifiedClick("Button2", Win_PatientOeffnen)
								WinWaitClose, % Win_PatientOeffnen,, 1
							; Fenster ist immer noch da? Dann sende ein ENTER.
								if WinExist(Win_PatientOeffnen)								{
									WinActivate, % Win_PatientOeffnen
									ControlFocus, Edit1, % Win_PatientOeffnen
									SendInput, {Enter}
								}

								If (A_Index > 10)
										return
								sleep, 100
					}

		}


return hwnd
}

Albismenu(mcmd, FTitel:="", wait:=2, methode:=1) {                                	;-- Aufrufen eines Menupunktes oder Toolbarkommandos

	; letzte Änderung: 05.05.2022

		InfoMsg := false

		If IsObject(FTitel) {
			WinTitle	:= FTitel.1
			WinText 	:= FTitel.2
			AltTitle  	:= FTitel.3
			AltText   	:= FTitel.4
		}
		else
			WinTitle	:= FTitel

	; prüft ob eines der erwarteten Fenster bereits geöffnet ist
		If (hwin := WinExist(WinTitle, WinText)) || (IsObject(FTitle) && (hAltwin := WinExist(AltTitle, AltText)))
			return hwin ? GetHex(hwin) : GetHex(hAltwin)

	; Menuaufruf, wenn kein Fenstername übergeben wurde
		If !WinTitle {

			If RegExMatch(methode, "i)(1|Post)") {
				PostMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				return 1
			}
			else { ; ACHTUNG: Sendmessage wartet auf eine Antwort, diese kann lange dauern (z.T. bis zu 5s)
				SendMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
				return ErrorLevel
			}

		}

	; Menuaufruf, bei Übergabe eines oder zweier Fensternamen
		If RegExMatch(methode, "i)(1|Post)")
			PostMessage	, 0x111, % mcmd,,, % "ahk_class OptoAppClass"
		else
			SendMessage, 0x111, % mcmd,,, % "ahk_class OptoAppClass"

	; Fenster abwarten
		maxRounds := wait*1000/20
		while (A_Index <= maxRounds) {
			sleep 20
			hWin 	:= WinExist(WinTitle	, WinText)
			hAltwin	:= WinExist(AltTitle	, AltText)
			If hWin || hAltwin
				return hwin ? GetHex(hwin) : GetHex(hAltwin)
		}

	hwin 	:= WinExist(WinTitle	, WinText)
	hAltwin	:= WinExist(AltTitle	, AltText)

return hWin ? GetHex(hwin) : hAltwin ? GetHex(hAltwin) : 0
}

AlbisMDIClientHandle() {                                                                             	;-- ermittelt das Handle des MDIClient (Basishandle für alle Unterfenster)
	; letzte Änderung: 10.02.2021
	ControlGet, hMdi, HWND,, MdiClient1, ahk_class OptoAppClass
return hMdi
}

AlbisMDIClientWindows() {                                                                         	;-- Namen, Klassen und Handles aller geöffneten MDI Fenster

	; Mdi	- [key : value]	- Object > global Mdi - im aufrufenden Skript / aufrufender Funktion einfügen!
	;         	- key             	- ist WinTitle
	;       	- value          	- ist die Fensterklasse, extra Key ist "MdiHwnd" für das Roothandle des MDI-Fenster
	;
	; letzte Änderung: 10.02.2021

	global Mdi
	Mdi := Object()   ; das Mdi-Objekt wird nur von dieser Funktion zurückgesetzt

	WinGet, MDIClientWinList, ControlListHWND, % "ahk_id " AlbisMDIClientHandle()
	Loop, Parse, MDIClientWinList, `n
		If InStr(class := WinGetClass(A_LoopField), "Afx:")
			Mdi[WinGetTitle(A_LoopField)]:= {"class": class, "ID": A_LoopField}

return Mdi
}

AlbisMDIChildHandle(MDITitle) {		                                                          	;-- ermittelt das Handle eines sub oder child Fenster innerhalb des Albis-MDI-Controls
return GetHex(FindChildWindow({"Class": "OptoAppClass"}, {"Title": MDITitle}, "Off"))
}

AlbisWinID() {                            			                                                    	;-- gibt die ID des übergeordneten Albisfenster zurück

	; letzte Änderung: 07.02.2021
	Loop
		If (AID := WinExist("ahk_class OptoAppClass"))
			return GetHex(AID)
		else If (A_Index > 40)
			return 0
		else if (A_Index <= 40)
			sleep, 20

return GetHex(AID)
}
;}

; Hilfsfunktionen

;{ anything else
MsgBoxMove(mx, my, wtitle, wtxt) {

	hwnd := WinExist(wtitle " ahk_class #32770", wtxt)
	mbx := GetWindowSpot(hwnd)
	SetWindowPos(hwnd, mx-mbx.W/2, my+10, mbx.W, mbx.H)

}

clocktoseconds(Time) {
return StrSplit(Time, ":").1*3600 + StrSplit(Time, ":").2*60 + StrSplit(Time, ":").3
}

nicenumber(telnr, formatter:="/") {
	static prefixes
	If !prefixes
		prefixes := "^(0)(151\d|160|17[015]|152\d|162|17[234]|157\d|159\d|163|17[6789]|15566|15888|" Addendum.prefixes ")"
return RegExReplace(telnr, prefixes, "$1$2" formatter)
}

plainnumber(telnr) {
return RegExReplace(telnr, "[^\d]")
}

duration(sec, Mode:=1) {

	min 	:= SubStr("00" Floor(sec/60), -1)
	sec 	:= SubStr("00" sec - min*60, -1)

return Mode=1 ? (min ":" sec) : (min+0>0 ? min " min." : sec "s.")
}

SetExplorerTheme(HCTL) { ; HCTL : handle of a ListView or TreeView control
   If (DllCall("GetVersion", "UChar") > 5) {
      VarSetCapacity(ClassName, 1024, 0)
      If DllCall("GetClassName", "Ptr", HCTL, "Str", ClassName, "Int", 512, "Int")
         If (ClassName = "SysListView32") || (ClassName = "SysTreeView32")
            Return !DllCall("UxTheme.dll\SetWindowTheme", "Ptr", HCTL, "WStr", "Explorer", "Ptr", 0)
   }
   Return False
}

;}

;{ Controls
ControlGetText(Control="", WTitle="", WText="", ExTitle="", ExText="") {          	;-- ControlGetText als Funktion
	ControlGetText, v, % Control, % WTitle, % WText, % ExTitle, % ExText
Return v
}

GetFocusedControlHwnd(hwnd:="A") {
	ControlGetFocus, FocusedControl	 , % (hwnd = "A") ? "A" : "ahk_id " hwnd
	ControlGet    	 , FocusedControlId, Hwnd,, % FocusedControl, % (hwnd = "A") ? "A" : "ahk_id " hwnd
return GetHex(FocusedControlId)
}

GetFocusedControlClassNN(hwnd:="A") {
	ControlGetFocus, FocusedControl	 , % (hwnd = "A") ? "A" : "ahk_id " hwnd
	ControlGet    	 , FocusedControlId, Hwnd,, % FocusedControl, % "ahk_id " hwnd
return Control_GetClassNN(hwnd, FocusedControlId)
}

Control_GetClassNN(hWnd, hCtrl) {
	; SKAN: www.autohotkey.com/forum/viewtopic.php?t=49471
 WinGet, CH, ControlListHwnd	, % "ahk_id " hWnd
 WinGet, CN, ControlList       	, % "ahk_id " hWnd
 LF:= "`n",  CH:= LF CH LF, CN:= LF CN LF,  S:= SubStr(CH,1,InStr(CH,LF hCtrl LF))
 S:= StrReplace(S,"`n","`n", RplCount)
 LP:= InStr(CN, "`n",,, RplCount) - 1
 Return SubStr(CN,LP+2,InStr(CN,LF,0,LP+2)-LP-2)
}

GetClassNN(Chwnd, Whwnd) {
	global _GetClassNN := {}
	_GetClassNN.Hwnd := Chwnd
	Detect := A_DetectHiddenWindows
	WinGetClass, Class, ahk_id %Chwnd%
	_GetClassNN.Class := Class
	DetectHiddenWindows, On
	EnumAddress := RegisterCallback("GetClassNN_EnumChildProc")
	DllCall("EnumChildWindows", "uint",Whwnd, "uint",EnumAddress)
	DetectHiddenWindows, %Detect%
	return, _GetClassNN.ClassNN, _GetClassNN:=""
}

GetControls(hwnd, class_filter:="", type_filter:="", info_filter:="") {                                     	  	;-- returns an array with ClassNN, ButtonTyp, Position.....

	;class_filter 	- comma separated list of classes        	you don't want to store
	;type_filter 	- comma separated list of control types 	you don't want to store
	;info_filter 	- comma separated list of classes       	you !want! to store

		If StrLen(info_filter) = 0
			info_filter:="hwnd,Pos,Enabled,Visible,Style,ExStyle"

		controls := Array()
		Control_Style           	:= "Style"
		Control_IsEnabled 	:= "Enabled"
		Control_IsVisible    	:= "Visible"
		Control_ExStyle      	:= "ExStyle"
		Control_Pos          	:= "Pos"
		Control_Handle      	:= "hwnd"

		WinGet, classnnList  	, ControlList        	, % "ahk_id " hwnd
		WinGet, controlIdList	, controllisthwnd	, % "ahk_id " hwnd

		For idx, classnn in StrSplit(classnnList, "`n")
			controls.Push({"classNN" : classnn})

		For idx, hwnd in StrSplit(controlIdList, "`n") {

			; class is dismissed
				RegExMatch(controls[idx].classNN, "i)[A-Z]+", class)
				If class in %class_filter%
					continue

			; informations
				hWin := "ahk_id " hwnd
				If RegExMatch(class, "Button") {

					bTyp:= GetButtonType(hwnd)
					if bTyp in %type_filter%
						continue

					controls[idx]["type"]:= bTyp
					If RegExMatch(bTyp, "(Radio|Checkbox)")
						controls[idx]["checked"]	:= ControlGet("Checked", "", "", hWin)
					else
						controls[idx]["text"]      	:= ControlGetText("", hWin)

				}
				else if RegExMatch(class, "(Edit|RichEdit)")	{
					controls[idx]["text"]       	:= ControlGetText("", hWin)
					controls[idx]["linecount"]	:= ControlGet("LineCount", "", "", hWin)
				}

				controls[idx]["text"]      	:= ControlGetText("", hWin)

			; filter informations
				If Control_Handle  	in %info_filter%
					controls[idx]["hwnd"]   	:= hwnd

				If Control_IsEnabled 	in %info_filter%
					controls[idx]["Enabled"]	:= ControlGet("Enabled", "", "", hWin)

				If Control_IsVisible 	in %info_filter%
					controls[idx]["Visible"]  	:= ControlGet("Visible", "", "", hWin)

				If Control_Style      	in %info_filter%
					controls[idx]["Style"]    	:= ControlGet("Style", "", "", hWin)

				If Control_ExStyle   	in %info_filter%
					controls[idx]["Exstyle"]  	:= ControlGet("ExStyle", "", "", hWin)

				If Control_Pos        	in %info_filter%
				{
					ControlGetPos, cx, cy, cw, ch,, % "ahk_id " hwnd
					controls[idx]["Pos"]        	:= cx "," cy "," cw "," ch
				}
		}

return controls
}

GetButtonType(hwndButton) {                                                                                                	;-- ermittelt welcher Art ein Button ist, liest dazu den Buttonstyle aus
	;Link: https://autohotkey.com/board/topic/101341-getting-type-of-control/
  static types := [ "Button"            	;BS_PUSHBUTTON
                     	, "Button"            	;BS_DEFPUSHBUTTON
                     	, "Checkbox"      	;BS_CHECKBOX
                     	, "Checkbox"      	;BS_AUTOCHECKBOX
                     	, "Radio"             	;BS_RADIOBUTTON
                     	, "Checkbox"      	;BS_3STATE
                     	, "Checkbox"      	;BS_AUTO3STATE
                     	, "Groupbox"      	;BS_GROUPBOX
                     	, "NotUsed"       	;BS_USERBUTTON
                     	, "Radio"             	;BS_AUTORADIOBUTTON
                     	, "Button"            	;BS_PUSHBOX
                     	, "AppSpecific"   	;BS_OWNERDRAW
                     	, "SplitButton"       	;BS_SPLITBUTTON    (vista+)
                     	, "SplitButton"       	;BS_DEFSPLITBUTTON (vista+)
                     	, "CommandLink"	;BS_COMMANDLINK    (vista+)
                     	, "CommandLink"]	;BS_DEFCOMMANDLINK (vista+)

  WinGet, btnStyle, Style, % "ahk_id " hwndButton
 return types[1+(btnStyle & 0xF)]
}

ControlGet(Cmd,Value="",Control="",WTitle="",WTxt="",ExTitle="",ExText="") {  	;-- ControlGet als Funktion
	ControlGet, v, % Cmd, % Value, % Control, % WTitle, % WTxt, % ExTitle, % ExText
Return v
}

VerifiedClick(CName, WTitle="", WText="", WinID="", WaitClose=false) {       	;-- 4 verschiedene Methoden um auf ein Control zu klicken

	; WaitClose eine Zahl größer 0 für die maximale Zeit die gewartet werden darf
	; eventuell vorhandene Kürzelzeichen <&> im Buttonnamen werden entfernt damit ein Vergleich Treffer erzielt
	; letzte Änderung: 23.09.2021

		CoordMode, Mouse, Screen
		EL := 0, CName := RegExReplace(CName, "[\&]", "")

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		if WinID
			WText := "", WTitle := "ahk_id " WinID
		else if RegExMatch(WTitle, "i)^(0x[A-F\d]+|[\d]+)$")
			WText := "", WTitle := "ahk_id " WTitle

	; 3 verschiedene Wege einen Buttonklick auszulösen
		ControlClick, % CName, % WTitle, % WText,,, NA
		If (EL := ErrorLevel) {                                                                            ; Misserfolg = 1 , Erfolg = 0
			ControlClick, % CName, % WTitle, % WText
			If (EL := ErrorLevel) {
               	SendMessage, 0x0201, 1, 0, % CName, % WTitle, % WText                 	;0x0201 - WM_Click
				EL := ErrorLevel = "FAIL" ? 1 : 0
				If EL {
					BlockInput, On                                                                       	; funktioniert nur mit Systemrechten
					WinGetPos    	, wx, wy,,, % WTitle, % WText
                   	ControlGetPos	, cx, cy, cw, ch, % CName, % WTitle, % WText
                   	MouseGetPos	, mx, my
                   	MouseClick   	, Left, % wx + cx + Floor(cw/2), % wy + cy + Floor(ch/2), 1, 0
                   	MouseMove  	, % mx, % my, 0
					BlockInput, Off
                    EL := 0
				}
			}
		}

		If WaitClose {
			WinWaitClose, % WTitle, % WText, % WaitClose
			EL := ErrorLevel                                                                                ; Zeitlimit überschritten = 1, sonst 0
		}

return (EL = 0 ? 1 : 0)
}

VerifiedSetFocus(CName, WTitle="", WText="", WinID="", activate=false) {       	;-- setzt den Eingabefokus und überprüft das dieser auch gesetzt wurde

	; Rückgabeparameter: 	erfolgreich - 	das Handle des Controls, ansonsten 0 (false)
	; letzte Änderung: 26.09.2021

		static tms

	; hwnd (WinID) des Fensters ermitteln
		If WText {
			tms := A_TitleMatchModeSpeed
			SetTitleMatchMode, Slow
		}
		ControlID	:= WinID && !CName ? WinID : ""
		WinID    	:= WinID ? WinID : RegExMatch(WTitle, "i)^(0x[A-F\d]+|\d+)$") ? GetHex(WinTitle) : WTitle ? WinExist(WTitle, WText)
		WTitle    	:= ("ahk_id " WinID)
		WText   	:= ""

	; Fenster aktivieren nach Bedarf
		if activate {
			WinActivate	 , % WTitle
			WinWaitActive, % WTitle,,  1
		}

	; Focus setzen und überprüfen
		If CName {
			while (!InStr(GetFocusedControlClassNN(WinID), CName) && A_Index < 21) {
               	wIndex := A_Index
              	ControlFocus, % CName, % WTitle
				focusEL := ErrorLevel
				If (A_Index > 1)
					sleep 70
			}
		}
		else {
			while (GetFocusedControlHwnd() <> ControlID && A_Index < 21) {
               	wIndex := A_Index
              	ControlFocus, % CName, % WTitle
				focusEL := ErrorLevel
				If (A_Index > 1)
					sleep 70
			}
		}

	; Titlematchmodespeed zurücksetzen
		If tms
			SetTitleMatchMode, % tms
		tms := ""

		;~ SciTEOutput(A_ThisFunc ": EL=" focusEL ", wI=" wIndex )

return wIndex >= 20 ? GetFocusedControlHwnd() : false
}

VerifiedSetText(CName="", NewText="", WTitle="", delay=100, WText="") {    	;-- erweiterte ControlSetText Funktion

	; kontrolliert ob der Text tatsächlich eingefügt wurde und man kann noch eine Verzögerung übergeben
	; delay = Zeit in ms zwischen den Versuchen
		abb	 := delay > 2000 ? 20 : Floor(2000//delay)	; maximal 2 Sekunden wird versucht Text in das Steuerelement einzutragen
		delay -= 40

	; leeren des Fenster-Titel und Textes wenn ein Handle übergeben wurde
		If RegExMatch(WTitle, "^0x[\w]+$")
			WTitle	:= RegExMatch(WTitle, "^0x[\w]+$")	? "ahk_id " WTitle	: WTitle
		else if RegExMatch(WTitle, "^\d+$", digits)
			WTitle	:= StrLen(WTitle) = StrLen(digits)     	? "ahk_id " digits 	: WTitle
		else
			WTitle	:= "ahk_id " (WinID := GetHex(WinExist(WTitle, WText)))

	; WText wird nicht mehr benötigt
		WText := ""

	; eintragen
		while (ControlGetText(CName, WTitle) <> NewText) 	{

			If (A_Index >= abb)
				  return 0
			else if (A_Index > 1)
				sleep % delay

			ControlSetText, % CName, % NewText, % WTitle
			If !ErrorLevel    	; schläft ein extra Ründchen
				sleep 40
			else {
				If VerifiedSetFocus(CName, WTitle) {
					ControlSendRaw, % CName, % NewText, % WinTitle
					;~ SciTEOutput(A_ThisFunc ": ControlSendRaw !" )
					Sleep 70
				}
			}

		}

return ControlGetText(CName, WTitle) = NewText ? true : false
}

;}

;{ Windows
GetWindowSpot(hWnd) {                                             	;-- gets window position
    NumPut(VarSetCapacity(WININFO, 60, 0), WININFO)
    DllCall("GetWindowInfo", "Ptr", hWnd, "Ptr", &WININFO)
    wi := Object()
    wi.X    	:= NumGet(WININFO, 4	, "Int")
    wi.Y    	:= NumGet(WININFO, 8	, "Int")
    wi.W   	:= NumGet(WININFO, 12, "Int") 	- wi.X
    wi.H    	:= NumGet(WININFO, 16, "Int") 	- wi.Y
    wi.CX  	:= NumGet(WININFO, 20, "Int")
    wi.CY  	:= NumGet(WININFO, 24, "Int")
    wi.CW	:= NumGet(WININFO, 28, "Int") 	- wi.CX
    wi.CH  	:= NumGet(WININFO, 32, "Int") 	- wi.CY
	wi.S    	:= NumGet(WININFO, 36, "UInt")
    wi.ES   	:= NumGet(WININFO, 40, "UInt")
	wi.AC  	:= NumGet(WININFO, 44, "UInt")
    wi.BW 	:= NumGet(WININFO, 48, "UInt")
    wi.BH  	:= NumGet(WININFO, 52, "UInt")
	wi.A    	:= NumGet(WININFO, 56, "UShort")
    wi.V    	:= NumGet(WININFO, 58, "UShort")
Return wi
}

SetWindowPos(hWnd, x, y, w, h, hWndInsertAfter := 0, uFlags := 0x40) {                                		;--works better than the internal command WinMove - why?

	/*  ; https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-setwindowpos

	SWP_NOSIZE                       	:= 0x0001	; Retains the current size (ignores the cx and cy parameters).
	SWP_NOMOVE                 	:= 0x0002	; Retains the current position (ignores X and Y parameters).
	SWP_NOZORDER              	:= 0x0004	; Retains the current Z order (ignores the hWndInsertAfter parameter).
	SWP_NOREDRAW             	:= 0x0008	; Does not redraw changes.
	SWP_NOACTIVATE   	        	:= 0x0010	; Does not activate the window.
	SWP_DRAWFRAME            	:= 0x0020	; Draws a frame (defined in the window's class description) around the window.
	SWP_FRAMECHANGED     	:= 0x0020	; Applies new frame styles set using the SetWindowLong function.
	SWP_SHOWWINDOW        	:= 0x0040	; Displays the window.
	SWP_HIDEWINDOW         	:= 0x0080	; Hides the window
	SWP_NOCOPYBITS            	:= 0x0100	; Discards the entire contents of the client area.
	SWP_NOOWNERZORDER 	:= 0x0200	; Does not change the owner window's position in the Z order.
	SWP_NOREPOSITION        	:= 0x0200	; Same as the SWP_NOOWNERZORDER flag.
	SWP_NOSENDCHANGING	:= 0x0400	; Prevents the window from receiving the WM_WINDOWPOSCHANGING message.
	SWP_DEFERERASE                	:= 0x2000	; Prevents generation of the WM_SYNCPAINT message.
	SWP_ASYNCWINDOWPOS	:= 0x4000	; This prevents the calling thread from blocking its execution while other threads process the request.

	 */

Return DllCall("SetWindowPos", "Ptr", hWnd, "Ptr", hWndInsertAfter, "Int", x, "Int", y, "Int", w, "Int", h, "UInt", uFlags)
}

WinGetClass(hwnd) {                                                                                                                	;-- schnellere Fensterfunktion
	if (hwnd is not Integer)
		hwnd := GetDec(hwnd)
	VarSetCapacity(sClass, 80, 0)
	DllCall("GetClassNameW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass)+1)
	wclass := sClass, sClass := ""
Return wclass
}

WinGet(hwnd, cmd) {                                                                                                                	;-- Wrapper
	WinGet, res, % cmd, % "ahk_id " hwnd
return res
}

WinGetTitle(hwnd) {                                                                                                                   	;-- schnellere Fensterfunktion
	if (hwnd is not Integer)
		hwnd :=GetDec(hwnd)
	vChars := DllCall("user32\GetWindowTextLengthW", "Ptr", hWnd) + 1
	VarSetCapacity(sClass, vChars << !!A_IsUnicode, 0)
	DllCall("user32\GetWindowTextW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass) + 1)
	wtitle := sClass, sClass := ""
Return wtitle
}

WinGetText(hwnd) {                                                                                                                  	;-- Wrapper
	WinGetText, wtext, % "ahk_id " hwnd
Return wtext
}

FindChildWindow(Parent, Child, DetectHiddenWindow="On") {                                                  	;{-- finds childWindow Hwnds of the parent window

/*                                                                                     	READ THIS FOR MORE INFORMATIONS
                                			    	a function from AHK-Forum : https://autohotkey.com/board/topic/46786-enumchildwindows/
                                                                                      it has been modified by IXIKO on May 09, 2018

	-finds childWindow handles from a parent window by using Name and/or class or only the WinID of the parentWindow
	-it returns a comma separated list of hwnds or nothing if there's no match

	-Parent parameter is an object(). Pass the following {Key:Value} pairs like this - WinTitle: "Name of window", WinClass: "Class (NN) Name", WinID: ParentWinID
	FindChildWindow({"ID":hwnd, "exe":"albis"}, {"class":"#32770"})
*/

		detect:= A_DetectHiddenWindows
		global ChildTitle, ChildNN, ChildClass, ChildExe, active_id
		global ChildHwnds

		ChildHwnds := ""

	; build ParentWinTitle parameter from ParentObject
		If Parent.WinID
			ParentWinTitle:= "ahk_id " Parent.ID
		else
			ParentWinTitle:= Parent.Title " ahk_class " Parent.Class

		ChildTitle  	:= Child.Title
		ChildClass	:= Child.class
		ChildNN   	:= Child.classnn
		ChildExe    	:= Child.Exe

		;DetectHiddenWindows, % DetectHiddenWindow  ; Due to fast-mode, this setting will go into effect for the callback too.
		DetectHiddenWindows, Off
		WinGet, active_id, ID, % ParentWinTitle

	; For performance and memory conservation, call RegisterCallback() only once for a given callback:
		if !EnumAddress  ; Fast-mode is okay because it will be called only from this thread:
			EnumAddress := RegisterCallback("EnumChildWindow") ; , "Fast")

		result:= DllCall("EnumChildWindows", "UInt", active_id, "UInt", EnumAddress, "UInt", 0)

		DetectHiddenWindows, % detect

return RTrim(ChildHwnds, ";")
}
; sub of FindChildWindow
EnumChildWindow(hwnd, lParam) {                                                                                             	;--sub function of FindChildWindow

	global ChildHwnds
	global ChildTitle, ChildNN, ChildClass, ChildExe, active_id

	If ChildNN
		classMatched := InStr(GetClassNN(hwnd, active_id), ChildNN) 	? 1 : 0
	else if ChildClass
		classMatched := InStr(WinGetClass(hwnd), ChildClass)           	? 1 : 0
	else
		classMatched := 1

	ProcName := WinGet(hwnd, "ProcessName")

	If InStr(WinGetTitle(hwnd), ChildTitle) && classMatched && InStr(ProcName, ChildExe)
		ChildHwnds.= GetHex(hwnd) "`;"

return true  ; Tell EnumWindows() to continue until all windows have been enumerated.
}
;}

;}

;{ Listview
LV_EX_GettemRect(hLV, Column, Row := 1, LVIR := 0, ByRef RECT := "") {   ; Retrieves information about the bounding rectangle for a subitem in a list-view control.

   ; LVM_GETSUBITEMRECT = 0x1038 -> http://msdn.microsoft.com/en-us/library/bb761075(v=vs.85).aspx
   VarSetCapacity(RECT, 16, 0)
   NumPut(LVIR, RECT, 0, "Int")
   NumPut(Column-1, RECT, 4, "Int")

   SendMessage, 0x100E, % (Row - 1), % &RECT,, % "ahk_id " hLV
   If (ErrorLevel = 0)
      Return False

   If (Column = 1) && ((LVIR = 0) || (LVIR = 3))
      NumPut(NumGet(RECT, 0, "Int") + LV_EX_GetColumnWidth(HLV, 1), RECT, 8, "Int")

	res := {}
	res.X 	:= NumGet(RECT,  0, "Int")
	res.Y 	:= NumGet(RECT,  4, "Int")
	res.R 	:= NumGet(RECT,  8, "Int")
	res.B 	:= NumGet(RECT, 12, "Int")
	res.W	:= res.R-res.X
	res.H	:= res.B-res.Y

	;~ SciTEOutput("H:" cJSON.Dump(res, 1))

Return res
}

LV_EX_GetColumnWidth(HLV, Column) { 	; Gets the width of a column in report or list view.
   ; LVM_GETCOLUMNWIDTH = 0x101D -> http://msdn.microsoft.com/en-us/library/bb774915(v=vs.85).aspx
   SendMessage, 0x101D, % (Column - 1), 0, , % "ahk_id " . HLV
   Return ErrorLevel
}

LVM_GetItemPosition(HLV, row) {

	static LVM_GETITEMPOSITION := 0x1010, LVM_SCROLL := 0x1014

	VarSetCapacity(pt, 8, 0)
	SendMessage % LVM_GETITEMPOSITION, % row-1, % &pt,, % "ahk_id " hLV
	PixelX := NumGet(pt, 0, "int")
	PixelY := NumGet(pt, 4, "int")

return {"X":PixelX, "Y":PixelY}
}

LV_GetCountPerPage(hLV) {
	static LVM_GETCOUNTPERPAGE  := 0x1028,
	SendMessage % LVM_GETCOUNTPERPAGE,,,, % "ahk_id " hLV
return ErrorLevel
}

;}


/*

class callmanager extends Fritzbox_CallMonitor {	; additional class to collect, count and show the phone call data

	;~ /* connects to Albis Patient Database to handle incoming call data

		;~ -  shows incoming calls with names

	;~ */

	static outfilter 	:= ["NR", "PRIVAT", "GESCHL", "NAME", "VORNAME", "GEBURT", "ALTER", "PLZ", "ORT", "STRASSE", "HAUSNUMMER"
								, "TELEFON", "TELEFON2", "TELEFAX", "ARBEIT", "LAST_BEH", "MORTAL"]
	static	PatDB, _init

	SomeInits() {

		;~ static _init := this.SomeInits()
		;~ PatDb	:= new PatDBF(Addendum.AlbisDBPath, outfilter)
		;~ PatDB 	:= ReadPatientDBF(Addendum.AlbisDBPath, outfilter, "allData")
		;~ SciTEOutput("PatDB: " PatDB.Count())
		;~ For Pat
		;~ For Pat

	return true
	}

}


OnReceive(rec:="") {     	; forwards received fritzbox string to dispatcher func

	global callmon
	static _init, managingfunc

	If !_init {
		_init := true
		If !IsObject(callmon.managingfunc)
			throw ".managingfunc: must be a BoundFunc object!"
		managingfunc := callmon.managingfunc
	}

	Critical
	recv	:= RegExReplace( (!IsObject(rec) ? rec : rec.recvText()), "[\n\r]+")
	%managingfunc%(call)


	;~ callmon.recv(call)

}

Recv(params*) {


		static msgstack

		Critical

		SciTEOutput("params: " params.Count())
		For pIndex, param in params {

			SciTEOutput(pIndex ":")

			If IsObject(param) {

				JSONData.Save(A_Temp "\param" pIndex ".json", param, true,, 1, "UTF-8")
				Run % A_Temp "\param" pIndex ".json"

				rec := param
				txt := rec.recvText()
				;~ this.msgstack.InsertAt(1, "obj: " txt)
				this.msgstack.Push("obj: " txt)

				For rKey, rVal in rec
					If !IsObject(rVal)
						SciTEOutput(" -- " rKey ": " rVal)
					else {
						SciTEOutput(" -- " rKey ":" rVal.Count())
						For sKey, sVal in rVal
							If !IsObject(sVal)
								SciTEOutput(" ---- " sKey ": " sVal)
							else
								SciTEOutput(" ---- " sKey ": is object(" rVal.Count() ")")
					}

			}
			else {

				;~ this.msgstack.InsertAt(1, "str: " param)
				this.msgstack.Push("str: " param)
				SciTEOutput(" - param" pIndex ": "param)

			}
		}

		recv := RegExReplace( (!IsObject(rec) ? rec : rec.recvText()), "[\n\r]+")
		;~ If IsObject(this.savefile) && IsObject(rec)
			;~ this.savefile.Write(recv "`n")
		;~ If IsObject(rec) {
		;~ RegExMatch(recv, "O)(?<Date>\d+\.\d+\.d+)\s"
									;~ . "(?<Time>\d+:\d+:\d+);"
									;~ . "(?<Status>.*?);"
									;~ . "(?<>.*?);"
									;~ . "(?<>.*?);"
									;~ . "(?<>.*?);")
		;~ }
		;~ this.msgstack.InsertAt(1, recv)
		;~ If (this.msgstack.Count() > 16)
			;~ this.msgstack.Pop()

		For idx, line in this.msgstack
			If Trim(line)
				msg.= line "`n"
			ToolTip, % "          Fritzbox Anrufe`n----------------------------------`n"msg,1600, 1, 15

}
rRecv(rec:="") {


		Critical
		static msgstack := []
		recv := RegExReplace( (!IsObject(rec) ? rec : rec.recvText()), "[\n\r]+")
		If IsObject(this.savefile) && IsObject(rec)
			this.savefile.Write(recv "`n")
		;~ If IsObject(rec) {
		;~ RegExMatch(recv, "O)(?<Date>\d+\.\d+\.d+)\s"
									;~ . "(?<Time>\d+:\d+:\d+);"
									;~ . "(?<Status>.*?);"
									;~ . "(?<>.*?);"
									;~ . "(?<>.*?);"
									;~ . "(?<>.*?);")
		;~ }
		msgstack.InsertAt(1, recv)
		If (msgstack.Count() > 6)
			msgstack.Pop()

		For idx, line in this.msgstack
			If Trim(line)
				msg.= line "`n"
		ToolTip, % msg,1600, 1, 15

}

*/


#Include %A_ScriptDir%\..\..\
#include include\Addendum_Datum.ahk
#include include\Addendum_DB.ahk
#include include\Addendum_DBase.ahk
#include include\Addendum_Internal.ahk
#include include\Addendum_PraxTT.ahk
#include lib\class_LV_Colors.ahk
#include lib\class_socket.ahk
#include lib\class_cJSON.ahk
#include lib\GDIP_all.ahk
#include lib\SciTEOutput.ahk
#include lib\Sift.ahk
#Include lib\TreeListView.ahk
#Include lib\VarTreeGui.ahk
#Include lib\VarTreeObjectNode.ahk
#Include lib\VarEditGui.ahk




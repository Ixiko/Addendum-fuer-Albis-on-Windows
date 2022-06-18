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
;                        			by Ixiko started in September 2017 - last change 05.06.2022 - this file runs under Lexiko's GNU Licence
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
	CoordMode, ToolTip, Screen
	CoordMode, Mouse, Screen
	CoordMode, Menu, Screen

	Menu, Tray, Icon, % A_ScriptDir "\assets\Fritzboxanrufmonitor.ico"

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

	;~ uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
	;~ SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
	;~ FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")
	;~ DllCall(SetPreferredAppMode, "int", 1) ; Dark
	;~ DllCall(FlushMenuThemes)

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
		ini.WriteLine("Version=0.3a")
		ini.Close()

		InputBox	, TelNumbers
						, Fritzbox Anrufmonitor Initialisierung
						, % 	"Hier eine Kommagetrennte Liste zu überwachender Telefonnummern eingeben.`n"
						  . 	"Vorwahlnummern weglassen! (z.B. 123331, 123356, 1237891)"
						, , 500, 160
		If ErrorLevel
			ExitApp
		IniWrite, % RegExReplace(Trim(TelNumbers), "[\s,\|\-\+\_\\\/]", "|"), % Addendum.fboxini, Telefone, ueberwachen

		InputBox	, NoTel
						, Fritzbox Anrufmonitor Initialisierung
						, % "Hier eine Komma-getrennte Liste mit Telefonnummern aller erreichbarer Geräte eintragen`n"
						  .  "z.B. Tel1=123331, Tel2=123332, Fax=123333, AB1=123334, AB2=123335"
						, , 500, 160
		If ErrorLevel
				ExitApp
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

	  ; Fritzbox IP-Adresse
		IniRead, FritzBoxIP	, % Addendum.fboxini, Fritzbox, IP-Adresse
		If RegExMatch(FritzBoxIP, "i)^(|ERROR)$") || !RegExMatch(FritzBoxIP, "i)^(\d+\.\d+\.\d+\.\d+|(https*\:\/\/)*fritz\.box)$") {

			FritzBoxIP := "fritz.box"
			FritzBoxIPInputBox:
			InputBox	, FritzBoxIP
							, Fritzbox Anrufmonitor Initialisierung
							, % "`n                                 Bitte geben Sie die IP-Adresse Ihrer Fritzbox ein.`n`n"
							.  	  "╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍`n"
							. 	  "                                                     🖧 Н Ӏ Ν Ԝ Ε Ӏ Ѕ 🖧`n"
							.  	  "╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍╍`n"
							. 	  "Sie können hier wie in jedem Internetbrowser die URL (fritz.box) der Fritzbox eingeben. "
							.	  "Das Skript wird dann die IPAdresse automatisch ermitteln.`n"
							.	  "Wenn Sie allerdings einen VPN Zugang für das Internet auf ihrem Client installiert haben, wird das Skript die Adresse nicht ermitteln können."
							,, 600, 290,,,,, % FritzBoxIP
			If ErrorLevel
				ExitApp

			If !FritzBoxIP || !RegExMatch(FritzBoxIP, "i)^(\d+\.\d+\.\d+\.\d+|(https*\:\/\/)*fritz\.box)$") {
				MsgBox, 0x1025, % StrReplace(A_ScriptName, ".ahk"), % "Sie haben keine gültige IP-Adresse eingegeben!`nNochmal oder Abbruch", 12
				IfMsgBox, Retry
					gosub FritzBoxIPInputBox
				ExitApp
			}

			If RegExMatch(FritzBoxIP, "i)^(\d+\.\d+\.\d+\.\d+|(https*\:\/\/)*fritz\.box)$") {

				txt1 := "Ermittle die IP-Adresse Ihrer Fritzbox"
				txt2 := "Prüfe Online-Status Ihrer Fritzbox [IP Adresse: "
				If RegExMatch(FritzBoxIP, "i)fritz\.box")
					txt := txt1, startwith:=1
				else
					txt := txt2, startwith:=2

			  ; Info
				Gui, fmi: new, -Caption +AlwaysOnTop
				Gui, fmi: Color, % "c" "99B898", % "c" "355C7D"
				Gui, fmi: Font, s13 q5
				Gui, fmi: Add, Text, % "xm ym w550 Center Border vIpCheckT", % txt "`n  "
				Gui, fmi: Show,, Überprüfung IP-Adresse


				If (startwith = 1) {
					Sleep 3000
					startwith ++
					FritzBoxIP := IPHelper.ResolveHostname("fritz.box")
					If !RegExMatch(FritzBoxIP, "\d+\.\d+\.\d+\.\d+") {
						MsgBox, 0x1025, % StrReplace(A_ScriptName, ".ahk"), % "Es konnte keine Adresse aufgelöst werden!`nAdresse jetzt ändern?", 12
						Gui, fmi: Destroy
						FritzBoxIP := "192.168.178."
						IfMsgBox, Retry
							gosub FritzBoxIPInputBox
						IfMsgBox, Cancel
							ExitApp
						IfMsgBox, Timeout
							ExitApp
					}
					else {
						GuiControl, fmi:, IpCheckT, % "Ihre Fritzbox hat die IP-Adresse: " FritzBoxIP
						Sleep 3000
					}
				}

				If (startwith = 2) {

					GuiControl, fmi:, IpCheckT, % txt2 . FritzBoxIP "]"
					Ping := IPHelper.Ping(FritzBoxIP, 3000)
					Sleep 5000
					If !Ping {
						MsgBox, 0x1022, % StrReplace(A_ScriptName, ".ahk"), % "Das Gerät mit der IP-Adresse: " FritzboxIP " ist nicht online!`nAdresse jetzt ändern?", 12
						IfMsgBox, Retry
						{
							FritzBoxIP := "192.168.178."
							gosub FritzBoxIPInputBox
						}
						IfMsgBox, Cancel
							ExitApp
						IgnoreOfflineIP := true
						Gui, fmi: Destroy
					}
					else {
						GuiControl, fmi:, IpCheckT, % "Die Fritzbox mit der IP " FritzBoxIP " ist online!`nDie IP Adresse wird gespeichert und anschließend verwendet."
						func_call := Func("fmiGuiTimer").Bind("fmi")
						SetTimer, % func_call, -8000
					}

				}

			}

			IniWrite, % FritzBoxIP, % Addendum.fboxini, FritzBox, IP-Adresse

		}


	  ; Vorwahlnummern Deutschland aufbereiten
		If FileExist(fvorwahl := A_ScriptDir "\deutschland_vorwahl.json") {
			tmp := JSONData.Load(fvorwahl, "", "UTF-8")
			For index, prefix in tmp
				prefixes .= (index>1 ? "|" : "") prefix.number
			Addendum.prefixes := prefixes
			tmp := fvorwahl := index := prefix:= prefixes := ""
		}

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

		prefixes := Fax := index := d := TelNumbers := ""
		val := shunts := shuntIndex := shunt := ""

  ;}

  ; CallersDB laden
	callersDB := new phonecallers(Addendum.AlbisDBPath)

  ; Gui zeigen
	callGui()

  ; Anrufmonitor starten
	callmon	:= new Fritzbox_Callmonitor(FritzBoxIP, {"savefilepath":"C:\tmp\cm.txt", "managingfunc":"CallManager"})
	connected	:= callmon.Connect()
	SciTEOutput(" > connected: " connected)

  ; bei erfolgreicher Verbindung wird die Objektvariable .connected auf wahr gesetzt
	callmon._OnReceive("savepath: " callmon.savepath)

  ; aktuelles Protokoll anzeigen wenn vorhanden
	callGui_Load()

  ; Neustart  um 0 Uhr
	func_call := Func("callGui_Exit").Bind(true)
	SetTimer, % func_call, % -1*TimerTime("00:00 Uhr")

return

;{  Hotkeys
!Esc::
	callmon.Disconnect()
ExitApp

^+a::
	callGui_Exit(true)
return
;}

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
			If (StrLen(options.managingfunc)>0 && IsFunc(options.managingfunc))
				this.managingfunc := Func(options.managingfunc)
			;~ else if IsFunc(options.managingfunc)
				;~ this.managingfunc := options.managingfunc
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


        if (StrLen(method)=0 && !IsObject(method)) {           	; For %fn%() or fn.()
			SciTEOutput(" > method call 1 is invoked")
            return this.Call(args*)
		}
        else if IsObject(method) {                                       	; If this function object is being used as a method.
			SciTEOutput(" > method call 2 is invoked")
            return this.Call(method, args*)
		}

    }

	_OnReceive(rec:="")                                                                   	{ 	;-- forwards received fritzbox string to dispatcher func

		static init, managingfunc

		SciTEOutput(" > Fritz!Monitor: received data " A_Hour ":" A_Min ":" A_Sec)

		If !init {
			init := true
			If !IsObject(this.managingfunc)
				throw ".managingfunc: must be a BoundFunc object!"
			managingfunc := this.managingfunc
		}

		Critical

		If IsObject(rec) {
			recv	:= rec.recvText()
			SciTEOutput(" > " recv)
			this.stringfile := FileOpen(this.savepath "\" A_YYYY A_MM A_DD "_FbxCMStrings.txt", "a", "UTF-8")
			this.stringfile.WriteLine(recv := RegExReplace(recv, "[\n\r]+"))
			this.stringfile.Close()
			call := this.CallStringParser(recv)
			%managingfunc%(call)
		}
		else {
			SciTEOutput(" > " rec)
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

		/*  					1			  |        	2  		| 3| 4|      5		|	  6    	|	  7
				08.06.22 11:21:41;	CALL				; 1; 14; 443261; 443260	;	SIP0;
				08.06.22 11:21:53;	CONNECT		; 1; 14; 443260;
				08.06.22 11:22:16;	DISCONNECT	; 1; 24;
		*/

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

	GetNameFromNumber(nr, mergenames=1)	{       	; sucht Patientennamen in der Albis-Datenbank

		nr := RegExReplace(nr, "[^\d]")
		matches := Array()

		For PatID, Pat in this.PatientsDB {
			For TelIndex, TelNr in Pat.TELNR       	; Patiententelefonnummern heraussuchen
				If RegExMatch(nr, TelNr "$") {   	 	; (TelNr && )

					If (!matches.Count() || !mergenames)
						matches.Push([PatID, TelNr, (PAT.VORNAME " " Pat.NAME)])
					else
						For pIndex, pData in matches
							If InStr(pData.3, " " Pat.NAME) {
							  ; verhindert >2 Vornamen (Klaus & Bernd & Renate Müller)
								If !Instr(matches[pIndex].3, "&") {
									matches[pIndex].3 := PAT.VORNAME " & " matches[pIndex].3
									matches[pIndex].1 .= "|" PatID
								}
								break
							}
						;~ continue
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

	SciTEOutput(cJSON.Dump(call, 1))

 ; execution depending on event
	Switch call.event {

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; RING (registers incoming call)
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
				callerstack[cID].3 := callerstack[cID].3 ? callerstack[cID].3 : 1                                                 	; Anruf wurde entgegen genommen
			}
			else if !callstack[conID].ignore{
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "🆎")
				LV_Modify(callerstack[cID].2, "Col5", call.time)
				callerstack[cID].4 := 1                                                                                                       	; Anrufbeantworter ist ran gegangen  📤
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

		  ; call was answered, calls with suppressed phone number are also moved to the 2nd list.
			If callstack[conID].ignore {

				Gui, CV: Default
				Gui, CV: ListView, Day
				LV_Modify(callstack[conID].Day, "Col3", (callstack[conID].connect ? "" : "✗") (RegExMatch(callstack[conID].to, Addendum.Fax) ? "℻" : ""))
				LV_Modify(callstack[conID].Day, "Col6", call.time)
				callstack[conID] := "-"

			}
			else {

			  ; connection string (constr) and  the number of the caller (callNr)
				constr		:= 	callstack[conID]
				callNr 		:= 	callerstack[cID].3<2 ? constr.from : constr.to

			  ; Call was answered by picking up the handset are moved to the second listview. Calls with suppressed phone number are also postponed.
				If (callerstack[cID].3>0 || InStr(constr.caller, "unbekannt")) && (constr.shuntnr<>40 && constr.duration > 0) {

					talktime += constr.duration

					Gui, CV: Default
					Gui, CV: ListView, Callers
					LV_Delete(callerstack[cID].2)

				; Search for entries still present from previous calls and then remove them
					maxRows := LV_GetCount() + 1
					Loop {
						LV_GetText(LVTelNr	, maxRows - A_Index, 2)
						LV_GetText(LVCalls	, maxRows - A_Index, 4)
						LV_GetText(LVTime	, maxRows - A_Index, 5)
						LVTelnr := plainnumber(LVTelNr)
						If (callNr = LVTelNr) {
							constr.calls += LVCalls
							LV_Delete(A_Index)
						}
					}

				; show summarized call data for answered calls
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

	global Callers, Day, hCV, BLV := Object(), hLV1, hLV2, hCVT1, LVColors, ontop

  ; Fenstergröße
	IniRead, winSize, % A_ScriptDir "\FritzboxCallMonitor.ini", % "sonstiges", % "Fensterposition"
	winSize	:=  InStr(winSize, "ERROR") ? "" : winSize
	RegExMatch(winSize, "x\s*(?<X>\-*\d+)", win)
	RegExMatch(winSize, "y\s*(?<Y>\-*\d+)", win)
	winSize	:= (winX<0 ? "" : "x" winX) (winY<0 ? "" : " y" winY " " )

  ; andere Fensterparameter
	LVNames 	:= ["Callers", "Day"]
	colSizes 	:= {"Callers":[230, 140, 50, "60 Center", 125], "Day":[230, 140, 50, 60, "55 Center", 70]}
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

  ; Gui
	Gui, CV: new, % "hwndhCV " ( ontop ? "AlwaysOnTop" : "")
	Gui, CV: Color    	, % "c" "355C7D" , % "c" "99B898"
	Gui, CV: Margin	, 0 , 5
	Gui, CV: Font    	, s10 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "xm ym   	w" LVWidth["Day"]+wplus    	" Center cWhite hwndhCVT1"                  	, aktuelle Anrufe
	opt := " hwndCVhOnTop vCVOnTop gcallGui_Handler"
	Gui, CV: Add     	, Button  	, % "x" LVWidth["Callers"]+wplus-22 " y2 w20 h20 " opt, % (ontop ? "🔐" : "🔓")
	Gui, CV: Font    	, s10 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, ListView	, % "xm y+2   	w" LVWidth["Callers"]+wplus 	" r15 " LVOpt "1  vCallers"  	, Anrufer|Telefon|Status|Anrufe|Uhrzeit

	Gui, CV: Font    	, s10 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "y+2    	w" LVWidth["Day"]+wplus    	" Center cWhite"                  	, angenommene Anrufe
	Gui, CV: Font    	, s10 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, ListView	, % "y+0      	w" LVWidth["Day"]+wplus    	" r20 " LVOpt "2 vDay"       	, Anrufer|Telefon|Status|Dauer|Anrufe|Uhrzeit

	WinSet, ExStyle, 0x0, % "ahk_id " hLV1
	WinSet, ExStyle, 0x0, % "ahk_id " hLV2

  ; die Höhe der Listview's soll ein exaktes Vielfaches der Zeilenhöhe sein
	If IsObject(itemrect := LV_EX_GettemRect(hLV2, 1, 1, 3)) {
		Lv := GetWindowSpot(hLV2)
		SetWindowPos(hLV2, Lv.X, Lv.Y, LV.W, itemrect.H * LV_GetCount())
		;~ SciTEOutput("H: " itemrect.H "*" LV_GetCount() " = " itemrect.H * LV_GetCount())
	}

	Gui, CV: Font    	, s8 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, StatusBar, hwndhSbar, % ""
	;~ DllCall("uxtheme\SetWindowTheme", "ptr", hSbar, "str", "DarkMode_Explorer", "ptr", 0)

	Gui, CV: Show    	, % winSize := (winX<0 ? "" : "x" winX) (winY<0 ? "" : " y" winY " " ) , Fritz!Monitor

  ; Listview 1 aktivieren
	ControlClick,, % "ahk_id " hLV1,, Left, 1

  ; Spaltenbreiten anpassen
	For LVIndex, LVName in LVNames {
		Gui, CV: ListView, % LVName
		For colNr, colOpt in colSizes[LVName]
			LV_ModifyCol(colNr, colOpt)
	}

  ; bringt Farbe in die Listview
	BLV["Callers"] := new LV_Colors(hLV1)
	BLV["Callers"].Critical := 200
	BLV["Callers"].SelectionColors("0x0078D6")

	BLV["Day"] := new LV_Colors(hLV2)
	BLV["Day"].Critical := 200
	BLV["Day"].SelectionColors("0x0078D6")

return
CVGuiClose:
CVGuiEscape:
	WinMinimize, % "ahk_id " hCV
return
}

callGui_Load(datestr:="")                                                                           	{	;-- loads phone data from saved Fritzbox files

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
	If (wqs.X>-8 && wqs.X>-8)
		IniWrite, % "x" wqs.X " y" wqs.Y, % A_ScriptDir "\FritzboxCallMonitor.ini", % "sonstiges", % "Fensterposition"

	If !FileExist(statspath := A_ScriptDir "\telefon_statistik.ini") {
		IniWrite, % "Anzahl Anrufe;versch.Telefonnummern; davon Unbekannt;ausgehende Anrufe;Gesprächszeit" , % statspath, _Datenstruktur
	}
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
	static cmenu, init

	Critical

	If !init {
		init := true
		;~ funcCM := Func("callGui_CM").Bind("remove")
		;~ Menu, cmenu, Add, von Anruferliste nehmen, % funcCM
		;~ funcCM := Func("callGui_CM").Bind("copy")
		;~ Menu, cmenu, Add, Telefonnummer kopieren, % funcCM
		;~ funcCM := Func("callGui_CM").Bind("plaincopy")
		;~ Menu, cmenu, Add, Telefonnummer (nur Zahlen) kopieren, % funcCM
		;~ Menu, cmneu, Add
		;~ funcCM := Func("callGui_CM").Bind("opencase", PatID)
		;~ Menu, cmenu, Add, Karteikarte öffnen, % funcCM
	}

	If (StrLen(A_GuiControl)=0) || (StrLen(EventInfo)=0) || !EventInfo
		return

	Gui, CV: Default
	Gui, CV: ListView, % (clickedLV := A_GuiControl)
	LV_GetText(clkLVCaller	, (clickedLVRow := A_EventInfo), 1)
	LV_GetText(clkLVTelNr	, clickedLVRow, 2)
	MouseGetPos, mx, my

  ; Doppelklick
	If (A_GuiEvent = "DoubleClick") {
		callGui_Rename(EventInfo, A_GuiControl)
	}

  ; Kontextmenu
	else if (A_GuiEvent = "RightClick") {

		If IsObject(cmenu)
			Menu, cmenu, DeleteAll
		funcCM1 	:= Func("callGui_CM").Bind("copy", clickedLV, clickedLVRow)
		funcCM2	:= Func("callGui_CM").Bind("plaincopy", clickedLV, clickedLVRow)

		If (clickedLV = "Callers") {
			funcCM := Func("callGui_CM").Bind("remove", clickedLV, clickedLVRow)
			Menu, cmenu, Add, von Anruferliste nehmen, % funcCM
		}

		Menu, cmenu, Add, Telefonnummer kopieren                 	, % funcCM1
		Menu, cmenu, Add, Telefonnummer (nur Zahlen) kopieren	, % funcCM2
		Menu, cmneu, Add

		;~ funcCM := Func("callGui_CM").Bind("opencase", PatID)
		;~ Menu, cmenu, Add, Karteikarte öffnen, % funcCM

		Menu, cmenu, Show, % mx-20, % my

	}

  ; einfacher Klick
	else if (A_GuiEvent = "Normal") {
		If !LVHandlerProc
			callGui_MarkCallers(clickedLV, clickedLVRow)
	}



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

	Gui, CV:Default
	Gui, CV:ListView, % LVclicked

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

	Gui, CV:Default
	Gui, CV:ListView, % altLV

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

callGui_CM(cmd, LVclicked:="", rowclicked:=0)                                         	{ ;-- context menu

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

		SciTEOutput("itemmoved: " itemmoved ", " itembesttime)
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

	global uphones

	MouseGetPos, mx, my

	Gui, CV: Default
	Gui, CV: ListView, % LVName

	LV_GetText(uNAME	, LVRow, 1)
	LV_GetText(uTEL    	, LVRow, 2)

	InputBox	, callername, Fritzbox Anrufmonitor, % "Sie ändern: " uName " der Telefonnummer: " uTEL
					,, 300, 140,,,,, % (!RegExMatch(uName, "i)unbekannt.*Name") ? uName : "")

	If RegExMatch(callername, "^\s*$")
		return

	If RegExMatch(uName, "i)unbekannt.*Anruf") && !RegExMatch(callername, "i)unbekannt.*Anruf") {
		Gui, CV: Default
		Gui, CV: ListView, % LVName
		LV_Modify(LVRow, 1, callername)
		IniWrite, % callername, % Addendum.fboxini, Telefonbuch, % (telnr := RegExReplace(uTEL, "[^\d]"))
		uphones.Delete(TelNr)
		callGui_SBSetText()
	}

}


; Hilfsfunktionen
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

AlbisMDIChildHandle(MDITitle) {		                                                              	;-- determines the handle of a sub or child window within the Albis MDI control
return GetHex(FindChildWindow({"Class": "OptoAppClass"}, {"Title": MDITitle}, "Off"))
}

FindChildWindow(Parent, Child, DetectHiddenWindow="On") {                      	;{-- finds childWindow Hwnds of the parent window

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

WinGetClass(hwnd) {                                                                                    	;-- fast window function
	if (hwnd is not Integer)
		hwnd := GetDec(hwnd)
	VarSetCapacity(sClass, 80, 0)
	DllCall("GetClassNameW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass)+1)
	wclass := sClass, sClass := ""
Return wclass
}

WinGet(hwnd, cmd) {                                                                                    	;-- Wrapper
	WinGet, res, % cmd, % "ahk_id " hwnd
return res
}

WinGetTitle(hwnd) {                                                                                       	;-- schnellere Fensterfunktion
	if (hwnd is not Integer)
		hwnd :=GetDec(hwnd)
	vChars := DllCall("user32\GetWindowTextLengthW", "Ptr", hWnd) + 1
	VarSetCapacity(sClass, vChars << !!A_IsUnicode, 0)
	DllCall("user32\GetWindowTextW", "UInt", hWnd, "Str", sClass, "Int", VarSetCapacity(sClass) + 1)
	wtitle := sClass, sClass := ""
Return wtitle
}

LV_EX_GettemRect(HLV, Column, Row := 1, LVIR := 0, ByRef RECT := "")  {       	;-- Retrieves information about the bounding rectangle for a subitem in a list-view control.
   ; LVM_GETSUBITEMRECT = 0x1038 -> http://msdn.microsoft.com/en-us/library/bb761075(v=vs.85).aspx
   VarSetCapacity(RECT, 16, 0)
   NumPut(LVIR, RECT, 0, "Int")
   NumPut(Column - 1, RECT, 4, "Int")
   SendMessage, 0x100E, % (Row - 1), % &RECT,, % "ahk_id " . HLV
   If (ErrorLevel = 0)
      Return False
   If (Column = 1) && ((LVIR = 0) || (LVIR = 3))
      NumPut(NumGet(RECT, 0, "Int") + LV_EX_GetColumnWidth(HLV, 1), RECT, 8, "Int")
   Result := {}
   Result.X := NumGet(RECT,  0, "Int"), Result.Y := NumGet(RECT,  4, "Int")
   Result.R := NumGet(RECT,  8, "Int"), Result.B := NumGet(RECT, 12, "Int")
   Result.W := Result.R - Result.X,     Result.H := Result.B - Result.Y
   Return Result
}

LV_EX_GetColumnWidth(HLV, Column) {                                                       	;-- Gets the width of a column in report or list view.
   ; LVM_GETCOLUMNWIDTH = 0x101D -> http://msdn.microsoft.com/en-us/library/bb774915(v=vs.85).aspx
   SendMessage, 0x101D, % (Column - 1), 0, , % "ahk_id " . HLV
   Return ErrorLevel
}

GetWindowSpot(hWnd) {                                                                                 	;-- gets window position
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
	wi.Ac  	:= NumGet(WININFO, 44, "UInt")
    wi.BW 	:= NumGet(WININFO, 48, "UInt")
    wi.BH  	:= NumGet(WININFO, 52, "UInt")
	wi.A    	:= NumGet(WININFO, 56, "UShort")
    wi.V    	:= NumGet(WININFO, 58, "UShort")
Return wi
}

SetWindowPos(hWnd, x, y, w, h, hWndInsertAfter := 0, uFlags := 0x40) {    		;-- works better than the internal command WinMove - why?

	/*  ; https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-setwindowpos

	SWP_NOSIZE                       	:= 0x0001	; Retains the current size (ignores the cx and cy parameters).
	SWP_NOMOVE                 	:= 0x0002	; Retains the current position (ignores X and Y parameters).
	SWP_NOZORDER              	:= 0x0004	; Retains the current Z order (ignores the hWndInsertAfter parameter).
	SWP_NOREDRAW             	:= 0x0008	; Does not redraw changes.
	SWP_NOACTIVATE   	       	:= 0x0010	; Does not activate the window.
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

fmiGuiTimer(guiname) {
	Gui, %guiname%: Destroy
}

ScanSubnet(addresses:="") {	                                                                          	;-- pings IP ranges & returns active IP's

	/*
	Progress, % "B1 M P0 T cW202842 cB8088C2 cTFFFFFF zH25 w" 500 " WM400 WS500", % "ermittle verbundene Netzwerkgeräte", % "Fritzbox IP-Adresse fehlt ...", % "warte auf IP Adresse ...", Segoe UI
		hProgressWindow := WinExist("warte auf IP Adresse")
		adr := ScanSubnet()
		;~ WinGetPos	,,,, wh	, % "ahk_id " hwnd
		;~ WinGetPos	,,,, th	, ahk_class Shell_TrayWnd
		;~ WinMove		, % "ahk_id " hwnd,,, % A_ScreenHeight-wh-th

		Progress, Off

	 */

	BL := A_BatchLines
	IPAddresses := []
	If !addresses{	;scan all active network adapters/connections for active IP's... if no ip's were specified...
		colItems := ComObjGet( "winmgmts:" ).ExecQuery("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")._NewEnum
		while colItems[objItem]
			IPAddresses.Push(objItem.IPAddress[0])
	}

	; pings
	rVal := []
	For each, IP in IPAddresses 	{

		;~ If (!InStr(addr, IPAddress) )
			;~ addr .= (addr ? A_Space "or Address = '" : "Address = '" ) . IP . "'"
		addrArr := StrSplit(IP, ".")
		Loop 256
			addr .= (addr ? A_Space "or Address = '" : "Address = '" ) . addrArr.1 "." addrArr.2 "." addrArr.3 "." A_Index-1 "'"

		SciTEOutput(each ": " ip)

	}

	SciTEOutput(addr)

	For each, IPAddress in IPAddresses 	{
		;~ colPings := ComObjGet( "winmgmts:" ).ExecQuery("Select * From Win32_PingStatus where " addr "")._NewEnum
		colPings := ComObjGet( "winmgmts:" ).ExecQuery("Select * From Win32_PingStatus where " addr "")._NewEnum
		While colPings[objStatus]
			If (oS:=(objStatus.StatusCode="" || objStatus.StatusCode<>0)) {
				rVal.Push(objStatus.Address)
				deviceName := IPHelper.ReverseLookup(objStatus.Address)
				ControlSetText, Static2, % "IP-Adresse [" objStatus.Address "] Nr. "  rVal.Count()  ", Gerätename: " deviceName, % "ahk_id " hProgressWindow
				ControlSetText, Static1, % "IP-Adresse [" objStatus.Address "] Statuscode: "  objStatus.StatusCode  ", Gerätename: " deviceName, % "ahk_id " hProgressWindow
			}

		addr := ""	;the quota on Win32_PingStatus prevents more than roughtly two ip ranges being scanned simultaneously...so each range is scanned individually.
	}
	Return rVal
}

/*
    IPHelper by jNizM
    https://github.com/jNizM/AHK_Scripts/blob/master/src/net/Class_IPHelper.ahk

    MsgBox % IPHelper.ResolveHostname("google-public-dns-a.google.com")    ; -> 8.8.8.8
    MsgBox % IPHelper.ReverseLookup("8.8.8.8")                             ; -> google-public-dns-a.google.com
    MsgBox % IPHelper.Ping("8.8.8.8")                                      ; -> 24

    MsgBox % IPHelper.ResolveHostname("autohotkey.com")                    ; -> 104.24.122.247
    MsgBox % IPHelper.ReverseLookup("104.24.122.247")                      ; -> 104.24.122.247 (because no reverse pointer is set)
    MsgBox % IPHelper.Ping("autohotkey.com")                               ; -> 129

*/

class IPHelper{

    static hWS2_32   := DllCall("LoadLibrary", "str", "ws2_32.dll",   "ptr")
    static hIPHLPAPI := DllCAll("LoadLibrary", "str", "iphlpapi.dll", "ptr")

    Ping(addr, timeout := 1000)    {
        if !(RegExMatch(addr, "^((|\.)\d{1,3}){4}$"))
            addr := this.ResolveHostname(addr)
        in_addr := this.inet_addr(addr)
        hICMP := this.IcmpCreateFile()
        rtt := this.IcmpSendEcho(hICMP, in_addr, timeout)
        return rtt, this.IcmpCloseHandle(hICMP)
    }

    ResolveHostname(hostname)    {
        this.WSAStartup()
        ip_addr := this.getaddrinfo(hostname)
        return ip_addr, this.WSACleanup()
    }

    ReverseLookup(ip_addr)    {
        this.WSAStartup()
        in_addr := this.inet_addr(ip_addr)
        hostname := this.getnameinfo(in_addr)
        return hostname, this.WSACleanup()
    }

    ; ===========================================================================================================================
    ; WSAStartup                                                  https://msdn.microsoft.com/en-us/library/ms742213(v=vs.85).aspx
    ; ===========================================================================================================================
    WSAStartup()    {
        static WSASIZE := 394 + (A_PtrSize - 2) + A_PtrSize
        VarSetCapacity(WSADATA, WSASIZE, 0)
        if (DllCall("ws2_32\WSAStartup", "ushort", 0x0202, "ptr", &WSADATA) != 0)
            throw Exception("WSAStartup failed", -1)
        return true
    }

    ; ===========================================================================================================================
    ; WSACleanup                                                  https://msdn.microsoft.com/en-us/library/ms741549(v=vs.85).aspx
    ; ===========================================================================================================================
    WSACleanup()    {
        if (DllCall("ws2_32\WSACleanup") != 0)
            throw Exception("WSACleanup failed: " DllCall("ws2_32\WSAGetLastError"), -1)
        return true
    }

    ; ===========================================================================================================================
    ; getaddrinfo                                                 https://msdn.microsoft.com/en-us/library/ms738520(v=vs.85).aspx
    ; ===========================================================================================================================
    getaddrinfo(hostname)    {
        VarSetCapacity(addrinfo, 16 + 4 * A_PtrSize, 0)
        NumPut(2, addrinfo, 4, "int") && NumPut(1, addrinfo, 8, "int") && NumPut(6, addrinfo, 12, "int")
        if (DllCall("ws2_32\getaddrinfo", "astr", hostname
                                        , "ptr",  0
                                        , "ptr",  &addrinfo
                                        , "ptr*", result) != 0)
            throw Exception("getaddrinfo failed: " DllCall("ws2_32\WSAGetLastError"), -1), this.WSACleanup()
        addr := StrGet(this.inet_ntoa(NumGet(NumGet(result+0, 16 + 2 * A_PtrSize) + 4, 0, "uint")), "cp0")
        return addr, this.freeaddrinfo(result)
    }

    ; ===========================================================================================================================
    ; freeaddrinfo                                                https://msdn.microsoft.com/en-us/library/ms737931(v=vs.85).aspx
    ; ===========================================================================================================================
    freeaddrinfo(addrinfo)    {
        DllCall("ws2_32\freeaddrinfo", "ptr", addrinfo)
    }

    ; ===========================================================================================================================
    ; getnameinfo                                                 https://msdn.microsoft.com/en-us/library/ms738532(v=vs.85).aspx
    ; ===========================================================================================================================
    getnameinfo(in_addr)    {
        static NI_MAXHOST := 1025
        size := VarSetCapacity(sockaddr, 16, 0), NumPut(2, sockaddr, 0, "short") && NumPut(in_addr, sockaddr, 4, "uint")
        VarSetCapacity(hostname, NI_MAXHOST, 0)
        if (DllCall("ws2_32\getnameinfo", "ptr",  &sockaddr
                                        , "int",  size
                                        , "ptr",  &hostname
                                        , "uint", NI_MAXHOST
                                        , "ptr",  0
                                        , "uint", 0
                                        , "int",  0))
            throw Exception("getnameinfo failed: " DllCall("ws2_32\WSAGetLastError"), -1), this.WSACleanup()
        return StrGet(&hostname+0, NI_MAXHOST, "cp0")
    }

    ; ===========================================================================================================================
    ; inet_addr                                                   https://msdn.microsoft.com/en-us/library/ms738563(v=vs.85).aspx
    ; ===========================================================================================================================
    inet_addr(ip_addr)    {
        in_addr := DllCall("ws2_32\inet_addr", "astr", ip_addr, "uint")
        if !(in_addr) || (in_addr = 0xFFFFFFFF)
            throw Exception("inet_addr failed", -1)
        return in_addr
    }

    ; ===========================================================================================================================
    ; inet_ntoa                                                   https://msdn.microsoft.com/en-us/library/ms738564(v=vs.85).aspx
    ; ===========================================================================================================================
    inet_ntoa(in_addr)    {
        if !(buf := DllCall("ws2_32\inet_ntoa", "uint", in_addr, "ptr"))
            throw Exception("inet_ntoa failed", -1)
        return buf
    }

    ; ===========================================================================================================================
    ; IcmpCreateFile                                              https://msdn.microsoft.com/en-us/library/aa366045(v=vs.85).aspx
    ; ===========================================================================================================================
    IcmpCreateFile()    {
        if !(hIcmpFile := DllCall("iphlpapi\IcmpCreateFile", "ptr"))
            throw Exception("IcmpCreateFile failed", -1)
        return hIcmpFile
    }

    ; ===========================================================================================================================
    ; IcmpSendEcho                                                https://msdn.microsoft.com/en-us/library/aa366050(v=vs.85).aspx
    ; ===========================================================================================================================
    IcmpSendEcho(hIcmpFile, in_addr, timeout)    {
        size := VarSetCapacity(buf, 32 + 8, 0)
        if !(DllCall("iphlpapi\IcmpSendEcho", "ptr",    hIcmpFile
                                            , "uint",   in_addr
                                            , "ptr",    0
                                            , "ushort", 0
                                            , "ptr",    0
                                            , "ptr",    &buf
                                            , "uint",   size
                                            , "uint",   timeout
                                            , "uint"))
            throw Exception("IcmpSendEcho failed", -1)
        return (rtt := NumGet(buf, 8, "uint")) < 1 ? 1 : rtt
    }

    ; ===========================================================================================================================
    ; IcmpCloseHandle                                             https://msdn.microsoft.com/en-us/library/aa366043(v=vs.85).aspx
    ; ===========================================================================================================================
    IcmpCloseHandle(hIcmpFile)    {
        if !(DllCall("iphlpapi\IcmpCloseHandle", "ptr", hIcmpFile))
            throw Exception("IcmpCloseHandle failed", -1)
        return true
    }
}


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


#include %A_ScriptDir%\..\..\include\Addendum_Datum.ahk
#include %A_ScriptDir%\..\..\include\Addendum_DB.ahk
#include %A_ScriptDir%\..\..\include\Addendum_DBase.ahk
#include %A_ScriptDir%\..\..\include\Addendum_Internal.ahk

#Include %A_ScriptDir%\..\..\lib\class_LV_Colors.ahk
#include %A_ScriptDir%\..\..\lib\class_cJSON.ahk
#include %A_ScriptDir%\..\..\lib\class_json.ahk
#include %A_ScriptDir%\..\..\lib\class_socket.ahk
#include %A_ScriptDir%\..\..\lib\GDIP_all.ahk
#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#include %A_ScriptDir%\..\..\lib\Sift.ahk
;~ #include %A_ScriptDir%\..\..\include\Addendum_Stackify.ahk



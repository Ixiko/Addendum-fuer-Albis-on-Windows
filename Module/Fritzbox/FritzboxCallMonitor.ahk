; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;                                                             	☎	Addendum Fritzbox Anrufmonitor V0.6a  	☎
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;
;                      ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
;                        Was unterscheidet diesen Anrufmonitor von anderen, wie dem CallMonitor oder dem JAnrufmonitor?
;                      ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;
;    	Dieser Anrufmonitor ist zunächst speziell für die Nutzung mit Albis, da er auf Daten aus Albis-Datenbanken zurück greift, um die
;     	Anrufer zu identifizieren. Die eigentliche Aufgabe ist das Filtern und Zählen von Anrufen, damit jederzeit klar mit wem bereits
;		telefoniert wurde und mit wem nicht. Deshalb werden keine einzelnen Anrufe gezeigt, sondern gebündelt nach Anrufer dargestellt.
;		So wird man in der Hektik einer Sprechstunde, in der sowieso nie alle Anrufe entgegen genommen werden können, zumindestens
;		vermeintlich dringende Anrufe erkennen können (ohne das man Sprechstunden AB benötigt).
;		weitere mögliche Vorteile: 	- nicht gespeicherte der geänderte Telefonnummern erkennen
;												- den Anrufer mit seinem Namen begrüßen
;												- Statistik der Anrufzahlen, Anrufer und gesamten Gesprächsdauer
;
;
;
;		Abhängigkeiten:	siehe includes
;
;	                    			This part of Addendum für Albis on Windows
;                        			by Ixiko started in September 2017 - last change 20.12.2022 - this file runs under Lexiko's GNU Licence
; ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
; ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
;
;
;


;~ FileDelete, % A_Temp "\chars.txt"

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

  ; Scriptversion
	tmp := FileOpen(A_ScriptFullPath, "r", "UTF-8").Read()
	RegExMatch(tmp, "Anrufmonitor (?<Version>V\d+(\.\d+\pL*)*)", FAM)
	Addendum.FAMVersion := FAMVersion ? FAMVersion : "V?.?"

  ; Traymenu
	Menu, Tray, Icon, % A_ScriptDir "\assets\Fritzboxanrufmonitor.ico"
	Menu, Tray, NoStandard
	fn := func("callProperties")
	Menu, Tray, Add, % "Fritz!box Anrufmonitor " Addendum.FAMVersion, % fn
	Menu, Tray, Add
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

		IBoxTitle := "Fritzbox Anrufmonitor Initialisierung"

	  Label_FritzboxIP:
		;~ fboxIP := "192.168.100.254"
		InputBox	, fboxIP, % IBoxTitle
						, % "Ein paar wenige Angaben sind vor dem Start des Anrufmonitors notwendig.`n"
						.	 "In der Einstellungsdatei [" RegExReplace(Addendum.fboxini, "i)^[\w\:\_\-\\\$\s]+\\") "] lassen`n"
						.	 "sich die Einstellungen jederzeit ändern.`n"
						. 	 "Bitte geben Sie zuerst die IP Adresse Ihrer Fritzbox ein.",, 500, 150,,, % fboxIP
		fboxIP := RegExReplace(fboxIP, "\s")
		If (!fboxIP || !RegExMatch(fboxIP, "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$")) {
			MsgBox, 0x1000	, Fritzbox Anrufmonitor Initialisierung
										, % !fboxIP     	? "Sie haben nichts eingegeben. Das Skript wird jetzt abgebrochen."
																: "Bitte nur Zahlen und Punkte eingeben und keine '"
																. (fboxIP := RegExReplace(fboxIP, "[^\d\.]")) "'!"
			If !localprefix
				ExitApp
			goto Label_FritzboxIP
		}
		If !IPHelper.Ping(fboxIP) {
			MsgBox, 0x1000	, Fritzbox Anrufmonitor Initialisierung
										, % "Die Fritzbox mit der Adresse: " fboxIP " antwortet nicht.`n"
										. 	  "Bitte überprüfen Sie die eingegebene Adresse und die Fritzbox!'"
			goto Label_FritzboxIP
		}
		fboxIP := RegExReplace(fboxIP, "[^\d\.]")


	  Label_localprefix:
		InputBox	, localprefix, % IBoxTitle
						, % "Bitte geben Sie Ihre Ortsvorwahl ein.",, 500, 150,,, % localprefix
		localprefix := RegExReplace(localprefix, "\s")
		If (!localprefix || !RegExMatch(localprefix, "^0\d+$")) {
			MsgBox, 0x1000	, Fritzbox Anrufmonitor Initialisierung
										, % !localprefix 	? "Sie haben nichts eingegeben. Das Skript wird jetzt abgebrochen."
																: "Bitte nur Zahlen eingeben und keine '"
																. (localprefix := RegExReplace(localprefix, "[^\d]")) "'!"
			If !localprefix
				ExitApp
			goto Label_localprefix
		}
		localprefix := RegExReplace(localprefix, "[^\d]")


	  Label_NoTel:
		InputBox	, NoTel, % IBoxTitle
						, % 	"Hier geben Sie bitte eine Komma-getrennte Liste mit Gerätebezeichnungen`n"
						.     	"und Telefonnummern aller Telefoniegeräte ein.`n"
						.   	"z.B. Tel1=123331, Tel2=123332, Fax=123333, AB1=123334, AB2=123335",, 500, 120,,, % NoTel
		NoTel := RegexReplace(NoTel, "\s")
		If (!NoTel || !RegExMatch(NoTel, "i)^[\wäöüß\=,\s]+$")) {
			MsgBox, 0x1000, Fritzbox Anrufmonitor Initialisierung
									, % !NoTel      	? "Sie haben nichts eingegeben. Das Skript wird jetzt abgebrochen."
												     		: "Bitte nur Zahlen, Buchstaben des deutsche Alphabet, Komma und Leerzeichen eingeben und nicht '"
															. (NoTel := RegExReplace(NoTel, "i)[^\wäöüß\=,\s]")) "'!"
			If !NoTel
				ExitApp
			goto Label_NoTel
		}
		NoTel := RegExReplace(NoTel, "i)[^\wäöüß\=,]")


	  Label_TelNumbers:
		InputBox	, TelNumbers, % IBoxTitle
						, % 	"Hier eine Kommagetrennte Liste zu überwachender eigener Telefonnummern eingeben.`n"
						.   	"Vorwahl weglassen! (z.B. 123331, 123356, 1237891)",, 500, 120,,, % TelNumbers
		TelNumbers := RegexReplace(TelNumbers, "\s")
		If (!RegExMatch(TelNumbers, "^[\d,]+$")) {
			MsgBox, 0x1000	, Fritzbox Anrufmonitor Initialisierung
										, % !TelNumbers 	? "Sie haben nichts eingegeben. Das Skript wird jetzt abgebrochen."
																	: "Bitte nur Zahlen, Komma und Leerzeichen eingeben und keine '"
																	. (TelNumbers := RegExReplace(TelNumbers, "[^\d,]")) "'!"
			If !TelNumbers
				ExitApp
			goto Label_TelNumbers
		}
		TelNumbers 	:= RegExReplace(TelNumbers, ",", "|")


		Firstinitialisation := true
		ini := FileOpen(Addendum.fboxini , "w", "UTF-16")
		ini.WriteLine("; ############################################################")
		ini.WriteLine("; ")
		ini.WriteLine(";                                                  Einstellungen für den AHK Fritzbox Anrufmonitor")
		ini.WriteLine("; ")
		ini.WriteLine("; ############################################################")
		ini.WriteLine("[Allgemein]")
		ini.WriteLine("Version=" Addendum.FAMVersion)
		ini.Close()

	  ; Einstellungen speichern
		IniWrite, % fboxIP      	, % Addendum.fboxini, Fritzbox, 	IP
		IniWrite, % localprefix  	, % Addendum.fboxini, Telefone, 	lokale_Vorwahl
		IniWrite, % TelNumbers	, % Addendum.fboxini, Telefone, 	Ueberwachen

	}

; Vorwahlverzeichnis der Bundesnetzagentur herunterladen, in SQLite DB umwandeln
	BNZAcsv := A_ScriptDir "\NVONB.INTERNET.20221122.OZRNB.csv"
	BNZASQlite := A_ScriptDir "\NVONB.INTERNET.db"

;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; Einstellungen
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{

  ; Fritzbox-IP
	If !fboxIP
		IniRead, fboxIP	, % Addendum.fboxini, Fritzbox, IP
	Addendum.fboxIP := InStr(fboxIP, "ERROR") || !fboxIP ? "192.168.100.1" : RegExReplace(fboxIP, "[^\d\.]")
	If (Addendum.fboxIP != fboxIP)
		IniWrite, % Addendum.fboxIP, % Addendum.fboxini, Fritzbox, IP

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

  ; Telefonnummern aller Gerätenamen die "Fax" enthalten ermitteln
	Addendum.Fax := ""
	IniRead, AlleGeraete, % Addendum.fboxini, andere_Geraete
	For each, line in StrSplit(AlleGeraete, "`n", "`r") {
		Geraet      	:= Trim(StrSplit(line, "=").1)
		GeraeteNr 	:= Trim(StrSplit(line, "=").2)
		If (Geraet ~= "i)Fax")
			Addendum.Fax .= (Addendum.Fax ? "" : "(") RegExReplace(GeraeteNr, "[^\d]") "|"
	}
	If Addendum.Fax
		Addendum.Fax := RTrim(Addendum.Fax , "|") ")"

  ; Vorwahlnummern Deutschland aufbereiten
	Addendum.prefixes := LoadCallPrefixes()

  ; lokale Ortsvorwahl auslesen
	If !localprefix
		IniRead, localprefix, % Addendum.fboxini, Telefone, lokale_Vorwahl
	Addendum.localprefix := RegExReplace(localprefix, "[^\d]")

  ; mit shunt ist die interne Nummer der Telefoniegeräte gemeint.
  ; um die shunt Nummer z.B. eines Anrufbeantworters zu erhalten muss die Sicherungsdatei der
  ; Fritzbox Callmonitors Strings gesichtet werden. Das Skript kann nur anhand dieser Nummern unterscheiden
  ; an welchem Gerät ein Anruf entgegen genommen wurde.
	IniRead, val, % Addendum.fboxini, Shunts
	shunts := Object()
	For index, shunt in StrSplit(val, "`n", "`r") {
		If RegExMatch(shunt, "(?<nr>\d+)\s*=\s*(?<name>[\pL\s\-\_]+)", shunt)
			shunts[shuntnr] := shuntname
	}
	If (shunts.Count() > 0) {
		Addendum.shunts := shunts
	}

	Fax := index := d := TelNumbers := line := localprefix := ""
	AlleGeraete := Geraet := GeraeteNr := Telblocked := fboxIP := ""
	val := shunt := shunts := d := d1 := d3 := ""

  ;}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; los gehts
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
  ; CallersDB laden (muss global sein)
	callersDB := new phonecallers(Addendum.AlbisDBPath, ">2010")


  ; Gui zeigen
	callGui()

  ; Anrufmonitor-Objekt initialisieren
	callmon	:= new Fritzbox_Callmonitor(Addendum.fboxIP	, { "savefilepath"                    	: "C:\tmp\CMStrings-" A_YYYY ".txt"
																							,	 "archivepath"                     	: Addendum.DBPath "\Telefon"
																							,	 "localprefix"                        	: Addendum.localprefix
															                             	, 	 "managingfunc"                	: "CallManager"
																							,	 "FritzMessageDisplayFunc" 	: "callGui_CMViewer"})

  ; Verbindung zur Fritzbox herstellen
	callmon.Connect()

  ; bei erfolgreicher Verbindung wird die Objektvariable .connected auf wahr gesetzt
	callmon._OnReceive("savepath: " callmon.savepath)

  ; Anrufe eines Tages anzeigen. Leer lassen für den aktuellen Tag oder eine Datum der Form YYYYMMDD (Jahr2stelligerMonat2stelligerTag)
	callGui_Load("")

  ; Neustart  um 0 Uhr
	func_call := Func("callGui_Exit").Bind(true)
	SetTimer, % func_call, % -1*TimerTime("00:00 Uhr")
	func_call := ""

return

  ; ein paar Hotkeys ;{
!Esc::
callmon.Disconnect()
ExitApp
^+a::
	callGui_Exit(true)
return
^+!r::
	callGui_Exit(true)
return
;}

;}


; 🕽🕻
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; 🕿 🕿  🕻 ☏  F  ☏  R  ☏  I  ☏  T  ☏  Z  ☏  B  ☏  O  ☏  X  ☏  🕿 🕿  ☏  C  ☏  A  ☏  L  ☏  L  ☏  M  ☏  O  ☏  N  ☏  I  ☏  T  ☏  O  ☏  R  ☏ 🕽    🕿 🕿
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
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

		; CM Stringsviewer function
			If IsFunc(options.FritzMessageDisplayFunc)
				this.FritzMessageDisplayFunc := options.FritzMessageDisplayFunc

		; Save received strings
		; Only the data of the completed communication is continuously stored in this file.
			If options.savefilepath {
				savefilepath := options.savefilepath
				SplitPath, savefilePath, fname, fpath
				If !InStr(FileExist(fPath), "D")
					throw A_ThisFunc ": this is no valid path <<" fpath ">>"
				this.savefilepath 	:= savefilepath
				this.savepath      	:= fpath
			}
			else
				this.savepath := A_Temp

		; Archive path
			If options.archivepath
				this.archivepath := FilePathCreate(options.archivepath) ? options.archivepath : ""

		; File path for the current communication data
			this.currentfilepath := this.savepath "\" A_YYYY A_MM A_DD "_FbxCMStrings.txt"

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
		this.Disconnect()
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
			this.CMStringReceived := true
			recv	:= rec.recvText()
			this.stringfile := FileOpen(this.currentfilepath, "a", "UTF-8")
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

		;		date       time          event                     conID             from           		      	to                 	?
		;	19.12.22 19:08:00;	RING					;     	0		;		01622768469 	;	88333          	; SIP0;
		;	19.12.22 19:08:02;	RING                	;    	1  	;   	09955561764 	;	674853         	; SIP1;
		;		date       time          event                     conID             shuntnr        		     	from              	?
		;	19.12.22 19:08:22;	CONNECT        	;		0   	;   	14                   	;	01622768469	;
		;	19.12.22 19:08:10;	CONNECT        	;		1   	;   	14                   	;	09955561764	;
		;	19.12.22 19:09:01;	DISCONNECT		;		1		;		0;
		;	19.12.22 19:22:22;	DISCONNECT		;		0		;		0;
		static dispfunc

		;recv := A_DD "." A_MM "." SubStr(A_YYYY, 3, 2) " " Substr("00" A_Hour, -1) ":"Substr("00" A_Min, -1) ":"Substr("00" A_Sec, -1) "; COMPLETED; " callerID ";"

	; connects a function to simultaneous display incoming message strings
		If this.ShowFritzMessages && IsFunc(this.FritzMessageDisplayFunc) {
			fn := this.FritzMessageDisplayFunc
			dispFunc := Func(fn)
		}
		If this.ShowFritzMessages && IsFunc(dispfunc)
			%dispfunc%(str)

	; RegExMatch to split the first parts of the incoming message
		RegExMatch(str, "Oi)(?<date>\d+\.\d+\.\d+)\s+(?<time>\d+:\d+:\d+);(?<event>\w+);(?<conID>\d+);", callh)
		call := {"date":callh.date, "time":callh.time, "event":callh.event, "conID":callh.conID}
		o := StrSplit(str, ";")
		Switch call.event {
			Case "RING":
				call.from           	:= o.4
				call.to                	:= o.5
				call.SIP               	:= o.6
			Case "CALL":
				call.shuntnr          	:= o.4
				call.from           	:= o.5
				call.to               	:= o.6
				call.SIP               	:= o.7
			Case "CONNECT":
				call.shuntnr          	:= o.4
				call.from           	:= o.5
			Case "DISCONNECT":
				call.duration        	:= o.4
		}

	return call
	}

	SaveCalls(call)                                                                             	{	;-- save call data as json string

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

		datestr := !datestr ? A_YYYY A_MM A_DD : InStr(datestr, ".") ? StrSplit(datestr, ".").3 StrSplit(datestr, ".").2 StrSplit(datestr, ".").1 : datestr
		If FileExist(this.loadfilepath := this.savepath "\" datestr "_FbxCMStrings.txt")
			calls := StrSplit(FileOpen(this.loadfilepath, "r", "UTF-8").Read(), "`n", "`r")

	return calls
	}

	ArchiveCalls(datestr:="")                                                              	{	;-- ##moves local CMStrings File to archive path

	  ; zurück falls Speicherpfade nicht angegeben oder nicht vorhanden sind
		If !this.savefilepath || !this.archivepath || !FilePathCreate(this.archivepath)
			return
	  ; feststellen welche Datei das jüngste Erstellungsdatum hat, diese Datei wird behalten
		datestr := !datestr ? A_YYYY A_MM A_DD : datestr

	}

	CallsFilename()                                                                         	{	;-- gibt den Namen der aktuellen CMStrings Datei zurück
	return this.savefilepath
	}

	GetNextCallsFilename(sDate, addDays)                                      	{ 	;-- gets filename for a archived CMStrings file

	  ; addDays - how many days to add (positive) and to substract if negative

		result := {"found":false}

		cDate := ConvertToDBASEDate(sDate) + 0
		CVYear := SubStr(cDate, 1, 4)
		Loop {

			cDate 	+= 	% addDays, days
			cDate 	:=	SubStr(cDate, 1, 8)
			cYear 	:= 	SubStr(cDate, 1, 4)

		  ; Abbruch wenn kein Verzeichnis für gespeicherte CMStrings gibt
			If (lYear != cYear) {	; prüft pro Jahr nur einmal
				lYear := cYear
				If !InStr(FileExist(Addendum.DBPath "\Telefon\" cYear), "D")
					return result
			}

		; wenn es die Datei gibt
			If FileExist((ffp := Addendum.DBPath "\Telefon\" cYear "\" cDate "_FbxCMStrings.txt"))
				return {	"found"      	: true
						, 	"fullfilepath"	: ffp
						, 	"newDate"		: ConvertDBASEDate(cDate)
						, 	"newGuiStr"	: DayOfWeek(cDate, "short", "yyyyMMdd") ", " ConvertDBASEDate(cDate)}

			;~ SciTEOutput(cDay)

		} until (A_Index = 400)	; nur als Notfallausstieg, falls in Endlosscheife

	return result
	}

}


; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; 🖁 🖁 🖁P 🖁H 🖁O 🖁N 🖁E 🖁C 🖁A🖁L 🖁L 🖁E 🖁R 🖁S 🖁 🖁 🖁 🖁 🖁A🖁N 🖁R 🖁U 🖁F 🖁E 🖁R 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁 🖁
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
class phonecallers                                                                                     	{	;-- lädt Telefonnummer-Daten aus verschiedenen Albis-Datenbank-Dateien

	; UEBARZT.dbf - weitere Telefonnummern

	__New(DBPath:="", opt:="")                          	{

	  ; Rückwärtssuche: 	"https://www.dasoertliche.de/?form_name=search_inv&ph=" telnumber
	  ; 							"http://www.google.de/search?num=0&q=%s" "$number"
	  ; Ver­zeich­nis über zu­ge­teil­te Ruf­num­mern­blö­cke:
	  ; Registrierte Anbieter. Welche Anbieter bei der Bundesnetzagentur registriert sind, können hier heruntergeladen werden
	  ;								"https://www.bundesnetzagentur.de/SharedDocs/Downloads/DE/Sachgebiete/Telekommunikation/Unternehmen_Institutionen/Nummerierung/"
	  ;	                 			. "Rufnummern/ONRufnr/Verz_zuget_ONB/Verzeichnis_Anbieter.pdf?__blob=publicationFile&v=17" --> zip Datei ~29MB gepackt

	  ; opt = ">2010"   - lädt alle Daten von Patienten deren letzter Behandlungstag nicht vor dem Jahr 2010 lagen.


	  ; Albis DBASE Datenbankenpfad
		this.AlbisDBPath := DBPath

	  ; lokale Vorwahl
		this.localprefix := Addendum.localprefix

	  ; Vorwahlverzeichnis Mobilfunk
	    this.mobilprefixes := {"de": { "(0151\d|0160|017[015])"                  	: "Telekom"
								        		, 	"(0152\d|0162|017[234])"                  	: "Vodaphone"
								    	     	, 	"(0157\d|0159\d|0163|017[6789])"	: "O2"
									          	, 	"015566"                                            	: "1&1 Drillisch"
												, 	"015888"                                           	: "TelcoVillage"
												,	"0118"                                                 	: "Auskunftsdienste"
												,	"0180"                                                 	: "Servicerufnummer"
												,	"0700"                                                 	: "Persönliche Rufnummer"
												,	"0800"	                                              	: "entgeltfreier Telefondienst"
												,	"0900"                                               	: "Premiumdienst"}}

	  ; Vorwahlverzeichnis Deutschland
		If !FileExist(fVorwahl := A_ScriptDir "\deutschland_vorwahl.json") {

			MsgBox, 0x1004, % A_ScriptName, % "Die Datei ('\deutschland_vorwahl.json') mit den deutschen`n"
													     		.	"Vorwahlnummern konnte nicht gefunden werden.`n"
												     			.	"Ohne die Daten dieser Datei ist eine Anruferidentifikation nicht möglich.`n"
													    		. 	"Erlauben Sie die Datei aus dem Internet zu laden?"
			IfMsgBox, No
				ExitApp
			DownloadLink := "https://github.com/Ixiko/Addendum-fuer-Albis-on-Windows/blob/master/Module/Fritzbox/deutschland_vorwahl.json"
			URLDownloadToFile, % DownloadLink, % fVorwahl
			If !FileExist(fVorwahl) {
				MsgBox, 0x1000, % A_ScriptName,  % "Die erforderliche Datei konnte nicht heruntergeladen werden.`n"
																		. "Daher kann das Skript nicht weiter ausgeführt werden.`n"
																		. "Der Download-Link zur Datei wird in die Zwischenablage`n"
																		. " kopiert, nachdem Sie auf Okay gedrückt haben."
				Clipboard := DownloadLink
				ClipWait, 3
				ExitApp
			}

		}

	  ; Vorwahlverzeichnis laden
		this.prefixes :=  cJSON.Load(FileOpen(fVorwahl, "r", "UTF-8").Read())

	  ; für Vorwahlnummererkennung rxString erstellen
		this.rxprefixes := "^(0)(15[1279]\d|16[023]|17[0-9]|15566|15888|118|180|700|800|900|" 				; ")"
		For prefindex, prefix in this.prefixes
			this.rxprefixes .= (prefindex>1 ? "|" : "") prefix.number
		this.rxprefixes .= ")"

	  ; Optionen für einen zeitliche Eingrenzung der Datenübernahme aus der PATIENT.dbf
	    RegExMatch(opt, "\>(?<ear>\d{4})", y)
		If (StrLen(year)>=4)
			this.ItemsSince := year SubStr("0101", 1, 8-StrLen(year))    ; füllt immer auf 8 Zeichen auf

	  ; Daten aus den jeweiligen Datenbanken laden
		UEBArzt        	:= this.LoadUEBArzt()           	; UEBArzt TEL TEL2 Fax
		PatTelNrs       	:= this.LoadPATTELNR()       	; PATTELNR.dbf PATNR TELNRNORM
		this.PatientsDB 	:= this.LoadPatients()          	; lädt Patientendaten

	  ; Daten aus PATIENT.dbf und PATTELNR.dbf zusammenführen
		For PatID, PatData in this.PatientsDB
			If IsObject(PatTelNrs[PatID])
				this.PatientsDB[PatID].TELNR := PatTelNrs[PatID]

	  ; Daten aus PATIENT.dbf und UEBARZT.dbf zusammenführen
		For UID, data in UEBArzt
			this.PatientsDB[UID] := data

		clipboard := cJSON.Dump(this.PatientsDB, 1)

	  ; Daten aus der Addendum FritzFaxbox Telefonnummern Datei
		this.FaxSender 	:= this.LoadFaxSender()

		;this.ExtrasDB	:= this.LoadExtrasDB()

	}

	LoadPatients()                                            	{      	; lädt Patienten-Daten

		infilter		:= ["NR", "GESCHL", "NAME", "VORNAME", "GEBURT", "MORTAL", "LAST_BEH"]
		outfilter		:= ["GESCHL", "NAME", "VORNAME", "GEBURT", "MORTAL"]

	  ; Daten auslesen
		tmpDB 	 	:= ReadPatientDBF(this.AlbisDBPath, infilter
													, "EMail=0 " (!this.ItemsSince ? "allData" : "")		; wenn ein Jahr übergebem wurde dann nicht die gesamten Daten laden
													, this.ItemsSince)   												;

		PatientDB	:= Object()
		For PatID, Pat in tmpDB {
			PatientDB[PatID] := Object()
			For index, key in outfilter
				If tmpDB[PatID][key]                                                                            	; soll keine "wertlosen" Schlüssel übernehmen
					PatientDB[PatID][key] := tmpDB[PatID][key]
		}
		tmpDB := ""

	return PatientDB
	}

	LoadPATTELNR()                                        	{       	; lädt nur Telefonnummern aus einer Albisdatenbank

		addo    	:= 0
		matches	:= Object()
		dbf       	:= new DBASE(this.AlbisDBPath "\PATTELNR.dbf")
		res        	:= dbf.OpenDBF()

		Loop % dbf.records {
			obj    	:= dbf.ReadRecord(["PATNR", "TELNRKLAR", "TELNRNORM"])
			TelNr	:= RegExReplace(obj.TELNRNORM, "[^\d\+]")
			TelNr	:= RegExReplace(TelNr, "^(\+|00)49", "0")
			addo  +=	  TelNr ~= "^(0|\+)" 	? 0 : 1
			TelNr 	:= (TelNr ~= "^(0|\+)" 	? "" : Addendum.localprefix) TelNr

		  ; Telefonnummer entfernen um Zusatzinfo zu extrahieren z.B. 01991234 564 48 Tochter
			tmpObj	:= {"nr":TelNr, "nfo":phonextrainfo(obj.TELNRKLAR)}

		  ; weniger als 5 Ziffern hat keine Telefonnummer
			If (StrLen(TelNr) > 4)
				If !IsObject(matches[obj.PATNR])
					matches[obj.PATNR] := [tmpobj]
				else {
			; vermeidet Telefonnummernduplikate beim selben Patienten
					telnrExist := false
					For each, data in matches[obj.PATNR]
						if (data.nr = tmpObj.nr) {
							telnrExist := true
							If tmpObj.nfo && !data.nfo
								data.nfo := tmpObj.nfo
							break
						}
					If !telnrExist
						matches[obj.PATNR].Push(tmpObj)
				}
		}

		res         	:= dbf.CloseDBF()
		dbf        	:= ""

	return matches
	}

	LoadUEBArzt()                                            	{       	; Telefon-/Faxnummern Überweisungsärzte aus UEBArzt.dbf

		static telkeys := ["TEL", "TEL2", "FAX"]

		Tel   	:= Object()
		dbf 	:= new DBASE(this.AlbisDBPath "\UEBARZT.dbf")
		res   	:= dbf.OpenDBF()

		Loop % dbf.records {

			obj  	:= dbf.ReadRecord(["ID", "ANREDE", "Titel", "Name", "VORNAME", "ORT", "FACH", "TEL", "TEL2", "FAX"])

		; keine Telefondaten - dann weiter
			If (!obj.TEL && !obj.TEL2 && !obj.FAX)
				continue

		; Telefondaten umwandeln in ein verarbeitbares Format
			oTEL := []
			For each, key in telkeys
				If obj[key]
					oTEL.Push({"nr"   	: this.formnumber(obj[key], obj.ORT)
								 ,	 "type"	: key
								 , 	 "nfo"	: phonextrainfo(obj[key])})

		; Daten zu dem Objekt hinzufügen
			If (oTEL.Count()>0) {
				Tel["U" obj.ID] := {"NAME"      	: Trim(obj.Name)
										, 	"VORNAME"	: Trim(obj.Vorname)
										, 	"ANREDE"   	: Trim(obj.ANREDE " " obj.TITEL)
										, 	"FACH"			: Trim(obj.FACH)
										, 	"ORT"        	: Trim(obj.Ort)
										,	"TelNr"   		: oTEL}

			; entfernt "wertlose" Schlüssel
				For key, val in Tel["U" obj.ID]
					If (key != "TelNr") {
						If (val = "")
							Tel["U" obj.ID].Delete(key)
					}
					else
						for tkey, tval in val
							if (tval = "")
								Tel["U" obj.ID].TelNr.Delete(tkey)

			}

		}

		res         	:= dbf.CloseDBF()
		dbf        	:= ""

	return Tel
	}

	LoadFaxSender()                                        	{      	; lädt Faxnummern aus Addendums FritzFaxbox

		IniRead, path, % Addendum.Ini, ScanPool, FritzFaxbox_Telefonbuch
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

	GetPrefixLocation(prefix)                            	{        	; Ortsnamen- oder Anbietersuche zur Vorwahl

		For mobil_prefix, mobil_provider in this.mobilprefixes.de
			If (prefix ~= mobil_prefix)
				return mobil_provider

		For pfxIndex, item in this.prefixes
			If (prefix = item.number)
				return item.Name

	return
	}

	GetLocationPrefix(location)                        		{       	; Vorwahl zum Ort ermitteln, wenn möglich
		return this.prefixes[location]
	}

	GetNameFromNumber(nr, mergenames=1)	{       	; Anruferidentifikation aus verschiedenen Quellen

	;                                                                                             Position
	; matches Struktur := Array([PatID (0 für nicht Patienten)               1
	;	    								, Telefon-/Faxnummer                         2
	;										, Patienten-/Firmenname                     3
	;                                     	, Datenquelle (Integer)])                       4

		matches := Array()
		nr := RegExReplace(nr, "^(\+|00)\d\d", "0")               	; +4940999888777 -> 040999888777
		nr := RegExReplace(nr, "[^\d]")
		If !nr
			return

	 ; Telefonnummern aus
	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 ; PATTELNR.dbf, UEBARZT.dbf und PATIENT.dbf
	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		For PatID, Pat in this.PatientsDB {
			For TelIndex, Tel in Pat.TELNR      {          	; Patiententelefonnummern heraussuchen
				If RegExMatch(nr, "^" Tel.nr "$") {   	 	; (TelNr && )

				  ; nur der erster Vorname wird verwendet
					VORNAME := StrSplit(PAT.VORNAME, " ").1

				  ; erster Treffer oder ein Zusammenlegen von Namen ist nicht gewünscht
					If (!matches.Count() || !mergenames)
						matches.Push([	PatID
											, 	Tel.Nr
											, 	(VORNAME ? VORNAME " " : "") . Pat.NAME . (Tel.nfo ? " (" Tel.nfo ")" : "")
											, 	0x1])                                                                                               	; 0x1 = Daten stammen aus DBase Datenbanken
					else
						For pIndex, pData in matches
							If InStr(pData.3, " " Pat.NAME) {
								If !Instr(pData.3, "&") && !Instr(pData.3, VORNAME)                                         	; verhindert das mehr als 2 Vornamen (Klaus & Bernd & Renate Müller)
									matches[pIndex].3 := VORNAME " & " matches[pIndex].3                        		; in der Ausgabe angezeigt werden, doch jede gefundene
								matches[pIndex].1 .= "|" PatID                                                                     		; Patienten-ID wird übernommen (der String ist später die ID des Anrufes)
								break
							}
				}

			}
		}

	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 ; _DB\Telefon\Telefonnummern.txt (Addendum)
	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		For callNr, obj in this.FaxSender {
			If RegExMatch(nr, "^" callNr "$")
				If !matches.Count()
					matches.Push([0, callNr, Trim(obj.sdr), 0x2])
				else {
					For each, item in matches
						item.4 := item.4 | 0x2
				}
		}

	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 ; Addendum.fboxini
	 ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		IniRead, callername, % Addendum.fboxini, Telefonbuch, % "T" "0" LTrim(nr, "0")
		callername := Trim(callername)

		If (callername := InStr(callername, "ERROR") ? "" : callername) {

			If InStr(callername, ",")
				callername := StrSplit(callername, ",", A_Space).2 " " StrSplit(callername, ",", A_Space).1

			If !matches.Count()
				matches.Push([0, nr, callername, 0x4])
			else {
				For each, item in matches
					item.4 := item.4 | 0x4
			}

		}

		if (matches.Count() > 1)
			SciTEOutput("matches>1:`n" cJSON.Dump(matches, 1))

	return matches
	}

	formnumber(telnr, location:="")                  	{        	; Telefonnummern formatieren (mit Vorwahl versehen)

	; entfernt alle Zeichen ausser Zahlen und Pluszeichen, entfernt danach die deutsche Ländervorwahl
		telnr  	:= plainnumber(telnr)
		telnr  	:= RegExReplace(telnr, "^(\+|00)49", "0")

	; und fügt Ortsvorwahlnummern hinzu, wenn möglich
		If !RegExMatch(telnr, this.rxprefixes, matchedprefix)
			If location && (locationprefix := this.GetLocationPrefix(location))
				telnr := locationprefix . LTrim(telnr, "0")
			else
				telnr := this.localprefix . LTrim(telnr, "0")

	return Trim(telnr)
	}

	removelocal(telnr)                                       	{
	return RegExReplace(telnr	, "^" this.localprefix "\/")
	}

	nicenumber(telnr, formatter:="/")                 	{        	; formatiert Telefonnummern schön
	return RegExReplace(telnr, this.rxprefixes "\" formatter "*", "$1$2" formatter)
	}

	niceremove(telnr, formatter:="/")                   	{        	; nicenumber und removelocal
		telnr := this.nicenumber(telnr, formatter)
	return this.removelocal(telnr)
	}
}

; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
; 💼  💼  💼  💼  💼  💼  💼  💼  💼 C 💼 A 💼 L 💼 L 💼 M 💼 A 💼 N 💼 A 💼 G 💼 E 💼 R 💼  💼  💼  💼  💼  💼  💼  💼  💼  💼  💼  💼
; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
callManager(call, load:=false)                                                                   	{	;-- behandelt die Verbindungsnachrichten der Fritzbox


	/* Funktionsbeschreibung / Objektverwendung

		- - - - - -
		Objekte:
		- - - - - -
		call         	: 	enthält die als key:value Objekt erhaltenen Daten der Fritzbox zum aktuellen Anrufvorgang

							.conID	= Zähler für alle offenen Verbindungen, DISCONNECT gibt die connection ID wieder frei
											der Zähler beginnt mit einer Null (aus techn. Gründen wird im Skript eine 1 addiert)

							.event 	= Anrufstatus    -- neue Verbindungen  --
															  (	RING              	=	Klingeln bei eingehenden Anrufen
																CALL             	=	ausgehendes Telefonat
																 -- nach Verbindungsaufbau --
																CONNECT    	= Telefonat wurde angenommen
																DISCONNECT	= Telefonat beendet)

							.from 	=	Nummer des Anrufers
							.to    	=	bei eingehenden Anrufen die vom Anrufer gewählte Telefonnummer
											bei ausgehenden Telefonaten die Nummer des Angerufenen

							.caller	=	enthält nach Identifizierung der Telefonnummer den Namen des Anrufer


		callstack  	:
		callerstack	:	enthält Daten der aktuellen Anrufe (Anrufe die nicht abgeschlossen wurden)
		callers     	:	Anruferidentifizierung. Enthält Daten nach erfolgreicher Identifiizerung einer Telefonnummer
		uphones  	:	Zähler unbekannte Telefonnummern
		phones 	:	Zähler für alle Anrufe

	 */

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -.
	; AUSGABE-SYMBOLE:                                                                                          	|
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -.
	;          es klingelt          	|                  FAX           	|          Anrufbeantworter      		|
	;              	🔔                 	|              	℻             	|              	  🆎                   	|
	;  ausgehendes Telefonat	|      es wird gesprochen 	|         Anruf ist beendet         	|
	;              	☎                	|              	📞             	|                   ✗| ✓                 	|  ✓ ✔
	;        Fax erhalten         	|      AB wurde abspielt   	|  Anrufer wurde nicht bedient	|
	;            	📥℻            	|          	    🆎            	|                	✗🔕                 	|
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -.

	global 	Callers, Day, hCV, phones, uphones, ophones, ringCount, callsCount, talktime
	global 	callManagerIsActive
	global	BLV, hLV1, hLV2, LVColors, aldo
	global 	callstack
	global 	callerstack
	static 	tables   	:= Object()
	static 	cMinit    	:= false
	static 	teltime 		:= 0

	If !cMinit {
		cMinit   	:= true
		phones 	:= Object()
	    uphones 	:= Object()									; unbekannte Telefonnummern
		ophones	:= Object()
		talktime  	:= ringCount := callsCount := 0
		aldo     	:= 0
		callstack   	:= Array()
		callerstack	:= Object()
	}

	callManagerIsActive := true

  ; AHK arrays starts with an index of 1
	conID := call.conID := call.conID + 1

  ; removes the connectionID-key/value
	call.Delete("conID")

	If (call.SIP = "SIP1")
		return

 ; execution depending on event
	Switch call.event {

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; RING    	 = 	eingehende Anrufe -> call.from
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "RING":

			call.Delete("event")
			call.blocked 	:= false
			call.ignore 	:= false

		  ; call.from und call.to erhalten die lokale Vorwahlnummer
		  ; Telefonnummer identifizieren
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If call.from {

				call.from :=  !(call.from ~= "^(0|\+)") ? callersDB.formnumber(call.from) : call.from
				callers := callersDB.GetNameFromNumber(call.from)
				for idx, caller in callers {
					t 	.= (t 	? "|" : "") 	caller.3                                                              	; .3 Name (ansonsten enthält t den/die Namen des Anrufers)
					p 	.= (p 	? "|" : "") 	caller.1                                                             	; .1 Patienten Nr (PatID) in Albis
				}

			}
			else
				call.from := "#"

		; Anrufe an bestimmte Nummern ignorieren (z.B. Fax)
		; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			call.blocked := false
			call.ignore := !RegExMatch(call.to, Addendum.TelNumbers) ? true : false       	; Anruf an bestimmte Telefoniegeräte z.B. an ein Faxgerät nicht zählen,
																																; wenn eine Nutzereinstellung dazu gemacht wurde
			device := Addendum.devices[call.to]                                                           	; Namen des angerufenen Gerätes feststellen

	    ; blockierte Telefonnummern werden nicht angezeigt
	    ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If RegExMatch(call.from, Addendum.Telblocked)                                            	; setzt flags um eine Verbindung später nicht zu zählen
				call.blocked := call.ignore := true

		 ; unbekannte Nummer aufnehmen und zählen
		 ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If (!t && !call.ignore && !call.blocked && call.from != "#")  {                         	; weder ignorierte noch blockierte Nummer (Unterschied?) call.from != "#" &&
				If !IsObject(uphones[call.from])
					uphones[call.from] := [uphones.Count()+1, 1]                                     	; .1 nummeriert die unbek. Anrufer durch
				else
					uphones[call.from].2 += 1                                                                 	; .2 zählt die Anrufe der unbek. Telefonnummer
			}

		  ; Anrufe zu ignorierten Geräten werden ebenso gezählt
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If call.ignore && !call.blocked                                                                     	; ignorierte aber nicht blockierte Nummern werden extra gezählt
				ophones[device] := !ophones.haskey(device) ? 1 : ophones[device]+1

		  ; zählt alle Anrufe, welche
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If !call.ignore && !call.blocked {
				ringCount ++
				If call.from
					phones[call.from] := !phones.haskey(call.from) ? 1 : phones[call.from]+1
			}

		  ; Daten zusammenfassen
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			call.caller     	:=  t  ? t	: call.from!="#" ? "unbekannte Nummer " (uphones[call.from].1) : "Nummer unterdrückt"
			call.callerID  	:= (p ? p	: call.from!="#" ? call.from : call.from="#" ? "#":"") 	; ist entweder die PatID oder die Telefonnummer des Anrufers
																																; oder # bei unterdrückten Telefonnummern
			cID              	:= (call.callerID ~= "^(0|\+)" ? "" : "#") call.callerID	            	; fügt ein Rautezeichen (#) bei PatID's hinzu

		  ; Debugging
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If tx
			  SciTEOutput("TelNr unbekannt: (" call.conID ") " call.from (call.callerID ? ", " call.callerID : ""))


		  ; 📞🔔🔕🆎
		  ; Anruf anzeigen
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			Gui, CV: Default
			If !call.ignore && !call.blocked
				If !IsObject(callerstack[cID]) {

					Gui, CV: ListView, Callers
					LVRow := !LV_GetCount() ? 1 : LV_GetCount()+1
					callstack[conID].DayRow := LVRow

				  ; virtuelle Tabelle (callerstack)                                   Spalte
					callerstack[cID] := [1, LVRow, 0, 0]                       	; 1    	= Anzahl der Anrufe
																								; 2     	= Listviewzeile
																								; 3     	= wer hat den Anruf entgegengenommen:
																								;		    	- 1: ein Mensch
																								; 	    		- 2: der AB
																								;	    		- 3: ausgehender Anruf (kein Rückruf)
																								;           	- 0: hat aufgelegt o. wartet auf Gesprächsannahme
																								; 4     	= Fax

				  ; Anruf in der Listview anzeigen
					Gui, CV: ListView, Callers
					LV_Add(""	, call.caller
									, callersDB.niceremove(call.from)
									, RegExMatch(plainnumber(call.to), Addendum.Fax) ? "((🔔)) ℻" : "((🔔))"
									, 1
									, call.time)

					BLV["Callers"].Row(LVRow, 0x99B898, 0x0)
					SendMessage, 0x102A, % LVRow-1,,, % "ahk_id " hLV1

				}
				else if IsObject(callerstack[cID]) {

				  ; Anrufzähler hochsetzen
					callerstack[cID].1 += 1

					Gui, CV: ListView, Callers
					LV_Modify(callerstack[cID].2, "Col3", "((🔔))")                     		; Status
					LV_Modify(callerstack[cID].2, "Col4", callerstack[cID].1)         	; Anrufe
					LV_Modify(callerstack[cID].2, "Col5", call.time)						; Uhrzeit

				}

			If tx
				FileAppend, % A_Hour ":" A_Min ":" A_Sec ":: " cJSON.Dump(callerstack, 1)
								. "`n- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -`n", % A_Temp "\chars.txt", UTF-8

			call.calls                     	:= callerstack[cID].1
			callstack[conID]         	:= call
			callstack[conID].Day  	:= DayRow

	;}

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; CALL        = 	ausgehender Anruf -> call.to
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "CALL":

			call.Delete("event")

		  ; call.from und call.to erhalten die lokale Vorwahlnummer
		  ; Anrufer identifizieren
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If call.to {

				call.to	:= !(call.to ~= "^(0|\+)") ? callersDB.formnumber(call.to) : call.to
				callers	:= callersDB.GetNameFromNumber(call.to)                                    	; Anrufer identifizieren
				for idx, caller in callers {
					t 	.= (t 	? "|" : "") 	caller.3                                                                  	; .3 Name (ansonsten enthält t den/die Namen des Anrufers)
					p 	.= (p 	? "|" : "") 	caller.1                                                                    	; .1 Patienten Nummern (PatIDs) in Albis
				}

			}
			else
				call.to := "#"

		 ; unbekannte Nummer aufnehmen und zählen
		 ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			If !t  {
				If !IsObject(uphones[call.to])
					uphones[call.to] := [uphones.Count()+1, 1]                                                 	; .1 nummeriert die unbek. Anrufer durch
				else
					uphones[call.to].2 += 1                                                                              	; .2 zählt die Anrufe der unbek. Telefonnummer
			}

		  ; Text zuordnen
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			call.caller         	:=  t  ? t	: call.to ? "unbekannte Nummer " (uphones[call.to].1) : "Nummer unterdrückt"
			call.callerID      	:= (p ? p	: call.to ? call.to: "")
			cID                  	:= (call.callerID ~= "^(0|\+|#)" ? "" : "#") call.callerID           	; cID ist nur für callerstack
			phones[call.to] 	:= !phones.haskey(call.to) ? 1 : phones[call.to]+1                  	; zählt die Anrufe der Telefonnummer
			callsCount ++                                                                                               	; zählt den Anruf

		  ; ausgehender Anruf
		  ; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
			Gui, CV: Default
			call.state := "☎"
			If !IsObject(callerstack[cID]) {                                                                         	; ist KEIN Rückruf  ➧ 🡀 🡆🡄

				Gui, CV: ListView, Callers
				LVRow := !LV_GetCount() ? 1 : LV_GetCount()+1
				callerstack[cID] := [1, LVRow, 3]
				callstack[conID].DayRow := LVRow
				Gui, CV: ListView, Callers
				LV_Add(""	, Trim(call.caller)                                                                        	; Name
								, "🡆" callersDB.niceremove(call.to)                                              	; Telefonnummer
								, "☎"                                                                                       	; Status
								, callerstack[cID].1                                                                      	; Anrufzähler
								, call.time)

			}
			else {                                                                                                            	; ist ein Rückruf

				callerstack[cID].1 += 1
				callerstack[cID].3  := 1
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "☎")
				LV_Modify(callerstack[cID].2, "Col4", callerstack[cID].1)
				LV_Modify(callerstack[cID].2, "Col5", call.time)

			}


			call.calls                  	:= callerstack[cID].1
			callstack[conID]        	:= call
			If IsObject(callerstack[cID]) {
				call.state := "🆎"
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "🆎")
				LV_Modify(callerstack[cID].2, "Col5", call.time)
				callerstack[cID].3 := 3 	; zurück gerufen

			}

	;}

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; CONNECT
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "CONNECT":

			cID := (callstack[conID].callerID ~= "^(0|\+|#)" ? "" : "#") callstack[conID].callerID  	; cID ist nur für callerstack
			callstack[conID].shuntnr := call.shuntnr
			callstack[conID].time  	:= call.time
			callstack[conID].connect := 1
			;~ SciTEOutput("C: " cID)
			to := plainnumber(callstack[conID].to)
			callerstack[cID].3 := 	callstack[conID].shuntnr=40   	? 2                               	 	; 2 = der Anrufbeantworter ging ran (siehe unter "Ring")
										: 	RegExMatch(to, Addendum.Fax) ? 4                                 		; 4 = Fax
										: 	callerstack[cID].3=3               	? 3                                    	; 3 = entgegengenommen
										: 	1
			Gui, CV: Default

		  ; Gespräch wurde angenommen und jetzt wird gesprochen
			If !callstack[conID].ignore && !callstack[conID].blocked {   ; 40 wäre der AB
				Gui, CV: ListView, Callers
				LV_Modify(callerstack[cID].2, "Col3", "📞" (callerstack[cID].3 	= 1            	? "___"
																			: callerstack[cID].3 	= 2            	? "🆎"
																			: callerstack[cID].3 	= 3         		? "☎"
																			: callerstack[cID].3 	= 4         		? "℻": ""))
				LV_Modify(callerstack[cID].2, "Col5", callstack[conID].time)
			}
			;~ else {
				;~ to := RegExReplace(call.to, "^" Addendum.localprefix)
				;~ Gui, CV: ListView, Day
				;~ LV_Modify(callstack[conID].Day, "Col3", RegExMatch(to, Addendum.Fax) ? "📞℻" : "")
				;~ LV_Modify(callstack[conID].Day, "Col6", call.time)
			;~ }
	;}

	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻
	; DISCONNECT ✓
	; ⸻⸻⸻⸻⸻⸻⸻⸻⸻⸻;{
		Case "DISCONNECT":

			If !IsObject(callstack[conID]) {
				SciTEOutput("kein callstack eintrag vorhanden: " conID )
				return
			}

			cID := (callstack[conID].callerID ~= "^(0|\+|#)" ? "" : "#") callstack[conID].callerID  	; cID ist nur für callerstack
			callstack[conID].duration := call.duration ? call.duration : 0

		  ; ignorierte/blockierte Telefonnummer
		  ; (callerstack[cID].3 - wenn 1 dann wurde mit dem Anrufer gesprochen, wenn 2 dann ging der AB ran, wenn 3 dann war  es ein ausgehendes Telefonat)
			If (callstack[conID].ignore && callerstack[cID].3~="(1|3)") {

				to := plainnumber(callstack[conID].to)
				Gui, CV: ListView, Day
				LV_Modify(callstack[conID].Day, "Col3", "✓" (RegExMatch(to, Addendum.Fax) 	? "℻"
																				: callerstack[cID].3 = "3"         		? "☎" : ""))
				LV_Modify(callstack[conID].Day, "Col4", duration(callstack[conID].duration))
				LV_Modify(callstack[conID].Day, "Col5", callerstack[cID].1)
				LV_Modify(callstack[conID].Day, "Col6", callstack[conID].time)
				callstack[conID] := "-"

			}

		  ; Anruf wurde entgegengenommen. Anrufe mit unterdrückter Telefonnummer werden ebenso in die 2.Liste verschoben
		  ; (cID.3 - wenn 1 dann wurde mit d. Anrufer gesprochen, wenn 2 dann ging der AB ran, wenn 3 dann war  es ein ausgehendes Telefonat)
			else If !callstack[conID].ignore && !callstack[conID].blocked {

				constr		:= 	callstack[conID]                                                	; connection string

			  ; Nummer des Anrufers oder des Angerufenen
				from := callersDB.nicenumber(constr.from)
				to := callersDB.nicenumber(constr.to)
				callNr 		:= 	callerstack[cID].3<=2 	?	    	callersDB.removelocal(from)
									:	callerstack[cID].3=3		?	"🡆" 	callersDB.removelocal(to)
																			:	     	callersDB.removelocal(to)
				mcallNr := plainnumber(StrReplace(callNr, "🡆"))
				mTime 	:= clocktoseconds(constr.Time)

			  ; Anruf angenommen
				If (callerstack[cID].3 = 3) {

					talktime += constr.duration

					Gui, CV: Default
					Gui, CV: ListView, Callers
					LV_Delete(LVRow := callerstack[cID].2)
					For each, item in callerstack
						If (item.2 > LVRow)
							item.2 -= 1

					Gui, CV: ListView, Day
					LV_Add(""	, Trim(constr.caller)
									, callNr
									, (RegExMatch(to, Addendum.Fax) 	? "℻"
												: callerstack[cID].3 = "1" 	? "✓"
												: callerstack[cID].3 = "2" 	? "✓🆎"
												: callerstack[cID].3 = "3" 	? "✓"
												: callerstack[cID].3 = "4" 	? "℻" : "✓🆎")
									, duration(constr.duration)
									, constr.calls
									, constr.time)

				;{ noch vorhandene Einträge aus vorherigen Anrufen entfernen
					;~ maxRows := LV_GetCount()
					;~ Loop % maxRows {                                                              	; von letzter zur ersten Reihe
						;~ row := maxRows-A_Index+1
						;~ LV_GetText(LVTelNr	, row, 2)
						;~ LV_GetText(LVCalls	, row, 4)
						;~ LV_GetText(LVTime	, row, 5)
						;~ LVTelnr := plainnumber(StrReplace(LVTelNr, "🡆"))
						;~ LVTime := clocktoseconds(LVTime)
						;~ If (mcallNr = LVTelNr && mTime >= LVTime)
							;~ LV_Delete(A_Index)

					;~ }
				;}

					If !load
						callmon.SaveCalls(callstack[conID])

					callstack[conID] := "-"
					callerstack.Delete(cID)

				}

			; DISCONNECT ohne CONNECT
				else {

					Gui, CV: Default
					Gui, CV: ListView, Callers
					;~ LV_GetText(LVName	, callerstack[cID].2, 1)
					;~ SciTEOutput(cID ", " LVName)
					LV_Modify(callerstack[cID].2, "Col3", "✗" (callerstack[cID].3=2 ? "🆎": callerstack[cID].3=2 ? "🔕" : ""))

				}
				constr := ""
		}
	;}


	}

	callManagerIsActive := false

	callGui_SBSetText()

}


; Gui
callGui()                                                                                                      	{	;-- die grafische Oberfläche

	;{ Variablen
	global Callers, Day, hCV, hLV1, hLV2, hCVT1, ontop, BLV := Object()
	global CVDate, CVBack, CVForward, CVhDate, CVhBack, CVhForward, CVProps,CVCMStr
	global TThwnd1, TThwnd2
	global LVNames 	:= ["Callers", "Day"]
	global colSizes  	:= {"Callers":[230,180,50,"50 Center",125], "Day":[230,180,50,50,"50 Center",70]}
	global LVColors	:= ["0x99B898", "0x355C7D"]

  ; Fenstergröße
	IniRead, winSize, % A_ScriptDir "\FritzboxCallMonitor.ini", % "sonstiges", % "Fensterposition_" Addendum.compname
	winSize	:=  InStr(winSize, "ERROR") || !winSize ? "" : winSize
	RegExMatch(winSize, "x\s*(?<X>\-*\d+)", win)
	RegExMatch(winSize, "y\s*(?<Y>\-*\d+)", win)
	winSize	:= (winX<0 ? "" : "x" winX) (winY<0 ? "" : " y" winY " " )

  ; andere Fensterparameter
	wplus      	:= 25
	LVWidth   	:= Object()
	LVOpt		:= "NoSortHdr AltSubmit gcallGui_LVHandler hwndhLV"
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
	gopt := " gcallGui_Handler", opt := " hwndCVhOnTop vCVOnTop" gOpt

	Gui, CV: new		, % "hwndhCV " ( ontop ? "AlwaysOnTop" : "")
	Gui, CV: Color    	, % "c" StrReplace(LVColors.2, "0x") , % "c" StrReplace(LVColors.1, "0x")
	Gui, CV: Margin	, 0 , 5

	Gui, CV: Font    	, s10 q5 bold, Segoe Script
	Gui, CV: Add     	, Text	  	, % "xm ym-3	           	cWhite  vCVBack hwndCVhBack"             	gOpt                         	, % " ⮪"
	Gui, CV: Font    	, s8 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "x+2 ym	w110  	cWhite  vCVDate hwndCVhDate Center"   	gOpt                            	, % "DD, 00.00.0000"
	Gui, CV: Font    	, s10 q5 bold, Segoe Script
	Gui, CV: Add     	, Text	  	, % "x+2 ym-3          	cWhite  vCVForward hwndCVhForward"  	gOpt                             	, % "⮫"

	LVNamesW := colSizes.Callers.2 + colSizes.Callers.3
	Gui, CV: Font    	, s10 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "x" colSizes.Callers.1 " ym w" LVNamesW " Left cWhite hwndhCVT1"                            	, aktuelle Anrufe
	Gui, CV: Add     	, Button  	, % "x" LVWidth["Callers"]+wplus-3*20-2*3 " y2 w20 h20 hwndCVhCMSTR vCVCMSTR " gOpt	, % "📜"
	Gui, CV: Add     	, Button  	, % "x" LVWidth["Callers"]+wplus-2*20-3 " y2 w20 h20 hwndCVhProps vCVProps " gOpt   	, % "⚙"
	Gui, CV: Add     	, Button  	, % "x+3 y2 w20 h20 hwndCVhOnTop vCVOnTop " gOpt                                               	, % (ontop ? "🔐" : "🔓")

	Gui, CV: Font    	, s10 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, ListView	, % "xm y+0 w" LVWidth["Callers"]+wplus 	" r15 " LVOpt "1  vCallers"                               	, Anrufer|Telefon|Status|Anrufe|Uhrzeit

	Gui, CV: Font    	, s10 q5 Normal, Segoe Script
	Gui, CV: Add     	, Text	  	, % "x" colSizes.Callers.1 " y+2 w" LVNamesW " Left cWhite"                                             	, angenommene Anrufe
	Gui, CV: Font    	, s10 q5 Normal, Futura Bk Bt
	Gui, CV: Add    	, ListView	, % "xm y+0 w" LVWidth["Day"]+wplus    	" r20 " LVOpt "2 vDay"                                    	, Anrufer|Telefon|Status|Dauer|Anrufe|Uhrzeit

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

  ; ToolTips anlegen
	AddTooltip(CVhDate     	, "für Datumswahl mit der linken Maustaste anklicken")
	AddTooltip(CVhBack     	, "Klicken um das Telefonprotokoll `ndes vorherigen Tages zu sehen")
	AddTooltip(CVhForward, "Klicken um das Telefonprotokoll `n des nächsten Tages zu sehen")
	AddTooltip(CVhOnTop	, "Fenster immer vorn anzeigen`n(Allways on top)")
	AddTooltip(CVhProps 	, "Programmeinstellungen`nFritz!Monitor Telefonbuch bearbeiten")
	AddTooltip(CVhCMStr	, "Nachrichten der Fritz!Box mitlesen")

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
	OnMessage(0x0200, "callGui_WM_MOUSEMOVE", 3)

return
CVGuiClose:
CVGuiEscape:
	WinMinimize, % "ahk_id " hCV
return
}

callGui_Load(datestr:="")                                                                           	{	;-- lädt Telefondaten aus gespeicherten Fritzbox-Dateien

	global CV, Calles, Day

	calls := callmon.LoadCalls(datestr)
	DayOfWeek := DayOfWeek(dayDate := (!dateStr ? A_YYYY A_MM  A_DD : datestr), "short", "yyyyMMdd")

	Gui, CV: Default
	GuiControl, CV:, CVDate, % DayOfWeek ", " SubStr(dayDate, 7, 2) "." SubStr(dayDate, 5, 2) "." SubStr(dayDate, 1, 4)
	;~ GuiControl, CV: -Redraw, Callers
	;~ GuiControl, CV: -Redraw, Day

	For callidx, callString in calls
		If callString {
			call := callmon.CallStringParser(callString)
			callManager(call, true)
		}

	;~ GuiControl, CV: +Redraw, Callers
	;~ GuiControl, CV: +Redraw, Day

}

callGui_Exit(onlyReload:=true)                                                                      	{	;-- Schließen der Gui

	global hCV

	wqs := GetWindowSpot(hCV)
	If (wqs.X>-8 && wqs.Y>-8)
		IniWrite, % "x" wqs.X " y" wqs.Y, % A_ScriptDir "\FritzboxCallMonitor.ini", % "sonstiges", % "Fensterposition_" Addendum.compname

	If !FileExist(statspath := A_ScriptDir "\telefon_statistik.ini")
		IniWrite, % "Anzahl Anrufe;versch.Telefonnummern; davon Unbekannt;ausgehende Anrufe;Gesprächszeit" , % statspath, _Datenstruktur

	IniWrite, % callGui_SBSetText(), % statspath, % A_YYYY "-" A_MM, % SubStr("0" A_DD, -1)

	If onlyReload
		Reload

ExitApp
}

callGui_Handler(hwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")           	{	;-- g-Label Buttonhandler

	global ontop, hCV, CVhOntop, CVOnTop, hLV1, hCVT1, CVBack, CVDate, CVForward, CVCMStr

	Critical

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; AlwaysOnTop On/Off
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If (A_GuiControl="CVOnTop")   	{

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

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Einstellungsfenster
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else if (A_GuiControl="CVProps") 	{
		callProperties()
	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Anzeige für Kommunikationsstrings der Firtzbox ein-/aussschalten
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else if (A_GuiControl="CVCMStr") 	{

		callmon.ShowFritzMessages := !callmon.ShowFritzMessages
		callGui_CMViewer(callmon.ShowFritzMessages ? "Show":"Close")

	}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Protokolldatum ändern
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	else {

		Gui, CV: Submit, NoHide
		GuiControlGet, tmp, CV:, CVDate
		RegExMatch(tmp, "\d+\.\d+\.\d+", CVDay)
		;~ SciTEOutput("ach das klickt doch: " A_GuiControl ", " htmp ", " tmp " , " CVDay)

	  ; Vor- / Zurück
		If (A_GuiControl~="(CVBack|CVForward)") {

			nextCalls := callmon.GetNextCallsFilename(CVDay, A_GuiControl = "CVBack" ? -1 : 1)

		}
	  ; Klick auf das Datum direkt
		else if (A_GuiControl="CVDate") {

		}

		If nextCalls.found {
			GuiControl, CV:, CVDate	, % nextCalls.newGuiStr
			GuiControl, CV: Enable		, % (A_GuiControl = "CVBack" ? "CVForward" : "CVBack")
			nextCalls := callmon.GetNextCallsFilename(nextCalls.newDate, A_GuiControl = "CVBack" ? -1 : 1)
			If !nextCalls.found
				GuiControl, CV: Disable, % A_GuiControl
		}

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

	If (StrLen(A_GuiControl)=0 || StrLen(EventInfo)=0 || !EventInfo)
		return

	Gui, CV: Default
	Gui, CV: ListView, % (clickedLV := A_GuiControl)
	LV_GetText(clkLVCaller	, (clickedLVRow := A_EventInfo), 1)
	LV_GetText(clkLVTelNr	, clickedLVRow, 2)
	MouseGetPos, mx, my

  ; Doppelklick
	If (A_GuiEvent = "DoubleClick") 	{
		clk := callGui_RenameCheck(clkLVCaller, clkLVTelNr)
		If !clk.RenameIsLocked
			callGui_Rename(EventInfo, A_GuiControl)
	}

  ; Kontextmenu
	else if (A_GuiEvent = "RightClick")	{

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
		;~ SciTEOutput(A_ThisFunc ": RenameIsLocked=" clk.RenameIsLocked ", " clk.PatID ", " clickedLV ", clkLVCaller= " clkLVCaller ", clkLVTelNr= " clkLVTelNr)

		NoTelNumber := false
		If (!clkLVTelNr ~= "\d") || (clkLVCaller ~= "i)Nummer\s+unterdrückt")
			NoTelNumber := true

		MenusAdded := 0
		If (clickedLV = "Callers")                           	{
			Menu, cmenu, Add, von Anruferliste nehmen       	, % funcCM3
			MenusAdded ++
		}
		If !clk.RenameIsLocked 	{
			Menu, cmenu, Add, Anrufernamen ändern          	, % funcCM4
			MenusAdded ++
		}
		If clk.PatID                                             	{
			funcCM5	:= Func("callGui_CM").Bind("opencase" 	, clickedLV, clickedLVRow, clk.PatID)
			Menu, cmenu, Add, Karteikarte öffnen                	, % funcCM5
			MenusAdded ++
		}
		If !NoTelNumber                                   	{
			Menu, cmenu, Add
			Menu, cmenu, Add, Telefonnr. kopieren                 	, % funcCM1
			Menu, cmenu, Add, Telefonnr. (nur Zahlen) kopieren	, % funcCM2
			MenusAdded ++
		}

		If MenusAdded
			Menu, cmenu, Show, % mx-20, % my

	}

  ; einfacher Klick
	else if (A_GuiEvent = "Normal") 	{
		If !LVHandlerProc
			callGui_MarkCallers(clickedLV, clickedLVRow)
	}



}

callGui_RenameCheck(Caller, TelNr)                                                        		{ 	;-- prüft ob eine Telefonnummer geändert werden darf

	global callersDB

	 RenameIsLocked := false
   If RegExMatch(Caller, "i)Nummer.*unterdrückt") 	{
	RenameIsLocked := true
	clkLVPatID := ""
	}
   else if !RegExMatch(Caller, "i)unbekannter.*Anruf") 	{
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

callGui_MarkCallers(LVclicked, rowclicked:=0)                                           	{	;-- Anrufer farbig markieren

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
		If !InStr(altLVState, "ä") {
			BLV[altLV].Row(A_Index, (teltel ? 0xD77800 : LVColors.1), (teltel ? 0xFFFFFF : 0x000000))
			SendMessage, 0x102A, % A_Index-1,,, % "ahk_id " (altLV = "Callers" ? hLV1 : hLV2)
		}

	}

	LVHandlerProc := false

}

callGui_CM(cmd, LVclicked:="", rowclicked:=0, clkPatID:="")                        	{	;-- Kontextmenuhandler

	Gui, CV: Default
	Gui, CV: ListView, % LVclicked
	MouseGetPos, mx, my

	If InStr(cmd, "copy")                                                  	{

		LV_GetText(TelNr, rowclicked, 2)
		Clipboard := TelNr := cmd = "plaincopy" ? RegExReplace(TelNr, "[^\d]") : TelNr
		ClipWait, 2
		ToolTip, % TelNr " kopiert", % mx-20, % my-20, 1
		SetTimer, TTipOff, -3000

	}
	else If (cmd = "rename")                                         	{

		callGui_Rename(rowclicked, LVclicked)

	}
	else if  InStr(cmd, "remove") && (LVclicked="Callers")	{

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

		itembesttime := !itembesttime ? LV_GetCount()+1 : itembesttime
		If !itemmoved
			LV_Insert(itembesttime,, Caller, TelNr, state,, calls, timestr)

		callGui_SBSetText()

		;recv := A_DD "." A_MM "." SubStr(A_YYYY, 3, 2) " " Substr("00" A_Hour, -1) ":"Substr("00" A_Min, -1) ":"Substr("00" A_Sec, -1) "; COMPLETED; " callerID ";"

	}

return

TTipOff:
	ToolTip,,,, 1
return
}

callGui_SBSetText(txt:="", sbpos:=1)                                                             	{	;-- Statuszeilenanzeige ändern

	global ringCount, phones, uphones, callsCount, talktime

	Gui, CV: Default

  ; Listenzähler
	Gui, CV: ListView, Callers
	CallersWait := LV_GetCount()
	Gui, CV: ListView, Day
	CallersTalked := LV_GetCount()

  ; Statustext erstellen
	txt := !txt ? (ringCount " Anruf" (ringCount ? "e":"") " von " phones.Count() " Telnr. "
					. (uphones.Count() ? uphones.Count() " Telnr. unbekannt. " : "")
					. "Anrufe: "
					. (CallersTalked ? CallersTalked " abgeschlossen. " : "")
					. (CallersWait ? CallersWait : "Keiner") " offen. "
					. (callsCount>0 ? callsCount 	" ausgehend. " : " ")
					. "Gesprächszeit: " (talktime>59 ? "~ ":" ") duration(talktime, 2)) : txt
	SB_SetText(txt , sbpos)

						;~ . (callsCount>0 ? callsCount 	" ausgehende" (callsCount=1 ? "r":"") " Anruf" (callsCount=1 ? "":"e") ". " : "")
						;~ . (CallersTalked ? CallersTalked " Anruf" (CallersTalked=1 ? "":"e") " abgeschlossen. " : "")
return ringCount ";" phones.Count() ";" uphones.Count() ";" callsCount ";" talktime
}

callGui_Rename(LVRow, LVName)                                                              	{	;-- Anruferbezeichnungen ändern

	global uphones, LVNames                   ; LVNames = Caller, Day

	MouseGetPos, mx, my

	Gui, CV: Default
	Gui, CV: ListView, % LVName

	LV_GetText(uNAME	, LVRow, 1)
	LV_GetText(uTEL    	, LVRow, 2)

	uTEL := RegExReplace(uTEL, "^\d+?🡆")

	InputBox	, callername, Fritzbox Anrufmonitor, % "Sie ändern: " uName " der Telefonnummer: " uTEL
					,, 300, 140,,,,, % (!RegExMatch(uName, "i)unbekannt.*Anruf") ? uName : "")

	If (RegExMatch(callername, "^\s*$") || RegExMatch(callername, "i)unbekannt.*Anruf") || callername = uName || ErrorLevel)
		return

  ; lokale Vorwahl  bei Bedarf hinzufügen
	telnr := callersDB.formnumber(uTEL)

  ; geänderten / oder neuen Namen speichern
	IniWrite, % callername, % Addendum.fboxini, Telefonbuch, % "T" plainnumber(telnr)  	; als reine Zahlenfolge speichern

   ; Telefonnummer aus uphones (Objekt für unbekannte Telefonnummern) entfernen
	uphones.Delete(TelNr)

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

  ; Statuszeile ändern
	callGui_SBSetText()

}

callGui_LVReplace(replacestring:="")                                                           	{	;-- Listvieweinträge ändern

	global CV, LVNames

	; replace Callers 	Name if Number is 01517453231 with Alfons Test - oder
	; replace Both   	Name if Number is 01517453231 with Alfons Test - oder
	; replace Both    	Number if Name is Alfons Test with 01893453434

	RegExMatch(replacestring	, "i)^\s*replace\s+"
											. "(?<LVName>\pL+)\s+if\s+"        	; Callers, Day, Both
											. "(?<ColName>\pL+)\s+is\+"			; Spaltenbezeichnung
											. "(?<MatchStr>.+?)\s+with\s+"    	; Muster für die Übereinstimmung
											. "(?<With>.+)")                             	; Ersatztext

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

callGui_WM_MOUSEMOVE(wParam, lParam, Msg, Hwnd)                          	{	;-- Listview Tooltips

   ; LVM_HITTEST   -> docs.microsoft.com/en-us/windows/desktop/Controls/lvm-hittest
   ; LVHITTESTINFO -> docs.microsoft.com/en-us/windows/desktop/api/Commctrl/ns-commctrl-taglvhittestinfo
	;  ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡ ℡
	;🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄
	;          es klingelt          	🟄                FAX             	|          Anrufbeantworter      		|
	;              	🔔                 	🟄              	℻             	|                	  🆎                    	|
	;🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄
	;  ausgehendes Telefonat	🟄      es wird gesprochen 	|         Anruf ist beendet         	|
	;              	☎                	🟄             	📞             	|                  ✗|✓                   	|
	;🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄
	;        Fax erhalten         	🟄      AB wurde abspielt   	|  Anrufer wurde nicht bedient	|
	;            	📥℻            	🟄         	    📤            	|                	✗🔕                 	|
	;🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄 🟄
	; ✆   📞   🕼    🕽    🕻   🕾   🕿    ☏     ☎   🖁   ℻   ℡  🖷

   global colSizes, hLV1, hLV2, hCV, CPUNBLK, CPhUNBLK, CPTX1, phone_prefix, clippi
   global callstack, callerstack
   static 	TelefonX1, TelefonX2, StatusX1, StatusX2, LVy1, LVy2, cell_TelNr
   static  cell_CallStatus, msgRuns := false, Xinit := true
   static zaehldazu := 0
   static  statusmsg := {"℻"    	: "Fax"
					    		, "🔔"    	: "Es klingelt."
					    		, "📞"    	: "im Gespräch"
					    		, "☎"    	: "ausgehendes Telefonat"
					    		, "✗"     	: "Anruf wurde nicht angenommen"
					    		, "✗🔕" 	: "Anrufer wurde nicht bedient`nund hat aufgelegt"
					    		, "✗🆎" 	: "Anrufer wurde nicht bedient.`nDer AB wurde abgespielt."
					    		, "✓"	    	: "Anruf wurde entgegen genommen."}

	; 0x2717, 0xD83D, 0xDD15

   CoordMode, ToolTip, Screen
   If Xinit {
		colW := []
		For each, column in colSizes.Callers {
			RegExMatch(column, "\d+", colWidth)
			colW.Push(colWidth)
		}
		TelefonX1 :=                 	colW.1
		TelefonX2 := TelefonX1 + colW.2
		StatusX1	:= TelefonX2	+ 1
		StatusX2	:= StatusX1	+ colW.3
		Xinit := false
   }

    MouseGetPos, mx, my, mWin, mCtrl, 3
	If (msgRuns || !hwnd || (hwnd != hLV1 &&  hwnd != hLV2 && hwnd != CPhUNBLK && hwnd != mCtrl))
		return
	msgRuns := true

	Critical

	; ToolTips mit zusätzlichen Informationen in der Anruferliste
	If (A_GuiControl ~= "i)(Callers|Day)") {

	  Gui, CV: Default
	  VarSetCapacity(LVHTI, 24, 0)                                                                                                 	; LVHITTESTINFO
	  NumPut(lParam & 0xFFFF, LVHTI, 0, "Int"), NumPut((lParam >> 16) & 0xFFFF, LVHTI, 4, "Int")
	  Item := DllCall("SendMessage", "Ptr", Hwnd, "UInt", 0x1012, "Ptr", 0, "Ptr", &LVHTI, "Int")           	; LVM_HITTEST

	  If (Item >= 0) && (NumGet(LVHTI, 8, "UInt") & 0x0E) {                                                          	; LVHT_ONITEM

		  ; Inhalt der Spalte 3 und 4 lesen
			Gui, CV: ListView, % A_GuiControl
			LV_GetText(cell_TelNr    	, row:=Item+1, 2)
			LV_GetText(cell_CallStatus	, row, 3)

		  ; X Position des Mauszeigers in der Gui
			X := lParam & 0xFFFF

		  ; Mauszeiger in Spalte Telefon oder Status dann weiter
			If (TTipFor :=  (X >= TelefonX1 && X <= TelefonX2)	? 1
							:	  (X >= StatusX1 	&& X <= StatusX2)  	? 2 : 0) {

			  ; Bildschirm und Fensterposition
				cellPos	:= LV_EX_GettemRect(hwnd, TTipFor+1, row, 3)                                             	; , Spalte  	-> TTipFor+1 ergibt genau die gewünschte Spalte
																																					; , Zeile   	-> ist die Variable row
																																					; , LVIR=3 -> LVIR_LABEL
																																					;					Gibt das umgebende Rechteck des gesamten
																																					;					Elements zurück, einschließlich des Symbols
																																					;					und der Beschriftung.
				CV    	:= GetWindowSpot(hCV) , LV := GetWindowSpot(Hwnd)
				LVy     	:= LV.Y + cellPos.Y - cellPos.H

			  ; Nachrichten erstellen
				If (TTipFor=1) {
					If RegExMatch(cell_TelNr, "\d+(?=\/)", prefix)                                                          ; der Ortsname zur Vorwahlnummer wird ermittelt
						TT(callersDB.GetPrefixLocation(prefix), CV.X+TelefonX1, LVy)
					else if !(prefix ~= "(#)")
						TT(callersDB.GetPrefixLocation(Addendum.localprefix), CV.X+TelefonX1, LVy)
				}
			; Eingehender oder ausgehender Anruf
				else If (TTipFor=2) {
					cID := (callstack[conID].callerID ~= "^(0|\+|#)" ? "" : "#") callstack[conID].callerID  	; cID ist nur für callerstack
					For callstatus, msgpart in statusmsg
						TTmsg .= callstatus = cell_CallStatus ? msgpart : ""
					TTmsg .= "`nStatus: " callerstack[cID].3
					StrReplace(TTmsg, "`n",, EOL)
					TT(RTrim(TTmsg, "`n"), CV.X+StatusX1, LVy - (EOL*14))
				}

			}
	  }


   }

  ; on mouse hover die Schrift fetter machen
	If (A_GuiControl = "CPUNBLK" || mCtrl = CPhUNBLK) {
		Gui, CP: Default
		GuiControlGet, status, Enabled, % CPhFBIP1
		If status {
			Critical Off
			return
		}
		GuiControl, CP:, CPTX1     	, % "Fritzbox IP!"
		GuiControlGet, ctxt, CP:, % CPhUNBLK
		Gui, CP: Font, s11 Bold, Futura Bk Bt
		GuiControl, CP:, % CPhUNBLK, % ctxt
		SetTimer, CPBoldOff, -500
	}

	msgRuns 	:= false
	cell_TelNr := cell_CallStatus := ""
	Critical Off

return

CPBoldOff:
	Gui, CP: Default
	GuiControl, CP:, CPTX1     	, % "Fritzbox IP"
	GuiControlGet, ctxt, CP:, % CPhUNBLK
	Gui, CP: Font, s11 Normal, Futura Bk Bt
	GuiControl, CP:, % CPhUNBLK, % ctxt
return

}

TT(msg:="", X:="", Y:="", duration:=500) {

	ToolTip, % msg, % X, % Y, 1
	SetTimer, WM_TT_Off, % -1*duration

return
WM_TT_Off:
	ToolTip,,,,1
return
}

callGui_ScriptVars()                                                                                     	{

	ListVars

}


; Einstellungen
callProperties()                                                                                             	{	;-- Einstellungen vornehmen, Telefonbuch bearbeiten

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Variablen
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	global LVColors, hCV, hCP, ontop, CPVars, CPFBIP1, CPFBIP2, CPFBIP3, CPFBIP4, CPhFBIP1, CPhFBIP2, CPhFBIP3, CPhFBIP4
	global CPUNBLK, CPhUNBLK, CPTX1, CPTELD, CPBLKTEL, CPLCPFX, CPLVTB
	static CPTMP, CPTX2, CPTX3, CPTX4, CPTX5
	static CPSAVE, CPCANCEL, CPhLVTB, CPPG1, CPPG2
	static TelNumbers, Telblocked
	static needSave := false
	static PBook
	gth := " hwndCPhTmp", gh := " gcallProps_Handler"

  ; Telefonbuchdaten lesen
	CPVars := Object()
	CPVars.PBook := Object()
	IniRead, tmp, % Addendum.fboxini, Telefonbuch
	If RegExMatch(tmp, "T\d+\s*\=\s*\pL")
		For each, line in StrSplit(tmp, "`n", "`r")
			CPVars.PBook["T" StrReplace(StrSplit(line, "=").1, "T")] := StrSplit(line, "=").2
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Gui Start
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	Gui, CP: new		, % "hwndhCP " ( ontop ? "AlwaysOnTop" : "")
	Gui, CP: Color    	, % StrReplace(LVColors.2, "0x")                                                      	, % "c" StrReplace(LVColors.1, "0x")
	Gui, CP: Margin	, 10 , 5
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Titel
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	Gui, CP: Font    	, s14 q5 bold, Segoe Script
	Gui, CP: Add     	, Picture  	, % "x100 ym	 w24 h-1  "                                                 	, % A_ScriptDir "\assets\Fritzboxanrufmonitor.ico"
	Gui, CP: Add     	, Text	  	, % "x+5        cWhite  "                                                  	, % "Fritz!Anrufmonitor " Addendum.FAMVersion
	Gui, CP: Add     	, Picture  	, % "x+2   	 w24 h-1  "                                                	, % A_ScriptDir "\assets\Fritzboxanrufmonitor.ico"
	Gui, CP: Add     	, Progress	, % "xm y+3	 w470 h2 cWhite cCPPG1"                         	, 100
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Fritzbox IP
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	Gui, CP: Font    	, s10 q5 Bold, Segoe Script
	Gui, CP: Add     	, Text	  	, % "xm y+15 cWhite  vCPTX1"                                          	, % "Fritzbox IP "
	Gui, CP: Font    	, s11 q5 Normal, Futura Bk Bt
	Gui, CP: Add     	, Text	  	, % "x+5 w18 h18 cWhite vCPUNBLK hwndCPhUNBLK " gh 	, % "🔒 "  ;  🔓
	CPVars.fboxIP := Addendum.fboxIP
	Loop 4 {
	  ; IP Felder
		Gui, CP: Font   	, s10 q5 Normal cBlack, Futura Bk Bt
		Gui, CP: Add   	, Edit 	  	, %  (A_Index=1 ? "xm y+2 ":"x+1") " w30 h18 r1 vCPFBIP" A_Index " Limit3 Number "
												. 	  "hwndCPhFBIP" A_Index . gh, % StrSplit(Addendum.fboxIP, ".")[A_Index]
		hwnd := % "CPhFBIP" A_Index
		hwnd := %hwnd%
		WinSet, Style 	, 0x50012000, % "ahk_id " hwnd
		WinSet, ExStyle	, 0x00000000, % "ahk_id " hwnd
		Control, Disable,,, % "ahk_id " hwnd
		GuiControl, CP: Move, % "CPFBIP" A_Index, % "h18"
	  ; Punkt
		Gui, CP: Font  	, s10 q5 Normal cWhite, Futura Bk Bt
		If (A_Index != 4)
			Gui, CP: Add, Text 	  	, % "x+1  "                                                                     	, % "."
	}
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Ortsvorwahl
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	CPVars.localprefix := Addendum.localprefix
	GuiControlGet, cp	, CP: Pos, CPFBIP1
	GuiControlGet, dp, CP: Pos, CPFBIP4
	width := dpX+dpW-cpX
	Gui, CP: Font    	, s10 q5 Bold, Segoe Script
	Gui, CP: Add     	, Text	  	, % "xm y+9 cWhite  vCPTX2"                                          	, % "lokale Ortsvorwahl"
	Gui, CP: Font    	, s10 q5 Normal cBlack, Futura Bk Bt
	Gui, CP: Add    	, Edit 	  	, % "y+1 w" width " r1 vCPLCPFX" gth                               	, % Addendum.localprefix
	GuiControl, CP: Move, % "CPLCPFX" 	, % "h18"
	WinSet, Style 	, 0x50012000, % "ahk_id " CPhTmp
	WinSet, ExStyle	, 0x00000000, % "ahk_id " CPhTmp
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; überwachte Telefone
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	; ueberwachen, blockierte_Anrufer, lokale_Vorwahl
	CPVars.TelNumbers := RegExReplace(Addendum.TelNumbers, "[\(\)]")
	GuiControlGet, cp, CP:Pos, CPTX1
	Gui, CP: Font    	, s10 q5 Bold, Segoe Script
	Gui, CP: Add     	, Text	  	, % "x+15 y" cpY " w150 Center cWhite vCPTX3"               	, % "überwachte Telefone"
	Gui, CP: Font    	, s10 q5 Normal cBlack, Futura Bk Bt
	Gui, CP: Add    	, Edit 	  	, % "y+1 w150 r4 vCPTELD" gth                                       	, % RegExReplace(StrReplace(Addendum.TelNumbers, "|", "`n"), "[\(\)]")
	WinSet, Style 	, 0x50013000, % "ahk_id " CPhTmp
	WinSet, ExStyle	, 0x00000000, % "ahk_id " CPhTmp
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; blockierte Telefonnummern
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
    CPVars.Telblocked := RegExReplace(Addendum.Telblocked, "[\(\)]")
	For each, telnr in StrSplit(CPVars.Telblocked, "|")
		Telblocked .= (Telblocked ? "`n" : "") nicenumber(telnr)
	Gui, CP: Font    	, s10 q5 Bold, Segoe Script
	Gui, CP: Add     	, Text	  	, % "x+15 y" cpY " w150 Center cWhite  vCPTX4"          	, % "blockierte Anrufer"
	Gui, CP: Font    	, s10 q5 Normal cBlack, Futura Bk Bt
	Gui, CP: Add    	, Edit 	  	, % "y+1 w150 r4 vCPBLKTEL" gth                                	, % Telblocked
	WinSet, Style 	, 0x50013000, % "ahk_id " CPhTmp
	WinSet, ExStyle	, 0x00000000, % "ahk_id " CPhTmp
	;}

  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  ; Fritz!Anrufmonitor Telefonbuch
  ; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -;{
	GuiControlGet, cp, CP:Pos, CPLCPFX
	GuiControlGet, dp, CP:Pos, CPBLKTEL
	width := dpX+dpW-cpX
	GuiControl, CP: Move, % "CPPG1" 	, % "w" (guiCW := width)

	Gui, CP: Font    	, s12 q5 Bold, Segoe Script
	Gui, CP: Add     	, Text	  	, % "xm y+10 w" width " Center cWhite  vCPTX5"          	, % "Fritz!Anrufmonitor Telefonbuch"
	Gui, CP: Add     	, Progress	, % "xm y+0	 w" width " h2 cWhite cCPPG2"                     	, 100
	Gui, CP: Font    	, s10 q5 Normal cBlack, Futura Bk Bt

	; kein Headerdragdrop
	LVOpt := "LV0x10000128 -LV0x10 hwndCPhLVTB"
	Gui, CP: Add     	, ListView 	, % "xm y+1 w" width " r20 vCPLVTB " LVOpt gth           	, % "Name|Telefonnummer"
	WinSet, Style 	, 0x50010205, % "ahk_id " CPhTmp
	WinSet, ExStyle	, 0x00000000, % "ahk_id " CPhTmp
	For pnr, pcaller in CPVars.PBook
		LV_Add("", pcaller, nicenumber(StrReplace(pnr, "T")))
	LV_ModifyCol(1, guiCW-160-25)
	LV_ModifyCol(2, 160)
	;}

	Gui, CP: Add     	, Button	  	, % "xm y+10  vCPSAVE "     	gh                                  	, % "Änderungen speichern"
	Gui, CP: Add     	, Button	  	, % "x+25      vCPCANCEL "	gh                                 	, % "Abbruch"

	Gui, CP: show		, w600 AutoSize , % "Fritz!Anrufmonitor Einstellungen"

	AddTooltip(CPhUNBLK 	, "Hier klicken um die IP Eingabe freizuschalten.")
	AddTooltip(CPhLVTB 		, "F2 oder Doppelclick in der`n1.Spalte um Namen zu ändern.")

   ; WM_MOUSEMOVE handler
	OnMessage(0x0200, "callGui_WM_MOUSEMOVE", 3)

return

CPGuiClose:
CPGuiEscape:
	If !needSave
		Gui, CP: Destroy
return
}

callProps_Handler(hwnd:="", GuiEvent:="", EventInfo:="", ErrLevel:="")          	{	;-- gHandler für callProperties

	global CPVars

	Switch A_GuiControl {

		Case "CPUNBLK":
			callProps_IPStatus()
			return

		Case "CPCANCEL":
			callProps_IPStatus("Disable")
			GuiVars := callProps_GetAndCheck()
			If GuiVars.Count() {
				SciTEOutput(GuiVars.Count() ": " cJSON.Dump(GuiVars,1))
				MsgBox, 0x1004, % StrReplace(A_ScriptName, ".ahk"), % "Wollen Sie die vorgenommenen Änderungen speichern?", 5
				IfMsgBox, Yes
					save:=1
			}
			Gui, CP: Destroy
			return

		Case "CPSAVE":
			callProps_IPStatus("Disable")
			Gui, CP: Destroy


	}

}

callProps_IPStatus(status:="")                                                                        	{	;-- aktiviert/inaktiviert die IP Eingabefelder

	global CPhFBIP1, CPhFBIP2, CPhFBIP3, CPhFBIP4, CPUNBLK, CP

	Gui, CP: Default
	If !status {
		GuiControlGet, status, Enabled, % CPhFBIP1
		status := status ? "Disable" : "Enable"
	}
	Loop 4 {
		hwnd := % "CPhFBIP" A_Index
		hwnd := %hwnd%
		Control, % status,,, % "ahk_id " hwnd
	}
	GuiControl, CP:, CPUNBLK, % (status="Enabled" ? "🔓" : "🔒")

}

callProps_GetAndCheck()                                                                           	{	;-- prüft auf Änderungen in den Eingabefeldern

	global CP, CPVars, CPFBIP1, CPFBIP2, CPFBIP3, CPFBIP4, CPLCPFX, CPTELD, CPBLKTEL, CPLVTB

	Gui, CP: Default
	Gui, CP: Submit, NoHide

	changed := 0
	GuiVars := CPVars.Clone()

  ; Fritzbox IP
	ip := CPFBIP1 "." CPFBIP2 "." CPFBIP3 "." CPFBIP4
	If (GuiVars.fboxIP = ip)
		GuiVars.Delete("fboxIP")
	else
		GuiVars.fboxIP := ip

  ; Vorwahl
	If (CPVars.localprefix = CPLCPFX)
		GuiVars.Delete("localprefix")
	else
		GuiVars.localprefix := CPLCPFX

  ; überwachte Telefonnummern
	ctxt := RegExReplace(CPTELD, "(\n{2,}|\n\s*\n)", "`n")
	ctxt := RegExReplace(ctxt, "\n", "|")
	ctxt := RegExReplace(ctxt, "[^\d\|]", "")
	ctxt := RTrim(ctxt, "|")
	If (GuiVars.TelNumbers = ctxt)
		GuiVars.Delete("TelNumbers")
	else
		GuiVars.TelNumbers := ctxt

  ; blockierte Telefonnummern
	ctxt := RegExReplace(CPBLKTEL, "(\n{2,}|\n\s+\n)", "`n")
	ctxt := RegExReplace(ctxt, "\n", "|")
	ctxt := RegExReplace(ctxt, "[^\d\|]", "")
	ctxt := RTrim(ctxt, "|")
	If (GuiVars.Telblocked = ctxt)
		GuiVars.Delete("Telblocked")
	else
		GuiVars.Telblocked := ctxt

  ; Telefonbuch
	Gui, CP: ListView, CPLVTB
	Loop % LV_GetCount() {

		LV_GetText(pcaller	, A_Index, 1)
		LV_GetText(pnr   	, A_Index, 2)
		pnr := "T" RegExReplace(pnr, "[^\d]")

		If (GuiVars.PBook[pnr] != pcaller)
			GuiVars.PBook[pnr] := pcaller
		else
			GuiVars.PBook.Delete(pnr)

	}

	If !GuiVars.PBook.Count()
		GuiVars.Delete("PBook")


return GuiVars
}

LoadCallPrefixes()                                                                                      	{

  ; Vorwahlnummern Deutschland aufbereiten
	If FileExist(fVorwahl := A_ScriptDir "\deutschland_vorwahl.json") {
		tmp := cJSON.Load(FileOpen(fVorwahl, "r", "UTF-8").Read())
		For index, prefix in tmp
			prefixes .= (index>1 ? "|" : "") prefix.number
		return prefixes
	}

return
}


; Kommunikation mitlesen
callGui_CMViewer(msg:="")                                                                      		{

	global LVColors, ontop, hCV, hCM, CMhStr, CMStr, CMFile
	static fileIs, fileIsTime
	static CMWidth := 360

	SciTEOutput("1 callmon.ShowFritzMessages: " msg )

  ; neue Nachricht hinzufügen
	If !(msg ~= "i)^\s*(Show|Close)") {
		Edit_Append(CMhStr, msg)
		return
	}

  ; schließe Fenster
	If (msg = "Close" && WinExist("Fritzbox Nachrichten ahk_class AutoHotkeyGUI")) {
		Gui, CM: Show, Hide
		callmon.ShowFritzMessages := false
		return
	}

  ; zeige Fenster
	If (msg = "Show") {

		CVWin := GetWindowSpot(hCV)

		If !WinExist("Fritzbox Nachrichten ahk_class AutoHotkeyGUI") {

			Gui, CM: new		, % "hwndhCM " ( ontop ? "AlwaysOnTop" : "")
			Gui, CM: Color    	, % "c" StrReplace(LVColors.1, "0x"), % StrReplace(LVColors.2, "0x")
			Gui, CM: Margin	, 5 , 5


			Gui, CM: Font    	, s8 q5 	, Segoe Script
			Gui, CM: Add     	, Text	  	, % "xm ym w" CMWidth " cBlack vCMFile Backgroundtrans "                 	, % ""
			GuiControlGet, cp, CM: Pos, CMFile

			CMWin := GetWindowSpot(hCM)

			Gui, CM: Font    	, s8 q5 Normal cWhite, Consolas
			h := CVWin.CH - cpH - 4*CMWin.BH + 1
			Gui, CM: Add    	, Edit 	  	, % "y+1 w" CMWidth  " h" h " vCMStr HwndCMhStr "      	, % ""
			WinSet, Style     	, 0x50213840, % "ahk_id " CMhStr
			WinSet, ExStyle   	, 0x00000000, % "ahk_id " CMhStr

			Gui, CM: Show  	, % "x" CVWin.X-CMWidth-4*CMWin.BW " y" CVWin.Y   , % "Fritzbox Nachrichten"

		}
		else
			Gui, CM: Show  	, % "x" CVWin.X-CMWidth-4*CVWin.BW " y" CVWin.Y   , % "Fritzbox Nachrichten"

	}

	If (fileIs != callmon.currentfilepath) {
		comstrings := FileOpen(callmon.currentfilepath, "r", "UTF-8").Read()
		GuiControl, CM:, CMFile, % callmon.currentfilepath
		GuiControl, CM:, CMStr, % RegExReplace(comstrings, "[\n\r]+", "`r`n")
		ControlSend,, % "{LControl Down}{End}{LControl Up}", % "ahk_id " CMhStr
	}


return
CMGuiClose:
CMGuiEscape:
	Gui, CM: Show, Hide
	callmon.ShowFritzMessages := false
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
									SendInput {Enter}
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

	; prüft, ob eines der erwarteten Fenster bereits geöffnet ist
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

;{ verschiedenes
MsgBoxMove(mx, my, wtitle, wtxt)                                                                	{

	hwnd := WinExist(wtitle " ahk_class #32770", wtxt)
	mbx := GetWindowSpot(hwnd)
	SetWindowPos(hwnd, mx-mbx.W/2, my+10, mbx.W, mbx.H)

}

clocktoseconds(Time)                                                                                 	{	;-- Zeitzeichenfolge (12:00:13) in Sekunden
return StrSplit(Time, ":").1*3600 + StrSplit(Time, ":").2*60 + StrSplit(Time, ":").3
}

nicenumber(telnr, formatter:="/")                                                               	{ 	; formatiert Telefonnummern schön
	static prefixes
	If !prefixes
		prefixes := "^(0)(15[1279]\d|16[023]|17[0-9]|15566|15888|118|180|700|800|900|" Addendum.prefixes ")"
return RegExReplace(telnr, prefixes "\" formatter "*", "$1$2" formatter)
}

plainnumber(telnr)                                                                                      	{	; entfernt alles das weder Zahl noch ein Pluszeichen ist
return RegExReplace(telnr, "[^\d\+]")
}

phonextrainfo(telnr)                                                                                     	{	; entfernt die Telefonnummer erhält Zusatztext
return LTrim(RegExReplace(telnr, "^.*?(\s(\pL)|$)", "$1"))
}

duration(sec, Mode:=1)                                                                             	{	; Formatierung für Zeitstrings

	hour	:= Floor(sec/3600)
	sec := sec - (hour+0)*3600
	min 	:= SubStr("00" Floor(sec/60), -1)
	sec 	:= SubStr("00" sec - (min+0)*60, -1)

return Mode=1 ? (min ":" sec)
		: Mode=2	? (hour+0>0 	? hour "h " min " min."
						:   min+0>0  	? min " min."
						:   sec "s.")
		: 0
}

SetExplorerTheme(HCTL)                                                                              	{ 	; HCTL: Handle eines ListView- oder TreeView-Steuerelements
   If (DllCall("GetVersion", "UChar") > 5) {
      VarSetCapacity(ClassName, 1024, 0)
      If DllCall("GetClassName", "Ptr", HCTL, "Str", ClassName, "Int", 512, "Int")
         If (ClassName = "SysListView32") || (ClassName = "SysTreeView32")
            return !DllCall("UxTheme.dll\SetWindowTheme", "Ptr", HCTL, "WStr", "Explorer", "Ptr", 0)
   }
   return false
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

GetControls(hwnd, class_filter:="", type_filter:="", info_filter:="") {                   	;-- gibt ein Array mit ClassNN, ButtonTyp, Position..... zurück

	; class_filter 	- kommagetrennte Liste von Klassen, die nicht gespeichert werden sollen
	; type_filter 	- kommagetrennte Liste von Kontrolltypen, die nicht gespeichert werden sollen
	; info_filter 	- kommagetrennte Liste der Klassen, die gespeichert werden sollen.

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

GetButtonType(hwndButton) {                                                                        	;-- ermittelt welcher Art ein Button ist, liest dazu den Buttonstyle aus
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

AddTooltip(p1,p2:="",p3="") {

	/*  Funktion: AddTooltip v2.0
	;
	; Beschreibung:
	;
	; Hinzufügen/Aktualisieren von Tooltips zu GUI-Steuerelementen.
	;
	; Parameter:
	;
	; p1 - Handle auf ein GUI-Steuerelement.  Alternativ auf "Activate" setzen, um die
	; die Tooltip-Steuerung zu aktivieren, "AutoPopDelay", um die Autopop-Verzögerungszeit einzustellen,
	; "Deaktivieren", um die Tooltip-Steuerung zu deaktivieren, oder "Titel", um den
	; Tooltip-Titel festzulegen.
	;
	; p2 - Wenn p1 das Handle zu einem GUI-Control enthält, sollte dieser Parameter
	; den Tooltip-Text enthalten.  Beispiel: "Mein Tooltip".  Auf null gesetzt, wird der
	; Tooltip zu löschen, der an das Steuerelement angehängt ist.  Wenn p1="AutoPopDelay", setzen Sie auf die
	; gewünschte Autopop-Verzögerungszeit in Sekunden.  Beispiel: 10. Hinweis: Die maximale
	; Autopop-Verzögerungszeit beträgt ~32 Sekunden.  Wenn p1="Title", setzen Sie den Titel der
	; der QuickInfo ein.  Beispiel: "Bob's Tooltips".  Auf null setzen, um den Tooltip zu entfernen
	; Titel.  Siehe den Abschnitt *Titel & Icon* für weitere Informationen.
	;
	; p3 - Tooltip-Symbol.  Siehe den Abschnitt *Titel & Icon* für weitere Informationen.
	;
	; Rückgabe:
	;
	; Das Handle auf das Tooltip-Steuerelement.
	;
	; Voraussetzungen:
	;
	; AutoHotkey v1.1+ (alle Versionen).
	;
	; Titel & Symbol:
	;
	Um den Titel der QuickInfo festzulegen, setzen Sie den Parameter p1 auf "Titel" und den Parameter p2
	; Parameter auf den gewünschten Tooltip-Titel.  Beispiel: AddTooltip("Titel", "Bob's
	; QuickInfos"). Um den Titel der QuickInfo zu entfernen, setzen Sie den Parameter p2 auf null.  Beispiel:
	; AddTooltip("Titel","").
	;
	; Der Parameter p3 bestimmt das Symbol, das zusammen mit dem Titel angezeigt wird,
	; falls vorhanden.  Wenn er nicht angegeben oder auf 0 gesetzt wird, wird kein Symbol angezeigt.  Um ein
	; Standardsymbol anzuzeigen, geben Sie einen der Standard-Icon-Identifikatoren an.  Siehe die
	; statischen Variablen der Funktion finden Sie eine Liste der möglichen Werte.  Beispiel:
	; AddTooltip("Titel", "Mein Titel",4).  Um ein benutzerdefiniertes Icon anzuzeigen, geben Sie ein Handle
	; auf ein Bild (Bitmap, Cursor oder Icon).  Wenn ein benutzerdefiniertes Icon angegeben wird, wird eine
	; Kopie des Symbols durch das Tooltip-Fenster erstellt, so dass bei Bedarf das Original
	; Icon jederzeit nach dem Setzen des Titels und des Icons zerstört werden kann.
	;
	; Das Setzen eines Tooltip-Titels führt in vielen Fällen nicht zum gewünschten Ergebnis.
	Der Titel (und das Symbol, falls angegeben) wird auf jedem Tooltip angezeigt, der
	; dieser Funktion hinzugefügt wird.
	;
	; Bemerkungen:
	;
	; Die Tooltip-Steuerung ist standardmäßig aktiviert.  Es besteht keine Notwendigkeit, das ; ; Tooltip-Steuerelement zu "aktivieren"
	; aktivieren, es sei denn, es wurde zuvor "deaktiviert".
	;
	Diese Funktion gibt den Handle auf das Tooltip-Control zurück, so dass bei Bedarf
	; zusätzliche Aktionen mit dem Tooltip-Steuerelement außerhalb dieser Funktion durchgeführt werden können
	; Funktion ausgeführt werden können.  Einmal erstellt, verwendet diese Funktion dasselbe Tooltip-Control wieder.
	Wird das Tooltip-Steuerelement außerhalb dieser Funktion ; zerstört, schlagen nachfolgende
	; Aufrufe dieser Funktion fehlschlagen.
	;
	; Kredit und Geschichte:
	;
	; Ursprünglicher Autor: Superfraggle
	; * Post: <http://www.autohotkey.com/board/topic/27670-add-tooltips-to-controls/>
	;
	; Aktualisiert zur Unterstützung von Unicode: art
	; * Post: <http://www.autohotkey.com/board/topic/27670-add-tooltips-to-controls/page-2#entry431059>
	;
	; Zusätzlich: jballi.
	; Fehlerkorrekturen.  Unterstützung für x64 hinzugefügt.  Modify-Parameter wurde entfernt.  hinzugefügt.
	; zusätzliche Funktionalität, Konstanten und Dokumentation.
	;
	 */

    Static hTT

          ;-- Misc. constants
          ,CW_USEDEFAULT:=0x80000000
          ,HWND_DESKTOP :=0

          ;-- Tooltip delay time constants
          ,TTDT_AUTOPOP:=2
                ;-- Set the amount of time a tooltip window remains visible if
                ;   the pointer is stationary within a tool's bounding
                ;   rectangle.

          ;-- Tooltip styles
          ,TTS_ALWAYSTIP:=0x1
                ;-- Indicates that the tooltip control appears when the cursor
                ;   is on a tool, even if the tooltip control's owner window is
                ;   inactive.  Without this style, the tooltip appears only when
                ;   the tool's owner window is active.

          ,TTS_NOPREFIX:=0x2
                ;-- Prevents the system from stripping ampersand characters from
                ;   a string or terminating a string at a tab character.
                ;   Without this style, the system automatically strips
                ;   ampersand characters and terminates a string at the first
                ;   tab character.  This allows an application to use the same
                ;   string as both a menu item and as text in a tooltip control.

          ;-- TOOLINFO uFlags
          ,TTF_IDISHWND:=0x1
                ;-- Indicates that the uId member is the window handle to the
                ;   tool.  If this flag is not set, uId is the identifier of the
                ;   tool.

          ,TTF_SUBCLASS:=0x10
                ;-- Indicates that the tooltip control should subclass the
                ;   window for the tool in order to intercept messages, such
                ;   as WM_MOUSEMOVE.  If this flag is not used, use the
                ;   TTM_RELAYEVENT message to forward messages to the tooltip
                ;   control.  For a list of messages that a tooltip control
                ;   processes, see TTM_RELAYEVENT.

          ;-- Tooltip icons
          ,TTI_NONE                            	:=0
          ,TTI_INFO                             	:=1
          ,TTI_WARNING          	         	:=2
          ,TTI_ERROR                           	:=3
          ,TTI_INFO_LARGE                 	:=4
          ,TTI_WARNING_LARGE        	:=5
          ,TTI_ERROR_LARGE              	:=6

          ;-- Extended styles
          ,WS_EX_TOPMOST              	:=0x8

          ;-- Messages
          ,TTM_ACTIVATE                   	:=0x401                    ;-- WM_USER + 1
          ,TTM_ADDTOOLA               	:=0x404                    ;-- WM_USER + 4
          ,TTM_ADDTOOLW                 	:=0x432                    ;-- WM_USER + 50
          ,TTM_DELTOOLA           	        :=0x405                    ;-- WM_USER + 5
          ,TTM_DELTOOLW          	        :=0x433                    ;-- WM_USER + 51
          ,TTM_GETTOOLINFOA          	:=0x408                    ;-- WM_USER + 8
          ,TTM_GETTOOLINFOW  		:=0x435                    ;-- WM_USER + 53
          ,TTM_SETDELAYTIME           	:=0x403                    ;-- WM_USER + 3
          ,TTM_SETMAXTIPWIDTH      	:=0x418                    ;-- WM_USER + 24
          ,TTM_SETTITLEA                   	:=0x420                    ;-- WM_USER + 32
          ,TTM_SETTITLEW                  	:=0x421                    ;-- WM_USER + 33
          ,TTM_UPDATETIPTEXTA  	     	:=0x40C                    ;-- WM_USER + 12
          ,TTM_UPDATETIPTEXTW        	:=0x439                    ;-- WM_USER + 57

    ;-- Save/Set DetectHiddenWindows
    l_DetectHiddenWindows:=A_DetectHiddenWindows
    DetectHiddenWindows On

    ;-- Tooltip control exists?
    if !hTT  {
        ;-- Create Tooltip window
        hTT:=DllCall("CreateWindowEx"
            ,"UInt",WS_EX_TOPMOST                            	;-- dwExStyle
            ,"Str","TOOLTIPS_CLASS32"                         	;-- lpClassName
            ,"Ptr",0                                                      	;-- lpWindowName
            ,"UInt",TTS_ALWAYSTIP|TTS_NOPREFIX          ;-- dwStyle
            ,"UInt",CW_USEDEFAULT                             	;-- x
            ,"UInt",CW_USEDEFAULT                             	;-- y
            ,"UInt",CW_USEDEFAULT                             	;-- nWidth
            ,"UInt",CW_USEDEFAULT                             	;-- nHeight
            ,"Ptr",HWND_DESKTOP                               	;-- hWndParent
            ,"Ptr",0                                                        	;-- hMenu
            ,"Ptr",0                                                        	;-- hInstance
            ,"Ptr",0                                                        	;-- lpParam
            ,"Ptr")                                                         	;-- Return type

        ;-- Disable visual style
        ;   Note: Uncomment the following to disable the visual style, i.e.
        ;   remove the window theme, from the tooltip control.  Since this
        ;   function only uses one tooltip control, all tooltips created by this
        ;   function will be affected.
		;;;;;        DllCall("uxtheme\SetWindowTheme","Ptr",hTT,"Ptr",0,"UIntP",0)

        ;-- Set the maximum width for the tooltip window
        ;   Note: This message makes multi-line tooltips possible
        SendMessage TTM_SETMAXTIPWIDTH,0,A_ScreenWidth,,ahk_id %hTT%
	}

    ;-- Other commands
    if p1 is not Integer
    {
        if (p1="Activate")
            SendMessage TTM_ACTIVATE,True,0,,ahk_id %hTT%

        if (p1="Deactivate")
            SendMessage TTM_ACTIVATE,False,0,,ahk_id %hTT%

        if (InStr(p1,"AutoPop")=1)  ;-- Starts with "AutoPop"
            SendMessage TTM_SETDELAYTIME,TTDT_AUTOPOP,p2*1000,,ahk_id %hTT%

        if (p1="Title")            {

            ;-- If needed, truncate the title
            if (StrLen(p2)>99)
                p2:=SubStr(p2,1,99)

            ;-- Icon
            if p3 is not Integer
                p3:=TTI_NONE

            ;-- Set title
            SendMessage A_IsUnicode ? TTM_SETTITLEW:TTM_SETTITLEA,p3,&p2,,ahk_id %hTT%
            }

        ;-- Restore DetectHiddenWindows
        DetectHiddenWindows % l_DetectHiddenWindows

        ;-- Return the handle to the tooltip control
        Return hTT
        }

    ;-- Create/Populate the TOOLINFO structure
    uFlags:=TTF_IDISHWND|TTF_SUBCLASS
    cbSize:=VarSetCapacity(TOOLINFO,(A_PtrSize=8) ? 64:44,0)
    NumPut(cbSize,      TOOLINFO,0,"UInt")              ;-- cbSize
    NumPut(uFlags,      TOOLINFO,4,"UInt")              ;-- uFlags
    NumPut(HWND_DESKTOP,TOOLINFO,8,"Ptr")               ;-- hwnd
    NumPut(p1,          TOOLINFO,(A_PtrSize=8) ? 16:12,"Ptr")
        ;-- uId

    ;-- Check to see if tool has already been registered for the control
    SendMessage, A_IsUnicode ? TTM_GETTOOLINFOW:TTM_GETTOOLINFOA, 0, &TOOLINFO,, % "ahk_id " hTT

    l_RegisteredTool:=ErrorLevel

    ;-- Update the TOOLTIP structure
    NumPut(&p2,TOOLINFO,(A_PtrSize=8) ? 48:36,"Ptr")
        ;-- lpszText

    ;-- Add, Update, or Delete tool
    if l_RegisteredTool         {

        if StrLen(p2)
            SendMessage,A_IsUnicode ? TTM_UPDATETIPTEXTW:TTM_UPDATETIPTEXTA,0,&TOOLINFO,, % "ahk_id " hTT
         else
            SendMessage,A_IsUnicode ? TTM_DELTOOLW:TTM_DELTOOLA,0,&TOOLINFO,, % "ahk_id " hTT

	} else
        if StrLen(p2)
            SendMessage,A_IsUnicode ? TTM_ADDTOOLW:TTM_ADDTOOLA,0,&TOOLINFO,, % "ahk_id " hTT

    ;-- Restore DetectHiddenWindows
    DetectHiddenWindows %l_DetectHiddenWindows%

    ;-- Return the handle to the tooltip control
    Return hTT
    }

Edit_Append(hEdit, txt)                                                                      		{         	;-- Modified version by SKAN
	Local        ; Original by TheGood on 09-Apr-2010 @ autohotkey.com/board/topic/52441-/?p=328342
	L :=	DllCall("SendMessage", "Ptr",hEdit, "UInt",0x0E, "Ptr",0 , "Ptr",0)             	; WM_GETTEXTLENGTH
	    	DllCall("SendMessage", "Ptr",hEdit, "UInt",0xB1, "Ptr",L , "Ptr",L)              	; EM_SETSEL
	    	DllCall("SendMessage", "Ptr",hEdit, "UInt",0xC2, "Ptr",0, "Str",txt "`r`n`r`n")   	; EM_REPLACESEL
	If RegExMatch(txt, "[\n\r]+$")
		ControlSend,, {Enter}, % "ahk_id " hEdit
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

	SWP_NOSIZE                     	:= 0x0001 ; Behält die aktuelle Größe bei (ignoriert die Parameter cx und cy).
	SWP_NOMOVE                 	:= 0x0002 ; Behält die aktuelle Position bei (ignoriert die Parameter X und Y).
	SWP_NOZORDER              	:= 0x0004 ; Behält die aktuelle Z-Reihenfolge bei (ignoriert den Parameter hWndInsertAfter).
	SWP_NOREDRAW             	:= 0x0008 ; Zeichnet Änderungen nicht neu.
	SWP_NOACTIVATE             	:= 0x0010 ; Aktiviert das Fenster nicht.
	SWP_DRAWFRAME             	:= 0x0020 ; Zeichnet einen Rahmen (definiert in der Klassenbeschreibung des Fensters) um das Fenster.
	SWP_FRAMECHANGED     	:= 0x0020 ; Wendet neue Rahmenstile an, die mit der Funktion SetWindowLong festgelegt wurden.
	SWP_SHOWWINDOW       	:= 0x0040 ; Zeigt das Fenster an.
	SWP_HIDEWINDOW         	:= 0x0080 ; Blendet das Fenster aus
	SWP_NOCOPYBITS           	:= 0x0100 ; Verwirft den gesamten Inhalt des Client-Bereichs.
	SWP_NOOWNERZORDER 	:= 0x0200 ; Verändert nicht die Position des Eigentümerfensters in der Z-Reihenfolge.
	SWP_NOREPOSITION        	:= 0x0200 ; Dasselbe wie das SWP_NOOWNERZORDER-Flag.
	SWP_NOSENDCHANGING 	:= 0x0400 ; Verhindert, dass das Fenster die Nachricht WM_WINDOWPOSCHANGING erhält.
	SWP_DEFERERASE             	:= 0x2000 ; Verhindert die Erzeugung der WM_SYNCPAINT-Nachricht.
	SWP_ASYNCWINDOWPOS 	:= 0x4000 ; Dies verhindert, dass der aufrufende Thread seine Ausführung blockiert, während andere Threads die Anfrage bearbeiten.

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

/*                                                                             	LESEN SIE DIES FÜR WEITERE INFORMATIONEN

                                			    	eine Funktion aus dem AHK-Forum: https://autohotkey.com/board/topic/46786-enumchildwindows/
                                                                                      sie wurde von IXIKO am 09. Mai 2018 geändert


	-findet ChildWindow-Handles von einem Parent-Fenster mit Hilfe von Name und/oder Klasse oder nur der WinID des ParentWindow
	-gibt eine kommaseparierte Liste von hwnds zurück oder nichts, wenn es keine Übereinstimmung gibt

	-Parameter Parent ist ein Objekt(). Übergeben Sie die folgenden {Key:Value} Paare wie folgt - WinTitle: "Name des Fensters", WinClass: "Klasse (NN) Name", WinID: ParentWinID
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
   VarSetCapacity(RECT, 16, 0), NumPut(LVIR, RECT, 0, "Int"), NumPut(Column-1, RECT, 4, "Int")

   If !DllCall("SendMessage", "Ptr", hLV, "UInt", 0x100E, "Ptr", Row-1, "Ptr", &RECT, "Int")
      return false

   If (Column=1 && (LVIR=0 || LVIR = 3))
      NumPut(NumGet(RECT, 0, "Int") + LV_EX_GetColumnWidth(hLV, 1), RECT, 8, "Int")

	cell    	:= {X:NumGet(RECT,  0, "Int"), Y:NumGet(RECT,  4, "Int"), R:NumGet(RECT,  8, "Int"),B:NumGet(RECT, 12, "Int")}
	cell.W	:= cell.R-cell.X
	cell.H	:= cell.B-cell.Y

	;~ SciTEOutput("H:" cJSON.Dump(cell, 1))

Return cell
   ;~ SendMessage, 0x100E, % (Row - 1), % &RECT,, % "ahk_id " hLV
   ;~ If !ErrorLevel
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

LV_EX_ItemText(HLV, Row, Column := 1, MaxChars := 257) {
   ; LVM_GETITEMTEXT -> http://msdn.microsoft.com/en-us/library/bb761055(v=vs.85).aspx
   Static LVM_GETITEMTEXT := A_IsUnicode ? 0x1073 : 0x102D ; LVM_GETITEMTEXTW : LVM_GETITEMTEXTA
   Static OffText := 16 + A_PtrSize
   Static OffTextMax := OffText + A_PtrSize
   VarSetCapacity(ItemText, MaxChars << !!A_IsUnicode, 0)
   ;~ LV_EX_LVITEM(LVITEM, , Row, Column)
   NumPut(&ItemText, LVITEM, OffText, "Ptr")
   NumPut(MaxChars, LVITEM, OffTextMax, "Int")
   SendMessage, % LVM_GETITEMTEXT, % (Row - 1), % &LVITEM, , % "ahk_id " . HLV
   VarSetCapacity(ItemText, -1)
Return ItemText
}

;}

;{ Monitorfunktionen
ScreenDims(MonNr:=1) {	                                                                                 		;-- returns a key:value pair of screen dimensions

	Sysget, MonitorInfo, Monitor, % MonNr
	X	:= MonitorInfoLeft
	Y	:= MonitorInfoTop
	W	:= MonitorInfoRight   	- MonitorInfoLeft
	H 	:= MonitorInfoBottom 	- MonitorInfoTop

	DPI    	:= A_ScreenDPI
	Orient	:= (W>H)?"L":"P"
	yEdge	:= DllCall("GetSystemMetrics", "Int", SM_CYEDGE)
	yBorder	:= DllCall("GetSystemMetrics", "Int", SM_CYBORDER)

 return {X:X, Y:Y, W:W, H:H, L:MonitorInfoLeft, R:MonitorInfoRight ,DPI:DPI, OR:Orient, yEdge:yEdge, yBorder:yBorder}
}

GetMonitorIndexFromWindow(windowHandle) {                                                     	;-- returns Monitorindex at window location

	; Starts with 1.
	; https://autohotkey.com/board/topic/69464-how-to-determine-a-window-is-in-which-monitor/

	monitorIndex := 1
	VarSetCapacity(monitorInfo, 40)
	NumPut(40, monitorInfo)

	if (monitorHandle := DllCall("MonitorFromWindow", "uint", windowHandle, "uint", 0x2)) && DllCall("GetMonitorInfo", "uint", monitorHandle, "uint", &monitorInfo) {

		monitorLeft  		:= NumGet(monitorInfo,  4, "Int")
		monitorTop    	:= NumGet(monitorInfo,  8, "Int")
		monitorRight  	:= NumGet(monitorInfo, 12, "Int")
		monitorBottom 	:= NumGet(monitorInfo, 16, "Int")
		workLeft      		:= NumGet(monitorInfo, 20, "Int")
		workTop       	:= NumGet(monitorInfo, 24, "Int")
		workRight     		:= NumGet(monitorInfo, 28, "Int")
		workBottom    	:= NumGet(monitorInfo, 32, "Int")
		isPrimary     		:= NumGet(monitorInfo, 36, "Int") & 1

		SysGet, MonCount, MonitorCount
		Loop % MonCount 	{                                    		; Compare location to determine the monitor index.
			SysGet, tempMon, Monitor, %A_Index%
			if ((monitorLeft = tempMonLeft) && (monitorTop = tempMonTop) && (monitorRight = tempMonRight) && (monitorBottom = tempMonBottom))
				return monitorIndex
		}

	}

return monitorIndex
}

GetMonitorAt(Lx, Ly, Ldefault:=1) {                                                                           	;-- get  index of the monitor containing the specified x and y co-ordinates.

	; https://autohotkey.com/board/topic/19990-windowpad-window-moving-tool/page-2
	; letzte Änderung: 27.09.2020

    SysGet, Lm, MonitorCount
    Loop % Lm {   ; Check if the window is on this monitor.

        SysGet, Mon, Monitor, % A_Index
        if (Lx >= MonLeft && Lx <= MonRight && Ly >= MonTop && Ly <= MonBottom)
            return A_Index

    }

return LDefault
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
#include lib\class_cJSON.ahk
#include lib\class_IPHelper.ahk
#include lib\class_LV_Colors.ahk
#include lib\class_socket.ahk
#include lib\class_Loaderbar.ahk
#include lib\GDIP_all.ahk
#include lib\SciTEOutput.ahk
#include lib\Sift.ahk
#Include lib\TreeListView.ahk
#Include lib\VarTreeGui.ahk
#Include lib\VarTreeObjectNode.ahk
#Include lib\VarEditGui.ahk




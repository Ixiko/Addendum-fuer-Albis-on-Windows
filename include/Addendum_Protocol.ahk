; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                             	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                             	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                              	by Ixiko started in September 2017 - last change 29.09.2021 - this file runs under Lexiko's GNU Licence
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; PROTOKOLLIEREN, FEHLERMELDUNGEN                                                                                                                                                                             	(05)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) FehlerProtokoll                        	(02) TProtGetDay 							(03) TProtMatchID 							(04) Telegram_Send                        	(03) TrayTip
;1
FehlerProtokoll(Exception, MsgToTelegram:= 1) {                                                	;-- erstellt ein Protokoll


	; REMARK: this function uses a global object "Addendum", this object must contain data to Addendum paths and Telegram data (please see Addendum.ahk)

	; MsgToTelegram - use the number of your stored chatbot in Addendum.ini
	; new since 24.01.2020 - function can handle an Exception object and custom strings

	; some variables
		inform  	:= []
		MaxLen		:= 0
		fillstr     	:= ""
		q          	:= Chr(0x22) ; 0x22 = "
		Range   	:= 3                                         	; means +- lines to print to errorlog (if I change the code the exeception.line is not corresponding later)

	; generates the timecode
		FormatTime, time, % A_Now, dd.MM.yyyy HH:mm:ss

		If IsObject(Exception)		{

				startLine	:= (Exception.line - Floor(Range/2))
				faultcode	:= "`tbetroffener Code (Bereich +-" Range "):`n"

			; saves a range of code lines to file protocoll
				faultFile 	:= FileOpen(Exception.File, "r", "UTF-8").Read()
				faultLine	:= StrSplit(faultFile, "`n", "`r")
				faultcode 	:= "`t`t/*`n"
				Loop % Range
					faultcode .= "`t`t`t" SubStr("00000" (startLine+A_Index-1), -4) ": " faultLine[(startLine+A_Index-1)] "`n"
				faultcode	.= " `t`t*/"

				eMsg:= RegexReplace(exception.Message, "(\{)", q "$1" q)
				eMsg:= RegexReplace(eMsg, "(\})", q "$1" q )

			; generates output text
				towrite1 :=   "Fehler im Skript`t: "       	A_ScriptName
				towrite1 .=	" - timecode: "               	time
				towrite1 .=	", Client: "                      	A_ComputerName
				towrite1 .=	", "                                	eMsg                   	" "
				towrite2 :=	"{`n`tSkriptpfad      `t: "	Exception.File    	"`n"
				towrite2 .=	"`tFehler bei Zeile`t: "    	Exception.Line    	"`n"
				towrite2 .=	"`t"                               	faultcode           	"`n"
				towrite2 .=	"`tFehlermeldung`t: "    	eMsg                  	"`n"
				towrite2 .=	"`tAuslöser         `t: "     	Exception.What 	"`n"
				towrite2 .=	"`tzusätzl. Info   `t: "      	Exception.Extra  	"`n"
				towrite2 .=	"`tA_LastError     `t: "     	A_LastError       	"`n"
				towrite2 .=	"`tLast ErrorLevel`t: "     	ErrorLevel          	"`n"
				towrite2 .=	"}`n"
		}
		else if RegExMatch(Exception, "Skript Fehlerausgabe:\s*([\w+_.$]+)", Skriptname) 	{
				RegExMatch(Exception, "Error:\s*([\w+_.$!]+)", eMsg)
				RegExMatch(Exception, "--->\s*(?<LineNr>\d+)\s*\:\s*(?<LineText>.*?)\n", xcptn_)
				RegExMatch(Exception, "[A-Z]:[\w\\_\säöüÄÖÜ.]+\.ahk", SkriptPfad)

			; generates output text
				towrite1 :=   "Fehler im Skript`t: "       	SkriptName1
				towrite1 .=	" - timecode: "               	time
				towrite1 .=	", Client: "                      	A_ComputerName
				towrite1 .=	", "                                	eMsg1                  	" "
				towrite2 :=	"{`n`tSkriptpfad      `t: "	SkriptPfad         	"`n"
				towrite2 .=	"`tFehler bei Zeile`t: "    	xcptn_LineNr      	"`n"
				towrite2 .=	"`t"                               	xcptn_LineText      	"`n"
				towrite2 .=	"`tFehlermeldung`t: "    	eMsg1                 	"`n"
				towrite2 .=	"`tA_LastError     `t: "     	A_LastError       	"`n"
				towrite2 .=	"`tLast ErrorLevel`t: "     	ErrorLevel          	"`n"
				towrite2 .=	"}`n"
		}

	; save generated protocol
		;~ SciTEOutput(towrite1 towrite2)
		FileAppend, % towrite1 towrite2, % Addendum.LogPath "\ErrorLogs\Fehlerprotokoll-" A_MM A_YYYY ".txt"


return true
}
;2
TProtGetDay(datestring, client)                                        	{                 	; Tagesprotokoll einlesen

	; datestring Format: dd.MM.yyyy oder yyyyMMdd

	static datestring_old, SectionDate, TPFullPath

	If (datestring_old != datestring) {

		If !RegExMatch(datestring, "(?<dd>\d+)\.(?<MM>\d+)\.(?<YYYY>\d+)", d)
			RegExMatch(datestring, "^(?<YYYY>\d{4})(?<MM>\d{2})(?<dd>\d{2})", d)

		lDate := dYYYY dMM ddd
		datestring_old := datestring

		FormatTime, weekdayshort, % lDate, ddd
		FormatTime, weekdaylong	, % lDate, dddd
		FormatTime, month       	, % lDate, MMMM

		SectionDate 	:= weekdaylong "|" ddd "." dMM
		TPFullPath		:= Addendum.DBPath "\Tagesprotokolle\" dYYYY "\" dMM "-" month "_TP.txt"

	}

	If !FileExist(TPFullPath)
		return "no file: " TPFullPath

	IniRead, TProtoTmp, % TPFullPath, % SectionDate, % client
 	TProto := 	InStr(TProtoTmp, "Error") 	? "Error"
				: 	InStr(TProtoTmp, ",")       	? StrSplit(RTrim(TProtoTmp, ","), ",")
				                                        	: StrSplit(TProtoTmp, ">", "<")
	If !Tproto[TProto.Count()]
		TProto.RemoveAt(TProto.Count())

return TPRoto
}
;3
TProtMatchID(matchID, datestring:="", client:="")             	{               	; sucht nach bereits gespeicherten PatID's

  ; datestring und client leer lassen um mit dem aktuellen Tagesprotokoll zu arbeiten

  ; bei Angabe eines Datums wird das jeweilige Tagesprotokoll geladen
	Tprot  	:= datestring ? TProtGetDay(datestring, (client ? client : compname)) : Addendum.TProtocol

  ; Tagesprotokoll enthält keine Daten
	If (!IsObject(TProt) || TProt.Count() = 0)
		return false

  ; vergleicht gespeicherte PatientenIDs mit der Übergebenen
	For each, PatIDTimeStr in TProt {
		RegExMatch(PatIDTimeStr, "O)(?<ID>\d+)\((?<Hour>\d+):(?<Min>\d+)\)", Pat)
		If (Pat.ID = matchID)
			return SubStr("00" Pat.Hour, -1) ":" SubStr("00" Pat.Min, -1)
	}

return false
}
;4
Telegram_Send(telegramBotKey, telegramChatId, textMessage) {				     		;-- another way to send a message

    WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
    WinHTTP.Open("POST", Format("https://api.telegram.org/bot{1}/sendMessage?chat_id={2}&text={3}", telegramBotKey, telegramChatId, textMessage), 0)
	WinHTTP.Send()
    TrayTip Text Sent to Telegram, % textMessage

Return
}
;5
TrayTip(Title, Msg, Seconds:=2) {
	If !InStr(compname, "SP1")
		return
	TrayTip, % Title, % Msg, % Seconds
}


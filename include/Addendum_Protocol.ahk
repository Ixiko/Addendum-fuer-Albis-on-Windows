﻿; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                              	Automatisierungs- oder Informations Funktionen für das AIS-Addon: "Addendum für Albis on Windows"
;                                                                                            	!diese Bibliothek wird von fast allen Skripten benötigt!
;                                                            	by Ixiko started in September 2017 - last change 27.04.2020 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; PROTOKOLLIEREN, FEHLERMELDUNGEN                                                                                                                                                                                                      	(04)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; (01) FehlerProtokoll                        	(02) ExceptionHelper                      	(03) Telegram_Send                          	(04) TrayTip
;1
FehlerProtokoll(Exception, MsgToTelegram:= 1) {                                                                   	;-- Addendum.ahk stürzt auf manchen Clients einfach ab, Funktion erstellt ein Protokoll hoffentlich

	; REMARK: this function uses a global object "Addendum", this object must contain data to Addendum paths and Telegram data (please see Addendum.ahk)

	; MsgToTelegram - use the number of your stored chatbot in Addendum.ini
	; 24.01.2020 - function can handle an Exception object and custom strings
	; 01.04.2020 - it happens that an error was detected all 3-5 second over night. function stores the last exception and compares it with the new one, if its matches they function will immediatly return,
	;						remark: it will not send a message or write the exception to the protocol file!

	; dependancies: \lib\crypt.ahk

	; static
		static lastExceptionObject
		static lastExceptionText

	; some variables
		ExcHit   	:= 0
		inform  	:= []
		MaxLen	:= 0
		fillstr     	:= ""
		q          	:= Chr(0x22)                        	; 0x22 = "
		Range   	:= 3                                      	; means +- lines to print to errorlog (if I change the code the exeception.line is not corresponding later)

		IniRead, lastMsg, % Addendum.AddendumIni, % compname, % "letzte_Fehlermeldung"
		RegExMatch(lastMsg, "^\s*time\:\s*(?<time>\d+)\shash\:\s(?<hash>\w+)", msg)


	; protocol filename
		protocol	:= Addendum["AddendumDir"] "\logs'n'data\ErrorLogs\Fehlerprotokoll-" A_MM A_YYYY ".txt"

	; generates the timecode
		FormatTime, time, % A_Now, dd.MM.yyyy HH:mm:ss

	; loads and analyze last written exception to supress an excess of sending error messages

	; generates a readable protocol
		If IsObject(Exception)  	{

			startLine	:= (Exception.line - Range)
			startLine	+= 1
			faultcode	:= "`tbetroffener Code (Bereich +-" Range "):`n"

		; saves a range of code lines to file protocoll
			FileRead, faultFile, % A_ScriptFullPath
			faultLine	:= StrSplit(faultFile, "`n", "`r")

			Loop, % Range
				faultcode .= "`t`t`t" SubStr("00000" (startLine+A_Index), -4) ": " q faultLine[(startLine+A_Index)] q "`n"
			faultcode	:= RTrim(faultcode, "`n")

			eMsg	:= RegexReplace(exception.Message, "(\{)", q "$1" q)
			eMsg 	:= RegexReplace(eMsg, "(\})", q "$1" q )
			eMsg1	:= RegExReplace(eMsg, "Action\:", "Action  `t:")
			eMsg1	:= RegExReplace(eMsg1, "Params\:", "Params `t:")
			eMsg2	:= RegExReplace(eMsg, "Action\:", "                           `tAction`t:")
			eMsg2	:= RegExReplace(eMsg2, "Params\:", "                           `t$1")
			RegExMatch(exception.Message, "^(.*?)\:", reason)

		; generates output text (the first is for telegram , the second is stored in protocol file )
			towrite1 :=	"Skript   `t: " A_ScriptName ", timecode: " time	"`n"
			towrite1 .=	"Client   `t: "                    	A_ComputerName	"`n"
			towrite1 .=	"Reason`t: "                    	eMsg                   		"`n"

			towrite2 :=	"MsgNr`t: "                    	lastMessageNr   	 "`n"
			towrite2 .=	"Skript   `t: " 						A_ScriptName		 ", "
			towrite2 .=	"timecode: " 		       			time                   	 ", "
			towrite2 .= 	"Client: " 							A_ComputerName "`n"
			towrite2 .=	"Reason `t:"                   	reason1             	 "`n"
			towrite2 :=	"{`n`tSkriptpfad        `t: "	Exception.File      	 "`n"
			towrite2 .=	"`tFehler bei Zeile`t: "     	Exception.Line      	 "`n"
			towrite2 .=	"`t"                                  	faultcode             	 "`n"
			towrite2 .=	"`tFehlermeldung`t: "     	eMsg2                 	 "`n"
			towrite2 .=	"`tAuslöser         `t: "       	Exception.What   	 "`n"
			towrite2 .=	"`tzusätzl. Info   `t: "        	Exception.Extra    	 "`n"
			towrite2 .=	"`tA_LastError     `t: "       	A_LastError          	 "`n"
			towrite2 .=	"`tLast ErrorLevel `t: "      	ErrorLevel            	 "`n"
			towrite2 .=	"}`n"

		; hash parts of informations to prevent sending every second an error msg
			hashthis := A_ScriptName "," A_ComputerName "," Exception.File "," Exception.Line "," faultcode
			newHash	:= CryptHash(&hashthis, StrLen(hashthis), "MD5")

		; save this exception object
			lastException := exception
	}
		else if RegExMatch(Exception, "Skript Fehlerausgabe:\s*([\w+_.$]+)", Skriptname) {

			; same with Exception coming as string, comparing and return from function for output
				If (StrLen(lastExceptionText) > 0) && (lastExceptionText = Exception)
					return

			; save ExceptionText
				lastExceptionText := Exception

			; RegEx the content from Exception string
				RegExMatch(Exception, "Error:\s*([\w+_.$!]+)", eMsg)
				RegExMatch(Exception, "--->\s*(?<LineNr>\d+)\s*\:\s*(?<LineText>.*?)\n", xcptn_)
				RegExMatch(Exception, "[A-Z]:[\w\\_\säöüÄÖÜ.]+\.ahk", SkriptPfad)

			; generates output text
				towrite1 :=   "Fehler im Skript`t: "       	SkriptName1
				towrite1 .=	" - timecode: "               	time
				towrite1 .=	", Client: "                      	A_ComputerName
				towrite1 .=	", "                                	eMsg1                 	"`n"
				towrite2 :=	"{`n`tSkriptpfad      `t: "	SkriptPfad            	"`n"
				towrite2 .=	"`tFehler bei Zeile`t: "    	xcptn_LineNr       	"`n"
				towrite2 .=	"`t"                               	xcptn_LineText      	"`n"
				towrite2 .=	"`tFehlermeldung`t: "    	eMsg1                 	"`n"
				towrite2 .=	"`tA_LastError     `t: "     	A_LastError          	"`n"
				towrite2 .=	"`tLast ErrorLevel`t: "     	ErrorLevel            	"`n"
				towrite2 .=	"}`n"

			; hash parts of informations to prevent sending every second an error msg
				hashthis := SkriptName1 "," A_ComputerName "," SkriptPfad "," xcptn_LineNr "," eMsg1
				newHash	:= CryptHash(&hashthis, StrLen(hashthis), "MD5")

		}

	; compare this message with last one
		If ((A_TickCount - msgTime) < 10000) || (msgHash = newHash)
			return

	; save generated protocol
		FileAppend, % towrite1 towrite2, % protocol
		IniWrite, A_TickCount " " newHash, % Addendum.AddendumIni, % compname, % "letzte_Fehlermeldung"

	; sends a message to preferred telegram bot option is set
		If (MsgToTelegram > 0) && IsObject(Addendum.Telegram) &&
		{
				If Addendum.Telegram.Active
		    	{
						BotNr	:= MsgToTelegram
						Telegram_Send(Addendum.Telegram[BotNr].Token, Addendum.Telegram[BotNr].ChatID, towrite1 ", Auslöser: " Exception.What ", Zeile(" Exception.Line "): " faultLine[Exception.Line])
						IniWrite, % A_TickCount, % AddendumDir "\Addendum.ini", % Addendum.Telegram[BotNr].BotName "_LastMsgTime"
						FileAppend, % "message send to telegram`n`n", % protocol
				}
		}


return true
}
;2
ExceptionHelper(libPath, SearchCode, ErrorMessage, codeline) {													;-- searches for the given SearchCode in an uncompiled script as a help to throw exceptions

	If !A_IsCompiled {

		;Fehlerfunktion bei Eingabe eines falschen Parameter
			FileRead, Pfunc, %AddendumDir%\libPath
			Loop, Parse, Pfunc, `n
			{
					If Instr(A_LoopField, SearchCode) {
						scriptline:= A_Index
						ScriptText:= A_LoopField
						break
					}
			}

			Exception(ErrorMessage)

	} else {

			msg=
					(Ltrim
					This message is shown, because the script wanted
					to call a function that works only in uncompiled scripts.

					A function was called to show a runtime error.
					This function was called from %A_ScriptName%
					at line: %codeline%. The code to show ist:
					%SearchCode%
					with the following error-message:
					%ErrorMessage%
					)

			MsgBox, % "Addendum für AlbisOnWindows - " A_ScriptName,  %msg%

	}

}
;3
Telegram_Send(telegramBotKey, telegramChatId, textMessage) {				                    			;-- another way to send a message

    WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
    WinHTTP.Open("POST", Format("https://api.telegram.org/bot{1}/sendMessage?chat_id={2}&text={3}", telegramBotKey, telegramChatId, textMessage), 0)
	;WinHTTP.SetRequestHeader("Content-Type", "application/json")
	WinHTTP.Send()

    TrayTip Text Sent to Telegram, %textMessage%

Return
}
;4
TrayTip(Title, Msg, Seconds:=2) {

	If !InStr(compname, "SP1")
		return

	TrayTip, % Title, % Msg, % Seconds
}

;intern
CryptHash(pData, nSize, SID = "CRC32", nInitial = 0) {
	CALG_SHA := CALG_SHA1 := 1 + CALG_MD5 := 0x8003
	If Not	CALG_%SID%
	{
		FormatI := A_FormatInteger
		SetFormat, Integer, H
		sHash := DllCall("ntdll\RtlComputeCrc32", "Uint", nInitial, "Uint", pData, "Uint", nSize, "Uint")
		SetFormat, Integer, %FormatI%
		StringUpper,	sHash, sHash
		StringReplace,	sHash, sHash, X, 000000
		Return	SubStr(sHash,-7)
	}

	DllCall("advapi32\CryptAcquireContextA", "UintP", hProv, "Uint", 0, "Uint", 0, "Uint", 1, "Uint", 0xF0000000)
	DllCall("advapi32\CryptCreateHash", "Uint", hProv, "Uint", CALG_%SID%, "Uint", 0, "Uint", 0, "UintP", hHash)
	DllCall("advapi32\CryptHashData", "Uint", hHash, "Uint", pData, "Uint", nSize, "Uint", 0)
	DllCall("advapi32\CryptGetHashParam", "Uint", hHash, "Uint", 2, "Uint", 0, "UintP", nSize, "Uint", 0)
	VarSetCapacity(HashVal, nSize, 0)
	DllCall("advapi32\CryptGetHashParam", "Uint", hHash, "Uint", 2, "Uint", &HashVal, "UintP", nSize, "Uint", 0)
	DllCall("advapi32\CryptDestroyHash", "Uint", hHash)
	DllCall("advapi32\CryptReleaseContext", "Uint", hProv, "Uint", 0)

	FormatI := A_FormatInteger
	SetFormat, Integer, H
	Loop,	%nSize%
		sHash .= SubStr(*(&HashVal + A_Index - 1), -1)
	SetFormat, Integer, %FormatI%
	StringReplace,	sHash, sHash, x, 0, All
	StringUpper,	sHash, sHash
	Return	sHash
}
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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

				startLine	:= (Exception.line - Range)
				startLine	+= 1
				faultcode	:= "`tbetroffener Code (Bereich +-" Range "):`n"

			; saves a range of code lines to file protocoll
				FileRead, faultFile, % A_ScriptFullPath
				faultLine	:= StrSplit(faultFile, "`n", "`r")
				Loop, % Range
					faultcode .= "`t`t`t" SubStr("00000" (startLine+A_Index), -4) ": " faultLine[(startLine+A_Index)] "`n"
				faultcode	:= RTrim(faultcode, "`n")

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
		FileAppend, % towrite1 towrite2, % Addendum["AddendumDir"] "\logs'n'data\ErrorLogs\Fehlerprotokoll-" A_MM A_YYYY ".txt"


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
;5
HashThisString(str) {
return CryptHash(&str, StrLen(str), "MD5")
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
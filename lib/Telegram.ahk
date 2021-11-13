;------------------------------ Telegram Bot Functions Library! (hopefully easy to understand) -------------------------
;------------------------------------------- this library contains code from gregster ----------------------------------------
;----------------------- https://autohotkey.com/boards/viewtopic.php?f=5&t=42031&hilit=Telegram -------------
									    TBFLib_Version:= "0.3", TBFLib_RevisonDate:= "29.04.2018"
;--------------------------------- written for use with Addendum fÃ¼r ALBIS on WINDOWS ------------------------------
;---------------------------------- written by Ixiko and now gregster (first edition 2017) ---------------------------------
;--------------------------- please report errors and suggestions to me: Ixiko@mailbox.org ---------------------------
;------------------------ use subject: "Addendum" so that you don't end up in the spam folder ------------------------
;--------------------------------- GNU Lizenz - can be found in main directory  - 2017 ---------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------

/*
the Telegram Api supports GET und POST HTTP methods

this function library will return all informations it receives from Telegram as associative Arrays
	for example if you use GetMe you will receive an array with the following contents:
		Value:= array["update_id"] or Value:= array_you_choose.first_name
		I have chosen the same keys that can be read in the Api description at telegram.org
			So I hope the output values are easier to understand.


*/

;#include <json.ahk>


;{-- Telgram.api Commands

Telegram_GetChatID(BotToken) {

	global nooption

	straw := DownloadToString("https://api.telegram.org/bot" BotToken "/get", "utf-8")

	if (straw = nooption)
		return "no_message"

	obj := Telegram_parse2obj(straw)

return obj
}

Telegram_getUpdates(BotToken, offset = "") { 													;GetMessages incl. all id's , tags and so on

	global nooption
	obj	:= Object()
	url	: = "https://api.telegram.org/bot" BotToken "/getUpdates?offset=" offset
	straw:=DownloadToString(url, "utf-8")

	if (straw = nooption)
		return "no_message"

return straw
}

Telegram_GetUpdatesT(token, offset="", updlimit=100, timeout=0) { 				;includes limit and timeout
	If (updlimit>100)
		updlimit := 100
	; Offset = Identifier of the first update to be returned.
	url := "https://api.telegram.org/bot" token "/getupdates?offset=" offset "&limit=" updlimit "&timeout=" timeout
	updjson := URLDownloadToVar(url,"GET")
return updjson
}

Telegram_getMe(BotToken) { 																			;getting your id; first_name , Username(bot-name)

	url = https://api.telegram.org/bot%botToken%/getMe
	straw:=DownloadToString(url, "utf-8")

	idn:= ExtractFromStringV2(straw, "id" . """:" , "`," , 1, 0, 0)
	ib:= ExtractFromStringV2(straw, "is_bot" . """:" , "`," , 1, 0, 0)
	fn:= ExtractFromStringV2(straw, "first_name" . """:""" , "`," , 1, 0, 1)
	un:= ExtractFromStringV2(straw, "username" . """:""" , "`}" , 1, 0, 1)

	GetMe:= Object("id", idn, "is_bot", ib, "first_name", fn, "username", un)

return GetMe
}

Telegram_SendMessage(BotToken, chat_id, msg) { 											;Send a message to the chat_id you received through the Telegram_getUpdates () function

	;zum ersetzen von Umlauten damit deutsche Buchstaben gesendet werden kÃ¶nnen
			StringReplace, msg, msg, ä,       `%C3`%A4    , All
			StringReplace, msg, msg, ö,       `%C3`%B6    , All
			StringReplace, msg, msg, ü,       `%C3`%BC    , All
			StringReplace, msg, msg, Ä,       `%C3`%84    , All
			StringReplace, msg, msg, Ö,       `%C3`%96    , All
			StringReplace, msg, msg, Ü,       `%C3`%9C    , All
			StringReplace, msg, msg, Space,   `%20        , All
			StringReplace, msg, msg, ß,       `%C3`%9F    , All
			StringReplace, msg, msg, °,       `%C2`%B0    , All

	;url:= "https://api.telegram.org/bot" . BotToken . "/sendmessage?chat_id=" . Chat_ID . "%&text=" . msg . "`""
	url = https://api.telegram.org/bot%BotToken%/sendmessage?chat_id=%chat_id%&text=%msg%
	straw:=DownloadToString(url, "utf-8")

	return straw
}

Telegram_SendText(token, text, from_id, replyMarkup="",  parseMode="" ) { 		;its like SendMessage - this function is written by gregster
		url := "https://api.telegram.org/bot" token "/sendmessage?chat_id=" from_ID "&text=" text "&reply_markup=" replyMarkup "&parse_mode=" parseMode
		json_message := URLDownloadToVar(url,"GET")
		return json_message
}

Telegram_Somefunction(botToken, msg, from_id, mtext) { 									;  - this function is written by gregster
		if (msg.callback_query.id!="")						;  After the user presses a callback button, Telegram clients will display a progress bar until you call answerCallbackQuery
		{
			cbQuery_id := msg.callback_query.id
			; Notification and alert are optional
			url := "https://api.telegram.org/bot" botToken "/answerCallbackQuery?text=Notification&callback_query_id=" cbQuery_id "&show_alert=true"
			json_message := URLDownloadToVar(url,"GET")
				;msgbox % json_message
		}
		text := "You chose " mtext
		Telegram_SendText(botToken, text, from_id)        ;  url encoding for messages :  %0A = newline (\n)   	%2b = plus sign (+)		%0D for carriage return \r        single Quote ' : %27			%20 = space
		return
}

Telegram_Inlinebuttons(botToken, msg,  from_id, mtext="") { 								; add inline buttons  - this function is written by gregster
	keyb=																	; json string:
	(
		{"inline_keyboard":[ [{"text": "Command One" , "callback_data" : "Command1"}, {"text" : "Some button (Cmd3)",  "callback_data" : "Command3"} ] ], "resize_keyboard" : true }
	)
	url := "https://api.telegram.org/bot" botToken "/sendMessage?text=Keyboard added&chat_id=" from_id "&reply_markup=" keyb
	json_message := URLDownloadToVar(url,"GET")   ; find this function at the next code box below
	return json_message
}

Telegram_RemoveKeyb(botToken, msg, from_id, mtext="") { 								; Remove custom keyboard  - this function is written by gregster
	keyb=
	(
		{"remove_keyboard" : true }
	)
	url := "https://api.telegram.org/bot" botToken "/sendMessage?text=Keyboard removed&chat_id=" from_id "&reply_markup=" keyb
	json_message := URLDownloadToVar(url,"GET")
		;msgbox % json_message
	return json_message
}

Telegram_Send(telegramBotKey, telegramChatId, textMessage) {							;-- another way to send a message

    WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
    WinHTTP.Open("POST", Format("https://api.telegram.org/bot{1}/sendMessage?chat_id={2}&text={3}", telegramBotKey, telegramChatId, textMessage), 0)
	;WinHTTP.SetRequestHeader("Content-Type", "application/json")
	WinHTTP.Send()

    TrayTip Text Sent to Telegram, %textMessage%

Return
}

Telegram_getUpdate(telegramBotKey, offset = "") { 													;GetMessages incl. all id's , tags and so on

		global nooption
		obj:= Object()
		url = https://api.telegram.org/bot%BotToken%/getUpdates?offset=%offset%
		straw:=DownloadToString(url, "utf-8")

		if straw = nooption
			return "no_message"

		WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
		WinHTTP.Open("POST", Format("https://api.telegram.org/bot{1}/getUpdates?offset={2}", telegramBotKey, offset), 0)
		;WinHTTP.SetRequestHeader("Content-Type", "application/json")
		WinHTTP.Send()

		TrayTip Text Sent to Telegram, %textMessage%



return WinHTTP.Request()
}

;}

;{-- additional important Telegram_funtions - these do not belong to the official telegram.api

Telegram_parse2obj(str) { ;this is my first telegram_json wrapper, is slow but works
   ;in the first i used json2obj but the function json2obj did not find the last part of the telegram message
   ;I understand the RegEx method only inadequate therefore I parse with my awkward function
   ;I hope someone find a way to speed up my function. At the moment I do not have the time to learn enough about RegEx.

;{ -DEFINITION OF REPLACE STRINGS AND REPLACE WITH - FIRST REPLACE CHAR'S LIKE {}[]()" THEN REPLACE WORDS
   rps1 = `n
   rw1 =
   rps2  = `r
   rw2  =
   rps3  = `{
   rw3 =
   rps4 = `}
   rw4 =
   rps5 = `]
   rw5 =
   rps6 = `"
   rw6 =
   rps7 = message:
   rw7 =
   rps8 = from:
   rw8 =
   rps9 = chat:id
   rw9 = chat_id
   rps10 = chat: id
   rw10 = chat_id
   rps11 = result:`[
   rw11 =
   rps12 = `\u00fc
   rw12 = ü
;}

      arr:= Object()
      str:= RegExReplace(str, "((?:{)[\s\S][^{}]+(?:}))", "", "", 1, 1)
      i:= 0
		str1:= str
      Loop, 12
      {
			  i += 1
			  r = % rps%i%
			  w = % rw%i%
			  StringReplace, str, str, %r%, %w%, All
      }

	;For a better overview when debugging I exchanged the comma against linefeed
	StringReplace, str, str, `, , `n, All

      Loop, Parse, str, `n
      {
               lf:=Trim(A_LoopField)
               if (lf <> "") {
					   dpoint:=Instr(lf, "`:")
					   key:= Substr(lf, 1, dpoint - 1)
					   value:= SubStr(lf, dpoint + 1, StrLen(lf) - dpoint)
					   arr[key]:= value
				}
	  }

return arr
}

Telegram_collectMessage(msg) {

	global Telegram_Brain		;i call it Brain, but its no brain, its only a collection of Messages and possible answers you can teach your bot
	global MyCommandList


return MyCommand
}

;}

;{-- currently not implemented functions
Telegram_setWebhook(url, certificate, max_connections, allowed_updates) {
; url as string , certificate as InputFile (optional), max_connections as Integer (optional), allowed_updates as Array of String
; url must be https - enabled! currently used ports are  443, 80, 88, 8443


}

;}

;{-- Useful functions for Telegram commands

DownloadToString(url, encoding="utf-8") {

	;wonderful function by nnik Posted 24 January 2013 - 09:41 PM
	;https://autohotkey.com/board/topic/89509-bestimmten-inhalt-einer-webseite-auslesen-und-in-txt-schreiben/
  static a := "AutoHotkey/" A_AhkVersion
  if (!DllCall("LoadLibrary", "str", "wininet") || !(h := DllCall("wininet\InternetOpen", "str", a, "uint", 1, "ptr", 0, "ptr", 0, "uint", 0, "ptr")))
    return 0
  c := s := 0, o := ""
  if (f := DllCall("wininet\InternetOpenUrl", "ptr", h, "str", url, "ptr", 0, "uint", 0, "uint", 0x80000000, "ptr", 0, "ptr")) {

    while (DllCall("wininet\InternetQueryDataAvailable", "ptr", f, "uint*", s, "uint", 0, "ptr", 0) && s>0) {

				  VarSetCapacity(b, s, 0)
				  DllCall("wininet\InternetReadFile", "ptr", f, "ptr", &b, "uint", s, "uint*", r)
				  o .= StrGet(&b, r>>(encoding="utf-16"||encoding="cp1200"), encoding)
    }

		DllCall("wininet\InternetCloseHandle", "ptr", f)
  }

  DllCall("wininet\InternetCloseHandle", "ptr", h)
  return o

}

URLDownloadToVar(url,method) {
	hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
	hObject.Open(method,url)
	hObject.Send()
	return hObject.ResponseText
}

ExtractFromStringV2(string, ParseString1, ParseString2, offset, cutleft, cutright ) {
	;if extracted string contains signs before and after - use cutleft as a number of signs to cut from left side of string after the end of the beginning of the last parsestring ....

	Len1:=StrLen(ParseString1)
	PosFound:=InStr(string, ParseString1, , offset) + Len1 + cutleft
	Length:=InStr(string, ParseString2, , PosFound + offset) - PosFound - cutright

	If (Pos1 = 0) or (Pos2 = 0)
			return "no match"

return SubStr(string, PosFound, Length)
}




;}


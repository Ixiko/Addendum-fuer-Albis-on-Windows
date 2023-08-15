; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;               															⚫      Telegram		⚫
;
;      Funktionen:   			▫ Callback Funktion für die TelegramBot Klasse
;                       			▫ Wrapper für Telegram API Methoden (Übergabe von Voreinstellungen)
;
;      Basisskript: 	    	Addendum.ahk
;
;		Abhängigkeiten: 	lib\class_TelegramBot.ahk
;		                        	lib\SciteOutput.ahk
;
;	                    	Addendum für Albis on Windows
;                        	by Ixiko started in September 2017 - letzte Änderung 15.09.2022 - this file runs under Lexiko's GNU Licence
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



Telegram_SendDocument(DocumentPath, sendTo:="", Notification:="")    	{	        		;-- versendet Dokumentdateien

	global TBot

	BotSet := {	"BotName"        	: Addendum.Telegram.Bots[1].botname
	        		,	"Token"         		: Addendum.Telegram.Bots[1].Token
	        		,	"LastOffset"      		: Addendum.Telegram.Bots[1].LastOffset
	        		,	"LastMsgDate"  		: Addendum.Telegram.Bots[1].LastMsgDate
	        		,	"ReplyTo"        		: Addendum.Telegram.Bots[1].ReplyTo
	        		,	"updateInterval"		: false
	        		,	"dialogInterval"  	: 5
	        		,	"DialogTimeout" 	: 60                                                        ; time without new messages to automatic switch back to long poll interval
	        		, 	"callbackfunc"    	: "Telegram_MessageReceiver"
	        		,	"log"                    	: true
	        		, 	"debug"             	: false}



	chatID := Addendum.Telegram.Chats[sendTo]
	PraxTT("versende Dokument: " DocumentPath "`nan " sendTo " [ID: " chatID "]" , "5 1")

	TBot := new TelegramBot(BotSet)
	response 	:= TBot.SendDocument(chatID, DocumentPath, Notification)
	If TBot.log
		FileAppend, % A_YYYY A_MM A_DD " " A_Hour ":" A_Min ":" A_Sec " " (IsObject(response) ? cJSON.Dump(response) : response)
		, % Addendum.DBPath "\logs\Telegram.txt", UTF-8
	TBot := ""

}

Telegram_MessageReceiver() {

}
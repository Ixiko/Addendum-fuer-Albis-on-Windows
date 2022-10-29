;-----------------------------------------------------------------------------------------------------------------------------------
;------------------------------------------------------- TELEGRAM BOT -------------------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------
													Version:= "0.2" , VersionDate:= "14.09.2022"

;-------------------------------------------- Addendum für AlbisOnWindows -----------------------------------------------
;-----------------------------------------------------------------------------------------------------------------------------------

; 	starting a small answer bot

/* example

	global q := Chr(0x22)
	global TBot


	settings := {	"BotName"        	: botname
					,	"Token"         		: Token
					,	"chatID"         		: ChatId
					,	"LastOffset"      		: LastOffset
					,	"LastMsgDate"  		: LastMsgDate
					,	"ReplyTo"        		: ReplyTo
					,	"updateInterval"		: false
					,	"dialogInterval"  	: 5
					,	"DialogTimeout" 	: 60                                                        ; time without new messages to automatic switch back to long poll interval
					, 	"callbackfunc"    	: "Telegram_Messagehandler"
					,	"log"                    	: true
					, 	"debug"             	: true}


	TBot := new TelegramBot(settings)
	response 	:= TBot.SendDocument(TBot.ChatId, "C:\Windows\System32\licence.rtf")
	rjson := cJSON.Load(response)
	TBot.print(cJSON.Dump(rjson, 1))
	TBot := ""


Telegram_Messagehandler() {

		static Reply:=Object()

		messages:= TBot.GetNewMessages()
		TBot.print("received: " (MsgMax := messages.MaxIndex()) " new messages")

		; loop through message queue
		Loop % MsgMax {

				msg := messages[A_Index]
				TBot.print("  " A_Index ". from: " msg.username ", msg: " q msg.text q)

			; get user and check if known
				user := TBot.GetUser(msg.ChatId)
				If !IsObject(user) {
						TBot.print("unknown user with chat_id: " msg.ChatId)
						continue
				}

			; check if it's our admin
				If RegExMatch(user.auth, "i)admin") {

					; start reply service if not
						username:= user.username
						;TBot.print("user " q user " with chatID: " msg.ChatId " " q " is admin. Starting reply service now.")
						If TBot.newReply(username, msg.ChatId, user.auth) {
							TBot.print("reply service for user " q user "(" msg.ChatId ")" q " is established.")
							TBot.SendReply(username, msg.Text)
						}

				}

		}

		TBot.print("set new offset: " messages[MsgMax].updateID)
		old := TBot.GetupDatesOpt.offset
		TBot.GetupDatesOpt.offset := messages[MsgMax].updateID + 1
		TBot.print("last offset: " old ", new: " TBot.GetupDatesOpt.offset)



}
 */


class TelegramBot {

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; setup/handle bots
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	__New(settings) {                                                                          	; setup a new bot

			; Settings are: "BotRealName", "updateInterval" in seconds, "CallbackFunc" "OnMessage_TelegramInbox" , log:=true, debug:=false

			this.BotName        	:= settings.BotName
			this.Token               	:= settings.Token
			this.ChatId               	:= settings.ChatId
			this.LastOffset          	:= settings.LastOffset
			this.LastMsgDate     	:= settings.LastMsgDate
			this.URL			    	:= "https://api.telegram.org/bot" this.Token "/"
			this.CallbackFunc  	:= settings.CallbackFunc
			this.ShortName      	:= RegExReplace(settings.BotName, "\_*bot", "")
			this.updateInterval    	:= settings.updateInterval
			this.dialogInterval  	:= settings.dialogInterval
			this.DialogTimeout	:= settings.DialogTimeout
			this.chats					:= Object()

			; for GetUpdates() api calls
			this.GetUpdatesOpt:={	"offset"  	: settings.offset
											,	"updlimit"	: (settings.updlimit	? settings.updlimit : 100)
											, 	"timeout"	: (settings.timeout	? settings.timeout	: 0)}

			this.debug         	:= settings.debug
			this.filterReady   	:= 0

			For i, schar in StrSplit("!§$%&/()=?\/```´{}[]@*+~'-_:.,;")
				this.escaped .= "\" schar

	}

	__Delete() {

		  ; stop
			If IsFunc(this.UpdatesTimer) {
				timer := this.UpdatesTimer
				SetTimer, % timer, Delete
			}
			this := ""

	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; get/set variables in TelegramBot object
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	_Get(aName) {
		return this[aName]
	}

	_Set(aName, aValue) {

		aValueLast := this[aName]
		Switch aName {

			Case "updateInterval":
				this.SetUpdatesInterval(aValue)

			Default:
				this[aName] := aValue

		}

	return aValueLast
	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; user settings gui's
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	BotsConfigGui(BotConfigFile) {

		static T1, T2, T3, BName, BAdd, BDel, BToken, BReply, BSave, BDiscard, Bots, BotNames, botNr

		Bots := cJSON.Load(FileOpen(BotConfigFile, "r").Read())
		For botNr, BotObj in Bots
				BotNames .= A_Index ": " BotObj.Name "|"

		botNr:=1

		FontSize 	:= (A_ScreenWidth > 1920) ? 12 : 8
		FontStyle1	:= "s" FontSize+4	" cBlack	 Bold     	q5"
		FontStyle2	:= "s" FontSize+1	" cWhite Bold     	q5"
		FontStyle3	:= "s" FontSize   	" cBlack	 Normal	q5"
		FontStyle4	:= "s" FontSize-1  	" cWhite Normal	q5"

		Gui, BC: -DPIScale +HWNDhBC
		Gui, BC: Margin, 10, 10
		Gui, BC: Color	, 4EA6E0, c8FC8F2

		Gui, BC: Font	, % FontStyle1, Lucida Sans Unicode
		Gui, BC: Add	, ListBox, x90 ym r1                     	, % RTrim(BotNames, "|")

		Gui, BC: Font	, % FontStyle3
		Gui, BC: Add	, Button	, x+20              	 vBAdd 	, % "Add a bot"
		Gui, BC: Add	, Button	, x+20              	 vBDel 	, % "Delete this bot"

		Gui, BC: Font	, % FontStyle2
		Gui, BC: Add	, Text	, xm y+25  vT1               	, % "name"
		GuiControlGet, p, BC: Pos, T1
		Gui, BC: Font	, % FontStyle3
		Gui, BC: Add	, Edit		, % "x90 y" pY " w800 r1 	 vBName"	,	% Bots[botNr].Name

		Gui, BC: Font	, % FontStyle2
		Gui, BC: Add	, Text	, xm  vT2                          	, % "token"
		GuiControlGet, p, BC: Pos, T2
		Gui, BC: Font	, % FontStyle3
		Gui, BC: Add	, Edit		, % "x90 y" pY " w800 r1 vBToken"	, % Bots[botNr].Token

		Gui, BC: Font	, % FontStyle2
		Gui, BC: Add	, Text	, xm  vT3                         	, % "reply to"

		Gui, BC: Font	, % FontStyle4
		GuiControlGet, p, BC: Pos, T3
		Gui, BC: Add	, Text	, % "xm y" pY+pH+3 " w" pW " Center" 	, % "(filter)"
		Gui, BC: Font	, % FontStyle3
		Gui, BC: Add	, Edit 	, % "x90 y" pY " w800 r20 vBReply"  	, % ""

		Gui, BC: Add	, Button	, x90              	 vBSave     	, % "Save Config"
		Gui, BC: Add	, Button	, x+20          	 vBDiscard	, % "Discard Config"
		Gui, BC: Show	, AutoSize, Telegram bot configuration

	return

	BCGuiClose:
	BCGuiEscape:
		Gui, BC: Destroy
	return

	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; timer methods
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	StartUpdates() {	                                                                            	; starting GetUpdates calls with a given interval

		; starting onmessage callback if wanted
			If this.CallbackFunc && IsFunc(this.CallbackFunc) {
					Gui, OnMsg: +HWNDhWnd
					Gui, OnMsg: +LastFound
					this.OnMsgCallbackhWnd := hWnd
					OnMessage(0x5555, this.CallbackFunc)
			} else {
					throw "Callback is not a function.`nThis script will be terminated"
					ExitApp
			}

		; starting the getupdates timer method
			timer := this.UpdatesTimer := ObjBindMethod(this, "UpdatesHandler")
			SetTimer, % timer, % (this.updateInterval * 1000)
			this.debug ? this.print("update interval is set to " this.updateInterval "s") : ""

		; get first updates before starting the timer
			this.UpdatesHandler()

	}

	StopUpdates() {	                                                                               	; stop GetUpdates calls

		; stop sendmessage callbacks
		OnMessage(0x5555, "")

		; stop timer's
		timer := this.UpdatesTimer
		SetTimer % timer, Off
		RestoreTimer := this.RestoreTimer
		SetTimer, % RestoreTimer, Delete

		Gui, OnMsg: Destroy

		this.debug ? this.print("getUpdates stopped") : ""

	}

	SetUpdatesInterval(new_updateInterval) {                                          	; set update interval

		; if new_updateInterval is zero the timer function ist stopped

		this.updateInterval := new_updateInterval
		timer := this.UpdatesTimer
		SetTimer, % timer, % this.interval ? (this.interval * 1000) : "Off"
		this.debug ? this.print("update interval is changed to " this.interval "s") : ""

	}

	UpdatesHandler() {                                                                        	; timed function to get updates from telegram

		static tRound

		tRound++
		this.debug ? this.print(" Timer round: " tRound) : ""

		response := this.GetUpdates()
		If response.ok
			this.FilterSenders(response.result)

	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; filter methods
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	FilterSenders(chats) {                                                                        	; filters by sender who is answered

	  ; function uses a white list of users to auto answer (must match ChatID and username)

		If !IsObject(this.messages)
			this.messages := Array()

		For idx, body in chats		{

			If      	body.hasKey("message")     	{

				message 	:= body.message
				tgram		:= {	 "updateID"	: body.update_id
										, "fromID"  	: message.from.id
										, "chatID"   	: message.chat.id
										, "date"      	: message.date
										, "type"       	: message.chat.type
										, "username"	: message.from.username
										, "isbot"      	: message.from.is_bot ? true : false}

				If message.haskey("text") {
					tgram.content 	:= "text"
					tgram.text		   	:= message.text
				}

			}
			else if	body.haskey("channel_post") 	{

				message 	:= body.channel_post
				tgram		:= {	 "updateID"	: body.update_id
										, "chatID"   	: message.chat.id
										, "chatTitle"   	: message.chat.title
										, "date"      	: message.date
										, "type"       	: message.chat.type
										, "username"	: message.chat.username}

				If message.haskey("photo") {
					tgram.content 	:= "image"
					tgram.photo   	:= message.photo
				}

			}

		; build a new message object with the only needed senders to reply to
			For rIdx, toReply in this.replyTo				{
				If (tgram.ChatId = toReply.ChatId) && (tgram.username = toReply.username) {
					this.messages.Push(tgram)
					break
				}
			}

		}

		;
		If this.messages.Count() {

		  ; callback by send a message to call a user defined function
			SendMessage, 0x5555, 1, 0,, % "ahk_id " this.OnMsgCallbackhWnd

		  ; set timer interval to a faster response time and set a timer to fallback to the long polling interval
			this.SetUpdatesInterval(this.dialogInterval)
			this.RestoreTimer := ObjBindMethod(this, "SetUpdatesInterval", this.updateInterval)
			RestoreTimer := this.RestoreTimer
			DialogTimeout := -1 * this.DialogTimeout * 1000

				;SetTimer, % RestoreTimer, % DialogTimeout

		}

		this.filterReady += 1

	}

	GetNewMessages() {                                                                     	; get prepared FilterSenders messages
	return this.messages
	}

	GetUser(ChatId) {                                                                         	; get data from user (replyTo configuration!)
		For idx, sender in this.replyTo
			If (sender.ChatId = ChatId)
				return sender
	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; chat
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	newChat(username, chatID) {	; speichert Daten zu laufenden Chats

		If !IsObject(this.chats)
			this.chats := Object()

		If !this.chats.haskey(username)
			this.chats[username] := chatId

	}

	userChat(username) {
		return this.chats[username].chatId
	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; reply user methods
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	newReply(username, chatID, userAuth) {

		If !IsObject(this.reply)
			this.reply := Object()

		If !IsObject(this.reply[username]) {
			this.reply[username] := {"chatID":chatID, "auth":userAuth}
			return 1
		}
		else if this.reply.haskey(username)
			return 2

		return 0
	}

	SendReply(username, msgText) {



		this.debug ? this.print("reply service receiving: " q msgText q " from user: " username) : ""

		If RegExMatch(msgText, "i)^\s*\#(?<cmd>[a-zäöüß\s" escaped "]+)", chat_) {
				;MsgBox, % "u: " this.chat.username "`nmT: " messageText

				this.debug ? this.print("receiving command (" q chat_cmd q ") from user: " this.reply[username].username) : ""
				this.cmd := ObjBindMethod(this, chat_cmd, this.reply[username].ChatId)
				fn := this.cmd
				;~ If IsFunc(this[chat_cmd]) { ; ??
				If IsFunc(fn) {
					 SetTimer, % fn, -10
				} else {
					this.debug ?  ? this.print("unknown command (" q chat_cmd q ") received from user: " this.chats[ChatId].username) : ""
				}
		} else {
				this.debug ? this.print("no command received from user: " this.chats[ChatId].username) : ""
		}


	}

	cmds(ChatId) {	; sends inline keyboard with avaible commands to user

		;result := this.SendText(ChatId, "Du hast mich gefragt was ich kann....`n`nNichts ist meine Antwort.")
		result := this.InlineButtons("Was möchtest Du tun?", ChatId)
		this.print("result: " result.result)

	}

	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; debug methods
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	print(Text:="", Clear:=0, LineBreak:=1, Exit:=0) {                             	; print to SciTE output pane for debugging

		static LinesOut := 0

		;If !this.Debug
		;	return

		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

		If InStr(Text, "ShowLinesOut") {
			SciObj.Output("SciteOutput function has printed " LinesOut " lines.`n")
			return
		}

		If Clear || ((StrLen(Text) = 0) && (LinesOut = 0))
			SendMessage, SciObj.Message(0x111, 420)                   	;If clear=1 Clear output window
		If LineBreak
			Text .= "`r`n"                                                                	;If LineBreak=1 append text with `r`n

		SciObj.Output("[" this.ShortName "] " Text)                             	;send text to SciTE output pane

		LinesOut += StrSplit(Text, "`n").MaxIndex()

		If (Exit=1) 	{
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Telegram api calls wrapper
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	GetMe() {                                                                                       	;-- get information about the bot itself
		query := this.URL "getMe"
	return this.HttpPost(query, "application/json")
	}

	GetUpdates() {		                                                                         	;-- get updates from bot
		query := Format(this.URL "getupdates?offset={1}&limit={2}&timeout={3}", this.GetUpdatesOpt.offset, this.GetUpdatesOpt.updlimit, this.GetUpdatesOpt.timeout)
	return this.HttpPost(query, "application/json")
	}

	InlineButtons(msg, ChatId, mtext:="") {                                           	;-- add inline buttons
		keyb=									; json string:
		(
			{"inline_keyboard":[ [{"text": "Hausbesuch eintragen" , "callback_data" : "Hausbesuch"}, {"text" : "Befunde lesen",  "callback_data" : "Befunde"} ] ], "resize_keyboard" : true }
		)
		query := Format(this.URL "sendMessage?text={1}&chat_id={2}&reply_markup={3}", msg, ChatId, keyb)
		Return this.HttpPost(query, "application/json")
	}

	ShowKeyboard(kb_type, key_arr, ChatId, mtext:="") {                        	;-- sends query for display an inline- oder replykeyboard

		q := Chr(0x22)
		;keyb:= "{" q kb_type "_keyboard" q ":"[ [{"text": "Hausbesuch eintragen" , "callback_data" : "Hausbesuch"}, {"text" : "Befunde lesen",  "callback_data" : "Befunde"} ] ], "resize_keyboard" : true }

		query := Format(this.URL "sendMessage?text={1}&chat_id={2}&reply_markup={3}", msg, ChatId, keyb)
		Return this.HttpPost(query, "application/json")
	}

	RemoveKeyb(msg, ChatId, mtext="")  {		                                     	;-- remove custom keyboard
		keyb		:= "{" q "remove_keyboard" q " : true }"
		query	:= Format(this.URL "sendMessage?text=Keyboard removed&chat_id={1}&reply_markup={2}", tfrom_id, keyb)
		Return this.HttpPost(query, "application/json")
	}

	SendDocument(chatID, file, caption:="" )	{                                   	;-- media send: image
		;-- you could add more options; compare the Telegram API docs
		return this.UploadFormData(Format(this.URL "sendDocument?caption={1}&parse_mode=html", caption), {"chat_id" : chatID, "document" : [file] })
	}

	SendDocumentByCurl(chatID, file)	{                                                 	;-- media send: documents
		;-- you could add more options; compare the Telegram API docs
		; curl -X  POST "https://api.telegram.org/bot123456:abcde1234ABCDE/sendDocument" -F chat_id=-419426123 -F document="@/home/mirco/misto.txt"
		Run, % A_ScriptDir "\lib\curl.exe -X POST " q this.URL "sendDocument" q " -F chat_id=" ChatId " -F document=" q "@" file q
		return ErrorLevel ;this.UploadFormData(Format(this.URL "sendDocument?caption={1}", caption), {"chat_id" : ChatId, "document" : [file] })
	}

	SendPhoto(chatID, file, caption:="" )	{                                           	;-- media send: documents pdf max.size 50MB
		;-- you could add more options; compare the Telegram API docs
		return this.UploadFormData(Format(this.URL "sendPhoto?caption={1}", caption), {"chat_id" : ChatId, "photo" : [file] })
	}

	SendText(ChatId, textMsg) {			                                            		;-- send a text message
	return this.HttpPost(Format(this.URL "sendMessage?chat_id={1}&text={2}", chatID, this.uriEncode(textMsg)), "application/json")
	}

	SomeFunction(botToken, msg, from_id, mtext) {
		if (msg.callback_query.id)	{					;  After the user presses a callback button, Telegram clients will display a progress bar until you call answerCallbackQuery
			query := Format(this.URL "answerCallbackQuery?text=Notification&callback_query_id={1}&show_alert=true", msg.callback_query.id)
			response := this.HttpPost(query, "application/json")
		}
		text := "You chose " mtext
		this.SendText(from_id, text)        ;  url encoding for messages :  %0A = newline (\n)   	%2b = plus sign (+)		%0D for carriage return \r        single Quote ' : %27			%20 = space
	return response
	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; Telegram api helper functions
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	DownloadFile(FilePath) {
        query   := this.URL . StrSplit(FilePath,"/").2
		;InetGet(Query, fileName)
	Return True
	}

	GetFilePath(FileID) {                                                                        	; get the path of a file on Telegram Server by its File ID
		response := this.HttpPost(this.URL "getFile?fileid=" FileID, "application/json")
	return response.result.filepath
	}
	;}

	;----------------------------------------------------------------------------------------------------------------------------------------------
	; HTTP Methods
	;----------------------------------------------------------------------------------------------------------------------------------------------;{
	HttpPost(query, content_Type, waitResponse:=false) {

		WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
		WinHTTP.Open("POST", query, 0)
		WinHTTP.SetRequestHeader("Content-Type", content_Type)
		WinHTTP.Send()

		If waitResponse
			WinHTTP.WaitForResponse()

		response := cJSON.Load(this.lastResponse := WinHTTP.ResponseText)

		If !response.Ok && this.debug
			this.print(cJSON.Dump(response, 1) "`n   // " query)

		WinHTTP := ""												; free COM object

	return response
	}

	uriEncode(str) { ; v 0.3 / (w) 24.06.2008 by derRaphael / zLib-Style release
		b_Format := A_FormatInteger
		data := ""
		SetFormat, Integer, H
		Loop,Parse,str
			if ((Asc(A_LoopField)>0x7f) || (Asc(A_LoopField)<0x30) || (asc(A_LoopField)=0x3d))
				data .= "%" . ((StrLen(c:=SubStr(ASC(A_LoopField),3))<2) ? "0" . c : c)
			Else
				data .= A_LoopField
		SetFormat, Integer, % b_format
	return data
	}

	UriEncodeF(Uri, RE="[0-9A-Za-z]") {  ; F = Fast
	VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0), StrPut(Uri, &Var, "UTF-8")
	While Code := NumGet(Var, A_Index - 1, "UChar")
		Res .= (Chr:=Chr(Code)) ~= RE ? Chr : Format("%{:02X}", Code)
	Return, Res
	}

	UploadFormData(url_str, objParam)	{						                       	;-- Upload multipart/form-data

		CreateFormData(postData, hdr_ContentType, objParam)
		WinHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
		WinHTTP.Open("POST", url_str, true)
		WinHTTP.SetRequestHeader("Content-Type", hdr_ContentType)
		; WinHTTP.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko")				; ???????
		WinHTTP.Option(6) := False ; No auto redirect
		WinHTTP.Send(postData)
		WinHTTP.WaitForResponse()
		jsonResponse := WinHTTP.ResponseText
		WinHTTP :=	""											; free COM object

	return cJSON.Load(jsonResponse)							; will return a JSON string that contains, among many other things, the file_id of the uploaded file
	}
	;}


}

class ReplyService extends TelegramBot {

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; setup reply services
		;----------------------------------------------------------------------------------------------------------------------------------------------;{
		__NEW(BotName, ChatId, username, auth) {                                      	; setup a new reply service for a specific username

				this.base.reply := Object()
				this.base.reply.ChatId := ChatId
				this.base.reply.username:= username
				this.base.reply.auth:= auth
				this.base.reply.Token := this.Token
				;this.reply.bot :=

		}

		__Delete(ChatId) {

				this.base.reply[ChatId] := ""

		}
		;}

		;----------------------------------------------------------------------------------------------------------------------------------------------
		; message content parsers/handlers
		;----------------------------------------------------------------------------------------------------------------------------------------------;{

		GetToken() {

				MsgBox, % "T1: " this.Token "`nT2: " this.reply.Token

		}



		;}



	}



; CreateFormData() by tmplinshi, AHK Topic: https://autohotkey.com/boards/viewtopic.php?t=7647
; Thanks to Coco: https://autohotkey.com/boards/viewtopic.php?p=41731#p41731
; Modified version by SKAN, 09/May/2016
; https://www.autohotkey.com/boards/viewtopic.php?t=67426

CreateFormData(ByRef retData, ByRef retHeader, objParam) {
	New CreateFormData(retData, retHeader, objParam)
}

Class CreateFormData {

	__New(ByRef retData, ByRef retHeader, objParam) {

		local CRLF := "`r`n", i, k, v, str, pvData
		; Create a random Boundary
		Local Boundary := this.RandomBoundary()
		Local BoundaryLine := "------------------------------" . Boundary

		this.Len := 0 ; GMEM_ZEROINIT|GMEM_FIXED = 0x40
		this.Ptr := DllCall( "GlobalAlloc", "UInt",0x40, "UInt",1, "Ptr"  )          ; allocate global memory

	  ; Loop input paramters
		For k, v in objParam
			If IsObject(v) {
				For i, FileName in v	{
					str := BoundaryLine . CRLF
					     . "Content-Disposition: form-data; name=""" . k . """; filename=""" . FileName . """" . CRLF
					     . "Content-Type: " . this.MimeType(FileName) . CRLF . CRLF
				  this.StrPutUTF8( str )
				  this.LoadFromFile( Filename )
				  this.StrPutUTF8( CRLF )
				}
			} Else {
				str := BoundaryLine . CRLF
				     . "Content-Disposition: form-data; name=""" . k """" . CRLF . CRLF
				     . v . CRLF
				this.StrPutUTF8( str )
			}

		this.StrPutUTF8( BoundaryLine . "--" . CRLF )

      ; Create a bytearray and copy data in to it.
		retData := ComObjArray( 0x11, this.Len ) ; Create SAFEARRAY = VT_ARRAY|VT_UI1
		pvData  := NumGet( ComObjValue( retData ) + 8 + A_PtrSize )
		DllCall( "RtlMoveMemory", "Ptr",pvData, "Ptr",this.Ptr, "Ptr",this.Len )

		this.Ptr := DllCall( "GlobalFree", "Ptr",this.Ptr, "Ptr" )                   ; free global memory

    retHeader := "multipart/form-data; boundary=----------------------------" . Boundary
	}

	StrPutUTF8(str) {
    Local ReqSz := StrPut( str, "utf-8" ) - 1
    this.Len += ReqSz                                  ; GMEM_ZEROINIT|GMEM_MOVEABLE = 0x42
    this.Ptr := DllCall( "GlobalReAlloc", "Ptr",this.Ptr, "UInt",this.len + 1, "UInt", 0x42 )
    StrPut( str, this.Ptr + this.len - ReqSz, ReqSz, "utf-8" )
  }

	LoadFromFile(Filename) {
    Local objFile := FileOpen( FileName, "r" )
    this.Len += objFile.Length                     ; GMEM_ZEROINIT|GMEM_MOVEABLE = 0x42
    this.Ptr := DllCall( "GlobalReAlloc", "Ptr",this.Ptr, "UInt",this.len, "UInt", 0x42 )
    objFile.RawRead( this.Ptr + this.Len - objFile.length, objFile.length )
    objFile.Close()
  }

	RandomBoundary() {
		str := "0|1|2|3|4|5|6|7|8|9|a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z"
		Sort, str, D| Random
		str := StrReplace(str, "|")
	return SubStr(str, 1, 12)
	}

	MimeType(FileName) {
		n := FileOpen(FileName, "r").ReadUInt()
		return (n           	= 0x474E5089	) ? "image/png"
				: (n           	= 0x38464947	) ? "image/gif"
				: (n&0xFFFF 	= 0x4D42       	) ? "image/bmp"
				: (n&0xFFFF 	= 0xD8FF      	) ? "image/jpeg"
				: (n&0xFFFF 	= 0x4949     	) ? "image/tiff"
				: (n&0xFFFF 	= 0x4D4D    	) ? "image/tiff"
				: (n&0xFFFF 	= 0x10000   	) ? "image/ico"
				: "application/octet-stream"
}

}


; 	Includes                                                                   	;{

;------------ allgemeine Bibliotheken ---------------------------------
;~ #include %A_ScriptDir%\
;~ #Include class_cJSON.ahk
;~ #Include SciTEOutput.ahk

;}


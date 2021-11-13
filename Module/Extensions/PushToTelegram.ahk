; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                                       Addendum PushToTelegram
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
/*

	Das Weiterleiten einer EMail nach Telegram geschrieben nur in VisualBasic war für mich nicht möglich.
	Wie liest man Daten aus einer ini Datei? Wie encodiert man einen String nach UTF-8?
	Jede Änderung aufgrund eines Fehlers im VBA Skriptes hatte weitere Fehler zu Folge. Also sind hier
	nur kurze VBA Skripte

	Die kurzen VBA Skripte sind das was fehlerfrei ging. Diese zu den Makro's in Outlook zu packen (Module).
	1. PushToTelegram - 		für manuelles Versenden von Nachrichten (Nachricht selektieren und Makro aufrufen)
	2. AutoPushToTelegram - 	als Skript mit Outlookregeln verwendbar. Ein getriggertes Event startet
											dieses Skript und die Mail wird per Telegram versendet.

	Das Skript leitet nicht nur eine EMails per Telegram weiter, es übernimmt vorrübergehend das
	Kopieren der zuvor gespeicherten Faxanhänge aus dem per Argument übergebenen Ordner in das
	Verzeichnis für die Befunde.

	Was wird benötigt:

	⓵. ein vorbereiteter Telegram Bot, sein Token und eine ChatID um ihm Nachrichten senden zu können
	⓶. dieses Skript einmal aufrufen um andere Einstellungen machen zu können
	⓷. den unteren VBA Code in die Outlook Makros kopieren und gewünschte Regeln erstellen

Public Sub PushToTelegram()

    Dim msg As MailItem
    Dim mAtts As Attachments
    Dim mAtt As Attachment
    Dim cmdln As String
    Dim text As String
    Dim PushToPath As String

    Dim retVal

    PushToPath = "C:\Program Files\AutoHotkey\Autohotkey.exe /f ""M:\Praxis\Skripte\Skripte Neu\Addendum für AlbisOnWindows\Module\Extensions\PushToTelegram.ahk"" "

    For Each msg In Application.ActiveExplorer.Selection

        'Set mAtts = msg.Attachments
        text = " ""EMail von: " & msg.SenderEmailAddress & " [" & msg.SenderName & " " & msg.Sender & "]" & vbCrLf _
				& "erhalten um:" & vbTab & Now & vbCrLf & "Betreff:   " & vbTab & msg.Subject & vbCrLf & "Body:      " & vbTab & msg.Body & ""

        'Send Message to AHK Script
        cmdln = PushToPath & "" & text & ""
        retVal = Shell(cmdln, vbNormalNoFocus)


    Next msg

    'SendToTelegram Mail


End Sub

' frisch eingetroffene Mails per Event zu Telegram weiterleiten
Public Sub AutoPushToTelegram(mail As Outlook.MailItem)

    Dim mAtts As Attachments
    Dim mAtt As Attachment
    Dim cmdln As String
    Dim retVal

    Set mAtts = mail.Attachments

    text = " ""EMail von: " & msg.SenderEmailAddress & " [" & msg.SenderName & " " & msg.Sender & "]" & vbCrLf & "erhalten um:" & vbTab & Now & vbCrLf & "Betreff:   " & vbTab & msg.Subject & vbCrLf & "Body:      " & vbTab & msg.Body & ""

    'Send Message to AHK Script
    cmdln = PushToPath & "" & text & ""
    retVal = Shell(cmdln, vbNormalNoFocus)


End Sub


	begonnen am: 03.11.2021
	letzte Änderung: 03.11.2021

 */

 ; Variablen Albis Datenbankpfad / Addendum Verzeichnis                     	;{

	#SingleInstance               	, Force
	#MaxThreads                  	, 200
	#MaxThreadsBuffer       	, On
	#KeyHistory                  	, 0
	SetBatchLines                	, -1
	ListLines                        	, Off

	global Props, PatDB, adm

  ; Pfad zu Albis und Addendum ermitteln
	RegRead, AlbisPath, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\CG\ALBIS\Albis on Windows, Installationspfad
	RegExMatch(A_ScriptDir, ".*(?=\\Module)", AddendumDir)

  ; Addendum Objekt
	adm                 	:= Object()
	adm.Dir            	:= AddendumDir
	adm.Ini              	:= AddendumDir "\Addendum.ini"
	adm.DBPath      	:= AddendumDir "\logs'n'data\_DB"
	adm.AlbisPath    	:= AlbisPath
	adm.AlbisDBPath 	:= adm.AlbisPath "\DB"
	adm.compname	:= StrReplace(A_ComputerName, "-")                                                                 	; der Name des Computer auf dem das Skript läuft
	adm.Telegram  	:= Object()

   ; es wurde kein Messagetext mitgesendet
	If !A_Args.Length()
		ExitApp

	messageToSend      	:= A_Args[1]
	FolderToMoveFrom  	:= A_Args[2]
	Attachements        	:= StrSplit(A_Args[3], "|")

  ; ein Aufruf nur mit dem Verzeichnisnamen initialisiert den Ini-Pfad der Funktion
	workini := IniReadExt(adm.Ini)

  ; Einstellungen für PushToTelegram
	adm.ManageAttachments :=  IniReadExt(compname, "PushToTelegram_ManageAttachments", "Ja")

 ;}

	If  RegExMatch(messageToSend, "\-\-AttachmentManagement") {

	  ; Skript soll kein Management der Faxanhänge durchführen dann wird es hier beendet
		If !adm.ManageAttachments
			ExitApp

	  ; Erstaufruf? Einstellungen setzen und überprüfen
		FolderToMoveFrom := RegExReplace(FolderToMoveFrom, "\\\s*$") . "\"
		If !InStr(FileExist(FolderToMoveFrom), "D") {
			MsgBox, 0x1000,  Addendum PushToTelegram, % "Da stimmt etwas nicht mit dem lokalen Speicherpfad für die Faxe:`n" . FolderToMoveFrom
			ExitApp
		}
		adm.BefundOrdner :=  IniReadExt("ScanPool", "BefundOrdner")
		If (!adm.BefundOrdner || InStr(adm.BefundOrdner, "Error") || !InStr(FileExist(adm.BefundOrdner), "D")) {
			MsgBox, 0x1000, Addendum PushToTelegram, % "Das Skript hat von Outlook die Aufforderung erhalten lokal gespeicherte Faxe zu verarbeiten....."
			ExitApp
		}

		files := Array()
		Loop, Files, % FolderToMoveFrom . "Fax *.pdf"
			files.Push(A_LoopFileFullPath)

		/*
		For idx, filepath in files {

			SplitPath, filePath, fName, fPath, fExt, fNoExt
			If !FileExist(filepath)
				continue
		  ; identische Dateinamen = identische Dateien ?
			If FileExist(filePathMoveTo := adm.BefundOrdner "\" fname) {

			 ; Dateigrößen oder CRC stimmen nicht überein
				sizeIsEqual	:= !FileSizeIsEqual(filePathMoveTo, filePath)
				crcIsEqual 	:= !FileCRCIsEqual(filePathMoveTo, filePath)
				If (!sizeIsEqual || !crcIsEqual) {

				  ; neuen Dateinamen generieren
					nextfile := false
					while FileExist(filePathMoveTo := adm.BefundOrdner "\" fNoExt "_" A_Index "." fExt) {
						If (FileSizeIsEqual(filePathMoveTo, filePath) && FileCRCIsEqual(filePathMoveTo, filePath)) {
							nextfile := true
							break
						}
					}
					If nextfile
						continue

				}
				else If (sizeIsEqual && crcIsEqual) {
					continue
				}

			}
			 */

		FileMove, % filepath, % filePathMoveTo
		TrayTip, PushToTelgram, % "Ein Fax mit dem Namen: " fname "`n nach " adm.BefundOrdner " verschoben.", 2

		;~ }

		ExitApp
	}

  ; Mailbot laden
	MailBot := LoadMailBot()

  ; Mail Daten und Betreff und Text für Lesbarkeit umformen
	messageToSend := TransformMailToTelgram(messageToSend)

 ; Mailinhalt senden und das wars auch schon
	Telegram.SendText(adm.Telegram.MailBot.Token, adm.Telegram.MailBot.ChatId, messageToSend)


ExitApp


LoadMailBot()                                                                                    		{        	;-- Einstellungen laden oder bei erster Skriptausführung Einstellungen abfragen

	; ist ein Mailbot eingetragen?
	MailBot	:= IniReadExt("Telegram", "MailBot")

  ; Bot Daten aus der Ini lesen
	LoadTelegramBots()

 ; wieviele Bots gibt es denn?
	If !IsObject(adm.Telegram.MailBot) && !adm.Telegram.Bots.Count() {

		BotCall := IniReadExt("Telegram", "MailBotCall_Failed", 1)
		TrayTip, % A_ScriptName, % "Es sind noch keine Daten zu Telegram Bots hinterlegt"

		BotName:
		InputBox, BotName, % A_ScriptName, % "Zunächst benötige ich den Namen des vorbereiteten Bots",, 400, 150,,, % BotName
		If !RegExMatch(BotName, "i)bot\s*$") {
			MsgBox, 0x1003, % A_ScriptName, % "Der Name eines Telegram-Bot endet auf bot`n"
																.	 "Wenn Du später eine korrekte Ausführung erwartest,`nsollten wir uns jetzt an die vorgebene Namenskonvention halten.`n "
																. 	 "Bitte korrigiere Deine Eingabe. Ich gebe Dir noch eine Chance!"
			fail := 1
			IfMsgBox, Cancel
				ExitApp
			IfMsgBox, No
				ExitApp
			goto BotName
		}
		IniWrite, BotName, % adm.Ini, % "Telegram", % "BotName"

		BotToken:
		InputBox, BotToken, % A_ScriptName, % "Jetzt brauche ich das BotToken",, 400, 150,,, % BotToken
		If (!RegExMatch(BotToken, "i)\d+\:[\w_\-\:\;\.\\+,]{20,}$") || StrLen(BotToken) < 30) {
			msg1 := "Mit Deinem BotToken stimmt etwas nicht: " (StrLen(BotToken) < 30 ?  "es ist zu kurz" : "die vewendeten Zeichen entsprechen nicht den Vorgaben")
			msg2 := fail = 1 ? "So kommen wir nicht voran. Das Token muss korrekt sein!" : "Das Token sieht in etwa so aus: 831435972:HGABRALFIG3nmCia8sp3lirvfX_Sk3RUE4Az."
			msg3 := fail = 1 ? "So wirst Du nie fertig. Konzentration!" : "Okay noch eine Runde!"
			MsgBox, 0x1003, % A_ScriptName, % msg1 "`n"
																.	msg2 "`n "
																.	msg3 "`n"
			fail ++

			IfMsgBox, Cancel
				ExitApp
			IfMsgBox, No
				ExitApp
			goto BotToken
		}
		IniWrite, BotToken, % adm.Ini, % "Telegram", % BotName "_BotToken"

		msg1 := fail = 1 ? "besteht sie doch nur aus Zahlen" : fail = 2 ? "besteht sie doch nur aus 0,1,2,3....9" : "errinnert Sie uns doch an unseren Bankkontostand (1435454534)."
		ChatID:
		InputBox, ChatID, % A_ScriptName, % "Nichts leichter als eine ChatID, ",, 400, 150,,, % ChatID
		If (!RegExMatch(ChatID, "i)^\d+$") || StrLen(ChatID) < 8) {
			msg1 := "Mit Deiner ChatID stimmt etwas nicht: " (StrLen(ChatID) < 8 ?  "sie ist zu kurz!" : "Du hast nicht nur Zahlen verwendet!")
			msg2 := fail = 1 ? "So kommen wir nicht voran. Die ChatID muss korrekt sein!" : fail = 2 ? "Ist nicht allzu schwer" : "Okay, " fail ". Fail-Runde,  ich glaub weiter an Dich!"
			msg3 := fail = 1 ? "So wirst Du nie fertig. Konzentration!" : fail = 2 ? "Nochmal!" : fail= 2 ? "Du hast es drauf!" : "0123456789 ich hab Geduld!"
			MsgBox, 0x1003, % A_ScriptName, % msg1 "`n"
																.	msg2 "`n "
																.	msg3 "`n"
           fail ++

			IfMsgBox, Cancel
				ExitApp
			IfMsgBox, No
				ExitApp
			goto ChatID
		}
		IniWrite, ChatID, % adm.Ini, % "Telegram", % BotName "_ChatID"

	}

 ; welcher ist der Mailbot?
	If !adm.Telegram.MailBot {

		MailBot	:= IniReadExt("Telegram", "MailBot", adm.Telegram.Bots[1].BotName)
		If InStr(adm.Telegram.MailBot, "ERROR")
			MailBot	:= adm.Telegram.Bots[1].BotName

	} else
		MailBot := adm.Telegram.MailBot

  ; wir haben einen Mailbot


return MailBot
}
LoadTelegramBots() 												    								{      	;-- Telegram Bot Einstellungen

	MailBot	:= IniReadExt("Telegram", "MailBot")
	If (!InStr(MailBot, "ERROR") && StrLen(MailBot) > 0) {
		adm.Telegram.MailBot := ReadBots(MailBot)
		return
	}

	adm.Telegram.Bots := Array()
	BotNr := 0
	Loop {
		If InStr(BotName:=IniReadExt("Telegram", "Bot" BotNr+1), "Error")
			break
		adm.Telegram.Bots.Push(ReadBots(BotName))
		If (BotNr = 1)
			MailBot := BotName
		BotNr ++
	}

}
ReadBots(BotName)                                                                            		{        	;-- Telegram-Bot Daten
return {	"BotName"            	: BotName
		, 	"Token"                 	: IniReadExt("Telegram", BotName "_Token")
		, 	"ChatID"                 	: IniReadExt("Telegram", BotName "_ChatID")
		, 	"Active"                  	: IniReadExt("Telegram", BotName "_Active", 0)}
}


TransformMailToTelgram(messageToSend) 												{

  ; mehrfache Folgen von Zeilenendezeichen entfernen
	messageToSend := RegExReplace(messageToSend, "[\n\r]{2}[\n\r]{2}")
	messageToSend := RegExReplace(messageToSend, "([\n\r]{1,})[\n\r]{1,}", "$1")
	messageToSend := RegExReplace(messageToSend, "erhalten um:\s*", "`nerhalten um: ")
	messageToSend := RegExReplace(messageToSend, "Betreff:\s*", "`nBetreff: ")
	messageToSend := RegExReplace(messageToSend, "Body:\s*", "**`nBody:`n`n")

return messageToSend
}
class Telegram 																							{        	;-- mini Telegramklasse - finde meine andere nicht

	SendText(telegramBotKey, telegramChatId, textMessage) {				                    			;-- another way to send a message

		WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
		WinHTTP.Open("POST", Format("https://api.telegram.org/bot{1}/sendMessage?chat_id={2}&text={3}", telegramBotKey, telegramChatId, uriencode(textMessage)), 0)
		;WinHTTP.SetRequestHeader("Content-Type", "application/json")
		WinHTTP.Send()
		response := WinHTTP.responseText

		TrayTip,  % "Mail an " adm.Telegram.MailBot.BotName " weitergeleitet!", % "Telegram hat geantwortet: `n" response, 4

	Return response
	}

}


URIEncode(str, encoding := "UTF-8")  														{
   VarSetCapacity(var, StrPut(str, encoding))
   StrPut(str, &var, encoding)

   while code := NumGet(Var, A_Index - 1, "UChar")  {
      bool := (code > 0x7F || code < 0x30 || code = 0x3D)
      UrlStr .= bool ? "%" . Format("{:02X}", code) : Chr(code)
   }
   Return UrlStr
}
CreateFormData(ByRef retData, ByRef retHeader, objParam) 						{
	New CreateFormData(retData, retHeader, objParam)
}
CreateFormData_WinInet(ByRef retData, ByRef retHeader, objParam) 			{
	New CreateFormData(safeArr, retHeader, objParam)

	size := safeArr.MaxIndex() + 1
	VarSetCapacity(retData, size, 1)
	DllCall("oleaut32\SafeArrayAccessData", "ptr", ComObjValue(safeArr), "ptr*", pdata)
	DllCall("RtlMoveMemory", "ptr", &retData, "ptr", pdata, "ptr", size)
	DllCall("oleaut32\SafeArrayUnaccessData", "ptr", ComObjValue(safeArr))
}
Class CreateFormData 																				{

	__New(ByRef retData, ByRef retHeader, objParam) {

		CRLF := "`r`n"

		Boundary := this.RandomBoundary()
		BoundaryLine := "------------------------------" . Boundary

		; Loop input paramters
		binArrs := []
		fileArrs := []
		For k, v in objParam
		{
			If IsObject(v) {
				For i, FileName in v
				{
					str := BoundaryLine . CRLF
					     . "Content-Disposition: form-data; name=""" . k . """; filename=""" . FileName . """" . CRLF
					     . "Content-Type: " . this.MimeType(FileName) . CRLF . CRLF
					fileArrs.Push( BinArr_FromString(str) )
					fileArrs.Push( BinArr_FromFile(FileName) )
					fileArrs.Push( BinArr_FromString(CRLF) )
				}
			} Else {
				str := BoundaryLine . CRLF
				     . "Content-Disposition: form-data; name=""" . k """" . CRLF . CRLF
				     . v . CRLF
				binArrs.Push( BinArr_FromString(str) )
			}
		}

		binArrs.push( fileArrs* )

		str := BoundaryLine . "--" . CRLF
		binArrs.Push( BinArr_FromString(str) )

		retData := BinArr_Join(binArrs*)
		retHeader := "multipart/form-data; boundary=----------------------------" . Boundary
	}

	RandomBoundary() {
		str := "0|1|2|3|4|5|6|7|8|9|a|b|c|d|e|f|g|h|i|j|k|l|m|n|o|p|q|r|s|t|u|v|w|x|y|z"
		Sort, str, D| Random
		str := StrReplace(str, "|")
		Return SubStr(str, 1, 12)
	}

	MimeType(FileName) {
		n := FileOpen(FileName, "r").ReadUInt()
		Return (n        = 0x474E5089) ? "image/png"
		     : (n        = 0x38464947) ? "image/gif"
		     : (n&0xFFFF = 0x4D42    ) ? "image/bmp"
		     : (n&0xFFFF = 0xD8FF    ) ? "image/jpeg"
		     : (n&0xFFFF = 0x4949    ) ? "image/tiff"
		     : (n&0xFFFF = 0x4D4D    ) ? "image/tiff"
		     : "application/octet-stream"
	}

}
BinArr_FromString(str) 																				{

	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type  	:= 2 ; adTypeText
	oADO.Mode 	:= 3 ; adModeReadWrite
	oADO.Open
	oADO.Charset	:= "UTF-8"
	oADO.WriteText(str)

	oADO.Position	:= 0
	oADO.Type    	:= 1 ; adTypeBinary
	oADO.Position	:= 3 ; Skip UTF-8 BOM

return oADO.Read, oADO.Close
}
BinArr_FromFile(FileName) 																		{

	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Open
	oADO.LoadFromFile(FileName)

return oADO.Read, oADO.Close
}
BinArr_Join(Arrays*) 																					{

	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type  	:= 1 ; adTypeBinary
	oADO.Mode 	:= 3 ; adModeReadWrite
	oADO.Open
	For i, arr in Arrays
		oADO.Write(arr)
	oADO.Position	:= 0

return oADO.Read, oADO.Close
}
BinArr_ToString(BinArr, Encoding := "UTF-8") 											{

	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type  	:= 1 ; adTypeBinary
	oADO.Mode 	:= 3 ; adModeReadWrite
	oADO.Open
	oADO.Write(BinArr)

	oADO.Position := 0
	oADO.Type  	:= 2 ; adTypeText
	oADO.Charset	:= Encoding

return oADO.ReadText, oADO.Close
}
BinArr_ToFile(BinArr, FileName) 																{

	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Open
	oADO.Write(BinArr)
	oADO.SaveToFile(FileName, 2)
	oADO.Close

}


IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {          	;-- eigene IniRead funktion für Addendum

	/* Beschreibung

		-	beim ersten Aufruf der Funktion nur den Pfad und Namen zur ini Datei übergeben!
				workini := IniReadExt("C:\temp\Addendum.ini")

		-	die Funktion behandelt einen geschriebenen Wert welcher ein "ja" oder "nein" ist, als logische Wahrheitswerte (wahr und falsch).
		-	UTF-16-Zeichen (verwendet in .ini Dateien) werden per Default in UTF-8 Zeichen umgewandelt

		letzte Änderung: 22.10.2021

	 */

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^([A-Z]:\\|\\\\[A-Z]{2,}).*")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox, 0x1024
							, % "Addendum für AlbisOnWindows"
							, % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet."
							, 10
				ExitApp
			}
			WorkIni	:= SectionOrFullFilePath
			admDir	:= RegExMatch(WorkIni, "^([A-Z]:\\|\\\\[A-Z]{2,}).*?AlbisOnWindows", rxDir) ? rxDir : Addendum.Dir
			return WorkIni
		}

	; Workini ist nicht definiert worden, dann muss das komplette Skript abgebrochen werden
		If !WorkIni {
			MsgBox,, Addendum für AlbisOnWindows, %	"Bei Aufruf von IniReadExt muss als erstes`n"
																			. 	"der Pfad zur ini Datei übergeben werden.`n"
																			.	"Das Skript wird jetzt beendet.", 10
			ExitApp
		}

	; Section, Key einlesen, ini Encoding in UTF.8 umwandeln
		IniRead, OutPutVar, % WorkIni, % SectionOrFullFilePath, % Key
		If convert
			OutPutVar := StrUtf8BytesToText(OutPutVar)

	; Bearbeiten des Wertes vor Rückgabe
		If InStr(OutPutVar, "ERROR")
			If DefaultValue  { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				OutPutVar := DefaultValue
				IniWrite, % DefaultValue, % WorkIni, % SectionOrFullFilePath, % key
				If ErrorLevel
					TrayTip, % A_ScriptName, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]", 2
			}
			else return "ERROR"

		If InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "%\.exe$") && !RegExMatch(OutPutVar, "i)[A-Z]\:\\")
				return GetAppImagePath(OutPutVar)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}
StrUtf8BytesToText(vUtf8) 																			{      	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}
GetAppImagePath(appname) 																	{         	;-- Installationspfad eines Programmes

	headers:= {	"DISPLAYNAME"                  	: 1
					,	"VERSION"                         	: 2
					, 	"PUBLISHER"             	         	: 3
					, 	"PRODUCTID"                    	: 4
					, 	"REGISTEREDOWNER"        	: 5
					, 	"REGISTEREDCOMPANY"    	: 6
					, 	"LANGUAGE"                     	: 7
					, 	"SUPPORTURL"                    	: 8
					, 	"SUPPORTTELEPHONE"       	: 9
					, 	"HELPLINK"                        	: 10
					, 	"INSTALLLOCATION"          	: 11
					, 	"INSTALLSOURCE"             	: 12
					, 	"INSTALLDATE"                  	: 13
					, 	"CONTACT"                        	: 14
					, 	"COMMENTS"                    	: 15
					, 	"IMAGE"                            	: 16
					, 	"UPDATEINFOURL"            	: 17}

   appImages := GetAppsInfo({mask: "IMAGE", offset: A_PtrSize*(headers["IMAGE"]-1)})
   Loop, Parse, appImages, "`n"
	If Instr(A_loopField, appname)
		return A_loopField

return ""
}
GetAppsInfo(infoType) 																				{

	static CLSID_EnumInstalledApps := "{0B124F8F-91F0-11D1-B8B5-006008059382}"
		, IID_IEnumInstalledApps     	:= "{1BC752E1-9046-11D1-B8B3-006008059382}"

        , DISPLAYNAME            	:= 0x00000001
        , VERSION                    	:= 0x00000002
        , PUBLISHER                  	:= 0x00000004
        , PRODUCTID                	:= 0x00000008
        , REGISTEREDOWNER    	:= 0x00000010
        , REGISTEREDCOMPANY	:= 0x00000020
        , LANGUAGE                	:= 0x00000040
        , SUPPORTURL               	:= 0x00000080
        , SUPPORTTELEPHONE  	:= 0x00000100
        , HELPLINK                     	:= 0x00000200
        , INSTALLLOCATION     	:= 0x00000400
        , INSTALLSOURCE         	:= 0x00000800
        , INSTALLDATE              	:= 0x00001000
        , CONTACT                  	:= 0x00004000
        , COMMENTS               	:= 0x00008000
        , IMAGE                        	:= 0x00020000
        , READMEURL                	:= 0x00040000
        , UPDATEINFOURL        	:= 0x00080000

   pEIA := ComObjCreate(CLSID_EnumInstalledApps, IID_IEnumInstalledApps)

   while DllCall(NumGet(NumGet(pEIA+0) + A_PtrSize*3), Ptr, pEIA, PtrP, pINA) = 0  {
      VarSetCapacity(APPINFODATA, size := 4*2 + A_PtrSize*18, 0)
      NumPut(size, APPINFODATA)
      mask := infoType.mask
      NumPut(%mask%, APPINFODATA, 4)

      DllCall(NumGet(NumGet(pINA+0) + A_PtrSize*3), Ptr, pINA, Ptr, &APPINFODATA)
      ObjRelease(pINA)
      if !(pData := NumGet(APPINFODATA, 8 + infoType.offset))
         continue
      res .= StrGet(pData, "UTF-16") . "`n"
      DllCall("Ole32\CoTaskMemFree", Ptr, pData)  ; not sure, whether it's needed
   }
   Return res
}
 /*
GetFileSize(path, units="") 																			{         ;-- FileGetSize wrapper
	FileGetSize, fSize, % path, % units
return fSize
}
FileSizeIsEqual(filepath1, filepath2) 															{
return (GetFileSize(filepath1)=GetFileSize(filepath2) ? true:false)
}
LC_FileCRC32(sFile := "", cSz := 4) 															{
	Bytes := ""
	cSz := (cSz < 0 || cSz > 8) ? 2**22 : 2**(18 + cSz)
	VarSetCapacity(Buffer, cSz, 0)
	hFil := DllCall("Kernel32.dll\CreateFile", "Str", sFile, "UInt", 0x80000000, "UInt", 3, "Int", 0, "UInt", 3, "UInt", 0, "Int", 0, "UInt")
	if (hFil < 1)
		return hFil
	hMod := DllCall("Kernel32.dll\LoadLibrary", "Str", "Ntdll.dll")
	CRC32 := 0
	DllCall("Kernel32.dll\GetFileSizeEx", "UInt", hFil, "Int64", &Buffer), fSz := NumGet(Buffer, 0, "Int64")
	loop % (fSz // cSz + !!Mod(fSz, cSz))
	{
		DllCall("Kernel32.dll\ReadFile", "UInt", hFil, "Ptr", &Buffer, "UInt", cSz, "UInt*", Bytes, "UInt", 0)
		CRC32 := DllCall("Ntdll.dll\RtlComputeCrc32", "UInt", CRC32, "UInt", &Buffer, "UInt", Bytes, "UInt")
	}
	DllCall("Kernel32.dll\CloseHandle", "Ptr", hFil)
	SetFormat, Integer, % SubStr((A_FI := A_FormatInteger) "H", 0)
	CRC32 := SubStr(CRC32 + 0x1000000000, -7)
	DllCall("User32.dll\CharLower", "Str", CRC32)
	SetFormat, Integer, %A_FI%
	return CRC32, DllCall("Kernel32.dll\FreeLibrary", "Ptr", hMod)
}
FileCRCIsEqual(filepath1, filepath2) 															{
	fileCRC1 := LC_FileCRC32(filepath1)
	fileCRC2 := LC_FileCRC32(filepath2)
	SciTEOutput("1: " fileCRC1)
	SciTEOutput("2: " fileCRC2)
return (fileCRC1=fileCRC2 ? true : false)
}
 */

#include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
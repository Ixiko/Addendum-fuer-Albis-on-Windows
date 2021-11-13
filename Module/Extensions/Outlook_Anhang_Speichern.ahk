; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;                                          	** FAXEINGANG SPEICHERN **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;
;		Beschreibung:
;
; 		Inhalt:
;       Abhängigkeiten:
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Outlook_Anhang_Speichern.ahk Letzte Änderung:    	21.09.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#NoEnv
#Persistent
SetBatchLines, -1


; -------------------------------------------------------------------------------------------------------------------------------------------------------
; Variablen, Pfade und Einstellungen
; -------------------------------------------------------------------------------------------------------------------------------------------------------;{
		global compname	:= StrReplace(A_ComputerName, "-")                    	; der Name des Computer auf dem das Skript läuft
		global adm        	:= Object()
		adm.scriptname := RegExReplace(A_ScriptName, "\.ahk$")

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Pfad zu Addendum und den Einstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		RegExMatch(A_ScriptDir, "[A-Z]\:.*?AlbisOnWindows", AddendumDir)
		adm.ini	:= AddendumDir "\Addendum.ini"
		If !FileExist(adm.ini) {
			MsgBox, 0x1024, % adm.scriptname, % "Die Einstellungsdatei ist nicht vorhanden!`n[" adm.ini "]"
			ExitApp
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Dateipfade
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		workini	:= IniReadExt(adm.ini)
		adm.DBPath := IniReadExt("Addendum", "AddendumDBPath")           	; Datenbankverzeichnis
		If InStr(adm.DBPath, "ERROR") {
			MsgBox, 0x1000, % adm.scriptname, % "In der Addendum.ini ist kein Pfad für`ndie Addendum Datendateien hinterlegt!`n"
			ExitApp
		}
		If !FilePathCreate(adm.DBPath "\sonstiges") {
			MsgBox, 0x1000, % adm.scriptname, % "Datenbankpfad konnte nicht angelegt werden!"
			ExitApp
		}
		adm.BefundOrdner   	:= IniReadExt("ScanPool", "BefundOrdner")

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Outlook Attachment Filter-Einstellungen
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		adm.ol := Object()
		adm.ol.Folder            	:= IniReadExt(compname, "Outlook_Folder", "Posteingang")
		adm.ol.SubjectIn         	:= IniReadExt(compname, "Outlook_EMailSubjectFilter_In", "i)Fax\s")
		adm.ol.SubjectOut     	:= IniReadExt(compname, "Outlook_EMailSubjectFilter_Out", "########")
		adm.ol.SubjectData    	:= IniReadExt(compname, "Outlook_EMailSubject_Daten", "i)von\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)")
		adm.ol.AttachmentIn  	:= IniReadExt(compname, "Outlook_AttachmentFilter_In", ".*")
		adm.ol.AttachmentOut	:= IniReadExt(compname, "Outlook_AttachmentFilter_Out", "########")

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Outlook prüfen der COM Verbindung, Verbindung bleibt erhalten
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		ol := new Outlook(adm.DBPath, adm.ol)
		olStatus := ol.Connect()
		If (olStatus = "no connection") {
			MsgBox, 0x1000, % StrReplace(A_ScriptName, ".ahk")
						, % "Es konnte keine Verbindung zu Outlook hergestellt werden.`n"
						.	  "Bitte prüfen Sie ihre Outlookinstallation!"
			ExitApp
		}

	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	; Speicherpfad für Anhänge
	; - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Attachmentpath	:= IniReadExt(compname, "Outlook_AttachmentSavePath")
		If (InStr(Attachmentpath, "ERROR") || StrLen(Attachmentpath) = 0) {

			AskAttachmentSavePath:
			Attachmentpath := SelectFolder(3, "Speicherpfad für Outlook-Anhänge auswählen")
			MsgBox	, %  (!Attachmentpath	?	0x1005 : 0x1003)
						, % StrReplace(A_ScriptName, ".ahk")
						, % (!Attachmentpath 	? 	"Sie haben keinen Pfad gewählt.`nMöchten Sie die Speichpfadauswahl wiederholen?"
												:	"Gewählter Dateiordner:`n<< " Attachmentpath " >>`nSoll der Ordnerpfad übernommen werden?")
			IfMsgBox, Cancel
				ExitApp
			IfMsgBox, Retry
				goto AskAttachmentSavePath

			IniWrite, % Attachmentpath, % adm.ini, % compname, % "Outlook_AttachmentSavePath"

		}

	  ; wichtig!: der Ordner in welchem die Anhänge abgespeichert werden soll muss dem Outlook-Objekt übergeben werden
		ol.AttachmentSavePath := AttachmentPath

;}

; Outlook-Events abfangen
	ol.CatchMailEvents()

; Outlook-Events und Outlook Verbindung beenden
	;~ ol.Disconnect()


return

class Outlook {                                     ; Outlook Class for automated attachment filtering and extraction

		; this is just a basic class. At the moment with prepared specialization for Fritzbox fax attachments.
		; last modification: 21.09.2021

		__New(DBPath, Props, SaveFolder:="") {

			this.DBPath 	:= DBPath
			this.props  	:= Object()
			If !IsObject(Props) {
				this.props.olFolder             	:= "Posteingang"
				this.props.olSubjectIn         	:= "i)Fax\s"
				this.props.olSubjectOut       	:= "########"
				this.props.olSubjectData    	:= "i)von\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)"
				this.props.olAttachmentIn   	:= "."
				this.props.olAttachmentOut	:= "########"
			} else {
				this.props.olFolder             	:= Props.Folder
				this.props.olSubjectIn         	:= Props.SubjectIn
				this.props.olSubjectOut       	:= Props.SubjectOut
				this.props.olSubjectData    	:= Props.SubjectData
				this.props.olAttachmentIn   	:= Props.AttachmentIn
				this.props.olAttachmentOut	:= Props.AttachmentOut
			}

		}

		Connect()                                                              	{

			try this.ol := ComObjActive("Outlook.Application")
			catch
				try
				this.ol := ComObjCreate("Outlook.Application")
					catch
						return (this.olStatus := "no connection")

		return (this.olStatus := "okay")
		}

		Disconnect()                                                            	{

		  ; Outlook Event-Objekt stoppen und danach entfernen
			If this.MailEvents {
				ComObjConnect(this.ol)
				this.ol.Events := false
			}

		  ; Outlook COM-Verbindung unterbrechen durch entfernen des Objekts
			If IsObject(this.ol)
				this.ol := ""

		}

		CatchMailEvents()                                                   	{

			; The storage path for the attachments must be known
				If (!this.AttachmentSavePath || !FilePathExist(this.AttachmentSavePath)) {
					msg := "Function cannot be executed properly.`nCause of error: "
					throw A_ThisFunc ":`nProblem: " msg
												. (!this.AttachmentSavePath ? " Empty storage path"
												: " this file path << " this.AttachmentSavePath " >>  does not exist")
				}

			ComObjConnect(this.ol, new this.email_events(this))
			this.MailEvents := true

		}

		class email_events {

			__New(parent:="") {
					this.parent := parent
			}

			NewMailEx(args*) {

				EntryIDCollection := args.1
				try
					mail := this.parent.ol.GetNamespace("MAPI").GetItemFromID(EntryIDCollection)
					catch
						throw A_ThisFunc ": this is no mailitem object"

				attachments := mail.attachments

				SciTEOutput(A_ThisFunc ": " args.Count() " | [" mail.subject ", " mail.sendername ", " attachments.Count() "]") ;

				If (RegExMatch(mail.subject, "i)WG.*Test") && mail.sendername = "arzt@praxis-clemenz.de")
					mail.delete()
				else if attachments
					this.parent.MailFilter(mail)

			}

			MailEx(args*) {
				SciTEOutput(A_ThisFunc ": " args.Count())
			}

		}

		MailFilter(mail) {

			; es bietet sich an nicht nur die Fritzbox FaxMails zu bearbeiten, sondern alle Mails zu untersuchen
			; es sollen bei ausser bei den Fritzbox-Mails
			;	- die Nachrichten auf bestimmte Stichworte, bekannte Namen oder Absendeadressen untersucht werden
			;	- bei einem Treffer können verschiedene Aktionen erfolgen:
			;		- Rezeptbestellungen - Versand per Telegram an die Anmeldung
			;		- Patientendokumente im Anhang - isolieren, untersuchen und wenn ungefährlich extrahieren und in den Befundordner kopieren
			;		- Patient welche eine EMail gesendet hat ins Albis Wartezimmer aufnehmen, so daß man sieht wer geschrieben hat
			;			- Mailtext könnte automatisch übernommen werden


		}

		ExtractMailAttachments(mail) {

			; 1. mail item zunächst filtern anhand Betreff, Absender (bei Faxeingang über die Fritzbox ist der Absender eine eigene Mailadresse)

		}

		ExtractAttachments(folderName, SJct_In, SJct_Out, AMFilterIn, AMFilterOut, SaveTo) {

				EMailsWithAttachment := unread := AttachmentCounter := 0
				IF StrLen(SJct_In) = 0
					SJct_In := "i)Fax\s"
				IF StrLen(AMFilterIn) = 0
					AMFilterIn := "i)von\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)"

				if !IsObject(this.ol)
					this.Connect()

			olNameSpace := this.ol.GetNameSpace("MAPI")
			olFolder := olNameSpace.Folders(1).Folders(folderName)
			Loop % olFolder.items.count 	{

				email := olFolder.items.Item(A_Index)
				If email.unread
					unread ++

				if email.unread && RegExMatch(email.Subject, SJct_In) && !RegExMatch(email.Subject, SJct_Out) 	{

					; EMail als gelesen markieren
						email.unread := 0
						EMailsWithAttachment ++

					; Posteingang protokollieren
						If EMailsWithAttachment = 1
							FileAppend, % this.datestamp() , % this.DBPath "\outlook.txt", UTF-8

					; FaxAbsender und Faxnummer
						RegExMatch(email.Subject, "i)Fax\svon\s*(?<Absender>[\w\s]+)*\s\(*(?<Nummer>\d\d+)", Fax)
						MBoxMsg .= "Absender:`t"  	FaxAbsender "`n"
						MBoxMsg .= "Nummer: `t" 	FaxNummer "`n"

					; Posteingang protokollieren 2
						FileAppend, % "`t[" FaxAbsender " | " FaxNummer "]", % this.DBPath "\outlook.txt", UTF-8

					; Anhänge entpacken
						Attachments := email.Attachments
						Loop % Attachments.Count		{
							thisattachment := Attachments.Item(A_Index)
							AttachmentCounter ++
							if RegExMatch(thisattachment.DisplayName, AMFilterIn ) 	{
								Fullpath := PathToSaveTo "\" thisattachment.DisplayName
								thisattachment.SaveAsFile(FullPath)
							}

					}
				}
			}

			; Posteingang protokollieren 3
				If EMailsWithAttachment > 0
					FileAppend, % "`n" , % this.DBPath "\outlook.txt", UTF-8

				MsgBox, % "           " A_DDDD ", " A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min "`n"   ;{
							.	 "--------------------------------------------------`n"
							. 	 (AttachmentCounter > 0 ? 	unread " ungelesene EMail" (EMailsWithAttachment > 1 ? "s" : "") "`n" : "")
							. 	 (AttachmentCounter > 0 ? 	EMailsWithAttachment " EMail" (EMailsWithAttachment > 1 ? "s" : "") " hatten Anhänge`n" : "")
							.	 (AttachmentCounter > 0 ? 	AttachmentCounter " " (AttachmentCounter > 1 ? "Anhänge wurden" : "Anhang wurde") " gespeichert`n" : "")
							.	 (AttachmentCounter = 0 ?	"              Heute gab es kein Fax!`n" : "")
							.	 "--------------------------------------------------"  (AttachmentCounter = 0 ?	"" : "`n")
							.	 (AttachmentCounter > 0 ? 	"                  Liste der Absender               `n" : "")
							.	 (AttachmentCounter > 0 ? 	"--------------------------------------------------`n" : "")
							.	 (AttachmentCounter > 0 ? 	Trim(MBoxMsg, "`n") "`n" : "")
							.	 (AttachmentCounter > 0 ? 	"--------------------------------------------------`n" : "")
							.	 (AttachmentCounter > 0 ? "                        " PathToSaveTo "`n" : "")
							.	 (AttachmentCounter > 0 ? 	"--------------------------------------------------" : "")
				;}

		}


		datestamp(nr:=1) {                                                                                                	;-- für Protokolle
			If (nr = 1)
				return (A_DD "." A_MM "." A_YYYY ", " A_Hour ":" A_Min ":" A_Min)
			else
				return (A_Hour ":" A_Min ":" A_Sec)
		}

}

ExtractAttachmentOutlook(folderName, EmailSubject, EmailSubjectData, attachmentname, PathToSaveTo) {

	EMailsWithAttachment := unread := AttachmentCounter := 0

	try ol := ComObjActive("Outlook.Application")
	catch
		try
     		this.ol := ComObjCreate("Outlook.Application")
				catch
					throw Exception("Kann nicht mit Outlook verbinden.")

	olNameSpace := ol.GetNameSpace("MAPI")
	olFolder := olNameSpace.Folders(1).Folders("Posteingang")
	MaxMails := olFolder.items.count
	Loop % olFolder.items.count 	{

		email := olFolder.items.Item(A_Index)
		If email.unread
			unread ++

		if email.unread && RegExMatch(email.Subject, EmailSubject)	{

			email.unread := 0
			Attachments := email.Attachments
			EMailsWithAttachment ++

			RegExMatch(email.Subject, EmailSubjectData, Fax)
			t .= "Absender:`t"  	FaxAbsender "`n"
			t .= "Nummer: `t" 	FaxNummer "`n"

			Loop % Attachments.Count		{
				thisattachment := Attachments.Item(A_Index)
				AttachmentCounter ++
				if RegExMatch(thisattachment.DisplayName , attachmentname ) 	{
					Fullpath := PathToSaveTo "\" thisattachment.DisplayName
					thisattachment.SaveAsFile(FullPath)
				}

			}
		}
	}

	MsgBox, % "           " A_DDDD ", " A_DD "." A_MM "." A_YYYY " " A_Hour ":" A_Min "`n"
				.	 "--------------------------------------------------`n"
				.	 MaxMails " Mails enthält der Ordner`n"
				. 	 (EMailsWithAttachment > 0 ? 	unread " ungelesene EMail" (EMailsWithAttachment > 1 ? "s" : "") "`n" : "")
				. 	 (EMailsWithAttachment > 0 ? 	EMailsWithAttachment " EMail" (EMailsWithAttachment > 1 ? "s" : "") " hatten Anhänge`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	AttachmentCounter " " (EMailsWithAttachment > 1 ? "Anhänge wurden" : "Anhang wurde") " gespeichert`n" : "")
				.	 (EMailsWithAttachment = 0 ?	"Keine passenden Mails mit Anhang gefunden!`n" : "")
				.	 "--------------------------------------------------"  (EMailsWithAttachment = 0 ?	"" : "`n")
				.	 (EMailsWithAttachment > 0 ? 	"                  Liste der Absender               `n" : "")
				.	 (EMailsWithAttachment > 0 ? 	"--------------------------------------------------`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	Trim(t, "`n") "`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	"--------------------------------------------------`n" : "")
				.	 (AttachmentCounter > 0 ? "                        " PathToSaveTo "`n" : "")
				.	 (EMailsWithAttachment > 0 ? 	"--------------------------------------------------" : "")


}

;{ Hilfsfunktionen
IniReadExt(SectionOrFullFilePath, Key:="", DefaultValue:="", convert:=true) {                                          	;-- eigene IniRead funktion für Addendum

	; beim ersten Aufruf der Funktion !nur! Übergabe des ini Pfades mit dem Parameter SectionOrFullFilePath
	; die Funktion behandelt einen geschriebenen Wert der "ja" oder "nein" ist, als Wahrheitswert, also true oder false
	; UTF-16 in UTF-8 Zeichen-Konvertierung
	; Der Pfad in Addendum.Dir wird einer anderen Variable übergeben. Brauche dann nicht immer ein globales Addendum-Objekt
	; letzte Änderung: 31.01.2021

		static admDir
		static WorkIni

	; Arbeitsini Datei wird erkannt wenn der übergebene Parameter einen Pfad ist
		If RegExMatch(SectionOrFullFilePath, "^[A-Z]\:.*\\")	{
			If !FileExist(SectionOrFullFilePath)	{
				MsgBox,, % "Addendum für AlbisOnWindows", % "Die .ini Datei existiert nicht!`n`n" WorkIni "`n`nDas Skript wird jetzt beendet.", 10
				ExitApp
			}
			WorkIni := SectionOrFullFilePath
			If RegExMatch(WorkIni, "[A-Z]\:.*?AlbisOnWindows", rxDir)
				admDir := rxDir
			else
				admDir := Addendum.Dir
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
			If (StrLen(DefaultValue) > 0) { ; Defaultwert vorhanden, dann diesen Schreiben und Zurückgeben
				OutPutVar := DefaultValue
				IniWrite, % DefaultValue, % WorkIni, % SectionOrFullFilePath, % key
				If ErrorLevel
					MsgBox, % "Der Defaultwert <" DefaultValue "> konnte geschrieben werden.`n`n[" WorkIni "]"
			}
			else return "ERROR"
		else if InStr(OutPutVar, "%AddendumDir%")
				return StrReplace(OutPutVar, "%AddendumDir%", admDir)
		else if RegExMatch(OutPutVar, "i)^\s*(ja|nein)\s*$", bool)
				return (bool1= "ja") ? true : false

return Trim(OutPutVar)
}

IniAppend(value, filename, section, key) {                                                                                 	;-- vorhandenen Werten weitere Werte hinzufügen

	IniRead, schreib, % filename, % section, % key
	If Instr(schreib, "Error")
		schreib:=""
	IniWrite, % schreib . value, % filename, % section, % key	;, UTF-8

}

StrUtf8BytesToText(vUtf8) {                                                                                                       	;-- Umwandeln von Text aus .ini Dateien in UTF-8
	if A_IsUnicode 	{
		VarSetCapacity(vUtf8X, StrPut(vUtf8, "CP0"))
		StrPut(vUtf8, &vUtf8X, "CP0")
		return StrGet(&vUtf8X, "UTF-8")
	} else
		return StrGet(&vUtf8, "UTF-8")
}

FilePathCreate(path) {                                                                                                              	;-- erstellt einen Dateipfad falls dieser noch nicht existiert

	If !FilePathExist(path) {
		FileCreateDir, % path
		If ErrorLevel
			return 0
		else
			return 1
	}

return 1
}

FilePathExist(path) {                                                                                                                 	;-- prüft ob ein Dateipfad vorhanden ist
	If !InStr(FileExist(path "\"), "D")
		return 0
return 1
}

isFullFilePath(path) {                                                                                                               	;-- prüft Pfadstring auf die Angabe eines Laufwerkes
	If RegExMatch(path, "[A-Z]\:\\")
		return 1
return 0
}

SelectFolder(FSFOptions,FSFText="",hWndOwner="0") {
   ; Common Item Dialog -> msdn.microsoft.com/en-us/library/bb776913%28v=vs.85%29.aspx
   ; IFileDialog        -> msdn.microsoft.com/en-us/library/bb775966%28v=vs.85%29.aspx
   ; IShellItem         -> msdn.microsoft.com/en-us/library/bb761140%28v=vs.85%29.aspx
   Static OsVersion 	:= DllCall("GetVersion", "UChar")
   Static Show         	:= A_PtrSize * 3
   Static SetOptions 	:= A_PtrSize * 9
   Static SetTitle      	:= A_PtrSize * 17
   Static GetResult  	:= A_PtrSize * 20
   SelectedFolder := ""
   If (OsVersion < 6) { ; IFileDialog requires Win Vista+
      FileSelectFolder, SelectedFolder,, %FSFOptions%, %FSFText%
      Return SelectedFolder
   }
   If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
      Return ""
   VTBL := NumGet(FileDialog + 0, "UPtr")
   DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", 0x00000028, "UInt") ; FOS_NOCHANGEDIR | FOS_PICKFOLDERS
   If (FSFText != "")
      DllCall(NumGet(VTBL + SetTitle, "UPtr"), "Ptr", FileDialog, "Str", FSFText, "UInt")
   If !DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", hWndOwner, "UInt") {
      If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt") {
         GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
         If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr) ; SIGDN_DESKTOPABSOLUTEPARSING
            SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
         ObjRelease(ShellItem)
      }
   }
   ObjRelease(FileDialog)
   Return SelectedFolder
}

SciTEOutput(Text:="", Clear=false, LineBreak=true, Exit=false) {     	; modified version for Addendum für Albis on Windows

	; last change 17.08.2020

	; some variables
		static	LinesOut           	:= 0
			, 	SCI_GETLENGTH	:= 2006
			,	SCI_GOTOPOS		:= 2025

	; gets Scite COM object
		try
			SciObj := ComObjActive("SciTE4AHK.Application")           	;get pointer to active SciTE window
		catch
			return                                                                            	;if not return

	; move Caret to end of output pane to prevent inserting text at random positions
		SendMessage, 2006,,, Scintilla2, % "ahk_id " SciObj.SciteHandle
		endPos := ErrorLevel
		SendMessage, 2025, % endPos,, Scintilla2, % "ahk_id " SciObj.SciteHandle

	; shows count of printed lines in case output pane was erased
		If InStr(Text, "ShowLinesOut") {
			;SciObj.Output("SciteOutput function has printed " LinesOut " lines.`n")
			return
		}

	; Clear output window
		If (Clear=1) || ((StrLen(Text) = 0) && (LinesOut = 0))
			SendMessage, SciObj.Message(0x111, 420)

	; send text to SciTE output pane
		If (StrLen(Text) != 0) {
			Text .= (LineBreak ? "`r`n": "")
			SciObj.Output(Text)
			LinesOut += StrSplit(Text, "`n", "`r").MaxIndex()
		}

		If Exit {
			MsgBox, 36, Exit App?, Exit Application?                         	;If Exit=1 ask if want to exit application
			IfMsgBox,Yes, ExitApp                                                       	;If Msgbox=yes then Exit the appliciation
		}

}
;}




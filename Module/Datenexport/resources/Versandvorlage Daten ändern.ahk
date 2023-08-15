
CreateShippingLetter(templatePath, LetterSavePath, ContentData) {

	;~ doc.validateOnParse := false
	;~ doc.resolveExternals := false


	; XML-Datei laden
	doc := ComObjCreate("Msxml2.DOMDocument.6.0")
	doc.async := false
	doc.loadXML(xmlstr := FileOpen(templatePath, "r", "UTF-8").Read())

	; Überprüfen, ob das Laden der XML erfolgreich war
	if (doc.parseError.errorCode != 0) {
			MsgBox, % "Fehler beim Laden der XML-Datei:`n" doc.parseError.reason
			ExitApp
	}

	; Wert ändern

	anrname := "Dr. med. V. Lach-Meer"
	strasseNr := "An Glockenklingel 12 B"
	PLZORT := "11112 Im Felde"

	svgBase := "//*[name()='svg']"
	svgPathValues := { "*[name()='g' and @id='Unterzeichnerbereich']//*[name()='text' and @id='text_NameUnterzeichner']"    	:	ContentData.Sendername
								, "*[name()='g' and @id='EMailfeld']//*[name()='text' and @id='Text_EMailadresse']"	                       	:	ContentData.SenderMail
								, "*[name()='text' and @id='OrtDatum']//*[name()='tspan' and @id='tspan606']"	                            	:	ContentData.SenderCity ", den " A_DD "." A_MM "." A_YYYY
								, "*[name()='g' and @id='Empfängerfeld']//*[name()='text' and @id='text_AnrTitelName']"	                   	:	anrname
								, "*[name()='g' and @id='Empfängerfeld']//*[name()='text' and @id='text_StrasseNr']"	                    	:	strasseNr
								, "*[name()='g' and @id='Empfängerfeld']//*[name()='text' and @id='text_PLZOrt']"	                        	:	PLZORT}

	For xpath, newvalue in svgPathValues {

		node := doc.selectSingleNode(xp := svgBase . "//" . xpath)
		oldvalue := node.text

		If IsObject(node)
			node.text := newvalue
		else
			oldvalue := "noNode"

		t .= xp "`n  " oldvalue " ==> " newValue "`n"

	}

	FileAppend, % t "`n", *
	recipientName := SQLiteGetRecipient(ContentData.bsnr)
	recipienName 	:= RegExReplace(RecipientName, "^.*?\..*?\.\s*")
	doc.Save((newsvg := A_ScriptDir "\Versandanschreiben_" RecipientName "_" A_DD "." A_MM "." A_YYYY  ".svg"))
	Run, % "msedge.exe " """" newsvg """"

}


	ExitApp
	svgElements := doc.selectNodes("//*[name()='svg']//*")  ; and @id='svg-connector"
	FileAppend, % svgElements.Length "`n", *
	Loop % svgElements.Length {
		svgElement := svgElements.Item[A_Index - 1]
		t .= svgElement.nodename "[" svgElement.getAttribute("id")  "]  `n   " RegExReplace(svgElement.text, "\s{2,}", " ") "`n`n"
	}
		FileAppend, % "parseError: " doc.parseError.errorCode "`n", *

	ExitApp
	if (node) {
			node.text := newValue
		FileAppend, % node.x "`n", *
	} else {
			FileAppend, % "Das Element wurde nicht gefunden. " StrLen(xmlstr) "`nparserError: " doc.parseError.reason "`n", *
			ExitApp
	}

	; Geänderte XML speichern
	;~ node := doc.selectSingleNode("//svg/g[@id='Unterzeichnerbereich']/text[@id='text_NameUnterzeichner']")

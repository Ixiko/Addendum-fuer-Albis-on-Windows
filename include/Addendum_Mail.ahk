
path :=""
If !FileExist(path) {
	MsgBox, keen Pfad
	ExitApp
}

OutlookSendMail("alf@gmx.de", "für die Einreise", body,, path)



Outlook_SendMail(to, subject, body:="", cc:="", Attachments:="") {

		ol := ComObjActive("Outlook.Application").CreateItem(0)
		ol.Subject := Subject
		ol.body := body
		If IsObject(Attachments)
			For nr, filepath in Attachments
				ol.Attachments.Add(filePath)
			else
				ol.Attachments.Add(Attachments)
		ol.display

}

/*
		html := "<html><head><style>"
		html .= "body, p {text-align: 'left'; background: 'none'; font-family: 'calibri'; font-size: 16pt; overflow-y:auto;}"
		html .= "table {border: 1px solid black}"
		html .= "tr {background-color: #f5f5f5;}"
		html .= "th {font-family: 'calibri'; font-size: 13pt; text-align: 'center';border-bottom: 1px solid blue}"
		html .= "td {font-family: 'calibri'; font-size: 11pt; text-align: 'right';border-bottom: 1px solid #ddd}"
		html .= "</style></head><body><p>Dear John:<br><br>See the table below:</p><table>`n"
		loop 5			{
			html .= "<tr>" , rn := a_index
			loop 18
				html .= rn=1 ? "<th width=75> col " a_index "</th>`n" : "<td>" rn * (a_index) - 1 "</td>`n"
			html .= "</tr>`n"
			}
		html .= "</table></html>"

		Ol.BodyFormat := 2 												; HTML
		;Ol.To := "xxx@gmail.com"
		Ol.HTMLBody := html
		;ol.Mail.Recipients.Add(recepient)
 */



;#Persistent
#NoEnv
;SetBatchLines, -1

;URL Quellverzeichnis + Beispielname
;http://www.kbv.de/tools/ebm/html/01100_2901333561637848669504.html

EBMURL:= "http://www.kbv.de/tools/ebm/html/"
OffPath:= "ebm\html"
;URLDownloadToFile, %EBMURL%\01100*.html, D:\01100.html
;If errorlevel	
;	MSgBox, Das hat wohl nicht funktioniert.	

/*
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("GET", "http://www.kbv.de/tools/ebm/html/01100*.html", true)
whr.Send()
; Durch das 'true' oben und dem Aufruf unten bleibt das Skript ansprechbar.
whr.WaitForResponse()
version := whr.ResponseText
MsgBox % version
*/



Exit

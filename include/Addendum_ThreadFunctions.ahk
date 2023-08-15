
;~ #Persistent
global Addendum

;~ Addendum := CriticalObject((&obj))
Addendum := ObjShare("lresult")
Addendum.Print("DBPath")
MsgBox % "DBPath: " Addendum.DBPath

ClassifyFiles()

ClassifyFiles(showProgress:=true)                     	{

	classifier := new autonamer(Addendum.DBPath "\Dictionary"
											, {"pdfpath":Addendum.BefundOrdner, "pdfdatapath":Addendum.DBPath "\sonstiges\PdfDaten.json"}
											,  "Debug=false ShowCleanupedText=false RemoveStopWords=true Save_immediately=false")

	If IsObject(documents := classifier.readDirectory("only_unnamed=true")) {

	  ; einfache Progress Gui
		If showprogress {
			Progress	, % "B2 zH25 w600 WM400 WS500"
						.  	 " 	cW" 	Addendum.Default.BgColor1
						. 	 " 	cB" 	Addendum.Default.PRGColor
						. 	 " 	cT" 	Addendum.Default.FntColor
						, % "Klassifiziere Dokument: "
						, % "Addendum für AlbisOnWindows"
						, % "Dokumentklassifizierung"
						, % Addendum.Default.Font
			hprg := WinExist("Dokumentklassifizierung ahk_class AutoHotkey2")
		}

	  ; Dateiliste klassifizieren
		factor := 100/documents.Count()
		For fIndex, file in documents {

			If showprogress {
				rc := Round(fIndex * factor)
				ControlSetText, Static2, % "Klassifiziere Dokument [" SubStr("000" fIndex, -1*(StrLen(documents.Count())-1)) "/" documents.Count() "]: " file.name
									, % "ahk_id " hprg
				Progress % rc
			}

			If (!file.isSearchable || file.category || !FileExist(file.path "\" file.name))
				continue

			doctxt	:= classifier.getDocumentText(file.path "\" file.name)
			result 	:= classifier.matchscore(doctxt, file.name)
			val    	:= result.Delete("txt")

		  ; die besten 3 Ergebnisse werden herausgesucht
			bestof 	:= [Array(), Array(), Array()]
			For score, data in result.titles {

				if (score ~= "(" (result.max1 ? result.max1 : "#") "|" (result.max2 ? result.max2 : "#") "|" (result.max3 ? result.max3 : "#")  ")") {
						obj := {"category":data.maintitle, "content":(data.subTitle="--" ? "" : data.subTitle)}
							If (score = result.max1)
							bestof.1.Push(obj)
						else If (score = result.max2)
							bestof.2.Push(obj)
						else If (score = result.max3)
							bestof.3.Push(obj)
				}

			}

		  ; Entfernen leerer Einträge
			Loop 3
			 If (bestof[4-A_Index].Count() = 0)
				bestof.RemoveAt(4-A_Index)

		  ; bestes Ergebnis sichern
			file.category 	:= bestof.1.1.category
			file.content 	:= bestof.1.1.content
			file.bestof    	:= bestof.titles

		}

	  ; Backup der PDF-Daten.json Datei
		If !InStr(FileExist(Addendum.DBPath "\sonstiges\_Backup"), "D")
			FileCreateDir, % Addendum.DBPath "\sonstiges\_Backup"
		Loop 6	{	; maximal 6 Backups

			If !FileExist(fpath := Addendum.DBPath "\sonstiges\_Backup\_backup-PdfDaten" A_Index ".json") {
				fpath := Addendum.DBPath "\sonstiges\_Backup\_PdfDaten"  (A_Index > 1 ? A_Index-1 : 1) ".json"
				break
			}

			FileGetTime, ftime, % fpath , M
			If (!ftimeLast || ftimeLast > ftime) {
				ftimeLast := ftime
				fpath := Addendum.DBPath "\sonstiges\_Backup\_backup-PdfDaten" A_Index ".json"
			}

		}

		FileCopy, % Addendum.DBPath "\sonstiges\PdfDaten.json", % fpath, 1
		FileOpen(Addendum.DBPath "\sonstiges\PdfDaten.json", "w", "UTF-8").Write(cJSON.Dump(documents, 1))
	}

return documents
}

#include %A_ScriptDir%\Addendum_Autonaming.ahk
#include %A_ScriptDir%Addendum_PdfHelper.ahk
#include %A_ScriptDir%Addendum_Internal.ahk
#include %A_ScriptDir%\..\..\..\..\lib\class_cJSON.ahk
#include %A_ScriptDir%\..\..\..\..\lib\SciTEOutput.ahk

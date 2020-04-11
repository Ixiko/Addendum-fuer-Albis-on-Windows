Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit


' Makro 'NachrichtlichAn' zur Verwendung mit Albis on Windows'
' Copyright Albis Arzteservice Product GmbH 1996-2008
'
Dim p__$(), s__$(), Anz, d$

Public Sub MAIN()
ReDim p__$(10, 21)
ReDim s__$(21)
Dim DateinameAlt$
Dim i
Dim j
Dim w
Anz = 0
d$ = ""
    '                      !!!   Pfad eventuell ändern   !!!
    d$ = "c:\albiswin\listen\tmpfrm.txt"

    ' den Zustand vor der Ersetzung retten
    ' Pfad mit Speichern, sonst kommt Datei auf anderen Pfad, AOM 11.2.1998
    ' DateinameAlt$ = WindowName$(0)
    DateinameAlt$ = WordBasic.[FileNameFromWindow$]()
    
    WordBasic.FileSaveAs Name:="$TEMP$.DOC", Format:=0
    
    SuchtexteInintialisieren
    Anz = ÜArztDateiLesen
    
    ' für alle Briefe (zu benachrichtigende Ärzte = Anz Zeilen der Übergabedatei) alle Suchbegriffe finden und ersetzen
        For i = 1 To Anz        ' i zählt die Briefe
            ' alten Zustand laden
            WordBasic.FileOpen Name:="$TEMP$.DOC", Revert:=1
            
            ' alle Platzhalter durch übergebene Datendelder (Spalten) ersetzen (ohne Nachrichtlich An)
            For j = 1 To 20
                SuchenAusschneidenErsetzen i, j
            Next j
            ' jetzt noch NachrichtlichAn
            NachrichtlichAn (i)
            
            ' Abfrage ob drucken
            ' w = WordBasic.MsgBox("Soll der Brief gedruckt werden ?", 4) ' j/n-Antwort
            Dialogs(wdDialogFilePrint).Show
            'If w = -1 Then
               ' WordBasic.FilePrint AppendPrFile:=0, Range:="0", _
                             PrToFileName:="", From:="", _
                             To:="", Type:=0, NumCopies:="1", _
                             Pages:="", Order:=0, _
                             PrintToFile:=0, Collate:=1, _
                             FileName:=""
               ' Dialogs(wdDialogFilePrint).Show
                
             'End If
            
        Next i
        ' letzte Ersetzung als Datei speichern
        WordBasic.FileSaveAs Name:=DateinameAlt$, Format:=0
        ' WordBasic.DocClose
End Sub

Private Sub SuchenAusschneidenErsetzen(pos, cnt)
    Dim flagge As Boolean
    flagge = False
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect
        flagge = True
    End If
    
    WordBasic.StartOfDocument
    On Error Resume Next
    WordBasic.EditFind Find:=s__$(cnt), _
                     Direction:=0, _
                     WholeWord:=0, _
                     MatchCase:=0, _
                     PatternMatch:=0, _
                     Format:=0, _
                     Wrap:=0
    On Error GoTo -1: On Error GoTo 0
    If WordBasic.EditFindFound() Then
        ' alles auf einem Schlag ersetzen
        WordBasic.EditReplace Find:=s__$(cnt), _
                           Replace:=p__$(pos, cnt), _
                           Direction:=0, _
                           MatchCase:=0, _
                           WholeWord:=0, _
                           PatternMatch:=0, _
                           ReplaceAll:=1, _
                           Format:=0, _
                           Wrap:=0
    End If
    If ActiveDocument.ProtectionType = wdNoProtection And flagge = True Then
        ActiveDocument.Protect (wdAllowOnlyComments)
    End If

End Sub

Private Sub NachrichtlichAn(pos)
Dim Nan$
Dim nl$
Dim n
Dim nPosNachrichtlich

    ' Position im Array für Nachrichtlich An
    nPosNachrichtlich = 21

    ' Text zusammenstellen
    Nan$ = ""
    nl$ = ""
    For n = 1 To Anz
        If n <> pos Then
            If Nan$ <> "" Then
                nl$ = "^z"
            End If
            Nan$ = Nan$ + nl$ + "^t" + p__$(n, nPosNachrichtlich)
        End If
    Next n

    WordBasic.StartOfDocument
    On Error Resume Next
    WordBasic.EditFind Find:=s__$(21), _
                     Direction:=0, _
                     WholeWord:=0, _
                     MatchCase:=0, _
                     PatternMatch:=0, _
                     Format:=0, _
                     Wrap:=0
    On Error GoTo -1: On Error GoTo 0
    If WordBasic.EditFindFound() Then
        ' alles auf einem Schlag ersetzen
        WordBasic.EditReplace Find:=s__$(nPosNachrichtlich), _
                           Replace:=Nan$, _
                           Direction:=0, _
                           MatchCase:=0, _
                           WholeWord:=0, _
                           PatternMatch:=0, _
                           ReplaceAll:=1, _
                           Format:=0, _
                           Wrap:=0
        Application.ScreenRefresh
    End If
End Sub

Private Function ÜArztDateiLesen()
Dim i
Dim j
Dim nPosNachrichtlich
Dim p_
    ' Ü-Arzt-Datei öffnen (s.o.)
    Open d$ For Input As 1
    ' Position im Array für Nachrichtlich An
    nPosNachrichtlich = 21
    ' alle Datensätze (zu benachrichtigende Ärzte) einlesen und in Tabelle aufnehemen
    ' Anzahl gelesener Datensätze (zu benachrichtigende Ärzte) zurückgeben
    i = 0
    While Not EOF(1)
        i = i + 1
        If i <= 10 Then
            Input #1, p__$(i, 1), p__$(i, 2), p__$(i, 3), p__$(i, 4), p__$(i, 5), _
                p__$(i, 6), p__$(i, 7), p__$(i, 8), p__$(i, 9), p__$(i, 10), _
                p__$(i, 11), p__$(i, 12), p__$(i, 13), p__$(i, 14), p__$(i, 15), _
                p__$(i, 16), p__$(i, 17), p__$(i, 18), p__$(i, 19)
            For j = 0 To 19
                p__$(i, j) = WordBasic.[RTrim$](WordBasic.[LTrim$](p__$(i, j)))
                p_ = InStr(p__$(i, j), ";")
                While p_ > 0
                    If p_ < Len(p__$(i, j)) Then
                        p__$(i, j) = Mid(p__$(i, j), 1, p_ - 1) + "," + _
                                                Mid(p__$(i, j), p_ + 1)
                    Else
                        p__$(i, j) = Mid(p__$(i, j), 1, p_ - 1) + ","
                    End If
                    p_ = InStr(p__$(i, j), ";")
                Wend
            Next j
            
            ' Uebarztstraße (ebenfalls aus Spalte 7)
            p__$(i, 20) = p__$(i, 7)
            
            ' nachrichtlich bauen
            p__$(i, nPosNachrichtlich) = p__$(i, 5) ' Name
            If p__$(i, 6) <> "" Then
                p__$(i, nPosNachrichtlich) = p__$(i, 6) + " " + p__$(i, nPosNachrichtlich) ' Vorname
            End If
            If p__$(i, 4) <> "" Then
                p__$(i, nPosNachrichtlich) = p__$(i, 4) + " " + p__$(i, nPosNachrichtlich) ' Zusatz
            End If
            If p__$(i, 3) <> "" Then
                p__$(i, nPosNachrichtlich) = p__$(i, 3) + " " + p__$(i, nPosNachrichtlich) ' Titel
            End If
            If p__$(i, 8) <> "" And p__$(i, 8) <> "0" Then
                p__$(i, nPosNachrichtlich) = p__$(i, nPosNachrichtlich) + " " + p__$(i, 8) ' Plz
            End If
            If p__$(i, 9) <> "" Then
                p__$(i, nPosNachrichtlich) = p__$(i, nPosNachrichtlich) + " " + p__$(i, 9) ' Ort
            End If
            If p__$(i, 10) <> "" Then
                p__$(i, nPosNachrichtlich) = p__$(i, nPosNachrichtlich) + " (" + p__$(i, 10) + ")" ' Fach
            End If
        Else
            WordBasic.MsgBox "Mehr als 10 Ü-Arzt-Adressen.^zMakro nicht ausführbar!"
            i = 10
        End If
    Wend
    Close 1
    ÜArztDateiLesen = i
End Function

Private Sub SuchtexteInintialisieren()
    s__$(1) = "$Übarztnr#"
    s__$(2) = "$Übarztanrede#"
    s__$(3) = "$Übarzttitel#"
    s__$(4) = "$Übarztzusatz#"
    s__$(5) = "$Übarztname#"
    s__$(6) = "$Übarztvorname#"
    s__$(7) = "$Übarztstrasse#"
    s__$(8) = "$Übarztplz#"
    s__$(9) = "$Übarztort#"
    s__$(10) = "$Übarztfach#"
    s__$(11) = "$Übarzttel#"
    s__$(12) = "$Übarzttel2#"
    s__$(13) = "$Übarztfax#"
    s__$(14) = "$Übarztaltanrede#"
    s__$(15) = "$Übarztinfo#"
    s__$(16) = "$Übarztsprechzeit#"
    s__$(17) = "$ÜbarztAltAnschrift#"
    s__$(18) = "$ÜbarztPraxbez#"
    s__$(19) = "$ÜbarztLanr#"
    ' ab hier befinden sich zusätzliche Platzhalter für die keine Daten von ALBIS on Windows übergeben werden
    s__$(20) = "$Übarztstraße#"         ' Alternative zu Übarztstrasse
                                           ' => falls Position geändert wird anpassen in Sub ÜArztDateiLesen => siehe dort Kommentar ' straße
    s__$(21) = "$nachrichtlich#"        ' hier werden alle Empfänger/Übärzte eingesetzt
                                           ' => falls Position geändert wird anpassen in Sub ÜArztDateiLesen
                                           ' und Sub NachrichtlichAn => jeweils siehe dort Kommentar ' Position im Array für Nachrichtlich An
End Sub





Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
' Makro 'ALBISMenuPunkte' zur Verwendung mit Albis on Windows'
' Copyright Albis Arzteservice Product GmbH 1998 - 2008
'

Option Explicit ' Variablen sind explizit zu deklarieren

Sub ALBISMenüAnfügen()
'
' Funktion ALBISMenüAnfügen
' fügt einen neuen Menüpunkt mit Namen ALBIS am Ende der Menüleiste an
' 24.09.98, AOM
'
    ' evtl. hier Variablen deklarieren, wenn Explicit gesetzt
    Dim ActiveMenuBar As CommandBar
    Dim ALBISMenu As CommandBarControl
    Dim ALBISSubMenuMakros As CommandBarControl
    Dim ALBISSubMenuWechselZuAoW As CommandBarControl
    Dim ALBISSubMenuEinfuegen As CommandBarControl
    Dim ALBISSubMenuPatientPersonalien As CommandBarControl
    Dim ALBISSubMenuPatientVersicherung As CommandBarControl
    Dim ALBISSubMenuPatientDaten As CommandBarControl
    Dim ALBISSubMenuUebarzt As CommandBarControl
    Dim ALBISSubMenu As CommandBarControl
    Dim ALBISSubMenuEinnahme As CommandBarControl
    Dim ALBISSubMenuArztPraxis As CommandBarControl
    Dim ALBISSubMenuMahnPraxGeb As CommandBarControl
    
    ' Dies ist kein Untermenü!:
    Dim ALBISSubMenuEintrag As CommandBarButton
    
    Dim bGefunden As Boolean
    
    Dim nItem As Integer
    
    
    'Menüpunkt in der Hauptmenüleiste anlegen
    Set ActiveMenuBar = CommandBars.ActiveMenuBar
    
    bGefunden = False
    
    For nItem = 1 To ActiveMenuBar.Controls.Count
        If ActiveMenuBar.Controls(nItem).Caption = "ALB&IS" Then
            Set ALBISMenu = ActiveMenuBar.Controls(nItem)
            
            Dim cmp As CommandBarPopup
            Set cmp = ActiveMenuBar.Controls(nItem)
            ' Menüpunkte des Menüs "einfügen" holen:
            Set cmp = cmp.Controls(2)
            If (cmp.Controls(2).Caption = "Laufendes Quartal") Then
                bGefunden = True
            End If
            
            If Not bGefunden Then
                ALBISMenu.Delete
            End If
            Exit For
        End If
    Next nItem
    
    If Not bGefunden Then
        Set ALBISMenu = ActiveMenuBar.Controls.Add(Type:=msoControlPopup)
        ALBISMenu.Caption = "ALB&IS"
        ALBISMenu.TooltipText = "Spezielle Funktionen zur Textverarbeitung mit Albis on Windows"
        
        ' Menüpunkt Wechsel zu AoW
        Set ALBISSubMenuWechselZuAoW = ALBISMenu.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuWechselZuAoW.Caption = "&Wechseln zu AoW"
        ALBISSubMenuWechselZuAoW.TooltipText = "Wechselt zu Albis on Windows"
        ALBISSubMenuWechselZuAoW.Style = msoButtonAutomatic
        ALBISSubMenuWechselZuAoW.OnAction = "WechselZuAoW"
        
        ' Untermenü <Einfügen>
        Set ALBISSubMenuEinfuegen = ALBISMenu.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuEinfuegen.Caption = "&Einfügen"
        ALBISSubMenuEinfuegen.TooltipText = "Platzhalter in Textvorlage einfügen"
        
        ' Druckdatum
        Set ALBISSubMenuEintrag = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Druckdatum"
        ALBISSubMenuEintrag.Parameter = "$Druckdatum#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Laufendes Quartal
        Set ALBISSubMenuEintrag = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Laufendes Quartal"
        ALBISSubMenuEintrag.Parameter = "$LfdQuartal#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Untermenü <Personalien Patient>
        Set ALBISSubMenuPatientPersonalien = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuPatientPersonalien.Caption = "&Personalien Patient"
        ALBISSubMenuPatientPersonalien.TooltipText = "Platzhalter für Personalien des Patienten in Textvorlage einfügen"
        ' Untermenü <Versicherungsdaten Patient>
        Set ALBISSubMenuPatientVersicherung = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuPatientVersicherung.Caption = "&Versicherung Patient"
        ALBISSubMenuPatientVersicherung.TooltipText = "Platzhalter für Versicherungsdaten des Patienten in Textvorlage einfügen"
        ' Untermenü <Daten Patient>
        Set ALBISSubMenuPatientDaten = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuPatientDaten.Caption = "&Daten Patient"
        ALBISSubMenuPatientDaten.TooltipText = "Platzhalter für weitere Daten des Patienten in Textvorlage einfügen"
        
        ' Untermenü <Arzt/Praxisdaten>
        Set ALBISSubMenuArztPraxis = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuArztPraxis.Caption = "&Arzt- und Praxisdaten"
        ALBISSubMenuArztPraxis.TooltipText = "Platzhalter für Daten des Überweisungsarztes in Textvorlage einfügen"
        
        ' Untermenü <Überweisungsarzt>
        Set ALBISSubMenuUebarzt = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuUebarzt.Caption = "&Überweisungsarzt"
        ALBISSubMenuUebarzt.TooltipText = "Platzhalter für Daten des Überweisungsarztes in Textvorlage einfügen"
        
        ' Untermenü <Einnahmeverordnung>
        Set ALBISSubMenuEinnahme = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuEinnahme.Caption = "Einnahme&verordnung"
        ALBISSubMenuEinnahme.TooltipText = "Platzhalter für Daten der Einnahmeverordnung in Textvorlage einfügen"
        
        ' Untermenü <Mahnung Praxisgebühr>
        Set ALBISSubMenuMahnPraxGeb = ALBISSubMenuEinfuegen.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuMahnPraxGeb.Caption = "&Mahnung Praxisgebühr"
        ALBISSubMenuMahnPraxGeb.TooltipText = "Platzhalter für Mahnung Praxisgebühr in Textvorlage einfügen"
        
        ' Menüpunkte einfügen
        ' PatNr
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Nummer"
        ALBISSubMenuEintrag.Parameter = "$Patnr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Patient(in)
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Patient(in)"
        ALBISSubMenuEintrag.Parameter = "$in#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' geehrte(r)
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Geehrte(r)"
        ALBISSubMenuEintrag.Parameter = "$r#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anrede
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Anrede"
        ALBISSubMenuEintrag.Parameter = "$Anrede#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anrede2
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anrede&2"
        ALBISSubMenuEintrag.Parameter = "$Anrede2#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Titel
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Vorname#")
        ALBISSubMenuEintrag.Caption = "&Titel"
        ALBISSubMenuEintrag.Parameter = "$Titel#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Zusatz
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Z&usatz"
        ALBISSubMenuEintrag.Parameter = "$Zusatz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Vorsatz Wort
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$VorsWort#")
        ALBISSubMenuEintrag.Caption = "Vors. &Wort"
        ALBISSubMenuEintrag.Parameter = "$VorsWort#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Nachname
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Nachname"
        ALBISSubMenuEintrag.Parameter = "$Nachname#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Vorname
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Vorname"
        ALBISSubMenuEintrag.Parameter = "$Vorname#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Straße
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Straße"
        ALBISSubMenuEintrag.Parameter = "$Straße#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Zusatz Straßenadresse
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$StraßeZusatz#")
        ALBISSubMenuEintrag.Caption = "Zusatz St&raßenadresse"
        ALBISSubMenuEintrag.Parameter = "$StraßeZusatz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Land
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Land"
        ALBISSubMenuEintrag.Parameter = "$Land#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' PLZ
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Pl&z"
        ALBISSubMenuEintrag.Parameter = "$Plz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Ort
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Ort"
        ALBISSubMenuEintrag.Parameter = "$Ort#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Postfach
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Postfach#")
        ALBISSubMenuEintrag.Caption = "&Postfach"
        ALBISSubMenuEintrag.Parameter = "$Postfach#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Postfachadresse Land
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$PostfachLand#")
        ALBISSubMenuEintrag.Caption = "Postfachadresse &Land"
        ALBISSubMenuEintrag.Parameter = "$PostfachLand#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Postfachadresse PLZ
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$PostfachPLZ#")
        ALBISSubMenuEintrag.Caption = "Post&fachadresse PLZ"
        ALBISSubMenuEintrag.Parameter = "$PostfachPLZ#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Postfachadresse Ort
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$PostfachOrt#")
        ALBISSubMenuEintrag.Caption = "Postfacha&dresse Ort"
        ALBISSubMenuEintrag.Parameter = "$PostfachOrt#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Geburtsdatum
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Geburtsdatum"
        ALBISSubMenuEintrag.Parameter = "$Gebdatum#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Alter
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Alter"
        ALBISSubMenuEintrag.Parameter = "$Alter#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Telefon
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Telefon-Nr."
        ALBISSubMenuEintrag.Parameter = "$Telefon#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' 2. Telefon
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "2. Telefon-Nr."
        ALBISSubMenuEintrag.Parameter = "$Telefon2#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' FAX
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Fa&x-Nr."
        ALBISSubMenuEintrag.Parameter = "$FAX#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' EMail
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&E-Mail-Addresse"
        ALBISSubMenuEintrag.Parameter = "$EMAIL#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Arbeitgeber
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Arbeitgeber"
        ALBISSubMenuEintrag.Parameter = "$Arbeitgeber#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Patient seit
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Patient seit"
        ALBISSubMenuEintrag.Parameter = "$Patseit#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' BG-Adresse
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "BG-Adresse"
        ALBISSubMenuEintrag.Parameter = "$BGAd#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' BG-Adresse (lang)
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "BG-Adresse (lang)"
        ALBISSubMenuEintrag.Parameter = "$BGAdLang#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Bankleitzahl
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Bankleitzahl / BIC"
        ALBISSubMenuEintrag.Parameter = "$BLZ#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
      
        ' Name Kreditinstitut
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Name Kredit&institut"
        ALBISSubMenuEintrag.Parameter = "$Kreditinstitut#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
       
        ' KontoNr.
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "&Kontonummer / IBAN"
        ALBISSubMenuEintrag.Parameter = "$Kontonr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Kontoinhaber Name
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Kontoinhaber Name"
        ALBISSubMenuEintrag.Parameter = "$KontoinhaberName#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Kontoinhaber Adresse
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientPersonalien.Controls.Add(Type:=msoControlButton, Parameter:="$Nachname#")
        ALBISSubMenuEintrag.Caption = "Kontoinhaber Adresse"
        ALBISSubMenuEintrag.Parameter = "$KontoinhaberAdresse#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
      
        
        ' Kasse
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Kasse"
        ALBISSubMenuEintrag.Parameter = "$Kasse#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Privatkasse
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Privatkasse"
        ALBISSubMenuEintrag.Parameter = "$Privatkasse#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Vknr
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&VKNR"
        ALBISSubMenuEintrag.Parameter = "$Vknr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Ikz
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&IK"
        ALBISSubMenuEintrag.Parameter = "$Ikz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Versicherten-Nr.
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Versicherten-&Nr."
        ALBISSubMenuEintrag.Parameter = "$Versnr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Versicherterstatus
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Versicherter&status"
        ALBISSubMenuEintrag.Parameter = "$VersStatus#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Versicherterstatusergänzung
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Status&ergänzung"
        ALBISSubMenuEintrag.Parameter = "$VersStatusErg#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Gültigkeit der KVK
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Gültigkeit der KVK"
        ALBISSubMenuEintrag.Parameter = "$VersGuelt#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Gebührenfrei
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Gebühr &frei bis"
        ALBISSubMenuEintrag.Parameter = "$Freibis#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-Anrede
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Anrede"
        ALBISSubMenuEintrag.Parameter = "$HvAnrede#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-Nachname
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Nachname"
        ALBISSubMenuEintrag.Parameter = "$HvNachname#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-Vorname
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Vorname"
        ALBISSubMenuEintrag.Parameter = "$HvVorname#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-Straße
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Straße"
        ALBISSubMenuEintrag.Parameter = "$HvStraße#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-PLZ
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Plz"
        ALBISSubMenuEintrag.Parameter = "$HvPlz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-Ort
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Ort"
        ALBISSubMenuEintrag.Parameter = "$HvOrt#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' HV-Geburt
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientVersicherung.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "HV-&Gebursdatum"
        ALBISSubMenuEintrag.Parameter = "$HvGebdatum#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anmerkung
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anmer&kung"
        ALBISSubMenuEintrag.Parameter = "$Anmerkung#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anmerkung Zeile 1
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anmerk. Zeile &1"
        ALBISSubMenuEintrag.Parameter = "$WIZ1#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anmerkung Zeile 2
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anmerk. Zeile &2"
        ALBISSubMenuEintrag.Parameter = "$WIZ2#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anmerkung Zeile 3
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anmerk. Zeile &3"
        ALBISSubMenuEintrag.Parameter = "$WIZ3#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anmerkung Zeile 4
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anmerk. Zeile &4"
        ALBISSubMenuEintrag.Parameter = "$WIZ4#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anmerkung Zeile 5
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Anmerk. Zeile &5"
        ALBISSubMenuEintrag.Parameter = "$WIZ5#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' AU
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&AU bis"
        ALBISSubMenuEintrag.Parameter = "$AU#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Allergie
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "A&llergie"
        ALBISSubMenuEintrag.Parameter = "$Allergie#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Anamnese
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Ana&mnese"
        ALBISSubMenuEintrag.Parameter = "$Anamnese#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
                
        ' BMI
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "BMI"
        ALBISSubMenuEintrag.Parameter = "$BMI#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Cave
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Cave"
        ALBISSubMenuEintrag.Parameter = "$Cave#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Dauerdiagnosen
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Dauer&diagnosen"
        ALBISSubMenuEintrag.Parameter = "$Dauerdiag#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Dauermedikamente
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Dauer&medikamente"
        ALBISSubMenuEintrag.Parameter = "$Dauermedi#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Größe
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Größe"
        ALBISSubMenuEintrag.Parameter = "$Größe#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
                
        ' Gewicht
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Gewicht"
        ALBISSubMenuEintrag.Parameter = "$Gewicht#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Kontrolltermine
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Kont&rolltermine"
        ALBISSubMenuEintrag.Parameter = "$Kontrolltermine#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Laborwerte
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Laborwerte"
        ALBISSubMenuEintrag.Parameter = "$Laborwerte#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Letztes Behandlungsdatum RBN, 01.10.2002
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Letztes &Behandlungsdatum"
        ALBISSubMenuEintrag.Parameter = "$BehDat#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Manueller Hinweis
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Manueller Hin&weis"
        ALBISSubMenuEintrag.Parameter = "$Hinweis#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
       
        ' Mutterschutz
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Mutterschutz"
        ALBISSubMenuEintrag.Parameter = "$Mutterschutz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
       
        ' Operation
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Operation"
        ALBISSubMenuEintrag.Parameter = "$Operation#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Rechnungsbetrag Praxisgebühr
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Rechnungsb&etrag (nur Einzugsermächtigung)"
        ALBISSubMenuEintrag.Parameter = "$Rechnungsbetrag#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Röntgennummer
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Röntgennummer"
        ALBISSubMenuEintrag.Parameter = "$Röntgennummer#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Schwangerschaftswoche
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Schwangerschaftswoche"
        ALBISSubMenuEintrag.Parameter = "$SSW#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Therapie
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Therapie"
        ALBISSubMenuEintrag.Parameter = "$Therapie#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Termin (nur Terminzettel F6)
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Termine (nur Termin&zettel!)"
        ALBISSubMenuEintrag.Parameter = "$Termine;z+#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Unfall
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Unfall"
        ALBISSubMenuEintrag.Parameter = "$Unfall#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Verwendungszweck
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Verwendungszweck (nur Einzugsermächtigung)"
        ALBISSubMenuEintrag.Parameter = "$Verwendungszweck#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Voraussichtlicher Entbindungstermin
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Voraussichtlicher Entbindungstermin"
        ALBISSubMenuEintrag.Parameter = "$VET#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"

        ' AusnahmeIndikation
        Set ALBISSubMenuEintrag = ALBISSubMenuPatientDaten.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Ausnahme Indikation"
        ALBISSubMenuEintrag.Parameter = "$AusnahmeIndikation#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        

        ' Überweisungsarzt Praxisbezeichnung
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Praxisbezeichnung"
        ALBISSubMenuEintrag.Parameter = "$ÜbarztPraxbez#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt KV-Nummer/BSNR
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Arzt-Nr."
        ALBISSubMenuEintrag.Parameter = "$Übarztnr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt LANR
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "LANR"
        ALBISSubMenuEintrag.Parameter = "$ÜbarztLanr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Anrede
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Anrede"
        ALBISSubMenuEintrag.Parameter = "$Übarztanrede#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Titel
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Titel"
        ALBISSubMenuEintrag.Parameter = "$Übarzttitel#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Namenszusatz
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Z&usatz"
        ALBISSubMenuEintrag.Parameter = "$Übarztzusatz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Nachname
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Nachname"
        ALBISSubMenuEintrag.Parameter = "$Übarztname#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Vorname
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Vorname"
        ALBISSubMenuEintrag.Parameter = "$Übarztvorname#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Straße
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Straße"
        ALBISSubMenuEintrag.Parameter = "$Übarztstraße#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt PLZ
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Pl&z"
        ALBISSubMenuEintrag.Parameter = "$Übarztplz#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Ort
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Ort"
        ALBISSubMenuEintrag.Parameter = "$Übarztort#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Fachrichtung
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Fachrichtung"
        ALBISSubMenuEintrag.Parameter = "$Übarztfach#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Telefon
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Telefon-Nr."
        ALBISSubMenuEintrag.Parameter = "$Übarzttel#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt 2. Telefon-Nr.
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "2. Telefon-Nr."
        ALBISSubMenuEintrag.Parameter = "$Übarzttel2#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt FAX
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Fa&x-Nr."
        ALBISSubMenuEintrag.Parameter = "$Übarztfax#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt E-Mail
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&E-Mail-Addresse"
        ALBISSubMenuEintrag.Parameter = "$ÜbarztEmail#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt alternative Anrede
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "alt. Anrede"
        ALBISSubMenuEintrag.Parameter = "$Übarztaltanrede#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Info
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Info"
        ALBISSubMenuEintrag.Parameter = "$Übarztinfo#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt Sprechzeit
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Spre&chzeit"
        ALBISSubMenuEintrag.Parameter = "$Übarztsprechzeit#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt alternative Anschrift
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "alt. Anschrift"
        ALBISSubMenuEintrag.Parameter = "$ÜbarztAltAnschrift#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        ' Überweisungsarzt nachrichtlich
        Set ALBISSubMenuEintrag = ALBISSubMenuUebarzt.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Nachrichtlich"
        ALBISSubMenuEintrag.Parameter = "$nachrichtlich#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
                        
        ' Tag der Verordnung:
        Set ALBISSubMenuEintrag = ALBISSubMenuEinnahme.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Tag der Verordnung"
        ALBISSubMenuEintrag.Parameter = "$RezDatum[]#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Einnahmeverordnung Medikament
        Set ALBISSubMenuEintrag = ALBISSubMenuEinnahme.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&Packungsbezeichnung"
        ALBISSubMenuEintrag.Parameter = "$RezMed[]#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Einnahmeverordnung morgens
        Set ALBISSubMenuEintrag = ALBISSubMenuEinnahme.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&morgens"
        ALBISSubMenuEintrag.Parameter = "$morgens[]#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Einnahmeverordnung mittags
        Set ALBISSubMenuEintrag = ALBISSubMenuEinnahme.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "m&ittags"
        ALBISSubMenuEintrag.Parameter = "$mittags[]#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Einnahmeverordnung abends
        Set ALBISSubMenuEintrag = ALBISSubMenuEinnahme.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&abends"
        ALBISSubMenuEintrag.Parameter = "$abends[]#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Einnahmeverordnung nachts
        Set ALBISSubMenuEintrag = ALBISSubMenuEinnahme.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&nachts"
        ALBISSubMenuEintrag.Parameter = "$nachts[]#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Mahnung Praxisgebühr
        Set ALBISSubMenuEintrag = ALBISSubMenuMahnPraxGeb.Controls.Add(Type:=msoControlButton)
        With ALBISSubMenuEintrag
            .Caption = "Mahnbetrag"
            .Parameter = "%$PraxgebMahnBetrag#"
            .Style = msoButtonCaption
            .OnAction = "ALBISPlatzhalterEinfuegen"
        End With
        Set ALBISSubMenuEintrag = ALBISSubMenuMahnPraxGeb.Controls.Add(Type:=msoControlButton)
        With ALBISSubMenuEintrag
            .Caption = "Mahn-Quartal"
            .Parameter = "$PraxgebMahnQuartal#"
            .Style = msoButtonCaption
            .OnAction = "ALBISPlatzhalterEinfuegen"
        End With
        Set ALBISSubMenuEintrag = ALBISSubMenuMahnPraxGeb.Controls.Add(Type:=msoControlButton)
        With ALBISSubMenuEintrag
            .Caption = "Mahngebühr/Porto"
            .Parameter = "$PraxgebMahnBetragExtra#"
            .Style = msoButtonCaption
            .OnAction = "ALBISPlatzhalterEinfuegen"
        End With
        Set ALBISSubMenuEintrag = ALBISSubMenuMahnPraxGeb.Controls.Add(Type:=msoControlButton)
        With ALBISSubMenuEintrag
            .Caption = "Mahn-Frist (zahlbar bis)"
            .Parameter = "$PraxgebMahnFrist#"
            .Style = msoButtonCaption
            .OnAction = "ALBISPlatzhalterEinfuegen"
        End With
        
        ' Arzt/Praxisdaten:
        
        ' Arztname
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Arzt-&Name"
        ALBISSubMenuEintrag.Parameter = "$ArztName#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Arztfachrichtung
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Arzt-&Fachrichtung"
        ALBISSubMenuEintrag.Parameter = "$ArztFachrichtung#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Arzt-Mail
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Arzt E-&Mail-Adresse"
        ALBISSubMenuEintrag.Parameter = "$ArztEmail#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' ArztKvNrForm
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "&KV-Nr. für Formular-/Briefbedruckung"
        ALBISSubMenuEintrag.Parameter = "$ArztKvNrForm#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' ArztKvNrAbr
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "KV-Nr. für &Abrechnung (nur in besonderen Fällen!)"
        ALBISSubMenuEintrag.Parameter = "$ArztKvNrAbr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' ArztKnappschaftsnr
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Knapps&chaftsnummer"
        ALBISSubMenuEintrag.Parameter = "$ArztKnappschaftsnr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' LANR
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "LANR"
        ALBISSubMenuEintrag.Parameter = "$Lanr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Praxisstraße
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Praxis-&Straße"
        ALBISSubMenuEintrag.Parameter = "$PraxisStraße#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Praxis-PLZ
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Praxis-&PLZ"
        ALBISSubMenuEintrag.Parameter = "$PraxisPLZ#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Praxis-Ort
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Praxis-&Ort"
        ALBISSubMenuEintrag.Parameter = "$PraxisOrt#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Praxis-Telefon
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Praxis-Telefon"
        ALBISSubMenuEintrag.Parameter = "$PraxisTelefon#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Praxis-Fax
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Praxis-Telefa&x"
        ALBISSubMenuEintrag.Parameter = "$PraxisFax#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Praxis-Mail
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Praxis &E-Mail-Adresse"
        ALBISSubMenuEintrag.Parameter = "$PraxisEmail#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' Stempel
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "Stempe&l"
        ALBISSubMenuEintrag.Parameter = "$Stempel#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
        
        ' BSNR
        Set ALBISSubMenuEintrag = ALBISSubMenuArztPraxis.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "BSNR"
        ALBISSubMenuEintrag.Parameter = "$Bsnr#"
        ALBISSubMenuEintrag.Style = msoButtonCaption 'nur Text
        ALBISSubMenuEintrag.OnAction = "ALBISPlatzhalterEinfuegen"
                        
        ' Untermenü Makros
        Set ALBISSubMenuMakros = ALBISMenu.Controls.Add(Type:=msoControlPopup)
        ALBISSubMenuMakros.Caption = "&Makros"
        ALBISSubMenuMakros.TooltipText = "Liste aller Makros"
                
        ' dessen Menüpunkte einfügen (NachrichtlichAn)
        Set ALBISSubMenuEintrag = ALBISSubMenuMakros.Controls.Add(Type:=msoControlButton)
        ALBISSubMenuEintrag.Caption = "NachrichtlichAn"
        ALBISSubMenuEintrag.TooltipText = "Startet das Makro NachrichtlichAn"
        ALBISSubMenuEintrag.DescriptionText = "Startet das Makro NachrichtlichAn"
        ALBISSubMenuEintrag.Style = msoButtonAutomatic
        ALBISSubMenuEintrag.OnAction = "NachrichtlichAn.MAIN"
        
        
        
        
    End If
    
    
End Sub


Sub ALBISPlatzhalterEinfuegen()
' Fügt den entsprechenden Platzhalter des Menüpunktes an der aktuell selektierten Stelle im aktiven Dokument ein
' 25.09.98, AOM
    Dim ActiveDoc As Document
    Dim ActiveWindow As Window
    
    If Documents.Count Then
        Set ActiveDoc = Application.ActiveDocument
        If ActiveDoc.Windows.Count Then
            Set ActiveWindow = ActiveDoc.ActiveWindow
            If ActiveWindow.WindowState <> wdWindowStateMinimize Then

                If ActiveWindow.Selection.Active Then
                    ' die aktuelle Selection ist aktiv
                    With CommandBars.ActionControl
                        ' Parameter des aufrufenden Menüpunktes einsetzen
                        Selection.TypeText Text:=.Parameter
                    End With
                 End If
            End If
        End If
    End If
End Sub

Sub AutoExec()
    ALBISMenüAnfügen
End Sub


Sub WechselZuAoW()
'
' Wechselt zurück zu Albis on Windows. Copyright Albis Ärzteservice Product GmbH & Co KG 1999
'
    On Error Resume Next
    If Tasks.Exists("Albis on Windows") Then
        Tasks("Albis on Windows").Activate
    End If
End Sub




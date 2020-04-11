; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                                                  	  ADDENDUM AUTO-ICD-DIAGNOSEN                                                                                      	. . . . . . . . . .
; . . . . . . . . . .                                                                                              letzte Änderung 01.09.2019                                                                                           	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                  - effektive Eingabe von Diagnosen in Albis on Windows durch Autohotkey Hotstrings                                                     	. . . . . . . . . .
; . . . . . . . . . .                                                  - Hotstrings sind so kurz gehalten das diese noch eindeutig sind                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                  - keine Verwendung kryptischer Kürzel                                                                                                                         	. . . . . . . . . .
; . . . . . . . . . .                                                  - sofortige Diagnoseexpandierung nach Erkennung des richtigen Wortbeginnes                                                            	. . . . . . . . . .
; . . . . . . . . . .                                                  - Abfrage der Seitenlokalisation bei entsprechenden Diagnosen                                                                                    	. . . . . . . . . .
; . . . . . . . . . .                                                  - Mehrfachauswahl von Diagnosen möglich                                                                                                                 	. . . . . . . . . .
; . . . . . . . . . .                                                  - Diagnosenketten z.B. bei Diabetes mit Folgekomplikationen mit einem Klick                                                               	. . . . . . . . . .
; . . . . . . . . . .                                                  - Thesaurus für Diagnosen (nicht vollständig)                                                                                                               	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                                  Achtung: dieses Skript benötigt Addendum.ahk oder Addendum_Functions.ahk                                                             	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . .                                        ROBOTIC PROCESS AUTOMATION FOR THE GERMAN MEDICAL SOFTWARE "ALBIS ON WINDOWS"                                    	. . . . . . . . . .
; . . . . . . . . . .                                               BY IXIKO STARTED IN SEPTEMBER 2017 - THIS FILE RUNS UNDER LEXIKO'S GNU LICENCE                                             	. . . . . . . . . .
; . . . . . . . . . .                                          	   RUNS WITH AUTOHOTKEY_H AND AUTOHOTKEY_L IN 32 OR 64 BIT UNICODE VERSION                                             	. . . . . . . . . .
; . . . . . . . . . .                                                                     THIS SOFTWARE IS ONLY AVAIBLE IN GERMAN LANGUAGE                                                                    	. . . . . . . . . .
; . . . . . . . . . .                                                                                                                                                                                                                                   	. . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
; . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
;RegEx AutoDiagnose einsetzen : (\:R\*\:)(\w*)(\:\:)(.*)(\};) -> :XR*:\2::AutoDiagnose("\4}", "")
^#!s::
HotstringStatitistik(AddendumDir "\Module\Addendum\Include\AutoDiagnosen.ahk")
return

#If ( (WinActive("ahk_class OptoAppClass") && InStr(AlbisGetActiveControl("contraction"), "dia"))  ) && NeueDiagnose(ADInterception)   ;{          || WinActive("Überweisung")
;	 A
:XR*:ACE::                                                                                                                                                                        	;{ ACE Hemmer Husten
AutoDiagnose("ACE Hemmer Husten {T88.7G}")
return ;}
:XR*:Adip::                                                                                                                                                                        	;{ Adipositas alle Grade
AutoDiagnose("Adipositas Grad I über 18 Jahre {E66.00G}|Adipositas Grad II über 18 Jahre {E66.01G}|Adipositas Grad III über 18 Jahre {E66.02G}")
return ;}
:XR*:Adrenogen::                                                                                                                                                               	;{ Adrenogenitale Störung
AutoDiagnose("Adrenogenitale Störung, nicht näher bezeichnet {E25.9G}", "")
return	;}
:XR*:Alkohol::                                                                                                                                                                    	;{ Alkoholkrankheit
:XR*:C2::
AutoDiagnose("C2-Abusus {F10.2G}", "")
return ;}
:XR*:Alzhei::                                                                                                                                                                      	;{ Alzheimer-Krankheit
AutoDiagnose("Alzheimer-Krankheit (F00.9*G) {+G30.9G}; Demenz bei Alzheimer-Krankheit (G30.9+G) {*F00.9G}")
return ;}
:XR*:Analfist::                                                                                                                                                                    	;{ Analfistel
AutoDiagnose("Analfistel {K60.3G}")
return ;}
:XR*:Analfiss::                                                                                                                                                                    	;{ Akute Analfissur
AutoDiagnose("Akute Analfissur {K60.0G}", "")
return ;}
:XR*:AnginaP::                                                                                                                                                                   	;{ Angina pectoris
AutoDiagnose("Angina pectoris, stabile {I20.9G}", "")
return ;}
:XR*:App::                                                                                                                                                                         	;{ Akute Appendizitis
AutoDiagnose("Akute Appendizitis, nicht näher bezeichnet {K35.8G}", "")
return ;}
:XR*:Angsts::                                                                                                                                                                     	;{ Generalisierte Angststörung
AutoDiagnose("Generalisierte Angststörung {F41.1G}", "")
return ;}
:XR*:Aortena::                                                                                                                                                                   	;{ Aortenaneurysma
AutoDiagnose("Aneurysma der Aorta abdominalis ohne Angabe einer Ruptur {I71.4G}|Aneurysma der Aorta thoracica ohne Angabe einer Ruptur {I71.2G}|Aortenaneurysma, thorakoabdominal, ohne Angabe einer Ruptur {I71.6G}", "M")
return ;}
:XR*:Ather::                                                                                                                                                                       	;{ Atherom
:XR*:Funkel::
:XR*:Trichile::
AutoDiagnose("Atherom {L72.1G}", "S")
return ;}
:XR*:Autismussp::                                                                                                                                                              	;{ Autismusspektrumstörung, kombinierte
:XR*:Asper::                                                                                                  ; Asperger Syndrom
AutoDiagnose("kombinierte Autismusspektrumstörung {F84.5G}", "")
return	;}
:XR*:Arthri::                                                                                                                                                                       	;{ Arthritis, reaktiv
:XR*:Gelenkent::
AutoDiagnose("Reaktive Arthritis, nicht näher bezeichnet: Finger Hand {M02.94G}", "S")
return
;}
:XRC*:AB::                                                                                                                                                                        	;{ Asthma bronchiale
:XR*:Asth::
AutoDiagnose("Asthma bronchiale {J45.0G}", "")
return ;}
:XRC*:AV::                                                                                                                                                                        	;{ AV_Block div.
:XR*:Wencke::
:XR*:Mobitz::
:XR*:Herzbl::
:XR*:Hemibl::
:XR*:Rechtssch::
:XR*:Linkssch::
AutoDiagnose("AV-Block I° {I44.0G}|AV-Block II° {I44.1G}|AV-Block III° {I44.2G}|Atrioventrikulärer Block {I44.3G}|Linksanteriorer Hemiblock {I44.4G}|Linksposteriorer Hemiblock {I44.5G}|Inkompletter Linksschenkelblock {I44.6G}|Kompletter Linksschenkelblock {I44.7G}|Rechtsfaszikulärer Block {I45.0}|Kompletter Rechtsschenkelblock {I45.1G}|inkompletter Rechtsschenkelblock {I45.1G}|Wolff-Parkinson-White-Syndrom {I45.6G}", "")
return	;}
;	 B
:XR*:Baker::                                                                                                                                                                      	;{ Baker-Zyste
:XR*:Synovialz::
AutoDiagnose("Baker-Zyste {M71.2G}", "S")
return ;}
:XR*:Bandsch::                                                                                                                                                                  	;{ Bandscheibenvorfall
AutoDiagnose("Lumbale und sonstige Bandscheibenschäden mit Radikulopathie (G55.1*) {+M51.1G}; Kompression von Nervenwurzeln und Nervenplexus bei Bandscheibenschäden (M50-M51+) {*G55.1G}", "S")
return ;}
:XR*:Basali::                                                                                                                                                                      	;{ Basaliom
AutoDiagnose("Basaliom des Gesichtes {C44.3G}|Basaliom der behaarten Kopfhaut {C44.4G}|Basaliom des Augenlides, einschließlich Kanthus {C44.1G}|Basaliom der Lippenhaut {C44.0G}|Basaliom des Ohres und des äußeren Gehörganges {C44.2G}", "MS")
return ;}
:XR*:Bauchs::                                                                                                                                                                    	;{ Bauchschmerz
:XR*:Oberb::
:XR*:Mittelb::
:XR*:Unterb::
AutoDiagnose("Schmerzen im Bereich des Oberbauches {R10.1G}|Schmerzen im Bereich des Mittelbauches {R10.2G}|Schmerzen im Bereich des Unterbauches {R10.3G}", "")
return ;}
:XR*:Beckens::                                                                                                                                                                   	;{ Beckenschiefstand / Hüftgelenksdysplasie
AutoDiagnose("Beckenschiefstand +<#> cm {M95.5G}|Hüftgelenksdysplasie {Q65.8G}", "MS")
return ;}
:XR*:Belastungsr::                                                                                                                                                              	;{ Belastungsreaktion, akute
:XRC*:ABR::                                                                    	; Abkürzung für ...
:XR*:Krisenrea::                                                             	; Krisenreaktion
:XR*:psychischeD::                                                            	; psychische Dekompensation
AutoDiagnose("Akute Belastungsreaktion {F43.0G}", "")
return ;}
:XR*:Block::                                                                                                                                                                       	;{ Blockierung HWS/BWS/LWS
AutoDiagnose("LWS Blockierung {M99.83G}|Blockierung BWS {M99.82G}|Blockierung HWS {M99.81G}", "M")
return ;}
:XRC*:BB::                                                                                                                                                                         	;{ Blockierung BWS
AutoDiagnose("Blockierung BWS {M99.82G}", "")
return ;}
:XRC*:BL::                                                                                                                                                                         	;{ Blockierung, LWS
AutoDiagnose("LWS Blockierung {M99.83G}", "")
return ;}
:XR*:Border::                                                                                                                                                                     	;{ Borderline, Emotional instabile Persönlichkeitsstörung vom Borderline-Typ
:XR*:Emotional::
AutoDiagnose("Emotional instabile Persönlichkeitsstörung vom Borderline-Typ {F60.31G}", "")
return	;}
:XR*:Borrel::                                                                                                                                                                      	;{ Borrelliose / Lyme
:XR*:Lyme::
AutoDiagnose("Lyme-Krankheit {A69.2G}")
return ;}
:XR*:Bronchitis::                                                                                                                                                                 	;{ Bronchitis, Akute
AutoDiagnose("Akute Bronchitis {J20.9G}", "")
return ;}
:XR*:Bronchial::                                                                                                                                                                 	;{ Bronchialkarzinom
:XR*:Lungenk::
AutoDiagnose("metastasiertes Lungenkarzinom {C34.8G}|Metastasen in der Lunge bei <#> {C78.0G}|Bronchial-Ca -Hauptbronchus {C34.0G}|Bronchial-Ca -Lungenoberlappen {C34.1G}", "S")
return ;}
:XR*:Bursitis::                                                                                                                                                                     	;{ Bursitis versch. Lokalisationen
AutoDiagnose("Bursitis trochanterica {M70.6G}|Bursitis im Schulterbereich {M75.5G}|Bursitis olecrani {M70.2G}|Bursitis praepatellaris {M70.4G}", "MS")
return ;}
;	 C
:XR*:Candi::                                                                                                                                                                      	;{ Candida / Soorösophagitis
:XR*:Soor::
AutoDiagnose("Candida-Ösophagitis {B37.81G}")
return ;}
:XR*:Carot::                                                                                                                                                                      	;{ Carotisstenose
AutoDiagnose("Arteria carotis Stenose {I65.2G}", "S")
return ;}
:XR*:Cholezysto::                                                                                                                                                               	;{ Cholezystolithiasis
:XR*:Gallenblasens::
:XR*:Gallens::
AutoDiagnose("Gallenblasenstein {K80.20G}|Gallenblasenstein ohne Cholezystitis ohne Gallenwegsobstruktion {K80.20G}|Gallenblasenstein ohne Cholezystitis mit Gallenwegsobstruktion {K80.21G}")
return ;}
:XR*:Claustr::                                                                                                                                                                    	;{ Claustrophobie
:XR*:Platzan::
AutoDiagnose("Claustrophobie {F40.2G}")
return ;}
:XR*:Colitis::                                                                                                                                                                      	;{ Colitis ulcerosa
AutoDiagnose("Colitis ulcerosa {K51.9G}", "")
return	;}
:XR*:COPD::                                                                                                                                                                     	;{ COPD
:XR*:obstruktive::
AutoDiagnose("COPD {J44.89G}|exacerbierte COPD {J44.19G}|COPD, FEV1 <35% {J44.90G}|COPD, FEV1 >=35% und <50% {J44.91G}|COPD, FEV1 >=50 % und <70 % {J44.92G}|COPD, FEV1 >=70% {J44.93G}")
return
;}
:XR*:Coxaval::                                                                                                                                                                      	;{ Coxa valga antetorta
AutoDiagnose("Coxa valga antetorta {M21.05G}", "S")
Return ;}
:XR*:Coxar::                                                                                                                                                                      	;{ Coxarthrose
:XR*:Hüftgelenksa::
AutoDiagnose("Coxarthrose {M16.0G}", "S")
return ;}
;	  D
:XR*:Daumena::                                                                                                                                                                	;{ Daumengrundgelenk, Rhizarthrose / reaktive Arthritis des Daumengrundgelenkes
:XR*:Rhiza::
AutoDiagnose("Rhizarthrose, nicht näher bezeichnet {M18.9G}|reaktive Arthrititis des Daumengrundgelenkes {M02.84G}", "MS")
return ;}
:XR*::Demenz::                                                                                                                                                                  	;{ Demenz, versch. Formen
AutoDiagnose("Subkortikale vaskuläre Demenz {F01.2G}")
return ;}
:XR*:Depres::                                                                                                                                                                    	;{ Depression
AutoDiagnose("Leichte depressive Episode {F32.0G}|Mittelgradige depressive Episode {F32.1G}|Schwere depressive Episode ohne psychotische Symptome {F32.2G}|Rezidivierende depressive Störung, gegenwärtig leichte Episode {F33.0G}|Rezidivierende depressive Störung, gegenwärtig mittelgradige Episode {F33.1G}|Rezidivierende depressive Störung, gegenwärtig schwere Episode {F33.2G}")
return ;}
:XR*:Diabet::                                                                                                                                                                      	;{ Diabetes mellitus Typ 2
AutoDiagnose("Diabetes mellitus Typ 2 {E11.90G}|Diabetes mit Polyneuropathie:Diabetes mellitus mit multiplen Komplikationen {E14.72G}; Diabetische Polyneuropathie (E10-E14+, vierte Stelle .4) {*G63.2G}; Glomeruläre Krankheiten bei Diabetes mellitus (E10-E14+, vierte Stelle .2) {*N08.3G}|diabetische Gastropathie [G59.0*, G63.2*, G73.0*, G99.0*] {+E11.40G}")
return ;}
:XR*:Diar::                                                                                                                                                                        	;{ Gastroenteritis
:XRC*:GE::
AutoDiagnose("Gastroenteritis {A09.0G}", "")
return ;}
:XR*:Divertikuli::                                                                                                                                                                	;{ Divertikulitis des Dickdarmes
:XR*:Sigmad::
AutoDiagnose("Divertikulitis des Dickdarmes ohne Perforation, Abszess, Blutung {K57.32G}|Divertikulitis des Dickdarmes mit Perforation ohne Abszess ohne Blutung {K57.22G}|Divertikulitis des Dickdarmes mit Perforation, Abszess und Blutung {K57.23G}|Divertikulitis des Dickdarmes ohne Perforation ohne Abszess mit Blutung {K57.33G}; ", "")
return	;}
:XR*:Divertikulo::                                                                                                                                                                	;{ Divertikulose des Dickdames
AutoDiagnose("Divertikulose des Dickdarmes {K57.30G}|Divertikulose des Dickdarmes ohne Perforation oder Abszess, mit Blutung {K57.31G}", "")
return	;}
:XR*:Dyskal::                                                                                                                                                                     	;{ Dyskalkulie
:XR*:Rechens::                                          	;Rechenstörung
AutoDiagnose("Dyskalkulie {R48.8G}", "")
return ;}
;	  E
:XR*:Eisenm::                                                                                                                                                                    	;{ Eisenmangelanämie
:XR*:mikrozytär::                                                             	; mikrozytäre Anämie
AutoDiagnose("Eisenmangelanämie {D50.9G}", "")
return ;}
:XR*:Epist::                                                                                                                                                                        	;{ Epistaxis
:XR*:Nasenb::
AutoDiagnose("Epistaxis {R04.0G}", "S")
return ;}
:XR*:Epico::                                                                                                                                                                       	;{ Epicondylitis ulnaris humeri
:XR*:Tenn::                                                                          	; Tennisarm
AutoDiagnose("Epicondylitis ulnaris humeri {M77.0G}|Epicondylitis radialis humeri {M77.1G}; ", "S")
return	;}
:XR*:Enzeph::                                                                                                                                                                    	;{ Enzephalopathie, Mikrovaskulär
AutoDiagnose("Mikrovaskuläre Enzephalopathie {I67.9G}", "")
return ;}
:XR*:exacer::                                                                                                                                                                     	;{ exacerbierte COPD
AutoDiagnose("exacerbierte COPD {J44.19G}", "")
return ;}
;	  F
:XR*:Fers::                                                                                                                                                                         	;{ Fersensporn
:XR*:Achil::
:XR*:Kalkan::
AutoDiagnose("Fersensporn {M77.3G}", "S")
return ;}
:XR*:Fettl::                                                                                                                                                                         	;{ Fettleber / Steatosis hepatis
:XR*:Steat::
AutoDiagnose("Fettleber {K76.0G}")
return ;}
:XR*:Folgen::                                                                                                                                                                     	;{ Folgen eines ....
AutoDiagnose("Folgen eines Schlaganfall {I69.4G}|Folgen eines intraspinalen Abszesses {G09G}|Folgen einer Myelitis{G09G}|Folgen einer intrazerebralen Blutung {I69.1G}", "M")
return ;}
:XR*:Folsä::                                                                                                                                                                       	;{ Folsäure-Mangelanämie
AutoDiagnose("Arzneimittelinduzierte Folsäure-Mangelanämie {D52.1G}", "")
return ;}
;	  G
:XR*:Gastroe::                                                                                                                                                                   	;{ Gastroenteritis / Durchfallerkrankung
:XR*:Durchf::
AutoDiagnose("Gastroenteritis {A09.0G}")
return ;}
:XR*:Gicht::                                                                                                                                                                       	;{ Gicht
:XR*:Urikop::                                                                                        	; Urikopathie
AutoDiagnose("idiopathische Gicht {M10.07G}", "S")
return ;}
:XR*:Gonar::                                                                                                                                                                     	;{ Gonarthrose / Chrondromalacia patellae / Kniegelenkserguß
:XR*:Kniegelenksa::
AutoDiagnose("Gonarthrose {M17.9G}|Chondromalacia patellae {M22.4G}|Kniegelenkserguß {M25.46G}", "MS")
return ;}
:XR*:Gürtel::                                                                                                                                                                      	;{ Gürtelrose / Herpes zoster
:XR*:HerpesZ::
:XR*:Zoster::
AutoDiagnose("Zoster [Herpes zoster] ohne Komplikation {B02.9G}", "S")
return ;}
;	  H
:XR*:Hämorr::                                                                                                                                                                   	;{ Hämorrhoiden
AutoDiagnose("Hämorrhoiden {K64.9G}", "")
return ;}
:XR*:Hämatoc::                                                                                                                                                                 	;{ Hämatochezie
:XR*:Melä::
:XR*:Teer::	;Teerstuhl
AutoDiagnose("Meläna {K92.1G}")
return ;}
:XR*:Helico::                                                                                                                                                                      	;{ Helicobacter pylori
AutoDiagnose("Helicobacter pylori [H. pylori] als Ursache von Krankheiten {!B98.0G}", "")
return ;}
:XR*:Hemipa::                                                                                                                                                                   	;{ Hemiparese
AutoDiagnose("Spastische Hemiparese {G81.1G}", "S")
return ;}
:XR*:Hepato::                                                                                                                                                                      	;{ Hepatomegalie
:XR*:Leberver::                                                                                                  ; Lebervergrößerung
:XR*:Lerbersch::                                                                                                 ; Lerberschwellung
AutoDiagnose("Hepatomegalie {R16.0G};", "")
return	;}
:XR*:Hepatitis::                                                                                                                                                                  	;{ Akute Virushepatitis
:XR*:Virushep::                                                                                                  ; Virushepatitis
:XR*:Leberent::                                                                                                  ; Leberentzündung
AutoDiagnose("Akute Virushepatitis {B17.9G}", "")
return	;}
:XR*:Hernia::	                                                                                                                                                                    	;{ Hernia inguinalis
:XR*:Hernie::
:XR*:Leisten::
AutoDiagnose("Hernia inguinalis mit Einklemmung und ohne Gangrän, nicht als Rezidivhernie bezeichnet {K40.30G}", "S")
return ;}
:XR*:Herzinf::                                                                                                                                                                     	;{ Herzinfarkt, Myokardinfarkt
:XR*:Myokardinf::
:XR*:Hinterwandi::
:XR*:Vorderwandi::
AutoDiagnose("Herzinfarkt {I21.9G}|Vorderwandinfarkt des Herzens {I21.0G}|Hinterwandinfarkt des Herzens {I21.1G}|Akuter transmuraler Myokardinfarkt an sonstigen Lokalisationen {I21.2G}|Akuter subendokardialer Myokardinfarkt {I21.4G}")
return ;}
:XR*:Herzinsu::                                                                                                                                                                  	;{ Herzinsuffizienz, Links-/Rechtsherzinsuffizienz
:XR*:Myokardins::
:XR*:Linksherzins::
:XR*:Rechtsherzins::
AutoDiagnose("Linksherzinsuffizienz ohne Beschwerden {I50.11G}|Linksherzinsuffizienz NYHA II {I50.12G}|Linksherzinsuffizienz NYHA III {I50.14G}|Linksherzinsuffizienz NYHA IV {I50.14G}|Linksherzinsuffizienz, nicht näher bezeichnet {I50.19G}|Primäre Rechtsherzinsuffizienz {I50.00G}|Sekundäre Rechtsherzinsuffizienz {I50.01G}")
return ;}
:XR*:Herzspi::                                                                                                                                                                    	;{ Herzspitzenaneurysma
AutoDiagnose("Herzspitzenaneurysma {I25.3G}", "")
return ;}
:XR*:Hüftgelenksd::                                                                                                                                                            	;{ Hüftgelenksdysplasie
:XR*:Hüftdysp::                                                                                       	; Hüftdysplasie
:XR*:Hüftgelenksve::                                                                              	; Hüftgelenksverrenkung
AutoDiagnose("Hüftgelenksdysplasie {Q65.8G}", "S")
return ;}
:XR*:Hydroz::                                                                                                                                                                    	;{ Hydrozele
:XR*:Wasserbr::                                                                                	; Wasserbruch
AutoDiagnose("Hydrozele {N43.3G}", "S")
return ;}
:XR*:Hyperch::                                                                                                                                                                   	;{ Hypercholesterinämie
:XR*:Hyperlipoprot::	; Hyperlipoproteinämie
AutoDiagnose("Hypercholesterinaemie {E78.0G}", "")
return ;}
:XR*:Hyperg::                                                                                                                                                                    	;{ Hyperglykämie
:XR*:Überz::
AutoDiagnose("Hyperglykämie {R73.9G}", "")
return ;}
:XR*:Hyperl::                                                                                                                                                                     	;{ Hyperlipidaemie
AutoDiagnose("Hyperlipidaemie {E78.5G}", "")
return ;}
:XR*:Hyperpa::                                                                                                                                                                    	;{ Hyperparathyreoidismus
AutoDiagnose("Primärer Hyperparathyreoidismus {E21.0G}", "")
return ;}
:XR*:Hyperpr::                                                                                                                                                                    	;{ Hyperprolaktinämie
AutoDiagnose("Hyperprolaktinämie {E22.1G}", "")
return	;}
:XR*:Hyperte::                                                                                                                                                                   	;{ hypertensive Episode
AutoDiagnose("hypertensive Episode {I10.91G}", "")
return ;}
:XR*:Hyperth::                                                                                                                                                                   	;{ Hyperthyreose diverse
AutoDiagnose("Hyperthyreose {E05.9G}", "")
return
:XR*:Based::AutoDiagnose("Morbus Basedow {E05.0G}", "")
:XR*:Hashi::AutoDiagnose("Hashimoto {E06.3G}", "")
:XR*:Autoimmunthyre::
:XR*:Thyreoi::
AutoDiagnose("Autoimmunthyreoiditis {E06.3G}", "")
return ;}
:XR*:Hyperto::                                                                                                                                                                   	;{ Hypertonus
:XR*:Blutho::
AutoDiagnose("Hypertonie {I10.0G}|essentielle Hypertonie {I10.0G}")
return ;}
:XR*:Hypogl::                                                                                                                                                                    	;{ Hypoglykämie
:XR*:Unterz::
AutoDiagnose("Hypoglykämie {E16.2G}", "")
return	;}
:XR*:Hypok::                                                                                                                                                                     	;{ Hypokaliämie
:XR*:Kaliumm::	;Kaliummangel
AutoDiagnose("Hypokaliämie {E87.6G}", "")
return ;}
:XR*:Hypoth::                                                                                                                                                                    	;{ Hypothyreose
AutoDiagnose("Hypothyreose {E03.9G}", "")
return ;}
:XR*:laten::                                                                                                                                                                        	;{ latente Hypothyreose
AutoDiagnose("latente Hypothyreose {E03.9G}", "")
return ;}
:XR*:Hypoto::                                                                                                                                                                    	;{ Hypotonie - orthostatisch / idiopathisch / durch Arzneimittel
:XR*:Orthos::
AutoDiagnose("Orthostatische Hypotonie {I95.1G}|Hypotonie durch Arzneimittel {I95.2G}|Idiopathische Hypotonie {I95.0G}|Sonstige Hypotonie {I95.8G}")
return ;}
:XR*:Harnw::                                                                                                                                                                      	;{ Harnwegsinfekt
:XRC*:HTI::
:XR*:Zystit::
:XR*:Blasenentz::
AutoDiagnose("Harnwegsinfektion {N39.0G}")
return ;}
;	  I
:XR*:Impet::                                                                                                                                                                       	;{ Impetigo contagiosa
:XR*:Grind::                                                                               	; Grindflechte, Grindblasen, Eitergrind
AutoDiagnose("Impetigo contagiosa Kopf, <#> {L01.0G}", "S")
return ;}
:XR*:Imping::                                                                                                                                                                    	;{ Impingment-Syndrom Schulter, Rotatorenmanschettenruptur
:XR*:Rotator::
AutoDiagnose("Impingement-Syndrom der Schulter {M75.4G}; Bursitis im Schulterbereich {M75.5G}|Läsionen der Rotatorenmanschette {M75.1RG}", "MS")
return ;}
:XR*:Insek::                                                                                                                                                                       	;{ Insektenstich
AutoDiagnose("Insektenstich {T63.4G}", "")
return ;}
;	  K
:XR*:KHK::                                                                                                                                                                          	;{ KHK
AutoDiagnose("KHK {I25.19G}", "")
return ;}
:XR*:Kardiom::                                                                                                                                                                    	;{ Kardiomegalie
AutoDiagnose("Kardiomegalie {I51.7G}", "")
return ;}
:XR*:Katar::                                                                                                                                                                        	;{ Katarakt
:XR*:grauer::                                                                                                                   	; grauer Star
AutoDiagnose("Katarakt {H26.9G}", "S")
return ;}
:XR*:Knick::                                                                                                                                                                        	;{ Knick-Senk-Spreizfuss
AutoDiagnose("Knick-Senk-Spreizfuss {M21.4G}", "S")
return ;}
:XR*:Kniegelenkse::                                                                                                                                                            	;{ Kniegelenkserguß
AutoDiagnose("Kniegelenkserguß {M25.46G}", "S")
return ;}
:XR*:Kopf::                                                                                                                                                                        	;{ Kopfschmerz
AutoDiagnose("Kopfschmerz {G44.8G}", "")
return ;}
:XR*:Kontaktd::                                                                                                                                                                  	;{ Kontaktdermatitis
:XR*:Kontaktex::	;Kontaktexzem
AutoDiagnose("Allergische Kontaktdermatitis {L23.9G}", "")
return ;}
:XR*:Krampf::                                                                                                                                                                     	;{ Krampf, Muskelkrämpfe
:XR*:Krämpfe::
:XR*:Muskelkr::
AutoDiagnose("Krämpfe und Spasmen der Muskulatur {R25.2G}", "S")
return ;}
;	  L
:XR*:Laktos::                                                                                                                                                                      	;{ Laktoseintoleranz
AutoDiagnose("Laktoseintoleranz, nicht näher bezeichnet {E73.9G}", "")
return ;}
:XR*:Lager::                                                                                                                                                                       	;{ Lagerungsschwindel, Benigner paroxysmaler Schwindel
AutoDiagnose("Benigner paroxysmaler Schwindel {H81.1G}", "")
return ;}
:XR*:Leberz::                                                                                                                                                                      	;{ Leberzirrhose, dekompensiert
AutoDiagnose("dekompensierte Leberzirrhose {K74.6G}", "")
return ;}
:XR*:Leses::                                                                                                                                                                       	;{ Lese- und Rechtschreibstörung
:XR*:Rechtssc::
AutoDiagnose("Lese- und Rechtschreibstörung {F81.0G}", "")
return ;}
:XR*:Lipom::                                                                                                                                                                       	;{ Lipom
AutoDiagnose("Lipom {D17.3G}", "S")
return ;}
:XR*:Lungenem::                                                                                                                                                                	;{ Lungenembolie ohne Angabe eines akuten Cor pulmonale
AutoDiagnose("Lungenembolie ohne Angabe eines akuten Cor pulmonale {I26.9G}", "S")
return ;}
:XR*:Lungenm::                                                                                                                                                                  	;{ Lungenmetastasen, Sekundäre bösartige Neubildung der Lunge
AutoDiagnose("Sekundäre bösartige Neubildung der Lunge {C78.0G}", "S")
return ;}
:XR*:Luxat::                                                                                                                                                                        	;{ Luxation des Schultergelenkes
AutoDiagnose("Luxation des Schultergelenkes {S43.00G}", "S")
return
:XR*:Schultergelenkslux::AutoDiagnose("Luxation des Schultergelenkes {S43.00G}", "S")
;}
:XRC*:LSB::                                                                                                                                                                        	;{ Linksschenkelblock
AutoDiagnose("Linksschenkelblock {I44.7G}", "")
return ;}
:XR*:Lympha::                                                                                                                                                                    	;{ Lymphadenitis versch. Lokalisationen
AutoDiagnose("Akute Lymphadenitis an Gesicht, Kopf und Hals {L04.0G}|Akute Lymphadenitis am Rumpf {L04.1G}|Akute Lymphadenitis an der oberen Extremität {L04.2G}|Akute Lymphadenitis an der unteren Extremität {L04.3G}", "MS")
return ;}
:XR*:Lymphö::                                                                                                                                                                    	;{ Lymphödem
AutoDiagnose("Lymphödem, nicht näher bezeichnet {I89.09G}", "S")
return ;}
;	  M
:XR*:Macu::                                                                                                                                                                       	;{ Makuladegeneration
AutoDiagnose("Makuladegeneration {H35.3G}", "S")
return ;}
:XR*:Mageng::                                                                                                                                                                   	;{ Magengeschwür, Ulcus ventriculi
AutoDiagnose("Ulcus ventriculi {K25.3G}", "")
return ;}
:XR*:MammaCa::                                                                                                                                                              	;{ Mamma Carcinom
:XR*:Brustkr::
AutoDiagnose("Mamma-Ca {C50.9RG}", "")
return ;}
:XR*:Maras::                                                                                                                                                                      	;{ Marasmus / Kachexie
:XR*:Kachex::
AutoDiagnose("Alimentärer Marasmus {E41G}", "")
return ;}
:XR*:Menie::                                                                                                                                                                      	;{ Ménière-Krankheit
AutoDiagnose("Ménière-Krankheit {H81.0G}")
return ;}
:XR*:Meninge::                                                                                                                                                                  	;{ Meningeom
AutoDiagnose("Meningeom {D32.9G}", "S")
return ;}
:XR*:Meteo::                                                                                                                                                                      	;{ Meteorismus
AutoDiagnose("Meteorismus {R14G}", "")
return ;}
:XR*:Migr::                                                                                                                                                                        	;{ Migräne
AutoDiagnose("Migräne {G43.0G}", "")
return ;}
:XR*:Mitrali::                                                                                                                                                                      	;{ Mitralklappeninsuffizienz
:XR*:Mitralklappeni::
AutoDiagnose("Mitralinsuffizienz {I34.0G}")
return ;}
:XR*:Mukov::                                                                                                                                                                      	;{ Mukoviszidose
AutoDiagnose("Mukoviszidose {E84.9G}", "")
return ;}
:XR*:Myog::                                                                                                                                                                       	;{ Myogelosen
AutoDiagnose("Myogelosen {M62.89G}", "")
return ;}
:XR*:Myotrig::                                                                                                                                                                    	;{ Myogelosen , Triggerpunkt LWS
AutoDiagnose("Myogelosen {M62.89G} Triggerpunkt LWS {M62.88G}", "")
return ;}
;	  N
:XR*:Narkol::                                                                                                                                                                     	;{ Narkolepsie und Kataplexie
AutoDiagnose("Narkolepsie und Kataplexie {G47.4G}", "")
return ;}
:XR*:Nebenw::                                                                                                                                                                   	;{ Nebenwirkungen, unerwünschte
AutoDiagnose("unerwünschte Nebenwirkungen, <#> {T78.8G}")
return ;}
:XR*:Niereni::                                                                                                                                                                    	;{ Niereninsuffizienz alle Stadien
AutoDiagnose("Chronische Nierenkrankheit, Stadium 2 {N18.2}|Chronische Nierenkrankheit, Stadium 3 {N18.3}|Chronische Nierenkrankheit, Stadium 4 {N18.4}|Chronische Nierenkrankheit, Stadium 5 {N18.5}|Chronische Nierenkrankheit, Stadium 1 {N18.1}")
return ;}
:XR*:Nierenkolik::                                                                                                                                                               	;{ Ureterkolik durch Stein
:XR*:Ureterk::                                                                     	; Ureterkolik
AutoDiagnose("Nierenkolik{N20.1G}", "S")
return	;}
;	  O
:XR*:Obsti::                                                                                                                                                                       	;{ Obstipation bei Kolontransitstörung
AutoDiagnose("Obstipation bei Kolontransitstörung {K59.00G}", "")
return ;}
:XR*:Ödem::                                                                                                                                                                      	;{ Ödem
AutoDiagnose("Ödem, Unterschenkel {R60.9BG}", "")
return ;}
:XR*:Orient::                                                                                                                                                                      	;{ Orientierungsstörung
AutoDiagnose("Orientierungsstörung, nicht näher bezeichnet {R41.0G}", "")
return ;}
:XR*:Otitisex::                                                                                                                                                                     	;{ Otitis externa
:XR*:Gehörgangse::
:XRC*:OE::
AutoDiagnose("eitrige Otitis externa {H60.8G}",1)
return ;}
:XR*:Otitism::                                                                                                                                                                     	;{ Otitis media
:XR*:Mittelohr::
:XRC*:OM::
AutoDiagnose("Akute eitrige Otitis media {H66.0G}", 1)
return ;}
:XR*:OSAS::                                                                                                                                                                       	;{ Obstruktive Schlafapnoe
:XR*:Schlafa::
AutoDiagnose("Obstruktives Schlafapnoesyndrom {G47.31G}")
return ;}
;	  P
:XR*:Panik::                                                                                                                                                                       	;{ Panikstörung, Episodisch paroxysmale Angst
AutoDiagnose("Episodisch paroxysmale Angst {F41.0G}", "")
return ;}
:XR*:Parki::                                                                                                                                                                        	;{ Parkinson alle Stadien
AutoDiagnose("Primärer M.Parkinson (fehlende oder geringe Beeinträchtigung) {G20.00G}|Primärer M.Parkinson (mäßige bis schwere Beeinträchtigung) {G20.10G}|Primärer M.Parkinson-Syndrom (schwerste Beeinträchtigung) {G20.20G}")
return ;}
:XR*:pavk::                                                                                                                                                                        	;{ Periphere Gefäßkrankheit
AutoDiagnose("Periphere Gefäßkrankheit {I73.9G}", "")
return ;}
:XR*:Pankreasin::                                                                                                                                                               	;{ Pankreasinsuffizienz
AutoDiagnose("Pankreasinsuffizienz {K86.8G}", "")
return ;}
:XR*:Pect::                                                                                                                                                                         	;{ Angina pectoris
AutoDiagnose("Angina pectoris, V.a. {I20.9V}; KHK {I25.19V}", "")
return ;}
:XR*:Phar::                                                                                                                                                                        	;{ Pharyngitis, Akute
AutoDiagnose("Akute Pharyngitis, G. {J02.9G}", "")
return ;}
:XR*:Phleg::                                                                                                                                                                        	;{ Phlegmone versch. Lokalisationen
AutoDiagnose("Phlegmone an den Fingern {L03.01G}|Phlegmone an den Zehen {L03.02G}|Phlegmone an der oberen Extremität {L03.10G}|Phlegmone an der unteren Extremität {L03.11G}|Phlegmone im Gesicht {L03.2G}|Phlegmone am Rumpf {L03.3G}|Phlegmone, nicht näher bezeichnet {L03.9G}", "MS")
return ;}
:XR*:Pneumo::                                                                                                                                                                   	;{ Pneumonie
:XR*:Lungenen::                                                                   	; Lungenentzündung
AutoDiagnose("Pneumonie {J15.8G}", "S")
return ;}
:XR*:Pity::                                                                                                                                                                          	;{ Pityriasis versicolor
:XR*:Mala::                                              	; Malassezia furfur
:XR*:Kleienp::                                          	; Kleienpilzflechte
:XR*:Kleief::                                            	; Kleieflechte
AutoDiagnose("Pityriasis versicolor {B36.0G}", "")
return	;}
:XR*:Plattf::                                                                                                                                                                        	;{ Plattfuß
AutoDiagnose("Plattfuß [Pes planus] {M21.4G}", "S")
return ;}
:XR*:Polyarthri::                                                                                                                                                                 	;{ Polyarthritis, chronische
:XR*:chronischeP::
:XRC*:PCP::
:XR*:Gelenkrheu::
AutoDiagnose("Chronische Polyarthritis {M06.90G}")
return ;}
:XR*:Polyarthro::                                                                                                                                                                	;{ Polyarthrose
AutoDiagnose("Polyarthrose, nicht näher bezeichnet, bds. {M15.9BG}", "")
return ;}
:XR*:Polymya::                                                                                                                                                                   	;{ Polymyalgia rheumatica
AutoDiagnose("Polymyalgia rheumatica {M35.3G}", "")
return ;}
:XR*:Polyneu::                                                                                                                                                                     	;{ Polyneuropathie, hereditäre
:XR*:PNP::
AutoDiagnose("Hereditäre und idiopathische Neuropathie {G60.9G}", "S")
return ;}
:XR*:Posttrau::                                                                                                                                                                     	;{ Posttraumatische Belastungsstörung
AutoDiagnose("Posttraumatische Belastungsstörung {F43.1G}", "")
return ;}
:XR*:Prell::                                                                                                                                                                         	;{ Prellung diverse
:XR*:Prell::AutoDiagnose("Prellung des Knies {S80.0G}|Prellung des Kopfes {S00.95G}|Prellung des Fußes {S90.3G}", "MS")
:XR*:Kopfprell::AutoDiagnose("Prellung des Kopfes {S00.95G}", "S")
:XR*:Knieprell::AutoDiagnose("Prellung des Knies {S80.0G}", "S")
:XR*:Fußprell::AutoDiagnose("Prellung des Fußes {S90.3G}", "S")
;}
:XR*:Problem::                                                                                                                                                                   	;{ Problem / Kontaktanlässe
AutoDiagnose("Problem mit Bezug auf die Lebensführung {Z72.9G}|Kontaktanlässe mit Bezug auf die soziale Umgebung {Z60G}", "M")
return ;}
:XR*:Prostatah::                                                                                                                                                                 	;{ Prostatahyperplasie, benigne
:XRC*:Prostatav::                                                                   	; Prostatavergrößerung
:XRC*:BPH::
AutoDiagnose("benigne Prostatahyperplasie {N40G}", "")
return ;}
:XR*:Prostatak::                                                                                                                                                                 	;{ Prostata-Ca
:XR*:Prostata?C::
:XR*:ProstataC::
AutoDiagnose("Prostata-Ca {C61G}", "")
return ;}
:XR*:pulmonaleHy::                                                                                                                                                           	;{ pulmonale Hypertonie
:XR*:Lungenhoch::
AutoDiagnose("sekundäre pulmonale Hypertonie {I27.28G}")
return ;}
:XR*:Lebensf::                                                                                                                                                                    	;{ Problem mit Bezug auf die Lebensführung
AutoDiagnose("Problem mit Bezug auf die Lebensführung {Z72.9G}", "")
return ;}
:XR*:Psor::                                                                                                                                                                         	;{ Psoriasis
AutoDiagnose("Psoriasis {L40.0G}|Psoriasis pustulosa palmoplantaris {L40.3G}", "")
return ;}
:XR*:Pyelon::                                                                                                                                                                      	;{ Pyelonephritis
:XR*:Nierenbe::
AutoDiagnose("Akute Pyelonephritis {N10G}", "S")
return ;}
;	  Q
:XR*:Quin::                                                                                                                                                                        	;{ Quincke Ödem, Angioneurotisches Ödem
AutoDiagnose("Angioneurotisches Ödem {T78.3G}", "")
return ;}
;	  R
:XR*:Reflux::                                                                                                                                                                      	;{ Refluxösophagitis
AutoDiagnose("Refluxösophagitis {K21.0G}", "")
return ;}
:XR*:Retrop::                                                                                                                                                                      	;{ Retropatellararthrose, Chondromalacia patellae
AutoDiagnose("Chondromalacia patellae {M22.4BG}", "")
return ;}
:XR*:Riesenz::                                                                                                                                                                     	;{ Riesenzellarteriitis, Sonstige
AutoDiagnose("Sonstige Riesenzellarteriitis {M31.6LG}", "")
return ;}
:XR*:Rippenpr::                                                                                                                                                                  	;{ Rippenprellung, Thoraxprellung
:XR*:Thoraxpr::
AutoDiagnose("Prellung des Thorax {S20.2G}", "S")
return ;}
:XRC*:RSB::                                                                                                                                                                       	;{ Rechtsschenkelblock
AutoDiagnose("Rechtsschenkelblock {I45.1G}", "")
return ;}
;	  S
:XR*:Schar::                                                                                                                                                                       	;{ Scharlach
:XR*:Streptokokkena::	; Streptokokkenangina
AutoDiagnose("Scharlach {A38G}", "")
return ;}
:XR*:Schilddrüsenk::                                                                                                                                                          	;{ Schilddrüsenkarzinom
:XR*:Strumama::		; Struma maligna
AutoDiagnose("Bösartige Neubildung der Schilddrüse {C73G}", "")
return ;}
:XR*:Schizop::                                                                                                                                                                    	;{ Schizophrenie, nur Sonstige
AutoDiagnose("Sonstige Schizophrenie {F20.8G}", "")
return	;}
:XR*:Schlafs::                                                                                                                                                                     	;{ Schlafstörungen, diverse
:XR*:Einschlaf::
:XR*:Durchschlaf::
:XR*:Albträ::
AutoDiagnose("Ein- und Durchschlafstörungen {G47.0G}|Krankhaft gesteigertes Schlafbedürfnis {G47.1G}|Sonstige Schlafstörungen {G47.8G}|Schlafstörung, nicht näher bezeichnet {G47.9G}|Albträume [Angstträume] {F51.5G}", "M")
return ;}
:XR*:Schmerz::                                                                                                                                                                   	;{ chronisches Schmerzsyndrom
:XR*:chronischerS::
AutoDiagnose("Sonstiger chronischer Schmerz {R52.2G}", "")
return ;}
:XR*:Schnitt::                                                                                                                                                                      	;{ Schnittwunde + Leistungskomplex
DiagnoseUndZiffer("Schnittwunde {T14.1G}", "02300", "S")
return ;}
:XR*:Schwerh::                                                                                                                                                                   	;{ Schwerhörigkeit, diverse
:XR*:Presby::
:XR*:Altersschw::
:XR*:Hörver::
AutoDiagnose("Presbyakusis{H91.1G}|Hörverlust {H91.9G}", "S")
return ;}
:XR*:Schwindel::                                                                                                                                                                 	;{ Schwindel und Taumel
:XR*:Vertig::
AutoDiagnose("Schwindel und Taumel {R42G}", "")
return ;}
:XR*:SenkSp::                                                                                                                                                                     	;{ Knick-Senk-Spreizfuss
AutoDiagnose("Knick-Senk-Spreizfuss {M21.4G}", "S")
return ;}
:XR*:SickSin::                                                                                                                                                                     	;{ Sick-Sinus-Syndrom
AutoDiagnose("Sick-Sinus-Syndrom {I49.5G}", "S")
return ;}
:XR*:Sod::                                                                                                                                                                         	;{ Sodbrennen
AutoDiagnose("Sodbrennen {R12G}", "")
return ;}
:XR*:Somat::                                                                                                                                                                     	;{ Somatisierungsstörung
:XRC*:PSY::                                                                                                  ; PSYchosomatik
AutoDiagnose("Somatisierungsstörung {F45.0G};", "")
return	;}
:XR*:Spondylol::                                                                                                                                                                	;{ Spondylolisthesis
AutoDiagnose("Spondylolisthesis {M43.19G}", "")
return ;}
:XR*:Spondylos::                                                                                                                                                                	;{ Spondylosis derformans
AutoDiagnose("Spondylosis derformans {M47.99G}", "")
return ;}
:XR*:Suprave::                                                                                                                                                                   	;{ Supraventrikuläre Tachykardie
AutoDiagnose("Supraventrikuläre Tachykardie {I47.1G}", "")
return ;}
:XR*:Stauungsl::                                                                                                                                                                 	;{ Stauungsleber
AutoDiagnose("Chronische Stauungsleber {K76.1G}", "")
return ;}
:XR*:Sturzn::                                                                                                                                                                      	;{ Sturzneigung
AutoDiagnose("Sturzneigung {R29.6G}", "")
return ;}
:XR*:polyzystischeOv::                                                                                                                                                        	;{ Syndrom polyzystischer Ovarien
AutoDiagnose("Syndrom polyzystischer Ovarien {E28.2G}", "")
return	;}
:XR*:Synk::                                                                                                                                                                        	;{ Synkope und Kollaps
AutoDiagnose("Synkope und Kollaps {R55G}", "")
return ;}
:XR*:SVES::                                                                                                                                                                        	;{ SVES
AutoDiagnose("SVES {I49.4G}", "")
return ;}
;	  T
:XR*:Tachyk::                                                                                                                                                                     	;{ Tachykardie
AutoDiagnose("Tachykardie, nicht näher bezeichnet {R00.0G}", "")
return ;}
:XR*:Tia::                                                                                                                                                                           	;{ TIA, Zerebrale transitorische Ischämie
AutoDiagnose("Zerebrale transitorische Ischämie, nicht näher bezeichnet {G45.92G}", "")
return ;}
:XR*:Tinea::                                                                                                                                                                       	;{ Tinea pedis
:XR*:Fußp::
AutoDiagnose("Tinea pedis {B35.3G}", "S")
return ;}
:XR*:Tonsi::                                                                                                                                                                       	;{ Tonsillitis, akut ; Scharlach
:XR*:Mandel::
AutoDiagnose("Akute Tonsillitis durch sonstige näher bezeichnete Erreger {J03.8G}|Scharlach {A38G}")
return ;}
:XR*:Trige::                                                                                                                                                                       	;{ Trigeminusneuralgie
AutoDiagnose("Trigeminusneuralgie {G50.0G}", "S")
return ;}
:XR*:Trigg::                                                                                                                                                                       	;{ Triggerpunkt LWS
AutoDiagnose("Triggerpunkt LWS {M62.88G}", "")
return ;}
:XR*:TVT::                                                                                                                                                                         	;{ Tiefe Venenthrombose
:XR*:Throm::
:XR*:Venent::
AutoDiagnose("Tiefe Venenthrombose {I80.1G}", "S")
return ;}
;	  U
:XR*:Ulcus::                                                                                                                                                                        	;{ Ulcus ventriculi	- {vk20sc039} - Spacetaste
AutoDiagnose(JDia, "Ulcus")
return ;}
:XR*:Ungui::                                                                                                                                                                        	;{ Unguis incarnatus
:XR*:eingew::
:XR*:Zehnagel::
AutoDiagnose("Unguis incarnatus {L60.0G}", "S")
return ;}
:XR*:Urt::                                                                                                                                                                          	;{ Urtikaria - idiopathisch/chronisch
AutoDiagnose("Idiopathische Urtikaria {L50.1G}|Chronische Urtikaria {L50.8G}", "S")
return ;}
;	  V
:XR*:Variz::                                                                                                                                                                        	;{ Varizen
AutoDiagnose("Varizen der unteren Extremitäten ohne Ulzeration oder Entzündung, bds. {I83.9G}", "S")
return ;}
:XR*:Ventrikels::                                                                                                                                                                 	;{ Ventrikelseptumdefekt
AutoDiagnose("Ventrikelseptumdefekt {Q21.0G}", "")
return ;}
:XR*:VHF::                                                                                                                                                                         	;{ VHF div. / Vorhofflattern
:XR*:Vorhoff::
AutoDiagnose("Vorhofflimmern, permanent {I48.2G}|Vorhofflimmern, paroxysmal {I48.0G}|Vorhofflimmern, persistierend {I48.1G}|Vorhofflattern, typisch {I48.3G}|Vorhofflimmern und Vorhofflattern, nicht näher bezeichnet {I48.9G}")
return ;}
:XR*:Virusi::                                                                                                                                                                       	;{ Virusinfektion
:XRC*:VI::
AutoDiagnose("Virusinfektion{B34.9G}", "")
return ;}
:XR*:Virusw::                                                                                                                                                                     	;{ Viruswarzen
:XR*:Warze::
AutoDiagnose("Viruswarzen {B07}", "")
return	;}
:XR*:VitaminD::                                                                                                                                                                 	;{ Vitamin D Mangel
AutoDiagnose("Vitamin D Mangel {E64.8G}", "")
return ;}
:XR*:Vorhoft::                                                                                                                                                                     	;{ Vorhofthrombus
:XR*:intrakardiale::
AutoDiagnose("Vorhofthrombus {I51.3G}", "")
return	;}
;	 W
:XR*:Wirbelf::                                                                                                                                                                    	;{ Wirbelfraktur, Fraktur der Wirbelsäule
:XR*:Wirbelkörperf::
:XR*:Wirbelsint::
AutoDiagnose("Fraktur der Wirbelsäule, <#> {T08.0G}")
return ;}
:XRC*:WPW::                                                                                                                                                                    	;{ Wolff-Parkinson-White-Syndrom
:XR*:Wolff::
:XR*:White::
AutoDiagnose("Wolff-Parkinson-White-Syndrom {I45.6G}", "")
return ;}
;	  Z
:XR*:Zecke::                                                                                                                                                                      	;{ Zeckenbiß + Leistungskomplex
DiagnoseUndZiffer("Zeckenbiß {T14.03G}", "02300", 0)
return ;}
:XR*:Zentralv::                                                                                                                                                                   	;{ Zentralvenenthrombose
:XR*:Augenv::
AutoDiagnose("Zentralvenenthrombose {H34.8G}", "S")
return ;}
:XR*: Zwangss::                                                                                                                                                                 	;{  Zwangsstörungen
AutoDiagnose(" Zwangsstörungen {F42.8G}", "")
return	;}

;------- FILTERWORTE ------
:XR*:Atrioven::AutoDiagnose("Atrioventrikulärer","Atrioventrikulär")
:XR*:inkomp::AutoDiagnose("inkomplette(r)","inkomplett|unvollständig")
:XR*:kompl::AutoDiagnose("komplette(r)","komplett|vollständig")
:XR*:perman::AutoDiagnose("permanentes","permanent")
;:XR*:Prell::AutoDiagnose("Prellung","Prellung")
:XR*:prim::AutoDiagnose("primäre(r)","primär")
:XR*:seku::AutoDiagnose("sekundäre(r)","sekundär")
;:XR*:Suprav::AutoDiagnose("Supraventrikuläre","Supraventrikulär")


:XR*:ADStats::HotstringStatitistik(AddendumDir "\Module\Addendum\Include\AutoDiagnosen.ahk")
#If
;}


AutoDiagnose(DX, Options := "") {												                        						;-- Auto-ICD-Diagnosen

	; Options: S	- Seitenangabe zur Diagnose notwendig
	;           	M	- Mehrfachauswahl von Diagnosen möglich

	; ------------------------------------------------------------------------
	; Variablen, Liste und Optionen sichern
	; ------------------------------------------------------------------------ ;{

		global hACB, hACBLb, ACB, ACBLb, fchwnd
		static DiaTmp, DiaAnzahl, goLeft, tDX, seite, lastPos, dpiF:= screenDims().DPI / 96
		static FilterWords, lastwords
		static oACB := Object()

	; kurz die Tastatureingabe blockieren damit der User nach Hotstringauslösung nicht sofort weiter schreiben kann
		BlockInputShort(500)

		If IsObject(DX)
		{
				AutoPrediction(DX, Options)
		}
		else
		{
				If !RegExMatch(DX, "\{.*\}")
				{
						FilterWords   	.= Options "|"							; z.B. :XR::inkompl:AutoDiagnose("inkompletter", "inkomplett|unvollständig") - die Worte in Options sollten Wortstämme sein
						lastwords      	.= DX "|"
						ADInterception	:= true
						SendRaw, % DX
						Send, {Space}
						return
				}

				tDX	            	:= DX
				seite	            	:= InStr(Options, "S")
				Multi	            	:= InStr(Options, "M")
				ADInterception	:= false                                      	; eventuell muss diese flag  noch anders zurück gesetzt werden können

			; zuvor eingegebene Worte aus dem RichEdit Control entfernen
				If (StrLen(lastwords) > 0)
						RemoveLastWords(lastwords)

				lastwords:= ""


		}
	;}

	; ------------------------------------------------------------------------
	; Zeilen für die Ausgabe zählen - maximal 10 gleichzeitig!
	; ------------------------------------------------------------------------ ;{
		If !InStr(DX, "|")
		{
				Diagnose:= DX
				gosub Seitenangabe
				return
		}
		else
		{
			;- Erstellen des Listboxinhaltes, ist die Liste leer, dann gibt es die Diagnose für diese Wortkombination nicht oder noch nicht
				PrevDia  	 := FilterDiagnosen(DX, FilterWords)
				If (PrevDia = "")
					PrevDia := FilterDiagnosen(DX, "")

			;- zählt die Anzahl der Diagnosen für die Listbox
				Loop, Parse, PrevDia, `|
					If A_Index < 11
						rows:= A_Index
					else
						break

			;- Filterliste und letzte Worte leeren
				FilterWords:= ""

			;- Gui aufrufen
				AutoCompleteGui(PrevDia, rows,  "x" A_CaretX " y" A_CaretY " NA Hide", Multi)
				a:= GetWindowSpot(hACB), c:= GetWindowSpot(fchwnd:= AlbisGetActiveControl("hwnd"))
				Gui, ACB: Show, % "x" Floor(A_CaretX // DpiF) " y" (c.Y - a.H) " NA"
				If Multi
					GuiControl, ACB: Focus, ACBLb
		}
	;}

	; ------------------------------------------------------------------------
	; Eingabefokus zurückgeben ins RichEdit-Control von Albis
	; ------------------------------------------------------------------------ ;{
		WinActivate, ahk_class OptoAppClass
		ControlFocus,, ahk_id %fchwnd%
		return
	;}

AutoCBLBres:   	;{ Ausgabe der Diagnosen nach Albis

	; ------------------------------------------------------------------------
	; ausgewählten Eintrag oder Einträge aus der Liste heraussuchen
	; ------------------------------------------------------------------------ ;{
		Gui, ACB: Submit
		If !lastpos 	;1. Ablauf: Eintragen der Diagnose
		{
				ACBLb:= RegExReplace(ACBLb, "\s*\(.*\)")
				diaWahl    	:= StrSplit(ACBLb, "|")
				diaAuswahl	:= StrSplit(tDX, "|")
				Diagnose  	:= ""

				For Each, Item in diaAuswahl
				{
						Loop, % diaWahl.MaxIndex()
						{
							If InStr(Item, diaWahl[A_Index]) = 1
								If RegExMatch(A_LoopField, "(?<=\:).*", str)
									Diagnose .= RegExReplace(str, "\}\s*\;*\s*", "}; ")
								else
									Diagnose .= RegExReplace(Item, "\}\s*\;*\s*", "}; ")
						}
				}
		}
			else  	;2. Ablauf: Eintragen der Seitenangabe
		{
				Diagnose	:= ACBLb
				lastpos     	:= 0
		}
	;}

	Seitenangabe:
	; ------------------------------------------------------------------------
	; Eingabefokus zurückgeben ins RichEdit-Control von Albis
	; ------------------------------------------------------------------------ ;{
		WinActivate, ahk_class OptoAppClass
		ControlFocus,, % "ahk_id " fchwnd
		If WinExist("AutoDiagnose")
			Gui, ACB: Destroy
	;}

	; ------------------------------------------------------------------------
	; Seitenangabe ergänzen lassen, falls notwendig
	; ------------------------------------------------------------------------ ;{
		If seite > 0
		{
			; auf 0 setzen damit dieser Abschnitt beim 2.Durchlauf nicht aufgerufen wird
				seite:= startpos:= DiaAnzahl:= 0

			; sucht die letzte Diagnose in einer Diagnosereihe heraus und fragt dann dort nach der Seitenangabe
				Loop
					If (startpos:= RegExMatch(Diagnose, "O)([G|A|V|Z|\s*]\})", "", startpos +1))
						lastpos := startpos, DiaAnzahl:= A_Index
					else
						break

			; temporär ein 'L" an die gefundene Position einfügen und den Diagnosentext senden
				DiaTmp := Diagnose := SubStr(Diagnose, 1, lastpos - 1) "L" SubStr(Diagnose, lastpos, StrLen(Diagnose))
				SendRaw, % Diagnose

			; den Cursor zu dieser Position bringen, die Caret-Position speichern und dann erst das 'L' selektieren
				aX:= A_CaretX, aY:= A_CaretY
				MoveCursorSelect("Left " (goLeft:= StrLen(Diagnose) - lastpos + 1), "Left 1")

			; Gui für die Auswahl der Seitenangabe erstellen
				AutoCompleteGui("L|R|B| ", 4, "x" aX " y" aY " NA Hide", 0)
				a:= GetWindowSpot(hACB), c:= GetWindowSpot(fchwnd:= AlbisGetActiveControl("hwnd"))
				Gui, ACB: Show, % "x" Floor(A_CaretX // DpiF) " y" (c.Y + c.H - a.H) " NA"
				return
		}
		else If RegExMatch(Diagnose, "^(L|R|B|\s)")
		{
				If DiaAnzahl > 1
				{
					; setzt bei allen Diagnosen in einer Diagnosenkette die Seitenangabe hinzu
						DiaSend := ""
						Loop, Parse, DiaTmp, `;, A_Space
							DiaSend .= RegExReplace(A_LoopField, "[RLG\s]*([GAVZ]\})", Diagnose "$1; ")

					; Text im Albiscontrol ändern
						CText := StrReplace(ControlGetText("", "ahk_id " fchwnd), DiaTmp, DiaSend)
						VerifiedSetText("", CText, fchwnd)

					; Cursor an die letzte Eingabeposition setzen
						Send, % "{Right " (goRight:= StrLen(CText) - InStr(CText, DiaSend) + 1) "}"
				}
				else
				{
						Send, % "{Right " (goLeft) "}{Space}"
				}
		}
		else if (sPos:= InStr(Diagnose, "<#>"))
		{
				SendRaw, % StrReplace(Diagnose, "<#>", "")
				MoveCursorSelect("Left " (goLeft:= StrLen(Diagnose) - sPos + 1), "Right 3")
				;ToolTip, % "Bitte den Freitext ergänzen!", % A_CaretX - 20, % A_CaretY - 10, 5
				;SetTimer, TTOff, -3000
		}

		if seite = 0
		{
				SendRaw, % Diagnose
				DiaTmp := Diagnose
		}

	;}

return ;}

ACBListbox:                                                                       	;{
	if (A_GuiEvent="DoubleClick")
		gosub AutoCBLBres
return ;}

TTOff:                                                                               	;{
	ToolTip,,,, 5
return ;}
}

FilterDiagnosen(DX, Filter:="") {

	; entfernt alle Diagnosen aus der Diagnosenliste welche nicht die Filterworte enthalten
		If (Filter <> "")
		{
				Filter:= StrSplit(RTrim(Filter, "|"), "|")

				Loop, Parse, DX, `|
					Loop, % Filter.MaxIndex()
						If InStr(A_LoopField, Filter[A_Index])
						{
								tmpDX.= A_LoopField "|"
								continue
						}

				DX := RTrim(tmpDX, "|")
		}

		Loop, Parse, DX, `|
			If InStr(A_LoopField, ":")
			{
					RegExMatch(A_LoopField, "(?<=\:)(.*)", TmpDiaListe)
					RegExMatch(A_LoopField, "^.*(?=\:)", TmpPrevDia)
					PrevDia .= TmpPrevDia " (Diagnosenreihe)|"
			}
			else
					PrevDia .= A_LoopField "|"

return RTrim(PrevDia, "|")
}

AutoPrediction(DX, trunk) {											                                       						;-- die smartere Auto-ICD Funktion soll das werden

	;For hotstring in DX[trunk]
	;	Hotstring()

}

AutoCompleteGui(Content, rows, GuiOptions, Multi:= 0, FontOptions:= "s11 Normal q5", FontName:= "Futura Bk Bt") {

		global hACB, hACBLb, ACB, ACBLb

	; ------------------------------------------------------------------------
	; AutoComplete Gui erstellen
	; ------------------------------------------------------------------------ ;{
		Gui, ACB: New 	, -SysMenu -Caption +AlwaysonTop +ToolWindow +HWNDhACB ;0x98200000
		Gui, ACB: Margin	, 0, 0
		If Multi
		{
				Gui, ACB: Color, c172842
				RegExMatch(FontOptions, "(?<=s)\d+", fsize)
				smallFont:= RegExReplace(FontOptions, "(?<=s)\d+", fsize - 1)
				Gui, ACB: Font 	, % smallFont , % FontName
				Gui, ACB: Add	, Text, % "w" LBEX_CalcIdealWidth(0, Content, "|", FontOptions, FontName) " HWNDhTACB cWhite CENTER", mehrere Diagnosen auswählbar (Strg + Mausklick)
				t:= GetWindowSpot(hTACB)
		}
		Gui, ACB: Font  	, % FontOptions, % FontName
		Gui, ACB: Add  	, ListBox, % (Multi = 0 ? "": "xm y" T.H+2 " ") "Choose1 HWNDhACBLb vACBLb gACBListbox" (Multi = 0 ? "": " Multi")  " r" rows " w" LBEX_CalcIdealWidth(0, Content, "|", FontOptions, FontName), % Content	;+0x8

		Gui, ACB: Show	, % GuiOptions, AutoDiagnose
	;}

		AutoCompleteGuiHotkeys("On")

return hACB

ACBMoveDown: ;{
	ControlSend,, {Down}, % "ahk_id " hACB
return
ACBMoveUp:
	ControlSend,, {Up}, % "ahk_id " hACB
return ;}
ACBCheck: ;{
	MouseGetPos, mx, my, hWin
	If (hWin = hACB)
	{
		MouseClick, Left, % mx, % my
		return
	}
;}
ACBGuiClose: ;{
		If InStr(A_ThisHotkey, "RButton")
			Send, {RButton}
		Gui, ACB: Destroy
		AutoCompleteGuiHotkeys("Off")
		hACB:= hACBLb:= ""
return ;}

}

AutoCompleteGuiHotkeys(status:="On") {                                                                               	;-- Hotkeys für die Nutzereingaben im AutoCompleteGui

	Hotkey, IfWinExist	, AutoDiagnose ahk_class AutoHotkeyGUI
	Hotkey, Enter     	, AutoCBLBres   	, % status
	Hotkey, Down   	, ACBMoveDown	, % status
	Hotkey, Up    		, ACBMoveUp   	, % status
	Hotkey, Esc	    	, ACBGuiClose  	, % status
	Hotkey, Left	    	, ACBGuiClose  	, % status
	Hotkey, Right	    	, ACBGuiClose   	, % status
	Hotkey, LButton  	, ACBCheck       	, % status
	Hotkey, RButton  	, ACBGuiClose   	, % status
	Hotkey, IfWinExist

}

DiagnoseUndZiffer(Diagnose, Ziffer, Seitenangabe) {

	SendRaw, % Diagnose
	Send, {Tab}
	SendRaw, % "lko"
	Send, {Tab}
	SendRaw, % "02300"
	Send, {Tab}
	If Seitenangabe
			Send, {ShiftUp}{Left 2}{ShiftDown}

return
}

CaretPos(ControlId) {                                                                                                              	;-- Get start and End Pos of the selected string - Get Caret pos if no string is selected
	;https://autohotkey.com/boards/viewtopic.php?p=27979#p27979
	DllCall("User32.dll\SendMessage", "Ptr", ControlId, "UInt", 0x00B0, "UIntP", Start, "UIntP", End, "Ptr")
	SendMessage, 0xB1, -1, 0, , % "ahk_id" ControlId
	DllCall("User32.dll\SendMessage", "Ptr", ControlId, "UInt", 0x00B0, "UIntP", CaretPos, "UIntP", CaretPos, "Ptr")
	if (CaretPos = End)
	SendMessage, 0xB1, % Start, % End, , % "ahk_id" ControlId	;select from left to right ("caret" at the End of the selection)
		else
	SendMessage, 0xB1, % End, % Start, , % "ahk_id" ControlId	;select from right to left ("caret" at the Start of the selection)
	CaretPos++	;force "1" instead "0" to be recognised as the beginning of the string!
return, CaretPos
}

LBEX_CalcIdealWidth(HLB, Content := "", Delimiter := "|", FontOptions := "", FontName := "") {  ;-- zum Berechnen der optimalen Breite einer Listbox
   DestroyGui := MaxW := 0
   If !(HLB) {
      If (Content = "")
         Return -1
      Gui, LB_EX_CalcContentWidthGui: Font, % FontOptions, % FontName
      Gui, LB_EX_CalcContentWidthGui: Add, ListBox, hwndHLB, % Content
      DestroyGui := True
   }
   ControlGet, Content, List,,, % "ahk_id " HLB
   Items := StrSplit(Content, "`n")
   SendMessage, 0x31, 0, 0,, % "ahk_id " HLB ; WM_GETFONT
   HFONT	:= ErrorLevel
   HDC  	:= DllCall("User32.dll\GetDC", "Ptr", HLB, "UPtr")
   DllCall("Gdi32.dll\SelectObject", "Ptr", HDC, "Ptr", HFONT)
   VarSetCapacity(SIZE, 8, 0)
   For Each, Item In Items {
      DllCall("Gdi32.dll\GetTextExtentPoint32", "Ptr", HDC, "Ptr", &Item, "Int", StrLen(Item), "UIntP", Width)
      MaxW := Width > MaxW ? Width : MaxW
   }
   DllCall("User32.dll\ReleaseDC", "Ptr", HLB, "Ptr", HDC)
   If (DestroyGui)
      Gui, LB_EX_CalcContentWidthGui: Destroy
   Return MaxW + 8 ; + 8 for the margins
}

BlockInputShort(time) {                                                                                                             	;-- verhindert nach Zeitvorgabe Tasten- und Mauseingaben
	BlockInput, On
	SetTimer, BlockInputOff, % "-" time
return
BlockInputOff:
	BlockInput, Off
return
}

NeueDiagnose(ADInterception) {                                                                                              	;-- verhindert das Hotstrings beim Korrigieren einer Diagnose  ausgelöst werden

		; Funktion liest den Inhalt des RichEditControls aus und ermittelt die Position des Eingabecursor (Caret)
		; steht der Cursor innerhalb eines Diagnosetextes oder genau am Anfang wird false zurückgegeben, ansonsten true
		; dies verhindert das ein Hotstring während der manuellen Änderung einer Diagnosebezeichnung ausgelöst wird

		static lastactive
		;ToolTip, % "Zeile: " A_LineNumber "`nADInterception: " ADInterception, 800, 600, 7
		Critical

		hactiveID    	:= AlbisGetActiveControl("hwnd")
		If (hactiveID <> lastactive)
				ADInterception := 0, lastactive := hactiveID

		If (ADInterception = true)
				return true

		thisHotString	:= RegExReplace(A_ThisHotkey, "\:.*\:")
		controlText	:= ControlGetText("", "ahk_id " hactiveID)
		CaretPos		:= CaretPos(hactiveID) - StrLen(ThisHotString)
		TextUpToCP	:= SubStr(controlText, 1, CaretPos) SubStr(controlText, CaretPos + StrLen(thisHotString), 1)
		If RegExMatch(TextUpToCP, "(;\s*$)|(^\s*$)") || (controlText = "")
				return true

return false
}

MoveCursorSelect(Move, Select) {                                                                                             	;-- Cursor verschieben und anschließend eine Anzahl an Zeichen selektieren
	Send, % "{" Move "}"
	Sleep, 200
	Send, % "{Shift Down}{" Select "}{Shift Up}"
	Sleep, 300
	Send, % "{Shift Up}"
	Send, % "{LControl Up}"
	Send, % "{Control Up}"
}

HotstringStatitistik(file) {                                                                                                         	;-- zeigt die Anzahl der Abkürzungen und Diagnosen an

	abbreviations := 0
	diagnosis		 := 0

	FileRead,f, % file
	Loop, parse, f, `n, `r
	{
			If RegExMatch(A_LoopField, "\:\w+\:\:")			;\w+\:\:
					abbreviations ++, continue
			else if RegExMatch(A_LoopField, "(\{[\w\.]*\})")
			{
					RegExReplace(A_LoopField, "(\{[\w\.]*\})", "", rplcount)
					diagnosis += rplcount
			}
	}

	text =
	(LTrim
	    HOTSTRINGSTATISTIK FÜR AUTO-ICD-DIAGNOSEN
	---------------------------------------------------------------------
	Abkürzungen: `t%abbreviations%
	Diagnosen:      `t%diagnosis%

	Möchten Sie mehr sehen, dann drücken Sie auf 'Ja'
	)
	MsgBox, 4, Addendum für Albis on Windows, % text

}

RemoveLastWords(lastwords) {                                                                                                	;-- entfernt die Worte aus einem Edit oder RichEdit-Control

		fchwnd		:= GetFocusedControl()
		ctext      	:= ControlGetText(fchwnd)
		words   	:= StrSplit(lastwords, "|")

		; die RichEdit Befehle bringen Albis machmal zum Absturz oder funktionieren nicht
		Loop, % words.MaxIndex()
		{
				CaretPos	:= CaretPos(fchwnd)
				idx			:= 1 + words.MaxIndex() - A_Index
				word	:= words[idx]
				if word = ""
						break
				wordStart	:= InStr(CText, word)
				wordEnd	:= wordstart + StrLen(word)
				goCaret	:= wordEnd - CaretPos
				;ToolTip, % "word: " word "`nwordstart: " wordstart "`nwordend: " wordEnd "`nCaretPos: " CaretPos "`ngoCaret: " goCaret, 800, 500, 5

				If goCaret <> 0
				{
						If goCaret < 0
							Send, % "{Right " (goCaret * -1) "}"
						else
							Send, % "{Left " goCaret "}"
				}

				Send, % "{BackSpace " StrLen(word) +1 "}"
		}
}



﻿; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;     	∑ ** Addendum_MDCalc **
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;		Description:	    	a class library of medical formulas
;
; 		Contents:
;
;       Dependencies:		none
;
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;       ☢ REMARKS:      		I have carefully checked the algorithms as far as possible.
;										However, I cannot guarantee the accuracy of the results of your calculations!
;
;									1. Because, programming often produces errors that are not always immediately recognizable.
;									2. Because, as the only programmer, I may not have the skills to correctly implement algorithms
;									    and therefore not be able to check the correctness of the outputs.
;									3. Because, my specialist medical understanding is insufficient without me suspecting it.
;									4. Because, people make mistakes!
;									5. Because, I cannot calculate a lack of specialist knowledge on the part of the user.
;
;									    Most of the formulas used have limited informative value under already known, possibly also
; 									    unknown circumstances. For each formula there is a definition of the criteria for which this is
; 										best used. These criteria have been summarized by professional societies over the years.
; 										Before using the algorithms, you should inform yourself in detail!
;       -------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;	    Addendum für Albis on Windows by Ixiko started in September 2017 - this file runs under Lexiko's GNU Licence
;       Addendum_Calc started:           	13.01.2021
;       Addendum_Calc last change:    	15.01.2021
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

/* 	MDRD

		GFR (ml/min/1,73 m2) = 186 x (KreatininSerum, mg/dl)-1,154 x (Alter, Jahre)-0,203 x (0,742 bei Frauen)

		gekürzte MDRD-Formel (Modification of Diet in Renal Disease) :
			GFR (ml/min/1,73m2) = 186 x S-Krea^-1,154 x Alter^-0,203 [x 0,742 nur bei Frauen] [x 1,21 bei Patienten mit schwarzer Hautfarbe]
			Korrektur auf Körperoberfläche: GFR durch KOF aus Nomogramm teilen!

		lange MDRD-Formel (genauer als Kurzformel):
			GFR (ml/min/1,73m2) = 170 x S-Krea -0,999 x (Harnstoff/2,144) -0,170 x (Albumin/10)
													+0,318 x Alter -0,176 [x 0,742 nur bei Frauen] [x 1,21 bei Patienten mit schwarzer Hautfarbe]
			Korrektur auf Körperoberfläche: GFR durch KOF aus Nomogramm teilen!

						Quelle:
			http://www.laborlexikon.de/Lexikon/Infoframe/k/Kreatinin-Clearance.htm

 */

/*  	Berechnung der Körperoberfläche aus Gewicht und Körpergröße:

		-	nach DuBois, veröffentlicht 1916: Körperoberfläche [m²] = Gewicht[kg]^0,425 x Größe[cm]^0,725 x 0.007184
		-	nach Mosteller, veröffentlicht 1897: Körperoberfläche [m²] = ((Gewicht[kg] x Größe[cm])/3600)^0,5

		Korrektur der Körperoberfläche:  bei Schätzung der GFR nach Cockroft und Gault:
		-	die beim Patienten individuell errechneten Werte müssen mittels eines Nomogrammes auf 1,73 m2
			Körperoberfläche mittels der untenstehenden Korrekturformel umgerechnet werden
		-	Korrekturformel auf KOF:
			C-Krea [ml/min. x 1,73 m2] = 	C-Krea x 1,73
															——————
																  KOF


*/

class MDCalc {

	BMI(weightkg, heightmeters) {                                                                        	;-- i never calculate the BMI, but maybe you!
		; https://github.com/jonjlee/schapp/blob/master/CalcFuns.ahk
	return Round(weightkg / (heightmeters * heightmeters),1)
	}

	WaistToHeightRatio(waistcfcm, heightcm, precision:=2) {                                  	;-- waist to height ratio

		; waistcfcm = Waist circumference [cm]
	return Round(waistcfcm/heightcm, precision)
	}

	highDoseAmoxicillin(kg) {                                                                                 	;-- Amoxicillin - Medication advice as text

		; source: https://github.com/jonjlee/schapp/blob/master/CalcFuns.ahk
		; ### NOT TESTED!!

		suspdosing	:= [400/5, 250/5]
		capdosing 	:= [250]
		adultdose   	:= 875 * 2

		hi := kg * 100, lo := kg * 80

		if (lo <= adultdose && hi >= adultdose) {
			return "875mg BID (use adult dose, which is < high dose amox)"
		} else {
			dosing := 400/5, dose := lo/2
			inQuarterMLs := Round(Ceil(dose / dosing * 4) / 4, 1)
			mg := round(dosing * inQuarterMLs)
			return mg "mg (" inQuarterMLs "ml) BID (" Round(mg*2/kg) "m/k/d - Amoxicillin (hohe Dosis: 80-90 mg/kgKG/Tag geteilt in 2 Dosen)"
		}
	}

	KOF(heightcm, weight, precision:=2, formula:="Dubois") {                             	;-- correction of body surface base on height and weight for adults and children

		; ⚕	5 formula to calculate the body surface of adults, 1 for children
			If (formula = "Dubois")
				KOF := (weight**0.425) * (heightcm**0.725) * 0.007184
			else if (formula = "Mosteller")
				KOF :=  ((weight * heightcm)/3600)**0.5
			else If (formula = "Takahira")
				KOF := (weight**0.425) * (heightcm**0.725) * 0,007241

		; Haycock formula: 	It is more precise than other formulas, especially for smaller body surfaces.
		; 		 weight units: 	[kg]
		; 				 source: 	https://flexikon.doccheck.com/de/Haycock-Formel
			else If (formula = "Haycock")
				KOF := (weight**0.5378) * (heightcm**0.3964) * 0.02426

		; Gehan-George: 		It is based on the analysis of data from 401 patients
		; 	   weight units: 		[kg]
		; 			   source: 		https://flexikon.doccheck.com/de/Gehan-George-Formel
			else If (formula = "Gehan-George")
				KOF := (weight**0.42246) * (heightcm**0.51456) * 0.0235

		; Boyd formula:     	should be used to calculate the body surface of children.
		;   weight units:       	[g]
		; 		    source: 		https://flexikon.doccheck.com/de/Boyd-Formel
			else If (formula = "Boyd")	{
				If (weight <= 300)
					return 0
				KOF := (weight**0.7285 - (0.0188 * log(weight) * (heightcm**0.3))) * 0.0003207 ; ### NOT TESTED!
			}

	return Round(KOF, precision)  ; KOF [m²]
	}

	LabUnitsConverter(Lab, LabVal, Unit, Precision:=3, lang:="") {                        	;-- laboratory units converter

		/* ⚕	Description of laboratory unit converting function

			⊛ Parameters:
				⚬ Lab    	= name of laboratory unit (can be english or german)
				⚬ LabVal 	= value to convert
				⚬ Unit   	= unit to convert from (can be SI or C)
				⚬ Precision= how many decimal places should the conversion value have
				⚬ lang   	= use this for a formatted output of the unit in another language

			⊛ The conversion from Conventional unit to SI unit is done by multiplying it with the conversion factor (CF)
				, from SI to convention unit by dividing

			⊛ this table is used to get conversion factors:
				https://accessmedicine.mhmedical.com/content.aspx?bookId=1069&sectionId=60775149


		*/
		/*     	Description Of CFactors Object Structure

			⊛ indexed objects with conversion data for laboratory units
				⚬ keys with fixed names are:
					- SI	= Standard International Unit
					- C	= Conventional Unit
					- CF	= Conversion Factor
				⚬ Key with names of laboratory values in German and English:
					- The values contain the abbreviations of the language [en for English and de for German]
					- There are also short names for the laboratory units avaible marked with an -s

		*/

		static CFactors := {1: {	"Kreatinin" 	: "de", "Krea"  	: "de-s",	"Creatinine"	: "eng", "SI" 	: "µmol/l" , "C" 	: "mg/dl", "CF" : 88.4}
								,	 2: {	"Harnstoff"		: "de", "HST"   	: "de-s", "Urea"			: "eng", "SI"	: "mmol/l", "C"	: "mg/dl", "CF" : 0.357}
								,	 3: {	"Harnstoff"		: "de", "HST"   	: "de-s", "Urea"			: "eng", "SI"	: "mmol/l", "C"	: "mg/dl", "CF" : 0.357}}

		For CFIndex, oLab in CFactors
			If oLab.haskey(Lab) {
				If (!StrLen(lang) > 0) {
					return (Unit = "SI") ?	{"LabVal":Round(LabVal/oLab.CF, 3), "Unit":oLab.C, "CF":oLab.CF}
												:	{"LabVal":Round(LabVal*oLab.CF, 2), "Unit":oLab.SI, "CF":oLab.CF}
				}
				else {
					For key, val in oLab
						If (val = lang)
							return key
				}
			}

	}

	CKD_EPI(Scr, Units, height, weight, age, sex, african:=false, precision:=2) {	   	;--

		/* ⚕	CKD-EPI-Formular for eGFR calculation

			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
				∑	GFR = 141  x  (min(SKr/κ, 1)^α)  x  (max(SKr/κ, 1)^-1.209)  x  0.993^Alter
					For women, the value is also multiplied by 1.018, for black skin by 1.159.
			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
			   ⚠This equation should only be used for patients 18 and younger.
					All estimating equations used in adult populations become less accurate as GFR increases, such as in people with normal kidney function.
					This equation is recommended when eGFR values above 60 mL/min/1.73 m² are desired.
			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
					units md/dl to µmol/l
					#1 mg/dl  = 88.4 µmol/l
			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
			    ⛖ If Scr in [mg/dl] use 	< units = "C" >
						κ 	= 0.742 	(female)	, 	 0.9  	 (male)
						α 	= -0.329 	(female)	, 	-0.411	(male)
					If Scr in [µmol/L] use	< units = "SI" >
						κ 	= 65.59 	(female)	, 	 79.6	(male)
						α 	= -0.329	(female)	, 	-0.411	(male)
			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
				⊛	Scr 		= serum creatinine
					κ & α	= sex base factors
					min 		= minimum	of SKr/κ and 1
					max 		= maximum	of SKr/κ and 1
					age 		= age in years
					sex		= "f" for female and "m" for male
					african	= true if black coloured African or African American
			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
			  ℹ The above formula is based on a standardized body surface area of 1.73 m2. They are of limited significance for:
    					⚬	Children and adolescents
	    				⚬	very old age
		    			⚬	very overweight or underweight
			    		⚬	extreme muscle mass (bodybuilding)
			    		⚬	Skeletal muscle disease
			────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
					adapted from python: https://github.com/gyonmyon/Glomerular_Filtration_Rate/blob/master/gfr.py


		 */

		female  	:= (sex = "f"	)	? 1.018 : 1
		african   	:= (african   	)	? 1.159 : 1
		α          	:= (sex = "m")	? (-0.411) : (-0.329)
		F      		:= -1.209

		If  (sex = "m") {
			If (units = "SI")
				κ := 79.6
			else
				κ := 0.9
		}
		else if (sex = "f") {
			If  (units = "SI")
				κ := 65.59
			else
				κ := 0.742
		}

		minScr	:= min((Scr/κ)	, 1)
		maxScr	:= max((Scr/κ)	, 1)
		eGFR 	:= 141 * ( minScr**α ) * ( maxScr**F ) * ( 0.993**age ) * female * african

	return Round(eGFR, precision)
	}

	WeightBasedDose(kg, mgPerKg, suspConc, maxMg) {                                    	;--

		; source: https://github.com/jonjlee/schapp/blob/master/CalcFuns.ahk
		; ### NOT TESTED!

		maxML = maxMg / suspConc
		mL := kg * mgPerKg / suspConc
		mL := Round(Ceil(mL * 2) / 2, 1)
		mL := (mL > maxML) ? maxML : mL
		dose := round(mL * suspConc)

    return dose "mg (" mL "mL) q4h (" Round(dose/kg, 1) "mg/kg, " Round(suspConc*5, 1) "mg/5mL)"
	}

	ntproBNPcorrection(NTproBNP, eGFR) {                                                            	;--

		/*  ⚕	NT-pro-BNP-Korrektur

			Wie Studien zeigen ist die NT-proBNP-Konzentration umgekehrt proportional zur Nierenfunktion. Die Ursache dafür ist, dass das Molekül auch über die Niere eliminiert wird.
			Somit besteht, besonders bei Nierenfunktionsstörungen, das Risiko, dass der NT-proBNP-Wert die kardiale Situation nicht korrekt abbildet. Die Formel erlaubt eine
			Interpretation des NT-proBNP-Wertes unter Berücksichtigung der Nierenfunktion anhand der eGFR. Dadurch steigt die Spezifität, gleichzeitig bleiben der hohe negativ
			prädiktive Wert (NPV) und die hohe Sensitivität erhalten.

			Hinweis
			Die Formel sollte nicht bei eGFR über 75 (ml/min/1.73 m²) verwendet werden.

			Literatur
			Luchner A, Weidemann A, Willenbrock R, Philipp S, Heinicke N, Rambausek M, Mehdorn U, Frankenberger B, Heid IM, Eckardt KU, Holmer SR.
			Improvement of the cardiac marker N-terminal-pro brain natriuretic peptide through adjustment for renal function: a stratified multicenter trial.
			Clin Chem Lab Med 2010;48(1):121–128

			Quelle: https://www.bioscientia.info/rechner-medizinische-formeln/nt-pro-bnp-korrektur/index.html


		*/



	}


}





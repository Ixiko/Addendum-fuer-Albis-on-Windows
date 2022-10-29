; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
;       Addendum_Calc last change:    	05.04.2022
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

/* 	MDRD

		—————————————————————————————————————————————————————————
		Deutsch
		—————————————————————————————————————————————————————————
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


		—————————————————————————————————————————————————————————
		english
		—————————————————————————————————————————————————————————
		GFR (ml / min / 1.73 m2) = 186 x (creatinine serum, mg / dl) -1.154 x (age, years) -0.203 x (0.742 in women)

		Abbreviated MDRD formula (Modification of Diet in Renal Disease):
		GFR (ml / min / 1.73m2) = 186 x S-Krea ^ -1.154 x age ^ -0.203 [x 0.742 for women only] [x 1.21 for black patients]
		Correction on body surface: divide GFR by KOF from nomogram!

		long MDRD formula (more precisely as a short formula):
		GFR (ml / min / 1.73m2) = 170 x S-Krea -0.999 x (urea / 2.144) -0.170 x (albumin / 10)
		+0.318 x age -0.176 [x 0.742 for women only] [x 1.21 for black patients]
		Correction on body surface: divide GFR by KOF from nomogram!

		Source:
		http://www.laborlexikon.de/Lexikon/Infoframe/k/Kreatinin-Clearance.htm

*/

/*  	Calculation of the body surface from weight and height:

		—————————————————————————————————————————————————————————
		Deutsch
		—————————————————————————————————————————————————————————
		-	nach DuBois, veröffentlicht 1916: Körperoberfläche [m²] = Gewicht[kg]^0,425 x Größe[cm]^0,725 x 0.007184
		-	nach Mosteller, veröffentlicht 1897: Körperoberfläche [m²] = ((Gewicht[kg] x Größe[cm])/3600)^0,5

		Korrektur der Körperoberfläche:  bei Schätzung der GFR nach Cockroft und Gault:
		-	die beim Patienten individuell errechneten Werte müssen mittels eines Nomogrammes auf 1,73 m2
			Körperoberfläche mittels der untenstehenden Korrekturformel umgerechnet werden
		-	Korrekturformel auf KOF:
			C-Krea [ml/min. x 1,73 m2] = 	C-Krea x 1,73
															——————
																  KOF

		—————————————————————————————————————————————————————————
		english
		—————————————————————————————————————————————————————————
		- according to DuBois, published 1916:  	body surface area [m²] = weight [kg] ^ 0.425 x height [cm] ^ 0.725 x 0.007184
		- according to Mosteller, published 1897: 	body surface area [m²] = ((weight [kg] x height [cm]) / 3600) ^ 0.5

		Correction of the body surface: when estimating the GFR according to Cockroft and Gault:
		- 	the values calculated individually for the patient must be on 1.73 m2 using a nomogram
			Body surface can be converted using the correction formula below
		- 	Correction formula on KOF:
			C-crea [ml / min. x 1.73 m2] = C-Krea x 1.73
															——————
																KOF

*/

/* 	weitere Formeln

		Quelle: https://www.uniklinikum-saarland.de/de/einrichtungen/kliniken_institute/zentrallabor/formeln_und_scores/gfr_kalkulator/

		(*) Formel nach Grubb et al (Cystatin-C-basiert)
			Formel: GFR[ml/min/1.73m2] = 130 * (S-CysC[mg/l])-1.069 * (Alter[Jahre])-0.117 - 7

		(*) Cockcroft-Gault-Formel (Kreatinin-basiert)
			Formel: GFR[ml/min] = (140 - Alter[Jahre]) * (S-Crea[mg/dl])-1 * (Gewicht[kg] * 72-1)
			Korrekturfaktor: für Frauen * 0.85

		(*) CKD-EPI-Formel (Kreatinin-basiert)
			Formel: GFR[ml/min/1.73m2] = 141 * min(S-Crea[mg/dl]/κ, 1)α * max(S-Crea[mg/dl]/κ, 1)-1.209 * 0.993Alter[Jahre] * 1.018 [if female] * 1.159 [if black]
			Korrekturfaktor: für Frauen * 1.018
			Korrekturfaktor: für Schwarze * 1.159
			κ is 0.7 for females and 0.9 for males,
			α is -0.329 for females and -0.411 for males,
			min(S-Crea/κ, 1) indicates the minimum of S-Crea/κ or 1,
			max(S-Crea/κ, 1) indicates the maximum of S-Crea/κ or 1.

		(*) CKD-EPI Formel (Cystatin-C-basiert)
			Formel: GFR[ml/min/1.73m2] = 133 * min(S-CysC[mg/l]/0.8, 1)-0.499 * max(S-CysC[mg/l]/0.8, 1)-1.328 * 0.996Alter[Jahre]
			Korrekturfaktor: für Frauen * 0.932
			min(S-CysC/0.8, 1) indicates the minimum of S-CysC/0.8 or 1,
			max(S-CysC/0.8, 1) indicates the maximum of S-CysC/0.8 or 1.

		(*) CKD-EPI Formel (Kreatinin- und Cystatin-C-basiert)
			Formel: GFR[ml/min/1.73m2] = 135 * min(S-Crea[mg/dl]/κ, 1)α * max(S-Crea[mg/dl]/κ, 1)-0.601 * min(S-CysC[mg/l]/0.8, 1)-0.375 * max(S-CysC[mg/l]/ 0.8, 1)-0.711 * 0.995Alter[Jahre]
			Korrekturfaktor: für Frauen * 0.969
			Korrekturfaktor: für Schwarze * 1.08
			κ is 0.7 for females and 0.9 for males,
			α is -0.248 for females and -0.207 for males,
			min(S-Crea/κ, 1) indicates the minimum of S-Crea/κ or 1,
			max(S-Crea/κ, 1) indicates the maximum of S-Crea/κ or 1,
			min(S-CysC/0.8, 1) indicates the minimum of S-CysC/0.8 or 1,
			max(S-CysC/0.8, 1) indicates the maximum of S-CysC/0.8 or 1.

		(*) Mosteller RD. Simplified calculation of body-surface area. NEJM 1987; 317: 1098-9
			Formel: KOF = Wurzel(Größe[cm] * Gewicht[kg] / 3600)

*/


class MDCalc {

	BMI(weightkg, heightmeters) {                                                                        	;-- i never need the calculation of a BMI, but maybe you!
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

	LabUnits(unitName, Labvalue, Unit:="", Precision:=2, lang:="") {                               	;-- laboratory units converter

		/* Description for LabUnitsConverter()

			――――――――――――――――――――――――――――――――――――――――――――――――
				⚕	laboratory unit converting function
			――――――――――――――――――――――――――――――――――――――――――――――――
				⊛ Parameters:
					⚬ unitName  	= name of laboratory unit (can be english or german)
					⚬ LabValue  	= value to convert
					⚬ Unit        	= unit to convert to (if there's only one choice, leave this parameter empty)
					⚬ Precision	= how many decimal places should the conversion value have (rounded)
					⚬ lang   		= use this for a formatted output of the unit in another language

				⊛ Conversion:
					⚬ Conventional units to SI units:  	C unit is multiplied with the Conversion-Factor (CF)
					⚬ SI units to Conventional units:  	SI unit is divided with CF

				⊛ this table is used to get conversion factors:
					https://accessmedicine.mhmedical.com/content.aspx?bookId=1069&sectionId=60775149

			 ――――――――――――――――――――――――――――――――――――――――――――――――
			   ⚕	CFactors Object Structure
			――――――――――――――――――――――――――――――――――――――――――――――――
				⊛ indexed objects with conversion data for laboratory units
					⚬ keys with fixed names are:
						- SI	= Standard International Unit
						- C	= Conventional Unit
						- CF	= Conversion Factor
					⚬ Key with names of laboratory values in German and English:
						- The values contain the abbreviations of the language [en for English and de for German]
						- There are also short names for the laboratory units avaible marked with an -s

		*/

		static CFactors := {1: {	"Kreatinin"     	: "de", "Krea"  	: "de-s",	"Creatinine"     	: "eng", "Crea"  	: "eng-s", "SI" : "µmol/l" , "C" : "mg/dl"	, "CF" : 88.4}
								,	 2: {	"Harnstoff"	    	: "de", "HST"   	: "de-s", "Urea"    			: "eng", "Urea"	: "eng-s", "SI" : "mmol/l", "C" : "mg/dl"	, "CF" : 0.357}
								,	 3: {	"Glucose"	    	: "de", "Glucose": "de-s", "Glucose" 	    	: "eng", "Gluc"   	: "eng-s", "SI" : "mmol/l", "C" : "mg/dl"	, "CF" : 0.0555}}
								,	 3: {	"Hämoglobin"	: "de", "Hb"     	: "de-s", "Haemoglobin"	: "eng", "Hb"     	: "eng-s", "SI" : "mmol/l", "C" : "g/dl"   	, "CF" : ‭1.6103059581320450885668276972625‬}}

		RegExMatch(Labvalue, "O)(?<value>[\d+\.\,])\s*(?<unit>[A-Za-zµΩω\/]+)", lab_)

		For CFIndex, oLab in CFactors
			If oLab.haskey(unitName)
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

	CKD_EPI(Scr, Units, height, weight, age, sex, african:=false, precision:=2) {	   	;-- eGFR (CKD-EPI-Formula)

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

	ntproBNPcorrection(NTproBNP, eGFR) {                                                            	;-- no code until today

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

	ConvertHbA1c(percent:="", mmolmol:="", precision:=2) {                                 	;-- converts HbA1c mmol/mol <--> mg/dl

		; used formula:   HbA1c [mmol/mol] = (HbA1c [%] - 2,15) x 10,929
		; Take the appropriate parameter so that the correct conversion formula is used,
		; so to convert a HbA1c with percentage to mg/dl write 'ConvertHbA1c(9.7, "")'.
		; mmol/mol to mg/dl then works like 'ConvertHbA1c("", 3.2).
		; The units for the values are logically omitted.
		; Link: https://www.bbraun.de/de/produkte-und-therapien/diabetes/omnitest/blutzuckerlangzeitwert-berechnen.html

		converted := percent ? (percent - 2.15) * 10.929 : (mmolmol / 10.929) + 2.15

		If precision
			return Round(converted, precision)

	return converted
	}

	Ganzoni(ActualHb, TargetHb, Weight) {                                                            	;-- Ganzoni Equation for Iron Deficiency Anemia

		; Calculates iron deficit for dosing iron.
		; Ganzoni gives us the formula to use with haemoglobin units of g/l or g/dl.
		; I'm working daily with mmol/l. And this function automatically converts the values to a suitable format if necessary.
		; You pass the values in convenient notation e.g. irondeficit := Ganzoni("7.7 mmol/l", "9.2 mmol/l", "64kg") or also irondeficit := Ganzoni("12.4 g/dl", "14.8 g/dl", "64kg").

	}




}






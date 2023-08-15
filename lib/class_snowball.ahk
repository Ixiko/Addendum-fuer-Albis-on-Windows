/* a porter stemmer - snowball - javascript wrapper V0.2a Ixiko 2023

	The Porter Stemmer is an algorithm for stem form reduction of English words. It was developed by Martin Porter in the 1980s
	and is one of the most widely used methods for automated stem form reduction.

	The Porter Stemmer uses a set of rules to reduce English words to their stem form. For example, it removes endings such as
	"-s," "-es," "-ed," and "-ing" to preserve the root form of the word. This can help reduce the number of different forms of words
	in a text, making it easier to understand the text.

	The Porter Stemmer used here is called "Snowball" because it is an extension of the algorithm called "Snowball". The Snowball
	extension allows the Porter Stemmer to be applied to different languages, including German, French, Italian, Spanish, and others.
	The name "Snowball" refers to the idea that the algorithm keeps growing like a snowball, spreading to more and more
	languages and application domains.

	The main purpose of the Porter Stemmer algorithm is stem form reduction of words to facilitate text analysis and processing.
	By applying Porter Stemmer, different forms of the same word can be reduced to a common base form, which can help
	reduce the number of different words in a text. This in turn can help improve the accuracy of text analysis and processing algorithms.

	Another application of the Porter Stemmer algorithm is to aid in the creation of text by using it in automated text generation,
	such as in automated report generation or chatbot design. By reducing words to their basic form, texts can be generated more
	efficiently because fewer different words need to be considered.

	found here:
		https://github.com/klaemo/snowball-german/blob/master/index.js

	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
	🌞  A big thank goes to 'just me' from the German Autohotkey forum, without whom I would not have got this library working.
	-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -

	 𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰
	  Brief explanation of the call and the function output
	 𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰

		• 	Because it is a class, multiple instances are possible. Each new instant needs only one parameter, which is the natural language
			you want to 'stem'. Use 'en' or 'de' to select the two built-in algorithms.

				sn_de := new snowball("de")
					or
				sn_en := new snowball("en")

		•	next step is already the stemmer algorithm. Pass a text string or a filepath and pass an object with options as the second parameter.
			Your passed text string is cleaned from superfluous characters before it is stemmed.
			At the moment there are only two options:

				noDD=1    	: 	Set to true if duplicate stems are to be removed
										if set to false you will get back the text in the original order of the words.
				minLen=0 	: 	stems every word no matter the length

	 𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰
		to run this example you need to have cJSON.ahk 0.4.1 (it's fast because its using machine code)
			https://www.autohotkey.com/boards/viewtopic.php?f=6&t=92320&hilit=cJSON&sid=3f262e01345dcdde7314de14357f1598
	 𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰𓋰

	Example:

		fname := "C:\temp\test.txt"
		toStem := FileOpen(fname, "r", "UTF-8").Read()

		sn_de := new snowball("de")
		rs := sn_de.stem(toStem, {"noDD":true, "minLen":3})
		MsgBox, % cJson.Dump(rs, 1)
		sn_de := ""

*/




class snowball {


	; code contains javascript code for stemming german and/or englisch words with a specific language dependend algorithm.
	static code := {"de" : "
	(LTRIM
	function stem (word) {

			word = word.replace(/([aeiouyäöü])u([aeiouyäöü])/g, '$1U$2');
			word = word.replace(/([aeiouyäöü])y([aeiouyäöü])/g, '$1Y$2');

			 word = word.replace(/ß/g, 'ss');

			var r1Index = word.search(/[aeiouyäöü][^aeiouyäöü]/);
			var r1 = '';
			if (r1Index != -1) {
				r1Index += 2;
				r1 = word.substring(r1Index);
			}

			var r2Index = -1;
			var r2 = '';

			if (r1Index != -1) {
				r2Index = r1.search(/[aeiouyäöü][^aeiouyäöü]/);
				if (r2Index != -1) {
					r2Index += 2;
					r2 = r1.substring(r2Index);
					r2Index += r1Index;
				} else {
					r2 = '';
				}
			}

			if (r1Index != -1 && r1Index < 3) {
				r1Index = 3;
				r1 = word.substring(r1Index);
			}

			var a1Index = word.search(/(em|ern|er)$/g);
			var b1Index = word.search(/(e|en|es)$/g);
			var c1Index = word.search(/([bdfghklmnrt]s)$/g);
			if (c1Index != -1) {
				c1Index++;
			}
			var index1 = 10000;
			var optionUsed1 = '';
			if (a1Index != -1 && a1Index < index1) {
				optionUsed1 = 'a';
				index1 = a1Index;
			}
			if (b1Index != -1 && b1Index < index1) {
				optionUsed1 = 'b';
				index1 = b1Index;
			}
			if (c1Index != -1 && c1Index < index1) {
				optionUsed1 = 'c';
				index1 = c1Index;
			}

			if (index1 != 10000 && r1Index != -1) {
				if (index1 >= r1Index) {
					word = word.substring(0, index1);
					if (optionUsed1 == 'b') {
						if (word.search(/niss$/) != -1) {
							word = word.substring(0, word.length -1);
						}
					}
				}
			}

			var a2Index = word.search(/(en|er|est)$/g);
			var b2Index = word.search(/(.{3}[bdfghklmnt]st)$/g);
			if (b2Index != -1) {
				b2Index += 4;
			}

			var index2 = 10000;
			var optionUsed2 = '';
			if (a2Index != -1 && a2Index < index2) {
				optionUsed2 = 'a';
				index2 = a2Index;
			}
			if (b2Index != -1 && b2Index < index2) {
				optionUsed2 = 'b';
				index2 = b2Index;
			}

			if (index2 != 10000 && r1Index != -1) {
				if (index2 >= r1Index) {
					word = word.substring(0, index2);
				}
			}

			var a3Index = word.search(/(end|ung)$/g);
			var b3Index = word.search(/[^e](ig|ik|isch)$/g);
			var c3Index = word.search(/(lich|heit)$/g);
			var d3Index = word.search(/(keit)$/g);
			if (b3Index != -1) {
				b3Index ++;
			}

			var index3 = 10000;
			var optionUsed3 = '';
			if (a3Index != -1 && a3Index < index3) {
				optionUsed3 = 'a';
				index3 = a3Index;
			}
			if (b3Index != -1 && b3Index < index3) {
				optionUsed3 = 'b';
				index3 = b3Index;
			}
			if (c3Index != -1 && c3Index < index3) {
				optionUsed3 = 'c';
				index3 = c3Index;
			}
			if (d3Index != -1 && d3Index < index3) {
				optionUsed3 = 'd';
				index3 = d3Index;
			}

			if (index3 != 10000 && r2Index != -1) {
				if (index3 >= r2Index) {
					word = word.substring(0, index3);
					var optionIndex = -1;
					// var optionSubsrt = '';
					if (optionUsed3 == 'a') {
						optionIndex = word.search(/[^e](ig)$/);
						if (optionIndex != -1) {
							optionIndex++;
							if (optionIndex >= r2Index) {
								word = word.substring(0, optionIndex);
							}
						}
					} else if (optionUsed3 == 'c') {
						optionIndex = word.search(/(er|en)$/);
						if (optionIndex != -1) {
							if (optionIndex >= r1Index) {
								word = word.substring(0, optionIndex);
							}
						}
					} else if (optionUsed3 == 'd') {
						optionIndex = word.search(/(lich|ig)$/);
						if (optionIndex != -1) {
							if (optionIndex >= r2Index) {
								word = word.substring(0, optionIndex);
							}
						}
					}
				}
			}

			word = word.replace(/U/g, 'u');
			word = word.replace(/Y/g, 'y');
			word = word.replace(/ä/g, 'a');
			word = word.replace(/ö/g, 'o');
			word = word.replace(/ü/g, 'u');

			return word;
		};
		)"

					, "en" : "
	(LTRIM
	var stemmer = (function(){
	  var step2list = {
		  'ational' : 'ate',
		  'tional' : 'tion',
		  'enci' : 'ence',
		  'anci' : 'ance',
		  'izer' : 'ize',
		  'bli' : 'ble',
		  'alli' : 'al',
		  'entli' : 'ent',
		  'eli' : 'e',
		  'ousli' : 'ous',
		  'ization' : 'ize',
		  'ation' : 'ate',
		  'ator' : 'ate',
		  'alism' : 'al',
		  'iveness' : 'ive',
		  'fulness' : 'ful',
		  'ousness' : 'ous',
		  'aliti' : 'al',
		  'iviti' : 'ive',
		  'biliti' : 'ble',
		  'logi' : 'log'
		},

		step3list = {
		  'icate' : 'ic',
		  'ative' : '',
		  'alize' : 'al',
		  'iciti' : 'ic',
		  'ical' : 'ic',
		  'ful' : '',
		  'ness' : ''
		},

		c = '[^aeiou]',          // consonant
		v = '[aeiouy]',          // vowel
		C = c + '[^aeiouy]*',    // consonant sequence
		V = v + '[aeiou]*',      // vowel sequence

		mgr0 = '^(' + C + ')?' + V + C,               // [C]VC... is m>0
		meq1 = '^(' + C + ')?' + V + C + '(' + V + ')?$',  // [C]VC[V] is m=1
		mgr1 = '^(' + C + ')?' + V + C + V + C,       // [C]VCVC... is m>1
		s_v = '^(' + C + ')?' + v;                   // vowel in stem

	  function dummyDebug() {}

	  function realDebug() {
		console.log(Array.prototype.slice.call(arguments).join(' '));
	  }

	  return function (w, debug) {
		var
		  stem,
		  suffix,
		  firstch,
		  re,
		  re2,
		  re3,
		  re4,
		  debugFunction,
		  origword = w;

		if (debug) {
		  debugFunction = realDebug;
		} else {
		  debugFunction = dummyDebug;
		}

		if (w.length < 3) { return w; }

		firstch = w.substr(0,1);
		if (firstch == 'y') {
		  w = firstch.toUpperCase() + w.substr(1);
		}

		// Step 1a
		re = /^(.+?)(ss|i)es$/;
		re2 = /^(.+?)([^s])s$/;

		if (re.test(w)) {
		  w = w.replace(re,'$1$2');
		  debugFunction('1a',re, w);

		} else if (re2.test(w)) {
		  w = w.replace(re2,'$1$2');
		  debugFunction('1a',re2, w);
		}

		// Step 1b
		re = /^(.+?)eed$/;
		re2 = /^(.+?)(ed|ing)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  re = new RegExp(mgr0);
		  if (re.test(fp[1])) {
			re = /.$/;
			w = w.replace(re,'');
			debugFunction('1b',re, w);
		  }
		} else if (re2.test(w)) {
		  var fp = re2.exec(w);
		  stem = fp[1];
		  re2 = new RegExp(s_v);
		  if (re2.test(stem)) {
			w = stem;
			debugFunction('1b', re2, w);

			re2 = /(at|bl|iz)$/;
			re3 = new RegExp('([^aeiouylsz])\\1$');
			re4 = new RegExp('^' + C + v + '[^aeiouwxy]$');

			if (re2.test(w)) {
			  w = w + 'e';
			  debugFunction('1b', re2, w);

			} else if (re3.test(w)) {
			  re = /.$/;
			  w = w.replace(re,'');
			  debugFunction('1b', re3, w);

			} else if (re4.test(w)) {
			  w = w + 'e';
			  debugFunction('1b', re4, w);
			}
		  }
		}

		// Step 1c
		re = new RegExp('^(.*' + v + '.*)y$');
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  w = stem + 'i';
		  debugFunction('1c', re, w);
		}

		// Step 2
		re = /^(.+?)(ational|tional|enci|anci|izer|bli|alli|entli|eli|ousli|ization|ation|ator|alism|iveness|fulness|ousness|aliti|iviti|biliti|logi)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  suffix = fp[2];
		  re = new RegExp(mgr0);
		  if (re.test(stem)) {
			w = stem + step2list[suffix];
			debugFunction('2', re, w);
		  }
		}

		// Step 3
		re = /^(.+?)(icate|ative|alize|iciti|ical|ful|ness)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  suffix = fp[2];
		  re = new RegExp(mgr0);
		  if (re.test(stem)) {
			w = stem + step3list[suffix];
			debugFunction('3', re, w);
		  }
		}

		// Step 4
		re = /^(.+?)(al|ance|ence|er|ic|able|ible|ant|ement|ment|ent|ou|ism|ate|iti|ous|ive|ize)$/;
		re2 = /^(.+?)(s|t)(ion)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  re = new RegExp(mgr1);
		  if (re.test(stem)) {
			w = stem;
			debugFunction('4', re, w);
		  }
		} else if (re2.test(w)) {
		  var fp = re2.exec(w);
		  stem = fp[1] + fp[2];
		  re2 = new RegExp(mgr1);
		  if (re2.test(stem)) {
			w = stem;
			debugFunction('4', re2, w);
		  }
		}

		// Step 5
		re = /^(.+?)e$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  re = new RegExp(mgr1);
		  re2 = new RegExp(meq1);
		  re3 = new RegExp('^' + C + v + '[^aeiouwxy]$');
		  if (re.test(stem) || (re2.test(stem) && !(re3.test(stem)))) {
			w = stem;
			debugFunction('5', re, re2, re3, w);
		  }
		}

		re = /ll$/;
		re2 = new RegExp(mgr1);
		if (re.test(w) && re2.test(w)) {
		  re = /.$/;
		  w = w.replace(re,'');
		  debugFunction('5', re, re2, w);
		}

		// and turn initial Y back to y
		if (firstch == 'y') {
		  w = firstch.toLowerCase() + w.substr(1);
		}


		return w;
	  }
	})();
	)"}

    __New(language:="de") 	{	; opens an instance

            this.lang 	:= !language ~= "^\s*\w\w\s*$" && !this.code[language] ? "de" : language
			this.js     	:= new ActiveScript("JScript")
			this.js.Exec(this.code[this.lang])

    }

	__Delete()                     	{

			If IsObject(this.js)
				this.js := ""

	}

	stem(words, opt:="")         	{   ; give a text to receive an array of snowballed words

		/* Description

			words:	 The function detects if one or more words were passed. If more than one word is passed, the stemmed words are
						 returned in a linear array, otherwise only the stemmed word is returned as string.
			opt: 		 an Object with following keys:
						 noDD=1 	: Set to true if duplicate stems are to be removed
						 minLen=0 	: stems every word no matter the length

		*/

		words := Trim(words)
		If (StrSplit(words, A_Space).Count() = 1)
			return this.js.stem(words)
		else if FileExist(words)
			words := FileOpen(words, "r", "UTF-8").Read()

		this.stems := Array()
		If !IsObject(opt)
			opt:= Object()
		opt.minLen := opt.minLen ? opt.minLen : 0

		words := RegExReplace(words, "[^\pL\d\-]", " ")
		words := RegExReplace(words, "\s\d+", " ")
		words := RegExReplace(words, "\d+\s", " ")
		words := RegExReplace(words, "\s{2,}", " ")

	; pushes every stemmed word to an array, leave blank stemmed (however this happens)
		For each, word in StrSplit(words, A_Space)
			If (StrLen(word)>=opt.minLen && !word ~= "\d")
				If (stemmed := this.js.stem(word))
					this.stems.Push(stemmed "|" word)

	 ; removes doublettes sort if remove option is set to true
		If (opt.NoDD) {
			For each, stem in this.stems
				stems .= (StrLen(stems)>0?"`n":"") stem
			Sort, stems, U
			this.stems := StrSplit(stems, "`n")
		}

	return this.stems
	}

}


 /* ActiveScript for AutoHotkey v1.1
 *
 *  Provides an interface to Active Scripting languages like VBScript and JScript,
 *  without relying on Microsoft's ScriptControl, which is not available to 64-bit
 *  programs.
 *
 *  License: Use, modify and redistribute without limitation, but at your own risk.
 *
 */
class ActiveScript extends ActiveScript._base{
    __New(Language)    {
        if this._script := ComObjCreate(Language, ActiveScript.IID)
            this._scriptParse := ComObjQuery(this._script, ActiveScript.IID_Parse)
        if !this._scriptParse
            throw Exception("Invalid language", -1, Language)
        this._site := new ActiveScriptSite(this)
        this._SetScriptSite(this._site.ptr)
        this._InitNew()
        this._objects := {}
        this.Error := ""
        this._dsp := this._GetScriptDispatch()  ; Must be done last.
        try
            if this.ScriptEngine() = "JScript"
                this.SetJScript58()
    }

    SetJScript58()    {
        static IID_IActiveScriptProperty := "{4954E0D0-FBC7-11D1-8410-006008C3FBFC}"
        if !prop := ComObjQuery(this._script, IID_IActiveScriptProperty)
            return false
        VarSetCapacity(var, 24, 0), NumPut(2, NumPut(3, var, "short") + 6)
        hr := DllCall(NumGet(NumGet(prop+0)+4*A_PtrSize), "ptr", prop, "uint", 0x4000
            , "ptr", 0, "ptr", &var), ObjRelease(prop)
        return hr >= 0
    }

    Eval(Code)    {
        pvar := NumGet(ComObjValue(arr:=ComObjArray(0xC,1)) + 8+A_PtrSize)
        this._ParseScriptText(Code, 0x20, pvar)  ; SCRIPTTEXT_ISEXPRESSION := 0x20
        return arr[0]
    }

    Exec(Code)    {
        this._ParseScriptText(Code, 0x42, 0)  ; SCRIPTTEXT_ISVISIBLE := 2, SCRIPTTEXT_ISPERSISTENT := 0x40
        this._SetScriptState(2)  ; SCRIPTSTATE_CONNECTED := 2
    }

    AddObject(Name, DispObj, AddMembers := false)    {
        static a, supports_dispatch ; Test for built-in IDispatch support.
            := a := ((a:=ComObjArray(0xC,1))[0]:=[42]) && a[0][1]=42
        if IsObject(DispObj) && !(supports_dispatch || ComObjType(DispObj))
            throw Exception("Adding a non-COM object requires AutoHotkey v1.1.17+", -1)
        this._objects[Name] := DispObj
        this._AddNamedItem(Name, AddMembers ? 8 : 2)  ; SCRIPTITEM_ISVISIBLE := 2, SCRIPTITEM_GLOBALMEMBERS := 8
    }

    _GetObjectUnk(Name)    {
        return !IsObject(dsp := this._objects[Name]) ? dsp  ; Pointer
            : ComObjValue(dsp) ? ComObjValue(dsp)  ; ComObject
            : &dsp  ; AutoHotkey object
    }

    class _base
    {
        __Call(Method, Params*)        {
            if ObjHasKey(this, "_dsp")
                try
                    return (this._dsp)[Method](Params*)
                catch e
                    throw Exception(e.Message, -1, e.Extra)
        }

        __Get(Property, Params*)        {
            if ObjHasKey(this, "_dsp")
                try
                    return (this._dsp)[Property, Params*]
                catch e
                    throw Exception(e.Message, -1, e.Extra)
        }

        __Set(Property, Params*)        {
            if ObjHasKey(this, "_dsp")            {
                Value := Params.Pop()
                try
                    return (this._dsp)[Property, Params*] := Value
                catch e
                    throw Exception(e.Message, -1, e.Extra)
            }
        }
    }

    _SetScriptSite(Site)    {
        hr := DllCall(NumGet(NumGet((p:=this._script)+0)+3*A_PtrSize), "ptr", p, "ptr", Site)
        if (hr < 0)
            this._HRFail(hr, "IActiveScript::SetScriptSite")
    }

    _SetScriptState(State)    {
        hr := DllCall(NumGet(NumGet((p:=this._script)+0)+5*A_PtrSize), "ptr", p, "int", State)
        if (hr < 0)
            this._HRFail(hr, "IActiveScript::SetScriptState")
    }

    _AddNamedItem(Name, Flags)    {
        hr := DllCall(NumGet(NumGet((p:=this._script)+0)+8*A_PtrSize), "ptr", p, "wstr", Name, "uint", Flags)
        if (hr < 0)
            this._HRFail(hr, "IActiveScript::AddNamedItem")
    }

    _GetScriptDispatch()    {
        hr := DllCall(NumGet(NumGet((p:=this._script)+0)+10*A_PtrSize), "ptr", p, "ptr", 0, "ptr*", pdsp:=0)
        if (hr < 0)
            this._HRFail(hr, "IActiveScript::GetScriptDispatch")
        return ComObject(9, pdsp, 1)
    }

    _InitNew()    {
        hr := DllCall(NumGet(NumGet((p:=this._scriptParse)+0)+3*A_PtrSize), "ptr", p)
        if (hr < 0)
            this._HRFail(hr, "IActiveScriptParse::InitNew")
    }

    _ParseScriptText(Code, Flags, pvarResult)    {
        VarSetCapacity(excp, 8 * A_PtrSize, 0)
        hr := DllCall(NumGet(NumGet((p:=this._scriptParse)+0)+5*A_PtrSize), "ptr", p
            , "wstr", Code, "ptr", 0, "ptr", 0, "ptr", 0, "uptr", 0, "uint", 1
            , "uint", Flags, "ptr", pvarResult, "ptr", 0)
        if (hr < 0)
            this._HRFail(hr, "IActiveScriptParse::ParseScriptText")
    }

    _HRFail(hr, what)    {
        if e := this.Error
        {
            this.Error := ""
            throw Exception("`nError code:`t" this._HRFormat(e.HRESULT)
                . "`nSource:`t`t" e.Source "`nDescription:`t" e.Description
                . "`nLine:`t`t" e.Line "`nColumn:`t`t" e.Column
                . "`nLine text:`t`t" e.LineText, -3)
        }
        throw Exception(what " failed with code " this._HRFormat(hr), -2)
    }

    _HRFormat(hr)    {
        return Format("0x{1:X}", hr & 0xFFFFFFFF)
    }

    _OnScriptError(err) ; IActiveScriptError err
    {
        VarSetCapacity(excp, 8 * A_PtrSize, 0)
        DllCall(NumGet(NumGet(err+0)+3*A_PtrSize), "ptr", err, "ptr", &excp) ; GetExceptionInfo
        DllCall(NumGet(NumGet(err+0)+4*A_PtrSize), "ptr", err, "uint*", srcctx, "uint*", srcline, "int*", srccol) ; GetSourcePosition
        DllCall(NumGet(NumGet(err+0)+5*A_PtrSize), "ptr", err, "ptr*", pbstrcode:=0) ; GetSourceLineText
        code := StrGet(pbstrcode, "UTF-16"), DllCall("OleAut32\SysFreeString", "ptr", pbstrcode)
        if fn := NumGet(excp, 6 * A_PtrSize) ; pfnDeferredFillIn
            DllCall(fn, "ptr", &excp)
        wcode := NumGet(excp, 0, "ushort")
        hr := wcode ? 0x80040200 + wcode : NumGet(excp, 7 * A_PtrSize, "uint")
        this.Error := {HRESULT: hr, Line: srcline, Column: srccol, LineText: code}
        static Infos := "Source,Description,HelpFile"
        Loop Parse, % Infos, `,
            if pbstr := NumGet(excp, A_Index * A_PtrSize)
                this.Error[A_LoopField] := StrGet(pbstr, "UTF-16"), DllCall("OleAut32\SysFreeString", "ptr", pbstr)
        return 0x80004001 ; E_NOTIMPL (let Exec/Eval get a fail result)
    }

    __Delete()    {
        if this._script
        {
            DllCall(NumGet(NumGet((p:=this._script)+0)+7*A_PtrSize), "ptr", p)  ; Close
            ObjRelease(this._script)
        }
        if this._scriptParse
            ObjRelease(this._scriptParse)
    }

    static IID := "{BB1A2AE1-A4F9-11cf-8F20-00805F2CD064}"
    static IID_Parse := A_PtrSize=8 ? "{C7EF7658-E1EE-480E-97EA-D52CB4D76D17}" : "{BB1A2AE2-A4F9-11cf-8F20-00805F2CD064}"
}
class ActiveScriptSite{
    __New(Script)
    {
        ObjSetCapacity(this, "_site", 3 * A_PtrSize)
        NumPut(&Script
        , NumPut(ActiveScriptSite._vftable("_vft_w", "31122", 0x100)
        , NumPut(ActiveScriptSite._vftable("_vft", "31125232211", 0)
            , this.ptr := ObjGetAddress(this, "_site"))))
    }

    _vftable(Name, PrmCounts, EIBase)
    {
        if p := ObjGetAddress(this, Name)
            return p
        ObjSetCapacity(this, Name, StrLen(PrmCounts) * A_PtrSize)
        p := ObjGetAddress(this, Name)
        Loop Parse, % PrmCounts
        {
            cb := RegisterCallback("_ActiveScriptSite", "F", A_LoopField, A_Index + EIBase)
            NumPut(cb, p + (A_Index-1) * A_PtrSize)
        }
        return p
    }
}
_ActiveScriptSite(this, a1:=0, a2:=0, a3:=0, a4:=0, a5:=0){
    Method := A_EventInfo & 0xFF
    if A_EventInfo >= 0x100  ; IActiveScriptSiteWindow
    {
        if Method = 4  ; GetWindow
        {
            NumPut(0, a1+0) ; *phwnd := 0
            return 0 ; S_OK
        }
        if Method = 5  ; EnableModeless
        {
            return 0 ; S_OK
        }
        this -= A_PtrSize     ; Cast to IActiveScriptSite
    }
    ;else: IActiveScriptSite
    if Method = 1  ; QueryInterface
    {
        iid := _AS_GUIDToString(a1)
        if (iid = "{00000000-0000-0000-C000-000000000046}"  ; IUnknown
         || iid = "{DB01A1E3-A42B-11cf-8F20-00805F2CD064}") ; IActiveScriptSite
        {
            NumPut(this, a2+0)
            return 0 ; S_OK
        }
        if (iid = "{D10F6761-83E9-11cf-8F20-00805F2CD064}") ; IActiveScriptSiteWindow
        {
            NumPut(this + A_PtrSize, a2+0)
            return 0 ; S_OK
        }
        NumPut(0, a2+0)
        return 0x80004002 ; E_NOINTERFACE
    }
    if Method = 5  ; GetItemInfo
    {
        a1 := StrGet(a1, "UTF-16")
        , (a3 && NumPut(0, a3+0))  ; *ppiunkItem := NULL
        , (a4 && NumPut(0, a4+0))  ; *ppti := NULL
        if (a2 & 1) ; SCRIPTINFO_IUNKNOWN
        {
            if !(unk := Object(NumGet(this + A_PtrSize*2))._GetObjectUnk(a1))
                return 0x8002802B ; TYPE_E_ELEMENTNOTFOUND
            ObjAddRef(unk), NumPut(unk, a3+0)
        }
        return 0 ; S_OK
    }
    if Method = 9  ; OnScriptError
        return Object(NumGet(this + A_PtrSize*2))._OnScriptError(a1)

    ; AddRef and Release don't do anything because we want to avoid circular references.
    ; The site and IActiveScript are both released when the AHK script releases its last
    ; reference to the ActiveScript object.

    ; All of the other methods don't require implementations.
    return 0x80004001 ; E_NOTIMPL
}
_AS_GUIDToString(pGUID){
    VarSetCapacity(String, 38*2)
    DllCall("ole32\StringFromGUID2", "ptr", pGUID, "str", String, "int", 39)
    return String
}
/* JsRT for AutoHotkey v1.1
 *
 *  Utilizes the JavaScript engine that comes with IE11.
 *
 *  License: Use, modify and redistribute without limitation, but at your own risk.
 *
 */
class JsRT extends ActiveScript._base {
    __New()    {
        throw Exception("This class is abstract. Use JsRT.IE or JSRT.Edge instead.", -1)
    }

    class IE extends JsRT    {

        __New()        {
            if !this._hmod := DllCall("LoadLibrary", "str", "jscript9", "ptr")
                throw Exception("Failed to load jscript9.dll", -1)
            if DllCall("jscript9\JsCreateRuntime", "int", 0, "int", -1
                , "ptr", 0, "ptr*", runtime:=0) != 0
                throw Exception("Failed to initialize JsRT", -1)
            DllCall("jscript9\JsCreateContext", "ptr", runtime, "ptr", 0, "ptr*", context:=0)
            this._Initialize("jscript9", runtime, context)
        }
    }

    class Edge extends JsRT    {
        __New()        {
            if !this._hmod := DllCall("LoadLibrary", "str", "chakra", "ptr")
                throw Exception("Failed to load chakra.dll", -1)
            if DllCall("chakra\JsCreateRuntime", "int", 0
                , "ptr", 0, "ptr*", runtime:=0) != 0
                throw Exception("Failed to initialize JsRT", -1)
            DllCall("chakra\JsCreateContext", "ptr", runtime, "ptr*", context:=0)
            this._Initialize("chakra", runtime, context)
        }

        ProjectWinRTNamespace(namespace)
        {
            return DllCall("chakra\JsProjectWinRTNamespace", "wstr", namespace)
        }
    }

    _Initialize(dll, runtime, context)    {
        this._dll := dll
        this._runtime := runtime
        this._context := context
        DllCall(dll "\JsSetCurrentContext", "ptr", context)
        DllCall(dll "\JsGetGlobalObject", "ptr*", globalObject:=0)
        this._dsp := this._JsToVt(globalObject)
    }

    __Delete()    {
        this._dsp := ""
        if dll := this._dll
        {
            DllCall(dll "\JsSetCurrentContext", "ptr", 0)
            DllCall(dll "\JsDisposeRuntime", "ptr", this._runtime)
        }
        DllCall("FreeLibrary", "ptr", this._hmod)
    }

    _JsToVt(valref)    {
        VarSetCapacity(variant, 24, 0)
        DllCall(this._dll "\JsValueToVariant", "ptr", valref, "ptr", &variant)
        ref := ComObject(0x400C, &variant), val := ref[], ref[] := 0
        return val
    }

    _ToJs(val)    {
        VarSetCapacity(variant, 24, 0)
        ref := ComObject(0x400C, &variant) ; VT_BYREF|VT_VARIANT
        ref[] := val
        DllCall(this._dll "\JsVariantToValue", "ptr", &variant, "ptr*", valref:=0)
        ref[] := 0
        return valref
    }

    _JsEval(code)    {
        e := DllCall(this._dll "\JsRunScript", "wstr", code, "uptr", 0, "wstr", "source.js"
            , "ptr*", result:=0)
        if e
        {
            if DllCall(this._dll "\JsGetAndClearException", "ptr*", excp:=0) = 0
                throw this._JsToVt(excp)
            throw Exception("JsRT error", -2, format("0x{:X}", e))
        }
        return result
    }

    Exec(code)    {
        this._JsEval(code)
    }

    Eval(code)    {
        return this._JsToVt(this._JsEval(code))
    }

    AddObject(name, obj, addMembers := false)    {
        if addMembers
            throw Exception("AddMembers=true is not supported", -1)
        this._dsp[name] := obj
    }
}

ComDispatch0(this){
    static vtable := ComDispatch0_VTable()
    static id_to_name := [], name_to_id := []

    obj := {}, obj.SetCapacity("_", 2*A_PtrSize)
    obj_mem := obj.GetAddress("_")
    ,NumPut(&obj, NumPut(vtable, obj_mem+0))
    ,ObjAddRef(&obj)

    ,obj.name_to_id := name_to_id
    ,obj.id_to_name := id_to_name
    ,obj.pointer    := obj_mem
    ,obj.this       := this

    return ComObject(9, obj_mem, 1)
}
ComDispatch0_VTable(){
    static vtable
    if !VarSetCapacity(vtable) {
        VarSetCapacity(vtable, 7 * A_PtrSize)
        for idx, cnt in [3,1,1,2,4,6,9]
            NumPut(RegisterCallback("ComDispatch0_", "", cnt, (idx-1)), vtable, (idx-1)*A_PtrSize)
    }
    return &vtable
}
ComDispatch0_Unwrap(ComObject){
    static vtable := ComDispatch0_VTable()
    return ComObjType(ComObject) = 9 && NumGet(ComObjValue(ComObject)) == vtable
        ?  Object(NumGet(ComObjValue(ComObject)+A_PtrSize)).this
        :  ComObject
}
ComDispatch0_(this_, prm0 := 0, prm1 := 0, prm2 := 0, prm3 := 0, prm4 := 0, prm5 := 0, prm6 := 0, prm7 := 0, prm8 := 0){
    Critical

    ; Get our object
    this := Object(this_ptr := NumGet(this_+A_PtrSize))

    goto cd0_%A_EventInfo%

cd0_0: ; IUnknown::QueryInterface
    ; Beware of the hack code!
     iid1 := NumGet(prm0+0, "Int64")
    ,iid2 := NumGet(prm0+8, "Int64")
    if (iid2 = 0x46000000000000C0) && (!iid1 || iid1 = 0x20400)
    {
        NumPut(this_, prm1+0), ObjAddRef(this_ptr)
        return 0
    }
    else
    {
        NumPut(0, prm1+0)
        return 0x80004002 ; E_NOINTERFACE
    }

cd0_1: ; IUnknown::AddRef
    return ObjAddRef(this_ptr)

cd0_2: ; IUnknown::Release
    return ObjRelease(this_ptr)

cd0_3: ; IDispatch::GetTypeInfoCount
    NumPut(0, prm0+0, "UInt")
    return 0

cd0_4: ; IDispatch::GetTypeInfo
    return 0x80004001 ; E_NOTIMPL

    ; All the funny 0xFF... masking in the code below is because
    ; of the x64 calling convention. For parameters whose size is
    ; < 64 bits, the upper bits are garbage. So we clear them out.

cd0_5: ; IDispatch::GetIDsOfNames
    status := 0, name := StrGet(NumGet(prm1+0), "UTF-16")
    if !(dispid := this.name_to_id[name])
    {
        dispid := this.id_to_name.Push(name)
        this.name_to_id[name] := dispid
    }
    NumPut(dispid, prm4 + 0, "int")
    Loop, % (prm2 & 0xFFFFFFFF) - 1
        NumPut(-1, prm4 + A_Index*4, "int") ; DISPID_UNKNOWN: -1
        , status := 0x80020006 ; DISP_E_UNKNOWNNAME
    return status

cd0_6: ; IDispatch::Invoke
    prm3 &= 0xFFFF
    name := this.id_to_name[prm0 &= 0xFFFFFFFF]
    if (name = "" && prm0 != 0)
        return 0x80020003 ; DISP_E_MEMBERNOTFOUND
     paramarray := NumGet(prm4+0)
    ,nparams := NumGet(prm4+2*A_PtrSize, "UInt")
    ,params := []
    if !nparams
        goto cd0_call
    else if NumGet(prm4+2*A_PtrSize+4, "UInt") != ((prm3 & 12) ? 1 : 0)
        return 0x80020007 ; DISP_E_NONAMEDARGS

    ; Make a SAFEARRAY out of the raw VARIANT array
    static pad := A_PtrSize = 8 ? 4 : 0, sizeof_SAFEARRAY := 20+pad+A_PtrSize, sizeof_VARIANT := 8+2*A_PtrSize
     VarSetCapacity(SAhdr, sizeof_SAFEARRAY, 0)
    ,NumPut(1, SAhdr, 0, "UShort")
    ,NumPut(0x0812, SAhdr, 2, "UShort") ; FADF_STATIC | FADF_FIXEDSIZE | FADF_VARIANT
    ,NumPut(sizeof_VARIANT, SAhdr, 4, "UInt")
    ,NumPut(paramarray, SAhdr, 12+pad)
    ,NumPut(nparams, SAhdr, 12+pad+A_PtrSize, "UInt")
    ,params_safearray := ComObject(0x200C, &SAhdr)

    ; Copy the parameters to a regular AutoHotkey array
    Loop % nparams
        ObjPush(params, params_safearray[idx := nparams - A_Index])
    Loop % nparams
    {
        a := params[A_Index]
        while ComObjType(a) = 0x400C ; VT_BYREF | VT_VARIANT
            NumPut(1, NumPut(ComObjValue(a), SAhdr, 12+pad), "UInt")
           ,a := params_safearray[0]
        params[A_Index] := ComDispatch0_Unwrap(a)
    }
    if (prm3 & 12)
        value := ObjPop(params)

cd0_call:
    ; Prepare a SAFEARRAY of VARIANT for converting the return value.
    ret := NumGet(ComObjValue(retarr:=ComObjArray(0xC,1)) + 8+A_PtrSize)

    ; Call the function
    try
    {
        if ((prm3 & 3) = 1)  ; DISPATCH_METHOD and not DISPATCH_PROPERTYGET
        {
            if (prm0 = 0) ; DISPID_VALUE: "call" the object itself
                name := this.this, retarr[0] := %name%(params*)
            else
                retarr[0] := (this.this)[name](params*)
        }
        else  ; Property
        {
            if (prm0 != 0) ; != DISPID_VALUE
                ObjInsertAt(params, 1, name)
            retarr[0] := (prm3 & 12)
                ? ((this.this)[params*] := value)
                : ((this.this)[params*])
        }
    }
    catch e
    {
        ; Clear caller-supplied VARIANT.
        if prm5
        Loop % sizeof_VARIANT // 8
            NumPut(0, prm5+8*(A_Index-1), "Int64")
        ; Fill exception info
        if prm6
        {
            NumPut(0, prm6+0) ; wCode, wReserved, padding
            NumPut(cd0_BSTR(e.what), prm6+A_PtrSize) ; bstrSource
            NumPut(cd0_BSTR(e.message), prm6+2*A_PtrSize) ; bstrDescription
            NumPut(cd0_BSTR(e.file ":" e.line), prm6+3*A_PtrSize) ; bstrHelpFile
            NumPut(0, prm6+4*A_PtrSize) ; dwHelpContext, padding
            NumPut(0, prm6+5*A_PtrSize) ; pvReserved
            NumPut(0, prm6+6*A_PtrSize) ; pfnDeferredFillIn
            NumPut(0x80020009, prm6+7*A_PtrSize, "UInt") ; scode
        }
        return 0x80020009 ; DISP_E_EXCEPTION
    }
    if prm5
    ; MOVE the converted return value to the caller-supplied VARIANT.
    Loop % sizeof_VARIANT // 8
    {
        idx := 8*(A_Index-1)
        NumPut(NumGet(ret+idx, "Int64"), prm5+idx, "Int64")
        NumPut(0, ret+idx, "Int64")
    }

    return 0
}
cd0_BSTR(ByRef a){
    return DllCall("oleaut32\SysAllocString", "wstr", a, "ptr")
}


#Include %A_ScriptDir%\..\..\lib\SciTEOutput.ahk
#Include %A_ScriptDir%\..\..\lib\class_cJson.ahk


	/* 	static en := "
	(
	var stemmer = (function(){
	  var step2list = {
		  "ational" : "ate",
		  "tional" : "tion",
		  "enci" : "ence",
		  "anci" : "ance",
		  "izer" : "ize",
		  "bli" : "ble",
		  "alli" : "al",
		  "entli" : "ent",
		  "eli" : "e",
		  "ousli" : "ous",
		  "ization" : "ize",
		  "ation" : "ate",
		  "ator" : "ate",
		  "alism" : "al",
		  "iveness" : "ive",
		  "fulness" : "ful",
		  "ousness" : "ous",
		  "aliti" : "al",
		  "iviti" : "ive",
		  "biliti" : "ble",
		  "logi" : "log"
		},

		step3list = {
		  "icate" : "ic",
		  "ative" : "",
		  "alize" : "al",
		  "iciti" : "ic",
		  "ical" : "ic",
		  "ful" : "",
		  "ness" : ""
		},

		c = "[^aeiou]",          // consonant
		v = "[aeiouy]",          // vowel
		C = c + "[^aeiouy]*",    // consonant sequence
		V = v + "[aeiou]*",      // vowel sequence

		mgr0 = "^(" + C + ")?" + V + C,               // [C]VC... is m>0
		meq1 = "^(" + C + ")?" + V + C + "(" + V + ")?$",  // [C]VC[V] is m=1
		mgr1 = "^(" + C + ")?" + V + C + V + C,       // [C]VCVC... is m>1
		s_v = "^(" + C + ")?" + v;                   // vowel in stem

	  function dummyDebug() {}

	  function realDebug() {
		console.log(Array.prototype.slice.call(arguments).join(' '));
	  }

	  return function (w, debug) {
		var
		  stem,
		  suffix,
		  firstch,
		  re,
		  re2,
		  re3,
		  re4,
		  debugFunction,
		  origword = w;

		if (debug) {
		  debugFunction = realDebug;
		} else {
		  debugFunction = dummyDebug;
		}

		if (w.length < 3) { return w; }

		firstch = w.substr(0,1);
		if (firstch == "y") {
		  w = firstch.toUpperCase() + w.substr(1);
		}

		// Step 1a
		re = /^(.+?)(ss|i)es$/;
		re2 = /^(.+?)([^s])s$/;

		if (re.test(w)) {
		  w = w.replace(re,"$1$2");
		  debugFunction('1a',re, w);

		} else if (re2.test(w)) {
		  w = w.replace(re2,"$1$2");
		  debugFunction('1a',re2, w);
		}

		// Step 1b
		re = /^(.+?)eed$/;
		re2 = /^(.+?)(ed|ing)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  re = new RegExp(mgr0);
		  if (re.test(fp[1])) {
			re = /.$/;
			w = w.replace(re,"");
			debugFunction('1b',re, w);
		  }
		} else if (re2.test(w)) {
		  var fp = re2.exec(w);
		  stem = fp[1];
		  re2 = new RegExp(s_v);
		  if (re2.test(stem)) {
			w = stem;
			debugFunction('1b', re2, w);

			re2 = /(at|bl|iz)$/;
			re3 = new RegExp("([^aeiouylsz])\\1$");
			re4 = new RegExp("^" + C + v + "[^aeiouwxy]$");

			if (re2.test(w)) {
			  w = w + "e";
			  debugFunction('1b', re2, w);

			} else if (re3.test(w)) {
			  re = /.$/;
			  w = w.replace(re,"");
			  debugFunction('1b', re3, w);

			} else if (re4.test(w)) {
			  w = w + "e";
			  debugFunction('1b', re4, w);
			}
		  }
		}

		// Step 1c
		re = new RegExp("^(.*" + v + ".*)y$");
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  w = stem + "i";
		  debugFunction('1c', re, w);
		}

		// Step 2
		re = /^(.+?)(ational|tional|enci|anci|izer|bli|alli|entli|eli|ousli|ization|ation|ator|alism|iveness|fulness|ousness|aliti|iviti|biliti|logi)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  suffix = fp[2];
		  re = new RegExp(mgr0);
		  if (re.test(stem)) {
			w = stem + step2list[suffix];
			debugFunction('2', re, w);
		  }
		}

		// Step 3
		re = /^(.+?)(icate|ative|alize|iciti|ical|ful|ness)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  suffix = fp[2];
		  re = new RegExp(mgr0);
		  if (re.test(stem)) {
			w = stem + step3list[suffix];
			debugFunction('3', re, w);
		  }
		}

		// Step 4
		re = /^(.+?)(al|ance|ence|er|ic|able|ible|ant|ement|ment|ent|ou|ism|ate|iti|ous|ive|ize)$/;
		re2 = /^(.+?)(s|t)(ion)$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  re = new RegExp(mgr1);
		  if (re.test(stem)) {
			w = stem;
			debugFunction('4', re, w);
		  }
		} else if (re2.test(w)) {
		  var fp = re2.exec(w);
		  stem = fp[1] + fp[2];
		  re2 = new RegExp(mgr1);
		  if (re2.test(stem)) {
			w = stem;
			debugFunction('4', re2, w);
		  }
		}

		// Step 5
		re = /^(.+?)e$/;
		if (re.test(w)) {
		  var fp = re.exec(w);
		  stem = fp[1];
		  re = new RegExp(mgr1);
		  re2 = new RegExp(meq1);
		  re3 = new RegExp("^" + C + v + "[^aeiouwxy]$");
		  if (re.test(stem) || (re2.test(stem) && !(re3.test(stem)))) {
			w = stem;
			debugFunction('5', re, re2, re3, w);
		  }
		}

		re = /ll$/;
		re2 = new RegExp(mgr1);
		if (re.test(w) && re2.test(w)) {
		  re = /.$/;
		  w = w.replace(re,"");
		  debugFunction('5', re, re2, w);
		}

		// and turn initial Y back to y
		if (firstch == "y") {
		  w = firstch.toLowerCase() + w.substr(1);
		}


		return w;
	  }
	})();
	)"

 */

	/* 	static code_stemmer_en_de := "
					(
							function stemmer_de(word) {
								word = word.replace(/([aeiouyäöü])u([aeiouyäöü])/g, '$1U$2');
								word = word.replace(/([aeiouyäöü])y([aeiouyäöü])/g, '$1Y$2');
								word = word.replace(/ß/g, 'ss');
								var r1Index = word.search(/[aeiouyäöü][^aeiouyäöü]/);
								var r1 = '';
								if (r1Index != -1) {
									r1Index += 2;
									r1 = word.substring(r1Index);
								}

								var r2Index = -1;
								var r2 = '';

								if (r1Index != -1) {
									r2Index = r1.search(/[aeiouyäöü][^aeiouyäöü]/);
									if (r2Index != -1) {
										r2Index += 2;
										r2 = r1.substring(r2Index);
										r2Index += r1Index;
									} else {
										r2 = '';
									}
								}

								if (r1Index != -1 && r1Index < 3) {
									r1Index = 3;
									r1 = word.substring(r1Index);
								}

								var a1Index = word.search(/(em|ern|er)$/g);
								var b1Index = word.search(/(e|en|es)$/g);
								var c1Index = word.search(/([bdfghklmnrt]s)$/g);
								if (c1Index != -1) {
									c1Index++;
								}
								var index1 = 10000;
								var optionUsed1 = '';
								if (a1Index != -1 && a1Index < index1) {
									optionUsed1 = 'a';
									index1 = a1Index;
								}
								if (b1Index != -1 && b1Index < index1) {
									optionUsed1 = 'b';
									index1 = b1Index;
								}
								if (c1Index != -1 && c1Index < index1) {
									optionUsed1 = 'c';
									index1 = c1Index;
								}

								if (index1 != 10000 && r1Index != -1) {
									if (index1 >= r1Index) {
										word = word.substring(0, index1);
										if (optionUsed1 == 'b') {
											if (word.search(/niss$/) != -1) {
												word = word.substring(0, word.length -1);
											}
										}
									}
								}

								var a2Index = word.search(/(en|er|est)$/g);
								var b2Index = word.search(/(.{3}[bdfghklmnt]st)$/g);
								if (b2Index != -1) {
									b2Index += 4;
								}

								var index2 = 10000;
								var optionUsed2 = '';
								if (a2Index != -1 && a2Index < index2) {
									optionUsed2 = 'a';
									index2 = a2Index;
								}
								if (b2Index != -1 && b2Index < index2) {
									optionUsed2 = 'b';
									index2 = b2Index;
								}

								if (index2 != 10000 && r1Index != -1) {
									if (index2 >= r1Index) {
										word = word.substring(0, index2);
									}
								}

								var a3Index = word.search(/(end|ung)$/g);
								var b3Index = word.search(/[^e](ig|ik|isch)$/g);
								var c3Index = word.search(/(lich|heit)$/g);
								var d3Index = word.search(/(keit)$/g);
								if (b3Index != -1) {
									b3Index ++;
								}

								var index3 = 10000;
								var optionUsed3 = '';
								if (a3Index != -1 && a3Index < index3) {
									optionUsed3 = 'a';
									index3 = a3Index;
								}
								if (b3Index != -1 && b3Index < index3) {
									optionUsed3 = 'b';
									index3 = b3Index;
								}
								if (c3Index != -1 && c3Index < index3) {
									optionUsed3 = 'c';
									index3 = c3Index;
								}
								if (d3Index != -1 && d3Index < index3) {
									optionUsed3 = 'd';
									index3 = d3Index;
								}

								if (index3 != 10000 && r2Index != -1) {
									if (index3 >= r2Index) {
										word = word.substring(0, index3);
										var optionIndex = -1;
										// var optionSubsrt = '';
										if (optionUsed3 == 'a') {
											optionIndex = word.search(/[^e](ig)$/);
											if (optionIndex != -1) {
												optionIndex++;
												if (optionIndex >= r2Index) {
													word = word.substring(0, optionIndex);
												}
											}
										} else if (optionUsed3 == 'c') {
											optionIndex = word.search(/(er|en)$/);
											if (optionIndex != -1) {
												if (optionIndex >= r1Index) {
													word = word.substring(0, optionIndex);
												}
											}
										} else if (optionUsed3 == 'd') {
											optionIndex = word.search(/(lich|ig)$/);
											if (optionIndex != -1) {
												if (optionIndex >= r2Index) {
													word = word.substring(0, optionIndex);
												}
											}
										}
									}
								}

								word = word.replace(/U/g, 'u');
								word = word.replace(/Y/g, 'y');
								word = word.replace(/ä/g, 'a');
								word = word.replace(/ö/g, 'o');
								word = word.replace(/ü/g, 'u');

								return word;
							}

							var stemmer_en = (function () {
								"use strict";

								const regExps = {}
								regExps.c = '[^aeiou]';
								regExps.v = '[aeiouy]';
								regExps.C = regExps.c + '[^aeiouy]*';
								regExps.V = regExps.v + '[aeiou]*';
								regExps.M_gr_0 = new RegExp('^(' + regExps.C + ')?' + regExps.V + regExps.C);
								regExps.M_eq_1 = new RegExp('^(' + regExps.C + ')?' + regExps.V + regExps.C + '(' + regExps.V + ')?$');
								regExps.M_gr_1 = new RegExp('^(' + regExps.C + ')?' + regExps.V + regExps.C + regExps.V + regExps.C);
								regExps.v_in_stem = new RegExp('^(' + regExps.C + ')?' + regExps.v);
								regExps.nonstd_S = /^(.+?)(ss|i)es$/;
								regExps.std_S = /^(.+?)([^s])s$/;
								regExps.E = /^(.+?)e$/;
								regExps.LL = /ll$/;
								regExps.EED = /^(.+?)eed$/;
								regExps.ED_or_ING = /^(.+?)(ed|ing)$/;
								regExps.Y = /^(.+?)y$/;
								regExps.nonstd_gp1 = /^(.+?)(ational|tional|enci|anci|izer|bli|alli|entli|eli|ousli|ization|ation|ator|alism|iveness|fulness|ousness|aliti|iviti|biliti|logi)$/;
								regExps.nonstd_gp2 = /^(.+?)(icate|ative|alize|iciti|ical|ful|ness)$/;
								regExps.nonstd_gp3 = /^(.+?)(al|ance|ence|er|ic|able|ible|ant|ement|ment|ent|ou|ism|ate|iti|ous|ive|ize)$/;
								regExps.S_or_T_with_ION = /^(.+?)(s|t)(ion)$/;
								regExps.has_C_and_v_but_doesnt_end_with_AEIOUWXY = new RegExp('^' + regExps.C + regExps.v + '[^aeiouwxy]$');

								const suffixList = {
									group1: {
										'ational': 'ate',
										'tional': 'tion',
										'enci': 'ence',
										'anci': 'ance',
										'izer': 'ize',
										'bli': 'ble',
										'alli': 'al',
										'entli': 'ent',
										'eli': 'e',
										'ousli': 'ous',
										'ization': 'ize',
										'ation': 'ate',
										'ator': 'ate',
										'alism': 'al',
										'iveness': 'ive',
										'fulness': 'ful',
										'ousness': 'ous',
										'aliti': 'al',
										'iviti': 'ive',
										'biliti': 'ble',
										'logi': 'log'
									},
									group2: {
										'icate': 'ic',
										'ative': '',
										'alize': 'al',
										'iciti': 'ic',
										'ical': 'ic',
										'ful': '',
										'ness': ''
									}
								};

								return function (w) {
									if (w.length < 3)
										return w;
									if (w.charAt(0) === "y")
										w = w.charAt(0).toUpperCase() + w.substr(1);
									if (regExps.nonstd_S.test(w))
										w = w.replace(regExps.nonstd_S, '$1$2');
									else if (regExps.std_S.test(w))
										w = w.replace(regExps.std_S, '$1$2');
									if (regExps.EED.test(w)) {
										var stem = (regExps.EED.exec(w) || [])[1];
										if (regExps.M_gr_0.test(w))
											w = w.substr(0, w.length - 1);
									}
									else if (regExps.ED_or_ING.test(w)) {
										var stem = (regExps.ED_or_ING.exec(w) || [])[1];
										if (regExps.v_in_stem.test(stem)) {
											w = stem;
											if (/(at|bl|iz)$/.test(w))
												w = w + 'e';
											else if (new RegExp('([^aeiouylsz])\\1$').test(w))
												w = w.substr(0, w.length - 1);
											else if (new RegExp('^' + regExps.C + regExps.v + '[^aeiouwxy]$').test(w))
												w = w + 'e';
										}
									}
									if (regExps.Y.test(w)) {
										var stem = (regExps.Y.exec(w) || [])[1];
										if (regExps.v_in_stem.test(stem))
											w = stem + 'i';
									}
									if (regExps.nonstd_gp1.test(w)) {
										var result = regExps.nonstd_gp1.exec(w) || [];
										var stem = result[1];
										var suffix = result[2];
										if (regExps.M_gr_0.test(stem))
											w = stem + suffixList.group1[suffix];
									}
									if (regExps.nonstd_gp2.test(w)) {
										var result = regExps.nonstd_gp2.exec(w) || [];
										var stem = result[1];
										var suffix = result[2];
										if (regExps.M_gr_0.test(stem))
											w = stem + suffixList.group2[suffix];
									}
									if (regExps.nonstd_gp3.test(w)) {
										var result = regExps.nonstd_gp3.exec(w) || [];
										var stem = result[1];
										if (regExps.M_gr_1.test(stem))
											w = stem;
									}
									else if (regExps.S_or_T_with_ION.test(w)) {
										var result = regExps.S_or_T_with_ION.exec(w) || [];
										var stem = result[1] + result[2];
										if (regExps.M_gr_1.test(stem))
											w = stem;
									}
									if (regExps.E.test(w)) {
										var result = regExps.E.exec(w) || [];
										var stem = result[1];
										if (regExps.M_gr_1.test(stem) || (regExps.M_eq_1.test(stem) && !(regExps.has_C_and_v_but_doesnt_end_with_AEIOUWXY.test(stem))))
											w = stem;
									}
									if (regExps.LL.test(w) && regExps.M_gr_1.test(w))
										w = w.substr(0, w.length - 1);
									if (w.charAt(0) === "Y")
										w = w.charAt(0).toLowerCase() + w.substr(1);
									return w;
								};
							})()

							var stemmer = {
								english: stemmer_en,
								german: stemmer_de,
							}

							try {
								if (module && module.exports) module.exports = stemmer;
							} catch (e) {

							}
							)"
				*/



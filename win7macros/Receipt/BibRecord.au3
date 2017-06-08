#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: BibRecord
 Script set: Receipt Cataloging (Receipt)

 Script Function:
	This script goes from the order record into the bib record. It copies the
	data from the bib record and proceeds to check it against the listed
	criteria for receipt cataloging processing. If the item does not meet the
	listed criteria, the script will automatically start the Create item - Non DLC
	script.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(Order record open)

 Last revised: 6/03/11
			   Added new BEST call number designations and oversize variable

			   6/29/10
			   Updated #Include to include TS custom function library
			   Updated window controls for JRE version

 PLEASE NOTE - This script uses a custom UDF library called TSCustomFunction.
			   The script will not run properly (if launched from .au3 file)
			   if that file is not included in the include folder in the
			   AutoIt directory.

 Copyright (C): 2009 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------

;######### INCLUDES AND OPTIONS #########
#Include <TSCustomFunction.au3>
AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("SendKeyDelay", 50)
TraySetIcon(@DesktopDir & "\Receipt\Images\black B.ico")

;######### DECLARE VARIABLES #########
Dim $REVIEW, $T85, $T86, $T87, $T88, $T89
dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $row1_MASTER, $acqtypecheck, $row2_MASTER, $row3_MASTER
Dim $BCODE3
Dim $CAT_DATE
Dim $LANG
Dim $BCODE1
Dim $COUNTRY
Dim $SKIP
Dim $BCODE2
Dim $BIB_LOC_1
Dim $CAT_DATE_R, $TODAY_DATE
Dim $decide
Dim $MARC_LEAD_A, $MARC_LEAD, $MARC_LEAD_L
Dim $040_A, $040, $040_DLC, $040_DNLM
Dim $042_A, $042, $042_PCC
Dim $050_A, $050_A_N, $050, $050_Z, $050_Z_A, $050_Z_N, $050_Z_L, $050_PZ
Dim $245_2_A, $245_2
Dim $300_A, $300, $300_C_A, $300_C_A_1,$300_C, $FOLIO, $300_E, $300_E_A, $300_E_1
Dim $050_1, $050_2, $050_2_S, $050_2_K, $050_345
Dim $REF
Dim $IS_IMC
Dim $260_A
Dim $008_D_A, $008_D
Dim $050_D
Dim $260, $260_C
Dim $490_A, $830_A, $440_A
Dim $SF_NAME, $FUND, $multi, $multi_A
dim $BIB_LOC_L2, $BIB_LOC_L1
Dim $300_C_X, $260_C1, $260_L
Dim $LOCATION
Dim $FUND
Dim $BIB_LOC
Dim $SF_590, $SF_7XX, $SF_NAME
Dim $LOCAL
Dim $008_C
Dim $dean
Dim $FUND_A
Dim $multi_loc
Dim $multi_loc_a
Dim $OVERSIZE

;################################ MAIN ROUTINE #################################
;start loading variables from order record script
$LOCATION = _LoadVar("$LOCATION")
$LOCAL = _LoadVar("$LOCAL")
$FUND = _LoadVar("$FUND")
$REF = _LoadVar("$REF")
$T85 = _LoadVar("$T85")
$T86 = _LoadVar("$T86")
$T87 = _LoadVar("$T87")
$T88 = _LoadVar("$T88")
$dean = _LoadVar("$dean")
$multi_loc = _LoadVar("$multi_loc")


$REVIEW = 1
$T89 = ""
$FOLIO = 0
$OVERSIZE = 0

;focus on order record screen, go to bib record
If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Millennium record", "Please open the order record.")
	Exit
EndIf
Sleep(0100)
_SendEx("!g")
Sleep(0100)
_SendEx("e")
WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")

;select all and copy bib record information, parse data into array
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;~ ##### start fixed data setting #####
;row 1 includes lang/cat date/ bcode3
$row1_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 1)
$BCODE3 = _ArrayPop($row1_MASTER)
$CAT_DATE = _ArrayPop($row1_MASTER)
$LANG = _ArrayPop($row1_MASTER)
$LANG = StringTrimLeft($LANG, 5)
$CAT_DATE = StringTrimLeft($CAT_DATE, 8)
$BCODE3 = StringTrimLeft($BCODE3, 6)

;row 2 includes skip, bcode1, country
$row2_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 2)
$COUNTRY = _ArrayPop($row2_MASTER)
$BCODE1 = _ArrayPop($row2_MASTER)
$SKIP = _ArrayPop($row2_MASTER)
$SKIP = StringTrimLeft($SKIP, 5)
$BCODE1 = StringTrimLeft($BCODE1, 6)
$COUNTRY = StringTrimLeft($COUNTRY, 7)

;row 3 includes bib location and bcode2
$row3_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 3)
$BCODE2 = _ArrayPop($row3_MASTER)
$BIB_LOC_1= _ArrayPop($row3_MASTER)
$BCODE2 = StringTrimLeft($BCODE2, 6)
$BIB_LOC_1 = StringTrimLeft($BIB_LOC_1, 9)
$BIB_LOC_1 = StringLeft($BIB_LOC_1, 5)
$BIB_LOC_1 = StringStripWS($BIB_LOC_1, 8)
;~ ##### end fixed data setting #####

;~ ##### start CAT_DATE check #####
$CAT_DATE_R = StringInStr($CAT_DATE, "-  -")
If $CAT_DATE_R = 0 Then
	$decide = MsgBox(4, "Already received", "This item was already received on " & $CAT_DATE & ". Do you want to override the CDATE?")
	Switch $decide
		Case 6
			$TODAY_DATE = "t"
		Case 7
			$TODAY_DATE = "r"
	EndSwitch
EndIf
;~ ##### end CAT_DATE check #####


;~ ##### start MARC LEADER 8 check #####
$MARC_LEAD = _arrayItemString($BIB_ARRAY_MASTER, "MARC Leader	")
$MARC_LEAD =  StringTrimLeft($MARC_LEAD, 12)
$MARC_LEAD_L = StringMid($MARC_LEAD, 18, 1)
If $MARC_LEAD_L = 8 Then
	$T85 = $T85 & "This record has an 8 in the 30th position of the MARC LEADER." & @CR
	$T89 = "SEND TO COPY CATALOGING"
	$REVIEW = 0
EndIf
;~ ##### end MARC LEADER 8 check #####

;~ ##### start DLC/DLC check #####
$040_A = _ArraySearch($BIB_ARRAY_MASTER, "MARC	040	 	 	", 0, 0, 0, 1)
If $040_A > -1 Then
	$040 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $040_A, $040_A)
	$040 =  StringTrimLeft($040, 13)
	$040_DLC = StringInStr($040, "DLC|cDLC")
	$040_DNLM = StringInStr($040, "DNLM/DLC|cDLC")
	$042_A = _ArraySearch($BIB_ARRAY_MASTER, "MARC	042	 	 	", 0, 0, 0, 1)
		If $042_A > -1 Then
			$042 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $042_A, $042_A)
			$042 =  StringTrimLeft($042, 13)
			$042_PCC = StringInStr($042, "pcc")
			If $040_DLC = 0 And $040_DNLM = 0 Then
				If $042_PCC = 0 Then
					$T85 = $T85 & "This is a non-DLC/PCC record" & @CR
					$T89 = "SEND TO COPY CATALOGING"
					$REVIEW = 0
				EndIf
			EndIf
		Else
			If $040_DLC = 0 AND $040_DNLM = 0 Then
				$T85 = $T85 & "This is a non-DLC record" & @CR
				$T89 = "SEND TO COPY CATALOGING"
				$REVIEW = 0
			EndIf
		EndIf
Else
		$T85 =  $T85 & "This record does not have an 040 field" & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
EndIf
;~ ##### end DLC/DLC check #####

;~ ##### start 050 check #####
$050_A = _ArrayFindAll($BIB_ARRAY_MASTER, "CALL #	050", 0, 0, 0, 1)
$050_A_N = UBound($050_A)
If $050_A_N > 0 Then
	If $050_A_N > 1 Then
		$T85 = $T85 & "This record has multiple 050 fields" & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
	Else
		$050_A = _ArraySearch($BIB_ARRAY_MASTER, "CALL #	050", 0, 0, 0, 1)
		If $050_A > -1 Then
			$050 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $050_A, $050_A)
			$050 =  StringTrimLeft($050, 15)
			$050_Z = StringInStr($050, "Z", 1, 1, 1, 1)
			If $050_Z <> 0 Then
				$050_Z_A = StringSplit($050, "|")
				$050_Z = _ArrayToString($050_Z_A, @TAB, 1, 1)
				$050_Z_N = StringMid($050_Z, 2, 4)
				If $050_Z_N > 1200 Then
					$T85 = $T85 & "The call number is larger than Z1200.  " & @CR
					$T89 = "SEND TO COPY CATALOGING"
					$REVIEW = 0
				EndIf
			EndIf
			$050_PZ = StringInStr($050, "PZ", 1, 1, 1, 1)
			If $050_PZ <>0 AND $LOCATION = "kngli" Then
				$T85 = $T85 & "This record has a PZ in the 050 field in the bibl record, but KNGLI in the order record." & @CR
				$T89 = "SEND to Deb Sayers"
				$REVIEW = 0
			EndIf
		EndIf
	EndIf
Else
	$T85 = $T85 & "This record does not have an 050 field" & @CR
	$T89 = "SEND TO COPY CATALOGING"
	$REVIEW = 0
EndIf
;~ ##### end 050 check #####

;~ ##### start SKIP check #####
$245_2 = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
$245_2 =  StringMid($245_2, 13, 1)
If $SKIP <> $245_2 Then
	MsgBox(64, "Skip does not match 245 second indicator", "Skip does not match the 245 second indicator." & @CR & "Please adjust and click ok to continue.")
EndIf
;~ ##### end SKIP check #####


;~ ##### start 300 check #####
$300_A = _ArraySearch($BIB_ARRAY_MASTER, "DESCRIPT.	300", 0, 0, 0, 1)
If $300_A > -1 Then
	$300 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $300_A, $300_A)
	$300 =  StringTrimLeft($300, 18)
	If $300 = "p.|ccm" Then
		$T85 = $T85 & "This record is a CIP." & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
	Else
		$300_C_A = StringSplit($300, "|")
		$300_C_A_1 = _ArraySearch($300_C_A, "cm", 0, 0, 0, 1)
		$300_C = _ArrayToString($300_C_A, @TAB, $300_C_A_1, $300_C_A_1)
		$300_C_X = StringInStr($300_C, "x")
		If $300_C_X > 0 Then
			$300_C_X = $300_C_X + 2
			$300_C_X = StringMid($300_C, $300_C_X, 2)
		EndIf
		$300_C = StringMid($300_C, 2, 2)
		If $300_C > 30 Or $300_C_X > 30 Then
			$FOLIO = 1
		ElseIf ($300_C > 26 AND $300_C <= 30) OR ($300_C_X > 26 AND $300_C_X <= 30) Then
			$OVERSIZE = 1
		ElseIf $300_C = "cm" Then
			$T85 = $T85 & "This record does not have a measurement in 300 |c." & @CR
			$T89 = "SEND TO COPY CATALOGING"
			$REVIEW = 0
		EndIf
		$300_E = StringInStr($300, "|e")
		If $300_E > 0 Then
			If $BIB_LOC = "mul" Then
				$T85 = $T85 & "This Music item has an accompanying material." & @CR
				$T89 = "SEND TO Music Librarian"
				$REVIEW = 0
				$300_E = ""
			Else
			$300_E_A = $300_C_A_1 + 1
			$300_E = _ArrayToString($300_C_A, @TAB, $300_E_A, $300_E_A)
			$300_E = StringTrimLeft($300_E, 1)

			$300_E_1 = StringInStr($300_E, "(")
			$300_E = StringLeft($300_E, $300_E_1)
			$300_E = StringTrimRight($300_E, 1)
			EndIf
		Else
			$300_E = "none"
		EndIf
	EndIf
Else
	$T85 = $T85 & "This record does not have a 300 field" & @CR
	$T89 = "SEND TO COPY CATALOGING"
	$300_E = "none"
	$REVIEW = 0
EndIf
;~ ##### end 300 check #####

;~ ##### start acq call number/location assignment and check #####
If $BIB_LOC_1 = "acq" Then
	$050_1 = StringLeft($050, 1)
	$050_2 = StringMid($050, 2, 1)
	Switch $050_1
		Case "M"
			$LOCATION = "muli"
		Case "N"
			$LOCATION = "aali"
		Case "Q", "R", "S"
			If $OVERSIZE = 1 Then
				$LOCATION = "scov"
			Else
				$LOCATION = "scli"
			endif
		Case "T"
			If $050_2 = "R" Then
				$LOCATION = "aali"
			Else
				If $OVERSIZE = 1 Then
					$LOCATION = "scov"
				Else
					$LOCATION = "scli"
				endif
			EndIf
		Case "G"
			$050_2_S = StringRegExp($050_2, "[A-CE4-9]")
			$050_2_K = StringRegExp($050_2, "[FGHKLMNRTV12]")
			If $050_2_S > 0 Then
				If $OVERSIZE = 1 Then
					$LOCATION = "scov"
				Else
					$LOCATION = "scli"
				endif
			ElseIf $050_2_K > 0 Then
				$LOCATION = "kngli"
			ElseIf $050_2 = 3 Then
				$050_345 = StringMid($050, 3, 3)
				If $050_345 >= 160 Then
					If $OVERSIZE = 1 Then
						$LOCATION = "scov"
					Else
						$LOCATION = "scli"
					endif
				Else
					$LOCATION = "kngli"
				EndIf
			EndIf
		Case "B"
			$050_2_S = StringRegExp($050_2, "[F]")
			$050_2_K = StringRegExp($050_2, "[^F]")
			If $050_2_S > 0 Then
				$050_345 = StringMid($050, 3, 4)
				If $050_345 <= 1156 Then
					If $OVERSIZE = 1 Then
						$LOCATION = "scov"
					Else
						$LOCATION = "scli"
					endif
				Else
					$LOCATION = "kngli"
				EndIf
			ElseIf $050_2_K > 0 Then
				$LOCATION = "kngli"
			EndIf
		Case "H"
			$050_2_S = StringRegExp($050_2, "[B-J]")
			$050_2_K = StringRegExp($050_2, "[^B-J]")
			If $050_2_S > 0 AND $FOLIO = 0 Then
				$LOCATION = "scli"
			ElseIf $050_2_S > 0 AND $OVERSIZE = 1 Then
				$LOCATION = "scov"
			ElseIf $050_2_K > 0 OR $FOLIO = 1 Then
				$LOCATION = "kngli"
			EndIf
		Case Else
			$LOCATION = "kngli"
	EndSwitch
	_StoreVar("$LOCATION")
EndIf

Sleep(1500)

If $REF = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "r")
	_StoreVar("$LOCATION")
EndIf
;~ ##### end check #####

;~ ##### start multi check #####
If $BIB_LOC_1 = "multi" Then
	$T85 = $T85 & "This record has a location of multi" & @CR
	$T89 = "SEND TO COPY CATALOGING OR ADDED COPIES"
	$REVIEW = 0
EndIf
;~ ##### end multi check #####

;~ ##### start imju check #####
If $LOCATION = "imju" Then
	$decide = InputBox("Choose IMC Juv location", "Choose which type of juv book this item is. Type in the corresponding letter." & @CR & "A: Juv (Fic and Nonfic)" & @CR & "B: Easy" & @CR & "C: Big Book"  & @CR & "D: IMC Reader" & @CR & "E: IMC LangMat", "", "", 200, 210)
	$decide = StringUpper($decide)
	Switch $decide
		Case "A"
			$LOCATION = "imju"
		Case "B"
			$LOCATION = "imje"
		Case "C"
			$LOCATION = "imjb"
		Case "D"
			$LOCATION = "imrdr"
		Case "E"
			$LOCATION = "imlng"
	EndSwitch
			_StoreVar("$LOCATION")
EndIf
;~ ##### end imju check #####

;~ ##### start folio/oversize check #####
$IS_IMC = StringInStr($LOCATION, "im")
If $FOLIO = 1 AND $IS_IMC = 0 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
	_StoreVar("$LOCATION")
EndIf
If $OVERSIZE = 1 AND $LOCATION = "scli" Then
	$LOCATION = StringReplace($LOCATION, "li", "ov")
	_StoreVar("$LOCATION")
EndIf
;~ ##### end folio check #####

;~ ##### start dean material location check #####
;this sets the default bib location to kngli for dean's office materials
If $dean = 1 Then
	$LOCATION = "kngli"
EndIf
;~ ##### end dean material location check #####

;##### RUN BIB LOCATION SCRIPT #####
Run(@DesktopDir & "\Receipt\Bib - check ILOC.exe")
	While ProcessExists("Bib - check ILOC.exe")
		Sleep(0400)
	WEnd
$BIB_LOC = _LoadVar("$BIB_LOC")
$REF = _LoadVar("$REF")
Sleep(0300)
;##### END RUN BIB LOCATION SCRIPT #####

;~ ##### start bcode1 check #####
$BCODE1 = StringLeft($BCODE1, 1)
If $BCODE1 = "s" OR $BCODE1 = "b" Then
	$T85 = $T85 & "This is a SERIAL" & @CR
	$T89 = "SEND TO SERIALS"
	$REVIEW = 0
EndIf
;~ ##### end bcode1 check #####

;~ ##### start bcode2 check #####
$BCODE2 = StringLeft($BCODE2, 1)
If $BCODE1 = "c" Then
	$T85 = $T85 & "This record is for a musical score" & @CR
	$T89 = "SEND to Music Cataloging"
	$REVIEW = 0
EndIf
;~ ##### end bcode2 check #####

;~ ##### start bcode3 check #####
;NOT LOOPED YET
If $LOCAL = -1 Then
	If $BCODE3 <> "-" Then
		MsgBox(64, "BCODE3 is not set to '-'", "BCODE3 is not set to '-'." & @CR & "Please correct and click ok to continue.")
	EndIf
ElseIf $dean = 1 Then
	If $BCODE3 <> "s" Then
		MsgBox(64, "BCODE3 is not set to 's'", "BCODE3 is not set to 's'." & @CR & "Please correct and click ok to continue.")
	EndIf
EndIf
;~ ##### end bcode3 check #####

;~ ##### start date check #####
$260_A = _ArraySearch($BIB_ARRAY_MASTER, "IMPRINT	260", 0, 0, 0, 1)
If $MARC_LEAD_L <> 8 AND $050_A_N = 1 AND $260_A > -1 Then
	;search for 008 date
	$008_D_A = _ArraySearch($BIB_ARRAY_MASTER, "MARC	008", 0, 0, 0, 1)
	If $008_D_A > -1 Then
		$008_D = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $008_D_A, $008_D_A)
		$008_D =  StringMid($008_D, 21, 4)
		$008_D = StringStripWS($008_D, 8)
		;check for conference entry
		$008_C = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $008_D_A, $008_D_A)
		$008_C =  StringMid($008_C, 43, 1)
		$008_C = StringStripWS($008_C, 8)
	EndIf
	;search for 050 date
	$050_D = StringRight($050, 4)
	$050_D = StringStripWS($050_D, 8)
	;search for 260 date
	$260 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $260_A, $260_A)
	$260_L = StringLen($260)
	$260_C1 = StringInStr($260, "|c")
	$260_C = StringMid($260, $260_C1)
	$260_C = StringTrimLeft($260_C, 2)
	$260_C = StringRegExpReplace($260_C, "[][c?]", "")
	$260_C = StringStripWS($260_C, 8)
	If $260_C <> $008_D OR $008_D <> $050_D OR $260_C <> $050_D Then
		;conference check
		If $008_C <> 1 Then
			$T85 = $T85 & "The 008, 050, and 260 |c dates do not match." & @CR
			$T89 = "SEND TO COPY CATALOGING"
			$REVIEW = 0
		EndIf
	EndIf
EndIf
;~ ##### end date check #####

;~ 	##### start series check #####
$440_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	440", 0, 0, 0, 1)
$490_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	490", 0, 0, 0, 1)
$830_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	830", 0, 0, 0, 1)
If $REVIEW = 1 Then
	If $490_A > -1 AND $830_A = -1 Then
		$T85 = $T85 & "The 490 field does not have a corresponding 830 field." & @CR
		$T89 = "SEND TO COPY CATALOGING"
		$REVIEW = 0
	EndIf
	If $440_A > -1 OR $830_A > -1 Then
		$decide = MsgBox(4, "Check the series.", "Check the series." & @CR & "Should this item be classed separately?")
		Switch $decide
			Case 6
			Case 7
				$T88 = $T88 & "Classed together." & @CR
				$T89 = "SEND TO SERIALS"
				$REVIEW = 0
				$LOCATION = "none"
		EndSwitch
		WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	EndIf
EndIf
;~ 	##### end series check #####

;START SAVING VARIABLES FOR ITEM SCRIPTS
_StoreVar("$300_E")
_StoreVar("$LOCATION")
_StoreVar("$BIB_LOC_1")

Dim $slash, $4
;~ ##### START PROCESSING RECORD #####
If $REVIEW = 1 Then
	;~ ##### start Special Fund check #####
	#cs
		Fun happy fund parsing workaround!
		It turns out that the new fund formatting for 2009B played havoc with
		parsing out fund codes for single fund purchases. Instead of having a
		space between the fund code and the fund name, they are lumped together
		in one solid line. I have a workaround that checks for the first digit
		of the fund line to determine if it's a special fund code. This in part
		works because all but one special fund start with "4" and none have
		a slash in them.

		As far as I know, the multi fund parsing has not run into the same
		issue from the 2009B upgrade, so I have left the original parsing
		intact.
	#ce
	$slash = StringInStr($FUND, "/")
	$4 = StringLeft($FUND, 1)
	If $slash = 0 And $4 = 4 Then
		$FUND_A = StringRegExp($FUND, "(4\D{1,3})", 1)
		$FUND = _ArrayPop($FUND_A)
	EndIf
	$FUND = StringStripWS($FUND, 8)
	_StoreVar("$FUND")
	If $FUND <> "multi" then
		Run(@DesktopDir & "\Receipt\Special Fund List.exe")
			While ProcessExists("Special Fund List.exe")
				Sleep(0400)
			WEnd
		$SF_NAME = _LoadVar("$SF_NAME")
		If $SF_NAME <> "none" Then
			$SF_590 = _LoadVar("$SF_590")
			$SF_7XX = _LoadVar("$SF_7XX")
			WinActive("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}n")
			Sleep(0100)
			_SendEx("590{TAB}")
			sleep(0100)
			Switch $BIB_LOC
				Case "kngl"
					_SendEx("King copy{SPACE}")
				Case "scl"
					_SendEx("Science copy{SPACE}")
				Case "aal"
					_SendEx("ArtArch copy{SPACE}")
				Case "imc"
					_SendEx("IMC copy{SPACE}")
				Case "mul"
					_SendEx("Music copy{SPACE}")
			EndSwitch
			_SendEx($SF_590)
			Sleep(0100)
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}b")
			Sleep(0100)
			_SendEx($SF_7XX)
		EndIf
	ElseIf $FUND = "multi" Then
		$multi = _LoadVar("$multi")
		$multi_A = StringSplit($multi, ",")
		Do
			$FUND = _ArrayPop($multi_A)
			$FUND = StringStripWS($FUND, 8)
				Run(@DesktopDir & "\Receipt\Special Fund List.exe")
					While ProcessExists("Special Fund List.exe")
						Sleep(0400)
					WEnd
				$SF_NAME = _LoadVar("$SF_NAME")
			If $SF_NAME <> "none" Then
				$SF_590 = _LoadVar("$SF_590")
				$SF_7XX = _LoadVar("$SF_7XX")
				$SF_7XX = WinActive("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
				_SendEx("^{END}")
				Sleep(0100)
				_SendEx("{ENTER}n")
				Sleep(0100)
				_SendEx("590{TAB}")
				sleep(0100)
				Switch $BIB_LOC
					Case "kngl"
						_SendEx("King copy{SPACE}")
					Case "scl"
						_SendEx("Science copy{SPACE}")
					Case "aal"
						_SendEx("ArtArch copy{SPACE}")
					Case "imc"
						_SendEx("IMC copy{SPACE}")
					Case "mul"
						_SendEx("Music copy{SPACE}")
				EndSwitch
				_SendEx($SF_590)
				Sleep(0100)
				_SendEx("^{END}")
				Sleep(0100)
				_SendEx("{ENTER}b")
				Sleep(0100)
				_SendEx($SF_7XX)
			EndIf
		Until $multi_A = ""
	EndIf
	;~ ##### end Special Fund check #####
	SLEEP(0400)
	If $BIB_LOC = "multi" Then ;usually if item is for reference
		$BIB_LOC_L1 = _LoadVar("$BIB_LOC_L1")
		$BIB_LOC_L2 = _LoadVar("$BIB_LOC_L2")
		_SendEx("^{HOME}")
		Sleep(0100)
		_SendEx("{DOWN 2}")
		Sleep(0300)
		_SendEx($BIB_LOC)
		Sleep(0100)
		WinWaitActive("[TITLE:Edit Data; CLASS:SunAwtDialog]")
		Sleep(0300)
		_SendEx($BIB_LOC_L1)
		Sleep(0300)
		_SendEx("!a")
		Sleep(0300)
		_SendEx($BIB_LOC_L2)
		Sleep(0300)
		_SendEx("!o")
	ElseIf $multi_loc <> "" Then
		_SendEx("^{HOME}")
		Sleep(0100)
		_SendEx("{DOWN 2}")
		Sleep(0300)
		_SendEx("multi")
		Sleep(0100)
		WinWaitActive("[TITLE:Edit Data; CLASS:SunAwtDialog]")
		Sleep(0300)
		$multi_loc_a = StringSplit($multi_loc, ",")
		Do
			$LOCATION = _ArrayPop($multi_loc_a)
			$LOCATION = StringRegExpReplace($LOCATION, "[(\d)]", "")
			$LOCATION = StringStripWS($LOCATION, 8)
			_StoreVar("$LOCATION")
			Run(@DesktopDir & "\Receipt\Bib - check ILOC.exe")
			While ProcessExists("Bib - check ILOC.exe")
				Sleep(0400)
			WEnd
			$BIB_LOC = _LoadVar("$BIB_LOC")
			$REF = _LoadVar("$REF")
			If $BIB_LOC = "multi" Then ; if item is for reference
				$BIB_LOC_L1 = _LoadVar("$BIB_LOC_L1")
				$BIB_LOC_L2 = _LoadVar("$BIB_LOC_L2")
				Sleep(0300)
				_SendEx($BIB_LOC_L1)
				Sleep(0300)
				_SendEx("!a")
				Sleep(0300)
				_SendEx($BIB_LOC_L2)
			Else
				Sleep(0300)
				_SendEx($BIB_LOC)
			EndIf
			Sleep(0300)
			_SendEx("!a")
			Sleep(0300)
		Until $multi_loc_a = ""
		Sleep(0300)
		_SendEx("!o")
	Else
		Sleep(0400)
	EndIf
	MsgBox(64, "Review the content lines before adding item.", "Review the content lines before adding item.", 5) ;done!
Else
	If $T85 <> "" OR $T86 <> "" OR $T87 <> "" Or $T88 <> "" Or $T89 <> "" Then
		MsgBox(0, "Receipt Cataloging", $T85 & @CR & $T86 & @CR & $T87 & @CR & $T88 & @CR & $T89)
	EndIf
	Sleep(0400)
	Run(@DesktopDir & "\Receipt\Create item - Non DLC.exe")
EndIf
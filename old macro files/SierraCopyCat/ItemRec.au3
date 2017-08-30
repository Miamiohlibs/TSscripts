#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: ItemRec
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script completes the copy cataloging process by entering the appropriate
	information in the bib record as well as perform quality control on select
	fields. The script then opens the item record and enters the appropriate
	codes and messages, creating additional item records if necesary.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(bib record open)
					Record view properties - Summary retrieval view, item
					summary view

 Last revised: 6/29/10
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
AutoItSetOption("SendKeyDelay", 30)
TraySetIcon(@DesktopDir & "\SierraCopyCat\Images\green D.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $UPD
Dim $ICODE1
Dim $row1_MASTER, $row2_MASTER, $row3_MASTER
Dim $BCODE3
dim $CAT_DATE
Dim $LANG
Dim $BCODE1
Dim $COUNTRY
Dim $SKIP
dim $BCODE2
Dim $BIB_LOC_1
dim $CAT_DATE_R
Dim $TODAY_DATE
dim $decide
Dim $BIB_LOC
Dim $LANG_CHECK
Dim $008_L_A
dim $008_L
Dim $COUNTRY_CHECK
Dim $008_C_A
Dim $008_C
Dim $SKIP_CHECK
Dim $LOCATION
Dim $ICODE1_1
Dim $ITYPE
Dim $LABELLOC
Dim $VOL, $vol_300, $vol_300_multi
Dim $BARCODE
Dim $DUP_CHECK
Dim $LAB_LOC
Dim $ACCOMP
Dim $C_INI
dim $dean
Dim $300_A, $300, $300_C_A, $300_C_A_1, $300_C, $300_C_X, $FOLIO, $300_E, $300_E_A, $300_E_1
Dim $260_A, $008_D_A, $008_D, $050_D, $050_A, $050, $260, $260_L, $260_C1, $008, $008_CON, $260_C, $090_A
Dim $245_2_A, $245_2
Dim $MULTI_LOC_A, $MULTI_LOC, $MULTI_REF, $kngr, $aar, $scr, $docr, $mdir, $doc, $imr
Dim $shelfready
Dim $IS_IMC
Dim $scfo, $scov
Dim $KNGFO
Dim $260, $260_C
Dim $264, $264_C
Dim $050, $050_B
Dim $008, $008_date
Dim $digCheck
Dim $conf
Dim $300error, $300errors, $300ok
Dim $donor_loc
Dim $SF_590, $SF_7XX, $SF_NAME, $SF_FIB
Dim $AAFO
Dim $don_name
Dim $264_C_fix, $264_fix, $260_fix
Dim $440_A, $490_A, $800_A, $810_A, $811_A, $830_A, $series
Dim $856, $956
Dim $subject
Dim $REF, $REF_1
Dim $OLOCATION
Global $050_array[2]
Dim $050_loc
Dim $049, $MIAS, $MIAA, $MIAH, $MIBN

$REF = _LoadVar("$REF")
$SF_NAME = _LoadVar("$SF_NAME")
$SF_590 = _LoadVar("$SF_590")
$SF_7XX = _LoadVar("$SF_7XX")
$SF_FIB = _LoadVar("$SF_FIB")
$OLOCATION = _LoadVar("$OLOCATION")
$ACCOMP = _LoadVar("$ACCOMP")
;################################ MAIN ROUTINE #################################
_Initial()
;focus on Millennium bib record
;If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;	WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the bib record.")
;	Exit
;EndIf

;/Emily/
#cs
If WinExists("[REGEXPTITLE:Sierra � Miami University Libraries � .* � [b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:Sierra � Miami University Libraries � .* � [b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra � Miami University Libraries � .* � [b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf
#ce

If WinExists("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the bib record.")
	Exit
EndIf

;force cataloger to verify that call number is unique

$decide = MsgBox(48+1, "Verify call number is unique", "Place the cursor in the call number field and type CTRL-G. Ensure call number is unique. Press OK if it is, and press cancel to fix.")
if $decide = 2 Then
	Exit
EndIf



;select all and copy bib record information, parce data into array
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;~ ##### start fixed data setting #####

;row 1 includes lang/cat date/ bcode3
$row1_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 1)
$BCODE3 = _ArrayPop($row1_MASTER)
$CAT_DATE = _ArrayPop($row1_MASTER)
$LANG = _ArrayPop($row1_MASTER)
$LANG = StringTrimLeft($LANG, 5)

;row 2 includes skip, bcode1, country
$row2_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 2)
$COUNTRY = _ArrayPop($row2_MASTER)
$BCODE1 = _ArrayPop($row2_MASTER)
$SKIP = _ArrayPop($row2_MASTER)
$SKIP = StringTrimLeft($SKIP, 5)
$COUNTRY = StringTrimLeft($COUNTRY, 7)

;row 3 includes bib location and bcode2
$row3_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 3)
$BCODE2 = _ArrayPop($row3_MASTER)
$BCODE2 = StringTrimLeft($BCODE2, 6)
$BIB_LOC_1 = _ArrayPop($row3_MASTER)
$BIB_LOC_1 = StringTrimLeft($BIB_LOC_1, 9)
$BIB_LOC_1 = StringLeft($BIB_LOC_1, 5)
$BIB_LOC_1 = StringStripWS($BIB_LOC_1, 8)
If $BIB_LOC_1 = "multi" Then
	$REF_1 = 1
Else
	$REF_1 = 0
EndIf
If $REF_1 = "0" Then
	If $BIB_LOC_1 <> $OLOCATION Then
		$decide = MsgBox(1, "Location mismatch", "The bib loc and the order location do not match.  Press OK to ignore, or cancel to fix.")
		If $decide = 2 Then
			Exit
		EndIf
	EndIf
EndIf


$049 = _arrayItemString($BIB_ARRAY_MASTER, "MARC	049")
$049 = StringRight($049, 4)
If $BIB_LOC_1 = "scl" Then
	If $049 <> "MIAS" Then
		MsgBox(0, "049", "The bib location and 049 fields do not match.")
		Exit
	EndIf
	ElseIf $BIB_LOC_1 = "kngl" Then
	If $049 <> "MIAA" Then
		MsgBox(0, "049", "The bib location and 049 fields do not match.")
		Exit
	EndIf
	ElseIf $BIB_LOC_1 = "aal" Then
	If $049 <> "MIAH" Then
		MsgBox(0, "049", "The bib location and 049 fields do not match.")
		Exit
	EndIf
EndIf






;~ ##### end fixed data setting #####

;- #### check for a multi volume set based on the 300 field of the bib ####
$vol_300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$vol_300_multi = StringInStr($vol_300, "volumes")
If $vol_300_multi > 2 Then
   $vol = 1
   ; MsgBox(0, "Volumes", $vol)
Else
   $vol = 0
EndIf

;get the call number to check against the location coding - still not implemented, just collecting variables
$050_loc = _arrayItemString($BIB_ARRAY_MASTER, "CALL #	050")
If $050_loc = -1 Then ; if the 050 is missing this checks quickly to see if there's a 090 instead. Instead of creating new variables, this just uses the existing variables.
	$050_loc = _arrayItemString($BIB_ARRAY_MASTER, "CALL #	090")
EndIf
$050_loc = StringTrimLeft($050_loc, 15)
_StoreVar("$050_loc")
;msgbox(0, "050 trim", $050_loc)


;##### start date check #####
;NOT LOOPED YET - should probably loop to make sure that person actually fixes the date...

;gather data to check dates (007 === end of 050 === |c of 26X)
$008 = _arrayItemString($BIB_ARRAY_MASTER, "MARC	008") ;these simply pull the string that starts with this value
$008_date = StringMid($008, 21, 4) ;this pulls out a substring from the main string that is 21 characters in from the left and four characters long
$050 = _arrayItemString($BIB_ARRAY_MASTER, "CALL #	050")
If $050 = -1 Then ; if the 050 is missing this checks quickly to see if there's a 090 instead. Instead of creating new variables, this just uses the existing variables.
	$050 = _arrayItemString($BIB_ARRAY_MASTER, "CALL #	090")
EndIf

;$050_B = StringRight($050, 4) ;this pulls out a substring that is four characters from the end of the main string
$050 = StringRight($050, 5)

$050_array = StringRegExp($050, '(\b[12][^a-zA-Z]{3})', 1)
If @error == 0 Then
	$050_B = $050_array[0]
	;MsgBox(0, "050", $050_B)
Else
	$decide = MsgBox(1, "Call number?", "There might be a problem with the call number. Click OK to ignore and Cancel to fix.")
	If $decide == 2 Then
		Exit
		EndIf
EndIf


If StringIsDigit($050_B) = 0 Then
	$decide = MsgBox(1, "Call Number Date", "The call number seems to be missing the date.  Click OK to ignore, or press cancel to fix.")
	If $decide = 2 Then
		Exit
	EndIf
EndIf

$260 = _arrayItemString($BIB_ARRAY_MASTER, "IMPRINT	260")
$260_C = StringInStr($260, $050_B, 0, 1, 1); will return 0 if the date from the 050 field is not found in the 260 (which means it's not present or it needs to check for a 264

#cs
$260_fix = StringRegExp($260, "(\[).{5}") ; have to cehck for the date in brackets in imprint field and remove brackets if needed
;MsgBox(0, "hello", $260_fix)
If $260_fix = 1 Then
	$260_C = StringRight($260, 6)
	$260_C = StringTrimLeft($260_C, 1)
	$260_C = StringTrimRight($260_C, 1)
	;msgbox(0, "check", $260_C)
Else
	$260_C = StringRight($260, 4) ;StringRight($260, 4) ;this pulls out a substring that is four characters from the end of the main string
	;msgbox(0, "check", $260_C)
EndIf
#ce

;msgbox(0, "260", $260_C)

If $260_C = 0 Then

	$264 = _arrayItemString($BIB_ARRAY_MASTER, "IMPRINT	264")
	$264_C = StringInStr($264, $050_B, 0, 1, 1); returns 0 if 050_B string is not found (i.e., there's no date in an imprint field in this record
	If $264_C = 0 Then ;if $264_C = 0 we know that $260_C = 0 as well, so there's no date
	$conf = MsgBox(1, "whoops", "The dates in the 008, Call # and Imprint fields don't match.  If this is a Conference Proceeding or otherwise correct, please click OK to proceed anyway. Click cancel to fix.")
	If $conf = 1 Then
		MsgBox(0, "OK!", "Moving on...")
	ElseIf $conf == 2 Then
		MsgBox(0, "Ok!", "Please click OK and adjust the date and re-run the D function.")
			 Exit
	EndIf
	Else
	$264_C = 1
	EndIf
	ElseIf $260_C > 0 Then
	$260_C = 1

EndIf

#cs
$264_fix = StringRegExp($264, "(\[).{5}") ; have to cehck for the date in brackets in imprint field and remove brackets if needed
;MsgBox(0, "hello", $264_fix)
If $264_fix = 1 Then
	$264_C = StringRight($264, 6)
	$264_C = StringTrimLeft($264_C, 1)
	$264_C = StringTrimRight($264_C, 1)
	;msgbox(0, "check", $264_C)
Else
	$264_C = StringRight($264, 4) ;this pulls out a substring that is four characters from the end of the main string
	;msgbox(0, "check", $264_C)
EndIf




; now we check the dates against each other
If StringLen($260_C) < 4 Then ;if our program sees taht the $260_C (260 field subfield C) is 0, that means it's empty which hopefully just means that we're dealing with a record with a 264 instead of a 260
   if StringLen($264_C) > 0 Then
	  if $264_C = $050_B and $050_B = $008_date Then
		 ;msgBox(0, "success!", "Your dates match!")
	  Else
		 $conf = MsgBox(1, "whoops", "The dates don't match.  If this is a Conference Proceeding or otherwise correct, please click OK to proceed anyway.")
		 If $conf == 1 Then
			 MsgBox(0, "Ok!", "OK, we'll keep going!")
		 ElseIf $conf == 2 Then
			 MsgBox(0, "Ok!", "Please click OK and adjust the date and re-run the D function.")
			 Exit
		 EndIf
	  EndIf
   Else
	  MsgBox(0, "oh dear!", "There's a big problem here with the 264!")
	  Exit
   EndIf
ElseIf StringLen($260_C) > 0 Then
   if $260_C = $050_B and $050_B = $008_date Then
	  ;msgBox(0, "success!", "Your dates match!")
	  Else
		 $conf = MsgBox(1, "whoops", "The dates don't match. If this is a Conference Proceeding or otherwise correct, please click OK to proceed anyway")
		 If $conf == 1 Then
			 MsgBox(0, "Ok!", "OK, we'll keep going!")
		 ElseIf $conf == 2 Then
			 MsgBox(0, "Ok!", "Please click OK and adjust the date and re-run the D function.")
			 Exit
		 EndIf
	  EndIf

Else
   MsgBox(0, "oh dear!", "There's a big problem here with the dates!")
   Exit
EndIf
#ce
Sleep(0500)

$300error = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$300errors = StringInStr($300error, ";|c", 1)
if $300errors < 2 Then
	$300ok = MsgBox(1, "Check 300 Field", "The 300 field appears to be missing data. Click OK to fix it or Cancel to ignore it.")
	if $300ok == 1 Then
		Exit
	EndIf
EndIf

;~ 	##### start series check #####
$440_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	440", 0, 0, 0, 1)
$490_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	490", 0, 0, 0, 1)
$800_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	800", 0, 0, 0, 1)
$810_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	810", 0, 0, 0, 1)
$811_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	811", 0, 0, 0, 1)
$830_A = _ArraySearch($BIB_ARRAY_MASTER, "SERIES	830", 0, 0, 0, 1)

$series = $800_A + $810_A + $811_A + $830_A
;MsgBox(0, "series", $series)

;_ArraySearch yields -1 as the value if there's an error (-1 would indicate that field is missing) so this needs to test to see if any of the 8XX field arrays are more than -1

If $490_A > -1  and $series = -4 Then
		$decide = msgbox(1, "Series Problem", "The 4XX field does not seem have a corresponding 8XX field. Click OK to ignore it, or cancel to stop and fix it.")
		If $decide = 2 Then
			Exit
	EndIf
EndIf
;~ 	##### end series check #####

;- #### check that it has a subject field ###
$subject = _ArraySearch($BIB_ARRAY_MASTER, "SUBJECT	", 0, 0, 0, 1)
;$subject_2 = _ArraySearch($BIB_ARRAY_MASTER, "SUBJECT 651", 0, 0, 0, 1)
;$subject_p = _ArraySearch($BIB_ARRAY_MASTER, "SUBJECT 600
If $subject = -1 Then
	$decide = MsgBox(1, "No subject data", "No subject field found.  If this item needs a subject field, press CANCEL to add one. Otherwise press OK to continue.")
	If $decide = 2 Then
		Exit
	EndIf
EndIf

;- #### check for existence of 856 and prompt to delete ####
$856 = _ArraySearch($BIB_ARRAY_MASTER, "MARC	856", 0, 0, 0, 1)
If $856 >-1 Then
	$decide = MsgBox(1, "856 Found", "There's an 856 in this record. Please delete if it's not required and then press OK to continue. Press cancel to stop macro.")
	If $decide = 2 Then
		Exit
	EndIf
EndIf


if $REF = 1 AND $REF_1 = 0 Then ; enter the extra BLOC if needed for a reference item

	_SendEx("^{HOME}")
	Sleep(0200)
	_SendEx("{DOWN 2}")
	Sleep(0200)
	_SendEx("^e")
	Sleep(0300)
	_SendEx("!a")
	Sleep(0400)
	_SendEx("ref")
	Sleep(0300)
	_SendEx("!o")
	Sleep(0300)

EndIf


; ### if there's a donor fund invovled, start adding the content lines ###
if StringLen($SF_NAME) > 1 Then
	If $SF_FIB = 1 Then
		$don_name = InputBox("Donor Name", "Enter the donor/group's name here (last name, first name if applicable)")
	_SendEx("^{HOME}")
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	_SendEx($SF_7XX)
	Sleep(0300)
	_SendEx($don_name)
	Sleep(0100)
	_SendEx(",|edonor")
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("n590{TAB}")
	Sleep(0100)
	Switch $BIB_LOC_1
		Case "kngl"
			$donor_loc = "King Library copy "
		Case "kngli"
			$donor_loc = "King Library copy "
		Case "scl"
			$donor_loc = "BEST Library copy "
		Case "scli"
			$donor_loc = "BEST Library copy "
		Case "imc"
			$donor_loc = "IMC copy "
		Case "aal"
			$donor_loc = "Art and Architecture Library copy "
		Case "aali"
			$donor_loc = "Art and Architecture Library copy "
	EndSwitch
	Sleep(0100)
	_SendEx($donor_loc)
	Sleep(0100)
	_SendEx($SF_590)
	Sleep(0300)
	_SendEx($don_name)
	Else
	_SendEx("^{HOME}")
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	;_SendEx("b7102")
	;Sleep(0100)
	;_SendEx("{TAB}")
	;Sleep(0100)
	_SendEx($SF_7XX)
	Sleep(0300)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("n590{TAB}")
	Sleep(0100)
	Switch $BIB_LOC_1
		Case "kngl"
			$donor_loc = "King Library copy "
		Case "kngli"
			$donor_loc = "King Library copy "
		Case "scl"
			$donor_loc = "BEST Library copy "
		Case "scli"
			$donor_loc = "BEST Library copy "
		Case "imc"
			$donor_loc = "IMC copy "
		Case "aal"
			$donor_loc = "Art and Architecture Library copy "
		Case "aali"
			$donor_loc = "Art and Architecture Library copy "
	EndSwitch
	Sleep(0100)
	_SendEx($donor_loc)
	Sleep(0100)
	_SendEx($SF_590)
	Sleep(0300)
	EndIf
EndIf


;get user initials, enter 947 field
$C_INI = _Initial()

$UPD = _LoadVar("$UPD") ;declared in barcodetorecord
Sleep(0400)
_SendEx("^{HOME}")
Sleep(0100)
_SendEx("{TAB 2}t")
Sleep(0100)
_SendEx("^{END}")
Sleep(0100)
_SendEx("{ENTER}y947{SPACE 2}")
Sleep(0400)
_SendEx($C_INI)


;check to see if item has accompanying material and if the item is over 27cm (goes into oversized/folio collection)
$300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$300 = StringTrimLeft($300, 18)
$300_C_A = StringSplit($300, "|")
$300_C = _arrayItemString($300_C_A, "cm")
$300_C_X = StringInStr($300_C, "x")
If $300_C_X > 0 Then
	$300_C_X = $300_C_X + 2
	$300_C_X = StringMid($300_C, $300_C_X, 2)
EndIf
$300_C = StringMid($300_C, 2, 2)
If $BIB_LOC_1 = "scl" AND $300_C >= 27 Then
	$SCOV = 1
EndIf
If $SCOV = 1 and $300_C >= 31 Then
	$SCFO = 1
	$SCOV = 0
EndIf
If $BIB_LOC_1 = "kngl" AND $300_C >= 31 Then
	$KNGFO = 1
EndIf
If $BIB_LOC_1 = "aal" and $300_C >= 30 Then
	$AAFO = 1
EndIf


;accompanying material search
$300_E = StringInStr($300, "|e")
If $300_E > 0 Then
	$300_E_A = $300_C_A_1 + 1
	$300_E = _ArrayToString($300_C_A, @TAB, $300_E_A, $300_E_A)
	$300_E = StringTrimLeft($300_E, 1)
	$300_E_1 = StringInStr($300_E, "(")
	$300_E = StringLeft($300_E, $300_E_1)
	$300_E = StringTrimRight($300_E, 1)
ElseIf $ACCOMP > -1 Then
	$decide = MsgBox(1, "Missing material?", "The 300 field might be missing subfield E for accompanying material. Press OK to ignore it and cancel to fix it.")
	If $decide = 2 Then
		MsgBox(0, "Stopping...", "The macro will stop. Please refer to the Wiki for proper coding of subfield E.")
		Exit
	Else
		Sleep(0200)
	EndIf

Else
	$300_E = "none"
EndIf

;~ ##### start language check #####
;looped - cannot continue until lang codes are matching
$LANG_CHECK = 1 ;when this turns 0, script exits loop and continues
Do
	$008_L = _arrayItemString($BIB_ARRAY_MASTER, "MARC	008") ;lang code in 008
	$008_L = StringMid($008_L, 49, 3)
	$LANG = StringLeft($LANG, 3) ;III lang code
	If $LANG <> $008_L Then
		MsgBox(64, "Mismatch between Language Fields", "Please make sure that the information is correct in the Lang and 008 fields. " & @CR & "Please adjust and click ok to continue.")
		;following doublechecks to see if fix has been done
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace($BIB_REC, "[\r\n]+", "fnord")
		$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)
		$row1_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 1)
		$BCODE3 = _ArrayPop($row1_MASTER)
		$CAT_DATE = _ArrayPop($row1_MASTER)
		$LANG = _ArrayPop($row1_MASTER)
		$LANG = StringTrimLeft($LANG, 5)
	Else
		$LANG_CHECK = 0
	EndIf
Until $LANG_CHECK = 0
;~ ##### end language check #####

;~ ##### start country check #####
;looped - cannot continue until cty codes are matching
$COUNTRY_CHECK = 1 ;when this turns 0, script exits loop and continues
Do
	$008_C = _arrayItemString($BIB_ARRAY_MASTER, "MARC	008") ;cty code in 008
	$008_C = StringMid($008_C, 29, 3)
	$008_C = StringStripWS($008_C, 8)
	$COUNTRY = StringLeft($COUNTRY, 3)
	$COUNTRY = StringStripWS($COUNTRY, 8) ;III cty code
	If $COUNTRY <> $008_C Then
		MsgBox(64, "Mismatch between Country Fields", "Please make sure that the information is correct in the Country and 008 fields. " & @CR & "Please adjust and click ok to continue.")
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace($BIB_REC, "[\r\n]+", "fnord")
		$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)
		$row2_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 2)
		$COUNTRY = _ArrayPop($row2_MASTER)
		$COUNTRY = StringTrimLeft($COUNTRY, 7)
	Else
		$COUNTRY_CHECK = 0
	EndIf
Until $COUNTRY_CHECK = 0
;~ ##### end country check #####


;~ ##### start skip check #####
;looped - cannot continue until skip codes are matching
$SKIP_CHECK = 1 ;when this turns 0, script exits loop and continues
Do
	$245_2 = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
	$245_2 = StringMid($245_2, 13, 1)
	If $SKIP <> $245_2 Then
		MsgBox(64, "Skip does not match 245 second indicator", "Skip does not match the 245 second indicator." & @CR & "Please adjust and click ok to continue.")
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace($BIB_REC, "[\r\n]+", "fnord")
		$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)
		$row2_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 2)
		$COUNTRY = _ArrayPop($row2_MASTER)
		$BCODE1 = _ArrayPop($row2_MASTER)
		$SKIP = _ArrayPop($row2_MASTER)
		$SKIP = StringTrimLeft($SKIP, 5)
	Else
		$SKIP_CHECK = 0
	EndIf
Until $SKIP_CHECK = 0
;~ ##### End skip check #####

;~ ##### Start location setting #####
;determines item location code
Switch $BIB_LOC_1
	Case "kngl"
		$LOCATION = "kngli"
	Case "scl"
		$LOCATION = "scli"
	Case "aal"
		$LOCATION = "aali"
	Case "doc"
		$LOCATION = "docli"
	Case "mdl"
		$LOCATION = "mdli"
	Case "imc"
		$LOCATION = "imc"
	Case "multi" ;multi usually occurs with reference items
		$MULTI_LOC = _arrayItemString($BIB_ARRAY_MASTER, "LOCATIONS	")
		$MULTI_LOC = StringTrimLeft($MULTI_LOC, 10)
		$MULTI_LOC_A = StringSplit($MULTI_LOC, ",")
		$MULTI_REF = _ArraySearch($MULTI_LOC_A, "ref", 0, 0, 0, 1)
		If $MULTI_REF > -1 Then ;if it's a reference item
			$aar = _ArraySearch($MULTI_LOC_A, "aal", 0, 0, 0, 1)
			$kngr = _ArraySearch($MULTI_LOC_A, "kngl", 0, 0, 0, 1)
			$scr = _ArraySearch($MULTI_LOC_A, "scl", 0, 0, 0, 1)
			$doc = _ArraySearch($MULTI_LOC_A, "doc", 0, 0, 0, 1)
			$mdir = _ArraySearch($MULTI_LOC_A, "mdli", 0, 0, 0, 1)
			$imr = _ArraySearch($MULTI_LOC_A, "imc", 0, 0, 0, 1)
			Select
				Case $aar > -1
					$LOCATION = "aar"
					$ITYPE = 21
				Case $kngr > -1
					$LOCATION = "kngr"
					$ITYPE = 21
				Case $scr > -1
					$LOCATION = "scr"
					$ITYPE = 21
				Case $doc > -1
					$LOCATION = "docr"
					$ITYPE = 21
				Case $mdir > -1
					$LOCATION = "mdlr"
					$ITYPE = 21
				Case $imr > -1
					$LOCATION = "imr"
					$ITYPE = 21
			EndSelect
		Else ;if it's not a reference item (added copy/vol)
			$decide = InputBox("Multi Location", "Which bib location is this item going to?")
			$decide = StringLower($decide)
			Switch $decide
				Case "kngl"
					$LOCATION = "kngli"
				Case "scl"
					$LOCATION = "scli"
				Case "aal"
					$LOCATION = "aali"
				Case "doc"
					$LOCATION = "docli"
				Case "mdl"
					$LOCATION = "mdli"
				Case "imc"
					$LOCATION = "imc"
			EndSwitch
		EndIf
EndSwitch

;check to see if item belongs in folio
If $FOLIO = 1 And $LOCATION <> "mdli" Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf

;since IMC has a slew of location codes, this one is manual entry for now
If $LOCATION = "imc" Then
	$decide = InputBox("IMC Item Location Code", "Enter the IMC Item Location code below", "", "", 150, 150)
	$LOCATION = $decide
EndIf
;~ ##### End location setting #####

;store location variable for bib - check ILOC script
_StoreVar("$LOCATION")

;~ set location/labelloc/itype codes
Run(@DesktopDir & "\SierraCopyCat\Bib - check ILOC.exe")
While ProcessExists("Bib - check ILOC.exe")
	Sleep(0400)
WEnd

Sleep(0200)

;load variables from bib - check ILOC and barcodetorecord scripts
$LABELLOC = _LoadVar("$LABELLOC")
$ITYPE = _LoadVar("$ITYPE")
$LAB_LOC = _LoadVar("$LAB_LOC")
;$VOL = _LoadVar("$VOL") tbhis script now checks for multiple volumes in the bib record 300 field
$ICODE1 = _LoadVar("$ICODE1")
$ACCOMP = _LoadVar("$ACCOMP")
$dean = _LoadVar("$dean")

;~ #### size check and change location endings to fo or ov ####

#cs
$IS_IMC = StringInStr($LOCATION, "im")
If $FOLIO = 1 AND $IS_IMC = 0 And $LOCATION <> "scli" AND $KNGFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
	_StoreVar("$LOCATION")
 ElseIf $FOLIO = 1 And $IS_IMC = 0 And $LOCATION = "scli" AND $SCFO <> 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "ov")
	_StoreVar("$LOCATION")
 ElseIf $FOLIO = 1 And $IS_IMC = 1 And $SCFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
	_StoreVar("LOCATION")
EndIf
#ce
If $SCOV = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "ov")
EndIf
If $SCFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf
If $KNGFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf
If $AAFO = 1 Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf

If $REF = 1 Then
	$ITYPE = 21
	_StoreVar("$ITYPE")
	$LOCATION = StringReplace($LOCATION, "li", "r")
EndIf

Sleep(0100)

;~ ##### start bcode2 (score) check #####
$BCODE2 = StringLeft($BCODE2, 1)
If $BCODE2 = "c" Then
	$ITYPE = 12
EndIf
;~ ##### end bcode2 (score) check #####

;MsgBox(0, "location check", $LOCATION)

; SAVE RECORD AND GO INTO ITEM RECORD
;focus on record
_SendEx("^s")
Sleep(0100)
_SendEx("^s")
Sleep(0100)
$DUP_CHECK = WinExists("[TITLE:Perform duplicate checking?; CLASS:SunAwtDialog]")
If $DUP_CHECK = 1 Then
	_SendEx("n")
	;WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	;/Emily/
	WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
EndIf
Sleep(1000)
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0300)
_SendEx("i")
Sleep(0400)
_SendEx("!l")
Sleep(0400)

;wait for item record focus
;WinWaitActive("[REGEXPTITLE:\A[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;/Emily/
WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;start item record data entry
Sleep(0100)
_SendEx("^a") ;probably don't to select all here, recommend deletion - CB 8/28/17
Sleep(0200)
_SendEx("^{HOME}")
Sleep(0300)
_SendEx("{TAB}")
Sleep(0400)
If $ICODE1 = 83 Then ;still not sure what's going on here.
	_SendEx("{DEL 2}")
	Sleep(0300)
Else
	;_SendEx("{DEL}")
	Sleep(0300)
EndIf
_SendEx("{RIGHT 1}")
Sleep(0300)
_SendEx($ICODE1) ;enters second (or third for middletown)
Sleep(0100)
_SendEx("3")
Sleep(0300)

_SendEx($LOCATION) ;edit location code here
Sleep(0400)

_SendEx("{DOWN}")
Sleep(0600)
_SendEx("{TAB}")
If $ICODE1 = 83 Then ;middletown
	_SendEx("l")
ElseIf $dean = 1 Then
	_SendEx("k")

EndIf

;$shelfready = InputBox("Shelf Ready Status", "Is this a shelf-ready item?" & @CR & "y - yes" & @CR & "n - no", "")

;Switch $shelfready
	;Case "n"
	$decide = MsgBox(4, "", "Is the item a paperback book?")
	If $decide = 6 Then
			Sleep(0100)
			_SendEx("r")
		ElseIf $decide = 7 Then
			;Sleep(0100)
			If $REF = 1 Then
				_SendEx("o")
			Else

			_SendEx("-")
			EndIf
		EndIf

	;EndSwitch
	;Case "y"
		;Sleep(0100)
		;_SendEx("-")
	;EndSwitch

; shelf ready paperbacks and hardbacks get status - for available
; non-shelf ready paperbacks get "r" and hardbacks get "l"

Sleep(0100)
_SendEx("{TAB 2}")
Sleep(0100)
_SendEx($ITYPE)
Sleep(0100)
_SendEx("^{END}")
Sleep(0300)
; Enter volume information if needed for item #1 after asking if necessary as sometimes volume info is already present in item record
If $vol = 1 Then

$decide = MsgBox(4, "Volume Designation", "Does item 1 need volume information?")
If $decide = 6 Then
$decide = InputBox("Volume label", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
_SendEx("{ENTER}")
Sleep(0100)
_SendEx("v{TAB}")
Sleep(0100)
_SendEx($decide)
Sleep(0400)
EndIf
EndIf


;determine if the item needs a location label
#cs taking out label loc JPM
If $LAB_LOC = 1 Then
	Sleep(0200)
ElseIf $LAB_LOC = 0 And $LABELLOC <> "" Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0500)
	_SendEx($LABELLOC)
	Sleep(0400)
EndIf
#ce
Sleep(0400)
If $dean = 1 Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("x{TAB}")
	Sleep(0500)
	_SendEx("Delivered to Dean's Office on item created date." & $C_INI)
EndIf
;save!
_SendEx("^s")
;determine if the item needs an accompanying material item note
If $300_E <> "none" And $ACCOMP = -1 Then
	$300_E = StringUpper($300_E)
	;focus on item window
	_SendEx("^{HOME}")
	Sleep(0500)
	_SendEx("{TAB}")
	Sleep(0500)
	_SendEx("{DOWN 5}")
	Sleep(0500)
	_SendEx("p")
	Sleep(0500)
	_SendEx("^{END}")
	Sleep(0500)
	_SendEx("{ENTER}")
	Sleep(0500)
	_SendEx("m{TAB}")
	Sleep(0500)
	_SendEx($300_E & "IN POCKET")
	Sleep(0500)
	_SendEx("^s")
EndIf
;determine if there needs to be more item records created
If $VOL = 1 Then
;looped - will continue loop until you answer no to the popup asking if you want to create another record
	Do
		$decide = MsgBox(4, "More Volumes", "Do you need to attach another item record?")
		Switch $decide
			Case 6 ;Yes answer
				;create new item record
				Sleep(1000)
				_SendEx("!g")
				Sleep(0200)
				_SendEx("o")
				Sleep(0200)
				_SendEx("a")
				Sleep(0500)
				_SendEx("!w")
				Sleep(0200)
				_SendEx("i")
				Sleep(0200)
				_SendEx("{TAB}")
				Sleep(0200)
				_SendEx("!n")
				Sleep(0400)
				;WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]", "")
				;start item record data entry
				_SendEx("^{HOME}")
				Sleep(0100)
				_SendEx("{TAB}")
				Sleep(0100)
				_SendEx($LOCATION)
				Sleep(0100)
				_SendEx("{TAB 2}")
				Sleep(0100)
				_SendEx("{DEL 2}")
				Sleep(0100)
				_SendEx($ICODE1)
				_SendEx("3")
				Sleep(0100)
				_SendEx("{TAB 3}")
				;$decide = MsgBox(4, "", "Is the item a paperback book?")
				;Switch $decide
					;Case 6
						;Sleep(0100)
						;_SendEx("r")
					;Case 7
						;Sleep(0100)
						_SendEx("-")
				;EndSwitch
				Sleep(0100)
				_SendEx("{TAB 2}")
				Sleep(0100)
				_SendEx($ITYPE)
				Sleep(0100)
				_SendEx("^{END}")
				Sleep(0300)
				_SendEx("{ENTER}")
				Sleep(0200)
				;If $LABELLOC <> "" Then
					;_SendEx("{ENTER}")
					;Sleep(0100)
					;_SendEx("l{TAB}")
					;Sleep(0100)
					;_SendEx($LABELLOC)
				;EndIf
				Sleep(0400)
				;volume number designation
				$decide = InputBox("Volume label", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
				_SendEx("{ENTER}")
				Sleep(0100)
				_SendEx("v{TAB}")
				Sleep(0100)
				_SendEx($decide)
				Sleep(0400)
				If $BARCODE = "N" Then
					_SendEx("^s")
				Else
					_SendEx("{ENTER}")
					Sleep(0100)
					_SendEx("b{TAB}")
					Sleep(0400)
					$decide = InputBox("Barcode", "Scan in Barcode", "")
					Sleep(0100)
					_SendEx($decide)
					Sleep(0100)
					_SendEx("^s")
				EndIf
				Sleep(0400)
			Case 7 ;No - ends loop
				$VOL = 0
		EndSwitch
	Until $VOL = 0
EndIf

If $dean = 1 Then
	MsgBox(0, "Dean's Office Materials", "Delete Catalog for Dean's office note.")
EndIf

_ClearBuffer()
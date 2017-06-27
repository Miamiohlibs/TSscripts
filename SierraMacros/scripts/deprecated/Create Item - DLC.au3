#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.4.4
 Authors:
 Martin Patrick, Academic Resident Librarian, Miami University
		patricm@miamioh.edu OR martinpatrick@outlook.com
 Jason Paul Michel, User Experience Librarian, Miami University
		micheljp@miamioh.edu
 Becky Yoose, Bibliographic Systems Librarian, Miami University
		b.yoose@gmail.com

 Name of script: ItemCreate

 Script Function:
	This script finalizes the bib record, creates a new
	item record, and inputs the appropriate codes in the item record.

 Programs used: III Sierra, Windows 8.1

 Last revised: 04/2015

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
;~ AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("SendKeyDelay", 30)
TraySetIcon(@WorkingDir & "\Images\greenD.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $IS_MUS
Dim $ITEM_STAT
Dim $row1_MASTER, $row2_MASTER, $row3_MASTER
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
Dim $COPIES
Dim $C_INI
Dim $BIB_LOC
Dim $RLOC
Dim $LANG_CHECK
Dim $008_L_A, $008_L
Dim $COUNTRY_CHECK
Dim $008_C_A, $008_C
Dim $SKIP_CHECK
Dim $LOCATION
Dim $ICODE1_1
Dim $ITYPE
Dim $LABELLOC
Dim $VOL
Dim $BARCODE
Dim $IS_DOC
Dim $ADD_ITEM
Dim $300_E
Dim $T25
Dim $REF, $REF_1
Dim $ITYPE
Dim $var, $goaway
Dim $LOCAL
Dim $EBOOK
Dim $reserve
Dim $245_2_A, $245_2
Dim $DUP_CHECK
Dim $dean
Dim $multi_loc
Dim $end
Dim $approval
Dim $score, $score_check_1, $score_check_2
Dim $FORM
Dim $donor_loc
Dim $SF_590, $SF_7XX, $SF_NAME, $SF_FIB
Dim $AAFO
Dim $don_name
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
Dim $BIB_LOC_1, $BCODE2
Dim $264_C_fix, $264_fix, $260_fix
Dim $OLOCATION

;################################ MAIN ROUTINE #################################
;start loading variables from order record script
$300_E = _LoadVar("$300_E")
$LOCATION = _LoadVar("$LOCATION")
$RLOC = _LoadVar("$RLOC")
$BARCODE = _LoadVar("$BARCODE")
$BIB_LOC = _LoadVar("$BIB_LOC")
$COPIES = _LoadVar("$COPIES")
$ICODE1_1 = _LoadVar("$ICODE1_1")
$LABELLOC = _LoadVar("$LABELLOC")
$LOCATION = _LoadVar("$LOCATION")
$RLOC = _LoadVar("$RLOC")
$T25 = _LoadVar("$T25")
$VOL = _LoadVar("$VOL")
$REF = _LoadVar("$REF")
$ITYPE = _LoadVar("$ITYPE")
$LOCAL = _LoadVar("$LOCAL")
$EBOOK = _LoadVar("$EBOOK")
$reserve = _LoadVar("$reserve")
$dean = _LoadVar("$dean")
$multi_loc = _LoadVar("$multi_loc")
$approval = _LoadVar("$approval")
$FORM = _LoadVar("$FORM")
$SF_NAME = _LoadVar("$SF_NAME")
$SF_590 = _LoadVar("$SF_590")
$SF_7XX = _LoadVar("$SF_7XX")
$SF_FIB = _LoadVar("$SF_FIB")
$C_INI = _LoadVar("$C_INI")


;focus on bib record screen
;If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the bib record.")
;	Exit
;EndIf

If WinExists("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf

;Sierra workaround / focus on bib record screen
;If WinExists("[TITLE:Sierra; CLASS:SunAwtFrame]") Then
;	WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
;	WinWaitActive ("[TITLE:Sierra; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Sierra record", "Please open the bib record.")
;	Exit
;EndIf

;force cataloger to verify that call number is unique

$decide = MsgBox(48+1, "Verify call number is unique", "Place the cursor in the call number field and type CTRL-G. Ensure call number is unique. Press OK if it is, and press cancel to fix.")
if $decide = 2 Then
	Exit
EndIf


;select all and copy bib record information, parce data into array
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

;~ ##### end fixed data setting #####



; double check if music score to adjust item record
$score_check = _ArraySearch($BIB_ARRAY_MASTER, "DESCRIPT.	300	 	 	", 0, 0, 0, 1)
;MsgBox(0, "title", $score_check)
$score_check_1 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $score_check, $score_check) ; arguements are ARRAY, DELIMITER, START, END
$score_check_2 = StringInStr($score_check_1, "score")
;msgBox(0, "title", $score_check_2)
if $score_check_1 > 2 Then
   $score = 1
   _StoreVar("$score")
Else
   $score = 0
   _StoreVar("$score")
EndIf

If $FORM = "w" Then
	$ITYPE = 12
EndIf

sleep(0500)

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
$050 = StringRight($050, 7)
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

Sleep(0500)

$300error = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$300errors = StringInStr($300error, ";|c", 1)
if $300errors < 2 Then
	$300ok = MsgBox(1, "Check 300 Field", "The 300 field appears to be missing data. Click OK to fix it or Cancel to ignore it.")
	if $300ok == 1 Then
		Exit
	EndIf
EndIf

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
	$LOCATION = StringReplace($LOCATION, "li", "ov")
EndIf
If $SCOV = 1 and $300_C >= 31 Then
	$SCFO = 1
	$SCOV = 0
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf
If $BIB_LOC_1 = "kngl" AND $300_C >= 31 Then
	$KNGFO = 1
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf
If $BIB_LOC_1 = "aal" and $300_C >= 30 Then
	$AAFO = 1
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf

#cs - merged this into the above block for simplicity
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
#ce

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

;enter bib location code and 947
$C_INI = _Initial()
sleep(0400)
_SendEx("^{HOME}")
Sleep(0100)
_SendEx("{TAB 2}t")
Sleep(0100)
_SendEx("^{HOME}")
Sleep(0100)
If $BIB_LOC = "multi" Then
	SLEEP(0100)
Else
	_SendEx("{DOWN 2}")
	_SendEx($BIB_LOC)
EndIf
Sleep(0100)
_SendEx("^{END}")
Sleep(0100)
_SendEx("{ENTER}y947{SPACE 2}")
Sleep(0400)
_SendEx($C_INI)
If $RLOC = "b" OR $RLOC = "n" Then
	_SendEx("{SPACE}upd")
	Sleep(0200)
EndIf

;~ ##### start language check #####
;looped - cannot continue until lang codes are matching
$LANG_CHECK = 1 ;when this turns 0, script exits loop and continues
Do
	$008_L = _arrayItemString($BIB_ARRAY_MASTER, "MARC	008")
	$008_L =  StringMid($008_L, 49, 3)
	$LANG = StringLeft($LANG, 3)
	If $LANG <> $008_L Then
		MsgBox(64, "Mismatch between Language Fields", "Please make sure that the information is correct in the Lang and 008 fields. " & @CR & "Please adjust and click ok to continue.")
		_DataCopy()
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
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
	$008_C = _arrayItemString($BIB_ARRAY_MASTER, "MARC	008")
	$008_C =  StringMid($008_C, 29, 3)
	$008_C = StringStripWS($008_C, 8)
	$COUNTRY = StringLeft($COUNTRY, 3)
	$COUNTRY = StringStripWS($COUNTRY, 8)
	If $COUNTRY <> $008_C Then
		MsgBox(64, "Mismatch between Country Fields", "Please make sure that the information is correct in the Country and 008 fields. " & @CR & "Please adjust and click ok to continue.")
		_DataCopy()
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
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
	$245_2 =  StringMid($245_2, 13, 1)
	If $SKIP <> $245_2 Then
		MsgBox(64, "Skip does not match 245 second indicator", "Skip does not match the 245 second indicator." & @CR & "Please adjust and click ok to continue.")
		_DataCopy()
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
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

; SAVE RECORD AND CREATE NEW ITEM RECORD
_SendEx("^s")
Sleep(0100)
$DUP_CHECK = WinExists("[TITLE:Perform duplicate checking?; CLASS:SunAwtDialog]")
If $DUP_CHECK = 1 Then
	_SendEx("n")
	WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
EndIf
Sleep(0800)
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("a")
;WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
;WinWaitActive ("[TITLE:Sierra; CLASS:SunAwtFrame]")
Sleep(0800)
MouseClick("left", 667, 200, 2)
Sleep(0200)
_SendEx("!w")
Sleep(0200)
_SendEx("i")
Sleep(0200)
If $approval <> 1 Then
_SendEx("!n")
Sleep(0200)
Else
   _SendEx("!l")
   Sleep(0200)
EndIf

IF WinExists("[TITLE:Attach New Records?]") Then
	_SendEx("y")
EndIf

;wait for item record focus
;Sierra needs MouseClick
;WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]")
MouseClick("left", 667, 519, 1)
Sleep(1000)
_SendEx("^{HOME}")

;T25 variable affects status.... look at bib- check Iloc File
; LOCAL USE CHECK
If $LOCAL > -1 Then
	$ITYPE = 26
endif

;dean's office materials setting
If $dean = 1 Then
	$LOCATION = "dean"
EndIf

;multiple locations indicated in teh order record
If $multi_loc <> "" Then
	$multi_a = StringSplit($multi_loc, ",")
	$LOCATION = _ArrayToString($multi_a, @TAB, 1, 1)
	$LOCATION = StringRegExpReplace($LOCATION, "[(\d)]", "")
	$LOCATION = StringStripWS($LOCATION, 8)
	MsgBox(0, "multiloc", $LOCATION)
	_StoreVar("$LOCATION")
	Run(@WorkingDir & "\scripts\Bib - check ILOC.exe")
	While ProcessExists("Bib - check ILOC.exe")
		Sleep(0400)
	WEnd
	$ITYPE = _LoadVar("$ITYPE")
	$LABELLOC = _LoadVar("$LABELLOC")
	$T25 = _LoadVar("$T25")
EndIf

If $REF = 1 Then
	$ITYPE = 21
	_StoreVar("$ITYPE")
	$LOCATION = StringReplace($LOCATION, "li", "r")
EndIf


;start item record data entry
Sleep(0400)
If $COPIES > 1 Then
	$decide = InputBox("Copies", "Enter copy designation (e.g., '1' OR '2')", "")
	_SendEx("{DEL 2}")
	Sleep(0100)
	_SendEx($decide)
	Sleep(0100)
	_SendEx("{TAB}")
	Sleep(0100)
Else
	_SendEx("{TAB}")
	Sleep(0100)
EndIf
Sleep(0300)
_SendEx($LOCATION)
Sleep(0100)
_SendEx("{TAB 2}")
Sleep(0100)
_SendEx("{DEL}")
Sleep(0100)
_SendEx($ICODE1_1)
_SendEx("2")
Sleep(0100)
_SendEx("{TAB 3}")
If $dean = 1 Then
	Sleep(0100)
	_SendEx("k")
Else
	sleep(0100)
	If $REF = 1 Then
		_SendEx("o")
		Else
   _SendEx("-")
		EndIf

EndIf
Sleep(0100)
_SendEx("{TAB 2}")
Sleep(0100)
_SendEx($ITYPE)
Sleep(0100)
_SendEx("^{END}")
Sleep(0300)
#cs
If $LABELLOC <> "" Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0100)
	_SendEx($LABELLOC)
EndIf
#ce

SLEEP(0400)
If $VOL > 1 Then
	If $LOCATION = "murc" Then
		$decide = "AudioD"
	Else
		$decide = InputBox("Volume lable", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
	EndIf
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("v{TAB}")
	Sleep(0100)
	_SendEx($decide)
EndIf
Sleep(0400)
If $dean = 1 Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("x{TAB}")
	Sleep(0100)
	_SendEx("Delivered to Dean's Office on item created date." & $C_INI)
EndIf
SLEEP(0400)
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

SLEEP(0400)
;determine if the item needs an accompanying material item note
If $300_E = 0 Then
		sleep(0100)
Else
	$IS_MUS = StringInStr($LOCATION, "mu")
	If $IS_MUS = 0 Then
			$300_E = StringUpper($300_E)
			_SendEx("^{HOME}")
			Sleep(0100)
			_SendEx("{TAB}")
			Sleep(0100)
			_SendEx("{DOWN 5}")
			Sleep(0100)
			_SendEx("p")
			Sleep(0100)
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}")
			Sleep(0100)
			_SendEx("m{TAB}")
			Sleep(0300)
			_SendEx($300_E & "IN POCKET")
			Sleep(0100)
			_SendEx("^s")
	Else
		$ADD_ITEM = 1
	EndIf
EndIf
SLEEP(0400)

;determine if there needs to be more item records created
If $VOL > 1 OR $COPIES > 1 Then
;~ 	looped - will continue loop until you answer no to the popup asking if you want to create another record
	$end = 0
	Do
		$decide = MsgBox(4, "More Volumes/Copies", "Do you need to attach another item record?")
		Switch $decide
			Case 6
				;determine which location the item should be going to if multiple locations for multi copies
				;And now for the very messy part of the script. Apologies for anyone who has to deal with the following
				If $multi_loc <> "" Then
					_StoreVar("$multi_loc")
					Run(@DesktopDir & "\SierraReceipt - gmd\multi locations item record.exe")
					While ProcessExists("multi locations item record.exe")
						Sleep(0400)
					WEnd
					$LOCATION = ClipGet()
					$LOCATION = StringRegExpReplace($LOCATION, "[(\d)]", "")
					$LOCATION = StringStripWS($LOCATION, 8)
					_StoreVar("$LOCATION")
					Run(@DesktopDir & "\SierraReceipt - gmd\Bib - check ILOC.exe")
					While ProcessExists("Bib - check ILOC.exe")
						Sleep(0400)
					WEnd
					$ITYPE = _LoadVar("$ITYPE")
					$LABELLOC = _LoadVar("$LABELLOC")
					$T25 = _LoadVar("$T25")
				EndIf
				;ok, it's safe to open your eyes again...
				Sleep(1000)
				_SendEx("!g")
				Sleep(0200)
				_SendEx("o")
				Sleep(0200)
				;_SendEx("a")
				;Sleep(0500)
				;_SendEx("!w")
				;Sleep(0200)
				_SendEx("i")
				Sleep(0300)
				_SendEx("!n")
				Sleep(0200)
				_SendEx("^{HOME}")
				Sleep(0500)
				If $COPIES > 1 Then
					$decide = InputBox("Copies", "Enter copy designation (e.g., '1' OR '2')", "")
					_SendEx("{DEL 2}")
					Sleep(0100)
					_SendEx($decide)
					Sleep(0100)
					_SendEx("{TAB}")
					Sleep(0100)
				Else
					_SendEx("{TAB}")
					Sleep(0100)
				EndIf
				_SendEx($LOCATION)
				Sleep(0100)
				_SendEx("{TAB 2}")
				Sleep(0100)
				_SendEx("{DEL}")
				Sleep(0100)
				_SendEx($ICODE1_1)
				_SendEx("1")
				Sleep(0100)
				_SendEx("{TAB 3}")
				;$decide = MsgBox(4, "","Is the item a paperback book?")
					;Switch $decide
						;Case 6
							sleep(0100)
							_SendEx("-")
						;Case 7
						;	sleep(0100)
							;_SendEx("l")
					;EndSwitch
				Sleep(0100)
				_SendEx("{TAB 2}")
				Sleep(0100)
				_SendEx($ITYPE)
				Sleep(0100)
				_SendEx("^{END}")
				Sleep(0300)
				;If $LABELLOC <> "" Then
					_SendEx("{ENTER}")
					Sleep(0100)
					;_SendEx("l{TAB}")
					;Sleep(0100)
					;_SendEx($LABELLOC)
				;EndIf
				SLEEP(0400)
				If $VOL > 1 Then
					If $LOCATION = "murc" Then
						$decide = "AudioD"
					Else
						$decide = InputBox("Volume lable", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
					EndIf
					_SendEx("{ENTER}")
					Sleep(0100)
					_SendEx("v{TAB}")
					Sleep(0100)
					_SendEx($decide)
				EndIf
				SLEEP(0400)
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
				SLEEP(0400)
			Case 7
				$end = 1
			EndSwitch
	Until $end = 1
EndIf

If $T25 <> 0 Then
	MsgBox(0, "Reserves", "RESERVE ITEM -- GIVE TO PROCESSING.")
EndIf

If $reserve = 1 Then
	MsgBox(0, "Reserves", "Flag for reserves. Insert red flag for rush, insert blue flag for King reserve.")
EndIf

If $dean = 1 Then
	MsgBox(0, "Dean's Office Materials", "Insert red flag for rush processing.")
EndIf

;determine if new item record for accompanying material needs creating
If $ADD_ITEM = 1 Then
	Sleep(0100)
	_SendEx("!g")
	Sleep(0200)
	_SendEx("o")
	Sleep(0200)
	_SendEx("a")
	;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	Sleep(0200)
	_SendEx("!g")
	Sleep(0200)
	_SendEx("o")
	Sleep(0200)
	_SendEx("i")
	Sleep(0200)
	_SendEx("!n")
	Sleep(0100)
	WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]")
	Sleep(0100)
		$decide = MsgBox(4, "","Is the item a non-music CD?")
			Switch $decide
				Case 6
					$LOCATION = "muc"
					$VOL = "disc"
					$ITYPE = 18
					$LABELLOC = "Music Circ"
				Case 7
					$LOCATION = "murc"
					$VOL = "AudioD"
					$ITYPE = 1
					$LABELLOC = "Music"
			EndSwitch
	_SendEx("^{HOME}")
	Sleep(0100)
	_SendEx("{TAB}")
	Sleep(0100)
	_SendEx($LOCATION)
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx("{DEL}")
	Sleep(0100)
	_SendEx($ICODE1_1)
	_SendEx("1")
	Sleep(0100)
	_SendEx("{TAB 3}")
	Sleep(0100)
	#cs
	$decide = MsgBox(4, "","Does the disc accompany a paperback book?")
		Switch $decide
			Case 6
				sleep(0100)
				_SendEx("r")
			Case 7
			#ce
   sleep(0100)
   _SendEx("-")
		; EndSwitch
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx($ITYPE)
	Sleep(0100)
	_SendEx("^{END}")
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0100)
	_SendEx($LABELLOC)
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("v{TAB}")
	Sleep(0100)
	_SendEx($VOL)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("b{TAB}")
	Sleep(0200)
	$decide = InputBox("Barcode", "Scan in Barcode", "")
	Sleep(0100)
	_SendEx($decide)
	Sleep(0100)
	_SendEx("^s")
EndIf

If $EBOOK = 1 Then
	MsgBox(64, "Paperwork to Jennifer", "Please send the accompanying paperwork to Jennifer for URL connection & cataloging.")
Endif

_ClearBuffer()

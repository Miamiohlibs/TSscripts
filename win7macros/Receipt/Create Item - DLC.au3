#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: Create Item - DLC
 Script set: Receipt Cataloging (Receipt)

 Script Function:
	This script inputs the appropriate codes in the bib record, creates a new
	item record, and inputs the appropriate codes in the item record.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(bibliographic record open)

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
;~ AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("SendKeyDelay", 30)
TraySetIcon(@DesktopDir & "\Receipt\Images\green D.ico")

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
Dim $REF
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


;focus on bib record screen
If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Millennium record", "Please open the bib record.")
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
;~ ##### end fixed data setting #####

sleep(0500)

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
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
EndIf
Sleep(1000)
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("a")
WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Sleep(0800)
_SendEx("!w")
Sleep(0200)
_SendEx("i")
Sleep(0200)
_SendEx("!n")
Sleep(0200)
IF WinExists("[TITLE:Attach New Records?]") Then
	_SendEx("y")
EndIf

;wait for item record focus
WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]")
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
	_StoreVar("$LOCATION")
	Run(@DesktopDir & "\Receipt\Bib - check ILOC.exe")
	While ProcessExists("Bib - check ILOC.exe")
		Sleep(0400)
	WEnd
	$ITYPE = _LoadVar("$ITYPE")
	$LABELLOC = _LoadVar("$LABELLOC")
	$T25 = _LoadVar("$T25")
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
	$decide = MsgBox(4, "","Is the item a paperback book?")
		Switch $decide
			Case 6
				sleep(0100)
				_SendEx("r")
			Case 7
				sleep(0100)
				_SendEx("l")
		EndSwitch
EndIf
Sleep(0100)
_SendEx("{TAB 2}")
Sleep(0100)
_SendEx($ITYPE)
Sleep(0100)
_SendEx("^{END}")
Sleep(0300)
If $LABELLOC <> "" Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0100)
	_SendEx($LABELLOC)
EndIf
SLEEP(0400)
If $VOL > 1 Then
	If $LOCATION = "murc" Then
		$decide = "AudioD"
	Else
		$decide = InputBox("Volume lable", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
	EndIf
	Sleep(0300)
	_SendEx("{ENTER}")
	Sleep(0300)
	_SendEx("v{TAB}")
	Sleep(0300)
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
	Sleep(0500)
	_SendEx($decide)
	Sleep(0300)
	_SendEx("^s")
EndIf

SLEEP(0400)
;determine if the item needs an accompanying material item note
If $300_E = "none" Then
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
					Run(@DesktopDir & "\Receipt\multi locations item record.exe")
					While ProcessExists("multi locations item record.exe")
						Sleep(0400)
					WEnd
					$LOCATION = ClipGet()
					$LOCATION = StringRegExpReplace($LOCATION, "[(\d)]", "")
					$LOCATION = StringStripWS($LOCATION, 8)
					_StoreVar("$LOCATION")
					Run(@DesktopDir & "\Receipt\Bib - check ILOC.exe")
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
				_SendEx("a")
				Sleep(0500)
				_SendEx("!w")
				Sleep(0200)
				_SendEx("i")
				Sleep(0200)
				_SendEx("!n")
				Sleep(0200)
				WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]", "")
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
				$decide = MsgBox(4, "","Is the item a paperback book?")
					Switch $decide
						Case 6
							sleep(0100)
							_SendEx("r")
						Case 7
							sleep(0100)
							_SendEx("l")
					EndSwitch
				Sleep(0100)
				_SendEx("{TAB 2}")
				Sleep(0100)
				_SendEx($ITYPE)
				Sleep(0100)
				_SendEx("^{END}")
				Sleep(0300)
				If $LABELLOC <> "" Then
					_SendEx("{ENTER}")
					Sleep(0100)
					_SendEx("l{TAB}")
					Sleep(0100)
					_SendEx($LABELLOC)
				EndIf
				SLEEP(0400)
				If $VOL > 1 Then
					If $LOCATION = "murc" Then
						$decide = "AudioD"
					Else
						$decide = InputBox("Volume lable", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
					EndIf
					Sleep(0300)
					_SendEx("{ENTER}")
					Sleep(0300)
					_SendEx("v{TAB}")
					Sleep(0300)
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
					Sleep(0300)
					_SendEx($decide)
					Sleep(0300)
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
	WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
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
	$decide = MsgBox(4, "","Does the disc accompany a paperback book?")
		Switch $decide
			Case 6
				sleep(0100)
				_SendEx("r")
			Case 7
				sleep(0100)
				_SendEx("l")
		EndSwitch
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
	Sleep(0300)
	_SendEx($decide)
	Sleep(0100)
	_SendEx("^s")
EndIf

If $EBOOK = 1 Then
	MsgBox(64, "Paperwork to Jennifer", "Please send the accompanying paperwork to Jennifer for URL connection & cataloging.")
Endif

_ClearBuffer()
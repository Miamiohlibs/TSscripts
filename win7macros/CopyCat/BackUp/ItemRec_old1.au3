AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("SendKeyDelay", 30)

TraySetIcon(@DesktopDir & "\CopyCat\Images\green D.ico")

#Include <TSCustomFunction.au3>

Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER, $UPD, $ICODE1
Dim $row1, $row1_prep, $row1_MASTER, $row2, $row2_prep, $row2_MASTER, $row3, $row3_prep, $row3_MASTER
Dim $BCODE3, $CAT_DATE, $LANG, $BCODE1, $COUNTRY, $SKIP, $BCODE2, $BIB_LOC_1, $CAT_DATE_R, $TODAY_DATE, $decide
Dim $BIB_LOC, $LANG_CHECK, $008_L_A, $008_L, $COUNTRY_CHECK, $008_C_A, $008_C, $SKIP_CHECK, $LOCATION, $ICODE1_1, $decide, $ITYPE, $LABELLOC, $VOL, $BARCODE
Dim $DUP_CHECK, $VOL, $LAB_LOC, $ACCOMP
Dim $var, $C_INI
Dim $300_A, $300, $300_C_A, $300_C_A_1, $300_C, $300_C_X, $FOLIO, $300_E, $300_E_A, $300_E_1
Dim $260_A, $008_D_A, $008_D, $050_D, $050_A, $050, $260, $260_L, $260_C1, $008, $008_CON, $260_C, $090_A
Dim $245_2_A, $245_2
Dim $MULTI_LOC_A, $MULTI_LOC, $MULTI_REF, $kngr, $aar, $scr, $docr, $mdir, $doc, $imr

$UPD = _LoadVar("$UPD")

If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Millennium record", "Please open the bib record.")
	Exit
EndIf
_DataCopy()


$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;~ _ArrayDisplay($BIB_ARRAY_MASTER)

;~ ################# start fixed data setting ##################

;row 1 includes lang/cat date/ bcode3
$row1 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, 1, 1)
$row1_prep = StringRegExpReplace ($row1, "\t+[A-Z]", "fnord")
$row1_MASTER = StringSplit($row1_prep, "fnord", 1)
$BCODE3 = _ArrayPop($row1_MASTER)
$CAT_DATE = _ArrayPop($row1_MASTER)
$LANG = _ArrayPop($row1_MASTER)
$LANG = StringTrimLeft($LANG, 5)

;row 2 includes skip, bcode1, country
$row2 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, 2, 2)
$row2_prep = StringRegExpReplace ($row2, "\t+[A-Z]", "fnord")
$row2_MASTER = StringSplit($row2_prep, "fnord", 1)
$COUNTRY = _ArrayPop($row2_MASTER)
$BCODE1 = _ArrayPop($row2_MASTER)
$SKIP = _ArrayPop($row2_MASTER)
$SKIP = StringTrimLeft($SKIP, 5)
$COUNTRY = StringTrimLeft($COUNTRY, 7)

;row 3 includes bib location and bcode2
$row3 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, 3, 3)
$row3_prep = StringRegExpReplace ($row3, "\t+[A-Z]", "fnord")
$row3_MASTER = StringSplit($row3_prep, "fnord", 1)
$BCODE2 = _ArrayPop($row3_MASTER)
$BIB_LOC_1= _ArrayPop($row3_MASTER)
$BIB_LOC_1 = StringTrimLeft($BIB_LOC_1, 9)
$BIB_LOC_1 = StringLeft($BIB_LOC_1, 5)
$BIB_LOC_1 = StringStripWS($BIB_LOC_1, 8)

;~ ################# end fixed data setting ##################

sleep(0500)

$C_INI = _Initial()

sleep(0400)
_SendEx("^{HOME}")
Sleep(0100)
_SendEx("{TAB 2}t")
Sleep(0100)
_SendEx("^{END}")
Sleep(0100)
_SendEx("{ENTER}y947{SPACE 2}")
Sleep(0400)
_SendEx($C_INI)
If $UPD = 1 Then
	_SendEx("{SPACE}upd")
EndIf

$300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$300 =  StringTrimLeft($300, 18)
$300_C_A = StringSplit($300, "|")
$300_C = _arrayItemString($300_C_A, "cm")
$300_C_X = StringInStr($300_C, "x")
If $300_C_X > 0 Then
	$300_C_X = $300_C_X + 2
	$300_C_X = StringMid($300_C, $300_C_X, 2)
EndIf
$300_C = StringMid($300_C, 2, 2)
If $300_C > 30 Or $300_C_X > 30 Then
	$FOLIO = 1
EndIf
$300_E = StringInStr($300, "|e")
If $300_E > 0 Then
	$300_E_A = $300_C_A_1 + 1
	$300_E = _ArrayToString($300_C_A, @TAB, $300_E_A, $300_E_A)
	$300_E = StringTrimLeft($300_E, 1)
	$300_E_1 = StringInStr($300_E, "(")
	$300_E = StringLeft($300_E, $300_E_1)
	$300_E = StringTrimRight($300_E, 1)
Else
	$300_E = "none"
EndIf


;~ ###################### start date check #############################

;NOT LOOPED YET

;search for 008 date
$008 = _arrayItemString($BIB_ARRAY_MASTER, "MARC	008")
$008_D =  StringMid($008, 21, 4)
$008_D = StringStripWS($008_D, 8)

;search for 050 date
$050_A = _ArraySearch($BIB_ARRAY_MASTER, "CALL #	050", 0, 0, 0, 1)
$090_A = _ArraySearch($BIB_ARRAY_MASTER, "CALL #	090", 0, 0, 0, 1)
If $050_A > -1 Then
	$050 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $050_A, $050_A)
ElseIf $090_A > -1 Then
	$050 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $090_A, $090_A)
EndIf
$050 =  StringTrimLeft($050, 15)
$050_D = StringRight($050, 4)
$050_D = StringStripWS($050_D, 8)

;search for 260 date
$260 = _arrayItemString($BIB_ARRAY_MASTER, "IMPRINT	260")
$260_L = StringLen($260)
$260_C1 = StringInStr($260, "|c")
$260_C = StringMid($260, $260_C1)
$260_C = StringTrimLeft($260_C, 2)
$260_C = StringRegExpReplace($260_C, "[][c?]", "")
$260_C = StringStripWS($260_C, 8)

If $260_C <> $008_D OR $008_D <> $050_D OR $260_C <> $050_D Then
	$008_CON = StringMid($008, 43, 1)
	IF $008_CON = 0 Then
		MsgBox(64, "Possible mismatch between date fields?", "Please make sure that the dates in the 008, 050 and 260 fields match if they are supposed to. " & @CR & "Please adjust (if necessary) and click ok to continue.")
	EndIf
EndIf

;~ ###################### end date check ###############################



;~ ##################### start language check ######################
$LANG_CHECK = 1

Do
	$008_L= _arrayItemString($BIB_ARRAY_MASTER, "MARC	008")
	$008_L =  StringMid($008_L, 49, 3)
	$LANG = StringLeft($LANG, 3)
	If $LANG <> $008_L Then
		MsgBox(64, "Mismatch between Language Fields", "Please make sure that the information is correct in the Lang and 008 fields. " & @CR & "Please adjust and click ok to continue.")
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
		$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)
		$row1 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, 1, 1)
		$row1_prep = StringRegExpReplace ($row1, "\t+[A-Z]", "fnord")
		$row1_MASTER = StringSplit($row1_prep, "fnord", 1)
		$BCODE3 = _ArrayPop($row1_MASTER)
		$CAT_DATE = _ArrayPop($row1_MASTER)
		$LANG = _ArrayPop($row1_MASTER)
		$LANG = StringTrimLeft($LANG, 5)
	Else
		$LANG_CHECK = 0
	EndIf
Until $LANG_CHECK = 0
;~ ##################### end language check ########################


;~ ##################### start country check #######################
$COUNTRY_CHECK = 1

Do
	$008_C= _arrayItemString($BIB_ARRAY_MASTER, "MARC	008")
	$008_C =  StringMid($008_C, 29, 3)
	$008_C = StringStripWS($008_C, 8)
	$COUNTRY = StringLeft($COUNTRY, 3)
	$COUNTRY = StringStripWS($COUNTRY, 8)
	If $COUNTRY <> $008_C Then
		MsgBox(64, "Mismatch between Country Fields", "Please make sure that the information is correct in the Country and 008 fields. " & @CR & "Please adjust and click ok to continue.")
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
		$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)
		$row2 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, 2, 2)
		$row2_prep = StringRegExpReplace ($row2, "\t+[A-Z]", "fnord")
		$row2_MASTER = StringSplit($row2_prep, "fnord", 1)
		$COUNTRY = _ArrayPop($row2_MASTER)
		$COUNTRY = StringTrimLeft($COUNTRY, 7)
	Else
		$COUNTRY_CHECK = 0
	EndIf
Until $COUNTRY_CHECK = 0
;~ ##################### end country check #########################


;~ ##################### start skip check ##########################
$SKIP_CHECK = 1

Do
	$245_2 = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
	$245_2 =  StringMid($245_2, 13, 1)
	If $SKIP <> $245_2 Then
		MsgBox(64, "Skip does not match 245 second indicator", "Skip does not match the 245 second indicator." & @CR & "Please adjust and click ok to continue.")
		$BIB_REC = ClipGet()
		$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
		$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)
		$row2 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, 2, 2)
		$row2_prep = StringRegExpReplace ($row2, "\t+[A-Z]", "fnord")
		$row2_MASTER = StringSplit($row2_prep, "fnord", 1)
		$COUNTRY = _ArrayPop($row2_MASTER)
		$BCODE1 = _ArrayPop($row2_MASTER)
		$SKIP = _ArrayPop($row2_MASTER)
		$SKIP = StringTrimLeft($SKIP, 5)
	Else
		$SKIP_CHECK = 0
	EndIf
Until $SKIP_CHECK = 0
;~ ##################### End skip check ############################


;~ ############### Start location setting ################

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
	Case "multi"
		$MULTI_LOC = _arrayItemString($BIB_ARRAY_MASTER, "LOCATIONS	")
		$MULTI_LOC = StringTrimLeft($MULTI_LOC, 10)
		$MULTI_LOC_A = StringSplit($MULTI_LOC, ",")
		$MULTI_REF = _ArraySearch($MULTI_LOC_A, "ref", 0, 0, 0, 1)
		If $MULTI_REF > -1 Then
			$aar = _ArraySearch($MULTI_LOC_A, "aal", 0, 0, 0, 1)
			$kngr = _ArraySearch($MULTI_LOC_A, "kngl", 0, 0, 0, 1)
			$scr = _ArraySearch($MULTI_LOC_A, "scl", 0, 0, 0, 1)
			$doc = _ArraySearch($MULTI_LOC_A, "doc", 0, 0, 0, 1)
			$mdir = _ArraySearch($MULTI_LOC_A, "mdli", 0, 0, 0, 1)
			$imr = _ArraySearch($MULTI_LOC_A, "imc", 0, 0, 0, 1)
			Select
				case $aar > -1
					$LOCATION = "aar"
				case $kngr > -1
					$LOCATION = "kngr"
				case $scr > -1
					$LOCATION = "scr"
				case $doc > -1
					$LOCATION = "docr"
				case $mdir > -1
					$LOCATION = "mdlr"
				case $imr > -1
					$LOCATION = "imr"
			EndSelect
		Else
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

If $FOLIO = 1 AND $LOCATION <> "mdli" Then
	$LOCATION = StringReplace($LOCATION, "li", "fo")
EndIf

If $LOCATION = "imc" Then
	$decide = InputBox("IMC Item Location Code", "Enter the IMC Item Location code below", "", "", 150, 150)
	$LOCATION = $decide
EndIf
;~ ############### End location setting ################

_StoreVar("$LOCATION")


;~ set location/labelloc/itype codes

Run(@DesktopDir & "\CopyCat\Bib - check ILOC.exe")
	While ProcessExists("Bib - check ILOC.exe")
		Sleep(0400)
	WEnd


$LABELLOC = _LoadVar("$LABELLOC")
$ITYPE = _LoadVar("$ITYPE")
$LAB_LOC = _LoadVar("$LAB_LOC")
$VOL = _LoadVar("$VOL")
$ICODE1 = _LoadVar("$ICODE1")
$ACCOMP = _LoadVar("$ACCOMP")


; SAVE RECORD AND GO INTO ITEM RECORD
;focus on record
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
_SendEx("y")


WinWaitActive ("[REGEXPTITLE:\A[i][0-9ax]{8}; CLASS:SunAwtFrame]")

Sleep(0100)
_SendEx("^{HOME}")
Sleep(0100)
_SendEx("{TAB}")
Sleep(0100)
_SendEx($LOCATION)
Sleep(0200)
_SendEx("{TAB 2}")
Sleep(0400)
If $ICODE1 = 83 Then
	_SendEx("{DEL 2}")
	Sleep(0300)
Else
	_SendEx("{DEL}")
	Sleep(0300)
EndIf
_SendEx($ICODE1)
_SendEx("3")
Sleep(0300)
_SendEx("{DOWN}")

Sleep(0600)
_SendEx("{TAB}")
If $ICODE1 = 83 Then
	_SendEx("l")
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



If $LAB_LOC = 1 Then
	Sleep(0200)
Elseif $LAB_LOC = 0 and $LABELLOC <> "" Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0100)
	_SendEx($LABELLOC)
	Sleep(0400)
EndIf

_SendEx("^s")


If $300_E <> "none" AND $ACCOMP = -1 Then
	$300_E = StringUpper($300_E)
	;focus on item window
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
	Sleep(0100)
	_SendEx($300_E & "IN POCKET")
	Sleep(0100)
	_SendEx("^s")
EndIf

If $VOL = 1 Then
	Do
	$decide = MsgBox(4, "More Volumes", "Do you need to attach another item record?")
	Switch $decide
		Case 6
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
			Sleep(0100)
			_SendEx("{TAB}")
			Sleep(0100)
			_SendEx($LOCATION)
			Sleep(0100)
			_SendEx("{TAB 2}")
			Sleep(0100)
			_SendEx("{DEL}")
			Sleep(0100)
			_SendEx($ICODE1)
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

			$decide = InputBox("Volume lable", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")

			_SendEx("{ENTER}")
			Sleep(0100)
			_SendEx("v{TAB}")
			Sleep(0100)
			_SendEx($decide)

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
			$VOL = 0
		EndSwitch
	Until $VOL = 0
EndIf

_ClearBuffer()
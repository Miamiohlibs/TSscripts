#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: ebook
 Script set: Order (order)

 Script Function:


	##### NOTE ##### The user must highlight and copy the entire GOBI record
	before starting the script. Copy everything, including the yellow box
	area for notes.

 Programs used: GOBI 3
					(Script tested in Firefox 3+ and IE 7+)
				Millennium Acquisitions Module JRE v 1.6.0_02
					(Main search screen open)


 Last revised:

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
#include <IE.au3>
#include <date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $GOBI_INFO, $GOBI_INFO_PREP, $GOBI_INFO_ARRAY_MASTER
Dim $WIN_TITLE
Dim $TITLE
Dim $ISBN
Dim $SERIES
Dim $decide
Dim $win
Dim $BROWSER
Dim $ISBN_N
Dim $AMAZON_A
Dim $FUND, $FUND_N, $FUND_S, $multi
Dim $DISCOUNT, $COST, $COST_L
Dim $REF_A, $REF_B, $REF
Dim $AV_A, $AV
Dim $HOLD_A, $HOLD
Dim $AC_A, $AC
Dim $MISSING_A, $MISSING
Dim $D
Dim $FUND_C
Dim $CALL_N, $DIFF_ED
Dim $msg
Dim $text
Dim $EBOOK
Dim $COST_A
Dim $SUPO, $MUPO
Dim $ORD_NOTE
Dim $SF

;################################ MAIN ROUTINE #################################
_ClearBuffer()

;ie work around part 1
$WIN_TITLE = WinGetTitle("GOBI 3")
$BROWSER = StringInStr($WIN_TITLE, "Mozilla Firefox")

;grab gobi info from clipboard and parse out into array
$GOBI_INFO = ClipGet()
$GOBI_INFO_PREP = StringRegExpReplace ($GOBI_INFO, "[\r\n\t]+", "fnord")
$GOBI_INFO_ARRAY_MASTER = StringSplit($GOBI_INFO_PREP, "fnord", 1)
;~ _ArrayDisplay($GOBI_INFO_ARRAY_MASTER)

;start setting individual variables from master array
$TITLE = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Title:")
$TITLE = StringTrimLeft($TITLE, 6)
;~ MsgBox(0, '', $TITLE)

$SERIES = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Series Title:")
$SERIES = StringTrimLeft($SERIES, 13)
;~ MsgBox(0, '', $SERIES)

$ISBN = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "ISBN:")
$ISBN = StringTrimLeft($ISBN, 5)
$ISBN = StringStripWS($ISBN, 8)
;~ MsgBox(0, '', $ISBN)


$FUND = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Fund Code:")
$FUND = StringTrimLeft($FUND, 10)
;~ MsgBox(0, '', $FUND)


$CALL_N = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "LC Class:")
$CALL_N = StringTrimLeft($CALL_N, 9)
;~ MsgBox(0, '', $CALL_N)


$EBOOK = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Binding:eBook")

;IE workaround part 2- cut and pasting text from IE different than in FF
If $BROWSER = 0 Then
	$ISBN_N = StringInStr($ISBN, "ISBN:")
	$ISBN = StringMid($ISBN, $ISBN_N, 18)
	$ISBN = StringTrimLeft($ISBN, 5)
	$FUND_N = StringInStr($FUND, "Fund Code:")
	$FUND_S = StringInStr($FUND, "/")
	If $FUND_S = 0 Then
		$FUND = StringMid($FUND, $FUND_N, 13)
		$FUND = StringTrimLeft($FUND, 10)
		$FUND = StringStripWS($FUND, 8)
	Else
		$FUND = StringTrimRight($FUND, 14)
		$FUND = StringTrimLeft($FUND, $FUND_N)
		$FUND = StringTrimLeft($FUND, 10)
	EndIf
	$COST_L = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "+ebrary Single User Option (SUPO)")
	$COST_A = StringRegExp($COST_L, "[\d]{1,3}[.\d]{3}", 1)
	$COST = _ArrayToString($COST_A, @TAB, 0, 0)
Else ;FF browser
	$COST_L = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Single User Option (SUPO)", 0, 0, 0, 1)
	$COST_A = $COST_L + 1 ;price is found next element down
	$COST = _ArrayToString($GOBI_INFO_ARRAY_MASTER, @TAB, $COST_A, $COST_A)
	$COST = StringRegExpReplace($COST, "(USD)", "")
EndIf
;end IE workaround

$COST = StringStripWS($COST, 8)
;~ MsgBox(0, '', $COST)

;single check
$SUPO = StringInStr($COST_L, "SUPO")

Select
	Case $SUPO > 0
		$ORD_NOTE = "s"
	CASE Else
		MsgBox(64, "SUPO not available", "The ebook does not have a SUPO option listed. The script will now exit.")
		Exit
EndSelect

;multiple funds check. Multiple funds in GOBI should be separated with a /
$FUND_S = StringInStr($FUND, "/")
$FUND_C = StringInStr($FUND, "/")
If $FUND_S > 0 or $FUND_C > 0 Then
 $multi = 1
 _StoreVar("$multi")
 Else
	 $multi = 0
EndIf


;~ #### start checking for various order notes #####



;holds
$HOLD_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Hold for", 0, 0, 0, 1)
If $HOLD_A > -1 Then
	$HOLD = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Order Notes 1:")
	$HOLD = StringTrimLeft($HOLD, 14)
	_StoreVar("$HOLD")
EndIf

;missing - means that we already have the record in the system (theoretically)
$MISSING_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Missing since", 0, 0, 0, 1)
If $MISSING_A > -1 Then
	$MISSING = 1
Else
	$MISSING = 0
EndIf

;start search in III
If WinExists("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]") Then
	WinActivate("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "nope", "Please log into Millennium.")
	Exit
EndIf

;start search with title
WinWaitActive("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]")
search($TITLE, "t")
Sleep(0400)

;if there is a "missing" message in gobi info, script will exit after III title search
If $MISSING = 1 Then
	Exit
EndIf

;if title search brings up a record, ask if it's an added copy/vol or diff ed.
If $win = 1 Then
	$decide = MsgBox(68, "Added copy/volume?", "Is this an added copy/volume?")
	Switch $decide
		Case 6
			Exit
		Case 7
	EndSwitch
	$decide = MsgBox(68, "Different edition?", "Is this a different edition?")
	Switch $decide
		Case 6
			$DIFF_ED = InputBox("Call number", "Please enter the diff ed. call number.", $CALL_N)
		Case 7
	EndSwitch
	WinClose("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitClose("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
	$win = 0
EndIf

#cs
input box asks user if they want to search by title, isbn, series, call number, or go straight to OCLC
hitting cancel button shuts down the script when existing records are found in system
#ce

While $win = 0
	choice()
	Switch $decide
		Case "C"
			search($CALL_N, "c")
		Case "I"
			search($ISBN, "i")
		Case "S"
			search($SERIES, "q")
		Case "T"
			search($TITLE, "t")
		Case "O"
			ExitLoop
		Case @error ;el cancel button
			Exit
	EndSwitch
WEnd

_createRecord()

;~ ########### local functions ##############
Func choice()
	$decide = InputBox("Search options", "What search do you want to perform next? Type in the corresponding letter." & @CR & "C: Call #"& @CR & "I: ISBN" & @CR & "S: Series" & @CR & "T: Title"  & @CR & "O: Create BIB/ORDER records", "", "", 200, 230)
	$decide = StringUpper($decide)
EndFunc

Func search ($x, $y)
	If $x <> "" Then
	WinActivate("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]")
	WinWaitActive("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]")
	Sleep(0100)
	_SendEx("!n")
	Sleep(0100)
	_SendEx("!n")
	Sleep(0100)
	_SendEx($y)
	Sleep(0100)
	_SendEx($x)
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(1100)
	$win = WinExists("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
		If $win = 1 And $y = "c" Then
			$decide = MsgBox(68, "Different edition?", "Is this a different edition?")
			Switch $decide
				Case 6
					$DIFF_ED = InputBox("Call number", "Please enter the diff ed. call number.", $CALL_N)
				Case 7
					WinClose("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
					WinWaitClose("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
					$win = 0
				EndSwitch
			$win = 0
		elseif $win = 1 and $y <>"c" Then
			$decide = MsgBox(68, "Added copy/volume?", "Is this an added copy/volume?")
				Switch $decide
					Case 6
						Exit
					Case 7
				EndSwitch
		EndIf
		$win = 0
	Else
		MsgBox(64, "Nothing to search", "GOBI did not supply the information needed to perform this search.")
	EndIf
EndFunc


Func _createRecord()
;~ 	MsgBox(0, '', $TITLE)
	Local $VENDOR, $DUP_CHECK
	Sleep(1000)
	_SendEx("^n")
	WinActivate("[TITLE:Add New Record; CLASS:SunAwtDialog]")
	WinWaitActive("[TITLE:Add New Record; CLASS:SunAwtDialog]")
	Sleep(0200)
	_SendEx("{DOWN 2}")
	Sleep(0200)
	_SendEx("onl")
	Sleep(0500)
	_SendEx("{TAB 2}")
	Sleep(0200)
	_SendEx("@")
	Sleep(0200)
	_SendEx("{UP}{LEFT}")
	Sleep(0200)
	_SendEx("m")
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	_SendEx("T24510" & $TITLE)
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	_SendEx("I020  " & $ISBN)
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	_SendEx("I05014" & $CALL_N)
	Sleep(0200)
	_SendEx("^s")
	Sleep(0500)
	$DUP_CHECK = WinExists("[TITLE:Perform duplicate checking?; CLASS:SunAwtDialog]")
	If $DUP_CHECK = 1 Then
		_SendEx("n")
	EndIf
	WinActivate("[TITLE:New ORDER; CLASS:SunAwtFrame]", "")
	WinWaitActive("[TITLE:New ORDER; CLASS:SunAwtFrame]", "")
	_SendEx("{DOWN}")
	$VENDOR = InputBox ("Vendor", "Enter vendor code for item", "ybp")
	WinActivate("[TITLE:New ORDER; CLASS:SunAwtFrame]", "")
	WinWaitActive("[TITLE:New ORDER; CLASS:SunAwtFrame]", "")
	_SendEx("onlbk")
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx($COST)
	Sleep(0100)
	_SendEx("{TAB}")
	Sleep(0100)
	_SendEx("b")
	Sleep(0100)
	_SendEx("{TAB 3}")
	Sleep(0100)
	_SendEx("e")
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx("b")
	Sleep(0100)
	_SendEx("{TAB 4}")
	Sleep(0100)
	$SF = StringRegExp($FUND, "4[\D]+")
	If $SF = 1 Then
		_SendEx($FUND)
	ElseIf $SF = 0 Then
		_SendEx($FUND & "/1")
	EndIf
	Sleep(0100)
	_SendEx("{TAB 10}")
	Sleep(0100)
	_SendEx("s")
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx($VENDOR)
	Sleep(0100)
	_SendEx("{TAB 4}")
	Sleep(0100)
	_SendEx("b")
	Sleep(0100)
	Exit
EndFunc

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: GOBItoIII-OCLC
 Script set: Order (order)

 Script Function:
	This script takes the information copied from the GOBI record and searches
	Amazon via API. Depending on if there is a match, an IE window to the
	Amazon product may or may not open. Regardless of the Amazon search
	outcome, the script searches Millennium by title (automatically) and by
	other information. If the item is not in III, then the script searches
	OCLC Connexion by ISBN. Otherwise, if the item is in III, it will prompt
	a check to determine if it is an added copy/volume order.

	##### NOTE ##### The user must highlight and copy the entire GOBI record
	before starting the script. Copy everything, including the yellow box
	area for notes.

 Programs used: GOBI 3
					(Script tested in Firefox 3+ and IE 7+)
				Millennium Acquisitions Module JRE v 1.6.0_02
					(Main search screen open)
				OCLC Connexion Client v 2.20
					(logged in)

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
Dim $PROFILE
Dim $DATE_P, $PROFILE_A, $DATE_D, $20, $16, $COST_D
Dim $D
Dim $FUND_C, $UK
Dim $CALL_N, $DIFF_ED
Dim $oIE, $GUIActiveX
Dim $msg
Dim $text
Dim $APRICE_A, $APRICE, $DIFF, $MARKETPLACE, $close, $AMAZON_SEARCH, $AMAZON_L

;################################ MAIN ROUTINE #################################
_ClearBuffer()

;ie work around part 1
$WIN_TITLE = WinGetTitle("GOBI 3")
$BROWSER = StringInStr($WIN_TITLE, "Mozilla Firefox")

;grab gobi info from clipboard and parse out into array
$GOBI_INFO = ClipGet()
$GOBI_INFO_PREP = StringRegExpReplace ($GOBI_INFO, "[\r\n\t]+", "fnord")
$GOBI_INFO_ARRAY_MASTER = StringSplit($GOBI_INFO_PREP, "fnord", 1)

;start setting individual variables from master array
$TITLE = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Title:")
$TITLE = StringTrimLeft($TITLE, 6)

$SERIES = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Series Title:")
$SERIES = StringTrimLeft($SERIES, 13)

$ISBN = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "ISBN:")
$ISBN = StringTrimLeft($ISBN, 5)
$ISBN = StringStripWS($ISBN, 8)

$FUND = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Fund Code:")
$FUND = StringTrimLeft($FUND, 10)
;store fund variable for future use
_StoreVar("$FUND")

$CALL_N = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "LC Class:")
$CALL_N = StringTrimLeft($CALL_N, 9)

;figure out which cost to use - search for slip discount
$DISCOUNT = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Slip Discount:", 0, 0, 0, 1)
$UK = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Est. US Net:", 0, 0, 0, 1)
If $DISCOUNT > -1  And $UK > -1 Then ;if discount was already figured in
	$COST = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Est. US Net:")
	$COST = StringTrimLeft($COST, 12)
	$COST = StringReplace($COST, "USD", "")
	$COST = StringStripWS($COST, 8)
Else ;if discount was not figured in
	$COST = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "US List:")
	$COST = StringTrimLeft($COST, 8)
	; START FIGURING 20% OR 16% DISCOUNT
	$PROFILE = _arrayItemString($GOBI_INFO_ARRAY_MASTER, "Handled On Approval YBP-US:")
	IF $PROFILE > -1 Then ;start preparing profile date to determine how long item was profiled
		$PROFILE = StringReplace($PROFILE, "Handled On Approval YBP-US:", "")
		$PROFILE = StringLeft($PROFILE, 10)
		$PROFILE_A = StringSplit($PROFILE, "/")
		If $PROFILE_A[1] < 10 Then
			$PROFILE_A[1] = "0" & $PROFILE_A[1]
		EndIf
		If $PROFILE_A[2] < 10 Then
			$PROFILE_A[2] = "0" & $PROFILE_A[2]
		EndIf
		$PROFILE = $PROFILE_A[3] & "/" & $PROFILE_A[1] & "/" & $PROFILE_A[2]
		$DATE_P  =  @YEAR & "/" & @MON	& "/" & @MDAY
		$PROFILE = StringStripWS($PROFILE, 8)
		$DATE_D = _DateDiff('Y', $PROFILE, $DATE_P)
		If $DATE_D >= 1 Then
			$16 = 1 ;been profiled for more than a year
		Else
			$20 = 1 ;been profiled for less than a year
		EndIf
	Else
		$16 = 1 ;not profiled
	EndIf
EndIf
_StoreVar("$COST")

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
	_StoreVar("$FUND")
	$COST_L = StringInStr($COST, "US Status:")
	If $COST_L > 0 Then
		$COST = StringLeft($COST, 5)
		_StoreVar("$COST")
	EndIf
Else ;firefox
	If $DISCOUNT = -1 Then
		$COST= StringTrimRight($COST, 3)
		_StoreVar("$COST")
	EndIf
EndIf
;end IE workaround

MsgBox(0, "", $FUND)


;Modify cost variable for discount percentage if appropirate
If $DISCOUNT = -1 Then
	$COST = Number($COST)
	Select
		Case $20 = 1
			$D = 0.2
			$COST_D = $COST * $D
			$COST_D = StringFormat("%.2f",$COST_D)
			$COST = $COST - $COST_D
			$COST = StringFormat("%.2f",$COST)
		Case $16 = 1
			$D = 0.16
			$COST_D = StringFormat("%.2f",$COST_D)
			$COST_D = $COST * $D
			$COST = $COST - $COST_D
			$COST = StringFormat("%.2f",$COST)
	EndSelect
	_StoreVar("$COST")
EndIf

;multiple funds check. Multiple funds in GOBI should be separated with a /
$FUND_S = StringInStr($FUND, "/")
$FUND_C = StringInStr($FUND, "/")
If $FUND_S > 0 or $FUND_C > 0 Then
 $multi = 1
 _StoreVar("$multi")
 Else
	 $multi = 0
EndIf

;##### start checking for various order notes #####
;Reference
$REF_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Cat for King ref", 0, 0, 0, 1)
$REF_B = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "catalog for general reference", 0, 0, 0, 1)
If $REF_A > -1 OR $REF_B > -1 Then
	$REF = 1
Else
	$REF = 0
EndIf
_StoreVar("$REF")

;added volume
$AV_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Added vol", 0, 0, 0, 1)
If $AV_A > -1 Then
	$AV = 1
Else
	$AV = 0
EndIf

;added copy
$AC_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Added copy", 0, 0, 0, 1)
If $AC_A > -1 Then
	$AC = 1
Else
	$AC = 0
EndIf

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

;check amazon items in amazon. does a keyword search
$AMAZON_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Order Notes 1:amazon", 0, 0, 0, 1)
IF $AMAZON_A > -1 Then
_IECreate ("http://www.amazon.com/s/ref=nb_ss?url=search-alias%3Daps&field-keywords=" & $ISBN & "&x=0&y=0")
EndIf
;##### end checking for various order notes #####

$ISBN = StringStripWS($ISBN, 8)
;~ $ISBN = "0380711648" ;Marketplace test
;~ $ISBN = "9780765680785" ;Marketplace test - no new books

;~ ##### GUI FOR AMAZON API SEARCH #####
$oIE = _IECreateEmbedded ()
GUICreate("Embedded Web control Test", 200, 200, _
		(@DesktopWidth - 640) / 2, (@DesktopHeight - 580) / 2)
$GUIActiveX = GUICtrlCreateObj($oIE, 10, 10, 150, 150)

GUISetState(@SW_HIDE) ;window does not appear to user.

_IENavigate ($oIE, "http://techserv.lib.muohio.edu/~yoosebj/amazon.php?isbn=" & $ISBN)

;~ ##### END GUI code #####

;the API php script has certain text that it brings back from its search
;the script checks for certain terms to determine where the item is available,
;and if the item is cheaper in Amazon than in YBP.
$text = _IEDocReadHTML($oIE)
$APRICE_A = StringRegExp($text, "[$0-9.]{2,}[0-9]{2}", 1)
$APRICE = _ArrayToString($APRICE_A, @TAB, 0, 0)
$APRICE = StringRegExpReplace($APRICE, "[$]", "")
$DIFF =  $COST - $APRICE
If $DIFF > 0 Then
	$MARKETPLACE = StringInStr($text, "Marketplace")
	If $MARKETPLACE > 0 Then
		If $APRICE = "" Then
			$AMAZON_SEARCH = 0
		Else
			$APRICE = $APRICE + 3.99 ;3.99 is S+H
			$AMAZON_L = "Amazon Marketplace"
			$AMAZON_SEARCH = 1
		EndIf
	Else
		$AMAZON_L = "Amazon.com"
		$AMAZON_SEARCH = 1
	EndIf
ElseIf $DIFF <= 0 Then
	$close = 1
EndIf

If $AMAZON_SEARCH = 1 Then ;amazon is cheaper
	$decide = MsgBox(68, "Search Amazon?", $AMAZON_L & " lists a lower price at " & $APRICE & "." & @CR & "Do you want to go to the Amazon item page?")
	Switch $decide
		Case 6
		$COST = $APRICE ;Set price to amazon's price
		_StoreVar("$COST")
		_StoreVar("$AMAZON_SEARCH")
		_IECreate ("http://www.amazon.com/s/ref=nb_ss?url=search-alias%3Daps&field-keywords=" & $ISBN & "&x=0&y=0", 0, 1, 0)
		Case 7
	EndSwitch
	$close = 1 ;closes hidden gui window
Else
	$close = 1 ;closes hidden gui window
EndIf

;~ ##### start gui close####
While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
		Case $close = 1
			ExitLoop
	EndSelect
WEnd
GUIDelete()
;~ ##### end gui close####


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

;option to create order record for added copy/vol
If $AV = 1 Or $AC = 1 Then
	$decide = MsgBox(68, "Added copy/volume?", "Would you like to create an order record for this added copy/vol?")
	Switch $decide
		Case 6
			_createACVrecord()
		Case 7
			Exit
	EndSwitch
EndIf

;if title search brings up a record, ask if it's an added copy/vol or diff ed.
If $win = 1 Then
	$decide = MsgBox(68, "Added copy/volume?", "Is this an added copy/volume?")
	Switch $decide
		Case 6
			_createACVrecord()
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
		Case "R"
			$DIFF_ED = InputBox("Call number", "Please enter the diff ed. call number.", $CALL_N)
		Case @error ;el cancel button
			Exit
	EndSwitch
WEnd

_OCLCSearch("bn:", $ISBN)
_StoreVar("$DIFF_ED")

;~ ########### local functions ##############
Func choice()
	$decide = InputBox("Search options", "What search do you want to perform next? Type in the corresponding letter." & @CR & "C: Call #"& @CR & "I: ISBN" & @CR & "S: Series" & @CR & "T: Title"  & @CR & "O: OCLC" & @CR & "R: Record Diff Ed. call #", "", "", 200, 230)
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
	$win = WinExists("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
		If $win = 1 And $y = "c" Then
			$decide = MsgBox(68, "Different edition?", "Is this a different edition?")
			Switch $decide
				Case 6
					$DIFF_ED = InputBox("Call number", "Please enter the diff ed. call number.", $CALL_N)
				Case 7
					WinClose("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
					WinWaitClose("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
					$win = 0
				EndSwitch
			$win = 0
		elseif $win = 1 and $y <>"c" Then
			$decide = MsgBox(68, "Added copy/volume?", "Is this an added copy/volume?")
				Switch $decide
					Case 6
						_createACVrecord()
					Case 7
				EndSwitch
		EndIf
		$win = 0
	Else
		MsgBox(64, "Nothing to search", "GOBI did not supply the information needed to perform this search.")
	EndIf
EndFunc


Func _createACVrecord()
	Local $LOCATION, $VENDOR
	WinActivate("[REGEXPTITLE:\A[bio][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[bio][0-9ax]{8}; CLASS:SunAwtFrame]")
	Sleep(1000)
	_SendEx("!g")
	Sleep(0200)
	_SendEx("o")
	Sleep(0200)
	_SendEx("a")
	Sleep(0500)
	_SendEx("!w")
	Sleep(0200)
	_SendEx("o")
	Sleep(0200)
	_SendEx("!n")
	Sleep(0200)
	If WinExists("[TITLE:Options; CLASS:SunAwtDialog]") Then
		MsgBox(64, "Options/template choices", "Click OK after selecting appropriate choices for order record options and template.")
	EndIf
	WinActivate("[TITLE:New ORDER; CLASS:SunAwtDialog]", "")
	WinWaitActive("[TITLE:New ORDER; CLASS:SunAwtDialog]", "")
	_SendEx("{DOWN}")
	$LOCATION = InputBox ("Location", "Enter location code for item")
	$VENDOR = InputBox ("Vendor", "Enter vendor code for item", "ybp")
	WinActivate("[REGEXPTITLE:\A[bio][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[bio][0-9ax]{8}; CLASS:SunAwtFrame]")
	_SendEx($LOCATION)
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx($COST)
	Sleep(0100)
	_SendEx("{TAB}")
	Sleep(0100)
	_SendEx("n")
	Sleep(0100)
	_SendEx("{TAB 5}")
	Sleep(0100)
	_SendEx("b")
	Sleep(0100)
	_SendEx("{TAB 4}")
	Sleep(0100)
	_SendEx($FUND)
	Sleep(0100)
	_SendEx("{TAB 12}")
	Sleep(0100)
	_SendEx($VENDOR)
	Sleep(0100)
	_SendEx("^{END}")
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("z")
	Sleep(0100)
	_SendEx("ADDED ")
	Exit
EndFunc

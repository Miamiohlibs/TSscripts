AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)


#include <Array.au3>
#Include <Misc.au3>
#include <IE.au3> 
#include <date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <SaveTransferVariables.au3>

Dim $GOBI_INFO, $GOBI_INFO_PREP, $GOBI_INFO_ARRAY_MASTER, $WIN_TITLE
Dim $TITLE
Dim $ISBN
Dim $SERIES
Dim $decide, $win, $BROWSER
Dim $ISBN_N, $AMAZON_A, $multi
Dim $FUND, $FUND_N, $FUND_S, $multi, $DISCOUNT, $COST, $COST_L
Dim $REF_A, $REF_B, $REF, $AV_A, $AV, $HOLD_A, $HOLD, $AC_A, $AC, $MISSING_A, $MISSING
Dim $PROFILE, $DATE_P, $PROFILE_A, $DATE_D, $20, $16, $COST_D, $D, $FUND_C, $UK, $CALL_N, $DIFF_ED
Dim $oIE, $GUIActiveX
Dim $msg, $text, $APRICE_A, $APRICE, $DIFF, $MARKETPLACE, $close, $AMAZON_SEARCH, $AMAZON_L

;stupid ie work around
$WIN_TITLE = WinGetTitle("GOBI 3")
$BROWSER = StringInStr($WIN_TITLE, "Mozilla Firefox")


;grab gobi info from clipboard and parse out into array
$GOBI_INFO = ClipGet()
$GOBI_INFO_PREP = StringRegExpReplace ($GOBI_INFO, "[\r\n\t]+", "fnord")
$GOBI_INFO_ARRAY_MASTER = StringSplit($GOBI_INFO_PREP, "fnord", 1)
;~ _ArrayDisplay($GOBI_INFO_ARRAY_MASTER)

;start setting individual variables from master array
$TITLE = arrayItemString($GOBI_INFO_ARRAY_MASTER, "Title:")
$TITLE = StringTrimLeft($TITLE, 6)

$SERIES = arrayItemString($GOBI_INFO_ARRAY_MASTER, "Series Title:")
$SERIES = StringTrimLeft($SERIES, 13)

$ISBN = arrayItemString($GOBI_INFO_ARRAY_MASTER, "ISBN:")
$ISBN = StringTrimLeft($ISBN, 5)
$ISBN = StringStripWS($ISBN, 8)

$FUND = arrayItemString($GOBI_INFO_ARRAY_MASTER, "Fund Code:")
$FUND = StringTrimLeft($FUND, 10)

$CALL_N = arrayItemString($GOBI_INFO_ARRAY_MASTER, "LC Class:")
$CALL_N = StringTrimLeft($CALL_N, 9)

StoreVar("$FUND")
;~ MsgBox(0, "", $FUND)

;figure out which cost to use - search for slip discount
$DISCOUNT = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Slip Discount:", 0, 0, 0, 1)
$UK = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Est. US Net:", 0, 0, 0, 1)
;~ MsgBox(64, "", $DISCOUNT)
If $DISCOUNT > -1  And $UK > -1 Then
	$COST = arrayItemString($GOBI_INFO_ARRAY_MASTER, "Est. US Net:")
	$COST = StringTrimLeft($COST, 12)
	$COST = StringReplace($COST, "USD", "")
	$COST = StringStripWS($COST, 8)
Else
	$COST = arrayItemString($GOBI_INFO_ARRAY_MASTER, "US List:")
	$COST = StringTrimLeft($COST, 8)
	
	; START FIGURING 20% OR 16% DISCOUNT

	$PROFILE = arrayItemString($GOBI_INFO_ARRAY_MASTER, "Handled On Approval YBP:")
	IF $PROFILE <> "" Then
		$PROFILE = StringReplace($PROFILE, "Handled On Approval YBP:", "")
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
			$16 = 1
;~ 			MsgBox(0, "", "hi")
		Else
			$20 = 1
;~ 			MsgBox(0, "", "yo")
		EndIf
	EndIf
EndIf
StoreVar("$COST")
;~ MsgBox(0, "", $COST)


;stupid IE workaround - cut and pasting text from IE different than in FF

If $BROWSER = 0 Then 
	$ISBN_N = StringInStr($ISBN, "ISBN:")
	$ISBN = StringMid($ISBN, $ISBN_N, 18)
	$ISBN = StringTrimLeft($ISBN, 5)
;~ 	MsgBox(0, "", $ISBN)

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
	StoreVar("$FUND")
;~ 	MsgBox(0, "", $FUND)

	$COST_L = StringInStr($COST, "US Status:")
	If $COST_L > 0 Then
		$COST = StringLeft($COST, 5)
		StoreVar("$COST")
	EndIf
;~ 	MsgBox(0, "", $COST)
Else
	If $DISCOUNT = -1 Then
		$COST = StringTrimRight($COST, 3)
		StoreVar("$COST")
	EndIf
EndIf
;end stupid IE workaround


;calculate discount if not already figured in
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
	StoreVar("$COST")
EndIf


;multiple funds check. Multiple funds in GOBI should be separated with a /
$FUND_S = StringInStr($FUND, "/")
$FUND_C = StringInStr($FUND, "/")
If $FUND_S > 0 or $FUND_C > 0 Then
 $multi = 1
 StoreVar("$multi")
 Else
	 $multi = 0
EndIf

;start checking for various order notes

;Reference
$REF_A = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "Cat for King ref", 0, 0, 0, 1)
$REF_B = _ArraySearch($GOBI_INFO_ARRAY_MASTER, "catalog for general reference", 0, 0, 0, 1)
If $REF_A > -1 OR $REF_B > -1 Then
	$REF = 1
Else
	$REF = 0
EndIf
;~ MsgBox(0, "", $REF)
StoreVar("$REF")

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
	$HOLD = arrayItemString($GOBI_INFO_ARRAY_MASTER, "Order Notes 1:")
	$HOLD = StringTrimLeft($HOLD, 14)
	StoreVar("$HOLD")
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

$ISBN = StringStripWS($ISBN, 8)
;~ $ISBN = "0380711648" ;Marketplace test

;~ ####################### GUI TOOL #############################
_IEErrorHandlerRegister ()

$oIE = _IECreateEmbedded ()
GUICreate("Embedded Web control Test", 200, 200, _
		(@DesktopWidth - 640) / 2, (@DesktopHeight - 580) / 2)
$GUIActiveX = GUICtrlCreateObj($oIE, 10, 10, 150, 150)

GUISetState(@SW_HIDE)

_IENavigate ($oIE, "http://techserv.lib.muohio.edu/~yoosebj/amazon.php?isbn=" & $ISBN)
; Waiting for user to close the window

;~ ############################## END GUI ########################

$COST = "29.30"

$text = _IEDocReadHTML($oIE)
;~ MsgBox(0, "", $text)
$APRICE_A = StringRegExp($text, "[$0-9.]{2,}[0-9]{2}", 1)
;~ _ArrayDisplay($APRICE_A)
$APRICE = _ArrayToString($APRICE_A, @TAB, 0, 0)
$APRICE = StringRegExpReplace($APRICE, "[$]", "")
MsgBox(0, "", $APRICE & " and " & $COST)
$DIFF =  $COST - $APRICE
If $DIFF > 0 Then
	$MARKETPLACE = StringInStr($text, "Marketplace")
	If $MARKETPLACE > 0 Then
		$APRICE = $APRICE + 3.99 ;3.99 is S+H
		$AMAZON_L = "Amazon Marketplace"
		$AMAZON_SEARCH = 1
;~ 		$close = MsgBox (0, "", "Amazon Marketplace is cheaper")
	Else	
;~ 		$close= MsgBox (0, "", "Amazon's cheaper")
		$AMAZON_L = "Amazon.com"
		$AMAZON_SEARCH = 1
	EndIf
ElseIf $DIFF <= 0 Then 
;~ 	$close = MsgBox (0, "", "YBP's cheaper")
	$close = 1
EndIf
;~ MsgBox(0, "", $DIFF)


If $AMAZON_SEARCH = 1 Then
	$decide = MsgBox(68, "Search Amazon?", $AMAZON_L & " lists a lower price at " & $APRICE & "." & @CR & "Do you want to go to the Amazon item page?")
	Switch $decide
		Case 6
		$COST = $APRICE ;Set price to amazon's price
		StoreVar("$COST")
		StoreVar("$AMAZON_SEARCH")
		_IECreate ("http://www.amazon.com/s/ref=nb_ss?url=search-alias%3Daps&field-keywords=" & $ISBN & "&x=0&y=0", 0, 1, 0)
		Case 7
	EndSwitch
	$close = 1
EndIf

;~ #####gui close####
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
;~ #####gui close####




;start search in III
If WinExists("[TITLE:Millennium Acquisitions; CLASS:com.iii.milmainpanel$3]") Then
	WinActivate("[TITLE:Millennium Acquisitions; CLASS:com.iii.milmainpanel$3]")
Else 
	MsgBox(64, "nope", "Please log into Millennium.")
	Exit

EndIf

;start search with title
WinWaitActive("[TITLE:Millennium Acquisitions; CLASS:com.iii.milmainpanel$3]")
search($TITLE, "t")
Sleep(0400)

;if there is a "missing/added copy/added vol" message in gobi info, script will exit after III title search
If $MISSING = 1 OR $AV = 1 Or $AC = 1 Then
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
	WinClose("[REGEXPTITLE:[bo]; CLASS:com.iii.ShowRecord$2]")
	WinWaitClose("[REGEXPTITLE:[bo]; CLASS:com.iii.ShowRecord$2]")
	$win = 0
EndIf

;input box asks user if they want to search by title, isbn, series, call number, or go straight to OCLC
;hitting cancel button shuts down the script when existing records are found in system
$win = WinActive("[CLASS:com.iii.ShowRecord$2]")
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


If WinExists("OCLC Connexion") Then
	WinActivate("OCLC Connexion")
Else 
	MsgBox(64, "nope", "Please open and log into Connexion. Click ok when you are logged in")
EndIf
WinWaitActive("OCLC Connexion")
Sleep(0100)
_SendEx("{F2}")
Sleep(0100)
_SendEx("bn:")
Sleep(0100)
_SendEx($ISBN)
Sleep(0100)
_SendEx("{ENTER}")


StoreVar("$DIFF_ED")


;~ ########### functions ##############

Func arrayItemString($y, $x)
	
	Local $z
	Local $w
	$w = _ArraySearch($y, $x, 0, 0, 0, 1)
	If $w > -1 Then
		$z = _ArrayToString($y, @TAB, $w, $w)
		Return $z
	EndIf
	
EndFunc



Func choice()

	$decide = InputBox("Search options", "What search do you want to perform next? Type in the corresponding letter." & @CR & "C: Call #"& @CR & "I: ISBN" & @CR & "S: Series" & @CR & "T: Title"  & @CR & "O: OCLC" & @CR & "R: Record Diff Ed. call #", "", "", 200, 230)
	$decide = StringUpper($decide)
			
EndFunc
	


Func search ($x, $y)
	If $x <> "" Then	
	WinActivate("[TITLE:Millennium Acquisitions; CLASS:com.iii.milmainpanel$3]")
	WinWaitActive("[TITLE:Millennium Acquisitions; CLASS:com.iii.milmainpanel$3]")
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
	Sleep(1000)
	$win = WinExists("[REGEXPTITLE:[bo]; CLASS:com.iii.ShowRecord$2]")
		If $win = 1 And $y = "c" Then
			$decide = MsgBox(68, "Different edition?", "Is this a different edition?")
			Switch $decide
				Case 6 
					$DIFF_ED = InputBox("Call number", "Please enter the diff ed. call number.", $CALL_N)
				Case 7
					WinClose("[REGEXPTITLE:[bo]; CLASS:com.iii.ShowRecord$2]")
					WinWaitClose("[REGEXPTITLE:[bo]; CLASS:com.iii.ShowRecord$2]")
					$win = 0
			EndSwitch
		EndIf
	Else
		MsgBox(64, "Nothing to search", "GOBI did not supply the information needed to perform this search.") 
	EndIf
EndFunc

;########### fuction to fix sticking mod keys #############
Func _SendEx($ss)
    Local $iT = TimerInit()
    
    While _IsPressed("10") Or _IsPressed("11") Or _IsPressed("12")
        If TimerDiff($iT) > 1000 Then
            MsgBox(262144, "Warning", "Shift, Ctrl and Alt Keys need to be released to proceed!" & @CR & "Click OK to release keys.")
			_SendEx("{CTRLUP}")
			_SendEx("{SHIFTUP}")
			_SendEx("{ALTUP}")
		EndIf
    WEnd
    Send($ss)
    
EndFunc ;==>_SendEx

;############## end function to fix sticking mod keys #########















;~ _IECreate ("http://www.amazon.com/s/ref=nb_ss?url=search-alias%3Daps&field-keywords=" & $ISBN & "&x=0&y=0")

;~ _IECreate ("http://ecs.amazonaws.com/onca/xml?AWSAccessKeyId=AKIAIWTC6E4GYVJRVALA&IdType=ISBN&ItemId=" & $ISBN & "&Operation=ItemLookup&ResponseGroup=Offers&SearchIndex=Books&Service=AWSECommerceService&Timestamp=2010-03-02T20%3A52%3A17.000Z&Version=2009-03-31&Signature=2SN7QtyMzpIuk45Qi3LTL3P0hZUsyiWf%2BI4oMUdyHc4%3D")


;~ this works - unsigned
;~ access key AKIAIWTC6E4GYVJRVALA
;~ secret key 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
;~ http://ecs.amazonaws.com/onca/xml?Service=AWSECommerceService
;~ &Version=2009-03-31
;~ &Operation=ItemLookup
;~ &SearchIndex=Books
;~ &IdType=ISBN
;~ &ItemId=9780393068337
;~ &ResponseGroup=OfferSummary <--gives lowest new price





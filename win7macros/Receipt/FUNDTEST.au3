#Include <TSCustomFunction.au3>

;~ $ORD_REC = ClipGet()
;~ $ORD_REC_PREP = StringRegExpReplace ($ORD_REC, "[\r\n]+", "fnord")
;~ $ORD_ARRAY_MASTER = StringSplit($ORD_REC_PREP, "fnord", 1)


;~ ;row 4 includes claim, fund, status
;~ $row4_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 4)
;~ $STATUS = _ArrayPop($row4_MASTER)
;~ $FUND = _ArrayPop($row4_MASTER)
;~ $CLAIM = _ArrayPop($row4_MASTER)
;~ $STATUS = StringTrimLeft($STATUS, 6)
;~ $FUND = StringTrimLeft($FUND, 4)
;~ $CLAIM = StringTrimLeft($CLAIM, 6)



;~ MsgBox(0, '', $FUND)

;~ ;matches for slash, parse out fund number
;~ $slash = StringInStr($FUND, "/")
;~ If $slash > 0 Then
;~ 	$fund_1 = StringRegExp($FUND, "(\w{1,5}/?\d?)", 1)
;~ 	_ArrayDisplay($fund_1)
;~ 	$slash_fund = _ArrayPop($fund_1)
;~ 	MsgBox(0, '', $slash_fund)
;~ EndIf

;~ ;now for the fun part! special fund codes
;~ $4 = StringLeft($FUND, 1)
;~ If $slash = 0 And $4 = 4 Then
;~ 	$fund_1 = StringRegExp($FUND, "(4\D{1,3})", 1)
;~ 	_ArrayDisplay($fund_1)
;~ 	$slash_fund = _ArrayPop($fund_1)
;~ 	MsgBox(0, '', $slash_fund)
;~ EndIf

;~ ;fun times with 61-64 codes
;~ $6X = StringLeft($FUND, 2)
;~ If $slash = 0 and ($6X = 61 OR $6X = 62 Or $6X = 63 Or $6X = 64 Or $6X = "7h" or $6X = "7s" or $6X = "8c" Or $6X = "8h" Or $6X = "8s") Then
;~ 	$slash_fund = StringLeft($FUND, 5)
;~ 	MsgBox(0, '', $slash_fund)
;~ EndIf

$multi_loc = "kngr (1),docr (2)"
$printerarray = StringSplit($multi_loc, ",")
;~ _ArrayDisplay($multi_a)
;~ Do
;~ 			$LOCATION = _ArrayPop($multi_A)
;~ 			$LOC = $LOCATION
;~ 			$LOCATION = StringRegExpReplace($LOCATION, "[(\d)]", "")
;~ 			$LOCATION = StringStripWS($LOCATION, 8)
;~ 			MsgBox(0, '', $LOCATION)
;~ until $multi_a = ""


;~ $printerarray = $multi_a

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
Opt("GUIOnEventMode", 1)


; create the GUI
$hGUI = GUICreate("Multiple Item Locations", 220, 130, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))

GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")
GUICtrlCreateLabel("Select item location for next item.", 10, 10)
$hCombo = GUICtrlCreateCombo("", 10, 50, 200, 20)
$1 = GUICtrlCreateButton("OK", 80, 90, 50, 30)
GUICtrlSetOnEvent($1, "loc")


; now use the array you have just generated to populate the combo
$sList = "|"
For $i = 1 To UBound($printerarray) -1
    $sList &= $printerarray[$i] & "|"
Next

GUICtrlSetData($hCombo, $sList)
$close = 0
$sCurrSel = ""
GUISetState(@SW_SHOW)
While $close = 0
	Sleep(1000)
	If $close = 1 Then
		GUIDelete($hGUI)
		ExitLoop
	EndIf
WEnd


Func loc()
	$sCurrSel = GUICtrlRead($hCombo)
	$iIndex = _ArraySearch($printerarray, $sCurrSel, 0, 0, 0, 1)
	$sText_1 = $printerarray[$iIndex]
	ClipPut($sText_1)
	$close = 1
EndFunc


Func CLOSEClicked()
  Exit
EndFunc


;~ Exit
;~ While 1

;~     Switch GUIGetMsg()
;~         Case $GUI_EVENT_CLOSE
;~             Exit
;~     EndSwitch

;~     If GUICtrlRead($hCombo) <> $sCurrSel Then
;~         ; Set new selection value
;~         $sCurrSel = GUICtrlRead($hCombo)
;~         ; Search for the value in the array
;~         $iIndex = _ArraySearch($printerarray, $sCurrSel, 0, 0, 0, 1)
;~         ; Read the other array elements
;~         $sText_1 = $printerarray[$iIndex]
;~         ; Show the data
;~         GUICtrlSetData($hLabel, $sText_1)
;~     EndIf

;~ WEnd

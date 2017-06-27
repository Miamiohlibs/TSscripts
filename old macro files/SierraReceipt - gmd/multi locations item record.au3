#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Include <TSCustomFunction.au3>
$multi_loc= _LoadVar("$multi_loc")
;~ $multi_loc = "kngr (1),docr (2)"
$multi_loc_a = StringSplit($multi_loc, ",")
;~ _ArrayDisplay($multi_loc_a)


;######### GUI #########
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
Opt("GUIOnEventMode", 1)
$hGUI = GUICreate("Multiple Item Locations", 220, 130, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")
GUICtrlCreateLabel("Select item location for next item.", 10, 10)
$hCombo = GUICtrlCreateCombo("", 10, 50, 200, 20)
$1 = GUICtrlCreateButton("OK", 80, 90, 50, 30)
GUICtrlSetOnEvent($1, "loc")
$sList = "|"
For $i = 1 To UBound($multi_loc_a) -1
	$sList &= $multi_loc_a[$i] & "|"
Next
GUICtrlSetData($hCombo, $sList)
$sCurrSel = ""
$close = 0
GUISetState(@SW_SHOW)
While $close = 0
	Sleep(1000)
	If $close = 1 Then
		GUIDelete($hGUI)
	EndIf
WEnd

;######### start GUI functions #########
Func loc()
	$sCurrSel = GUICtrlRead($hCombo)
	$iIndex = _ArraySearch($multi_loc_a, $sCurrSel, 0, 0, 0, 1)
	$sText_1 = $multi_loc_a[$iIndex]
	ClipPut($sText_1)
	$close = 1
EndFunc

Func CLOSEClicked()
  Exit
EndFunc
;######### end GUI functions #########
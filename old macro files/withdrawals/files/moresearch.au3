#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\Images\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#Include <String.au3>
#include <Array.au3>
#Include <Misc.au3>
#Include <TSCustomFunction.au3>


Dim $winsize
Dim $decision
Dim $barcode
Dim $BIB_ARRAY_MASTER
Dim $BIB_REC
Dim $BIB_REC_PREP
Dim $OCLC_NUM, $OCLC_NUM_A
Dim $oIE_1, $oIE_2
Dim $q
dim $GUIActiveX_1, $GUIActiveX_2



$decision = inputbox("Search for what?", "Enter o for OCLC# search, i for ISBN search, or t for Title search")
If $decision = "o" Then


$q = InputBox("OCLC#", "Enter the OCLC number.")


;~ $OCLC_NUM = "45487059"


;_IEErrorHandlerRegister()

$oIE_1 = _IECreateEmbedded ()
$oIE_2 = _IECreateEmbedded ()
GUICreate("SWORD Book Search", 1200, 800,0,0, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
$GUIActiveX_1 = GUICtrlCreateObj($oIE_1, 10, 40, 580, 700)
$GUIActiveX_2 = GUICtrlCreateObj($oIE_2, 590, 40, 580, 700)


;GUISetState(@SW_SHOW)       ;Show GUI
WinSetState("SWORD Book Search", "", @SW_SHOW)

_IENavigate ($oIE_1, "http://catalog.libraries.wright.edu/search/?searchtype=o&searcharg=" & $q)
_IENavigate ($oIE_2, "http://uclid.uc.edu/search/o?SEARCH=" & $q)

; Waiting for user to close the window
While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
	EndSelect
WEnd

GUIDelete()

Exit
ElseIf $decision = "t" Then
	$q = inputbox("Title", "Please enter the title to search by.")

$oIE_1 = _IECreateEmbedded ()
$oIE_2 = _IECreateEmbedded ()
GUICreate("SWORD Book Search", 1200, 800,0,0, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
$GUIActiveX_1 = GUICtrlCreateObj($oIE_1, 10, 40, 580, 700)
$GUIActiveX_2 = GUICtrlCreateObj($oIE_2, 590, 40, 580, 700)


;GUISetState(@SW_SHOW)       ;Show GUI
WinSetState("SWORD Book Search", "", @SW_SHOW)

	_IENavigate ($oIE_1, "http://catalog.libraries.wright.edu/search/?searchtype=t&searcharg=" & $q)
	_IENavigate ($oIE_2, "http://uclid.uc.edu/search/t?SEARCH=" & $q)
	While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
	EndSelect
WEnd

GUIDelete()
	Exit
ElseIf $decision = "i" Then
	$q = inputbox("ISBN", "Please enter the ISBN to search by.")
	$oIE_1 = _IECreateEmbedded ()
$oIE_2 = _IECreateEmbedded ()
GUICreate("SWORD Book Search", 1400, 1000,0,0, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
$GUIActiveX_1 = GUICtrlCreateObj($oIE_1, 10, 40, 680, 800)
$GUIActiveX_2 = GUICtrlCreateObj($oIE_2, 690, 40, 680, 800)


;GUISetState(@SW_SHOW)       ;Show GUI
WinSetState("SWORD Book Search", "", @SW_SHOW)

_IENavigate ($oIE_1, "http://catalog.libraries.wright.edu/search/?searchtype=i&searcharg=" & $q)
_IENavigate ($oIE_2, "http://uclid.uc.edu/search/i?SEARCH=" & $q)
While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
	EndSelect
WEnd

GUIDelete()
Exit
EndIf
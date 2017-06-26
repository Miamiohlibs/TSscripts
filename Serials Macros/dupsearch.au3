#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#Include <String.au3>
#include <Array.au3>
#Include <Misc.au3>

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



WinActivate("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
WinWaitActive ("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")

Sleep(0100)
ClipPut("")
Sleep(0100)
ClipPut("")
Sleep(0100)
_SendEx("^{HOME}")
Sleep(0100)
_SendEx("^a")
Sleep(0100)
_SendEx("^c")
Sleep(0100)
_SendEx("^c")



;set bib_rec
;set bib_rec variable

$BIB_REC = ClipGet()

;replace tabs and line breaks and split record into an array

$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "\n?\r?\t?(?!IMC|DVD|MLC|ROM)[A-Z1-3/.']{3,8}[[:blank:]A-Z#.]{0,7}\t", "Fnord")

$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "Fnord", 1)


$OCLC_NUM_A = _ArraySearch($BIB_ARRAY_MASTER, "001	", 0, 0, 0, 1)

If $OCLC_NUM_A > -1 Then

	$OCLC_NUM = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $OCLC_NUM_A, $OCLC_NUM_A)
	$OCLC_NUM =  StringTrimLeft($OCLC_NUM, 8)
	$OCLC_NUM =  StringTrimRight($OCLC_NUM, 1)
	$OCLC_NUM =  StringStripWS($OCLC_NUM, 8)

EndIf


;~ $OCLC_NUM = "45487059"




$oIE_1 = _IECreateEmbedded ()
$oIE_2 = _IECreateEmbedded ()
GUICreate("SWORD Book Search", 1200, 800,0,0, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
$GUIActiveX_1 = GUICtrlCreateObj($oIE_1, 10, 40, 580, 700)
$GUIActiveX_2 = GUICtrlCreateObj($oIE_2, 590, 40, 580, 700)


GUISetState()       ;Show GUI

_IENavigate ($oIE_1, "http://wsuol2.wright.edu/search/o?SEARCH=" & $OCLC_NUM & "&sortdropdown=-&searchscope=7&submit=Search")
_IENavigate ($oIE_2, "http://uclid.uc.edu/search/o?SEARCH=" & $OCLC_NUM & "&submit.x=48&submit.y=16&searchscope=36")

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

;start of search strings from wright state
;http://wsuol2.wright.edu/search/o?SEARCH=5139880&sortdropdown=-&searchscope=1&submit=Search
;~ #######breakdown#######
;~ 		http://wsuol2.wright.edu/search/o - base url for OCLC search
;~ 		?SEARCH=5139880 - oclc number
;~ 		sortdropdown=-&searchscope=1&submit=Search - system sorted and limited to jounral title


;~ start of search strings from uc
;~ http://uclid.uc.edu/search/o?SEARCH=5139880&submit.x=26&submit.y=17&searchscope=14
;~ #######breakdown#######
;~ 		http://uclid.uc.edu/search/o - base url for oclc search
;~ 		?SEARCH=5139880 - oclc Number
;~ 		submit.x=26&submit.y=17&searchscope=14 - system sorted and limited to periodicals/serials



;~ oclc number 4359563 brings up no results in wright
;~ oclc number 40047301 brings up no results in uc


;start of search string from ohiolink
;~ #######breakdown#######
;~ http://olc1.ohiolink.edu/search~S0?/o5139880/o5139880/1,1,1,B/detlframeset&FF=o5139880&1,1,
;~ 		http://olc1.ohiolink.edu/search~S0?/ - base url for oclc search
;~ 		o5139880 - oclc number (with o attached)
;~ 		1,1,1,B/detlframeset&FF=o5139880&1,1, - holdings info...? with oclc number

;~ idea.... 5 letter codes. Example source page code :
;~ 	<tr class="holdingsws3ug"><td><a name="ws3ug">
;~ 	uc - ci3ug
;~ 	wright - ws3ug
;~ 	miami - mu3ug
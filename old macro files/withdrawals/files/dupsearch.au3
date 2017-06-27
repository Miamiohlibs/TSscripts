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


;############## end function to fix sticking mod keys #########


If WinExists("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
WinActivate("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")

$winsize = WinGetPos("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
_SendEx("^s")
Sleep(0200)
_SendEx("!q")
sleep(0300)
EndIf

;scan in library barcode, searches catalog by barcode
$barcode = InputBox("Barcode", "Scan the item's barcode", "")
If $barcode = "" Then
	Exit
Else
_IIIsearch("b", $barcode)
EndIf
sleep(0900)

_SendEx("^+e")
Sleep(0500)

;WinActivate("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
;WinWaitActive ("[REGEXPTITLE:[b]; CLASS:com.iii.ShowRecord$2]")
WinWaitActive("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0300)
_SendEx("^+e")

$decision = inputbox("Search for what?", "Enter o for OCLC# search, i for ISBN search, or t for Title search")
If $decision = "o" Then


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
;$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "Fnord")
;$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "Fnord", 1)

;replace tabs and line breaks and split record into an array

$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "\n?\r?\t?(?!IMC|DVD|MLC|ROM)[A-Z1-3/.']{3,8}[[:blank:]A-Z#.]{0,7}\t", "Fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "Fnord", 1)


$OCLC_NUM_A = _arrayItemString($BIB_ARRAY_MASTER, "001")
;msgbox(0, "OCLC", $OCLC_NUM_A)

If $OCLC_NUM_A > -1 Then

	;$OCLC_NUM = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $OCLC_NUM_A, $OCLC_NUM_A)
	$OCLC_NUM_A =  StringTrimLeft($OCLC_NUM_A, 8)
	;OCLC_NUM =  StringTrimRight($OCLC_NUM_A, 1)
	$OCLC_NUM_A =  StringStripWS($OCLC_NUM_A, 8)
	$q = $OCLC_NUM_A
;MsgBox(0, "OCLC", $OCLC_NUM)
EndIf


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
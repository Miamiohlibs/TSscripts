#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#Include <String.au3>
#include <Array.au3>
#Include <Misc.au3>
Opt("GUIResizeMode", $GUI_DOCKAUTO)
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


If WinExists("[REGEXPTITLE:\A[b][0-9ax]{7,8}]") Then
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{7,8}]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{7,8}]")
Else
	MsgBox(64, "No bib record", "Please open the bib record in edit view, close this message box, and run the script again.")
	Exit
EndIf

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
;replace tabs and line breaks and split record into an array

$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)


$decide = InputBox("Type of Search", "What do you want to search by?" & @CR & "Enter the corresponding letter." & @CR & @CR & "O - OCLC Number" & @CR & "T - Title" & @CR & "I - ISBN/ISSN", "")
$decide = StringUpper($decide)

Switch $decide
	Case "O" ;OCLC #
		$OCLC_NUM_A = _ArraySearch($BIB_ARRAY_MASTER, "OCLC #	001	", 0, 0, 0, 1)
		If $OCLC_NUM_A > -1 Then
			$OCLC_NUM = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $OCLC_NUM_A, $OCLC_NUM_A)
			$OCLC_NUM =  StringTrimLeft($OCLC_NUM, 15)
			$OCLC_NUM =  StringStripWS($OCLC_NUM, 8)
		EndIf
		$search_1 = "http://wsuol2.wright.edu/search/o?SEARCH=" & $OCLC_NUM & "&sortdropdown=-&searchscope=7&submit=Search"
		$search_2 = "http://uclid.uc.edu/search/o?SEARCH=" & $OCLC_NUM & "&submit.x=48&submit.y=16&searchscope=36"
		$osearch = "http://olc1.ohiolink.edu/search/o?SEARCH="  & $OCLC_NUM
;~ 		$search_2 = "http://www.lib.muohio.edu/multifacet/books/" & $OCLC_NUM & "?field=oclc&sort=relevance&fn[]=item_campus&fv[]=SWORD"
	Case "T" ;Title search (245 |a)
		$245_A = _ArraySearch($BIB_ARRAY_MASTER, "TITLE	245", 0, 0, 0, 1)
		If $245_A > -1 Then
			$245 = _ArrayToString($BIB_ARRAY_MASTER, @TAB, $245_A, $245_A)
			$245_A2 = StringSplit($245, "|")
			$245_A = _ArraySearch($245_A2, "TITLE	245", 0, 0, 0, 1)
			$245 = _ArrayToString($245_A2, @TAB, $245_A, $245_A)
			$245 = StringTrimLeft($245, 14)
			$245DASH = StringRegExp($245, "[:/]", 0) ;take out punctuation at the end of subfield a
			If $245DASH = 1 Then
				$245 = StringTrimRight($245, 2)
			EndIf
		EndIf
		$search_1 = "http://wsuol2.wright.edu/search/t?SEARCH=" & $245 & "&sortdropdown=-&searchscope=7&submit=Search"
		$search_2 = "http://uclid.uc.edu/search/t?SEARCH=" & $245 & "&submit.x=48&submit.y=16&searchscope=36"
		$245 = StringReplace($245, "{space}", "{+}")
		$osearch = "http://olc1.ohiolink.edu/search/t?SEARCH=" & $245
;~ 		$search_2 = "http://www.lib.muohio.edu/multifacet/books/"& $245 &"?field=title_phrase&sort=title_asc&fn[]=item_campus&fv[]=SWORD"
	Case "I" ;ISBN
		$ISBN = InputBox("ISN Search", "Scan in the ISBN/ISSN barcode", "") ;Scan isbn from the physical item
		$search_1 = "http://wsuol2.wright.edu/search/i?SEARCH=" & $ISBN & "&sortdropdown=-&searchscope=7&submit=Search"
		$search_2 = "http://uclid.uc.edu/search/i?SEARCH=" & $ISBN & "&submit.x=48&submit.y=16&searchscope=36"
		$osearch = "http://olc1.ohiolink.edu/search/i?SEARCH=" & $ISBN
;~ 		$search_2 = "http://www.lib.muohio.edu/multifacet/books/"& $ISBN & "?field=isn&sort=relevance&fn[]=item_campus&fv[]=SWORD"
EndSwitch




$oIE_1 = _IECreateEmbedded ()
$oIE_2 = _IECreateEmbedded ()
GUICreate("SWORD Book Search", 800, 600,0,0, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
$GUIActiveX_1 = GUICtrlCreateObj($oIE_1, 10, 10, 390, 580)
$GUIActiveX_2 = GUICtrlCreateObj($oIE_2, 410, 10, 390, 580)


GUISetState(@SW_MAXIMIZE)       ;Show GUI

_IENavigate ($oIE_1, $search_1)
_IENavigate ($oIE_2, $search_2)

; Waiting for user to close the window
While 1
	$msg = GUIGetMsg()
	Select
		Case $msg = $GUI_EVENT_CLOSE
			ExitLoop
	EndSelect
WEnd

GUIDelete()

$decide = MsgBox(4, "Search Ohiolink?", "Do you want to search Ohiolink?")

Switch $decide
	Case 6

		$oIE_1 = _IECreateEmbedded ()
		GUICreate("SWORD Book Search", 800, 600,0,0, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
		$GUIActiveX_1 = GUICtrlCreateObj($oIE_1, 10, 10, 780, 580)

		GUISetState(@SW_MAXIMIZE)       ;Show GUI
		_IENavigate ($oIE_1, $osearch)

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

	Case 7
		Exit
EndSwitch

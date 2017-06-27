#Include <TSCustomFunction.au3>

;~ Dim $BIB_NUM


;~ $BIB_NUM = WinGetTitle("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")

;~ ;focus on Notepad: toggle with code above to testing purposes
;~ WinActivate("Untitled")
;~ WinWaitActive("Untitled")

;~ ;send 949 line
;~ _SendEx("^{END}")
;~ Sleep(0100)
;~ _SendEx("{ENTER}949{SPACE 2}*recs-b;b3--;ov-." & $BIB_NUM)



#include <MsgBoxConstants.au3>

Example()

Func Example()
    ; Retrieve the window text of the active window.
    Local $sText = WinGetText("[ACTIVE]")

    ; Display the window text.
    MsgBox($MB_SYSTEMMODAL, "", $sText)
EndFunc   ;==>Example



_ClearBuffer
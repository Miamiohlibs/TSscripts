#include <MsgBoxConstants.au3>
#include <TSCustomFunction.au3>

Example()

Func Example()
	; Retrieve the window text of the active window.
	Local $BIB_NUM = WinGetTitle("[REGEXPTITLE:[A-z]{8}; CLASS:SunAwtFrame]")
	Local $sText = WinGetText($BIB_NUM)

;~ ;focus on Notepad: toggle with code above to testing purposes
	WinActivate("Untitled")
	WinWaitActive("Untitled")

	_SendEx($sText)
EndFunc   ;==>Example

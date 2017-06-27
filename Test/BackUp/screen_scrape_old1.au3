#include <MsgBoxConstants.au3>
#include <TSCustomFunction.au3>

Example()

Func Example()
	; Retrieve the window text of the active window.
	Local $sText = WinGetText("[TITLE:Sierra; CLASS:SunAwtFrame]")

	;~ ;focus on Notepad: toggle with code above to testing purposes
	WinActivate("Untitled")
	WinWaitActive("Untitled")

	_SendEx($sText)
EndFunc   ;==>Example

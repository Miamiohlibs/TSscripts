;Playing around with stuff!
#include <GuiConstantsEx.au3>
#Include <GuiButton.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>


; GUI
$guiheight = 70
$guiwidth = 240
$mainwindow = GuiCreate("Order macro box", $guiwidth, $guiheight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
GUISetBkColor(0x4E9258)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")



$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\order\Images\take1_gui.ico")
GUICtrlSetOnEvent($1, "gobitoiii")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\order\Images\red 9_gui.ico")
GUICtrlSetOnEvent($2, "oclc")

$3 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\order\Images\blue O_gui.ico")
GUICtrlSetOnEvent($3, "OCLCtoiii")


$6 = GUICtrlCreateButton("Sample button", 170, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\order\Images\stopsign_gui.ico")
GUICtrlSetOnEvent($6, "kill")

	
	
GUISetState(@SW_SHOW)


While 1
	Sleep(1000)
WEnd


Func gobitoiii()
	Sleep(0400)
	Run(@DesktopDir & "\order\GOBItoIII-OCLC.exe")
EndFunc


Func oclc()
	Sleep(0400)
	Run(@DesktopDir & "\order\OCLC049and949.exe")

EndFunc

Func OCLCtoiii()
	Sleep(0400)
	Run(@DesktopDir & "\order\OCLCtoIII.exe")

EndFunc





Func kill()

		ProcessClose("GOBItoIII-OCLC.exe")
		ProcessClose("OCLC049and949.exe")
		ProcessClose("OCLCtoIII.exe")

EndFunc

Func CLOSEClicked()
  ;Note: at this point @GUI_CTRLID would equal $GUI_EVENT_CLOSE, 
  ;and @GUI_WINHANDLE would equal $mainwindow 
  Exit
EndFunc
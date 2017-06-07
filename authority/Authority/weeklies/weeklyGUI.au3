;Playing around with stuff!
#include <GuiConstantsEx.au3>
#Include <GuiButton.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>

Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)

HotKeySet("{F11}", "newsearch")

; GUI
$guiheight = 70
$guiwidth = 190
$mainwindow = GuiCreate("Weekly Report macro box", $guiwidth, $guiheight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
GUISetBkColor(0xFDD017)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\take1_gui.ico")
GUICtrlSetOnEvent($1, "prepheadings")

$4 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\Whistler XP Icon 11_gui.ico")
GUICtrlSetOnEvent($4, "newsearch")


$6 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\stopsign_gui.ico")
GUICtrlSetOnEvent($6, "kill")

GUISetState(@SW_SHOW)

While 1
	Sleep(1000)
WEnd

Func prepheadings()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\weeklies\file_prep_weekly.exe")
EndFunc

Func newsearch()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\weeklies\New_search_weekly.exe")

EndFunc


Func kill()
		ProcessClose("file_prep_weekly.exe")
		ProcessClose("New_search_weekly.exe")
EndFunc

Func CLOSEClicked()
  ;Note: at this point @GUI_CTRLID would equal $GUI_EVENT_CLOSE,
  ;and @GUI_WINHANDLE would equal $mainwindow
  Exit
EndFunc
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
$guiwidth = 440
$mainwindow = GuiCreate("Authority Headings Report macro box", $guiwidth, $guiheight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
GUISetBkColor(0xE55B3C)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\take1_gui.ico")
GUICtrlSetOnEvent($1, "justauto1")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\blue 2_gui.ico")
GUICtrlSetOnEvent($2, "justauto15")

$3 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\black 3_gui.ico")
GUICtrlSetOnEvent($3, "justauto2")

$4 = GUICtrlCreateButton("Sample button", 170, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\green 4_gui.ico")
GUICtrlSetOnEvent($4, "justauto25")

$5 = GUICtrlCreateButton("Sample button", 220, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\Whistler XP Icon 10_gui.ico")
GUICtrlSetOnEvent($5, "prepheadings")

$6 = GUICtrlCreateButton("Sample button", 270, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\Whistler XP Icon 11_gui.ico")
GUICtrlSetOnEvent($6, "newsearch")

$7 = GUICtrlCreateButton("Sample button", 320, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\Magnifier_gui.ico")
GUICtrlSetOnEvent($7, "moresearch")

$8 = GUICtrlCreateButton("Sample button", 370, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\Authority\Images\stopsign_gui.ico")
GUICtrlSetOnEvent($8, "kill")

GUISetState(@SW_SHOW)

While 1
	Sleep(1000)
WEnd

Func justauto1()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\justautomated_1.exe")
EndFunc

Func justauto15()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\justautomated_1.5.exe")
EndFunc

Func justauto2()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\justautomated_2.exe")
EndFunc

Func justauto25()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\justautomated_2.5.exe")
EndFunc

Func prepheadings()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\headings_prep.exe")

EndFunc

Func newsearch()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\New_search.exe")

EndFunc

Func moresearch()
	Sleep(0400)
	Run(@DesktopDir & "\Authority\More_search.exe")

EndFunc

Func kill()
		ProcessClose("justautomated_1.exe")
		ProcessClose("justautomated_1.5.exe")
		ProcessClose("justautomated_2.exe")
		ProcessClose("justautomated_2.5.exe")
		ProcessClose("headings_prep.exe")
		ProcessClose("New_search.exe")
		ProcessClose("More_search.exe")
EndFunc

Func CLOSEClicked()
  ;Note: at this point @GUI_CTRLID would equal $GUI_EVENT_CLOSE,
  ;and @GUI_WINHANDLE would equal $mainwindow
  Exit
EndFunc
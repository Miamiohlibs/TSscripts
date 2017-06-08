#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: CopyCatGUI
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script creates a toolbar with buttons that launch various scripts in the
	copy cataloging set.

 Programs used: n/a

 Last revised: 6/29/10
			   Added comments

 Copyright (C): 2009 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------

;######### INCLUDES AND OPTIONS #########
#include <GuiConstantsEx.au3>
#Include <GuiButton.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>

;######### GUI #########
$guiheight = 70
$guiwidth = 390
$mainwindow = GuiCreate("Copy Cat macro box", $guiwidth, $guiheight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
GUISetBkColor(0x8B2252)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

;######### BUTTONS #########
$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\take1_gui.ico")
GUICtrlSetOnEvent($1, "barcodetotitle")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\Connex.ico")
GUICtrlSetOnEvent($2, "Connex")

$3 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\red 9_gui.ico")
GUICtrlSetOnEvent($3, "OCLCoverlay")

$5 = GUICtrlCreateButton("Sample button", 170, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\blue 5_gui.ico")
GUICtrlSetOnEvent($5, "bibandindex")

$7  = GUICtrlCreateButton("Sample button", 220, 10, 50,50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\black S_gui.ico")
GUICtrlSetOnEvent($7, "specialfunds")

$4 = GUICtrlCreateButton("Sample button", 270, 10, 50,50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\green D_gui.ico")
GUICtrlSetOnEvent($4, "Done")

$6 = GUICtrlCreateButton("Sample button", 320, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\CopyCat\Images\stopsign_gui.ico")
GUICtrlSetOnEvent($6, "kill")

;######### GUI STATE #########
GUISetState(@SW_SHOW)
While 1
	Sleep(1000)
WEnd

;######### GUI FUNCTIONS #########
Func barcodetotitle()
	Sleep(0400)
	Run(@DesktopDir & "\CopyCat\barcodetorecord.exe")
EndFunc

Func Connex()
	Sleep(0400)
	Run(@DesktopDir & "\CopyCat\OCLCRec.exe")
EndFunc

Func OCLCoverlay()
	Sleep(0400)
	Run(@DesktopDir & "\CopyCat\949.exe")
EndFunc

Func bibandindex()
	Sleep(0400)
	Run(@DesktopDir & "\CopyCat\504.exe")
EndFunc

Func specialfunds()
	Sleep(0400)
	Run(@DesktopDir & "\CopyCat\SpecialFundNote.exe")
EndFunc

Func Done()
	Sleep(0400)
	Run(@DesktopDir & "\CopyCat\ItemRec.exe")
EndFunc

Func kill()
		ProcessClose("barcodetorecord.exe")
		ProcessClose("OCLCRec.exe")
		ProcessClose("504.exe")
		ProcessClose("949.exe")
		ProcessClose("SpecialFundNote.exe")
		ProcessClose("ItemRec.exe")
EndFunc

Func CLOSEClicked()
  Exit
EndFunc
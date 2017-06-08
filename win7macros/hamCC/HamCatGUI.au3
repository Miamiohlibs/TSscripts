#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: HamCatGUI
 Script set: Hamilton Copy Cataloging (hamCC)

 Script Function:
	This script creates a toolbar with buttons that launch various scripts in the
	Hamilton copy cataloging set.

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
$guiwidth = 190
$mainwindow = GuiCreate("Ham Cat macro box", $guiwidth, $guiheight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
GUISetBkColor(0x00FF00)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

;######### BUTTONS #########
$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\hamCC\Images\iiirunner.ico")
GUICtrlSetOnEvent($1, "inIII")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\hamCC\Images\Connex.ico")
GUICtrlSetOnEvent($2, "notinIII")

$6 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\hamCC\Images\stopsign_gui.ico")
GUICtrlSetOnEvent($6, "kill")

;######### GUI STATE #########
GUISetState(@SW_SHOW)
While 1
	Sleep(1000)
WEnd

;######### GUI FUNCTIONS #########
Func inIII()
	Sleep(0400)
	Run(@DesktopDir & "\hamCC\copyinIII.exe")
EndFunc

Func notinIII()
	Sleep(0400)
	Run(@DesktopDir & "\hamCC\copynotinIII.exe")
EndFunc

Func kill()
		ProcessClose("copyinIII.exe")
		ProcessClose("copynotinIII.exe")
EndFunc

Func CLOSEClicked()
  Exit
EndFunc
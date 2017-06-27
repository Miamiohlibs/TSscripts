#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\..\Transition\macros\SierraMacros\scripts\images\autoiticon.ico
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.4.4
 Author: Martin Patrick, Academic Resident Librarian in Technical Services
 Miami University
 patricm@miamioh.edu

 Name of script: withdrawGUI
 Script set: Withdrawals

 Script Function:
	This script creates a toolbar with buttons that launch various scripts in the
	withdrawal processing set.

 Programs used: n/a

 Last revised:5/2015


 Copyright (C): 2015 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------

;######### INCLUDES AND OPTIONS #########
#include <GuiConstantsEx.au3>
#Include <GuiButton.au3>
#include <WindowsConstants.au3>

;######### GUI #########
$mainwindow = GuiCreate("Withdrawals Box", 290, 70, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
GUISetBkColor(0xF5002F)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

;######### BUTTONS #########
$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\Images\take1new.ico")
GUICtrlSetOnEvent($1, "dupsearch")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\Images\search.ico")
GUICtrlSetOnEvent($2, "moresearch")

$3 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\Images\bluedubs.ico")
GUICtrlSetOnEvent($3, "withdraw")

$4 = GUICtrlCreateButton("Sample button", 170, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\Images\sword3.ico")
GUICtrlSetOnEvent($4, "sword")

$5 = GUICtrlCreateButton("Sample button", 220, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\Images\stopsign2.ico")
GUICtrlSetOnEvent($5, "kill")

;######### GUI STATE #########
GUISetState(@SW_SHOW)
While 1
	Sleep(1000)
WEnd

;######### GUI FUNCTIONS #########
Func dupsearch()
	Sleep(0400)
	Run("dupsearch.exe")
EndFunc

Func moresearch()
	Sleep(0400)
	Run("moresearch.exe")
	EndFunc

Func withdraw()
	Sleep(0400)
	Run("withdraw.exe")
EndFunc

Func sword()
	Sleep(0400)
	Run("sword.exe")
EndFunc

Func kill()
		ProcessClose("dupsearch.exe")
		ProcessClose("withdraw.exe")
		ProcessClose("sword.exe")
EndFunc

Func CLOSEClicked()
  Exit
EndFunc
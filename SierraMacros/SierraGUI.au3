#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\patricm\Google Drive\Transition\macros\SierraMacros\scripts\images\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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
$guiwidth = 440
$mainwindow = GuiCreate("Sierra macro box", $guiwidth, $guiheight, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
GUISetBkColor(0xF5002F)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

;######### BUTTONS #########
$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\take1new.ico")
GUICtrlSetOnEvent($1, "FindRecord")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\blueO.ico")
GUICtrlSetOnEvent($2, "OrderRecord")

$3 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\blackB.ico")
GUICtrlSetOnEvent($3, "BibRecord")

$4 = GUICtrlCreateButton("Sample button", 170, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\oclc.ico")
GUICtrlSetOnEvent($4, "Connex")

$5 = GUICtrlCreateButton("Sample button", 220, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\red9.ico")
GUICtrlSetOnEvent($5, "OCLCoverlay")

$6 = GUICtrlCreateButton("Sample button", 270, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\blue5.ico")
GUICtrlSetOnEvent($6, "bibandindex")

$7 = GUICtrlCreateButton("Sample button", 320, 10, 50,50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\greenD.ico")
GUICtrlSetOnEvent($7, "Done")

$8 = GUICtrlCreateButton("Sample button", 370, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @WorkingDir & "\scripts\Images\stopsign2.ico")
GUICtrlSetOnEvent($8, "kill")

;######### GUI STATE #########
GUISetState(@SW_SHOW)
While 1
	Sleep(1000)
WEnd

;######### GUI FUNCTIONS #########
Func FindRecord()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\FindRecord.exe")
EndFunc

Func OrderRecord()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\OrderRecord.exe")
EndFunc

Func BibRecord()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\BibRecord.exe")
EndFunc

Func Connex()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\OCLCRec.exe")
EndFunc

Func OCLCoverlay()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\949.exe")
EndFunc

Func bibandindex()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\504.exe")
EndFunc

#cs
Func specialfunds()
	Sleep(0400)
	Run(@WorkingDir & "\SpecialFundNote.exe")
EndFunc
#ce

Func Done()
	Sleep(0400)
	Run(@WorkingDir & "\scripts\ItemCreate.exe")
EndFunc

Func kill()
		ProcessClose("FindRecord.exe")
		ProcessClose("OCLCRec.exe")
		ProcessClose("504.exe")
		ProcessClose("949.exe")
		ProcessClose("SpecialFundNote.exe")
		ProcessClose("ItemCreate.exe")
		ProcessClose("OrderRecord.exe")
EndFunc

Func CLOSEClicked()
  Exit
EndFunc
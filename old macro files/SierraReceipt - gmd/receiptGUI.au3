#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: receiptGUI
 Script set: Receipt Cataloging (Receipt)

 Script Function:
	This script creates a toolbar with buttons that launch various scripts in the
	receipt cataloging set.

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

;######### GUI #########
$mainwindow = GuiCreate("Receipt Cataloging Box", 390, 70, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP, $WS_MAXIMIZEBOX, $WS_SIZEBOX), BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GuiSetIcon(@SystemDir & "\mspaint.exe", 0)
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
GUISetBkColor(0x6591CD)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")

;######### BUTTONS #########
$1 = GUICtrlCreateButton("Sample button", 20, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\take1_gui.ico")
GUICtrlSetOnEvent($1, "isbntotitle")

$2 = GUICtrlCreateButton("Sample button", 70, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\blue O_gui.ico")
GUICtrlSetOnEvent($2, "order")

$3 = GUICtrlCreateButton("Sample button", 120, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\black B_gui.ico")
GUICtrlSetOnEvent($3, "bib")

$4 = GUICtrlCreateButton("Sample button", 170, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\green D_gui.ico")
GUICtrlSetOnEvent($4, "DLC")

$5 = GUICtrlCreateButton("Sample button", 220, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\red N_gui.ico")
GUICtrlSetOnEvent($5, "nonDLC")

$6 = GUICtrlCreateButton("Sample button", 270, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\rda.ico")
GUICtrlSetOnEvent($6, "rda")

$7 = GUICtrlCreateButton("Sample button", 320, 10, 50, 50, $BS_ICON)
GUICtrlSetImage(-1, @DesktopDir & "\SierraReceipt - gmd\Images\stopsign_gui.ico")
GUICtrlSetOnEvent($7, "kill")

;######### GUI STATE #########
GUISetState(@SW_SHOW)
While 1
	Sleep(1000)
WEnd

;######### GUI FUNCTIONS #########
Func isbntotitle()
	Sleep(0400)
	Run(@DesktopDir & "\SierraReceipt - gmd\isbnToTitle.exe")
EndFunc

Func order()
	Sleep(0400)
	Run(@DesktopDir & "\SierraReceipt - gmd\OrderRecord.exe")
EndFunc

Func bib()
	Sleep(0400)
	Run(@DesktopDir & "\SierraReceipt - gmd\BibRecord.exe")
EndFunc

Func DLC()
	Sleep(0400)
	Run(@DesktopDir & "\SierraReceipt - gmd\Create Item - DLC.exe")
EndFunc

Func nonDLC()
	Sleep(0400)
	Run(@DesktopDir & "\SierraReceipt - gmd\Create item - Non DLC.exe")
EndFunc

Func rda()
	Sleep(0400)
	Run(@DesktopDir & "\SierraReceipt - gmd\rda.exe")
EndFunc


Func kill()
		ProcessClose("Create Item - Non DLC.exe")
		ProcessClose("Create Item - DLC.exe")
		ProcessClose("BibRecord.exe")
		ProcessClose("OrderRecord.exe")
		ProcessClose("isbnToTitle.exe")
EndFunc

Func CLOSEClicked()
  Exit
EndFunc
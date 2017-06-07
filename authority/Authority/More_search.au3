#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: More_search
 Script set: Headings Files (Authority)

 Script Function: Replaces the following ME scripts
					subheading search macros


 Programs used:

 Last revised:

 PLEASE NOTE - This script uses a custom UDF library called TSCustomFunction.
			   The script will not run properly (if launched from .au3 file)
			   if that file is not included in the include folder in the
			   AutoIt directory.

 Copyright (C): 2010 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------
;######### INCLUDES AND OPTIONS #########
#Include <TSCustomFunction.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>


Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKAUTO)
AutoItSetOption("WinTitleMatchMode", 2)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $heading_term
Dim $OCLC_search_prefix
Dim $III_search_prefix
Dim $subdivision_a
Dim $subdivision_1
Dim $subdivision_2
Dim $subdivision_3
Dim $subdivision_4
Dim $x, $y
Dim $mainwindow, $iii, $oclc, $done
dim $radio1, $radio2, $radio3, $radio4, $radio5, $radio6, $search, $tab


;################################ MAIN ROUTINE #################################

$heading_term = _LoadVar("$heading_term")
$OCLC_search_prefix = _LoadVar("$OCLC_search_prefix")
$III_search_prefix = _LoadVar("$III_search_prefix")
$subdivision_a = _LoadVar("$subdivision_a")
$subdivision_1 = _LoadVar("$subdivision_1")
$subdivision_2 = _LoadVar("$subdivision_2")
$subdivision_3 = _LoadVar("$subdivision_3")
$subdivision_4 = _LoadVar("$subdivision_4")

;######### set hot keys #########
HotKeySet("!{NUMPAD0}", "main_first_mil")
HotKeySet("!{NUMPAD1}", "main_mil")
HotKeySet("!{NUMPAD2}", "first_mil")
HotKeySet("!{NUMPAD3}", "second_mil")
HotKeySet("!{NUMPAD4}", "third_mil")
HotKeySet("!{NUMPAD5}", "fourth_mil")

HotKeySet("^{NUMPAD0}", "main_first_oclc")
HotKeySet("^{NUMPAD1}", "main_oclc")
HotKeySet("^{NUMPAD2}", "first_oclc")
HotKeySet("^{NUMPAD3}", "second_oclc")
HotKeySet("^{NUMPAD4}", "third_oclc")
HotKeySet("^{NUMPAD5}", "fourth_oclc")

HotKeySet("{F8}", "skip_search")
HotKeySet("{F11}", "start_new_search")

;######### GUI FOR MORE SEARCHING #########

$mainwindow = GUICreate("", 90, 70, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_TABSTOP)) ; will create a dialog box that when displayed is centered
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked", $mainwindow)

$done = GUICtrlCreateButton("Done searching!", 10, 10, 70, 50, $BS_MULTILINE)
GUICtrlSetOnEvent($done, "CLOSEClicked")

GUISetState(@SW_SHOW)
While 1
	Sleep(1000)
WEnd

;######### LOCAL FUNCTIONS #########

Func _OCLCauthsearch($x, $y)
	If WinExists("OCLC Connexion") Then
		WinActivate("OCLC Connexion")
	Else
		MsgBox(64, "nope", "Please open Connexion. Click ok when you are logged in")
	EndIf
	Send("!{F2}")
	WinWaitActive("Browse Authority File")
	ControlSend("Browse Authority File", "", "[CLASS:Edit; INSTANCE:2]", $y)
	Sleep (0200)
	Send("{TAB}")
	Sleep (0200)
	Send ("{HOME}")
	Sleep (0200)
	Send ($x)
	Sleep (0200)
	Send("{ENTER}")
EndFunc
;~ ******* III search *****

Func main_first_mil()
	If $subdivision_1 <> "" Then
		_IIIsearch($III_search_prefix, $subdivision_a &" "&$subdivision_1)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func main_mil()
	_IIIsearch($III_search_prefix, $subdivision_a)
EndFunc

Func first_mil()
	If $subdivision_1 <> "" Then
		_IIIsearch($III_search_prefix, $subdivision_1)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func second_mil()
	If $subdivision_2 <> "" Then
		_IIIsearch($III_search_prefix, $subdivision_2)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func third_mil()
	If $subdivision_3 <> "" Then
		_IIIsearch($III_search_prefix, $subdivision_3)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func fourth_mil()
	If $subdivision_4 <> "" Then
		_IIIsearch($III_search_prefix, $subdivision_4)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

;~ ******* OCLC search *****

Func main_first_oclc()
	If $subdivision_1 <> "" Then
		If WinExists("OCLC Connexion") Then
			WinActivate("OCLC Connexion")
		Else
			MsgBox(64, "nope", "Please open Connexion. Click ok when you are logged in")
		EndIf
		Send("!{F2}")
		WinWaitActive("Browse Authority File")
		ControlSend("Browse Authority File", "", "[CLASS:Edit; INSTANCE:2]", $subdivision_a)
		Sleep (0200)
		Send("{TAB}")
		Sleep (0200)
		Send ("{HOME}")
		Sleep (0200)
		Send ($OCLC_search_prefix)
		Sleep (0200)
		Send("{TAB}")
		Sleep (0200)
		Send($subdivision_1)
		Sleep (0200)
		Send("{ENTER}")
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func main_oclc()
	_OCLCauthsearch($OCLC_search_prefix, $subdivision_a)
EndFunc

Func first_oclc()
	If $subdivision_1 <> "" Then
		_OCLCauthsearch($OCLC_search_prefix, $subdivision_1)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func second_oclc()
	If $subdivision_2 <> "" Then
		_OCLCauthsearch($OCLC_search_prefix, $subdivision_2)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func third_oclc()
	If $subdivision_3 <> "" Then
		_OCLCauthsearch($OCLC_search_prefix, $subdivision_3)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

Func fourth_oclc()
	If $subdivision_4 <> "" Then
		_OCLCauthsearch($OCLC_search_prefix, $subdivision_4)
	Else
		MsgBox(0, "Nothing to search", "The subdivision does not exist.")
	EndIf
EndFunc

;~ ******* Skip search of next heading *****

Func skip_search()
	WinActivate("Microsoft Word")
	WinWaitActive("Microsoft Word")
	Send("^h")
	WinWaitActive("Find and Replace")
	Send("!d")
	Send("Field:")
	Send("{ENTER}")
	Send("!{f4}")
	Send("{END}+{HOME}^c")
	Send("!hfc")
	Send("{DOWN}{RIGHT}{ENTER}")
	Send("{END}{HOME}{DELETE}{END}")
EndFunc

Func CLOSEClicked()
	Exit
EndFunc

Func start_new_search()
	_ClearBuffer()
	Run(@DesktopDir & "\Authority\New_search.exe")
	Exit
EndFunc
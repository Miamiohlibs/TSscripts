#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: New_search_weekly
 Script set: Weekly Files (Authority)

 Script Function: Replaces the following ME scripts
					F11
					F9
					F12

 Programs used: Millennium Cataloging
				OCLC Connexion

 Last revised: 4/1/2011

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
#include <IE.au3>


AutoItSetOption("WinTitleMatchMode", 2)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $heading_term
Dim $left_bracket
Dim $heading_number
Dim $decide

;################################ MAIN ROUTINE #################################

_ClearBuffer()

WinActivate("Microsoft Word")
WinWaitActive("Microsoft Word")

_SendEx("^h")
WinWaitActive("Find and Replace")
_SendEx("!d")
Sleep(0100)
_SendEx("{^}p1")
Sleep(0100)
_SendEx("{ENTER}")
Sleep(0100)
_SendEx("!{f4}")
Sleep(0100)
_SendEx("{RIGHT}")
Sleep(0100)
_SendEx("{END}+{HOME}^c")
Sleep(0100)
_SendEx("!hfc")
Sleep(0100)
_SendEx("{DOWN}{ENTER}")
Sleep(0100)
_SendEx("{END}{HOME}Z")
Sleep(0100)
$heading_term = ClipGet()

;MsgBox(0, "", $heading_term)

$heading_term = StringReplace($heading_term, "CANCEL", " ")
$heading_term = StringReplace($heading_term, "ADD GEOG", " ")
$heading_term = StringReplace($heading_term, "CHANGE GEOG", " ")

$left_bracket = StringInStr($heading_term, "[")
If $left_bracket = 0 Then
	Sleep (0100)
Else
	$heading_number = StringTrimLeft ($heading_term, $left_bracket)
	$heading_number = StringReplace($heading_number, "sh", "")
	$heading_number = StringReplace($heading_number, "]", "")
	$heading_number = StringStripWS($heading_number, 8)
	$heading_term = StringLeft($heading_term, $left_bracket)
EndIf
$heading_term = StringTrimRight($heading_term, 1)
$heading_term = StringRegExpReplace($heading_term, "\A1[0-9]{2}", "")
$heading_term = StringStripWS($heading_term, 7)

;~ MsgBox(0, '', $heading_number)

_IIIsearch("d", $heading_term)

$decide = MsgBox(68, "OCLC Search", "Do you want to search OCLC?")

Switch $decide
	Case 6
		_OCLCauthsearch("ln", $heading_number)
	Case 7
		Exit
EndSwitch

;######### LOCAL FUNCTIONS #########

Func _OCLCauthsearch($x, $y)
	If WinExists("OCLC Connexion") Then
		WinActivate("OCLC Connexion")
	Else
		MsgBox(64, "nope", "Please open Connexion. Click ok when you are logged in")
	EndIf
	_SendEx("+{F2}")
	WinWaitActive("Search LC Names and Subjects")
	ControlSend("Search LC Names and Subjects", "", "[CLASS:Edit; INSTANCE:1]", $x & ":" & $y)
	_SendEx("{ENTER}")
EndFunc
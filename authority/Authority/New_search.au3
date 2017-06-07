#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: New_search
 Script set: Headings Files (Authority)

 Script Function: Replaces the following ME scripts
					F11
					Search MILCAT
					Search OCLC

 Programs used: Microsoft Word 2010
				Millennium Cataloging Module
				OCLC Connexion version 2.20

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
Dim $xzo_check
Dim $970_check
Dim $first_tag_number
Dim $second_third_tag_numbers
Dim $OCLC_search_prefix
Dim $III_search_prefix
Dim $subdivision_a
Dim $subdivision_1
Dim $subdivision_2
Dim $subdivision_3
Dim $subdivision_4
Dim $x, $y
Dim $mainwindow, $iii, $oclc
dim $radio1, $radio2, $radio3, $radio4, $radio5, $radio6, $search, $tab

;################################ MAIN ROUTINE #################################
_ClearBuffer()

WinActivate("Microsoft Word")
WinWaitActive("Microsoft Word")

;find a valid heading to search, change color and delect the first character in each heading entry
Do
	_SendEx("^h")
	Sleep(0200)
	WinWaitActive("Find and Replace")
	Sleep(0200)
	_SendEx("!d")
	Sleep(0200)
	_SendEx("Field:")
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	_SendEx("!{f4}")
	Sleep(0200)
	_SendEx("{END}+{HOME}^c")
	Sleep(0200)
	_SendEx("!hfc")
	Sleep(0200)
	_SendEx("{DOWN}{ENTER}")
	Sleep(0200)
	_SendEx("{END}{HOME}{DELETE}{END}")
	$heading_term = ClipGet()
	$xzo_check = StringInStr($heading_term, "xzo")
	$970_check = StringInStr($heading_term, "v970")
Until $xzo_check = 0 And $970_check = 0

;parse out 3 digit tag to determine search strategy
$heading_term = StringTrimLeft($heading_term, 8)
$first_tag_number = StringLeft($heading_term, 1)
$second_third_tag_numbers = StringLeft($heading_term, 3)
$second_third_tag_numbers = StringTrimLeft($second_third_tag_numbers, 1)

;~ MsgBox(0, "", $first_tag_number & $second_third_tag_numbers)

_StoreVar("$heading_term")
_StoreVar("$first_tag_number")
_StoreVar("$second_third_tag_numbers")

Run(@DesktopDir & "\Authority\parse_headings.exe")
	While ProcessExists("parse_headings.exe")
		Sleep(0100)
	WEnd

$heading_term = _LoadVar("$heading_term")
$OCLC_search_prefix = _LoadVar("$OCLC_search_prefix")
$III_search_prefix = _LoadVar("$III_search_prefix")
$subdivision_a = _LoadVar("$subdivision_a")
$subdivision_1 = _LoadVar("$subdivision_1")
$subdivision_2 = _LoadVar("$subdivision_2")
$subdivision_3 = _LoadVar("$subdivision_3")
$subdivision_4 = _LoadVar("$subdivision_4")

If $first_tag_number = 6 Then
	_IIIsearch($III_search_prefix, $subdivision_a)
Else
	_OCLCauthsearch($OCLC_search_prefix, $subdivision_a)
EndIf


Run(@DesktopDir & "\Authority\More_search.exe")


;######### LOCAL FUNCTIONS #########
Func _OCLCauthsearch($x, $y)
	If WinExists("OCLC Connexion") Then
		WinActivate("OCLC Connexion")
	Else
		MsgBox(64, "nope", "Please open Connexion. Click ok when you are logged in")
	EndIf
	_SendEx("!{F2}")
	WinWaitActive("Browse Authority File")
	ControlSend("Browse Authority File", "", "[CLASS:Edit; INSTANCE:2]", $y)
	Sleep (0200)
	_SendEx("{TAB}")
	Sleep (0200)
	Send ("{HOME}")
	Sleep (0200)
	Send ($x)
	Sleep (0200)
	_SendEx("{ENTER}")
EndFunc

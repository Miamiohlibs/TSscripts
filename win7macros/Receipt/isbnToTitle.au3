#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: isbnToTitle
 Script set: Receipt Cataloging (Receipt)

 Script Function:
	This script prompts the user to scan in the item's ISBN. The script will
	search for the item, brings up the bib record, and then searches that
	item again with the title from the bib record.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(Main catalog screen open)
				#### Record view proporties - make sure Order is selected

 Last revised: 6/29/10
			   Updated #Include to include TS custom function library
			   Updated window controls for JRE version

 PLEASE NOTE - This script uses a custom UDF library called TSCustomFunction.
			   The script will not run properly (if launched from .au3 file)
			   if that file is not included in the include folder in the
			   AutoIt directory.

 Copyright (C): 2009 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------

;######### INCLUDES AND OPTIONS #########
#Include <TSCustomFunction.au3>
AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
TraySetIcon(@DesktopDir & "\Receipt\Images\take1.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $245, $245_A2
Dim $ISBN
Dim $300, $300_E, $300_E1

;################################ MAIN ROUTINE #################################
;check to see if an existing record is open, close it and clear buffer
If WinExists("[REGEXPTITLE:\A[bio][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinKill ("[REGEXPTITLE:\A[bio][0-9ax]{7,8}; CLASS:SunAwtFrame]")
EndIf
_ClearBuffer()

;focus on main search screen
If WinExists("[TITLE:Millennium Cataloging; CLASS:SunAwtFrame]") Then
	WinActivate("[TITLE:Millennium Cataloging; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "nope", "Please log into Millennium.")
	Exit
EndIf

;ask for ISBN, search III
$ISBN = InputBox("ISBN", "Scan the item's ISBN", "")
_IIIsearch("i", $ISBN)

;go to bib record
;~ WinWaitActive("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
;~ Sleep(0200)
WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0400)
_SendEx("!g")
Sleep(0400)
_SendEx("!v")

WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;select all and copy bib record information, parse data into array
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;245 title search
$245 =  _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
$245_A2 = StringSplit($245, "|")
$245 =  _arrayItemString($245_A2, "TITLE	245")
$245 = StringTrimLeft($245, 14)

;300 accompanying material search for internal note
$300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300")
$300_E = StringInStr($300, "|e")
If $300_E > 0 Then
	$300_E = StringMid($300, $300_E)
	$300_E = StringTrimLeft($300_E, 2)
	$300_E1 = StringInStr($300_E, "(")
	If $300_E1 > 0 Then
		$300_E = StringLeft($300_E, $300_E1)
		$300_E = StringTrimRight($300_E, 1)
	EndIf
	$300_E = StringStripCR($300_E)
Else
	$300_E = "none"
EndIf
_StoreVar("$300_E") ;store for OrderRec use to create internal note

;close record and search III catalog with title
Sleep(0100)
WinKill("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0100)
_IIIsearch("t", $245)
Sleep(0800)

;dup check - if the order record doesn't pop up in a certain amount of time, then script prompts for a dup check
If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	Exit
Else
	MsgBox(0,"Check possible dup", "There are similar titles in the system. Please select correct title.")
EndIf
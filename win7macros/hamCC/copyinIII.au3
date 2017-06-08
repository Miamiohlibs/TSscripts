#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: copyinIII
 Script set: Hamilton Copy Cataloging (hamCC)

 Script Function:
	This script deals with Hamilton items that are in III. It enters appropriate
	codes in both the bib and item record and then searches OCLC for the item.
	The user then pushes the F8 key to set holdings in OCLC.

	#### This script has events that depend on user clicks or button entries.
	Please see script below for more details.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(Bib record open)
					Record view properties - Summary retrieval view, item
					summary view
				OCLC Connexion Client v 2.20
					(logged in)

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
TraySetIcon(@DesktopDir & "\hamCC\Images\iiirunner.ico")

;######### DECLARE VARIABLES #########
Dim $dll
dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $OCLC_NUM
dim $049_A, $049_D
Dim $DUP_CHECK
Dim $decide

$dll = DllOpen("user32.dll")

;################################ MAIN ROUTINE #################################
;focus on Millennium bib record
If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Millennium record", "Please open the bib record.")
	Exit
EndIf

;select all and copy bib record information, parce data into array
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;search for OCLC number
$OCLC_NUM= _arrayItemString($BIB_ARRAY_MASTER, "OCLC #	001	 	 	")
$OCLC_NUM =  StringTrimLeft($OCLC_NUM, 15)

;### USER INPUT - wait for left mouse click
While 1
	Sleep(0010)
	If _IsPressed("01", $dll) Then ;enter 049 and bib location information
		Sleep(0100)
		_SendEx("|aOHMM")
		Sleep(0100)
		_SendEx("^{HOME}")
		Sleep(0100)
		_SendEx("{DOWN 2}")
		Sleep(0100)
		_SendEx("^e")
		Sleep(0100)
		WinWaitActive("[TITLE:Edit Data; CLASS:SunAwtDialog]")
		Sleep(0100)
		_SendEx("!a")
		Sleep(0100)
		_SendEx("hal")
		Sleep(0100)
		_SendEx("!o")
		Sleep(0100)
	ExitLoop
	EndIf
WEnd

;save
_SendEx("^s")
Sleep(0100)
$DUP_CHECK = WinExists("[TITLE:Perform duplicate checking?; CLASS:SunAwtDialog]")
If $DUP_CHECK = 1 Then
	_SendEx("n")
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
EndIf
Sleep(1000)

;create new item record
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("i")
Sleep(0300)
_SendEx("!n")
WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]", "")
_SendEx("^{HOME}")
Sleep(0100)
_SendEx("{TAB}")
Sleep(0100)
_SendEx("hali")
Sleep(0200)
_SendEx("{TAB 2}")
Sleep(0200)
_SendEx("{DEL}")
Sleep(0100)
_SendEx("823")
Sleep(0100)
_SendEx("{TAB 3}")
Sleep(0100)
_SendEx("l")
Sleep(0100)
_SendEx("{ENTER}")
Sleep(0100)
_SendEx("b{TAB}")
Sleep(0400)
$decide = InputBox("Barcode", "Scan in Barcode", "")
Sleep(0100)
_SendEx($decide)
Sleep(0100)
_SendEx("^s")
Sleep(0100)

;search item in OCLC
_OCLCSearch("{#}", $OCLC_NUM)

;### USER INPUT - wait for F8 key press
While 1
	Sleep(0010)
	If _IsPressed("77", $dll) Then ;Update holdings
		Sleep(0100)
		_SendEx("^{F4}")
		Sleep(0100)
	ExitLoop
	EndIf
WEnd

DllClose($dll)
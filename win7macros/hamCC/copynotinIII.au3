#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: copynotinIII
 Script set: Hamilton Copy Cataloging (hamCC)

 Script Function:
	This script deals with Hamilton items that are not in III. The script
	performs an ISBN search in OCLC. After the user is done with the record,
	the script inserts the appropriate 949 field and exports the record into
	III.

	#### This script has events that depend on user clicks or button entries.
	Please see script below for more details.

 Programs used:	OCLC Connexion Client v 2.20
					(logged in)

 Last revised: 6/29/10
			   Updated #Include to include TS custom function library

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
TraySetIcon(@DesktopDir & "\hamCC\Images\Connex.ico")

;######### DECLARE VARIABLES #########
Dim $ISBN
dim $done
Dim $decide

;######### SET HOT KEY #########
HotKeySet("^h", "export") ; see export function below

;################################ MAIN ROUTINE #################################
;search oclc by isbn
$ISBN = InputBox("ISBN Search", "Scan in the ISBN barcode", "")
_OCLCSearch("bn:", $ISBN)

;have script loop sleep until hotkey is pushed
$done = 0
While 1
	Sleep(0010)
	If $done = 1 Then
		ExitLoop
	EndIf
WEnd

;##### LOCAL FUNCTIONS #####
Func export()
	Sleep(0100)
	_SendEx("^{END}")
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("949  *recs-b;b3--;ov-.b") ;create bib record
	Sleep(0100)
	_SendEx("{ENTER}")
	_SendEx("949 1^dl hali ^dq 823 ^ds l ^di ") ;create item record
	Sleep(0100)
	$decide = InputBox("Barcode", "Scan in Barcode", "")
	Sleep(0100)
	_SendEx($decide)
	Sleep(0500)
	_SendEx("{F8}")
	Sleep(0100)
	_SendEx("{F5}")
	Sleep(0100)
	$done = 1 ;end loop
EndFunc
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: 949
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script takes the bib record number from the open Millennium record,
	closes the Millennium record, and inserts a 949 line in the OCLC Connexion
	record. The 949 field is set to overwrite the existing III bib record.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(bibliographic record open)
				OCLC Connexion Client v 2.20
					(on record screen)

 Last revised: 6/29/10
			   Updated #Include to include TS custom function library
			   Updated window controls for JRE version

 PLEASE NOTE - This script uses a custom UDF library called TSCustomFunction.
			   The script will not run properly (if launched from .au3 file)
			   if that file is not included in the include folder in the
			   AutoIt directory.

 Copyright (C): 2009 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.s

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------

;######### INCLUDES AND OPTIONS #########
#Include <TSCustomFunction.au3>
AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
TraySetIcon(@DesktopDir & "\CopyCat\Images\red 9.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_NUM

;################################ MAIN ROUTINE #################################
;focus on Millennium bib record screen
If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the bib record.")
	Exit
EndIf

;Get bib number from window title
$BIB_NUM = WinGetTitle("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
_SendEx("!{F4}")

;~ ;focus on Connexion
;~ WinActivate("OCLC Connexion")
;~ WinWaitActive("OCLC Connexion")

;focus on Notepad: toggle with code above to testing purposes
WinActivate("Untitled")
WinWaitActive("Untitled")

;send 949 line
_SendEx("^{END}")
Sleep(0100)
_SendEx("{ENTER}949{SPACE 2}*recs-b;b3--;ov-." & $BIB_NUM)
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: 504
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script prompts the user with questions about bibliography and index
	and inserts the 504 field (with the appropriate notes) in the OCLC Connexion
	record.

 Programs used: OCLC Connexion Client v 2.20
					(on record screen)

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
TraySetIcon(@DesktopDir & "\CopyCat\Images\blue 5.ico")

;######### DECLARE VARIABLES #########
Dim $PG_NUM, $INDEX

;################################ MAIN ROUTINE #################################
;focus on Connexion screen
;~ If WinExists("OCLC Connexion") Then
;~ 	WinActivate("OCLC Connexion")
;~ Else
;~ 	MsgBox(64, "nope", "Please open and log into Connexion. Click ok when you are logged in")
;~ EndIf

;testing script for the code block above
If WinExists("Untitled") Then
	WinActivate("Untitled")
Else
	MsgBox(64, "nope", "please open connexion and log in")
	EndIf


;Start prompts for information about page numbers and index
$PG_NUM = InputBox("Page Numbers?", "Enter page number(s) of bibliography, if appropriate.", "")
$INDEX = MsgBox(4, "Index?", "Does the item have an index?")
Sleep(0100)

;Send line to Connexion client record
_SendEx("{ENTER}504{SPACE 2}Includes bibliographical references")
If $PG_NUM <> "" Then
	_SendEx("{SPACE}(p. "& $PG_NUM &")")
EndIf
If $INDEX = 6 Then
	_SendEx("{SPACE}and index")
EndIf
_SendEx(".")
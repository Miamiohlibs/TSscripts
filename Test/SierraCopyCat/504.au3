#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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
TraySetIcon(@DesktopDir & "\SierraCopyCat\Images\blue 5.ico")

;######### DECLARE VARIABLES #########
Dim $PG_NUM, $INDEX, $BIB_REF
Dim $choose

;################################ MAIN ROUTINE #################################
;focus on Connexion screen
$choose = InputBox("Sierra or OCLC?", "Enter s for Sierra or o for OCLC")
if $choose = "o" Then
	If WinExists("OCLC Connexion") Then
		WinActivate("OCLC Connexion")
	Else
		MsgBox(64, "nope", "Please open and log into Connexion. Click ok when you are logged in")
	EndIf



;Start prompts for information about page numbers and index
$BIB_REF = MsgBox(4, "Bib ref?", "Does this item have bibliographic references?")
If $BIB_REF = 6 Then

$PG_NUM = InputBox("Page Numbers?", "Enter page number(s) of bibliography, if appropriate.", "")
$BIB_REF = " bibliographic references"
Else
	$BIB_REF = ""
EndIf
$INDEX = MsgBox(4, "Index?", "Does the item have an index?")
If $INDEX = 7 AND $BIB_REF = "" Then
	Exit
Else

Sleep(0100)
EndIf

;Send line to Connexion client record

_SendEx("{ENTER}504{SPACE 2}Includes"& $BIB_REF)
If $PG_NUM <> "" Then
	_SendEx("{SPACE}(pages "& $PG_NUM &")")
EndIf
If $INDEX = 6 Then
	If $BIB_REF = "" Then
		_SendEx("{SPACE}index")
	Else

	_SendEx("{SPACE}and index")
	EndIf
EndIf
_SendEx(".")

	ElseIf $choose = "s" Then
WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")

	$BIB_REF = MsgBox(4, "Bib ref?", "Does this item have bibliographic references?")
If $BIB_REF = 6 Then

$PG_NUM = InputBox("Page Numbers?", "Enter page number(s) of bibliography, if appropriate.", "")
$BIB_REF = " bibliographic references"
Else
	$BIB_REF = ""
EndIf
$INDEX = MsgBox(4, "Index?", "Does the item have an index?")
If $INDEX = 7 AND $BIB_REF = "" Then
	Exit
Else

Sleep(0100)
EndIf


_SendEx("{ENTER}n504{SPACE 2}Includes"& $BIB_REF)
If $PG_NUM <> "" Then
	_SendEx("{SPACE}(pages "& $PG_NUM &")")
EndIf
If $INDEX = 6 Then
	If $BIB_REF = "" Then
		_SendEx("{SPACE}index")
	Else

	_SendEx("{SPACE}and index")
	EndIf

EndIf
	EndIf
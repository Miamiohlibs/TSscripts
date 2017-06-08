#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: justautomated_1.5
 Script set: Just Automated names (Authority)

 Script Function: Replaces the following ME scripts
					-control+kp2
					-right mouse click

 Programs used: Internet Explorer (8+)
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
Dim $batchsearch_file_path
Dim $OTHER_filename
Dim $mil_tel_check
Dim $arn_check

;################################ MAIN ROUTINE #################################
$batchsearch_file_path = @DesktopDir & "\authorities\batchsearch\"
If Not FileExists($batchsearch_file_path) Then
	DirCreate(@DesktopDir & "\authorities\batchsearch")
EndIf

WinActivate("Save As")
WinWaitActive("Save As")
Sleep(0500)

_SendEx("!n")
_SendEx("^c")
$OTHER_filename = clipget()
_SendEx($batchsearch_file_path & $OTHER_filename)
Sleep(0400)

$mil_tel_check = StringRegExp($OTHER_filename, "(mil|tel)")
$arn_check = StringInStr($OTHER_filename, "arn")

Select
	Case $mil_tel_check = 1
		_SendEx("{ENTER}")
	Case $arn_check > 0
		_SendEx("{ENTER}")
		WinActivate("http://dog.lib.muohio.edu/~rshanley/authorities/")
		WinWaitActive("http://dog.lib.muohio.edu/~rshanley/authorities/")
		_SendEx("!{F4}")
	Case Else
		exit
EndSelect

;Going into OCLC...
If WinExists("OCLC Connexion") Then
	WinActivate("OCLC Connexion")
Else
	MsgBox(64, "nope", "Please open Connexion. Click ok when you are logged in")
EndIf
WinWaitActive("OCLC Connexion")
Sleep(0200)
_SendEx("!bp")
WinWaitActive("Process Batch")
Sleep(0200)
_SendEx("{DOWN}!k")
WinWaitActive("Enter Authority Batch Search Keys")
Sleep(0200)
_SendEx("!i")
WinWaitActive("Open")
Sleep(0200)
_SendEx("!n")
Sleep(0200)
_SendEx($batchsearch_file_path & $OTHER_filename)
Sleep(0200)
_SendEx("{ENTER}")
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
Sleep(0200)
_SendEx("y")
WinWaitActive("Enter Authority Batch Search Keys")
Sleep(0200)
_SendEx("!{F4}")
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
Sleep(0200)
_SendEx("y")
WinWaitActive("Process Batch")
Sleep(0200)
_SendEx("!s")
Sleep(0200)
_SendEx("{ENTER}")
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: justautomated_2.5
 Script set: Just Automated names (Authority)

 Script Function: Replaces the following ME scripts
					-control+kp4
				  *** To copy the file path of the .txt document before running
				  the script, push the Shift key down while right clicking on
				  the file name and select Copy As Path.

 Programs used: OCLC Connexion version 2.20

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
Dim $number_of_records
Dim $exported_records
Dim $i
Dim $number
Dim $oclc_arn_filepath

;################################ MAIN ROUTINE #################################

$oclc_arn_filepath = ClipGet()
$oclc_arn_filepath = StringReplace($oclc_arn_filepath, '"', "")

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
_SendEx($oclc_arn_filepath)
Sleep(0200)
_SendEx("{ENTER}")
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
Sleep(0500)
_SendEx("!y")
WinWaitActive("Enter Authority Batch Search Keys")
Sleep(0500)
_SendEx("!{F4}")
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
Sleep(0500)
_SendEx("!y")
WinWaitActive("Process Batch")
Sleep(0500)
_SendEx("!s")
Sleep(0200)
_SendEx("{ENTER}")

MsgBox(64, "Continue?", "Click OK when you want to continue the script.")

;time to export the file to III through OCLC
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
Sleep(0200)
_SendEx("!usl")
WinWaitActive("Search Local Authority Save File")
Sleep(0200)
_SendEx("!r")
Sleep(0200)
_SendEx("{ENTER}")
WinWaitActive("Local Authority Save File List (DefaultAuth.auth.db)")
_SendEx("{END}+{HOME}")

$number_of_records = InputBox("Number of records", "How many records to export?")

Do
	$number = StringIsInt($number_of_records)
	If StringIsInt($number_of_records) = 0 Then
		$number_of_records = InputBox("Number of records", "How many records to export?" & @CR & "Please enter a valid number.")
	EndIf
	$number = StringIsInt($number_of_records)
Until $number = 1
_SendEx("{F5}")

For $i = 1 To $number_of_records Step 1
	WinWaitActive("OCLC Gateway Export Status")
	Sleep(0500)
	WinKill("OCLC Gateway Export Status")
Next

$exported_records = MsgBox(4, "Export Complete?", "Did all the records export?")
Switch $exported_records
	Case 6
		_SendEx("!ac")
		Sleep(0200)
		WinActivate("OCLC Connexion")
		WinWaitActive("OCLC Connexion")
		Sleep(0200)
		_SendEx("y{ENTER}")
		WinWaitActive("Local Authority Save File List (DefaultAuth.auth.db)")
		WinKill("Local Authority Save File List (DefaultAuth.auth.db)")
	Case 7
		;Not sure what to do at this point. Usually staff person manually exports the rest
	EndSwitch
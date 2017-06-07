#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: justautomated_2
 Script set: Just Automated names (Authority)

 Script Function: Replaces the following ME scripts
					-control+kp3
					-control+kp4

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
Dim $oclc_authority_records_filepath
Dim $number_of_records
Dim $exported_records
Dim $i
Dim $number
Dim $oclc_arn_filepath

;################################ MAIN ROUTINE #################################

$oclc_authority_records_filepath = @DesktopDir & "\authorities\oclc_authority_records.txt"

If FileExists($oclc_authority_records_filepath) Then
	FileDelete($oclc_authority_records_filepath)
EndIf

If WinExists("OCLC Connexion") Then
	WinActivate("OCLC Connexion")
Else
	MsgBox(64, "nope", "Please open Connexion. Click ok when you are logged in")
EndIf

;create OCLC file for upload
WinWaitActive("OCLC Connexion")
Sleep(0200)
_SendEx("+{F3}")
Sleep(0200)
WinWaitActive("Search Local Authority Save File")
Sleep(0200)
_SendEx("!r")
Sleep(0200)
_SendEx("{ENTER}")
WinWaitActive("Local Authority Save File List (DefaultAuth.auth.db)")
Sleep(0200)
_SendEx("{END}+{HOME}")
Sleep(0200)
_SendEx("^p")

while Not FileExists($oclc_authority_records_filepath)
	Sleep(0100)
WEnd

WinActivate("Local Authority Save File List")
WinWaitActive("Local Authority Save File List")
Sleep(0500)
_SendEx("!ac")
Sleep(0500)
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
Sleep(0500)
_SendEx("y")

;upload OCLC file
_IECreate("http://dog.lib.muohio.edu/~rshanley/authorities/oclcupload.php")
WinWaitActive("http://dog.lib.muohio.edu/~rshanley/authorities/oclcupload.php")
WinWaitActive("Choose File to Upload")
Sleep(0500)

ControlSend("Choose File to Upload", "", "[CLASS:Edit; INSTANCE:1]", $oclc_authority_records_filepath)
;staff person needs to push enter to upload file
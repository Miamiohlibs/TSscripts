#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: justautomated_1
 Script set: Just Automated names (Authority)

 Script Function: Replaces the following ME scripts
					-control+kp1

 Programs used: Outlook (2010)
				Internet Explorer (8+)
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
Dim $email_text
Dim $headings_file_path
dim $mil_filename
Dim $tel_filename
Dim $filename_prefix
Dim $today_date
Dim $file_name
Dim $window_title
Dim $batchsearch_file_path
Dim $OTHER_filename
Dim $mil_tel_check
Dim $arn_check

;################################ MAIN ROUTINE #################################
_ClearBuffer()

;~ ;grab text from email, create local file
$window_title = "Mail from the Library"

ControlFocus($window_title, "", "[CLASS:_WwG; INSTANCE:1]")
ControlSend($window_title, "", "[CLASS:_WwG; INSTANCE:1]", "^a^c")

$email_text = ClipGet()

$headings_file_path = @DesktopDir & "\authorities\Headings reports\"


If Not FileExists($headings_file_path) Then
	DirCreate(@DesktopDir & "\authorities\Headings reports\")
EndIf

;determine if file is a millennium or telnet file, determine file prefix
$tel_filename = StringInStr($email_text, "Headings used for FIRST Time", 1)
$mil_filename = StringInStr($email_text, "Headings used for the first time", 1)

Select
	Case $mil_filename > 0
		$filename_prefix = "mil"
	Case $tel_filename > 0
		$filename_prefix = "tel"
	Case Else
		MsgBox(64, "Nope", "This email does not appear to be from the Headings Reports. The macro will now exit.")
		Exit
EndSelect

$today_date = @YEAR & @MON & @MDAY
$file_name = $headings_file_path & $filename_prefix & $today_date & ".txt"

FileWrite($file_name, $email_text)

;open web page to upload local file
_IECreate("http://dog.lib.muohio.edu/~rshanley/authorities/fileupload.php")
WinWaitActive("http://dog.lib.muohio.edu/~rshanley/authorities/fileupload.php")

;wait for the staff person to click on the Browse button
WinWaitActive("Choose File to Upload")
Sleep(0500)

ControlSend("Choose File to Upload", "", "[CLASS:Edit; INSTANCE:1]", $file_name)
;staff person needs to push enter to upload file


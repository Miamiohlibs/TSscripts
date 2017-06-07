#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: headings_prep
 Script set: Headings Files (Authority)

 Script Function: Replaces the following ME scripts
					F6
					F7
				  *** To copy the file path of the .txt document before running
				  the script, push the Shift key down while right clicking on
				  the file name and select Copy As Path.

 Programs used: Microsoft Word 2010

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
#include <Word.au3>

AutoItSetOption("WinTitleMatchMode", 2)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $NOT_filepath
Dim $NOT_contents
Dim $oWordApp

;################################ MAIN ROUTINE #################################
$NOT_filepath = ClipGet()
$NOT_filepath = StringReplace($NOT_filepath, '"', "")

$NOT_contents = FileRead($NOT_filepath)

$NOT_contents = StringRegExpReplace($NOT_contents, "Page[0-9]+", "")
$NOT_contents = StringReplace($NOT_contents, "|c", " ")
$NOT_contents = StringReplace($NOT_contents, "|e", @CR & "|e")
$NOT_contents = StringReplace($NOT_contents, "|4", @CR & "|4")
$NOT_contents = StringReplace($NOT_contents, "|o", @CR & "|o")
$NOT_contents = StringRegExpReplace($NOT_contents, "\r[[:space:]+]", @CR)
$NOT_contents = StringReplace($NOT_contents, "|n", @CR)
$NOT_contents = StringReplace($NOT_contents, ",|d", " ")
$NOT_contents = StringReplace($NOT_contents, "|d", " ")

;f7 macro start
;~ $NOT_contents = StringReplace($NOT_contents, @CR, " ")
$NOT_contents = StringReplace($NOT_contents, "|a", @CR & " ")
$NOT_contents = StringReplace($NOT_contents, "|b", " ")
$NOT_contents = StringReplace($NOT_contents, "|q", " ")
$NOT_contents = StringReplace($NOT_contents, "Field: ", @CR & @CR & "Field: ")
$NOT_contents = StringReplace($NOT_contents, "Indexed as ", @CR & "Indexed as ")
$NOT_contents = StringReplace($NOT_contents, "Preceded by ", @CR & "Preceded by ")
$NOT_contents = StringReplace($NOT_contents, "Followed by ", @CR & "Followed by ")
$NOT_contents = StringReplace($NOT_contents, "From: ", @CR & "From: ")
$NOT_contents = StringReplace($NOT_contents, "Catalog Date:", @CR & "Catalog Date:")
$NOT_contents = StringReplace($NOT_contents, "|t", @CR)
$NOT_contents = StringRegExpReplace($NOT_contents, "(,\|)[a-z0-9]", "")
$NOT_contents = StringReplace($NOT_contents, @CR & @CR, @CR)

FileOpen($NOT_filepath, 2)
FileWrite($NOT_filepath, $NOT_contents)
FileClose($NOT_filepath)

$oWordApp = _WordCreate("")
_WordDocOpen($oWordApp, $NOT_filepath)
Sleep(0200)
WinActivate("Microsoft Word")
WinWaitActive("Microsoft Word")
Send("^o")
WinWaitActive("Open")
Sleep(0500)
Send("!n")
Send($NOT_filepath)
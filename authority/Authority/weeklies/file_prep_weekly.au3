#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: file_prep_weekly
 Script set: Weeklies

 Script Function: Replaces the following ME scripts
					F6
				*** To copy the file path of the .txt document before running
				the script, push the Shift key down while right clicking on
				the file name and select Copy As Path.

 Programs used: Microsoft Word (2010)

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
#include <Word.au3>


AutoItSetOption("WinTitleMatchMode", 2)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
;~ AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $WEEKLIES_filepath
Dim $WEEKLIES_contents
Dim $oWordApp


;################################ MAIN ROUTINE #################################



$WEEKLIES_filepath = ClipGet()
$WEEKLIES_filepath = StringReplace($WEEKLIES_filepath, '"', "")

$WEEKLIES_contents = FileRead($WEEKLIES_filepath)

$WEEKLIES_contents = StringReplace($WEEKLIES_contents, "[sp", "[sh")
$WEEKLIES_contents = StringReplace($WEEKLIES_contents, @LF & @TAB, @CR & @CRLF)
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\[\w{3} Subd Geog\]", "")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\n.....110 1", @CR & @CR &"110 aa1")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\n\*\s+1", @CR & @CRLF &"1")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\n\([AC,]+\)\s+1", @CR & @CRLF &"1")
$WEEKLIES_contents = StringReplace($WEEKLIES_contents, "aa1", "1")
$WEEKLIES_contents = StringReplace($WEEKLIES_contents, @CR & "10 ", @CR & "110 ")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\s+100", @CR & @CRLF & @CRLF &"100")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\s+110", @CR & @CRLF & @CRLF &"110")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\s+150", @CR & @CRLF & @CRLF &"150")
$WEEKLIES_contents = StringRegExpReplace($WEEKLIES_contents, "\s+151", @CR & @CRLF & @CRLF &"151")

For $i = 0 To 6 Step 1
	$WEEKLIES_contents = StringReplace($WEEKLIES_contents, "   ", @CR & " ")
	$WEEKLIES_contents = StringReplace($WEEKLIES_contents, "  ", @CR & " ")
Next

$WEEKLIES_contents = StringReplace($WEEKLIES_contents, "[sh " & @CR, "[sh")
$WEEKLIES_contents = StringReplace($WEEKLIES_contents, @CR & " [sh", " [sh")
$WEEKLIES_contents = StringReplace($WEEKLIES_contents, "—", "--")

;~ $WEEKLIES_filepath = @DesktopDir & "\Weekly folder\"

;~ If Not FileExists($WEEKLIES_filepath) Then
;~ 	DirCreate(@DesktopDir & "\Weekly folder")
;~ EndIf

;~ $WEEKLIES_filename = @MON & @MDAY & @YEAR & ".txt"

FileOpen($WEEKLIES_filepath, 2)
FileWrite($WEEKLIES_filepath, $WEEKLIES_contents)
FileClose($WEEKLIES_filepath)

$oWordApp = _WordCreate("")
_WordDocOpen($oWordApp, $WEEKLIES_filepath)
Sleep(0200)
WinActivate("Microsoft Word")
WinWaitActive("Microsoft Word")
_SendEx("^o")
WinWaitActive("Open")
Sleep(0500)
_SendEx("!n")
_SendEx($WEEKLIES_filepath)
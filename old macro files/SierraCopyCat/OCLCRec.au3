#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: OCLCRec
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script searches OCLC Connexion by title, OCLC #, or ISBN. Title and
	OCLC # information are derived from the bib record. ISBN is from the
	physical item

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(bib record open)
				OCLC Connexion Client v 2.20
					(logged in)

 Last revised: 6/29/10
			   Updated #Include to include TS custom function library
			   Updated window controls for JRE version

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
TraySetIcon(@DesktopDir & "\SierraCopyCat\Images\Connex.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
dim $OCLC_NUM
Dim $decide
dim $245, $245_A2, $245DASH
Dim $ISBN

;################################ MAIN ROUTINE #################################
;focus on Millenium bib record
#cs If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the bib record and try again.")
	Exit
EndIf
#ce

; Focus on Sierra bib record
If WinExists("[REGEXPTITLE:Sierra � Miami University Libraries � .* � [b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:Sierra � Miami University Libraries � .* � [b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra � Miami University Libraries � .* � [b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the bib record and try again.")
	Exit
EndIf


;copy and parse bibliographic record data
_DataCopy()
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "fnord", 1)

;prompt asking what type of search user wants to do
$decide = InputBox("Type of Search", "What do you want to search by?" & @CR & "Enter the corresponding letter." & @CR & @CR & "O - OCLC Number" & @CR & "T - Title" & @CR & "I - ISBN", "")
$decide = StringUpper($decide)

Switch $decide
	Case "O" ;OCLC #
		$OCLC_NUM = _arrayItemString($BIB_ARRAY_MASTER, "OCLC #	001	")
		$OCLC_NUM =  StringTrimLeft($OCLC_NUM, 15)
		$OCLC_NUM =  StringStripWS($OCLC_NUM, 8)
		_OCLCSearch("{#}", $OCLC_NUM)
	Case "T" ;Title search (245 |a)
		$245 = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	245")
		$245_A2 = StringSplit($245, "|")
		$245 = _arrayItemString($245_A2, "TITLE	245")
		$245 = StringTrimLeft($245, 14)
		$245DASH = StringRegExp($245, "[:/]", 0) ;take out punctuation at the end of subfield a
		If $245DASH = 1 Then
			$245 = StringTrimRight($245, 2)
		EndIf
		_OCLCSearch("ti:", $245)
	Case "I" ;ISBN
		$ISBN = InputBox("ISBN Search", "Scan in the ISBN barcode", "") ;Scan isbn from the physical item
		_OCLCSearch("bn:", $ISBN)

EndSwitch
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: parse_headings
 Script set: Headings Files (Authority)

 Script Function: Replaces the following ME scripts
					parse headings for daily lists

 Programs used: None

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
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $heading_term ;t9 in ME script
Dim $first_tag_number
Dim $second_third_tag_numbers
Dim $OCLC_search_prefix
Dim $III_search_prefix
Dim $search_heading_term ;t14 in ME script
Dim $subdivision_array
Dim $number_of_subdivisions
Dim $subdivision_a
Dim $subdivision_1
Dim $subdivision_2
Dim $subdivision_3
Dim $subdivision_4
Dim $i, $x, $y

;################################ MAIN ROUTINE #################################

$heading_term = _LoadVar("$heading_term")
$first_tag_number = _LoadVar("$first_tag_number")
$second_third_tag_numbers = _LoadVar("$second_third_tag_numbers")

;test variable
;~ $heading_term = "FIELD: d650 0 Salamanders|zSouthern States|xIdentification "
;~ $first_tag_number = "6"
;~ $second_third_tag_numbers = "50"


If $first_tag_number = 6 Then
	$III_search_prefix = "d"
	Switch $second_third_tag_numbers
		Case "00", "90"
			$OCLC_search_prefix = "p"
		Case "10", "91"
			$OCLC_search_prefix = "c"
		Case "80", "85", "30"
			$OCLC_search_prefix = "t"
		Case "50", "51"
			$OCLC_search_prefix = "ll"
	EndSwitch
Else
	Switch $second_third_tag_numbers
		Case "00", "90"
			$III_search_prefix = "a"
			$OCLC_search_prefix = "p"
		Case "10", "91"
			$III_search_prefix = "a"
			$OCLC_search_prefix = "c"
		Case "80", "85"
			$III_search_prefix = "a"
			$OCLC_search_prefix = "t"
		Case "11"
			$III_search_prefix = "a"
			$OCLC_search_prefix = "c"
		Case "50", "51"
			$III_search_prefix = "d"
			$OCLC_search_prefix = "ll"
		Case "30"
			$III_search_prefix = "t"
			$OCLC_search_prefix = "t"
	EndSwitch
EndIf

If $second_third_tag_numbers <> "00" Then
	$heading_term = StringReplace($heading_term, ",", "")
EndIf

$heading_term = StringTrimLeft($heading_term, 5)
$search_heading_term = $heading_term

$heading_term = StringRegExpReplace($heading_term, "(\|[4evxz])", "--")
$heading_term = StringReplace($heading_term, "|y", " ")

$search_heading_term = StringReplace($search_heading_term, "|z", "|x", 1)
$search_heading_term = StringReplace($search_heading_term, "|z", "--")

;~ MsgBox(0, "", $search_heading_term)
;~ MsgBox(0, "", $heading_term)

$subdivision_array = StringSplit($search_heading_term, "|")

$number_of_subdivisions = UBound($subdivision_array)

If $number_of_subdivisions > 2 Then
	$x = 1
	For $i = 2 to $number_of_subdivisions Step 1
		$y = _ArrayToString($subdivision_array, @TAB, $i, $i)
		$y = StringTrimLeft($y, 1)
		Assign("subdivision_" & $x, $y)
		$x = $x + 1
	Next
EndIf

$subdivision_a = _ArrayToString($subdivision_array, @TAB, 1, 1)

_StoreVar("$heading_term")
_StoreVar("$OCLC_search_prefix")
_StoreVar("$III_search_prefix")
_StoreVar("$subdivision_a")
_StoreVar("$subdivision_1")
_StoreVar("$subdivision_2")
_StoreVar("$subdivision_3")
_StoreVar("$subdivision_4")

#cs
notes -
	array starts at 0
	element 1 is subfield a
	elements after 1 are subdivisions
	the subfield character is the first one in the string
#ce

;~ _ArrayDisplay($subdivision_array)
;~ MsgBox(0, '', $number_of_subdivisions)
;~ MsgBox(64, "", $subdivision_1)
;~ MsgBox(64, "", $subdivision_2)
;~ MsgBox(64, "", $subdivision_3)
;~ MsgBox(64, "", $subdivision_4)
;~ MsgBox(0, "", $subdivision_a)
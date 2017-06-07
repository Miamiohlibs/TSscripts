#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: OCLC049and949
 Script set: Order (order)

 Script Function:
	This script fills out the 049 and 949 fields in the OCLC Connexion record.

	##### NOTE ##### Place cursor at the end of the 049 field before
	running the script

 Programs used: OCLC Connexion Client v 2.20
					(record open)

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
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $049
Dim $LOCATION
Dim $decide
Dim $FUND
Dim $PRICE
Dim $Date
Dim $VENDOR
Dim $multi
Dim $YEAR
Dim $REF
Dim $AMAZON_SEARCH
Dim $VENDCODE
Dim $SF_CHECK


;################################ MAIN ROUTINE #################################
;load variables from gobitoIII-OCLC
$FUND = _LoadVar("$FUND")
$multi = _LoadVar("$multi")
$PRICE = _LoadVar("$COST")
$REF = _LoadVar("$REF")
$AMAZON_SEARCH = _LoadVar("$AMAZON_SEARCH")


$YEAR = @YEAR
$YEAR = StringTrimLeft($YEAR, 2)
$Date = @MON & "-" & @MDAY & "-" & $YEAR

;focus on connexion window
If WinExists("OCLC Connexion") Then
	WinActivate("OCLC Connexion")
Else
	MsgBox(64, "nope", "Please open and log into Connexion. Click ok when you are logged in")
EndIf

;determine location of item for 049 field

	$049 = InputBox("Location options", "Where will the item go to? Type in the corresponding letter." & @CR & "A: Art/Arch" & @CR & "G: Gov Docs" & @CR & "J: IMC Juv" & @CR & "K: King" & @CR & "M: Music"  & @CR & "S: Science" & @CR & "MT: Middletown", "", "", 200, 230)
	$049 = StringUpper($049)
	If $REF = 1 Then
		Switch $049
			Case "A"
				$LOCATION = "aar"
			Case "G"
				$LOCATION = "docr"
			Case "J"
				$LOCATION = "imr"
			Case "K"
				$LOCATION = "kngr"
			Case "M"
				$LOCATION = "mur"
			Case "S"
				$LOCATION = "scr"
			Case "MT"
				$LOCATION = "mdr"
		EndSwitch
	Else
		Switch $049
			Case "A"
				$LOCATION = "aali"
			Case "G"
				$LOCATION = "docr" ;keeping this as reserve...
			Case "J"
				$LOCATION = "imju"
			Case "K"
				$LOCATION = "kngli"
			Case "M"
				$LOCATION = "muli"
			Case "S"
				$LOCATION = "scli"
			Case "MT"
				$LOCATION = "mdli"
		EndSwitch
	EndIf

_StoreVar("$LOCATION") ;store location code for below script
Run(@DesktopDir & "\order\Bib - check ILOC.exe") ;determine 049 code
While ProcessExists("Bib - check ILOC.exe")
	Sleep(0400)
WEnd

$049 = _LoadVar("$049")
If $LOCATION = "onlbk" Then
	$049 = "MIAA"
EndIf

;edit 049 line
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
_SendEx("+{HOME}")
Sleep(0400)
_SendEx($049)
Sleep(0400)

;check fund for special fund number
$SF_CHECK = StringRegExp($FUND, "(4\D{1,3})")

;create 949 line
_SendEx("^{END}")
Sleep(0100)

;IMC JUV 949 line
If $LOCATION = "imju" Then
	WinActivate("OCLC Connexion")
	WinWaitActive("OCLC Connexion")
	_SendEx("{ENTER}949{SPACE 2}*recs-bo;po-n;ot-r;b3--;ins-acq;fm-u;rl-b;bl-b;br-imju;fd-67/1;ep-")
	_SendEx($PRICE)
	Sleep(0100)
	_SendEx(";od-" & $Date & ";vd-amazn;")
	Exit
EndIf

;Regular 949 line
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
_SendEx("{ENTER}949{SPACE 2}*recs-bo;")
If $AMAZON_SEARCH = 1 Then
	_SendEx("ot-r") ; Amazon order
Else
	_SendEx("po-n;ot-n") ;oxford/middletown
EndIf
_SendEX(";b3--;ins-acq;")
_SendEx("fm-u;rl-")
If $LOCATION = "mdli" Then
	_SendEx("m;bl-b") ;middletown
Elseif $AMAZON_SEARCH = 1 Then
	_SendEx("b;bl-b") ;amazon
Else
	_SendEx("n") ;oxford
EndIf
_SendEx(";br-")
$decide = MsgBox(4, "Correct location code?", "Is " & $LOCATION & " the correct location code?")
	Switch $decide
		Case 6
		Case 7
			$LOCATION = InputBox("Manual location entry", "Please enter the location code below.")
		EndSwitch
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
_SendEx($LOCATION)
If $LOCATION = "aar" OR $LOCATION = "imr" OR $LOCATION = "kngr" Or $LOCATION = "mur" Or $LOCATION = "scr" Or $LOCATION = "mdr" or $LOCATION = "docr"Then
	$REF = True
	_StoreVar ("$REF")
EndIf
_StoreVar ("$LOCATION")
Sleep(0100)
_SendEx(";fd-")
If $049 = "OHMM" Then
	_SendEx("M") ;middletown has an M in front of fund codes
	If $FUND = "DIS" Then ;in case they used a DIS code in GOBI
		$FUND = "56"
		$decide = MsgBox(4, "Correct fund code?", "Is " & $FUND & " the correct fund code?")
		Switch $FUND
			Case 6
			Case 7
				$FUND = InputBox("Manual fund entry", "Please enter the fund code below.")
		EndSwitch
	EndIf
EndIf
If $multi = 1 Then
	_SendEx("") ;if there are multiple funds listed then the script will leave this part blank. The OCLCtoIII script will fill out the funds in the order record.
Else
	_SendEx($FUND)
EndIf
Sleep(0100)
If $049 <> "OHMM" AND $SF_CHECK = 0 Then
	_SendEx("/1") ;this changes every FY. This really needs to be automated so I don't have to go back each FY and change this.
ElseIf $SF_CHECK = 1 AND $049 <> "OHMM" Then
	Sleep(0100)
EndIf
_SendEx(";ep-")
_SendEx($PRICE)
Sleep(0100)
_SendEx(";od-" & $Date & ";vd-")
If $AMAZON_SEARCH = 1 Then
	$VENDCODE = "amazn";amazon vendor code
Else
	$VENDCODE = "ybp" ;ybp code is ybp
EndIf
$VENDOR = InputBox("Vendor", "Enter the vendor code.", $VENDCODE)
WinActivate("OCLC Connexion")
WinWaitActive("OCLC Connexion")
_SendEx($VENDOR & ";")
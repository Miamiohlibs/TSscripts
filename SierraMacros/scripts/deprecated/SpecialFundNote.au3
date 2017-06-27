#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=images\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: SpecialFundNote
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script searches the order record for special fund codes. If one is
	found, the script goes into the bib record and

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(order record open)
					### Remember to switch the summary view back to the item
					record after running this script!

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
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("MustDeclareVars", 1)
TraySetIcon(@DesktopDir & "\CopyCat\Images\black S.ico")

;######### DECLARE VARIABLES #########
Dim $winsize
dim $ORD_REC, $ORD_REC_PREP, $ORD_ARRAY_MASTER
Dim $row4, $row4_prep, $row4_MASTER
Dim $STATUS
Dim $FUND
dim $multi_a, $multi
Dim $SF_NAME, $SF_590, $SF_7XX
Dim $decide
Dim $slash, $4, $FUND_A

;################################ MAIN ROUTINE #################################
;focus on Millennium order record
;If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;	WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the order record.")
;	Exit
;EndIf

;/Emily/
If WinExists("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf

;click on data fields
;$winsize = WinGetPos("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;/Emily/
$winsize = WinGetPos("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]")

_WinClick($winsize)
;WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;/Emily/
WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;copy and parse order record data
_DataCopy()
$ORD_REC = ClipGet()
$ORD_REC_PREP = StringRegExpReplace ($ORD_REC, "[\r\n]+", "fnord")
$ORD_ARRAY_MASTER = StringSplit($ORD_REC_PREP, "fnord", 1)

;~ ##### start fixed data setting #####
;row 4 includes claim, fund, status
$row4_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 4)
$STATUS = _ArrayPop($row4_MASTER)
$FUND = _ArrayPop($row4_MASTER)
$FUND = StringTrimLeft($FUND, 4)
;~ ##### end fixed data setting #####

;check for multi funds note
If $FUND = "multi" Then
 $multi = _arrayItemString($ORD_ARRAY_MASTER, "FUNDS	")
 $multi = StringTrimLeft($multi, 6)
EndIf

;go into bib record
;WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;/Emily/
WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0100)
_SendEx("!g")
Sleep(0100)
_SendEx("e")
;WinActivate("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;/Emily/
WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Sleep(0100)
_SendEx("^{END}")
Sleep(0100)

;prepare fund variable for Special Fund List script
#cs
	Fun happy fund parsing workaround!
	It turns out that the new fund formatting for 2009B played havoc with
	parsing out fund codes for single fund purchases. Instead of having a
	space between the fund code and the fund name, they are lumped together
	in one solid line. I have a workaround that checks for the first digit
	of the fund line to determine if it's a special fund code. This in part
	works because all but one special fund start with "4" and none have
	a slash in them.

	As far as I know, the multi fund parsing has not run into the same
	issue from the 2009B upgrade, so I have left the original parsing
	intact.
#ce
$slash = StringInStr($FUND, "/")
$4 = StringLeft($FUND, 1)
If $slash = 0 And $4 = 4 Then
	$FUND_A = StringRegExp($FUND, "(4\D{1,3})", 1)
	$FUND = _ArrayPop($FUND_A)
EndIf
$FUND = StringStripWS($FUND, 8)
_StoreVar("$FUND")

If $FUND <> "multi" then
	Run(@DesktopDir & "\SierraCopyCat\Special Fund List.exe")
		While ProcessExists("Special Fund List.exe")
			Sleep(0400)
		WEnd
	$SF_NAME = _LoadVar("$SF_NAME") ;determined from Special Fund List script
	If $SF_NAME <> "none" Then ;enter 590 and 79x fields
		$SF_590 = _LoadVar("$SF_590")
		$SF_7XX = _LoadVar("$SF_7XX")
		;WinActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
		;/Emily/
		WinActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")

		_SendEx("^{END}")
		Sleep(0100)
		_SendEx("{ENTER}n")
		Sleep(0100)
		_SendEx("590{TAB}")
		sleep(0100)
		$decide = InputBox("Location?", "What location does the item belong to? Type in the corresponding letter." & @CR & "A: Art/Arch" & @CR & "I: IMC" & @CR & "K: King"  & @CR & "S: Science", "", "", 200, 210)
		$decide = StringUpper($decide)
		Switch $decide
			Case "kngl"
				_SendEx("King copy{SPACE}")
			Case "scl"
				_SendEx("Science copy{SPACE}")
			Case "aal"
				_SendEx("ArtArch copy{SPACE}")
			Case "imc"
				_SendEx("IMC copy{SPACE}")
		EndSwitch
		_SendEx($SF_590)
		Sleep(0100)
		_SendEx("^{END}")
		Sleep(0100)
		_SendEx("{ENTER}b")
		Sleep(0100)
		_SendEx($SF_7XX)
	EndIf
ElseIf $FUND = "multi" Then ;if there are multiple fund numbers
	$multi = _LoadVar("$multi")
	$multi_A = StringSplit($multi, ",") ;place in array, then loop through each array element
	Do
		$FUND = _ArrayPop($multi_A)
		$FUND = StringStripWS($FUND, 8)
			Run(@DesktopDir & "\SierraCopyCat\Special Fund List.exe")
				While ProcessExists("Special Fund List.exe")
					Sleep(0400)
				WEnd
			$SF_NAME = _LoadVar("$SF_NAME")
		If $SF_NAME <> "none" Then
			$SF_590 = _LoadVar("$SF_590")
			$SF_7XX = _LoadVar("$SF_7XX")
			;$SF_7XX = WinActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
			;/Emily/
			$SF_7XX = WinActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")

			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}n")
			Sleep(0100)
			_SendEx("590{TAB}")
			sleep(0100)
			$decide = InputBox("Location?", "What location does the item belong to? Type in the corresponding letter." & @CR & "A: Art/Arch" & @CR & "I: IMC" & @CR & "K: King"  & @CR & "S: Science", "", "", 200, 210)
			$decide = StringUpper($decide)
			Switch $decide
				Case "kngl"
					_SendEx("King copy{SPACE}")
				Case "scl"
					_SendEx("Science copy{SPACE}")
				Case "aal"
					_SendEx("ArtArch copy{SPACE}")
				Case "imc"
					_SendEx("IMC copy{SPACE}")
			EndSwitch
			_SendEx($SF_590)
			Sleep(0100)
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}b")
			Sleep(0100)
			_SendEx($SF_7XX)
		EndIf
	Until $multi_A = ""
EndIf
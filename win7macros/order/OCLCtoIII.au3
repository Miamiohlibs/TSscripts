#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: OCLCtoIII
 Script set: Order (order)

 Script Function:
	This script reads the information from the OCLC Connexion Gateway Status box.
	The script determines from the text, along with other variables declared from
	previous scripts, if it needs to go into the order record for further editing.
	In either case, the script copies the PO number to the clipboard for future
	use and closes the Status box.

	##### NOTE ##### While the script saves the order record, it cannot automate
	the entries for the various pop up boxes that come with saving the record.
	The user must manually close these pop up boxes.

 Programs used:	Millennium Acquisitions Module JRE v 1.6.0_02
					(Main search screen open)
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
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("MouseCoordMode", 0)

;######### DECLARE VARIABLES #########
Dim $winsize
Dim $OCLC_TXT
Dim $OCLC_TXT_PREP, $OCLC_TXT_ARRAY_MASTER
Dim $ORD_NUM
Dim $ON_HOLD
Dim $decide
Dim $SF
Dim $FUND, $FUND_TEMP, $FUND1
Dim $SF_NAME, $SF_NAME_1
Dim $ORD_REC, $ORD_REC_PREP, $ORD_ARRAY_MASTER, $row4_MASTER
Dim $STATUS
Dim $FUND
Dim $multi, $multi_A
Dim $note
Dim $REF
Dim $HOLD
Dim $LOCATION
Dim $OCLC_TITLE
Dim $DIFF_ED

;################################ MAIN ROUTINE #################################
;load variables from gobitoIII-OCLC and OCLC049and949
$REF = _LoadVar("$REF")
$FUND = _LoadVar("$FUND")
$multi = _LoadVar("$multi")
$HOLD = _LoadVar("$HOLD")
$LOCATION = _LoadVar ("$LOCATION")
$DIFF_ED = _LoadVar("$DIFF_ED")

;focus on gateway export status box
If WinExists("[TITLE:OCLC Gateway Export Status; CLASS:#32770]") Then
	WinActivate("[TITLE:OCLC Gateway Export Status; CLASS:#32770]")
Else
	MsgBox(64, "nope", "Please log into OCLC")
	Exit
EndIf

;grab status box text and parse information
$OCLC_TXT = WinGetText("[TITLE:OCLC Gateway Export Status; CLASS:#32770]")
$OCLC_TXT_PREP = StringRegExpReplace ($OCLC_TXT, "[\r\n\t]+", "fnord")
$OCLC_TXT_ARRAY_MASTER = StringSplit($OCLC_TXT_PREP, "fnord", 1)

$ON_HOLD = _arrayItemString($OCLC_TXT_ARRAY_MASTER, "hold for order record")

$ORD_NUM = _arrayItemString($OCLC_TXT_ARRAY_MASTER, "  ORDER          #")
$ORD_NUM = StringTrimLeft($ORD_NUM, 20)
$ORD_NUM = StringStripWS($ORD_NUM, 8)

;~ Search for special funds in fund/multi variables
If $multi = 0 Then
	Run(@DesktopDir & "\order\Special Fund List.exe")
	While ProcessExists("Special Fund List.exe")
		Sleep(0400)
	WEnd
	$SF_NAME = _LoadVar("$SF_NAME")
	If $SF_NAME <> "none" Then
		$SF = 1
	EndIf
Elseif $multi = 1 Then
	$multi_A = StringSplit($FUND, ",")
	$FUND_TEMP = $FUND
	Do
		$FUND = _ArrayPop($multi_A)
		$FUND = StringStripWS($FUND, 8)
			Run(@DesktopDir & "\Receipt\Special Fund List.exe")
				While ProcessExists("Special Fund List.exe")
					Sleep(0400)
				WEnd
			$SF_NAME = _LoadVar("$SF_NAME")
				If $SF_NAME <> "none" Then
					$SF = 1
				EndIf
	Until $multi_A = ""
	$FUND = $FUND_TEMP
EndIf

;grab title from OCLC window title:
$OCLC_TITLE = WinGetTitle("OCLC Connexion")
$OCLC_TITLE = StringReplace($OCLC_TITLE, "OCLC Connexion - [Online WorldCat: ", "")
$OCLC_TITLE = StringTrimRight($OCLC_TITLE, 1)

;If nothing else needs to be done to the order record, exit script.
If $ON_HOLD = -1 And $REF = 0 And $multi = 0 And $HOLD = "" And $SF = 0 Then
	$ORD_NUM = StringTrimLeft($ORD_NUM, 1)
	ClipPut($ORD_NUM)
	WinActivate("[TITLE:OCLC Gateway Export Status; CLASS:#32770]")
	WinKill("[TITLE:OCLC Gateway Export Status; CLASS:#32770]")
	_ClearBuffer()
	exit
EndIf

;if notes need to be added or if the record is on hold, script will search for
;order record by title from OCLC record
;focus on Millennium window, search by title from oclc record
If WinExists("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]") Then
	WinActivate("[TITLE:Millennium Acquisitions; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "nope", "Please log into Millennium.")
EndIf
_IIIsearch("t", $OCLC_TITLE)

;check to see if there is an order record
If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
		WinActivate("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	Else
		MsgBox(64, "nope", "Please go into the order record." & @CR & "Make sure that the item can be accessed by a title search.")
EndIf

WinWaitActive("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;click on data fields
$winsize = WinGetPos("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
WinWaitActive("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")

Sleep(0300)
ClipPut("")
Sleep(0100)
_SendEx("^{HOME}")
Sleep(0700)

;take status off off hold
If $ON_HOLD > -1 Then
	Sleep(0300)
	_SendEx("^{HOME}")
	Sleep(0300)
	_SendEx("{TAB 4}")
	Sleep(0300)
	_SendEx("{DOWN 3}")
	Sleep(0300)
	_SendEx("o")
	Sleep(0700)
EndIf
;previous notes from gobi info

;make note for reference items
If $REF = 1 Then
	orderNote("Cat for Ref")
	Sleep(0700)
EndIf

;make note for diff ed items
If $DIFF_ED <> "" Then
	orderNote("DIFF ED " & $DIFF_ED)
	Sleep(0700)
EndIf

;enter multiple fund numbers
If $multi = 1 Then
	$multi = StringRegExpReplace($multi, "[/,]", "/")
	$multi_A = StringSplit($multi, "/")
	_SendEx("{TAB 2}")
	Sleep(0400)
	_SendEx("{DOWN 3}")
	Sleep(0400)
	_SendEx("^e")
	Do
		$FUND1 = _ArrayPop($multi_A)
		$FUND1 = StringStripWS($FUND1, 8)
		_SendEx($LOCATION)
		Sleep(0400)
		_SendEx("{TAB}")
		Sleep(0400)
		_SendEx($FUND1)
		Sleep(0400)
		_SendEx("!a")
	Until $multi_A = ""
	Sleep(0400)
	_SendEx("!r")
	Sleep(0400)
	_SendEx("!r")
	Sleep(0400)
	_SendEx("!o")
EndIf

;enter special funds notes
_DataCopy()
$ORD_REC = ClipGet()
$ORD_REC_PREP = StringRegExpReplace ($ORD_REC, "[\r\n]+", "fnord")
$ORD_ARRAY_MASTER = StringSplit($ORD_REC_PREP, "fnord", 1)
$row4_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 4)
$STATUS = _ArrayPop($row4_MASTER)
$FUND = _ArrayPop($row4_MASTER)
$FUND = StringTrimLeft($FUND, 4)
If $FUND = "multi" Then
	$multi = _arrayItemString($ORD_ARRAY_MASTER, "FUNDS	")
	$multi = StringTrimLeft($multi, 6)
EndIf
If $FUND <> "multi" then
	Run(@DesktopDir & "\order\Special Fund List.exe")
	While ProcessExists("Special Fund List.exe")
		Sleep(0400)
	WEnd
	$SF_NAME = _LoadVar("$SF_NAME")
	If $SF_NAME <> "none" Then
		$SF_NAME_1 = "BOOKPLATE & DONOR NOTES REQUIRED FOR " & $SF_NAME & " FUND"
		orderNote($SF_NAME_1)
	EndIf
	ElseIf $FUND = "multi" Then
		$multi_A = StringSplit($multi, ",")
		Do
			$FUND = _ArrayPop($multi_A)
			$FUND = StringStripWS($FUND, 8)
			Run(@DesktopDir & "\order\Special Fund List.exe")
				While ProcessExists("Special Fund List.exe")
					Sleep(0400)
				WEnd
			$SF_NAME = _LoadVar("$SF_NAME")
			If $SF_NAME <> "none" Then
				$SF_NAME_1 = "BOOKPLATE & DONOR NOTES REQUIRED FOR " & $SF_NAME & " FUND"
				orderNote($SF_NAME_1)
			EndIf
		Until $multi_A = ""
EndIf

;see if there are any other notes that need to be entered
Do
	$decide = InputBox("Order note options", "What note needs to be entered? Type in the corresponding letter." & @CR & "A: Accompanying materials" & @CR & "D: DIFF ED" & @CR & "O: OTHER", "", "", 160, 180)
	$decide = StringUpper($decide)
	Switch $decide
		Case "A"
			$note = InputBox("Accompanying note", "Enter estimated accompanying material below.")
			orderNote($note)
		Case "D"
			$note = InputBox("Diff ed note", "Enter call # of earlier ed below.")
			orderNote("DIFF ED " & $note)
		Case "O"
			$note = InputBox("Order note", "Enter order note below")
			orderNote($note)
	EndSwitch
Until $decide = @error

Sleep(0100)
_SendEx("^s")
	;need to get window information about PO queue...

;~ 	WinWait("[TITLE:WARNING; CLASS:SunAwtDialog]", "")
;~ 	_SendEx("y")

;copy order record number
$ORD_NUM = StringTrimLeft($ORD_NUM, 1)
ClipPut($ORD_NUM)

;close OCLC popup window
If WinExists("[TITLE:OCLC Gateway Export Status; CLASS:#32770]") Then
	WinActivate("[TITLE:OCLC Gateway Export Status; CLASS:#32770]")
	WinKill("[TITLE:OCLC Gateway Export Status; CLASS:#32770]")
Else
	MsgBox(64, "nope", "Please go to Connexion.")
EndIf
Sleep(0300)

WinActivate("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;if the item has a hold request indicated in GOBI
If $HOLD <> "" Then
	MsgBox(64, "Hold information", "The item has the following hold note:" & @cr & $HOLD & @cr & "Please place the item on hold. Click OK when finished")
EndIf

_ClearBuffer()

;~ ######################## Local Functions #######################

Func orderNote($y)
	WinActivate("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:\A[bo][0-9ax]{8}; CLASS:SunAwtFrame]")
	_SendEx("^{END}")
	_SendEx("^{END}")
	Sleep(0700)
	_SendEx("{ENTER}")
	Sleep(0400)
	_SendEx("z")
	Sleep(0400)
	_SendEx($y)
EndFunc
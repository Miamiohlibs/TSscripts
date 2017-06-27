#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\Images\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.4.4
 Author: Martin Patrick, Academic Resident Librarian in Technical Services
 Miami University
 patricm@miamioh.edu

 Name of script: sword.exe
 Script set: Withdrawals

 Script Function:
	This script prompts the user through the process of sending an item to
	offsite long-term storage called SWORD (Southwest Ohio Regional Depository)

 Programs used: n/a

 Last revised:5/2015


 Copyright (C): 2015 by Miami University Libraries.  Libraries
 may freely use and adapt this macro with due credit.  Commercial use
 prohibited without written permission.

 For more information about the functions/commands below, please see the online
 AutoIt help file at http://www.autoitscript.com/autoit3/docs/

#ce ----------------------------------------------------------------------------
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#Include <String.au3>
#include <Array.au3>
#Include <Misc.au3>
#Include <TSCustomFunction.au3>

;make sure bib record is active
;Search LOCATIONS from bib_rec for swdep if not add it
Dim $location
Dim $clip
Dim $nextStep
Dim $year



$decision = msgbox(1, "Bib Record", "Please open the Bib Record. Press OK when ready, or Cancel to abort.")
If $decision = 2 Then
	Exit
EndIf

;prompt for the location to which the item should be sent
$location = InputBox("From where?", "Enter the number that corresponds to location." & @crlf & "1 - King Mono" & @CRLF & "2 - King Per" & @CRLF & "3 - BEST" & @CRLF & "4 - Music" & @crlf & "5 - Art/Arch" & @crlf & "6 - Documents" & @crlf & "7 - IMC" & @crlf & "8 - Hamilton" & @crlf & "9 - Withdraw", "", "", 75, 300)

If $location = 1 Then
	$iloc = "rg2tk"
ElseIf $location = 2 Then
	$iloc = "rg2tk"
ElseIf $location = 3 Then
	$iloc = "rg2ts"
ElseIf $location = 4 Then
	$iloc = "rg2tm"
ElseIf $location = 5 Then
	$iloc = "rg2ta"
Elseif $location = 6 Then
	$iloc = "rg2td"
ElseIf $location = 7 Then
	$iloc = "rg2ti"
ElseIf $location = 8 Then
	$iloc = "rg2th"
ElseIf $location = 9 Then
	$iloc = "wd"
EndIf
_StoreVar("$iloc")


;this part of code will simply witdraw the item
If $iloc = "wd" Then
	$nextStep = MsgBox(4, "Suppress?", "Does this bib have just one copy?")
	If $nextStep = 7 Then
		msgbox(1, "Delete MIAA", "Delete the Oxford location codes in the 049 field and the Location field header. Press OK when finished to proceed or Cancel to abort.")
		If $nextStep = 2 Then
	Exit
EndIf
	ElseIf $nextStep = 6 Then
		msgbox(1, "Supress", "Change BCODE3 to s, and then press OK to proceed or Cancel to abort.")
		If $nextStep = 2 Then
	Exit
EndIf
	EndIf
Sleep(0200)
$nextStep = msgbox(1, "OCLC", "Please copy the OCLC to a text file in order to delete holdings later. Press OK to proceed or Cancel to Abort.")
If $nextStep = 2 Then
	Exit
EndIf

_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("i")
Sleep(0500)
_SendEx("!l")
Sleep(0500)
If WinExists("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Oops", "Please open an item record.")
	Exit
EndIf

_SendEx("^a")
Sleep(0100)
_SendEx("^c")
Sleep(0400)
$clip = ClipGet()
If $clip = "" Then
	MsgBox(64, "Oops", "Please open an item record.")
	Exit
EndIf

_SendEx("^{HOME}")
Sleep(0200)
_SendEx("{DOWN 2}")
Sleep(0400)
_SendEx("s")
Sleep(0200)
_SendEx("{TAB 2}")
Sleep(0300)
_SendEx("k")
Sleep(0200)
_SendEx("{ENTER}")
Sleep(0200)
_SendEx("x")
Sleep(0200)
_SendEx("{TAB}")
Sleep(0300)
$year = InputBox(0, "Please enter the year in the format MMM YYYY (AUG 2015)")
_SendEx("WITHDRAWN ")
Sleep(0200)
_SendEx($year)
Sleep(0400)
_SendEx("^s")
Exit
EndIf





;this part of code will process an item for sword
Sleep(0400)
WinActivate("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
$winsize = WinGetPos("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
Sleep(0400)
_SendEx("^{HOME}")
Sleep(0500)
_SendEx("{DOWN 2}")
Sleep(0500)
_SendEx("^e")
Sleep(0500)
_sendEx("!a")
Sleep(0300)
_sendex("swdep")
Sleep(0700)


$nextStep = MsgBox(65, "Ok", "Delete location code from MARC 049 and Location header edit box that's open, and then press OK to move to the item record. Press cancel to abort.")
if $nextStep = 2 Then
	Exit
Else
	;switch to item record view
_SendEx("^s")
Sleep(0200)
_SendEx("^s")
Sleep(0200)
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("i")
Sleep(0500)
EndIf

$nextStep = MsgBox(65, "Item", "Please choose the item, open the item record, and then click OK to begin process. Press cancel to abort.")
If $nextStep = 2 Then
	Exit
EndIf

;this section checks to see if an bib record is open or not
If WinExists("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Oops", "Please open an item record.")
	Exit
EndIf

_SendEx("^a")
Sleep(0100)
_SendEx("^c")
Sleep(0400)
$clip = ClipGet()
If $clip = "" Then
	MsgBox(64, "Oops", "Please open an item record.")
	Exit
EndIf
; end item open check

_SendEx("^{HOME}")
Sleep(0200)
_SendEx("{TAB}")
Sleep(0200)
_SendEx($iloc)
Sleep(0400)
_SendEx("{TAB 5}")
Sleep(0400)
_SendEx("k")
Sleep(0200)
_SendEx("{ENTER}")
Sleep(0200)
_SendEx("x")
Sleep(0200)
_SendEx("{TAB}")
Sleep(0300)
_SendEx("@SWORD ")
Sleep(0200)
$year = InputBox(0, "Please enter the fiscal year in the format 2014/15")
_SendEx($year)
Sleep(0400)
Exit
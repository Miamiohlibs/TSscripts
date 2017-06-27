#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\Images\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.4.4
 Author: Martin Patrick, Academic Resident Librarian in Technical Services
 Miami University
 patricm@miamioh.edu

 Name of script: withdraw.exe
 Script set: Withdrawals

 Script Function:
	This script prompts the user through the process of updating a bib and item record
	to withdrawn status.

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
#Include <Date.au3>

Dim $date, $date_month, $date_year
Dim $clip
Dim $decision

AutoItSetOption("WinTextMatchMode", 1)
AutoItSetOption("WinTitleMatchMode", 3)



$decision = msgbox(1, "Bib Record", "Please open the Bib Record. Press OK when ready, or Cancel to abort.")
If $decision = 2 Then
	Exit
EndIf
#cs
If WinExists("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "nope", "Please open the bib record.")
	Exit
EndIf
#ce



$nextStep = MsgBox(4, "Suppress?", "Does this bib have just one copy?")
	If $nextStep = 7 Then
		msgbox(1, "Delete MIAA", "Delete the location codes in the 049 field and the Location field header. Press OK when finished to proceed or Cancel to abort.")
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
$nextStep = msgbox(1, "OCLC", "Please copy the OCLC to a text file in order to update holdings later. Press OK to proceed or Cancel to Abort.")
If $nextStep = 2 Then
	Exit
EndIf
Sleep(0200)

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
_SendEx("!l")
Sleep(0500)


If WinExists("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "nope", "Please open the item record.")
	Exit
EndIf
Sleep(0300)
$winsize = WinGetPos("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)

Sleep(0300)

_SendEx("^{HOME}")
Sleep(0300)
_SendEx("{DOWN 2}")
Sleep(0300)
_SendEx("s")
Sleep(0400)
_SendEx("{TAB 2}")
Sleep(0400)
_SendEx("k")
Sleep(0400)
_SendEx("{ENTER}")
Sleep(0300)
_SendEx("x")
Sleep(0300)
_SendEx("{TAB}")
Sleep(0100)
_SendEx("{TAB}")
$date = InputBox("Year", "Enter the date in MMM YYYY format (MAY 2015)")
_SendEx("WITHDRAWN " & $date)
Exit
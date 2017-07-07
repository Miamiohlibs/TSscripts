#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Craig Boman, Discovery Services Librarian, Miami University
		 bomanca@miamioh.edu or craig.boman@gmail.com

 Name of script: GOBItoIII-OCLC
 Script set: Order (order)

 Script Function:
	This script takes the information copied from the GOBI record and searches
	Amazon via API. Depending on if there is a match, an IE window to the
	Amazon product may or may not open. Regardless of the Amazon search
	outcome, the script searches Millennium by title (automatically) and by
	other information. If the item is not in III, then the script searches
	OCLC Connexion by ISBN. Otherwise, if the item is in III, it will prompt
	a check to determine if it is an added copy/volume order.

	##### NOTE ##### The user must highlight and copy the entire GOBI record
	before starting the script. Copy everything, including the yellow box
	area for notes.

 Programs used: GOBI 3
					(Script tested in Firefox 3+ and IE 7+)
				Millennium Acquisitions Module JRE v 1.6.0_02
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
#include <IE.au3>
#include <date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
AutoItSetOption("WinTitleMatchMode", 4)
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $GOBI_INFO, $GOBI_INFO_PREP, $GOBI_INFO_ARRAY_MASTER
Dim $WIN_TITLE
Dim $TITLE
Dim $ISBN
Dim $SERIES
Dim $decide
Dim $win
Dim $BROWSER
Dim $ISBN_N
Dim $AMAZON_A
Dim $FUND, $FUND_N, $FUND_S, $multi
Dim $DISCOUNT, $COST, $COST_L
Dim $REF_A, $REF_B, $REF
Dim $AV_A, $AV
Dim $HOLD_A, $HOLD
Dim $AC_A, $AC
Dim $MISSING_A, $MISSING
Dim $PROFILE
Dim $DATE_P, $PROFILE_A, $DATE_D, $20, $16, $COST_D
Dim $D
Dim $FUND_C, $UK
Dim $CALL_N, $DIFF_ED
Dim $oIE, $GUIActiveX
Dim $msg
Dim $text
Dim $APRICE_A, $APRICE, $DIFF, $MARKETPLACE, $close, $AMAZON_SEARCH, $AMAZON_L
Dim $oShell, $oHTTP, $HTMLSource

;################################ MAIN ROUTINE #################################
_ClearBuffer()

;~ ;ie work around part 1
;~ $WIN_TITLE = WinGetTitle("GOBI 3")
;~ $BROWSER = StringInStr($WIN_TITLE, "Mozilla Firefox")


;~ WinActivate("Untitled")
;~ 	Local $oHTTP = open.
;~ 	ConsoleWrite($oHTTP)

;~ $oHTTP = ObjCreate("winhttp.winhttprequest")
;~ $oHTTP.Open("Get", "https://www.rstudio.com/products/rstudio/")
;~ $oHTTP.Send()
;~ $HTMLSource = $oHTTP.Responsetext


;~ $ISBN = "9780838914144"
;~ $ISBN = InputBox("Barcode", "Please scan barcode")
;~ Local $oIE =_IECreate("https://www.amazon.com/gp/search/ref=sr_adv_b/?search-alias=stripbooks&unfiltered=1&field-keywords=&field-author=&field-title=&field-isbn=" & $ISBN)


;~ Local $oIE = _IECreate("www.autoitscript.com")
;~ Sleep(5000)
;~ _IENavigate($oIE, "http://www.autoitscript.com/forum/index.php?")
;~ Sleep(5000)
;~ _IENavigate($oIE, "http://www.autoitscript.com/forum/index.php?showforum=9")
;~ _IEQuit($oIE)



;~ $oShell = ObjCreate("shell.application")
;~ $oShell.MinimizeAll


;Get user initials based on username to put into 947 field
Local $var = EnvGet("TEMP")
local $C_INI = _LoadVar("$C_INI")

$var = StringTrimLeft($var, 9)
$var = StringTrimRight($var, 19)

Switch $var
	Case "alexanpk"
		$C_INI = "pk"
	Case "barbouh2"
		$C_INI = "hlb"
	Case "patricm"
		$C_INI = "mp"
	Case "spencert"
		$C_INI = "rts"
	Case "keyessl"
		$C_INI = "sk"
	Case "stepanm"
		$C_INI = "ms"
    Case "bazelejw"
		$C_INI = "jwb"
    Case "abneymd"
		$C_INI = "ma"
    Case "smithjl9"
	    $C_INI = "js"
    Case "cliftks"
	    $C_INI = "kc"
	Case "bomanca"
		$C_INI = "cb"
	case Else
		$C_INI = "999"
EndSwitch



ShellExecute("notepad","","")
sleep(100)
WinActivate("Untitled")
send($C_INI)
sleep(1000)
send("{enter}")
send("This is a test.")
Sleep(1000)
send("!{F4}")
send("{RIGHT}{ENTER}")




_ClearBuffer()


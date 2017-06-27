#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Fileversion=4.0.1.12
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: Create Item - Non DLC
 Script set: Receipt Cataloging (Receipt)

 Script Function:
	This script reates a new item record, inputs the appropriate codes in
	the item record, and prompts the user to send the item to the appropriate
	person/department.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(bibliographic/order record open)

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
AutoItSetOption("WinTitleMatchMode", 2)
Opt("WinSearchChildren", 1)
AutoItSetOption("MustDeclareVars", 1)
TraySetIcon(@DesktopDir & "\SierraReceipt - gmd\Images\red N.ico")

;######### DECLARE VARIABLES #########
Dim $COPIES
Dim $BIB_LOC
Dim $FUND_M, $FUND
Dim $ITEM_STAT
Dim $STATUS
Dim $LOCATION
Dim $ICODE1_1
Dim $ITYPE
Dim $LABELLOC
Dim $LOCAL
Dim $VOL
Dim $decide
Dim $BARCODE
Dim $T25
Dim $ITEM_REC, $ITEM_REC_PREP, $ITEM_ARRAY_MASTER
Dim $ICODE1_A, $ICODE1
Dim $IS_DOC
Dim $ADD_ITEM
Dim $300_E_E, $300_E
Dim $T85, $T86, $T87, $T88, $T89
Dim $IS_MUS
Dim $BIB_LOC_1
Dim $EBOOK
Dim $reserve
Dim $dean
Dim $approval
Dim $notification
Dim $firm
Dim $FORM

;################################ MAIN ROUTINE #################################
;start loading variables from order/bib record script
$FUND = _LoadVar("$FUND")
$T85 = _LoadVar("$T85")
$T86 = _LoadVar("$T86")
$T87 = _LoadVar("$T87")
$T88 = _LoadVar("$T88")
$T89 = _LoadVar("$T89")
$BARCODE = _LoadVar("$BARCODE")
$BIB_LOC = _LoadVar("$BIB_LOC")
$COPIES = _LoadVar("$COPIES")
$ICODE1_1 = _LoadVar("$ICODE1_1")
$LABELLOC = _LoadVar("$LABELLOC")
$LOCATION = _LoadVar("$LOCATION")
$ITYPE = _LoadVar("$ITYPE")
$VOL = _LoadVar("$VOL")
$300_E = _LoadVar("$300_E")
$300_E_E = Eval($300_E)
$LOCAL = _LoadVar("$LOCAL")
$BIB_LOC_1 = _LoadVar("$BIB_LOC_1")
$EBOOK = _LoadVar("$EBOOK")
$reserve = _LoadVar("$reserve")
$dean = _LoadVar("$dean")
$approval = _LoadVar("$approval")
$firm = _LoadVar("$firm")
$notification = _LoadVar("$notification")
$FORM = _LoadVar("$FORM")

;msgbox(0, "icode", $ICODE1_1)
;focus on bib/order record screen
;If WinExists("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;	WinWaitActive ("[REGEXPTITLE:\A[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the record.")
;	Exit
;EndIf

If WinExists("[REGEXPTITLE:[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[bo][0-9ax]{7,8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf

;Sierra workaround / focus on bib record screen
;If WinExists("[TITLE:Sierra; CLASS:SunAwtFrame]") Then
;	WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
;	WinWaitActive ("[TITLE:Sierra; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Sierra record", "Please open the bib record.")
;	Exit
;EndIf


Sleep(1000)

;prep fund variable for processing
$FUND = StringLeft($FUND, 5)
$FUND = StringStripWS($FUND, 8)
If $FUND <> "multi" Then
	$FUND_M = StringLeft($FUND, 1)
EndIf
Sleep(0200)

$STATUS = "i" ;default status
If $FORM = "w" Then
	$ITYPE = 12
EndIf
;MsgBox(0, "itype", $ITYPE)


Select
	Case $FUND = "4a" ;textbook reserve fund
		$LOCATION = $LOCATION
		$ITYPE = "56"
		$LABELLOC = "King Textbook Reserve"
		$STATUS = "l"
	Case $FUND = "multi" ;default to ccat
		$LOCATION = "ccat"
	Case $FUND_M = "m" ;middletown ccat
		$LOCATION = "mcat"
	Case $LOCATION = "none" ;added vol/copy item location
		$LOCATION = "none"
	Case $LOCATION = "game1" ;game lab location
		$LOCATION = "ccat"
		$ITYPE = "21"
		$LABELLOC = "GameLab"
		$ICODE1_1 = "88"
	Case Else
		$LOCATION = "ccat" ;default location
EndSelect

;switch to new item record
;If WinExists("[REGEXPTITLE:\A[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")Then
;	Sleep(1000)
;	_SendEx("!g")
;	Sleep(0600)
;	_SendEx("o")
;	Sleep(0600)
;	_SendEx("a")
;Else
Sleep(1000)
_SendEx("!g")
Sleep(0600)
_SendEx("o")
Sleep(0600)
_SendEx("a")
;EndIf
;WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;WinWaitActive ("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
Sleep(0800)
MouseClick("left", 667, 200, 2)
Sleep(0500)
_SendEx("!w")
Sleep(0600)
_SendEx("i")
Sleep(0600)
_SendEx("{ENTER}")
Sleep(0300)
If $approval <> 1 Then
	_SendEx("!n")
	Sleep(0800)
Else
   _SendEx("!l")
   Sleep(0800)
EndIf


;IF WinExists("[TITLE:Attach New Records?; CLASS:SunAwtDialog]") Then
;	_SendEx("y")
;EndIf

;wait for item record focus
;WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]", "")
;MouseClick("left", 667, 519, 1)
;_SendEx("^{HOME}")

;T25 variable affects status.... look at bib- check Iloc File
;If $LOCAL > -1 Then
;	$ITYPE = 26
;endif

;start item record data entry
Sleep(0600)
If $COPIES > 1 Then
	$decide = InputBox("Copies", "Enter copy designation (e.g., '1' OR '2')", "")
	_SendEx("{DEL 2}")
	Sleep(0100)
	_SendEx($decide)
	Sleep(0100)
	_SendEx("{TAB}")
	Sleep(0300)
Else
	_SendEx("{TAB}")
	Sleep(0600)
EndIf
_SendEx($LOCATION)
Sleep(0300)
_SendEx("{TAB 2}")
Sleep(0300)
_SendEx("{DEL}")
Sleep(0300)
_SendEx($ICODE1_1)
Sleep(0300)
_SendEx("{TAB 3}")
Sleep(0300)
_SendEx($STATUS)
Sleep(0300)
_SendEx("{TAB 2}")
Sleep(0400)
_SendEx($ITYPE)
Sleep(0300)
_SendEx("{END}")
Sleep(0100)
#cs
If $LABELLOC <> "" Then
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0100)
	_SendEx($LABELLOC)
EndIf
#ce

Sleep(0400)
If $VOL > 1 Then
	If $LOCATION = "murc" Then
		$decide = "AudioD"
	Else
		$decide = InputBox("Volume label", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
	EndIf
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("v{TAB}")
	Sleep(0100)
	_SendEx($decide)
EndIf
SLEEP(0400)
If $BARCODE = "N" Then
	_SendEx("^s")
Else
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("b{TAB}")
	Sleep(0400)
	$decide = InputBox("Barcode", "Scan in Barcode", "")
	Sleep(0100)
	_SendEx($decide)
	Sleep(0100)
	_SendEx("^s")
EndIf

SLEEP(0400)


;add note for dean's material
If $dean = 1 Then
	Sleep(0100)
	_SendEx("^{END}")
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("x{TAB}")
	Sleep(0100)
	_SendEx("Catalog for Dean's Office")
	Sleep(0100)
	_SendEx("^s")
EndIf

;determine if accompanying material note needs to be added
;If $300_E = 0 or $300_E_E = "" Then
;		sleep(0100)
;Else
$IS_MUS = StringInStr($BIB_LOC_1, "mu")
If $300_E <> 0 AND $IS_MUS = 0 Then
			$300_E = StringUpper($300_E)
			;_SendEx("^{HOME}")
			Sleep(0100)
			;_SendEx("{TAB}")
			Sleep(0100)
			;_SendEx("{DOWN 5}")
			_SendEx("{UP 3}")
			Sleep(0100)
			_SendEx("{RIGHT 2}")
			Sleep(0100)
			_SendEx("p")
			Sleep(0100)
			_SendEx("^{END}")
			Sleep(0100)
			_SendEx("{ENTER}")
			Sleep(0100)
			_SendEx("m{TAB}")
			Sleep(0100)
			_SendEx($300_E & "IN POCKET")
			Sleep(0100)
			_SendEx("^s")
	ElseIf $300_E <> 0 AND $IS_MUS = 1 Then
		$ADD_ITEM = 1
	Else
		sleep(0200)
	;EndIf
EndIf

;CREATE new item record for accompanying material
If $ADD_ITEM = 1 Then
	Sleep(0100)
	_SendEx("!g")
	Sleep(0200)
	_SendEx("o")
	Sleep(0200)
	_SendEx("a")
	;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")
	Sleep(0200)
	_SendEx("!g")
	Sleep(0200)
	_SendEx("o")
	Sleep(0200)
	_SendEx("i")
	Sleep(0200)
	_SendEx("!n")
	Sleep(0100)
	WinWaitActive("[TITLE:New ITEM; CLASS:SunAwtFrame]")
	Sleep(0100)
	$decide = MsgBox(4, "","Is the item a non-music CD?")
	Switch $decide
		Case 6
			$LOCATION = "muc"
			$VOL = "disc"
			$ITYPE = 18
			$ITEM_STAT = "r"
			$LABELLOC = "Music Circ"
		Case 7
			$LOCATION = "murc"
			$VOL = "AudioD"
			$ITYPE = 1
			$ITEM_STAT = "r"
			$LABELLOC = "Music"
	EndSwitch
	_SendEx("^{HOME}")
	Sleep(0100)
	_SendEx("{TAB}")
	Sleep(0100)
	_SendEx($LOCATION)
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx("{DEL}")
	Sleep(0100)
	_SendEx($ICODE1_1)
	_SendEx("1")
	Sleep(0100)
	_SendEx("{TAB 3}")
	Sleep(0100)
	_SendEx($ITEM_STAT)
	Sleep(0100)
	_SendEx("{TAB 2}")
	Sleep(0100)
	_SendEx($ITYPE)
	Sleep(0100)
	_SendEx("^{END}")
	Sleep(0100)
	#cs
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("l{TAB}")
	Sleep(0100)
	_SendEx($LABELLOC)
	Sleep(0100)
	#ce
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("v{TAB}")
	Sleep(0100)
	_SendEx($VOL)
	_SendEx("{ENTER}")
	Sleep(0100)
	_SendEx("b{TAB}")
	Sleep(0200)
	$decide = InputBox("Barcode", "Scan in Barcode", "")
	Sleep(0100)
	_SendEx($decide)
	Sleep(0100)
	_SendEx("^s")
EndIf

; Start preparation to bump
Select
	Case $T25 <> 0
		MsgBox(0, "Reserves", "RESERVE ITEM -- GIVE TO PROCESSING.")
	Case $LOCATION = "mcat"
		MsgBox(0, "MIDDLETOWN RECORD", "Unless this is a DVD, send to Carol with pink slip in book.")
	Case $FUND = "4a"
		MsgBox(0, "RESERVES TEXTBOOKS", "PLACE PINK RUSH RESERVES TEXTBOOK FLAG IN BOOK." & @CR & "Give book to Linda for Processing.")
	Case $approval = 1 Or $notification = 1
		MsgBox(0, "Copy Cataloging", "Send to copy cataloging." & @CR & $T85 & @CR & $T86 & @CR & $T87 & @CR & $T88 & @CR & $T89)
	Case $reserve = 1
		MsgBox(0, "Reserves", "Flag for reserves. Insert red flag for rush, insert blue flag for King reserve.")
	Case $dean = 1
		MsgBox(0, "Dean's Materials", "Insert red flag for rush. Send to copy cataloging.")
	Case $LABELLOC = "GameLab"
		MsgBox(0, "Video Game", "Give to Catalog Librarian.")
	;Case Else
		;MsgBox(0, "Copy Cataloging", "If not an IMC Juv, place blue copy in book." & @CR & $T85 & @CR & $T86 & @CR & $T87 & @CR & $T88 & @CR & $T89)
EndSelect

If $EBOOK = 1 Then
	MsgBox(64, "Paperwork to Jennifer", "Please send the accompanying paperwork to Jennifer for URL connection & cataloging.")
Endif

_ClearBuffer()
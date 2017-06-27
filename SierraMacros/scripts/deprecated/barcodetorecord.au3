#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
	#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: barcodetorecord
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script starts the copy cataloging process. It takes the local
	library barcode and searches Millennium. If the item record opens,
	the script copies the item record data and then opens the bib record.
	The script copies the bib data and will	search OCLC for the item if necesary.

 Programs used: Millennium Cataloging Module JRE v 1.6.0_02
					(Main catalog search screen)
					Record view properties - Summary retrieval view, item
					summary view
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
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("MustDeclareVars", 1)
TraySetIcon(@WorkingDir & "\Images\take1new.ico")


;######### DECLARE VARIABLES #########
Dim $winsize
dim $barcode
Dim $UPD
Dim $ITEM_REC_PREP, $ITEM_REC, $ITEM_ARRAY_MASTER
Dim $row2_MASTER
Dim $ICODE1
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $300
Dim $OCLC_NUM
Dim $VOL
dim $LAB_LOC
Dim $ACCOMP
Dim $dean
Dim $row3_MASTER, $BCODE2, $BIB_LOC_1
Dim $LEADER, $LEADER_3
Dim $TITLE, $245_A2, $245DASH
Dim $SF_NAME, $SF_590, $SF_7XX
Dim $fundrow_MASTER, $ORD_REC, $ORD_REC_PREP, $ORD_ARRAY_MASTER
Dim $FUND, $CLAIM, $STATUS, $FUND2, $RLOC, $EPRICE, $LOCATION, $REF
DIM $COP, $ILOVC, $LPAT, $row1_MASTER
Dim $RDATE, $CODE4, $ACQTYPE, $acqtypecheck, $BLOC, $FORM, $CDATE, $TLOC, $ODATE, $COPIES
Dim $VENDOR, $ORDNOTE, $CODE1, $row4_MASTER, $row5_MASTER, $row6_MASTER, $row7_MASTER, $row8_MASTER
Dim $ORDTYPE, $ORD_LANG, $CODE2, $RACTION, $CODE3, $OLOCATION
Dim $clipcheck

;################################ MAIN ROUTINE #################################
;focus on Millennium main catalog screen
;If WinExists ("[REGEXPTITLE:\A[bio][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
;	WinClose("[REGEXPTITLE:\A[bio][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;EndIf
;/Emily/
;If WinExists ("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [bio][0-9ax]{7,8}; CLASS:SunAwtFrame]" Then
;	WinClose("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [bio][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;EndIf

_ClearBuffer()
Sleep(0300)
ClipPut("")
$ITEM_REC = ClipGet()
;If WinExists("[TITLE:Millennium Cataloging; CLASS:SunAwtFrame]") Then
;	WinActivate("[TITLE:Millennium Cataloging; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "nope", "Please log into Millennium.")
;	Exit
;EndIf

;WinWaitActive("[TITLE:Millennium Cataloging; CLASSS:SunAwtFrame]")

; if in multi window mode, make sure to close exising item or bib record
If WinExists("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
WinActivate("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")

$winsize = WinGetPos("[REGEXPTITLE:[bi][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
_SendEx("^s")
Sleep(0200)
_SendEx("!q")
sleep(0300)
EndIf

Local $DUP_CHECK = WinExists("[TITLE:Perform duplicate checking?; CLASS:SunAwtDialog]")
If $DUP_CHECK = 1 Then
	_SendEx("n")
	;WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	;WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	;/Emily/
	WinActivate("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[b][0-9ax]{8}; CLASS:SunAwtFrame]")
EndIf

#cs;scan in library barcode, searches catalog by barcode
$barcode = InputBox("Barcode", "Scan the item's barcode", "")
If $barcode = "" Then
	Exit
Else
_IIIsearch("b", $barcode)
EndIf
sleep(1200)
#ce

;wait for item record to appear/focus, left mouse click on data field
;WinWaitActive("[REGEXPTITLE:\A[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;$winsize = WinGetPos("[REGEXPTITLE:\A[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")


#cs
WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
$winsize = WinGetPos("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
#ce



;select all and copy item record information, parce data into array
While $ITEM_REC = "" ;this prevents the macro from moving on too quickly. basically, it can't move on until it's copied the item
	If WinExists("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]") Then
WinActivate("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")

$winsize = WinGetPos("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
_WinClick($winsize)
EndIf
_DataCopy()
$ITEM_REC = ClipGet()
$ITEM_REC_PREP = StringRegExpReplace ($ITEM_REC, "[\r\n]+", "fnord")
$ITEM_ARRAY_MASTER = StringSplit($ITEM_REC_PREP, "fnord", 1)

;grab item location in case it's from standing orders or doesn't have an order record for some reason
$row1_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 1)
$LOCATION = _ArrayPop($row2_MASTER)
$LPAT = _ArrayPop($row2_MASTER)
$COP = _ArrayPop($row2_MASTER)
$LOCATION = StringTrimLeft($LOCATION, 8)
$LOCATION = StringLeft($LOCATION, 5)
$LOCATION = StringStripWS($LOCATION, 8)

;search for ICODE - determines UPD in 947
$row2_MASTER = _IIIfixed($ITEM_ARRAY_MASTER, 2)
$ICODE1 = _arrayItemString ($row2_MASTER, "CODE1")
$ICODE1 = StringTrimLeft($ICODE1, 6)
$ICODE1 = StringStripWS($ICODE1, 8)

If $ICODE1 = 2 OR $ICODE1 = 83 Then
	$UPD = 0 ;no upd code in 947
Else
	$UPD = 1 ;upd code in 947
EndIf

;determine if additional item records need to be created
$VOL = _ArraySearch($ITEM_ARRAY_MASTER, "VOLUME", 0, 0, 0, 1)
If $VOL > -1 Then
	$VOL = 1
EndIf

;determine if item is dean's office item
$dean = _ArraySearch($ITEM_ARRAY_MASTER, "Catalog for Dean's Office", 0, 0, 0, 1)
If $dean > -1 Then
	$dean = 1
EndIf

#cs -- deprecated as printer no longer needs this field
;determine if label location is present
$LAB_LOC = _ArraySearch($ITEM_ARRAY_MASTER, "LABEL LOC", 0, 0, 0, 1)
If $LAB_LOC > -1 Then
	$LAB_LOC = 1 ;label location not present, will determine later on if record needs one
Else
	$LAB_LOC = 0
EndIf
#ce

;determine if item has accompying item note, will determine later on if record needs one
$ACCOMP = _ArraySearch($ITEM_ARRAY_MASTER, "IN POCKET", 0, 0, 0, 1)
_StoreVar("$ACCOMP")
If $ITEM_REC <> "" Then ExitLoop
WEnd

Sleep(0200)

WinWaitActive("[REGEXPTITLE:[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
;check order record for fund codes! whee!!! check all the things!!
Sleep(0100)
_SendEx("!g")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("o")
Sleep(0200)
_SendEx("!l")
Sleep(0600)

WinWaitActive("[REGEXPTITLE:[o][0-9ax]{7,8}; CLASS:SunAwtFrame]")

;select all and copy order record information, parse data into array
_SendEx("^{HOME}")
Sleep(0200)
_SendEx("^a")
Sleep(0100)
_SendEx("^a")
Sleep(0100)
_SendEx("^c")
Sleep(0300)

$ORD_REC = ClipGet()
$ORD_REC_PREP = StringRegExpReplace($ORD_REC, "[\r\n]+", "fnord")
$ORD_ARRAY_MASTER = StringSplit($ORD_REC_PREP, "fnord", 1)

If $ACCOMP = -1 Then

$ACCOMP = _ArraySearch($ORD_ARRAY_MASTER, "Rec'd", 0, 0, 0, 1)

_StoreVar("$ACCOMP")
EndIf


;~ ##### start fixed data setting #####
;row 1 includes acqtype/code4/rdate
$row1_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 1)
$RDATE = _ArrayPop($row1_MASTER)
$CODE4 = _ArrayPop($row1_MASTER)
$ACQTYPE = _ArrayPop($row1_MASTER)
$RDATE = StringTrimLeft($RDATE, 5)
$CODE4 = StringTrimLeft($CODE4, 5)

;~ Bloody acqtype workaround - sometimes the A in Acq type doesn't get captured
$acqtypecheck = StringInStr($ACQTYPE, "ACQ TYPE")
If $acqtypecheck <> 0 Then
	$ACQTYPE = StringTrimLeft($ACQTYPE, 9)
EndIf

;row 2 includes location, eprice, rloc
$row2_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 2)
$RLOC = _ArrayPop($row2_MASTER)
$EPRICE = _ArrayPop($row2_MASTER)
$LOCATION = _ArrayPop($row2_MASTER)
$RLOC = StringTrimLeft($RLOC, 4)
$EPRICE = StringTrimLeft($EPRICE, 7)
$LOCATION = StringTrimLeft($LOCATION, 9)
$LOCATION = StringLeft($LOCATION, 5)
$LOCATION = StringStripWS($LOCATION, 8)
_StoreVar("$LOCATION")
$OLOCATION = $LOCATION
$OLOCATION = StringReplace($OLOCATION, "li", "l")
_StoreVar("$OLOCATION")

;MsgBox(0, "variables", $RLOC & $EPRICE & $LOCATION)

;row 3 includes cdate, form, bloc
$row3_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 3)
$BLOC = _ArrayPop($row3_MASTER)
$FORM = _ArrayPop($row3_MASTER)
$CDATE = _ArrayPop($row3_MASTER)
$BLOC = StringTrimLeft($BLOC, 4)
$FORM = StringTrimLeft($FORM, 4)
$CDATE = StringTrimLeft($CDATE, 6)

;row 4 includes claim, fund, status
$row4_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 4)
$STATUS = _ArrayPop($row4_MASTER)
$FUND = _ArrayPop($row4_MASTER)
$CLAIM = _ArrayPop($row4_MASTER)
$STATUS = StringTrimLeft($STATUS, 6)
$FUND = StringTrimLeft($FUND, 4)
$CLAIM = StringTrimLeft($CLAIM, 6)

$FUND2 = StringLeft($FUND, 3); this is for checking the fund code against the special fund list
_StoreVar("$FUND2")

;row 5 includes copies, odate, tloc
$row5_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 5)
$TLOC = _ArrayPop($row5_MASTER)
$ODATE = _ArrayPop($row5_MASTER)
$COPIES = _ArrayPop($row5_MASTER)
$TLOC = StringTrimLeft($TLOC, 4)
$ODATE = StringTrimLeft($ODATE, 5)
$COPIES = StringTrimLeft($COPIES, 7)

;row 6 includes code1, ordnote, vendor
$row6_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 6)
$VENDOR = _ArrayPop($row6_MASTER)
$ORDNOTE = _ArrayPop($row6_MASTER)
$CODE1 = _ArrayPop($row6_MASTER)
$VENDOR = StringTrimLeft($VENDOR, 6)
$ORDNOTE = StringTrimLeft($ORDNOTE, 8)
$CODE1 = StringTrimLeft($CODE1, 6)

;row 7 includes code2, ordtype, lang
$row7_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 7)
$ORD_LANG = _ArrayPop($row7_MASTER)
$ORDTYPE = _ArrayPop($row7_MASTER)
$CODE2 = _ArrayPop($row7_MASTER)
$ORD_LANG = StringTrimLeft($ORD_LANG, 4)
$ORDTYPE = StringTrimLeft($ORDTYPE, 8)
$CODE2 = StringTrimLeft($CODE2, 6)

;row 8 includes code3, raction, volumes
$row8_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 8)
$VOL = _ArrayPop($row8_MASTER)
$RACTION = _ArrayPop($row8_MASTER)
$CODE3 = _ArrayPop($row8_MASTER)
$VOL = StringTrimLeft($VOL, 7)
$RACTION = StringTrimLeft($RACTION, 7)
$CODE3 = StringTrimLeft($CODE3, 6)
;~ ##### end fixed data setting #####

$ACCOMP

;msgbox(0, "loc", $LOCATION)

If $LOCATION = "kngr" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATION = "aar" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATION = "scr" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATION = "mur" Then
	$REF = 1
	_StoreVar("$REF")
ElseIf $LOCATIOn = "imr" Then
	$REF = 1
	_StoreVar("$REF")
Else
	$REF = 0
	_StoreVar("$REF")
EndIf

#cs
;row 4 includes claim, fund, status
$fundrow_MASTER = _IIIfixed($ORD_ARRAY_MASTER, 4)
$STATUS = _ArrayPop($fundrow_MASTER)
$FUND = _ArrayPop($fundrow_MASTER)
$CLAIM = _ArrayPop($fundrow_MASTER)
$STATUS = StringTrimLeft($STATUS, 6)
$FUND = StringTrimLeft($FUND, 4)
$CLAIM = StringTrimLeft($CLAIM, 6)

$FUND2 = StringLeft($FUND, 3)
_StoreVar("$FUND2")

;MsgBox(0, "fund", $FUND2)
#ce

MsgBox(0, "Order notes", "Please make a mental note of any relevant order record notes!")


Run(@DesktopDir & "\SierraCopyCat\Special Fund List.exe")
While ProcessExists("Special Fund List.exe")
	Sleep(0400)
WEnd
$SF_NAME = _LoadVar("$SF_NAME")
$SF_590 = _LoadVar("$SF_590")
$SF_7XX = _LoadVar("$SF_7XX")


;going to bib record
Sleep(0100)
_SendEx("^+e")

;wait for bib record focus
;WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{7,8}; CLASS:SunAwtFrame]")

WinWaitActive("[REGEXPTITLE:[b][0-9ax]{7,8}; CLASS:SunAwtFrame]")

If $ICODE1 = "" Then
	MsgBox(0, "Uh-oh!", "It seems Sierra flaked out on this run. Please close the bib record and start the macro again before proceeding.")
	Exit
EndIf


;select all and copy bib record information, parse data into array
_DataCopy()
_SendEx("^{HOME}") ;de-highlights record. common courtesy for the catalogers
$BIB_REC = ClipGet()
$BIB_REC_PREP = StringRegExpReplace ($BIB_REC, "[\r\n]+", "Fnord")
$BIB_ARRAY_MASTER = StringSplit($BIB_REC_PREP, "Fnord", 1)
_StoreVar("$BIB_ARRAY_MASTER")


;row 3 includes bib location and bcode2
$row3_MASTER = _IIIfixed($BIB_ARRAY_MASTER, 3)
$BCODE2 = _ArrayPop($row3_MASTER)
$BIB_LOC_1= _ArrayPop($row3_MASTER)
$BIB_LOC_1 = StringTrimLeft($BIB_LOC_1, 9)
$BIB_LOC_1 = StringLeft($BIB_LOC_1, 5)
$BIB_LOC_1 = StringStripWS($BIB_LOC_1, 8)

;get encoding level from leader, if level 3 open OCLC record
$LEADER = _arrayItemString($BIB_ARRAY_MASTER, "MARC Leader	")
$LEADER =  StringTrimLeft($LEADER, 12)
$LEADER = StringStripWS($LEADER, 8)
$LEADER_3 = StringMid($LEADER, 16, 1)

;300 search - if incomplete open OCLC record
$300 = _arrayItemString($BIB_ARRAY_MASTER, "DESCRIPT.	300	 	 	")
$300 =  StringTrimLeft($300, 18)
$300 = StringStripWS($300, 8)

;(300 is incomplete OR item is Middletown) AND encoding level != 3 - search OCLC by OCLC #
If ($300 = "p.|ccm" Or $300 = "p.;|ccm" Or $300 = -1  or $300 = "p.cm" OR $ICODE1 = 83) AND $LEADER_3 <> 3 Then
	$OCLC_NUM = _arrayItemString($BIB_ARRAY_MASTER, "OCLC #	001	 	 	")
	$OCLC_NUM = StringTrimLeft($OCLC_NUM, 15)
	_OCLCSearch("{#}", $OCLC_NUM)
EndIf

;Item is STO OR encoding level = 3 - search OCLC by title (subfield a)
If $ICODE1 = 3 OR $LEADER_3 = 3 Then
	$TITLE = _arrayItemString($BIB_ARRAY_MASTER, "TITLE	")
	If $LEADER_3 = 3 Then
		$TITLE = StringTrimLeft($TITLE, 6)
		$245_A2 = StringSplit($TITLE, "|")
		$TITLE = _arrayItemString($245_A2, "245	")
		$TITLE = StringTrimLeft($TITLE, 8)
		$245DASH = StringRegExp($TITLE, "[:/]", 0)
		If $245DASH = 1 Then
			$TITLE = StringTrimRight($TITLE, 2)
		EndIf
	Else
		$TITLE = StringTrimLeft($TITLE, 6)
	EndIf
	_OCLCSearch("ti:", $TITLE)
EndIf

;store variables for ItemRec script
_StoreVar("$UPD")
_StoreVar("$LAB_LOC")
_StoreVar("$VOL")
_StoreVar("$ICODE1")
_StoreVar("$ACCOMP")
_StoreVar("$BIB_LOC_1")
_StoreVar("$dean")
_StoreVar("$SF_NAME")
_StoreVar("$SF_590")
_StoreVar("$SF_7XX")

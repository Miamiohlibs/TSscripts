#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Authors:Craig Boman, Discovery Services Libarian, Miami University
			bomanca@miamioh.edu OR craig.boman@gmail.com
		 Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: ItemRec
 Script set: Copy Cataloging (CopyCat)

 Script Function:
	This script completes the copy cataloging process by entering the appropriate
	information in the bib record as well as perform quality control on select
	fields. The script then opens the item record and enters the appropriate
	codes and messages, creating additional item records if necesary.

 Programs used: Sierra Cataloging Module JRE v 1.8.0_45
					(bib record open)
					Record view properties - Summary retrieval view, item
					summary view

 Last revised: 8/29/17
			   Updated #Include to include TS custom function library
			   Updated window controls for JRE version

 PLEASE NOTE - This script uses a custom UDF library called TSCustomFunction.
			   The script will not run properly (if launched from .au3 file)
			   if that file is not included in the include folder in the
			   AutoIt directory.

 Copyright (C): 2017 by Miami University Libraries.  Libraries
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
AutoItSetOption("SendKeyDelay", 30)
TraySetIcon(@DesktopDir & "\SierraCopyCat\Images\green D.ico")

;######### DECLARE VARIABLES #########
Dim $BIB_REC, $BIB_REC_PREP, $BIB_ARRAY_MASTER
Dim $UPD
Dim $ICODE1
Dim $row1_MASTER, $row2_MASTER, $row3_MASTER
Dim $BCODE3
dim $CAT_DATE
Dim $LANG
Dim $BCODE1
Dim $COUNTRY
Dim $SKIP
dim $BCODE2
Dim $BIB_LOC_1
dim $CAT_DATE_R
Dim $TODAY_DATE
dim $decide
Dim $BIB_LOC
Dim $LANG_CHECK
Dim $008_L_A
dim $008_L
Dim $COUNTRY_CHECK
Dim $008_C_A
Dim $008_C
Dim $SKIP_CHECK
Dim $LOCATION
Dim $ICODE1_1
Dim $ITYPE
Dim $LABELLOC
Dim $VOL, $vol_300, $vol_300_multi
Dim $BARCODE
Dim $DUP_CHECK
Dim $LAB_LOC
Dim $ACCOMP
Dim $C_INI
dim $dean
Dim $300_A, $300, $300_C_A, $300_C_A_1, $300_C, $300_C_X, $FOLIO, $300_E, $300_E_A, $300_E_1
Dim $260_A, $008_D_A, $008_D, $050_D, $050_A, $050, $260, $260_L, $260_C1, $008, $008_CON, $260_C, $090_A
Dim $245_2_A, $245_2
Dim $MULTI_LOC_A, $MULTI_LOC, $MULTI_REF, $kngr, $aar, $scr, $docr, $mdir, $doc, $imr
Dim $shelfready
Dim $IS_IMC
Dim $scfo, $scov
Dim $KNGFO
Dim $260, $260_C
Dim $264, $264_C
Dim $050, $050_B
Dim $008, $008_date
Dim $digCheck
Dim $conf
Dim $300error, $300errors, $300ok
Dim $donor_loc
Dim $SF_590, $SF_7XX, $SF_NAME, $SF_FIB
Dim $AAFO
Dim $don_name
Dim $264_C_fix, $264_fix, $260_fix
Dim $440_A, $490_A, $800_A, $810_A, $811_A, $830_A, $series
Dim $856, $956
Dim $subject
Dim $REF, $REF_1
Dim $OLOCATION
Global $050_array[2]
Dim $050_loc
Dim $049, $MIAS, $MIAA, $MIAH, $MIBN

$REF = _LoadVar("$REF")
$SF_NAME = _LoadVar("$SF_NAME")
$SF_590 = _LoadVar("$SF_590")
$SF_7XX = _LoadVar("$SF_7XX")
$SF_FIB = _LoadVar("$SF_FIB")
$OLOCATION = _LoadVar("$OLOCATION")
$ACCOMP = _LoadVar("$ACCOMP")

;REMOVE IN PRODUCTION - test variables
$ICODE1 = "14"
$LOCATION = "kngli"
$vol = 1
$dean = "1"
$VOL = "1"
$ACCOMP = -1
$300_E = "atlas"
$BARCODE = "N"
$ITYPE = "26"
;REMOVE IN PRODUCTION - test variables

;################################ MAIN ROUTINE #################################
_Initial()

;focus on Millennium bib record
;If WinExists("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
;	WinActivate("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;	WinWaitActive("[REGEXPTITLE:\A[b][0-9ax]{8}; CLASS:SunAwtFrame]")
;Else
;	MsgBox(64, "Millennium record", "Please open the bib record.")
;	Exit
;EndIf

;/Emily/
#cs
If WinExists("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:Sierra · Miami University Libraries · .* · [b][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64, "Sierra record", "Please open the order record.")
	Exit
EndIf
#ce


If WinExists("[REGEXPTITLE:[i][0-9ax]{8}; CLASS:SunAwtFrame]") Then
	WinActivate("[REGEXPTITLE:[i][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[i][0-9ax]{8}; CLASS:SunAwtFrame]")
Else
	MsgBox(64+65536, "Sierra record", "Please open the item record.")
	Exit
EndIf



Func _windowFocus()
	WinExists("[REGEXPTITLE:[i][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinActivate("[REGEXPTITLE:[i][0-9ax]{8}; CLASS:SunAwtFrame]")
	WinWaitActive("[REGEXPTITLE:[i][0-9ax]{8}; CLASS:SunAwtFrame]")
EndFunc


Func _itemEdits()

   ;~ ##### start fixed data setting #####

   ;wait for item record focus
   ;WinWaitActive("[REGEXPTITLE:\A[i][0-9ax]{7,8}; CLASS:SunAwtFrame]")
   ;/Emily/

   _windowFocus()

   ;start item record data entry
   Sleep(0100)
   _SendEx("^a" & "^{HOME}" & "{TAB}") ;probably don't need to select all here, recommend deletion - CB 8/28/17
   Sleep(0200)

	  ;start icode1 edits
   If $ICODE1 = 83 Then ;still not sure what's going on here.
	   _SendEx("{DEL 2}")
	   Sleep(0300)
	EndIf


   Sleep(0300)
   _SendEx($ICODE1 & "3" & "{TAB}") ;enters second (or third for middletown)
   Sleep(0300)

	  ;status update
   Sleep(0600)
   If $ICODE1 = 83 Then ;middletown
	   _SendEx("l")
   ElseIf $dean = 1 Then
	   _SendEx("k")

   EndIf




   ;$shelfready = InputBox("Shelf Ready Status", "Is this a shelf-ready item?" & @CR & "y - yes" & @CR & "n - no", "")

   ;Switch $shelfready
	   ;Case "n"
	   $decide = MsgBox(4+65536, "", "Is the item a paperback book?") ;we should not need to ask, add variable
	   If $decide = 6 Then
			   Sleep(0100)
			   _SendEx("{LEFT}r")
		   ElseIf $decide = 7 Then
			   ;Sleep(0100)
			   If $REF = 1 Then
				   _SendEx("{LEFT}o")
			   Else

			   _SendEx("-")
			   EndIf
		   EndIf

	   ;EndSwitch
	   ;Case "y"
		   ;Sleep(0100)
		   ;_SendEx("-")
	   ;EndSwitch

   ; shelf ready paperbacks and hardbacks get status - for available
   ; non-shelf ready paperbacks get "r" and hardbacks get "l"



	  ;start $ITYPE EDIT
   Sleep(0100)
   _SendEx("{TAB 4}" & $ITYPE)
   Sleep(0100)

   _SendEx("{TAB 3}")
   If $300_E <> "none" And $ACCOMP = -1 Then
	   $300_E = StringUpper($300_E)
	   _SendEx("p") ;update imessage field for check parts message
	   Sleep(0500)
   EndIf

	  ;start Location edit
   Sleep(0100)
   _SendEx("{TAB 4}" & $LOCATION)
   Sleep(0400)




   _SendEx("^{END}")
   Sleep(0300)
   ; Enter volume information if needed for item #1 after asking if necessary as sometimes volume info is already present in item record
   If $vol = 1 Then
	  $decide = MsgBox(4+65536, "Volume Designation", "Does item 1 need volume information?")
	  If $decide = 6 Then
		 $decide = InputBox("Volume label", "Enter volume designation (e.g., 'v.2' OR 'Suppl.')", "")
		 _SendEx("{ENTER}" & "v{TAB}")
		 Sleep(0100)
	     _SendEx($decide)
		 Sleep(0400)
	  EndIf
   EndIf



	  ;add note for delivered to dean's office
   Sleep(0400)
   If $dean = 1 Then
	   _SendEx("{ENTER}" & "x{TAB}")
	   Sleep(0500)
	   _SendEx("Delivered to Dean's Office on item created date." & $C_INI)
	EndIf



   ;save!
   _SendEx("^s")
   ;determine if the item needs an accompanying material item note
   If $300_E <> "none" And $ACCOMP = -1 Then
	   $300_E = StringUpper($300_E)
	   _SendEx("{ENTER}")
	   Sleep(0500)
	   _SendEx("m{TAB}" & $300_E & " IN POCKET")
	   Sleep(0500)
	   _SendEx("^s")
	EndIf

EndFunc



Func _newItem() ;creates new items
   Sleep(1000)
	_SendEx("!g" & "o" & "a" & "!w" & "i" & "{TAB}")
	Sleep(0200)
	_SendEx("!n")
	Sleep(0400)
    _SendEx("^s")
EndFunc

_itemEdits() ;runs function above



;determine if there needs to be more item records created
If $VOL = 1 Then
;looped - will continue loop until you answer no to the popup asking if you want to create another record
	Do
		$decide = MsgBox(4+65536, "More Volumes", "Do you need to attach another item record?")
		Switch $decide
			Case 6 ;Yes answer
				;create new item record
				;_windowFocus()
				Sleep(0100)
				_newItem() ; creates new item record

			   Sleep(0100)
				_itemEdits() ;should be able to call the function above in this loop
			   Sleep(0100)
			Case 7 ;No - ends loop
				$VOL = 0
		EndSwitch
	Until $VOL = 0
 EndIf



If $dean = 1 Then
	MsgBox(0+65536, "Dean's Office Materials", "Delete Catalog for Dean's office note.")
 EndIf



_ClearBuffer()
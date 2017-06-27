#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: Special Fund List
 Script set: Multiple

 Script Function:
	This script checks the call number against our local shelving practices, to make sure the call number
	and the order/bib locations all match before allowing it to proceed.  This is a QC script

 Programs used: n/a

Last revised: November 2014
				Updated away from switch case to if then as switch case was not work with variables this way. Now it works.

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
AutoItSetOption("MustDeclareVars", 1)

;######### DECLARE VARIABLES #########
Dim $050_loc, $050_loc_1
Dim $050_regexp[3]
Dim $050_letters
Dim $050_numbers
Dim $050_check


;################################ MAIN ROUTINE #################################
;load fund variable from other scripts
$050_loc = _LoadVar("$050_loc")
$050_regexp = StringRegExp($050_loc, '(^\b[A-Z]{1,}).([0-9]{1,})', 1); this pulls out for example, the first letters of the call number plus numbers up to the first period
If @error == 0 Then
	$050_loc = $050_regexp[0]
Else
	Exit
EndIf

If $050_loc = "G70.2" Then
	MsgBox(0, "Ken", "Is this a map?")
	Exit
Else
	$050_regexp = StringRegExp($050_loc, '(^\b[A-Z]{1,})', 1)
	If @error == 0 Then
		$050_letters = $050_regexp[0]
	EndIf
	$050_regexp = StringRegExp($050_loc, '[0-9]{1,}.([0-9]{1})', 1)
	If @error == 0 Then
		$050_numbers = $050_regexp[0]
	EndIf
EndIf
; now we have the first 1-3 letters of the call number split out from the first set of cutters and can start testing that against the locations

If StringInStr($050_letters, "BF") Then
	If $050_numbers > 1156.9999 Then
		$050_check = "kngl"
		_StoreVar("050_check")
	Else
		$050_check = "scl"
		_StoreVar("050_check")
	EndIf
	ElseIf StringInStr($050_letters, "BF") Then






	If $050_loc = "4b " Then; the macro grabs the first three characters of the fund code (easiest method) so two digit fund codes need to have an extra space for the match
		$SF_NAME = "Baer"
		$SF_590 = "purchased with monies from the Paul W. & Elsa Thoma Baer Fund"
		$SF_7XX = "b7901 Baer, Elsa Thoma,|edonor" & @CR & "b7901 Baer, Paul W.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4bi" Then
		$SF_NAME = "Bierly"
		$SF_590 = "donated in honor of Dr. Willis (Andy) W. Wertz, MU 1932, Dept. of Architecture, 1937-1973, by RIchard Bierly, MU 1954."
		$SF_7XX = "b7901 Bierly, Richard,|edonor" & @CR & "b7901 Wertz, Willis W.,|ehonoree"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4c " Then
		$SF_NAME = "Wertz"
		$SF_590 = "purchased with monies from the Willis W. Wertz Library (Architecture) Acquisitions Fund"
		$SF_7XX = "b7901 Wertz, Willis W.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4dr" Then
		$SF_NAME = "Drake"
		$SF_590 = "purchased with monies from the  the Vivian L. Drake Fund"
		$SF_7XX = "b7901 Drake, Vivian L.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4du" Then
		$SF_NAME = "Duncan"
		$SF_590 = "purchased with monies from the Richard & Dorothy Duncan Fund"
		$SF_7XX = "b7901 Duncan, Richard,|edonor" & @CR & "b7901 Duncan, Dorothy,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4e " Then
		$SF_NAME = "Class of 1920"
		$SF_590 = "purchased with monies from the Class of 1920 University Libraries Enrichment Fund"
		$SF_7XX = "b7912 Class of 1920,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4eb" Then
		$SF_NAME = "Elec.Bus.Res."
		$SF_590 = "purchased with monies from the University Libraries Electronic Business Resources Fund"
		$SF_7XX = "b7912 University Libraries Electronic Business Resources,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4er" Then
		$SF_NAME = "Erwin"
		$SF_590 = "purchased with monies from the David & Susan Erwin Library Acquisitions Fund"
		$SF_7XX = "b7901 Erwin, David,|edonor" & @CR & "b7901 Erwin, Susan,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4f " Then
		$SF_NAME = "Mozingo"
		$SF_590 = "purchased with monies from the Todd Mozingo Memorial Fund"
		$SF_7XX = "b7901 Mozingo, Todd,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4ft" Then ;clipped because it's looking for three characters only
		$SF_NAME = "Foote"
		$SF_590 = "purchased in memory of Diane Altman Ball '54"
		$SF_7XX = "b7901 Ball, Diane Altman,|ehonoree" & @CR & "b7901 Foote, Janet,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4gb" Then
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "b7901 Hovorka, Sophie Nickel,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4gg" Then
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "b7901 Hovorka, Sophie Nickel,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4gp" Then
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "b7901 Hovorka, Sophie Nickel,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4h " Then
		$SF_NAME = "G.Hill R. R."
		$SF_590 = "purchased with monies from the Gretchen Hill Children’s Literature Reading Area Fund"
		$SF_7XX = "b7901 Hill, Gretchen,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4ia" Then
		$SF_NAME = "Iandoli"
		$SF_590 = "purchased with monies from the Marie Iandoli Library Acquisitions Fund"
		$SF_7XX = "b7901 Iandoli, Marie,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4j " Then
		$SF_NAME = "Wesolowski"
		$SF_590 = "purchased with monies from the John Wesolowski Memorial Libraries Endowment"
		$SF_7XX = "b7901 Wesolowski, John,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4k " Then
		$SF_NAME = "I. Williams"
		$SF_590 = "This item was acquired through funding from the William D. Mulliken Library Acquisitions Fund"
		$SF_7XX = "b7901 Williams, Isabella Riggs,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4l " Then
		$SF_NAME = "Mulliken"
		$SF_590 = "purchased with monies from the Isabella Riggs Williams Library Fund"
		$SF_7XX = "b7901 Mulliken, William D.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4ky" Then
		$SF_NAME = "Kirby"
		$SF_590 = "purchased in honor of Dr. Jack T. Kirby, Emeritus, Professor, History (1965-2002) by FILL IN THIS BLANK"
		$SF_7XX = "b7901 Kirby, Jack T.,|ehonoree" & @CR & "b7901 FILL IN THE BLANK"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4lf" Then
		$SF_NAME = "Lyons-Felstead"
		$SF_590 = "purchased in memory of Anna Mae Duvall Lyons '32"
		$SF_7XX = "b7901 Lyons, Anna Mae Duvall,|ehonoree" & @CR & "b7901 Lyons-Felstead, Joan,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4m " Then
		$SF_NAME = "Sinclair"
		$SF_590 = "purchased with monies from the Robert B. Sinclair Library Collections Fund"
		$SF_7XX = "b7901 Sinclair, Robert B.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4n " Then
		$SF_NAME = "Schlachter"
		$SF_590 = "purchased with monies from the Karl & Roberta Schlachter Library Fund"
		$SF_7XX = "b7901 Schlachter, Karl,|edonor" & @CR & "b7901 Schlachter, Roberta,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4naw" Then
		$SF_NAME = "NAWPA"
		$SF_590 = "purchased with monies from the Native American Women Playwrights Archive Fund"
		$SF_7XX = "b7912 Native American Women Playwrights Archive,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4o " Then
		$SF_NAME = "Havighurst"
		$SF_590 = "purchased with monies from the Walter E. Havighurst Special Collections Library Fund"
		$SF_7XX = "b7901 Havighurst, Walter R.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4p " Then
		$SF_NAME = "Peterson/Defoe"
		$SF_590 = "purchased with monies from the Spiro Peterson Center for Defoe Studies"
		$SF_7XX = "b7901 Peterson, Spiro,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4pr" Then
		$SF_NAME = "Pratt Amer. Lit."
		$SF_590 = "purchased with monies from the William C. Pratt American Literature Acquisitions Fund"
		$SF_7XX = "b7901 Pratt, William C.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4q " Then
		$SF_NAME = "NEH"
		$SF_590 = "purchased with monies from the NEH Challenge Fund"
		$SF_7XX = "b7912 NEH Challenge,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4r " Then
		$SF_NAME = "Alumni"
		$SF_590 = "purchased with monies donated by "
		$SF_7XX = "b7911 "
		$SF_FIB = 1
		_StoreVar("$SF_FIB")
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4s " Then
		$SF_NAME = "Covington-Williams"
		$SF_590 = "purchased with monies from the Covington-Williams Families Fund"
		$SF_7XX = "b7903 Covington-Williams Families,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4sc" Then ; should be 4sch, but 4sc doesn't bump into anything and since we're only dealing in the first three chars.
		$SF_NAME = "Shafer"
		$SF_590 = "purchased with monies from the Alice Schafer Fund"
		$SF_7XX = "b7901 Schafer, Alice,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4sm" Then
		$SF_NAME = "Swenson-Mount"
		$SF_590 = "given in memory of Edward & Gladys Swenson and Robert L. Mount Children’s Literature Fund"
		$SF_7XX = "b7901 Swenson, Edward,|ehonoree" & @CR & "b7901 Swenson, Gladys,|ehonoree" & @CR & "b7901 Mount, Robert L.,|ehonoree"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4t " Then
		$SF_NAME = "Brown-Lane"
		$SF_590 = "purchased with monies from the Mariyette Brown-Lane Library Endowment Fund"
		$SF_7XX = "b7901 Brown-Lane, Mariyette,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4u " Then
		$SF_NAME = "Halstead"
		$SF_590 = "purchased with monies from the Halstead Collection Fund"
		$SF_7XX = "b7903 Halstead Collection,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4v " Then
		$SF_NAME = "Richey"
		$SF_590 = "purchased with monies from the Sam W. Richey Collection of Southern Confederacy Fund"
		$SF_7XX = "b7901 Richey, Sam. W.,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4wa" Then
		$SF_NAME = "Joanne Ward IMC"
		$SF_590 = "purchased with monies from the Joanne Stalzer Ward Instructional Materials Center Support Fund"
		$SF_7XX = "b7901 Ward, Joanne Stalzer,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4x " Then
		$SF_NAME = "Fletcher"
		$SF_590 = "purchased with monies from the Alice Harries Fletcher Special Collections Fund"
		$SF_7XX = "b7901 Fletcher, Alice Harries,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "4y " Then
		$SF_NAME = "Sissman"
		$SF_590 = "purchased with monies from the Ben Sissman-WWII Intelligence History Library Fund"
		$SF_7XX = "b7901 Sissman, Ben,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "54 " Then
		$SF_NAME = "Dean’s Special Fund"
		$SF_590 = "purchased with monies from the University Libraries Acquisitions Fund"
		$SF_7XX = "b7912 University Libraries Acquisitions Fund,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")

	ElseIf $050_loc = "wes" Then
		$SF_NAME = "Weslowski"
		$SF_590 = "purchased with monies from the John Weslowski Memorial Libraries Endowment Fund"
		$SF_7XX = "b7901 Weslowksi, John,|edonor"
		_StoreVar("$SF_NAME")
		_StoreVar("$SF_590")
		_StoreVar("$SF_7XX")


	Else
		$SF_NAME = 0

	EndIf

Exit

;store variables for use in other scripts
;_StoreVar("$SF_NAME")
;_StoreVar("$SF_590")
;_StoreVar("$SF_7XX")

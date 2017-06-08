#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: Becky Yoose, Bibliographic Systems Librarian, Miami University
		 yoosebj@muohio.edu OR b.yoose@gmail.com

 Name of script: Special Fund List
 Script set: Multiple

 Script Function:
	This script takes the fund variable and determines if that fund is a
	special fund. If so, the fund name, 590, and 79X variables are set
	and saved to be used in creating notes and fields in order and
	bibliographic records

 Programs used: n/a

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
Dim $FUND
Dim $SF_NAME ;used for order record notes
Dim $SF_590 ;bib record field
Dim $SF_7XX ;bib record fields, can contain multiple 79X entries

;################################ MAIN ROUTINE #################################
;load fund variable from other scripts
$FUND = _LoadVar("$FUND")

Switch $FUND

	Case "4b"
		$SF_NAME = "Baer"
		$SF_590 = "purchased with monies from the Paul W. & Elsa Thoma Baer Fund"
		$SF_7XX = "7901 Baer, Elsa Thoma,|edonor" & @CR & "b7901 Baer, Paul W.,|edonor"

	Case "4bi"
		$SF_NAME = "Bierly"
		$SF_590 = "donated in honor of Dr. Willis (Andy) W. Wertz, MU 1932, Dept. of Architecture, 1937-1973, by RIchard Bierly, MU 1954."
		$SF_7XX = "7901 Bierly, Richard,|edonor" & @CR & "b7901 Wertz, Willis W.,|ehonoree"

	Case "4c"
		$SF_NAME = "Wertz"
		$SF_590 = "purchased with monies from the Willis W. Wertz Library (Architecture) Acquisitions Fund"
		$SF_7XX = "7901 Wertz, Willis W.,|edonor"

	Case "4du"
		$SF_NAME = "Duncan"
		$SF_590 = "purchased with monies from the Richard & Dorothy Duncan Fund"
		$SF_7XX = "7901 Duncan, Richard,|edonor" & @CR & "b7901 Duncan, Dorothy,|edonor"

	Case "4e"
		$SF_NAME = "Class of 1920"
		$SF_590 = "purchased with monies from the Class of 1920 University Libraries Enrichment Fund"
		$SF_7XX = "7912 Class of 1920,|edonor"

	Case "4eb"
		$SF_NAME = "Elec.Bus.Res."
		$SF_590 = "purchased with monies from the University Libraries Electronic Business Resources Fund"
		$SF_7XX = "7912 University Libraries Electronic Business Resources,|edonor"

	Case "4er"
		$SF_NAME = "Erwin"
		$SF_590 = "purchased with monies from the David & Susan Erwin Library Acquisitions Fund"
		$SF_7XX = "7901 Erwin, David,|edonor" & @CR & "b7901 Erwin, Susan,|edonor"

	Case "4f"
		$SF_NAME = "Mozingo"
		$SF_590 = "purchased with monies from the Todd Mozingo Memorial Fund"
		$SF_7XX = "7901 Mozingo, Todd,|edonor"

	Case "4fte"
		$SF_NAME = "Foote"
		$SF_590 = "purchased in memory of Diane Altman Ball '54"
		$SF_7XX = "7901 Ball, Diane Altman,|ehonoree" & @CR & "b7901 Foote, Janet,|edonor"

	Case "4gb"
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "7901 Hovorka, Sophie Nickel,|edonor"

	Case "4gg"
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "7901 Hovorka, Sophie Nickel,|edonor"

	Case "4gp"
		$SF_NAME = "Hovorka"
		$SF_590 = "purchased with monies from the Sophie Nickel Hovorka Library Fund"
		$SF_7XX = "7901 Hovorka, Sophie Nickel,|edonor"

	Case "4h"
		$SF_NAME = "G.Hill R. R."
		$SF_590 = "purchased with monies from the Gretchen Hill Children’s Literature Reading Area Fund"
		$SF_7XX = "7901 Hill, Gretchen,|edonor"

	Case "4ia"
		$SF_NAME = "Iandoli"
		$SF_590 = "purchased with monies from the Marie Iandoli Library Acquisitions Fund"
		$SF_7XX = "7901 Iandoli, Marie,|edonor"

	Case "4j"
		$SF_NAME = "Wesolowski"
		$SF_590 = "purchased with monies from the John Wesolowski Memorial Libraries Endowment"
		$SF_7XX = "7901 Wesolowski, John,|edonor"

	Case "4k"
		$SF_NAME = "I. Williams"
		$SF_590 = "This item was acquired through funding from the William D. Mulliken Library Acquisitions Fund"
		$SF_7XX = "7901 Williams, Isabella Riggs,|edonor"

	Case "4l"
		$SF_NAME = "Mulliken"
		$SF_590 = "purchased with monies from the Isabella Riggs Williams Library Fund"
		$SF_7XX = "7901 Mulliken, William D.,|edonor"

	Case "4lf"
		$SF_NAME = "Lyons-Felstead"
		$SF_590 = "purchased in memory of Anna Mae Duvall Lyons '32"
		$SF_7XX = "7901 Lyons, Anna Mae Duvall,|ehonoree" & @CR & "b7901 Lyons-Felstead, Joan,|edonor"

	Case "4m"
		$SF_NAME = "Sinclair"
		$SF_590 = "purchased with monies from the Robert B. Sinclair Library Collections Fund"
		$SF_7XX = "7901 Sinclair, Robert B.,|edonor"

	Case "4n"
		$SF_NAME = "Schlachter"
		$SF_590 = "purchased with monies from the Karl & Roberta Schlachter Library Fund"
		$SF_7XX = "7901 Schlachter, Karl,|edonor" & @CR & "b7901 Schlachter, Roberta,|edonor"

	Case "4naw"
		$SF_NAME = "NAWPA"
		$SF_590 = "purchased with monies from the Native American Women Playwrights Archive Fund"
		$SF_7XX = "7912 Native American Women Playwrights Archive,|edonor"

	Case "4o"
		$SF_NAME = "Havighurst"
		$SF_590 = "purchased with monies from the Walter E. Havighurst Special Collections Library Fund"
		$SF_7XX = "7901 Havighurst, Walter R.,|edonor"

	Case "4p"
		$SF_NAME = "Peterson/Defoe"
		$SF_590 = "purchased with monies from the Spiro Peterson Center for Defoe Studies"
		$SF_7XX = "7901 Peterson, Spiro,|edonor"

	Case "4pr"
		$SF_NAME = "Pratt Amer. Lit."
		$SF_590 = "purchased with monies from the William C. Pratt American Literature Acquisitions Fund"
		$SF_7XX = "7901 Pratt, William C.,|edonor"

	Case "4q"
		$SF_NAME = "NEH"
		$SF_590 = "purchased with monies from the NEH Challenge Fund"
		$SF_7XX = "7912 NEH Challenge,|edonor"

	Case "4s"
		$SF_NAME = "Covington-Williams"
		$SF_590 = "purchased with monies from the Covington-Williams Families Fund"
		$SF_7XX = "7903 Covington-Williams Families,|edonor"

	Case "4sch"
		$SF_NAME = "Shafer"
		$SF_590 = "purchased with monies from the Alice Schafer Fund"
		$SF_7XX = "7901 Schafer, Alice,|edonor"

	Case "4sm"
		$SF_NAME = "Swenson-Mount"
		$SF_590 = "given in memory of Edward & Gladys Swenson and Robert L. Mount Children’s Literature Fund"
		$SF_7XX = "7901 Swenson, Edward,|ehonoree" & @CR & "b7901 Swenson, Gladys,|ehonoree" & @CR & "b7901 Mount, Robert L.,|ehonoree"

	Case "4t"
		$SF_NAME = "Brown-Lane"
		$SF_590 = "purchased with monies from the Mariyette Brown-Lane Library Endowment Fund"
		$SF_7XX = "7901 Brown-Lane, Mariyette,|edonor"

	Case "4u"
		$SF_NAME = "Halstead"
		$SF_590 = "purchased with monies from the Halstead Collection Fund"
		$SF_7XX = "7903 Halstead Collection,|edonor"

	Case "4v"
		$SF_NAME = "Richey"
		$SF_590 = "purchased with monies from the Sam W. Richey Collection of Southern Confederacy Fund"
		$SF_7XX = "7901 Richey, Sam. W.,|edonor"

	Case "4wa"
		$SF_NAME = "Joanne Ward IMC"
		$SF_590 = "purchased with monies from the Joanne Stalzer Ward Instructional Materials Center Support Fund"
		$SF_7XX = "7901 Ward, Joanne Stalzer,|edonor"

	Case "4x"
		$SF_NAME = "Fletcher"
		$SF_590 = "purchased with monies from the Alice Harries Fletcher Special Collections Fund"
		$SF_7XX = "7901 Fletcher, Alice Harries,|edonor"

	Case "4y"
		$SF_NAME = "Sissman"
		$SF_590 = "purchased with monies from the Ben Sissman-WWII Intelligence History Library Fund"
		$SF_7XX = "7901 Sissman, Ben,|edonor"

	Case "54"
		$SF_NAME = "Dean’s Special Fund"
		$SF_590 = "purchased with monies from the University Libraries Acquisitions Fund"
		$SF_7XX = "7912 University Libraries Acquisitions Fund,|edonor"

	Case Else
		$SF_NAME = "none"

EndSwitch

;store variables for use in other scripts
_StoreVar("$SF_NAME")
_StoreVar("$SF_590")
_StoreVar("$SF_7XX")

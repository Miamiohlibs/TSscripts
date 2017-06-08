DISCLAIMER
----------------
Macros cannot check for some human error in the ordering process. Therefore, if the order team inserts an incorrect location code in the order record, then that mistake is not noticed by the script and will accept that code as correct. If you are wondering why a script is not processing an item properly, it is probably caused by a key indicator (used by the macro to determine processing) that is incorrect or missing.

Otherwise, enjoy and report any bugs so they can be squashed accordingly.
----------------


***version 0.9.1 (4/2/09)***

Fixed following issues:

*260 |c brackets - should no longer matter if there are brackets in the 260|c date for date match
*300 |c # x #cm folio - should take into acount both numbers if there is more than 1 measurement for folio consideration
*Checking series to determine classing will happen only when record will not be bumped
*Added item records for non-bump volume set - macro should automatically ask if another item record should be created for multiple volumes
*300 |e messages and codes should work now!


/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*


***Version 0.9 (3/30/09)***


What the buttons mean and do (3/30/09)

"Take one" icon
Same as F4 macro. Scan in the barcode, pulls up bib record, copies record, closes record, does a title search. If there are more than one search result (aka if the order record does not automatically pop up) the macro will prompt you to check for possible dups in the search results.


"O" icon
Similar to "o" macro. Macro will automatically move the mouse pointer into the data fields and click to place the cursor in the data field. The macro will check/do the following:
	*E-price check
	*RDATE check 
	*RLOC check
	*ORDTYPE check
	*Middletown DVD check and bump
	*Multiple copies and volume
		**Pop up will ask you if you have received all copies/volumes. If yes, macro will prompt you to enter # of copies and will generate internal note. If no, macro will do same as above but will save the record before going on.
	*Textbook fund check and bump
	*STATUS check and bump
	*Music score form check and bump
	*MLC CD check, bump and print
	*IMC DVD check and bump
	*Added vol/copy check and bump
	*300 |e internal note (if appropriate) - only if item is marked to be received


"B" icon
Similar to "b" macro. Pulls up bib record, copies record. The macro will check/do the following:
	*CDATE check
	*MARC Leader 8 check
	*DLC|DLC - DNLM/DLC|DLC - pcc check
	*050 check
	*Bib location check if record is approval
	*SKIP check
	*300 check for p.|ccm, folio (> 30 cm), music book with |e
	*multi location check
	*IMC Juv Location check
	*BCODE1 thru 3 check
	*008, 050, 260 date check
	*Series check
	*Special funds check - only if item is marked to be received
	*Reference location input - only if item is marked to be received
The macro will also determine the item location from the order location code (for Firm and Notification) or from the call number (for approval)


"D" icon
D for "done" or "DLC". This macro replaces the "l" and "r" macro. Macro will insert the bib location, initials, and save bib record. Switches over to create a new item record and inserts appropriate information:
	*ICODE1 
	*location
	*STATUS
		**Macro will prompt asking if item is a paperback. If so, or if the item needs to go to conservation, press "Yes" and macro will insert "r" in STATUS. Otherwise, STATUS will be "l". 
	*copy/vol number
	*label loc if appropriate
	*accompanying material message code and note
	*barcode
Macro will insert and save record. If the item is a DOC with an accompanying item, macro will generate a new item record for said item. From here you can press the "take one" icon to automatically close the item record and start another item.


"N" icon
N for "non-DLC" or "Nope". This macro replaces the "i" macro. Macro will automatically run if item is bumped by the "O" or "B" macros. Macro will create a new item record and inserts appropriate information:
	*ICODE1 
	*location
	*STATUS
	*copy/vol number
	*label loc if appropriate
	*accompanying material message code and note
	*barcode
Prompt will note what color slip to put in item and any other instructions. If the item is a DOC with an accompanying item, macro will generate a new item record for said item. From here you can press the "take one" icon to automatically close the item record and start another item.


"Stop Sign" icon
Panic button. If you need to stop a script for any reason, click on the stop sign.


~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
New features (3/30/09)

*Taskbar icon matches toolbar icon - when a script is running, the corresponding taskbar icon will match the running script's toolbar icon. Ex. - a "B" will be where the running man used to be when the "B" macro is running.

*300 |e order record internal note - macro will insert an internal note in the order record. Pop up will ask to see if all material has been received. If so, macro will insert note. If not, macro will stop.

*Macro will automatically bump music books that have accompanying material

*O macro will automatically move the mouse pointer into the data fields and click to place the cursor in the data field. No longer need to manually click on the data fields yourself!

*Macro will now accept the following material as DLC|DLC: 
**040 field - DNLM/DLC|cDLC
**042 field - pcc

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Known issues (3/30/09)

*Added volumes - macro will not add item records for volume sets that are DLC/pcc
*Brackets in 260 |c field - macro will bump items with bracketed 260|c dates
*Tool bar will not resize
*Taskbar icon will not close after script has been stopped using the panic button - will only close after you mouse over the taskbar.
*Any speed issues - caused by III and, unless III magically improves its interface and processing, unlikely going to be solved any time soon.

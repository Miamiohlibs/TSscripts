;~ fun happy function list time
;~ use control+f to find specific functions

#include-once

;~ some of the function libraries that are used for the functions below
#include <Misc.au3>
#include <Array.au3>
#include <WinAPI.au3>

;~ ****************************************************************************
;~ 			*******			UNIVERSAL FUNCTIONS			*******
;~ ****************************************************************************



;########### function to fix sticking mod keys #############
Func _SendEx($ss)
    Local $iT = TimerInit()

    While _IsPressed("10") Or _IsPressed("11") Or _IsPressed("12")
        If TimerDiff($iT) > 1000 Then
            MsgBox(262144, "Warning", "Shift, Ctrl and Alt Keys need to be released to proceed!" & @CR & "Click OK to release keys.")
			_SendEx("{CTRLUP}")
			_SendEx("{SHIFTUP}")
			_SendEx("{ALTUP}")
		EndIf
    WEnd
    Send($ss)

EndFunc ;==>_SendEx

;############## end function to fix sticking mod keys #########


;########### functions to save and transfer variables #############
Func _StoreVar($VarName)
	RegWrite("HKCU64\Environment\AutoItVarBuffer",$VarName,"REG_SZ",Execute($VarName))
EndFunc

Func _LoadVar($VarName)
    Return RegRead("HKCU64\Environment\AutoItVarBuffer",$VarName)
EndFunc

Func _ClearBuffer()
    RegDelete("HKCU64\Environment\AutoItVarBuffer")
EndFunc
;########### end function to save and transfer variables #############


;########### function to copy window contents #############
Func _DataCopy()
	Sleep(0100)
	ClipPut("")
	Sleep(0100)
	_SendEx("^{HOME}")
	Sleep(0100)
	_SendEx("^a")
	Sleep(0100)
	_SendEx("^c")
	Sleep(0100)
	_SendEx("^c")
EndFunc
;########### end function to copy window contents #############



;########### function to turn array elements into strings #############
Func _arrayItemString($y, $x)

	Local $z
	Local $w
	$w = _ArraySearch($y, $x, 0, 0, 0, 1)
	If $w > -1 Then
		$z = _ArrayToString($y, @TAB, $w, $w)
		Return $z
	Else
		Return $w
	EndIf

EndFunc
;########### end function to turn array elements into strings #############


;########### function to search OCLC #############
Func _OCLCSearch($a, $x)
	;If WinExists("OCLC Connexion") Then
		WinActivate("OCLC Connexion")
		WinWaitActive("OCLC Connexion")
	;Else
		;MsgBox(64, "nope", "Please open and log into Connexion. Click ok AFTER you are logged in")
	;EndIf
	Sleep(0300)
	_SendEx("!c")
	Sleep(0300)
	_SendEx("{RIGHT}")
	Sleep(0300)
	_SendEx("{ENTER}")
	Sleep(0300)
	_SendEx($a)
	Sleep(0100)
	_SendEx($x)
	Sleep(0100)
	_SendEx("{ENTER}")
EndFunc
;########### End function to search OCLC #############


;########### function to determine 947 initials #############

Func _Initial()
	local $var
	local $C_INI
	$var = EnvGet("TEMP")

	$var = StringTrimLeft($var, 9)
	$var = StringTrimRight($var, 19)

	Switch $var
		Case "alexanpk"
			$C_INI = "pka"
		Case "barbouh2"
			$C_INI = "hlb"
		Case "keyessl"
			$C_INI = "slk"
		Case "stepanm"
			$C_INI = "mls"
		Case "barbouh2"
			$C_INI = "hlb"
	    Case "marlinsa"
			$C_INI = "sab"
		Case "bomanca"
			$C_INI = "cb"
		case Else
			$C_INI = InputBox("initials", "Please enter your initials!")
	EndSwitch
	Return $C_INI
EndFunc

;########### End function to determine 947 initials #############



;########### function to do a III search from the main screen #############
Func _IIIsearch($a, $b)
	If WinExists("[TITLE:Sierra; CLASS:SunAwtFrame]") Then
		WinActivate("[TITLE:Sierra; CLASS:SunAwtFrame]")
	Else
		MsgBox(64, "nope", "Please log into Sierra.")
		Exit
	EndIf
	WinWaitActive("[TITLE:Sierra; CLASS:SunAwtFrame]")
	Sleep(0500)
	;_SendEx("!n")
	;upgrade v.1.1 change
	_SendEx("!t")
	Sleep(0100)
	;_SendEx("!n")
	_SendEx("s")
	Sleep(0200)
	_SendEx("{ENTER}")
	Sleep(0200)
	_SendEx($a)
	Sleep(0200)
	_SendEx($b)
	Sleep(0100)
	_SendEx("{ENTER}")
	Sleep(0400)
EndFunc

;########### function to do a III search from the main screen #############


;########### function to left mouse click on data fields in III #############
Func _WinClick($x)
	Local $winwidth, $winheight
	AutoItSetOption("MouseCoordMode", 0)
	$winwidth = _ArrayToString($x, "", 2, 2)
	$winheight = _ArrayToString($x, "", 3, 3)
	$winwidth *= .3
	$winheight *= .7
	MouseMove($winwidth, $winheight)
	MouseClick("left")
EndFunc
;########### function to left mouse click on data fields in III #############


;########### function to help create variables from III fixed fields #############
Func _IIIfixed($a, $b)
	local $x, $y, $z
	$x = _ArrayToString($a, @TAB, $b, $b)
	$y = StringRegExpReplace ($x, "\t+[A-Z]", "fnord")
	$z = StringSplit($y, "fnord", 1)
	Return $z
EndFunc
;########### End function to help create variables from III fixed fields #############
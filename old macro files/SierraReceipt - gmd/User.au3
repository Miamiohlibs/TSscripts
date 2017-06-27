#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\autoiticon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <TSCustomFunction.au3>

$var = EnvGet("TEMP")
MsgBox(4096, "Path variable is:", $var)
;~ ClipPut($var)
$var = StringTrimLeft($var, 12)
$var = StringTrimRight($var, 14)

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
	case Else
		$C_INI = "999"
EndSwitch
msgbox(0, "ini", $C_INI)

_StoreVar("$C_INI")
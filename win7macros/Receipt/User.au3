#include <SaveTransferVariables.au3>

$var = EnvGet("TEMP")
MsgBox(4096, "Path variable is:", $var)
;~ ClipPut($var)
$var = StringTrimLeft($var, 12)
$var = StringTrimRight($var, 14)

Switch $var
	Case "flanagk"
		$C_INI = "kmf"
	Case "klumbca"
		$C_INI = "cak"
	Case "scottjs"
		$C_INI = "jjs"
	Case "spencert"
		$C_INI = "rts"
	Case "hanseld"
		$C_INI = "dlh"
	Case "yoosebj"
		$C_INI = "bjy"
	case Else
		$C_INI = "999"
EndSwitch
	
StoreVar("$C_INI")
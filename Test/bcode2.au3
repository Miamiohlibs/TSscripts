
#include <MsgBoxConstants.au3>

Example()

Func Example()
   Local $bibid = '4530689'
   Local $output = ""

   ; Run Notepad and wait for the Notepad process to close.
    Local $iReturn = RunWait("C:\\Users\\bomanca\\AppData\\Local\\Programs\\Python\\Python36\\python.exe", "C:\\Users\\bomanca\\Documents\\GitHub\\python\\bcode2.py", $bibid)
    ProcessWaitClose($iReturn)

   $output &= StdoutRead($iReturn)

    ; Display the return code of the Notepad process.
    MsgBox($MB_SYSTEMMODAL, "", "The return code from Notepad was: " & $output)

EndFunc   ;
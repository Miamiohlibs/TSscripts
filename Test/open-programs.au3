㴿;######### DECLARE VARIABLES #########


  Local $Action ;
  Local $FileName;
  ;STRING cParams, ;
  ;STRING cDir, ;
  ;INTEGER nShowWin


  $FileName = "C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe"
  $Action = "open"



;ShellExecute("chrome","","") ;chrome works


;can't get it to open spotify or keepass
;ShellExecute("keepass.exe","c:\\Program Files (x86)\KeePass Password Safe 2\","")

;attempting to open explorer
ShellExecute("explorer.exe","C:\Program Files (x86)\KeePass Password Safe 2\","")
Opt("WinDetectHiddenText", 1)
Opt("WinSearchChildren", 1)


AutoItSetOption("WinTitleMatchMode", 2)

WinActivate("Mail from the Library")
WinWaitActive("Mail from the Library")


ClipPut("")

Send("^a")
Send("^c")

; reset titlematchmode
AutoItSetOption("WinTitleMatchMode", 4)

If WinExists("[TITLE:OCLC Connexion]") Then
		WinActivate("[TITLE:OCLC Connexion]")
	Else 
		Run(@ProgramFilesDir & "\OCLC\Connexion\Program\Connex.exe")
	EndIf



WinWait("[TITLE:OCLC Connexion]", "menuStrip1", 2)

Send("!b")
Sleep(0100)
Send("h")


WinWait("[TITLE:Batch Holdings by OCLC Number]")

Send("^v")
Sleep(0100)
Send("!u")
Sleep(0100)
Send("{TAB}")
Send("{SPACE}")
Send("{TAB}")
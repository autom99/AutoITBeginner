HotKeySet('h','HotKey1')
HotKeySet('p','HotKey2')
HotKeySet('x','ExitProgram')

While 1
	Sleep(50)
WEnd

Func HotKey1()
	ConsoleWrite("The Hotkey was presseed.!" & @CRLF)
EndFunc

Func HotKey2()
;~ 	MsgBox(0,"This hotkey worked","This hotkey was pressed")
	Run('notepad.exe')
EndFunc

Func ExitProgram()
	Exit
EndFunc
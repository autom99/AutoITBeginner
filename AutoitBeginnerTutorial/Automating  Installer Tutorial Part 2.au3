#RequireAdmin
AutoItSetOption('MouseCoordMode',0)
AutoItSetOption('SendKeyDelay',5)

;Run(@ScriptDir & '\ccsetup558.exe')

;********* Method of Clicking/Checking/Unchecking ************

; Method 1
;~ WinActivate('','View License Agreement')
;~ MouseClick('primary',43,171,0)
;~ Sleep(100)

; Method 2
;ControlClick('CCleaner v5.58 Setup','','button')

; Method 3
;ControlCommand()

;********* Method of Typing text into TEXTAREA /textboxes ************

; Method 1
;~ WinActivate('Untitled - Notepad')
;~ Send("This is a test to see this text typed into the box",1)

; Method 2
;~ $text = "This is a test to see this text typed into the box"
;~ ClipPut($text)
;~ WinActivate('Untitled - Notepad')
;~ Send('^V')

; Method 3
;~ ControlSend("Untitled - Notepad","","Edit1","This is a test to see this text typed into the box")

; Method 4
ControlSetText("Untitled - Notepad","","Edit1","This is a test to see this text typed into the box")






Global $message,$name

$message = "Hello and Welcome"

$name = "Whisely"

MsgBox(0,$message,"My name is " & $name)

MsgBox(0,$message,$name)

MsgBox(0,$message,"My name is " & $name & "."  & "This will be display off after 2 sec..!!",2)
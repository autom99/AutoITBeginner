Global $buttonPressed

$buttonPressed = MsgBox(4,"Select Choice","Press Yes or No")

If $buttonPressed = 6 Then
	MsgBox(0,"You have selected","You have pressed YES button!")
ElseIf $buttonPressed = 7 Then
	MsgBox(0,"You have selected","You have pressed NO button!")
EndIf



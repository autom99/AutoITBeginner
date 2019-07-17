Global $first,$second,$third,$fourth

$first = 100
$second = 125
$third = $first + $second
$fourth = 1250

;Local $res = $fourth % 2
;MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & ' $res' & @CRLF & @CRLF & 'Return:' & @CRLF &  $res) ;### Debug MSGBOX
;ConsoleWrite("$fourth % 2" & $res)

If $second > $first Then
	MsgBox(0,"Second is greater than First","Second is " & $second)
EndIf

;MsgBox(0,"Result","$first + $second = " & $third)
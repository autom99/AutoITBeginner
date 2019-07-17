Global $first,$second,$third,$fourth

$first = 1

$second = 5

;~ Do
;~ 	MsgBox(0,"Text",$first)
;~ 	$first += 1
;~ Until $first = 5


;~ While $first <> 5
;~ 	MsgBox(0,"Text",$first)
;~ 	$first += 1
;~ WEnd

;~ For $a = 1 To 10  ;Step + 1
;~ 	MsgBox(0,"Number",$a)
;~ Next

For $a = $first To $second   ;Step + 1
	MsgBox(0,"Number",$a)
Next
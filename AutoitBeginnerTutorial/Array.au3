#include <String.au3>

$Source = BinaryToString(InetRead("https://www.youtube.com/watch?v=IeH5iKfI8DY&list=PLNpExbvcyUkOJvgxtCPcKsuMTk9XwoWum"),1)
$FirstChunks = _StringBetween($Source,'content="','">')

For $a In $FirstChunks
	ConsoleWrite($a & @CRLF)
	Sleep(1000)
Next

;~ Global $Array[10],$Variable

;~ $Array[0] = 0
;~ $Array[1] = 123
;~ $Array[9] = "Finish"
;~ $Array[19] = "Success"

;~ For $a = 0 To 9
;~ 	ConsoleWrite($Array[$a] & @CRLF)
;~ Next

;~ For $a = 0 To (UBound($Array) - 1)  ;9
;~ 	ConsoleWrite($Array[$a] & @CRLF)
;~ Next

;~ For $a In $Array
;~ 	ConsoleWrite($a & @CRLF)
;~ Next



#include "..\..\Includes\DotNet.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	; Start the latest version of .NET Framework
	; DotNet_Start() is automatically called by the other functions in DotNet.au3
	; You only have to call DotNet_Start() if you need an older version of .NET Framework
	DotNet_Start()
	If @error Then Return ConsoleWrite( "DotNet_Start ERR" & @CRLF )
	ConsoleWrite( "DotNet_Start OK" & @CRLF )
EndFunc

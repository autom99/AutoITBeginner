#include "..\..\Includes\DotNet.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	Local $oCode = DotNet_LoadCScode( FileRead( "CodeCS (C# syntax error).cs" ), "System.Windows.Forms.dll" )
	If @error Then Return ConsoleWrite( "DotNet_LoadCScode ERR" & @CRLF )
	ConsoleWrite( "DotNet_LoadCScode OK" & @CRLF )
	Local $oFoo = DotNet_CreateObject( $oCode, "Foo" )
	If IsObj( $oFoo ) Then $oFoo.Test()
EndFunc

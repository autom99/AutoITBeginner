#include "..\..\Includes\DotNet.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	Local $oCode = DotNet_LoadVBcode( FileRead( "CodeVB.vb" ), "System.Windows.Forms.dll" )
	If @error Then Return ConsoleWrite( "DotNet_LoadVBcode ERR" & @CRLF )
	ConsoleWrite( "DotNet_LoadVBcode OK" & @CRLF )
	; Creates an object as an instance of the "Foo" class
	Local $oFoo = DotNet_CreateObject( $oCode, "Foo" )
	If IsObj( $oFoo ) Then $oFoo.Test()
EndFunc

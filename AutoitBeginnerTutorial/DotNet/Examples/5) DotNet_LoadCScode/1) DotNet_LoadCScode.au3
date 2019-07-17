#include "..\..\Includes\DotNet.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	; Compiles C# source code on the fly
	; Creates a .NET assembly DLL-file in memory
	; Loads the memory assembly file into the default domain
	Local $oCode = DotNet_LoadCScode( FileRead( "CodeCS.cs" ), "System.Windows.Forms.dll" ) ; You must add the "System.Windows.Forms.dll"
	If @error Then Return ConsoleWrite( "DotNet_LoadCScode ERR" & @CRLF )                   ; assembly because of line 1 in the C# code.
	ConsoleWrite( "DotNet_LoadCScode OK" & @CRLF )
	Local $oFoo = DotNet_CreateObject( $oCode, "Foo" ) ; $oFoo is an object of the "Foo" class in the C# code
	If IsObj( $oFoo ) Then $oFoo.Test()                ; Test is a method (function) of the $oFoo object
EndFunc

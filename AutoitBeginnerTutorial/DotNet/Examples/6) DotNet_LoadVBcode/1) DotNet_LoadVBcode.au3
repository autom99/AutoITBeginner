#include "..\..\Includes\DotNet.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	; Compiles VB source code on the fly
	; Creates a .NET assembly DLL-file in memory
	; Loads the memory assembly file into the default domain
	Local $oCode = DotNet_LoadVBcode( FileRead( "CodeVB.vb" ), "System.Windows.Forms.dll" ) ; You must add the "System.Windows.Forms.dll"
	If @error Then Return ConsoleWrite( "DotNet_LoadVBcode ERR" & @CRLF )                   ; assembly because of line 1 in the VB code.
	ConsoleWrite( "DotNet_LoadVBcode OK" & @CRLF )
	Local $oFoo = DotNet_CreateObject( $oCode, "Foo" ) ; $oFoo is an object of the "Foo" class in the VB code
	If IsObj( $oFoo ) Then $oFoo.Test()                ; Test is a method (function) of the $oFoo object
EndFunc

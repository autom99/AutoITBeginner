#include "..\..\Includes\DotNet.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	Local $oComErrFunc = ObjEvent( "AutoIt.Error", "ComErrFunc" ) ; See COMUtils.au3
	Local $oCode = DotNet_LoadVBcode( FileRead( "CodeVB.vb" ), "System.Windows.Forms.dll" )
	If @error Then Return ConsoleWrite( "DotNet_LoadVBcode ERR" & @CRLF )
	ConsoleWrite( "DotNet_LoadVBcode OK" & @CRLF )
	Local $oFoo = DotNet_CreateObject( $oCode, "Foo" )
	If IsObj( $oFoo ) Then $oFoo.Tst() ; Typing error: Tst instead of Test
	#forceref $oComErrFunc
EndFunc

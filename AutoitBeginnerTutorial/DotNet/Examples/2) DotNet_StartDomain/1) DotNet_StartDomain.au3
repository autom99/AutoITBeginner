#include "..\..\Includes\DotNetUtils.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	Local $oMyDomain
	; Creates and starts a user domain to run .NET code
	; A default domain is automatically created by the other functions in DotNet.au3
	; You only need to call DotNet_StartDomain() if you want to specify a base directory for the domain
	DotNet_StartDomain( $oMyDomain )
	If Not IsObj( $oMyDomain ) Then Return ConsoleWrite( "DotNet_StartDomain ERR" & @CRLF )
	ConsoleWrite( "DotNet_StartDomain OK" & @CRLF )
	DotNet_ListDomainsEx()
EndFunc

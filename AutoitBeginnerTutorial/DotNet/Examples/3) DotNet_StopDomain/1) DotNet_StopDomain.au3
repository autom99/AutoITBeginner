#include "..\..\Includes\DotNetUtils.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	Local $oMyDomain

	DotNet_StartDomain( $oMyDomain, "MyDomain", "C:\Windows\Temp" )
	If Not IsObj( $oMyDomain ) Then Return ConsoleWrite( "DotNet_StartDomain ERR" & @CRLF )
	ConsoleWrite( "DotNet_StartDomain OK" & @CRLF )
	DotNet_ListDomainsEx()

	; Unloads .NET assemblies loaded into the user domain and stops the domain
	DotNet_StopDomain( $oMyDomain )
	ConsoleWrite( "DotNet_StopDomain OK" & @CRLF )
	DotNet_ListDomainsEx()

	DotNet_StartDomain( $oMyDomain, "MyDomain", "C:\Windows\Temp" )
	If Not IsObj( $oMyDomain ) Then Return ConsoleWrite( "DotNet_StartDomain ERR" & @CRLF )
	ConsoleWrite( "DotNet_StartDomain OK" & @CRLF )
	DotNet_ListDomainsEx()

	; Unloads .NET assemblies loaded into the user domain and stops the domain
	DotNet_StopDomain( $oMyDomain )
	ConsoleWrite( "DotNet_StopDomain OK" & @CRLF )
	DotNet_ListDomainsEx()
EndFunc

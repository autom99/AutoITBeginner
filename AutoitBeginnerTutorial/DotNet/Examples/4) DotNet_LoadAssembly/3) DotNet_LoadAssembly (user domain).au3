#include "..\..\Includes\DotNetUtils.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	Local $oMyDomain
	; Create and start a user domain to run .NET code
	; Set a base directory to load .NET assemblies into the domain
	DotNet_StartDomain( $oMyDomain, "MyDomain", @ScriptDir )
	DotNet_ListDomainsEx()

	; Loads a .NET assembly DLL-file into a user domain
	Local $oXPTable = DotNet_LoadAssembly( "XPTable.dll", $oMyDomain )
	If @error Then Return ConsoleWrite( "DotNet_LoadAssembly ERR" & @CRLF )
	ConsoleWrite( "DotNet_LoadAssembly OK" & @CRLF )
	DotNet_ListAssembliesEx( $oMyDomain )
	#forceref $oXPTable
EndFunc

#include "..\..\Includes\DotNetUtils.au3"

Opt( "MustDeclareVars", 1 )

Example()

Func Example()
	; Loads a .NET assembly DLL-file into the default domain
	Local $oXPTable = DotNet_LoadAssembly( "XPTable.dll" )
	If @error Then Return ConsoleWrite( "DotNet_LoadAssembly ERR" & @CRLF )
	ConsoleWrite( "DotNet_LoadAssembly OK" & @CRLF )
	DotNet_ListDomainsEx()
	DotNet_ListAssembliesEx()
	#forceref $oXPTable
EndFunc

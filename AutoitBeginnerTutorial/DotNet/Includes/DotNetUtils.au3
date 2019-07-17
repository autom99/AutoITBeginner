#include-once
#include <Array.au3>
#include "DotNet.au3"

Func DotNet_ListDomains()
	Local $pEnum, $pDomain, $oDomain
	Local $aDomains[100][2], $iDomains = 0
	Local $oRuntimeHost = DotNet_Start()
	$oRuntimeHost.EnumDomains( $pEnum )
	While Not $oRuntimeHost.NextDomain( $pEnum, $pDomain )
		$oDomain = ObjCreateInterface( $pDomain, $sIID__AppDomain, $sTag__AppDomain )
		$oDomain.get_FriendlyName( $aDomains[$iDomains][0] )
		$oDomain.get_BaseDirectory( $aDomains[$iDomains][1] )
		$iDomains += 1
	WEnd
	ReDim $aDomains[$iDomains][2]
	Return $aDomains
EndFunc

Func DotNet_ListDomainsEx()
	Local $aDomains = DotNet_ListDomains()
	_ArrayDisplay( $aDomains, ".NET Framework Domains", "", 0, Default, "Friendly Name|Base Directory to load assemblies" )
EndFunc

Func DotNet_ListAssemblies( $oAppDomain = 0 )
	Local $pAssemblies, $aAssemblies
	If Not $oAppDomain Then _
		$oAppDomain = DotNet_GetDefaultDomain()
	$oAppDomain.getAssemblies( $pAssemblies )
	AccVars_SafeArrayToArray( $pAssemblies, $aAssemblies )
	Local $nAsms = UBound( $aAssemblies ), $aAsmInfo[$nAsms][3], $oAssembly
	For $i = 0 To $nAsms - 1
		$oAssembly = ObjCreateInterface( $aAssemblies[$i], $sIID__Assembly, $sTag__Assembly )
		$oAssembly.get_GlobalAssemblyCache( $aAsmInfo[$i][0] )
		$oAssembly.get_FullName( $aAsmInfo[$i][1] )
		$oAssembly.get_Location( $aAsmInfo[$i][2] )
	Next
	Return $aAsmInfo
EndFunc

Func DotNet_ListAssembliesEx( $oAppDomain = 0 )
	Local $aAssemblies = DotNet_ListAssemblies( $oAppDomain )
	_ArrayDisplay( $aAssemblies, "Loaded Assemblies in Domain", "", 0, Default, "GAC|Assembly Full Name|Assembly Location" )
EndFunc

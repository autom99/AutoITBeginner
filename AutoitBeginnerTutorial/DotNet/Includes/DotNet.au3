#include-once
#include "Interfaces.au3"
#include "AccVarsUtilities.au3"

Func DotNet_Start( $sVersion = "" )
	Static $pRuntimeHost = 0, $oRuntimeHost
	If $pRuntimeHost Then Return $oRuntimeHost

	If $sVersion = "" Then
		Local $sPath = @WindowsDir & "\Microsoft.NET\Framework" & ( @AutoItX64 ? "64" : "" ) & "\"
		Local $hSearch = FileFindFirstFile( $sPath & "v?.*" ), $sFolder
		If $hSearch <> -1 Then
			While 1
				$sFolder = FileFindNextFile( $hSearch )
				Local $iError = @error, $iExtended = @extended
				If $iExtended = 1 And FileExists( $sPath & $sFolder & "\mscorlib.dll" ) And $sVersion < $sFolder Then $sVersion = $sFolder
				If $iError Then ExitLoop
			WEnd
		EndIf
	ElseIf Not FileExists( @WindowsDir & "\Microsoft.NET\Framework" & ( @AutoItX64 ? "64" : "" ) & "\" & $sVersion ) Then
		Return SetError( 1,0,0 )
	EndIf

	Local Const $tagGUID = "struct; ulong Data1;ushort Data2;ushort Data3;byte Data4[8]; endstruct"

	Local $tCLSID_CorRuntimeHost = DllStructCreate( $tagGUID )
	GUIDFromStringEx( $sCLSID_CorRuntimeHost, $tCLSID_CorRuntimeHost )

	Local $tIID_ICorRuntimeHost = DllStructCreate( $tagGUID )
	GUIDFromStringEx( $sIID_ICorRuntimeHost, $tIID_ICorRuntimeHost )

	Local $aRet = DllCall( "MSCorEE.dll", "long", "CorBindToRuntimeEx", "wstr", $sVersion, "ptr", NULL, "dword", 0, _
	                       "struct*", $tCLSID_CorRuntimeHost, "struct*", $tIID_ICorRuntimeHost, "ptr*", 0 )
	If Not ( @error = 0 And $aRet[0] = 0 And $aRet[6] ) Then Return SetError( 1,0,0 )

	$pRuntimeHost = $aRet[6]
	$oRuntimeHost = ObjCreateInterface( $pRuntimeHost, $sIID_ICorRuntimeHost, $sTag_ICorRuntimeHost )
	Return $oRuntimeHost
EndFunc

Func DotNet_StartDomain( ByRef $oAppDomain, $sFriendlyName = "", $sBaseDirectory = "" )
	Local $oDefDomain = DotNet_GetDefaultDomain()

	Local $pType, $oType
	$oDefDomain.GetType( $pType )
	$oType = ObjCreateInterface( $pType, $sIID__Type, $sTag__Type )

	Local $psaEmpty = SafeArrayCreateEmpty( $VT_VARIANT )
	Local $aArguments[5], $pSafeArray
	$aArguments[0] = $sFriendlyName
	$aArguments[2] = $sBaseDirectory
	$aArguments[4] = False
	AccVars_ArrayToSafeArray( $aArguments, $pSafeArray )

	$oType.InvokeMember_3( "CreateDomain", $BindingFlags_DefaultValue, $psaEmpty, $psaEmpty, $pSafeArray, $oAppDomain )
EndFunc

Func DotNet_StopDomain( ByRef $oAppDomain )
	Local $oRuntimeHost = DotNet_Start()
	$oRuntimeHost.UnloadDomain( Ptr( $oAppDomain ) )
	$oAppDomain = 0
EndFunc

; Internal function
Func DotNet_GetDefaultDomain()
	Static $pDefDomain = 0, $oDefDomain
	If $pDefDomain Then Return $oDefDomain

	Local $oRuntimeHost = DotNet_Start()
	$oRuntimeHost.Start()

	$oRuntimeHost.GetDefaultDomain( $pDefDomain )
	$oDefDomain = ObjCreateInterface( $pDefDomain, $sIID__AppDomain, $sTag__AppDomain )
	Return $oDefDomain
EndFunc

Func DotNet_LoadAssembly( $sAssemblyName, $oAppDomain = 0 )
	If Not $oAppDomain Then $oAppDomain = DotNet_GetDefaultDomain()

	Local $pType, $oType
	$oAppDomain.GetType( $pType )
	$oType = ObjCreateInterface( $pType, $sIID__Type, $sTag__Type )

	Local $pAssembly, $oAssembly
	$oType.get_Assembly( $pAssembly )
	$oAssembly = ObjCreateInterface( $pAssembly, $sIID__Assembly, $sTag__Assembly )

	Local $pAssemblyType, $oAssemblyType
	$oAssembly.GetType( $pAssemblyType )
	$oAssemblyType = ObjCreateInterface( $pAssemblyType, $sIID__Type, $sTag__Type )

	Local $aAssemblyName = [ $sAssemblyName ], $pSafeArray
	AccVars_ArrayToSafeArray( $aAssemblyName, $pSafeArray )

	Local $pNetCode
	$oAssemblyType.InvokeMember_3( "LoadFrom", $BindingFlags_DefaultValue, 0, 0, $pSafeArray, $pNetCode )
	; We first try to load the .NET assembly with the LoadFrom method

	; If LoadFrom method fails, we try to load the .NET assembly using the LoadWithPartialName method
	If Not Ptr( $pNetCode ) Then
		Local $iPos = StringInStr( $sAssemblyName, ".", Default, -1 )
		If StringRight( $sAssemblyName, StringLen( $sAssemblyName ) - $iPos ) = "dll" Then _ ; Use name of .NET assembly without DLL-
			$sAssemblyName = StringLeft( $sAssemblyName, $iPos - 1 )                           ; extension in LoadWithPartialName method.
		$aAssemblyName[0] = $sAssemblyName
		AccVars_ArrayToSafeArray( $aAssemblyName, $pSafeArray )
		$oAssemblyType.InvokeMember_3( "LoadWithPartialName", $BindingFlags_DefaultValue, 0, 0, $pSafeArray, $pNetCode )
		; LoadWithPartialName searches for the .NET assembly in the path given by $sBaseDirectory in DotNet_StartDomain
	EndIf

	If Not Ptr( $pNetCode ) Then Return SetError( 2,0,0 )
	Return ObjCreateInterface( $pNetCode, $sIID__Assembly, $sTag__Assembly )
EndFunc

Func DotNet_LoadCScode( $sCode, $sReferences = "", $oAppDomain = 0, $sFileName = "", $sCompilerOptions = "" )
	Local $oNetCode = DotNet_LoadCode( $sCode, $sReferences, "System", "Microsoft.CSharp.CSharpCodeProvider", $oAppDomain, $sFileName, $sCompilerOptions )
	If @error Then Return SetError( @error,0,0 )
	Return $oNetCode
EndFunc

Func DotNet_LoadVBcode( $sCode, $sReferences = "", $oAppDomain = 0, $sFileName = "", $sCompilerOptions = "" )
	Local $oNetCode = DotNet_LoadCode( $sCode, $sReferences, "System", "Microsoft.VisualBasic.VBCodeProvider", $oAppDomain, $sFileName, $sCompilerOptions )
	If @error Then Return SetError( @error,0,0 )
	Return $oNetCode
EndFunc

; Internal function
Func DotNet_LoadCode( $sCode, $sReferences, $sProviderAssembly, $sProviderType, $oAppDomain = 0, $sFileName = "", $sCompilerOptions = "" )
	If Not $oAppDomain Then $oAppDomain = DotNet_GetDefaultDomain()

	Local $oAsmProvider, $oCodeProvider, $oCodeCompiler, $oAsmSystem, $oPrms
	If IsObj( $oAppDomain ) Then $oAsmProvider = DotNet_LoadAssembly( $sProviderAssembly, $oAppDomain )
	If IsObj( $oAsmProvider ) Then $oAsmProvider.CreateInstance( $sProviderType, $oCodeProvider )
	If IsObj( $oCodeProvider ) Then $oCodeCompiler = $oCodeProvider.CreateCompiler()
	If IsObj( $oCodeCompiler ) Then $oAsmSystem = $sProviderAssembly = "System" ? $oAsmProvider : DotNet_LoadAssembly( "System", $oAppDomain )
	If Not IsObj( $oAsmSystem ) Then Return SetError( 3,0,0 )

	Local $aReferences = StringSplit( StringStripWS( StringRegExpReplace( $sReferences, "\s*\|\s*", "|" ), 3 ), "|", 2 ) ; 3 = $STR_STRIPLEADING + $STR_STRIPTRAILING, 2 = $STR_NOCOUNT
	Local $pSafeArray = AccVars_ArrayToSafeArrayOfVartype( $aReferences, $VT_BSTR )
	Local $psaVariant = AccVars_SafeArrayToSafeArrayOfVariant( $pSafeArray )
	Local $psaEmpty = SafeArrayCreateEmpty( $VT_VARIANT )

	; Create $oPrms object
	$oAsmSystem.CreateInstance_3( "System.CodeDom.Compiler.CompilerParameters", True, 0, 0, $psaVariant, 0, $psaEmpty, $oPrms )
	If Not IsObj( $oPrms ) Then Return SetError( 4,0,0 )

	; Set parameters for compiler
	$oPrms.OutputAssembly = $sFileName
	$oPrms.GenerateInMemory = ( $sFileName = "" )
	$oPrms.GenerateExecutable = ( StringRight( $sFileName, 4 ) = ".exe" )
	$oPrms.CompilerOptions = $sCompilerOptions
	$oPrms.IncludeDebugInformation = True

	; Compile code
	Local $oCompilerRes = $oCodeCompiler.CompileAssemblyFromSource( $oPrms, $sCode )

	; Compiler errors?
	If $oCompilerRes.Errors.Count() Then
		Local $b = False
		For $err In $oCompilerRes.Errors
			If $b Then ConsoleWrite( @CRLF )
			ConsoleWrite( "Line: " & $err.Line & @CRLF & "Column: " & $err.Column & @CRLF & $err.ErrorNumber & ": " & $err.ErrorText & @CRLF )
			$b = True
		Next
		Return SetError( 5,0,0 )
	EndIf

	If $sFileName Then
		Return $oCompilerRes.PathToAssembly()
	Else
		Local $pNetCode = $oCompilerRes.CompiledAssembly()
		Return ObjCreateInterface( $pNetCode, $sIID__Assembly, $sTag__Assembly )
	EndIf
EndFunc

Func DotNet_CreateObject( ByRef $oNetCode, $sClassName, $v3 = Default, $v4 = Default, $v5 = Default, $v6 = Default, $v7 = Default, $v8 = Default, $v9 = Default )
	Local $aParams = [ $v3, $v4, $v5, $v6, $v7, $v8, $v9 ], $oObject = 0

	If @NumParams = 2 Then
		$oNetCode.CreateInstance_2( $sClassName, True, $oObject )
		Return $oObject
	EndIf

	Local $iArgs = @NumParams - 2, $aArgs[$iArgs], $pSafeArray
	For $i = 0 To $iArgs - 1
		$aArgs[$i] = $aParams[$i]
	Next
	AccVars_ArrayToSafeArray( $aArgs, $pSafeArray )

	Local $psaEmpty = SafeArrayCreateEmpty( $VT_VARIANT )

	$oNetCode.CreateInstance_3( $sClassName, True, 0, 0, $pSafeArray, 0, $psaEmpty, $oObject )
	Return $oObject
EndFunc

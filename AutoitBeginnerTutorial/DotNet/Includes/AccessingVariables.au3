#include-once
#include "Variant.au3"
#include "SafeArray.au3"
#include "COMUtils.au3"

Global $oAccVars_Object = AccVars_Init(), $hAccVars_MethodFunc


; --- The first part of the UDF creates $oAccVars_Object and implements method functions ---

Func AccVars_Init()
	; Three locals copied from "IUIAutomation MS framework automate chrome, FF, IE, ...." by junkew
	; https://www.autoitscript.com/forum/index.php?showtopic=153520
	Local $sCLSID_CUIAutomation = "{FF48DBA4-60EF-4201-AA87-54103EEF594E}"
	Local $sIID_IUIAutomation = "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}"
	Local $stag_IUIAutomation = _
		"AccVars_VariableToVariant01 hresult(variant*);" & _
		"AccVars_VariableToVariant02 hresult(variant*;variant*);"

	; Create AccVars object (Automation object)
  Local $oAccVars_Object = ObjCreateInterface( $sCLSID_CUIAutomation, $sIID_IUIAutomation, $stag_IUIAutomation )
  If Not IsObj( $oAccVars_Object ) Then Return SetError(1,0,1)

	; Replace original methods with AccVars_VariableToVariantXY methods
	Local $pVariableToVariant, $pAccessVariable = Ptr( $oAccVars_Object() ), $sFuncName, $sFuncParams = "ptr", $iOffset, $iPtrSize = @AutoItX64 ? 8 : 4
	For $i = 1 To 2
		$sFuncName = "AccVars_VariableToVariant" & StringFormat( "%02i", $i )
		$sFuncParams &= ";ptr*"
		$iOffset = ( 3 + $i - 1 ) * $iPtrSize
		$pVariableToVariant = DllCallbackGetPtr( DllCallbackRegister( $sFuncName, "long", $sFuncParams ) )
		ReplaceVTableFuncPtr( $pAccessVariable, $iOffset, $pVariableToVariant )
	Next

	Return $oAccVars_Object
EndFunc

Func AccVars_VariableToVariant01( $pSelf, $pVariant01 )
	$hAccVars_MethodFunc( $pVariant01 )
	Return 0 ; $S_OK (COM constant)
	#forceref $pSelf
EndFunc

Func AccVars_VariableToVariant02( $pSelf, $pVariant01, $pVariant02 )
	$hAccVars_MethodFunc( $pVariant01, $pVariant02 )
	Return 0 ; $S_OK (COM constant)
	#forceref $pSelf
EndFunc


; --- The last part of the UDF creates a set of easy to use functions to call the object methods ---

Func AccessVariables01( $hAccVars_Method, ByRef $vVariable01 )
	$hAccVars_MethodFunc = $hAccVars_Method
	$oAccVars_Object.AccVars_VariableToVariant01( $vVariable01 )
EndFunc

Func AccessVariables02( $hAccVars_Method, ByRef $vVariable01, ByRef $vVariable02 )
	$hAccVars_MethodFunc = $hAccVars_Method
	$oAccVars_Object.AccVars_VariableToVariant02( $vVariable01, $vVariable02 )
EndFunc

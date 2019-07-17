#include-once
#include "AccessingVariables.au3"

; --- AccVars_ArrayToSafeArray ---

; The AutoIt array and the safearray are arrays of variants
Func AccVars_ArrayToSafeArray( ByRef $aArray, ByRef $pSafeArray )
	AccessVariables01( AccVars_ArrayToSafeArrayConvert, $aArray )
	AccVars_ArrayToSafeArrayData( $pSafeArray )
EndFunc

Func AccVars_ArrayToSafeArrayConvert( $pArray )
	; <<<< On function entry the native AutoIt array is converted to a safearray contained in a variant >>>>

	; Get safearray pointer
	Local $pSafeArrayFromArray = DllStructGetData( DllStructCreate( "ptr", $pArray + 8 ), 1 )

	Local $pSafeArray
	SafeArrayCopy( $pSafeArrayFromArray, $pSafeArray )
	If @error Then Return SetError(1,0,0)

	AccVars_ArrayToSafeArrayData( $pSafeArray, 1 )
EndFunc

Func AccVars_ArrayToSafeArrayData( ByRef $pSafeArray, $bSet = 0 )
	Static $pStaticSafeArray
	If $bSet Then
		$pStaticSafeArray = $pSafeArray
	Else
		$pSafeArray = $pStaticSafeArray
	EndIf
EndFunc


; --- AccVars_SafeArrayToArray ---

; The safearray and the AutoIt array are arrays of variants
Func AccVars_SafeArrayToArray( ByRef $pSafeArray, ByRef $aArray )
	AccessVariables02( AccVars_SafeArrayToArrayConvert, $pSafeArray, $aArray )
EndFunc

Func AccVars_SafeArrayToArrayConvert( $pvSafeArray, $pArray )

	; --- Get safearray information ---

	; $pvSafeArray is a variant that contains a pointer
	Local $pSafeArray = DllStructGetData( DllStructCreate( "ptr", $pvSafeArray + 8 ), 1 )

	; Array type
	Local $iVarType
	SafeArrayGetVartype( $pSafeArray, $iVarType )
	Switch $iVarType
		Case $VT_I2, $VT_I4   ; Signed integers
		Case $VT_R4, $VT_R8   ; 4/8 bytes floats
		Case $VT_BSTR         ; Basic string
		Case $VT_BOOL         ; Boolean type
		Case $VT_UI4, $VT_UI8 ; 4/8 bytes unsigned integers
		Case $VT_VARIANT      ; Variant data type
		Case $VT_UNKNOWN      ; IUnknown pointer
			; $pSafeArray is not compatible with a native AutoIt array
			; Convert $pSafeArray to a compatible safearray in $pSafeArray2
			Local $tSafeArray = DllStructCreate( $tagSAFEARRAY, $pSafeArray )
			Local $tSafeArrayBound = DllStructCreate( $tagSAFEARRAYBOUND )
			Local $iElements = DllStructGetData( $tSafeArray, "cElements" )
			DllStructSetData( $tSafeArrayBound, "cElements", $iElements )
			DllStructSetData( $tSafeArrayBound, "lLbound", DllStructGetData( $tSafeArray, "lLbound" ) )
			Local $pSafeArray2 = SafeArrayCreate( $VT_VARIANT, 1, $tSafeArrayBound ), $pArrayData, $pArrayData2
			SafeArrayAccessData( $pSafeArray2, $pArrayData2 )
			SafeArrayAccessData( $pSafeArray, $pArrayData )
			$iVarType = @AutoItX64 ? $VT_UI8 : $VT_UI4
			Local $iVarSize = @AutoItX64 ? 24 : 16
			Local $iPtrSize = @AutoItX64 ? 8 : 4
			For $i = 0 To $iElements - 1 ; Set variant type and data
				DllStructSetData( DllStructCreate( "word", $pArrayData2 + $i * $iVarSize ), 1, $iVarType )
				DllStructSetData( DllStructCreate( "uint_ptr", 8 + $pArrayData2 + $i * $iVarSize ), 1, DllStructGetData( DllStructCreate( "ptr", $pArrayData + $i * $iPtrSize ), 1 ) )
			Next
			SafeArrayUnaccessData( $pSafeArray )
			SafeArrayUnaccessData( $pSafeArray2 )
			SafeArrayCopy( $pSafeArray2, $pSafeArray )
			SafeArrayDestroy( $pSafeArray2 )
		Case Else
			Return SetError(1,0,0)
	EndSwitch

	; --- Set $pArray to match an array ---

	; Set vt element to $VT_ARRAY + $iVarType
	DllStructSetData( DllStructCreate( "word", $pArray ), 1, $VT_ARRAY + $iVarType )

	; Set data element to safearray pointer
	DllStructSetData( DllStructCreate( "ptr", $pArray + 8 ), 1, $pSafeArray )

	; <<<< On function exit the safearray contained in a variant is converted to a native AutoIt array >>>>
EndFunc


; --- AccVars_ArrayToSafeArrayOfVartype ---

; The returned safearray is an array of the specified variant type
Func AccVars_ArrayToSafeArrayOfVartype( ByRef $aArray, $iVartype )
	Local $tsaBound = DllStructCreate( $tagSAFEARRAYBOUND ), $pSafeArray, $pSafeArrayData, $iArray = UBound( $aArray ), $iPtrSize = @AutoItX64 ? 8 : 4
	DllStructSetData( $tsaBound, "cElements", $iArray )
	DllStructSetData( $tsaBound, "lLbound", 0 )
	$pSafeArray = SafeArrayCreate( $iVartype, 1, $tsaBound )
	SafeArrayAccessData( $pSafeArray, $pSafeArrayData )
	For $i = 0 To $iArray - 1
		DllStructSetData( DllStructCreate( "ptr", $pSafeArrayData + $iPtrSize * $i ), 1, SysAllocString( $aArray[$i] ) )
	Next
	SafeArrayUnaccessData( $pSafeArray )
	Return $pSafeArray
EndFunc


; --- AccVars_SafeArrayToSafeArrayOfVariant ---

; Returns a safearray with one variant element
; The variant element contains the input safearray
Func AccVars_SafeArrayToSafeArrayOfVariant( ByRef $pSafeArrayIn )
	Local $iVartype, $tsaBound = DllStructCreate( $tagSAFEARRAYBOUND ), $pSafeArrayOut, $pSafeArrayOutData
	SafeArrayGetVartype( $pSafeArrayIn, $iVartype )
	DllStructSetData( $tsaBound, "cElements", 1 )
	DllStructSetData( $tsaBound, "lLbound", 0 )
	$pSafeArrayOut = SafeArrayCreate( $VT_VARIANT, 1, $tsaBound )
	SafeArrayAccessData( $pSafeArrayOut, $pSafeArrayOutData )
	DllStructSetData( DllStructCreate( "word", $pSafeArrayOutData ), 1, $iVartype + $VT_ARRAY )
	DllStructSetData( DllStructCreate( "ptr", $pSafeArrayOutData + 8 ), 1, $pSafeArrayIn )
	SafeArrayUnaccessData( $pSafeArrayOut )
	Return $pSafeArrayOut
EndFunc


; --- AccVars_VariantToVariable ---

Func AccVars_VariantToVariable( ByRef $pVariant )
	Switch DllStructGetData( DllStructCreate( "word", $pVariant ), 1 )
		Case $VT_I4, $VT_I8 ; 4/8 bytes signed integer
			Return DllStructGetData( DllStructCreate( "int", $pVariant + 8 ), 1 )
		Case $VT_R8 ; 8 bytes double
			Return DllStructGetData( DllStructCreate( "double", $pVariant + 8 ), 1 )
		Case $VT_BSTR ; Basic string
			Return SysReadString( DllStructGetData( DllStructCreate( "ptr", $pVariant + 8 ), 1 ) )
		Case $VT_BOOL ; 2 bytes boolean
			Return DllStructGetData( DllStructCreate( "short", $pVariant + 8 ), 1 )
		Case $VT_UI4, $VT_UI8 ; 4/8 bytes unsigned integer
			Return DllStructGetData( DllStructCreate( "ptr", $pVariant + 8 ), 1 )
		Case Else
			Return SetError(1,0,0)
	EndSwitch
EndFunc


; --- AccVars_VariableToVariant ---

Func AccVars_VariableToVariant( $vVariable, ByRef $pVariant )
	Switch VarGetType( $vVariable )
		Case "Bool"
			DllStructSetData( DllStructCreate( "word", $pVariant ), 1, $VT_BOOL )
			DllStructSetData( DllStructCreate( "short", $pVariant + 8 ), 1, $vVariable )
		Case "Double"
			DllStructSetData( DllStructCreate( "word", $pVariant ), 1, $VT_R8 )
			DllStructSetData( DllStructCreate( "double", $pVariant + 8 ), 1, $vVariable )
		Case "Int32"
			DllStructSetData( DllStructCreate( "word", $pVariant ), 1, $VT_I4 )
			DllStructSetData( DllStructCreate( "int", $pVariant + 8 ), 1, $vVariable )
		Case "Ptr"
			DllStructSetData( DllStructCreate( "word", $pVariant ), 1, @AutoItX64 ? $VT_UI8 : $VT_UI4 )
			DllStructSetData( DllStructCreate( "ptr", $pVariant + 8 ), 1, $vVariable )
		Case "String"
			DllStructSetData( DllStructCreate( "word", $pVariant ), 1, $VT_BSTR )
			DllStructSetData( DllStructCreate( "ptr", $pVariant + 8 ), 1, SysAllocString( $vVariable ) )
		Case Else
			Return SetError(1,0,0)
	EndSwitch
EndFunc

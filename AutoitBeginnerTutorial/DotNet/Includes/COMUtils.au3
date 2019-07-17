#include-once

Func CoCreateInstance( $rclsid, $pUnkOuter, $ClsContext, $riid, ByRef $ppv )
	Local $aRet = DllCall( "ole32.dll", "long", "CoCreateInstance", "struct*", $rclsid, "ptr", $pUnkOuter, "dword", $ClsContext, "struct*", $riid, "ptr*", 0 )
	If @error Then Return SetError(1,0,1)
	$ppv = $aRet[5]
	Return $aRet[0]
EndFunc

; Add these lines to a script to activate the error handler:
;Local $oComErrFunc = ObjEvent( "AutoIt.Error", "ComErrFunc" )
;#forceref $oComErrFunc
Func ComErrFunc( $oError )
ConsoleWrite( @ScriptName & "(" & $oError.scriptline & "): ==> COM Error intercepted!" & @CRLF & _
	@TAB & "Err.number is: " & @TAB & @TAB & "0x" & Hex( $oError.number ) & @CRLF & _
	@TAB & "Err.windescription:" & @TAB & $oError.windescription & _
	@TAB & "Err.description is: " & @TAB & $oError.description & @CRLF & _
	@TAB & "Err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
	@TAB & "Err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
	@TAB & "Err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
	@TAB & "Err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
	@TAB & "Err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
	@TAB & "Err.retcode is: " & @TAB & "0x" & Hex( $oError.retcode ) & @CRLF )
EndFunc

; Copied from "Hooking into the IDispatch interface" by monoceres
; https://www.autoitscript.com/forum/index.php?showtopic=107678
Func ReplaceVTableFuncPtr( $pVTable, $iOffset, $pNewFunc )
	Local $pPointer = DllStructGetData( DllStructCreate( "ptr", $pVTable ), 1 ) + $iOffset, $PAGE_EXECUTE_READWRITE = 0x40
	Local $pOldFunc = DllStructGetData( DllStructCreate( "ptr", $pPointer ), 1 ) ; Get the original function pointer in VTable
	Local $aRet = DllCall( "Kernel32.dll", "int", "VirtualProtect", "ptr", $pPointer, "long", @AutoItX64 ? 8 : 4, "dword", $PAGE_EXECUTE_READWRITE, "dword*", 0 ) ; Unprotect memory
	DllStructSetData( DllStructCreate( "ptr", $pPointer ), 1, $pNewFunc ) ; Replace function pointer in VTable with $pNewFunc function pointer
	DllCall( "Kernel32.dll", "int", "VirtualProtect", "ptr", $pPointer, "long", @AutoItX64 ? 8 : 4, "dword", $aRet[4], "dword*", 0 ) ; Protect memory
	Return $pOldFunc ; Return original function pointer
EndFunc

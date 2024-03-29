#Include-once
#include "wd_core.au3"

#Region Copyright
#cs
	* WD_Helper.au3
	*
	* MIT License
	*
	* Copyright (c) 2018 Dan Pollak
	*
	* Permission is hereby granted, free of charge, to any person obtaining a copy
	* of this software and associated documentation files (the "Software"), to deal
	* in the Software without restriction, including without limitation the rights
	* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	* copies of the Software, and to permit persons to whom the Software is
	* furnished to do so, subject to the following conditions:
	*
	* The above copyright notice and this permission notice shall be included in all
	* copies or substantial portions of the Software.
	*
	* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	* SOFTWARE.
#ce
#EndRegion Copyright

#Region Many thanks to:
#cs
	- Jonathan Bennett and the AutoIt Team
	- Thorsten Willert, author of FF.au3, which I've used as a model
	- Michał Lipok for all his feedback / suggestions
#ce
#EndRegion Many thanks to:


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_NewTab
; Description ...: Helper function to create new tab using Javascript
; Syntax ........: _WD_NewTab($sSession[, $lSwitch = True[, $iTimeout = -1]])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $lSwitch             - [optional] Switch session context to new tab? Default is True.
;                  $iTimeout            - [optional] Period of time to wait before exiting function
;                  $sURL                - [optional] URL to be loaded in new tab
;                  $sFeatures           - [optional] Comma-separated list of requested features of the new tab
; Return values .: Success      - String representing handle of new tab
;                  Failure      - blank string
;                  @ERROR       - $_WD_ERROR_Success
;                  				- $_WD_ERROR_GeneralError
;                  				- $_WD_ERROR_Timeout
; Author ........: Dan Pollak
; Modified ......: 01/12/2019
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_NewTab($sSession, $lSwitch = True, $iTimeout = -1, $sURL = "", $sFeatures = "")
	Local Const $sFuncName = "_WD_NewTab"
	Local $sTabHandle = '', $sLastTabHandle, $hWaitTimer, $iTabIndex, $aTemp

	If $iTimeout = -1 Then $iTimeout = $_WD_DefaultTimeout

	Local $aHandles = _WD_Window($sSession, 'handles')

	If @error <> $_WD_ERROR_Success Or Not IsArray($aHandles) Then
		Return SetError(__WD_Error($sFuncName, $_WD_ERROR_Exception), 0, $sTabHandle)
	EndIf

	Local $iTabCount = UBound($aHandles)

	; Get handle to current last tab
	$sLastTabHandle = $aHandles[$iTabCount - 1]

	; Get handle for current tab
	Local $sCurrentTabHandle = _WD_Window($sSession, 'window')

	If @error = $_WD_ERROR_Success Then
		; Search for current tab handle in array of tab handles. If not found,
		; then make the current tab handle equal to the last tab
		$iTabIndex = _ArraySearch($aHandles, $sCurrentTabHandle)

		If @error Then
			$sCurrentTabHandle = $sLastTabHandle
			$iTabIndex = $iTabCount - 1
		EndIf
	Else
		_WD_Window($sSession, 'Switch', '{"handle":"' & $sLastTabHandle & '"}')
		$sCurrentTabHandle = $sLastTabHandle
		$iTabIndex = $iTabCount - 1
	EndIf

	_WD_ExecuteScript($sSession, "window.open(arguments[0], '', arguments[1])", '"' & $sURL & '","' & $sFeatures & '"')

	If @error <> $_WD_ERROR_Success Then
		Return SetError(__WD_Error($sFuncName, $_WD_ERROR_Exception), 0, $sTabHandle)
	EndIf

	$hWaitTimer = TimerInit()

	While 1
 		$aTemp = _WD_Window($sSession, 'handles')

		If UBound($aTemp) > $iTabCount Then
			$sTabHandle = $aTemp[$iTabIndex + 1]
			ExitLoop
		EndIf

		If TimerDiff($hWaitTimer) > $iTimeout Then Return SetError(__WD_Error($sFuncName, $_WD_ERROR_Timeout), 0, $sTabHandle)

		Sleep(10)
	WEnd

	If $lSwitch Then
		_WD_Window($sSession, 'Switch', '{"handle":"' & $sTabHandle & '"}')
	Else
		_WD_Window($sSession, 'Switch', '{"handle":"' & $sCurrentTabHandle & '"}')
	EndIf

	Return SetError($_WD_ERROR_Success, 0, $sTabHandle)
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_Attach
; Description ...: Helper function to attach to existing browser tab
; Syntax ........: _WD_Attach($sSession, $sString[, $sMode = 'title'])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sString             - String to search for
;                  $sMode               - [optional] One of the following search modes:
;                               | Title (Default)
;                               | URL
;                               | HTML
; Return values .: Success      - String representing handle of matching tab
;                  Failure      - blank string
;                  @ERROR       - $_WD_ERROR_Success
;                  				- $_WD_ERROR_InvalidDataType
;                  				- $_WD_ERROR_NoMatch
;                  				- $_WD_ERROR_GeneralError
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_Attach($sSession, $sString, $sMode = 'title')
	Local Const $sFuncName = "_WD_Attach"
	Local $sTabHandle = '', $lFound = False, $sCurrentTab = '', $aHandles

	$aHandles = _WD_Window($sSession, 'handles')

	If @error = $_WD_ERROR_Success Then
		$sMode = StringLower($sMode)
		$sCurrentTab = _WD_Window($sSession, 'window')

		For $sHandle In $aHandles

			_WD_Window($sSession, 'Switch', '{"handle":"' & $sHandle & '"}')

			Switch $sMode
				Case "title", "url"
					If StringInStr(_WD_Action($sSession, $sMode), $sString) > 0 Then
						$lFound = True
						$sTabHandle = $sHandle
						ExitLoop
					EndIf

				Case 'html'
					If StringInStr(_WD_GetSource($sSession), $sString) > 0 Then
						$lFound = True
						$sTabHandle = $sHandle
						ExitLoop
					EndIf

				Case Else
					Return SetError(__WD_Error($sFuncName, $_WD_ERROR_InvalidDataType, "(Title|URL|HTML) $sOption=>" & $sMode), 0, $sTabHandle)
			EndSwitch
		Next

		If Not $lFound Then
			; Restore prior active tab
			If $sCurrentTab <> '' Then
				_WD_Window($sSession, 'Switch', '{"handle":"' & $sCurrentTab & '"}')
			EndIf

			Return SetError(__WD_Error($sFuncName, $_WD_ERROR_NoMatch), 0, $sTabHandle)
		EndIf
	Else
		Return SetError(__WD_Error($sFuncName, $_WD_ERROR_GeneralError), 0, $sTabHandle)
	EndIf

	Return SetError($_WD_ERROR_Success, 0, $sTabHandle)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_LinkClickByText
; Description ...: Simulate a mouse click on a link with text matching the provided string
; Syntax ........: _WD_LinkClickByText($sSession, $sText[, $lPartial = True])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sText               - Text to find in link
;                  $lPartial            - [optional] Search by partial text? Default is True.
; Return values .: Success      - None
;                  Failure      - Sets @error to non-zero
;                  @ERROR       - $_WD_ERROR_Success
;                  				- $_WD_ERROR_Exception
;                  				- $_WD_ERROR_NoMatch
;                  @EXTENDED    - WinHTTP status code
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_LinkClickByText($sSession, $sText, $lPartial = True)
	Local Const $sFuncName = "_WD_LinkClickByText"

	Local $sElement = _WD_FindElement($sSession, ($lPartial) ? $_WD_LOCATOR_ByPartialLinkText : $_WD_LOCATOR_ByLinkText, $sText)

	Local $iErr = @error

	If $iErr = $_WD_ERROR_Success Then
		_WD_ElementAction($sSession, $sElement, 'click')
		$iErr = @error

		If $iErr <> $_WD_ERROR_Success Then
			Return SetError(__WD_Error($sFuncName, $_WD_ERROR_Exception), $_WD_HTTPRESULT)
		EndIf
	Else
		Return SetError(__WD_Error($sFuncName, $_WD_ERROR_NoMatch), $_WD_HTTPRESULT)
	EndIf

	Return SetError($_WD_ERROR_Success)
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_WaitElement
; Description ...: Wait for a element to be found  in the current tab before returning
; Syntax ........: _WD_WaitElement($sSession, $sStrategy, $sSelector[, $iDelay = 0[, $iTimeout = -1[, $lVisible = False]]])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sStrategy           - Locator strategy. See defined constant $_WD_LOCATOR_* for allowed values
;                  $sSelector           - Value to find
;                  $iDelay              - [optional] Milliseconds to wait before checking status
;                  $iTimeout            - [optional] Period of time to wait before exiting function
;                  $lVisible            - [optional] Check visibility of element?
; Return values .: Success      - 1
;                  Failure      - 0 and sets the @error flag to non-zero
;                  @error       - $_WD_ERROR_Success
;                  				- $_WD_ERROR_Timeout
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_WaitElement($sSession, $sStrategy, $sSelector, $iDelay = 0, $iTimeout = -1, $lVisible = False)
	Local Const $sFuncName = "_WD_WaitElement"
	Local $iErr, $iResult = 0, $sElement, $lIsVisible = True

	If $iTimeout = -1 Then $iTimeout = $_WD_DefaultTimeout

	Sleep($iDelay)

	Local $hWaitTimer = TimerInit()

	While 1
		$sElement = _WD_FindElement($sSession, $sStrategy, $sSelector)
		$iErr = @error

		If $iErr = $_WD_ERROR_Success Then
			If $lVisible Then
				$lIsVisible = _WD_ElementAction($sSession, $sElement, 'displayed')
			EndIf

			If $lIsVisible = True Then
				$iResult = 1
				ExitLoop
			EndIf

;~ 		ElseIf $iErr <> $_WD_ERROR_NoMatch Then
;~ 			ExitLoop
		EndIf

		If (TimerDiff($hWaitTimer) > $iTimeout) Then
			$iErr = $_WD_ERROR_Timeout
			ExitLoop
		EndIf

		Sleep(1000)
	WEnd

	Return SetError(__WD_Error($sFuncName, $iErr), 0, $iResult)
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_GetMouseElement
; Description ...: Retrieves reference to element below mouse pointer
; Syntax ........: _WD_GetMouseElement($sSession)
; Parameters ....: $sSession            - Session ID from _WDCreateSession
; Return values .: Element ID returned by web driver
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://stackoverflow.com/questions/24538450/get-element-currently-under-mouse-without-using-mouse-events
; Example .......: No
; ===============================================================================================================================
Func _WD_GetMouseElement($sSession)
	Local Const $sFuncName = "_WD_GetMouseElement"
	Local $sResponse, $sJSON, $sElement
	Local $sScript = "return Array.from(document.querySelectorAll(':hover')).pop()"

	$sResponse = _WD_ExecuteScript($sSession, $sScript, '')
	$sJSON = Json_Decode($sResponse)
	$sElement = Json_Get($sJSON, "[value][" & $_WD_ELEMENT_ID & "]")

	If $_WD_DEBUG = $_WD_DEBUG_Info Then
		ConsoleWrite($sFuncName & ': ' & $sElement & @CRLF)
		ConsoleWrite($sFuncName & ': ' & IsObj($sElement) & @CRLF)
	EndIf

Return SetError($_WD_ERROR_Success, 0, $sElement)
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_GetElementFromPoint
; Description ...:
; Syntax ........: _WD_GetElementFromPoint($sSession, $iX, $iY)
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $iX                  - an integer value.
;                  $iY                  - an integer value.
; Return values .: Success      - Element ID returned by web driver
;                  Failure      - blank string
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_GetElementFromPoint($sSession, $iX, $iY)
	Local $sResponse, $sElement, $sJSON
    Local $sScript = "return document.elementFromPoint(arguments[0], arguments[1]);"
	Local $sParams = $iX & ", " & $iY

	$sResponse = _WD_ExecuteScript($sSession, $sScript, $sParams)
	$sJSON = Json_Decode($sResponse)
	$sElement = Json_Get($sJSON, "[value][" & $_WD_ELEMENT_ID & "]")

	Return SetError($_WD_ERROR_Success, 0, $sElement)
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_LastHTTPResult
; Description ...: Return the result of the last WinHTTP request
; Syntax ........: _WD_LastHTTPResult()
; Parameters ....: None
; Return values .: Result of last WinHTTP request
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_LastHTTPResult()
	Return $_WD_HTTPRESULT
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_GetFrameCount
; Description ...: This will return how many frames/iframes are in your current window/frame. It will not traverse to nested frames.
; Syntax ........: _WD_GetFrameCount()
; Parameters ....:
; Return values .: Success      - Numeric count of frames, 0 or positive number
;                  Failure      - ""
; Author ........: Decibel
; Modified ......: 2018-04-27
; Remarks .......:
; Related .......:
; Link ..........: https://www.w3schools.com/jsref/prop_win_length.asp
; Example .......: No
; ===============================================================================================================================
Func _WD_GetFrameCount($sSession)
    Local $sResponse, $sJSON, $iValue

    $sResponse = _WD_ExecuteScript($sSession, "return window.frames.length")
    $sJSON = Json_Decode($sResponse)
    $iValue = Json_Get($sJSON, "[value]")

    Return Number($iValue)
EndFunc ;==>_WD_GetFrameCount


; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_IsWindowTop
; Description ...: This will return a boolean of the session being at the top level, or in a frame(s).
; Syntax ........: _WD_IsWindowTop()
; Parameters ....:
; Return values .: Success      - Boolean response
;                  Failure      - ""
; Author ........: Decibel
; Modified ......: 2018-04-27
; Remarks .......:
; Related .......:
; Link ..........: https://www.w3schools.com/jsref/prop_win_top.asp
; Example .......: No
; ===============================================================================================================================
Func _WD_IsWindowTop($sSession)
    Local $sResponse, $sJSON
    Local $blnResult

    $sResponse = _WD_ExecuteScript($sSession, "return window.top == window.self")
    $sJSON = Json_Decode($sResponse)
    $blnResult = Json_Get($sJSON, "[value]")

    Return $blnResult
EndFunc ;==>_WD_IsWindowTop

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_FrameEnter
; Description ...: This will enter the specified frame for subsequent WebDriver operations.
; Syntax ........: _WD_FrameEnter($sSession, $sIndexOrID)
; Parameters ....:
; Return values .: Success      - True
;                  Failure      - WD Response error message (E.g. "no such frame")
; Author ........: Decibel
; Modified ......: 2018-04-27
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_FrameEnter($sSession, $sIndexOrID)
    Local $sOption
    Local $sResponse, $sJSON
    Local $sValue

    ;*** Encapsulate the value if it's an integer, assuming that it's supposed to be an Index, not ID attrib value.
    If IsInt($sIndexOrID) = True Then
        $sOption = '{"id":' & $sIndexOrID & '}'
    Else
		$sOption = '{"id":{"' & $_WD_ELEMENT_ID & '":"' & $sIndexOrID & '"}}'
    EndIf

    $sResponse = _WD_Window($sSession, "frame", $sOption)
    $sJSON = Json_Decode($sResponse)
    $sValue = Json_Get($sJSON, "[value]")

    ;*** Evaluate the response
    If $sValue <> Null Then
        $sValue = Json_Get($sJSON, "[value][error]")
    Else
        $sValue = True
    EndIf

    Return $sValue
EndFunc ;==>_WD_FrameEnter

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_FrameLeave
; Description ...: This will leave the current frame, to its parent, not necessarily the Top, for subsequent WebDriver operations.
; Syntax ........: _WD_FrameLeave()
; Parameters ....:
; Return values .: Success      True
;                  Failure      - WD Response error message (E.g. "chrome not reachable")
; Author ........: Decibel
; Modified ......: 2018-04-27
; Remarks .......: ChromeDriver and GeckoDriver respond differently for a successful operation
; Related .......:
; Link ..........: https://www.w3.org/TR/webdriver/#switch-to-parent-frame
; Example .......: No
; ===============================================================================================================================
Func _WD_FrameLeave($sSession)
    Local $sOption
    Local $sResponse, $sJSON, $asJSON
    Local $sValue

    $sOption = '{}'

    $sResponse = _WD_Window($sSession, "parent", $sOption)
    ;Chrome--
    ;   Good: '{"value":null}'
    ;   Bad: '{"value":{"error":"chrome not reachable"....
    ;Firefox--
    ;   Good: '{"value": {}}'
    ;   Bad: '{"value":{"error":"unknown error","message":"Failed to decode response from marionette","stacktrace":""}}'

    $sJSON = Json_Decode($sResponse)
    $sValue = Json_Get($sJSON, "[value]")

    ;*** Is this something besides a Chrome PASS?
    If $sValue <> Null Then
        ;*** Check for a nested JSON object
        If Json_IsObject($sValue) = True Then
            $asJSON = Json_ObjGetKeys($sValue)

            ;*** Is this an empty nested object
            If UBound($asJSON) = 0 Then ;Firefox PASS
                $sValue = True
            Else ;Chrome and Firefox FAIL
                $sValue = $asJSON[0] & ":" & Json_Get($sJSON, "[value][" & $asJSON[0] & "]")
            EndIf
        EndIf
    Else ;Chrome PASS
        $sValue = True
    EndIf

    Return $sValue
EndFunc ;==>_WD_FrameLeave

; #FUNCTION# ===========================================================================================================
; Name ..........: _WD_HighlightElement
; Description ...:
; Syntax ........: _WD_HighlightElement($sSession, $sElement[, $iMethod = 1])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sElement            - Element ID from _WDFindElement
;                  $iMethod             - [optional] an integer value. Default is 1.
;                  1=style -> Highlight border dotted red
;                  2=style -> Highlight yellow rounded box
;                  3=style -> Highlight yellow rounded box + border  dotted red
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Danyfirex
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/192730-webdriver-udf-help-support/?do=findComment&comment=1396643
; Example .......: No
; ===============================================================================================================================
Func _WD_HighlightElement($sSession, $sElement, $iMethod = 1)
    Local Const $aMethod[] = ["border: 2px dotted red", _
            "background: #FFFF66; border-radius: 5px; padding-left: 3px;", _
            "border:2px dotted  red;background: #FFFF66; border-radius: 5px; padding-left: 3px;"]
    If $iMethod < 1 Or $iMethod > 3 Then $iMethod = 1
	Local $sJsonElement = '{"' & $_WD_ELEMENT_ID & '":"' & $sElement & '"}'
    Local $sResponse = _WD_ExecuteScript($sSession, "arguments[0].style='" & $aMethod[$iMethod - 1] & "'; return true;", $sJsonElement)
    Local $sJSON = Json_Decode($sResponse)
    Local $sResult = Json_Get($sJSON, "[value]")
    Return ($sResult = "true" ? SetError(0, 0, $sResult) : SetError(1, 0, False))
EndFunc   ;==>_WD_HighlightElement

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_HighlightElements
; Description ...:
; Syntax ........: _WD_HighlightElements($sSession, $aElements[, $iMethod = 1])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $aElements           - an array of Elements ID from _WDFindElement
;                  $iMethod             - [optional] an integer value. Default is 1.
;                  1=style -> Highlight border dotted red
;                  2=style -> Highlight yellow rounded box
;                  3=style -> Highlight yellow rounded box + border  dotted red
; Return values .: Success      - True
;                  Failure      - False
;                  @Extended Number of Highlighted Elements
; Author ........: Danyfirex
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/192730-webdriver-udf-help-support/?do=findComment&comment=1396643
; Example .......: No
; ===============================================================================================================================
Func _WD_HighlightElements($sSession, $aElements, $iMethod = 1)
    Local $iHighlightedElements = 0
    For $i = 0 To UBound($aElements) - 1
        $iHighlightedElements += (_WD_HighlightElement($sSession, $aElements[$i], $iMethod) = True ? 1 : 0)
    Next
    Return ($iHighlightedElements > 0 ? SetError(0, $iHighlightedElements, True) : SetError(1, 0, False))
EndFunc   ;==>_WD_HighlightElements

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_LoadWait
; Description ...: Wait for a browser page load to complete before returning
; Syntax ........: _WD_LoadWait($sSession[, $iDelay = 0[, $iTimeout = -1[, $sElement = '']]])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $iDelay              - [optional] Milliseconds to wait before checking status
;                  $iTimeout            - [optional] Period of time to wait before exiting function
;                  $sElement            - [optional] Element ID to confirm DOM invalidation
; Return values .: Success      - 1
;                  Failure      - 0 and sets the @error flag to non-zero
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_LoadWait($sSession, $iDelay = 0, $iTimeout = -1, $sElement = '')
	Local Const $sFuncName = "_WD_LoadWait"
	Local $iErr, $sResponse, $sJSON, $sReadyState

	If $iTimeout = -1 Then $iTimeout = $_WD_DefaultTimeout

	If $iDelay Then Sleep($iDelay)

	Local $hLoadWaitTimer = TimerInit()

	While True
		If $sElement <> '' Then
			_WD_ElementAction($sSession, $sElement, 'name')

			If $_WD_HTTPRESULT = $HTTP_STATUS_NOT_FOUND Then $sElement = ''
		Else
			$sResponse = _WD_ExecuteScript($sSession, 'return document.readyState', '')
			$iErr = @error

			If $iErr Then
				ExitLoop
			EndIf

			$sJSON = Json_Decode($sResponse)
			$sReadyState = Json_Get($sJSON, "[value]")

			If $sReadyState = 'complete' Then ExitLoop
		EndIf

		If (TimerDiff($hLoadWaitTimer) > $iTimeout) Then
			$iErr = $_WD_ERROR_Timeout
			ExitLoop
		EndIf

		Sleep(100)
	WEnd

	If $iErr Then
		Return SetError(__WD_Error($sFuncName, $iErr, ""), 0, 0)
	EndIf

	Return SetError($_WD_ERROR_Success, 0, 1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_Screenshot
; Description ...:
; Syntax ........: _WD_Screenshot($sSession[, $sElement = ''[, $nOutputType = 1]])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sElement            - [optional] Element ID from _WDFindElement
;                  $nOutputType         - [optional] One of the following output types:
;                               | 1 - String (Default)
;                               | 2 - Binary
;                               | 3 - Base64

; Return values .: None
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_Screenshot($sSession, $sElement = '', $nOutputType = 1)
	Local Const $sFuncName = "_WD_Screenshot"
	Local $sResponse, $sResult, $iErr

	If $sElement = '' Then
		$sResponse = _WD_Window($sSession, 'Screenshot')
		$iErr = @error
	Else
		$sResponse = _WD_ElementAction($sSession, $sElement, 'Screenshot')
		$iErr = @error
	EndIf

	If $iErr = $_WD_ERROR_Success Then
		Switch $nOutputType
			Case 1 ; String
				$sResult = BinaryToString(_Base64Decode($sResponse))

			Case 2 ; Binary
				$sResult = _Base64Decode($sResponse)

			Case 3 ; Base64

		EndSwitch
	Else
		$sResult = ''
	EndIf

	Return SetError(__WD_Error($sFuncName, $iErr), 0, $sResult)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_jQuerify
; Description ...: Inject jQuery library into current session
; Syntax ........: _WD_jQuerify($sSession)
; Parameters ....: $sSession            - Session ID from _WDCreateSession
; Return values .: None
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://sqa.stackexchange.com/questions/2921/webdriver-can-i-inject-a-jquery-script-for-a-page-that-isnt-using-jquery
; Example .......: No
; ===============================================================================================================================
Func _WD_jQuerify($sSession)
Local $jQueryLoader = _
"(function(jqueryUrl, callback) {" & _
"    if (typeof jqueryUrl != 'string') {" & _
"        jqueryUrl = 'https://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js';" & _
"    }" & _
"    if (typeof jQuery == 'undefined') {" & _
"        var script = document.createElement('script');" & _
"        var head = document.getElementsByTagName('head')[0];" & _
"        var done = false;" & _
"        script.onload = script.onreadystatechange = (function() {" & _
"            if (!done && (!this.readyState || this.readyState == 'loaded' " & _
"                    || this.readyState == 'complete')) {" & _
"                done = true;" & _
"                script.onload = script.onreadystatechange = null;" & _
"                head.removeChild(script);" & _
"                callback();" & _
"            }" & _
"        });" & _
"        script.src = jqueryUrl;" & _
"        head.appendChild(script);" & _
"    }" & _
"    else {" & _
"        jQuery.noConflict();" & _
"        callback();" & _
"    }" & _
"})(arguments[0], arguments[arguments.length - 1]);"

_WD_ExecuteScript($sSession, $jQueryLoader, "[]", True)

Do
	Sleep(250)
	_WD_ExecuteScript($sSession, "jQuery")
Until @error = $_WD_ERROR_Success

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_ElementOptionSelect
; Description ...: Find and click on an option from a Select element
; Syntax ........: _WD_ElementOptionSelect($sSession, $sStrategy, $sSelector[, $sStartElement = ""])
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sStrategy           - Locator strategy. See defined constant $_WD_LOCATOR_* for allowed values
;                  $sSelector           - Value to find
;                  $sStartElement       - [optional] Element ID of element to use as starting point
; Return values .: None
;                  @ERROR       - $_WD_ERROR_Success
;                  				- $_WD_ERROR_NoMatch
;                  @EXTENDED    - WinHTTP status code
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_ElementOptionSelect($sSession, $sStrategy, $sSelector, $sStartElement = "")
    Local $sElement = _WD_FindElement($sSession, $sStrategy, $sSelector, $sStartElement)

    If @error = $_WD_ERROR_Success Then
        _WD_ElementAction($sSession, $sElement, 'click')
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_ConsoleVisible
; Description ...: Control visibility of the webdriver console app
; Syntax ........: _WD_ConsoleVisible([$lVisible = False])
; Parameters ....: $lVisible            - [optional] Set to true to hide the console
; Return values .: None
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_ConsoleVisible($lVisible = False)
	Local $sFile = StringRegExpReplace($_WD_DRIVER, "^.*\\(.*)$", "$1")
	Local $pid, $pid2, $hWnd = 0, $aWinList

	$pid = ProcessExists($sFile)

	If $pid Then
		$aWinList=WinList("[CLASS:ConsoleWindowClass]")

		For $i=1 To $aWinList[0][0]
			$pid2 = WinGetProcess($aWinList[$i][1])

			If $pid2 = $pid Then
				$hWnd=$aWinList[$i][1]
				ExitLoop
			EndIf
		Next

		If $hWnd<>0 Then
			WinSetState($hWnd, "", $lVisible ? @SW_SHOW : @SW_HIDE)
		EndIf
	EndIf

EndFunc   ;==>_WD_ConsoleVisible

; #FUNCTION# ====================================================================================================================
; Name ..........: _WD_GetShadowRoot
; Description ...:
; Syntax ........: _WD_GetShadowRoot($sSession, $sStrategy, $sSelector, $sStartElement = "")
; Parameters ....: $sSession            - Session ID from _WDCreateSession
;                  $sStrategy           - Locator strategy. See defined constant $_WD_LOCATOR_* for allowed values
;                  $sSelector           - Value to find
;                  $sStartElement       - [optional] a string value. Default is "".
; Return values .: Success      - Element ID returned by web driver
;                  Failure      - ""
;                  @ERROR       - $_WD_ERROR_Success
;                  				- $_WD_ERROR_Exception
;                  				- $_WD_ERROR_NoMatch
;                  @EXTENDED    - WinHTTP status code
; Author ........: Dan Pollak
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WD_GetShadowRoot($sSession, $sStrategy, $sSelector, $sStartElement = "")
	Local Const $sFuncName = "_WD_GetShadowRoot"

	Local $sResponse, $sResult, $sJsonElement, $oJson
	Local $sElement = _WD_FindElement($sSession, $sStrategy, $sSelector, $sStartElement)
	Local $iErr = @error

	If $iErr = $_WD_ERROR_Success Then
		$sJsonElement = '{"' & $_WD_ELEMENT_ID & '":"' & $sElement & '"}'
		$sResponse = _WD_ExecuteScript($sSession, "return arguments[0].shadowRoot", $sJsonElement)
		$oJson = Json_Decode($sResponse)
		$sResult  = Json_Get($oJson, "[value][" & $_WD_ELEMENT_ID & "]")
    EndIf

	If $_WD_DEBUG = $_WD_DEBUG_Info Then
		ConsoleWrite($sFuncName & ': ' & $sResult & @CRLF)
	EndIf

	Return SetError(__WD_Error($sFuncName, $iErr), $_WD_HTTPRESULT, $sResult)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Base64Decode
; Description ...:
; Syntax ........: _Base64Decode($input_string)
; Parameters ....: $input_string        - string to be decoded
; Return values .: Decoded string
; Author ........: trancexx
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/81332-_base64encode-_base64decode/
; Example .......: No
; ===============================================================================================================================
Func _Base64Decode($input_string)

    Local $struct = DllStructCreate("int")

    Local $a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", _
            "str", $input_string, _
            "int", 0, _
            "int", 1, _
            "ptr", 0, _
            "ptr", DllStructGetPtr($struct, 1), _
            "ptr", 0, _
            "ptr", 0)

    If @error Or Not $a_Call[0] Then
        Return SetError(1, 0, "") ; error calculating the length of the buffer needed
    EndIf

    Local $a = DllStructCreate("byte[" & DllStructGetData($struct, 1) & "]")

    $a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", _
            "str", $input_string, _
            "int", 0, _
            "int", 1, _
            "ptr", DllStructGetPtr($a), _
            "ptr", DllStructGetPtr($struct, 1), _
            "ptr", 0, _
            "ptr", 0)

    If @error Or Not $a_Call[0] Then
        Return SetError(2, 0, ""); error decoding
    EndIf

    Return DllStructGetData($a, 1)

EndFunc   ;==>_Base64Decode

#include-once
#include <Array.au3>
#include <GuiTab.au3>
#include <WinAPI.au3>
#include <GuiStatusBar.au3>
#include <ColorConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <DateTimeConstants.au3>
#include <ListViewConstants.au3>

; -------------------------------------------------------------------------------------------------

Func _KODAParser_DoFile($sFileOrXML, $iWidth, $iHeight)
	; returned object that will contain control IDs
	Local $oForm = _objCreate()

	; XML DOM object
	Local $oXML = ObjCreate("Microsoft.XMLDOM")
	$oXML.Async = False

	; load XML
	If FileExists($sFileOrXML) Then
		Local $hF = FileOpen($sFileOrXML, 512)
		$sFileOrXML = FileRead($hF)
		FileClose($hF)
	EndIf
	If Not $oXML.LoadXML($sFileOrXML) Then Return SetError(1, 0, Null)

	; create GUI
	__KODAParser_createGUI($oForm, $oXML, $iWidth, $iHeight)
	If @error Then Return SetError(1 + @error, 0, Null)

	; create controls
	Local $oControls = $oXML.selectNodes("/object/components/object")
	__KODAParser_createControls($oForm, $oControls)

	; clean and return
	_objDel($oForm, "___###gui_properties###___")
	$oControls = 0
	$oXML = 0

	Return $oForm
EndFunc

Func __KODAParser_createGUI($oForm, $oXML, $iWidth, $iHeight)
	Local $oObject = $oXML.selectSingleNode("/object")
	If $oObject.getAttribute("type") <> "TAForm" Then Return SetError(1, 0, Null)

	Local $oProperties = __KODAParser_readObjectProperties($oObject)

	Switch $oProperties.Item("Position")
		Case "poDesktopCenter"
			$oProperties.Item("Left") = -1
			$oProperties.Item("Top") = -1
		;TODO: $poFixed?
	EndSwitch

	; temporary store GUI properties object (could be used by control creation)
	; deleted in _KODAParser_DoFile befor returning
	_objSet($oForm, "___###gui_properties###___", $oProperties)

	; ajust window size (tooo headache!)
;~ 	$oProperties.Item("Width") = $oProperties.Item("Width") - (_WinAPI_GetSystemMetrics($SM_CXSIZEFRAME) * 2)
;~ 	$oProperties.Item("Height") = $oProperties.Item("Height") - (_WinAPI_GetSystemMetrics($SM_CYSIZEFRAME) * 2); caption, menu, status bar?

	Local $hGUI = GUICreate( _
		$oProperties.Item("Caption"), _
		$iWidth, $iHeight, _ ; $oProperties.Item("Width"), $oProperties.Item("Height"), _
		$oProperties.Item("Left"), $oProperties.Item("Top"), _
		$oProperties.Item("Style"), $oProperties.Item("ExStyle"), _
		Eval($oProperties.Item("ParentForm")) _ ;TODO: test
	)

	;TODO: move/resize GUI

	; font
	Local $aFont = __KODAParser_processFont($oForm, $oProperties)
	GUISetFont($aFont[0], $aFont[1], $aFont[2], $aFont[3], $hGUI)

	; color
	Local $iColor = __KODAParser_identifiers_colors(_objGet($oProperties, "Color", ""))
	If $iColor <> "" Then GUISetBkColor($iColor, $hGUI)

	; cursor
	GUISetCursor(__KODAParser_identifiers_cursor($oProperties.Item("Cursor")), 0, $hGUI)

	; store GUI handle
	_objSet($oForm, "#hwnd#", $hGUI)
	_objSet($oForm, $oObject.getAttribute("name"), $hGUI)

	; clean memory (really useful?)
	$oProperties = 0
	$oObject = 0
EndFunc

Func __KODAParser_createControls($oForm, $oObjects, $iXOffset = 0, $iYOffset = 0, $vUserData = Null)
	Local $iCtrlID, $oObject, $oProperties

	; we first sort the objects/controls to create in order to respect TabOrder
	; also, this function will make sure that any PopupMenu declaration comes after it's parent control
	Local $aObjects = __KODAParser_sortObjects($oObjects)

	For $iObjID = 0 To UBound($aObjects) - 1
		$oObject = $aObjects[$iObjID][0]
		$oProperties = $aObjects[$iObjID][1]

		If _objExists($oProperties, "Left") Then _objSet($oProperties, "Left", _objGet($oProperties, "Left") + $iXOffset)
		If _objExists($oProperties, "Top") Then _objSet($oProperties, "Top", _objGet($oProperties, "Top") + $iYOffset)

		Switch $oObject.getAttribute("type")
			; ---
;~ 			Case "TAMenu" ; main GUI menu
			Case "TMainMenu"
				Local $oComponents = $oObject.selectNodes("components/object")
				__KODAParser_createControls($oForm, $oComponents, 0, 0, -1)
				$oComponents = 0
				ContinueLoop
			; ---
			Case "TAContextMenu"
				; after sortin, we are sure that the 'Associate' control exists because menus are created the last
				$iCtrlID = GUICtrlCreateContextMenu(_objGet($oForm, _objGet($oProperties, "Associate"), -1))

				Local $oPopupMenu = $oObject.selectNodes("//object[@type='TPopupMenu' and @name='" & _objGet($oProperties, "WrappedName") & "']")
				__KODAParser_createControls($oForm, $oPopupMenu, 0, 0, $iCtrlID)
				$oPopupMenu = 0
			; ---
			Case "TPopupMenu"
				If Not $vUserData Then ContinueLoop

				Local $oComponents = $oObject.selectNodes("components/object")
				__KODAParser_createControls($oForm, $oComponents, 0, 0, $vUserData)
				$oComponents = 0
				ContinueLoop
			; ---
			Case "TAMenuItem"
				Local $oSubMenuItems = $oObject.selectNodes("components/object")
				If $oSubMenuItems.length > 0 Then
					$iCtrlID = GUICtrlCreateMenu( _
						_objGet($oProperties, "Caption", ""), _
						$vUserData _
					)
					__KODAParser_createControls($oForm, $oSubMenuItems, 0, 0, $iCtrlID)
				Else
					$iCtrlID = GUICtrlCreateMenuItem( _
						_objGet($oProperties, "Caption", ""), _
						$vUserData, _
						-1, _
						_objGet($oProperties, "RadioItem", False) ? 1 : 0 _
					)
				EndIf
				$oSubMenuItems = 0
			; ---
			Case "TALabel"
				$iCtrlID = GUICtrlCreateLabel( _
					$oProperties.Item("Caption"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TAButton"
				$iCtrlID = GUICtrlCreateButton( _
					$oProperties.Item("Caption"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				; do not change background color if default (to not fall in windows xp style)
				If $oProperties.Item("Color") = "clBtnFace" Then _objDel($oProperties, "Color")
			; ---
			Case "TAInput"
				$iCtrlID = GUICtrlCreateInput( _
					$oProperties.Item("Text"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TAEdit"
				$aLines = $oProperties.Item("Lines.Strings")
				$iCtrlID = GUICtrlCreateEdit( _
					IsArray($aLines) ? _ArrayToString($aLines, @CRLF, 1) : "", _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TACheckbox"
				$iCtrlID = GUICtrlCreateCheckbox( _
					$oProperties.Item("Caption"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TARadio"
				$iCtrlID = GUICtrlCreateRadio( _
					$oProperties.Item("Caption"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TAList"
				$iCtrlID = GUICtrlCreateList("", _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				$aItems = _objGet($oProperties, "Items.Strings", Null)
				If IsArray($aItems) Then GUICtrlSetData($iCtrlID, _ArrayToString($aItems, "|", 1))
			; ---
			Case "TACombo"
				If _objGet($oProperties, "ItemIndex", -1) > 0 Then
					_objSet($oProperties, "Text", "")
				EndIf

				$iCtrlID = GUICtrlCreateCombo( _
					$oProperties.Item("Text"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				$aItems = _objGet($oProperties, "Items.Strings", Null)
				$iItemIndex = _objGet($oProperties, "ItemIndex", -1)
				If IsArray($aItems) Then GUICtrlSetData($iCtrlID, _
					_ArrayToString($aItems, "|", 1), _
					$iItemIndex > 0 ? $aItems[$iItemIndex] : "" _
				)
			; ---
			Case "TAGroup"
				$iCtrlID = GUICtrlCreateGroup( _
					$oProperties.Item("Caption"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				Local $oComponents = $oObject.selectNodes("components/object")
				__KODAParser_createControls($oForm, $oComponents, $oProperties.Item("Left"), $oProperties.Item("Top"))
				$oComponents = 0

				GUICtrlCreateGroup("", -99, -99, 1, 1)
			; ---
			Case "TAPic"
				$iCtrlID = GUICtrlCreatePic( _
					$oProperties.Item("PicturePath"), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TAIcon" ;TODO: display bug (peut être rafraichir la fenêtre?)
				$iCtrlID = GUICtrlCreateIcon( _
					_objGet($oProperties, "CurstomPath", $oProperties.Item("PicturePath")), -1 * ($oProperties.Item("PictureIndex") + 1), _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TADummy"
				$iCtrlID = GUICtrlCreateDummy()
			; ---
			Case "TAControlGroup"
				GUIStartGroup()
				Local $oComponents = $oObject.selectNodes("components/object")
				__KODAParser_createControls($oForm, $oComponents, $oProperties.Item("Left"), $oProperties.Item("Top"))
				$oComponents = 0
				GUIStartGroup()
			; ---
			Case "TASlider"
				$iCtrlID = GUICtrlCreateSlider( _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				GUICtrlSetLimit($iCtrlID, _objGet($oProperties, "Max", 100), _objGet($oProperties, "Min", 0))
				GUICtrlSetData($iCtrlID, _objGet($oProperties, "Position", _objGet($oProperties, "Min", 0)))
			; ---
			Case "TAProgress"
				$iCtrlID = GUICtrlCreateProgress( _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

;~ 				GUICtrlSetLimit($iCtrlID, _objGet($oProperties, "Max", 100), _objGet($oProperties, "Min", 0)) ; limits are ignored by Progress control
				GUICtrlSetData($iCtrlID, _objGet($oProperties, "Position", _objGet($oProperties, "Min", 0))) ; 0-100
			; ---
			Case "TADate"
				Local $sText = StringSplit($oProperties.Item("Date"), "/")
				$sText = $sText[3] & "/" & $sText[2] & "/" & $sText[1]
				If BitAND($oProperties.Item("CtrlStyle"), $DTS_TIMEFORMAT) = $DTS_TIMEFORMAT Then $sText = $oProperties.Item("Time")

				$iCtrlID = GUICtrlCreateDate( _
					$sText, _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				If _objGet($oProperties, "Format") Then GUICtrlSendMsg($iCtrlID, $DTM_SETFORMATW, 0, _objGet($oProperties, "Format"))
			; ---
			Case "TAMonthCal"
				Local $sText = StringSplit($oProperties.Item("Date"), "/")
				$sText = $sText[3] & "/" & $sText[2] & "/" & $sText[1]

				Local $iCtrlID = GUICtrlCreateMonthCal( _
					$sText, _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
			; ---
			Case "TATreeView"
				$iCtrlID = GUICtrlCreateTreeView( _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				; treeView items are stored in an (opaque?) binary format
			; ---
			Case "TAListView"
				Local $sText = ""
				Local $aColumns = _objGet($oProperties, "Columns", Null)
				If IsArray($aColumns) Then
					For $i = 1 To $aColumns[0]
						$sText &= _objGet($aColumns[$i], "Caption", "") & "|"
					Next
					$sText = StringTrimRight($sText, 1)
				EndIf

				$iCtrlID = GUICtrlCreateListView( _
					$sText, _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)

				For $i = 1 To $aColumns[0]
					GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, $i - 1, _objGet($aColumns[$i], "Width", 50))
				Next

				; listView items are stored in an (opaque?) binary format
			; ---
			Case "TATab"
				$iCtrlID = GUICtrlCreateTab( _
					$oProperties.Item("Left"), $oProperties.Item("Top"), $oProperties.Item("Width"), $oProperties.Item("Height"), _
					$oProperties.Item("CtrlStyle"), $oProperties.Item("CtrlExStyle") _
				)
				__KODAParser_createControls($oForm, $oObject.selectNodes("components/object"), $oProperties.Item("Left"), $oProperties.Item("Top"), $iCtrlID)

				GUICtrlSetState(_objGet($oForm, _objGet($oProperties, "ActivePage", ""), 0), $GUI_SHOW)
			; ---
			Case "TTabSheet"
				GUICtrlCreateTabItem($oProperties.Item("Caption"))

				$aDispRect = _GUICtrlTab_GetDisplayRect(GUICtrlGetHandle($vUserData))
				__KODAParser_createControls($oForm, $oObject.selectNodes("components/object"), $iXOffset + $aDispRect[0], $iYOffset + $aDispRect[1])

				GUICtrlCreateTabItem("")
			; ---
;~ 			Case "TAStatusBar"
;~ 			Case "TAIPAddress"
;~ 			Case "TAToolBar"
			Case Else
				ContinueLoop
		EndSwitch
		; ---
		; font
		Local $aFont = __KODAParser_processFont($oForm, $oProperties)
		GUICtrlSetFont($iCtrlID, $aFont[0], $aFont[1], $aFont[2], $aFont[3])

		; colors
		$iColor = __KODAParser_identifiers_colors(_objGet($oProperties, "Color", ""))
		If $iColor <> "" Then GUICtrlSetBkColor($iCtrlID, $iColor)

		$iColor = __KODAParser_identifiers_colors(_objGet($oProperties, "Font.Color", ""))
		If $iColor <> "" Then GUICtrlSetColor($iCtrlID, $iColor)

		; cursor
		GUICtrlSetCursor($iCtrlID, __KODAParser_identifiers_cursor(_objGet($oProperties, "Cursor", "")))

		; hint
		GUICtrlSetTip($iCtrlID, _objGet($oProperties, "Hint", ""))

		; visible
		If Not _objGet($oProperties, "Visible", True) Then GUICtrlSetState($iCtrlID, $GUI_HIDE)

		; enable
		If Not _objGet($oProperties, "Enabled", True) Then GUICtrlSetState($iCtrlID, $GUI_DISABLE)

		; checked
		If _objGet($oProperties, "Checked", False) Then GUICtrlSetState($iCtrlID, $GUI_CHECKED)

		; resizing
		Local $aResizing = $oProperties.Item("Resizing")
		If IsArray($aResizing) And $aResizing[0] > 0 Then
			Local $iResizing = 0
			For $i = 1 To $aResizing
				$iResizing += __KODAParser_identifiers_docking($aResizing[$i])
			Next
			GUICtrlSetResizing($iCtrlID, $iResizing)
		EndIf

		; ---
		; store control in returned object (if non empty name)
		If String($oObject.getAttribute("name")) Then _objSet($oForm, $oObject.getAttribute("name"), $iCtrlID)

		$oObject = 0
		$oProperties = 0
	Next
EndFunc

Func __KODAParser_sortObjects($oObjects)
	Local $i = 0, $aObjects[$oObjects.length][3] ; $oObject, $oProperties, $iTabOrder
	For $oObject In $oObjects
		$aObjects[$i][0] = $oObject
		$aObjects[$i][1] = __KODAParser_readObjectProperties($oObject)

		; default to a very big taborder (so that elements such as menus and popuMenus will be the last ones)
		; add + $i to preserve original order for these controls
		$aObjects[$i][2] = _objGet($aObjects[$i][1], "TabOrder", 1000000 + $i)
		$i += 1
	Next
	_ArraySort($aObjects, 0, 0, 0, 2)
	ReDim $aObjects[UBound($aObjects)][2]
	Return $aObjects
EndFunc

; -------------------------------------------------------------------------------------------------

Func __KODAParser_readObjectProperties($oObject)
	Local $oRet = _objCreate()
	For $oProperty In $oObject.selectNodes("properties/property")
		_objSet($oRet, $oProperty.getAttribute("name"), __KODAParser_readProperty($oProperty))
	Next
	Return $oRet
EndFunc

Func __KODAParser_readProperty($oProperty)
	Switch $oProperty.getAttribute("vt")
		Case "Binary" ;TODO: test
			Local $oBins = $oProperty.selectNodes("bin")
			If $oBins.length <= 0 Then Return Binary("")

			Local $aBins[$oBins.length], $i = 0
			For $oBin In $oBins
				$aBins[$i] = Binary($oBin.text)
				$i += 1
			Next
			Return Binary("0x" & _ArrayToString($aBins, ""))
		; ---
		Case "Collection" ; collection of properties
			Local $oItems = $oProperty.selectNodes("collection/item")
			Local $aRet[1] = [0]
			If $oItems.length > 0 Then
				ReDim $aRet[$oItems.length + 1]
				$aRet[0] = $oItems.length
				$i = 1
				For $oItem In $oItems
					$aRet[$i] = _objCreate()
					For $oItemProp In $oItem.selectNodes("property")
						_objSet($aRet[$i], $oItemProp.getAttribute("name"), __KODAParser_readProperty($oItemProp))
					Next
					$i += 1
				Next
			EndIf
			Return $aRet
		; ---
		Case "List" ; list of strings (only?)
			Local $aRet[1] = [0]
			For $oLi In $oProperty.selectNodes("list/li") ;TODO: possible bad performance for (very) big lists
				_ArrayAdd($aRet, __KODAParser_readProperty($oLi))
				$aRet[0] += 1
			Next
			Return $aRet
		; ---
		Case "Set" ; list of Identifiers
			Local $aRet[1] = [0]
			If $oProperty.text Then
				$aRet = StringSplit($oProperty.text, ",")
				For $i = 1 To $aRet[0]
					$aRet[$i] = StringStripWS($aRet[$i], 3)
				Next
			EndIf
			Return $aRet
		; ---
		Case "False"
			Return False
		Case "True"
			Return True
		Case "Ident" ; identifier
			Return String($oProperty.text)
		Case "Int8", "Int16", "Int32"
			Return Int($oProperty.text)
		Case "Single", "Extended"
			Return Number($oProperty.text)
		Case "String", "UTF8String", "WString" ;TODO: handle unicode strings?
			Return String($oProperty.text)
		Case Else
			Return String($oProperty.text)
	EndSwitch
EndFunc

;TODO: test on non-highDPI displays
Func __KODAParser_processFont($oForm, $oProperties)
	Local $aRet[4] ; size, weight, attributes, name

	; GUI props are used for PixelsPerInch, and to default to GUI font if no font info are provided for the control
	Local $oGUIProps = _objGet($oForm, "___###gui_properties###___")

	; calculate font point size (https://support.microsoft.com/en-us/help/74299/info-calculating-the-logical-height-and-point-size-of-a-font)
	$aRet[0] = Round(Abs(_objGet($oProperties, "Font.Height", _objGet($oGUIProps, "Font.Height")) * 72 / _objGet($oGUIProps, "PixelsPerInch")))

	; font attributes
	Local $aAttribs = _objGet($oProperties, "Font.Style", _objGet($oGUIProps, "Font.Style"))
	$aRet[1] = _ArraySearch($aAttribs, "fsBold", 1) > 0 ? 800 : 400

	$aRet[2] = 0
	For $i = 1 To $aAttribs[0]
		$aRet[2] += __KODAParser_identifiers_fontStyle($aAttribs[$i])
	Next

	; font name
	$aRet[3] = _objGet($oProperties, "Font.Name", _objGet($oGUIProps, "Font.Name"))

	Return $aRet
EndFunc

; -------------------------------------------------------------------------------------------------
; Identifiers

Func __KODAParser_identifiers_fontStyle($sIdent)
	Local $aDef[][] = [ _
		["fsItalic", 2], _
		["fsUnderline", 4], _
		["fsStrikeOut", 8] _
	]
	For $i = 0 To UBound($aDef) - 1
		If $sIdent = $aDef[$i][0] Then Return $aDef[$i][1]
	Next
	Return 0 ; normal font
EndFunc

Func __KODAParser_identifiers_docking($sIdent)
	Local $aDef[][] = [ _
		["DockAuto", 1], _
		["DockLeft", 2], _
		["DockRight", 4], _
		["DockTop", 32], _
		["DockBottom", 64], _
		["DockHCenter", 8], _
		["DockVCenter", 128], _
		["DockWidth", 256], _
		["DockHeight", 512] _
	]
	For $i = 0 To UBound($aDef) - 1
		If $sIdent = $aDef[$i][0] Then Return $aDef[$i][1]
	Next
	Return 0 ; no docking value
EndFunc

Func __KODAParser_identifiers_cursor($sIdent)
	Local $aDef[][] = [ _
		["crAppStart", 1], _
		["crArrow", 2], _
		["crCross", 3], _
		["crDefault", -1], _
		["crDrag", 2], _
		["crHandPoint", 0], _
		["crHelp", 4], _
		["crHourGlass", 15], _
		["crHSplit", 2], _
		["crBeam", 5], _
		["crMultiDrag", 2], _
		["crNo", 7], _
		["crNoDrop", 7], _
		["crSizeAll", 9], _
		["crSizeNESW", 10], _
		["crSizeNS", 11], _
		["crSizeNWSE", 12], _
		["crSizeWE", 13], _
		["crSQLWait", 2], _
		["crUpArrow", 14], _
		["crVSplit", 2] _
	]
	For $i = 0 To UBound($aDef) - 1
		If $aDef[$i][0] = $sIdent Then Return $aDef[$i][1]
	Next
	Return -1 ; $crDefault
EndFunc

Func __KODAParser_identifiers_colors($sIdent)
	If Not $sIdent Then Return ""

	Local $aDef[][] = [ _ ; standard colors (unused: $COLOR_MEDBLUE)
		["clDefault", $CLR_DEFAULT], _
		["clNone", $CLR_NONE], _
		["clBlack", $COLOR_BLACK], _
		["clMaroon", $COLOR_MAROON], _
		["clGreen", $COLOR_GREEN], _
		["clOlive", $COLOR_OLIVE], _
		["clNavy", $COLOR_NAVY], _
		["clPurple", $COLOR_PURPLE], _
		["clTeal", $COLOR_TEAL], _
		["clGray", $COLOR_GRAY], _
		["clSilver", $COLOR_SILVER], _
		["clRed", $COLOR_RED], _
		["clLime", $COLOR_LIME], _
		["clYellow", $COLOR_YELLOW], _
		["clBlue", $COLOR_BLUE], _
		["clFuchsia", $COLOR_FUCHSIA], _
		["clAqua", $COLOR_AQUA], _
		["clWhite", $COLOR_WHITE], _
		["clMoneyGreen", $COLOR_MONEYGREEN], _
		["clSkyBlue", $COLOR_SKYBLUE], _
		["clCream", $COLOR_CREAM], _
		["clMedGray", $COLOR_MEDGRAY] _
	]
	For $i = 0 To UBound($aDef) - 1
		If $aDef[$i][0] = $sIdent Then Return $aDef[$i][1]
	Next

	Dim $aDef[][] = [ _ ; system colors (unused: $COLOR_3DHIGHLIGHT, $COLOR_3DHILIGHT, $COLOR_3DSHADOW)
		["clActiveBorder", $COLOR_ACTIVEBORDER], _
		["clActiveCaption", $COLOR_ACTIVECAPTION], _
		["clAppWorkSpace", $COLOR_APPWORKSPACE], _
		["clBackground", $COLOR_BACKGROUND], _
		["clBtnFace", $COLOR_BTNFACE], _
		["clBtnHighlight", $COLOR_BTNHIGHLIGHT], _
		["clBtnHilight", $COLOR_BTNHILIGHT], _ ; same as $COLOR_BTNHIGHLIGHT (added)
		["clBtnShadow", $COLOR_BTNSHADOW], _
		["clBtnText", $COLOR_BTNTEXT], _
		["clCaptionText", $COLOR_CAPTIONTEXT], _
		["clDesktop", $COLOR_DESKTOP], _ ; same as $COLOR_BACKGROUND (added)
		["clGradientActiveCaption", $COLOR_GRADIENTACTIVECAPTION], _
		["clGradientInactiveCaption", $COLOR_GRADIENTINACTIVECAPTION], _
		["clGrayText", $COLOR_GRAYTEXT], _
		["clHighlight", $COLOR_HIGHLIGHT], _
		["clHighlightText", $COLOR_HIGHLIGHTTEXT], _
		["clHotLight", $COLOR_HOTLIGHT], _
		["clInactiveBorder", $COLOR_INACTIVEBORDER], _
		["clInactiveCaption", $COLOR_INACTIVECAPTION], _
		["clInactiveCaptionText", $COLOR_INACTIVECAPTIONTEXT], _
		["clInfoBk", $COLOR_INFOBK], _
		["clInfoText", $COLOR_INFOTEXT], _
		["clMenu", $COLOR_MENU], _
		["clMenuBar", $COLOR_MENUBAR], _
		["clMenuHilight", $COLOR_MENUHILIGHT], _
		["clMenuHighlight", $COLOR_MENUHILIGHT], _ ; HILIGH is the same as HIGHLIGHT (I think! :p)
		["clMenuText", $COLOR_MENUTEXT], _
		["clScrollBar", $COLOR_SCROLLBAR], _
		["cl3DFace", $COLOR_3DFACE], _ ; (added)
		["cl3DHighlight", $COLOR_3DHIGHLIGHT], _ ; (added)
		["cl3DHilight", $COLOR_3DHILIGHT], _ ; (added)
		["cl3DShadow", $COLOR_3DSHADOW], _ ; (added)
		["cl3DDkShadow", $COLOR_3DDKSHADOW], _
		["cl3DLight", $COLOR_3DLIGHT], _
		["clWindow", $COLOR_WINDOW], _
		["clWindowFrame", $COLOR_WINDOWFRAME], _
		["clWindowText", $COLOR_WINDOWTEXT] _
	]
	For $i = 0 To UBound($aDef) - 1
		If $aDef[$i][0] = $sIdent Then Return _WinAPI_GetSysColor($aDef[$i][1])
	Next
	Return ""
EndFunc

; -------------------------------------------------------------------------------------------------
; Scripting Dictionary helpers

Func _objCreate($bCaseSensitiveKeys = False)
	Local $oObj = ObjCreate("Scripting.Dictionary")
	$oObj.CompareMode = ($bCaseSensitiveKeys ? 0 : 1)
	Return $oObj
EndFunc

Func _objSet(ByRef $oObj, $sKey, $vValue, $bOverwrite = True)
	If Not IsObj($oObj) Then Return SetError(1, 0, False)
	$sKey = String($sKey)
	If $oObj.Exists($sKey) And Not $bOverwrite Then Return SetError(2, 0, False)
	$oObj.Item($sKey) = $vValue
	Return True
EndFunc

Func _objGet(Const ByRef $oObj, $sKey, $vDefaultValue = "")
	$sKey = String($sKey)
	If Not IsObj($oObj) Or Not $oObj.Exists($sKey) Then Return SetError(1, 0, $vDefaultValue)
	Return $oObj.Item($sKey)
EndFunc

Func _objDel(ByRef $oObj, $sKey)
	$sKey = String($sKey)
	If Not IsObj($oObj) Or Not $oObj.Exists($sKey) Then Return SetError(1, 0, False)
	$oObj.Remove($sKey)
	Return True
EndFunc

Func _objEmpty(ByRef $oObj)
	If Not IsObj($oObj) Then Return SetError(1, 0, False)
	$oObj.RemoveAll()
EndFunc

Func _objExists(Const ByRef $oObj, $sKey)
	If Not IsObj($oObj) Then Return SetError(1, 0, False)
	Return $oObj.Exists(String($sKey))
EndFunc

Func _objCount(Const ByRef $oObj)
	If Not IsObj($oObj) Then Return SetError(1, 0, -1)
	Return $oObj.Count
EndFunc

Func _objKeys(Const ByRef $oObj)
	If Not IsObj($oObj) Then Return SetError(1, 0, Null)
	Return $oObj.Keys()
EndFunc

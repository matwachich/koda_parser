#NoTrayIcon
#include <Array.au3>
#include "kodaparser.au3"

$oForm = _KODAParser_Do("testing.kxf", 330, 194)

GUISetState(@SW_SHOW, _KODAForm_HWnd($oForm))

$aGroupCtrls = _KODAForm_CtrlChildrenNames($oForm, "ControlGroup1")
_ArrayDisplay($aGroupCtrls, "Group control names")

$aGroupCtrls = _KODAForm_CtrlChildrenIDs($oForm, "ControlGroup1")
_ArrayDisplay($aGroupCtrls, "Group control IDs")

While 1
	Switch GUIGetMsg()
		Case -3
			Exit
		Case _KODAForm_CtrlID($oForm, "Button1")
			MsgBox(64, "Button click", "Button 1" & @CRLF & @CRLF & "Input: " & GUICtrlRead(_KODAForm_CtrlID($oForm, "Input1")))
		Case _KODAForm_CtrlID($oForm, "Button2")
			GUICtrlSetData(_KODAForm_CtrlID($oForm, "Label1"), GUICtrlRead(_KODAForm_CtrlID($oForm, "Input1")))
	EndSwitch
WEnd

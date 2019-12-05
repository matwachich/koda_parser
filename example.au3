#NoTrayIcon
#include "kodaparser.au3"

$oForm = _KODAParser_Do("testing.kxf", 330, 154)

GUISetState(@SW_SHOW, _KODAForm_HWnd($oForm))

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

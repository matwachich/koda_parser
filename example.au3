#NoTrayIcon
#include "kodaparser.au3"

$oForm = _KODAParser_DoFile("testing.kxf")

GUISetState(@SW_SHOW, $oForm.Item("hGUI"))

While 1
	Switch GUIGetMsg()
		Case -3
			Exit
		Case $oForm.Item("Button1")
			MsgBox(64, "Button click", "Button 1")
		Case $oForm.Item("Button2")
			MsgBox(64, "Button click", "Button 2")
	EndSwitch
WEnd

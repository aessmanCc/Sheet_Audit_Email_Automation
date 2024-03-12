Attribute VB_Name = "Send_All"
Option Explicit

Sub Send_All()
Dim messageMac, inputMac

messageMac = "Do you want to generate Audit"
inputMac = Application.InputBox(messageMac, Title:="Audit")

If inputMac = "n" Or inputMac = "N" Or inputMac = "No" Or inputMac = "NO" Or inputMac = 0 Or inputMac = vbCancel Or inputMac = "" Then Exit Sub

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call Sheet_Loop

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Audit Complete."

End Sub

Attribute VB_Name = "mdlClienteAPIs"
Option Explicit

Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Declare Sub Sleep Lib "kernel32" (ByVal lngMilisegundos As Long)

Public Sub AlwaysOnTop(frmID As Form, OnTop As Boolean)
    On Error Resume Next
' True lo pone OnTop
' False le quita el OnTop
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop Then
        OnTop = SetWindowPos(frmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(frmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub


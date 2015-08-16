Attribute VB_Name = "OnTop"
Option Explicit

#If Win16 Then
    Private Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'*************************************************************************
'* Function: KeepOnTop(F As Form)
'*
'*
'*************************************************************************
'* Description: Keep form on top.
'*
'*
'*************************************************************************
'* Parameters: Form Control
'*
'*************************************************************************
'* Notes: The SetWindowPos API call gets turned off if the form is
'*        minimized.  Put this code in the resize event to make sure
'*        your form stays on top.
'*
'*************************************************************************
'* Returns:
'*************************************************************************
Public Sub KeepOnTop(ByVal hwnd As Long)

    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'Call SetWindowPos(Frm.hWnd, WS_EX_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)

End Sub

Public Sub KillOnTop(ByVal hwnd As Long)

    Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'Call SetWindowPos(Frm.hWnd, WS_EX_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'Call SetWindowLong(Frm.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)

End Sub


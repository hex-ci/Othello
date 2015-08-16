Attribute VB_Name = "modCustomForm"
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'常数声明
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CYCAPTION = 4

'模块级变量声明
Private MyRgn As Long

Public Sub CreateRectForm(ByVal hWnd As Long, ByVal Width As Long, ByVal Height As Long)
    Dim CaptionHeight As Long
    Dim BorderX As Long
    Dim BorderY As Long

    On Error Resume Next

    If MyRgn <> 0 Then Exit Sub

    CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    BorderX = GetSystemMetrics(SM_CXBORDER)
    BorderY = GetSystemMetrics(SM_CYBORDER)

    MyRgn = CreateRectRgn(BorderX + 2, CaptionHeight + BorderY + 2, Width - BorderX - 2, Height - BorderY - 2)
    Call SetWindowRgn(hWnd, MyRgn, True)

    WindowWidth = GetTwipX(Width - BorderX - 6)
    WindowHeight = GetTwipY(Height - BorderY - CaptionHeight - 6)
    gsngCaptionHeight = GetTwipY(CaptionHeight)
    gsngBorderX = GetTwipX(BorderX + 2)
    gsngBorderY = GetTwipY(BorderY + 2)
End Sub

Public Sub DeleteRgn()
    If MyRgn <> 0 Then Call DeleteObject(MyRgn)
End Sub

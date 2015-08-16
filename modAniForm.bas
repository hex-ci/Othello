Attribute VB_Name = "modAniForm"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
'Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const R2_NOTXORPEN = 10
Private Const SW_SHOW = 5
Private Const SW_HIDE = 0

'Public Sub AniShowFrm(ByVal Frm As Long, Optional ByVal Speed As Long = 20)
'    Dim hDC As Long
'    Dim rcCurrent As RECT
'    Dim rcNew As RECT
'    Dim Step1 As Long
'    Dim Step2 As Long
'    Dim i As Long
'
'    On Error Resume Next
'
'    hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
'    Call SetROP2(hDC, R2_NOTXORPEN)
'    Call GetWindowRect(Frm, rcCurrent)
'    Step1 = (rcCurrent.Right - rcCurrent.Left) / Speed / 2
'    Step2 = (rcCurrent.Bottom - rcCurrent.Top) / Speed / 2
'    With rcCurrent
'        .Left = (.Right - .Left) \ 2 + .Left
'        .Right = .Left
'        .Top = (.Bottom - .Top) \ 2 + .Top
'        .Bottom = .Top
'    End With
'    For i = 1 To Speed
'        Call Rectangle(hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom)
'        Call Sleep(30)
'        Call Rectangle(hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom)
'        With rcCurrent
'            .Left = .Left - Step1
'            .Top = .Top - Step2
'            .Bottom = .Bottom + Step2
'            .Right = .Right + Step1
'        End With
'        'DoEvents
'    Next i
'    Call DeleteDC(hDC)
'End Sub
'
'Public Sub AniUnloadFrm(ByRef objFrm As Form, Optional ByVal Speed As Long = 20)
'    Dim hDC As Long
'    Dim rcCurrent As RECT
'    Dim rcNew As RECT
'    Dim Step1 As Long
'    Dim Step2 As Long
'    Dim i As Long
'    Dim OldCapture As Long
'
'    On Error Resume Next
'
'    'Call ShowWindowAsync(Frm, SW_HIDE)
'    'Call ShowWindow(Frm, SW_HIDE)
'    OldCapture = GetCapture()
'    Call SetCapture(objFrm.hWnd)
'    Call objFrm.Hide
'    DoEvents
'
'    hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
'    Call SetROP2(hDC, R2_NOTXORPEN)
'    Call GetWindowRect(objFrm.hWnd, rcCurrent)
'    Step1 = (rcCurrent.Right - rcCurrent.Left) / Speed / 2
'    Step2 = (rcCurrent.Bottom - rcCurrent.Top) / Speed / 2
'    For i = 1 To Speed
'        Call Rectangle(hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom)
'        Call Sleep(30)
'        Call Rectangle(hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom)
'        With rcCurrent
'            .Left = rcCurrent.Left + Step1
'            .Top = rcCurrent.Top + Step2
'            .Bottom = rcCurrent.Bottom - Step2
'            .Right = rcCurrent.Right - Step1
'        End With
'        'DoEvents
'    Next i
'    Call DeleteDC(hDC)
'
'    Call SetCapture(OldCapture)
'End Sub

Public Sub AniRotateShowFrm(ByVal Frm As Long, Optional ByVal Speed As Long = 20)
    Dim PPP1(3) As POINTAPI
    Dim PPP2(3) As POINTAPI
    Dim cx As Long
    Dim cy As Long
    Dim hDC As Long
    Dim rcCurrent As RECT
    Dim rcNew As RECT
    Dim Step1 As Long
    Dim Step2 As Long
    Dim ii As Long
    Dim Radian As Single

    On Error Resume Next

    hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call GetWindowRect(Frm, rcCurrent)
    cx = (rcCurrent.Right - rcCurrent.Left) \ 2 + rcCurrent.Left
    cy = (rcCurrent.Bottom - rcCurrent.Top) \ 2 + rcCurrent.Top
    PPP1(0).X = cx - 1
    PPP1(0).Y = cy - 1
    PPP1(1).X = PPP1(0).X + 1
    PPP1(1).Y = PPP1(0).Y - 1
    PPP1(2).X = PPP1(0).X + 1
    PPP1(2).Y = PPP1(0).Y + 1
    PPP1(3).X = PPP1(0).X - 1
    PPP1(3).Y = PPP1(0).Y - 1
    Step1 = (rcCurrent.Right - rcCurrent.Left) / Speed / 2
    Step2 = (rcCurrent.Bottom - rcCurrent.Top) / Speed / 2
    For Radian = 0 To 3.14159 Step 3.13159 / Speed
        PPP1(0).X = PPP1(0).X - Step1
        PPP1(0).Y = PPP1(0).Y - Step2
        PPP1(1).X = PPP1(1).X + Step1
        PPP1(1).Y = PPP1(1).Y - Step2
        PPP1(2).X = PPP1(2).X + Step1
        PPP1(2).Y = PPP1(2).Y + Step2
        PPP1(3).X = PPP1(3).X - Step1
        PPP1(3).Y = PPP1(3).Y + Step2
        For ii = 0 To 3
            PPP2(ii).X = (PPP1(ii).X - cx) * Cos(Radian) + (PPP1(ii).Y - cx) * Sin(Radian) + cx
            PPP2(ii).Y = (PPP1(ii).Y - cy) * Cos(Radian) - (PPP1(ii).X - cy) * Sin(Radian) + cy
        Next ii
        Call Polygon(hDC, PPP2(0), 4)
        Call Sleep(30)
        Call Polygon(hDC, PPP2(0), 4)
        'DoEvents
    Next Radian
    Call DeleteDC(hDC)
End Sub

Public Sub AniRotateUnloadFrm(ByRef objFrm As Form, Optional ByVal Speed As Long = 20)
    Dim PPP1(3) As POINTAPI
    Dim PPP2(3) As POINTAPI
    Dim cx As Long
    Dim cy As Long
    Dim hDC As Long
    Dim rcCurrent As RECT
    Dim rcNew As RECT
    Dim Step1 As Long
    Dim Step2 As Long
    Dim ii As Long
    Dim Radian As Single
    Dim OldCapture As Long

    On Error Resume Next

    'Call ShowWindowAsync(Frm, SW_HIDE)
    'Call ShowWindow(Frm, SW_HIDE)
    OldCapture = GetCapture()
    Call SetCapture(objFrm.hWnd)
    Call objFrm.Hide
    DoEvents

    hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call GetWindowRect(objFrm.hWnd, rcCurrent)
    cx = (rcCurrent.Right - rcCurrent.Left) \ 2 + rcCurrent.Left
    cy = (rcCurrent.Bottom - rcCurrent.Top) \ 2 + rcCurrent.Top
    PPP1(0).X = rcCurrent.Left
    PPP1(0).Y = rcCurrent.Top
    PPP1(1).X = rcCurrent.Right
    PPP1(1).Y = rcCurrent.Top
    PPP1(2).X = rcCurrent.Right
    PPP1(2).Y = rcCurrent.Bottom
    PPP1(3).X = rcCurrent.Left
    PPP1(3).Y = rcCurrent.Bottom
    Step1 = (rcCurrent.Right - rcCurrent.Left) / Speed / 2
    Step2 = (rcCurrent.Bottom - rcCurrent.Top) / Speed / 2
    For Radian = 0 To -3.14159 Step -3.14159 / Speed
        PPP1(0).X = PPP1(0).X + Step1
        PPP1(0).Y = PPP1(0).Y + Step2
        PPP1(1).X = PPP1(1).X - Step1
        PPP1(1).Y = PPP1(1).Y + Step2
        PPP1(2).X = PPP1(2).X - Step1
        PPP1(2).Y = PPP1(2).Y - Step2
        PPP1(3).X = PPP1(3).X + Step1
        PPP1(3).Y = PPP1(3).Y - Step2
        For ii = 0 To 3
            PPP2(ii).X = (PPP1(ii).X - cx) * Cos(Radian) + (PPP1(ii).Y - cx) * Sin(Radian) + cx
            PPP2(ii).Y = (PPP1(ii).Y - cy) * Cos(Radian) - (PPP1(ii).X - cy) * Sin(Radian) + cy
        Next ii
        Call Polygon(hDC, PPP2(0), 4)
        Call Sleep(30)
        Call Polygon(hDC, PPP2(0), 4)
        'DoEvents
    Next Radian
    Call DeleteDC(hDC)

    Call SetCapture(OldCapture)
End Sub

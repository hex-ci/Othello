VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' 保存棋盘数据
Private mvarTables(7, 7) As Byte
'保持属性值的局部变量
Private mvarPanelHeight As Single '局部复制
'保持属性值的局部变量
Private mvarPanelWidth As Single '局部复制
'保持属性值的局部变量
Private mvarChessboardX As Single '局部复制
Private mvarChessboardY As Single '局部复制

'要引发该事件，请遵循下列语法使用 RaiseEvent:
'RaiseEvent DownChessMan[(arg1, arg2, ... , argn)]
Public Event DownChessMan(Col As Integer, Row As Integer, Man As Byte)

Public Function Umpire(ByVal Man As Byte) As Byte
    Dim MyNumber As Integer
    Dim YourNumber As Integer
    
    On Error Resume Next

    ' 计算自己的棋子数
    MyNumber = GetTotal(Man)
    
    ' 计算对方的棋子数
    YourNumber = GetTotal(ToPartner(Man))

    If MyNumber > YourNumber Then   ' 如果大于，胜利
        Umpire = GAME_WIN
    ElseIf MyNumber < YourNumber Then   ' 小于，失败
        Umpire = GAME_LOSE
    Else                                ' 其它情况，平局
        Umpire = GAME_DRAW
    End If
End Function

Public Sub Clear()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 7
        For j = 0 To 7
            mvarTables(j, i) = T_NONE
        Next j
    Next i
End Sub

Public Property Let ChessboardX(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ChessboardX = 5
    mvarChessboardX = vData
End Property
Public Property Get ChessboardX() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ChessboardX
    ChessboardX = mvarChessboardX
End Property
Public Property Let ChessboardY(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ChessboardY = 5
    mvarChessboardY = vData
End Property
Public Property Get ChessboardY() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ChessboardY
    ChessboardY = mvarChessboardY
End Property
Public Property Let PanelWidth(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PanelWidth = 5
    mvarPanelWidth = vData
End Property
Public Property Get PanelWidth() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PanelWidth
    PanelWidth = mvarPanelWidth
End Property
Public Property Let PanelHeight(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PanelHeight = 5
    mvarPanelHeight = vData
End Property
Public Property Get PanelHeight() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PanelHeight
    PanelHeight = mvarPanelHeight
End Property

Public Function GetCol(ByVal X As Single) As Integer
    GetCol = (X - ChessboardX + PanelWidth) \ PanelWidth - 1
End Function
Public Function GetRow(ByVal Y As Single) As Integer
    GetRow = (Y - ChessboardY + PanelHeight) \ PanelHeight - 1
End Function
Public Function GetX(ByVal Col As Integer) As Single
    GetX = mvarChessboardX + PanelWidth * Col
End Function
Public Function GetY(ByVal Row As Integer) As Single
    GetY = mvarChessboardY + PanelHeight * Row
End Function

Public Function GetMan(ByVal Col As Integer, ByVal Row As Integer) As Byte
    GetMan = mvarTables(Col, Row)
End Function
Public Function SetMan(ByVal Col As Integer, ByVal Row As Integer, ByVal Man As Byte) As Boolean
    mvarTables(Col, Row) = Man
    SetMan = True
End Function

Public Function GetTotal(ByVal Man As Byte) As Integer
    Dim Col As Integer
    Dim Row As Integer
    Dim Num As Integer
    'Num = 0
    On Error Resume Next

    For Row = 0 To 7
        For Col = 0 To 7
            If mvarTables(Col, Row) = Man Then Num = Num + 1
        Next Col
    Next Row
    GetTotal = Num
End Function

Public Function CanDown(ByVal Man As Byte) As Boolean
    Dim Col As Integer
    Dim Row As Integer

    On Error Resume Next

    For Row = 0 To 7
        For Col = 0 To 7
            If IsDown(Col, Row, Man) Then
                CanDown = True
                Exit Function
            End If
        Next Col
    Next Row
    CanDown = False
End Function

Public Function IsDown(ByVal Col As Integer, ByVal Row As Integer, ByVal Man As Byte) As Boolean
    Dim c As Integer
    Dim r As Integer
    Dim Value As Byte

    On Error GoTo ErrHandler

    If mvarTables(Col, Row) <> T_NONE Then
        IsDown = False
        Exit Function
    End If

    ' To Left-Bottom
    c = Col - 1: r = Row + 1
    Do While (c >= 0 And r <= 7)
        Value = mvarTables(c, r)
        If Value = Man And Col - c > 1 And r - Row > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
        End If
        c = c - 1: r = r + 1
    Loop
    
    ' To Left
    For c = Col - 1 To 0 Step -1
        Value = mvarTables(c, Row)
        If Value = Man And Col - c > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit For
        End If
    Next c
    
    ' To Left-Top
    c = Col - 1: r = Row - 1
    Do While (c >= 0 And r >= 0)
        Value = mvarTables(c, r)
        If Value = Man And Col - c > 1 And Row - r > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
        End If
        c = c - 1: r = r - 1
    Loop
    
    ' To Top
    For r = Row - 1 To 0 Step -1
        Value = mvarTables(Col, r)
        If Value = Man And Row - r > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit For
        End If
    Next r
    
    ' To Right-Top
    c = Col + 1: r = Row - 1
    Do While (c <= 7 And r >= 0)
        Value = mvarTables(c, r)
        If Value = Man And c - Col > 1 And Row - r > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
        End If
        c = c + 1: r = r - 1
    Loop
    
    ' To Right
    For c = Col + 1 To 7
        Value = mvarTables(c, Row)
        If Value = Man And c - Col > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit For
        End If
    Next c
    
    ' To Right-Bottom
    c = Col + 1: r = Row + 1
    Do While (c <= 7 And r <= 7)
        Value = mvarTables(c, r)
        If Value = Man And c - Col > 1 And r - Row > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
        End If
        c = c + 1: r = r + 1
    Loop

    ' To Bottom
    For r = Row + 1 To 7
        Value = mvarTables(Col, r)
        If Value = Man And r - Row > 1 Then
            IsDown = True
            Exit Function
        Else
            If Value = T_NONE Or Value = Man Then Exit For
        End If
    Next r

ErrHandler:
    IsDown = False
End Function

Private Sub SetDrawMan(ByVal StartCol As Integer, ByVal StartRow As Integer, ByVal EndCol As Integer, ByVal EndRow As Integer, ByVal Man As Byte)
    Dim c As Integer
    Dim r As Integer
    Dim ColSgn As Integer
    Dim RowSgn As Integer

    On Error Resume Next

    ColSgn = Sgn(EndCol - StartCol)
    RowSgn = Sgn(EndRow - StartRow)
    c = StartCol: r = StartRow
    Do While (((c <= EndCol And ColSgn > 0) Or (c >= EndCol And ColSgn < 0) Or ColSgn = 0) And ((r <= EndRow And RowSgn > 0) Or (r >= EndRow And RowSgn < 0) Or RowSgn = 0))
        mvarTables(c, r) = Man
        RaiseEvent DownChessMan(c, r, Man)
        c = c + ColSgn: r = r + RowSgn
    Loop
End Sub

Public Function DownMan(ByVal Col As Integer, ByVal Row As Integer, ByVal Man As Byte) As Integer
    Dim c As Integer
    Dim r As Integer
    Dim Value As Byte
    Dim Number As Integer
    Dim Temp As Integer

    On Error Resume Next

    If Man = T_NONE Then
        DownMan = 0
        Exit Function
    End If
    'Number = 0
    'temp = 0
    
    mvarTables(Col, Row) = Man
    RaiseEvent DownChessMan(Col, Row, Man)

    ' To Left-Bottom
    c = Col - 1: r = Row + 1
    Do While (c >= 0 And r <= 7)
        Value = mvarTables(c, r)
        If Value = Man And Col - c > 1 And r - Row > 1 Then
            Call SetDrawMan(Col - 1, Row + 1, c, r, Man)
            Number = Temp
            Exit Do
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
            Temp = Temp + 1
        End If
        c = c - 1: r = r + 1
    Loop

    ' To Left
    Temp = 0
    For c = Col - 1 To 0 Step -1
        Value = mvarTables(c, Row)
        If Value = Man And Col - c > 1 Then
            Call SetDrawMan(Col - 1, Row, c, Row, Man)
            Number = Number + Temp
            Exit For
        Else
            If Value = T_NONE Or Value = Man Then Exit For
            Temp = Temp + 1
        End If
    Next c

    ' To Left-Top
    Temp = 0
    c = Col - 1: r = Row - 1
    Do While (c >= 0 And r >= 0)
        Value = mvarTables(c, r)
        If Value = Man And Col - c > 1 And Row - r > 1 Then
            Call SetDrawMan(Col - 1, Row - 1, c, r, Man)
            Number = Number + Temp
            Exit Do
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
            Temp = Temp + 1
        End If
        c = c - 1: r = r - 1
    Loop

    ' To Top
    Temp = 0
    For r = Row - 1 To 0 Step -1
        Value = mvarTables(Col, r)
        If Value = Man And Row - r > 1 Then
            Call SetDrawMan(Col, Row - 1, Col, r, Man)
            Number = Number + Temp
            Exit For
        Else
            If Value = T_NONE Or Value = Man Then Exit For
            Temp = Temp + 1
        End If
    Next r

    ' To Right-Top
    Temp = 0
    c = Col + 1: r = Row - 1
    Do While (c <= 7 And r >= 0)
        Value = mvarTables(c, r)
        If Value = Man And c - Col > 1 And Row - r > 1 Then
            Call SetDrawMan(Col + 1, Row - 1, c, r, Man)
            Number = Number + Temp
            Exit Do
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
            Temp = Temp + 1
        End If
        c = c + 1: r = r - 1
    Loop

    ' To Right
    Temp = 0
    For c = Col + 1 To 7
        Value = mvarTables(c, Row)
        If Value = Man And c - Col > 1 Then
            Call SetDrawMan(Col + 1, Row, c, Row, Man)
            Number = Number + Temp
            Exit For
        Else
            If Value = T_NONE Or Value = Man Then Exit For
            Temp = Temp + 1
        End If
    Next c

    ' To Right-Bottom
    Temp = 0
    c = Col + 1: r = Row + 1
    Do While (c <= 7 And r <= 7)
        Value = mvarTables(c, r)
        If Value = Man And c - Col > 1 And r - Row > 1 Then
            Call SetDrawMan(Col + 1, Row + 1, c, r, Man)
            Number = Number + Temp
            Exit Do
        Else
            If Value = T_NONE Or Value = Man Then Exit Do
            Temp = Temp + 1
        End If
        c = c + 1: r = r + 1
    Loop

    ' To Bottom
    Temp = 0
    For r = Row + 1 To 7
        Value = mvarTables(Col, r)
        If Value = Man And r - Row > 1 Then
            Call SetDrawMan(Col, Row + 1, Col, r, Man)
            Number = Number + Temp
            Exit For
        Else
            If Value = T_NONE Or Value = Man Then Exit For
            Temp = Temp + 1
        End If
    Next r
    DownMan = Number
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next

    For i = 0 To 7
        For j = 0 To 7
            mvarTables(j, i) = T_NONE
        Next j
    Next i
    mvarTables(3, 3) = T_BLACK
    mvarTables(3, 4) = T_WHITE
    mvarTables(4, 3) = T_WHITE
    mvarTables(4, 4) = T_BLACK

    'mvarTables(3, 3) = T_BLACK
    'mvarTables(3, 4) = T_BLACK
    'mvarTables(4, 3) = T_BLACK
    'mvarTables(4, 4) = T_BLACK

    mvarChessboardX = TablePos.X + 390
    mvarChessboardY = TablePos.Y + 355
    mvarPanelHeight = 540
    mvarPanelWidth = 540
End Sub

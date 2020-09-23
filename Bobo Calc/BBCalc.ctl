VERSION 5.00
Begin VB.UserControl BBCalc 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "BBCalc.ctx":0000
   PropertyPages   =   "BBCalc.ctx":0822
   ScaleHeight     =   330
   ScaleWidth      =   480
   ToolboxBitmap   =   "BBCalc.ctx":0848
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "BBCalc.ctx":0B5A
      Top             =   960
      Width           =   240
   End
   Begin VB.Menu mnuPUBase 
      Caption         =   ""
      Begin VB.Menu mnuPU 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuPU 
         Caption         =   "Paste"
         Index           =   1
      End
   End
End
Attribute VB_Name = "BBCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'All errors, both programatical and calculations, entirely free of charge

Public Enum ClipAction
    Copy = 0
    Paste = 1
End Enum
Public Enum OperationMode
    Division = 1
    Multiplication = 2
    Subtraction = 3
    Addition = 4
    SquareRoot = 5
    Percent = 6
    Reciprocal = 7
    RaiseToPower = 8
    Squared = 9
    Cubed = 10
End Enum
Public Enum TrigMode
    Sine = 0
    Cosine = 1
    Tangent = 2
    AntiSine = 3
    AntiCosine = 4
    AntiTangent = 5
    PI = 6
    RandomNum = 7
End Enum
Dim CurValue As Double
Dim CurOperator As Double
Dim curOpType As Long
Dim CurMemory As Double
Dim DoingOp As Boolean
Dim IsDecimal As Boolean
Dim CantDoMessage As Boolean
Dim IsRadNotDegree As Boolean
Private Const mPI = 3.14159265358979
Dim WithEvents DBox As PictureBox
Attribute DBox.VB_VarHelpID = -1
Public Event DisplayChanged()
Public Property Get AngleMode() As Boolean
    AngleMode = IsRadNotDegree
End Property
Public Property Let AngleMode(ByVal vNewValue As Boolean)
    IsRadNotDegree = vNewValue
End Property
Public Property Get CalcIcon() As Picture
    Set CalcIcon = Image1.Picture
End Property
Public Sub ShowBox(vNewValue As Object)
    Set DBox = vNewValue
    DBox.AutoRedraw = True
    DBox.ScaleMode = 1
    Clear
    On Error Resume Next
    DBox.SetFocus
End Sub
Public Property Let OperationType(ByVal vNewValue As OperationMode)
    If vNewValue > 0 Then
        If DoingOp Then
            CurValue = Operation
            UpdateDisplay CurValue
        End If
        curOpType = vNewValue
        DoingOp = True
        CurOperator = 0
        DBox.Tag = ""
        Select Case vNewValue
            Case 5, 7, 9, 10, 12, 14 'see Operation function
                DoResult
        End Select
    Else
        DoingOp = False
    End If
    On Error Resume Next
    DBox.SetFocus
End Property
Private Function Operation() As Double
    Select Case curOpType
        Case 1 'divide
            If CurOperator = 0 Then
                CantDoMessage = True
                Exit Function
            End If
            Operation = Abs(CurValue / CurOperator)
        Case 2 'multiply
            Operation = CurValue * CurOperator
        Case 3 'subtract
            Operation = CurValue - CurOperator
        Case 4 'add
            Operation = CurValue + CurOperator
        Case 5 'squareroot
            If CurValue < 0 Then 'cant be negative
                CantDoMessage = True
                Exit Function
            End If
            Operation = Sqr(CurValue)
        Case 6 'percent
            If CurOperator = 0 Then 'cant divide by zero
                CantDoMessage = True
                Exit Function
            End If
            Operation = CurValue / CurOperator * 100
        Case 7 'reciprocal
            Operation = 1 / CurValue
        Case 8 'raise to power of
            Operation = CurValue ^ CurOperator
        Case 9 'square
            Operation = CurValue ^ 2
        Case 10 'cube
            Operation = CurValue ^ 3
        Case 11 'mod
            Operation = CurValue Mod CurOperator
        Case 12 'integer
            Operation = Int(CurValue)
        Case 13 'Nth root
            Operation = CurValue ^ (1 / CurOperator)
        Case 14 'cubed root
            Operation = CurValue ^ (1 / 3) '
    End Select
End Function
Public Sub DoBackSpace()
    Dim temp As String, tmp As Double
    temp = lblDisplay.Caption
    If Right(temp, 1) = "." Then
        temp = Left(temp, Len(temp) - 2)
        IsDecimal = False
    Else
        temp = Left(temp, Len(temp) - 1)
        If Right(temp, 1) = "." Then IsDecimal = True
    End If
    tmp = Val(temp)
    If DoingOp Then
        CurOperator = tmp
    Else
        CurValue = tmp
    End If
    UpdateDisplay tmp
End Sub
Public Sub NumberInput(mNum As Integer)
    DBox.Tag = DBox.Tag & mNum
    PrintDisplay
    If DoingOp Then
        CurOperator = Val(DBox.Tag)
    Else
        CurValue = Val(DBox.Tag)
    End If
End Sub
Public Sub DoDecimalPoint()
    If InStr(DBox.Tag, ".") = 0 Then
        DBox.Tag = DBox.Tag + "."
        PrintDisplay
    End If
    IsDecimal = True
    On Error Resume Next
End Sub
Private Sub UpdateDisplay(mVal As Double)
    If mVal = 0 Then
        DBox.Tag = ""
    Else
        DBox.Tag = mVal
        If InStr(DBox.Tag, ".") = 0 And IsDecimal Then DBox.Tag = mVal & "."
    End If
    PrintDisplay
End Sub
Public Sub DoResult()
    If curOpType = 0 Then Exit Sub
    DoingOp = False
    Dim temp As String
    temp = Trim(Str(Operation))
    If CantDoMessage Then 'something illegal happened
        CantDoMessage = False
        MsgBox "Cannot comply with your request." + vbCrLf + "Current operation has been cancelled.", vbCritical, "Bobo Enterprises"
        ClearOperation
        Exit Sub
    End If
    If InStr(temp, ".") <> 0 Then IsDecimal = True
    CurValue = Val(temp)
    UpdateDisplay CurValue
    DBox.Tag = ""
End Sub

Public Sub Clear()
    curOpType = 0
    CurValue = 0
    CurOperator = 0
    UpdateDisplay 0
    IsDecimal = False
End Sub
Public Sub ClearOperation()
    curOpType = 0
    CurOperator = 0
    UpdateDisplay CurValue
End Sub
Public Sub KeyboardInput(KeyCode As Integer)
    Select Case KeyCode
        Case 8
            DoBackSpace
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
            NumberInput KeyCode - 48
        Case 190, 110
            DoDecimalPoint
        Case 96, 97, 98, 99, 100, 101, 102, 103, 104, 105
            NumberInput KeyCode - 96
        Case 111
            OperationType = 1 'divide
        Case 106
            OperationType = 2 'multiply
        Case 109
            OperationType = 3 'subtract
        Case 107
            OperationType = 4 'add
        Case 13
            DoResult 'equals
    End Select
    On Error Resume Next
    DBox.SetFocus
End Sub
Public Sub ChangeSign()
    Dim tmp As Double, st As Long 'positive-negative/negative-positive
    tmp = Val(DBox.Tag)
    tmp = -tmp
    If DoingOp Then
        CurOperator = tmp
    Else
        CurValue = tmp
    End If
    UpdateDisplay tmp
End Sub
Public Property Let TrigOp(Index As TrigMode)
    Dim angle As Double, tmp As Double
    On Error GoTo woops
    Select Case Index
        Case 0 'sin
            angle = Abs(Val(DBox.Tag) Mod 360)
            tmp = IIf(IsRadNotDegree, Sin(angle), Sin(angle * mPI / 180))
        Case 1 'cos
            angle = Abs(Val(DBox.Tag) Mod 360)
            tmp = IIf(IsRadNotDegree, Cos(angle), Cos(angle * mPI / 180))
        Case 2 'tan
            angle = Abs(Val(DBox.Tag) Mod 360)
            tmp = IIf(IsRadNotDegree, Tan(angle), Tan(angle * mPI / 180))
        Case 3 'sin-1
            angle = Val(DBox.Tag)
            If angle > 0 Then
                tmp = IIf(IsRadNotDegree, Atn(angle / Sqr(-angle * angle + 1)), Atn(angle / Sqr(-angle * angle + 1)) * 180 / mPI)
            ElseIf angle < 0 Then
                tmp = IIf(IsRadNotDegree, 360 - Atn(Abs(angle / Sqr(-angle * angle + 1))), 360 - Atn(Abs(angle / Sqr(-angle * angle + 1))) * 180 / mPI)
            End If
        Case 4 'cos-1
            angle = Val(DBox.Tag)
            tmp = IIf(IsRadNotDegree, (Atn(-angle / Sqr(-angle * angle + 1))) + (2 * Atn(1)), (Atn(-angle / Sqr(-angle * angle + 1)) * 180 / mPI) + (2 * Atn(1) * 180 / mPI))
        Case 5 'tan-1
            angle = Val(DBox.Tag)
            If angle > 0 Then
                tmp = IIf(IsRadNotDegree, Atn(angle), Atn(angle) * 180 / mPI)
            ElseIf angle < 0 Then
                tmp = IIf(IsRadNotDegree, 180 - Abs(Atn(angle)), 180 - Abs(Atn(angle) * 180 / mPI))
            End If
        Case 6 'pi
            tmp = mPI
        Case 7 'random number
            tmp = RandomNumber(1000, 1)
    End Select
    If DoingOp Then
        CurOperator = tmp
    Else
        CurValue = tmp
    End If
    UpdateDisplay tmp
    Exit Property
woops:
    MsgBox "Cannot comply with your request." + vbCrLf + "Current operation has been cancelled.", vbCritical, "Bobo Enterprises"
    ClearOperation

End Property
Private Function RandomNumber(Max As Double, min As Double) As Double
    Randomize Timer
    RandomNumber = (Max - min + 1) * Rnd + min
End Function
Public Sub MemFunction(Index As Integer)
    Select Case Index
        Case 0
            CurMemory = 0 'clear
        Case 1 'recall
            If DoingOp Then
                CurOperator = CurMemory
            Else
                CurValue = CurMemory
            End If
            UpdateDisplay CurMemory
        Case 2
            CurMemory = Val(lblDisplay.Caption) 'store value
        Case 3
            CurMemory = CurMemory + Val(lblDisplay.Caption) 'add value
    End Select
    On Error Resume Next
    DBox.SetFocus
End Sub
Public Function IsMemAvailable() As Boolean
    IsMemAvailable = (CurMemory <> 0)
End Function
Private Sub DBox_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyboardInput KeyCode
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Height = 330
    UserControl.Width = 480
End Sub

Private Sub PrintDisplay()
    DBox.Cls
    lblDisplay.Caption = DBox.Tag 'just to get the width of text so we can DBox.CurrentX for printing
    If InStr(lblDisplay.Caption, ".") = 0 Then lblDisplay.Caption = lblDisplay.Caption + "."
    If Left(lblDisplay.Caption, 1) = "." Then lblDisplay.Caption = "0" + lblDisplay.Caption
    DBox.CurrentX = DBox.Width - lblDisplay.Width - 60
    DBox.Print lblDisplay.Caption
    RaiseEvent DisplayChanged
    On Error Resume Next
    DBox.SetFocus
End Sub

Public Sub ClipboardAction(mDo As ClipAction)
    Select Case mDo
        Case 0 'copy
            Clipboard.Clear
            Clipboard.SetText lblDisplay.Caption
        Case 1 'paste
            DBox.Tag = Val(Clipboard.GetText)
            If InStr(DBox.Tag, ".") Then
                IsDecimal = True
            Else
                IsDecimal = False
            End If
            PrintDisplay
            If DoingOp Then
                CurOperator = Val(DBox.Tag)
            Else
                CurValue = Val(DBox.Tag)
            End If
    End Select
End Sub

Public Function CanPaste() As Boolean
    Dim z As Double, temp As String, temp1 As String
    'heck if the clipboard has strings
    If Not Clipboard.GetFormat(vbCFText) Then
        CanPaste = False
    Else
        temp = Clipboard.GetText
        'Is it a number ?
        z = Val(temp)
        If z = 0 Then
            'Even though it has no value, it could still be a number
            For z = 1 To Len(temp)
                temp1 = Mid(temp, z, 1)
                If Not temp1 Like "[0-9]" And temp1 <> "." Then
                    'No, it's not a number
                    CanPaste = False
                    Exit Function
                End If
            Next
            CanPaste = True
        Else
            CanPaste = True
        End If
    End If
End Function

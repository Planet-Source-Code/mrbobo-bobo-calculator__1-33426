VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5655
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin BoboCalc.BBCalc BBCalc1 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   582
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdBackSpace 
         Caption         =   "<<"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         ToolTipText     =   "Backspace"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C"
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   "Resets current constant and value cancelling all operations."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "CE"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   4
         ToolTipText     =   "Clears last operation."
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PicDisplay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         ScaleHeight     =   225
         ScaleWidth      =   2025
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
      Begin VB.OptionButton OptDegree 
         Caption         =   "deg"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton OptRadian 
         Caption         =   "rad"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   5415
      Begin VB.CommandButton cmdSci 
         Caption         =   "tan-1"
         Height          =   255
         Index           =   5
         Left            =   3720
         TabIndex        =   14
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSci 
         Caption         =   "cos-1"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   13
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSci 
         Caption         =   "sin-1"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   12
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSci 
         Caption         =   "tan"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSci 
         Caption         =   "cos"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   10
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSci 
         Caption         =   "sin"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSci 
         Height          =   255
         Index           =   6
         Left            =   4920
         Picture         =   "frmCalc.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Trigonometrical operation"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   735
      Begin VB.CommandButton cmdMem 
         Caption         =   "M+"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Add current value to memory"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdMem 
         Caption         =   "MS"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Store current value in memory"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdMem 
         Caption         =   "MR"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Memory recall"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem 
         Caption         =   "MC"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Clear memory"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   3000
      TabIndex        =   33
      Top             =   1320
      Width           =   2535
      Begin VB.CommandButton cmdOp 
         Height          =   375
         Index           =   1
         Left            =   120
         Picture         =   "frmCalc.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Operation - division"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "X"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   48
         ToolTipText     =   "Operation - multiplication"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "Operation - subtraction"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Operation - addition"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Height          =   375
         Index           =   5
         Left            =   720
         Picture         =   "frmCalc.frx":0EB6
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Operation - square root"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   44
         ToolTipText     =   "Operation - percentage - 30 (as a) % (of) 300 = 10)"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "1/x"
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   43
         ToolTipText     =   "Operation - reciprocal"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1920
         Picture         =   "frmCalc.frx":1440
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Operation - current value raised to the power of..."
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdEquals 
         Caption         =   "="
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "x ³"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1320
         TabIndex        =   40
         ToolTipText     =   "Operation - cubed"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "x ²"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   39
         ToolTipText     =   "Operation - sqared"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "mod"
         Height          =   375
         Index           =   11
         Left            =   1320
         TabIndex        =   38
         ToolTipText     =   "Operation - (16 mod 5 = 1),(74  mod 10 = 4)"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Height          =   375
         Index           =   13
         Left            =   1920
         Picture         =   "frmCalc.frx":19CA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Operation - nth root"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "int"
         Height          =   375
         Index           =   12
         Left            =   1320
         TabIndex        =   36
         ToolTipText     =   "Operation - integer"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdOp 
         Height          =   375
         Index           =   14
         Left            =   1320
         Picture         =   "frmCalc.frx":1F54
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Operation - cubed root"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdSci 
         Caption         =   "ran"
         Height          =   375
         Index           =   7
         Left            =   1920
         TabIndex        =   34
         ToolTipText     =   "Trigonometrical operation"
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   960
      TabIndex        =   20
      Top             =   1320
      Width           =   1935
      Begin VB.CommandButton cmdNum 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Numerical input"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Numerical input"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   30
         ToolTipText     =   "Numerical input"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   29
         ToolTipText     =   "Numerical input"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Numerical input"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   27
         ToolTipText     =   "Numerical input"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   26
         ToolTipText     =   "Numerical input"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Numerical input"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   24
         ToolTipText     =   "Numerical input"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "9"
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   23
         ToolTipText     =   "Numerical input"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdValue 
         Height          =   375
         Left            =   720
         Picture         =   "frmCalc.frx":24DE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Convert positive/negative"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdDecimal 
         Caption         =   "."
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         ToolTipText     =   "Decimal point"
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.Menu mnuEditbase 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   1
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Just a Calculator
'I put it in a control so you could change the interface easily

Private Sub BBCalc1_DisplayChanged()
    cmdOp(4).Enabled = (Val(PicDisplay.Tag) > -1)
    cmdSci(3).Enabled = (Val(PicDisplay.Tag) > -1 And Val(PicDisplay.Tag) < 1)
    cmdSci(4).Enabled = (Val(PicDisplay.Tag) > -1 And Val(PicDisplay.Tag) < 1)
End Sub
Private Sub cmdBackSpace_Click()
    BBCalc1.DoBackSpace
End Sub
Private Sub cmdClear_Click(Index As Integer)
    Select Case Index
        Case 0
            BBCalc1.Clear
        Case 1
            BBCalc1.ClearOperation
    End Select
End Sub
Private Sub cmdDecimal_Click()
    BBCalc1.DoDecimalPoint
End Sub
Private Sub cmdEquals_Click()
    BBCalc1.DoResult
End Sub
Private Sub cmdMem_Click(Index As Integer)
    BBCalc1.MemFunction Index
    cmdMem(1).Enabled = BBCalc1.IsMemAvailable
End Sub
Private Sub cmdNum_Click(Index As Integer)
    BBCalc1.NumberInput Index
End Sub
Private Sub cmdOp_Click(Index As Integer)
    BBCalc1.OperationType = Index
End Sub
Private Sub cmdSci_Click(Index As Integer)
    BBCalc1.TrigOp = Index
End Sub
Private Sub cmdValue_Click()
    BBCalc1.ChangeSign
End Sub

Private Sub Form_Load()
    BBCalc1.ShowBox PicDisplay
    Me.Icon = BBCalc1.CalcIcon
End Sub

Private Sub Form_Paint()
    PicDisplay.SetFocus
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
        Case 0
            BBCalc1.ClipboardAction Copy
        Case 1
            BBCalc1.ClipboardAction Paste
    End Select
End Sub

Private Sub mnuEditbase_Click()
    mnuEdit(1).Enabled = BBCalc1.CanPaste
End Sub

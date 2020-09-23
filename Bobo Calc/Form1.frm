VERSION 5.00
Object = "{686BD8BF-E960-46B0-80A0-CED2EAB6A9CF}#1.0#0"; "BBCalcControl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin BBCalcControl.BBCalc BBCalc1 
      Height          =   330
      Left            =   1200
      TabIndex        =   34
      Top             =   3120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   582
   End
   Begin VB.CommandButton cmdRecip 
      Caption         =   "1/x"
      Height          =   495
      Left            =   2760
      TabIndex        =   33
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdPercent 
      Caption         =   "%"
      Height          =   495
      Left            =   2760
      TabIndex        =   32
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdSqr 
      Caption         =   "sqr"
      Height          =   495
      Left            =   2760
      TabIndex        =   31
      Top             =   1800
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "rad"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   720
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "deg"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   480
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdRan 
      Caption         =   "ran"
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdPi 
      Caption         =   "pi"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdAnteTangent 
      Caption         =   "tan-1"
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdTangent 
      Caption         =   "tan"
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdAnteCoSine 
      Caption         =   "cos-1"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdCosine 
      Caption         =   "cos"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdAnteSine 
      Caption         =   "sin-1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "sin"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   3105
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   18
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   720
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   1320
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   720
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdDot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "±"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
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
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "X"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdSubtract 
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
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Backspace"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdCE 
      Caption         =   "CE"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnteCoSine_Click(Index As Integer)
    BBCalc1.TrigOp = AntiCosine
End Sub
Private Sub cmdAnteSine_Click(Index As Integer)
    BBCalc1.TrigOp = AntiSine
End Sub
Private Sub cmdAnteTangent_Click()
    BBCalc1.TrigOp = AntiTangent
End Sub
Private Sub cmdBS_Click()
    BBCalc1.DoBackSpace
End Sub
Private Sub cmdC_Click()
    BBCalc1.Clear
End Sub
Private Sub cmdCE_Click()
    BBCalc1.ClearOperation
End Sub
Private Sub cmdCosine_Click()
    BBCalc1.TrigOp = Cosine
End Sub
Private Sub cmdDivide_Click()
    BBCalc1.OperationType = Division
End Sub
Private Sub cmdDot_Click()
    BBCalc1.DoDecimalPoint
End Sub
Private Sub cmdEquals_Click()
    BBCalc1.DoResult
End Sub
Private Sub cmdMultiply_Click()
    BBCalc1.OperationType = Multiplication
End Sub
Private Sub cmdNum_Click(Index As Integer)
    BBCalc1.NumberInput Index
End Sub
Private Sub cmdPercent_Click()
    BBCalc1.OperationType = Percent
End Sub
Private Sub cmdPi_Click()
    BBCalc1.TrigOp = PI
End Sub
Private Sub cmdPlus_Click()
    BBCalc1.OperationType = Addition
End Sub
Private Sub cmdRan_Click()
    BBCalc1.TrigOp = RandomNum
End Sub
Private Sub cmdRecip_Click()
    BBCalc1.OperationType = Reciprocal
End Sub
Private Sub cmdSin_Click()
    BBCalc1.TrigOp = Sine
End Sub
Private Sub cmdSqr_Click()
    BBCalc1.OperationType = SquareRoot
End Sub
Private Sub cmdSubtract_Click()
    BBCalc1.OperationType = Subtraction
End Sub
Private Sub cmdTangent_Click()
    BBCalc1.TrigOp = Tangent
End Sub
Private Sub cmdValue_Click()
    BBCalc1.ChangeSign
End Sub
Private Sub Form_Load()
    Me.Icon = BBCalc1.CalcIcon
    BBCalc1.DisplayForm Me
    BBCalc1.ShowBox Picture1
End Sub
Private Sub Form_Paint()
    Picture1.SetFocus
End Sub

Private Sub Option1_Click()
    BBCalc1.AngleMode = False
End Sub
Private Sub Option2_Click()
    BBCalc1.AngleMode = True
End Sub

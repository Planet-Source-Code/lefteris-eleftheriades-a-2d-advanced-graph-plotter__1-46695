VERSION 5.00
Begin VB.Form ExpressionBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expression Builder"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   270
      Index           =   37
      Left            =   3315
      TabIndex        =   38
      ToolTipText     =   "The variable X"
      Top             =   735
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Mod "
      Height          =   270
      Index           =   49
      Left            =   2805
      TabIndex        =   50
      ToolTipText     =   "The result of an integer division   Divider Mod Dividor "
      Top             =   735
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abs"
      Height          =   270
      Index           =   48
      Left            =   3150
      TabIndex        =   49
      ToolTipText     =   "Returns the modulus of a number Abs(Number)"
      Top             =   1005
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   270
      Index           =   47
      Left            =   3345
      TabIndex        =   48
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   270
      Index           =   46
      Left            =   2250
      TabIndex        =   47
      Top             =   420
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   270
      Index           =   45
      Left            =   2610
      TabIndex        =   46
      Top             =   420
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   270
      Index           =   44
      Left            =   2985
      TabIndex        =   45
      Top             =   420
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   270
      Index           =   43
      Left            =   1140
      TabIndex        =   44
      Top             =   420
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   270
      Index           =   42
      Left            =   1500
      TabIndex        =   43
      Top             =   420
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   270
      Index           =   41
      Left            =   1875
      TabIndex        =   42
      Top             =   420
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   270
      Index           =   40
      Left            =   30
      TabIndex        =   41
      Top             =   420
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   270
      Index           =   39
      Left            =   390
      TabIndex        =   40
      Top             =   420
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   270
      Index           =   38
      Left            =   765
      TabIndex        =   39
      Top             =   420
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   ")"
      Height          =   270
      Index           =   36
      Left            =   2415
      TabIndex        =   37
      ToolTipText     =   "End Bracket"
      Top             =   735
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "("
      Height          =   270
      Index           =   35
      Left            =   2025
      TabIndex        =   36
      ToolTipText     =   "Begin Bracket"
      Top             =   735
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "^"
      Height          =   270
      Index           =   34
      Left            =   1620
      TabIndex        =   35
      ToolTipText     =   "At the power of"
      Top             =   735
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "/"
      Height          =   270
      Index           =   33
      Left            =   1215
      TabIndex        =   34
      ToolTipText     =   "Division"
      Top             =   735
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "*"
      Height          =   270
      Index           =   32
      Left            =   810
      TabIndex        =   33
      ToolTipText     =   "Multiplication"
      Top             =   735
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   270
      Index           =   31
      Left            =   405
      TabIndex        =   32
      ToolTipText     =   "Substraction"
      Top             =   735
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   270
      Index           =   30
      Left            =   15
      TabIndex        =   31
      ToolTipText     =   "Addition"
      Top             =   735
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sqr"
      Height          =   270
      Index           =   29
      Left            =   2625
      TabIndex        =   30
      ToolTipText     =   "Square root of a number Sqr(Number)"
      Top             =   1005
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log"
      Height          =   270
      Index           =   28
      Left            =   2100
      TabIndex        =   29
      ToolTipText     =   "Logaritm base 10 Log(Number)"
      Top             =   1005
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exp"
      Height          =   270
      Index           =   27
      Left            =   1575
      TabIndex        =   28
      ToolTipText     =   "Expotential at a power Exp(Power)"
      Top             =   1005
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rdeg"
      Height          =   270
      Index           =   26
      Left            =   1050
      TabIndex        =   27
      ToolTipText     =   "Converts Rads to Degrees"
      Top             =   1005
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drad"
      Height          =   270
      Index           =   25
      Left            =   525
      TabIndex        =   26
      ToolTipText     =   "Converts Degrees to Rads"
      Top             =   1005
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pi"
      Height          =   270
      Index           =   24
      Left            =   15
      TabIndex        =   25
      ToolTipText     =   "Ð = 3.1415"
      Top             =   1005
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LogN"
      Height          =   300
      Index           =   23
      Left            =   2760
      TabIndex        =   24
      ToolTipText     =   "Logarithm to base N LogN(Base, Number)"
      Top             =   2700
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HArccotan"
      Height          =   300
      Index           =   22
      Left            =   1845
      TabIndex        =   23
      ToolTipText     =   "Inverse Hyperbolic Cotangent of an angle analogy Harccotan(á)"
      Top             =   2700
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HArccosec"
      Height          =   300
      Index           =   21
      Left            =   930
      TabIndex        =   22
      ToolTipText     =   "Inverse Hyperbolic Cosecant of an angle analogy Harccosec(á)"
      Top             =   2700
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HArcsec"
      Height          =   300
      Index           =   20
      Left            =   15
      TabIndex        =   21
      ToolTipText     =   "Inverse Hyperbolic Secant of an angle analogy Harcsec(á)"
      Top             =   2700
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HArctan"
      Height          =   300
      Index           =   19
      Left            =   2760
      TabIndex        =   20
      ToolTipText     =   "Inverse Hyperbolic Tangent of an angle analogy Harctan(á)"
      Top             =   2415
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HArccos"
      Height          =   300
      Index           =   18
      Left            =   1845
      TabIndex        =   19
      ToolTipText     =   "Inverse Hyperbolic Cosine of an angle analogy Harccos(á)"
      Top             =   2415
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HArcsin"
      Height          =   300
      Index           =   17
      Left            =   930
      TabIndex        =   18
      ToolTipText     =   "Inverse Hyperbolic Sine of an angle analogy Harcsin(á)"
      Top             =   2415
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HCotan"
      Height          =   300
      Index           =   16
      Left            =   15
      TabIndex        =   17
      ToolTipText     =   "Hyperbolic Cotangent of an angle Hcotan(è)"
      Top             =   2415
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HCosec"
      Height          =   300
      Index           =   15
      Left            =   2760
      TabIndex        =   16
      ToolTipText     =   "Hyperbolic Cosecant of an angle Hcosec(è)"
      Top             =   2130
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HSec"
      Height          =   300
      Index           =   14
      Left            =   1845
      TabIndex        =   15
      ToolTipText     =   "Hyperbolic Secant of an angle Hsec(è)"
      Top             =   2130
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HTan"
      Height          =   300
      Index           =   13
      Left            =   930
      TabIndex        =   14
      ToolTipText     =   "Hyperbolic Tangent of an angle Htan(è)"
      Top             =   2130
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HCos"
      Height          =   300
      Index           =   12
      Left            =   15
      TabIndex        =   13
      ToolTipText     =   "Hyperbolic Cosine of an angle Hcos(è)"
      Top             =   2130
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HSin"
      Height          =   300
      Index           =   11
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "Hyperbolic Sine of an angle HSin(è)"
      Top             =   1845
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arccotan"
      Height          =   300
      Index           =   10
      Left            =   1845
      TabIndex        =   11
      ToolTipText     =   "Inverse cotangent of an angle analogy Arccotan(á)"
      Top             =   1845
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arcsec"
      Height          =   300
      Index           =   9
      Left            =   930
      TabIndex        =   10
      ToolTipText     =   "Inverse secant of an angle analogy Arcsec(á)"
      Top             =   1845
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arccos"
      Height          =   300
      Index           =   8
      Left            =   15
      TabIndex        =   9
      ToolTipText     =   "Inverse cosine of an angle analogy Arccos(á)"
      Top             =   1845
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arcsin"
      Height          =   300
      Index           =   7
      Left            =   2760
      TabIndex        =   8
      ToolTipText     =   "Inverse sine of an angle analogy Arcsin(á)"
      Top             =   1560
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Atn"
      Height          =   300
      Index           =   6
      Left            =   1845
      TabIndex        =   7
      ToolTipText     =   "Inverse tangent of an angle analogy Atn(á)"
      Top             =   1560
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sec"
      Height          =   300
      Index           =   5
      Left            =   930
      TabIndex        =   6
      ToolTipText     =   "Secant of an angle Tan(è)"
      Top             =   1560
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cosec"
      Height          =   300
      Index           =   4
      Left            =   15
      TabIndex        =   5
      ToolTipText     =   "Cosecant of an angle Cosec(è)"
      Top             =   1560
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cotan"
      Height          =   300
      Index           =   3
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Cotangent of an angle Cotan(è)"
      Top             =   1275
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tan"
      Height          =   300
      Index           =   2
      Left            =   1845
      TabIndex        =   3
      ToolTipText     =   "Tangent of an angle Tan(è)"
      Top             =   1275
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cos"
      Height          =   300
      Index           =   1
      Left            =   930
      TabIndex        =   2
      ToolTipText     =   "Cosine of an angle Cos(è)"
      Top             =   1275
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sin"
      Height          =   300
      Index           =   0
      Left            =   15
      TabIndex        =   1
      ToolTipText     =   "Sine of an angle Sin(è)"
      Top             =   1275
      Width           =   930
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   3660
   End
End
Attribute VB_Name = "ExpressionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
  If Index < 24 Or (Index >= 27 And Index <= 29) Or Index = 48 Then
     Text1.SelText = Command1(Index).Caption & "("
  Else
     Text1.SelText = Command1(Index).Caption
  End If
  Command1(36).FontBold = cCount("(", Text1.Text) > cCount(")", Text1.Text)
End Sub

Function cCount(Char$, Strng$) As Integer
Dim I%, C%
For I = 1 To Len(Strng$) - (Len(Char) - 1)
   If Mid(Strng$, I, Len(Char)) = Char Then C = C + 1
Next
cCount = C
End Function

Private Sub Form_Unload(Cancel As Integer)
  Plotter.Text1.Text = Text1.Text
End Sub

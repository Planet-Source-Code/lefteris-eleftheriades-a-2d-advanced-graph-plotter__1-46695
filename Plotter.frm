VERSION 5.00
Begin VB.Form Plotter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2D Graph Plotter"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "About"
      Height          =   345
      Left            =   5385
      TabIndex        =   13
      Top             =   7095
      Width           =   720
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Options"
      Height          =   345
      Left            =   4680
      TabIndex        =   12
      Top             =   7095
      Width           =   720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Copy"
      Height          =   345
      Left            =   5385
      TabIndex        =   11
      Top             =   6765
      Width           =   720
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Height          =   930
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   6540
      Width           =   4605
   End
   Begin VB.CommandButton Command3 
      Caption         =   "••"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4290
      TabIndex        =   8
      Top             =   6315
      Width           =   315
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4290
      TabIndex        =   9
      Top             =   6120
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   345
      Left            =   4680
      TabIndex        =   7
      Top             =   6765
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6105
      Left            =   30
      Picture         =   "Plotter.frx":0000
      ScaleHeight     =   403
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   403
      TabIndex        =   2
      Top             =   15
      Width           =   6105
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5430
         TabIndex        =   6
         Text            =   "2Ð"
         Top             =   3030
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2Ð"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2250
         TabIndex        =   5
         Top             =   -15
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2Ð"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   5085
         TabIndex        =   4
         Top             =   3030
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Plot"
      Height          =   510
      Left            =   4680
      TabIndex        =   1
      Top             =   6180
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   375
      TabIndex        =   0
      Text            =   "Sin(X)"
      Top             =   6165
      Width           =   3885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Y="
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   6165
      Width           =   450
   End
End
Attribute VB_Name = "Plotter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Script2 As New ScriptControl 'Microsoft Script Control 1.0
Private Script As Object
Dim Math As New MathFunctions

Private Sub Command1_Click()
On Error GoTo errh
Dim PX As Currency, PY As Currency
Dim Eq As String
Dim CPoints As Long
Dim Lim As Currency
Dim Increasing As Boolean
Dim I As Currency, Y As Currency
  Picture1.ForeColor = DrawColor
  Script.UseSafeSubset = False
  'Script2.UseSafeSubset = True
  If UCase(Text2.Text) = "Ð" Or UCase(Text2.Text) = "PI" Then
     Lim = 3.14159265358979
  ElseIf UCase(Text2.Text) = "2Ð" Or UCase(Text2.Text) = "2PI" Then
     Lim = 3.14159265358979 * 2
  Else
     Lim = Val(Text2.Text)
  End If
  PX = 987654321
  PY = 987654321
  Eq = Text1.Text
  CPoints = 0
  Text3.Text = ""
  'Script.ExecuteStatement "On Error resume next"
  Me.Caption = "2D Graph Plotter [Plotting...]"
  For I = -Lim To Lim Step Lim / 100
      Script.ExecuteStatement "X=" & Replace(I, ",", ".")
      Script.ExecuteStatement "On Error resume next: Y=" & Eq
      Y = Script.Eval("Y")
      If ShowPoints Then
        CPoints = CPoints + 1
        If CPoints = 20 Then
          CPoints = 0
          Picture1.DrawWidth = 3
          Picture1.PSet (201 + I * (200 / Lim), 201 - Y * (200 / Lim))
          Picture1.DrawWidth = 1
        End If
      End If
      
      If PX <> 987654321 Or X = -Lim Then
        '''''''''''If Round(Y, 0) = 0 Then Text3.Text = Text3.Text & "Cuts The X-axis (" & Round(X, 2) & "," & Round(Y, 2) & ")" & vbCrLf
        'Search for turning points
        If Y > PY And Not Increasing Then
           Increasing = True
           Text3.Text = Text3.Text & "Minimum Turning Point at (" & Round(PX, 2) & "," & Round(PY, 2) & ")" & vbCrLf
           If ShowTurningPoints Then
              Picture1.DrawWidth = 4
              Picture1.PSet (201 + I * (200 / Lim), 201 - Y * (200 / Lim)), RGB(255, 100, 0)
              Picture1.DrawWidth = 1
           End If
        ElseIf Y < PY And Increasing Then
           Increasing = False
           Text3.Text = Text3.Text & "Maximum Turning Point at (" & Round(PX, 2) & "," & Round(PY, 2) & ")" & vbCrLf
           If ShowTurningPoints Then
              Picture1.DrawWidth = 4
              Picture1.PSet (201 + I * (200 / Lim), 201 - Y * (200 / Lim)), RGB(0, 100, 255)
              Picture1.DrawWidth = 1
           End If
        End If
      
        Picture1.Line (201 + I * (200 / Lim), 201 - Y * (200 / Lim))-(201 + PX * (200 / Lim), 201 - PY * (200 / Lim))
      End If
      PX = I
      PY = Y
      DoEvents
  Next I
  Me.Caption = "2D Graph Plotter [Ready]"
errh:
  PX = 987654321
  Resume Next
End Sub

Private Sub Command2_Click()
  Picture1.Cls
  Text3.Text = ""
End Sub

Private Sub Command3_Click()
  ExpressionBuilder.Show vbModal
End Sub

Private Sub Command4_Click()
  Select Case Round(Rnd * 5)
    Case 0
          Text1.Text = "Sin(X)"
          Text2.Text = "2Ð"
    Case 1
          Text1.Text = "Tan(Drad*45)*X"
          Text2.Text = "10"
    Case 2
          Text1.Text = "Sin(Drad*X)*180"
          Text2.Text = "180"
    Case 3
          Text1.Text = "Atn(X)"
          Text2.Text = "1"
    Case 4
          Text1.Text = "logN(Exp(1),X)"
          Text2.Text = "8"
    Case 5
          Text1.Text = "ncr(10,abs(X/25.2))"
          Text2.Text = "252"
  End Select
End Sub

Private Sub Command5_Click()
   Clipboard.Clear
   Clipboard.SetData Picture1.Image, vbCFBitmap
End Sub

Private Sub Command6_Click()
  Options.Show
End Sub

Private Sub Command7_Click()
   MsgBox "(c) 2003 Lefteris Eleftherioades"
End Sub

Private Sub Form_Load()
Const P2 = 3.14159265358979 * 2
Const Rdeg = 1 / 57.2957795130823
Dim Y As Currency, I As Currency
  Set Script = CreateObject("ScriptControl")
  Script.Language = "VBScript"
  Script.AddObject "Math", Math, True 'Assign More Math Functions
  DoEvents
  Me.Show
  On Error Resume Next
  Randomize
'Kill the on load animation
If False Then
Me.Caption = "2D Graph Plotter [Loading...]"
DoEvents
again:
  Sh = 10
  F = Round(Rnd * 8.4)
  For Z = -220 To 220
  PX = -220
  PY = 0
  'Picture1.Cls
  DoEvents
  Picture1.Cls
  For I = -P2 To P2 Step P2 / Sh
      Select Case F
         Case 0: Y = 1 / Sin(I + 0.0174 * Z) + Z / 32
         Case 1: Y = Sin(I + 0.0174 * Z) + Z / 32
         Case 2: Y = Tan(I + 0.0174 * Z) + Z / 32
         Case 3: Y = HCotan(I + 0.0174 * Z) + Z / 32
         Case 4: Y = HCosec(I + 0.0174 * Z) + Z / 32
         Case 5: Y = HSec(I + 0.0174 * Z) + Z / 32
         Case 6: Y = HTan(I + 0.0174 * Z) + Z / 32
         Case 7: Y = HCos(I + 0.0174 * Z) + Z / 32
         Case 8: Y = Sin(I + Z / 5)
      End Select
      
      Picture1.Line (201 + I * (200 / (P2)), 201 - Y * (200 / (P2)))-(201 + PX * (200 / (P2)), 201 - PY * (200 / (P2))), RGB((255 - ((Z + 220) * (255 / 440))), 0, (Z + 220) * (255 / 440))
      PX = I
      PY = Y
  Next I
  DoEvents
  Next Z
  Me.Caption = "2D Graph Plotter [Ready]"
  Picture1.Cls
End If
   Label2(0).Visible = True
   Label2(1).Visible = True
 ' GoTo again
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Script = Nothing
  End
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
    Case 1: Text2.Move 160, 0
    Case 0: Text2.Move 361, 202
End Select
Text2.Visible = True
Text2.SetFocus
End Sub

Private Sub Picture1_Click()
   If Text2.Visible Then
      Text2.Visible = False
      Label2(0).Caption = Text2.Text
      Label2(1).Caption = Text2.Text
   End If
End Sub

Private Sub Picture2_Click()
Picture2.BackColor = Rnd * vbWhite
End Sub

Private Sub Text2_Change()
      Label2(0).Caption = Text2.Text
      Label2(1).Caption = Text2.Text
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
      Text2.Visible = False
      Label2(0).Caption = Text2.Text
      Label2(1).Caption = Text2.Text
  End If
End Sub

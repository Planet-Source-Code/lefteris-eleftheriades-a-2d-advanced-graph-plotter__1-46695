VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1875
      TabIndex        =   45
      Top             =   2085
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   390
      Left            =   1140
      TabIndex        =   44
      Top             =   2085
      Width           =   720
   End
   Begin VB.PictureBox Picture4 
      DragMode        =   1  'Automatic
      Height          =   195
      Left            =   3705
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   30
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color Pallet"
      Height          =   1140
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   1755
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   35
         Left            =   75
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   40
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   34
         Left            =   240
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   39
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   33
         Left            =   420
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   38
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   600
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   37
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   31
         Left            =   780
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   36
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   960
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   35
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   1140
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   34
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   1320
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   33
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   1500
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   32
         Top             =   690
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   60
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   31
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   240
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   30
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   405
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   29
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   28
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   765
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   27
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   945
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   26
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   25
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   1305
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   24
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000001&
         Height          =   195
         Index           =   15
         Left            =   1500
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   23
         Top             =   900
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   60
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   22
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000060FF&
         Height          =   195
         Index           =   1
         Left            =   240
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   21
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000080FF&
         Height          =   195
         Index           =   2
         Left            =   420
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   20
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H0000BBFF&
         Height          =   195
         Index           =   3
         Left            =   600
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   19
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   780
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   18
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H0000FFBB&
         Height          =   195
         Index           =   5
         Left            =   960
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   17
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H0000FF00&
         Height          =   195
         Index           =   6
         Left            =   1140
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   16
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00AAFF00&
         Height          =   195
         Index           =   7
         Left            =   1320
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   15
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFF00&
         Height          =   195
         Index           =   8
         Left            =   1500
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   14
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFAA00&
         Height          =   195
         Index           =   9
         Left            =   60
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   13
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF0000&
         Height          =   195
         Index           =   10
         Left            =   240
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   12
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF00AA&
         Height          =   195
         Index           =   11
         Left            =   405
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   11
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF00FF&
         Height          =   195
         Index           =   12
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   10
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00AA00FF&
         Height          =   195
         Index           =   13
         Left            =   765
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   9
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   945
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   8
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         Height          =   195
         Index           =   16
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   7
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00808080&
         Height          =   195
         Index           =   18
         Left            =   1305
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   6
         Top             =   420
         Width           =   165
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000001&
         Height          =   195
         Index           =   20
         Left            =   1500
         ScaleHeight     =   135
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   420
         Width           =   165
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   75
         X2              =   1680
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   75
         X2              =   1680
         Y1              =   645
         Y2              =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custom Color"
      Height          =   2025
      Left            =   1860
      TabIndex        =   0
      Top             =   15
      Width           =   2115
      Begin VB.PictureBox Picture1 
         Height          =   1755
         Left            =   1845
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   2
         Top             =   210
         Width           =   195
      End
      Begin VB.PictureBox Picture2 
         Height          =   1755
         Left            =   60
         ScaleHeight     =   1695
         ScaleWidth      =   1695
         TabIndex        =   1
         Top             =   210
         Width           =   1755
         Begin VB.Image Image1 
            Height          =   1710
            Left            =   0
            Picture         =   "Options.frx":000C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Graph"
      Height          =   915
      Left            =   45
      TabIndex        =   41
      Top             =   1125
      Width           =   1755
      Begin VB.CheckBox Check1 
         Caption         =   "Show Points"
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   43
         Top             =   285
         Width           =   1290
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Turn. Pts"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   42
         Top             =   570
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  ShowPoints = (Check1(0).Value = 1)
  ShowTurningPoints = (Check1(2).Value = 1)
  DrawColor = Picture4.BackColor
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Picture4.BackColor = DrawColor
  Check1(0).Value = -ShowPoints
  Check1(2).Value = -ShowTurningPoints
End Sub


Private Sub Form_Paint()
  Picture1.Cls
  AnalyzeColor DrawColor, Picture1
  Picture1.Refresh
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long
 If Button = 1 Then
  If (X > Picture2.Width - 50) Or (X < 0) Then Exit Sub
  If (Y > Picture2.Height - 50) Or (Y < 0) Then Exit Sub
  Col = Picture2.Point(X, Y)
  If Col = -1 Then Col = 0
  AnalyzeColor Col, Picture1
  Picture4.BackColor = Col
 End If
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long
 If Button = 1 Then
  If (X > Picture1.Width - 50) Or (X < 0) Then Exit Sub
  If (Y > Picture1.Height - 50) Or (Y < 0) Then Exit Sub
  Col = Picture1.Point(X, Y)
  If Col = -1 Then Col = 0
  Picture4.BackColor = Col
 End If
End Sub

Sub AnalyzeColor(Color&, Target As Object, Optional Frequency As Single = 2.75, Optional DraWWidth_ As Long = 3, Optional SM As Long = 56)
Dim I%, r&, G&, B&
   Target.DrawWidth = DraWWidth_
   GetPixel Color&, r&, G&, B& 'Analize Color to RGB
   If (r < 0) Or (G < 0) Or (B < 0) Then r& = r& + 1: G& = G& + 1: B& = B& + 1
   For I% = 1 To 20
    Target.Line (-10, SM - I% * Frequency)-(30, SM - I% * Frequency), _
    RGB(r& + ((255 - r&) / 20 * I%), G& + ((255 - G&) / 20 * I%), B& + ((255 - B&) / 20 * I%))
   Next I%
   For I% = 0 To 20
    Target.Line (-10, I% * Frequency + SM)-(30, I% * Frequency + SM), _
    RGB(r& - (r& / 20 * I%), G& - (G& / 20 * I%), B& - (B& / 20 * I%))
   Next I%
End Sub

Function GetPixel(ByVal Colour&, ByRef red&, ByRef green&, ByRef blue&)
 blue& = Int(Colour& / 65536) ' function to get the blue
 green& = Int((Colour& - (65536 * blue&)) / 256) ' function to get the green
 red& = Colour& - (blue& * 65536) - (green& * 256) ' function to get the red
End Function

Private Sub Picture3_Click(Index As Integer)
  Picture4.BackColor = Picture3(Index).BackColor
End Sub

Private Sub Picture3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
 'If UCase(Source.Name) = "PICTURE4" Or UCase(Source.Name) = "PICTURE6" Then
  Picture3(Index).BackColor = Picture4.BackColor
 'End If
End Sub

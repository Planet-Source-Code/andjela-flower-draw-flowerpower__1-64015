VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "TEN"
      Height          =   885
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   4635
      Width           =   1110
   End
   Begin VB.PictureBox Pic 
      Height          =   4365
      Left            =   45
      ScaleHeight     =   287
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   426
      TabIndex        =   1
      Top             =   75
      Width           =   6450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ONE"
      Height          =   885
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   4635
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub Command1_Click(Index As Integer)
Dim u As Long
Dim r As Long
Dim w As Long

If Index = 1 Then w = 9

For u = 0 To w
    r = Rnd * 200
    Flower Pic, Rnd * Pic.ScaleWidth, Rnd * Pic.ScaleHeight, r, 0, 5, Rnd * 360, RandomBoja, RandomBoja
Next

End Sub


Private Sub Flower(Ob As Variant, X As Long, Y As Long, r As Long, Optional r2 As Long, Optional BrojLatica As Long, Optional Angle As Long, Optional BojaLatica As Long, Optional BojaSredine As Long)
Dim pi As Double
Dim u As Long
Dim tacke(0 To 3) As POINTAPI
Dim sec As Double
Dim A As Double

If BojaLatica = 0 Then BojaLatica = vbYellow
If BojaSredine = 0 Then BojaSredine = vbRed

Ob.ScaleMode = 3
Ob.DrawWidth = 2
Ob.FillStyle = 0
Ob.FillColor = BojaLatica
Ob.ForeColor = BojaLatica

If BrojLatica < 2 Then BrojLatica = 5
If r2 = 0 Then r2 = r / 4

pi = 3.1415
sec = 2 * pi / BrojLatica
A = pi * Angle / 180

For u = 1 To BrojLatica
 tacke(0).X = Cos(sec * u + A) * r2 + X: tacke(0).Y = Sin(sec * u + A) * r2 + Y
 tacke(1).X = Cos(sec * u + A) * r + X: tacke(1).Y = Sin(sec * u + A) * r + Y
 tacke(2).X = Cos(sec * u - sec + A) * r + X: tacke(2).Y = Sin(sec * u - sec + A) * r + Y
 tacke(3).X = Cos(sec * u - sec + A) * r2 + X: tacke(3).Y = Sin(sec * u - sec + A) * r2 + Y

 PolyBezier Ob.hdc, tacke(0), 4
Next u

 FloodFill Ob.hdc, X, Y, BojaLatica
 Ob.FillStyle = 0
 Ob.FillColor = BojaSredine
 Ob.Circle (X, Y), r2

End Sub


Private Function RandomBoja() As Long
Dim r As Long
Dim g As Long
Dim b As Long
 r = Rnd * 255
 g = Rnd * 255
 b = Rnd * 255
 RandomBoja = r + g * 256 + b * 65536
End Function


Private Sub Form_Resize()
Pic.Top = Command1(0).Height
Pic.Left = 0
Pic.Width = Me.Width / 15 - 9
Pic.Height = (Me.Height - Command1(0).Height) / 15 - 84

Command1(0).Top = 0 'Me.Height - Command1(0).Height
Command1(1).Top = 0 'Me.Height - Command1(1).Height
Command1(0).Left = 0
Command1(1).Left = Command1(0).Width




End Sub



Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r As Long

If Button = 1 Then
   r = Rnd * 150 + 20
   Flower Pic, CLng(X), CLng(Y), CLng(r), 0, 5, Rnd * 360, RandomBoja, RandomBoja
End If


End Sub


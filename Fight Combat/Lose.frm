VERSION 5.00
Begin VB.Form Lose 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   7395
      Left            =   120
      Picture         =   "Lose.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "Lose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
End
End If
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width
Image1.Height = Me.Height
End Sub


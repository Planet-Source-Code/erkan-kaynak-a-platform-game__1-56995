VERSION 5.00
Begin VB.Form frmMesaj 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   687
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picmain 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmMesaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public txt As String
Public mfont As Integer

Public Sub YAZI()
Dim i As Integer
Dim sayi
Dim char As String
Dim basY As Integer
Dim boy As Integer
Dim x As Integer
Dim y As Integer

y = Me.ScaleHeight / 2
x = 100

Select Case mfont
    Case 0: basY = 0: boy = 16
    Case 1: basY = 16: boy = 32
End Select

For i = 1 To Len(txt)
    char = Mid(txt, i, 1)
    sayi = Asc(char) - 65
    BitBlt picmain.hDC, x + (i) * boy, y, boy, boy, frmMain.picABC.hDC, sayi * boy, basY, vbSrcCopy
Next
MsgBox (txt)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then Me.Hide
End Sub


Private Sub Form_Paint()
YAZI
End Sub

Private Sub picmain_Paint()
YAZI
End Sub

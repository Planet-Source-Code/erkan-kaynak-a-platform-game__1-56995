VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "erkan_27@yahoo.com"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picABC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   3840
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   735
      ScaleWidth      =   12945
      TabIndex        =   17
      Top             =   10200
      Width           =   12975
   End
   Begin VB.Timer timeTower 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11280
      Top             =   7080
   End
   Begin VB.PictureBox picTower 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   5040
      Picture         =   "Form1.frx":1F92C
      ScaleHeight     =   330
      ScaleWidth      =   660
      TabIndex        =   12
      Top             =   9240
      Width           =   720
   End
   Begin VB.Timer timeAtes 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10680
      Top             =   7080
   End
   Begin VB.PictureBox picAtes 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   360
      Picture         =   "Form1.frx":204C6
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   11
      Top             =   7800
      Width           =   270
   End
   Begin VB.PictureBox picPatlama 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2070
      Left            =   0
      Picture         =   "Form1.frx":20878
      ScaleHeight     =   2010
      ScaleWidth      =   3750
      TabIndex        =   10
      Top             =   9120
      Width           =   3810
   End
   Begin VB.Timer timeDusman 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11280
      Top             =   6000
   End
   Begin VB.PictureBox picDusman 
      AutoRedraw      =   -1  'True
      Height          =   1260
      Left            =   7320
      Picture         =   "Form1.frx":3925A
      ScaleHeight     =   1200
      ScaleWidth      =   1260
      TabIndex        =   9
      Top             =   8520
      Width           =   1320
   End
   Begin VB.PictureBox picAdam 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1125
      Left            =   1080
      ScaleHeight     =   1065
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Timer timeKapi 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10680
      Top             =   6000
   End
   Begin VB.PictureBox picKapi 
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   10800
      Picture         =   "Form1.frx":3B0DC
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   7
      Top             =   3720
      Width           =   915
   End
   Begin VB.PictureBox picAnahtar 
      AutoRedraw      =   -1  'True
      Height          =   330
      Left            =   10680
      Picture         =   "Form1.frx":3BC8A
      ScaleHeight     =   270
      ScaleWidth      =   600
      TabIndex        =   6
      Top             =   3240
      Width           =   660
   End
   Begin VB.PictureBox picCibik 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   11400
      Picture         =   "Form1.frx":3C53C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1680
      Width           =   285
   End
   Begin VB.PictureBox picPanel1 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   11040
      Picture         =   "Form1.frx":3C84E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   300
   End
   Begin VB.Timer timePanel 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10680
      Top             =   6480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   11280
      Top             =   6480
   End
   Begin VB.PictureBox picDuvar 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   10680
      Picture         =   "Form1.frx":3CB90
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   300
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   6000
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   7560
      Width           =   1095
   End
   Begin VB.PictureBox picHero 
      AutoRedraw      =   -1  'True
      Height          =   1260
      Left            =   3000
      Picture         =   "Form1.frx":3CED2
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   7560
      Width           =   2760
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   471
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   695
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      Begin VB.Frame frmSHOW 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   1200
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   7815
         Begin VB.Label label_1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "[SPACE] For Begin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2415
            TabIndex        =   16
            Top             =   2040
            Width           =   3105
         End
         Begin VB.Label label_0 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "kod adý : C.I.P.I.R"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   48
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   1125
            Left            =   0
            TabIndex        =   15
            Top             =   240
            Width           =   7875
         End
      End
   End
   Begin VB.Label labelEnerji 
      BackColor       =   &H00000000&
      Caption         =   "Energy : "
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shEnerji 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   8040
      Top             =   60
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   8040
      Top             =   60
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Bizimle ilgili þeyler
Dim mY As Integer
Dim mX As Integer
Dim dusAcik As Boolean
Dim mYon As Integer
Dim yukariAcik As Integer
Dim DusSayi As Integer
Dim panelUstu As Integer
Dim mAtesSayi As Integer

Dim duvarlar(50) As duvar
Dim paneller(10) As panel
Dim anahtarlar(5) As anahtar
Dim kapilar(5) As kapi
Dim dusmanlar(5) As dusman
Dim patlamalar(50) As patlama
Dim atesler(20) As ates
Dim toplar(10) As tower

Dim mBolum As Integer
Dim introAcik As Integer


Dim blasterFire As String
Dim patlamaFire  As String
Dim dusmanOlFile As String
Dim anahtarFile As String
Dim kapiFile As String
Dim secenekFile As String
Dim aciFile As String

'Sound in Memory variable
Dim SoundInMemory As String

Const YukariMesafe As Integer = 15

Public Sub git(yon As Integer, Optional d As Integer, Optional SabitYon As Integer)
Dim tX As Integer, tY As Integer
Dim basX As Integer, basY As Integer
Dim duvarNo As Integer
Dim anahtarNo As Integer

Static ileri As Integer

Static current_frame

If current_frame = 1 Then
    basX = 46
    current_frame = 0
Else
    basX = 0
    current_frame = 1
End If

If mYon = 4 And yon = 6 Then ileri = 0
If mYon = 6 And yon = 4 Then ileri = 0

tX = mX: tY = mY
Select Case yon
    Case 4:
        basY = 41
        ileri = 5
        
        'Optional Ones...
        If SabitYon > 0 Then
            If SabitYon = 4 Then basY = 41 Else basY = 0
            basX = 0
            mYon = SabitYon
        Else
            mYon = 4
        End If
        If d > 0 Then ileri = d
        If DuvarCarpma(mX - ileri, mY, 0) = False And panelKontrol(mX - ileri, mY, 0) = False And anahtarKontrol(mX - ileri, mY, anahtarNo) = False Then
            mX = mX - ileri
        Else
            If anahtarNo > -1 Then anahtarAcKapa (anahtarNo)
        End If
    Case 6:
        basY = 0
        ileri = 5
        'Optional Ones...
        If SabitYon > 0 Then
            If SabitYon = 4 Then basY = 41 Else basY = 0
            basX = 0
            mYon = SabitYon
        Else
            mYon = 6
        End If
        If d > 0 Then ileri = d
        If DuvarCarpma(mX + ileri, mY, 0) = False And panelKontrol(mX + ileri, mY, 0) = False And anahtarKontrol(mX + ileri, mY, anahtarNo) = False Then
            mX = mX + ileri
        Else
            If anahtarNo > -1 Then anahtarAcKapa (anahtarNo)
        End If
    Case 2:
        If mYon = 6 Then basY = 0 Else basY = 41
        basX = 46
         'Ekranda mýyým?
        If mY > 500 Then
            Enerji 0, 10000
        End If
        
        'Panele Bak
        If panelKontrol(mX, mY + DusSayi + 1, panelUstu) = True Then
            mY = paneller(panelUstu).y - 41
            dusAcik = False
            yukariAcik = 0
            GoTo son
        End If
        
        'Duvara Bak
        If DuvarCarpma(mX, mY + DusSayi + 1, duvarNo) = False Then
            DusSayi = DusSayi + 1
            mY = mY + DusSayi
        Else
            dusAcik = True
            mY = duvarlar(duvarNo).y - 41
        End If

    Case 8:
        If mYon = 6 Then basY = 0 Else basY = 41
        basX = 46
        If yukariAcik > 0 Then
            If DuvarCarpma(mX, mY - yukariAcik, duvarNo) = False And panelKontrol(mX, mY - yukariAcik, 0) = False And anahtarKontrol(mX, mY - yukariAcik, anahtarNo) = False Then
                yukariAcik = yukariAcik - 1
                mY = mY - yukariAcik
            Else
                yukariAcik = 0
                dusAcik = False
                
                If anahtarNo > -1 Then
                    anahtarAcKapa (anahtarNo)
                End If
            End If
        Else
            yukariAcik = 0
            dusAcik = False
        End If
End Select

son:
kapi_kontrol

'BitBlt picMain.hDC, tX, tY, 45, 40, picTemp.hDC, 0, 0, vbSrcCopy
BitBlt picMain.hDC, tX, tY, 45, 40, picAdam.hDC, 0, 0, vbSrcCopy
BitBlt picAdam.hDC, 0, 0, 45, 40, picMain.hDC, mX, mY, vbSrcCopy
BitBlt picMain.hDC, mX, mY, 45, 40, picHero.hDC, basX + 90, basY, vbSrcAnd
BitBlt picMain.hDC, mX, mY, 44, 40, picHero.hDC, basX, basY, vbSrcPaint

    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Enabled = True Then Exit Sub
If mBolum = -2 Then
    picMain.Picture = LoadPicture("")
    intro
    Exit Sub
End If

Select Case KeyCode
    Case 38: If introAcik = 0 Then introMenuSec 8: sndPlaySound secenekFile, SND_ASYNC Or SND_FILENAME
    Case 40: If introAcik = 0 Then introMenuSec 2: sndPlaySound secenekFile, SND_ASYNC Or SND_FILENAME
    Case 13:
            If introAcik = 0 Then
                Dim secenek As Integer
                introMenuSec secenek
                Select Case secenek
                    Case 0: mBolum = 0: bolumOncesiIntro
                    Case 1: About
                    Case 2: End
                End Select
            Else
                introAcik = 0
                intro
            End If
    Case 32: If introAcik = 2 Then introAcik = 0: intro
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 32:
    If frmSHOW.Visible = True Then
        frmSHOW.Visible = False
        
        If Left(label_0.Caption, 1) = "L" Then
            bolum_yukle
        Else
            bolumOncesiIntro
        End If
    Else
        atesAc mYon, mX, mY, 2, -1
        sndPlaySound blasterFire, SND_ASYNC Or SND_FILENAME
    End If
Case 13:
    If introAcik = 1 Then intro
Case 27: Enerji 0, 150
End Select

End Sub

Public Sub bolum_yukle(Optional reset As Boolean)
Dim i As Integer
'Hafýzayý boþaltalým

For i = 0 To UBound(duvarlar): duvarlar(i).acik = False:  Next
For i = 0 To UBound(paneller): paneller(i).acik = False: Next
For i = 0 To UBound(anahtarlar): anahtarlar(i).acik = False: Next
For i = 0 To UBound(kapilar): kapilar(i).acik = False: Next
For i = 0 To UBound(patlamalar): patlamalar(i).acik = False: Next
For i = 0 To UBound(atesler): atesler(i).acik = False: Next
For i = 0 To UBound(toplar): toplar(i).acik = False: toplar(i).atesEtti = False: Next

panelUstu = -1
picMain.Cls

If reset = True Then Exit Sub

Harita_Oku mBolum + 1

labelEnerji.Visible = True
shEnerji.Visible = True
Shape2.Visible = True
Duvarlari_Ciz

timerAcKapa (True)
End Sub

Public Sub Harita_Oku(bolum As Integer)
Dim satir As String
Dim bilesenler
Dim i As Integer
Dim duvarSayi As Integer
Dim panelSayi As Integer
Dim anahtarSayi As Integer
Dim kapiSayi As Integer
Dim dusmanSayi As Integer
Dim topSayi As Integer

Open App.Path & "/levels/level_" & bolum & ".txt" For Input As #1
While Not EOF(1)
    Input #1, satir
    If satir <> "" Then
        bilesenler = Split(satir, "-")
        Select Case bilesenler(0)
            Case "M": 'BENÝM BÝLGÝLERÝM
                    mX = bilesenler(1)
                    mY = bilesenler(2)
            Case "D": 'DUVAR AÇ
                With duvarlar(duvarSayi)
                    .acik = True
                    .yon = bilesenler(1)
                    .x = bilesenler(2)
                    .y = bilesenler(3)
                    .boy = bilesenler(4)
                    .tip = 0
                End With
                duvarSayi = duvarSayi + 1
            Case "P": 'PANEL AÇ
                With paneller(panelSayi)
                    .acik = True
                    .x = bilesenler(2)
                    .y = bilesenler(3)
                    .boy = bilesenler(4)
                    .mesafe = bilesenler(5)
                    .mesafeSayi = bilesenler(5)
                    .yon = bilesenler(6)
                End With
                panelSayi = panelSayi + 1
            Case "K": 'KAZIK AÇ
                With duvarlar(duvarSayi)
                    .acik = True
                    .yon = bilesenler(1)
                    .x = bilesenler(2)
                    .y = bilesenler(3)
                    .boy = bilesenler(4)
                    .tip = 1
                End With
                duvarSayi = duvarSayi + 1
            Case "A": 'ANAHTAR AÇ
                With anahtarlar(anahtarSayi)
                    .acik = True
                    .x = bilesenler(1)
                    .y = bilesenler(2)
                    .durum = False
                End With
                anahtarSayi = anahtarSayi + 1
            Case "Q": 'KAPI ACI
                With kapilar(kapiSayi)
                    .acik = True
                    .x = bilesenler(1)
                    .y = bilesenler(2)
                    .durum = True
                End With
                kapiSayi = kapiSayi + 1
            Case "E": 'DUSMAN AC
                With dusmanlar(dusmanSayi)
                    .acik = True
                    .tip = bilesenler(1)
                    .x = bilesenler(2)
                    .y = bilesenler(3)
                    .mesafe = bilesenler(4)
                    .mesafeSayi = .mesafe
                    .yon = bilesenler(5)
                End With
                dusmanSayi = dusmanSayi + 1
            Case "T": 'TOP AÇ
                With toplar(topSayi)
                    .acik = True
                    .yon = bilesenler(1)
                    .x = bilesenler(2)
                    .y = bilesenler(3)
                End With
                topSayi = topSayi + 1
        End Select
    End If
Wend

Close #1
End Sub

Public Sub Duvarlari_Ciz()
Dim d As Integer
Dim i As Integer
Dim y As Integer, x As Integer

For d = 0 To UBound(duvarlar)
    If duvarlar(d).acik = True Then
        y = duvarlar(d).y
        x = duvarlar(d).x
        If duvarlar(d).yon = 0 Then
            Select Case duvarlar(d).tip
            Case 0:
                For i = 0 To CInt(duvarlar(d).boy / 16) - 1
                    BitBlt picMain.hDC, x + 16 * i, y, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
                Next
            Case 1:
                For i = 0 To CInt(duvarlar(d).boy / 16) - 1
                    BitBlt picMain.hDC, x + 16 * i, y, 16, 16, picCibik.hDC, 0, 0, vbSrcCopy
                Next
            End Select
        Else
            For i = 0 To CInt(duvarlar(d).boy / 16) - 1
                BitBlt picMain.hDC, x, y + 16 * i, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
            Next
        End If
    End If
Next

'Draw Keys
For d = 0 To UBound(anahtarlar)
    If anahtarlar(d).acik = True Then
        anahtarAcKapa d, True
    End If
Next

'Draw Towers
For d = 0 To UBound(toplar)
    If toplar(d).acik = True Then
        Select Case toplar(d).yon
            Case 4: BitBlt picMain.hDC, toplar(d).x, toplar(d).y, 22, 22, picTower.hDC, 0, 0, vbSrcCopy
            Case 6: BitBlt picMain.hDC, toplar(d).x, toplar(d).y, 22, 22, picTower.hDC, 22, 0, vbSrcCopy
        End Select
    End If
Next
End Sub

Private Sub Form_Load()
'blasterFire = App.Path & "/sounds/blaster_3.wav"
'patlamaFire = App.Path & "/sounds/GUNSHOT.wav"
'dusmanOlFile = App.Path & "/sounds/fire2.wav"
blasterFire = App.Path & "/sounds/fire2.wav"
dusmanOlFile = App.Path & "/sounds/GUNSHOT.wav"
patlamaFire = App.Path & "/sounds/GUNSHOT.wav"
anahtarFile = App.Path & "/sounds/hit.wav"
kapiFile = App.Path & "/sounds/kapi.wav"
secenekFile = App.Path & "/sounds/secenek.wav"
aciFile = App.Path & "/sounds/aci.wav"
intro
End Sub

Private Sub picmain_Paint()
If frmSHOW.Visible = False And Timer1.Enabled = False And mBolum <> -2 Then
    intro
Else
    If mBolum <> -2 Then Duvarlari_Ciz
End If
End Sub

Private Sub timeAtes_Timer()
atesHareket 2
End Sub

Private Sub timeDusman_Timer()
dusmanHareket
patlamaHareket
End Sub



Private Sub timeKapi_Timer()
Static current_frame As Integer
Dim k As Integer

If current_frame = 2 Then current_frame = 0 Else current_frame = current_frame + 1

For k = 0 To UBound(kapilar)
    If kapilar(k).acik = True Then
        BitBlt picMain.hDC, kapilar(k).x, kapilar(k).y, 19, 17, picKapi.hDC, 19 * current_frame, 0, vbSrcCopy
    End If
Next
End Sub

Private Sub timePanel_Timer()
    panelHareket
    topAtes
End Sub

Private Sub Timer1_Timer()

'Run only once
If mBolum = -1 Then
    picMain.Cls
    Duvarlari_Ciz
    BitBlt picAdam.hDC, -50, -50, 45, 40, picMain.hDC, mX, mY, vbSrcCopy
    git 6
    mBolum = 0
End If

If (GetKeyState(vbKeyLeft) And KEY_DOWN) Then git 4
If (GetKeyState(vbKeyRight) And KEY_DOWN) Then git 6

If (GetKeyState(vbKeyUp) And KEY_DOWN) Then
    If yukariAcik = 0 And dusAcik = False Then yukariAcik = YukariMesafe: dusAcik = False
End If

If yukariAcik > 0 And dusAcik = False Then git 8
If dusAcik = True And yukariAcik = 0 Then git 2
If yukariAcik = 0 Then dusAcik = Not altKontrol(mX, mY + 1)
    


If dusAcik = False Then DusSayi = 0
End Sub

Public Function DuvarCarpma(x As Integer, y As Integer, duvarNo As Integer, Optional tip As Integer) As Boolean
Dim d As Integer
Dim carpma As Boolean
Dim boy As Integer
Dim En As Integer

'Tip :
'   bos = biz
'   1   = Enemy
'   2   = Fire
'   3   = Tower Fire

Select Case tip
    Case 1: boy = 39: En = 41
    Case 2: boy = 20: En = 20
    Case 3: boy = 6: En = 12
    Case Else: boy = 40: En = 45
End Select
    

For d = 0 To UBound(duvarlar)
    If duvarlar(d).acik = True Then
        carpma = True
        Select Case duvarlar(d).yon
        Case 0:
                If x > duvarlar(d).x + duvarlar(d).boy Then carpma = False
                If x + En < duvarlar(d).x Then carpma = False
                If y > duvarlar(d).y + 16 Then carpma = False
                If y + boy < duvarlar(d).y Then carpma = False
        Case 1:
                If x > duvarlar(d).x + 16 Then carpma = False
                If x + En < duvarlar(d).x Then carpma = False
                If y > duvarlar(d).y + duvarlar(d).boy Then carpma = False
                If y + boy < duvarlar(d).y Then carpma = False
        End Select
        If carpma = True Then
            DuvarCarpma = True
            duvarNo = d
            Exit Function
        End If
    End If
Next

DuvarCarpma = False
End Function

Public Function altKontrol(x As Integer, y As Integer) As Boolean
Dim duvarNo As Integer

        'On The Screen?
        If mY > 500 Then Enerji 0, 10000
        
If panelKontrol(x, y, panelUstu) = True Then altKontrol = True: Exit Function
If anahtarKontrol(x, y, 0) = True Then altKontrol = True: Exit Function
If DuvarCarpma(x, y, duvarNo) = True Then
    altKontrol = True
    If duvarlar(duvarNo).tip = 1 Then Enerji 0, 1
    Exit Function
End If

altKontrol = False
End Function

Public Sub panelHareket()
Dim p As Integer
Dim tX As Integer, tY As Integer, i As Integer
Dim gitYon As Integer
Dim eMy As Integer

    For p = 0 To UBound(paneller)
        If paneller(p).acik = True Then
            tX = paneller(p).x
            tY = paneller(p).y
            Select Case paneller(p).yon
                Case 4:
                        If paneller(p).mesafeSayi < 0 Then
                            paneller(p).mesafeSayi = paneller(p).mesafe
                            paneller(p).yon = 6
                        Else
                            paneller(p).x = tX - 1
                            paneller(p).mesafeSayi = paneller(p).mesafeSayi - 1
                            tX = tX + paneller(p).boy
                        End If
                        
                        If panelUstu = p Then
                             git 4, 1, mYon
                        Else
                            If panelCarpma(mX, mY, p) = True Then
                                git 4: git 4: git 4
                            End If
                        End If
                Case 6:
                        If paneller(p).mesafeSayi < 0 Then
                            paneller(p).mesafeSayi = paneller(p).mesafe
                            paneller(p).yon = 4
                        Else
                            paneller(p).x = tX + 1
                            paneller(p).mesafeSayi = paneller(p).mesafeSayi - 1
                        End If
                        
                        If panelUstu = p Then
                            git 6, 1, mYon
                        Else
                            If panelCarpma(mX, mY, p) = True Then
                                git 6: git 6: git 6
                            End If
                        End If
                Case 8:
                       If paneller(p).mesafeSayi < 0 Then
                            paneller(p).mesafeSayi = paneller(p).mesafe
                            paneller(p).yon = 2
                        Else
                            If panelUstu = p Then
                                
                                If DuvarCarpma(mX, mY - 16, 0, 0) = True Then
                                    Enerji 0, 2
                                Else
                                    paneller(p).y = tY - 1
                                    paneller(p).mesafeSayi = paneller(p).mesafeSayi - 1
                                End If
                                git 2, 16, mYon
                            Else
                                paneller(p).y = tY - 1
                                paneller(p).mesafeSayi = paneller(p).mesafeSayi - 1
                            End If
                        End If
                Case 2:
                       If paneller(p).mesafeSayi < 0 Then
                            paneller(p).mesafeSayi = paneller(p).mesafe
                            paneller(p).yon = 8
                       Else
                            If panelCarpma(mX, mY - 1, p) = True Then
                                   Enerji 0, 2
                            Else
                                paneller(p).y = tY + 1
                                paneller(p).mesafeSayi = paneller(p).mesafeSayi - 1
                            End If
                       End If
                       
            End Select
            
            
            'Gideceðim Yeri hafýzaya almalýyým
            'Select Case paneller(p).yon
            'Case 6:
            '    BitBlt picTemp.hDC, 0, 0, 16, 16, picMain.hDC, paneller(p).x + paneller(p).boy + 16, paneller(p).y, vbSrcCopy
            'Case 4:
            '    BitBlt picTemp.hDC, 0, 0, 16, 16, picMain.hDC, paneller(p).x - 16, paneller(p).y, vbSrcCopy
            'End Select

            BitBlt picMain.hDC, tX, tY, 16, 16, picTemp.hDC, 0, 0, vbSrcCopy
            For i = 0 To paneller(p).boy / 16
                If paneller(p).yon = 2 Or paneller(p).yon = 8 Then BitBlt picMain.hDC, tX + i * 16, tY, 16, 16, picTemp.hDC, 0, 0, vbSrcCopy
                BitBlt picMain.hDC, paneller(p).x + i * 16, paneller(p).y, paneller(p).boy, 16, picPanel1.hDC, 0, 0, vbSrcCopy
            Next
            
            
        End If
    Next

End Sub

Public Function panelCarpma(x As Integer, y As Integer, panelNo As Integer) As Boolean
Dim carpma As Boolean

carpma = True

    If x > paneller(panelNo).x + paneller(panelNo).boy + 16 Then carpma = False
    If x + 45 < paneller(panelNo).x Then carpma = False
    If y > paneller(panelNo).y + 16 Then carpma = False
    If y + 40 < paneller(panelNo).y Then carpma = False
    
panelCarpma = carpma
End Function

Public Function panelKontrol(x As Integer, y As Integer, panelNo As Integer) As Boolean

Dim p As Integer
panelNo = -1

    For p = 0 To UBound(paneller)
        If paneller(p).acik = True Then
            If panelCarpma(x, y, p) = True Then
                panelKontrol = True
                panelNo = p
                Exit Function
            End If
        End If
    Next

panelKontrol = False
End Function

Public Sub Enerji(islem As Integer, miktar As Integer)
If islem = 0 Then
    If shEnerji.Width - miktar > 0 Then
        shEnerji.Width = shEnerji.Width - miktar
    Else
        shEnerji.Width = 1
    End If
    If shEnerji.Width <= 1 Then
        dusAcik = False
        yukariAcik = 1
        patlamaAc 4, mX, mY
        git 4, 10000, 4
    End If
    sndPlaySound aciFile, SND_ASYNC Or SND_FILENAME
Else
    If shEnerji.Width + miktar <= 150 Then
        shEnerji.Width = shEnerji.Width + miktar
    Else
        shEnerji.Width = 150
    End If
End If

End Sub

Public Sub anahtarAcKapa(anahtarNo As Integer, Optional sabit As Boolean)
Dim a As Integer
Dim hepsiAcik As Boolean

    If sabit = False Then
        anahtarlar(anahtarNo).durum = Not anahtarlar(anahtarNo).durum
        sndPlaySound anahtarFile, SND_ASYNC Or SND_FILENAME
    End If
    'control the keys
    hepsiAcik = True
    For a = 0 To UBound(anahtarlar)
            If anahtarlar(a).acik = True And anahtarlar(a).durum = False Then hepsiAcik = False: Exit For
    Next
    If hepsiAcik = True Then
        If timeKapi.Enabled = False Then sndPlaySound kapiFile, SND_ASYNC Or SND_FILENAME
        timeKapi.Enabled = True
        
    Else
        timeKapi.Enabled = False
    End If
    
    If anahtarlar(anahtarNo).durum = False Then
        BitBlt picMain.hDC, anahtarlar(anahtarNo).x, anahtarlar(anahtarNo).y, 20, 18, picAnahtar.hDC, 0, 0, vbSrcCopy
    Else
        BitBlt picMain.hDC, anahtarlar(anahtarNo).x, anahtarlar(anahtarNo).y, 20, 18, picAnahtar.hDC, 20, 0, vbSrcCopy
    End If
End Sub

Public Function anahtarKontrol(x As Integer, y As Integer, anahtarNo As Integer) As Boolean
Dim carpma As Boolean
Dim a As Integer

anahtarNo = -1

For a = 0 To UBound(anahtarlar)
    If anahtarlar(a).acik = True Then
        carpma = True
        If x > anahtarlar(a).x + 20 Then carpma = False
        If x + 45 < anahtarlar(a).x Then carpma = False
        If y > anahtarlar(a).y + 18 Then carpma = False
        If y + 40 < anahtarlar(a).y Then carpma = False
        
        If carpma = True Then
            anahtarKontrol = True
            anahtarNo = a
            Exit Function
        End If
    End If
Next

anahtarKontrol = False
End Function

Public Sub kapiAcKapa(kapiNo As Integer, Optional sabit As Boolean)

    If sabit = False Then kapilar(kapiNo).durum = Not kapilar(kapiNo).durum
    
    timeKapi.Enabled = kapilar(kapiNo).durum
    sndPlaySound kapiFile, SND_ASYNC Or SND_FILENAME
End Sub

Public Sub dusmanHareket()
Dim d As Integer
Dim tX As Integer, tY As Integer
Dim basX As Integer, basY As Integer
Static current_frame(5) As Integer
Dim carpma As Boolean

    For d = 0 To UBound(dusmanlar)
        If dusmanlar(d).acik = True Then
            tX = dusmanlar(d).x
            tY = dusmanlar(d).y
            Select Case dusmanlar(d).tip
                Case 0: 'ENEMY
                        Select Case dusmanlar(d).yon
                            Case 6:
                                    basY = 40
                                    If dusmanlar(d).mesafe > 0 Then
                                        dusmanlar(d).x = tX + 10
                                        dusmanlar(d).mesafeSayi = dusmanlar(d).mesafeSayi - 10
                                        If dusmanlar(d).mesafeSayi <= 0 Then
                                            dusmanlar(d).mesafeSayi = dusmanlar(d).mesafe
                                            dusmanlar(d).yon = 4
                                        End If
                                    Else
                                        If DuvarCarpma(tX + 10, tY, 0, 1) = False Then
                                            dusmanlar(d).x = tX + 10
                                        Else
                                            dusmanlar(d).yon = 4
                                        End If
                                    End If
                             Case 4:
                                    basY = 0
                                    If dusmanlar(d).mesafe > 0 Then
                                        dusmanlar(d).x = tX - 10
                                        dusmanlar(d).mesafeSayi = dusmanlar(d).mesafeSayi - 10
                                        If dusmanlar(d).mesafeSayi <= 0 Then
                                            dusmanlar(d).mesafeSayi = dusmanlar(d).mesafe
                                            dusmanlar(d).yon = 6
                                        End If
                                    Else
                                        If DuvarCarpma(tX - 10, tY, 0, 1) = False Then
                                            dusmanlar(d).x = tX - 10
                                        Else
                                            dusmanlar(d).yon = 6
                                        End If
                                    End If
                        End Select
                        
            End Select
            
            If current_frame(d) = 0 Then
                basX = 0
                current_frame(d) = 1
            Else
                basX = 41
                current_frame(d) = 0
            End If
            
            BitBlt picMain.hDC, tX, tY, 39, 41, picTemp.hDC, 0, 0, vbSrcCopy
            BitBlt picMain.hDC, dusmanlar(d).x, dusmanlar(d).y, 39, 41, picDusman.hDC, basX, basY, vbSrcCopy
            
            If dusmanKontrol(mX, mY, d) = True Then
                dusmanOl (d)
                Enerji 0, 50
            End If
        End If
    Next
End Sub

Public Sub patlamaAc(tip As Integer, x As Integer, y As Integer)
Dim p As Integer

For p = 0 To UBound(patlamalar)
    If patlamalar(p).acik = False Then
            With patlamalar(p)
                .acik = True
                .tip = tip
                .x = x
                .y = y
                .current_frame = 0
                Select Case tip
                    Case 1: .frame = 6: sndPlaySound dusmanOlFile, SND_ASYNC Or SND_FILENAME
                    Case 2: .frame = 5: sndPlaySound patlamaFire, SND_ASYNC Or SND_FILENAME
                    Case 3: .frame = 5
                    Case 4: .frame = 5
                End Select
                
                Exit Sub
            End With
    End If
Next

End Sub


Public Sub patlamaHareket()
Dim p As Integer
Dim basX As Integer, basY As Integer
Dim En As Integer
Dim boy As Integer

For p = 0 To UBound(patlamalar)
    If patlamalar(p).acik = True Then
        Select Case patlamalar(p).tip
        Case 1: basY = 0: En = 40: boy = 35
        Case 2: basY = 36: En = 32: boy = 32 'My Filre
        Case 3: basY = 68: En = 24: boy = 18 'Tower Fire
        Case 4: basY = 86: En = 50: boy = 46 'We Die!
        End Select
        
        With patlamalar(p)
            .current_frame = .current_frame + 1
            If .current_frame >= .frame Then
                .acik = False
                'Clean Screen
                BitBlt picMain.hDC, .x, .y, En, boy, picTemp.hDC, 0, 0, vbSrcCopy
                BitBlt picAdam.hDC, 0, 0, 40, 45, picTemp.hDC, 0, 0, vbSrcCopy
                Duvarlari_Ciz
                If patlamalar(p).tip = 4 Then
                    game_over
                End If
            Else
                BitBlt picMain.hDC, .x, .y, En, boy, picPatlama.hDC, En * .current_frame, basY, vbSrcCopy
            End If
        End With
    End If
Next
End Sub

Public Function dusmanKontrol(x As Integer, y As Integer, dusmanNo As Integer, Optional tip As Boolean) As Boolean
Dim d As Integer
Dim boy As Integer, En As Integer
Dim carpma As Boolean

Select Case tip
    Case 2: En = 20: boy = 20
    Case Else: En = 45: boy = 40
End Select

For d = 0 To UBound(dusmanlar)
    If dusmanlar(d).acik = True Then
        carpma = True
        If x > dusmanlar(d).x + 40 Then carpma = False
        If x + En < dusmanlar(d).x Then carpma = False
        If y > dusmanlar(d).y + 37 Then carpma = False
        If y + boy < dusmanlar(d).y Then carpma = False
        
        If carpma = True Then
            dusmanKontrol = True
            dusmanNo = d
            Exit Function
        End If
    End If
Next

dusmanKontrol = False
End Function


Public Sub atesAc(yon As Integer, x As Integer, y As Integer, tip As Integer, gonderen As Integer)
Dim a As Integer

    For a = 0 To UBound(atesler)
        With atesler(a)
            If .acik = False Then
                .acik = True
                .tip = tip
                .gonderen = gonderen
                If yon = 6 Then
                    .x = x + 43
                Else
                    .x = x - 20
                End If
                .y = y + 15
                .yon = yon
                Exit Sub
            End If
        End With
    Next

End Sub

Public Sub atesHareket(Optional tip As Integer)
Dim a As Integer
Dim tX As Integer
Dim tY As Integer
Dim duvarNo As Integer
Dim dusmanNo As Integer
Dim ben As Boolean

'Ateþ Resmini Seçmek Ýçin
Dim basY As Integer
Dim En As Integer
Dim boy As Integer
Dim atesTip As Integer

' Ates Tip
'
'   My Fire          :   2
'   Tower Fire       :   3

dusmanNo = -1
    For a = 0 To UBound(atesler)
        With atesler(a)

            If .acik = True And .tip = tip Then
                tX = .x
                tY = .y
            
                If .tip = 2 Then
                    basY = 0
                    En = 14
                    boy = 14
                ElseIf .tip = 3 Then
                    basY = 14
                    En = 12
                    boy = 6
                End If
                
                    
                Select Case .yon
                    Case 6:
                        If DuvarCarpma(.x + 20, .y, duvarNo, tip) = False And anahtarKontrol(.x, .y, 0) = False And dusmanKontrol(.x, .y, dusmanNo, False) = False And banaCarpma(.x + 20, .y, tip, ben) = False Then
                            .x = tX + 10
                        Else
                            .acik = False
                            BitBlt picMain.hDC, .x, tY, En, boy, picTemp.hDC, 0, 0, vbSrcCopy
                            patlamaAc tip, .x, .y
                            If dusmanNo <> -1 Then
                                dusmanOl (dusmanNo)
                            End If
                            
                            If ben = True Then
                                Enerji 0, 50
                                git 6, 20, mYon
                            End If
                            
                            If .gonderen > -1 Then toplar(.gonderen).atesEtti = False
                        End If
                    Case 4:
                        If DuvarCarpma(.x - 10, .y, duvarNo, tip) = False And anahtarKontrol(.x - 10, .y, 0) = False And dusmanKontrol(.x, .y, dusmanNo, False) = False And banaCarpma(.x, .y, tip, ben) = False Then
                            .x = tX - 10
                        Else
                            .acik = False
                            BitBlt picMain.hDC, tX, tY, En, boy, picTemp.hDC, 0, 0, vbSrcCopy
                            patlamaAc tip, .x, .y
                            If dusmanNo <> -1 Then
                                dusmanOl (dusmanNo)
                            End If
                            
                            If ben = True Then
                                Enerji 0, 50
                                git 4, 20, mYon
                            End If
                            
                            If .gonderen > -1 Then toplar(.gonderen).atesEtti = False
                        End If
                End Select
                
                If .acik = True Then
                    BitBlt picMain.hDC, tX, tY, En, boy, picTemp.hDC, 0, basY, vbSrcCopy
                    BitBlt picMain.hDC, .x, .y, En, boy, picAtes.hDC, 0, basY, vbSrcCopy
                End If
            End If
            

        End With
    Next
End Sub

Public Sub dusmanOl(d As Integer)
                dusmanlar(d).acik = False
                BitBlt picMain.hDC, dusmanlar(d).x, dusmanlar(d).y, 39, 41, picTemp.hDC, 0, 0, vbSrcCopy
                patlamaAc 1, dusmanlar(d).x, dusmanlar(d).y
End Sub

Public Sub topAtes()
Dim t As Integer
Dim acikSayi As Integer
Dim topSayi As Integer

'Açýk olan 1 li atesleri sayacaðýz.
'toplardan fazlaysa yapma bekle

For t = 0 To UBound(atesler)
    If atesler(t).acik = True And atesler(t).tip = 1 Then acikSayi = acikSayi + 1
Next

For t = 0 To UBound(toplar)
    If toplar(t).acik = True Then topSayi = topSayi + 1
Next

If acikSayi >= topSayi Then
    Exit Sub
End If

For t = 0 To UBound(toplar)
    With toplar(t)
        If .acik = True And .atesEtti = False Then
            atesAc .yon, .x + 1, .y - 10, 3, t
            .atesEtti = True
        End If
    End With
Next
End Sub

Private Sub timeTower_Timer()
atesHareket 3
End Sub

Public Function banaCarpma(x As Integer, y As Integer, tip As Integer, ben As Boolean) As Boolean
Dim En As Integer
Dim boy As Integer

Select Case tip
    Case 3: 'TOWER ATES
        En = 12
        boy = 6
End Select

banaCarpma = True
If mX > x + En Then banaCarpma = False
If mX + 45 < x Then banaCarpma = False
If mY > y + boy Then banaCarpma = False
If mY + 40 < y Then banaCarpma = False

ben = banaCarpma

End Function

Public Sub intro()
Dim i As Integer

    
    bolum_yukle True
    dusAcik = False
    yukariAcik = 0
    shEnerji.Width = 150
    
    mBolum = -1
    For i = 0 To 672 Step 16
        BitBlt picMain.hDC, i, 0, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
        BitBlt picMain.hDC, i, 448, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
    Next
    
    YAZI "C.I.P.I.R.", 180, 80, 1
    YAZI "NEW GAME", 200, 150, 0
    YAZI "ABOUT", 200, 200, 0
    YAZI "EXIT", 200, 250, 0
    
    introMenuSec 2
    introMenuSec 8
End Sub

Public Sub game_over(Optional win As Boolean)
Dim i
    picMain.Cls
    
    labelEnerji.Visible = False
    shEnerji.Visible = False
    Shape2.Visible = False
    
    For i = 0 To 672 Step 16
        BitBlt picMain.hDC, i, 0, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
        BitBlt picMain.hDC, i, 448, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
    Next
    
    timerAcKapa False

    If win = True Then
        YAZI "YOU WIN", 160, 200, 1
        introAcik = 2
    Else
        YAZI "GAME OVER", 150, 200, 1
        introAcik = 2
    End If
End Sub
Public Sub kapi_kontrol()
Dim k As Integer
Dim carpma As Boolean

If timeKapi.Enabled = False Then Exit Sub

For k = 0 To UBound(kapilar)
    If kapilar(k).acik = True Then
        carpma = True
                
        If mX > kapilar(k).x + 19 Then carpma = False
        If mX + 45 < kapilar(k).x Then carpma = False
        If mY > kapilar(k).y + 17 Then carpma = False
        If mY + 40 < kapilar(k).y Then carpma = False
        
        If carpma = True Then
            mBolum = mBolum + 1
            'CONTROL THE END!!!!!
            If mBolum = 6 Then
                dusAcik = False
                yukariAcik = 1
                mBolum = 0
                game_over True
                Exit Sub
            End If
            '**
            '
            bolumOncesiIntro
            'bolum_yukle
        End If
    End If
    
Next
End Sub

Public Sub bolumOncesiIntro()
Dim i
    picMain.Cls
    timerAcKapa (False)
    
    bolum_yukle True
    
    frmSHOW.Visible = True
    
    label_0.Caption = "Level " & (mBolum + 1)
    
    labelEnerji.Visible = False
    shEnerji.Visible = False
    Shape2.Visible = False
    For i = 0 To 672 Step 16
        BitBlt picMain.hDC, i, 0, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
        BitBlt picMain.hDC, i, 448, 16, 16, picDuvar.hDC, 0, 0, vbSrcCopy
    Next
End Sub



Public Sub YAZI(txt As String, x As Integer, y As Integer, font As Integer)
Dim i As Integer
Dim sayi
Dim char As String
Dim basY As Integer
Dim boy As Integer

Select Case font
    Case 0: basY = 0: boy = 16
    Case 1: basY = 16: boy = 32
End Select

For i = 1 To Len(txt)
    char = Mid(txt, i, 1)
    sayi = Asc(char) - 65
    BitBlt picMain.hDC, x + (i) * boy, y, boy, boy, picABC.hDC, sayi * boy, basY, vbSrcCopy
Next
End Sub

Public Sub timerAcKapa(durum As Boolean)
Timer1.Enabled = durum
timeDusman.Enabled = durum
timePanel.Enabled = durum
timeTower.Enabled = durum
timeAtes.Enabled = durum
End Sub

Public Sub introMenuSec(yon As Integer)
Static secenek As Integer
Dim i As Integer
Dim eSecenek As Integer

'Nerede olduðumuzu göndrelim
If yon = 0 Then yon = secenek: Exit Sub

eSecenek = secenek

If yon = 8 Then
    If secenek > 0 Then secenek = secenek - 1 Else secenek = 2
End If

If yon = 2 Then
    If secenek < 2 Then secenek = secenek + 1 Else secenek = 0
End If

picMain.Line (200, 140 + eSecenek * 50)-(400, 175 + eSecenek * 50), vbBlack, B
picMain.Line (200, 140 + secenek * 50)-(400, 175 + secenek * 50), vbYellow, B
End Sub


Public Sub About()
Dim i As Integer
    picMain.Cls
    introAcik = 0
    mBolum = -2
    picMain.Picture = LoadPicture(App.Path & "/about.gif")
    
End Sub

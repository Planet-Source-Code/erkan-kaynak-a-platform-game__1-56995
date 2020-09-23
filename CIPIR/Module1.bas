Attribute VB_Name = "Module1"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long
                 
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'''sndPlaySound Constants
Public Const SND_ALIAS = &H10000
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10

Public Const KEY_TOGGLED As Integer = &H1

Public Const KEY_DOWN As Integer = &H1000

Type duvar
    acik As Boolean
    yon As Integer
    x As Integer
    y As Integer
    boy As Integer
    tip As Integer '0- Duvar, 1- Kazýk
End Type
    
Type panel
    acik As Boolean
    yon As Integer
    x As Integer
    y As Integer
    boy As Integer
    mesafe As Integer
    mesafeSayi As Integer
End Type

Type anahtar
    acik  As Boolean
    x As Integer
    y As Integer
    durum As Boolean
End Type

Type kapi
    acik As Boolean
    x As Integer
    y As Integer
    durum As Boolean
End Type

Type dusman
    acik As Boolean
    tip As Integer
    x As Integer
    y As Integer
    yon As Integer
    mesafe As Integer
    mesafeSayi As Integer
End Type

Type patlama
    acik As Boolean
    tip As Integer
    x As Integer
    y As Integer
    frame As Integer
    current_frame As Integer
End Type
    
Type ates
    acik As Boolean
    yon As Integer
    x As Integer
    y As Integer
    tip As Integer
    gonderen As Integer
End Type

Type tower
    acik As Boolean
    yon As Integer
    x As Integer
    y As Integer
    atesEtti As Boolean
End Type

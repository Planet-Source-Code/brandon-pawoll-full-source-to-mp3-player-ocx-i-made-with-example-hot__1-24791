VERSION 5.00
Begin VB.UserControl MP3Play 
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   ScaleHeight     =   1035
   ScaleWidth      =   990
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "UserControl1.ctx":0312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "MP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Total As String * 255
Private Pos As String * 255

Private Sub UserControl_Initialize()
Paused = False
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Public Sub LoadMP3(MP3 As String)
mciSendString "close mpeg", 0, 0, 0
mciSendString "open " & MP3 & " type MPEGVideo Alias mpeg", 0&, 0&, 0&
End Sub

Public Sub PlayMP3()
mciSendString "play mpeg", 0, 0, 0
End Sub

Public Sub StopMP3()
mciSendString "close mpeg", 0, 0, 0
End Sub

Public Function GetPercent() As Integer
Dim Ttl As String
Dim Ps As String
Dim Percent As String
mciSendString "set mpeg time format frames", Total, 255, 0&
mciSendString "status mpeg length", Total, 255, 0&
mciSendString "status mpeg position", Pos, 255, 0&
Ttl = Val(Total)
Ps = Val(Pos)
If Ttl = "0" Then Exit Function
Percent = Ps * 100 / Ttl
GetPercent = CInt(Percent)
End Function

Public Sub Pause()
mciSendString "pause mpeg", 0, 0, 0
End Sub

Public Sub UnPause()
mciSendString "resume mpeg", 0, 0, 0
End Sub


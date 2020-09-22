VERSION 5.00
Object = "{D20FB24D-7228-11D5-AC90-000103279643}#2.0#0"; "Mp3Player.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MP3 Player Example using mp3player.ocx"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   600
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UnPause"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin MP3Player.MP3Play MP3Play1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "0 %"
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Percent Through MP3:"
      Height          =   195
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "Powered by: Brandon's MP3 Play .OCX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Ans As String
Ans = InputBox("Type the path to the mp3 you want to play")
MP3Play1.LoadMP3 (Ans)
End Sub

Private Sub Command2_Click()
MP3Play1.PlayMP3
Timer1 = True
End Sub

Private Sub Command3_Click()
MP3Play1.Pause
End Sub

Private Sub Command4_Click()
MP3Play1.UnPause
End Sub

Private Sub Command5_Click()
MP3Play1.StopMP3
End Sub

Private Sub Timer1_Timer()
Label3 = MP3Play1.GetPercent & " %"
End Sub

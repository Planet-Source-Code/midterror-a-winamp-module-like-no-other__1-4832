VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Winamp Module Sample"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSongInfo 
      Caption         =   "Song Info"
      Height          =   1575
      Left            =   2880
      TabIndex        =   16
      Top             =   1560
      Width           =   1815
      Begin VB.Label lblChannels 
         Caption         =   "Channels: "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblBitrate 
         Caption         =   "Bitrate: "
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblSampleRate 
         Caption         =   "Samplerate: "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.VScrollBar ScrSongNum 
      Height          =   375
      Left            =   1320
      Max             =   1
      Min             =   2
      TabIndex        =   7
      Top             =   2160
      Value           =   2
      Width           =   255
   End
   Begin VB.Frame frmePlayList 
      Caption         =   "Play List Songs:"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2655
      Begin VB.CommandButton cmdSetSongNum 
         Caption         =   "Set Song #"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   670
         Width           =   975
      End
      Begin VB.CommandButton cmdDeletePlayList 
         Caption         =   "&Delete Current Playlist"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblScrollNum 
         Caption         =   "Song #: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   670
         Width           =   1095
      End
      Begin VB.Label lblCurrentSongNum 
         Caption         =   "Current Song Number: "
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdSetPosition 
      Caption         =   "Set Position"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.HScrollBar scrSetPos 
      Height          =   255
      Left            =   0
      Max             =   1
      TabIndex        =   5
      Top             =   1200
      Width           =   3615
   End
   Begin VB.HScrollBar ScrSongPos 
      Height          =   255
      Left            =   0
      Max             =   1
      TabIndex        =   4
      Top             =   840
      Width           =   4695
   End
   Begin VB.HScrollBar ScrVolume 
      Height          =   255
      LargeChange     =   25
      Left            =   1275
      Max             =   255
      SmallChange     =   10
      TabIndex        =   0
      Top             =   0
      Value           =   255
      Width           =   2295
   End
   Begin VB.CommandButton cmdSetPanning 
      Caption         =   "Set Panning"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.HScrollBar ScrPanning 
      Height          =   255
      LargeChange     =   25
      Left            =   1275
      Max             =   255
      SmallChange     =   10
      TabIndex        =   2
      Top             =   240
      Value           =   113
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   120
   End
   Begin VB.CommandButton cmdSetVolume 
      Caption         =   "Set Volume"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Height          =   1215
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label lblLengthPos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Song Length/Position"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: "
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label lblWinamp 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Winamp is:"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDeletePlayList_Click()
'Clears winamp's playlist
    FindWinamp
    DeletePlayList
End Sub

Private Sub cmdSetPanning_Click()
'Sets the panning
    FindWinamp
    SetPanning ScrPanning
End Sub

Private Sub cmdSetPosition_Click()
'Sets the position of the song to
'scrSetPos.value seconds
    FindWinamp
    SetCurrentSongPosition scrSetPos.Value, 0
End Sub

Private Sub cmdSetSongNum_Click()
'Sets the play list song number
    SetPlayListPosition (ScrSongNum.Value - 1)
    PlaySong
End Sub

Private Sub cmdSetVolume_Click()
'Sets the volume
    FindWinamp
    SetVolume ScrVolume
End Sub

Private Sub Form_Load()
'Sets lblInfo caption
    lblInfo.Caption = "Thanks for downloading this sample. This is only a sample. It does not show everything the module can do. It can do A LOT more. Look in the module (Winamp.bas), its all commented, so it shouldn't be too hard to understand. Good Luck. PS: feel free to e-mail me comments, help, suggestions, anything at midterror@hotmail.com"
End Sub

Private Sub ScrSongNum_Change()
'Changes the label to fit what the value
'Of the scrollbar is equal to
    lblScrollNum.Caption = "Song #: " & ScrSongNum.Value
End Sub

Private Sub Timer1_Timer()
Dim RC As Long
On Error Resume Next

'Finds out if Winamp is open and responds by
'setting the label to ON/OFF
    RC = FindWinamp
    If RC Then
        lblWinamp.Caption = "Winamp is: On"
    Else
        lblWinamp.Caption = "Winamp is: Off"
    End If
    
'finds out if winamp is playing, stopped, or paused
    RC = IsPlaying
    If RC = 1 Then
        lblStatus.Caption = "Status: Playing"
    ElseIf RC = 0 Then
        lblStatus.Caption = "Status: Stopped"
    Else
        lblStatus.Caption = "Status: Paused"
    End If
    
'Sets the scrollbars' max to the song length
    ScrSongPos.Max = GetSongLength
    scrSetPos.Max = ScrSongPos.Max

'Sets the scroll bar to the current position
'of the song
    ScrSongPos.Value = GetCurrentSongPosition / 1000
    
'If winamp has a song in the list then
'Find its length and current position and
'Set a lable to them
    If GetSongLength <> 0 Then
        lblLengthPos.Caption = "Length: " & ScrSongPos.Max & _
        " seconds    Current Position: " & ScrSongPos.Value & " seconds"
    Else
        lblLengthPos.Caption = "No Song To Get Info On"
    End If
    
'Sets the scrollbar to have a max of the
'number of songs in the playlist
    ScrSongNum.Min = GetPlayListLength
    
'Sets the frame caption to the total songs
'In the playlist
    frmePlayList.Caption = "Play List Songs: " & ScrSongNum.Min

'Sets the label caption to the current song
'number being played in the playlist
    lblCurrentSongNum.Caption = "Current Song Number: " & GetPlayListPosition + 1
    
'Gets information about the mp3 being played
    lblSampleRate.Caption = "Samplerate: " & GetSamplerate & " KHz"
    lblBitrate.Caption = "Bitrate: " & GetBitrate & " KBps"
    lblChannels.Caption = "Channels: " & GetChannels
    
End Sub

Attribute VB_Name = "WinAmpMod"
'WinampMod - winamp.bas
'By MidTerror - midterror@hotmail.com
'Feel free to mail comments, bugs, advice
'I'd appreciate it if you gave me some credit
'If you make a program out of this. If not
'then it's ok, but I wouldn't mind seeing the
'program, so feel free to send me programs you
'make with this.

Option Explicit

'All the Declarations
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageCDS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_USER = &H400
Public Const WM_WA_IPC = WM_USER
Public Const WM_COPYDATA = &H4A
Public Const WM_COMMAND = &H111

Public hwnd_winamp As Long

Public Const IPC_DELETE = 101
Public Const IPC_ISPLAYING = 104
Public Const IPC_GETOUTPUTTIME = 105
Public Const IPC_JUMPTOTIME = 106
Public Const IPC_WRITEPLAYLIST = 120
Public Const IPC_SETPLAYLISTPOS = 121
Public Const IPC_SETVOLUME = 122
Public Const IPC_SETPANNING = 123
Public Const IPC_GETLISTLENGTH = 124
Public Const IPC_SETSKIN = 200
Public Const IPC_GETSKIN = 201
Public Const IPC_GETLISTPOS = 125
Public Const IPC_GETINFO = 126
Public Const IPC_GETEQDATA = 127
Public Const IPC_PLAYFILE = 100
Public Const IPC_CHDIR = 103
Public Const WINAMP_OPTIONS_EQ = 40036
Public Const WINAMP_OPTIONS_PLEDIT = 40040
Public Const WINAMP_VOLUMEUP = 40058
Public Const WINAMP_VOLUMEDOWN = 40059
Public Const WINAMP_FFWD5S = 40060
Public Const WINAMP_REW5S = 40061
Public Const WINAMP_BUTTON1 = 40044
Public Const WINAMP_BUTTON2 = 40045
Public Const WINAMP_BUTTON3 = 40046
Public Const WINAMP_BUTTON4 = 40047
Public Const WINAMP_BUTTON5 = 40048
Public Const WINAMP_BUTTON1_SHIFT = 40144
Public Const WINAMP_BUTTON4_SHIFT = 40147
Public Const WINAMP_BUTTON5_SHIFT = 40148
Public Const WINAMP_BUTTON1_CTRL = 40154
Public Const WINAMP_BUTTON2_CTRL = 40155
Public Const WINAMP_BUTTON5_CTRL = 40158
Public Const WINAMP_FILE_PLAY = 40029
Public Const WINAMP_OPTIONS_PREFS = 40012
Public Const WINAMP_OPTIONS_AOT = 40019
Public Const WINAMP_HELP_ABOUT = 40041

Public Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As String
End Type


Public Function FindWinamp() As Long
'Find winamp window
'Returns 1 if winamp is open, 0 if not
    hwnd_winamp = FindWindow("Winamp v1.x", vbNullString)
    If hwnd_winamp Then FindWinamp = 1 Else FindWinamp = 0
End Function

Public Function DeletePlayList() As Long
'Clears the play list
    DeletePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_DELETE)
End Function
Public Function IsPlaying() As Long
'Returns:
'1 If playing
'3 if paused
'0 if stopped
    IsPlaying = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_ISPLAYING)
End Function

Public Function GetCurrentSongPosition() As Double
'Finds the current song position in milliseconds
    GetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETOUTPUTTIME)
End Function

Public Function GetSongLength() As Long
'Finds the song length in Seconds
    GetSongLength = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETOUTPUTTIME)
End Function

Public Function SetCurrentSongPosition(Optional Seconds As Long, Optional Ms As Long)
'Sets the current position in the song
'Returns:
'0 if success
'1 if eof
'-1 if not playing
    SetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, (Seconds * 1000 + Ms), IPC_JUMPTOTIME)
End Function


Public Function WritePlayList() As Long
'Writes the current playlist to C:\WINAMP_DIR\Winamp.m3u
'And then finds the play position
'Now obsolete, but good for old version of winamp
'Look at GetPlayListPosition
    WritePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_WRITEPLAYLIST)
End Function

Public Function SetPlayListPosition(Position As Integer) As Long
'Sets which song to play (0 being first)
    SetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, Position, IPC_SETPLAYLISTPOS)
End Function

Public Function SetVolume(Volume As Integer) As Long
'Sets the volume (Volume must be between 0 - 255)
    SetVolume = SendMessage(hwnd_winamp, WM_WA_IPC, Volume, IPC_SETVOLUME)
End Function

Public Function SetPanning(PanPosition As Integer) As Long
'Sets the panning (PanPosition must be between 0 - 255)
    SetPanning = SendMessage(hwnd_winamp, WM_WA_IPC, PanPosition, IPC_SETPANNING)
End Function

Public Function GetPlayListLength() As Long
'Gets amount of songs in play list
    GetPlayListLength = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTLENGTH)
End Function


Public Function GetPlayListPosition() As Long
'Returns which song its playing in the playlist
'0 being first
    GetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTPOS)
End Function

Public Function GetSamplerate() As Long
'Gets the samplerate
    GetSamplerate = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETINFO)
End Function

Public Function GetBitrate() As Long
'Gets the bitrate
    GetBitrate = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETINFO)
End Function

Public Function GetChannels() As Long
'Gets the channel
    GetChannels = SendMessage(hwnd_winamp, WM_WA_IPC, 2, IPC_GETINFO)
End Function

Public Function GetEQBandData(BandNumber As Integer) As Long
'Get each EQ banddata (0 being the first, 9 being last)
'Returns 0 - 255
    If BandNumber > 9 Then Exit Function
    GetEQBandData = SendMessage(hwnd_winamp, WM_WA_IPC, BandNumber, IPC_GETEQDATA)
End Function

Public Function GetEQPreampValue() As Long
'Gets the preamp value (Between 0 - 255)
    GetEQPreampValue = SendMessage(hwnd_winamp, WM_WA_IPC, 10, IPC_GETEQDATA)
End Function

Public Function GetEQEnabled()
'1 if EQ is enabled
'0 if it isn't
    GetEQEnabled = SendMessage(hwnd_winamp, WM_WA_IPC, 11, IPC_GETEQDATA)
End Function

Public Function GetEQAutoLoad()
'1 if EQ is autoloaded
'0 if it isn't
    GetEQAutoLoad = SendMessage(hwnd_winamp, WM_WA_IPC, 12, IPC_GETEQDATA)
End Function

Public Function PlayFile(FileToPlay As String) As Long
'Adds FileToPlay to the play list
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_PLAYFILE
    CDS.lpData = FileToPlay
    CDS.cbData = Len(FileToPlay) + 1
    PlayFile = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function ChangeDirectory(Directory As String) As Long
'Changes directory
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_CHDIR
    CDS.lpData = Directory
    CDS.cbData = Len(Directory) + 1
    ChangeDirectory = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function ToggleEQWindow() As Long
'Turns on or off the EQ window
    ToggleEQWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_EQ, 0)
End Function

Public Function TogglePlayListWindow() As Long
'Turns on or off play list window
    TogglePlayListWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PLEDIT, 0)
End Function

Public Function VolumeUp() As Long
'Raises the volume a tiny bit
    VolumeUp = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEUP, 0)
End Function
Public Function VolumeDown() As Long
'Sets the volume down a tiny bit
    VolumeDown = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEDOWN, 0)
End Function

Public Function Rewind() As Long
'Rewinds by 5 seconds
    Rewind = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_REW5S, 0)
End Function

Public Function FastForward() As Long
'Fast forwards by 5 seconds
    FastForward = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_FFWD5S, 0)
End Function

Public Function PreviousTrack() As Long
'Plays the previous song
    PreviousTrack = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1, 0)
End Function

Public Function PlaySong() As Long
'Plays the current song
    PlaySong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2, 0)
End Function

Public Function PauseSong() As Long
'Pauses playing
    PauseSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON3, 0)
End Function
Public Function StopSong() As Long
'Stops playing
    StopSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4, 0)
End Function

Public Function NextTrack() As Long
'Plays the next song in the playlist
    NextTrack = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5, 0)
End Function

Public Function FadeStop() As Long
'slowly fades away until it stops
    FadeStop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4_SHIFT, 0)
End Function

Public Function FirstSong() As Long
'Goes to the first song in the play list
    FirstSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1_CTRL, 0)
End Function

Public Function LastSong() As Long
'Goes to the last song in the play list
    LastSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5_CTRL, 0)
End Function
Public Function OpenLocation() As Long
'Shows Open Location Dialog
    OpenLocation = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2_CTRL, 0)
End Function
Public Function LoadFile() As Long
'Shows Load a file dialog
    LoadFile = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_FILE_PLAY, 0)
End Function
Public Function ShowPreferences() As Long
'Shows Preferences Dialog
    ShowPreferences = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PREFS, 0)
End Function

Public Function ToggleAlwaysOnTop() As Long
'Turns Always On Top On and Off
    ToggleAlwaysOnTop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_AOT, 0)
End Function

Public Function ShowAbout() As Long
'Shows About Box
    ShowAbout = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_HELP_ABOUT, 0)
End Function

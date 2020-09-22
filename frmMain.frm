VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "sn0st0rm Multimedia"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   -660
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DDE"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cpOpen 
      Caption         =   "^"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      ToolTipText     =   "Open..."
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cpForward 
      Caption         =   ">"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   "Forward One Track"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cpStop 
      Caption         =   "#"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      ToolTipText     =   "Stop Track"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cpPause 
      Caption         =   "||"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      ToolTipText     =   "Pause Track"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cpPlay 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      ToolTipText     =   "Play Track"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cpBack 
      Caption         =   "<"
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      ToolTipText     =   "Back One Track"
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar scrollVolume 
      Height          =   135
      Left            =   4800
      Max             =   100
      TabIndex        =   0
      Top             =   480
      Value           =   100
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Width           =   3375
      Begin VB.TextBox songInfo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "frmMain.frx":49866
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.CommandButton CommandButton6 
      Height          =   375
      Left            =   4200
      Picture         =   "frmMain.frx":49876
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Open..."
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton CommandButton5 
      Height          =   375
      Left            =   3600
      Picture         =   "frmMain.frx":4AD6E
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Forward One Track"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton CommandButton4 
      Height          =   375
      Left            =   3000
      Picture         =   "frmMain.frx":4C266
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Stop Track"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton CommandButton3 
      Height          =   375
      Left            =   2400
      Picture         =   "frmMain.frx":4D75E
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Pause Track"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton CommandButton2 
      Height          =   375
      Left            =   1800
      Picture         =   "frmMain.frx":4EC56
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Play Track"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton CommandButton1 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":5014E
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Back One Track"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer timez 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   120
      Top             =   2280
   End
   Begin VB.Timer tmrBass 
      Interval        =   250
      Left            =   120
      Top             =   1800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5850
      TabIndex        =   16
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5700
      TabIndex        =   15
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5580
      TabIndex        =   14
      ToolTipText     =   "Size"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Menu"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CPU Drain"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label cpuAmount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      ToolTipText     =   "Current CPU Usage of Bass.dll"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "sn0st0rm Multimedia Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim config As Uematsu_CONFIG
Dim gblStatus As Integer
Const mStopped = 0
Const mPaused = 1
Const mPlay = 2

Private Sub CommandButton1_Click()
If frmPlaylist.playlist.ListCount = 0 Then Exit Sub
timez.Enabled = False
currentTrack = currentTrack - 1
If currentTrack = -1 Then currentTrack = (frmPlaylist.playlist.ListCount - 1)
PlaySomeMusic
timez.Enabled = True
UpdateCaption
scrollVolume.SetFocus
End Sub

Private Sub CommandButton2_Click()
If frmPlaylist.playlist.ListCount = 0 Then
 MsgBox "No tracks available.", vbExclamation, "Cannot Play Track"
 scrollVolume.SetFocus
Exit Sub
End If
If gblStatus = mPlay Then
  scrollVolume.SetFocus
  Exit Sub
End If
If gblStatus = mPaused Then
  Select Case MediaType
    Case MEDIA_STREAM
      BASS_ChannelResume STRM
      timez.Enabled = True
      gblStatus = mPlay
      scrollVolume.SetFocus
      Exit Sub
    Case MEDIA_MODULE
      BASS_ChannelResume ModHandle
      timez.Enabled = True
      gblStatus = mPlay
      scrollVolume.SetFocus
      Exit Sub
    Case MEDIA_MIDI
      gblStatus = mPlay
      scrollVolume.SetFocus
      Exit Sub 'MIDI cannot be paused
    Case MEDIA_SPC
      DoPause
      gblStatus = mPlay
      scrollVolume.SetFocus
      Exit Sub
  End Select
End If
PlaySomeMusic
gblStatus = mPlay
UpdateCaption
scrollVolume.SetFocus
End Sub

Private Sub CommandButton3_Click()
If gblStatus = mPaused Then
  CommandButton2_Click
  scrollVolume.SetFocus
  Exit Sub
End If
Select Case MediaType
  Case MEDIA_STREAM
    gblStatus = mPaused
    timez = False
    BASS_ChannelPause STRM
  Case MEDIA_MODULE
    gblStatus = mPaused
    timez = False
    BASS_ChannelPause ModHandle
  Case MEDIA_SPC
    gblStatus = mPaused
    DoPause
End Select
scrollVolume.SetFocus
End Sub

Private Sub CommandButton4_Click()
On Local Error Resume Next
gblStatus = mStopped
Select Case MediaType
  Case MEDIA_STREAM
    BASS_ChannelStop STRM
    If Err Then MsgBox "oh shit!"
    timez.Enabled = False
  Case MEDIA_MODULE
    BASS_ChannelStop ModHandle
    If Err Then MsgBox "Cannot stop module"
    timez.Enabled = False
  Case MEDIA_MIDI
    HarmonyStopMusic
  Case MEDIA_SPC
    DoStop
End Select
scrollVolume.SetFocus
End Sub

Private Sub CommandButton5_Click()
If frmPlaylist.playlist.ListCount = 0 Then Exit Sub
timez.Enabled = False
currentTrack = currentTrack + 1
If currentTrack >= frmPlaylist.playlist.ListCount Then currentTrack = 0
PlaySomeMusic
timez.Enabled = True
UpdateCaption
scrollVolume.SetFocus
End Sub

Private Sub CommandButton6_Click()
MenuForm.dlg.DialogTitle = "Add Track"
MenuForm.dlg.FileName = ""
MenuForm.dlg.Filter = "All Supported Media Formats|*.mp3;*.wav;*.mod;*.it;*.xm;*.s3m;*.mo3;*.mtm;*.spc;*.mid;*.midi;*.st0;*.sdl|Streamable Files|*.mp3;*.wav|Modules|*.mod;*.it;*.xm;*.s3m;*.mo3;*.mtm|SNES Music|*.spc|MIDI Sequences|*.mid;*.midi|Playlists|*.st0;*.sdl|All Files|*.*|"
On Local Error Resume Next
MenuForm.dlg.ShowOpen
If Err Then
  scrollVolume.SetFocus
  Exit Sub
End If
If UCase$(Right$(MenuForm.dlg.FileName, 4)) = ".ST0" Or UCase$(Right$(MenuForm.dlg.FileName, 4)) = ".SDL" Then
  If frmPlaylist.playlist.ListCount > 0 Then
    If MsgBox("This will clear the current playlist. Proceed?", vbYesNo Or vbQuestion, "Confirm Playlist Clear") = vbNo Then Exit Sub
  End If
  Dim trackname As String, fh As Integer
  frmPlaylist.playlist.Clear
  fh = FreeFile
  Open MenuForm.dlg.FileName For Input As #fh
  If UCase$(Right$(MenuForm.dlg.FileName, 4)) = ".SDL" Then Input #fh, listName
  Do
   Input #fh, trackname
   frmPlaylist.playlist.AddItem trackname
  Loop Until EOF(fh)
  Close fh
  frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
  currentTrack = 0
  CommandButton4_Click
  CommandButton2_Click
  Exit Sub
End If
frmPlaylist.playlist.AddItem MenuForm.dlg.FileName
frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
frmPlaylist.CommandButton1.Enabled = True
scrollVolume.SetFocus
End Sub

Private Sub cpBack_Click()
CommandButton1_Click
End Sub

Private Sub cpForward_Click()
CommandButton5_Click
End Sub

Private Sub cpOpen_Click()
CommandButton6_Click
End Sub

Private Sub cpPause_Click()
CommandButton3_Click
End Sub

Private Sub cpPlay_Click()
CommandButton2_Click
End Sub

Private Sub cpStop_Click()
CommandButton4_Click
End Sub

Private Sub Form_GotFocus()
Label1.ForeColor = &HFFFF
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
Command = CmdStr
End Sub

Private Sub Form_Load()
On Local Error Resume Next
If App.PrevInstance = True Then
  If Command$ <> "" Then
    Clipboard.Clear
    Clipboard.SetText "sn0st0rm-" & Trim$(Command$), vbCFText
  End If
  End
End If
frmMain.LinkMode = 1
Load MenuForm
Load frmPlaylist
If BASS_GetStringVersion <> "0.8" Then
 MsgBox "BASS version 0.8 was not loaded", vbCritical, "Error!"
 End
End If
If BASS_Init(-1, 44100, 0, Me.hwnd) = BASSFALSE Then
 MsgBox "Can't initialize digital sound system", vbCritical, "Error!"
 End
End If
If BASS_Start = BASSFALSE Then
 MsgBox "Can't start digital output", vbCritical, "Error!"
 End
End If
'start up MIDI engine
HarmonyCreate
HarmonyInitMidi
'initialize SPC emulation options
Dim inLong As Long
config.SmpRate = Val(GetSetting(App.EXEName, "SPC Emulation", "Sample Rate", "44100"))
inLong = GetSetting(App.EXEName, "SPC Emulation", "Bits Per Sample", 1)
Select Case inLong
  Case 0: config.BPS = 8
  Case 1: config.BPS = 16
  Case 2: config.BPS = 32
End Select
inLong = GetSetting(App.EXEName, "SPC Emulation", "Channels", 1)
config.NChn = inLong + 1
inLong = GetSetting(App.EXEName, "SPC Emulation", "Mixing Engine", 0)
config.MixingEngine = inLong + 1
inLong = GetSetting(App.EXEName, "SPC Emulation", "Interpolation", 0)
config.Interpolation = inLong
config.BufferLength = Val(GetSetting(App.EXEName, "SPC Emulation", "Buffer Length", "2000"))
config.MiscOpt = 0
inLong = GetSetting(App.EXEName, "SPC Emulation", "Lowpass Filter", 0)
If inLong = 1 Then config.MiscOpt = config.MiscOpt + 1
inLong = GetSetting(App.EXEName, "SPC Emulation", "Use Old Sample Routine", 0)
If inLong = 1 Then config.MiscOpt = config.MiscOpt + 2
inLong = GetSetting(App.EXEName, "SPC Emulation", "Surround Sound", 0)
If inLong = 1 Then config.MiscOpt = config.MiscOpt + 4
config.VisRate = 10
Uematsu_SetConfiguration config
inLong = GetSetting(App.EXEName, "SPC Emulation", "Mixing Priority", 3)
Select Case Val(inLong)
  Case 0: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_IDLE
  Case 1: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_LOWEST
  Case 2: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_BELOW_NORMAL
  Case 3: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_NORMAL
  Case 4: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_ABOVE_NORMAL
  Case 5: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_HIGHEST
  Case 6: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_TIME_CRITICAL
End Select
mircString = GetSetting(App.EXEName, "General Settings", "mIRC DDE Echo String", "/me is listening to: %s")
mircTail = GetSetting(App.EXEName, "General Settings", "Use sn0st0rm Tail", 1)


'finish startup
tmrBass.Enabled = True
taskbarstring = "...sn0st0rm Multimedia"
'get general settings from registry
scrolltaskbar = GetSetting(App.EXEName, "General Settings", "Scroll In Taskbar", 0)
mIRCecho = GetSetting(App.EXEName, "General Settings", "Echo to mIRC", 0)
OnTop = GetSetting(App.EXEName, "General Settings", "Always On Top", 0)
Me.Show
If OnTop = True Then Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
scrollVolume.Value = BASS_GetVolume
If Command <> "" Then
  Dim newFile As String
  If Left(Command, 1) = Chr(34) Then newFile = Mid(Command, 2, Len(Command) - 2) Else newFile = Command
  frmPlaylist.playlist.Clear
  frmPlaylist.playlist.AddItem newFile
  frmPlaylist.playlist.ListIndex = 0
  frmPlaylist.trackCount.Caption = "1"
  CommandButton2_Click
End If
End Sub

Private Sub Form_LostFocus()
Label1.ForeColor = &H7777
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload MenuForm
Unload frmPlaylist
HarmonyStopMusic
HarmonyTermMidi
HarmonyRelease
DoStop
BASS_Stop
BASS_Free
End
End Sub

Private Sub Label1_DblClick()
Label5_MouseUp 0, 0, 0, 0
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Button = 1 Then
 Dim ReturnVal As Long
 X = ReleaseCapture()
 ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload frmPlaylist
Unload MenuForm
HarmonyStopMusic
HarmonyTermMidi
HarmonyRelease
DoStop
BASS_Stop
BASS_Free
End
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.WindowState = 1
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu MenuForm.mnuAppBox
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If frmMain.Height = 3720 Then
 Me.cpBack.Visible = True
 Me.cpForward.Visible = True
 Me.cpOpen.Visible = True
 Me.cpPause.Visible = True
 Me.cpStop.Visible = True
 Me.cpPlay.Visible = True
 frmMain.Height = 255
Else
 Me.cpBack.Visible = False
 Me.cpForward.Visible = False
 Me.cpOpen.Visible = False
 Me.cpPause.Visible = False
 Me.cpStop.Visible = False
 Me.cpPlay.Visible = False
 frmMain.Height = 3720
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub scrollVolume_Change()
BASS_SetVolume scrollVolume.Value
End Sub

Private Sub songInfo_GotFocus()
Me.scrollVolume.SetFocus
End Sub

Private Sub timez_Timer()
If MediaType <> MEDIA_STREAM Then GoTo skipstuff
If BASS_ChannelIsActive(STRM) = BASSFALSE Then
 timez.Enabled = False
 currentTrack = currentTrack + 1
 If currentTrack >= frmPlaylist.playlist.ListCount Then currentTrack = 0
 PlaySomeMusic
 timez.Enabled = True
 UpdateCaption
End If
skipstuff:
End Sub

Private Sub tmrBass_Timer()
If MediaType = MEDIA_MIDI Or MediaType = MEDIA_SPC Then
 cpuAmount.Caption = "N/A"
 GoTo skipthis
End If
Dim p As Long
cpuAmount.Caption = Trim(Str(CInt(BASS_GetCPU))) + "%"
skipthis:
If scrolltaskbar = True Then
 neeples = Len(taskbarstring)
 scrollstep = scrollstep + 1
 If scrollstep > neeples Then scrollstep = 1
 leftside$ = Left(taskbarstring, scrollstep)
 rightside$ = Right(taskbarstring, neeples - scrollstep)
 frmMain.Caption = rightside$ + leftside$
End If
DoEvents
End Sub

Private Sub UpdateCaption()
On Local Error Resume Next
taskbarstring = "...sn0st0rm Multimedia - " + frmPlaylist.playlist.List(currentTrack)
frmMain.Caption = taskbarstring
scrollstep = 0
End Sub

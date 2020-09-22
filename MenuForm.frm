VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuForm 
   ClientHeight    =   465
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8475
   Icon            =   "MenuForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer checkClipboard 
      Interval        =   1000
      Left            =   1680
      Top             =   0
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuAppBox 
      Caption         =   "AppBox"
      Begin VB.Menu mnuShowPlaylist 
         Caption         =   "Show Playlist"
      End
      Begin VB.Menu sep7382645 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Skins"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configure Player"
         Begin VB.Menu mnuGenConfig 
            Caption         =   "General Configuration"
         End
         Begin VB.Menu sep982745687234 
            Caption         =   "-"
         End
         Begin VB.Menu mnuConfigSPCengine 
            Caption         =   "SPC Engine"
         End
      End
      Begin VB.Menu sep876384726 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu sep8723648756234 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "Playlist"
      Begin VB.Menu mnuAddTrack 
         Caption         =   "Add Track"
      End
      Begin VB.Menu mnuRemoveTrack 
         Caption         =   "Remove Track"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveTrackUp 
         Caption         =   "Move Track Up"
      End
      Begin VB.Menu mnuMoveTrackDown 
         Caption         =   "Move Track Down"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveAllFiles 
         Caption         =   "Remove All Files"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPlaylist 
         Caption         =   "Load Playlist"
      End
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub checkClipboard_Timer()
On Local Error Resume Next
Dim result As Long
If Clipboard.GetFormat(vbCFText) Then
  Dim getString As String, newFile As String
  getString = Clipboard.GetText
  If Left$(getString, 9) = "sn0st0rm-" Then
    newFile = Right$(getString, Len(getString) - 9)
    frmPlaylist.playlist.AddItem newFile
    frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
    frmPlaylist.CommandButton1.Enabled = True
    currentTrack = frmPlaylist.playlist.ListCount - 1
    Clipboard.Clear
    PlaySomeMusic
    taskbarstring = "...sn0st0rm Multimedia - " + frmPlaylist.playlist.List(currentTrack)
  End If
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuAddTrack_Click()
Me.dlg.DialogTitle = "Add Track"
Me.dlg.FileName = ""
Me.dlg.Filter = "All Supported Media Formats|*.mp3;*.wav;*.mod;*.it;*.xm;*.s3m;*.mo3;*.mtm;*.spc;*.mid;*.midi|Streamable Files|*.mp3;*.wav|Modules|*.mod;*.it;*.xm;*.s3m;*.mo3;*.mtm|SNES Music|*.spc|MIDI Sequences|*.mid;*.midi|All Files|*.*|"
On Local Error Resume Next
Me.dlg.ShowOpen
If Err Then Exit Sub
frmPlaylist.playlist.AddItem Me.dlg.FileName
frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
frmPlaylist.CommandButton1.Enabled = True
End Sub

Private Sub mnuConfigSPCengine_Click()
frmSPC.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuGenConfig_Click()
frmGeneral.Show
End Sub

Private Sub mnuHelpTopics_Click()
Beep
End Sub

Private Sub mnuLoadPlaylist_Click()
If frmPlaylist.playlist.ListCount > 0 Then
 Dim result As Integer
 result = MsgBox("This will wipe all the songs out of the current playlist! Proceed?", vbQuestion + vbYesNo, "Confirm Playlist Clear")
 If result = vbNo Then Exit Sub
 frmPlaylist.playlist.Clear
 listName = ""
 spindleFileName = ""
End If
dlg.DialogTitle = "Load Playlist"
dlg.Filter = "sn0st0rm Lists (*.st0)|*.st0|Spindle Lists (*.sdl)|*.sdl|All Files (*.*)|*.*|"
On Local Error Resume Next
dlg.ShowOpen
If Err Then Exit Sub
Dim trackname As String, fh As Integer
fh = FreeFile
Open dlg.FileName For Input As #fh
If dlg.FilterIndex = 2 Then Input #fh, listName
Do
 Input #fh, trackname
 frmPlaylist.playlist.AddItem trackname
Loop Until EOF(fh)
Close fh
frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
End Sub

Private Sub mnuMoveTrackDown_Click()
Dim choice1 As String
Dim choice2 As String
If frmPlaylist.playlist.ListIndex = -1 Or frmPlaylist.playlist.ListIndex = (frmPlaylist.playlist.ListCount - 1) Then Exit Sub
choice1 = frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex)
choice2 = frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex + 1)
frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex) = choice2
frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex + 1) = choice1
frmPlaylist.playlist.ListIndex = frmPlaylist.playlist.ListIndex + 1
frmPlaylist.playlist.Refresh
DoEvents
End Sub

Private Sub mnuMoveTrackUp_Click()
If frmPlaylist.playlist.ListIndex < 1 Then Exit Sub
Dim choice1 As String
Dim choice2 As String
choice1 = frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex)
choice2 = frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex - 1)
frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex) = choice2
frmPlaylist.playlist.List(frmPlaylist.playlist.ListIndex - 1) = choice1
frmPlaylist.playlist.ListIndex = frmPlaylist.playlist.ListIndex - 1
frmPlaylist.playlist.Refresh
DoEvents
End Sub

Private Sub mnuRemoveAllFiles_Click()
Dim result As Integer
result = MsgBox("This will wipe all the songs out of the current playlist! Proceed?", vbQuestion + vbYesNo, "Confirm Playlist Clear")
If result = vbNo Then Exit Sub
frmPlaylist.playlist.Clear
frmPlaylist.CommandButton1.Enabled = False
frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
listName = ""
spindleFileName = ""
End Sub

Private Sub mnuRemoveTrack_Click()
If frmPlaylist.playlist.ListIndex = -1 Then Exit Sub
frmPlaylist.playlist.RemoveItem frmPlaylist.playlist.ListIndex
frmPlaylist.trackCount.Caption = frmPlaylist.playlist.ListCount
If frmPlaylist.playlist.ListCount = 0 Then frmPlaylist.CommandButton1.Enabled = False
End Sub

Private Sub mnuShowPlaylist_Click()
frmPlaylist.Show
End Sub

Private Sub mnuSkins_Click()
MsgBox "Skins are not coded in yet.", vbExclamation, "Oops!"
End Sub

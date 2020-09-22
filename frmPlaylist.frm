VERSION 5.00
Begin VB.Form frmPlaylist 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "sn0st0rm Playlist Editor"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox playlist 
      BackColor       =   &H00000080&
      ForeColor       =   &H80000009&
      Height          =   3960
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton CommandButton3 
      Caption         =   "Hide Me"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Save List"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label trackCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tracks In List:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
MenuForm.dlg.FileName = ""
If spindleFileName <> "" Then
 If listName <> "" Then
   SaveSpindleList
   Exit Sub
 Else
  SaveNormalList
  Exit Sub
 End If
End If
MenuForm.dlg.DialogTitle = "Save Playlist"
MenuForm.dlg.Filter = "sn0st0rm List|*.st0|Spindle List|*.sdl|"
On Local Error Resume Next
MenuForm.dlg.ShowSave
If Err Then Exit Sub
If MenuForm.dlg.FilterIndex = 2 Then
 Do
  listName = InputBox("Spindle Lists require an identifying name, please enter one", "Create Spindle List")
 Loop Until listName <> ""
End If
spindleFileName = MenuForm.dlg.FileName
If listName <> "" Then
 SaveSpindleList
Else
 SaveNormalList
End If
End Sub

Private Sub CommandButton3_Click()
Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 Dim ReturnVal As Long
 X = ReleaseCapture()
 ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If
End Sub

Private Sub playlist_DblClick()
frmMain.timez.Enabled = False
currentTrack = frmPlaylist.playlist.ListIndex
PlaySomeMusic
frmMain.timez.Enabled = True
UpdateCaption2
End Sub

Private Sub playlist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 If playlist.ListCount = 0 Then
  MenuForm.mnuRemoveAllFiles.Enabled = False
 Else
  MenuForm.mnuRemoveAllFiles.Enabled = True
 End If
 If playlist.ListIndex < 1 Then
  MenuForm.mnuMoveTrackUp.Enabled = False
 Else
  MenuForm.mnuMoveTrackUp.Enabled = True
 End If
 If playlist.ListIndex = (playlist.ListCount - 1) Then
  MenuForm.mnuMoveTrackDown.Enabled = False
 Else
  MenuForm.mnuMoveTrackDown.Enabled = True
 End If
 If playlist.ListCount = 1 Then
  MenuForm.mnuMoveTrackDown.Enabled = False
  MenuForm.mnuMoveTrackUp.Enabled = False
 End If
 If playlist.ListIndex > -1 Then
  MenuForm.mnuRemoveTrack.Enabled = True
 Else
  MenuForm.mnuRemoveTrack.Enabled = False
  MenuForm.mnuMoveTrackDown.Enabled = False
  MenuForm.mnuMoveTrackUp.Enabled = False
 End If
 PopupMenu MenuForm.mnuPlaylist
End If
End Sub

Private Sub SaveSpindleList()
Dim fh As Integer
fh = FreeFile
Open spindleFileName For Output As #fh
Print #fh, listName
For aa = 0 To (playlist.ListCount - 1)
 Print #fh, playlist.List(aa)
Next aa
Close fh
End Sub

Private Sub SaveNormalList()
Dim fh As Integer
fh = FreeFile
Open spindleFileName For Output As #fh
For aa = 0 To (playlist.ListCount - 1)
 Print #fh, playlist.List(aa)
Next aa
Close fh
End Sub

Private Sub UpdateCaption2()
On Local Error Resume Next
taskbarstring = "...sn0st0rm Multimedia - " + playlist.List(currentTrack)
End Sub


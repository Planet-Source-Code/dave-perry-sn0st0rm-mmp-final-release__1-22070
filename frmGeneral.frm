VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Program Options"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "mIRC DDE String"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
      Begin VB.TextBox txtMircString 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "/me is listening to: %s"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chkRedTail 
         Caption         =   "Use sn0st0rm's red tail in DDE string"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   3375
      End
      Begin VB.CheckBox ddeEcho 
         Caption         =   "Echo via DDE to mIRC"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox alwaysOnTop 
         Caption         =   "Player screen always on top"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox scrollTitle 
         Caption         =   "Scroll title in taskbar"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'save settings to registry
SaveSetting App.EXEName, "General Settings", "Echo to mIRC", ddeEcho.Value
SaveSetting App.EXEName, "General Settings", "Scroll In Taskbar", scrollTitle.Value
SaveSetting App.EXEName, "General Settings", "Always On Top", alwaysOnTop.Value
SaveSetting App.EXEName, "General Settings", "mIRC DDE Echo String", txtMircString.Text
mircString = txtMircString.Text
SaveSetting App.EXEName, "General Settings", "Use sn0st0rm Tail", chkRedTail.Value
If chkRedTail = 1 Then mircTail = True Else mircTail = False
If alwaysOnTop.Value = 1 Then
 Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 OnTop = True
Else
 Call SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
 OnTop = False
End If
If ddeEcho.Value = 1 Then mIRCecho = True Else mIRCecho = False
If scrollTitle.Value = 1 Then
 scrolltaskbar = True
Else
 scrolltaskbar = False
 scrollstep = 0
 frmMain.Caption = taskbarstring
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If scrolltaskbar = True Then scrollTitle.Value = 1 Else scrollTitle.Value = 0
If OnTop = True Then alwaysOnTop.Value = 1 Else alwaysOnTop.Value = 0
If mIRCecho = True Then
  ddeEcho.Value = 1
Else
  ddeEcho.Value = 0
End If
If mircTail = False Then chkRedTail.Value = 0 Else chkRedTail.Value = 1
txtMircString.Text = mircString
End Sub

VERSION 5.00
Begin VB.Form frmSPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SNESAPU/Uematsu SPC Engine Configuration"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmSPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1935
      Left            =   3840
      TabIndex        =   13
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox mixPri 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtStereoSep 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "32768"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtPreamp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "30"
         Top             =   1260
         Width           =   1455
      End
      Begin VB.CheckBox useSurroundSound 
         Caption         =   "Surround Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Enable/disable surround sound"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox oldSample 
         Caption         =   "Old Sample Decompression Routine"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox lowPass 
         Caption         =   "Low Pass Filter"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Stereo Separation"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1600
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Preamp Level"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1300
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Mixing Priority"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1000
         Width           =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mixing Options"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox buffer 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "Set the buffer between 100 and 10000"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox interpol 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox engine 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox channels 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox bitsPerSample 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox sampleRate 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Set the sample rate between 8000 and 96000"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Buffer"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Interpolation"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Engine"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Channels"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Bits Per Sample"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Sample Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim config As Uematsu_CONFIG

Private Sub buffer_Change()
If Val(buffer.Text) > 10000 Then buffer.Text = "10000"
If Val(buffer.Text) < 100 Then buffer.Text = "100"
End Sub

Private Sub Command1_Click()
'apply new settings and update registry
SaveSetting App.EXEName, "SPC Emulation", "Sample Rate", sampleRate.Text
config.SmpRate = Val(sampleRate.Text)
SaveSetting App.EXEName, "SPC Emulation", "Bits Per Sample", bitsPerSample.ListIndex
Select Case bitsPerSample.ListIndex
  Case 0: config.BPS = 8
  Case 1: config.BPS = 16
  Case 2: config.BPS = 32
End Select
SaveSetting App.EXEName, "SPC Emulation", "Channels", channels.ListIndex
config.NChn = channels.ListIndex + 1
SaveSetting App.EXEName, "SPC Emulation", "Mixing Engine", engine.ListIndex
config.MixingEngine = engine.ListIndex + 1
SaveSetting App.EXEName, "SPC Emulation", "Interpolation", interpol.ListIndex
config.Interpolation = interpol.ListIndex
SaveSetting App.EXEName, "SPC Emulation", "Buffer Length", buffer.Text
config.BufferLength = Val(buffer.Text)
config.MiscOpt = 0
SaveSetting App.EXEName, "SPC Emulation", "Lowpass Filter", lowPass.Value
If lowPass.Value = 1 Then config.MiscOpt = config.MiscOpt + 1
SaveSetting App.EXEName, "SPC Emulation", "Use Old Sample Routine", oldSample.Value
If oldSample.Value = 1 Then config.MiscOpt = config.MiscOpt + 2
SaveSetting App.EXEName, "SPC Emulation", "Surround Sound", useSurroundSound.Value
If useSurroundSound.Value = 1 Then config.MiscOpt = config.MiscOpt + 4
config.VisRate = 10
Uematsu_SetConfiguration config
SaveSetting App.EXEName, "SPC Emulation", "Mixing Priority", mixPri.ListIndex
Select Case mixPri.ListIndex
  Case 0: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_IDLE
  Case 1: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_LOWEST
  Case 2: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_BELOW_NORMAL
  Case 3: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_NORMAL
  Case 4: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_ABOVE_NORMAL
  Case 5: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_HIGHEST
  Case 6: Uematsu_SetMixingThreadPriority THREAD_PRIORITY_TIME_CRITICAL
End Select
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

'    Mixing_Priority As Long
'Note: use const's defined up top of
'file but is usually best if left at 0
'    Amp As Long
'Preamp level: 30 is usually the best
'    Stereo_Seperation As Long
'Set Stereo seperation:
'Settings range from:
'0 - No seperation (mono)
'32768 - Normal (snes)
'65536 - Full (completly left / right, no center)



Private Sub Form_Load()
sampleRate.Text = GetSetting(App.EXEName, "SPC Emulation", "Sample Rate", "44100")
bitsPerSample.AddItem "8 bits"
bitsPerSample.AddItem "16 bits"
bitsPerSample.AddItem "32 bits"
bitsPerSample.ListIndex = GetSetting(App.EXEName, "SPC Emulation", "Bits Per Sample", 1)
channels.AddItem "Mono"
channels.AddItem "Stereo"
channels.ListIndex = GetSetting(App.EXEName, "SPC Emulation", "Channels", 1)
engine.AddItem "386"
engine.AddItem "Intel MMX"
engine.AddItem "AMD 3DNow!"
engine.AddItem "Intel SSE"
engine.ListIndex = GetSetting(App.EXEName, "SPC Emulation", "Mixing Engine", 0)
interpol.AddItem "None"
interpol.AddItem "Linear"
interpol.AddItem "Cubic"
interpol.AddItem "Gaussian"
interpol.ListIndex = GetSetting(App.EXEName, "SPC Emulation", "Interpolation", 0)
buffer.Text = GetSetting(App.EXEName, "SPC Emulation", "Buffer Length", "2000")
'VisRate As Long 'Rate of visualization in Hertz, 10-120 (leave alone for now)
'APR As Long 'Automatic preamp (0 off/1 on)
'APRThreshhold As Long 'automatic preamp threshold
'MiscOpt As Long 'Misc options (bit 1 = low pass filter, bit 2 = old sample decompression routine, bit 3 = "Surround" sound)
lowPass.Value = GetSetting(App.EXEName, "SPC Emulation", "Lowpass Filter", 0)
oldSample.Value = GetSetting(App.EXEName, "SPC Emulation", "Use Old Sample Routine", 0)
useSurroundSound.Value = GetSetting(App.EXEName, "SPC Emulation", "Surround Sound", 0)
mixPri.AddItem "Idle"
mixPri.AddItem "Lowest"
mixPri.AddItem "Below Normal"
mixPri.AddItem "Normal"
mixPri.AddItem "Above Normal"
mixPri.AddItem "Highest"
mixPri.AddItem "Time Critical"
mixPri.ListIndex = Val(GetSetting(App.EXEName, "SPC Emulation", "Mixing Priority", "3"))
End Sub

Private Sub sampleRate_Change()
If Val(sampleRate.Text) > 96000 Then sampleRate.Text = "96000"
If Val(sampleRate.Text) < 8000 Then sampleRate.Text = "8000"
End Sub

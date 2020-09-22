Attribute VB_Name = "modDeclares"
Public Declare Function HarmonyCreate Lib "harmony.dll" () As Long
Public Declare Function HarmonyFadeInMusic Lib "harmony.dll" (ByVal TimeFactor As Long) As Long
Public Declare Function HarmonyFadeOutMusic Lib "harmony.dll" (ByVal TimeFactor As Long) As Long
Public Declare Function HarmonyGetVersion Lib "harmony.dll" () As Long
Public Declare Function HarmonyInitMidi Lib "harmony.dll" () As Long
Public Declare Function HarmonyPlayMusic Lib "harmony.dll" (ByVal FileName As String) As Long
Public Declare Function HarmonyRelease Lib "harmony.dll" () As Long
Public Declare Function HarmonySetMusicPanpot Lib "harmony.dll" (ByVal Panpot As Long) As Long
Public Declare Function HarmonySetMusicSpeed Lib "harmony.dll" (ByVal Speed As Long) As Long
Public Declare Function HarmonySetMusicVolume Lib "harmony.dll" (ByVal NewVolume As Long) As Long
Public Declare Function HarmonyStopMusic Lib "harmony.dll" () As Long
Public Declare Function HarmonyTermMidi Lib "harmony.dll" () As Long
Public Declare Function HarmonyGetMusicPlaying Lib "harmony.dll" () As Long
Public Declare Function HarmonyGetMusicLooping Lib "harmony.dll" () As Long
Public Declare Function HarmonyCheckValidMidi Lib "harmony.dll" () As Long
Public Declare Function HarmonyGetMidiTick Lib "harmony.dll" () As Long

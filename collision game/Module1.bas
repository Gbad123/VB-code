Attribute VB_Name = "Module1"
Option Explicit

'diming for scores
Global m As Integer
Global x As Integer

'sound
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Enum SND_Settings
   SND_SYNC = &H0
   SND_ASYNC = &H1
   SND_NODEFAULT = &H2
   SND_MEMORY = &H4
   SND_LOOP = &H8
   SND_NOSTOP = &H10
   SW_SHOW = 5
End Enum
Public Sub Play(fname As String, Optional Settings As SND_Settings = SND_ASYNC)
Dim retval As Long
retval = sndPlaySound(fname, Settings)
End Sub



'Play "\Pathway and filname.wav"

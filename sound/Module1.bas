Attribute VB_Name = "Module1"
Option Explicit

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

Private Declare Function sndPlaySound2 Lib "winmm" Alias "sndPlaySoundB" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Enum SND_Settings2
   SND_SYNC2 = &H0
   SND_ASYNC2 = &H1
   SND_NODEFAULT2 = &H2
   SND_MEMORY2 = &H4
   SND_LOOP2 = &H8
   SND_NOSTOP2 = &H10
   SW_SHOW2 = 5
End Enum





Public Sub Play(fname As String, Optional Settings As SND_Settings = SND_ASYNC)
Dim retval As Long
retval = sndPlaySound(fname, Settings)
End Sub



'Play "\Pathway and filname.wav"
Public Sub Play2(fname2 As String, Optional Settings2 As SND_Settings2 = SND_ASYNC)
Dim retval2 As Long
retval2 = sndPlaySound(fname2, Settings2)
End Sub



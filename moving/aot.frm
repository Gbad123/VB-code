VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   120
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   4200
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   3000
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6000
      Top             =   2880
   End
   Begin VB.Label Lblend 
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Image img1 
      Height          =   3390
      Left            =   2280
      Picture         =   "aot.frx":0000
      Top             =   480
      Width           =   3345
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyRight
        tmright.Enabled = True
    Case vbKeyLeft
        tmrleft.Enabled = True
    Case vbKeyUp
        tmrup.Enabled = True
    Case vbKeyDown
        tmrdown.Enabled = True
    Case vbKeySpace
        End
        
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

tmright.Enabled = False
tmrleft.Enabled = False
tmrup.Enabled = False
tmrdown.Enabled = False

End Sub

Private Sub Lblend_Click()

End

End Sub

Private Sub tmrdown_Timer()

img1.Top = img1.Top + 250
    If img1.Top > frmain.Height Then
        img1.Top = 0 - frmain.Top
            End If
End Sub

Private Sub tmright_Timer()

img1.Left = img1.Left + 250
    If img1.Left > frmain.Width Then
        img1.Left = 0 - img1.Left
            End If

End Sub

Private Sub tmrleft_Timer()

img1.Left = img1.Left - 250
    If img1.Left < 0 Then
        img1.Left = frmain.Width - img1.Left
            End If
End Sub

Private Sub tmrup_Timer()

img1.Top = img1.Top - 250
    If img1.Top < 0 Then
        img1.Top = frmain.Height - img1.Top
            End If
End Sub

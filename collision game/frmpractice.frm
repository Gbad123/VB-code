VERSION 5.00
Begin VB.Form frmpractice 
   Caption         =   "Practice"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   Picture         =   "frmpractice.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrfire 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   5160
   End
   Begin VB.Timer tmrfollow 
      Interval        =   1
      Left            =   960
      Top             =   5160
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   4560
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   4560
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   3960
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   3960
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Image imgbullet 
      Height          =   600
      Left            =   3960
      Picture         =   "frmpractice.frx":15CA
      Top             =   3480
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imggun 
      Height          =   1755
      Left            =   3960
      Picture         =   "frmpractice.frx":1BF3
      Top             =   4080
      Width           =   930
   End
End
Attribute VB_Name = "frmpractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'image moving properly
Select Case KeyCode
    
    Case vbKeyRight
        tmright.Enabled = True
    Case vbKeyLeft
        tmrleft.Enabled = True
    Case vbKeyUp
        tmrup.Enabled = True
    Case vbKeyDown
        tmrdown.Enabled = True
    Case vbKeyX
        End
    Case vbKeySpace
        tmrfollow.Enabled = False
            tmrfire.Enabled = True
                imgbullet.Visible = True
                    Play App.Path & "\gun2.wav" 'sound
                    
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'image start off stopped
tmright.Enabled = False
    tmrleft.Enabled = False
        tmrup.Enabled = False
            tmrdown.Enabled = False
        
End Sub

Private Sub lbl1_Click()

'go back
Load frmain
    frmain.Show
        Unload frmpractice
        
End Sub

Private Sub tmrfire_Timer()

'fireing the bullet
imgbullet.Top = imgbullet.Top - 200
    If imgbullet.Top < 0 - imgbullet.Height Then
                
        tmrfire.Enabled = False
            tmrfollow.Enabled = True
                imgbullet.Visible = False
                
        End If
        
End Sub

Private Sub tmrdown_Timer()

'moving down
imggun.Top = imggun.Top + 250
    If imggun.Top > frmpractice.Height Then
        imggun.Top = 0 - frmpractice.Top
            End If
            
End Sub

Private Sub tmrfollow_Timer()

'bullet following the gun
imgbullet.Left = imggun.Left + imggun.Width / 2 - imgbullet.Width / 2
    imgbullet.Top = imggun.Top
    
End Sub

Private Sub tmright_Timer()

'moving right
imggun.Left = imggun.Left + 250
    If imggun.Left > frmpractice.Width Then
        imggun.Left = 0 - imggun.Left
            End If

End Sub

Private Sub tmrleft_Timer()

'moving left
imggun.Left = imggun.Left - 250
    If imggun.Left < 0 Then
        imggun.Left = frmpractice.Width - imggun.Left
            End If
            
End Sub

Private Sub tmrup_Timer()

'moving down
imggun.Top = imggun.Top - 250
    If imggun.Top < 0 Then
        imggun.Top = frmpractice.Height - imggun.Top
            End If
            
End Sub


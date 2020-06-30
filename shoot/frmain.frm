VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrcollision 
      Interval        =   1
      Left            =   4920
      Top             =   3360
   End
   Begin VB.Timer tmrtarget 
      Interval        =   10
      Left            =   840
      Top             =   240
   End
   Begin VB.Timer tmrfire 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   3360
   End
   Begin VB.Timer tmrfollow 
      Interval        =   1
      Left            =   240
      Top             =   3360
   End
   Begin VB.Label lblmiss 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblscore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape shptarget 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1440
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape shpbullet 
      Height          =   135
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpgun 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2520
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_load()

miss = 3

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    
    Case vbKeyRight
        shpgun.Left = shpgun.Left + 250
    Case vbKeyLeft
        shpgun.Left = shpgun.Left - 250
    Case vbKeySpace
        tmrfollow.Enabled = False
            tmrfire.Enabled = True
                shpbullet.Visible = True
                    Play App.Path & "\l.wav"
        
    End Select
    
End Sub

Private Sub tmrcollision_Timer()

Call collision(shpbullet, shptarget)

End Sub

Private Sub tmrfire_Timer()

shpbullet.Top = shpbullet.Top - 150
    If shpbullet.Top < 0 - shpbullet.Height Then
    
        miss = miss - 1
            lblmiss.Caption = miss
            
            If miss = 0 Then
                MsgBox "loser"
                    End
                    
                End If
                
        tmrfire.Enabled = False
            tmrfollow.Enabled = True
                shpbullet.Visible = False
                
        End If
End Sub

Private Sub tmrfollow_Timer()

shpbullet.Left = shpgun.Left + shpgun.Width / 2 - shpbullet.Width / 2
    shpbullet.Top = shpgun.Top
    
End Sub

Private Sub tmrtarget_Timer()

shptarget.Left = shptarget.Left - 50
    If shptarget.Left < 0 Then
        shptarget.Left = frmain.Width - shptarget.Left
            
        End If

End Sub

Sub collision(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
         
            shptarget.Left = 0 - shptarget.Width
                x = x + 1
                    lblscore.Caption = x
                        tmrfire.Enabled = False
                            tmrfollow.Enabled = True
                            
                End If
                End If
            
End Sub

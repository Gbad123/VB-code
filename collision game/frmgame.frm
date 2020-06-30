VERSION 5.00
Begin VB.Form frmgame 
   Caption         =   "Game"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   Picture         =   "frmgame.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer guncollision 
      Interval        =   1
      Left            =   5160
      Top             =   4680
   End
   Begin VB.Timer tmrcollide 
      Interval        =   1
      Left            =   5640
      Top             =   4680
   End
   Begin VB.Timer tmr2 
      Interval        =   10
      Left            =   9840
      Top             =   960
   End
   Begin VB.Timer tmr3 
      Interval        =   10
      Left            =   4200
      Top             =   2280
   End
   Begin VB.Timer tmr1 
      Interval        =   10
      Left            =   240
      Top             =   720
   End
   Begin VB.Timer tmrfire 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   4560
   End
   Begin VB.Timer tmrfollow 
      Interval        =   1
      Left            =   600
      Top             =   4560
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   4080
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   3600
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   4080
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   3600
   End
   Begin VB.Label lblscore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblmiss 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   5520
      Width           =   855
   End
   Begin VB.Shape shp3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape shp2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   9120
      Shape           =   3  'Circle
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shp1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgbullet 
      Height          =   600
      Left            =   2520
      Picture         =   "frmgame.frx":15CA
      Top             =   3600
      Width           =   585
   End
   Begin VB.Image imggun 
      Height          =   1755
      Left            =   2400
      Picture         =   "frmgame.frx":1BF3
      Top             =   4200
      Width           =   930
   End
End
Attribute VB_Name = "frmgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'moving properly
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

'start stopped
tmright.Enabled = False
    tmrleft.Enabled = False
        tmrup.Enabled = False
            tmrdown.Enabled = False
        
End Sub

Private Sub lbl1_Click()

'going back
Load frmain
    frmain.Show
        Unload frmgame
        
End Sub

Private Sub Form_Load()

'default stuff
m = 5
x = 0

End Sub

Private Sub guncollision_Timer()

'gun gets hit
Call collision2(imggun, shp1)
    Call collision2(imggun, shp2)
        Call collision2(imggun, shp3)

End Sub

Private Sub tmr1_Timer()

'target moving
shp1.Left = shp1.Left + 200
    If shp1.Left > frmgame.Width Then
        shp1.Left = 0 - shp1.Left
        
    End If

End Sub

Private Sub tmr2_Timer()

'target moving
shp2.Left = shp2.Left - 200
    If shp2.Left < 0 Then
        shp2.Left = frmgame.Width - shp2.Left
        
    End If
    
End Sub

Private Sub tmr3_Timer()

'target moving
shp3.Left = shp3.Left + 150
    If shp3.Left > frmgame.Width Then
        shp3.Left = 0 - shp3.Left
        
    End If
    
End Sub

Private Sub tmrcollide_Timer()

'collisions
Call collision(imgbullet, shp1)
    Call collision(imgbullet, shp2)
        Call collision(imgbullet, shp3)
        
End Sub

Private Sub tmrfire_Timer()

'fireing the bullet
imgbullet.Top = imgbullet.Top - 200
    If imgbullet.Top < 0 - imgbullet.Height Then
                
            'missing and showing the number of bullets left
            tmrfire.Enabled = False
                tmrfollow.Enabled = True
                    m = m - 1
             
            lblmiss.Caption = m
         
                'you missed to much LOSER
                If m = 0 Then
                    MsgBox "Loser", vbCritical, "LOSER"
                    
            'show how you did
            MsgBox "Your Hits and Misses", vbOKOnly, "End Game"
                MsgBox lblscore.Caption, vbOKOnly, "Hits"
                    MsgBox lblmiss.Caption, vbOKOnly, "Misses Left"
                        
                        'back to main
                        Load frmain
                            frmain.Show
                                Unload frmgame
                    
                End If
                
        'more for firing
        tmrfire.Enabled = False
            tmrfollow.Enabled = True
                imgbullet.Visible = False
                
        End If
        
End Sub

Private Sub tmrdown_Timer()

'moving down
imggun.Top = imggun.Top + 250
    If imggun.Top > frmgame.Height Then
        imggun.Top = 0 - frmgame.Top
        
    End If
    
End Sub

Private Sub tmrfollow_Timer()

'bullet stay with gun
imgbullet.Left = imggun.Left + imggun.Width / 2 - imgbullet.Width / 2
    imgbullet.Top = imggun.Top
    
End Sub

Private Sub tmright_Timer()

'moving right
imggun.Left = imggun.Left + 250
    If imggun.Left > frmgame.Width Then
        imggun.Left = 0 - imggun.Left
            End If

End Sub

Private Sub tmrleft_Timer()

'moving left
imggun.Left = imggun.Left - 250
    If imggun.Left < 0 Then
        imggun.Left = frmgame.Width - imggun.Left
            End If
            
End Sub

Private Sub tmrup_Timer()

'moving up
imggun.Top = imggun.Top - 250
    If imggun.Top < 0 Then
        imggun.Top = frmgame.Height - imggun.Top
            End If
            
End Sub

Sub collision(a, b)

'what happens with collision
If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
         
         'shapes reset
         shp1.Left = 0 - shp1.Width
            shp2.Left = 0 - shp2.Width
                shp3.Left = 0 - shp3.Width
                
        'shapes changing colours
        shp1.BackColor = vbBlue
            shp2.BackColor = vbGreen
                shp3.BackColor = vbYellow
        
        'shapes speeding up
        If lblscore.Caption = 5 Then
            shp1.Left = shp1.Left + 250
                shp2.Left = shp2.Left + 300
                    shp3.Left = shp3.Left - 200
                
        End If
                    
            Play App.Path & "\gun.wav" 'sound
            
            'more for collision
            tmrfollow.Enabled = ture
                tmrfire.Enabled = False
            
            'number of hits
            x = x + 1
            
                'display score
                lblscore.Caption = x
                
                    'when you win
                    If x = 10 Then
                        MsgBox "You Win", vbOKOnly, "WIN"
                        
            'show how you did
            MsgBox "Your Hits and Misses", vbOKOnly, "End Game"
                MsgBox lblscore.Caption, vbOKOnly, "Hits"
                    MsgBox lblmiss.Caption, vbOKOnly, "Misses Left"
                            
                            
                            'going back to main
                            Load frmain
                                frmain.Show
                                    Unload frmgame
                    
                    End If
                    
                            'more for collision
                    tmrfire.Enabled = False
                        tmrfollow.Enabled = True
                            
                End If
                End If
            
End Sub

Sub collision2(a, b)

'collision for gun
If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
    
    End
    
End If
    End If
    
End Sub


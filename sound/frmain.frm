VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrcollide 
      Interval        =   1
      Left            =   3480
      Top             =   1800
   End
   Begin VB.Timer tmrshp3 
      Interval        =   10
      Left            =   5040
      Top             =   4200
   End
   Begin VB.Timer tmrshp2 
      Interval        =   10
      Left            =   6240
      Top             =   2160
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   2160
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   2160
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   1200
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   3120
   End
   Begin VB.Shape shp3 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3120
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Shape shp2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   1455
      Left            =   6120
      Top             =   360
      Width           =   855
   End
   Begin VB.Shape shp1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyUp
        tmrup.Enabled = True
    Case vbKeyDown
        tmrdown.Enabled = True
    Case vbKeyRight
        tmright.Enabled = True
    Case vbKeyLeft
        tmrleft.Enabled = True
    Case vbKeyX
        End
    Case vbKeySpace
        Play App.Path & "\boing_2.wav"
        
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    tmrup.Enabled = False
    tmrdown.Enabled = False
    tmright.Enabled = False
    tmrleft.Enabled = False
    
End Sub

Private Sub tmrcollide_Timer()

Call collision(shp1, shp3)
    Call collision(shp1, shp2)
     
End Sub

Private Sub tmrdown_Timer()

shp1.Top = shp1.Top + 250
    If shp1.Top > frmain.Height Then
        shp1.Top = 0 - frmain.Height
            
            End If
            
End Sub

Private Sub tmright_Timer()

shp1.Left = shp1.Left + 250
    If shp1.Left > frmain.Height Then
        shp1.Left = 0 - frmain.Width
            
            End If
            
End Sub

Private Sub tmrleft_Timer()

shp1.Left = shp1.Left - 250
    If shp1.Left < 0 - shp1.Width Then
        shp1.Left = frmain.Width
        
            End If
        
End Sub

Private Sub tmrshp2_Timer()

shp2.Top = shp2.Top - 50
    If shp2.Top < 0 - shp1.Height Then
        shp2.Top = frmain.Height - shp2.Height
            
            End If
            
End Sub

Private Sub tmrshp3_Timer()

shp3.Left = shp3.Left - 50
    If shp3.Left < 0 - shp3.Width Then
        shp3.Left = frmain.Width
            
            End If
            
End Sub

Private Sub tmrup_Timer()

shp1.Top = shp1.Top - 250
    If shp1.Top < 0 - shp1.Height Then
        shp1.Top = frmain.Height - shp1.Height
        
            End If
            
End Sub

Sub collision(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then

Play App.Path & "\boing_2.wav"

        Load frmain2
            frmain2.Show
                Unload frmain
        
        
            End If
            End If
            
End Sub

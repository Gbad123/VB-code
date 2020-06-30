VERSION 5.00
Begin VB.Form frmain2 
   BackColor       =   &H0000FFFF&
   Caption         =   "Game"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrcollide 
      Interval        =   1
      Left            =   3240
      Top             =   1560
   End
   Begin VB.Timer tmrshp5 
      Interval        =   10
      Left            =   5880
      Top             =   3240
   End
   Begin VB.Timer tmrshp4 
      Interval        =   10
      Left            =   480
      Top             =   1200
   End
   Begin VB.Timer tmrimg2 
      Interval        =   10
      Left            =   4920
      Top             =   960
   End
   Begin VB.Timer tmrimg3 
      Interval        =   10
      Left            =   4080
      Top             =   4080
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2400
      Top             =   3240
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2400
      Top             =   2760
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   3240
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   2760
   End
   Begin VB.Shape shp5 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   5880
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape shp4 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   1  'Square
      Top             =   1800
      Width           =   615
   End
   Begin VB.Shape shp3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape shp2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4560
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape shp1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "frmain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyLeft
        tmrleft.Enabled = True
    Case vbKeyRight
        tmright.Enabled = True
    Case vbKeyDown
        tmrdown.Enabled = True
    Case vbKeyUp
        tmrup.Enabled = True
    Case vbKeySpace
        shp1.BackColor = vbBlue
        frmain.BackColor = vbRed
    Case vbKeyX
        End
        
End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    tmrleft.Enabled = False
    tmright.Enabled = False
    tmrdown.Enabled = False
    tmrup.Enabled = False
    
End Sub

Private Sub tmrcollide_Timer()

Call collision(shp1, shp2)
    Call collision(shp1, shp3)
        Call collision(shp1, shp4)
            Call collision(shp1, shp5)
            
End Sub

Private Sub tmrdown_Timer()

shp1.Top = shp1.Top + 350
    If shp1.Top > frmain.Height Then
        shp1.Top = 0 - shp1.Top
    End If
    
End Sub

Private Sub tmright_Timer()

shp1.Left = shp1.Left + 350
    If shp1.Left > frmain.Width Then
        shp1.Left = 0 - shp1.Left
    End If
    
End Sub
Private Sub tmrleft_Timer()

shp1.Left = shp1.Left - 350
    If shp1.Left < 0 - shp1.Width Then
        shp1.Left = frmain.Width
    End If

End Sub

Private Sub tmrshp5_Timer()

shp5.Top = shp5.Top + 500
    If shp5.Top > frmain.Height Then
        shp5.Top = 0 - shp5.Top
    End If
    
End Sub

Private Sub tmrup_Timer()

shp1.Top = shp1.Top - 350
    If shp1.Top < 0 - shp1.Height Then
        shp1.Top = frmain.Height
    End If

End Sub

Private Sub tmrimg2_Timer()

shp2.Left = shp2.Left - 500
    If shp2.Left < 0 - shp1.Width Then
        shp2.Left = frmain.Width
    End If
    
End Sub

Private Sub tmrimg3_Timer()

shp3.Left = shp3.Left + 500
    If shp3.Left > frmain.Width Then
        shp3.Left = 0 - shp3.Left
    End If
    
End Sub

Private Sub tmrshp4_Timer()

shp4.Top = shp4.Top - 500
    If shp4.Top < 0 - shp4.Height Then
        shp4.Top = frmain.Height
    End If
    
End Sub

Sub collision(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
    
        Play App.Path & "\gun.wav"
        
            Load frmain
                frmain.Show
                    Unload frmain2
            
                End If
                End If
            
End Sub


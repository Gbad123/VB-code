VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrshp4 
      Interval        =   10
      Left            =   1440
      Top             =   360
   End
   Begin VB.Timer tmrimg3 
      Interval        =   10
      Left            =   6600
      Top             =   480
   End
   Begin VB.Timer tmrimg2 
      Interval        =   10
      Left            =   240
      Top             =   1200
   End
   Begin VB.Timer tmrcollide 
      Interval        =   1
      Left            =   3480
      Top             =   1200
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   2520
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   3960
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   3240
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   3240
   End
   Begin VB.Shape shp4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   2040
      Top             =   360
      Width           =   735
   End
   Begin VB.Shape shp3 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape shp2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   600
      Shape           =   1  'Square
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shp1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      FillColor       =   &H0000FF00&
      Height          =   735
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmain"
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
        Shp1.BackColor = vbBlue
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

Call collision(Shp1, shp2)
    Call collision(Shp1, shp3)
        Call collision(Shp1, shp4)

End Sub
 
Private Sub tmrdown_Timer()

Shp1.Top = Shp1.Top + 350
    If Shp1.Top > frmain.Height Then
        Shp1.Top = 0 - Shp1.Top
    End If
    
End Sub

Private Sub tmright_Timer()

Shp1.Left = Shp1.Left + 350
    If Shp1.Left > frmain.Width Then
        Shp1.Left = 0 - Shp1.Left
    End If
    
End Sub

Private Sub tmrimg2_Timer()

shp2.Left = shp2.Left - 500
    If shp2.Left < 0 - Shp1.Width Then
        shp2.Left = frmain.Width
    End If
    
End Sub

Private Sub tmrimg3_Timer()

shp3.Left = shp3.Left + 500
    If shp3.Left > frmain.Width Then
        shp3.Left = 0 - shp3.Left
    End If
    
End Sub

Private Sub tmrleft_Timer()

Shp1.Left = Shp1.Left - 350
    If Shp1.Left < 0 - Shp1.Width Then
        Shp1.Left = frmain.Width
    End If

End Sub

Private Sub tmrshp4_Timer()

shp4.Top = shp4.Top - 500
    If shp4.Top < 0 - shp4.Height Then
        shp4.Top = frmain.Height
    End If
    
End Sub

Private Sub tmrup_Timer()

Shp1.Top = Shp1.Top - 350
    If Shp1.Top < 0 - Shp1.Height Then
        Shp1.Top = frmain.Height
    End If

End Sub

Sub collision(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
    
        Play App.Path & "\gun.wav"
        
            End
            
                End If
                End If
            
End Sub

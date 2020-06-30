VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrcollision 
      Interval        =   1
      Left            =   3480
      Top             =   2040
   End
   Begin VB.Timer tmrimg2 
      Interval        =   10
      Left            =   240
      Top             =   1680
   End
   Begin VB.Timer tmrimg3 
      Interval        =   10
      Left            =   5280
      Top             =   840
   End
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   4560
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   3240
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   3840
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   3960
   End
   Begin VB.Shape shp2 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape shp3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   240
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape shp1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   975
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
    Case vbKeyUp
        tmrup.Enabled = True
    Case vbKeyDown
        tmrdown.Enabled = True
    Case vbKeyX
        End
    
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    tmrleft.Enabled = False
    tmright.Enabled = False
    tmrup.Enabled = False
    tmrdown.Enabled = False
    
End Sub

Private Sub tmrcollision_Timer()

Call collision(shp1, shp2)
    Call collision(shp1, shp3)
    
End Sub

Private Sub tmrdown_Timer()

shp1.Top = shp1.Top + 250
    If shp1.Top > frmain.Height Then
        shp1.Top = 0 - shp1.Top
    End If
    
End Sub

Private Sub tmright_Timer()

shp1.Left = shp1.Left + 250
    If shp1.Left > frmain.Width Then
        shp1.Left = 0 - shp1.Left
    End If
    
End Sub

Private Sub tmrimg2_Timer()

shp3.Left = shp3.Left + 50
    If shp3.Left > frmain.Width Then
        shp3.Left = 0 - shp3.Left
    End If
    
End Sub

Private Sub tmrimg3_Timer()

shp2.Left = shp2.Left - 50
    If shp2.Left < 0 - shp1.Width Then
        shp2.Left = frmain.Width
    End If
    
End Sub

Private Sub tmrleft_Timer()

shp1.Left = shp1.Left - 250
    If shp1.Left < 0 - shp1.Width Then
        shp1.Left = frmain.Width
    End If
    
End Sub

Private Sub tmrup_Timer()

shp1.Top = shp1.Top - 250
    If shp1.Top < 0 - shp1.Top Then
        shp1.Top = frmain.Height
    End If
    
End Sub

Sub collision(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
    
        End
        
            End If
            End If
            
End Sub


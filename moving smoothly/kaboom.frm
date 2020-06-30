VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   3600
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   720
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5040
      Top             =   2400
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   2640
   End
   Begin VB.Shape shp1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   1815
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
        shp1.BackColor = vbRed
        frmain.BackColor = vbBlue
    Case vbKeyX
        End
    
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    tmrup.Enabled = False
    tmrdown.Enabled = False
    tmright.Enabled = False
    tmrleft.Enabled = False
    
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

Private Sub tmrleft_Timer()

shp1.Left = shp1.Left - 250
    If shp1.Left < 0 - shp1.Width Then
        shp1.Left = frmain.Width
    End If
        
End Sub

Private Sub tmrup_Timer()

shp1.Top = shp1.Top - 250
    If shp1.Top < 0 - shp1.Height Then
        shp1.Top = frmain.Height - shp1.Height
    End If
    
End Sub

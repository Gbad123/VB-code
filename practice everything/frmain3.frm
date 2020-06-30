VERSION 5.00
Begin VB.Form frmain3 
   BackColor       =   &H0000FFFF&
   Caption         =   "Controls"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6120
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrdown 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   3120
   End
   Begin VB.Timer tmrup 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   2520
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   3120
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   2520
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Left to move left, Right to move right, Up to move up, down to move down, x to end game."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape shp1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frmain3"
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

Private Sub lbl2_Click()

Load frmain
    frmain.Show
        Unload frmain3
        
Play App.Path & "\boing_2.wav"

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

Private Sub tmrshp4_Timer()

shp4.Top = shp4.Top - 500
    If shp4.Top < 0 - shp4.Height Then
        shp4.Top = frmain.Height
    End If
    
End Sub

Private Sub tmrup_Timer()

shp1.Top = shp1.Top - 350
    If shp1.Top < 0 - shp1.Height Then
        shp1.Top = frmain.Height
    End If

End Sub

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Snake Game"
   ClientHeight    =   9675
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrgrow 
      Interval        =   50
      Left            =   1440
      Top             =   3720
   End
   Begin VB.Timer tmrfollow 
      Interval        =   100
      Left            =   6000
      Top             =   2040
   End
   Begin VB.Timer tmrside 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6360
      Top             =   3000
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6000
      Top             =   2640
   End
   Begin VB.Timer tmr4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5640
      Top             =   3000
   End
   Begin VB.Timer tmr3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6000
      Top             =   3360
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   9600
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape shp2 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1920
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim length As Integer
    Dim tail As Integer
    Dim y As Integer
    Dim coun As Integer
    Dim hit As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyUp
        tmr1.Enabled = True
        tmr2.Enabled = False
        tmr3.Enabled = False
        tmr4.Enabled = False
    Case vbKeyDown
        tmr1.Enabled = False
        tmr2.Enabled = False
        tmr3.Enabled = True
        tmr4.Enabled = False
    Case vbKeyLeft
        tmr1.Enabled = False
        tmr2.Enabled = False
        tmr3.Enabled = False
        tmr4.Enabled = True
    Case vbKeyRight
        tmr1.Enabled = False
        tmr2.Enabled = True
        tmr3.Enabled = False
        tmr4.Enabled = False
        
End Select

End Sub

Private Sub tmr1_Timer()

shp1(0).Top = shp1(0).Top - shp1(0).Height

End Sub

Private Sub tmr2_Timer()

shp1(0).Left = shp1(0).Left + shp1(0).Width

End Sub

Private Sub tmr3_Timer()

shp1(0).Top = shp1(0).Top + shp1(0).Height

End Sub

Private Sub tmr4_Timer()

shp1(0).Left = shp1(0).Left - shp1(0).Width

End Sub

Private Sub tmrfollow_Timer()

If length > 0 Then
    Dim tail As Integer
    tail = length - 1

        For y = 0 To tail
    
            shp1(length - y).Left = shp1(tail - y).Left
            shp1(length - y).Top = shp1(tail - y).Top
                
        Next y
End If

End Sub

Private Sub tmrgrow_Timer()

If shp1(0).Left > shp2.Left - shp1(0).Width And shp1(0).Left < shp2.Left + shp2.Width Then
   If shp1(0).Top < shp2.Top + shp2.Height And shp1(0).Top > shp2.Top - shp2.Height Then
    
length = length + 1
        
lbl1.Caption = lbl1.Caption + 1


coun = coun + 1

        Load shp1(length)
        shp1(length).Visible = True

    shp2.Left = Rnd * Form1.Width
    shp2.Top = Rnd * Form1.Height
        
Exit Sub

    End If
End If
    
    
For hit = 5 To length

    If shp1(0).Left > shp1(hit - 1).Left - shp1(hit - 1).Width And shp1(0).Left < shp1(hit - 1).Left + shp1(hit - 1).Width Then
        If shp1(hit - 1).Top < shp1(0).Top + shp1(0).Height And shp1(hit - 1).Top > shp1(0).Top - shp1(hit - 1).Height Then
    
            MsgBox "You Lose", vbCritical, "Loser"
    
            End
    
        End If
    End If
    
Next hit

End Sub

Private Sub tmrside_Timer()

If shp1(0).Left < 0 Then
MsgBox "lose", vbOKOnly, "Loser"
End
End If

If shp1(0).Left > Form1.Width Then
MsgBox "you lose", vbOKOnly, "Loser"
End
End If

If shp1(0).Top < 0 Then
MsgBox "you lose", vbOKOnly, "Loser"
End
End If

If shp1(0).Top > Form1.Height Then
MsgBox "you lose", vbOKOnly, "Loser"
End
End If

End Sub

VERSION 5.00
Begin VB.Form frmain 
   BackColor       =   &H8000000E&
   Caption         =   "Main"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   FillColor       =   &H00C0C0FF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Bet"
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Timer tmr4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   1080
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Start"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmr3 
      Enabled         =   0   'False
      Interval        =   90
      Left            =   4680
      Top             =   1440
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2520
      Top             =   1440
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   1440
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Place Bet"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Money Own"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "$5000"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Image img3 
      Height          =   735
      Left            =   4320
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image img2 
      Height          =   735
      Left            =   2280
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image img1 
      Height          =   735
      Left            =   360
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Integer 'slot1
        Dim y As Integer 'slot2
            Dim z As Integer 'slot3
                Dim b As Integer 'won
                    Dim s As Integer 'money
                        Dim t As Integer 'bet

Private Sub cmd1_Click()

'start everthing
tmr1.Enabled = True
    tmr2.Enabled = True
        tmr3.Enabled = True
            tmr4.Enabled = True
            
'winnings
If x + y + z = 3 Then
    lbl2.Caption = "$" & s + 5000
End If
                
If x + y + z = 6 Then
    lbl2.Caption = "$" & s + 6000
End If

If x + y + z = 9 Then
    lbl2.Caption = "$" & s + 7000
End If
              
End Sub

Private Sub cmd2_Click()

'have the start button appear
cmd1.Visible = True
    cmd1.Enabled = True

'score keeper
s = lbl2.Caption
t = Val(txt1.Text)
lbl2.Caption = "$" & s - t
    
    'when low on money
    If lbl2.Caption < 500 Then
        MsgBox "Low on Money", vbCritical, "Money Problems"
    End If
    
'when out of money
If lbl2.Caption < 0 Then
    MsgBox "You Lose", vbCritical, "Loser"
        End
End If

'when too much money
If lbl2.Caption >= 15000 Then
    MsgBox "You Win", vbCritical, "Congradulations"
        End
End If

'count down
lbl1.Caption = 10

End Sub

Private Sub cmdend_Click()

'ending comments
MsgBox lbl2.Caption, vbOKOnly, "Your Earnings"
    End

End Sub

Private Sub tmr1_Timer()

'slot 1
x = (Rnd(10) * 3) + 1

    If x = 1 Then
        img1.Picture = LoadPicture(App.Path & "\cherry.bmp")
    End If
    
    If x = 2 Then
        img1.Picture = LoadPicture(App.Path & "\apple.bmp")
    End If
    
    If x = 3 Then
        img1.Picture = LoadPicture(App.Path & "\7.bmp")
    End If
    
End Sub

Private Sub tmr2_Timer()

'slot 2
y = (Rnd(10) * 3) + 1
    
    If y = 1 Then
        img2.Picture = LoadPicture(App.Path & "\cherry.bmp")
    End If
    
    If y = 2 Then
        img2.Picture = LoadPicture(App.Path & "\apple.bmp")
    End If
    
    If y = 3 Then
        img2.Picture = LoadPicture(App.Path & "\7.bmp")
    End If
    
End Sub

Private Sub tmr3_Timer()

'slot 3
z = (Rnd(10) * 3) + 1
    
    If z = 1 Then
        img3.Picture = LoadPicture(App.Path & "\cherry.bmp")
    End If
    
    If z = 2 Then
        img3.Picture = LoadPicture(App.Path & "\apple.bmp")
    End If
    
    If z = 3 Then
        img3.Picture = LoadPicture(App.Path & "\7.bmp")
    End If
    
End Sub

Private Sub tmr4_Timer()

'timer and everything stop
lbl1.Caption = lbl1.Caption - 1

    If lbl1.Caption = 0 Then
        tmr1.Enabled = False
            tmr2.Enabled = False
                tmr3.Enabled = False
                    tmr4.Enabled = False
                        cmd1.Enabled = False
                            cmd1.Value = False
    End If

End Sub

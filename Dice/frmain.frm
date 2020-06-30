VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   2040
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   2040
   End
   Begin VB.CommandButton cmdroll 
      Caption         =   "Roll"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5000"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image img2 
      Height          =   855
      Left            =   2640
      Top             =   960
      Width           =   975
   End
   Begin VB.Image img1 
      Height          =   855
      Left            =   600
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dim stuff
Option Explicit
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    Dim b As Integer
    Dim q As Integer
    Dim m As Integer
    Dim l As Integer
    
Private Sub cmdroll_Click()

'turn timer on
tmr1.Enabled = True
tmr2.Enabled = True

End Sub

Private Sub cmdstop_Click()

'all beepin posibilites for x & y
Select Case x & y

    Case 1 & 1
        MsgBox "roll again"
    Case 2 & 2
        MsgBox "roll again"
    Case 3 & 3
        MsgBox "roll again"
    Case 4 & 4
        MsgBox "roll again"
    Case 5 & 5
        MsgBox "roll again"
    Case 6 & 6
        MsgBox "roll again"
    
    Case 1 & 6
        MsgBox "you win"
            lbl3.Caption = "$" & q + 1000
    Case 6 & 1
        MsgBox "you win"
            lbl3.Caption = "$" & q + 1000
    Case 4 & 3
        MsgBox "you win"
            lbl3.Caption = "$" & q + 1000
    Case 3 & 4
        MsgBox "you win"
            lbl3.Caption = "$" & q + 1000
    Case 5 & 2
        MsgBox "you win"
            lbl3.Caption = "$" & q + 1000
    Case 2 & 5
        MsgBox "you win"
            lbl3.Caption = "$" & q + 1000
        
    Case 1 & 2
        MsgBox "Loser"
    Case 1 & 3
        MsgBox "Loser"
    Case 1 & 4
        MsgBox "Loser"
    Case 1 & 5
        MsgBox "Loser"
    
    Case 2 & 1
        MsgBox "Loser!"
    Case 2 & 3
        MsgBox "Loser!"
    Case 2 & 4
        MsgBox "Loser!"
    Case 2 & 6
        MsgBox "Loser!"
    
    Case 3 & 1
        MsgBox "Loser?"
    Case 3 & 2
        MsgBox "Loser?"
    Case 3 & 5
        MsgBox "Loser?"
    Case 3 & 6
        MsgBox "Loser?"
    
    Case 4 & 1
        MsgBox "(Loser)"
    Case 4 & 2
        MsgBox "(Loser)"
    Case 4 & 6
        MsgBox "(Loser)"
    Case 4 & 5
        MsgBox "(Loser)"
    
    Case 5 & 1
        MsgBox "Loser :o"
    Case 5 & 3
        MsgBox "Loser :o"
    Case 5 & 4
        MsgBox "Loser :o"
    Case 5 & 6
        MsgBox "Loser :o"
    
    Case 6 & 2
        MsgBox "Loser :p"
    Case 6 & 3
        MsgBox "Loser :p"
    Case 6 & 4
        MsgBox "Loser :p"
    Case 6 & 5
        MsgBox "Loser :p"
        
End Select

'timers off
tmr1.Enabled = False
tmr2.Enabled = False

'adding $
If x + y = 7 Then
    lbl3.Caption = "$" & q + 1000
End If

'lossing $
l = Val(txt1.Text)
m = lbl3.Caption
lbl3.Caption = m - l

End Sub

Private Sub tmr1_Timer()

'pictures for x
z = (Rnd(10) * 5) + 1

If z = 1 Then
    img1.Picture = LoadPicture(App.Path & "\1.bmp")
End If

If z = 2 Then
    img1.Picture = LoadPicture(App.Path & "\2.bmp")
End If

If z = 3 Then
    img1.Picture = LoadPicture(App.Path & "\3.bmp")
End If

If z = 4 Then
    img1.Picture = LoadPicture(App.Path & "\4.bmp")
End If

If z = 5 Then
    img1.Picture = LoadPicture(App.Path & "\5.bmp")
End If

If z = 6 Then
    img1.Picture = LoadPicture(App.Path & "\6.bmp")
End If

x = z

End Sub

Private Sub tmr2_Timer()

'pictures y
b = (Rnd(10) * 5) + 1

If b = 1 Then
    img2.Picture = LoadPicture(App.Path & "\1.bmp")
End If

If b = 2 Then
    img2.Picture = LoadPicture(App.Path & "\2.bmp")
End If

If b = 3 Then
    img2.Picture = LoadPicture(App.Path & "\3.bmp")
End If

If b = 4 Then
    img2.Picture = LoadPicture(App.Path & "\4.bmp")
End If

If b = 5 Then
    img2.Picture = LoadPicture(App.Path & "\5.bmp")
End If

If b = 6 Then
    img2.Picture = LoadPicture(App.Path & "\6.bmp")
End If

y = b

End Sub


VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   Picture         =   "frmain.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrpikachu2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   360
      Top             =   6000
   End
   Begin VB.Timer tmrfire 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8400
      Top             =   4320
   End
   Begin VB.Timer tmrfollow 
      Interval        =   1
      Left            =   7800
      Top             =   4320
   End
   Begin VB.Timer tmrcollision 
      Interval        =   1
      Left            =   3000
      Top             =   3840
   End
   Begin VB.Timer tmrpikachu 
      Interval        =   150
      Left            =   360
      Top             =   5280
   End
   Begin VB.Timer tmrsnipe 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9240
      Top             =   3600
   End
   Begin VB.Timer tmrdrop 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7200
      Top             =   6000
   End
   Begin VB.Timer tmrjump 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7680
      Top             =   6000
   End
   Begin VB.Timer tmrcloud 
      Interval        =   150
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer tmright 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8880
      Top             =   6000
   End
   Begin VB.Timer tmrleft 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8400
      Top             =   6000
   End
   Begin VB.Image imgpikachu2 
      Height          =   345
      Left            =   1320
      Picture         =   "frmain.frx":37C5
      Top             =   5760
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8880
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgpokeball 
      Height          =   300
      Left            =   6120
      Picture         =   "frmain.frx":3C6C
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgpikachu 
      Height          =   810
      Left            =   1320
      Picture         =   "frmain.frx":40C3
      Top             =   4560
      Width           =   570
   End
   Begin VB.Image imglatios 
      Height          =   885
      Left            =   8640
      Picture         =   "frmain.frx":42E0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgcloud3 
      Height          =   450
      Left            =   3120
      Picture         =   "frmain.frx":4BD4
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imgcloud2 
      Height          =   450
      Left            =   1080
      Picture         =   "frmain.frx":533C
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image imgcloud 
      Height          =   450
      Left            =   120
      Picture         =   "frmain.frx":5AA4
      Top             =   360
      Width           =   855
   End
   Begin VB.Image img1 
      Height          =   960
      Left            =   7920
      Picture         =   "frmain.frx":620C
      Top             =   4920
      Width           =   690
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim l As Integer 'left
        Dim r As Integer 'right
            Dim c As Integer 'up
                Dim d As Integer 'down
                    Dim p As Integer 'pikachu
                        Dim a As Integer 'pikachu2
                            Dim s As Integer 'score

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    
    Case vbKeyLeft
        tmrleft.Enabled = True
    Case vbKeyRight
        tmright.Enabled = True
    Case vbKeySpace
        tmrjump.Enabled = True
    Case vbKeyX
        tmrfire.Enabled = True
            Play App.Path & "\boing_2.wav"
    End Select
            
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    
    Case vbKeyLeft
        tmrleft.Enabled = False
    Case vbKeyRight
        tmright.Enabled = False
    Case vbKeySpace
        tmrjump.Enabled = False

End Select

End Sub

Private Sub tmrcloud_Timer()

imgcloud.Left = imgcloud.Left + 50
    If imgcloud.Left > frmain.Width Then
        imgcloud.Left = 0 - imgcloud.Width
    End If
    
imgcloud2.Left = imgcloud2.Left + 50
    If imgcloud2.Left > frmain.Width Then
        imgcloud2.Left = 0 - imgcloud2.Width
    End If
    
imgcloud3.Left = imgcloud3.Left + 50
    If imgcloud3.Left > frmain.Width Then
        imgcloud3.Left = 0 - imgcloud3.Width
    End If
    
End Sub

Private Sub tmrcollision_Timer()

Call collision(img1, imgpikachu)
    Call collision(img1, imgpikachu2)
    
Call collide(imgpokeball, imgpikachu)
    Call collide(imgpokeball, imgpikachu2)
    
End Sub

Private Sub tmrdrop_Timer()

d = d + 1
    If d = 5 Then
        d = 0
        tmrdrop.Enabled = False
    Else
        img1.Top = img1.Top + 100
    
    End If
    
End Sub

Private Sub tmrfire_Timer()

imgpokeball.Visible = True
    tmrfollow.Enabled = False
    
imgpokeball.Left = imgpokeball.Left - 500

    If imgpokeball.Left < 0 - imgpokeball.Width Then
        
        tmrfire.Enabled = False
            tmrfollow.Enabled = True
                imgpokeball.Visible = False
    End If
    
End Sub

Private Sub tmrfollow_Timer()

imgpokeball.Left = img1.Left = img1.Width / 2 - imgpokeball.Width / 2
    imgpokeball.Top = img1.Top
        imgpokeball.Left = img1.Left
    
End Sub

Private Sub tmright_Timer()

r = r + 1
    img1.Picture = LoadPicture(App.Path & ("\" & r & "r.gif"))
        If r = 4 Then r = -1
    
    imgcloud.Left = imgcloud.Left + 40
        imgcloud2.Left = imgcloud2.Left + 40
            imgcloud3.Left = imgcloud3.Left + 40
        
End Sub

Private Sub tmrjump_Timer()

c = c + 1
    If c = 5 Then
        c = 0
        tmrjump.Enabled = False
            tmrdrop.Enabled = True
    Else
        img1.Top = img1.Top - 100
    
    End If
End Sub

Private Sub tmrleft_Timer()

l = l + 1
    img1.Picture = LoadPicture(App.Path & ("\" & l & "l.gif"))
        If l = 4 Then l = -1
    
    imgcloud.Left = imgcloud.Left + 10
        imgcloud2.Left = imgcloud2.Left + 10
            imgcloud3.Left = imgcloud3.Left + 10
    
End Sub

Private Sub tmrpikachu_Timer()

p = p + 1
    imgpikachu.Picture = LoadPicture(App.Path & ("\" & p & "p.gif"))
        If p = 3 Then p = 0
    
    imgpikachu.Left = imgpikachu.Left + 100
    
            If imgpikachu.Left > frmain.Width Then
                imgpikachu.Left = 0 - imgpikachu.Width
            
            End If
            
            If s >= 15 Then
                imgpikachu.Left = imgpikachu.Left + 200
                
            End If
            
End Sub


Private Sub tmrpikachu2_Timer()

a = a + 1
    imgpikachu2.Picture = LoadPicture(App.Path & ("\" & a & "a.gif"))
        If a = 3 Then a = 0
    
    imgpikachu2.Left = imgpikachu2.Left + 100
    
            If imgpikachu2.Left > frmain.Width Then
                imgpikachu2.Left = 0 - imgpikachu2.Width
            
            End If
            
            If s >= 15 Then
                imgpikachu2.Left = imgpikachu2.Left + 250
                
            End If
            
End Sub

Private Sub tmrsnipe_Timer()
    
imglatios.Left = imglatios.Left - 250

End Sub

Sub collision(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
        
            End
            
                End If
                End If
            
End Sub

Sub collide(a, b)

If a.Left > b.Left - a.Width And a.Left < b.Left + b.Width Then
    If a.Top < b.Top + b.Height And a.Top > b.Top - a.Height Then
        
        b.Left = 0 - b.Width
            s = s + 1
                lbl1.Caption = s
                
If s >= 25 Then
    tmrsnipe.Enabled = True
        imglatios.Visible = True
         
End If

If s >= 20 Then
    tmrpikachu2.Enabled = True
        imgpikachu2.Visible = True
        
End If
            
If s >= 50 Then
    MsgBox "You Win!", vbOKOnly, "Winner"
        End
     
End If
                End If
                End If
                
End Sub

VERSION 5.00
Begin VB.Form frmain 
   BackColor       =   &H00FFFF00&
   Caption         =   "Main"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtotal 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cost"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame frm5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Colour"
      Height          =   3255
      Left            =   3840
      TabIndex        =   17
      Top             =   480
      Width           =   1215
      Begin VB.OptionButton opt15 
         BackColor       =   &H00FFFF00&
         Caption         =   "Black"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   975
      End
      Begin VB.OptionButton opt14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Red"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton opt13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Blue"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.PictureBox pic1 
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frm4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Chevrolet"
      Height          =   3255
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   1215
      Begin VB.OptionButton opt12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sonic"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   855
      End
      Begin VB.OptionButton opt11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cruze"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton opt10 
         BackColor       =   &H00FFFF00&
         Caption         =   "Spark"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame frm3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Toyota"
      Height          =   3255
      Left            =   2040
      TabIndex        =   8
      Top             =   480
      Width           =   1215
      Begin VB.OptionButton opt9 
         BackColor       =   &H00FFFF00&
         Caption         =   "Camry"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   735
      End
      Begin VB.OptionButton opt8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Lander"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton opt7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Matric"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame frm2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Honda"
      Height          =   3255
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
      Begin VB.OptionButton opt6 
         BackColor       =   &H00FFFF00&
         Caption         =   "CRV"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin VB.OptionButton opt5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Accord"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton opt4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Civic"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Dealer"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
      Begin VB.OptionButton opt3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Chevrolet"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Toyota"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Honda"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label lbltotal 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   5400
      TabIndex        =   21
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Shape shp1 
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   3840
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgaccord 
      Height          =   1095
      Left            =   2040
      Picture         =   "frmain.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgcrv 
      Height          =   1095
      Left            =   2040
      Picture         =   "frmain.frx":25026
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgcivic 
      Height          =   1095
      Left            =   2040
      Picture         =   "frmain.frx":4A190
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image img3 
      Height          =   1095
      Left            =   120
      Picture         =   "frmain.frx":6F0EA
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image img1 
      Height          =   1020
      Left            =   120
      Picture         =   "frmain.frx":7F1C4
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Image img2 
      Height          =   1035
      Left            =   120
      Picture         =   "frmain.frx":93166
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Integer
        Dim y As Integer
            Dim z As Integer

Private Sub cmdtotal_Click()

'total cost
If opt4.Value = True Then
    x = 1400
        End If
    
If opt5.Value = True Then
    x = 1500
        End If

If opt6.Value = True Then
    x = 1550
        End If

If opt7.Value = True Then
    x = 1460
        End If
        
If opt8.Value = True Then
    x = 1500
        End If
        
If opt9.Value = True Then
    x = 1600
        End If
    
If opt10.Value = True Then
    x = 1480
        End If
    
If opt11.Value = True Then
    x = 1560
        End If
        
If opt12.Value = True Then
    x = 1660
        End If

If opt13.Value = True Then
    y = 200
        End If
        
If opt14.Value = True Then
    y = 150
        End If

If opt15.Value = True Then
    y = 125
        End If

z = x + y
    lbltotal.Caption = "$" & z
    
End Sub

Private Sub opt1_Click()

'what happens when you select honda
If opt1.Value = True Then
    img1.Visible = True 'honda
        img2.Visible = False 'chevrolet
            img3.Visible = False 'toyota
            
                frm2.Visible = True
                    frm3.Visible = False
                        frm4.Visible = False
                            frm5.Visible = False
                        
                                pic1.Visible = False
            
End If

End Sub

Private Sub opt10_Click()

'what happenes when you select spark
If opt10.Value = True Then
    pic1.Visible = True
        pic1.Picture = LoadPicture(App.Path & "\spark.bmp")

frm5.Visible = True

End If

End Sub

Private Sub opt11_Click()

'what happens when you select cruze
If opt11.Value = True Then
    pic1.Visible = True
        pic1.Picture = LoadPicture(App.Path & "\cruze.bmp")

frm5.Visible = True

End If

End Sub

Private Sub opt12_Click()

'what happenes when you select sonic
If opt12.Value = True Then
    pic1.Visible = True
        pic1.Picture = LoadPicture(App.Path & "\sonic.bmp")

frm5.Visible = True

End If

End Sub

Private Sub opt13_Click()

'what happenes when you select blue
If opt13.Value = True Then
    shp1.Visible = True
        shp1.BackColor = vbBlue

End If

End Sub

Private Sub opt14_Click()

'what happenes when you select red
If opt14.Value = True Then
    shp1.Visible = True
        shp1.BackColor = vbRed

End If

End Sub

Private Sub opt15_Click()

'what happenes when you select black
If opt15.Value = True Then
    shp1.Visible = True
        shp1.BackColor = vbBlack
    
End If

End Sub

Private Sub opt2_Click()

'what happends when you select toyota
If opt2.Value = True Then
    img3.Visible = True 'toyota
        img1.Visible = False 'honda
            img2.Visible = False 'chevrolet
            
                frm2.Visible = False
                    frm3.Visible = True
                        frm4.Visible = False
                            frm5.Visible = False
                    
                                imgaccord.Visible = False
                                    imgcrv.Visible = False
                                        imgcivic.Visible = False
                                        
                                            pic1.Visible = False
                            
                            
End If

End Sub

Private Sub opt3_Click()

'what happends when you select chevrolet
If opt3.Value = True Then
    img2.Visible = True 'chevrolet
        img1.Visible = False 'honda
            img3.Visible = False 'toyta
            
                frm2.Visible = False
                    frm3.Visible = False
                        frm4.Visible = True
                            frm5.Visible = False
                    
                                imgaccord.Visible = False
                                    imgcrv.Visible = False
                                        imgcivic.Visible = False
                                    
                                            pic1.Visible = False

End If

End Sub

Private Sub opt4_Click()

'what happends whenyou select civic
If opt4.Value = True Then
    imgcivic.Visible = True
        imgcrv.Visible = False
            imgaccord.Visible = False
            
frm5.Visible = True

End If

End Sub

Private Sub opt5_Click()

'what happens when you select accord
If opt5.Value = True Then
    imgaccord.Visible = True
        imgcrv.Visible = False
            imgcivic.Visible = False

frm5.Visible = True

End If

End Sub

Private Sub opt6_Click()

'what happenes when you select crv
If opt6.Value = True Then
    imgcrv.Visible = True
        imgaccord.Visible = False
            imgcivic.Visible = False
    
frm5.Visible = True

End If

End Sub

Private Sub opt7_Click()

'what happens when you select matric
If opt7.Value = True Then
    pic1.Visible = True
        pic1.Picture = LoadPicture(App.Path & "\matric.bmp")
    
frm5.Visible = True

End If

End Sub

Private Sub opt8_Click()

'what happends when you select highlander
If opt8.Value = True Then
    pic1.Visible = True
        pic1.Picture = LoadPicture(App.Path & "/highlander.bmp")

frm5.Visible = True

End If

End Sub

Private Sub opt9_Click()

'what happends when you select camry
If opt9.Value = True Then
    pic1.Visible = True
        pic1.Picture = LoadPicture(App.Path & "\camry.bmp")
        
frm5.Visible = True

End If

End Sub


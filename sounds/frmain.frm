VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgun 
      Caption         =   "gun"
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdsound 
      Caption         =   "boing"
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgun_Click()

Play App.Path & "\gun.wav"

Load frmain3
frmain3.Show
Unload frmain

End Sub

Private Sub cmdsound_Click()

Play App.Path & "\boing_2.wav"

Load frmain2
frmain2.Show
Unload frmain

End Sub

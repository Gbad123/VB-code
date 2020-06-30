VERSION 5.00
Begin VB.Form fmrain 
   BackColor       =   &H80000007&
   Caption         =   "Main"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pickitten 
      Height          =   3015
      Left            =   3480
      Picture         =   "fmrain.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton cmdhelloworld 
      Caption         =   "Hello World"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Cmdend 
      Caption         =   "&End"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image imgkitten 
      Height          =   3270
      Left            =   0
      Picture         =   "fmrain.frx":140F
      Top             =   -240
      Width           =   3465
   End
End
Attribute VB_Name = "fmrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myname As String

Private Sub Cmdend_Click()

End

End Sub

Private Sub cmdhelloworld_Click()

myname = InputBox(" Enter your name", "Your Name", "Ebad Babar")
MsgBox myname, vbOKOnly, "Your name"

End Sub

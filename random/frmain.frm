VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr1 
      Interval        =   500
      Left            =   2520
      Top             =   3000
   End
   Begin VB.Label lbl1 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Integer

Private Sub tmr1_Timer()

x = (Rnd(10) * 3) + 1
    lbl1.Caption = x

End Sub

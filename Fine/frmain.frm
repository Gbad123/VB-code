VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "Go"
      Height          =   855
      Left            =   2400
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txt2 
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txt1 
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblanswer 
      BackStyle       =   0  'Transparent
      Height          =   3615
      Left            =   7200
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lbl2 
      Caption         =   "Driver Speed"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Caption         =   "Speed Limit"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    
Private Sub cmd1_Click()

a = txt1.Text
b = txt2.Text

c = b - a

    If c < 0 Then lblanswer.Caption = "you are within the limit"
    If c > 0 And c < 21 Then lblanswer.Caption = "$100 fine"
    If c > 20 And c < 31 Then lblanswer.Caption = "$270 fine"
    If c < 30 Then lblanswer.Caption = "$500 fine"
    
End Sub

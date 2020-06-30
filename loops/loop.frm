VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "go"
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txt1 
      Height          =   1335
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Integer
        Dim c As String
            Dim y As Integer

Private Sub cmd1_Click()

x = InputBox("whatever")

For y = 1 To x

c = c + "*"

    txt1.Text = txt1.Text & vbCrLf & c
    
Next y

End Sub

VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   480
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
    Dim y As Integer
    Dim s As Integer
    Dim b As Integer
    Dim q As Integer
        Dim a As Boolean
    
Private Sub cmdgo_Click()


x = Val(txt1.Text)
    y = Mid(x, 1, 1)
    s = Mid(x, 2, 1)
    b = Mid(x, 3, 1)
    q = Mid(x, 4, 1)
    
If y = s Or y = b Or y = q & s = b Or s = q Or s = y & b = q Or b = s Or b = y Then
    Print "no"

Else
    Print "yes"
    
End If

End Sub

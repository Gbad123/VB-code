VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "Go"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
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
            Dim c As String

Private Sub cmd1_Click()

c = "*"

For x = 10 To 1 Step -1
    c = c + "*"
    Print c; x
    
        For y = 1 To 3
            Print y
        Next y
        
Next x
    
End Sub

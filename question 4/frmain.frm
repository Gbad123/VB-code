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
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim d(1 To 4) As Integer
    Dim x As Integer
        Dim a As Boolean
            Dim s As String
    
Private Sub cmdgo_Click()

a = False

s = InputBox("")

    Do While a = False
    
        For x = 1 To 4
            d(x) = Mid(s, x, 1)
        Next x
        
    If d(1) = d(2) Then
        s = s + 1
        
        ElseIf d(1) = d(3) Then s = s + 1
        ElseIf d(1) = d(4) Then s = s + 1
        ElseIf d(2) = d(3) Then s = s + 1
        ElseIf d(2) = d(4) Then s = s + 1
        ElseIf d(3) = d(4) Then s = s + 1
        
        Else
            
            a = True
            
            MsgBox s
            
    End If
    
    Loop
    
End Sub

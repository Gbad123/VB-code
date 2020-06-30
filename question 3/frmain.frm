VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim s As String
        Dim x As Integer
        Dim v As Integer
    Dim y As String

Private Sub cmdgo_Click()

s = InputBox("enter")
v = Len(s)

For x = 1 To v
    y = Mid(s, x, 1)
    
    If y = "O" Then
        ElseIf y = "Z" Then
        ElseIf y = "N" Then
        ElseIf y = "S" Then
        ElseIf y = "H" Then
        ElseIf y = "X" Then
        ElseIf y = "I" Then
        
        Else
            MsgBox "NO"
        
        Exit Sub
    
    End If

Next x

    MsgBox "YES"
    
End Sub

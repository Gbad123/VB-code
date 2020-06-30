VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "Go"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
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
        Dim a As Boolean


Private Sub cmdgo_Click()

x = InputBox("Number", "Number")
    y = 20

Do While a = False
    
    x = x + 1
    Print x
    
        If x = 16 Then a = True
    
Loop

End Sub

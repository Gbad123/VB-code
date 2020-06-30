VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Integer
        Dim readin As String
        Dim s(1 To 4) As String

Private Sub cmdgo_Click()

readin = txt1.Text
For x = 1 To Len(readin)
s(x) = Mid$(readin, x, 1)
Next x



End Sub

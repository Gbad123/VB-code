VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt2 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
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

Private Sub cmdgo_Click()

a = txt1.Text
b = txt2.Text

c = (b - a) + b
MsgBox c, vbOKOnly, "Eldest Child"
End

End Sub

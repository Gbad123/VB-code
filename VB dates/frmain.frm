VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt2 
      Height          =   855
      Left            =   6240
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txt1 
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   1095
      Left            =   3240
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lbl2 
      Caption         =   "Enter Month"
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lbl1 
      Caption         =   "Enter Date"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
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

Private Sub cmdgo_Click()

x = Val(txt1.Text)
y = Val(txt2.Text)

If x > 31 Then
    MsgBox "ERROR", vbCritical, "ERROR"
    End
End If

If y > 12 Then
    MsgBox "ERROR", vbCritical, "ERROR"
    End
End If

If y > 2 Then MsgBox "AFTER", vbOKOnly, "TEST"
If y < 2 Then MsgBox "BEFORE", vbOKOnly, "TEST"

If y = 2 And x > 18 Then MsgBox "AFTER", vbOKOnly, "TEST"
If y = 2 And x < 18 Then MsgBox "BEFORE", vbOKOnly, "TEST"

If y = 2 And x = 18 Then MsgBox "SPECICAL", vbOKOnly, "TEST"

End Sub

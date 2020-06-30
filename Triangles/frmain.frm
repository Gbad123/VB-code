VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   855
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txt2 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lbl1 
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   1695
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

Private Sub cmdclear_Click()

txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
lbl1.Caption = ""

End Sub

Private Sub cmdgo_Click()

a = Val(txt1.Text)
b = Val(txt2.Text)
c = Val(txt3.Text)

If a = 60 And b = 60 And a = 60 Then MsgBox "equilateral", vbOKOnly, "Triangle"
If a = b Or a = c Or c = b and Then MsgBox "isocelese", vbOKOnly, "triangle"
If a <> b And b <> c And a <> c Then MsgBox "scalene", vbOKOnly, "triangle"

If a + b + c > 180 Then MsgBox "Not a triangle", vbCritical, "ERROR"


End Sub

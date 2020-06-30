VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "Find fish"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txt4 
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt3 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt2 
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Height          =   1695
      Left            =   6120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
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
    Dim d As Integer
    
Private Sub cmd1_Click()

a = txt1.Text
b = txt2.Text
c = txt3.Text
d = txt4.Text

If a > b And b > c And c > d Then
lbl1.Caption = "fish assending"
Else
lbl1.Caption = "no fiah"

If a < b And b < c And c < d Then
lbl1.Caption = "fish decending"
Else
lbl1.Caption = "no fish"

End If
End If

End Sub

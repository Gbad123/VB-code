VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Question 2"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txt2 
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   2040
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   1320
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Weekend"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Evening"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Daytime min"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim w As Double
    Dim b As Double
    Dim a As Double
    Dim p As Double
    Dim q As Double

Private Sub cmdgo_Click()

x = (Val(txt1.Text) - 100) * 0.25
    If Val(txt1.Text) < 100 Then
        x = 0
    End If
    
y = Val(txt2.Text) * 0.15
z = Val(txt3.Text) * 0.2

b = (Val(txt1.Text) - 250) * 0.45

    If Val(txt1.Text) < 250 Then
        b = 0
    End If
    
w = (Val(txt2.Text) * 0.35)
a = (Val(txt3.Text) * 0.25)

p = "$" & (x + y + z)
    lbl5.Caption = "$" & p
    
q = "$" & (a + b + w)
    lbl6.Caption = "$" & q

If p < q Then
    lbl7.Caption = "Plan A is cheaper"
End If

If p > q Then
    lbl7.Caption = "Plan B is cheaper"
End If

If p = q Then
    lbl7.Caption = "The plans are the same price"
End If

End Sub

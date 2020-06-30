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
   Begin VB.CommandButton cmdend 
      Caption         =   "&End"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txt2 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lbldivide 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.Label cmdtimes 
      Alignment       =   2  'Center
      Caption         =   "*"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblminus 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblplus 
      Alignment       =   2  'Center
      Caption         =   "+"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblequal 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblanswer 
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n1 As Double
Dim n2 As Double
Dim a As Double

Private Sub cmdend_Click()

End

End Sub

Private Sub cmdtimes_Click()

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 * n2
lblanswer.Caption = a

End Sub

Private Sub lbldivide_Click()

On Error GoTo error:

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 / n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox ("Error"), vbCritical, "Error Screen"

End Sub

Private Sub lblminus_Click()

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 - n2
lblanswer.Caption = a

End Sub

Private Sub lblplus_Click()

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 + n2
lblanswer.Caption = a

End Sub

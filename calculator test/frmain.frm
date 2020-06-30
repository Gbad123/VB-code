VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcolour 
      Caption         =   "C&olour"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "&End"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdtan 
      Caption         =   "Tan"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdcos 
      Caption         =   "Cos"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdsin 
      Caption         =   "Sin"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdcube 
      Caption         =   "*3"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdsquare 
      Caption         =   "*2"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdpower 
      Caption         =   "^"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmddivide 
      Caption         =   "/"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdtimes 
      Caption         =   "*"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "+"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txt2 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblanswer 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblequal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   855
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

Private Sub cmdadd_Click()

On Error GoTo error

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 + n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbCritical, "Error"

End Sub

Private Sub cmdclear_Click()

txt1.Text = ""
txt2.Text = ""
lblanswer.Caption = ""

End Sub

Private Sub cmdcolour_Click()

MsgBox ("Power wasting mode"), vbOKOnly, "Colour Change"

cmdcolour.BackColor = vbRed
cmdend.BackColor = vbRed
cmdclear.BackColor = vbRed

cmdsin.BackColor = vbBlue
cmdcos.BackColor = vbBlue
cmdtan.BackColor = vbBlue

cmdadd.BackColor = vbGreen
cmdtimes.BackColor = vbGreen
cmdminus.BackColor = vbGreen
cmddivide.BackColor = vbGreen

cmdpower.BackColor = vbYellow
cmdcube.BackColor = vbYellow
cmdsquare.BackColor = vbYellow

frmain.BackColor = vbBlack
lblanswer.ForeColor = vbWhite
lblequal.ForeColor = vbWhite

txt1.BackColor = vbBlack
txt2.BackColor = vbBlack

txt1.ForeColor = vbWhite
txt2.ForeColor = vbWhite

End Sub

Private Sub cmdcos_Click()

On Error GoTo error

n1 = Val(txt1.Text)
a = (n1 * 3.1416 / 180)
a = Cos(a)
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbOKOnly, "Error"

End Sub

Private Sub cmdcube_Click()

On Error GoTo error

n1 = Val(txt1.Text)
a = n1 * n1 * n1
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbYesNo, "Error"

End Sub

Private Sub cmddivide_Click()

On Error GoTo error

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 / n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbCritical, "Error"

End Sub

Private Sub cmdend_Click()

End

End Sub

Private Sub cmdminus_Click()

On Error GoTo error

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 - n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbCritical, "Error"

End Sub

Private Sub cmdpower_Click()

On Error GoTo error

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 ^ n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbYesNo, "Error"

End Sub

Private Sub cmdsin_Click()

On Error GoTo error

n1 = Val(txt1.Text)
a = (n1 * 3.1416 / 180)
a = Sin(a)
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbOKOnly, "Error"

End Sub

Private Sub cmdsquare_Click()

On Error GoTo error

n1 = Val(txt1.Text)
a = n1 * n1
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbYesNo, "Error"

End Sub

Private Sub cmdtan_Click()

On Error GoTo error

n1 = Val(txt1.Text)
a = (n1 * 3.1416 / 180)
a = Tan(a)
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbOKOnly, "Error"

End Sub

Private Sub cmdtimes_Click()

On Error GoTo error

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 * n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbCritical, "Error"

End Sub

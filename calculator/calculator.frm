VERSION 5.00
Begin VB.Form frmain 
   BackColor       =   &H0000FF00&
   Caption         =   "Main"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H000000FF&
      Caption         =   "&Clear"
      Height          =   255
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdtan 
      BackColor       =   &H00FFFF00&
      Caption         =   "tan"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton camcos 
      BackColor       =   &H00FFFF00&
      Caption         =   "cos"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdpower 
      BackColor       =   &H00FFFF00&
      Caption         =   "^"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdsin 
      BackColor       =   &H00FFFF00&
      Caption         =   "sin"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdcube 
      BackColor       =   &H00FFFF00&
      Caption         =   "*3"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdsquared 
      BackColor       =   &H00FFFF00&
      Caption         =   "*2"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmddivide 
      BackColor       =   &H00FFFF00&
      Caption         =   "/"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdtimes 
      BackColor       =   &H00FFFF00&
      Caption         =   "*"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdminus 
      BackColor       =   &H00FFFF00&
      Caption         =   "-"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFFF00&
      Caption         =   "+"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdend 
      BackColor       =   &H000000FF&
      Caption         =   "&End"
      Height          =   255
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txt2 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblanswer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1575
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

Private Sub camcos_Click()


On Error GoTo error

n1 = Val(txt1.Text)
a = (n1 * 3.1416 / 180)
a = Cos(a)
lblanswer.Caption = a

Exit Sub
error:
MsgBox (""), vbOKOnly, "Error"

End Sub

Private Sub cmdadd_Click()

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 + n2
lblanswer.Caption = a

End Sub

Private Sub cmdclear_Click()

txt1.Text = ""
txt2.Text = ""
lblanswer.Caption = ""

End Sub

Private Sub cmdcube_Click()

n1 = Val(txt1.Text)
a = n1 * n1 * n1
lblanswer.Caption = a

End Sub

Private Sub cmddivide_Click()

On Error GoTo error:

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 / n2
lblanswer.Caption = a

Exit Sub
error:
MsgBox ("Can't divide by 0"), vbCritical, "Error"

End Sub

Private Sub cmdend_Click()

End

End Sub

Private Sub cmdminus_Click()

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 - n2
lblanswer.Caption = a

End Sub

Private Sub cmdpower_Click()

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 ^ n2
lblanswer.Caption = a

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

Private Sub cmdsquared_Click()

n1 = Val(txt1.Text)
a = n1 * n1
lblanswer.Caption = a

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

n1 = Val(txt1.Text)
n2 = Val(txt2.Text)
a = n1 * n2
lblanswer.Caption = a

End Sub

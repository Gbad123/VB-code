VERSION 5.00
Begin VB.Form frmain 
   BackColor       =   &H0000FF00&
   Caption         =   "Main"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   FillColor       =   &H000080FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcolour 
      BackColor       =   &H00FF0000&
      Caption         =   "&Colour"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmddisappear 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go &away"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H00FFFF00&
      Caption         =   "S&how"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdend 
      BackColor       =   &H000000FF&
      Caption         =   "&End"
      Height          =   375
      Left            =   3840
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0C000&
      Caption         =   "Ebad Babar"
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcolour_Click()

lbl1.BackStyle = transparent
lbl1.ForeColor = vbBlue

End Sub

Private Sub cmddisappear_Click()

lbl1.Visible = False

End Sub

Private Sub cmdend_Click()

End

End Sub

Private Sub cmdshow_Click()

lbl1.Visible = True

End Sub


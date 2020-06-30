VERSION 5.00
Begin VB.Form framain 
   Caption         =   "Main"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Height          =   1095
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblanswer3 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   7320
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblanswer2 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   6240
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblanswer 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   6120
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of eyes"
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Antenna"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "framain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim a As Integer
    Dim b As Integer

Private Sub cmdclear_Click()

txt1.Text = ""
txt2.Text = ""
lblanswer.Caption = ""

End Sub

Private Sub cmdgo_Click()

a = txt1.Text
b = txt2.Text

If a > 2 And b < 5 Then lblanswer.Caption = "TroyMartian"
If a < 7 And b > 1 Then lblanswer2.Caption = "VladSaturnian"
If a < 3 And b < 4 Then lblanswer3.Caption = "GraemeMercurian"

End Sub

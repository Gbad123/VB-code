VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "clear"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      Height          =   735
      Left            =   4560
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txt2 
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Mama Bears Bowl Weight"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight of bowl 3"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight of bowl 2"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblanswer 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight of bowl 1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
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
Dim z As Integer

Private Sub cmd1_Click()

x = txt1.Text
y = txt2.Text
z = txt3.Text

If x > y And y > z Then lblanswer.Caption = y & " " & "bowl # 2"
If z > y And y > x Then lblanswer.Caption = y & " " & "bowl # 2"

If y > x And x > z Then lblanswer.Caption = x & " " & "bowl # 1"
If z > x And x > y Then lblanswer.Caption = x & " " & "bowl # 1"

If z > y And x > z Then lblanswer.Caption = z & " " & "bowl # 3"
If z > x And y > z Then lblanswer.Caption = z & " " & "bowl # 3"

End Sub

Private Sub cmdclear_Click()

txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
lblanswer.Caption = ""

End Sub

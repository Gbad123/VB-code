VERSION 5.00
Begin VB.Form frmain2 
   Caption         =   "part 2"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt2 
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Guess"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txt1 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Hide your sentence"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim str As String
            Dim x As String

Private Sub cmd1_Click()

str = txt1.Text

str = StrReverse(str)
    str = Replace$(str, "e", "i")
        str = Replace$(str, "a", "u")
            str = Replace$(str, "o", "a")
            
    Print str
    
End Sub

Private Sub cmd2_Click()

x = txt2.Text

If x = str Then
    MsgBox "Your are Right", vbOKOnly, "Secret"
        MsgBox txt1.Text, vbOKOnly, "Secret"
        
ElseIf x > str Then
    MsgBox "Your are Wrong", vbCritical, "Secret"
    
End If
    
End Sub

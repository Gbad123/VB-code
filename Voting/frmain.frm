VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   1095
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim str As String
    Dim a As Integer
    Dim b As Integer

Private Sub cmdgo_Click()

str = txt1.Text

a = UBound(Split(str, "A"))
b = UBound(Split(str, "B"))

If a > b Then MsgBox "A WINS", vbOKOnly, "WINNER"
If b > a Then MsgBox "B WINS", vbOKOnly, "WINNER"
If a = b Then MsgBox "TIE", vbCritical, "NO WINNER"

End Sub

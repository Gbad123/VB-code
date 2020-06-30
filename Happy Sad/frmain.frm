VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim str As String
        Dim x As Integer
        Dim y As Integer
    
Private Sub cmdgo_Click()

str = txt1.Text

Print UBound(Split(str, ":)"))
x = UBound(Split(str, ":)"))

Print x

If x > 0 Then MsgBox "i'm happy", vbOKCancel, "HAPPY"


Print UBound(Split(str, ":("))
y = UBound(Split(str, ":("))

Print y

If y > 0 Then MsgBox "i'm sad", vbOKCancel, "SAD"

If x = 0 And y = 0 Then MsgBox "NONE", vbOKOnly, "NONE"

End Sub

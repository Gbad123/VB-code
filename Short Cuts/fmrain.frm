VERSION 5.00
Begin VB.Form fmrain 
   Caption         =   "Main"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd2 
      Caption         =   "Clear"
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "GO"
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lbl3 
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Long term"
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter short form"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "fmrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()

If txt1.Text = "CU" Then lbl3.Caption = "see you"
If txt1.Text = ":-)" Then lbl3.Caption = "I’m happy"
If txt1.Text = ":-(" Then lbl3.Caption = " I’m unhappy"
If txt1.Text = ";-)" Then lbl3.Caption = " wink"
If txt1.Text = ":-P" Then lbl3.Caption = " stick out my tongue"
If txt1.Text = "(˜." Then lbl3.Caption = "˜) sleepy"
If txt1.Text = "TA" Then lbl3.Caption = " totally awesome"
If txt1.Text = "CCC" Then lbl3.Caption = " Canadian Computing Competition"
If txt1.Text = "CUZ" Then lbl3.Caption = " because"
If txt1.Text = "TY" Then lbl3.Caption = " thank - you"
If txt1.Text = "YW" Then lbl3.Caption = " you’re welcome"
If txt1.Text = "TTYL" Then lbl3.Caption = " talk to you later"

End Sub

Private Sub cmd2_Click()

txt1.Text = ""
lbl3.Caption = ""

End Sub

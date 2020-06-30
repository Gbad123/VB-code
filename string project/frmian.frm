VERSION 5.00
Begin VB.Form frmian 
   Caption         =   "Main"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt3 
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Part2"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txt2 
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txt1 
      Height          =   855
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdboom 
      Caption         =   "Lets Do it"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "What letter you want"
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Caption         =   "Replace A Letter"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim str As String
    Dim stright As String
        Dim stmid As String
            
Private Sub cmd3_Click()

Load frmain2
    frmain2.Show
        Unload frmian
        
End Sub

Private Sub cmdboom_Click()

str = InputBox("Enter A Sentence", "Name", "Ebad")
    Print str
    
End Sub

Private Sub cmdnext_Click()

Print UCase(str) 'everything uppercase
    Print LCase(str) 'everything lowercase

Print Replace$(str, "e", "x") 'all e's are now x'es

stright = Right$(str, 5)
    Print stright '5 characters from right
    
stmid = Mid$(str, 3, 8)
    Print stmid 'start at 3 and up to 8

Print Replace$(str, txt1.Text, txt2.Text) 'changes what you put in txt1 to what you put in txt2

Print StrReverse$(str) 'you sentence backwords

Print UBound(Split(str, txt3.Text))

End Sub

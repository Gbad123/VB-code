VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdretreave 
      Caption         =   "Retreave"
      Height          =   975
      Left            =   4080
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txt2 
      Height          =   975
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txt1 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim afile As Integer
    
Private Sub cmdretreave_Click()

afile = FreeFile
    Open App.Path & "\tester.csv" For Input As #afile
        Dim whatever As String
            Do While Not EOF(afile)
                Input #afile, whatever
                    txt2.Text = txt2.Text & whatever & vbCrLf
                        Loop
                            Close #afile

End Sub

Private Sub cmdsave_Click()

afile = FreeFile
    Open App.Path & "\tester.csv" For Output As #afile
        Print #afile, txt1.Text
            Close #afile

End Sub

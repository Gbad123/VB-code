VERSION 5.00
Begin VB.Form Frmain 
   Caption         =   "Main"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Command"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim x As Integer
        Dim n As String
            Dim v As String

Private Sub cmd1_Click()

x = Val(txt1.Text)
    n = txt2.Text
    
If x > 0 And x < 50 Then
    v = n & " " & "fail"
        lbl1.Caption = n & " " & "fail"
            ElseIf x >= 50 And x < 70 Then
                v = n & " " & "ok"
                    lbl1.Caption = n & " " & "ok"
                        ElseIf x >= 70 And x < 90 Then
                            v = n & " " & "great"
                                lbl1.Caption = n & " " & "great"
                                    Else:
                                        v = n & " " & "Excellent"
                                            lbl1.Caption = n & " " & "Excellent"

Call save(v)

End If

End Sub

Sub save()

afile = FreeFile
    Open App.Path & "\tester.csv" For Output As #afile
        Print #afile, txt1.Text
            Close #afile

End Sub

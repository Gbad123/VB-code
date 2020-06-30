VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdretreave 
      Caption         =   "Retreave"
      Height          =   615
      Left            =   5640
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdave 
      Caption         =   "Average"
      Height          =   735
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txt4 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txt3 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txt2 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblcomeback 
      BackStyle       =   0  'Transparent
      Height          =   2055
      Left            =   5640
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbl8 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblname 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblscore 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   3360
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Computers"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Science"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Math"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "English"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dimiing
Dim s As Integer 'science
    Dim m As Integer 'math
        Dim e As Integer 'english
            Dim c As Integer 'computers
                Dim a As Integer 'total
                    Dim afile As Integer 'save name
            
Private Sub cmdave_Click()

'what each stands for
e = Val(txt1.Text)
    c = Val(txt2.Text)
        m = Val(txt3.Text)
            s = Val(txt4.Text)

'final
a = (e + c + m + s) / 4
    lblscore.Caption = a & " " & txtname.Text
    
'how to do each
If e > 0 And e < 50 Then
    lbl5.Caption = "Fail!"
        ElseIf e >= 50 And e < 70 Then
            lbl5.Caption = "OK!"
                ElseIf e >= 70 And e < 90 Then
                    lbl5.Caption = "Great!"
                        Else:
                            lbl5.Caption = "Excellent!"
    
    End If
    
If c > 0 And c < 50 Then
    lbl6.Caption = "Fail!"
        ElseIf c >= 50 And c < 70 Then
            lbl6.Caption = "OK!"
                ElseIf c >= 70 And c < 90 Then
                    lbl6.Caption = "Great!"
                        Else:
                            lbl6.Caption = "Excellent!"
    
    End If
    
If m > 0 And m < 50 Then
    lbl7.Caption = "Fail!"
        ElseIf m >= 50 And m < 70 Then
            lbl7.Caption = "OK!"
                ElseIf m >= 70 And m < 90 Then
                    lbl7.Caption = "Great!"
                        Else:
                            lbl7.Caption = "Excellent!"
    
    End If
    
If s > 0 And s < 50 Then
    lbl8.Caption = "Fail!"
        ElseIf s >= 50 And s < 70 Then
            lbl8.Caption = "OK!"
                ElseIf s >= 70 And s < 90 Then
                    lbl8.Caption = "Great!"
                        Else:
                            lbl8.Caption = "Excellent!"
    
    End If
                            
End Sub

Private Sub cmdretreave_Click()

'retreave
afile = FreeFile
    Open App.Path & "\tester.text" For Input As #afile
        Dim whatever As String
        
            Do While Not EOF(afile)
                Input #afile, whatever
                    lblcomeback.Caption = lblcomeback.Caption & whatever & vbCrLf
                        Loop
                        
    Close #afile
                            
End Sub

Private Sub cmdsave_Click()

'save
afile = FreeFile
    Open App.Path & "\tester.text" For Output As #afile
    
        Print #afile, txt1.Text & " " & "English"
            Print #afile, txt2.Text & " " & "Computers"
                Print #afile, txt3.Text & " " & "Math"
                    Print #afile, txt4.Text & " " & "Science"
                            Print #afile, "Total Average" & " " & lblscore.Caption
                    
    Close #afile
            
End Sub


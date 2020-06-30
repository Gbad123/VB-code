VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
    Dim stleft As String
        Dim stright As String
            Dim stmid As String
                Dim leng As Integer

Private Sub cmd1_Click()

str = "Wow I really hate working with string functions"
    Print str 'this prints to the screen

stleft = Left$(str, 7)
    Print stleft 'this prints first 7 characters from the left

stright = Right$(str, 7)
    Print stright 'this prints last 7 characters of the sentence

stmid = Mid$(str, 6, 7)
    Print stmid 'this prints 7 characters starting from 6
    
leng = Len(str)
    Print leng 'prints the number of total characters
    
Print UCase(str) 'makes upper case everything

Print LCase(str) 'makes lower case everything

Print StrReverse$(str) 'makes everything backwords

Print Replace$(str, "t", "*") 'replaces all the t's with *'s

Print InStr(str, "r") 'where is the 1st r in the senetence

Print InStr(10, str, "r") 'starts at character 10 then looks for the first r

Print InStrRev(str, "r") 'looks for the r from the right side

Print UBound(Split(str, "t")) 'counts the number of t's in the whole sentence

End Sub

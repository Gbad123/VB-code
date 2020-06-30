VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtree 
      Caption         =   "Tree"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim d As String
        Dim b As Integer
        Dim x As Integer
        Dim y As Integer
    Dim space As String
        Dim p As Integer
    Dim q As String
    Dim l As String

Private Sub cmdtree_Click()

b = InputBox("How big the tree")

For x = 1 To b
    space = "     "
    
        For y = b To x Step -1
            space = space & " "
        
                For p = 1 To x Step -1
                    q = space & "[ ]"
                        l = q
                    
                Next p
            
        Next y
            

d = d & "%"

Print vbCrLf & space & d & vbCrLf & l

Next x

End Sub

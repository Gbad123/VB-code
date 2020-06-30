VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Mian"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton optnormal 
      Caption         =   "Normal"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton optmax 
      Caption         =   "Max"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Blue"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Red"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub opt1_Click()

If opt1.Value = True Then
    frmain.BackColor = vbRed

End If

End Sub

Private Sub opt2_Click()

If opt2.Value = True Then
    frmain.BackColor = vbBlue

End If

End Sub

Private Sub optmax_Click()

If optmax.Value = True Then
    frmain.WindowState = 2

End If

End Sub

Private Sub optnormal_Click()

If optnormal.Value = True Then
    frmain.WindowState = 0

End If

End Sub

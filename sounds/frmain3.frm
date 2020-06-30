VERSION 5.00
Begin VB.Form frmain3 
   Caption         =   "Tertiary"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "Gun sound"
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmain3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()

Load frmain
frmain.Show
Unload frmain3

End Sub

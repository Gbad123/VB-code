VERSION 5.00
Begin VB.Form frmain2 
   Caption         =   "Secondary"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Boing Sound"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()

Load frmain
frmain.Show
Unload frmain2

End Sub

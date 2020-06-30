VERSION 5.00
Begin VB.Form frmain2 
   Caption         =   "Scondary"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "&END"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "You hit the block"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
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

Private Sub cmdend_Click()

End

End Sub

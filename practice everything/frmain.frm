VERSION 5.00
Begin VB.Form frmain 
   BackColor       =   &H0000FFFF&
   Caption         =   "Title"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7260
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Avoid Collision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl2_Click()

Load frmain2
    frmain2.Show
        Unload frmain

Play App.Path & "\boing_2.wav"

End Sub

Private Sub lbl3_Click()

Load frmain3
    frmain3.Show
        Unload frmain
    
Play App.Path & "\boing_2.wav"

End Sub

Private Sub lbl4_Click()

End

End Sub

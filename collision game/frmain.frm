VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Start"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   Picture         =   "frmain.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Practice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lbl2 
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
      ForeColor       =   &H00FFFF80&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Psycho Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1455
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl2_Click()

'it's game time
Load frmgame
    frmgame.Show
        Unload frmain
        
End Sub

Private Sub lbl3_Click()

'learn what to do
Load frmcontrol
    frmcontrol.Show
        Unload frmain
        
End Sub

Private Sub lbl4_Click()

'why you play
Load frmpractice
    frmpractice.Show
        Unload frmain
        
End Sub

Private Sub lblend_Click()

'end
End

End Sub

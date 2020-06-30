VERSION 5.00
Begin VB.Form frmcontrol 
   Caption         =   "Control"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   Picture         =   "frmcontrol.frx":0000
   ScaleHeight     =   4545
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move using the arrow keys and use the space bar to shoot energy balls and take down your target."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   3480
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmcontrol.frx":9DC6
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
      Height          =   1695
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Practice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmcontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl1_Click()

'learn the controls
Load frmain
    frmain.Show
        Unload frmcontrol
        
End Sub

Private Sub lbl2_Click()

'practice it
Load frmpractice
    frmpractice.Show
        Unload frmcontrol
        
End Sub

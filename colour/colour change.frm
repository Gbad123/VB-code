VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   FillColor       =   &H0080FFFF&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "&Go"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtname 
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdcolour 
      Caption         =   "&Colour"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmddissapear 
      Caption         =   "&Dissapear"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "&Show"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "&End"
      Height          =   255
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblname 
      Caption         =   "Ebad Babar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcolour_Click()

'this code changes colour

frmain.BackColor = vbYellow
frmain.ForeColor = vbGreen

cmdend.BackColor = vbRed

cmdshow.BackColor = vbBlue

cmddissapear.BackColor = vbRed

cmdcolour.BackColor = vbGreen

lblname.BackColor = vbYellow
lblname.Font = vbGreen

txtname.BackColor = vbBlue
txtname.ForeColor = vbYellow

cmdgo.BackColor = vbGreen

End Sub

Private Sub cmddissapear_Click()

'makes label dissapear

lblname.Visible = False

End Sub

Private Sub cmdend_Click()

'end

End

End Sub


Private Sub cmdgo_Click()

'puts text in label

lblname.Caption = txtname.Text

End Sub

Private Sub cmdshow_Click()

'makes lable appear

lblname.Visible = True

End Sub


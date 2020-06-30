VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Player 2"
      Height          =   2775
      Left            =   4560
      TabIndex        =   16
      Top             =   1920
      Width           =   4215
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "1"
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "2"
         Height          =   375
         Left            =   1320
         TabIndex        =   27
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "3"
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "4"
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "5"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "6"
         Height          =   495
         Left            =   3240
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.Frame frm1 
      Caption         =   "Player 1"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
      Begin VB.TextBox txtsix 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtfive 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtfour 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtthree 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txttwo 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtone 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "6"
         Height          =   495
         Left            =   3240
         TabIndex        =   15
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "5"
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "2"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Caption         =   "Number of Rolls"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer

Dim y As Integer
Dim z As Integer

Dim q As Integer
Dim r As Integer

    'player one
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim e As Integer
    Dim f As Integer
    
    'player 2
    Dim g As Integer
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer

Private Sub cmdgo_Click()

x = txt1.Text

a = txtone.Text
b = txttwo.Text
c = txtthree.Text
d = txtfour.Text
e = txtfive.Text
f = txtsix.Text

y = (a * 1) + (b * 2) + (c * 3) + (d * 4) + (e * 5) + (f * 6)
z = a + b + c + d + e + f

'If z > x Then
'MsgBox "cheater player 1", vbCritical, "Cheater"
'End
'End If

g = Text6.Text
h = Text5.Text
i = Text4.Text
j = Text3.Text
k = Text2.Text
l = Text1.Text

q = (q * 1) + (h * 2) + (i * 3) + (j * 4) + (k * 5) + (l * 6)
r = q + h + i + j + k + l

'If r > x Then
'MsgBox "cheater player 2", vbCritical, "Cheater"
'End
'End If

If y > q Then MsgBox "Player one wins", vbOKOnly, "Winner"
If y < q Then MsgBox "Player two wins", vbOKOnly, "Winner"
If y = q Then MsgBox "NO ONE WINS", vbCritical, "TIE"

End Sub


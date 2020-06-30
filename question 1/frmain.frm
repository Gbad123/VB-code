VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Question 1"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm4 
      Caption         =   "Dessert"
      Height          =   1935
      Left            =   5400
      TabIndex        =   20
      Top             =   1320
      Width           =   1095
      Begin VB.OptionButton opt16 
         Caption         =   "4"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton opt15 
         Caption         =   "3"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton opt14 
         Caption         =   "2"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton opt13 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame frm3 
      Caption         =   "Drink"
      Height          =   1935
      Left            =   3960
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
      Begin VB.OptionButton opt12 
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton opt11 
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton opt10 
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton opt9 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frm2 
      Caption         =   "Side dish"
      Height          =   1935
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
      Begin VB.OptionButton opt8 
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton opt7 
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton opt6 
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton opt5 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "Go"
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame frm1 
      Caption         =   "Burger"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
      Begin VB.OptionButton opt4 
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton opt3 
         Caption         =   "3"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton opt2 
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton opt1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   2280
      TabIndex        =   26
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Colories"
      Height          =   495
      Left            =   2280
      TabIndex        =   25
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter the Dessert"
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter the Drink"
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter the Side Dish"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter the Burger"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
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
    Dim w As Integer

Private Sub cmdgo_Click()

If opt1.Value = True Then
    x = 461
End If

If opt2.Value = True Then
    x = 431
End If

If opt3.Value = True Then
    x = 420
End If

If opt4.Value = True Then
    x = 0
End If

If opt5.Value = True Then
    y = 100
End If

If opt6.Value = True Then
    y = 57
End If

If opt7.Value = True Then
    y = 70
End If

If opt8.Value = True Then
    y = 0
End If

If opt9.Value = True Then
    z = 130
End If

If opt10.Value = True Then
    z = 160
End If

If opt11.Value = True Then
    z = 118
End If

If opt12.Value = True Then
    z = 0
End If

If opt13.Value = True Then
    w = 167
End If

If opt14.Value = True Then
    w = 266
End If

If opt15.Value = True Then
    w = 75
End If

If opt16.Value = True Then
    w = 0
End If

lbl6.Caption = x + y + z + w

End Sub

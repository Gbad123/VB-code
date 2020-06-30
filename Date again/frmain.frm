VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmb1 
      Height          =   315
      ItemData        =   "frmain.frx":0000
      Left            =   240
      List            =   "frmain.frx":0061
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ListBox lst1 
      Height          =   840
      ItemData        =   "frmain.frx":00E1
      Left            =   2760
      List            =   "frmain.frx":0109
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txt3 
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdays 
      Caption         =   "Go"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblyear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblmonth 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblday 
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3840
      Width           =   615
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim day As Integer
Dim month As Integer
Dim year As Integer

Dim x As String


Private Sub cmdays_Click()

day = Val(cmb1.Text)
month = Val(lst1.Text)
year = Val(txt3.Text)
x = month & "/" & day & "/" & year

Print "Now:"; Tab(20); Now
Print DateDiff("d", x, Now)

Print "using 'yyyy':"; Tab(20); DateDiff("yyyy", #1/1/2001#, Now)
Print "Using 'q':"; Tab(20); DateDiff("q", #1/1/2001#, Now)
Print "Using 'm':"; Tab(20); DateDiff("m", #1/1/2001#, Now)
Print "Using 'y':"; Tab(20); DateDiff("y", #1/1/2001#, Now)
Print "Using 'd':"; Tab(20); DateDiff("d", #1/1/2001#, Now)
Print "Using 'w':"; Tab(20); DateDiff("w", #1/1/2001#, Now)
Print "Using 'ww':"; Tab(20); DateDiff("ww", #1/1/2001#, Now)

End Sub

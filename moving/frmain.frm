VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
   Begin VB.Shape shp1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   495
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    Case vbKeyRight
        shp1.Left = shp1.Left + 250
    Case vbKeyLeft
        shp1.Left = shp1.Left - 250
    Case vbKeyDown
        shp1.Top = shp1.Top + 250
    Case vbKeyUp
        shp1.Top = shp1.Top - 250
        
End Select
End Sub

Private Sub lblend_Click()

End

End Sub

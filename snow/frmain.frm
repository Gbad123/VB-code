VERSION 5.00
Begin VB.Form frmain 
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrmove 
      Interval        =   50
      Left            =   2640
      Top             =   5280
   End
   Begin VB.Timer tmrspawn 
      Interval        =   50
      Left            =   960
      Top             =   1200
   End
   Begin VB.Timer tmrsnow 
      Interval        =   1
      Left            =   360
      Top             =   1200
   End
   Begin VB.Timer tmr1 
      Interval        =   100
      Left            =   8640
      Top             =   5880
   End
   Begin VB.Image img3 
      Height          =   2100
      Left            =   600
      Picture         =   "frmain.frx":0000
      Top             =   3720
      Width           =   1620
   End
   Begin VB.Image img2 
      Height          =   960
      Left            =   8880
      Picture         =   "frmain.frx":02FF
      Top             =   6720
      Width           =   690
   End
   Begin VB.Image img1 
      Height          =   585
      Index           =   0
      Left            =   960
      Picture         =   "frmain.frx":0604
      Top             =   360
      Width           =   510
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'shaw's code
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hwnd As Long, _
                    ByVal nIndex As Long) As Long
     
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hwnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) As Long
                   
    Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal crKey As Long, _
                    ByVal bAlpha As Byte, _
                    ByVal dwFlags As Long) As Long
     
    Private Const GWL_STYLE = (-16)
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_LAYERED = &H80000
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2
     
    Dim mr As Integer
    
'dims
Option Explicit
    Dim l As Integer
        Dim coun As Integer
            Dim x As Integer
                Dim s As Integer

Private Sub Form_Load()

'shaw's code
Me.BackColor = vbCyan
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
        
End Sub

Private Sub img2_Click()

'end
End

End Sub

Private Sub img3_Click()
End Sub

Private Sub tmr1_Timer()

'moving
l = l + 1
    img2.Picture = LoadPicture(App.Path & ("\" & l & "l.gif"))
        If l = 3 Then l = -1
        
img2.Left = img2.Left - 200
    If img2.Left < 0 Then
        img2.Left = frmain.Width - img2.Left
        
    End If

        
End Sub

Private Sub tmrmove_Timer()

'moving on it's own
s = s + 1
    img3.Picture = LoadPicture(App.Path & ("\" & s & "s.gif"))
        If s = 2 Then s = -1
        
End Sub

Private Sub tmrsnow_Timer()

'snowing
For x = 0 To coun
    img1(x).Top = img1(x).Top + 30
        If img1(x).Top > frmain.Height Then
            img1(x).Top = 0
        End If
    
Next x

End Sub

Private Sub tmrspawn_Timer()

'spawninig
Randomize

    coun = coun + 1
    
    Load img1(coun)
        img1(coun).Visible = True
            img1(coun).Top = 0
                img1(coun).Left = Rnd(10) * frmain.Width
        
End Sub

VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   2640
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdalarm 
      Caption         =   "Stop"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer tmralarm 
      Interval        =   1
      Left            =   5520
      Top             =   1920
   End
   Begin VB.Timer tmr1 
      Interval        =   1
      Left            =   4320
      Top             =   1920
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "eg: 9:30:14 AM"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line linesecond 
      X1              =   1320
      X2              =   1320
      Y1              =   2160
      Y2              =   3360
   End
   Begin VB.Line linehour 
      X1              =   1320
      X2              =   1320
      Y1              =   2160
      Y2              =   1440
   End
   Begin VB.Line lineminute 
      X1              =   1320
      X2              =   2520
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
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
    
Const PI = 3.14159
Dim length As Integer
Dim Hours As Single, Minutes As Single, Seconds As Single
Dim TrueHours As Single

    Dim x As Integer

Private Sub cmdalarm_Click()

tmralarm.Enabled = False
tmr2.Enabled = False

End Sub

Private Sub cmdend_Click()

End

End Sub

Private Sub Form_Load()

'shaw's code
Me.BackColor = vbCyan
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY

End Sub

Private Sub tmr1_Timer()

lbl1.Caption = Date

lbl2.Caption = Time

Hours = Hour(Time)
Minutes = Minute(Time)
Seconds = Second(Time)
TrueHours = Hours + Minutes / 60

' I made all the X1 and Y1 equal in the form
linehour.X2 = 700 * Cos(PI / 180 * (30 * TrueHours - 90)) + linehour.X1
linehour.Y2 = 700 * Sin(PI / 180 * (30 * TrueHours - 90)) + linehour.Y1
    
lineminute.X2 = 950 * Cos(PI / 180 * (6 * Minutes - 90)) + linehour.X1
lineminute.Y2 = 950 * Sin(PI / 180 * (6 * Minutes - 90)) + linehour.Y1

linesecond.X2 = 1600 * Cos(PI / 180 * (6 * Seconds - 90)) + linehour.X1
linesecond.Y2 = 1600 * Sin(PI / 180 * (6 * Seconds - 90)) + linehour.Y1
    
End Sub

Private Sub tmr2_Timer()

Play App.Path & "\boing_2.wav"

End Sub

Private Sub tmralarm_Timer()

If lbl2.Caption = txt1.Text Then
    
    tmr2.Enabled = True
    
End If

End Sub

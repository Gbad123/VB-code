VERSION 5.00
Begin VB.Form frmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   Caption         =   "Travel Package"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmeal 
      BackColor       =   &H0000FF00&
      Caption         =   "Meal"
      Height          =   2175
      Left            =   5400
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
      Begin VB.OptionButton opt3 
         BackColor       =   &H0000FF00&
         Caption         =   "Some times"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H0000FF00&
         Caption         =   "No"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H0000FF00&
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.ComboBox cmb2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      ItemData        =   "frmain.frx":0000
      Left            =   6000
      List            =   "frmain.frx":0010
      TabIndex        =   4
      Text            =   "Time"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cmb1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmain.frx":0037
      Left            =   3000
      List            =   "frmain.frx":0047
      TabIndex        =   3
      Text            =   "Hotel"
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H0000FF00&
      Caption         =   "Go"
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox lst1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1560
      ItemData        =   "frmain.frx":0073
      Left            =   120
      List            =   "frmain.frx":0086
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   855
      Left            =   8040
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image pic1 
      Height          =   2295
      Left            =   120
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destinations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim x As Integer 'places
        Dim y As Integer 'hotel
            Dim z As Integer 'time
                Dim p As Integer 'meals
                    Dim t As Integer 'total

Private Sub cmdgo_Click()

'South America code
If lst1.Text = "South America" Then
    pic1.Picture = LoadPicture(App.Path & "\south america.bmp")
        pic1.Visible = True
            x = 5800
            
End If

'EU code
If lst1.Text = "EU" Then
    pic1.Picture = LoadPicture(App.Path & "\EU.bmp")
        pic1.Visible = True
            x = 5900


End If

'Middle east code
If lst1.Text = "Middle East" Then
    pic1.Picture = LoadPicture(App.Path & "\middle east.bmp")
        pic1.Visible = True
            x = 5000

End If

'East Asis code
If lst1.Text = "East Asia" Then
    pic1.Picture = LoadPicture(App.Path & "\east asia.bmp")
        pic1.Visible = True
            x = 6000
            
End If

'Holy Lands code
If lst1.Text = "Holy Lands" Then
    pic1.Picture = LoadPicture(App.Path & "\holy lands.bmp")
        pic1.Visible = True
            x = 4500

End If

'private code
If cmb1.Text = "Private" Then
    y = 3500

End If

'luxury code
If cmb1.Text = "Luxury" Then
    y = 2200

End If

'Middle class code
If cmb1.Text = "Middle Class" Then
    y = 1500
    
End If

'Economy code
If cmb1.Text = "Economy" Then
    y = 1900

End If

'1 week
If cmb2.Text = "1 Week" Then
    z = 700

End If

'2 weeks
If cmb2.Text = "2 Weeks" Then
    z = 900

End If

'3 weeks
If cmb2.Text = "3 Weeks" Then
    z = 1000
    
End If

'4 Weeks
If cmb2.Text = "4 Weeks" Then
    z = 2000

End If

'yes meals
If opt1.Value = True Then
    p = 500

End If

'sometimes
If opt3.Value = True Then
    p = 200

End If

'no meals
If opt2.Value = True Then
    p = 0
    
End If

t = p + x + y + z
    lbl2.Caption = "$" & t

End Sub

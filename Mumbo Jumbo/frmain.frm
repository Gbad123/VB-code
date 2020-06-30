VERSION 5.00
Begin VB.Form frmain 
   Caption         =   "Main"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txt1 
      Height          =   615
      Left            =   2760
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   2400
      Width           =   2775
   End
End
Attribute VB_Name = "frmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim str As String
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim h As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        Dim m As Integer
        Dim n As Integer
        Dim o As Integer
        Dim p As Integer
        Dim q As Integer
        Dim r As Integer
        Dim s As Integer
        Dim t As Integer
        Dim u As Integer
        Dim v As Integer
        Dim w As Integer
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
    
Private Sub cmdgo_Click()

str = txt1.Text

a = UBound(Split(str, "a"))
b = UBound(Split(str, "b"))
c = UBound(Split(str, "c"))
d = UBound(Split(str, "d"))
e = UBound(Split(str, "e"))
f = UBound(Split(str, "f"))
g = UBound(Split(str, "g"))
h = UBound(Split(str, "h"))
i = UBound(Split(str, "i"))
j = UBound(Split(str, "j"))
k = UBound(Split(str, "k"))
l = UBound(Split(str, "l"))
m = UBound(Split(str, "m"))
n = UBound(Split(str, "n"))
o = UBound(Split(str, "o"))
p = UBound(Split(str, "p"))
q = UBound(Split(str, "q"))
r = UBound(Split(str, "r"))
s = UBound(Split(str, "s"))
t = UBound(Split(str, "t"))
u = UBound(Split(str, "u"))
v = UBound(Split(str, "v"))
w = UBound(Split(str, "w"))
x = UBound(Split(str, "x"))
y = UBound(Split(str, "y"))
z = UBound(Split(str, "z"))

If a > 0 Then Print "a";
If b > 0 Then Print "bac";
If c > 0 Then Print "ced";
If d > 0 Then Print "dec"
If e > 0 Then Print "e";
If f > 0 Then Print "feg";
If g > 0 Then Print "geh";
If h > 0 Then Print "hig";
If i > 0 Then Print "i";
If j > 0 Then Print "jik";
If k > 0 Then Print "kil";
If l > 0 Then Print "lim";
If m > 0 Then Print "mon";
If n > 0 Then Print "nom";
If o > 0 Then Print "o";
If p > 0 Then Print "poq";
If q > 0 Then Print "qor";
If r > 0 Then Print "ros";
If s > 0 Then Print "sut";
If t > 0 Then Print "tus";
If u > 0 Then Print "u";
If v > 0 Then Print "vuw";
If w > 0 Then Print "wux";
If x > 0 Then Print "xuy";
If y > 0 Then Print "yuz";
If z > 0 Then Print "zu";

End Sub

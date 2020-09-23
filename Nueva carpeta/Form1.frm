VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alpha Channel     v. 1.0"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear  IMPORTANT!!"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   3
      Top             =   2280
      Value           =   50
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alpha Channel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   2325
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   1
      Top             =   120
      Width           =   2130
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   120
      Picture         =   "Form1.frx":417C
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   120
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "50 %"
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 0 To Picture1.Width Step Screen.TwipsPerPixelX

For ii = 0 To Picture1.Height Step Screen.TwipsPerPixelY
    Dim color As Long
    Dim color2 As Long
    
color = Picture1.Point(i, ii)
color2 = Picture3.Point(i, ii)

    Dim r As Integer, g As Integer, b As Integer
    Dim rback As Integer, gback As Integer, bback As Integer
GetRgb color, r, g, b
GetRgb color2, rback, gback, bback
percent = 100 / HScroll1.Value
Picture3.PSet (i, ii), RGB(rback - rback / percent + r / percent, gback - gback / percent + g / percent, bback - bback / percent + b / percent)
Next

Next
End Sub

Private Sub Command2_Click()
Picture3.Cls
End Sub

Private Sub HScroll1_Change()
Label1 = HScroll1.Value & " %"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim color As Long
    
color = Picture1.Point(X / 15, Y / 15)
End Sub

Sub GetRgb(ByVal color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    Dim temp As Long
    
    temp = (color And 255)
    red = temp And 255
    
    temp = Int(color / 256)
    green = temp And 255
    
    temp = Int(color / 65536)
    blue = temp And 255
       
End Sub

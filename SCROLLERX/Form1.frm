VERSION 5.00
Object = "{6DCB1BBF-F5A2-11D3-AAE2-0000CB5322C6}#1.0#0"; "ScrollerX.ocx"
Begin VB.Form Form1 
   Caption         =   "Scrolling Example of ScrollerX  (C) by Siegfried Rings"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture3 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   465
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   7005
   End
   Begin ScrollerX.SCROLLER SCROLLER1 
      Height          =   3255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5741
      Text            =   $"Form1.frx":13780
      Fontbackground  =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start/Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font-Background-Style"
      Height          =   1695
      Left            =   2640
      TabIndex        =   12
      Top             =   3480
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Forecolor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mode RED-Bar"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mode Fire"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mode Yellow/blue"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mode solid"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Backcolor"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3375
      Left            =   6960
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fontname && size"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scroll_Properties"
      Height          =   1695
      Left            =   5640
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
      Begin VB.CheckBox Check1 
         Caption         =   "Use Sinus"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Aktiviert
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   6
         Top             =   840
         Value           =   1
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1455
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   4
         Top             =   1200
         Value           =   4
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Scrollspeed"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   20
         TabIndex        =   2
         Top             =   240
         Value           =   2
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   3960
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Timer enabled"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Value           =   1  'Aktiviert
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bar_before As Boolean
Dim Bar_past As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
 SCROLLER1.Sinus = True
Else
 SCROLLER1.Sinus = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
 Timer1.Enabled = True
Else
 Timer1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
SCROLLER1.Action = Not SCROLLER1.Action
'If SCROLLER1.Action = True Then Timer1.Enabled = True

End Sub

Private Sub Command2_Click(index As Integer)
MsgBox "Put a CommonDialog in this Code;  I removed while problems with other machines"
'Select Case index
' Case 0
' CDLG.Color = SCROLLER1.ForeColor
' CDLG.Action = 3 'color
' SCROLLER1.ForeColor = CDLG.Color
' Case 1
' CDLG.Color = SCROLLER1.BackColor
' CDLG.Action = 3 'color
' SCROLLER1.BackColor = CDLG.Color
'End Select
'Command2(index).BackColor = CDLG.Color
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As Long

'init the backgroundpicture
SCROLLER1.SetBackgroundpicture Picture3.Image, 1, 1

'init the Bars
Picture1.Width = SCROLLER1.Width
Picture1.Height = 512
Picture2.Width = SCROLLER1.Width
Picture2.Height = 512

For i = 1 To 255
 c = RGB(i, i, i)
 Picture1.Line (1, i)-(Picture1.Width, i), c, BF
 c = RGB(i, i, 0)
 Picture2.Line (1, i)-(Picture2.Width, i), c, BF
Next i
For i = 256 To 512
 c = RGB(512 - i, 512 - i, 512 - i)
 Picture1.Line (1, i)-(Picture1.Width, i), c, BF
 c = RGB(512 - i, 512 - i, 0)
 Picture2.Line (1, i)-(Picture2.Width, i), c, BF
Next i

'init the vscroll and some defaults

VScroll1.Max = SCROLLER1.Height / Screen.TwipsPerPixelY
VScroll1.Value = VScroll1.Max / 2
SCROLLER1.YPosition = VScroll1.Value


VScroll2.Max = SCROLLER1.Height / Screen.TwipsPerPixelY
VScroll2.Value = 60
SCROLLER1.SinMaxima = VScroll2.Value

HScroll2.Value = 4
SCROLLER1.SinusStripe = HScroll2.Value

HScroll3.Value = 8
SCROLLER1.SinusPressing = HScroll3.Value

'init the Font's
Text1.Text = SCROLLER1.Font.Name
Text2.Text = SCROLLER1.Font.Size

End Sub

Private Sub HScroll1_Change()
SCROLLER1.Scrollspeed = HScroll1.Value
Frame4.Caption = "Speed:" & HScroll1.Value
End Sub

Private Sub HScroll2_Change()
SCROLLER1.SinusStripe = HScroll2.Value

End Sub

Private Sub HScroll3_Change()
SCROLLER1.SinusPressing = HScroll3.Value
End Sub

Private Sub Option1_Click(index As Integer)
SCROLLER1.Fontbackground = index
End Sub

Private Sub SCROLLER1_Click()
 MsgBox "Scrollpicture Click"
End Sub

Private Sub SCROLLER1_PaintBefore()
Static Y As Single
If Bar_before = False Then Exit Sub
Y = Y + 1
If Y > 200 Then Y = 0
SCROLLER1.PaintPicture Picture1.Image, 1, Y
End Sub

Private Sub SCROLLER1_Paintpast()
Static Y As Single
If Bar_past = False Then Exit Sub

Y = Y + 4
If Y > 200 Then Y = 0
'SCROLLER1.PLOTLINE 1, 1, 200 - y, 200, 200 - y + 20, vbGreen
SCROLLER1.PaintPicture Picture2.Image, 1, 200 - Y

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SCROLLER1.Font.Name = Text1.Text
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SCROLLER1.Font.Size = Text2.Text
End If

End Sub

Private Sub Timer1_Timer()
Static cc As Integer
cc = cc + 1
If cc = 7 Then cc = 1
Select Case cc
 Case 1
  Bar_before = True
 Case 2
  Bar_past = True
 Case 3
  Bar_before = True
  Bar_past = True
 Case 4
  SCROLLER1.Sinus = False
 Case 5
  Bar_before = False
 Case 6
  SCROLLER1.Sinus = True
  Bar_past = False
 Case 7
  HScroll1.Value = 4
 Case 8
  HScroll1.Value = 2
 
 Case 10
 
End Select

End Sub

Private Sub VScroll1_Change()
SCROLLER1.YPosition = VScroll1.Value

End Sub

Private Sub VScroll2_Change()
SCROLLER1.SinMaxima = VScroll2.Value


End Sub


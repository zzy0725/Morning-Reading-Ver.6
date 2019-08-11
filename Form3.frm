VERSION 5.00
Begin VB.Form Form3yy 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   ScaleHeight     =   10005
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "恢复任务"
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "缓存任务"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form3.frx":0000
      Top             =   1440
      Width           =   14895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "英语早读请读英语"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      TabIndex        =   4
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   612
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   120
      Width           =   612
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   840
      Shape           =   3  'Circle
      Top             =   120
      Width           =   612
   End
End
Attribute VB_Name = "Form3yy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Option Explicit
'窗口透明常数
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'窗口透明API
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H1
Const LWA_COLORKEY = &H2
Private Sub Form_Load()
Dim Rec As RECT, hRgn As Long
    GetWindowRect Me.hWnd, Rec
    hRgn = CreateRoundRectRgn(0, 0, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, 50, 50) '这里的两个50是分别用来设置圆角的宽和高的
    SetWindowRgn Me.hWnd, hRgn, True
    'Randomize
'R = Int(Rnd * 256) '红色
'G = Int(Rnd * 256) '绿色
'B = Int(Rnd * 256) '蓝色
'Me.BackColor = rgRGB(R, G, B)
Dim rtn As Long
Me.BackColor = RGB(256, 256, 256) '设置一下窗口的颜色
rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hWnd, RGB(256, 256, 256), 200, LWA_COLORKEY
'RGB(0, 0, 0)参数就是要透明掉的颜色
Timer1.Interval = 250
End Sub

Private Sub Label1_Click()
If MsgBox("退出?", vbYesNo) = vbYes Then
End
End If
End Sub

Private Sub Label2_Click()
Timer1.Enabled = False
Label5.Caption = "早读暂停"
Label2.Visible = False
Shape2.Visible = False
End Sub

Private Sub Label3_Click()
Timer1.Enabled = True
Label5.Caption = Time
Label2.Visible = True
Shape2.Visible = True
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Time
End Sub
Private Sub Command1_Click()
'save as
Open "C://早读/englishtoday.txt" For Output As #1
Print #1, Text1.Text
Close #1
End Sub

Private Sub Command2_Click()
'open as
'Dim nextline As String
Open "C://早读/englishtoday.txt" For Input As #1
Do Until EOF(1)
'Line Input #1, nextline
Text1.Text = Text1.Text + Chr(13) + Chr(10)
Loop
Close #1
End Sub

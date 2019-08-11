VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Left            =   840
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label10 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bug反馈"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   8400
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " ×"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   4540
      TabIndex        =   8
      Top             =   8200
      Width           =   1335
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "关于"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "化学"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      TabIndex        =   5
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "明明可以靠颜值，却偏偏要靠才华"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   5400
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "早读软件     V6"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "生物"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   8160
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "语文"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "英语"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape32 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   735
   End
   Begin VB.Shape Shape31 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape30 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape29 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Shape           =   3  'Circle
      Top             =   840
      Width           =   855
   End
   Begin VB.Shape Shape28 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Shape Shape27 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape Shape26 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   240
      Width           =   1815
   End
   Begin VB.Shape Shape25 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400040&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape Shape24 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Shape Shape23 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   7920
      Width           =   615
   End
   Begin VB.Shape Shape22 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
   Begin VB.Shape Shape21 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape20 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   975
   End
   Begin VB.Shape Shape19 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   -120
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Shape Shape18 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   960
      Width           =   735
   End
   Begin VB.Shape Shape17 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF80FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape16 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape15 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   8040
      Width           =   855
   End
   Begin VB.Shape Shape14 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape Shape13 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   120
      Width           =   2415
   End
   Begin VB.Shape Shape12 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   7920
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   600
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   6120
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'touming
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'窗口透明常数
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1


'jianbian    *********(IGNORE)**********
Private Declare Function SetWindowLongs Lib "user32" _
                    Alias "SetWindowLongsA" _
                    (ByVal hWnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) _
                    As Long
    Private Declare Function GetWindowLongs Lib "user32" _
                    Alias "GetWindowLongsA" _
                    (ByVal hWnd As Long, _
                    ByVal nIndex As Long) _
                    As Long
    
    Private Const GWL_EXSTYLES = (-20)
    Private Const LWA_ALPHAS As Long = &H2
    Private Const WS_EX_LAYEREDS As Long = &H80000
    
    Private Declare Function SetLayeredsWindowAttributes Lib "user32" _
                (ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Long, _
                ByVal dwFlags As Long) _
                As Long
Dim i As Byte
Private Sub cmdClose_Click()
    Timer2.Enabled = True
End Sub


Private Sub Form_Load()
'''''''''''''''''''''窗体透明'''''''''''''''
Me.BackColor = vbBlue

'Text1.BackColor = vbWhite

Dim rtn As Long

rtn = GetWindowLong(hWnd, GWL_EXSTYLE)

rtn = rtn Or WS_EX_LAYERED

SetWindowLong hWnd, GWL_EXSTYLE, rtn

SetLayeredWindowAttributes hWnd, vbBlue, 190, LWA_COLORKEY

'''''''''''''''''''''窗体渐变'''''''''''''''
'Dim p As Long
'    p = GetWindowLongs(Me.hwnd, GWL_EXSTYLE) '取得当前窗口属性
   ' Call SetWindowLongs(Me.hwnd, GWL_EXSTYLE, p Or WS_EX_LAYERED)
   ' '加上一个透明属性
  '  Call SetLayeredsWindowAttributes(Me.hwnd, 0, 0, LWA_ALPHA)
  '  Timer1.Interval = 1
  '  Timer1.Enabled = True
'End Sub
'Private Sub Timer1_Timer()
  '  i = i + 1
  '  Call SetLayeredWindowAttributes(Me.hwnd, 0, i, LWA_ALPHA)
 '   If i >= 255 Then Timer1.Enabled = False
'END IF
End Sub


Private Sub Label1_Click() 'YINGYU API
Me.Visible = False
Form7yybg.Visible = True
Form3yy.Visible = True
End Sub

Private Sub Label10_Click()
Me.Visible = False
frmSplash.Visible = True
End Sub

Private Sub Label2_Click() 'YUWEN API
Me.Visible = False
Form8ywbg.Visible = True
Form4yw.Visible = True
End Sub

Private Sub Label3_Click() 'SHENGWU API
Me.Visible = False
Form9swbg.Visible = True
Form5sw.Visible = True
End Sub

Private Sub Label7_Click() 'HUAXUE API
Me.Visible = False
Form10hxbg.Visible = True
Form6hx.Visible = True
End Sub

Private Sub Label8_Click() 'about
Me.Visible = False
Form2.Visible = True
End Sub

Private Sub Label9_Click() 'exit
End
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Time
End Sub

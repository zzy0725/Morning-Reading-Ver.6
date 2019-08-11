VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7387.64
   ScaleMode       =   0  'User
   ScaleWidth      =   10027.08
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bonjour Inc."
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   5
      Left            =   6240
      TabIndex        =   10
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "¸ßÈý¡¤Ò»°à"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   4
      Left            =   6600
      TabIndex        =   9
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÚÈ¨Ê¹ÓÃ:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   6240
      TabIndex        =   8
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "¼¼ÊõÍÅ¶Ó:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   2
      Left            =   6240
      TabIndex        =   7
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÀ¹¤Ö§³Ö:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   6240
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   1635
      Left            =   6240
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1515
      Left            =   6240
      Picture         =   "Form2.frx":36E15A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3315
   End
   Begin VB.Image imgLogo 
      Height          =   2145
      Left            =   2520
      Picture         =   "Form2.frx":69632C
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " ¡Á"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   748
      Left            =   0
      TabIndex        =   4
      Top             =   3343
      Width           =   735
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "´úÂë±àÐ´:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "V6"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developer Edition"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ôç¶ÁÈí¼þ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   8295
      Left            =   840
      Picture         =   "Form2.frx":696636
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   6120
      TabIndex        =   6
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'touming
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'´°¿ÚÍ¸Ã÷³£Êý
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
'Dim i As Byte
'Private Sub cmdClose_Click()
  '  Timer2.Enabled = True
'End Sub


Private Sub Form_Load()
'''''''''''''''''''''´°ÌåÍ¸Ã÷'''''''''''''''
Me.BackColor = vbBlue

'Text1.BackColor = vbWhite

Dim rtn As Long

rtn = GetWindowLong(hWnd, GWL_EXSTYLE)

rtn = rtn Or WS_EX_LAYERED

SetWindowLong hWnd, GWL_EXSTYLE, rtn

SetLayeredWindowAttributes hWnd, vbBlue, 190, LWA_COLORKEY

'''''''''''''''''''''´°Ìå½¥±ä'''''''''''''''
'Dim p As Long
'    p = GetWindowLongs(Me.hwnd, GWL_EXSTYLE) 'È¡µÃµ±Ç°´°¿ÚÊôÐÔ
   ' Call SetWindowLongs(Me.hwnd, GWL_EXSTYLE, p Or WS_EX_LAYERED)
   ' '¼ÓÉÏÒ»¸öÍ¸Ã÷ÊôÐÔ
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

Private Sub Label9_Click()
Me.Visible = False
Form1.Visible = True
End Sub

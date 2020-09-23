VERSION 5.00
Begin VB.Form frmCustom 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Custom Form"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCustom.frx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Show Transparent Message"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "Name"
      Top             =   920
      Width           =   1630
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enter Password"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1620
      Width           =   1630
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "Password"
      Top             =   1260
      Width           =   1630
   End
   Begin VB.CommandButton Command2 
      Caption         =   "--"
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   195
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Administration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   680
      Width           =   1695
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Custom Colors
Private Const vbGrey = 8421504
Private Const vbOffWhite = 16448250

Private Sub Form_Load()
Set clsTrans = New clsTransparency
clsTrans.DoTransparency Me, ConvertColor(Me.BackColor)   ' RGB(255, 255, 255) '
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frmMessage
Set clsTrans = Nothing
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Me.WindowState = 1
End Sub

Private Sub Command4_Click()
'                        (strMsg As String, _
                          clr As Long, _
                          size As Integer, _
                          Optional bBold As Boolean = True, _
                          Optional bItalic As Boolean = False, _
                          Optional bUnderline As Boolean = False, _
                          Optional fName As String = "MS Sans Serif", _
                          Optional fsclr1 As Long = vbGrey, _
                          Optional fsclr2 As Long = vbOffWhite)
Unload frmMessage
With frmMessage
      .Msg = "Wow||There's Text On My Desktop||I've Never Seen that Before!"
      .fClr = vbRed
      .sze = 14
      .bld = True
      .itlic = True
      .uLine = True
      .fName = "Bookman Old Style"
      .fShadowClr1 = vbGrey
      .fShadowClr2 = vbOffWhite
       ShowWindow .hWnd, SW_SHOWNOACTIVATE
      .Top = Screen.Height / 3
      .Left = Screen.Height / 3
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = vbLeftButton Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End If
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

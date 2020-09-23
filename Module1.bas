Attribute VB_Name = "Module1"
Option Explicit

Public Const SW_SHOWNOACTIVATE = 4

Public clsTrans As clsTransparency

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Function ConvertColor(tColor As Long) As Long
'Converts VB color constants to real color values
If tColor < 0 Then
   ConvertColor = GetSysColor(tColor And &HFF&)
 Else
   ConvertColor = tColor
End If
End Function

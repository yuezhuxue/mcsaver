Attribute VB_Name = "mdl_T"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal HWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
 
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1


Sub SetWinAlpha(Int_TouMing As Byte)
SetWindowLong frmshow.HWnd, GWL_EXSTYLE, GetWindowLong(frmshow.HWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes frmshow.HWnd, 0, Int_TouMing, LWA_ALPHA
End Sub

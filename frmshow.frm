VERSION 5.00
Begin VB.Form frmshow 
   BorderStyle     =   0  'None
   ClientHeight    =   2625
   ClientLeft      =   24510
   ClientTop       =   1425
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdPaste 
      Caption         =   "覆盖"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1860
   End
End
Attribute VB_Name = "frmshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'''窗口半透明声明开始
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'''窗口半透明声明结束
Dim h As Integer, m As Integer, s As Integer '分别存储时分秒


Public Function GetHMS() As Long
h = CInt(Mid(Time$, 1, 2))
m = CInt(Mid(Time$, 4, 2))
s = CInt(Mid(Time$, 7, 2))
'Debug.Print h
'Debug.Print m

End Function
Private Sub cmdPaste_Click() '覆盖
'没有备份文件就不执行
If Dir(App.Path + "\tmp.back") = "" Then
    MsgBox "程序目录没有发现备份文件！", , "tmp.back不存在"
    Exit Sub
End If

Call GetHMS
'覆盖前先备份防止误杀

    Dim Ft
    Set Ft = CreateObject("Scripting.FileSystemObject")
    If Dir(App.Path + "\【刷物】leveldat\") = "" Then MkDir App.Path + "\【刷物】leveldat\"
    Dim Strhc, Strmc, Strsc As String '标签显示用
    If h < 10 Then
        Strhc = "0" + CStr(h)
    Else
        Strhc = CStr(h)
    End If
    If m < 10 Then
        Strmc = "0" + CStr(m)
    Else
        Strmc = CStr(m)
    End If
    If s < 10 Then
        Strsc = "0" + CStr(s)
    Else
        Strsc = CStr(s)
    End If
    Ft.copyfile frm_timer.fram_Save.Caption, App.Path + "\【刷物】leveldat\" + Strhc + "_" + Strmc + "_" + Strsc + "level.dat"
    Kill frm_timer.fram_Save.Caption

'FileCopy frm_timer.Label3.Caption, App.Path + "\tmp.back"

Dim fs ' 先声明一个变体型
Set fs = CreateObject("Scripting.FileSystemObject") '创建文件系统对象fs
fs.copyfile App.Path + "\tmp.back", frm_timer.fram_Save.Caption
'使用该对象的copyfile方法将源文件复制到目标文件
If cmdPaste.BackColor = vbRed Then
    cmdPaste.BackColor = vbYellow
ElseIf cmdPaste.BackColor = vbYellow Then
    cmdPaste.BackColor = vbGreen
Else
    cmdPaste.BackColor = vbRed
End If
'cmdPaste.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End Sub

Private Sub cmdSave_Click()
If Dir(App.Path + "\tmp.back") <> "" Then Kill App.Path + "\tmp.back"
'FileCopy frm_timer.Label3.Caption, App.Path + "\tmp.back"
Dim fs ' 先声明一个变体型
Set fs = CreateObject("Scripting.FileSystemObject") '创建文件系统对象fs
fs.copyfile frm_timer.fram_Save.Caption, App.Path + "\tmp.back"
'使用该对象的copyfile方法将源文件复制到目标文件
If cmdSave.BackColor = vbRed Then
    cmdSave.BackColor = vbYellow
ElseIf cmdSave.BackColor = vbYellow Then
    cmdSave.BackColor = vbGreen
Else
    cmdSave.BackColor = vbRed
End If
'cmdSave.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End Sub

Private Sub Form_Load()
''''''窗口半透明代码开始
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA '  透明度为 0--255 之间的数
''''''窗口半透明代码结束

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmshow
Unload frm_timer
End Sub

Private Sub lbl_Click()
If frm_timer.Visible = True Then
    frm_timer.Visible = False
Else
    frm_timer.Visible = True
End If
End Sub

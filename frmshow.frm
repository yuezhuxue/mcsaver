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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdPaste 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
'''���ڰ�͸��������ʼ
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'''���ڰ�͸����������
Dim h As Integer, m As Integer, s As Integer '�ֱ�洢ʱ����


Public Function GetHMS() As Long
h = CInt(Mid(Time$, 1, 2))
m = CInt(Mid(Time$, 4, 2))
s = CInt(Mid(Time$, 7, 2))
'Debug.Print h
'Debug.Print m

End Function
Private Sub cmdPaste_Click() '����
'û�б����ļ��Ͳ�ִ��
If Dir(App.Path + "\tmp.back") = "" Then
    MsgBox "����Ŀ¼û�з��ֱ����ļ���", , "tmp.back������"
    Exit Sub
End If

Call GetHMS
'����ǰ�ȱ��ݷ�ֹ��ɱ

    Dim Ft
    Set Ft = CreateObject("Scripting.FileSystemObject")
    If Dir(App.Path + "\��ˢ�leveldat\") = "" Then MkDir App.Path + "\��ˢ�leveldat\"
    Dim Strhc, Strmc, Strsc As String '��ǩ��ʾ��
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
    Ft.copyfile frm_timer.fram_Save.Caption, App.Path + "\��ˢ�leveldat\" + Strhc + "_" + Strmc + "_" + Strsc + "level.dat"
    Kill frm_timer.fram_Save.Caption

'FileCopy frm_timer.Label3.Caption, App.Path + "\tmp.back"

Dim fs ' ������һ��������
Set fs = CreateObject("Scripting.FileSystemObject") '�����ļ�ϵͳ����fs
fs.copyfile App.Path + "\tmp.back", frm_timer.fram_Save.Caption
'ʹ�øö����copyfile������Դ�ļ����Ƶ�Ŀ���ļ�
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
Dim fs ' ������һ��������
Set fs = CreateObject("Scripting.FileSystemObject") '�����ļ�ϵͳ����fs
fs.copyfile frm_timer.fram_Save.Caption, App.Path + "\tmp.back"
'ʹ�øö����copyfile������Դ�ļ����Ƶ�Ŀ���ļ�
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
''''''���ڰ�͸�����뿪ʼ
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA '  ͸����Ϊ 0--255 ֮�����
''''''���ڰ�͸���������

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

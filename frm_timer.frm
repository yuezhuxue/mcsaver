VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_timer 
   Caption         =   "MC�浵�Զ�����"
   ClientHeight    =   8175
   ClientLeft      =   1185
   ClientTop       =   1170
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   10230
   Begin VB.Frame Frame1 
      Caption         =   "���Զ�����ģ�顿����״̬"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5760
      TabIndex        =   11
      Top             =   120
      Width           =   4215
      Begin VB.Label lblShow1 
         Caption         =   "�����ʱ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblShow2 
         Caption         =   "�����״̬"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label lblShow3 
         Caption         =   "����ʱ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   3735
      End
   End
   Begin VB.Frame fram_1 
      Caption         =   "С����͸���ȣ� 200"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.VScrollBar VScroll1 
         Height          =   2175
         Left            =   1560
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Value           =   200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ʾ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "�ƶ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��ɫ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1095
      End
   End
   Begin VB.Frame fram_Save 
      Caption         =   "�˴���ʾ�浵·��������ѡ��level.dat��·��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   9615
      Begin VB.Frame Frame2 
         Caption         =   "ѡ�񱣴��Դ����ɫΪ�أ���ɫΪ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   5175
         Begin VB.OptionButton Opt2 
            Caption         =   "����浵"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   17
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "����浵"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   16
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame fram_HowLong 
         Caption         =   "������ 5 ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   5175
         Begin VB.HScrollBar HScroll2 
            Height          =   735
            Left            =   120
            Max             =   60
            Min             =   1
            TabIndex        =   10
            Top             =   360
            Value           =   5
            Width           =   4935
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�����ֶ�����"
         Height          =   615
         Left            =   5640
         TabIndex        =   3
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H000000FF&
         Height          =   735
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4200
         Top             =   1680
      End
      Begin VB.CommandButton Command3 
         Caption         =   "����Դ�浵level.dat·��"
         Height          =   615
         Left            =   5640
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   4680
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "dat"
         FileName        =   "level.dat"
         Filter          =   "dat"
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   240
   End
End
Attribute VB_Name = "frm_timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim a As Integer
Dim strM As String, strS As String
Dim Int_Count As Double
Dim str_Temp As String

Dim h As Integer, m As Integer, s As Integer '�ֱ�洢ʱ����
Dim hs, ms, ss As Integer '��ʼʱ���
Dim Int_Time, Save_Time As Integer
Dim ha, ma, sa As Integer '��ǩ��ʾ��
Dim Strha, Strma, Strsa As String '��ǩ��ʾ��
'-----


Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function SetTopMostWindow(tHWND As Long, Topmost As Boolean) As Long
 If Topmost = True Then ''Make the window topmost
  SetTopMostWindow = SetWindowPos(tHWND, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 Else
  SetTopMostWindow = SetWindowPos(tHWND, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
  SetTopMostWindow = False
 End If
End Function



Public Function GetHMS() As Long
h = CInt(Mid(Time$, 1, 2))
m = CInt(Mid(Time$, 4, 2))
s = CInt(Mid(Time$, 7, 2))
'Debug.Print h
'Debug.Print m

End Function



Private Sub Chk1_Click()
If chk1.Value = 1 Then
chk1.BackColor = vbGreen
Timer2.Enabled = True
'����ʱ���
Call GetHMS
hs = h
ms = m
ss = s
'Me.Cls
'Print CStr(h) + ":" + CStr(m) + ":" + CStr(s) + " �����ã�"

Else
chk1.BackColor = vbRed
Timer2.Enabled = False
End If


End Sub

Private Sub cmd2_Click()
If frmshow.BorderStyle = 0 Then
frmshow.BorderStyle = 1
frmshow.Caption = frmshow.Caption
cmd2.Caption = "�̶�"
Else
frmshow.BorderStyle = 0
frmshow.Caption = frmshow.Caption
cmd2.Caption = "�ƶ�"
End If
End Sub

Private Sub Command1_Click()
frmshow.Show
Timer1.Interval = 1000 '���ü�ʱ��Ϊһ�뷢��һ��
Timer1.Enabled = True '�����ʱ��
End Sub

Private Sub Command2_Click()
'Call GetHMS
CommonDialog1.ShowColor
frmshow.lbl.ForeColor = CommonDialog1.Color
End Sub

Private Sub Command3_Click()
CommonDialog2.ShowOpen
fram_Save.Caption = CommonDialog2.FileName
If Dir(App.Path & "/saver.ini") <> "" Then Kill App.Path & "/saver.ini"
Open App.Path & "/saver.ini" For Output As #1
Print #1, fram_Save.Caption
Close #1
End Sub



Private Sub Form_Load()
Randomize
SetTopMostWindow frmshow.hwnd, True
Timer1.Enabled = False '�ȹرռ�ʱ��




If Dir(App.Path & "/saver.ini") <> "" Then
Open App.Path & "/saver.ini" For Input As #2
Line Input #2, str_Temp
frm_timer.fram_Save.Caption = str_Temp
Close #2
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frm_timer
Unload frmshow
End Sub




Private Sub HScroll2_Change()
Static Int_FenZhong As Byte
Int_FenZhong = HScroll2.Value
fram_HowLong.Caption = "������ " + CStr(Int_FenZhong) + " ����"
Save_Time = 60 * Int_FenZhong
End Sub

Private Sub Timer1_Timer()

'Label1.Caption = Time$
frmshow.lbl.Caption = lblShow3.Caption
End Sub


Private Sub Timer2_Timer()
Call GetHMS

Int_Time = 3600 * (h - hs) + 60 * (m - ms) + (s - ss)
Save_Time = 60 * HScroll2.Value
Debug.Print "������" + CStr(Int_Time)



If Int_Time < Save_Time Then
    Int_Time = Save_Time - Int_Time
    Call SaveMC
Else
    Int_Time = Int_Time Mod Save_Time
    Int_Time = Save_Time - Int_Time
    Call SaveMC
End If
    '----����ʱ��ǩ
'Int_Time = 180 - Int_Time
ma = CInt(Int_Time \ 60)
sa = Int_Time Mod 60
If ma < 10 Then Strma = "0" + CStr(ma) Else Strma = CStr(ma)
If sa < 10 Then Strsa = "0" + CStr(sa) Else Strsa = CStr(sa)
lblShow3.Caption = Strma + ":" + Strsa


'����1h������ʱ���
If m - ms >= 30 Then
hs = h
ms = m
ss = s
'Me.Cls
'Print CStr(h) + ":" + CStr(m) + ":" + CStr(s) + " �����ã�"
End If
End Sub


Public Function SaveMC() As Long
'ѡ�񱣴��Դ
If Opt1.Value = True Then
    Dim i As Integer
    If Int_Time = 1 Then
        Int_Count = Int_Count + 1
        lblShow1.Caption = CStr(h) + ":" + CStr(m) + ":" + CStr(s)
        lblShow2.Caption = "��" + CStr(Int_Count) + "�� ����ɹ���"
    '����·��
        str_Temp = fram_Save.Caption
        i = InStrRev(str_Temp, "\")
    '����·��
        str_Temp = Mid(str_Temp, 1, i - 1)
    'str_Temp = App.Path + "\000"
        MkDir App.Path + "\" + "��" + CStr(Int_Count) + "��" + CStr(h) + "_" + CStr(m) + "_" + CStr(s)
        CopyPath str_Temp, App.Path + "\" + "��" + CStr(Int_Count) + "��" + CStr(h) + "_" + CStr(m) + "_" + CStr(s)
    End If
'Debug.Print h
'Debug.Print m
ElseIf Opt2.Value = True Then
    If Int_Time = 1 Then
        lblShow1.Caption = CStr(h) + ":" + CStr(m) + ":" + CStr(s)
        lblShow2.Caption = CStr(h) + ":" + CStr(m) + ":" + CStr(s) + "����ɹ���"
    
        Dim Ft
        Set Ft = CreateObject("Scripting.FileSystemObject")
            If Dir(App.Path + "\���浵��leveldat\") = "" Then MkDir App.Path + "\���浵��leveldat\"
            Dim Strhb, Strmb, Strsb As String '��ǩ��ʾ��
            If h < 10 Then
                Strhb = "0" + CStr(h)
            Else
                Strhb = CStr(h)
            End If
            If m < 10 Then
                Strmb = "0" + CStr(m)
            Else
                Strmb = CStr(m)
            End If
            If s < 10 Then
                Strsb = "0" + CStr(s)
            Else
                Strsb = CStr(s)
            End If
        Ft.copyfile frm_timer.fram_Save.Caption, App.Path + "\���浵��leveldat\" + Strhb + "_" + Strmb + "_" + Strsb + "level.dat"
    End If
End If
End Function
Private Sub Command4_Click()
Dim i As Integer
Call GetHMS
Int_Count = Int_Count + 1
lblShow1.Caption = CStr(h) + ":" + CStr(m) + ":" + CStr(s)
lblShow2.Caption = "��" + CStr(Int_Count) + "�� ����ɹ���"
    '����·��
    str_Temp = fram_Save.Caption
    i = InStrRev(str_Temp, "\")
    '����·��
    str_Temp = Mid(str_Temp, 1, i - 1)
    'str_Temp = App.Path + "\000"
MkDir App.Path + "\" + "��" + CStr(Int_Count) + "��" + CStr(h) + "_" + CStr(m) + "_" + CStr(s)
CopyPath str_Temp, App.Path + "\" + "��" + CStr(Int_Count) + "��" + CStr(h) + "_" + CStr(m) + "_" + CStr(s)

'Debug.Print App.Path + "\" + "��" + CStr(Int_Count) + "��" + CStr(h) + "_" + CStr(m) + "_" + CStr(s)
'Debug.Print str_Temp
End Sub

Private Sub VScroll1_Change()
Static Int_TouMing As Byte
Int_TouMing = VScroll1.Value
fram_1.Caption = "ˢ�͸���ȣ� " + CStr(Int_TouMing)
SetWinAlpha Int_TouMing
End Sub

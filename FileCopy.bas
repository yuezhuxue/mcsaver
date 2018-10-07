Attribute VB_Name = "FileCopy"
Option Explicit

'------------------------------------
'定义win文件夹操作的函数（复制和删除，其它的(移动、改名)没用到）
'删除文件夹的函数：KillPath(path)
'复制文件夹的函数：CopyPath(mpath,tPath)
'------------------------------------
Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SILENT = &H4
Private Const FOF_NOERRORUI = &H400
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String 'only used if FOF_SIMPLEPROGRESS
End Type
'删除文件夹的函数：KillPath(path)
Public Function KillPath(ByVal sPath As String) As Boolean
On Error Resume Next
Dim udtPath As SHFILEOPSTRUCT
udtPath.hwnd = 0
udtPath.wFunc = FO_DELETE
udtPath.pFrom = sPath
udtPath.pTo = ""
udtPath.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
KillPath = Not CBool(SHFileOperation(udtPath))
End Function
'复制文件夹的函数：CopyPath(mpath,tPath)
Public Function CopyPath(ByVal mPath As String, ByVal tPath As String) As Boolean
On Error Resume Next
Dim shfileop As SHFILEOPSTRUCT
shfileop.hwnd = 0
shfileop.wFunc = FO_COPY
shfileop.pFrom = mPath
shfileop.pTo = tPath
shfileop.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI
CopyPath = Not CBool(SHFileOperation(shfileop))
End Function
'------------------------
'------------------------

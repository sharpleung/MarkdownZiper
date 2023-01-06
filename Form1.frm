VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MarkdownPicer  By Gqleung"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6615
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":444A
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版本：0.2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   210
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "inside Pandora's Box"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      TabIndex        =   4
      Top             =   2400
      Width           =   2700
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.plasf.cn"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "   程序默认将mardown文档中的网络图片或者本地绝对路径图片复制到img文件夹。"
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "将目标拖入图片处即可"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":4450
      Top             =   -120
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long



Function StrReplace(S As String, p As String, r As String) As String
 Dim re As regexp
 Set re = New regexp
 re.IgnoreCase = True
 re.Global = True
 re.Pattern = p
 StrReplace = re.replace(S, r)
End Function

Function RegexpReplace(regexp As String, target As String, replace As String) As String
 RegexpReplace = StrReplace(target, regexp, replace)
End Function
Function RegEx(r As String, target As String) As MatchCollection
 Dim re As regexp
 Dim mhs   As MatchCollection
 Set re = New regexp
 re.IgnoreCase = True
 re.Global = True
 re.Pattern = r
 Set RegEx = re.Execute(target)
End Function
'创建文件夹
Function mkdir(dirname As String) As String
Dim fso As New FileSystemObject
Dim newfolder As Folder
Dim temp As String
Dim self_dir As String
If (fso.FolderExists(dirname) = True) Then
    If MsgBox(dirname & "文件夹已经存在,是否自定义文件夹", vbYesNo, "请确认") = vbYes Then
        temp = Split(dirname, "\")(UBound(Split(dirname, "\")))
        temp = replace(dirname, temp, "")
        self_dir = InputBox("请输入新建文件夹名：", "新建文件夹")
        If self_dir <> "" Then
            dirname = temp & self_dir
            fso.CreateFolder (dirname)
        End If
    Else
        mkdir = dirname
        Exit Function
    End If
    Else
    fso.CreateFolder (dirname)
End If
Set newfolder = Nothing
mkdir = dirname
End Function
Function DownPic(url As String, dir As String, filename As String)
Dim fs As New FileSystemObject
Dim fpath As String
url = replace(url, ")", "")
url = replace(url, "(", "")
url = replace(url, "]", "")
url = replace(url, "src=", "")
url = replace(url, """", "")
fpath = dir & "\" & filename

If fs.FileExists(fpath) <> True Then
    If URLDownloadToFile(0, url, fpath, 0, 0) = 0 Then
           temp = Split(dir, "\")(UBound(Split(dir, "\")))
           Text1.Text = replace(Text1.Text, url, ".\" & temp & "\" & filename)
    End If
Else
    temp = Split(dir, "\")(UBound(Split(dir, "\")))
    Text1.Text = replace(Text1.Text, url, ".\" & temp & "\" & filename)
End If

End Function
Function checkpath(path As String) As Boolean
    Dim mhs   As MatchCollection
    path = replace(path, ")", "")
    path = replace(path, """", "")
    path = replace(path, "]", "")
    path = replace(path, "(", "")
    Set mhs = RegEx("^\./[\u4e00-\u9fa5_a-zA-Z0-9]|^\.\\[\u4e00-\u9fa5_a-zA-Z0-9]|^[\u4e00-\u9fa5_a-zA-Z0-9]+", path)
    If mhs.Count > 0 Then
    checkpath = True
    Exit Function
    End If
    checkpath = False
End Function
Function cy2file(fpath As String, dir As String, Repath As String)
        fpath = replace(fpath, "/", "\")
        temp = fpath
        temp = replace(temp, ")", "")
        temp = replace(temp, "(", "")
        temp = replace(temp, "]", "")
        temp = replace(temp, "src=", "")
        temp = replace(temp, """", "")
        filename = Split(fpath, "\")(UBound(Split(fpath, "\")))
        filename = replace(filename, ")", "")
        filename = replace(filename, """", "")
        filename = replace(filename, "]", "")
        filename = replace(filename, "(", "")
On Error GoTo catch
try:
        FileCopy temp, dir & "\" & filename  '复制文件
        tempdir = Split(dir, "\")(UBound(Split(dir, "\")))
        If Len(Repath) > 0 Then
            Text1.Text = replace(Text1.Text, Repath, ".\" & tempdir & "\" & filename)
        Else
            Text1.Text = replace(Text1.Text, temp, ".\" & tempdir & "\" & filename)
        End If
catch:
    Debug.Print "文件已经存在"
End Function
Function SavePic(content As String, url As String)
    Dim filename As String
    Dim dir As String
    Dim mhs   As MatchCollection
     Dim path   As MatchCollection
    Dim temp As String
    Dim tempdir As String
    Dim flag As Boolean
    Dim ext As String
    flag = False
    Set mhs = RegEx("]\((.*?)\)|src=""(.*?)""", content)
    If mhs.Count > 0 Then
    dir = mkdir(url & "\img")
    For i = 0 To mhs.Count - 1
        If InStr(1, mhs(i), "http://") Or InStr(1, mhs(i), "https://") Then
        temp = Split(mhs(i), "/")(UBound(Split(mhs(i), "/")))
        filename = replace(temp, ")", "")
        filename = replace(filename, """", "")
        filename = replace(filename, "]", "")
        filename = replace(filename, "(", "")
        ext = Split(filename, ".")(UBound(Split(filename, ".")))
        ext = UCase(ext)
            If InStr(1, ext, "PNG") Or InStr(1, ext, "JPG") Or InStr(1, ext, "JPEG") Or InStr(1, ext, "GIF") Or InStr(1, ext, "BMP") Then
                DownPic mhs(i), dir, filename
            End If
        ElseIf checkpath(mhs(i)) And InStr(1, mhs(i), ":") = 0 Then
            If flag = False Then
                If MsgBox("此文档中存在相对路径图片,是否一同复制到新文件夹", vbYesNo, "请确认") = vbYes Then
                    flag = True
                    temp = replace(mhs(i), ")", "")
                    temp = replace(temp, """", "")
                    temp = replace(temp, "]", "")
                    temp = replace(temp, "(", "")
                    temp = replace(temp, "src=", "")
                    Set path = RegEx("^\./[\u4e00-\u9fa5_a-zA-Z0-9]|^\.\\[\u4e00-\u9fa5_a-zA-Z0-9]", temp)
                    If path.Count > 0 Then
                        temp = replace(mhs(i), ".\\", "")
                        temp = replace(temp, "./", "")
                        cy2file url & "\" & temp, dir, temp
                    End If
                    Set path = RegEx("^[\u4e00-\u9fa5_a-zA-Z0-9_._\__!_@_#_\$_%_^_&_*_\-_\+]+/", temp)
                    If path.Count > 0 Then
                        cy2file url & "\" & temp, dir, temp
                    End If
                Else
                    GoTo e
                End If
            Else
                temp = replace(mhs(i), ")", "")
                temp = replace(temp, """", "")
                temp = replace(temp, "]", "")
                temp = replace(temp, "(", "")
                 temp = replace(temp, "src=", "")
                Set path = RegEx("^\./[\u4e00-\u9fa5_a-zA-Z0-9]|^\.\\[\u4e00-\u9fa5_a-zA-Z0-9]", temp)
                If path.Count > 0 Then
                    temp = replace(mhs(i), ".\\", "")
                    temp = replace(temp, "./", "")
                    cy2file url & "\" & temp, dir, temp
                End If
                Set path = RegEx("^[\u4e00-\u9fa5_a-zA-Z0-9_._\__!_@_#_\$_%_^_&_*_\-_\+]+/", temp)
                If path.Count > 0 Then
                    cy2file url & "\" & temp, dir, temp
                End If
            End If
        Else
            cy2file mhs(i), dir, ""
        End If
e:
    Next i
    End If
End Function

 Function WriteToFile(file, Message)
        Dim Stm1
        Set Stm1 = CreateObject("ADODB.Stream")
        Stm1.Type = 2
        Stm1.Open
        Stm1.Charset = "UTF-8"
        'Stm1.Charset = "Unicode"
        Stm1.position = Stm1.Size
        Stm1.WriteText Message
        Stm1.SaveToFile file, 2
        Stm1.Close
        Set Stm1 = Nothing
    End Function




Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim filename  As String
Dim temp As String
If Data.Files.Count > 0 Then
    filename = Data.Files(1)
    Dim objStream, strData
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (filename)
    strData = objStream.ReadText()
    Text1.Text = strData
    temp = Split(Data.Files(1), "\")(UBound(Split(Data.Files(1), "\")))
    temp = replace(Data.Files(1), "\" & temp, "")
    SavePic Text1.Text, temp
    objStream.Close
    Set objStream = Nothing
    If MsgBox("是否覆盖原文,否则生成副本", vbYesNo, "请确认") = vbYes Then
        WriteToFile filename, Text1.Text
    Else
        WriteToFile replace(filename, ".md", "_副本.md"), Text1.Text
    End If
     MsgBox "修改完成", vbOKOnly, "提示"
End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "У԰���Զ���¼ϵͳ"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox autorun 
      Caption         =   "�����Զ�����"
      Height          =   180
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox uid 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox PWD 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "���迪���Զ����У��뽫������źú��ٹ�ѡ���������޸�ע�����ȫ�������ʾ����ѡ������"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "����дУ԰���˺�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "�˺�:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "����:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "�����˺����뽫�ᱣ���ڱ����������ʹ�á�"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "�޸ĵ�¼��Ϣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Powered by ZGCloud"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "���ڼ��У԰��  �Ƿ������֤"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||�������ZGCloud���  д�������࣬�����ظ����Ͷ��ɹ�                            |||||||||
'|||||||||����305623673@qq.com ���������������������������ϵɾ��                      |||||||||
'|||||||||�����ж��ķ�����Դ������������ֽ�ֹ�޸�                                       |||||||||
'|||||||||��Դ�ڻ��������еĹ�����Դ���ʾ:                                              |||||||||
'|||||||||getHTTPPage()  ReadIniFile()   WriteIniFile()   Sleep()                        |||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'Dir([pathname], [Attributes as VbFileAttribute=vbNormal]) As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long
Function ReadIniFile(filename As String, AppName As String, KeyName As String) As String
Dim temp As String * 100
Dim n As Long
n = GetPrivateProfileString(AppName, KeyName, "", temp, Len(temp), filename)
ReadIniFile = Mid(temp, 1, n)
End Function
Function WriteIniFile(filename As String, AppName As String, KeyName As String, NewKeyName As String)
Dim n As Long
n = WritePrivateProfileString(AppName, KeyName, NewKeyName, filename)
End Function
Function getHTTPPage(url)
On Error Resume Next
  Dim http
  Set http = CreateObject("MSXML2.XMLHTTP")
 http.Open "GET", url, False
   getHTTPPage = http.Send()
'  MsgBox http.ReadyState
'  If http.ReadyState <> 4 Then
'  MsgBox "�޷����ӷ�����"
'  End
'     Exit Function
'  End If
  getHTTPPage = BytesToBstr(http.responseBody, "GB2312")
  Set http = Nothing
End Function
Function BytesToBstr(body, Cset)
      Dim objstream
      Set objstream = CreateObject("adodb.stream")
      objstream.Type = 1
      objstream.Mode = 3
      objstream.Open
      objstream.Write body
      objstream.position = 0
      objstream.Type = 2
      objstream.Charset = Cset
      BytesToBstr = objstream.ReadText
      objstream.Close
      Set objstream = Nothing
End Function
'Private Sub autorun_Click()
'Set w = CreateObject("wscript.shell")
'If autorun.Value = True Then
'w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
'Call WriteIniFile("d:\login.ini", "LOGIN", "auto", "True")
'End If
'If autorun.Value = False Then
'w.regdelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName
'Call WriteIniFile("d:\login.ini", "LOGIN", "auto", "False")
'End If
'End Sub
Private Sub Command1_Click()
End
End Sub
Private Sub Command2_Click()
If Dir("d:/login.ini") <> "" Then
Kill "d:/login.ini"
End If
If PWD.Text = "" Then
MsgBox "������������Ϣ��", , "Error"
Else
Call WriteIniFile("d:\login.ini", "LOGIN", "UID", uid.Text)
Call WriteIniFile("d:\login.ini", "LOGIN", "PWD", PWD.Text)
Form1.Height = 1455
checkid
End If
End Sub
Private Sub checknet()
On Error Resume Next
Dim S As String
Shell "cmd /c ipconfig -all>d:\1.txt", vbHide
Sleep (1000)
Open "d:\1.txt" For Binary As #1
a = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1
b = Split(a, "   DNS ������  . . . . . . . . . . . : ")(1)
c = Split(b, vbCrLf)(0)
Kill "d:/1.txt"
If c <> "10.10.252.4" Then
End
End If
End Sub
Private Sub First()
MsgBox "�����״�ʹ�ñ����������д�˺����롣", , "��ܰ��ʾ:"
Label3_Click
Label1.Caption = "�벹ȫ��½��Ϣ"
End Sub
Private Sub read()
uid.Text = ReadIniFile("d:\login.ini", "LOGIN", "UID")
PWD.Text = ReadIniFile("d:\login.ini", "LOGIN", "PWD")
getresult
End Sub
Private Sub getresult()
a = getHTTPPage("http://172.16.30.45/drcom/login?callback=dr1569456885718&DDDDD=" & ReadIniFile("d:\login.ini", "LOGIN", "UID") & "&upass=" & ReadIniFile("d:\login.ini", "LOGIN", "PWD") & "&0MKKey=123456&R1=0&R3=1&R6=0&para=00&v6ip=&_=1569456869405")
If InStr(1, a, "NID") = 0 Then
a = getHTTPPage("http://172.16.30.45/drcom/login?callback=dr1569456885718&DDDDD=" & ReadIniFile("d:\login.ini", "LOGIN", "UID") & "&upass=" & ReadIniFile("d:\login.ini", "LOGIN", "PWD") & "&0MKKey=123456&R1=0&R3=1&R6=0&para=00&v6ip=&_=1569456869405")
End If
If InStr(1, a, "error") = 0 Then
b = Split(a, Chr(34) & "NID" & Chr(34) & ":" & Chr(34))(1)
c = Split(b, Chr(34))(0)
Label1.Caption = "��ӭ��:" & c & "���ѳɹ���¼��"
Sleep (1500)
End
Else
MsgBox "��֤ϵͳ���ش�����Ϣ�������˺Ż������Ƿ����"
Form1.Height = 4125
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
If Dir("d:/login.ini") = "" Then
First
Else
read
End If
End Sub
Private Sub Form_Load()
Sleep (3000)
checknet
End Sub
Private Sub Label3_Click()
Form1.Height = 4125
End Sub
Private Sub checkid()
a = getHTTPPage("http://172.16.30.45/drcom/login?callback=dr1569456885718&DDDDD=" & ReadIniFile("d:\login.ini", "LOGIN", "UID") & "&upass=" & ReadIniFile("d:\login.ini", "LOGIN", "PWD") & "&0MKKey=123456&R1=0&R3=1&R6=0&para=00&v6ip=&_=1569456869405")
'MsgBox InStr(1, a, "error")
If InStr(1, a, "error") <> 0 Then
MsgBox "��֤ϵͳ���ش�����Ϣ�������˺Ż������Ƿ����"
Form1.Height = 4125
Else
b = Split(a, Chr(34) & "NID" & Chr(34) & ":" & Chr(34))(1)
c = Split(b, Chr(34))(0)
Label1.Caption = "��ӭ��:" & c & "���ѳɹ���¼��"
Sleep (3000)
End
End If
End Sub


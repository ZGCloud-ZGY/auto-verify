VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   LinkTopic       =   "Form2"
   ScaleHeight     =   2985
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "填好了"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox PWD 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox uid 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "您的账号密码将会保存在本机，请放心使用。"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "密码:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "账号:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请填写校园网账号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Begin VB.Form Update 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "班班通助手 ☆ 更新日志"
   ClientHeight    =   5340
   ClientLeft      =   2805
   ClientTop       =   5700
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5025
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "启航版 2012年12月01日 版本号 version 1.1.0"
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "- 恢复屏幕键盘为系统自带"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "- 新增U盘自动识别功能，更智能"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "- 界面美化，去除不协调的图标元素，按钮颜色调整"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "启程版 2012年11月26日 版本号 version 1.0.0"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "- 一键关机重启增加5秒确认，防止误按"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "- 增加360任务管理，自动关机（原创）功能"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "- 修复班班通客户端打不开问题，增加英语听力、迅雷"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "测试版 2012年11月24日 版本号 version 0.9"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "- 打开U盘：直接打开我的电脑中第一个U盘(必须存在)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "- 手写输入法：搜狗技术支持"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "- 放映课件：跳过幻灯片程序直接从第一张开始放映"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "下一版本：??版 2013年 版本号 version 2.0"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   4695
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "- 欢迎反馈意见建议"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   0
      Picture         =   "Update.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   6555
      TabIndex        =   14
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "访问我的博客"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "- 鼠标自动点击启动白板并全屏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   4200
         Width           =   4455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "- 末日狂欢修订版v1.2 智能课程表，U盘一键弹出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   3960
         Width           =   4455
      End
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://blog.sina.com.cn/imaegoo", 1
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form ZZBBT 
   BackColor       =   &H00FFFFFF&
   Caption         =   "班班通助手 v1.6 Final 正式版"
   ClientHeight    =   5310
   ClientLeft      =   6225
   ClientTop       =   5685
   ClientWidth     =   6495
   Icon            =   "ZZBBT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "ZZBBT.frx":4322
   ScaleHeight     =   5310
   ScaleWidth      =   6495
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5310
      Left            =   0
      Picture         =   "ZZBBT.frx":74E96
      ScaleHeight     =   5310
      ScaleWidth      =   6510
      TabIndex        =   0
      Top             =   0
      Width           =   6510
      Begin VB.Timer Timer9 
         Interval        =   9000
         Left            =   5880
         Top             =   2880
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   0
         Picture         =   "ZZBBT.frx":E5A0A
         ScaleHeight     =   5295
         ScaleWidth      =   6495
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   6495
         Begin VB.Timer Timer7 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   360
            Top             =   360
         End
         Begin VB.CommandButton Commandthank 
            BackColor       =   &H005EDFBF&
            Caption         =   "不用，谢谢"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   4080
            Width           =   1935
         End
         Begin VB.Timer Timer8 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   840
            Top             =   360
         End
         Begin VB.Label daojs 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   90
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2175
            Left            =   2040
            TabIndex        =   6
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "数学"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   3840
            TabIndex        =   5
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "当前课程——"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   1680
            TabIndex        =   4
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "秒后自动开启白板"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   240
            TabIndex        =   3
            Top             =   3360
            Width           =   6015
         End
      End
      Begin VB.CommandButton CommandData 
         BackColor       =   &H00FFFF80&
         Caption         =   "驱动重启"
         Height          =   255
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CommandExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "隐藏"
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdUnLoad 
         BackColor       =   &H00FFFF80&
         Caption         =   "安全移除"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CommandInternet 
         BackColor       =   &H00FFFF80&
         Caption         =   "重新检测"
         Height          =   255
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   3120
      End
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   0
         Top             =   2640
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   0
         Top             =   2160
      End
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   0
         Top             =   1680
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10500
         Left            =   0
         Top             =   1200
      End
      Begin VB.CommandButton CommandOFF 
         BackColor       =   &H006262FF&
         Caption         =   "◆ 一键关机!"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CommandReset 
         BackColor       =   &H0000FFFF&
         Caption         =   "◇ 一键重启!"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer Timer6 
         Interval        =   1000
         Left            =   0
         Top             =   3600
      End
      Begin VB.CommandButton CommandBBT 
         BackColor       =   &H00FED78D&
         Caption         =   "班班通客户端"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandPPT2 
         BackColor       =   &H00FDC660&
         Caption         =   "放映课件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton CommandPPT1 
         BackColor       =   &H00FED78D&
         Caption         =   "打开课件"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandBoard 
         BackColor       =   &H00FDC660&
         Caption         =   "写字白板"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton CommandHand 
         BackColor       =   &H006AE1FF&
         Caption         =   "手写输入法"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandKey 
         BackColor       =   &H006AE1FF&
         Caption         =   "屏幕键盘"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandU 
         BackColor       =   &H0040D9FF&
         Caption         =   "打开U盘"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton CommandD 
         BackColor       =   &H006AE1FF&
         Caption         =   "打开硬盘"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton CommandExplorer 
         BackColor       =   &H005EDFBF&
         Caption         =   "访问百度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton CommandPlayer 
         BackColor       =   &H0076E4C9&
         Caption         =   "Qvod播放器"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandMusic 
         BackColor       =   &H0076E4C9&
         Caption         =   "英语听力"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton CommandVideo 
         BackColor       =   &H0076E4C9&
         Caption         =   "启动迅雷7"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandAbout 
         BackColor       =   &H00B782FF&
         Caption         =   "关于 About"
         Height          =   255
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandAutoShut 
         BackColor       =   &H00A35EFF&
         Caption         =   "自动关机"
         Height          =   255
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CommandNB 
         BackColor       =   &H00B782FF&
         Caption         =   "系统监视"
         Height          =   255
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command360 
         BackColor       =   &H00B782FF&
         Caption         =   "任务管理"
         Height          =   255
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   0
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "请选择需要放映的幻灯片"
         Flags           =   12
         InitDir         =   "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "请选择需要预览、编辑的幻灯片"
         Flags           =   12
         InitDir         =   "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
      End
      Begin VB.CommandButton CommandUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "日志"
         Height          =   255
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer TimerInternet 
         Interval        =   60000
         Left            =   0
         Top             =   4080
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1020
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   3840
         Width           =   5775
      End
      Begin VB.CommandButton CommandEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "手写编辑"
         Height          =   255
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4900
         Width           =   855
      End
      Begin VB.CommandButton CommandClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "清空"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4900
         Width           =   495
      End
      Begin VB.CommandButton CommandOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "完成"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4900
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FCF5F1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "正在自检 ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Shape Shape5 
         Height          =   255
         Left            =   1440
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         Height          =   375
         Left            =   4920
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         Height          =   255
         Left            =   4920
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape Shape2 
         Height          =   255
         Left            =   3720
         Top             =   1320
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   3720
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   39
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   855
         Left            =   4440
         TabIndex        =   53
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "天，高考"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   3240
         TabIndex        =   52
         Top             =   3280
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "距月考还有"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   360
         TabIndex        =   51
         Top             =   3280
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "天"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   5760
         TabIndex        =   50
         Top             =   3280
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   39
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   1920
         TabIndex        =   49
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "天"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   6480
         TabIndex        =   48
         Top             =   3195
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Designed ＆UI ＆Powered by iMaeGoo_虹墨"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   4940
         Width           =   3510
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "============收起===↑↑↑===更多============"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   2880
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "============展开===↓↓↓===更多============"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   2880
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "=★当前00:00:00 已运行00分00秒★="
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label LabelU 
         BackColor       =   &H00FCF5F1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "正在检测 (F:\) ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   38
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FCF5F1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "正在检测网络 ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   4800
         X2              =   4800
         Y1              =   600
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   4920
         X2              =   6240
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "班班通系统"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "U盘状态"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "联网状态"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "教学工具——"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "常用工具——"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "多媒体工具——"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4680
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "WARNING... 系统工具——"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   4560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FEDC9C&
         BorderWidth     =   2
         X1              =   2280
         X2              =   2280
         Y1              =   2040
         Y2              =   4200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FEDC9C&
         BorderWidth     =   2
         X1              =   4320
         X2              =   4320
         Y1              =   2040
         Y2              =   4200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000011&
         X1              =   240
         X2              =   6240
         Y1              =   5040
         Y2              =   5040
      End
   End
End
Attribute VB_Name = "ZZBBT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Dim tm As String, mn As Integer, sc As Integer

'██████████叉号命令██████████
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Class
Unload USBMode
Unload Update
Unload Page
End
End Sub

Private Sub Label13_Click()
CommandPPT1.Visible = True
CommandBBT.Visible = True
CommandKey.Visible = True
CommandHand.Visible = True
CommandPlayer.Visible = True
CommandVideo.Visible = True
Label10.Visible = True
Command360.Visible = True
CommandNB.Visible = True
CommandAutoShut.Visible = True
CommandAbout.Visible = True
CommandData.Visible = True
CommandInternet.Visible = True
CommandUp.Visible = True
CommandReset.Visible = True
Label2.Visible = True
Label15.Visible = True
Label20.Visible = False
Label16.Visible = False
Label19.Visible = False
Label22.Visible = False
Label18.Visible = False
Text1.Visible = False
Line5.Visible = False
Label17.Visible = False
CommandEdit.Visible = False
CommandClear.Visible = False
CommandOK.Visible = False
Label13.Visible = False
End Sub

Private Sub Label15_Click()
CommandPPT1.Visible = False
CommandBBT.Visible = False
CommandKey.Visible = False
CommandHand.Visible = False
CommandPlayer.Visible = False
CommandVideo.Visible = False
Label10.Visible = False
Command360.Visible = False
CommandNB.Visible = False
CommandAutoShut.Visible = False
CommandAbout.Visible = False
CommandData.Visible = False
CommandInternet.Visible = False
CommandUp.Visible = False
CommandReset.Visible = False
Label2.Visible = False
Label15.Visible = False
Label20.Visible = True
Label16.Visible = True
Label19.Visible = True
Label22.Visible = True
Label18.Visible = True
Text1.Visible = True
Line5.Visible = True
Label17.Visible = True
CommandEdit.Visible = True
CommandClear.Visible = True
CommandOK.Visible = True
Label13.Visible = True
End Sub

'██████████启动命令██████████
'延时启动内容
Private Sub Timer4_Timer()
Label2.Caption = "Microsoft Windows XP"
If PingIP("220.181.6.19") = True Then
Label3.Caption = "网络正常! 可访问互联网"
Label3.ForeColor = &HFFFF00
Else
Label3.Caption = "异常! 貌似学校把网断了"
Label3.ForeColor = &H80000012
End If
Timer4.Enabled = False

Dim date1 As Date
date1 = Date
If #3/21/2013# - date1 < 0 Then
Label16.Caption = "000"
Else: Label16.Caption = Format(#3/21/2013# - date1, "000")
End If
Label22.Caption = Format(#6/7/2014# - date1, "000")

If Dir("D:\MyTools\gg.txt") = "" Then
Text1.Text = "公告文件“D:\MyTools\gg.txt”不存在，重新编辑以创建！"
Else
Open "D:\MyTools\gg.txt" For Input As #1
On Error Resume Next
Input #1, StringA
Close #1
Text1.Text = StringA
End If
End Sub

'███↓↓↓████优盘检测████↓↓↓███
Private Sub Timer2_Timer()
'U盘检测
Dim u As String
u = Dir("F:\")
If u = "" Then                                     '如果U盘还没插
CommandU.Enabled = False
cmdUnLoad.Enabled = False
LabelU.Caption = "未插入U盘 (F:\)"
Else
  CommandU.Enabled = True
  cmdUnLoad.Enabled = True
  USBMode.Show
  LabelU.Caption = "检测到U盘! 点击打开"
  LabelU.ForeColor = &HFF&
  Timer3.Enabled = True
  Timer2.Enabled = False
End If
End Sub

Private Sub LabelU_Click()
If LabelU.Caption = "检测到U盘! 点击打开" Then
  If USBMode.Visible = True Then
  Shell "C:\WINDOWS\explorer.exe F:\", 1
  Else
  USBMode.Show
  End If
End If
End Sub

Private Sub Timer3_Timer()
'U盘拔出检测
Dim u As String
u = Dir("F:\")
If u = "" Then                                     '如果U盘没了
  CommandU.Enabled = False
  cmdUnLoad.Enabled = False
  Unload USBMode
  LabelU.Caption = "未插入U盘 (F:\)"
  LabelU.ForeColor = &H0&
  Timer2.Enabled = True
  Timer3.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
'U盘闪动
If LabelU.Caption = "正在检测 (F:\) ..." Then
  If LabelU.ForeColor = &H80000012 Then
    LabelU.ForeColor = &HFF&
  Else: LabelU.ForeColor = &H80000012
  End If
End If
If LabelU.Caption = "检测到U盘! 点击打开" Then
  If LabelU.ForeColor = &H80000012 Then
    LabelU.ForeColor = &HFF00&
  Else: LabelU.ForeColor = &H80000012
  End If
End If
If LabelU.Caption = "未插入U盘 (F:\)" Then
LabelU.ForeColor = &H80000012
End If
'网络闪动
If Label3.Caption = "网络正常! 可访问互联网" Then
  If LabelU.Caption = "检测到U盘! 点击打开" Or LabelU.Caption = "正在检测 (F:\) ..." Then
    If LabelU.ForeColor = &HFF00& Or LabelU.ForeColor = &HFF& Then
      Label3.ForeColor = &HFFFF00
    Else: Label3.ForeColor = &H80000012
    End If
  Else
    If Label3.ForeColor = &H80000012 Then
      Label3.ForeColor = &HFFFF00
    Else: Label3.ForeColor = &H80000012
    End If
  End If
End If
If Label3.Caption = "异常! 貌似学校把网断了" Then
Label3.ForeColor = &H80000012
End If
'时间计数
tm = Format(Time(), "hh:mm:ss")
sc = sc + 1
If sc = 60 Then
sc = 0
mn = mn + 1
End If
Label1.Caption = "=★当前" & tm & " 已运行" & Format(mn, "000") & "分" & Format(sc, "00") & "秒★="
End Sub

'███↑↑↑████优盘检测████↑↑↑███

'██████████驱动重启██████████
Private Sub CommandData_Click()
Label2.Caption = "手动自检中,白板驱动将会重启 ..."
Shell "C:\Program Files\HiteBoard\HiteBoard\Driver\driver.exe"     '★电子板驱动程序
Timer4.Enabled = True
End Sub

'██████████安全移除██████████
'█本段代码及mod开头的两个模块作者【http://blog.csdn.net/chenhui530/】
'█已作精简修改
Private Sub cmdUnLoad_Click()
Dim blnIsUsb As Boolean, strPath As String
strPath = "F:\"                                                       '★移除哪个盘
blnIsUsb = True
    If CloseLockFileHandle(Left(strPath, 2), GetCurrentProcessId) Then
        If blnIsUsb Then
            If RemoveUsbDrive("\\.\" & Left(strPath, 2), True) Then
                cmdUnLoad.Enabled = False
                LabelU.Caption = "☆ 已安全弹出 ★ !!!"
                LabelU.ForeColor = &HC000&
            Else
                MsgBox "安全移除磁盘失败" & vbCrLf & "~~~~(>_<)~~~~" & vbCrLf & vbCrLf & "请先关闭U盘中已打开的文件" & vbCrLf & "或检查F盘是否为你的U盘", vbCritical, "出错了。。。"
            End If
        End If
    Else
        MsgBox "解除占用失败" & vbCrLf & "请先关闭U盘中已打开的文件 !!!" & vbCrLf & "o(︶︿︶)o" & vbCrLf & "或者直接拔了吧。", vbCritical, "注意"
    End If
End Sub

'███↓↓↓████检测联网████↓↓↓███
Private Sub CommandInternet_Click()
Label3.Caption = "手动检测网络 ..."
Label3.ForeColor = &H80000012
Timer5.Enabled = True
End Sub

'【PingIP代码】及【Molude1模块】由百度提供，作者未知
'“220.181.6.19”为百度主服务器IP
Private Sub Timer5_Timer()
If PingIP("220.181.6.19") = True Then
Label3.Caption = "网络正常! 可访问互联网"
Else
Label3.Caption = "异常! 貌似学校把网断了"
End If
Timer5.Enabled = False
End Sub
'███↑↑↑████检测联网████↑↑↑███

'██████████日志按钮██████████
Private Sub CommandUp_Click()
Update.Show
End Sub

'██████████隐藏按钮██████████
Private Sub CommandExit_Click()
If Class.Visible = True Then
  Class.Command1.Caption = "显示助手"
  ZZBBT.Hide
Else
  Unload Class
  Unload USBMode
  Unload Update
  End
End If
End Sub

'███↓↓↓████关机重启████↓↓↓███
Private Sub CommandOFF_Click()
'关机！
Shutdn.Show 1
End Sub

Private Sub CommandReset_Click()
'重启！
Restart.Show 1
End Sub
'███↑↑↑████关机重启████↑↑↑███

'████↓↓↓████16功能键████↓↓↓████

'███↓↓↓███写字白板███↓↓↓███
Private Sub CommandBoard_Click()
'①启动写字白板
CommandBoard.Enabled = False
CommandBoard.Caption = "...请稍候..."
Shell "C:\Program Files\HiteBoard\HiteBoard\Environment.exe", 1
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
'白板“请稍候”恢复程序
'模拟鼠标开始
SetCursorPos 1021, 328
Sleep 100
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Sleep 100
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Sleep 500

SetCursorPos 950, 546
Sleep 100
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Sleep 100
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'模拟鼠标结束
CommandBoard.Enabled = True
CommandBoard.Caption = "写字白板"
Timer1.Enabled = False
End Sub
'███↑↑↑███写字白板███↑↑↑███

Private Sub CommandPPT1_Click()
'②打开课件
Dim Filename1 As String
On Error GoTo ErrHandler
CommonDialog1.Filter = "PPT2003文件 (*.ppt)|*.ppt|PPT2007文件 (*.pptx)|*.pptx|全部文件 (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
Filename1 = CommonDialog1.FileName
Shell "C:\Program Files\Microsoft Office\Office12\POWERPNT.EXE" & " /O " & """" & Filename1 & """", vbMaximizedFocus
ErrHandler:
End Sub

Private Sub CommandPPT2_Click()
'③放映课件
Dim Filename2 As String
On Error GoTo ErrHandler
CommonDialog2.Filter = "PPT2003文件 (*.ppt)|*.ppt|PPT2007文件 (*.pptx)|*.pptx|全部文件 (*.*)|*.*"
CommonDialog2.FilterIndex = 1
CommonDialog2.ShowOpen
Filename2 = CommonDialog2.FileName
Shell "C:\Program Files\Microsoft Office\Office12\POWERPNT.EXE" & " /S " & """" & Filename2 & """", 1
'我的PPT软件在“D:\Program Files\Microsoft Office 2010\Office14\POWERPNT.EXE”
'班班通上的PPT软件在“C:\Program Files\Microsoft Office\Office12\POWERPNT.EXE”
'学校新安装的“D:\EasyTeach\Microsoft Office 2012\Office14\POWERPNT.EXE”
ErrHandler:
End Sub

Private Sub CommandBBT_Click()
'④打开班班通客户端
Shell "C:\Program Files\PstimWebClient\vcomie.exe", 1   '★设置班班通的（绝对）路径
End Sub

Private Sub CommandD_Click()
'⑤打开D盘
Shell "C:\WINDOWS\explorer.exe D:\", 1
End Sub

Private Sub CommandU_Click()
'⑥打开优盘
Shell "C:\WINDOWS\explorer.exe F:\", 1                       '★设置U盘盘符
End Sub

Private Sub CommandKey_Click()
'⑦屏幕键盘
Shell "C:\WINDOWS\system32\osk.exe", 1
End Sub

Private Sub CommandHand_Click()
'⑧手写输入法
Shell "hand.exe", 1
End Sub

Private Sub CommandExplorer_Click()
'⑨访问百度首页
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.baidu.com/", 1
Shell "hand.exe", 1
End Sub

Private Sub CommandMusic_Click()
'⑩打开音乐
Shell "C:\WINDOWS\explorer.exe D:\新建文件夹\", 1        '★设置音乐文件夹的路径
End Sub

Private Sub CommandPlayer_Click()
'⑾Qvod播放器
Shell "QvodPlayer\QvodPlayer.exe", 1
End Sub

Private Sub CommandVideo_Click()
'⑿打开迅雷
Shell "Thunder 7\Program\Thunder.exe", 1           '★设置迅雷7【相对】路径
End Sub

Private Sub Command360_Click()
'360任务管理
Shell "360taskmgr\360taskmgr.exe", 1
End Sub

Private Sub CommandNB_Click()
'NB自动关机
Shell "nb关机-v6.0\NBClose.exe", 1
End Sub

Private Sub CommandAutoShut_Click()
'自制自动关机
Shell "AutoShut.exe", 1
End Sub

Private Sub CommandAbout_Click()
'关于
MsgBox ("VB编程处女作 BY iMaeGoo_虹墨" & vbCrLf & "马越  纯个人制作 (累死了)" & vbCrLf & "班班通助手 和 自动关机小工具 为纯个人原创开源软件" & vbCrLf & "转载请注明" & vbCrLf & "作者微博：http://weibo.com/t1st" & vbCrLf & "Email：mail1st@qq.com" & vbCrLf & "个人小站努力建设中。。。。" & vbCrLf & "version 1.5　Build 20130308")
End Sub
'████↑↑↑████16功能键████↑↑↑████

'██████████今日课程██████████
Private Sub Label2_Click()
If Label2.Caption = "课程表已隐藏，点击恢复" Then
Class.Show
Label2.Caption = "Microsoft Windows XP"
Label2.ForeColor = &H80000012
End If
If Class.Visible = True Then
ZZBBT.Label2.Caption = "课程表已隐藏，点击恢复"
ZZBBT.Label2.ForeColor = &HFF&
Class.Hide
End If
End Sub

Private Sub Timer7_Timer()
daojs.Caption = Format(daojs.Caption - 1, "00")
If daojs.Caption = -1 Then
 Commandthank.Visible = False
 Label14.Caption = "请等待，白板启动可能较慢"
 Shell "C:\Program Files\HiteBoard\HiteBoard\Environment.exe", 1
 daojs.Caption = 13
 Timer8.Enabled = True
 Timer7.Enabled = False
End If
End Sub

Private Sub Timer8_Timer()
daojs.Caption = daojs.Caption - 1
If daojs.Caption = 12 Then
Timer1.Enabled = True
End If
If daojs.Caption = 0 Then
Picture2.Visible = False
Timer8.Enabled = False
End If
End Sub

Private Sub Commandthank_Click()
Timer7.Enabled = False
Picture2.Visible = False
End Sub

Private Sub TimerInternet_Timer()
If PingIP("220.181.6.19") = True Then
Label3.Caption = "网络正常! 可访问互联网"
Else
Label3.Caption = "异常! 貌似学校把网断了"
End If
End Sub

'公告

Private Sub CommandClear_Click()
Text1.Text = ""
End Sub

Private Sub CommandEdit_Click()
Shell "hand.exe", 1
Timer9.Enabled = False
Label17.Caption = "编辑模式"
Text1.Locked = False
CommandOK.Enabled = True
CommandClear.Enabled = True
CommandEdit.Enabled = False
End Sub

Private Sub CommandOK_Click()
If Text1.Text = "" Then
StringA = "暂无公告或您没有输入任何文字"
Else: StringA = Text1.Text
End If
Open "D:\MyTools\gg.txt" For Output As #1
Write #1, StringA
Close #1
Label17.Caption = "Designed ＆UI ＆Powered by iMaeGoo_虹墨"
Text1.Locked = True
Timer9.Enabled = True
CommandEdit.Enabled = True
CommandClear.Enabled = False
Shell "taskkill /f /im hand.exe"
CommandOK.Enabled = False
End Sub

Private Sub Timer9_Timer()
If Dir("D:\MyTools\gg.txt") = "" Then
Text1.Text = "公告文件“D:\MyTools\gg.txt”不存在，重新编辑以创建！"
Else
Open "D:\MyTools\gg.txt" For Input As #1
On Error Resume Next
Input #1, StringA
Close #1
Text1.Text = StringA
End If
End Sub

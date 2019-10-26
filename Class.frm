VERSION 5.00
Begin VB.Form Class 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "4班 今明课程"
   ClientHeight    =   9225
   ClientLeft      =   12855
   ClientTop       =   1785
   ClientWidth     =   1815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   0
      Picture         =   "Class.frx":0000
      ScaleHeight     =   9195
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.Timer Timer1 
         Interval        =   10000
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "隐藏助手"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8760
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "晚自习"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1575
         Left            =   240
         TabIndex        =   12
         Top             =   7080
         Width           =   1455
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   27
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   26
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -15
            TabIndex        =   14
            Top             =   240
            Width           =   1125
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -15
            TabIndex        =   13
            Top             =   840
            Width           =   1125
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "下午"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2775
         Left            =   240
         TabIndex        =   7
         Top             =   4200
         Width           =   1455
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   25
            Top             =   2160
            Width           =   345
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   24
            Top             =   1560
            Width           =   345
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   23
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   22
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -15
            TabIndex        =   11
            Top             =   240
            Width           =   1125
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -15
            TabIndex        =   10
            Top             =   840
            Width           =   1125
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -15
            TabIndex        =   9
            Top             =   1440
            Width           =   1125
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -15
            TabIndex        =   8
            Top             =   2040
            Width           =   1125
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "上午"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   3615
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "今天 明天"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   21
            Top             =   3000
            Width           =   340
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   20
            Top             =   2400
            Width           =   340
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   19
            Top             =   1800
            Width           =   340
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   18
            Top             =   1200
            Width           =   340
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   320
            Left            =   1080
            TabIndex        =   17
            Top             =   600
            Width           =   340
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -30
            TabIndex        =   6
            Top             =   480
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -30
            TabIndex        =   5
            Top             =   1080
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -30
            TabIndex        =   4
            Top             =   1680
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -30
            TabIndex        =   3
            Top             =   2280
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   27.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -30
            TabIndex        =   2
            Top             =   2880
            Width           =   1155
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "今天是 周＿"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim k11 As String, k12 As String, k13 As String, k14 As String, k15 As String, k16 As String, k17 As String, k18 As String, k19 As String, k110 As String, k111 As String, k21 As String, k22 As String, k23 As String, k24 As String, k25 As String, k26 As String, k27 As String, k28 As String, k29 As String, k210 As String, k211 As String, k31 As String, k32 As String, k33 As String, k34 As String, k35 As String, k36 As String, k37 As String, k38 As String, k39 As String, k310 As String, k311 As String, k41 As String, k42 As String, k43 As String, k44 As String, k45 As String, k46 As String, k47 As String, k48 As String, k49 As String, k410 As String, k411 As String, k51 As String, k52 As String, k53 As String, k54 As String, k55 As String, k56 As String, k57 As String, k58 As String, k59 As String, k510 As String, k511 As String, k61 As String, k62 As String, k63 As String, k64 As String, k65 As String, k66 As String, k67 As String, k68 As String
'Dim k69 As String, k610 As String, k611 As String

Private Sub Command1_Click()
If Command1.Caption = "隐藏助手" Then
ZZBBT.Hide
Command1.Caption = "显示助手"
Else
ZZBBT.Show
Command1.Caption = "隐藏助手"
End If
End Sub

Private Sub Form_Load()
Dim w As Integer
w = Weekday(Now, vbMonday)
Select Case (w)
  Case 1
    Label1.Caption = "语文"
    Label2.Caption = "语文"
    Label3.Caption = "自习"
    Label4.Caption = "英语"
    Label5.Caption = "英语"
    
    Label6.Caption = "生物"
    Label7.Caption = "物理"
    Label8.Caption = "数学"
    Label9.Caption = "自习"
    
    Label10.Caption = "物理"
    Label11.Caption = "生物"
    
    Label13.Caption = "今天是 周一"
    
    Label12.Caption = "物理"
    Label14.Caption = "英语"
    Label15.Caption = "数学"
    Label16.Caption = "语文"
    Label17.Caption = "生物"
    
    Label18.Caption = "化学"
    Label19.Caption = "数学"
    Label20.Caption = "通用"
    Label21.Caption = "自习"
    
    Label22.Caption = "生物"
    Label23.Caption = "数学"
    
  Case 2
    Label1.Caption = "物理"
    Label2.Caption = "英语"
    Label3.Caption = "数学"
    Label4.Caption = "语文"
    Label5.Caption = "生物"
    
    Label6.Caption = "化学"
    Label7.Caption = "数学"
    Label8.Caption = "通用"
    Label9.Caption = "自习"
    
    Label10.Caption = "生物"
    Label11.Caption = "数学"
    
    Label13.Caption = "今天是 周二"
    
    Label12.Caption = "数学"
    Label14.Caption = "数学"
    Label15.Caption = "语文"
    Label16.Caption = "化学"
    Label17.Caption = "物理"
    
    Label18.Caption = "英语"
    Label19.Caption = "音乐"
    Label20.Caption = "自习"
    Label21.Caption = "活动"
    
    Label22.Caption = "数学"
    Label23.Caption = "英语"
    
  Case 3
    Label1.Caption = "数学"
    Label2.Caption = "数学"
    Label3.Caption = "语文"
    Label4.Caption = "化学"
    Label5.Caption = "物理"
    
    Label6.Caption = "英语"
    Label7.Caption = "音乐"
    Label8.Caption = "自习"
    Label9.Caption = "活动"
    
    Label10.Caption = "数学"
    Label11.Caption = "英语"
    
    Label13.Caption = "今天是 周三"

    Label12.Caption = "物理"
    Label14.Caption = "生物"
    Label15.Caption = "数学"
    Label16.Caption = "体育"
    Label17.Caption = "化学"
    
    Label18.Caption = "语文"
    Label19.Caption = "英语"
    Label20.Caption = "自习"
    Label21.Caption = "自习"
    
    Label22.Caption = "化学"
    Label23.Caption = "语文"
    
  Case 4
    Label1.Caption = "物理"
    Label2.Caption = "生物"
    Label3.Caption = "数学"
    Label4.Caption = "体育"
    Label5.Caption = "化学"
    
    Label6.Caption = "语文"
    Label7.Caption = "英语"
    Label8.Caption = "自习"
    Label9.Caption = "自习"
    
    Label10.Caption = "化学"
    Label11.Caption = "语文"
    
    Label13.Caption = "今天是 周四"

    Label12.Caption = "化学"
    Label14.Caption = "语文"
    Label15.Caption = "生物"
    Label16.Caption = "英语"
    Label17.Caption = "物理"
    
    Label18.Caption = "体育"
    Label19.Caption = "数学"
    Label20.Caption = "自习"
    
  Case 5
    Label1.Caption = "化学"
    Label2.Caption = "语文"
    Label3.Caption = "生物"
    Label4.Caption = "英语"
    Label5.Caption = "物理"
    
    Label6.Caption = "体育"
    Label7.Caption = "数学"
    Label8.Caption = "自习"
    
    Label13.Caption = "今天是 周五"
    
  Case 7
    Label10.Caption = "物理"
    Label11.Caption = "化学"
    
    Label13.Caption = "今天是 周日"
    
    Label12.Caption = "语文"
    Label14.Caption = "语文"
    Label15.Caption = "自习"
    Label16.Caption = "英语"
    Label17.Caption = "英语"
    
    Label8.Caption = "生物"
    Label9.Caption = "物理"
    Label20.Caption = "数学"
    Label21.Caption = "自习"
    
    Label22.Caption = "物理"
    Label23.Caption = "生物"

  Case Else
    Label13.Caption = "今天是 周六"
End Select

If Label3.Caption = "自习" Then
Label3.ForeColor = &H80000011
End If

If Label4.Caption = "自习" Then
Label4.ForeColor = &H80000011
End If

If Label5.Caption = "自习" Then
Label5.ForeColor = &H80000011
End If

If Label6.Caption = "自习" Then
Label6.ForeColor = &H80000011
End If

If Label7.Caption = "自习" Then
Label7.ForeColor = &H80000011
End If

If Label8.Caption = "自习" Then
Label8.ForeColor = &H80000011
End If

If Label9.Caption = "自习" Then
Label9.ForeColor = &H80000011
End If

Dim t As Date
Dim kc As String
t = Time()
Select Case (t)
  Case #12:00:00 AM# To #7:19:59 AM#
    Frame1.ForeColor = &HFF&
  Case #7:20:00 AM# To #8:19:59 AM#
    Label1.ForeColor = &HFF&
    kc = Label1.Caption
  Case #8:20:00 AM# To #9:10:59 AM#
    Label2.ForeColor = &HFF&
    kc = Label2.Caption
  Case #9:11:00 AM# To #10:00:59 AM#
    Label3.ForeColor = &HFF&
    kc = Label3.Caption
  Case #10:01:00 AM# To #11:10:59 AM#
    Label4.ForeColor = &HFF&
    kc = Label4.Caption
  Case #11:11:00 AM# To #12:00:59 PM#
    Label5.ForeColor = &HFF&
    kc = Label5.Caption
  Case #12:01:00 PM# To #2:09:59 PM#
    Frame2.ForeColor = &HFF&
  Case #2:10:00 PM# To #3:10:59 PM#
    Label6.ForeColor = &HFF&
    kc = Label6.Caption
  Case #3:11:00 PM# To #4:00:59 PM#
    Label7.ForeColor = &HFF&
    kc = Label7.Caption
  Case #4:01:00 PM# To #4:50:59 PM#
    Label8.ForeColor = &HFF&
    kc = Label8.Caption
  Case #4:51:00 PM# To #5:59:59 PM#
    Label9.ForeColor = &HFF&
    kc = Label9.Caption
  Case #6:00:00 PM# To #6:19:59 PM#
    Frame3.ForeColor = &HFF&
  Case #6:20:00 PM# To #7:40:59 PM#
    Label10.ForeColor = &HFF&
    kc = Label10.Caption
  Case #7:41:00 PM# To #8:39:59 PM#
    Label11.ForeColor = &HFF&
    kc = Label11.Caption
  Case #8:40:00 PM# To #11:59:59 PM#
    Label11.ForeColor = &H80000012
  Case Else
    MsgBox "系统时间错误"
End Select
If kc = "数学" Or kc = "物理" Then
ZZBBT.Picture2.Visible = True
ZZBBT.Timer7.Enabled = True
ZZBBT.Label12.Caption = kc
End If
End Sub

Private Sub Timer1_Timer()
'灰色上下午 &H00808080&
'红当前课程 &H000000FF&
'黑其它课程 &H80000012&
t = Time()
Select Case (t)
  Case #12:00:00 AM# To #7:19:59 AM#
    Label11.ForeColor = &H80000012
    Frame1.ForeColor = &HFF&
  Case #7:20:00 AM# To #8:19:59 AM#
    Frame1.ForeColor = &H808080
    Label1.ForeColor = &HFF&
  Case #8:20:00 AM# To #9:10:59 AM#
    Label1.ForeColor = &H80000012
    Label2.ForeColor = &HFF&
  Case #9:11:00 AM# To #10:00:59 AM#
    Label2.ForeColor = &H80000012
    Label3.ForeColor = &HFF&
  Case #10:01:00 AM# To #11:10:59 AM#
    Label3.ForeColor = &H80000012
    Label4.ForeColor = &HFF&
  Case #11:11:00 AM# To #12:00:59 PM#
    Label4.ForeColor = &H80000012
    Label5.ForeColor = &HFF&
  Case #12:01:00 PM# To #2:09:59 PM#
    Label5.ForeColor = &H80000012
    Frame2.ForeColor = &HFF&
  Case #2:10:00 PM# To #3:10:59 PM#
    Frame2.ForeColor = &H808080
    Label6.ForeColor = &HFF&
  Case #3:11:00 PM# To #4:00:59 PM#
    Label6.ForeColor = &H80000012
    Label7.ForeColor = &HFF&
  Case #4:01:00 PM# To #4:50:59 PM#
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &HFF&
  Case #4:51:00 PM# To #5:59:59 PM#
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &HFF&
  Case #6:00:00 PM# To #6:19:59 PM#
    Label9.ForeColor = &H80000012
    Frame3.ForeColor = &HFF&
  Case #6:20:00 PM# To #7:40:59 PM#
    Frame3.ForeColor = &H808080
    Label10.ForeColor = &HFF&
  Case #7:41:00 PM# To #8:39:59 PM#
    Label10.ForeColor = &H80000012
    Label11.ForeColor = &HFF&
  Case #8:40:00 PM# To #11:59:59 PM#
    Label11.ForeColor = &H80000012
  Case Else
    MsgBox "系统时间错误"
End Select
End Sub

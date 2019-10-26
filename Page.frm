VERSION 5.00
Begin VB.Form Page 
   BorderStyle     =   0  'None
   Caption         =   "4班 - 班级公告板"
   ClientHeight    =   2295
   ClientLeft      =   6180
   ClientTop       =   2445
   ClientWidth     =   6495
   Icon            =   "Page.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      Picture         =   "Page.frx":4322
      ScaleHeight     =   2415
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   3840
         Picture         =   "Page.frx":34EBC
         ScaleHeight     =   1725
         ScaleWidth      =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   1815
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "Page.frx":4269E
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "删除“gx.txt”隐藏本提示"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Top             =   1920
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10秒后自动关闭"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   1800
      End
   End
End
Attribute VB_Name = "Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub Form_Load()
t = 10
End Sub

Private Sub Timer1_Timer()
t = t - 1
If t > -1 Then
Label1.Caption = t & "秒后自动关闭"
Else
Unload Me
End If
End Sub

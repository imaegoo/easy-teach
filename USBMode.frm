VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form USBMode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "班班通U盘助手"
   ClientHeight    =   1380
   ClientLeft      =   1470
   ClientTop       =   5685
   ClientWidth     =   4620
   Icon            =   "USBMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   5370
      Left            =   0
      Picture         =   "USBMode.frx":4322
      ScaleHeight     =   5310
      ScaleWidth      =   6510
      TabIndex        =   1
      Top             =   0
      Width           =   6570
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "打开U盘"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "关闭(&C)"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "放映课件"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdUnLoad 
         BackColor       =   &H00C0C0FF&
         Caption         =   "安全移除"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   4080
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "选择U盘中的幻灯片"
         InitDir         =   "F:\"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "!"
         Height          =   255
         Left            =   315
         TabIndex        =   6
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "△ 检测到您插入了U盘，请选择——"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "USBMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUnLoad_Click()
Dim blnIsUsb As Boolean, strPath As String
strPath = "F:\"                                                       '★移除哪个盘
blnIsUsb = True
    If CloseLockFileHandle(Left(strPath, 2), GetCurrentProcessId) Then
        If blnIsUsb Then
            If RemoveUsbDrive("\\.\" & Left(strPath, 2), True) Then
                MsgBox "安全移除U盘成功！", , "提示"
                cmdUnLoad.Enabled = False
            Else
                MsgBox "安全移除磁盘失败" & vbCrLf & "~~~~(>_<)~~~~" & vbCrLf & vbCrLf & "请先关闭U盘中已打开的文件" & vbCrLf & "或检查F盘是否为你的U盘", vbCritical, "出错了。。。"
            End If
        End If
    Else
        MsgBox "请先关闭打开中的U盘文件！！！", vbCritical, "提示"
    End If
End Sub

Private Sub Command1_Click()
Dim Filename2 As String
On Error GoTo ErrHandler
CommonDialog2.Filter = "PPT2003文件 (*.ppt)|*.ppt|PPT2007文件 (*.pptx)|*.pptx|全部文件 (*.*)|*.*"
CommonDialog2.FilterIndex = 1
CommonDialog2.ShowOpen
Filename2 = CommonDialog2.FileName
Shell "C:\Program Files\Microsoft Office\Office12\POWERPNT.EXE" & " /S " & """" & Filename2 & """", 1
Unload Me
ErrHandler:
End Sub

Private Sub Command2_Click()
Shell "C:\WINDOWS\explorer.exe F:\", 1
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

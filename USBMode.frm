VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form USBMode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ͨU������"
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
         Caption         =   "��U��"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ر�(&C)"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "��ӳ�μ�"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdUnLoad 
         BackColor       =   &H00C0C0FF&
         Caption         =   "��ȫ�Ƴ�"
         BeginProperty Font 
            Name            =   "����"
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
         DialogTitle     =   "ѡ��U���еĻõ�Ƭ"
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
         Caption         =   "�� ��⵽��������U�̣���ѡ�񡪡�"
         BeginProperty Font 
            Name            =   "����"
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
strPath = "F:\"                                                       '���Ƴ��ĸ���
blnIsUsb = True
    If CloseLockFileHandle(Left(strPath, 2), GetCurrentProcessId) Then
        If blnIsUsb Then
            If RemoveUsbDrive("\\.\" & Left(strPath, 2), True) Then
                MsgBox "��ȫ�Ƴ�U�̳ɹ���", , "��ʾ"
                cmdUnLoad.Enabled = False
            Else
                MsgBox "��ȫ�Ƴ�����ʧ��" & vbCrLf & "~~~~(>_<)~~~~" & vbCrLf & vbCrLf & "���ȹر�U�����Ѵ򿪵��ļ�" & vbCrLf & "����F���Ƿ�Ϊ���U��", vbCritical, "�����ˡ�����"
            End If
        End If
    Else
        MsgBox "���ȹرմ��е�U���ļ�������", vbCritical, "��ʾ"
    End If
End Sub

Private Sub Command1_Click()
Dim Filename2 As String
On Error GoTo ErrHandler
CommonDialog2.Filter = "PPT2003�ļ� (*.ppt)|*.ppt|PPT2007�ļ� (*.pptx)|*.pptx|ȫ���ļ� (*.*)|*.*"
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

VERSION 5.00
Begin VB.Form Update 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ͨ���� �� ������־"
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
      Caption         =   "������ 2012��12��01�� �汾�� version 1.1.0"
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "- �ָ���Ļ����Ϊϵͳ�Դ�"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "- ����U���Զ�ʶ���ܣ�������"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "- ����������ȥ����Э����ͼ��Ԫ�أ���ť��ɫ����"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���̰� 2012��11��26�� �汾�� version 1.0.0"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "- һ���ػ���������5��ȷ�ϣ���ֹ��"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "- ����360��������Զ��ػ���ԭ��������"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "- �޸����ͨ�ͻ��˴򲻿����⣬����Ӣ��������Ѹ��"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���԰� 2012��11��24�� �汾�� version 0.9"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "- ��U�̣�ֱ�Ӵ��ҵĵ����е�һ��U��(�������)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "- ��д���뷨���ѹ�����֧��"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "- ��ӳ�μ��������õ�Ƭ����ֱ�Ӵӵ�һ�ſ�ʼ��ӳ"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��һ�汾��??�� 2013�� �汾�� version 2.0"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   4695
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "- ��ӭ�����������"
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
         Caption         =   "�����ҵĲ���"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "- ����Զ���������װ岢ȫ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "- ĩ�տ��޶���v1.2 ���ܿγ̱�U��һ������"
         BeginProperty Font 
            Name            =   "����"
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

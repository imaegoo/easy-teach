VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "正在启动..."
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   -120
      MousePointer    =   11  'Hourglass
      Picture         =   "splash.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5700
      TabIndex        =   0
      Top             =   0
      Width           =   5700
      Begin VB.Timer Timer2 
         Interval        =   2000
         Left            =   1440
         Top             =   120
      End
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   960
         Top             =   120
      End
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer_Timer()
ZZBBT.Show
ZZBBT.Hide
Class.Show
Class.Hide
Timer.Enabled = False
End Sub

Private Sub Timer2_Timer()
Class.Show
If Dir("D:\MyTools\gx.txt") <> "" Then
Page.Show
End If
ZZBBT.Show
Unload Me
End Sub

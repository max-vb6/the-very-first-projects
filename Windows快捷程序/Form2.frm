VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "记事本"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "记事本"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "打开记事本"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "保存文本文档"
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
Open Environ("ALL USER SPRO FILE") & "D:\文本文档.txt" For Append As #1
Print #1, Text1.Text
Close #1
Kill "D:\文本文档.txt"
Open Environ("ALL USER SPRO FILE") & "D:\文本文档.txt" For Append As #1
Print #1, Text1.Text
Close #1
MsgBox "已保存文件在D盘！", 64, "提示"
End Sub

Private Sub Command9_Click()
Shell "C:\Windows\System32\notepad.exe"
End Sub

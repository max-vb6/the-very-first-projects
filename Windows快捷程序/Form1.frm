VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows快捷程序"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "打开任务管理器"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "系统程序"
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
      Begin VB.CommandButton Command8 
         Caption         =   "打开命令提示符"
         Height          =   495
         Left            =   2280
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "打开资源管理器"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "打开画图"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "锁定计算机"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "注销"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新启动"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关机"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "系统命令"
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu 系统 
      Caption         =   "系统(&S)"
      Begin VB.Menu 关机 
         Caption         =   "关机"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu 重新启动 
         Caption         =   "重新启动"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu 注销 
         Caption         =   "注销"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu 锁定计算机 
         Caption         =   "锁定计算机"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu 分割线 
         Caption         =   "-"
      End
      Begin VB.Menu 资源管理器 
         Caption         =   "资源管理器"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu 任务管理器 
         Caption         =   "任务管理器"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu 画图 
         Caption         =   "画图"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu 命令提示符 
         Caption         =   "命令提示符"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu 分割线2 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu 程序 
      Caption         =   "程序(&E)"
      Begin VB.Menu 记事本 
         Caption         =   "记事本"
         Shortcut        =   {F1}
      End
      Begin VB.Menu 计算器 
         Caption         =   "计算器"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu 关于 
      Caption         =   "关于(&A)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "C:\Windows\System32\Shutdown.exe -s -t 10"
End Sub

Private Sub Command2_Click()
Shell "C:\Windows\System32\Shutdown.exe -r -t 10"
End Sub

Private Sub Command3_Click()
Shell "C:\Windows\System32\Shutdown.exe -l"
End Sub

Private Sub Command4_Click()
Shell "C:\Windows\System32\rundll32.exe user32.dll,LockWorkStation"
End Sub

Private Sub Command5_Click()
Shell "C:\Windows\System32\taskmgr.exe"
End Sub

Private Sub Command6_Click()
Shell "C:\Windows\explorer.exe"
End Sub

Private Sub Command7_Click()
Shell "C:\Windows\System32\mspaint.exe"
End Sub

Private Sub Command8_Click()
Shell "C:\Windows\System32\cmd.exe"
End Sub

Private Sub 关机_Click()
Shell "C:\Windows\System32\Shutdown.exe -s -t 10"
End Sub

Private Sub 关于_Click()
Form4.Show
End Sub

Private Sub 画图_Click()
Shell "C:\Windows\System32\mspaint.exe"
End Sub

Private Sub 记事本_Click()
Form2.Show
End Sub

Private Sub 计算器_Click()
Form3.Show
End Sub

Private Sub 命令提示符_Click()
Shell "C:\Windows\System32\cmd.exe"
End Sub

Private Sub 任务管理器_Click()
Shell "C:\Windows\System32\taskmgr.exe"
End Sub

Private Sub 锁定计算机_Click()
Shell "C:\Windows\System32\rundll32.exe user32.dll,LockWorkStation"
End Sub

Private Sub 退出_Click()
End
End Sub

Private Sub 重新启动_Click()
Shell "C:\Windows\System32\Shutdown.exe -r -t 10"
End Sub

Private Sub 注销_Click()
Shell "C:\Windows\System32\Shutdown.exe -l"
End Sub

Private Sub 资源管理器_Click()
Shell "C:\Windows\explorer.exe"
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows��ݳ���"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4455
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command5 
      Caption         =   "�����������"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "ϵͳ����"
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
      Begin VB.CommandButton Command8 
         Caption         =   "��������ʾ��"
         Height          =   495
         Left            =   2280
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "����Դ������"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "�򿪻�ͼ"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���������"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ע��"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ػ�"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "ϵͳ����"
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu ϵͳ 
      Caption         =   "ϵͳ(&S)"
      Begin VB.Menu �ػ� 
         Caption         =   "�ػ�"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu ע�� 
         Caption         =   "ע��"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu ��������� 
         Caption         =   "���������"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu �ָ��� 
         Caption         =   "-"
      End
      Begin VB.Menu ��Դ������ 
         Caption         =   "��Դ������"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu ��������� 
         Caption         =   "���������"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu ��ͼ 
         Caption         =   "��ͼ"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu ������ʾ�� 
         Caption         =   "������ʾ��"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu �ָ���2 
         Caption         =   "-"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����(&E)"
      Begin VB.Menu ���±� 
         Caption         =   "���±�"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����(&A)"
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

Private Sub �ػ�_Click()
Shell "C:\Windows\System32\Shutdown.exe -s -t 10"
End Sub

Private Sub ����_Click()
Form4.Show
End Sub

Private Sub ��ͼ_Click()
Shell "C:\Windows\System32\mspaint.exe"
End Sub

Private Sub ���±�_Click()
Form2.Show
End Sub

Private Sub ������_Click()
Form3.Show
End Sub

Private Sub ������ʾ��_Click()
Shell "C:\Windows\System32\cmd.exe"
End Sub

Private Sub ���������_Click()
Shell "C:\Windows\System32\taskmgr.exe"
End Sub

Private Sub ���������_Click()
Shell "C:\Windows\System32\rundll32.exe user32.dll,LockWorkStation"
End Sub

Private Sub �˳�_Click()
End
End Sub

Private Sub ��������_Click()
Shell "C:\Windows\System32\Shutdown.exe -r -t 10"
End Sub

Private Sub ע��_Click()
Shell "C:\Windows\System32\Shutdown.exe -l"
End Sub

Private Sub ��Դ������_Click()
Shell "C:\Windows\explorer.exe"
End Sub

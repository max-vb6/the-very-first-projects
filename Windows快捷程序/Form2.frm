VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���±�"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4455
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame3 
      Caption         =   "���±�"
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
         Caption         =   "�򿪼��±�"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "�����ı��ĵ�"
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
Open Environ("ALL USER SPRO FILE") & "D:\�ı��ĵ�.txt" For Append As #1
Print #1, Text1.Text
Close #1
Kill "D:\�ı��ĵ�.txt"
Open Environ("ALL USER SPRO FILE") & "D:\�ı��ĵ�.txt" For Append As #1
Print #1, Text1.Text
Close #1
MsgBox "�ѱ����ļ���D�̣�", 64, "��ʾ"
End Sub

Private Sub Command9_Click()
Shell "C:\Windows\System32\notepad.exe"
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "无聊的程序  作者：MaxXing"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option2 
      Caption         =   "二进制"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "计算器"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
      Begin VB.OptionButton Option7 
         Caption         =   "梯度"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         Caption         =   "角度"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "弧度"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "十进制"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "÷"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "×"
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "-"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "请输入数字"
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   " 你无聊吗？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do
MsgBox ("你还无聊吗？")
Loop
End Sub

Private Sub Command2_Click()
MsgBox (Text1)
MsgBox ("要加啊？")
Do
MsgBox ("自己算！！")
Loop
End Sub

Private Sub Command3_Click()
MsgBox (Text1)
MsgBox ("要减啊？")
Do
MsgBox ("自己算！！")
Loop
End Sub

Private Sub Command4_Click()
MsgBox (Text1)
MsgBox ("要乘啊？")
Do
MsgBox ("自己算！！")
Loop
End Sub

Private Sub Command5_Click()
MsgBox (Text1)
MsgBox ("要除啊？")
Do
MsgBox ("自己算！！")
Loop
End Sub

Private Sub Text1_Change()
Text1 = Int(Text1)
End Sub

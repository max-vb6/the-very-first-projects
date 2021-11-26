VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "猜数字"
   ClientHeight    =   3630
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":3BCA
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   0
      Picture         =   "Form1.frx":3F71C
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   10
      Top             =   0
      Width           =   4695
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "领取奖励"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "提示"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "放弃"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "猜猜看"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      MaxLength       =   4
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "正确的数字"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "你猜的数字"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "在这里写你猜的数字"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu 游戏 
      Caption         =   "游戏(&G)"
      Index           =   1
      Begin VB.Menu 开始游戏 
         Caption         =   "开始游戏"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu 分割线1 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
         Index           =   3
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助(&A)"
      Index           =   4
      Begin VB.Menu 游戏说明 
         Caption         =   "游戏说明"
         Shortcut        =   ^G
      End
      Begin VB.Menu 关于游戏 
         Caption         =   "关于游戏..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
If Text1.Text = ("") Then
MsgBox "输入的数字不能为空！", 48, "提示"
Else
If Text1.Text = Text2.Text Then
MsgBox "恭喜您！您猜对了！", 64, "提示"
Label1.Caption = ("恭喜您！")
Label2.Visible = True
Label3.Visible = True
Text2.Visible = True
Text1.Enabled = False
Command5.Enabled = True
Command3.Enabled = True
Command4.Enabled = False
Command2.Enabled = False
Else
MsgBox "您猜错了！" + vbCrLf + "提示：所猜数字必须是一个由1、2、3、4随机组成的4位数。" + vbCrLf + "如：2143。", 64, "提示"
End If
End If
End Sub

Private Sub Command3_Click()
Dim a As String
a = MsgBox("您确定要放弃并退出吗？", 48 + vbYesNo, "提示")
If a = vbYes Then
End
End If
End Sub

Private Sub Command4_Click()
MsgBox "提示：" + Text3.Text + "413112" + Text2.Text + "243" + Text3.Text + "中有一个是对的。", 64, "提示"
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Form5.Show
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Else
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End If
'分割线
Randomize Timer
Dim a(1 To 4) As Integer '数组M
Dim b(1 To 4) As String '数组N
Dim i As Integer, k As Integer, t As Integer
For i = 1 To 4
a(i) = i
Next
For i = 1 To 4 '数组打乱
t = a(i)
k = Fix(Rnd * 4) + 1
a(i) = a(k)
a(k) = t
Next
For i = 1 To 4 '从M中随机取出N个数,不重复
b(i) = a(i)
Next
Text3.Text = Join(b(), "")
End Sub

Private Sub Timer1_Timer()
Unload Form5
Form1.Visible = True
Timer1.Enabled = False
End Sub

Private Sub 关于游戏_Click()
frmAbout.Show
End Sub

Private Sub 开始游戏_Click(Index As Integer)
Picture1.Visible = False
Label1.Caption = ("在这里写你猜的数字")
Label2.Visible = False
Label3.Visible = False
Text2.Visible = False
Text1.Enabled = True
Command5.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = True
'分割线
Text2.Text = ("")
Text1.Text = ("")
Randomize Timer
Dim a(1 To 4) As Integer '数组M
Dim b(1 To 4) As String '数组N
Dim i As Integer, k As Integer, t As Integer
For i = 1 To 4
a(i) = i
Next
For i = 1 To 4 '数组打乱
t = a(i)
k = Fix(Rnd * 4) + 1
a(i) = a(k)
a(k) = t
Next
For i = 1 To 4 '从M中随机取出N个数,不重复
b(i) = a(i)
Next
Text2.Text = Join(b(), "")
End Sub

Private Sub 退出_Click(Index As Integer)
End
End Sub

Private Sub 游戏说明_Click()
MsgBox "游戏说明：" + vbCrLf + "玩家必须猜出有游戏随机生成的一个由1~4排列组成的数组，否则游戏失败。", 64, "游戏说明"
End Sub

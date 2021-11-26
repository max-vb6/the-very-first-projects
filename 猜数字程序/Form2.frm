VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "领取您的奖励"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":3BCA
   ScaleHeight     =   2040
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "隐藏明文"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确认"
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查看明文"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "什么是验证特征码？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入验证特征码"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Text1.PasswordChar = ""
Form2.Command3.Visible = True
Form2.Command1.Visible = False
End Sub

Private Sub Command2_Click()
If Text1.Text = Form1.Text3.Text + "413112" + Form1.Text2.Text + "243" + Form1.Text3.Text Then
Unload Me
MsgBox "成功……", 64, "提示"
Form3.Show
Else
MsgBox "输入错误！不正确的验证特征码！", 48, "验证特征码不正确"
End If
End Sub

Private Sub Command3_Click()
Form2.Text1.PasswordChar = "*"
Form2.Command1.Visible = True
Form2.Command3.Visible = False
End Sub

Private Sub Label2_Click()
MsgBox "何为验证特征码？" + vbCrLf + "验证特征码就是由猜数字提示信息中的数字组成的验证码。", 64, "什么是验证特征码？"
End Sub

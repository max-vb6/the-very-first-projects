VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自定义画笔颜色..."
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   4815
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4815
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "预览"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "添加到列表(&A)"
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清空(&C)"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "颜色设置"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3015
      Begin VB.OptionButton Option3 
         Caption         =   "使用选择颜色"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "使用普通颜色"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "使用RGB颜色"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "选择(&S)"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "4210816"
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "255"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "255"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "选择颜色："
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
      TabIndex        =   9
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "设置普通颜色："
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
      TabIndex        =   5
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "设置RGB颜色："
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
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err2
If Option1.Value = True Then
Form1.ForeColor = RGB(Text1.Text, Text2.Text, Text3.Text)
End If
If Option2.Value = True Then
Form1.ForeColor = Text4.Text
End If
If Option3.Value = True Then
Form1.ForeColor = Combo1.Text
End If
Unload Me
Exit Sub
err2:
MsgBox "错误！请填写颜色！", 48, "错误"
Text1.Text = ("0")
Text2.Text = ("255")
Text3.Text = ("255")
Text4.Text = ("4210816")
Combo1.Text = ("")
End Sub

Private Sub Command2_Click()
On Error GoTo err
CommonDialog1.ShowColor
Combo1.Text = CommonDialog1.Color
Exit Sub
err:
End Sub

Private Sub Command3_Click()
Combo1.Clear
End Sub

Private Sub Command4_Click()
Combo1.AddItem (Combo1.Text)
End Sub

Private Sub Command5_Click()
On Error GoTo err
If Option1.Value = True Then
Me.BackColor = RGB(Text1.Text, Text2.Text, Text3.Text)
End If
If Option2.Value = True Then
Me.BackColor = Text4.Text
End If
If Option3.Value = True Then
Me.BackColor = Combo1.Text
End If
Exit Sub
err:
End Sub


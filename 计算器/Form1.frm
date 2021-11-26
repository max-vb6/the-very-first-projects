VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "计算器  作者：MaxXing"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9000
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "体积计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   8775
      Begin VB.CommandButton Command13 
         Caption         =   "求圆锥体积"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   27
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command12 
         Caption         =   "求圆柱体积"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   26
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton Command11 
         Caption         =   "求正方体体积"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command10 
         Caption         =   "求长方体体积"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   24
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "请输入长方体高或圆柱高或圆锥高"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "请输入长方体宽或圆柱、圆锥底面半径"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "请输入长方体长或正方体棱长"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "圆计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4560
      TabIndex        =   11
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command9 
         Caption         =   "求圆直径"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "求圆面积"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "求圆周长"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "请输入圆的半径"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command6 
         Caption         =   "立方"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "平方"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "÷"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "求立方值以这里的数字为准"
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "求平方值以这里的数字为准"
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "请输入第二个数字"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "请输入第一个数字"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox (Text1 + Text2 * 2 / 2)
End Sub

Private Sub Command10_Click()
MsgBox (Text3 * Text5 * Text6)
End Sub

Private Sub Command11_Click()
MsgBox (Text3 ^ 3)
End Sub

Private Sub Command12_Click()
MsgBox (Text5 ^ 2 * 3.14 * Text6)
End Sub

Private Sub Command13_Click()
MsgBox (Text5 ^ 2 * 3.14 * Text6 / 3)
End Sub

Private Sub Command2_Click()
MsgBox (Text1 - Text2)
End Sub

Private Sub Command3_Click()
MsgBox (Text1 * Text2)
End Sub

Private Sub Command4_Click()
MsgBox (Text1 / Text2)
End Sub

Private Sub Command5_Click()
MsgBox (Text1 ^ 2)
End Sub

Private Sub Command6_Click()
MsgBox (Text2 ^ 3)
End Sub

Private Sub Command7_Click()
MsgBox (Text4 * 2 * 3.14)
End Sub

Private Sub Command8_Click()
MsgBox (Text4 ^ 2 * 3.14)
End Sub

Private Sub Command9_Click()
MsgBox (Text4 * 2)
End Sub

Private Sub Text1_Change()
Text1 = Int(Text1)
End Sub

Private Sub Text2_Change()
Text2 = Int(Text2)
End Sub

Private Sub Text3_Change()
Text3 = Int(Text3)
End Sub

Private Sub Text4_Change()
Text4 = Int(Text4)
End Sub

Private Sub Text5_Change()
Text5 = Int(Text5)
End Sub

Private Sub Text6_Change()
Text6 = Int(Text6)
End Sub

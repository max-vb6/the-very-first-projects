VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计算器"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4485
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4485
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Caption         =   "数字键2"
      Height          =   615
      Left            =   360
      TabIndex        =   20
      Top             =   3000
      Width           =   3735
      Begin VB.CommandButton Command26 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command25 
         Caption         =   "2"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command24 
         Caption         =   "3"
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command23 
         Caption         =   "4"
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command22 
         Caption         =   "5"
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command21 
         Caption         =   "6"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command20 
         Caption         =   "7"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command19 
         Caption         =   "8"
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command18 
         Caption         =   "9"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command15 
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数字键1"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   3735
      Begin VB.CommandButton Command10 
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "9"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "7"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "计算器"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "+"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "×"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         Caption         =   "÷"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command16 
         Caption         =   "清零"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton Command17 
         Caption         =   "打开计算器"
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   1560
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = (Text2.Text + "1")
End Sub

Private Sub Command10_Click()
Text2.Text = (Text2.Text + "0")
End Sub

Private Sub Command11_Click()
MsgBox Text2.Text + Text3.Text * 1 / 1, 64, "计算结果"
End Sub

Private Sub Command12_Click()
MsgBox Text2.Text - Text3.Text, 64, "计算结果"
End Sub

Private Sub Command13_Click()
MsgBox Text2.Text * Text3.Text, 64, "计算结果"
End Sub

Private Sub Command14_Click()
MsgBox Text2.Text / Text3.Text, 64, "计算结果"
End Sub

Private Sub Command15_Click()
Text3.Text = (Text3.Text + "0")
End Sub

Private Sub Command16_Click()
Text2.Text = ("0")
Text3.Text = ("0")
End Sub

Private Sub Command17_Click()
Shell "C:\Windows\System32\calc.exe"
End Sub

Private Sub Command18_Click()
Text3.Text = (Text3.Text + "9")
End Sub

Private Sub Command19_Click()
Text3.Text = (Text3.Text + "8")
End Sub

Private Sub Command2_Click()
Text2.Text = (Text2.Text + "2")
End Sub

Private Sub Command20_Click()
Text3.Text = (Text3.Text + "7")
End Sub

Private Sub Command21_Click()
Text3.Text = (Text3.Text + "6")
End Sub

Private Sub Command22_Click()
Text3.Text = (Text3.Text + "5")
End Sub

Private Sub Command23_Click()
Text3.Text = (Text3.Text + "4")
End Sub

Private Sub Command24_Click()
Text3.Text = (Text3.Text + "3")
End Sub

Private Sub Command25_Click()
Text3.Text = (Text3.Text + "2")
End Sub

Private Sub Command26_Click()
Text3.Text = (Text3.Text + "1")
End Sub

Private Sub Command3_Click()
Text2.Text = (Text2.Text + "3")
End Sub

Private Sub Command4_Click()
Text2.Text = (Text2.Text + "4")
End Sub

Private Sub Command5_Click()
Text2.Text = (Text2.Text + "5")
End Sub

Private Sub Command6_Click()
Text2.Text = (Text2.Text + "6")
End Sub

Private Sub Command7_Click()
Text2.Text = (Text2.Text + "7")
End Sub

Private Sub Command8_Click()
Text2.Text = (Text2.Text + "8")
End Sub

Private Sub Command9_Click()
Text2.Text = (Text2.Text + "9")
End Sub

Private Sub Text2_Change()
Text2 = Int(Text2)
If Text2.Text > 1.11111111111111E+15 Then
MsgBox "溢出！", 48, "计算错误"
Text2.Text = ("0")
End If
End Sub

Private Sub Text3_Change()
Text3 = Int(Text3)
If Text3.Text > 1.11111111111111E+15 Then
MsgBox "溢出！", 48, "计算错误"
Text3.Text = ("0")
End If
End Sub

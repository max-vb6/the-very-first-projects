VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBS脚本编辑器"
   ClientHeight    =   2205
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   5055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5055
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":30BA
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu 文件 
      Caption         =   "文件(&F)"
      Begin VB.Menu 保存 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu 删除 
         Caption         =   "删除"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu 清除所有 
         Caption         =   "清除所有"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu 分割线 
         Caption         =   "-"
      End
      Begin VB.Menu 测试 
         Caption         =   "测试"
         Shortcut        =   {F1}
      End
      Begin VB.Menu 设置 
         Caption         =   "设置(&S)"
      End
      Begin VB.Menu 分割线1 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出(&E)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu 插入 
      Caption         =   "插入(&I)"
      Begin VB.Menu Dim 
         Caption         =   "Dim"
      End
      Begin VB.Menu 分割线2 
         Caption         =   "-"
      End
      Begin VB.Menu IfThen 
         Caption         =   "If Then"
      End
      Begin VB.Menu Else 
         Caption         =   "Else"
      End
      Begin VB.Menu EndIf 
         Caption         =   "End If"
      End
      Begin VB.Menu 分割线3 
         Caption         =   "-"
      End
      Begin VB.Menu Do 
         Caption         =   "Do"
      End
      Begin VB.Menu Loop 
         Caption         =   "Loop"
      End
      Begin VB.Menu ExitDo 
         Caption         =   "Exit Do"
      End
      Begin VB.Menu 分割线4 
         Caption         =   "-"
      End
      Begin VB.Menu WSShell 
         Caption         =   "WS Shell"
      End
      Begin VB.Menu WSPopup 
         Caption         =   "WS Popup"
      End
      Begin VB.Menu WSSleep 
         Caption         =   "WS Sleep"
      End
      Begin VB.Menu SAPISpVoice 
         Caption         =   "SAPI SpVoice"
      End
      Begin VB.Menu 分割线5 
         Caption         =   "-"
      End
      Begin VB.Menu Msgbox 
         Caption         =   "Msgbox"
      End
      Begin VB.Menu Inputbox 
         Caption         =   "Inputbox"
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

Private Sub Dim_Click()
Text1.Text = (Text1 + "Dim")
End Sub

Private Sub Do_Click()
Text1.Text = (Text1 + "Do")
End Sub

Private Sub Else_Click()
Text1.Text = (Text1 + "Else")
End Sub

Private Sub EndIf_Click()
Text1.Text = (Text1 + "End If")
End Sub

Private Sub ExitDo_Click()
Text1.Text = (Text1 + "Exit do")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open Environ("ALL USER SPRO FILE") & "D:\NewVBS.vbs" For Append As #1
Print #1, Text1.Text
Close #1
Kill "D:\NewVBS.vbs"
Open Environ("ALL USER SPRO FILE") & "D:\NewVBS.vbs" For Append As #1
Print #1, Text1.Text
Close #1
End
End Sub

Private Sub IfThen_Click()
Text1.Text = (Text1 + "If Then")
End Sub

Private Sub Inputbox_Click()
Text1.Text = (Text1 + "Inputbox")
End Sub

Private Sub Loop_Click()
Text1.Text = (Text1 + "Loop")
End Sub

Private Sub Msgbox_Click()
Text1.Text = (Text1 + "Msgbox")
End Sub

Private Sub SAPISpVoice_Click()
Text1.Text = (Text1 + "CreateObject(" + "chr(34)" + "&" + "SAPI." + "SpVoice" + "&" + "chr(34)" + ")." + "Speak")
End Sub

Private Sub WSPopup_Click()
Text1.Text = (Text1 + "WS.popup")
End Sub

Private Sub WSShell_Click()
Text1.Text = (Text1 + "set ws=createobjecet" + "(" + "chr(34)" + "&" + "wscript.Shell" + "&" + "chr(34)" + ")")
End Sub

Private Sub WSSleep_Click()
Text1.Text = (Text1 + "WS.sleep")
End Sub

Private Sub 保存_Click()
Open Environ("ALL USER SPRO FILE") & "D:\NewVBS.vbs" For Append As #1
Print #1, Text1.Text
Close #1
Kill "D:\NewVBS.vbs"
Open Environ("ALL USER SPRO FILE") & "D:\NewVBS.vbs" For Append As #1
Print #1, Text1.Text
Close #1
End Sub

Private Sub 测试_Click()
Shell "C:\Windows\System32\wscript.exe D:\NewVBS.vbs"
End Sub

Private Sub 关于_Click()
Form2.Show
End Sub

Private Sub 清除所有_Click()
Text1.Text = ("")
End Sub

Private Sub 删除_Click()
Open Environ("ALL USER SPRO FILE") & "D:\NewVBS.vbs" For Append As #1
Print #1, Text1.Text
Close #1
Kill "D:\NewVBS.vbs"
End Sub

Private Sub 设置_Click()
Shell "C:\Windows\System32\wscript.exe"
End Sub

Private Sub 退出_Click()
End
End Sub

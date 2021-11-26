VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "文本编辑器 - 无标题"
   ClientHeight    =   6540
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   9030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9030
   StartUpPosition =   2  '屏幕中心
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "新建"
            Description     =   "新建"
            Object.ToolTipText     =   "新建"
            Object.Tag             =   "新建"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "打开"
            Description     =   "打开"
            Object.ToolTipText     =   "打开"
            Object.Tag             =   "打开"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "保存"
            Description     =   "保存"
            Object.ToolTipText     =   "保存"
            Object.Tag             =   "保存"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "剪切"
            Description     =   "剪切"
            Object.ToolTipText     =   "剪切"
            Object.Tag             =   "剪切"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "复制"
            Description     =   "复制"
            Object.ToolTipText     =   "复制"
            Object.Tag             =   "复制"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "粘贴"
            Description     =   "粘贴"
            Object.ToolTipText     =   "粘贴"
            Object.Tag             =   "粘贴"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "普通"
            Description     =   "普通"
            Object.ToolTipText     =   "普通模式"
            Object.Tag             =   "普通"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "高级"
            Description     =   "高级"
            Object.ToolTipText     =   "高级模式"
            Object.Tag             =   "高级"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "字体"
            Description     =   "字体"
            Object.ToolTipText     =   "字体"
            Object.Tag             =   "字体"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "退出"
            Description     =   "退出"
            Object.ToolTipText     =   "退出"
            Object.Tag             =   "退出"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "关于"
            Description     =   "关于"
            Object.ToolTipText     =   "关于"
            Object.Tag             =   "关于"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "帮助"
            Description     =   "帮助"
            Object.ToolTipText     =   "帮助"
            Object.Tag             =   "帮助"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog5 
      Left            =   5280
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "宋体"
      FontSize        =   30
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   6510
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   5280
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5280
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "保存文件"
      FileName        =   "*.txt"
      Filter          =   $"Form1.frx":25EA
      FontName        =   "宋体"
      InitDir         =   "C:\桌面"
      Max             =   1000
   End
   Begin MSComDlg.CommonDialog CommonDialog4 
      Left            =   5280
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "打开文件..."
      FileName        =   "*.txt"
      Filter          =   $"Form1.frx":2682
      InitDir         =   "C:\桌面"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4935
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":271A
   End
   Begin VB.PictureBox Picture2 
      Height          =   4935
      Left            =   5880
      ScaleHeight     =   4875
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   720
      Width           =   2415
      Begin VB.Frame Frame3 
         Caption         =   "操作"
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
         Begin VB.CommandButton Command9 
            Caption         =   "复制文本"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton Command10 
            Caption         =   "粘贴文本"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "颜色"
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2175
         Begin VB.CommandButton Command1 
            Caption         =   "字体颜色"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton Command2 
            Caption         =   "背景颜色"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "锁定文本"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "字体"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton Command8 
            Caption         =   "自定义..."
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":27BC
      Top             =   720
      Width           =   5775
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6135
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "2010-11-16"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9878
            MinWidth        =   9878
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "PM 04:47"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":27C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2ADD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2DF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3611
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":392B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4145
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":445F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4C79
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5493
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5CAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5FC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":67E1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu 文件 
      Caption         =   "文件(&F)"
      Begin VB.Menu 新建文件 
         Caption         =   "新建文件(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu 打开文件 
         Caption         =   "打开文件(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu 保存文件 
         Caption         =   "保存文件(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu 操作 
      Caption         =   "操作(&O)"
      Begin VB.Menu 复制 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu 剪切 
         Caption         =   "剪切"
      End
      Begin VB.Menu 粘贴 
         Caption         =   "粘贴(&V)"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu 查找 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu 编辑字体 
         Caption         =   "编辑字体..."
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu 普通编辑模式 
         Caption         =   "普通编辑模式"
         Checked         =   -1  'True
      End
      Begin VB.Menu 高级编辑模式 
         Caption         =   "高级编辑模式"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助(&H)"
      Begin VB.Menu 显示帮助文档 
         Caption         =   "显示帮助文档(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu 关于 
         Caption         =   "关于(&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check5_Click()
If Check5.Value = 1 Then
Text1.Locked = True
RichTextBox1.Locked = True
End If
If Check5.Value = 0 Then
Text1.Locked = False
RichTextBox1.Locked = False
End If
End Sub

Private Sub Command1_Click()
On Error GoTo err7
CommonDialog2.ShowColor
Text1.ForeColor = (CommonDialog2.Color)
Exit Sub
err7:
End Sub

Private Sub Command10_Click()
RichTextBox1.SelText = Clipboard.GetText
End Sub

Private Sub Command2_Click()
On Error GoTo err8
CommonDialog3.ShowColor
Text1.BackColor = (CommonDialog3.Color)
RichTextBox1.BackColor = (CommonDialog3.Color)
Exit Sub
err8:
End Sub

Private Sub Command3_Click()
On Error GoTo err
Text1.FontSize = Text2.Text
Exit Sub
err:
End Sub

Private Sub Command5_Click()
Text1.Text = (vbCrLf + Text1.Text)
RichTextBox1.Text = (vbCrLf + RichTextBox1.Text)
End Sub

Private Sub Command6_Click()
SendKeys "{BACKSPACE}"
End Sub

Private Sub Command7_Click()
Text1.Text = (" " + Text1.Text)
RichTextBox1.Text = (vbCrLf + RichTextBox1.Text)
End Sub

Private Sub Command8_Click()
CommonDialog5.Flags = cdlCFEffects Or cdlCFBoth
CommonDialog5.ShowFont
Text1.Font.Name = CommonDialog5.FontName
Text1.Font.Size = CommonDialog5.FontSize
Text1.Font.Bold = CommonDialog5.FontBold
Text1.Font.Italic = CommonDialog5.FontItalic
Text1.Font.Underline = CommonDialog5.FontUnderline
Text1.FontStrikethru = CommonDialog5.FontStrikethru
RichTextBox1.Font.Name = CommonDialog5.FontName
RichTextBox1.Font.Size = CommonDialog5.FontSize
RichTextBox1.Font.Bold = CommonDialog5.FontBold
RichTextBox1.Font.Italic = CommonDialog5.FontItalic
RichTextBox1.Font.Underline = CommonDialog5.FontUnderline
RichTextBox1.Font.Strikethrough = CommonDialog5.FontStrikethru
End Sub

Private Sub Command9_Click()
If Text1.Visible = False Then
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End If
If Text1.Visible = True Then
Clipboard.Clear
Clipboard.SetText Text1.SelText
End If
End Sub

Private Sub Form_Resize()
On Error GoTo err10
StatusBar1.Panels(2).MinWidth = (Form1.Width - 2880)
If Text1.Visible = True Then
Text1.Height = (Form1.Height - 2280)
Text1.Width = (Form1.Width - 2600)
Picture2.Left = (Text1.Width + 50)
Picture2.Height = (Form1.Height - 2280)
End If
If Text1.Visible = False Then
RichTextBox1.Height = (Form1.Height - 2280)
RichTextBox1.Width = (Form1.Width - 2600)
Picture2.Left = (RichTextBox1.Width + 50)
Picture2.Height = (Form1.Height - 2280)
End If
Exit Sub
err10:
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu 文件, vbPopupMenuLeftAlign
Else
Exit Sub
End If
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu 操作, vbPopupMenuLeftAlign
Else
Exit Sub
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "新建"
新建文件_Click
Case "打开"
打开文件_Click
Case "保存"
保存文件_Click
Case "剪切"
剪切_Click
Case "复制"
复制_Click
Case "粘贴"
粘贴_Click
Case "退出"
退出_Click
Case "普通"
普通编辑模式_Click
Case "高级"
高级编辑模式_Click
Case "关于"
关于_Click
Case "字体"
编辑字体_Click
Case "帮助"
显示帮助文档_Click
End Select
End Sub

Private Sub 保存文件_Click()
On Error GoTo err6
CommonDialog1.ShowSave
If Text1.Visible = True Then
Open Environ("ALL USER SPRO FILE") & (CommonDialog1.FileName) For Append As #1
Print #1, Text1.Text
Close #1
End If
If Text1.Visible = False Then
RichTextBox1.SaveFile CommonDialog1.FileName, rtfRTF
End If
Me.Caption = ("文本编辑器 - " + CommonDialog1.FileName)
RichTextBox1.Locked = False
Text1.Locked = False
StatusBar1.Panels(2).Text = "目前状态：成功保存文件 - " + CommonDialog1.FileName
Exit Sub
err6:
End Sub

Private Sub 编辑字体_Click()
Command8_Click
End Sub

Private Sub 查找_Click()
Dim sFind As String
sFind = InputBox("请输入要查找的字、词：", "查找内容", sFind)
RichTextBox1.Find sFind
End Sub

Private Sub 打开文件_Click()
On Error GoTo err9
CommonDialog4.ShowOpen
RichTextBox1.LoadFile (CommonDialog4.FileName)
Text1.Text = RichTextBox1.Text
Me.Caption = ("文本编辑器 - " + CommonDialog4.FileName)
RichTextBox1.Locked = False
Text1.Locked = False
StatusBar1.Panels(2).Text = "目前状态：编辑文件 - " + CommonDialog4.FileName
Exit Sub
err9:
End Sub

Private Sub 复制_Click()
If Text1.Visible = False Then
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
End If
If Text1.Visible = True Then
Clipboard.Clear
Clipboard.SetText Text1.SelText
End If
StatusBar1.Panels(2).Text = "目前状态：已复制文本"
End Sub

Private Sub 高级编辑模式_Click()
高级编辑模式.Checked = True
普通编辑模式.Checked = False
RichTextBox1.Visible = True
Text1.Visible = False
Command1.Enabled = False
Command10.Enabled = True
粘贴.Enabled = True
End Sub

Private Sub 关于_Click()
frmAbout.Show
End Sub

Private Sub 剪切_Click()
If Text1.Visible = False Then
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""
End If
If Text1.Visible = True Then
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SelText = ""
End If
StatusBar1.Panels(2).Text = "目前状态：已剪切文本"
End Sub

Private Sub 普通编辑模式_Click()
高级编辑模式.Checked = False
普通编辑模式.Checked = True
RichTextBox1.Visible = False
Text1.Visible = True
Command1.Enabled = True
Command10.Enabled = False
粘贴.Enabled = False
End Sub

Private Sub 退出_Click()
End
End Sub

Private Sub 显示帮助文档_Click()
On Error GoTo err11
高级编辑模式_Click
RichTextBox1.LoadFile "Help.rtf"
RichTextBox1.Locked = True
Me.Height = (Me.Height - 1 + 1)
Me.Caption = ("文本编辑器 - " + "帮助文档")
RichTextBox1.BackColor = (&H80000005)
StatusBar1.Panels(2).Text = "目前状态：正在阅读帮助文档"
Exit Sub
err11:
If Error = 53 Then
MsgBox "帮助文档不存在！请联系作者！", 16, "错误"
End If
RichTextBox1.Height = (Form1.Height - 2280)
RichTextBox1.Width = (Form1.Width - 2600)
End Sub

Private Sub 新建文件_Click()
Text1.Text = ("")
RichTextBox1.Text = ("")
Me.Caption = ("文本编辑器 - " + "无标题")
RichTextBox1.Locked = False
Text1.Locked = False
StatusBar1.Panels(2).Text = "目前状态：新建文件成功"
End Sub

Private Sub 粘贴_Click()
If Text1.Visible = False Then
RichTextBox1.SelText = Clipboard.GetText
StatusBar1.Panels(2).Text = "目前状态：已粘贴文本"
End If
End Sub

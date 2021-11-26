VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "点西瓜"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8520
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   7440
      Picture         =   "Form1.frx":0000
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "游戏说明"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "点击开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7440
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2550
      Left            =   2400
      Picture         =   "Form1.frx":351A
      Top             =   120
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   1560
      Picture         =   "Form1.frx":10F34
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   7320
      X2              =   8520
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6495
      Left            =   7320
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If b Then
If KeyCode = 67 Then
Image2.Top = 360
PlaySoundResource 101
End If
Else
Beep
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If b Then
If KeyCode = 67 Then
Image2.Top = 120
c = c + 1
Label6.Caption = c
XiGua
End If
End If
End Sub

Private Sub Form_Load()
b = False
MsgBox "说明：" + vbCrLf + "在10秒内以你最快的速度连续按下“C”键" + vbCrLf + "即可获胜！", 64, "游戏说明"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 1
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackStyle = 0
End
End Sub

Private Sub Label4_Click()
Timer1.Enabled = True
Label2.Caption = 10
Timer2.Enabled = True
Image3.Visible = True
Image1.Picture = Form2.Image1(8).Picture
c = 0
End Sub

Private Sub Label5_Click()
MsgBox "说明：" + vbCrLf + "在10秒内以你最快的速度连续按下“C”键" + vbCrLf + "即可获胜！", 64, "游戏说明"
End Sub

Private Sub Timer1_Timer()
b = True
Label4.Caption = "游戏开始"
Label2.Caption = Int(Label2.Caption)
Label2.Caption = Label2.Caption - 1
If Label2.Caption = 0 Then
Label4.Caption = ("时间到" + vbCrLf + "点击重来")
Image2.Top = 120
Timer2.Enabled = False
Image3.Visible = False
Timer1.Enabled = False
b = False
End If
End Sub
Private Sub icoChange()
Ico = Ico + 1
If Ico = 4 Then
Ico = 0
End If
Image3.Picture = Form2.Image3(Ico).Picture
End Sub

Private Sub Timer2_Timer()
icoChange
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͼ��"
   ClientHeight    =   6540
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   6135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   6540
   ScaleWidth      =   6135
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   4245
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   855
         Left            =   5160
         MousePointer    =   1  'Arrow
         Picture         =   "Form1.frx":46D2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "������ɫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   120
         Width           =   4935
         Begin VB.CommandButton Command2 
            BackColor       =   &H00000000&
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FF0000&
            Height          =   615
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H000000FF&
            Height          =   615
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H0000FF00&
            Height          =   615
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   3720
            MaxLength       =   3
            MousePointer    =   3  'I-Beam
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   270
            Left            =   4080
            MaxLength       =   3
            MousePointer    =   3  'I-Beam
            TabIndex        =   13
            Text            =   "255"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Left            =   4440
            MaxLength       =   3
            MousePointer    =   3  'I-Beam
            TabIndex        =   12
            Text            =   "255"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command13 
            Caption         =   "�Զ���"
            Height          =   255
            Left            =   3720
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "���ʴ�С"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   1200
         Width           =   4935
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00808080&
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00404040&
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   3120
            MousePointer    =   3  'I-Beam
            TabIndex        =   4
            Text            =   "25"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command12 
            Caption         =   "�Զ���"
            Height          =   255
            Left            =   3120
            TabIndex        =   3
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command14 
         Caption         =   "��Ƥ��"
         Height          =   855
         Left            =   5160
         MousePointer    =   1  'Arrow
         Picture         =   "Form1.frx":481C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.jpg"
      DialogTitle     =   "��ͼƬ�ļ�"
      FileName        =   "*.jpg"
      Filter          =   $"Form1.frx":5036
      InitDir         =   "C:\����"
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.jpg"
      DialogTitle     =   "����ͼƬ�ļ�"
      FileName        =   "*.jpg"
      Filter          =   $"Form1.frx":50FF
      InitDir         =   "C:\����"
   End
   Begin VB.Menu �ļ� 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu ��ͼƬ 
         Caption         =   "��ͼƬ(&O)"
      End
      Begin VB.Menu �ر�ͼƬ 
         Caption         =   "�ر�ͼƬ(&C)"
      End
      Begin VB.Menu ����ͼƬ 
         Caption         =   "����ͼƬ(&S)"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����(&O)"
      Begin VB.Menu ������ɫ 
         Caption         =   "������ɫ"
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^B
         End
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^W
         End
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^U
         End
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^R
         End
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^G
         End
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^Y
         End
         Begin VB.Menu �� 
            Caption         =   "��"
            Shortcut        =   ^N
         End
         Begin VB.Menu l1 
            Caption         =   "-"
         End
         Begin VB.Menu �Զ���2 
            Caption         =   "�Զ���..."
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu ���ʴ�С 
         Caption         =   "���ʴ�С"
         Begin VB.Menu h1 
            Caption         =   "1"
            Shortcut        =   {F1}
         End
         Begin VB.Menu h5 
            Caption         =   "5"
            Shortcut        =   {F2}
         End
         Begin VB.Menu h10 
            Caption         =   "10"
            Shortcut        =   {F3}
         End
         Begin VB.Menu h15 
            Caption         =   "15"
            Shortcut        =   {F4}
         End
         Begin VB.Menu h20 
            Caption         =   "20"
            Shortcut        =   {F5}
         End
         Begin VB.Menu l2 
            Caption         =   "-"
         End
         Begin VB.Menu �Զ��� 
            Caption         =   "�Զ���..."
            Shortcut        =   {F6}
         End
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu ʹ�û��� 
         Caption         =   "ʹ�û���"
         Checked         =   -1  'True
      End
      Begin VB.Menu ʹ��ֱ�� 
         Caption         =   "ʹ��ֱ��"
      End
      Begin VB.Menu �ı䱳��ɫ 
         Caption         =   "�ı䱳��ɫ"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu ���ظ����� 
         Caption         =   "���ع�����"
         Shortcut        =   ^H
      End
      Begin VB.Menu ��ʾ������ 
         Caption         =   "��ʾ������"
         Shortcut        =   ^S
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu ��Ƥ�� 
         Caption         =   "��Ƥ��"
         Shortcut        =   ^E
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����(&A)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim paintnow As Boolean
Dim linenow As Boolean
Dim lineon As Boolean
Dim painton As Boolean
Dim lx As Long
Dim ly As Long
Private Sub Command1_Click()
Cls
End Sub

Private Sub Command10_Click()
DrawWidth = 15
End Sub

Private Sub Command11_Click()
DrawWidth = 20
End Sub

Private Sub Command12_Click()
On Error GoTo err
DrawWidth = Text1.Text
Exit Sub
err:
MsgBox "����", 48, "����"
Text1.Text = ("10")
End Sub

Private Sub Command13_Click()
On Error GoTo err2
ForeColor = RGB(Text2.Text, Text3.Text, Text4.Text)
Exit Sub
err2:
MsgBox "��������дRGB��ɫ��", 48, "����"
Text2.Text = ("0")
Text3.Text = ("255")
Text4.Text = ("255")
End Sub

Private Sub Command14_Click()
ForeColor = RGB(240, 240, 240)
End Sub

Private Sub Command2_Click()
ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Command3_Click()
ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Command4_Click()
ForeColor = RGB(0, 0, 255)
End Sub

Private Sub Command5_Click()
ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command6_Click()
ForeColor = RGB(0, 255, 0)
End Sub

Private Sub Command7_Click()
DrawWidth = 1
End Sub

Private Sub Command8_Click()
DrawWidth = 5
End Sub

Private Sub Command9_Click()
DrawWidth = 10
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub Form_Load()
DrawWidth = 10
ForeColor = RGB(0, 0, 0)
painton = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If painton Then
paintnow = True
ElseIf lineon Then
lx = X
ly = Y
linenow = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If paintnow Then
PSet (X, Y)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If painton Then
paintnow = False
ElseIf lineon Then
If linenow Then
Line (lx, ly)-(X, Y)
End If
linenow = False
End If
End Sub

Private Sub h1_Click()
DrawWidth = 1
End Sub

Private Sub h10_Click()
DrawWidth = 10
End Sub

Private Sub h15_Click()
DrawWidth = 15
End Sub

Private Sub h20_Click()
DrawWidth = 20
End Sub

Private Sub h5_Click()
DrawWidth = 5
End Sub

Private Sub ��_Click()
ForeColor = RGB(255, 255, 255)
End Sub

Private Sub ����ͼƬ_Click()
On Error Resume Next
CommonDialog2.ShowSave
  If CommonDialog2.FileName <> "" Then
     SavePicture Form1.Image, CommonDialog2.FileName
  End If
End Sub

Private Sub ��ͼƬ_Click()
On Error GoTo err3
CommonDialog1.ShowOpen
Form1.Picture = LoadPicture(CommonDialog1.FileName)
Exit Sub
err3:
End Sub

Private Sub �ı䱳��ɫ_Click()
Me.BackColor = Me.ForeColor
End Sub

Private Sub �ر�ͼƬ_Click()
Me.Picture = LoadPicture("")
End Sub

Private Sub ����_Click()
frmAbout.Show
End Sub

Private Sub ��_Click()
ForeColor = RGB(0, 0, 0)
End Sub

Private Sub ��_Click()
ForeColor = RGB(255, 0, 0)
End Sub

Private Sub ��_Click()
ForeColor = RGB(255, 255, 0)
End Sub

Private Sub ��_Click()
ForeColor = RGB(0, 0, 255)
End Sub

Private Sub ��_Click()
ForeColor = RGB(0, 255, 0)
End Sub

Private Sub ��_Click()
ForeColor = RGB(0, 255, 255)
End Sub

Private Sub ����_Click()
Cls
End Sub

Private Sub ʹ�û���_Click()
painton = True
lineon = False
ʹ�û���.Checked = True
ʹ��ֱ��.Checked = False
End Sub

Private Sub ʹ��ֱ��_Click()
painton = False
lineon = True
ʹ�û���.Checked = False
ʹ��ֱ��.Checked = True
End Sub

Private Sub �˳�_Click()
End
End Sub

Private Sub ��ʾ������_Click()
Picture1.Visible = True
End Sub

Private Sub ��Ƥ��_Click()
ForeColor = &H8000000F
ʹ�û���_Click
End Sub

Private Sub ���ظ�����_Click()
Picture1.Visible = False
End Sub

Private Sub �Զ���_Click()
Form2.Show
End Sub

Private Sub �Զ���2_Click()
Form3.Show
End Sub

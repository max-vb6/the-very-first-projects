VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ı��༭�� - �Զ�������"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4560
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�����������ƣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
Form1.Text1.Font.Name = Combo1.Text
Form1.RichTextBox1.Font = Combo1.Text
Combo1.AddItem Combo1.Text
Unload Me
Exit Sub
err:
MsgBox "��Ч���壡", 16, "����"
End Sub

Private Sub Command2_Click()

End Sub

Private Sub File1_Click()

End Sub


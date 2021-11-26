VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "¿ì½ÝËÑË÷"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10455
   FillColor       =   &H80000004&
   ForeColor       =   &H80000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   10455
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   5505
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SogouËÑË÷(&S)"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GoogleËÑË÷(&G)"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   5040
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "ÊäÈëËÑË÷ÄÚÈÝ..."
      Top             =   4560
      Width           =   10455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "°Ù¶ÈÒ»ÏÂ(&B)"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WebBrowser1.Navigate "http://www.baidu.com/s?wd=" + Text1.Text
End Sub

Private Sub Command2_Click()
WebBrowser1.Navigate "http://www.google.com.hk/search?q=" + Text1.Text
End Sub

Private Sub Command3_Click()
WebBrowser1.Navigate "http://www.sogou.com/sohu?query=" + Text1.Text
End Sub

Private Sub Command4_Click()
WebBrowser1.Navigate "http://cn.bing.com/search?q=" + Text1.Text + "&src=IE-SearchBox&FORM=IE8SRC"
End Sub

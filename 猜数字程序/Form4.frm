VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于猜数字"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   0
      Picture         =   "Form4.frx":3BCA
      ScaleHeight     =   3675
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C)2010 MaxXSoft.All rights reserved."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "程序版权："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

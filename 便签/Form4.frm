VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3060
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   4560
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu ·µ»Ø 
         Caption         =   "·µ»Ø"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu ÍË³ö 
         Caption         =   "ÍË³ö"
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Main"
      Begin VB.Menu Òþ²Ø 
         Caption         =   "Òþ²Ø"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu MainEnd 
         Caption         =   "ÍË³ö"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MainEnd_Click()
End
End Sub

Private Sub ·µ»Ø_Click()
Unload Form3
Form1.Show
End Sub

Private Sub ÍË³ö_Click()
End
End Sub

Private Sub Òþ²Ø_Click()
Form3.Text1.Text = Form1.Text1.Text
Unload Form1
Form3.Show
End Sub

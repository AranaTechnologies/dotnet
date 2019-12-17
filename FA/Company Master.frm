VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00004000&
   Caption         =   "Company Master"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form14"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu companymnu 
      Caption         =   "Company"
      Begin VB.Menu ccmnu 
         Caption         =   "Create Company"
      End
      Begin VB.Menu bcmnu 
         Caption         =   "Browse Company"
      End
      Begin VB.Menu clcmu 
         Caption         =   "Company Login"
      End
   End
   Begin VB.Menu exitmnu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcmnu_Click()
Form1.Show
End Sub

Private Sub ccmnu_Click()
Form2.Show
End Sub

Private Sub clcmu_Click()
Form12.Show
End Sub

Private Sub exitmnu_Click()
End
End Sub


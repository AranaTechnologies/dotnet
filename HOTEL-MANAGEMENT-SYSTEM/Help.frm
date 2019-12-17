VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Help 
   BackColor       =   &H00FFFF80&
   Caption         =   "Help"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   6105
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   6360
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   9551
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
      Location        =   ""
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Web.Navigate "C:\Bca\Hotel M.System\HELP\CUSTOMER.HTML"
End Sub

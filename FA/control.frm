VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00004000&
   Caption         =   "Control Window"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1140
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1575
      Left            =   2880
      TabIndex        =   1
      Top             =   3000
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Menu inventorymnu 
      Caption         =   "Inventory"
      Begin VB.Menu csgmnu 
         Caption         =   "Create Stock Group"
      End
      Begin VB.Menu bsgmnu 
         Caption         =   "Browse Stock Group"
      End
      Begin VB.Menu csimnu 
         Caption         =   "Create Stock Item"
      End
      Begin VB.Menu bsimnu 
         Caption         =   "Browse Stock Item"
      End
   End
   Begin VB.Menu journalmnu 
      Caption         =   "Journal"
      Begin VB.Menu jourmnu 
         Caption         =   "journal entry form"
      End
      Begin VB.Menu joubrwmnu 
         Caption         =   "journal browser"
      End
   End
   Begin VB.Menu accountmnu 
      Caption         =   "Account"
      Begin VB.Menu clgmnu 
         Caption         =   "Create Ledger Group"
      End
      Begin VB.Menu blgmnu 
         Caption         =   "Browse Ledger Group"
      End
      Begin VB.Menu clmnu 
         Caption         =   "Create Ledger"
      End
      Begin VB.Menu blmnu 
         Caption         =   "Browse Ledger"
      End
   End
   Begin VB.Menu repomnu 
      Caption         =   "Report Section"
   End
   Begin VB.Menu pwmnu 
      Caption         =   "Password"
   End
   Begin VB.Menu helpmnu 
      Caption         =   "Help"
   End
   Begin VB.Menu aboutmnu 
      Caption         =   "About"
      Begin VB.Menu sysmnu 
         Caption         =   "About the System"
      End
      Begin VB.Menu devmnu 
         Caption         =   "About the Developer"
      End
   End
   Begin VB.Menu exitmnu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub blgmnu_Click()
Form5.Show
End Sub

Private Sub blmnu_Click()
Form6.Show
End Sub

Private Sub bsgmnu_Click()
Form10.Show
End Sub

Private Sub bsimnu_Click()
Form9.Show
End Sub

Private Sub clgmnu_Click()
Form3.Show
End Sub

Private Sub clmnu_Click()
Form4.Show
End Sub

Private Sub csgmnu_Click()
Form8.Show
End Sub

Private Sub csimnu_Click()
Form7.Show
End Sub

Private Sub devmnu_Click()
Form13.Show
End Sub

Private Sub exitmnu_Click()
Unload Me
End Sub

Private Sub joubrwmnu_Click()
Form25.Show
End Sub

Private Sub jourmnu_Click()
Form16.Show
End Sub

Private Sub pwmnu_Click()
cpassfrm.Show
End Sub

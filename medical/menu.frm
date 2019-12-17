VERSION 5.00
Begin VB.Form menu 
   BackColor       =   &H00FFFF80&
   Caption         =   "Patient Directory"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Control Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3720
      TabIndex        =   0
      Top             =   6720
      Width           =   8175
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Search"
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "By Date"
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "By Name"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Patient Registration"
         Height          =   615
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Line Line5 
         X1              =   4320
         X2              =   4320
         Y1              =   840
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   6360
         X2              =   6360
         Y1              =   1920
         Y2              =   2280
      End
      Begin VB.Line Line3 
         X1              =   2520
         X2              =   2520
         Y1              =   1920
         Y2              =   2280
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   6360
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         X1              =   4320
         X2              =   4320
         Y1              =   1560
         Y2              =   1920
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Us: scriptonova@yahoo.com"
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   11640
      TabIndex        =   5
      Top             =   10080
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By: ScriptoNova"
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   11640
      TabIndex        =   4
      Top             =   9840
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   3540
      Left            =   5400
      Picture         =   "menu.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   2160
      Picture         =   "menu.frx":B34D
      Top             =   120
      Width           =   12000
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
entrymedical.Show
End Sub

Private Sub Command2_Click()
searchmedical.Show
End Sub

Private Sub Command3_Click()
datesearch.Show
End Sub


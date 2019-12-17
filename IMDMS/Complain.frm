VERSION 5.00
Begin VB.Form Form34 
   BackColor       =   &H0080C0FF&
   Caption         =   "Complain For Item"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form34"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
      Height          =   495
      Left            =   8520
      TabIndex        =   16
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10200
      TabIndex        =   14
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   5415
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   975
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Problem Found"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item_Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Complain Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()

End Sub

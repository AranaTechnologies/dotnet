VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form HTML 
   BackColor       =   &H0080C0FF&
   Caption         =   "HTML Editor"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton header 
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Publish"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton IMG 
      Caption         =   "Img"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton EXIT 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton BR 
      Caption         =   "Br"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton FONT_BOLD 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton FONT 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   7800
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BODY 
      Caption         =   "Body"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton HTML 
      Caption         =   "Html"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"HTML.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Click the buttons below to insert appropriate HTML tags."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BODY_Click()
RTB.SelRTF = "<BODY>" & Chr(10) & "</BODY>"
BODY.Enabled = False
End Sub
Private Sub BR_Click()
RTB.SelRTF = "<br>"
End Sub
Private Sub Command1_Click()
Cd1.Filter = "HTML |*.HTML| Text Only Format |*.txt| All Files |*.*"
Cd1.ShowSave
Text1.Text = Cd1.filename
RTB.SaveFile Cd1.filename, 1
End Sub
Private Sub Command2_Click()
Editor.Show
End Sub
Private Sub Command3_Click()
Form35.Text3.Text = Text1.Text
Form35.Show
End Sub
Private Sub EXIT_Click()
End
End Sub
Private Sub FONT_BOLD_Click()
RTB.SelRTF = "<B>" & Chr(10) & "</B>"
End Sub
Private Sub FONT_Click()
RTB.SelRTF = "<Font>" & Chr(10) & "</Font>"
'RTB.Text = RTB.Text + "<Font>" & Chr(10) + "</Font>"'
End Sub
Private Sub Form_Activate()
RTB.SetFocus
End Sub

Private Sub header_Click()
RTB.SelRTF = "<H1>" & Chr(10) & "</H1>"
End Sub

Private Sub HTML_Click()
RTB.Text = RTB.Text + "<HTML>" & Chr(10) & "</HTML>"
HTML.Enabled = False
End Sub
Private Sub IMG_Click()
CMDB.Filter = "All Files |*.jpg|*.gif"
CMDB.ShowOpen
End Sub

VERSION 5.00
Begin VB.Form Form45 
   BackColor       =   &H0080C0FF&
   Caption         =   "Sensitivity test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form45"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   27
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   26
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   9360
      TabIndex        =   25
      Text            =   " "
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   9360
      TabIndex        =   23
      Text            =   " "
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   9360
      TabIndex        =   21
      Text            =   " "
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   9360
      TabIndex        =   19
      Text            =   " "
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   9360
      TabIndex        =   17
      Text            =   " "
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   9360
      TabIndex        =   15
      Text            =   " "
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Text            =   " "
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Text            =   " "
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Text            =   " "
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Text            =   " "
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Text            =   " "
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Text            =   " "
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Chloromphenical"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   24
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cefixin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   22
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Sephuroxyme"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   20
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tetracyclin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   18
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Roxythromycin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Erythromycin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Penicilin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ampicilin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Zentamycin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cyprofloxacin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amoxycelin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Antibiotic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Entry  for Antibiotic Sensitivity Test"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Department of Microbiology"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "Form45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim fsys As New FileSystemObject

'PRINT ALL BUTTON'
Private Sub Command1_Click()
Dim OUTSTREAM As TextStream
Dim STR As String
Dim PREC As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Mca\Imdms\Reports\" & Trim(RST.Fields("employee_id")) & ".PRN", True, False)
PREC = Space(10) & "CENTRAL  DIAGNOSTIC  &  RESEARCH  CENTRE . "
OUTSTREAM.WriteLine PREC
PREC = Space(15) & "AK POINT,68B Acharya Prafulla Chandra Road, Kolkata - 700 009"
OUTSTREAM.WriteLine PREC
PREC = Space(13) & "Telephone : 2352 - 0114 / 2360-0206"
OUTSTREAM.WriteLine PREC
PREC = Space(15) & "E-mail : Central_Diag@vsnl.net.in"
OUTSTREAM.WriteLine PREC
PREC = Space(7) & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
PREC = Space(21) & "MICROBIOLOGY DEPARTMENT"
OUTSTREAM.WriteLine PREC
PREC = "*********************************************************************************"
OUTSTREAM.WriteLine PREC
Do While RST.EOF = False
PREC = Space(5) & RST.Fields("ANTIBIOTIC")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("AMOXYCELIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("CYPROFLOXACIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("ZENTAMYCIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("AMPICILIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("PENICILIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("ERYTHROMYCIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("ROXYTHROMYCIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("TETRACYCLIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("SEPHUROXYME")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("CEFIXIN")
OUTSTREAM.WriteLine PREC
PREC = Space(5) & RST.Fields("CHLOROMPHENICAL")
OUTSTREAM.WriteLine PREC
RST.MoveNext
Loop

OUTSTREAM.WriteLine
PREC = Space(15) & "*******************"
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
PREC = Space(40) & "Thanking You , "
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
PREC = Space(40) & "      for Central Diagnostic & Research Centre "
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
PREC = "*********************************************************************************"
OUTSTREAM.WriteLine PREC

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from ANTIBIOTIC_SENSITIVITY_TEST", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = RST
End Sub


 

End Sub

VERSION 5.00
Begin VB.Form Form43 
   BackColor       =   &H0080C0FF&
   Caption         =   "HIV Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form43"
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
      Left            =   7200
      TabIndex        =   7
      Top             =   7200
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
      Left            =   3480
      TabIndex        =   6
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Result if Negative"
      Height          =   1215
      Left            =   3240
      TabIndex        =   4
      Top             =   4920
      Width           =   5535
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Text            =   " "
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Result if Positive"
      Height          =   1215
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Text            =   " "
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Entry  for HIV I  and  HIV II"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Department of Serology"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form43"
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
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Mca\Imdms\Reports\" & Trim(RST.Fields("employee_id")) & ".PRN", True, False)
prec = Space(10) & "CENTRAL  DIAGNOSTIC  &  RESEARCH  CENTRE . "
OUTSTREAM.WriteLine prec
prec = Space(15) & "AK POINT,68B Acharya Prafulla Chandra Road, Kolkata - 700 009"
OUTSTREAM.WriteLine prec
prec = Space(13) & "Telephone : 2352 - 0114 / 2360-0206"
OUTSTREAM.WriteLine prec
prec = Space(15) & "E-mail : Central_Diag@vsnl.net.in"
OUTSTREAM.WriteLine prec
prec = Space(7) & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(21) & "SEROLOGY DEPARTMENT"
OUTSTREAM.WriteLine prec
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
Do While RST.EOF = False
prec = Space(5) & RST.Fields("RESULT IF POSITIVE")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("RESULT_IF NEGATIVE")
OUTSTREAM.WriteLine prec
RST.MoveNext
Loop

OUTSTREAM.WriteLine
prec = Space(15) & "*******************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(40) & "Thanking You , "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(40) & "      for Central Diagnostic & Research Centre "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from SEROLOGY_HIV_TEST", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = RST
End Sub


 

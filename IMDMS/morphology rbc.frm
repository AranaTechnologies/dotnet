VERSION 5.00
Begin VB.Form Form47 
   BackColor       =   &H0080C0FF&
   Caption         =   "Severe Annemia"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form47"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Result"
      Height          =   2775
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of RBC per 100 WBC"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Reticulocyte Count"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Entry  for Morphology of Red Blood Cell  (Severe Annemia)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Department of Haematology"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form47"
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
prec = Space(21) & "HAEMATOLOGY DEPARTMENT"
OUTSTREAM.WriteLine prec
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
Do While RST.EOF = False
prec = Space(5) & RST.Fields("RETICULYTE_COUNT ")
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
RST.Open "select * from MORPHOLOGY ", CNN, adOpenStatic, adLockOptimistic, adCmdText
End Sub


 





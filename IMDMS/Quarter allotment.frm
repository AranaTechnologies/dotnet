VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form26 
   BackColor       =   &H0080C0FF&
   Caption         =   "Browse Quarter Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form26"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Print Quarter Surrender"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
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
      Left            =   8760
      TabIndex        =   3
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Quarter Allotment Letter"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   7440
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5415
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9551
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Quarter Information "
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
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
'PRINT QUARTER ALLOTMENT LETTER'
Private Sub Command1_Click()
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Mca\Imdms\Reports\" & Trim(rst.Fields("QUARTER_NO")) & ".PRN", True, False)
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
prec = Space(21) & "SURRNEDER INFORMATION FOR YOUR QUARTER"
OUTSTREAM.WriteLine prec
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
prec = Space(5) & "To,"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("FIRST_NAME") & " " & rst.Fields("LAST_NAME")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("DESIGNATION")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("DATE_OF_BIRTH")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("DATE_OF_JOINING")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "Respected Sir,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "It is hereby intimated that,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "you  have been allotted a Quarter."
OUTSTREAM.WriteLine prec
prec = Space(15) & "Now  you can take over your Quarter."
OUTSTREAM.WriteLine prec
prec = Space(15) & "Your Quarter_No is given below. "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(15) & "    QUARTER_NO : " & rst.Fields("QUARTER_NO")
OUTSTREAM.WriteLine prec
prec = Space(15) & "*******************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(40) & "Thanking You , "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(50) & "     DR. R.K. SARKAR,"
OUTSTREAM.WriteLine
prec = Space(50) & " HEAD OF THE DEPARTMENT, BIOCHEMISTRY "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
End Sub

'EXIT BUTTON'
Private Sub Command2_Click()
rst.Close
cnn.Close
Unload Me
End Sub

Private Sub Command3_Click()
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Mca\Imdms\Reports\" & Trim(rst.Fields("QUARTER_NO")) & ".PRN", True, False)
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
prec = Space(21) & "ALLOTMENT FOR YOUR QUARTER"
OUTSTREAM.WriteLine prec
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
prec = Space(5) & "To,"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("FIRST_NAME") & " " & rst.Fields("LAST_NAME")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("DESIGNATION")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("DATE_OF_BIRTH")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("DATE_OF_JOINING")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "Respected Sir,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "It is hereby intimated that,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "you  have been allotted a Quarter."
OUTSTREAM.WriteLine prec
prec = Space(15) & "Now  you can take over your Quarter."
OUTSTREAM.WriteLine prec
prec = Space(15) & "Your Quarter_No is given below. "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(15) & "    QUARTER_NO : " & rst.Fields("QUARTER_NO")
OUTSTREAM.WriteLine prec
prec = Space(15) & "*******************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(40) & "Thanking You , "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(50) & "     DR. R.K. SARKAR,"
OUTSTREAM.WriteLine
prec = Space(50) & " HEAD OF THE DEPARTMENT, BIOCHEMISTRY "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
End Sub

Private Sub Form_Load()
cnn.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
rst.Open "select * from QUARTER_FILE where QUARTER_NO is NOT null", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst

End Sub

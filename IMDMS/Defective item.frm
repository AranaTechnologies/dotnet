VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form36 
   BackColor       =   &H0080C0FF&
   Caption         =   "Defective Item"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form36"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Repalcement Request"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Supplier"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   7680
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9763
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
      Caption         =   "Defective Item"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
'SEARCH SUPPLIER BUTTON'
Private Sub Command1_Click()
Dim filename As String




 End Sub

'PRINT REPLACEMENT REQUEST BUTTON'
Private Sub Command3_Click()
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Mca\Imdms\Reports\" & Trim(RST.Fields("Patient_id")) & ".PRN", True, False)
prec = Space(10) & "CENTRAL  DIAGNOSTIC  &  RESEARCH  CENTRE . "
OUTSTREAM.WriteLine prec
prec = Space(15) & "AK POINT,68B Acharya Prafulla Chandra Road, Kolkata - 700 009"
OUTSTREAM.WriteLine prec
prec = Space(12) & "Telephone : 2352 - 0114 / 2360-0206"
OUTSTREAM.WriteLine prec
prec = Space(21) & "E-mail : Central_Diag@vsnl.net.in"
OUTSTREAM.WriteLine prec
prec = Space(7) & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(21) & "REPLACEMENT REQUEST"
OUTSTREAM.WriteLine prec
prec = "*********************************************************************************"
 
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("VOUCHER_NO")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("ITEM_ID")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("PROBLEM_FOUND")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("STATUS")
OUTSTREAM.WriteLine prec
prec = Space(15) & "It is hereby intimated that,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "your VOUCHER_NO is given below."
OUTSTREAM.WriteLine prec
prec = Space(15) & "VOUCHER_NO :" & RST.Fields("ITEM_ID")
OUTSTREAM.WriteLine prec
prec = Space(15) & "*******************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(40) & "Thanking You , "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(50) & "      for Central Diagnostic & Research Centre "
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
End Sub

'EXIT BUTTON'
Private Sub Command4_Click()
RST.Close
CNN.Close
Unload Me
End Sub


Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from COMPLAIN_file where VOUCHER_NO is NOT NULL", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = RST
End Sub

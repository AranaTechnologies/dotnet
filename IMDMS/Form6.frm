VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H0080C0FF&
   Caption         =   "  "
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form5"
   ScaleHeight     =   3420
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Money Receipt"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   7560
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5595
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9869
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
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
      BackColor       =   &H0080C0FF&
      Caption         =   "Browse Patient Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim RST1 As New ADODB.Recordset
Dim fsys As New FileSystemObject

 
'PRINT MONEY RECEIPT BUTTON
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
prec = Space(21) & "MONEY RECEIPT"
OUTSTREAM.WriteLine prec
prec = "*********************************************************************************"
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("PATIENT_NAME")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("PATIENT_address1")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("PATIENT_address2")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("PIN")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("DOCTORS_NAME")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("DATE_OF_VISIT")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("DATE_OF_BIRTH")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("GENDER")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("FEES_TO_BE_PAID")
OUTSTREAM.WriteLine prec
prec = Space(5) & RST.Fields("FEES_STATUS")
OUTSTREAM.WriteLine prec
prec = Space(15) & "It is hereby intimated that,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "your PATIENT_ID is given below."
OUTSTREAM.WriteLine prec
prec = Space(15) & "PATIENT_ID :" & RST.Fields("patient_id")
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
RST.Open "select * from patient_information_file where patient_id is NOT NULL", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = RST
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form37 
   BackColor       =   &H0080C0FF&
   Caption         =   "Salary Slip"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form37"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Salary Slip"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
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
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Slip"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "Form37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset

Private Sub Command1_Click()
 
Dim OURSTREAM As TextStream
Dim STR As String
Dim PREC As String
Dim DA
Dim MA
Dim HRA
Dim PF
Dim GROSS
Dim NET
Dim INFACT
Dim ttax
Dim pay_tax
Dim rst3 As New ADODB.Recordset

Set OUTSTREAM = fsys.CreateTextFile("C:\MCA\IMDMS\REPORTS\" & Trim(RST.Fields("EMPLOYEE_NAME")) & ".PRNT", True, False)
OUTSTREAM.WriteLine
Set OUTSTREAM = fsys.CreateTextFile("C:\Mca\Hms\Reports\" & Trim(RST.Fields("EMPLOYEE_ID")) & ".PRN", True, False)
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
PREC = Space(5) & " Pay Slip of - " & RST.Fields("DOCTOR_ID") & " for the Month of :- " & Month(Date) & "," & Year(Date)
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
PREC = Space(5) & " Bill No. - " & Space(5) & "/P , DATED - " & Date
OUTSTREAM.WriteLine PREC
PREC = Space(5) & " SL. NO. of A.Roll " & Space(2) & " ___________ "
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
PREC = Space(5) & "EMPLOYEE Name - " & RST.Fields("EMPLOYEE_NAME") & " ( " & RST.Fields("DESIGNATION") & " )"
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine
INFACT = (Year(Date) - Year(RST.Fields("D_O_J")))
BASIC = (RST.Fields("BASIC") + (RST.Fields("INCRE") * INFACT))
PREC = Space(18) & "Basic - " & BASIC
OUTSTREAM.WriteLine PREC
OUTSTREAM.WriteLine

  DA = (BASIC * RST.Fields("DA") / 100)
    PREC = Space(13) & " DA (Rs.)  - " & Space(1) & DA
    OUTSTREAM.WriteLine PREC
  HRA = (BASIC * RST.Fields("HRA") / 100)
    PREC = Space(12) & " HRA (Rs.)  - " & Space(1) & HRA
    OUTSTREAM.WriteLine PREC
  PREC = Space(13) & " MA (Rs.)  - " & Space(4) & RST.Fields("MA")
    OUTSTREAM.WriteLine PREC
     PREC = Space(16) & "________________"
    OUTSTREAM.WriteLine PREC
  GROSS = (BASIC + DA + HRA + MA)
  MsgBox (GROSS)
  Set rst3 = CNN.Execute("select deduction from tax where " & GROSS & " between lower_limit and upper_limit")
  ttax = rst3.Fields("deduction")
  rst3.Close
  pay_tax = GROSS * (ttax / 100)

  PREC = Space(5) & "Gross Amount (Rs.) - " & GROSS
    OUTSTREAM.WriteLine PREC
  PREC = Space(5) & "Tax Amount (Rs.) - " & pay_tax
    OUTSTREAM.WriteLine PREC
    
    'OUTSTREAM.WriteLine
  PF = (BASIC * RST.Fields("PF") / 100)
    PREC = Space(7) & "(LESS) PF (Rs.) - " & Space(2) & PF
    OUTSTREAM.WriteLine PREC
    'PREC = Space(5) & " Less Co-Op. (Rs.) " & Space(1) & " ____________ "
    'OUTSTREAM.WriteLine PREC
    PREC = Space(16) & "__________________"
    OUTSTREAM.WriteLine PREC
  NET = (GROSS - PF)
     PREC = Space(5) & " Net Amount (Rs.) - " & Space(1) & NET
     OUTSTREAM.WriteLine PREC
     
OUTSTREAM.WriteLine
PREC = Space(32) & " ---------------------------------- "
OUTSTREAM.WriteLine PREC
PREC = Space(30) & " Signature of the Issuing Authority "
OUTSTREAM.WriteLine PREC
'RST1.Close
'CNN1.Close


End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from SALARY_FULES_FILE where DESIGNATION is NOT NULL", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = RST

End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Payreport_Browser 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Payreport"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   6030
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate  Payreport"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   5760
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   12648384
      HeadLines       =   1
      RowHeight       =   22
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
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "PAYREPORT  BROWSER  ENTRY  FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "Payreport_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub Command1_Click()
rst.Close
cnn.Close
Unload Me

End Sub

Private Sub Command2_Click()
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Bca\Hotel M.System\Reports\payreport.PRN", True, False)
rst.MoveFirst
While Not rst.EOF
prec = Space(25) & " PARK CHAIN OF HOTEL "
OUTSTREAM.WriteLine prec
prec = Space(20) & " 14/7 PARK STREET KOLKATA-700002"
OUTSTREAM.WriteLine prec
prec = Space(20) & "Telephone : 033-2282-4666,4667 "
OUTSTREAM.WriteLine prec
prec = Space(20) & "E-mail : TPBL@THEPARKHOTELS.COM"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Date : " & Date & Space(40) & "Time : " & Time()
OUTSTREAM.WriteLine prec

prec = Space(7) & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
OUTSTREAM.WriteLine prec

prec = Space(5) & "Emp Id   :      " & rst.Fields("emp_id")
OUTSTREAM.WriteLine prec
prec = Space(5) & " Name    :      " & rst.Fields("name")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Basic    :      " & rst.Fields("basic")
OUTSTREAM.WriteLine prec
prec = Space(5) & "DA       :      " & rst.Fields("da")
OUTSTREAM.WriteLine prec
prec = Space(5) & "TRA      :      " & rst.Fields("tra")
OUTSTREAM.WriteLine prec
prec = Space(5) & "HRA      :      " & rst.Fields("hra")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Total    :      " & rst.Fields("total")
OUTSTREAM.WriteLine prec


prec = Space(7) & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = "--------------------------------------cut from here----------------------"
OUTSTREAM.WriteLine prec
rst.MoveNext
Wend

prec = Space(65) & "-----------"
OUTSTREAM.WriteLine prec
prec = Space(65) & "(Signature)"
OUTSTREAM.WriteLine prec
rst.MoveFirst
MsgBox ("Printing is over")


End Sub

Private Sub Form_Load()
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=HTLMS;pwd=HTLMS1"
rst.Open " select e.emp_id , e.name , pr.basic, pr.da, pr.tra, pr.hra, pr.total from employee  e , payreport  pr where upper(e.emp_id)=upper(pr.emp_id)", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst

End Sub



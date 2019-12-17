VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Payment_Browser 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Payment"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   6540
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Generate  Money Receipt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   2400
      TabIndex        =   0
      Top             =   2520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4895
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
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   9360
      Picture         =   "Payment1.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1320
      Picture         =   "Payment1.frx":0167
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "   PAYMENT  BROWSER  ENTRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   4935
   End
End
Attribute VB_Name = "Payment_Browser"
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
Set OUTSTREAM = fsys.CreateTextFile("C:\Bca\Hotel M.System\Reports\money_receipt.PRN", True, False)
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

prec = Space(5) & "Customer id   : " & rst.Fields("customer_id")
OUTSTREAM.WriteLine prec
prec = Space(5) & " Name         : " & rst.Fields("name")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Amount        : " & rst.Fields("amount")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Description   : " & rst.Fields("description")
OUTSTREAM.WriteLine prec
prec = Space(5) & "P_Date        : " & rst.Fields("p_date")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Received_By   : " & rst.Fields("received_by")

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
rst.Open " select c.customer_id , c.name , p.amount, p.description, p.p_date, p.received_by  from customer  c , payment  p where upper(c.customer_id)=upper(p.customer_id)", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst

End Sub


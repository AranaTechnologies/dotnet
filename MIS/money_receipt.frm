VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmmoney_receipt 
   BackColor       =   &H00FF8080&
   Caption         =   "MONEY RECEIPT"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdgenerate_money 
      Caption         =   "GENERATE MONEY RECEIPT"
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   4320
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5530
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
End
Attribute VB_Name = "frmmoney_receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdgenerate_money_Click()

 
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\bca\aoes\reports\money_receipt.PRN", True, False)
rst.MoveFirst
While Not rst.EOF

 

prec = Space(35) & " ADVANCE ENTERPRICE "
OUTSTREAM.WriteLine prec
prec = Space(25) & " MADRAL ROAD ,NAIHATI 24 PARGANAS(NORTH) "
OUTSTREAM.WriteLine prec
prec = Space(20) & "Telephone : 2581-0612    MOBILE : 32032199 "
OUTSTREAM.WriteLine prec
prec = Space(21) & "E-mail : ADVANCE@REDIFMAIL.COM"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Date : " & Date & Space(40) & "Time : " & Time()
OUTSTREAM.WriteLine prec

prec = Space(7) & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
OUTSTREAM.WriteLine prec

prec = Space(5) & "Customer id   : " & rst.Fields("customer_id")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Customer name  : " & rst.Fields("customer_name")
OUTSTREAM.WriteLine prec
prec = Space(5) & "order number   : " & rst.Fields("order_no")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Amount          : " & rst.Fields("amount")


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
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis;pwd=mis1"
rst.Open " select c.customer_id , customer_name , o.order_no ,p.amount  from customer  c , payment  p,orders  o where upper(c.customer_id)=upper(o.customer_id) and upper(p.customer_id)= upper(o.customer_id)", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst

End Sub

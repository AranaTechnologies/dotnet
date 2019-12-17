VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ADuty_Browser 
   BackColor       =   &H00FFFFC0&
   Caption         =   "ADuty_Browser"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   6480
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   393216
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
   Begin VB.CommandButton cmdduty_chart 
      Caption         =   "PRINT DUTY CHART"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   6960
      Width           =   975
   End
   Begin VB.ComboBox cmbemployee_id 
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   5280
      Width           =   975
   End
   Begin VB.ComboBox cmb_duty 
      Height          =   315
      ItemData        =   "ADuty_Browser.frx":0000
      Left            =   5160
      List            =   "ADuty_Browser.frx":0010
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txt_date 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmd_search 
      Caption         =   "SEARCH"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txt_search 
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmd_go 
      Caption         =   "GO"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdallote_duty 
      Caption         =   "ALLOTE DUTY"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOTE  DUTY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMPLOYEE ID"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DUTY SIFT"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENTER THE DATE"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   6360
      Width           =   1695
   End
End
Attribute VB_Name = "ADuty_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject



Private Sub cmd_exit_Click()
rst.Close
cnn.Close
Unload Me

End Sub

Private Sub cmd_go_Click()
Dim flag As Integer
flag = 0
rst.MoveFirst
Do While rst.EOF = False
If Trim(UCase(rst.Fields("allote_date"))) = Trim(UCase(txt_search.Text)) Then

flag = 1
Exit Do
End If
rst.MoveNext
Loop

'Label1.Visible = False
'txtsearch.Visible = False
'cmdgo.Visible = False
'cmdsearch.Enabled = True
If flag = 0 Then
MsgBox ("No DATE bears the ALLOTE DATE " & txt_search.Text)
rst.MoveFirst
End If

End Sub

Private Sub cmd_search_Click()
txt_search.Text = ""

txt_search.SetFocus
End Sub

Private Sub cmdallote_duty_Click()
rst.AddNew
rst.Fields("emp_id") = cmbemployee_id.Text
rst.Fields("duty") = cmb_duty.Text
rst.Fields("allote_date") = txt_date.Text

rst.Update
cmb_duty.Text = ""
cmbemployee_id.Text = ""

End Sub

Private Sub cmdduty_chart_Click()
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\Bca\Hotel M.System\Reports\Print duty_chart.PRN", True, False)
rst.MoveNext
rst.MoveFirst
While Not rst.EOF
prec = Space(25) & " PARK CHAIN OF HOTEL"
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

prec = Space(5) & " Employee Id    : " & rst.Fields("emp_id")
OUTSTREAM.WriteLine prec
prec = Space(5) & " Duty           : " & rst.Fields("duty")
OUTSTREAM.WriteLine prec
prec = Space(5) & " Allot Date     : " & rst.Fields("allot_date")

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

rst.Open "select emp_id from employee", cnn, adOpenStatic, adLockOptimistic, adCmdText
While Not rst.EOF
cmbemployee_id.AddItem rst.Fields("emp_id")
rst.MoveNext
Wend
rst.Close

rst.Open "select * from duty_roster ", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
txt_date.Text = Date



End Sub



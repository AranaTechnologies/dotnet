VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form process_request_for_raw_materials 
   BackColor       =   &H00C0FFFF&
   Caption         =   "PROCESS REQUEST FOR RAW MATERIALS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Delivery Date"
      ForeColor       =   &H00004040&
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         Height          =   255
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(dd-mmm-yyyy; e.g.01-mar-2004)"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Delivery Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Print Delivery Note"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   16448
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Submit"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESS REQUEST FOR RAW MATERIALS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT REQUEST NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
End
Attribute VB_Name = "process_request_for_raw_materials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub Command1_Click()
'code of submit button
Dim rst1 As New ADODB.Recordset
rst1.Open "select ri.request_no,ri.dept_name,ri.request_date,i.raw_material_id,i.quantity,r.unit_of_measurement from request_info ri,issue_register i,raw_materials r where upper(ri.request_no)=upper(i.request_no) and upper(i.raw_material_id)=upper(r.raw_material_id) and upper(ri.request_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst1
End Sub

Private Sub Command2_Click()
'code of print delivery note button
rst.Close
rst.Open "select request_no from request_info where status='Pending'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst("request_no")
rst.MoveNext
Wend
Frame1.Visible = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command3_Click()
'code of exit button
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command4_Click()
'code of ok button
'this module will generate Delivery note of raw materials and update stock correespondingly
Command2.Enabled = True
Command3.Enabled = True
Dim sql As String
Dim prev_stock As Long
Dim curr_stock As Long
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Dim sl_no As Integer
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\delivery_note_raw_materials.DOC", True, False)
rst1.Open "select ri.dept_name,ri.request_date,i.raw_material_id,r.description,r.unit_of_measurement,i.quantity from request_info ri,issue_register i,raw_materials r where  upper(i.raw_material_id)=upper(r.raw_material_id) and upper(i.request_no)=upper(ri.request_no) and upper(i.request_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst1.EOF = False
sql = "select * from raw_materials where raw_material_id='" & rst1("raw_material_id") & "'"
Set rst3 = cnn.Execute(sql)
prev_stock = rst3("current_stock")
If prev_stock < rst1("quantity") Then
MsgBox ("The demand of '" & rst1("raw_material_id") & "' exceeds the current stock")
Exit Sub
End If
rst1.MoveNext
Wend
rst1.Close
rst1.Open "select ri.dept_name,ri.request_date,i.raw_material_id,r.description,r.unit_of_measurement,i.quantity from request_info ri,issue_register i,raw_materials r where  upper(i.raw_material_id)=upper(r.raw_material_id) and upper(i.request_no)=upper(ri.request_no) and upper(i.request_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "NICCO CORPORATION LIMITED"
OUTSTREAM.WriteLine prec
prec = Space(10) & "CABLE DIVISION" & Space(35) & "Telephone: (033)581-2131,2132,2133,6234"
OUTSTREAM.WriteLine prec
prec = Space(6) & "SHYAMNAGAR, P.O. ATHPUR" & Space(30) & "Fax No.  :(91)33-581-2940"
OUTSTREAM.WriteLine prec
prec = Space(8) & "North 24-Parganas" & Space(34) & "E-mail   : nclpur@cal3.vsnl.net.in"
OUTSTREAM.WriteLine prec
prec = Space(8) & "W.Bengal-743128"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(48) & "DELIVERY NOTE OF RAW MATERIALS"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "THE FOLLOWING MATERIAL(S) HAS/HAVE BEEN DELIVERED"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "TO:"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst1("dept_name") & "Department"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "ON  " & Date & "  AT  " & Time() & "  AGGAINST REQUEST NO.  " & Combo1.Text & "  DATE  " & rst1.Fields("request_date")
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
sl_no = 1
prec = Space(5) & "SL.NO." & Space(5) & "RAW MATERIAL ID" & Space(5) & "DESCRIPTION" & Space(24) & "QUANTITY" & Space(5) & "UNIT"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
While rst1.EOF = False
sql = "select * from raw_materials where raw_material_id='" & rst1("raw_material_id") & "'"
Set rst3 = cnn.Execute(sql)
prev_stock = rst3("current_stock")
prec = Space(5) & sl_no & Space(10 - Len(sl_no)) & rst1.Fields("raw_material_id") & Space(20 - Len(rst1.Fields("raw_material_id"))) & rst1.Fields("description") & Space(45 - Len(rst1.Fields("description")) - Len(rst1.Fields("quantity"))) & rst1.Fields("quantity") & Space(5) & rst1.Fields("unit_of_measurement")
OUTSTREAM.WriteLine prec
sl_no = sl_no + 1
sql = "update request_info set delivery_date='" & Text1.Text & "',status='Issued'where upper(request_no)='" & UCase(Combo1.Text) & "'"
Set r = cnn.Execute(sql)
curr_stock = prev_stock - rst1("quantity")
sql = "update raw_materials set current_stock=" & curr_stock & "where raw_material_id='" & rst1("raw_material_id") & "'"
Set r = cnn.Execute(sql)
sql1 = "insert into raw_material_stock values('" & Combo1.Text & "','" & rst1("raw_material_id") & "','" & rst1("description") & "','" & rst1("unit_of_measurement") & "'," & rst1("quantity") & ",'ISSUED','" & Text1.Text & "'," & prev_stock & "," & curr_stock & ")"
Set rst2 = cnn.Execute(sql1)
rst1.MoveNext
Wend
OUTSTREAM.WriteLine
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(72) & "(SIGNATURE OF THE STAFF)"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "THE ABOVE MENTIONED MATERIAL HAS BEEN RECEIVED" & Space(21) & "(SIGNATURE OF THE RECEIVER)"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------" & "CUT FROM HERE" & "-----------------------------------------"
OUTSTREAM.WriteLine prec
MsgBox ("Printing is over")
Frame1.Visible = False
End Sub

Private Sub Form_Load()
cnn.Open "DSN=MIS; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select request_no from request_info where status='Pending'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst("request_no")
rst.MoveNext
Wend
End Sub

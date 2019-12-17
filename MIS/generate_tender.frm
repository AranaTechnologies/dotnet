VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form generate_tender 
   BackColor       =   &H00C0FFFF&
   Caption         =   "GENERATE TENDER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00C0FFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFF00&
      Caption         =   "Print"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFF00&
      Caption         =   "OK"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "generate_tender.frx":0000
      Left            =   7080
      List            =   "generate_tender.frx":0002
      TabIndex        =   25
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFF00&
      Caption         =   "Open Existing Tender"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "GENERATE NEW TENDER"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2355
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TENDER INFORMATION ENTRY"
      ForeColor       =   &H00004040&
      Height          =   5295
      Left            =   360
      TabIndex        =   29
      Top             =   2760
      Width           =   8895
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   5760
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   5760
         TabIndex        =   40
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   5760
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFF00&
         Caption         =   "Submit"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Add New"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Next"
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Previous"
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         Height          =   375
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPT.NO."
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "OPENING DATE"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "CLOSING DATE"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4320
         TabIndex        =   48
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RAW MATERIAL ID"
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4320
         TabIndex        =   46
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TENDER NO."
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TENDER INFORMATION ENTRY"
      ForeColor       =   &H00004040&
      Height          =   3495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   8895
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "REQUIRED RAW MATERIALS ENTRY"
         ForeColor       =   &H00004040&
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   8775
         Begin VB.CommandButton Command18 
            BackColor       =   &H00FFFF00&
            Caption         =   "Exit"
            Height          =   375
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FFFF00&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFF00&
            Caption         =   "Previous"
            Height          =   375
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFFF00&
            Caption         =   "Next"
            Height          =   375
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFFF00&
            Caption         =   "Delete"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFF00&
            Caption         =   "Add"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFF00&
            Caption         =   "Submit"
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   6240
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   2400
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER QUANTITY"
            ForeColor       =   &H00004040&
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   14
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER RAW MATERIAL ID"
            ForeColor       =   &H00004040&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         Height          =   255
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER CLOSING DATE"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER OPENING DATE"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER DEPT. NO."
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TENDER NO."
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RAW MATERIALS BELOW REORDER LEVEL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "generate_tender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim rst4 As New ADODB.Recordset
Dim rst5 As New ADODB.Recordset
Dim rst6 As New ADODB.Recordset
Dim fsys As New FileSystemObject
Dim flag As Integer
Public Sub test()
'test whether the tender is pending or not
If rst6("status") = "Pending" Then
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text14.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command15.Enabled = True
Else
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text14.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command15.Enabled = False
End If
End Sub

Private Sub Command1_Click()
'code of generate new tender
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text13.Enabled = True
Command2.Enabled = True
Dim rst1 As New ADODB.Recordset
Dim sql As String
If Not rst2.BOF Or Not rst2.EOF Then
rst2.MoveLast
End If
rst2.AddNew
If Not rst3.BOF Or Not rst3.EOF Then
rst3.MoveLast
End If
rst3.AddNew
Frame1.Visible = True
If flag = 1 Then
Frame2.Visible = False
Frame1.Enabled = True
Command3.Enabled = True
Command4.Enabled = False
End If
flag = 1
Combo1.Enabled = False
Command16.Enabled = False
End Sub

Private Sub Command10_Click()
'code of submit(frame3) button
rst5.Update
rst6.Update
Command10.Enabled = False
Command11.Enabled = True
End Sub

Private Sub Command11_Click()
'code of addnew(frame3) button
Dim temp
temp = rst5("tender_no")
If Not rst5.BOF Or Not rst5.EOF Then
rst5.MoveLast
End If
rst5.AddNew
rst5("tender_no") = temp
Command11.Enabled = False
Command10.Enabled = True
Text11.SetFocus
End Sub

Private Sub Command13_Click()
'code of next(frame3) button
rst5.MoveNext
If rst5.EOF Then
MsgBox ("You are on the Last Record")
If rst5.RecordCount <> 0 Then
rst5.MoveLast
End If
End If
End Sub

Private Sub Command14_Click()
'code of previous(frame3) button
rst5.MovePrevious
If rst5.BOF Then
MsgBox ("You are on the First Record")
If rst5.RecordCount <> 0 Then
rst5.MoveFirst
End If
End If
End Sub

Private Sub Command15_Click()
'code of cancel(frame3) button
rst5.CancelUpdate
rst6.CancelUpdate
Text7.Text = rst5("tender_no")
Text11.Text = rst5("raw_material_id")
Text12.Text = rst5("quantity")
Text8.Text = rst6("dept_no")
Text9.Text = rst6("op_date")
Text10.Text = rst6("cl_date")
Text14.Text = rst6("status")
End Sub

Private Sub Command16_Click()
'code of ok button
Frame1.Visible = False
Frame3.Visible = True
rst5.Open "select * from tender_detail  where  tender_no='" & Combo1.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
rst6.Open "select * from tender_info where  tender_no='" & Combo1.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text7.DataSource = rst5
Set Text11.DataSource = rst5
Set Text12.DataSource = rst5
Set Text8.DataSource = rst6
Set Text9.DataSource = rst6
Set Text10.DataSource = rst6
Set Text14.DataSource = rst6
Text7.DataField = "tender_no"
Text11.DataField = "raw_material_id"
Text12.DataField = "quantity"
Text8.DataField = "dept_no"
Text9.DataField = "op_date"
Text10.DataField = "cl_date"
Text14.DataField = "status"
Combo1.Enabled = False
Command16.Enabled = False
test
End Sub

Private Sub Command17_Click()
'code of exit(frame3) button
rst5.Close
rst6.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command18_Click()
'code of exit(frame2) button
rst.Close
rst2.Close
rst3.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command19_Click()
'code of print button
Dim rst7 As New ADODB.Recordset
If flag = 1 Then
rst7.Open "select ti.status,ti.tender_no,ti.dept_no,ti.op_date,ti.cl_date,td.raw_material_id,td.quantity,r.description,r.unit_of_measurement from tender_info ti,tender_detail td,raw_materials r where ti.tender_no=td.tender_no and td.raw_material_id=r.raw_material_id and ti.tender_no='" & Text13.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Else
rst7.Open "select ti.status,ti.tender_no,ti.dept_no,ti.op_date,ti.cl_date,td.raw_material_id,td.quantity,r.description,r.unit_of_measurement from tender_info ti,tender_detail td,raw_materials r where ti.tender_no=td.tender_no and td.raw_material_id=r.raw_material_id and ti.tender_no='" & Text7.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
End If
If rst7("status") <> "Pending" Then
MsgBox ("This tender is already published")
Exit Sub
Else
Dim OUTSTREAM As TextStream
Dim str, sql As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\tender.doc", True, False)
Dim sl_no
sl_no = 1
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
prec = Space(30) & "TENDER FOR RAW MATERIALS"
OUTSTREAM.WriteLine prec
prec = Space(30) & "************************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "Tender No    :" & rst7("tender_no")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Dept No      :" & rst7("dept_no")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Opening Date :" & rst7("op_date")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Closing Date :" & rst7("cl_date")
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "Dear Sir/Madam,"
OUTSTREAM.WriteLine prec
prec = Space(20) & "We are asking a tender from you for following raw materials :"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "SL." & " MATERIAL" & " DESCRIPTION" & Space(20) & " UNIT" & Space(2) & "QUANTITY"
OUTSTREAM.WriteLine prec
prec = Space(5) & "NO." & "  CODE"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
sql = "update tender_info set status='Called' where tender_no='" & rst7("tender_no") & "'"
Set r = cnn.Execute(sql)
While rst7.EOF = False
prec = Space(5) & sl_no & Space(4 - Len(sl_no)) & rst7.Fields("raw_material_id") & Space(9 - Len(rst7.Fields("raw_material_id"))) & rst7.Fields("description") & Space(33 - Len(rst7.Fields("description"))) & rst7.Fields("unit_of_measurement") & Space(14 - Len(rst7.Fields("unit_of_measurement")) - Len(rst7.Fields("quantity"))) & rst7.Fields("quantity")
OUTSTREAM.WriteLine prec
sl_no = sl_no + 1
rst7.MoveNext
Wend
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
MsgBox ("Print is over")
Frame2.Visible = False
Frame1.Enabled = False
Command3.Enabled = True
Command4.Enabled = False
flag = 0
End If
End Sub

Private Sub Command2_Click()
'code of ok(frame1) button
Dim sql
rst2("tender_no") = Text1.Text
rst3("tender_no") = Text1.Text
rst3.Update
sql = "update tender_info set status='Pending' where upper(tender_no)='" & UCase(Text1.Text) & "'"
Set r = cnn.Execute(sql)
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text13.Enabled = False
Command2.Enabled = False
Frame2.Visible = True
End Sub
Private Sub Command3_Click()
'code of submit button
rst2.Update
Command3.Enabled = False
Command4.Enabled = True
Command4.SetFocus
End Sub

Private Sub Command4_Click()
'code of add button
Dim temp
temp = rst2("tender_no")
If Not rst2.BOF Or Not rst2.EOF Then
rst2.MoveLast
End If
rst2.AddNew
rst2.Fields("tender_no") = temp
Command3.Enabled = True
Command4.Enabled = False
Text5.SetFocus
End Sub

Private Sub Command5_Click()
If rst2.EOF = True Then
MsgBox ("Eof has occured")
Else
rst2.Delete
rst2.Update
If Not rst.BOF Or Not rst.EOF Then
If rst.RecordCount > 1 Then
rst.MoveNext
End If
End If

If rst.EOF = True Then
If rst.RecordCount > 1 Then
rst.MovePrevious
End If
End If
End If
End If
rst2.Close
rst2.Open "select * from trial", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
End Sub

Private Sub Command6_Click()
'code of next button
Dim temp
temp = rst2("tender_no")
rst2.MoveNext
If rst2.EOF = True Then
MsgBox ("you are on the last record")
rst2.MovePrevious
Else
If rst2("tender_no") = temp Then
Set Text5.DataSource = rst2
Set Text6.DataSource = rst2
Text5.DataField = "raw_material_id"
Text6.DataField = "quantity"
Else
rst2.MovePrevious
Set Text5.DataSource = rst2
Set Text6.DataSource = rst2
Text5.DataField = "raw_material_id"
Text6.DataField = "quantity"
MsgBox ("you are on the last record")
End If
End If
End Sub

Private Sub Command7_Click()
'code of previous button
Dim temp
temp = rst2("tender_no")
rst2.MovePrevious
If rst2.BOF = True Then
MsgBox ("you are on the first record")
rst2.MoveFirst
Else
If rst2("tender_no") = temp Then
Set Text5.DataSource = rst2
Set Text6.DataSource = rst2
Text5.DataField = "raw_material_id"
Text6.DataField = "quantity"
Else
rst2.MoveNext
Set Text5.DataSource = rst2
Set Text6.DataSource = rst2
Text5.DataField = "raw_material_id"
Text6.DataField = "quantity"
MsgBox ("you are on the first record")
End If
End If
End Sub

Private Sub Command8_Click()
'code of cancel button
rst2.CancelUpdate
Text5.Text = rst2("raw_material_id")
Text6.Text = rst2("quantity")
End Sub

Private Sub Command9_Click()
'code of Open Existing Tender button
Combo1.Enabled = True
Command16.Enabled = True
rst4.Open "select  tender_no from tender_info", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst4.EOF = False
Combo1.AddItem rst4.Fields("tender_no")
rst4.MoveNext
Wend
End Sub

Private Sub Form_Load()
cnn.Open "DSN=MIS; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from raw_materials r where r.current_stock < r.reorder_level", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
rst2.Open "select * from tender_detail", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst2
Set Text5.DataSource = rst2
Set Text6.DataSource = rst2
Text1.DataField = "tender_no"
Text5.DataField = "raw_material_id"
Text6.DataField = "quantity"
rst3.Open "select * from tender_info", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst3
Set Text2.DataSource = rst3
Set Text3.DataSource = rst3
Set Text4.DataSource = rst3
Text1.DataField = "tender_no"
Text2.DataField = "dept_no"
Text3.DataField = "op_date"
Text4.DataField = "cl_date"
Command4.Enabled = False
Combo1.Enabled = False
Command16.Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub


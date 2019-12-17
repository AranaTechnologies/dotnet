VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form evaluate_tender 
   BackColor       =   &H00C0FFFF&
   Caption         =   "EVALUATE TENDER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Print Purchase Order"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Submit"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3413
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
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EVALUATE TENDER"
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
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select  Approved Vendor ID"
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
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Tender No."
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
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "evaluate_tender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim fsys As New FileSystemObject
Dim rst As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Private Sub Command1_Click()
'code of submit button
Dim rst4 As New ADODB.Recordset
rst.Open " select t.tender_no,s.vendor_id,s.raw_material_id,r.description,r.unit_of_measurement,t.quantity,s.price,s.price * t.quantity Amount from tender_detail t,submitted_tender_form s,vendor v,raw_materials r where upper(t.tender_no)=upper(s.tender_no) and upper(t.raw_material_id)=upper(s.raw_material_id) and upper(s.vendor_id)=upper(v.vendor_id) and upper(s.raw_material_id)=upper(r.raw_material_id) and  upper(t.tender_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
rst4.Open "select distinct vendor_id from submitted_tender_form where upper(tender_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst4.EOF = False
Combo2.AddItem rst4.Fields("vendor_id")
rst4.MoveNext
Wend
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
'code of print purchase order
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\purchase_order.doc", True, False)
Dim sql As String
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim sl_no
sql = "update tender_info set status ='Done' where upper(tender_no) ='" & UCase(Combo1.Text) & "'"
Set r1 = cnn.Execute(sql)
rst.Close
rst3.Close
rst3.Open " select tender_no from tender_info  where status='Called' and status='Done'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst3.EOF = False
Combo1.AddItem rst3.Fields("tender_no")
rst3.MoveNext
Wend
sql = "select * from final"
Set rst1 = cnn.Execute(sql)
temp = rst1("purchase_order_no")
sql = "update final set purchase_order_no=purchase_order_no+1"
Set r2 = cnn.Execute(sql)
temp = Trim(temp)
sql = "insert into purchase_order_info(tender_no,purchase_order_no,vendor_id,status) values('" & UCase(Combo1.Text) & "','" & "PO" & temp & "','" & UCase(Combo2.Text) & "','Pending')"
Set r3 = cnn.Execute(sql)
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
prec = Space(30) & "PURCHASE ORDER"
OUTSTREAM.WriteLine prec
prec = Space(30) & "**************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "To:"
OUTSTREAM.WriteLine prec
rst2.Open " select t.tender_no,s.vendor_id,s.raw_material_id,r.description,r.unit_of_measurement,t.quantity,s.price,s.price * t.quantity Amount,v.name,v.building_no,v.street_name,v.city,v.state,v.pin,v.ph_no from tender_detail t,submitted_tender_form s,vendor v,raw_materials r where upper(t.tender_no)=upper(s.tender_no) and upper(t.raw_material_id)=upper(s.raw_material_id) and upper(s.vendor_id)=upper(v.vendor_id) and upper(s.raw_material_id)=upper(r.raw_material_id) and  upper(t.tender_no)='" & UCase(Combo1.Text) & "' and upper(s.vendor_id) = '" & UCase(Combo2.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
prec = Space(5) & rst2.Fields("name")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst2.Fields("building_no") & "," & rst2.Fields("street_name")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst2.Fields("city")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst2.Fields("state")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Pin- " & rst2.Fields("pin")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Phone No.- " & rst2.Fields("ph_no") & Space(50) & "Purchase Order No : " & "PO" & temp
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "Dear Sir/Madam,"
OUTSTREAM.WriteLine prec
prec = Space(20) & "We are pleased to hereby place an order on you for supply of the following as par"
OUTSTREAM.WriteLine prec
prec = Space(5) & "the tender form submitted from your end. The details are furnished below:"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "SL." & " MATERIAL" & " DESCRIPTION" & Space(20) & " UNIT" & Space(2) & "QUANTITY" & Space(3) & "RATE" & Space(10) & "AMOUNT"
OUTSTREAM.WriteLine prec
prec = Space(5) & "NO." & "  CODE"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
While rst2.EOF = False
sql = "insert into purchase_register values('" & "PO" & temp & "','" & rst2("raw_material_id") & "'," & rst2("quantity") & ",'" & rst2("unit_of_measurement") & "')"
Set r4 = cnn.Execute(sql)
prec = Space(5) & sl_no & Space(4 - Len(sl_no)) & rst2.Fields("raw_material_id") & Space(9 - Len(rst2.Fields("raw_material_id"))) & rst2.Fields("description") & Space(33 - Len(rst2.Fields("description"))) & rst2.Fields("unit_of_measurement") & Space(14 - Len(rst2.Fields("unit_of_measurement")) - Len(rst2.Fields("quantity"))) & rst2.Fields("quantity") & Space(7 - Len(rst2.Fields("price"))) & rst2.Fields("price") & Space(16 - Len(rst2.Fields("amount"))) & rst2.Fields("amount")
OUTSTREAM.WriteLine prec
sl_no = sl_no + 1
rst2.MoveNext
Wend
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
MsgBox ("Print is over")
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
'code of exit button
rst3.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Form_Load()
cnn.Open "DSN=MIS; provider=MSDASQL; uid=mis; pwd=mis1"
rst3.Open " select tender_no from tender_info  where status='Called' and status<>'Done'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst3.EOF = False
Combo1.AddItem rst3.Fields("tender_no")
rst3.MoveNext
Wend
End Sub

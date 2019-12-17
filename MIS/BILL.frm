VERSION 5.00
Begin VB.Form bill 
   BackColor       =   &H00C0FFFF&
   Caption         =   "BILL ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "EXIT"
      Height          =   375
      Left            =   5760
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "GENERATE PAY ORDER"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5160
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mmm-yyyy; e.g. 01-mar-2004)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mmm-yyyy; e.g. 01-mar-2004)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER RECEIVED DATE"
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
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BILL ENTRY"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT PURCHASE ORDER NO."
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
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER DELIVERY DATE"
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   3375
   End
End
Attribute VB_Name = "bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub Combo1_Click()
'code of combo box
Command1.Enabled = True
End Sub

Private Sub Command1_Click()
'code of generate pay order and update stock
Dim amount As Long
Dim total As Long
Dim sql As String
Dim sql1 As String
Dim prev_stock
Dim curr_stock
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim rst4 As New ADODB.Recordset
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\pay_order.doc", True, False)

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
prec = Space(35) & "PAY ORDER"
OUTSTREAM.WriteLine prec
prec = Space(35) & "*********"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "To"
OUTSTREAM.WriteLine prec
prec = Space(5) & "The Finance Officer."
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "sir/madam,"
OUTSTREAM.WriteLine prec
prec = Space(15) & "You are requested to pay the amount as par the Purchase Order No:" & Combo1.Text
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "SL." & " MATERIAL" & " DESCRIPTION" & Space(20) & " UNIT" & " QUANTITY" & Space(2) & " RATE" & "  AMOUNT"
OUTSTREAM.WriteLine prec
prec = Space(5) & "NO." & "  CODE"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine

sql = "update purchase_order_info set delivery_date='" & Text1.Text & "',received_date='" & Text2.Text & "',status= 'Purchased' where upper(purchase_order_no)='" & UCase(Combo1.Text) & "'"
Set r = cnn.Execute(sql)
'rst1.Open " select pr.purchase_order_no,pr.raw_material_id,pr.quantity,pi.tender_no,pi.vendor_id,r.description,r.unit_of_measurement,s.price,s.price * pr.quantity amount from purchase_register pr,purchase_order_info pi,raw_materials r,submitted_tender_form s where pr.purchase_order_no=pi.purchase_order_no and pr.raw_material_id=s.raw_material_id and pi.tender_no=s.tender_no and pi.vendor_id=s.vendor_id and pr.raw_material_id=r.raw_material_id and upper(pr.purchase_order_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText

rst1.Open " select * from purchase_register where upper(purchase_order_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst1.EOF = False
sql = "select pi.tender_no,s.price,r.description,r.unit_of_measurement from raw_materials r,purchase_order_info pi,submitted_tender_form s where upper(pi.tender_no)=upper(s.tender_no) and upper(pi.vendor_id)=upper(s.vendor_id) and r.raw_material_id=s.raw_material_id and s.raw_material_id='" & rst1("raw_material_id") & "'"
Set rst4 = cnn.Execute(sql)

'Set DataGrid1.DataSource = rst1
'rst1.Open "select pr.purchase_order_no,pr.raw_material_id,r.description,r.unit_of_measurement,pr.quantity from purchase_register pr,raw_materials r where  upper(pr.raw_material_id)=upper(r.raw_material_id) and upper(pr.purchase_order_no)='" & UCase(Combo1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
'While rst1.EOF = False
amount = rst1("quantity") * rst4("price")
total = total + amount
prec = Space(5) & sl_no & Space(4 - Len(sl_no)) & rst1.Fields("raw_material_id") & Space(9 - Len(rst1.Fields("raw_material_id"))) & rst4.Fields("description") & Space(33 - Len(rst4.Fields("description"))) & rst4.Fields("unit_of_measurement") & Space(6 - Len(rst4.Fields("unit_of_measurement"))) & rst1.Fields("quantity") & Space(11 - Len(rst1.Fields("quantity"))) & rst4.Fields("price") & Space(6 - Len(rst4.Fields("price"))) & amount
OUTSTREAM.WriteLine prec
sql = "select * from raw_materials where raw_material_id='" & rst1("raw_material_id") & "'"
Set rst3 = cnn.Execute(sql)
prev_stock = rst3("current_stock")
curr_stock = prev_stock + rst1("quantity")
sql1 = "insert into raw_material_stock values('" & Combo1.Text & "','" & rst1("raw_material_id") & "','" & rst4("description") & "','" & rst4("unit_of_measurement") & "'," & rst1("quantity") & ",'PURCHASED','" & Text2.Text & "'," & prev_stock & "," & curr_stock & ")"
Set rst2 = cnn.Execute(sql1)
sql = "update raw_materials set current_stock=" & curr_stock & "where raw_material_id='" & rst1("raw_material_id") & "'"
Set r3 = cnn.Execute(sql)
sl_no = sl_no + 1
rst1.MoveNext
Wend
Command1.Enabled = False
rst.Close
rst.Open " select purchase_order_no from purchase_order_info  where status='Pending'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst.Fields("purchase_order_no")
rst.MoveNext
Wend
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(66) & "Total : " & total
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(65) & "(Signature of Purchase Manager)"
OUTSTREAM.WriteLine prec
MsgBox ("Print is over")
End Sub

Private Sub Command2_Click()
'code of exit button
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open " select purchase_order_no from purchase_order_info  where status='Pending'", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst.Fields("purchase_order_no")
rst.MoveNext
Wend
End Sub


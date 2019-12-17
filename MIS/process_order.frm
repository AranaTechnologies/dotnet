VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmprocess_order 
   BackColor       =   &H00C0FFFF&
   Caption         =   "PROCESS ORDER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "EXIT"
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ENTER ISSUE DATE"
      ForeColor       =   &H00004000&
      Height          =   1815
      Left            =   960
      TabIndex        =   5
      Top             =   3240
      Width           =   5775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   735
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
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER ISSUE DATE"
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
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10320
      Top             =   3000
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdupdate_status 
      BackColor       =   &H00FFFF00&
      Caption         =   "UPDATE STATUS"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdgenerate_delivery 
      BackColor       =   &H00FFFF00&
      Caption         =   "GENERATE CURRENT DELIVERY NOTE"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdgenerate_bill 
      BackColor       =   &H00FFFF00&
      Caption         =   "GENERATE CURRENT BILL"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2990
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESS CUSTOMER ORDER"
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
      Height          =   735
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmprocess_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim fsys As New FileSystemObject
Dim rst As New ADODB.Recordset
Private Sub cmdgenerate_bill_Click()
'code of generate current bill
If rst("status") <> "done" Then
MsgBox ("Before generate bill you should update the status of the order")
Exit Sub
Else
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\current_bill.doc", True, False)
Dim sql As String
Dim rst1 As New ADODB.Recordset
sql = "select * from final"
Set rst1 = cnn.Execute(sql)
temp = rst1("bill_no")
sql = "update final set bill_no = bill_no + 1"
Set r = cnn.Execute(sql)
temp = Trim(temp)
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
prec = Space(48) & "B  I  L  L"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "To:"
OUTSTREAM.WriteLine prec
prec = Space(5) & "Customer ID: " & rst.Fields("customer_id")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("name_of_the_organization")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("building_no") & "," & rst.Fields("street_name")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("city")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("state")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Pin- " & rst.Fields("pin") & Space(70) & "Bill No: " & temp
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "PRODUCT ID" & Space(5) & "DESCRIPTION" & Space(24) & "QUANTITY" & Space(5) & "UNIT" & Space(6) & "RATE" & Space(7) & "AMOUNT"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("product_id") & Space(15 - Len(rst.Fields("product_id"))) & rst.Fields("description") & Space(35 - Len(rst.Fields("description"))) & rst.Fields("quantity") & Space(13 - Len(rst.Fields("quantity"))) & rst.Fields("unit_of_measurement") & Space(11 - Len(rst.Fields("unit_of_measurement"))) & rst.Fields("price") & Space(11 - Len(rst.Fields("price"))) & rst.Fields("amount")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(89) & "(Signature)"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "----------------------------------------" & "CUT FROM HERE" & "-----------------------------------------"
OUTSTREAM.WriteLine prec
MsgBox ("Printing is over")
rst.MoveFirst
End If
End Sub
Private Sub cmdgenerate_delivery_Click()
'code of generate delivery note
If rst("status") <> "done" Then
MsgBox ("Before generate delivery note you should update the status of the order")
Exit Sub
Else
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\current_delivery_note.DOC", True, False)
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
prec = Space(44) & "DELIVERY NOTE"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "THE FOLLOWING MATERIAL HAS BEEN DELIVERED"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "TO:"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("name_of_the_organization")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("building_no") & "," & rst.Fields("street_name")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("city")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("state")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Pin- " & rst.Fields("pin")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "ON  " & Date & "  AT  " & Time() & "  AGGAINST ORDER NO.  " & rst.Fields("CUSTOMER_ORDER_NO") & "  DATE  " & rst.Fields("ORDER_DATE")
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "PRODUCT ID" & Space(5) & "DESCRIPTION" & Space(24) & "QUANTITY" & Space(5) & "UNIT"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("product_id") & Space(15 - Len(rst.Fields("product_id"))) & rst.Fields("description") & Space(35 - Len(rst.Fields("description"))) & rst.Fields("quantity") & Space(13 - Len(rst.Fields("quantity"))) & rst.Fields("unit_of_measurement")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(72) & "(SIGNATURE OF THE STAFF)"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "THE ABOVE MENTIONED MATERIAL HAS BEEN RECEIVED" & Space(21) & "(SIGNATURE OF THE CUSTOMER)"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------" & "CUT FROM HERE" & "-----------------------------------------"
OUTSTREAM.WriteLine prec
rst.MoveFirst
MsgBox ("Printing is over")
End If
End Sub
Private Sub cmdupdate_status_Click()
'code of update status button
If rst("status") = "done" Then
MsgBox ("This record is already updated")
Exit Sub
Else
Frame1.Visible = True
End If
End Sub
Private Sub Command1_Click()
'code of ok button
Dim sql As String
Dim temp
Dim rst1 As New ADODB.Recordset
Dim curr_stock As Integer
Dim prev_stock As Integer
temp = rst("customer_order_no")
sql = "select current_stock from product where product_id='" & rst("product_id") & "'"
Set rst1 = cnn.Execute(sql)
prev_stock = rst1("current_stock")
If prev_stock < rst("quantity") Then
MsgBox ("There is not sufficient stock of " & rst("product_id"))
Frame1.Visible = False
Exit Sub
Else
curr_stock = prev_stock - rst("quantity")
sql = "insert into product_stock values('" & rst("customer_order_no") & "','" & rst("product_id") & "'," & rst("quantity") & ",'Issued','" & Text1.Text & "'," & curr_stock & ")"
Set r2 = cnn.Execute(sql)
sql = "update product set current_stock=" & curr_stock & "where product_id='" & rst("product_id") & "'"
Set r3 = cnn.Execute(sql)
sql = "update customer_order set status='done',issue_date='" & Text1.Text & "' where customer_order_no='" & temp & "'"
Set r1 = cnn.Execute(sql)
rst("status") = "done"
End If
Frame1.Visible = False
End Sub
Private Sub Command2_Click()
'CODE OF EXIT BUTTON
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub
Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis;pwd=mis1"
rst.Open " select o.customer_id ,c.name_of_the_organization , o.customer_order_no,o.order_date,c.city,c.state,c.pin,c.building_no,c.street_name,p.product_id,p.description,p.price,p.unit_of_measurement,o.quantity, p.price * o.quantity Amount,o.status  from  customer  c , product  p,customer_order  o where upper(c.customer_id)=upper(o.customer_id) and upper(p.product_id)= upper(o.product_id)", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
Frame1.Visible = False
End Sub


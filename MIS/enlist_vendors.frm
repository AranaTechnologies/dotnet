VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form enlist_vendors 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ENLIST VENDORS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "EXIT"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "GRADE"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter Grade"
      ForeColor       =   &H00004040&
      Height          =   1095
      Left            =   2160
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER GRADE"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Generate Acknowledgement"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enlist Vendor"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   12648447
      Enabled         =   -1  'True
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENLIST VENDOR"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "enlist_vendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim fsys As New FileSystemObject
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
'code of enlist vendor button
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim tno
If rst("vendor_id") = " " Then
sql = "select * from final"
Set rst1 = cnn.Execute(sql)
temp = rst1("vendor_id")
sql = "update final set vendor_id=vendor_id+1"
Set r = cnn.Execute(sql)
temp = Trim(temp)
tno = rst("trade_license_no")
sql = "update vendor set vendor_id= '" & "V" & temp & "' where trade_license_no = '" & tno & "'"
Set rst2 = cnn.Execute(sql)
rst.Close
rst.Open "select v.vendor_id,v.name,v.trade_license_no,v.building_no,v.street_name,v.city,v.state,v.pin,v.ph_no,v.email_id,v.grade from vendor v", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
rst.MoveNext
If rst.EOF = True Then
rst.MoveFirst
End If
Else
MsgBox ("This vendor is already enlisted")
End If
End Sub

Private Sub Command2_Click()
'code of generate acknowledgement button
If rst("vendor_id") = " " Then
MsgBox ("This vendor is not yet enlisted.")
Exit Sub
Else
Dim OUTSTREAM As TextStream
Dim str As String
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\acknowledgement.doc", True, False)
Dim sql As String
Dim rst1 As New ADODB.Recordset
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
prec = Space(30) & "ACKNOWLEDGEMENT LETTER OF YOUR REGISTRATION"
OUTSTREAM.WriteLine prec
prec = Space(30) & "**********************************************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "To:"
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("name")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("building_no") & "," & rst.Fields("street_name")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("city")
OUTSTREAM.WriteLine prec
prec = Space(5) & rst.Fields("state")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Pin- " & rst.Fields("pin")
OUTSTREAM.WriteLine prec
prec = Space(5) & "Phone No.- " & rst.Fields("ph_no")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "Dear Sir/Madam,"
OUTSTREAM.WriteLine prec
prec = Space(20) & "You will be glad to know that your application have been accepted by us. Now, you"
OUTSTREAM.WriteLine prec
prec = Space(5) & "become an enlisted vendor of our company."
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(40) & "YOUR VENDOR ID:- " & rst.Fields("vendor_id")
OUTSTREAM.WriteLine prec
prec = Space(40) & "********************"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(86) & "Thanking you,"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(86) & "(Purchase Manager)"
OUTSTREAM.WriteLine prec
rst.MoveFirst
MsgBox ("Printing is over")
End If
End Sub

Private Sub Command3_Click()
'code of ok(frame1) button
Dim sql As String
Dim tno
tno = rst("trade_license_no")
sql = "update vendor set grade='" & UCase(Text1.Text) & "'where trade_license_no='" & tno & "'"
Set r = cnn.Execute(sql)
Frame1.Visible = False
rst.Close
rst.Open "select v.vendor_id,v.name,v.trade_license_no,v.building_no,v.street_name,v.city,v.state,v.pin,v.ph_no,v.email_id,v.grade from vendor v", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
rst.MoveNext
If rst.EOF = True Then
rst.MoveFirst
End If
End Sub

Private Sub Command4_Click()
'code of grade button
If rst("vendor_id") <> " " Then
Frame1.Visible = True
Else
MsgBox ("This vendor is not enlisted")
End If
End Sub

Private Sub Command5_Click()
'CODE OF EXIT BUTTON
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select v.vendor_id,v.name,v.trade_license_no,v.building_no,v.street_name,v.city,v.state,v.pin,v.ph_no,v.email_id,v.grade from vendor v", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
Frame1.Visible = False
End Sub

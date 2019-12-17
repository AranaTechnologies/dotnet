VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form23 
   BackColor       =   &H0080C0FF&
   Caption         =   "Employee Information Entry"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form23"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "Browse"
      Height          =   375
      Left            =   9840
      TabIndex        =   34
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   375
      Left            =   10080
      TabIndex        =   33
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   7680
      TabIndex        =   32
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10200
      TabIndex        =   31
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
      Height          =   495
      Left            =   8520
      TabIndex        =   30
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   5400
      TabIndex        =   28
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6960
      TabIndex        =   27
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Employee Record"
      Height          =   495
      Left            =   2280
      TabIndex        =   26
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   480
      TabIndex        =   25
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   3840
      TabIndex        =   23
      Top             =   7680
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Office Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   6720
      TabIndex        =   14
      Top             =   2760
      Width           =   4815
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Increment"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Joining "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   5295
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3480
         TabIndex        =   9
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter Test Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   29
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Information Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim RST1 As New ADODB.Recordset
'SEARCH BUTTON'
Private Sub Command1_Click()
Label11.Visible = True
Text12.Visible = True
Command2.Visible = True
Command1.Enabled = False

End Sub
'BROWSE IMAGE'
Private Sub Command10_Click()

Dim filename As String
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
If Len(filename) = 0 Then
 filename = "C:\MCA\IMDMS\PICS\d.bmp"
End If
Image1.Picture = LoadPicture(filename)
RST.Fields("PATH") = filename
RST.Update

End Sub

'GO BUTTON'
Private Sub Command2_Click()
Dim flag As Integer
flag = 0
RST.MoveFirst
Do While RST.EOF = False
If Trim(UCase(RST.Fields("EMPLOYEE_ID"))) = Trim(UCase(Text4.Text)) Then
flag = 1
Exit Do
End If
RST.MoveNext
Loop
Label4.Visible = False
Text4.Visible = False
Command2.Visible = False
Command1.Enabled = True
If flag = 0 Then
MsgBox ("No EMPLOYEE bears the EMPLOYEE-ID " & Text4.Text)
RST.MoveFirst
End If

End Sub
'ADD NEW BUTTON'
Private Sub Command3_Click()
Dim filename As String
If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command4.Enabled = True
sql = "SELECT * FROM GENERATED_LAST"
Set RST1 = CNN.Execute(sql)
temp = RST1.Fields("EMPLOYEE_ID")
sql = "UPDATE GENERATED_LAST SET EMPLOYEE_ID=EMPLOYEE_ID+1"
Set RST1 = CNN.Execute(sql)
temp = Trim(temp)
temp1 = "EMPLOYEE" & temp
RST.Fields("employee_ID") = temp1


End Sub
'SUBMIT BUTTON'
Private Sub Command4_Click()
RST.Update
Command4.Enabled = False
Command3.Enabled = True
End Sub
'E-MAIL BUTTON'
Private Sub Command5_Click()
If Len(Trim(RST.Fields("SEND_EMAIL"))) <> 0 Then
Form13.Text1.Text = RST.Fields("SEND_EMAIL")
End If
RST.Close
CNN.Close
Form13.Show
End Sub

'EXIT BUTTON'
Private Sub Command6_Click()

RST.Close
CNN.Close
Unload Me
End Sub

Private Sub Command7_Click()
If Not RST.BOF Or Not RST.EOF Then
RST.MoveNext
If Not RST.EOF Then
End If
End If


If RST.EOF Then
MsgBox ("You are on the Last Record")
If RST.RecordCount <> 0 Then
RST.MoveLast
End If
End If
End Sub

'DELETE BUTTON'
Private Sub Command8_Click()
Dim response As Integer
Dim message As String
message = "Delete the record of " & UCase(Text1.Text) & "?"
response = MsgBox(message, 36, "Delete Record")
If response = 6 Then
If RST.EOF = True Then
MsgBox ("Eof has occured")
Else
RST.Delete
RST.Update
If Not RST.BOF Or Not RST.EOF Then
If RST.RecordCount > 1 Then
RST.MoveNext
End If
End If
If RST.EOF = True Then
If RST.RecordCount > 1 Then
RST.MovePrevious
End If

End If
End If
End If
RST.Close
RST.Open "select * from EMPLOYEE_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST
Set Text6.DataSource = RST
Set Text7.DataSource = RST
Set Text8.DataSource = RST
Set Text9.DataSource = RST
Set Text10.DataSource = RST
Set Text11.DataSource = RST

End Sub
'PREVIOUS BUTTON'
Private Sub Command9_Click()
If Not RST.BOF Or Not RST.EOF Then
RST.MovePrevious
If Not RST.BOF Then
End If
End If


If RST.BOF Then
MsgBox ("You are on the First Record")
If RST.RecordCount <> 0 Then
RST.MoveFirst
End If
If Not RST.BOF Or Not RST.EOF Then
RST.MoveFirst
End If
End If

End Sub

Private Sub Form_Activate()
Label11.Visible = False
Text12.Visible = False
Command2.Visible = False
End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from EMPLOYEE_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST
Set Text6.DataSource = RST
Set Text7.DataSource = RST
Set Text8.DataSource = RST
Set Text9.DataSource = RST
Set Text10.DataSource = RST
Set Text11.DataSource = RST
Set Image1.DataSource = RST
 
Text1.DataField = "employee_id"
Text2.DataField = "employee_name"
Text3.DataField = "address1"
Text4.DataField = "ADDRESS2"
Text5.DataField = "pin"
Text6.DataField = "telephone_no"
Text7.DataField = "date_of_birth"
Text8.DataField = "date_of_joining"
Text9.DataField = "basic"
Text10.DataField = "designation"
Text11.DataField = "date_of_increment"
If Len(RST.Fields("path")) <> 0 Then
Image1.Picture = LoadPicture(RST.Fields("path"))
End If
End Sub


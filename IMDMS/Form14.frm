VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H0080C0FF&
   Caption         =   "Quarter Information"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form14"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   375
      Left            =   9480
      TabIndex        =   19
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
      Height          =   495
      Left            =   8280
      TabIndex        =   17
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Quarter Record"
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quarter  Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5280
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   3840
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         TabIndex        =   8
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Quarter No"
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
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter Employee Id"
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
      Left            =   3480
      TabIndex        =   16
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quarter Information Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset

'SEARCH BUTTON'
Private Sub Command1_Click()
Label6.Visible = True
Text4.Visible = True
Command2.Visible = True
Command1.Enabled = False

End Sub
'GO BUTTON'
Private Sub Command2_Click()
Dim flag As Integer
flag = 0
RST.MoveFirst
Do While RST.EOF = False
If Trim(UCase(RST.Fields("QUARTER_NO"))) = Trim(UCase(Text4.Text)) Then
flag = 1
Exit Do
End If
RST.MoveNext
Loop
Label6.Visible = False
Text4.Visible = False
Command2.Visible = False
Command1.Enabled = True
If flag = 0 Then
MsgBox ("No QUARTER bears the QUARTER NO " & Text4.Text)
RST.MoveFirst
End If


End Sub
'ADD NEW'
Private Sub Command3_Click()
Dim filename As String

If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command4_Click()
RST.Update
Command4.Enabled = False
Command3.Enabled = True
End Sub
'SEND E-MAIL BUTTON'
Private Sub Command5_Click()
If Len(Trim(RST.Fields("SEND_EMAIL"))) <> 0 Then
Form13.Text1.Text = RST.Fields("SEND_EMAIL")
End If
RST.Close
CNN.Close
Form13.Show
End Sub
End Sub
'EXIT BUTTON'
Private Sub Command6_Click()
rst1.Close
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
RST.Open "select * from quarter_file", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
 


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
Label6.Visible = False
Text4.Visible = False
Command2.Visible = False
End Sub

Private Sub Form_Load()

CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from QUARTER_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
 
Text1.DataField = "QUARTER_NO"
Text2.DataField = "TYPE"
Text3.DataField = "EMPLOYEE_ID"
 
'ADD ITEM IN COMBOBOX'
rst1.Open "select * from QUARTER_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
While Not rst1.EOF
Combo1.AddItem (rst1.Fields("LOCATION"))
rst1.MoveNext
Wend
rst1.Close
Combo1.DataField = "LOCATION"
 
End Sub

VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H0080C0FF&
   Caption         =   "Test Information"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form7"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Previous"
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exitl"
      Height          =   360
      Left            =   9360
      TabIndex        =   17
      Top             =   7200
      Width           =   960
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Submit"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   7200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "  Test Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   4320
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      Begin VB.TextBox Text5 
         Height          =   1215
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   22
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Submit"
         Height          =   375
         Left            =   -1560
         TabIndex        =   13
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Top"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   10
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add New"
         Height          =   375
         Left            =   -3000
         TabIndex        =   9
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Test Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Amount of Sample in ml."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Type of Sample"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Test_Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Test_Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Test Information Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   120
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3315
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset

'SEARCH TEST RECORD BUTTON

Private Sub Command1_Click()
Label4.Visible = True
Text4.Visible = True
Command2.Visible = True
Command1.Enabled = False
End Sub

'Submit Button
 
Private Sub Command11_Click()
'RST.Update
Command11.Enabled = False
Command3.Enabled = True
If Len(Text1.Text) = 0 Then
SHOW_ERROR (10)
RST.Open
End If

End Sub

'GO BUTTON
Private Sub Command2_Click()
Dim flag As Integer
flag = 0
RST.MoveFirst
Do While RST.EOF = False
If Trim(UCase(RST.Fields("Test_id"))) = Trim(UCase(Text4.Text)) Then
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
MsgBox ("No Test bears the Test-Id " & Text4.Text)
RST.MoveFirst
End If
End Sub

End Sub


'ADD NEW
Private Sub Command3_Click()
Dim filename As String

If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command11.Enabled = True

sql = "SELECT * FROM GENERATED_LAST"
Set RST = CNN.Execute(sql)
temp = RST.Fields("TEST_ID")
sql = "UPDATE GENERATED_LAST SET TEST_ID=TEST_ID+1"
Set RST = CNN.Execute(sql)
temp = Trim(temp)
temp1 = "TEST" & temp
Text1.Text = temp1
End Sub

'EXIT BUTTON
Private Sub Command5_Click()
RST.Close
CNN.Close
Unload Me
End Sub
 

'NEXT BUTTON
Private Sub Command10_Click()
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
'DELETE BUTTON
Private Sub Command4_Click()

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
RST.Open "select * from Test_info_file", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text4.DataSource = RST


End Sub

Private Sub Form_Load()

CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from Test_info_file", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST

Text1.DataField = "Test_Id"
Text2.DataField = "test_name"
Text3.DataField = "type_of_sample"
Text4.DataField = "amount_of_sample"
Text5.DataField = "test_description"

Text1.DataField = "TEST_ID"
End Sub

Private Sub Form_Activate()

Command11.Enabled = False

End Sub

'PREVIOUS BUTTON
Private Sub Command12_Click()
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




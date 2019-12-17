VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0080C0FF&
   Caption         =   "Test Requisition Entry  "
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
      Height          =   495
      Left            =   8520
      TabIndex        =   17
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   375
      Left            =   10800
      TabIndex        =   12
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7680
      TabIndex        =   11
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Test Record"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Test Information"
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2160
         TabIndex        =   18
         Text            =   " "
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   435
         Left            =   2160
         TabIndex        =   8
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2160
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Status"
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
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Test_Id"
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
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Patient_Id "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Test  Requisition Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4440
      TabIndex        =   16
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label4 
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
      TabIndex        =   10
      Top             =   6720
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
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

 


'ADD NEW
Private Sub Command3_Click()
Dim filename As String

If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command4.Enabled = True

End Sub

'SUBMIT'
Private Sub Command4_Click()
RST.Update
Command4.Enabled = False
Command3.Enabled = True
If Len(Text1.Text) = 0 Then
SHOW_ERROR (20)
End If

End Sub

Private Sub Command5_Click()
If Len(Trim(RST.Fields("SEND_EMAIL"))) <> 0 Then
Form13.Text1.Text = RST.Fields("SEND_EMAIL")
End If
RST.Close
CNN.Close
Form13.Show
End Sub

'EXIT BUTTON
Private Sub Command6_Click()
RST.Close
CNN.Close
Unload Me
End Sub
 


'NEXT BUTTON
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


'DELETE BUTTON
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
RST.Open "select * from Test_Requisition_File", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
If Not RST.EOF Then
End If
End If


End Sub

Private Sub Form_Activate()
Label4.Visible = False
Text4.Visible = False
Command2.Visible = False
End Sub

Private Sub Form_Load()

CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from Test_Requisition_File", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
 
Text3.DataField = "Status"
Text2.DataField = "Test_Id"
Text1.DataField = "Patient_Id"

End Sub

'PREVIOUS BUTTON
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



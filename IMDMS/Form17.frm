VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H0080C0FF&
   Caption         =   "Department Test Info"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form17"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "Previous"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   360
      Left            =   10080
      TabIndex        =   13
      Top             =   7320
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E-Mail"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete"
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   11
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Submit"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   7320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Test Fees Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   5280
      TabIndex        =   1
      Top             =   1560
      Width           =   5295
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2640
         TabIndex        =   7
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Test_Fees"
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
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Test  Id"
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
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Department_name"
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
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Department test info"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
'SEND E-MAIL BUTTON'
Private Sub Command1_Click()
If Len(Trim(RST.Fields("SEND_EMAIL"))) <> 0 Then
Form13.Text1.Text = RST.Fields("SEND_EMAIL")
End If
RST.Close
CNN.Close
Form13.Show
End Sub
'NEXT'
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
'SUBMIT'
Private Sub Command11_Click()
If Not RST.BOF Or Not RST.EOF Then
RST.Update
Command11.Enabled = False
Command3.Enabled = True
'rst.Open
End If
End Sub
'PREVIOUS BUTTON'
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
'ADD NEW'
Private Sub Command3_Click()
Dim filename As String

If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command11.Enabled = True
End Sub
'EXIT BUTTON'
Private Sub Command5_Click()
RST.Close
CNN.Close
Unload Me
End Sub
'DELETE'
Private Sub Command6_Click(Index As Integer)

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
RST.Open "select * from DEPARTMENT_TEST_INFO", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
 


End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from DEPARTMENT_TEST_INFO", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST

Text1.DataField = "DEPARTMENT_NAME"
Text2.DataField = "TEST_ID"
Text3.DataField = "TEST_FEES"
 

End Sub

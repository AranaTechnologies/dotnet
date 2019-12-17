VERSION 5.00
Begin VB.Form Form32 
   BackColor       =   &H0080C0FF&
   Caption         =   "Tender Invitation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form32"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10200
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Required"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Id"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tender No."
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Invitation"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset

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
RST.Open "select * from Test_info_file", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text4.DataSource = RST



End Sub

Private Sub Command9_Click()
If Not RST.BOF Or Not RST.EOF Then
RST.MovePrevious
If Not RST.BOF Then
End If
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

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from TENDER_INVITATION_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
 
Text1.DataField = "TENDER_NO"
Text2.DataField = "ITEM_ID"
Text3.DataField = "QUANTITY_REQUIRED"

End Sub

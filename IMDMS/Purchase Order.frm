VERSION 5.00
Begin VB.Form Form31 
   BackColor       =   &H0080C0FF&
   Caption         =   "Purchase Order"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form31"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT PURCHASE ORDER"
      Height          =   615
      Left            =   5160
      TabIndex        =   16
      Top             =   7560
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10440
      TabIndex        =   15
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   7200
      TabIndex        =   14
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   5760
      TabIndex        =   12
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Order"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier _Id"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tender No"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form31"
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
If Len(Text1.Text) = 0 Then
SHOW_ERROR (20)
End If
If Not RST.EOF Or Not RST.BOF Then
RST.Update
Command4.Enabled = False
Command3.Enabled = True
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
RST.Open "select * from ORDER_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
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
RST.Open "select * from ORDER_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Text1.DataField = "ORDER_NO"
Text2.DataField = "TENDER_NO"
Text3.DataField = "SUPPLIER_ID"
Text4.DataField = "DATE_OF_ORDER"

End Sub

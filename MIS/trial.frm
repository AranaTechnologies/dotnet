VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "trial.frx":0000
      Left            =   480
      List            =   "trial.frx":0002
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "prev"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "next"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "addnew"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "submit"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "delete"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
Dim response As Integer
Dim message As String
message = "Delete the record of " & UCase(Text1.Text) & "?"
response = MsgBox(message, 36, "Delete Record")
If response = 6 Then
If rst.EOF = True Then
MsgBox ("Eof has occured")
Else
rst.Delete
rst.Update
If Not rst.BOF Or Not rst.EOF Then
If rst.RecordCount > 1 Then
rst.MoveNext
End If
End If

If rst.EOF = True Then
If rst.RecordCount > 1 Then
rst.MovePrevious
End If
End If
End If
End If
rst.Close
rst.Open "select * from trial", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
End Sub

Private Sub Command2_Click()
rst.Update
Command2.Enabled = False
Command3.Enabled = True
End Sub

Private Sub Command3_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
Command3.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command4_Click()
rst.MoveNext
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
End Sub

Private Sub Command5_Click()
rst.MovePrevious
If rst.BOF Then
MsgBox ("You are on the First Record")
If rst.RecordCount <> 0 Then
rst.MoveFirst
End If
End If
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from trial", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst.Fields("aaa")
rst.MoveNext
Wend
Set Text1.DataSource = rst
Text1.DataField = "aaa"

End Sub

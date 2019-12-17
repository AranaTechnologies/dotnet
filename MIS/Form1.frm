VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "delete"
      Height          =   495
      Left            =   960
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "previous"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "next"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "add new"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      Begin VB.CommandButton Command11 
         Caption         =   "cancel"
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "delete"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "previous"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "next"
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "add"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "submit"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset

'ok button
Private Sub Command1_Click()
Frame1.Visible = True
End Sub
'delete1 button
Private Sub Command10_Click()
Dim rst1 As New ADODB.Recordset
Dim sql
Dim temp
temp = rst("name")
rst1.Open "select name from trial where trial.name='" & temp & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText


While Not rst1.EOF = True
rst1.Delete
rst1.Update
rst1.MoveNext
Wend
rst.Close
rst.Open "select * from trial", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
End Sub
'cancel button
Private Sub Command11_Click()
rst.CancelUpdate
Text1.Text = rst("name")
Text2.Text = rst("id")
Text3.Text = rst("amt")
End Sub

'submit button
Private Sub Command2_Click()
rst.Update

End Sub
'add button
Private Sub Command3_Click()
Dim temp
temp = rst("name")
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
rst.Fields("name") = temp
End Sub
'add1 button
Private Sub Command4_Click()
Text1.Enabled = True
Text1.SetFocus
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
End Sub
'next1 button
Private Sub Command5_Click()
Dim temp
temp = rst("name")
While rst("name") = temp
rst.MoveNext
If rst.EOF Then
MsgBox ("you are on the last record")
rst.MoveLast
Exit Sub
End If
Wend
End Sub
'previous1 button
Private Sub Command6_Click()
Dim temp
temp = rst("name")
While rst("name") = temp
rst.MovePrevious
If rst.BOF Then
MsgBox ("you are on the first record")
rst.MoveFirst
Exit Sub
End If
Wend
End Sub
'next button
Private Sub Command7_Click()
rst.MoveNext
If rst.EOF = True Then
MsgBox ("you are on the last record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
End Sub
'previous button
Private Sub Command8_Click()
rst.MovePrevious
If rst.BOF = True Then
MsgBox ("you are on the first record")
If rst.RecordCount <> 0 Then
rst.MoveFirst
End If
End If
End Sub
'delete button
Private Sub Command9_Click()
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
Set Text2.DataSource = rst
Set Text3.DataSource = rst
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=trial; pwd=trial1"
rst.Open "select * from trial", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst

Text1.DataField = "name"
Text2.DataField = "id"
Text3.DataField = "amt"
Frame1.Visible = False
End Sub

Private Sub Text1_LostFocus()
Frame1.Visible = True
Text2.SetFocus
Text1.Enabled = False
End Sub

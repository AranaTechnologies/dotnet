VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "submit"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
End Sub

Private Sub Command2_Click()
Dim rst1 As New ADODB.Recordset
rst1.Open "select a from xyz", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst1.EOF = False
If Text1.Text = rst1("a") Then
AUTO_ERROR (1)
Exit Sub
End If
rst1.MoveNext
Wend
If Len(Text2.Text) = 0 Then
AUTO_ERROR (1)
Exit Sub
End If


rst.Update
End Sub

Private Sub Command3_Click()
'On Error GoTo e1
rst.MoveNext
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
'e1: AUTO_ERROR (1)
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from xyz", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst


Text1.DataField = "a"
Text2.DataField = "b"

End Sub

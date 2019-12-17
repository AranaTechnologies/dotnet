VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "submit"
      Height          =   855
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset


Private Sub Command2_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
rst.AddNew
End If
End Sub

Private Sub Command3_Click()
rst.Update
Dim temp
Dim sql

temp = 3
sql = "update b set x=" & temp & "where y='" & Text2.Text & "'"
Set r = cnn.Execute(sql)
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from a", cnn, adOpenStatic, adLockOptimistic, adCmdText
'rst1.Open "select * from b", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst

Text1.DataField = "x"
Text2.DataField = "y"

End Sub

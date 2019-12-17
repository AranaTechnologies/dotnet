VERSION 5.00
Begin VB.Form finished_product_rec_entry 
   BackColor       =   &H00C0FFFF&
   Caption         =   "FINISHED PRODUCT REC ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "UPDATE STOCK"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Next"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Previous"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Submit"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Add New"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mmm-yy;e.g. 01-nov-03)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUE NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT ID."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FINISHED PRODUCT REC ENTRY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "finished_product_rec_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Public Sub disable()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub
Public Sub enable()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
End Sub
'ADD NEW RECORD BUTTON
Private Sub Command1_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
enable
Command1.Enabled = False
Command2.Enabled = True
End Sub
'SUBMIT BUTTON
Private Sub Command2_Click()
rst.Update
Command2.Enabled = False
Command1.Enabled = True
End Sub


Private Sub Command3_Click()
'code of update stock button
Dim sql As String
Dim temp
Dim rst1 As New ADODB.Recordset
Dim curr_stock As Integer
Dim prev_stock As Integer
If rst("status") = "done" Then
MsgBox ("This record is already updated")
Exit Sub
Else
temp = rst("issue_no")
sql = "update received_product set status='done' where issue_no='" & temp & "'"
Set r1 = cnn.Execute(sql)
sql = "select current_stock from product where product_id='" & rst("product_id") & "'"
Set rst1 = cnn.Execute(sql)
prev_stock = rst1("current_stock")
curr_stock = prev_stock + rst("quantity")
sql = "insert into product_stock values('" & rst("issue_no") & "','" & rst("product_id") & "'," & rst("quantity") & ",'Received','" & Text4.Text & "'," & curr_stock & ")"
Set r2 = cnn.Execute(sql)
sql = "update product set current_stock=" & curr_stock & "where product_id='" & rst("product_id") & "'"
Set r3 = cnn.Execute(sql)
disable
End If
End Sub

Private Sub Command4_Click()
'code of previous button
rst.MovePrevious
If rst.BOF Then
MsgBox ("You are on the First Record")
If rst.RecordCount <> 0 Then
rst.MoveFirst
End If
End If
If rst("status") = "done" Then
disable
Else
enable
End If
End Sub

Private Sub Command5_Click()
'code of next button
rst.MoveNext
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
If rst("status") = "done" Then
disable
Else
enable
End If
End Sub

Private Sub Command6_Click()
'code of delete button
Dim response As Integer
Dim message As String
message = "Delete the record of " & UCase(Text1.Text) & "?"
response = MsgBox(message, 36, "Delete Record")
If response = 6 Then
If rst.EOF = True Then
MsgBox ("Eof has occured")
Else
If rst("status") = "done" Then
MsgBox ("You can't delete the record")
Exit Sub
Else
rst.Delete
rst.Update
End If
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
rst.Open "select * from received_product", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
End Sub

Private Sub Command7_Click()
'code of cancel button
rst.CancelUpdate
Text1.Text = rst("issue_no")
Text2.Text = rst("product_id")
Text3.Text = rst("quantity")
Text4.Text = rst("issue_date")
Text5.Text = rst("status")
End Sub

'EXIT BUTTON
Private Sub Command8_Click()
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub
Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from received_product", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Text1.DataField = "issue_no"
Text2.DataField = "product_id"
Text3.DataField = "quantity"
Text4.DataField = "issue_date"
Text5.DataField = "status"
If rst("status") = "done" Then
disable
Else
enable
End If
End Sub


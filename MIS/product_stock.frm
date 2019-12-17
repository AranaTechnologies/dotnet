VERSION 5.00
Begin VB.Form product_stock 
   BackColor       =   &H00004080&
   Caption         =   "PRODUCT STOCK ENTRY"
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
      Left            =   6360
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
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
      Height          =   405
      Left            =   6360
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&ADD NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "&SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "&PREVIOUS"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "&NEXT"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "&DELETE"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FFFF&
      Caption         =   "&EDIT"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FFFF&
      Caption         =   "&CANCEL"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FFFF&
      Caption         =   "E&XIT"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   3240
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   2400
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT STOCK ENTRY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "product_stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
'ADD NEW RECORD BUTTON
Private Sub Command1_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
Command1.Enabled = False
Command2.Enabled = True
End Sub
'SUBMIT BUTTON
Private Sub Command2_Click()
Dim sql As String
Dim prev_stock
Dim curr_stock
Dim rst1 As New ADODB.Recordset
rst1.Open "select product_id,total_stock from product_stock where upper(product_id)='" & UCase(Text1.Text) & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
If rst1.RecordCount <> 0 Then
rst1.MoveLast
prev_stock = rst1("total_stock")
Else
prev_stock = 0
End If
curr_stock = prev_stock + rst("quantity")
sql = "update product set current_stock = curr_stock where upper(product_id)='" & UCase(Text1.Text) & "'"
Set r = cnn.Execute(sql)
rst("transaction_code") = "Received"
rst("total_stock") = curr_stock
rst.Update
Command2.Enabled = False
Command1.Enabled = True
End Sub
'PREVIOUS BUTTON
Private Sub Command3_Click()
rst.MovePrevious
If rst.BOF Then
MsgBox ("You are on the First Record")
If rst.RecordCount <> 0 Then
rst.MoveFirst
End If
End If
End Sub
'NEXT BUTTON
Private Sub Command4_Click()
rst.MoveNext
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
End Sub
'DELETE BUTTON
Private Sub Command5_Click()
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
rst.Open "select * from product_stock", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst

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
rst.Open "select * from product_stock", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Text1.DataField = "product_id"
Text2.DataField = "quantity"
Text3.DataField = "stock_date"
End Sub


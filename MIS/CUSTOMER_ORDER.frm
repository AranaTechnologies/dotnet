VERSION 5.00
Begin VB.Form customer_order 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CUSTOMER ORDER ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form6"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
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
      Left            =   6000
      TabIndex        =   21
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   6000
      TabIndex        =   20
      Top             =   4080
      Width           =   1455
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
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   6000
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
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
      Height          =   345
      Left            =   6000
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text5 
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
      Height          =   345
      Left            =   6000
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
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
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF00&
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mmm-yy; e.g. 01-mar-03)"
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mmm-yy; e.g. 01-mar-03)"
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUE DATE"
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
      Left            =   360
      TabIndex        =   18
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CUSTOMER ORDER NO."
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
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CUSTOMER ID."
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
      Left            =   360
      TabIndex        =   16
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER PRODUCT ID."
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
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER QUANTITY"
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
      Left            =   360
      TabIndex        =   14
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER DATE "
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
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ORDER ENTRY"
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
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "customer_order"
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
Text5.Enabled = False
End Sub
Public Sub enable()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Sub
'ADD NEW RECORD BUTTON
Private Sub Command1_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
rst("status") = "Pending"
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
'PREVIOUS BUTTON
Private Sub Command3_Click()
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
'NEXT BUTTON
Private Sub Command4_Click()
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
rst.Open "select * from customer_order", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Set Text7.DataSource = rst
End Sub



Private Sub Command7_Click()
'code of cancel button
rst.CancelUpdate
Text1.Text = rst("customer_order_no")
Text2.Text = rst("customer_id")
Text3.Text = rst("product_id")
Text4.Text = rst("quantity")
Text5.Text = rst("order_date")
Text7.Text = rst("status")
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
rst.Open "select * from customer_order", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Set Text7.DataSource = rst
Text1.DataField = "customer_order_no"
Text2.DataField = "customer_id"
Text3.DataField = "product_id"
Text4.DataField = "quantity"
Text5.DataField = "order_date"
Text6.DataField = "issue_date"
Text7.DataField = "status"
If rst("status") = "done" Then
disable
Else
enable
End If
End Sub



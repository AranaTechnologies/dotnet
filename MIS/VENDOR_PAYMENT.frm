VERSION 5.00
Begin VB.Form vendor_payment 
   BackColor       =   &H00004080&
   Caption         =   "VENDOR PAYMENT ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   3375
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
      Left            =   3840
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "VENDOR_PAYMENT.frx":0000
      Left            =   3840
      List            =   "VENDOR_PAYMENT.frx":000D
      TabIndex        =   2
      Text            =   "CASH"
      Top             =   2520
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
      Height          =   345
      Left            =   3840
      TabIndex        =   3
      Top             =   3480
      Width           =   3135
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
      Left            =   3840
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
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
      Left            =   600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VENDOR PAYMENT ENTRY"
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
      Left            =   1800
      TabIndex        =   18
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER BILL NO."
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
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER AMOUNT"
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
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MODE OF PAYMENT"
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
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER DRAFT/CHEQUE ID."
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
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER DATE OF PAYMENT"
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
      Left            =   360
      TabIndex        =   12
      Top             =   4320
      Width           =   3495
   End
End
Attribute VB_Name = "vendor_payment"
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
rst.Open "select * from vendor_payment", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Combo1.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
End Sub
'EDIT BUTTON
Private Sub Command6_Click()
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
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
rst.Open "select * from vendor_payment", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Combo1.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Text1.DataField = "bill_no"
Text2.DataField = "amount"
Combo1.DataField = "mode_of_payment"
Text3.DataField = "d_c_id"
Text4.DataField = "date_of_payment"
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Text3.Enabled = False
Text4.Enabled = False

End Sub



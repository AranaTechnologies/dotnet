VERSION 5.00
Begin VB.Form CUSTOMER 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CUSTOMER INFORMATION ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
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
      TabIndex        =   11
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
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   975
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   5
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   9
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   7
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER DESIGNATION OF CONTACT PERSON"
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
      Left            =   480
      TabIndex        =   29
      Top             =   4440
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NAME OF CONTACT PERSON"
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
      Left            =   480
      TabIndex        =   28
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER INFORMATION ENTRY"
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
      Left            =   1320
      TabIndex        =   27
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NAME OF THE ORGANIZATION"
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
      Left            =   480
      TabIndex        =   25
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER BUILDING NO."
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
      Left            =   480
      TabIndex        =   24
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER STREET NAME"
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
      Left            =   480
      TabIndex        =   23
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CITY"
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
      Left            =   480
      TabIndex        =   22
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER STATE"
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
      Left            =   480
      TabIndex        =   21
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER PIN"
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
      Left            =   480
      TabIndex        =   20
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER PH. NO. OF CONTACT PERSON"
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
      Left            =   480
      TabIndex        =   19
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER E-MAIL ID."
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
      Left            =   480
      TabIndex        =   17
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "CUSTOMER"
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
rst.Open "select * from customer", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Set Text7.DataSource = rst
Set Text8.DataSource = rst
Set Text9.DataSource = rst
Set Text10.DataSource = rst
Set Text11.DataSource = rst
End Sub

Private Sub Command7_Click()
'code of cancel button
rst.CancelUpdate
Text1.Text = rst("customer_id")
Text2.Text = rst("name_of_the_organization")
Text3.Text = rst("building_no")
Text4.Text = rst("street_name")
Text5.Text = rst("city")
Text6.Text = rst("state")
Text7.Text = rst("pin")
Text8.Text = rst("email_id")
Text9.Text = rst("name_of_contact_person")
Text10.Text = rst("ph_no_of_contact_person")
Text11.Text = rst("designation_of_contact_person")
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
rst.Open "select * from customer", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Set Text7.DataSource = rst
Set Text8.DataSource = rst
Set Text9.DataSource = rst
Set Text10.DataSource = rst
Set Text11.DataSource = rst

Text1.DataField = "customer_id"
Text2.DataField = "name_of_the_organization"
Text3.DataField = "building_no"
Text4.DataField = "street_name"
Text5.DataField = "city"
Text6.DataField = "state"
Text7.DataField = "pin"
Text8.DataField = "email_id"
Text9.DataField = "name_of_contact_person"
Text10.DataField = "ph_no_of_contact_person"
Text11.DataField = "designation_of_contact_person"
End Sub


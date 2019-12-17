VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form15"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form15"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "E-Mail"
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   360
      Left            =   10320
      TabIndex        =   17
      Top             =   7920
      Width           =   960
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Previous"
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Submit"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   12
      Top             =   7920
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Item Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5520
      TabIndex        =   10
      Top             =   5400
      Width           =   5415
      Begin VB.TextBox Text5 
         Height          =   1335
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5760
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity in Stock in ml"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level in ml"
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
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Id"
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
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   120
      Picture         =   "Form15.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Information Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset
'SEND E-MAIL BUTTON'
Private Sub Command1_Click()
Form13.Show
End Sub

Private Sub Command10_Click()

If Not RST.BOF Or Not RST.EOF Then
RST.MoveNext
If Not RST.EOF Then
End If
End If


If RST.EOF Then
MsgBox ("You are on the Last Record")
If RST.RecordCount <> 0 Then
RST.MoveLast
End If
End If
End Sub

Private Sub Command11_Click()
'rst.Update
Command11.Enabled = False
Command3.Enabled = True
End Sub
'PREVIOUS BUTTON'
Private Sub Command12_Click()
If Not RST.BOF Or Not RST.EOF Then
RST.MovePrevious
If Not RST.BOF Then
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
'ADD BUTTON'
Private Sub Command3_Click()
Dim filename As String

If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command11.Enabled = True
sql = "SELECT * FROM GENERATED_LAST"
Set rst1 = CNN.Execute(sql)
temp = rst1.Fields("ITEM_ID")
sql = "UPDATE GENERATED_LAST SET ITEM_ID=ITEM_ID+1"
Set rst1 = CNN.Execute(sql)
temp = Trim(temp)
temp1 = "ITEM" & temp
RST.Fields("item_id") = temp1

End Sub
'EXIT BUTTON'
Private Sub Command5_Click()
RST.Close
CNN.Close
Unload Me
End Sub

Private Sub Command6_Click(Index As Integer)

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
RST.Open "select * from item_file ", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST



End Sub

Private Sub Form_Activate()
Command11.Enabled = False
End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from ITEM_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST

Text1.DataField = "ITEM_ID"
Text2.DataField = "ITEM_NAME"
Text3.DataField = "REORDER_LEVEL"
Text4.DataField = "QUANTITY_IN_STOCk"
Text5.DataField = "ITEM_DESCRIPTION"


End Sub

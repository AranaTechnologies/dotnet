VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H0080C0FF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "E-Mail"
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   7800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Department Details"
      Height          =   5775
      Left            =   4080
      TabIndex        =   7
      Top             =   1560
      Width           =   5775
      Begin VB.TextBox Text4 
         Height          =   1335
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3120
         TabIndex        =   17
         Top             =   3360
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   3120
         TabIndex        =   15
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   435
         Left            =   3120
         TabIndex        =   13
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "Form9.frx":0000
         Left            =   3000
         List            =   "Form9.frx":0016
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Telephone No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Head of the Department"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Department Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Department Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   360
      Left            =   10320
      TabIndex        =   6
      Top             =   7800
      Width           =   720
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Previous"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
      Height          =   390
      Left            =   600
      TabIndex        =   4
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Next"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Submit"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   1
      Top             =   7800
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   240
      Picture         =   "Form9.frx":0071
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   " Department Information Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim CNN1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim CNN2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
'SEND EMAIL BUTTON'
Private Sub Command1_Click()
If Len(Trim(rst.Fields("SEND_EMAIL"))) <> 0 Then
Form13.Text1.Text = rst.Fields("SEND_EMAIL")
End If
rst.Close
cnn.Close
Form13.Show
End Sub

'NEXT BUTTON'
Private Sub Command10_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveNext
If Not rst.EOF Then
End If
End If
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
End Sub
'SUBMIT'
Private Sub Command11_Click()
If Not rst.EOF Or Not rst.BOF Then
rst.Update
Command11.Enabled = False
Command3.Enabled = True
End If
End Sub

'ADD NEW'
Private Sub Command3_Click()
Dim filename As String
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
Command3.Enabled = False
Command11.Enabled = True
sql = "SELECT * FROM GENERATED_LAST"
Set RST1 = cnn.Execute(sql)
temp = RST1.Fields("DEPARTMENT_ID")
sql = "UPDATE GENERATED_LAST SET DEPARTMENT_ID=DEPARTMENT_ID+1"
Set RST1 = cnn.Execute(sql)
RST1.Close
temp = Trim(temp)
temp1 = "TEST" & temp
rst.Fields("DEPARTMENT_ID") = temp1
 
 
End Sub
'PREVIOUS'
Private Sub Command12_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MovePrevious
If Not rst.BOF Then
End If
End If
If rst.BOF Then
MsgBox ("You are on the First Record")
If rst.RecordCount <> 0 Then
rst.MoveFirst
End If
If Not rst.BOF Or Not rst.EOF Then
rst.MoveFirst
End If
End If

End Sub


'DELETE'
Private Sub Command6_Click(Index As Integer)
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
rst.Open "select * from department", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Combo1.DataSource = rst
Set Combo2.DataSource = rst

End Sub

Private Sub Form_Activate()
Command11.Enabled = False
End Sub

Private Sub Form_Load()
cnn.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
rst.Open "select * from Department", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Combo1.DataSource = rst
Set Combo2.DataSource = rst
Text1.DataField = "DEPARTMENT_ID"
Text2.DataField = "Location"
Text3.DataField = "Telephone_No"
Text4.DataField = "Description"
Combo1.DataField = "Department_name"
Combo2.DataField = "Head_of_the_department"

'ADD ITEM IN COMBOBOX'
'CNN2.Open "DSN=from oracle; PROVIDER=MSDASQL; UID=imdms; PWD=imdms1"
'RST2.Open " SELECT department_name FROM department WHERE UPPER(designation) = 'HEAD OF THE DEPARTMENT'", cnn, adOpenStatic, adLockOptimistic, adCmdText
'While RST2.EOF = False
'Combo1.AddItem (RST2.Fields("Department_name"))
'RST2.MoveNext
'Wend
'RST2.Close
'CNN2.Close
'Combo2.DataField = "head_of_the_department"
End Sub


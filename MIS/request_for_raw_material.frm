VERSION 5.00
Begin VB.Form request_for_raw_material 
   BackColor       =   &H00C0FFFF&
   Caption         =   "REQUEST FOR RAW MATERIALS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFF00&
      Caption         =   "OK"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5280
      TabIndex        =   42
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFF00&
      Caption         =   "OPEN"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "NEW REQUEST"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXISTING REQUEST"
      ForeColor       =   &H00004040&
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   9135
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFF00&
         Caption         =   "Previous"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Next"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Submit"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Add New"
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   6360
         TabIndex        =   33
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   2400
         TabIndex        =   29
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   6360
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2400
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   4800
         TabIndex        =   32
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Raw Material ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dept Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4680
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Request No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "REQUIRED RAW MATERIALS ENTRY"
      ForeColor       =   &H00004040&
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   9135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "ADD NEW"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "SUBMIT"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "DELETE"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF00&
         Caption         =   "NEXT"
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF00&
         Caption         =   "PREVIOUS"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF00&
         Caption         =   "CANCEL"
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Caption         =   "EXIT"
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER RAW MATERIAL ID"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "REQUEST INFORMATION ENTRY"
      ForeColor       =   &H00004040&
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   9135
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   7560
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER REQUEST DATE"
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
         Left            =   4440
         TabIndex        =   20
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER REQUEST NO"
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
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER DEPT. NAME"
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
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   3735
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REQUEST FOR RAW MATERIALS ENTRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
End
Attribute VB_Name = "request_for_raw_material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset
Dim rst4 As New ADODB.Recordset
Dim rst5 As New ADODB.Recordset
Dim rst6 As New ADODB.Recordset
Public Sub test()
'test whether the tender is pending or not
If rst6("status") = "Pending" Then
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
Command16.Enabled = True
Else
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Command16.Enabled = False
End If
End Sub

Private Sub Command1_Click()
'code of add button
Dim temp
temp = rst("request_no")
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
rst.Fields("request_no") = temp
Command2.Enabled = True
Command1.Enabled = False
Text2.SetFocus
End Sub

Private Sub Command10_Click()
'code of cancel(frame2) button
rst1.CancelUpdate
rst1.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command11_Click()
'code of addnew(frame3) button
Dim temp
temp = rst5("tender_no")
If Not rst5.BOF Or Not rst5.EOF Then
rst5.MoveLast
End If
rst5.AddNew
rst5("tender_no") = temp
Command11.Enabled = False
Command12.Enabled = True
Text9.SetFocus
End Sub

Private Sub Command12_Click()
'code of submit(frame3) button
rst5.Update
rst6.Update
Command12.Enabled = False
Command11.Enabled = True
End Sub

Private Sub Command14_Click()
'code of next(frame3) button
rst5.MoveNext
If rst5.EOF Then
MsgBox ("You are on the Last Record")
If rst5.RecordCount <> 0 Then
rst5.MoveLast
End If
End If
End Sub

Private Sub Command15_Click()
'code of previous(frame3) button
rst5.MovePrevious
If rst5.BOF Then
MsgBox ("You are on the First Record")
If rst5.RecordCount <> 0 Then
rst5.MoveFirst
End If
End If
End Sub

Private Sub Command17_Click()
'code of exit(frame3) button
rst5.Close
rst6.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command18_Click()
'code of Open  button
Combo1.Enabled = True
Command19.Enabled = True
rst4.Open "select  request_no from request_info", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst4.EOF = False
Combo1.AddItem rst4.Fields("request_no")
rst4.MoveNext
Wend
End Sub

Private Sub Command19_Click()
'code of ok button
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
rst5.Open "select * from issue_register  where  request_no='" & Combo1.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
rst6.Open "select * from request_info where  request_no='" & Combo1.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text6.DataSource = rst5
Set Text9.DataSource = rst5
Set Text10.DataSource = rst5
Set Text7.DataSource = rst6
Set Text8.DataSource = rst6
Text6.DataField = "request_no"
Text9.DataField = "raw_material_id"
Text10.DataField = "quantity"
Text7.DataField = "request_date"
Text8.DataField = "dept_name"
Combo1.Enabled = False
Command19.Enabled = False
test
End Sub

Private Sub Command2_Click()
'SUBMIT BUTTON
rst.Update
Command2.Enabled = False
Command1.Enabled = True
Command1.SetFocus
End Sub

Private Sub Command4_Click()
'code of next button
Dim temp
temp = rst("request_no")
rst.MoveNext
If rst.EOF = True Then
MsgBox ("you are on the last record")
rst.MovePrevious
Else
If rst("request_no") = temp Then
Set Text2.DataSource = rst
Set Text4.DataSource = rst
Text2.DataField = "raw_material_id"
Text4.DataField = "quantity"
Else
rst2.MovePrevious
Set Text2.DataSource = rst
Set Text4.DataSource = rst
Text2.DataField = "raw_material_id"
Text4.DataField = "quantity"
MsgBox ("you are on the last record")
End If
End If
End Sub

Private Sub Command5_Click()
'code of previous button
Dim temp
temp = rst("request_no")
rst.MovePrevious
If rst.BOF = True Then
MsgBox ("you are on the first record")
rst.MoveFirst
Else
If rst("request_no") = temp Then
Set Text2.DataSource = rst
Set Text4.DataSource = rst
Text2.DataField = "raw_material_id"
Text4.DataField = "quantity"
Else
rst.MoveNext
Set Text2.DataSource = rst
Set Text4.DataSource = rst
Text2.DataField = "raw_material_id"
Text4.DataField = "quantity"
MsgBox ("you are on the first record")
End If
End If

End Sub

Private Sub Command7_Click()
'EXIT BUTTON
rst.Close
rst1.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command8_Click()
'code of new request button
Frame3.Visible = False
Frame2.Visible = True
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
If Not rst1.BOF Or Not rst1.EOF Then
rst1.MoveLast
End If
rst1.AddNew
End Sub

Private Sub Command9_Click()
'code of ok (frame2)button
Dim sql As String
rst("request_no") = Text1.Text
rst1.Update
sql = "update request_info set status='Pending' where upper(request_no)='" & UCase(Text1.Text) & "'"
Set r = cnn.Execute(sql)
Frame1.Visible = True
Frame2.Enabled = False
Text2.SetFocus
End Sub
Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from issue_register", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text4.DataSource = rst
Text1.DataField = "request_no"
Text2.DataField = "raw_material_id"
Text4.DataField = "quantity"
rst1.Open "select * from request_info", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst1
Set Text3.DataSource = rst1
Set Text5.DataSource = rst1
Text1.DataField = "request_no"
Text3.DataField = "dept_name"
Text5.DataField = "request_date"
Command1.Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

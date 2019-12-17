VERSION 5.00
Begin VB.Form SUBMITTED_TENDER_FORM 
   BackColor       =   &H00C0FFFF&
   Caption         =   "SUBMITTED TENDER FORM ENTRY"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TENDER DETAILS"
      ForeColor       =   &H00004040&
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   9135
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         Height          =   375
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Caption         =   "Previous"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF00&
         Caption         =   "Next"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Add New"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Submit"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE/UNIT"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "RAW MATERIAL ID"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tender Details"
      ForeColor       =   &H00004040&
      Height          =   1935
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   9135
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFF00&
         Caption         =   "Previous"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFF00&
         Caption         =   "Next"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFF00&
         Caption         =   "Add New"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFF00&
         Caption         =   "Submit"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   6480
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE/UNIT"
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   4560
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RAW MTERIAL ID"
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFF00&
      Caption         =   "OPEN"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TENDER INFORMATION"
      ForeColor       =   &H00004040&
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   9135
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text6 
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
         Left            =   5040
         TabIndex        =   3
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "VENDOR ID"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TENDER NO"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER  DATE"
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
         Left            =   0
         TabIndex        =   4
         Top             =   3360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "NEW ENTRY"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUBMITTED TENDER FORM ENTRY"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "SUBMITTED_TENDER_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim flag


Private Sub Command1_Click()
'code of new entry button
flag = 1
Frame1.Visible = True
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
End Sub

Private Sub Command10_Click()
'code of open button
Frame1.Visible = True
flag = 2
End Sub

Private Sub Command11_Click()
'code of submit(frame3) button
rst3.Update
Command11.Enabled = False
Command12.Enabled = True
End Sub

Private Sub Command12_Click()
'code of add new(frame3) button
If Not rst3.BOF Or Not rst3.EOF Then
rst3.MoveLast
End If
rst3.AddNew
rst3("tender_no") = Combo1.Text
rst3("vendor_id") = Combo2.Text
Command12.Enabled = False
Command11.Enabled = True
End Sub

Private Sub Command14_Click()
'code of next(frame3) button
rst3.MoveNext
If rst3.EOF Then
MsgBox ("You are on the Last Record")
If rst3.RecordCount <> 0 Then
rst3.MoveLast
End If
End If
End Sub

Private Sub Command15_Click()
'code of previous(frame3) button
rst3.MovePrevious
If rst3.BOF Then
MsgBox ("You are on the First Record")
If rst3.RecordCount <> 0 Then
rst3.MoveFirst
End If
End If
End Sub

Private Sub Command16_Click()
'code of cancel(frame3) button
rst3.CancelUpdate
Text1.Text = rst3("raw_material_id")
Text4.Text = rst3("price")
End Sub

Private Sub Command17_Click()
'code of exit(frame3) button
rst3.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Command2_Click()
'code of ok button
Dim sql As String
Dim rst4 As New ADODB.Recordset
If flag = 1 Then
rst("tender_no") = Combo1.Text
rst("vendor_id") = Combo2.Text
Frame2.Visible = True
Frame3.Visible = False
Text2.SetFocus
Else
Frame2.Visible = False
Frame3.Visible = True
rst3.Open "select * from submitted_tender_form where tender_no='" & Combo1.Text & "'" & " and vendor_id= '" & Combo2.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst3
Set Text4.DataSource = rst3
Text1.DataField = "raw_material_id"
Text4.DataField = "price"
sql = "select * from tender_info where tender_no='" & Combo1.Text & "'"
Set rst4 = cnn.Execute(sql)
If rst4("status") <> "Called" Then
Text1.Enabled = False
Text4.Enabled = False
Command11.Enabled = False
Command12.Enabled = False
Command13.Enabled = False
Command16.Enabled = False
End If
End If
End Sub

Private Sub Command3_Click()
'code of SUBMIT BUTTON
rst.Update
Command3.Enabled = False
Command4.Enabled = True
Command4.SetFocus
End Sub

Private Sub Command4_Click()
'code of add new button
Dim temp1
Dim temp2
temp1 = rst("tender_no")
temp2 = rst("vendor_id")
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
rst.Fields("tender_no") = temp1
rst.Fields("vendor_id") = temp2
Command3.Enabled = True
Command4.Enabled = False
Text2.SetFocus
End Sub

Private Sub Command6_Click()
'code of next button
Dim temp1
Dim temp2
temp1 = rst("tender_no")
temp2 = rst("vendor_id")
rst.MoveNext
If rst.EOF = True Then
MsgBox ("you are on the last record")
rst.MovePrevious
Else
If rst("tender_no") = temp1 And rst("vendor_id") = temp2 Then
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Text2.DataField = "raw_material_id"
Text3.DataField = "price"
Else
rst.MovePrevious
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Text2.DataField = "raw_material_id"
Text3.DataField = "price"
MsgBox ("you are on the last record")
End If
End If
End Sub

Private Sub Command7_Click()
'code of previous button
Dim temp1
Dim temp2
temp1 = rst("tender_no")
temp2 = rst("vendor_id")
rst.MovePrevious
If rst.BOF = True Then
MsgBox ("you are on the first record")
rst.MoveNext
Else
If rst("tender_no") = temp1 And rst("vendor_id") = temp2 Then
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Text2.DataField = "raw_material_id"
Text3.DataField = "price"
Else
rst.MoveNext
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Text2.DataField = "raw_material_id"
Text3.DataField = "price"
MsgBox ("you are on the first record")
End If
End If
End Sub

Private Sub Command9_Click()
'code of exit button
rst.Close
rst1.Close
rst2.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Form_Load()
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from submitted_tender_form ", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Combo1.DataSource = rst
Set Combo2.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Combo1.DataField = "tender_no"
Combo2.DataField = "vendor_id"
Text2.DataField = "raw_material_id"
Text3.DataField = "price"
rst1.Open "select  tender_no from tender_info where status='Called'", cnn, adOpenStatic, adLockOptimistic, adCmdText
If rst1.EOF = True Then
MsgBox ("There is no more published tender")
rst.Close
rst1.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
Else
While rst1.EOF = False
Combo1.AddItem rst1.Fields("tender_no")
rst1.MoveNext
Wend
rst2.Open "select vendor_id from vendor", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst2.EOF = False
Combo2.AddItem rst2.Fields("vendor_id")
rst2.MoveNext
Wend
Command4.Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End If
End Sub




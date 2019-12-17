VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0080C0FF&
   Caption         =   "Patient Details  Entry"
   ClientHeight    =   8595
   ClientLeft      =   -3765
   ClientTop       =   -2670
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "E-Mail"
      Height          =   375
      Left            =   8880
      TabIndex        =   28
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add New"
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   360
      Left            =   10440
      TabIndex        =   25
      Top             =   7920
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Previous"
      Height          =   375
      Left            =   7200
      TabIndex        =   24
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Submit"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   7920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Department Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   21
         Top             =   6480
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   19
         Top             =   6000
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   17
         Top             =   5400
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   15
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   12
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   10
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2880
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   435
         Left            =   1920
         TabIndex        =   7
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   435
         Left            =   1920
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   435
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees Status"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fees to be Paid"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Visit"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctors Name"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Address"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Id"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   360
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Patient  Details Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset
'SEND E-MAIL BUTTON'
Private Sub Command2_Click()
If Len(Trim(rst.Fields("SEND_EMAIL"))) <> 0 Then
Form13.Text1.Text = rst.Fields("SEND_EMAIL")
End If
rst.Close
cnn.Close
Form13.Show
End Sub
'PREVIOUS BUTTON'
Private Sub Command3_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MovePrevious
    
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

'DELETE BUTTON'
Private Sub Command6_Click()
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
rst.Open "select * from Patient_Information_file", cnn, adOpenStatic, adLockOptimistic, adCmdText
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

'NEXT BUTTON'
Private Sub Command1_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveNext
End If
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
End Sub

'EXIT BUTTON
Private Sub Command5_Click()
rst.Close
cnn.Close
Unload Me
End Sub

'Submit Button
Private Sub Command7_Click()
If Len(Text1.Text) = 0 Then
SHOW_ERROR (20)
End If
If Not rst.EOF Or Not rst.BOF Then
rst.Update
Command7.Enabled = False
Command8.Enabled = True
End If
End Sub

'ADD NEW'
Private Sub Command8_Click()
Dim filename As String
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
Command8.Enabled = False
Command7.Enabled = True
sql = "SELECT * FROM GENERATED_LAST"
Set rst1 = cnn.Execute(sql)
temp = rst1.Fields("PATIENT_ID")
sql = "UPDATE GENERATED_LAST SET PATIENT_ID=PATIENT_ID+1"
Set rst1 = cnn.Execute(sql)
temp = Trim(temp)
temp1 = "TEST" & temp
rst.Fields("PATIENT_ID") = temp1
End Sub

Private Sub Form_Activate()
Command7.Enabled = False
End Sub

Private Sub Form_Load()
cnn.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
rst.Open "select * from patient_information_file", cnn, adOpenStatic, adLockOptimistic, adCmdText
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

Text1.DataField = "Patient_Id"
Text2.DataField = "patient_name"
Text3.DataField = "patient_address1"
Text4.DataField = "patient_address2"
Text5.DataField = "pin"
Text6.DataField = "DOCTORS_NAME"
Text7.DataField = "DATE_OF_VISIT"
Text8.DataField = "FEES_TO_BE_PAID"
Text9.DataField = "FEES_STATUS"
Text10.DataField = "DATE_OF_BIRTH"
Text11.DataField = "GENDER"

End Sub

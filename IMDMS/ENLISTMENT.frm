VERSION 5.00
Begin VB.Form Form18 
   BackColor       =   &H0080C0FF&
   Caption         =   "Enlistment"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form18"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "&Previous"
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
      Left            =   6720
      TabIndex        =   32
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Exit"
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
      Left            =   10200
      TabIndex        =   31
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Next"
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
      Left            =   3720
      TabIndex        =   30
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Submit"
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
      Left            =   2160
      TabIndex        =   29
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Delete"
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
      Left            =   5160
      TabIndex        =   28
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add New"
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
      Left            =   360
      TabIndex        =   27
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "E_Mail"
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
      Left            =   8520
      TabIndex        =   26
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   3840
      TabIndex        =   23
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Height          =   405
      Left            =   3840
      TabIndex        =   22
      Text            =   "  "
      Top             =   6450
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "CONTACT PERSON INFORMATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6360
      TabIndex        =   13
      Top             =   1920
      Width           =   5175
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Caption         =   "CONTACT NO."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   4935
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Residential"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "OFFICE INFORMATION"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Height          =   435
         Left            =   3600
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "E - Mail Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Pin"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Registration Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   25
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   " Turn Over Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Enlistment Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
'Submit Button
Private Sub Command2_Click()
If Not RST.EOF Or Not RST.BOF Then
RST.Update
Command2.Enabled = False
Command4.Enabled = True
End If
End Sub
'PREVIOUS'
Private Sub Command6_Click()
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
'NEXT BUTTON'
Private Sub Command3_Click()
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
'ADD NEW BUTTON
Private Sub Command4_Click()
Dim filename As String
If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command4.Enabled = False
Command2.Enabled = True
End Sub


'EXIT BUTTON'
Private Sub Command5_Click()
CNN.Close
RST.Close
Unload Me

End Sub
'DELETE BUTTON'
Private Sub Command7_Click()
Dim response As Integer
Dim message As String
message = "Delete the record of " & UCase(Text2.Text) & "?"
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
RST.Open "select * from ENLISTMENT_FOR_SUPPLIER", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST
Set Text6.DataSource = RST
Set Text7.DataSource = RST
Set Text8.DataSource = RST
Set Text9.DataSource = RST
Set Text10.DataSource = RST
Set Text11.DataSource = RST
Set Text12.DataSource = RST
End Sub

'SEND E-MAIL'
Private Sub Command8_Click()
If Len(Trim(RST.Fields("SEND_EMAIL"))) <> 0 Then
Form13Text1.Text = RST.Fields("SEND_EMAIL")
End If
RST.Close
CNN.Close
Form13.Show
End Sub
 

Private Sub Form_Activate()
 Command4.Enabled = False
End Sub

Private Sub Form_Load()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from ENLISTMENT_FOR_SUPPLIER", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text5.DataSource = RST
Set Text6.DataSource = RST
Set Text7.DataSource = RST
Set Text8.DataSource = RST
Set Text9.DataSource = RST
Set Text10.DataSource = RST
Set Text11.DataSource = RST
Set Text12.DataSource = RST
Text2.DataField = "NAME"
Text3.DataField = CONTACT_NO
Text4.DataField = "address1"
Text5.DataField = "ADDRESS2"
Text6.DataField = "pin"
Text8.DataField = "telephone_no"
Text7.DataField = "EMAIL_ID"
Text9.DataField = "NAME"
Text10.DataField = "DESIGNATION"
Text11.DataField = "FIRM_REGISTRATION_NO"
Text12.DataField = "TURN_OVER_AMOUNT"
End Sub
 
 
 
 
 
 
 



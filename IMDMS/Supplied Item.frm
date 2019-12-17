VERSION 5.00
Begin VB.Form Form41 
   BackColor       =   &H0080C0FF&
   Caption         =   "Supplied Item"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form41"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New "
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Previous"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9120
      TabIndex        =   9
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   6840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Supply Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No "
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
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher_No"
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
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Date_Of_Supply "
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
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Supplied Item Entry"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset

 

'Submit Button
 
Private Sub Command4_Click()
'RST.Update
Command4.Enabled = False
Command3.Enabled = True
RST.Open
End If

End Sub

 
 


'ADD NEW
Private Sub Command3_Click()
Dim filename As String

If Not RST.BOF Or Not RST.EOF Then
RST.MoveLast
End If
RST.AddNew
Command3.Enabled = False
Command4.Enabled = True

End Sub

'EXIT BUTTON
Private Sub Command6_Click()
RST.Close
CNN.Close
Unload Me
End Sub
 

'NEXT BUTTON
Private Sub Command7_Click()
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
'DELETE BUTTON
Private Sub Command8_Click()

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
RST.Open "select * from SUPPLIED_ITEM_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Set Text4.DataSource = RST
Set Text4.DataSource = RST


End Sub

Private Sub Form_Load()

CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from SUPPLIED_ITEM_FILE", CNN, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = RST
Set Text2.DataSource = RST
Set Text3.DataSource = RST
Text1.DataField = "ORDER_NO"
Text2.DataField = "VOUCHER_NO"
Text3.DataField = "DATE_OF_SUPPLY"
End Sub

Private Sub Form_Activate()

Command4.Enabled = False

End Sub

'PREVIOUS BUTTON
Private Sub Command9_Click()
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






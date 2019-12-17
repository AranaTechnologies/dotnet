VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   3720
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   9855
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Exit"
         Height          =   735
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7320
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Label24"
         Height          =   375
         Left            =   4320
         TabIndex        =   25
         Top             =   6720
         Width           =   4455
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Clinic"
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Label22"
         Height          =   375
         Left            =   4320
         TabIndex        =   23
         Top             =   6120
         Width           =   4455
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Time"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         Height          =   375
         Left            =   4320
         TabIndex        =   20
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   5160
         Width           =   4335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   4560
         Width           =   4575
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Reference By"
         Height          =   495
         Left            =   1680
         TabIndex        =   16
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Reference Date"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Patient Type"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mobile No"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Phone"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sex"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Age"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Date"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         Height          =   495
         Left            =   4320
         TabIndex        =   4
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Address"
         Height          =   495
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Patient Name"
         Height          =   495
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Us:scriptonova@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   12000
      TabIndex        =   27
      Top             =   10440
      Width           =   3255
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By: ScriptoNova"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   12000
      TabIndex        =   26
      Top             =   10200
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   3840
      Picture         =   "Form1.frx":05F0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Private Sub Command1_Click()
rst.Close
cnn.Close

Unload Me
End Sub

Private Sub Command2_Click()
Dim fsys As New FileSystemObject
Dim OUTSTREAM As TextStream
Dim PRNTREC As String
Set OUTSTREAM = fsys.CreateTextFile("c:\medical\reports\" & rst.Fields("pt_name") & ".txt", True, False)

PRNTREC = Space(5) & "                 P A T I E N T  R E P O R T"
OUTSTREAM.WriteLine PRNTREC
PRNTREC = Space(5) & "---------------------------------------------------------------"
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "NAME                     : " & rst.Fields("pt_name")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "ADDRESS                  : " & rst.Fields("address")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "DATE OF REGISTRATION     : " & rst.Fields("date_reg")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "AGE                      : " & rst.Fields("age")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "SEX                      : " & rst.Fields("sex")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "PHONE                    : " & rst.Fields("ph")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "MOBILE NO                : " & rst.Fields("mobile_no")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "REFERENCE BY             : " & rst.Fields("refd_by")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "DATE REFERENCE           : " & rst.Fields("date_ref")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "ARRIVAL TIME             : " & rst.Fields("arrival_time")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "PATIENT                  : " & rst.Fields("ptype")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "CLINIC                   : " & rst.Fields("ntype")
OUTSTREAM.WriteLine PRNTREC
OUTSTREAM.WriteLine
PRNTREC = Space(5) & "---------------------------------------------------------------"
OUTSTREAM.WriteLine PRNTREC


MsgBox ("The report is saved in : " & "c:\medical\reports\" & rst.Fields("pt_name") & ".txt")












End Sub

Private Sub Form_Load()

Dim sql As String

sql = "select * from patient where sl_no = " & Form2.Text1.Text
cnn.Open "DSN=fromaccess"
rst.Open sql, cnn, adOpenStatic, adLockOptimistic, adCmdText


Label2.Caption = rst.Fields("pt_name")
Label4.Caption = rst.Fields("address")
Label6.Caption = rst.Fields("date_reg")
Label9.Caption = rst.Fields("age")
Label10.Caption = rst.Fields("sex")
Label12.Caption = rst.Fields("ph")
Label17.Caption = rst.Fields("mobile_no")
Label20.Caption = rst.Fields("refd_by")
Label19.Caption = rst.Fields("date_ref")
Label22.Caption = rst.Fields("arrival_time")
Label18.Caption = rst.Fields("ptype")

Label24.Caption = rst.Fields("ntype")




End Sub

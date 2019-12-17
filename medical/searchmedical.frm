VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form searchmedical 
   BackColor       =   &H00FFFF80&
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14655
   ControlBox      =   0   'False
   ForeColor       =   &H00404080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   1320
      TabIndex        =   1
      Top             =   4920
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   12648447
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search"
      Height          =   1335
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   11175
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Left            =   11280
      TabIndex        =   10
      Top             =   10680
      Width           =   3135
   End
   Begin VB.Label Label3 
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
      Left            =   11280
      TabIndex        =   9
      Top             =   10440
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   8880
      Width           =   11055
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   4200
      Picture         =   "searchmedical.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Search Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "searchmedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn1 As New ADODB.Connection
Dim rst1 As New ADODB.Recordset
'DELETE BUTTON
Private Sub Command1_Click()
rst1.Delete
rst1.Update
End Sub
'DISPLAY BUTTON
Private Sub Command2_Click()
If Len(Form2.Text1.Text) > 0 Then
Form1.Show
Else
MsgBox ("Please select a record first !!!")
End If
End Sub

'EXIT BUTTON
Private Sub Command3_Click()
Unload Me
cnn1.Close
End Sub
'GO BUTTON
Private Sub Command4_Click()
Dim sql As String
If Len(Text1.Text) <> 0 Then
sql = "select * from patient where pt_name like '%" & Text1.Text & "%'"
rst1.Close
cnn1.Close


cnn1.Open "DSN=fromaccess"
rst1.Open sql, cnn1, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst1
Else
MsgBox ("Please enter a name")
End If
End Sub

Private Sub DataGrid1_Click()
Form2.Text1.Text = rst1.Fields("sl_no")
End Sub

Private Sub Form_Load()
Dim sql1 As String
sql1 = "select * from patient"
cnn1.Open "DSN=fromaccess"
rst1.Open sql1, cnn1, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst1
End Sub

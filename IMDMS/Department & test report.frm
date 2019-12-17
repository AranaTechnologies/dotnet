VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form34 
   BackColor       =   &H0080C0FF&
   Caption         =   "Department & Test Report"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11640
   LinkTopic       =   "Form34"
   ScaleHeight     =   8310
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   9240
      TabIndex        =   6
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Test Info"
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record Test Report"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   6720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Tests that are performed in this Department"
      Height          =   975
      Left            =   7440
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Name Of Department"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3615
      End
   End
   Begin VB.Menu Testmnu 
      Caption         =   "Test"
   End
   Begin VB.Menu Testreportmnu 
      Caption         =   "Test Report"
   End
   Begin VB.Menu Exitmnu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset

Private Sub Command1_Click()


If RTrim(UCase(Combo1.Text)) = "T1" Then
Form43.Show
End If

If RTrim(UCase(Combo1.Text)) = "T2" Then
Form46.Show
End If

If RTrim(UCase(Combo1.Text)) = "T3" Then
Form47.Show
End If

If RTrim(UCase(Combo1.Text)) = "T4" Then
Form45.Show
End If

End Sub

Private Sub Command2_Click()
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select test_id from DEPARTMENT_TEST_INFO where rtrim(DEPARTMENT_NAME) = '" & Label1.Caption & "'", CNN, adOpenStatic, adLockOptimistic, adCmdText

While Not RST.EOF
Combo1.AddItem (RST.Fields("TEST_ID"))
RST.MoveNext
Wend

End Sub

Private Sub Command3_Click()

RST.Close
CNN.Close
Unload Me

End Sub

Private Sub Form_Load()

End Sub

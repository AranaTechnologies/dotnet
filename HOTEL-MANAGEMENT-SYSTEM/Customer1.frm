VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Customer_Browser 
   BackColor       =   &H00C0C0FF&
   Caption         =   "CUSTOMER"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   12632319
      HeadLines       =   1
      RowHeight       =   22
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
         Name            =   "Arial"
         Size            =   11.25
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
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   8640
      Picture         =   "Customer1.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   840
      Picture         =   "Customer1.frx":016D
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER  BROWSER  ENTRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "Customer_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub Command1_Click()
rst.Close
cnn.Close
Unload Me

End Sub

Private Sub Form_Load()
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=HTLMS; pwd=HTLMS1"
rst.Open "select * from customer", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
End Sub
